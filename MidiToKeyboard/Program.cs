using Melanchall.DryWetMidi.Core;
using Melanchall.DryWetMidi.Multimedia;
using MidiToKeyboard.Application;
using MidiToKeyboard.Infrastructure;
using MidiToKeyboard.Ui;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace MidiToKeyboard
{
    /// <summary>
    /// MIDI 入力をキーボード入力に変換して送信するコンソールアプリケーションのメインクラス
    /// </summary>
    internal static class Program
    {
        #region 入力モード用変数

        private enum InputMode
        {
            VirtualKey,
            Scancode
        }

        // 起動時にユーザーが選択する（デフォルトは VirtualKey）
        private static InputMode _inputMode = InputMode.VirtualKey;

        #endregion

        #region マッピング JSON モデル

        // JSON ファイルの構造
        // {
        //   "sets": [
        //     { "name": "Default", "mappings": { "52": "v", "53": "s", ... } },
        //     { "name": "Alt", "mappings": { "52": "a", ... } }
        //   ]
        // }

        private class MappingFile
        {
            public List<MappingSet> sets { get; set; }
        }

        private class MappingSet
        {
            public string name { get; set; }
            public Dictionary<string, string> mappings { get; set; }
        }

        // ロード済みセットと現在使用中のマップ
        private static List<MappingSet> _mappingSets = new List<MappingSet>();
        private static Dictionary<int, char> _currentMapping = new Dictionary<int, char>();

        private const string MappingFileName = "mappings.json";

        #endregion

        #region ノート状態管理用変数（重複 KeyUp を防ぐため）

        // _activeNotes: 現在オンと見なしている MIDI ノート集合
        // _keyRefCount: マッピングされたキーボード文字ごとの押下参照カウント
        private static readonly object _stateLock = new object();
        private static readonly HashSet<int> _activeNotes = new HashSet<int>();
        private static readonly Dictionary<char, int> _keyRefCount = new Dictionary<char, int>();
        private static readonly WindowsKeyOutput _keyOutput = new WindowsKeyOutput();

        #endregion

        static Program()
        {
            _keyOutput.WarningOccurred += OnKeyOutputWarningOccurred;
        }

        #region メイン処理

        /// <summary>
        /// アプリケーションのエントリポイント
        /// MIDI デバイス選択、入力モード選択を行い受信を開始する
        /// </summary>
        /// <param name="args">コマンドライン引数（未使用）</param>
        static void Main(string[] args)
        {
            try
            {
                string mappingsPath = Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    MappingFileName);

                using (var midiInput = new DryWetMidiInput())
                {
                    var keyOutput = new WindowsKeyOutput();
                    var profileRepository = new JsonProfileRepository(mappingsPath);
                    var application = new MidiToKeyboardApplication(
                        midiInput,
                        keyOutput,
                        profileRepository);
                    var consoleUi = new ConsoleUi(
                        application,
                        midiInput,
                        profileRepository);

                    keyOutput.WarningOccurred += OnKeyOutputWarningOccurred;
                    consoleUi.Run();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"エラーが発生しました: {ex.Message}");
            }
        }

        #endregion

        #region マッピングファイル操作

        /// <summary>
        /// mappings.json を読み込み、_mappingSets を更新する
        /// </summary>
        private static void LoadMappingSets()
        {
            try
            {
                string mappingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, MappingFileName);
                if (!File.Exists(mappingsPath))
                {
                    Console.WriteLine("マッピングファイルが見つかりません。既定マッピングを使用します。");
                    LoadDefaultMapping();
                    return;
                }

                using (var mappingsStream = File.OpenRead(mappingsPath))
                {
                    _mappingSets = JsonSerializer.Deserialize<MappingFile>(mappingsStream)?.sets ?? new List<MappingSet>();
                }

                if (_mappingSets.Count == 0)
                {
                    Console.WriteLine("マッピングファイルにセットがありません。既定マッピングを使用します。");
                    LoadDefaultMapping();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"マッピングファイルの読み込みに失敗しました: {ex.Message}");
                LoadDefaultMapping();
            }
        }

        /// <summary>
        /// ユーザーにマッピングセット選択を促して_currentMapping を設定する
        /// </summary>
        private static void ChooseMappingSet()
        {
            if (_mappingSets.Count == 0)
            {
                LoadDefaultMapping();
                return;
            }

            Console.WriteLine();
            Console.WriteLine("使用するマッピングセットを選択してください:");
            for (int i = 0; i < _mappingSets.Count; i++)
            {
                Console.WriteLine($"  [{i}] {_mappingSets[i].name}");
            }
            Console.Write("番号を入力: ");

            if (!int.TryParse(Console.ReadLine(), out int mappingIndex) || mappingIndex < 0 || mappingIndex >= _mappingSets.Count)
                mappingIndex = 0;

            ApplyMappingSet(_mappingSets[mappingIndex]);
            Console.WriteLine($"選択セット: {_mappingSets[mappingIndex].name}");
        }

        /// <summary>
        /// 指定セットの mappings を _currentMapping に適用する
        /// </summary>
        /// <param name="set">選択されたマッピングセット</param>
        private static void ApplyMappingSet(MappingSet set)
        {
            _currentMapping.Clear();
            if (set?.mappings == null) return;

            foreach (var mapping in set.mappings)
            {
                if (!int.TryParse(mapping.Key, out int noteNumber)) continue;
                if (string.IsNullOrEmpty(mapping.Value)) continue;
                _currentMapping[noteNumber] = mapping.Value[0];
            }
        }

        /// <summary>
        /// 既定のマッピングを _currentMapping に登録する
        /// </summary>
        private static void LoadDefaultMapping()
        {
            _currentMapping.Clear();
            _currentMapping[52] = 'v';
            _currentMapping[53] = 's';
            _currentMapping[55] = 'd';
            _currentMapping[57] = 'f';
            _currentMapping[59] = 'g';
            _currentMapping[65] = 'h';
            _currentMapping[67] = 'j';
            _currentMapping[69] = 'k';
            _currentMapping[71] = 'l';
            _currentMapping[72] = 'n';
        }

        #endregion

        #region MIDIイベント処理

        /// <summary>
        /// MIDIイベントを受信した際のコールバック
        /// Note On/Off を判定して状態管理関数に委譲する
        /// </summary>
        /// <param name="sender">イベント送信元</param>
        /// <param name="e">受信した MIDI イベント情報</param>
        private static void OnMidiEventReceived(object sender, MidiEventReceivedEventArgs e)
        {
            var midiEvent = e.Event;
            char targetKey;

            if (midiEvent is NoteOnEvent noteOnEvent)
            {
                targetKey = MapMidiNoteToKey(noteOnEvent.NoteNumber);

                if (noteOnEvent.Velocity > 0)
                {
                    NoteOnReceived(noteOnEvent.NoteNumber, targetKey);
                    Console.WriteLine($"[MIDI ON] Note: {noteOnEvent.NoteNumber} -> キー: '{targetKey}' DOWN");
                }
                else
                {
                    NoteOffReceived(noteOnEvent.NoteNumber, targetKey);
                    Console.WriteLine($"[MIDI (vel=0)] Note: {noteOnEvent.NoteNumber} -> キー: '{targetKey}' UP");
                }
            }
            else if (midiEvent is NoteOffEvent noteOffEvent)
            {
                targetKey = MapMidiNoteToKey(noteOffEvent.NoteNumber);
                NoteOffReceived(noteOffEvent.NoteNumber, targetKey);
                Console.WriteLine($"[MIDI OFF] Note: {noteOffEvent.NoteNumber} -> キー: '{targetKey}' UP");
            }
        }

        /// <summary>
        /// Note On を受け取ったときの処理
        /// ノート状態とキー参照カウントを更新し、必要なら KeyDown を送信する
        /// </summary>
        /// <param name="noteNumber">MIDI ノート番号</param>
        /// <param name="key">マッピングされたキーボード文字（存在しない場合は '\0'）</param>
        private static void NoteOnReceived(int noteNumber, char key)
        {
            if (key == '\0') return;

            lock (_stateLock)
            {
                // 既にそのノートがアクティブなら重複のため無視
                if (!_activeNotes.Add(noteNumber))
                    return;

                // key の参照カウントを増やし、初回ならキー押下を送信
                if (_keyRefCount.TryGetValue(key, out int count))
                    _keyRefCount[key] = count + 1;
                else
                    _keyRefCount[key] = 1;

                if (_keyRefCount[key] == 1)
                {
                    // 選択されたモードで送信
                    if (_inputMode == InputMode.Scancode)
                        _keyOutput.SendScancodeKey(key, NativeMethods.KEYEVENTF_KEYDOWN);
                    else
                        _keyOutput.SendKeyInput(key, NativeMethods.KEYEVENTF_KEYDOWN);
                }
            }
        }

        /// <summary>
        /// Note Off を受け取ったときの処理
        /// ノート状態とキー参照カウントを更新し、必要なら KeyUp を送信する
        /// </summary>
        /// <param name="noteNumber">MIDI ノート番号</param>
        /// <param name="key">マッピングされたキーボード文字（存在しない場合は '\0'）</param>
        private static void NoteOffReceived(int noteNumber, char key)
        {
            if (key == '\0') return;

            lock (_stateLock)
            {
                // ノートがアクティブでなければ重複ノートオフと見なして無視
                if (!_activeNotes.Remove(noteNumber))
                    return;

                // key の参照カウントを減らし、0 になればキー離上を送信
                if (_keyRefCount.TryGetValue(key, out int count))
                {
                    count--;
                    if (count <= 0)
                    {
                        _keyRefCount.Remove(key);
                        // 選択されたモードで送信
                        if (_inputMode == InputMode.Scancode)
                            _keyOutput.SendScancodeKey(key, NativeMethods.KEYEVENTF_KEYUP);
                        else
                            _keyOutput.SendKeyInput(key, NativeMethods.KEYEVENTF_KEYUP);
                    }
                    else
                    {
                        _keyRefCount[key] = count;
                    }
                }
                else
                {
                    // カウント情報が無ければ既にキー離上を送った可能性があるので無視
                }
            }
        }

        /// <summary>
        /// MIDI ノート番号を対応するキーボード文字にマップする
        /// </summary>
        /// <param name="noteNumber">MIDI ノート番号</param>
        /// <returns>対応する文字、マッピングなしは '\0' を返す</returns>
        private static char MapMidiNoteToKey(int noteNumber)
        {
            if (_currentMapping.TryGetValue(noteNumber, out char keyChar))
                return keyChar;
            return '\0';
        }

        private static void OnKeyOutputWarningOccurred(object sender, string message)
        {
            Console.WriteLine(message);
        }

        #endregion
    }
}
