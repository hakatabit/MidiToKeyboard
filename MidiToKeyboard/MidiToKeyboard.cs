using Melanchall.DryWetMidi.Core;
using Melanchall.DryWetMidi.Multimedia;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace MidiToKeyboard
{
    /// <summary>
    /// MIDI 入力をキーボード入力に変換して送信するコンソールアプリケーションのメインクラス
    /// </summary>
    class MidiToKeyboard
    {
        #region Native / PInvoke

        /// <summary>
        /// Win32 呼び出し、構造体、定数をまとめたネイティブ用ヘルパクラス
        /// Win32 名は可視性を分離するためこの内部クラス内で保持する
        /// </summary>
        private static class NativeMethods
        {
            [DllImport("user32.dll", SetLastError = true)]
            internal static extern uint SendInput(uint nInputs, [In] INPUT[] pInputs, int cbSize);

            [DllImport("user32.dll")]
            internal static extern short VkKeyScan(char ch);

            [DllImport("user32.dll")]
            internal static extern uint MapVirtualKey(uint uCode, uint uMapType);

            [StructLayout(LayoutKind.Sequential)]
            internal struct INPUT
            {
                public uint type;
                public InputUnion U;
            }

            [StructLayout(LayoutKind.Explicit)]
            internal struct InputUnion
            {
                [FieldOffset(0)] public MOUSEINPUT mi;
                [FieldOffset(0)] public KEYBDINPUT ki;
                [FieldOffset(0)] public HARDWAREINPUT hi;
            }

            [StructLayout(LayoutKind.Sequential)]
            internal struct MOUSEINPUT
            {
                public int dx;
                public int dy;
                public uint mouseData;
                public uint dwFlags;
                public uint time;
                public IntPtr dwExtraInfo;
            }

            [StructLayout(LayoutKind.Sequential)]
            internal struct HARDWAREINPUT
            {
                public uint uMsg;
                public ushort wParamL;
                public ushort wParamH;
            }

            [StructLayout(LayoutKind.Sequential)]
            internal struct KEYBDINPUT
            {
                public ushort wVk;
                public ushort wScan;
                public uint dwFlags;
                public uint time;
                public IntPtr dwExtraInfo;
            }

            // Win32 定数
            internal const int INPUT_KEYBOARD = 1;
            internal const uint KEYEVENTF_KEYDOWN = 0x0000;
            internal const uint KEYEVENTF_KEYUP = 0x0002;
            internal const uint KEYEVENTF_EXTENDEDKEY = 0x0001;
            internal const uint KEYEVENTF_UNICODE = 0x0004;
            internal const uint KEYEVENTF_SCANCODE = 0x0008;

            internal const ushort VK_SHIFT = 0x10;
            internal const ushort VK_CONTROL = 0x11;
            internal const ushort VK_MENU = 0x12;
        }

        #endregion

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

        #endregion

        #region メイン処理

        private static InputDevice _midiDevice;

        /// <summary>
        /// アプリケーションのエントリポイント
        /// MIDI デバイス選択、入力モード選択を行い受信を開始する
        /// </summary>
        /// <param name="args">コマンドライン引数（未使用）</param>
        static void Main(string[] args)
        {
            try
            {
                // 利用可能なMIDIデバイスのリストを表示
                var devices = InputDevice.GetAll().ToList();
                Console.WriteLine("利用可能なMIDI入力デバイス:");

                if (devices.Count == 0)
                {
                    Console.WriteLine("デバイスが見つかりませんでした。");
                    return;
                }

                for (int i = 0; i < devices.Count; i++)
                {
                    Console.WriteLine($"  [{i}]: {devices[i].Name}");
                }

                // ユーザーにデバイスを選択させる
                Console.Write("使用するデバイスの番号を入力してください: ");
                if (int.TryParse(Console.ReadLine(), out int deviceIndex) && deviceIndex >= 0 && deviceIndex < devices.Count)
                {
                    _midiDevice = devices[deviceIndex];

                    // ユーザーに入力送信モードを選ばせる
                    Console.WriteLine();
                    Console.WriteLine("入力送信モードを選択してください:");
                    Console.WriteLine("  [1] 仮想キーコード (VK) - 仮想キー＋修飾キーで送信");
                    Console.WriteLine("  [2] スキャンコード (SC) - スキャンコードで送信（ゲーム等向け）");
                    Console.Write("番号を入力: ");
                    string modeInput = Console.ReadLine()?.Trim();
                    if (modeInput == "2")
                        _inputMode = InputMode.Scancode;
                    else
                        _inputMode = InputMode.VirtualKey;

                    // マッピングセットのロードと選択
                    LoadMappingSets();
                    ChooseMappingSet();

                    // MIDIイベントハンドラを設定
                    _midiDevice.EventReceived += OnMidiEventReceived;

                    // デバイスを開いて受信を開始
                    _midiDevice.StartEventsListening();

                    try
                    {
                        Console.WriteLine($"\n** '{_midiDevice.Name}' をリッスン中です。 **");
                        Console.WriteLine($"選択モード: {(_inputMode == InputMode.Scancode ? "Scancode" : "VirtualKey")}");
                        Console.WriteLine("MIDIノートオンイベントをキーボード押下に変換します。");
                        Console.WriteLine("任意のキーを押すと終了します...");

                        Console.ReadKey();
                    }
                    finally
                    {
                        // MIDI受信を停止して、デバイスを解放
                        _midiDevice.StopEventsListening();
                        _midiDevice.Dispose();
                    }
                }
                else
                {
                    Console.WriteLine("無効な選択です。");
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
                        SendScancodeKey(key, NativeMethods.KEYEVENTF_KEYDOWN);
                    else
                        SendKeyInput(key, NativeMethods.KEYEVENTF_KEYDOWN);
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
                            SendScancodeKey(key, NativeMethods.KEYEVENTF_KEYUP);
                        else
                            SendKeyInput(key, NativeMethods.KEYEVENTF_KEYUP);
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

        #endregion

        #region キーボードイベントに関する処理（仮想キーコード版）

        /// <summary>
        /// 指定された文字に対して仮想キー／修飾キーを用いてキーイベントを送信する
        /// VkKeyScan が失敗した場合は Unicode フォールバックを行う
        /// </summary>
        /// <param name="keyChar">送信対象の文字</param>
        /// <param name="keyEventFlag">KEYEVENTF_* のフラグ（KEYEVENTF_KEYDOWN / KEYEVENTF_KEYUP）</param>
        private static void SendKeyInput(char keyChar, uint keyEventFlag)
        {
            short vkWithState = NativeMethods.VkKeyScan(keyChar);
            if (vkWithState == -1)
            {
                // VkKeyScan でマップできない場合は Unicode 送信を試みる
                SendUnicodeKey(keyChar, keyEventFlag);
                return;
            }

            byte virtualKey = (byte)(vkWithState & 0xFF);
            byte shiftState = (byte)((vkWithState >> 8) & 0xFF);

            bool needShift = (shiftState & 1) != 0;
            bool needCtrl = (shiftState & 2) != 0;
            bool needAlt = (shiftState & 4) != 0;

            // キーダウン時は修飾を先に押し、キーアップ時は修飾を後で離す
            if (keyEventFlag == NativeMethods.KEYEVENTF_KEYDOWN)
            {
                if (needShift) SendSingleVk(NativeMethods.VK_SHIFT, NativeMethods.KEYEVENTF_KEYDOWN);
                if (needCtrl) SendSingleVk(NativeMethods.VK_CONTROL, NativeMethods.KEYEVENTF_KEYDOWN);
                if (needAlt) SendSingleVk(NativeMethods.VK_MENU, NativeMethods.KEYEVENTF_KEYDOWN);

                // 通常は仮想キーを送る
                SendSingleVk(virtualKey, NativeMethods.KEYEVENTF_KEYDOWN);
            }
            else // KEYEVENTF_KEYUP
            {
                // まず本体のキーを離す
                SendSingleVk(virtualKey, NativeMethods.KEYEVENTF_KEYUP);

                if (needAlt) SendSingleVk(NativeMethods.VK_MENU, NativeMethods.KEYEVENTF_KEYUP);
                if (needCtrl) SendSingleVk(NativeMethods.VK_CONTROL, NativeMethods.KEYEVENTF_KEYUP);
                if (needShift) SendSingleVk(NativeMethods.VK_SHIFT, NativeMethods.KEYEVENTF_KEYUP);
            }
        }

        /// <summary>
        /// 単一の仮想キーイベントを SendInput で送るヘルパ
        /// </summary>
        /// <param name="virtualKey">仮想キーコード</param>
        /// <param name="flags">送信フラグ</param>
        private static void SendSingleVk(ushort virtualKey, uint flags)
        {
            var input = new NativeMethods.INPUT
            {
                type = NativeMethods.INPUT_KEYBOARD,
                U = new NativeMethods.InputUnion
                {
                    ki = new NativeMethods.KEYBDINPUT
                    {
                        wVk = virtualKey,
                        wScan = 0,
                        dwFlags = flags,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            uint result = NativeMethods.SendInput(1, new NativeMethods.INPUT[] { input }, Marshal.SizeOf(typeof(NativeMethods.INPUT)));
            if (result == 0)
            {
                Console.WriteLine($"[エラー] SendInput 失敗 GetLastWin32Error: {Marshal.GetLastWin32Error()}");
            }
        }

        /// <summary>
        /// VkKeyScan が失敗した文字（Unicode）を送るためのフォールバック実装
        /// KEYEVENTF_UNICODE を用いて wScan に Unicode を入れて送信する
        /// </summary>
        /// <param name="keyChar">送る文字</param>
        /// <param name="keyEventFlag">KEYEVENTF_* フラグ</param>
        private static void SendUnicodeKey(char keyChar, uint keyEventFlag)
        {
            var input = new NativeMethods.INPUT
            {
                type = NativeMethods.INPUT_KEYBOARD,
                U = new NativeMethods.InputUnion
                {
                    ki = new NativeMethods.KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = (ushort)keyChar,
                        dwFlags = keyEventFlag | NativeMethods.KEYEVENTF_UNICODE,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            uint result = NativeMethods.SendInput(1, new NativeMethods.INPUT[] { input }, Marshal.SizeOf(typeof(NativeMethods.INPUT)));
            if (result == 0)
            {
                Console.WriteLine($"[エラー] Unicode SendInput 失敗 GetLastWin32Error: {Marshal.GetLastWin32Error()}");
            }
        }

        #endregion

        #region キーボードイベントに関する処理（スキャンコード版）

        /// <summary>
        /// 指定された文字に対してスキャンコード／修飾キーを用いてキーイベントを送信する
        /// VkKeyScan が失敗した場合は Unicode フォールバックを行う
        /// </summary>
        /// <param name="keyChar">送信対象の文字</param>
        /// <param name="keyEventFlag">KEYEVENTF_* のフラグ</param>
        private static void SendScancodeKey(char keyChar, uint keyEventFlag)
        {
            // VkKeyScan でマップできなければ Unicode フォールバック
            short vkWithState = NativeMethods.VkKeyScan(keyChar);
            if (vkWithState == -1)
            {
                SendUnicodeKey(keyChar, keyEventFlag);
                return;
            }

            byte virtualKey = (byte)(vkWithState & 0xFF);
            byte shiftState = (byte)((vkWithState >> 8) & 0xFF);

            bool needShift = (shiftState & 1) != 0;
            bool needCtrl = (shiftState & 2) != 0;
            bool needAlt = (shiftState & 4) != 0;

            // 修飾キーの scancode を取得（MapVirtualKey: MAPVK_VK_TO_VSC = 0）
            ushort scShift = (ushort)NativeMethods.MapVirtualKey(NativeMethods.VK_SHIFT, 0);
            ushort scCtrl = (ushort)NativeMethods.MapVirtualKey(NativeMethods.VK_CONTROL, 0);
            ushort scAlt = (ushort)NativeMethods.MapVirtualKey(NativeMethods.VK_MENU, 0);

            // 主キーの scancode
            ushort scanCode = (ushort)NativeMethods.MapVirtualKey(virtualKey, 0);
            bool extended = IsExtendedKeyForVk(virtualKey);

            if (keyEventFlag == NativeMethods.KEYEVENTF_KEYDOWN)
            {
                // 修飾を先に押す
                if (needShift) SendSingleScancode(scShift, false, NativeMethods.KEYEVENTF_SCANCODE);
                if (needCtrl) SendSingleScancode(scCtrl, false, NativeMethods.KEYEVENTF_SCANCODE);
                if (needAlt) SendSingleScancode(scAlt, false, NativeMethods.KEYEVENTF_SCANCODE);

                // 主キー押下
                SendSingleScancode(scanCode, extended, NativeMethods.KEYEVENTF_SCANCODE);
            }
            else // KEYEVENTF_KEYUP
            {
                // 主キー離上
                SendSingleScancode(scanCode, extended, NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);

                // 修飾を後で離す（逆順でも可）
                if (needAlt) SendSingleScancode(scAlt, false, NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);
                if (needCtrl) SendSingleScancode(scCtrl, false, NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);
                if (needShift) SendSingleScancode(scShift, false, NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);
            }
        }

        /// <summary>
        /// 単一の scancode イベントを SendInput で送るヘルパ
        /// </summary>
        /// <param name="scancode">送信するスキャンコード</param>
        /// <param name="extended">拡張キーフラグを付けるか</param>
        /// <param name="flags">送信フラグ</param>
        private static void SendSingleScancode(ushort scancode, bool extended, uint flags)
        {
            uint sendFlags = flags;
            if (extended) sendFlags |= NativeMethods.KEYEVENTF_EXTENDEDKEY;

            var input = new NativeMethods.INPUT
            {
                type = NativeMethods.INPUT_KEYBOARD,
                U = new NativeMethods.InputUnion
                {
                    ki = new NativeMethods.KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = scancode,
                        dwFlags = sendFlags,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            uint result = NativeMethods.SendInput(1, new NativeMethods.INPUT[] { input }, Marshal.SizeOf(typeof(NativeMethods.INPUT)));
            if (result == 0)
            {
                Console.WriteLine($"[エラー] Scancode SendInput 失敗 GetLastWin32Error: {Marshal.GetLastWin32Error()}");
            }
        }

        /// <summary>
        /// 指定した仮想キーが拡張キーに該当するか判定する
        /// </summary>
        /// <param name="virtualKey">仮想キーコード</param>
        /// <returns>拡張キーなら true</returns>
        private static bool IsExtendedKeyForVk(ushort virtualKey)
        {
            switch (virtualKey)
            {
                // これらは拡張キー扱い（テンキー以外の Insert/Delete/Home/End/PageUp/PageDown/矢印、右 Ctrl/右 Alt 等）
                case 0x2D: // VK_INSERT
                case 0x2E: // VK_DELETE
                case 0x24: // VK_HOME
                case 0x23: // VK_END
                case 0x21: // VK_PRIOR (PageUp)
                case 0x22: // VK_NEXT  (PageDown)
                case 0x27: // VK_RIGHT
                case 0x25: // VK_LEFT
                case 0x26: // VK_UP
                case 0x28: // VK_DOWN
                case 0xA3: // VK_RCONTROL
                case 0xA5: // VK_RMENU (Right Alt)
                case 0x6F: // VK_DIVIDE (numpad /) - often extended
                    return true;
                default:
                    return false;
            }
        }

        #endregion
    }
}
