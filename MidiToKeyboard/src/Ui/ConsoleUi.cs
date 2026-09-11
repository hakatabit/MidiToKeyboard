using System;
using System.Collections.Generic;
using System.Linq;
using MidiToKeyboard.Application;

namespace MidiToKeyboard.Ui
{
    /// <summary>
    /// コンソールを介した選択操作と実行中の通知表示を担当
    /// </summary>
    public sealed class ConsoleUi
    {
        private readonly MidiToKeyboardApplication _application;
        private readonly IMidiInput _midiInput;
        private readonly IProfileRepository _profileRepository;

        /// <summary>
        /// UI が利用するアプリケーションと一覧取得元の初期化
        /// </summary>
        /// <param name="application">MIDI 変換処理を制御するアプリケーション</param>
        /// <param name="midiInput">選択可能な MIDI デバイスの取得元</param>
        /// <param name="profileRepository">選択可能なプロファイルの取得元</param>
        public ConsoleUi(
            MidiToKeyboardApplication application,
            IMidiInput midiInput,
            IProfileRepository profileRepository)
        {
            if (application == null)
            {
                throw new ArgumentNullException(
                    nameof(application),
                    "アプリケーションは null にできません。");
            }

            if (midiInput == null)
            {
                throw new ArgumentNullException(
                    nameof(midiInput),
                    "MIDI 入力は null にできません。");
            }

            if (profileRepository == null)
            {
                throw new ArgumentNullException(
                    nameof(profileRepository),
                    "プロファイルリポジトリは null にできません。");
            }

            _application = application;
            _midiInput = midiInput;
            _profileRepository = profileRepository;
        }

        /// <summary>
        /// 対話形式の操作フローを実行し、終了時にアプリケーションを停止
        /// </summary>
        public void Run()
        {
            try
            {
                RunCore();
            }
            finally
            {
                _application.Stop();
            }
        }

        private void RunCore()
        {
            IReadOnlyList<MidiDeviceInfo> devices = _midiInput.EnumerateDevices();

            Console.WriteLine("利用可能なMIDI入力デバイス:");
            if (devices.Count == 0)
            {
                Console.WriteLine("デバイスが見つかりませんでした。");
                return;
            }

            for (int i = 0; i < devices.Count; i++)
            {
                Console.WriteLine("  [{0}]: {1}", i, devices[i].Name);
            }

            Console.Write("使用するデバイスの番号を入力してください: ");
            int deviceIndex;
            if (!int.TryParse(Console.ReadLine(), out deviceIndex) ||
                deviceIndex < 0 ||
                deviceIndex >= devices.Count)
            {
                Console.WriteLine("無効な選択です。");
                return;
            }

            Console.WriteLine();
            Console.WriteLine("入力送信モードを選択してください:");
            Console.WriteLine("  [1] 仮想キーコード (VK) - 仮想キー＋修飾キーで送信");
            Console.WriteLine("  [2] スキャンコード (SC) - スキャンコードで送信（ゲーム等向け）");
            Console.Write("番号を入力: ");
            string inputModeSelection = Console.ReadLine();
            InputMode inputMode = inputModeSelection != null && inputModeSelection.Trim() == "2"
                ? InputMode.Scancode
                : InputMode.VirtualKey;

            List<string> profileNames = _profileRepository.ListNames().ToList();

            Console.WriteLine();
            Console.WriteLine("使用するマッピングセットを選択してください:");
            if (profileNames.Count == 0)
            {
                Console.WriteLine("マッピングファイルにセットがありません。");
                return;
            }

            for (int i = 0; i < profileNames.Count; i++)
            {
                Console.WriteLine("  [{0}] {1}", i, profileNames[i]);
            }

            Console.Write("番号を入力: ");
            int profileIndex;
            if (!int.TryParse(Console.ReadLine(), out profileIndex) ||
                profileIndex < 0 ||
                profileIndex >= profileNames.Count)
            {
                profileIndex = 0;
            }

            Console.WriteLine("選択セット: {0}", profileNames[profileIndex]);

            _application.Start(
                devices[deviceIndex].Id,
                profileNames[profileIndex],
                inputMode);

            Console.WriteLine();
            Console.WriteLine("** '{0}' をリッスン中です。 **", devices[deviceIndex].Name);
            Console.WriteLine("選択モード: {0}",
                inputMode == InputMode.Scancode ? "Scancode" : "VirtualKey");
            Console.WriteLine("MIDIノートオンイベントをキーボード押下に変換します。");
            Console.WriteLine("任意のキーを押すと終了します...");
            Console.ReadKey();
        }

        /// <summary>
        /// MIDI 入力活動をコンソールに表示
        /// </summary>
        /// <param name="activity">表示する MIDI 入力活動</param>
        public void HandleMidiInputActivity(MidiInputActivity activity)
        {
            if (activity.Type == MidiInputActivityType.NoteOn)
            {
                Console.WriteLine(
                    "[MIDI ON] Note: {0} -> キー: '{1}' DOWN",
                    activity.NoteNumber,
                    activity.KeyChar);
            }
            else
            {
                Console.WriteLine(
                    "[MIDI OFF] Note: {0} -> キー: '{1}' UP",
                    activity.NoteNumber,
                    activity.KeyChar);
            }
        }

        /// <summary>
        /// キー出力から通知された警告をコンソールに表示
        /// </summary>
        /// <param name="sender">警告の通知元</param>
        /// <param name="message">表示する警告メッセージ</param>
        public void HandleWarning(object sender, string message)
        {
            Console.WriteLine(message);
        }
    }
}
