using System;
using System.Collections.Generic;
using System.Linq;
using MidiToKeyboard.Domain;

namespace MidiToKeyboard.Application
{
    /// <summary>
    /// MIDI 入力、プロファイル、およびキー出力の接続
    /// </summary>
    public sealed class MidiToKeyboardApplication
    {
        private readonly object _syncRoot = new object();
        private readonly IMidiInput _midiInput;
        private readonly IKeyOutput _keyOutput;
        private readonly IProfileRepository _profileRepository;

        private Profile _currentProfile;
        private MidiTranslator _midiTranslator;
        private bool _isStarted;

        /// <summary>
        /// UI に表示可能な MIDI 入力活動が発生したときの通知
        /// </summary>
        public event Action<MidiInputActivity> MidiInputActivityOccurred;

        /// <summary>
        /// アプリケーションを構成する依存関係の初期化
        /// </summary>
        /// <param name="midiInput">MIDI 入力</param>
        /// <param name="keyOutput">キー出力</param>
        /// <param name="profileRepository">プロファイルのリポジトリ</param>
        public MidiToKeyboardApplication(
            IMidiInput midiInput,
            IKeyOutput keyOutput,
            IProfileRepository profileRepository)
        {
            if (midiInput == null)
            {
                throw new ArgumentNullException(
                    nameof(midiInput),
                    "MIDI 入力は null にできません。");
            }

            if (keyOutput == null)
            {
                throw new ArgumentNullException(
                    nameof(keyOutput),
                    "キー出力は null にできません。");
            }

            if (profileRepository == null)
            {
                throw new ArgumentNullException(
                    nameof(profileRepository),
                    "プロファイルのリポジトリは null にできません。");
            }

            _midiInput = midiInput;
            _keyOutput = keyOutput;
            _profileRepository = profileRepository;
        }

        /// <summary>
        /// VirtualKey モードで MIDI 入力の変換を開始
        /// </summary>
        /// <param name="deviceId">使用する MIDI 入力デバイスの識別値</param>
        /// <param name="profileName">使用するプロファイル名</param>
        public void Start(string deviceId, string profileName)
        {
            Start(deviceId, profileName, InputMode.VirtualKey);
        }

        /// <summary>
        /// 指定した入力送信モードで MIDI 入力の変換を開始
        /// </summary>
        /// <param name="deviceId">使用する MIDI 入力デバイスの識別値</param>
        /// <param name="profileName">使用するプロファイル名</param>
        /// <param name="inputMode">キー入力の送信方式</param>
        public void Start(string deviceId, string profileName, InputMode inputMode)
        {
            Stop();

            IInputModeKeyOutput inputModeKeyOutput = _keyOutput as IInputModeKeyOutput;
            if (inputModeKeyOutput != null)
            {
                inputModeKeyOutput.SetInputMode(inputMode);
            }
            else if (inputMode != InputMode.VirtualKey)
            {
                throw new NotSupportedException(
                    "構成されたキー出力は入力送信モードの選択をサポートしていません。");
            }

            SetProfile(profileName);

            lock (_syncRoot)
            {
                _midiInput.MessageReceived += OnMidiMessageReceived;
                try
                {
                    _midiInput.Start(deviceId);
                    _isStarted = true;
                }
                catch
                {
                    _midiInput.MessageReceived -= OnMidiMessageReceived;
                    throw;
                }
            }
        }

        /// <summary>
        /// MIDI 入力の変換を停止し、実行中の状態を破棄
        /// </summary>
        public void Stop()
        {
            lock (_syncRoot)
            {
                if (_isStarted)
                {
                    _midiInput.MessageReceived -= OnMidiMessageReceived;
                    _midiInput.Stop();
                    _isStarted = false;
                }

                _currentProfile = null;
                _midiTranslator = null;
            }
        }

        /// <summary>
        /// MIDI 入力の変換に使用するプロファイルを設定
        /// </summary>
        /// <param name="profileName">使用するプロファイル名</param>
        public void SetProfile(string profileName)
        {
            Profile profile = _profileRepository.Load(profileName);
            if (profile == null)
            {
                throw new InvalidOperationException(
                    "プロファイルリポジトリが null を返しました。");
            }

            KeyPressState keyPressState = new KeyPressState();
            MidiTranslator midiTranslator =
                new MidiTranslator(profile.NoteMappings, keyPressState);

            lock (_syncRoot)
            {
                _currentProfile = profile;
                _midiTranslator = midiTranslator;
            }
        }

        private void OnMidiMessageReceived(MidiEvent midiEvent)
        {
            List<KeyAction> actions;
            MidiInputActivity activity;

            lock (_syncRoot)
            {
                if (_midiTranslator == null)
                {
                    return;
                }

                actions = _midiTranslator.Translate(midiEvent).ToList();
                activity = CreateMidiInputActivity(midiEvent);
            }

            foreach (KeyAction action in actions)
            {
                _keyOutput.Send(action);
            }

            Action<MidiInputActivity> activityHandler = MidiInputActivityOccurred;
            if (activity != null && activityHandler != null)
            {
                activityHandler(activity);
            }
        }

        private MidiInputActivity CreateMidiInputActivity(
            MidiEvent midiEvent)
        {
            MidiInputActivityType activityType;

            if (midiEvent.Type == MidiEventType.NoteOn)
            {
                activityType = MidiInputActivityType.NoteOn;
            }
            else if (midiEvent.Type == MidiEventType.NoteOff)
            {
                activityType = MidiInputActivityType.NoteOff;
            }
            else
            {
                return null;
            }

            char keyChar;
            if (!_currentProfile.NoteMappings.TryGetValue(midiEvent.NoteNumber, out keyChar))
            {
                keyChar = '\0';
            }

            return new MidiInputActivity(
                activityType,
                midiEvent.NoteNumber,
                keyChar);
        }
    }
}
