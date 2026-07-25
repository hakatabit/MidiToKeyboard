using System;
using System.Collections.Generic;
using System.Linq;

namespace MidiToKeyboard.Application
{
    public sealed class MidiToKeyboardApplication
    {
        private readonly object _syncRoot = new object();
        private readonly IMidiInput _midiInput;
        private readonly IKeyOutput _keyOutput;
        private readonly IProfileRepository _profileRepository;

        private MidiToKeyboard.Domain.Profile _currentProfile;
        private MidiToKeyboard.Domain.MidiTranslator _midiTranslator;
        private bool _isStarted;

        public event Action<MidiInputActivity> MidiInputActivityOccurred;

        public MidiToKeyboardApplication(
            IMidiInput midiInput,
            IKeyOutput keyOutput,
            IProfileRepository profileRepository)
        {
            if (midiInput == null)
                throw new ArgumentNullException(nameof(midiInput));

            if (keyOutput == null)
                throw new ArgumentNullException(nameof(keyOutput));

            if (profileRepository == null)
                throw new ArgumentNullException(nameof(profileRepository));

            _midiInput = midiInput;
            _keyOutput = keyOutput;
            _profileRepository = profileRepository;
        }

        public void Start(string deviceId, string profileName)
        {
            Start(deviceId, profileName, InputMode.VirtualKey);
        }

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
                    "The configured key output does not support input mode selection.");
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

        public void SetProfile(string profileName)
        {
            MidiToKeyboard.Domain.Profile profile = _profileRepository.Load(profileName);
            if (profile == null)
                throw new InvalidOperationException("Profile repository returned null.");

            MidiToKeyboard.Domain.KeyPressState keyPressState =
                new MidiToKeyboard.Domain.KeyPressState();
            MidiToKeyboard.Domain.MidiTranslator midiTranslator =
                new MidiToKeyboard.Domain.MidiTranslator(profile.NoteMappings, keyPressState);

            lock (_syncRoot)
            {
                _currentProfile = profile;
                _midiTranslator = midiTranslator;
            }
        }

        private void OnMidiMessageReceived(MidiToKeyboard.Domain.MidiEvent midiEvent)
        {
            List<MidiToKeyboard.Domain.KeyAction> actions;
            MidiInputActivity activity;

            lock (_syncRoot)
            {
                if (_midiTranslator == null)
                    return;

                actions = _midiTranslator.Translate(midiEvent).ToList();
                activity = CreateMidiInputActivity(midiEvent);
            }

            foreach (MidiToKeyboard.Domain.KeyAction action in actions)
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
            MidiToKeyboard.Domain.MidiEvent midiEvent)
        {
            MidiInputActivityType activityType;

            if (midiEvent.Type == MidiToKeyboard.Domain.MidiEventType.NoteOn)
            {
                activityType = MidiInputActivityType.NoteOn;
            }
            else if (midiEvent.Type == MidiToKeyboard.Domain.MidiEventType.NoteOff)
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
