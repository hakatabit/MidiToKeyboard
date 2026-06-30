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

        private global::MidiToKeyboard.Domain.Profile _currentProfile;
        private global::MidiToKeyboard.Domain.KeyPressState _keyPressState;
        private global::MidiToKeyboard.Domain.MidiTranslator _midiTranslator;
        private bool _isStarted;

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
            Stop();

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
                _keyPressState = null;
                _midiTranslator = null;
            }
        }

        public void SetProfile(string profileName)
        {
            global::MidiToKeyboard.Domain.Profile profile = _profileRepository.Load(profileName);
            if (profile == null)
                throw new InvalidOperationException("Profile repository returned null.");

            global::MidiToKeyboard.Domain.KeyPressState keyPressState =
                new global::MidiToKeyboard.Domain.KeyPressState();
            global::MidiToKeyboard.Domain.MidiTranslator midiTranslator =
                new global::MidiToKeyboard.Domain.MidiTranslator(profile.NoteMappings, keyPressState);

            lock (_syncRoot)
            {
                _currentProfile = profile;
                _keyPressState = keyPressState;
                _midiTranslator = midiTranslator;
            }
        }

        private void OnMidiMessageReceived(global::MidiToKeyboard.Domain.MidiEvent midiEvent)
        {
            List<global::MidiToKeyboard.Domain.KeyAction> actions;

            lock (_syncRoot)
            {
                if (_midiTranslator == null)
                    return;

                actions = _midiTranslator.Translate(midiEvent).ToList();
            }

            foreach (global::MidiToKeyboard.Domain.KeyAction action in actions)
            {
                _keyOutput.Send(action);
            }
        }
    }
}
