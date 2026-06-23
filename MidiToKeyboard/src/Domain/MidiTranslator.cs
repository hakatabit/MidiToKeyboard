using System;
using System.Collections.Generic;

namespace MidiToKeyboard.Domain
{
    public sealed class MidiTranslator
    {
        private readonly IReadOnlyDictionary<int, char> _noteToKeyMap;
        private readonly KeyPressState _keyPressState;

        public MidiTranslator(
            IReadOnlyDictionary<int, char> noteToKeyMap,
            KeyPressState keyPressState)
        {
            if (noteToKeyMap == null)
                throw new ArgumentNullException(nameof(noteToKeyMap));

            if (keyPressState == null)
                throw new ArgumentNullException(nameof(keyPressState));

            _noteToKeyMap = noteToKeyMap;
            _keyPressState = keyPressState;
        }

        public IEnumerable<KeyAction> Translate(MidiEvent midiEvent)
        {
            if (midiEvent == null)
                throw new ArgumentNullException(nameof(midiEvent));

            char key;
            if (!_noteToKeyMap.TryGetValue(midiEvent.NoteNumber, out key))
                return new List<KeyAction>();

            if (midiEvent.Type == MidiEventType.NoteOn)
            {
                if (_keyPressState.ShouldSendKeyDown(midiEvent.NoteNumber, key))
                    return new List<KeyAction> { KeyAction.KeyDown(key) };
            }
            else if (midiEvent.Type == MidiEventType.NoteOff)
            {
                if (_keyPressState.ShouldSendKeyUp(midiEvent.NoteNumber, key))
                    return new List<KeyAction> { KeyAction.KeyUp(key) };
            }

            return new List<KeyAction>();
        }
    }
}
