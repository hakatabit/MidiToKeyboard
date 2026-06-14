using System.Collections.Generic;

namespace MidiToKeyboard.Domain
{
    public sealed class KeyPressState
    {
        private readonly object _stateLock = new object();
        private readonly HashSet<int> _activeNotes = new HashSet<int>();
        private readonly Dictionary<char, int> _keyRefCount = new Dictionary<char, int>();

        public bool ShouldSendKeyDown(int noteNumber, char key)
        {
            if (key == '\0')
                return false;

            lock (_stateLock)
            {
                if (!_activeNotes.Add(noteNumber))
                    return false;

                int count;
                if (_keyRefCount.TryGetValue(key, out count))
                    _keyRefCount[key] = count + 1;
                else
                    _keyRefCount[key] = 1;

                return _keyRefCount[key] == 1;
            }
        }

        public bool ShouldSendKeyUp(int noteNumber, char key)
        {
            if (key == '\0')
                return false;

            lock (_stateLock)
            {
                if (!_activeNotes.Remove(noteNumber))
                    return false;

                int count;
                if (!_keyRefCount.TryGetValue(key, out count))
                    return false;

                count--;
                if (count <= 0)
                {
                    _keyRefCount.Remove(key);
                    return true;
                }

                _keyRefCount[key] = count;
                return false;
            }
        }
    }
}
