using System.Collections.Generic;

namespace MidiToKeyboard.Domain
{
    /// <summary>
    /// ノートとキーの押下状態を保持し、重複するキー操作を抑制
    /// </summary>
    public sealed class KeyPressState
    {
        private readonly object _stateLock = new object();
        private readonly HashSet<int> _activeNotes = new HashSet<int>();
        private readonly Dictionary<char, int> _keyRefCount = new Dictionary<char, int>();

        /// <summary>
        /// 指定したノートに対してキー押下を送信すべきかを判定
        /// </summary>
        /// <param name="noteNumber">MIDI ノート番号</param>
        /// <param name="key">ノートに割り当てられたキー文字</param>
        /// <returns>キー押下を送信する場合は <see langword="true"/></returns>
        public bool ShouldSendKeyDown(int noteNumber, char key)
        {
            if (key == '\0')
            {
                return false;
            }

            lock (_stateLock)
            {
                if (!_activeNotes.Add(noteNumber))
                {
                    return false;
                }

                int count;
                if (_keyRefCount.TryGetValue(key, out count))
                {
                    _keyRefCount[key] = count + 1;
                }
                else
                {
                    _keyRefCount[key] = 1;
                }

                return _keyRefCount[key] == 1;
            }
        }

        /// <summary>
        /// 指定したノートに対してキー解放を送信すべきかを判定
        /// </summary>
        /// <param name="noteNumber">MIDI ノート番号</param>
        /// <param name="key">ノートに割り当てられたキー文字</param>
        /// <returns>キー解放を送信する場合は <see langword="true"/></returns>
        public bool ShouldSendKeyUp(int noteNumber, char key)
        {
            if (key == '\0')
            {
                return false;
            }

            lock (_stateLock)
            {
                if (!_activeNotes.Remove(noteNumber))
                {
                    return false;
                }

                int count;
                if (!_keyRefCount.TryGetValue(key, out count))
                {
                    return false;
                }

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
