using System;
using System.Collections.Generic;

namespace MidiToKeyboard.Domain
{
    /// <summary>
    /// MIDI イベントをプロファイルに従ってキー操作へ変換
    /// </summary>
    public sealed class MidiTranslator
    {
        private readonly IReadOnlyDictionary<int, char> _noteToKeyMap;
        private readonly KeyPressState _keyPressState;

        /// <summary>
        /// ノート割り当てと押下状態の初期化
        /// </summary>
        /// <param name="noteToKeyMap">MIDI ノート番号とキー文字の割り当て</param>
        /// <param name="keyPressState">ノートとキーの押下状態</param>
        public MidiTranslator(
            IReadOnlyDictionary<int, char> noteToKeyMap,
            KeyPressState keyPressState)
        {
            if (noteToKeyMap == null)
            {
                throw new ArgumentNullException(
                    nameof(noteToKeyMap),
                    "ノート割り当ては null にできません。");
            }

            if (keyPressState == null)
            {
                throw new ArgumentNullException(
                    nameof(keyPressState),
                    "キー押下状態は null にできません。");
            }

            _noteToKeyMap = noteToKeyMap;
            _keyPressState = keyPressState;
        }

        /// <summary>
        /// MIDI イベントを送信すべきキー操作へ変換
        /// </summary>
        /// <param name="midiEvent">変換する MIDI イベント</param>
        /// <returns>送信すべきキー操作</returns>
        public IEnumerable<KeyAction> Translate(MidiEvent midiEvent)
        {
            if (midiEvent == null)
            {
                throw new ArgumentNullException(
                    nameof(midiEvent),
                    "MIDI イベントは null にできません。");
            }

            char key;
            if (!_noteToKeyMap.TryGetValue(midiEvent.NoteNumber, out key))
            {
                return new List<KeyAction>();
            }

            if (midiEvent.Type == MidiEventType.NoteOn)
            {
                if (_keyPressState.ShouldSendKeyDown(midiEvent.NoteNumber, key))
                {
                    return new List<KeyAction> { KeyAction.KeyDown(key) };
                }
            }
            else if (midiEvent.Type == MidiEventType.NoteOff)
            {
                if (_keyPressState.ShouldSendKeyUp(midiEvent.NoteNumber, key))
                {
                    return new List<KeyAction> { KeyAction.KeyUp(key) };
                }
            }

            return new List<KeyAction>();
        }
    }
}
