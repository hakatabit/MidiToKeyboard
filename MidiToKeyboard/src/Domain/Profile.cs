using System;
using System.Collections.Generic;

namespace MidiToKeyboard.Domain
{
    /// <summary>
    /// 名前付きの MIDI ノートとキー文字の割り当て
    /// </summary>
    public sealed class Profile
    {
        private readonly Dictionary<int, char> _noteMappings;

        /// <summary>
        /// プロファイル名
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// MIDI ノート番号とキー文字の割り当て
        /// </summary>
        public IReadOnlyDictionary<int, char> NoteMappings
        {
            get { return _noteMappings; }
        }

        /// <summary>
        /// プロファイル名とノート割り当ての初期化
        /// </summary>
        /// <param name="name">プロファイル名</param>
        /// <param name="noteMappings">MIDI ノート番号とキー文字の割り当て</param>
        public Profile(string name, IDictionary<int, char> noteMappings)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException(
                    "プロファイル名に null、空文字列、または空白のみの文字列は指定できません。",
                    nameof(name));
            }

            if (noteMappings == null)
            {
                throw new ArgumentNullException(
                    nameof(noteMappings),
                    "ノート割り当ては null にできません。");
            }

            Name = name;
            _noteMappings = new Dictionary<int, char>(noteMappings);
        }
    }
}
