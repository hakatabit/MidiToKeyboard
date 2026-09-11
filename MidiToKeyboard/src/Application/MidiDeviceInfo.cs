namespace MidiToKeyboard.Application
{
    /// <summary>
    /// 選択可能な MIDI 入力デバイス
    /// </summary>
    public sealed class MidiDeviceInfo
    {
        /// <summary>
        /// デバイスを識別する値
        /// </summary>
        public string Id { get; }

        /// <summary>
        /// ユーザーに表示するデバイス名
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// MIDI 入力デバイス情報の初期化
        /// </summary>
        /// <param name="id">デバイスを識別する値</param>
        /// <param name="name">ユーザーに表示するデバイス名</param>
        public MidiDeviceInfo(string id, string name)
        {
            Id = id;
            Name = name;
        }
    }
}
