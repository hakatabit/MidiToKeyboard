namespace MidiToKeyboard.Application
{
    /// <summary>
    /// UI に通知する MIDI 入力活動の種類
    /// </summary>
    public enum MidiInputActivityType
    {
        /// <summary>
        /// ノートの押下
        /// </summary>
        NoteOn,

        /// <summary>
        /// ノートの解放
        /// </summary>
        NoteOff
    }

    /// <summary>
    /// UI に表示する MIDI 入力活動
    /// </summary>
    public sealed class MidiInputActivity
    {
        /// <summary>
        /// MIDI 入力活動の種類
        /// </summary>
        public MidiInputActivityType Type { get; }

        /// <summary>
        /// MIDI ノート番号
        /// </summary>
        public int NoteNumber { get; }

        /// <summary>
        /// ノートに割り当てられたキー文字
        /// </summary>
        public char KeyChar { get; }

        /// <summary>
        /// MIDI 入力活動の初期化
        /// </summary>
        /// <param name="type">MIDI 入力活動の種類</param>
        /// <param name="noteNumber">MIDI ノート番号</param>
        /// <param name="keyChar">ノートに割り当てられたキー文字</param>
        public MidiInputActivity(
            MidiInputActivityType type,
            int noteNumber,
            char keyChar)
        {
            Type = type;
            NoteNumber = noteNumber;
            KeyChar = keyChar;
        }
    }
}
