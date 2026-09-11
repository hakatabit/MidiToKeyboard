namespace MidiToKeyboard.Domain
{
    /// <summary>
    /// アプリケーションが扱う MIDI イベントの種類
    /// </summary>
    public enum MidiEventType
    {
        NoteOn,
        NoteOff,
        ControlChange
    }

    /// <summary>
    /// 外部ライブラリに依存しない MIDI イベント
    /// </summary>
    public sealed class MidiEvent
    {
        /// <summary>
        /// MIDI イベントの種類
        /// </summary>
        public MidiEventType Type { get; }

        /// <summary>
        /// MIDI チャンネル
        /// </summary>
        public int Channel { get; }

        /// <summary>
        /// MIDI ノート番号
        /// </summary>
        public int NoteNumber { get; }

        /// <summary>
        /// ベロシティ
        /// </summary>
        public int Velocity { get; }

        /// <summary>
        /// コントロール番号
        /// </summary>
        public int ControlNumber { get; }

        /// <summary>
        /// コントロール値
        /// </summary>
        public int ControlValue { get; }

        private MidiEvent(
            MidiEventType type,
            int channel,
            int noteNumber,
            int velocity,
            int controlNumber,
            int controlValue)
        {
            Type = type;
            Channel = channel;
            NoteNumber = noteNumber;
            Velocity = velocity;
            ControlNumber = controlNumber;
            ControlValue = controlValue;
        }

        /// <summary>
        /// ノートオンイベントを作成
        /// </summary>
        /// <param name="channel">MIDI チャンネル</param>
        /// <param name="noteNumber">MIDI ノート番号</param>
        /// <param name="velocity">ベロシティ</param>
        /// <returns>作成した MIDI イベント</returns>
        public static MidiEvent NoteOn(int channel, int noteNumber, int velocity)
        {
            return new MidiEvent(
                MidiEventType.NoteOn,
                channel,
                noteNumber,
                velocity,
                0,
                0);
        }

        /// <summary>
        /// ノートオフイベントを作成
        /// </summary>
        /// <param name="channel">MIDI チャンネル</param>
        /// <param name="noteNumber">MIDI ノート番号</param>
        /// <returns>作成した MIDI イベント</returns>
        public static MidiEvent NoteOff(int channel, int noteNumber)
        {
            return new MidiEvent(
                MidiEventType.NoteOff,
                channel,
                noteNumber,
                0,
                0,
                0);
        }

        /// <summary>
        /// コントロールチェンジイベントを作成
        /// </summary>
        /// <param name="channel">MIDI チャンネル</param>
        /// <param name="controlNumber">コントロール番号</param>
        /// <param name="controlValue">コントロール値</param>
        /// <returns>作成した MIDI イベント</returns>
        public static MidiEvent ControlChange(int channel, int controlNumber, int controlValue)
        {
            return new MidiEvent(
                MidiEventType.ControlChange,
                channel,
                0,
                0,
                controlNumber,
                controlValue);
        }
    }
}
