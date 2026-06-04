namespace MidiToKeyboard.Domain
{
    public enum MidiEventType
    {
        NoteOn,
        NoteOff,
        ControlChange
    }

    public sealed class MidiEvent
    {
        public MidiEventType Type { get; }

        public int Channel { get; }
        public int NoteNumber { get; }
        public int Velocity { get; }
        public int ControlNumber { get; }
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
