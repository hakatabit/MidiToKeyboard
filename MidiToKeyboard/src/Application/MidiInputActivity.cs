namespace MidiToKeyboard.Application
{
    public enum MidiInputActivityType
    {
        NoteOn,
        NoteOff
    }

    public sealed class MidiInputActivity
    {
        public MidiInputActivityType Type { get; }
        public int NoteNumber { get; }
        public char KeyChar { get; }

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
