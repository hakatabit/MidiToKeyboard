namespace MidiToKeyboard.Domain
{
    public enum KeyActionType
    {
        KeyDown,
        KeyUp,
        KeyPress,
        UnicodeText
    }

    public sealed class KeyAction
    {
        public KeyActionType Type { get; }
        public char KeyChar { get; }
        public string Text { get; }

        private KeyAction(
            KeyActionType type,
            char keyChar,
            string text)
        {
            Type = type;
            KeyChar = keyChar;
            Text = text;
        }

        public static KeyAction KeyDown(char keyChar)
        {
            return new KeyAction(KeyActionType.KeyDown, keyChar, null);
        }

        public static KeyAction KeyUp(char keyChar)
        {
            return new KeyAction(KeyActionType.KeyUp, keyChar, null);
        }

        public static KeyAction KeyPress(char keyChar)
        {
            return new KeyAction(KeyActionType.KeyPress, keyChar, null);
        }

        public static KeyAction UnicodeText(string text)
        {
            return new KeyAction(KeyActionType.UnicodeText, '\0', text);
        }

    }
}
