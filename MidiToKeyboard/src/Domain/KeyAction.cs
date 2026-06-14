namespace MidiToKeyboard.Domain
{
    public enum KeyActionType
    {
        KeyDown,
        KeyUp,
        KeyPress,
        UnicodeText,
        ScanCode,
        VirtualKey
    }

    public sealed class KeyAction
    {
        public KeyActionType Type { get; }
        public char KeyChar { get; }
        public int VirtualKey { get; }
        public int ScanCode { get; }
        public string Text { get; }

        private KeyAction(
            KeyActionType type,
            char keyChar,
            int virtualKey,
            int scanCode,
            string text)
        {
            Type = type;
            KeyChar = keyChar;
            VirtualKey = virtualKey;
            ScanCode = scanCode;
            Text = text;
        }

        public static KeyAction KeyDown(char keyChar)
        {
            return new KeyAction(KeyActionType.KeyDown, keyChar, 0, 0, null);
        }

        public static KeyAction KeyUp(char keyChar)
        {
            return new KeyAction(KeyActionType.KeyUp, keyChar, 0, 0, null);
        }

        public static KeyAction KeyPress(char keyChar)
        {
            return new KeyAction(KeyActionType.KeyPress, keyChar, 0, 0, null);
        }

        public static KeyAction UnicodeText(string text)
        {
            return new KeyAction(KeyActionType.UnicodeText, '\0', 0, 0, text);
        }

    }
}
