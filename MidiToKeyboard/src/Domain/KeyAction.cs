namespace MidiToKeyboard.Domain
{
    /// <summary>
    /// 実行するキー操作の種類
    /// </summary>
    public enum KeyActionType
    {
        KeyDown,
        KeyUp,
        KeyPress,
        UnicodeText
    }

    /// <summary>
    /// MIDI 入力から変換されたキー操作
    /// </summary>
    public sealed class KeyAction
    {
        /// <summary>
        /// キー操作の種類
        /// </summary>
        public KeyActionType Type { get; }

        /// <summary>
        /// 操作対象のキー文字
        /// </summary>
        public char KeyChar { get; }

        /// <summary>
        /// Unicode 入力の対象文字列
        /// </summary>
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

        /// <summary>
        /// キーを押下する操作を作成
        /// </summary>
        /// <param name="keyChar">操作対象の文字</param>
        /// <returns>作成したキー操作</returns>
        public static KeyAction KeyDown(char keyChar)
        {
            return new KeyAction(KeyActionType.KeyDown, keyChar, null);
        }

        /// <summary>
        /// キーを解放する操作を作成
        /// </summary>
        /// <param name="keyChar">操作対象の文字</param>
        /// <returns>作成したキー操作</returns>
        public static KeyAction KeyUp(char keyChar)
        {
            return new KeyAction(KeyActionType.KeyUp, keyChar, null);
        }

        /// <summary>
        /// キーを押下して解放する操作を作成
        /// </summary>
        /// <param name="keyChar">操作対象の文字</param>
        /// <returns>作成したキー操作</returns>
        public static KeyAction KeyPress(char keyChar)
        {
            return new KeyAction(KeyActionType.KeyPress, keyChar, null);
        }

        /// <summary>
        /// Unicode 文字列を入力する操作を作成
        /// </summary>
        /// <param name="text">入力する文字列</param>
        /// <returns>作成したキー操作</returns>
        public static KeyAction UnicodeText(string text)
        {
            return new KeyAction(KeyActionType.UnicodeText, '\0', text);
        }
    }
}
