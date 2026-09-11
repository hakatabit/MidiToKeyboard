namespace MidiToKeyboard.Application
{
    /// <summary>
    /// Windows へキー入力を送信する方式
    /// </summary>
    public enum InputMode
    {
        /// <summary>
        /// 仮想キーコードを使用する方式
        /// </summary>
        VirtualKey,

        /// <summary>
        /// スキャンコードを使用する方式
        /// </summary>
        Scancode
    }
}
