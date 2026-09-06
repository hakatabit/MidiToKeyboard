namespace MidiToKeyboard.Application
{
    /// <summary>
    /// キー操作を出力する機能の定義
    /// </summary>
    public interface IKeyOutput
    {
        /// <summary>
        /// 指定されたキー操作の出力
        /// </summary>
        /// <param name="action">出力するキー操作</param>
        void Send(MidiToKeyboard.Domain.KeyAction action);
    }
}
