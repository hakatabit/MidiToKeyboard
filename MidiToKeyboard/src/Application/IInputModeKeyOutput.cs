namespace MidiToKeyboard.Application
{
    /// <summary>
    /// キー出力へ入力送信モードを設定する機能の定義
    /// </summary>
    public interface IInputModeKeyOutput
    {
        /// <summary>
        /// 入力送信モードの設定
        /// </summary>
        /// <param name="inputMode">設定する入力送信モード</param>
        void SetInputMode(InputMode inputMode);
    }
}
