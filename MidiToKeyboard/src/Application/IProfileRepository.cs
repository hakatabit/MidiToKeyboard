using System.Collections.Generic;

namespace MidiToKeyboard.Application
{
    /// <summary>
    /// プロファイルの列挙と読み込みを行う機能の定義
    /// </summary>
    public interface IProfileRepository
    {
        /// <summary>
        /// 指定された名前のプロファイルを読み込む
        /// </summary>
        /// <param name="name">読み込むプロファイル名</param>
        /// <returns>読み込んだプロファイル</returns>
        MidiToKeyboard.Domain.Profile Load(string name);

        /// <summary>
        /// 利用可能なプロファイル名の取得
        /// </summary>
        /// <returns>利用可能なプロファイル名</returns>
        IEnumerable<string> ListNames();
    }
}
