using System;
using System.Collections.Generic;

namespace MidiToKeyboard.Application
{
    /// <summary>
    /// MIDI 入力デバイスの列挙と受信を行う機能の定義
    /// </summary>
    public interface IMidiInput
    {
        /// <summary>
        /// MIDI メッセージを受信したときに発生するイベント
        /// </summary>
        event Action<MidiToKeyboard.Domain.MidiEvent> MessageReceived;

        /// <summary>
        /// 利用可能な MIDI 入力デバイスの取得
        /// </summary>
        /// <returns>利用可能な MIDI 入力デバイス</returns>
        IReadOnlyList<MidiDeviceInfo> EnumerateDevices();

        /// <summary>
        /// 指定された MIDI 入力デバイスからの受信を開始
        /// </summary>
        /// <param name="deviceId">受信する MIDI 入力デバイスの ID</param>
        void Start(string deviceId);

        /// <summary>
        /// MIDI 入力の受信を停止
        /// </summary>
        void Stop();
    }
}
