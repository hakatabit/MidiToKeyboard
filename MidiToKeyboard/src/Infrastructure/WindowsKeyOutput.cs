using MidiToKeyboard.Application;
using MidiToKeyboard.Domain;
using System;
using System.Runtime.InteropServices;

namespace MidiToKeyboard.Infrastructure
{
    public sealed class WindowsKeyOutput : IKeyOutput
    {
        public event EventHandler<string> WarningOccurred;

        public void Send(KeyAction action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(nameof(action));
            }

            switch (action.Type)
            {
                case KeyActionType.KeyDown:
                    SendKeyInput(action.KeyChar, NativeMethods.KEYEVENTF_KEYDOWN);
                    return;

                case KeyActionType.KeyUp:
                    SendKeyInput(action.KeyChar, NativeMethods.KEYEVENTF_KEYUP);
                    return;

                case KeyActionType.KeyPress:
                    SendKeyInput(action.KeyChar, NativeMethods.KEYEVENTF_KEYDOWN);
                    SendKeyInput(action.KeyChar, NativeMethods.KEYEVENTF_KEYUP);
                    return;

                case KeyActionType.UnicodeText:
                    SendUnicodeText(action.Text);
                    return;

                case KeyActionType.ScanCode:
                case KeyActionType.VirtualKey:
                    throw new NotSupportedException("Unsupported KeyActionType: " + action.Type);

                default:
                    throw new NotSupportedException("Unsupported KeyActionType: " + action.Type);
            }
        }

        private void OnWarningOccurred(string message)
        {
            EventHandler<string> handler = WarningOccurred;
            if (handler != null)
            {
                handler(this, message);
            }
        }

        private void SendUnicodeText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            foreach (char keyChar in text)
            {
                SendUnicodeKey(keyChar, NativeMethods.KEYEVENTF_KEYDOWN);
                SendUnicodeKey(keyChar, NativeMethods.KEYEVENTF_KEYUP);
            }
        }

        /// <summary>
        /// 指定された文字に対して仮想キー／修飾キーを用いてキーイベントを送信する
        /// VkKeyScan が失敗した場合は Unicode フォールバックを行う
        /// </summary>
        /// <param name="keyChar">送信対象の文字</param>
        /// <param name="keyEventFlag">KEYEVENTF_* のフラグ（KEYEVENTF_KEYDOWN / KEYEVENTF_KEYUP）</param>
        public void SendKeyInput(char keyChar, uint keyEventFlag)
        {
            short vkWithState = NativeMethods.VkKeyScan(keyChar);
            if (vkWithState == -1)
            {
                // VkKeyScan でマップできない場合は Unicode 送信を試みる
                SendUnicodeKey(keyChar, keyEventFlag);
                return;
            }

            byte virtualKey = (byte)(vkWithState & 0xFF);
            byte shiftState = (byte)((vkWithState >> 8) & 0xFF);

            bool needShift = (shiftState & 1) != 0;
            bool needCtrl = (shiftState & 2) != 0;
            bool needAlt = (shiftState & 4) != 0;

            // キーダウン時は修飾を先に押し、キーアップ時は修飾を後で離す
            if (keyEventFlag == NativeMethods.KEYEVENTF_KEYDOWN)
            {
                if (needShift) SendSingleVk(NativeMethods.VK_SHIFT, NativeMethods.KEYEVENTF_KEYDOWN);
                if (needCtrl) SendSingleVk(NativeMethods.VK_CONTROL, NativeMethods.KEYEVENTF_KEYDOWN);
                if (needAlt) SendSingleVk(NativeMethods.VK_MENU, NativeMethods.KEYEVENTF_KEYDOWN);

                // 通常は仮想キーを送る
                SendSingleVk(virtualKey, NativeMethods.KEYEVENTF_KEYDOWN);
            }
            else // KEYEVENTF_KEYUP
            {
                // まず本体のキーを離す
                SendSingleVk(virtualKey, NativeMethods.KEYEVENTF_KEYUP);

                if (needAlt) SendSingleVk(NativeMethods.VK_MENU, NativeMethods.KEYEVENTF_KEYUP);
                if (needCtrl) SendSingleVk(NativeMethods.VK_CONTROL, NativeMethods.KEYEVENTF_KEYUP);
                if (needShift) SendSingleVk(NativeMethods.VK_SHIFT, NativeMethods.KEYEVENTF_KEYUP);
            }
        }

        /// <summary>
        /// 単一の仮想キーイベントを SendInput で送るヘルパ
        /// </summary>
        /// <param name="virtualKey">仮想キーコード</param>
        /// <param name="flags">送信フラグ</param>
        private void SendSingleVk(ushort virtualKey, uint flags)
        {
            var input = new NativeMethods.INPUT
            {
                type = NativeMethods.INPUT_KEYBOARD,
                U = new NativeMethods.InputUnion
                {
                    ki = new NativeMethods.KEYBDINPUT
                    {
                        wVk = virtualKey,
                        wScan = 0,
                        dwFlags = flags,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            uint result = NativeMethods.SendInput(1, new NativeMethods.INPUT[] { input }, Marshal.SizeOf(typeof(NativeMethods.INPUT)));
            if (result == 0)
            {
                OnWarningOccurred($"[エラー] SendInput 失敗 GetLastWin32Error: {Marshal.GetLastWin32Error()}");
            }
        }

        /// <summary>
        /// VkKeyScan が失敗した文字（Unicode）を送るためのフォールバック実装
        /// KEYEVENTF_UNICODE を用いて wScan に Unicode を入れて送信する
        /// </summary>
        /// <param name="keyChar">送る文字</param>
        /// <param name="keyEventFlag">KEYEVENTF_* フラグ</param>
        private void SendUnicodeKey(char keyChar, uint keyEventFlag)
        {
            var input = new NativeMethods.INPUT
            {
                type = NativeMethods.INPUT_KEYBOARD,
                U = new NativeMethods.InputUnion
                {
                    ki = new NativeMethods.KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = (ushort)keyChar,
                        dwFlags = keyEventFlag | NativeMethods.KEYEVENTF_UNICODE,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            uint result = NativeMethods.SendInput(1, new NativeMethods.INPUT[] { input }, Marshal.SizeOf(typeof(NativeMethods.INPUT)));
            if (result == 0)
            {
                OnWarningOccurred($"[エラー] Unicode SendInput 失敗 GetLastWin32Error: {Marshal.GetLastWin32Error()}");
            }
        }

        /// <summary>
        /// 指定された文字に対してスキャンコード／修飾キーを用いてキーイベントを送信する
        /// VkKeyScan が失敗した場合は Unicode フォールバックを行う
        /// </summary>
        /// <param name="keyChar">送信対象の文字</param>
        /// <param name="keyEventFlag">KEYEVENTF_* のフラグ</param>
        public void SendScancodeKey(char keyChar, uint keyEventFlag)
        {
            // VkKeyScan でマップできなければ Unicode フォールバック
            short vkWithState = NativeMethods.VkKeyScan(keyChar);
            if (vkWithState == -1)
            {
                SendUnicodeKey(keyChar, keyEventFlag);
                return;
            }

            byte virtualKey = (byte)(vkWithState & 0xFF);
            byte shiftState = (byte)((vkWithState >> 8) & 0xFF);

            bool needShift = (shiftState & 1) != 0;
            bool needCtrl = (shiftState & 2) != 0;
            bool needAlt = (shiftState & 4) != 0;

            // 修飾キーの scancode を取得（MapVirtualKey: MAPVK_VK_TO_VSC = 0）
            ushort scShift = (ushort)NativeMethods.MapVirtualKey(NativeMethods.VK_SHIFT, 0);
            ushort scCtrl = (ushort)NativeMethods.MapVirtualKey(NativeMethods.VK_CONTROL, 0);
            ushort scAlt = (ushort)NativeMethods.MapVirtualKey(NativeMethods.VK_MENU, 0);

            // 主キーの scancode
            ushort scanCode = (ushort)NativeMethods.MapVirtualKey(virtualKey, 0);
            bool extended = IsExtendedKeyForVk(virtualKey);

            if (keyEventFlag == NativeMethods.KEYEVENTF_KEYDOWN)
            {
                // 修飾を先に押す
                if (needShift) SendSingleScancode(scShift, false, NativeMethods.KEYEVENTF_SCANCODE);
                if (needCtrl) SendSingleScancode(scCtrl, false, NativeMethods.KEYEVENTF_SCANCODE);
                if (needAlt) SendSingleScancode(scAlt, false, NativeMethods.KEYEVENTF_SCANCODE);

                // 主キー押下
                SendSingleScancode(scanCode, extended, NativeMethods.KEYEVENTF_SCANCODE);
            }
            else // KEYEVENTF_KEYUP
            {
                // 主キー離上
                SendSingleScancode(scanCode, extended, NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);

                // 修飾を後で離す（逆順でも可）
                if (needAlt) SendSingleScancode(scAlt, false, NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);
                if (needCtrl) SendSingleScancode(scCtrl, false, NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);
                if (needShift) SendSingleScancode(scShift, false, NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);
            }
        }

        /// <summary>
        /// 単一の scancode イベントを SendInput で送るヘルパ
        /// </summary>
        /// <param name="scancode">送信するスキャンコード</param>
        /// <param name="extended">拡張キーフラグを付けるか</param>
        /// <param name="flags">送信フラグ</param>
        private void SendSingleScancode(ushort scancode, bool extended, uint flags)
        {
            uint sendFlags = flags;
            if (extended) sendFlags |= NativeMethods.KEYEVENTF_EXTENDEDKEY;

            var input = new NativeMethods.INPUT
            {
                type = NativeMethods.INPUT_KEYBOARD,
                U = new NativeMethods.InputUnion
                {
                    ki = new NativeMethods.KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = scancode,
                        dwFlags = sendFlags,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            uint result = NativeMethods.SendInput(1, new NativeMethods.INPUT[] { input }, Marshal.SizeOf(typeof(NativeMethods.INPUT)));
            if (result == 0)
            {
                OnWarningOccurred($"[エラー] Scancode SendInput 失敗 GetLastWin32Error: {Marshal.GetLastWin32Error()}");
            }
        }

        /// <summary>
        /// 指定した仮想キーが拡張キーに該当するか判定する
        /// </summary>
        /// <param name="virtualKey">仮想キーコード</param>
        /// <returns>拡張キーなら true</returns>
        private bool IsExtendedKeyForVk(ushort virtualKey)
        {
            switch (virtualKey)
            {
                // これらは拡張キー扱い（テンキー以外の Insert/Delete/Home/End/PageUp/PageDown/矢印、右 Ctrl/右 Alt 等）
                case 0x2D: // VK_INSERT
                case 0x2E: // VK_DELETE
                case 0x24: // VK_HOME
                case 0x23: // VK_END
                case 0x21: // VK_PRIOR (PageUp)
                case 0x22: // VK_NEXT  (PageDown)
                case 0x27: // VK_RIGHT
                case 0x25: // VK_LEFT
                case 0x26: // VK_UP
                case 0x28: // VK_DOWN
                case 0xA3: // VK_RCONTROL
                case 0xA5: // VK_RMENU (Right Alt)
                case 0x6F: // VK_DIVIDE (numpad /) - often extended
                    return true;
                default:
                    return false;
            }
        }
    }
}
