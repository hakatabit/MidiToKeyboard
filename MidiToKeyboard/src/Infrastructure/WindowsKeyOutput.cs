using System;
using System.Runtime.InteropServices;
using MidiToKeyboard.Application;
using MidiToKeyboard.Domain;

namespace MidiToKeyboard.Infrastructure
{
    /// <summary>
    /// Windows の SendInput を使用したキー操作の送信
    /// </summary>
    public sealed class WindowsKeyOutput : IKeyOutput, IInputModeKeyOutput
    {
        private InputMode _inputMode = InputMode.VirtualKey;

        /// <summary>
        /// キー操作の送信に失敗したときの通知
        /// </summary>
        public event EventHandler<string> WarningOccurred;

        /// <inheritdoc />
        public void SetInputMode(InputMode inputMode)
        {
            if (inputMode != InputMode.VirtualKey && inputMode != InputMode.Scancode)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(inputMode),
                    inputMode,
                    "入力送信モードは VirtualKey または Scancode である必要があります。");
            }

            _inputMode = inputMode;
        }

        /// <inheritdoc />
        public void Send(KeyAction action)
        {
            if (action == null)
            {
                throw new ArgumentNullException(
                    nameof(action),
                    "キー操作は null にできません。");
            }

            switch (action.Type)
            {
                case KeyActionType.KeyDown:
                    SendKey(action.KeyChar, NativeMethods.KEYEVENTF_KEYDOWN);
                    return;

                case KeyActionType.KeyUp:
                    SendKey(action.KeyChar, NativeMethods.KEYEVENTF_KEYUP);
                    return;

                case KeyActionType.KeyPress:
                    SendKey(action.KeyChar, NativeMethods.KEYEVENTF_KEYDOWN);
                    SendKey(action.KeyChar, NativeMethods.KEYEVENTF_KEYUP);
                    return;

                case KeyActionType.UnicodeText:
                    SendUnicodeText(action.Text);
                    return;

                default:
                    throw new NotSupportedException(
                        "サポートされていない KeyActionType です: " + action.Type);
            }
        }

        private void SendKey(char keyChar, uint keyEventFlag)
        {
            if (_inputMode == InputMode.Scancode)
            {
                SendScancodeKey(keyChar, keyEventFlag);
            }
            else
            {
                SendKeyInput(keyChar, keyEventFlag);
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
        /// 指定された文字に対応する仮想キーイベントを送信
        /// </summary>
        /// <param name="keyChar">送信対象の文字</param>
        /// <param name="keyEventFlag">KEYEVENTF_KEYDOWN または KEYEVENTF_KEYUP</param>
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

            bool needsShift = (shiftState & 1) != 0;
            bool needsControl = (shiftState & 2) != 0;
            bool needsAlt = (shiftState & 4) != 0;

            // キーダウン時は修飾を先に押し、キーアップ時は修飾を後で離す
            if (keyEventFlag == NativeMethods.KEYEVENTF_KEYDOWN)
            {
                if (needsShift)
                {
                    SendSingleVirtualKey(NativeMethods.VK_SHIFT, NativeMethods.KEYEVENTF_KEYDOWN);
                }

                if (needsControl)
                {
                    SendSingleVirtualKey(NativeMethods.VK_CONTROL, NativeMethods.KEYEVENTF_KEYDOWN);
                }

                if (needsAlt)
                {
                    SendSingleVirtualKey(NativeMethods.VK_MENU, NativeMethods.KEYEVENTF_KEYDOWN);
                }

                SendSingleVirtualKey(virtualKey, NativeMethods.KEYEVENTF_KEYDOWN);
            }
            else
            {
                SendSingleVirtualKey(virtualKey, NativeMethods.KEYEVENTF_KEYUP);

                if (needsAlt)
                {
                    SendSingleVirtualKey(NativeMethods.VK_MENU, NativeMethods.KEYEVENTF_KEYUP);
                }

                if (needsControl)
                {
                    SendSingleVirtualKey(NativeMethods.VK_CONTROL, NativeMethods.KEYEVENTF_KEYUP);
                }

                if (needsShift)
                {
                    SendSingleVirtualKey(NativeMethods.VK_SHIFT, NativeMethods.KEYEVENTF_KEYUP);
                }
            }
        }

        private void SendSingleVirtualKey(ushort virtualKey, uint flags)
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

            uint result = NativeMethods.SendInput(
                1,
                new NativeMethods.INPUT[] { input },
                Marshal.SizeOf(typeof(NativeMethods.INPUT)));
            if (result == 0)
            {
                OnWarningOccurred($"[エラー] SendInput 失敗 GetLastWin32Error: {Marshal.GetLastWin32Error()}");
            }
        }

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

            uint result = NativeMethods.SendInput(
                1,
                new NativeMethods.INPUT[] { input },
                Marshal.SizeOf(typeof(NativeMethods.INPUT)));
            if (result == 0)
            {
                OnWarningOccurred($"[エラー] Unicode SendInput 失敗 GetLastWin32Error: {Marshal.GetLastWin32Error()}");
            }
        }

        /// <summary>
        /// 指定された文字に対応するスキャンコードイベントを送信
        /// </summary>
        /// <param name="keyChar">送信対象の文字</param>
        /// <param name="keyEventFlag">KEYEVENTF_KEYDOWN または KEYEVENTF_KEYUP</param>
        public void SendScancodeKey(char keyChar, uint keyEventFlag)
        {
            short vkWithState = NativeMethods.VkKeyScan(keyChar);
            if (vkWithState == -1)
            {
                SendUnicodeKey(keyChar, keyEventFlag);
                return;
            }

            byte virtualKey = (byte)(vkWithState & 0xFF);
            byte shiftState = (byte)((vkWithState >> 8) & 0xFF);

            bool needsShift = (shiftState & 1) != 0;
            bool needsControl = (shiftState & 2) != 0;
            bool needsAlt = (shiftState & 4) != 0;

            ushort shiftScanCode = (ushort)NativeMethods.MapVirtualKey(NativeMethods.VK_SHIFT, 0);
            ushort controlScanCode = (ushort)NativeMethods.MapVirtualKey(NativeMethods.VK_CONTROL, 0);
            ushort altScanCode = (ushort)NativeMethods.MapVirtualKey(NativeMethods.VK_MENU, 0);

            ushort scanCode = (ushort)NativeMethods.MapVirtualKey(virtualKey, 0);
            bool isExtendedKey = IsExtendedVirtualKey(virtualKey);

            if (keyEventFlag == NativeMethods.KEYEVENTF_KEYDOWN)
            {
                if (needsShift)
                {
                    SendSingleScancode(shiftScanCode, false, NativeMethods.KEYEVENTF_SCANCODE);
                }

                if (needsControl)
                {
                    SendSingleScancode(controlScanCode, false, NativeMethods.KEYEVENTF_SCANCODE);
                }

                if (needsAlt)
                {
                    SendSingleScancode(altScanCode, false, NativeMethods.KEYEVENTF_SCANCODE);
                }

                SendSingleScancode(scanCode, isExtendedKey, NativeMethods.KEYEVENTF_SCANCODE);
            }
            else
            {
                SendSingleScancode(
                    scanCode,
                    isExtendedKey,
                    NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);

                if (needsAlt)
                {
                    SendSingleScancode(
                        altScanCode,
                        false,
                        NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);
                }

                if (needsControl)
                {
                    SendSingleScancode(
                        controlScanCode,
                        false,
                        NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);
                }

                if (needsShift)
                {
                    SendSingleScancode(
                        shiftScanCode,
                        false,
                        NativeMethods.KEYEVENTF_SCANCODE | NativeMethods.KEYEVENTF_KEYUP);
                }
            }
        }

        private void SendSingleScancode(ushort scanCode, bool isExtendedKey, uint flags)
        {
            uint sendFlags = flags;
            if (isExtendedKey)
            {
                sendFlags |= NativeMethods.KEYEVENTF_EXTENDEDKEY;
            }

            var input = new NativeMethods.INPUT
            {
                type = NativeMethods.INPUT_KEYBOARD,
                U = new NativeMethods.InputUnion
                {
                    ki = new NativeMethods.KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = scanCode,
                        dwFlags = sendFlags,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };

            uint result = NativeMethods.SendInput(
                1,
                new NativeMethods.INPUT[] { input },
                Marshal.SizeOf(typeof(NativeMethods.INPUT)));
            if (result == 0)
            {
                OnWarningOccurred($"[エラー] Scancode SendInput 失敗 GetLastWin32Error: {Marshal.GetLastWin32Error()}");
            }
        }

        private bool IsExtendedVirtualKey(ushort virtualKey)
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
