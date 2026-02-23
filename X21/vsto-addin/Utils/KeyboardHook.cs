using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using X21.Logging;

namespace X21.Utils
{
    /// <summary>
    /// Local keyboard hook to capture shortcuts within Excel using Microsoft's recommended approach
    /// Ctrl+Shift+A - Focus AI chat text input in the task pane
    /// Ctrl+Shift+M - Fix formula in current cell
    /// Based on: https://learn.microsoft.com/en-us/archive/blogs/vsod/using-shortcut-keys-to-call-a-function-in-an-office-add-in
    /// </summary>
    public class KeyboardHook : IDisposable
    {
        private const int WH_KEYBOARD = 2;  // Local keyboard hook (Microsoft recommended)
        private const int HC_ACTION = 0;

        public event Action CtrlShiftAPressed;
        public event Action CtrlShiftMPressed;
        public event Action CtrlShiftYPressed;

        public delegate int LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        private static KeyboardHook _instance;

        public KeyboardHook()
        {
            _instance = this;
            Logger.Info("Creating local keyboard hook for multiple shortcuts...");
            SetHook();
        }

        public static void SetHook()
        {
            // Use local hook with current *native* (Win32) thread id.
            // `AppDomain.GetCurrentThreadId()` returned the Win32 thread id but is deprecated; use kernel32 instead.
            uint threadId = GetCurrentThreadId();
            _hookID = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, threadId);
            if (_hookID != IntPtr.Zero)
            {
                Logger.Info($"Local keyboard hook created successfully (threadId={threadId})");
            }
            else
            {
                var err = Marshal.GetLastWin32Error();
                Logger.Info($"Failed to create local keyboard hook (threadId={threadId}, win32Error={err})");
            }
        }

        public static void ReleaseHook()
        {
            if (_hookID != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hookID);
                _hookID = IntPtr.Zero;
                Logger.Info("Local keyboard hook released");
            }
        }

        private static int HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
            }

            if (nCode == HC_ACTION)
            {
                Keys keyData = (Keys)wParam;

                // Check if key was already pressed (to avoid repeated events)
                int PreviousStateBit = 31;
                Int64 bitmask = (Int64)Math.Pow(2, (PreviousStateBit - 1));
                bool keyWasAlreadyPressed = ((Int64)lParam & bitmask) > 0;

                Logger.Info($"Key detected: {keyData}, Already pressed: {keyWasAlreadyPressed}");

                // Check for different key combinations
                if (!keyWasAlreadyPressed)
                {
                    bool ctrlPressed = IsKeyDown(Keys.ControlKey);
                    bool shiftPressed = IsKeyDown(Keys.ShiftKey);

                    // Ctrl+Shift+A for AI Chat
                    if (keyData == Keys.A && ctrlPressed && shiftPressed)
                    {
                        Logger.Info("Ctrl+Shift+A combination detected - firing event");
                        _instance?.OnCtrlShiftADetected();
                        return 1; // Consume the key event
                    }

                    // Ctrl+Shift+M for Fix Formula
                    if (keyData == Keys.M && ctrlPressed && shiftPressed)
                    {
                        Logger.Info("Ctrl+Shift+M combination detected - firing formula fix event");
                        _instance?.OnCtrlShiftMDetected();
                        return 1; // Consume the key event
                    }

                    // Ctrl+Shift+Y for Apply/Revert Changes
                    if (keyData == Keys.Y && ctrlPressed && shiftPressed)
                    {
                        Logger.Info("Ctrl+Shift+Y combination detected - firing apply/revert event");
                        _instance?.OnCtrlShiftYDetected();
                        return 1; // Consume the key event
                    }
                }
            }

            return (int)CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        public static bool IsKeyDown(Keys keys)
        {
            return (GetKeyState((int)keys) & 0x8000) == 0x8000;
        }

        private void OnCtrlShiftADetected()
        {
            try
            {
                Logger.Info("OnCtrlShiftADetected called - invoking CtrlShiftAPressed event");
                CtrlShiftAPressed?.Invoke();
                Logger.Info("CtrlShiftAPressed event invoked successfully");
            }
            catch (Exception ex)
            {
                Logger.Info($"Exception in OnCtrlShiftADetected: {ex.Message}");
            }
        }

        private void OnCtrlShiftMDetected()
        {
            try
            {
                Logger.Info("OnCtrlShiftMDetected called - invoking CtrlShiftMPressed event");
                CtrlShiftMPressed?.Invoke();
                Logger.Info("CtrlShiftMPressed event invoked successfully");
            }
            catch (Exception ex)
            {
                Logger.Info($"Exception in OnCtrlShiftMDetected: {ex.Message}");
            }
        }

        private void OnCtrlShiftYDetected()
        {
            try
            {
                Logger.Info("OnCtrlShiftYDetected called - invoking CtrlShiftYPressed event");
                CtrlShiftYPressed?.Invoke();
                Logger.Info("CtrlShiftYPressed event invoked successfully");
            }
            catch (Exception ex)
            {
                Logger.Info($"Exception in OnCtrlShiftYDetected: {ex.Message}");
            }
        }

        public void Dispose()
        {
            ReleaseHook();
            _instance = null;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook,
            LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
            IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();
    }
}
