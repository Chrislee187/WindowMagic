using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OpenQA.Selenium;

namespace WebDriverElements.Samples.Feature.Tests
{

    public class WindowMagic
    {
        /// <summary>
        /// If the method can detect you are running from visual studio or the resharper test runner it moves the first window to match <paramref name="windowName"/> to the first monitor that doesn't have visual studio on it.
        /// Ideal to move browser windows on to another monitor.
        /// 
        /// <example>
        ///     new WindowMagic().MoveToNonVisualStudioMonitor("Firefox");
        /// </example>
        /// </summary>
        /// <param name="windowName"></param>
        /// <param name="maximise"></param>
        public void MoveToNonVisualStudioMonitor(string windowName, bool maximise = true)
        {
            var sourceAppHandle = GetWindowHandle(windowName);

            var moveToScreen = FindNonVisualStudioMonitor(sourceAppHandle);

            if (moveToScreen == null) return;

            var sourceRect = User32.GetWindowRectCS(sourceAppHandle);
            var width = maximise ? 200 : sourceRect.Right - sourceRect.Left;
            var height = maximise ? 200 : sourceRect.Bottom - sourceRect.Top;

            User32.SetWindowPos(sourceAppHandle, 0, moveToScreen.Bounds.Left, moveToScreen.Bounds.Top, width, height, User32.SetWindowPostFlags.SWP_NOZORDER | User32.SetWindowPostFlags.SWP_SHOWWINDOW);

            if(maximise) User32.ShowWindow(sourceAppHandle, User32.SW_SHOWMAXIMIZED);
        }

        private static Screen FindNonVisualStudioMonitor(IntPtr sourceAppHandle)
        {
            var devenvProcess = RunningFromVisualStudio();

            if (devenvProcess == null) return null;

            var devenvScreen = Screen.FromHandle(devenvProcess.MainWindowHandle);

            var sourceAppScreen = Screen.FromHandle(sourceAppHandle);

            var moveToScreen = Screen.AllScreens.First(s => s.DeviceName != devenvScreen.DeviceName);

            if (devenvScreen.DeviceName != sourceAppScreen.DeviceName) moveToScreen = null;
            return moveToScreen;
        }

        private static Process RunningFromVisualStudio()
        {
            var match = new []{"devenv", "JetBrains.ReSharper.TaskRunner"};

            var devenvProcess = Process.GetCurrentProcess();

            if (match.Any(m => devenvProcess.ProcessName.IndexOf(m) >= 0)) return devenvProcess;

            return null;
        }

        public void Move(string windowName, int x, int y)
        {
            var handle = GetWindowHandle(windowName);

            var rect = User32.GetWindowRectCS(handle);

            User32.SetWindowPos(handle, 0, x, y, rect.Right - rect.Left, rect.Bottom - rect.Top, User32.SetWindowPostFlags.SWP_NOZORDER | User32.SetWindowPostFlags.SWP_SHOWWINDOW);
        }

        private static IntPtr GetWindowHandle(string windowName)
        {
            var process = Process.GetProcesses().FirstOrDefault(p => p.MainWindowTitle.Contains(windowName));
            if (process == null)
            {
                throw new NoSuchWindowException(string.Format("Couldn't get process for window [{0}]", windowName));
            }


            var handle = process.MainWindowHandle;
            return handle;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        private class User32
        {
            #region user32::SetWindowPos()

            [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
            public static extern IntPtr SetWindowPos(IntPtr hWnd, int hWndInsertAfter, int x, int y, int cx, int cy, User32.SetWindowPostFlags wFlags);

            [Flags]
            public enum SetWindowPostFlags
            {
                SWP_NOSIZE = 0x1,
                SWP_NOMOVE = 0x2,
                SWP_NOZORDER = 0x4,
                SWP_SHOWWINDOW = 0x0040
            }

            #endregion

            #region user32::GetWindowRect()

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            public static extern bool GetWindowRect(HandleRef hWnd, out RECT lpRect);

            public static RECT GetWindowRectCS(HandleRef handleRef)
            {
                RECT rect;

                if (!GetWindowRect(handleRef, out rect))
                {
                    throw new NoSuchWindowException();
                }
                return rect;
            }

            public static RECT GetWindowRectCS(IntPtr handle)
            {
                return GetWindowRectCS(new HandleRef(new RefWrapper(handle), handle));
            }
            #endregion

            #region user32::ShowWindow()

            [DllImport("user32.dll")]
            public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

            public const int SW_SHOWMAXIMIZED = 3;

            #endregion
        }

        private class RefWrapper
        {
            private IntPtr _handle;

            public RefWrapper(IntPtr handle)
            {
                _handle = handle;
            }
        }
    }

}