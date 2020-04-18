using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Security;
using System.Threading;

namespace ConsoleApp1
{
    class Program
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [SuppressUnmanagedCodeSecurity]
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [SuppressUnmanagedCodeSecurity]
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("User32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(System.Drawing.Point p);

        [STAThread]
        static void Main(string[] args)
        {
            var proc = "\"C:\\SCAD Soft\\SCAD Office\\64\\scadx.exe\"";
            var param = "\"C:\\Users\\User\\Desktop\\4_26 — копия.SPR\"";
            var scadInfo = new ProcessStartInfo()
            {
                FileName = proc,
                Arguments = param,
                WindowStyle = ProcessWindowStyle.Maximized
            };
            Process scadProcess = Process.Start(scadInfo);

            while (scadProcess.MainWindowHandle == IntPtr.Zero) { }
            Thread.Sleep(2000);
            IntPtr scadMainWindow = scadProcess.MainWindowHandle;
            if (!SetForegroundWindow(scadMainWindow))
            {
                return;
            }

            int x = 130;
            int y = 460;
            IntPtr hWnd = WindowFromPoint(new System.Drawing.Point(x, y));
            int w = y << 16 | x;
            var z = PostMessage(hWnd, Win32.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)w);
            var v = PostMessage(hWnd, Win32.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)w);

            
        }
    }

    // <summary>
    /// Summary description for Win32.
    /// </summary>
    public class Win32
    {
        // The WM_COMMAND message is sent when the user selects a command item from 
        // a menu, when a control sends a notification message to its parent window, 
        // or when an accelerator keystroke is translated.
        public const int WM_CLOSE = 0x10;
        public const int WM_KEYDOWN = 0x100;
        public const int WM_KEYUP = 0x101;
        public const int WM_COMMAND = 0x111;
        public const int WM_LBUTTONDOWN = 0x201;
        public const int WM_LBUTTONUP = 0x202;
        public const int WM_LBUTTONDBLCLK = 0x203;
        public const int WM_RBUTTONDOWN = 0x204;
        public const int WM_RBUTTONUP = 0x205;
        public const int WM_RBUTTONDBLCLK = 0x206;
        public const int WM_SETTEXT = 0x000c;

        public const int VK_TAB = 0x09;
    }
}
