using System;
using System.Runtime.InteropServices;
using System.Security;

namespace SCADAutoRunner
{
    /// <summary>
    /// user32.dll functions definitions
    /// </summary>
    static class User32
    {
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [SuppressUnmanagedCodeSecurity]
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr FindWindow(string className, string windowTitle);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(System.Drawing.Point p);

        [DllImport("user32.dll", EntryPoint = "FindWindowEx", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr OpenProcess(ProcessAccessFlags processAccess, bool bInheritHandle, int processId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowUnicode(IntPtr hWnd);

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        public static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, int dwSize, AllocationType flAllocationType, MemoryProtection flProtect);

        [SuppressUnmanagedCodeSecurity]
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [Flags]
        public enum ProcessAccessFlags : uint
        {
            All = 0x001F0FFF,
            Terminate = 0x00000001,
            CreateThread = 0x00000002,
            VirtualMemoryOperation = 0x00000008,
            VirtualMemoryRead = 0x00000010,
            VirtualMemoryWrite = 0x00000020,
            DuplicateHandle = 0x00000040,
            CreateProcess = 0x000000080,
            SetQuota = 0x00000100,
            SetInformation = 0x00000200,
            QueryInformation = 0x00000400,
            QueryLimitedInformation = 0x00001000,
            Synchronize = 0x00100000
        }

        [Flags]
        public enum AllocationType
        {
            Commit = 0x1000,
            Reserve = 0x2000,
            Decommit = 0x4000,
            Release = 0x8000,
            Reset = 0x80000,
            Physical = 0x400000,
            TopDown = 0x100000,
            WriteWatch = 0x200000,
            LargePages = 0x20000000
        }

        [Flags]
        public enum MemoryProtection
        {
            Execute = 0x10,
            ExecuteRead = 0x20,
            ExecuteReadWrite = 0x40,
            ExecuteWriteCopy = 0x80,
            NoAccess = 0x01,
            ReadOnly = 0x02,
            ReadWrite = 0x04,
            WriteCopy = 0x08,
            GuardModifierflag = 0x100,
            NoCacheModifierflag = 0x200,
            WriteCombineModifierflag = 0x400
        }

        private const int TV_FIRST = 0x1100;
        public class TVM
        {
            public static uint TVM_GETNEXTITEM = (TV_FIRST + 10);
            public static uint TVM_GETITEMA = (TV_FIRST + 12);
            public static uint TVM_GETITEM = (TV_FIRST + 62);
            public static uint TVM_GETCOUNT = (TV_FIRST + 5);
            public static uint TVM_SELECTITEM = (TV_FIRST + 11);
            public static uint TVM_DELETEITEM = (TV_FIRST + 1);
            public static uint TVM_EXPAND = (TV_FIRST + 2);
            public static uint TVM_GETITEMRECT = (TV_FIRST + 4);
            public static uint TVM_GETINDENT = (TV_FIRST + 6);
            public static uint TVM_SETINDENT = (TV_FIRST + 7);
            public static uint TVM_GETIMAGELIST = (TV_FIRST + 8);
            public static uint TVM_SETIMAGELIST = (TV_FIRST + 9);
            public static uint TVM_GETISEARCHSTRING = (TV_FIRST + 64);
            public static uint TVM_HITTEST = (TV_FIRST + 17);
        }
    }
}
