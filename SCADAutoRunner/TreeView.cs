using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace SCADAutoRunner
{
    class TreeView
    {
        private const int MaxCorrectCoord = 10000;

        private IntPtr handle;

        public TreeView()
        {
            handle = Win32API.WindowFromPoint(new Point(Settings.DefaultX, Settings.DefaultY));
        }

        /// <summary>
        /// Получает пункт "Расчёт"->"Линейный" в дереве действий главного окна SCAD
        /// </summary>
        /// <returns>Handle пункта "Линеный"</returns>
        public IntPtr GetLinearCalcuationsItem()
        {
            IntPtr treeItem = IntPtr.Zero;
            while (treeItem == IntPtr.Zero)
            {
                handle = Win32API.WindowFromPoint(new Point(Settings.DefaultX, Settings.DefaultY));
                treeItem = Win32API.SendMessage(handle, TreeViewMessages.TVM_GETNEXTITEM, TreeViewMessages.GetRoot, IntPtr.Zero);
                treeItem = Win32API.SendMessage(handle, TreeViewMessages.TVM_GETNEXTITEM, TreeViewMessages.GetNext, treeItem);
                treeItem = Win32API.SendMessage(handle, TreeViewMessages.TVM_GETNEXTITEM, TreeViewMessages.GetChild, treeItem);
                Win32API.SendMessage(handle, TreeViewMessages.TVM_SELECTITEM, TreeViewMessages.SelectItem, treeItem);
            }
            return treeItem;
        }

        /// <summary>
        /// Получает пункт "Результаты"->"Документирование" в дереве действий главного окна SCAD
        /// </summary>
        /// <returns>Handle пункта "Документирование"</returns>
        public IntPtr GetDocResultItem()
        {
            IntPtr treeItem = IntPtr.Zero;
            while (treeItem == IntPtr.Zero)
            {
                handle = Win32API.WindowFromPoint(new Point(Settings.DefaultX, Settings.DefaultY));
                treeItem = Win32API.SendMessage(handle, TreeViewMessages.TVM_GETNEXTITEM, TreeViewMessages.GetRoot, IntPtr.Zero);
                treeItem = Win32API.SendMessage(handle, TreeViewMessages.TVM_GETNEXTITEM, TreeViewMessages.GetNext, treeItem);
                treeItem = Win32API.SendMessage(handle, TreeViewMessages.TVM_GETNEXTITEM, TreeViewMessages.GetNext, treeItem);
                treeItem = Win32API.SendMessage(handle, TreeViewMessages.TVM_GETNEXTITEM, TreeViewMessages.GetChild, treeItem);
                treeItem = Win32API.SendMessage(handle, TreeViewMessages.TVM_GETNEXTITEM, TreeViewMessages.GetNext, treeItem);
                treeItem = Win32API.SendMessage(handle, TreeViewMessages.TVM_GETNEXTITEM, TreeViewMessages.GetChild, treeItem);
                Win32API.SendMessage(handle, TreeViewMessages.TVM_SELECTITEM, TreeViewMessages.SelectItem, treeItem);
                Input.ShortDelay();
            }
            return treeItem;
        }

        /// <summary>
        /// Возвращает точку, расположенную по центру указанного пункта
        /// </summary>
        /// <param name="item">Handle пункта дерева</param>
        /// <returns>Точка по центру указанного пункта</returns>
        public Point GetItemPoint(IntPtr item)
        {
            RECT rct = new RECT();
            while (rct.Left > MaxCorrectCoord ||
                   rct.Top > MaxCorrectCoord ||
                   rct.Left <= 0 ||
                   rct.Top <= 0)
            {
                Win32API.GetWindowThreadProcessId(handle, out int pid);
                IntPtr process = Win32API.OpenProcess(
                    Win32API.ProcessAccessFlags.VirtualMemoryOperation |
                    Win32API.ProcessAccessFlags.VirtualMemoryRead |
                    Win32API.ProcessAccessFlags.VirtualMemoryWrite |
                    Win32API.ProcessAccessFlags.QueryInformation,
                    false, pid);
                IntPtr pBool = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(int)));
                Marshal.WriteInt32(pBool, 0, 1);
                int sizeRct = Marshal.SizeOf(typeof(RECT));
                IntPtr hRct = Marshal.AllocHGlobal(sizeRct);
                var hRecItem = Win32API.VirtualAllocEx(process, IntPtr.Zero, sizeRct, Win32API.AllocationType.Commit, Win32API.MemoryProtection.ReadWrite);
                Win32API.WriteProcessMemory(process, hRecItem, item, sizeRct, out IntPtr written);
                Win32API.SendMessage(handle, TreeViewMessages.TVM_GETITEMRECT, pBool, hRecItem);
                Win32API.ReadProcessMemory(process, hRecItem, hRct, sizeRct, out IntPtr read);
                rct = (RECT)Marshal.PtrToStructure(hRct, typeof(RECT));
                Win32API.CloseHandle(process);
            }

            Win32API.GetWindowRect(handle, out RECT treeRect);
            int x = (rct.Left + rct.Right) / 2 + treeRect.Left;
            int y = (rct.Top + rct.Bottom) / 2 + treeRect.Top;
            return new Point(x, y);
        }
    }

    /// <summary>
    /// Сообщения для работы с TreeView
    /// </summary>
    public static class TreeViewMessages
    {
        public static IntPtr GetRoot = (IntPtr)0x0;
        public static IntPtr GetNext = (IntPtr)0x1;
        public static IntPtr GetChild = (IntPtr)0x4;
        public static IntPtr SelectItem = (IntPtr)0x9;

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

        private const int TV_FIRST = 0x1100;
    }
}
