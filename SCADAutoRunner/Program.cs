using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Security;
using System.Threading;
using System.Collections.Generic;
using System.IO;

namespace SCADAutoRunner
{
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT
    {
        public int Left;        // x position of upper-left corner
        public int Top;         // y position of upper-left corner
        public int Right;       // x position of lower-right corner
        public int Bottom;      // y position of lower-right corner
    }

    class Program
    {
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
            if (!User32.SetForegroundWindow(scadMainWindow))
            {
                return;
            }

            int x = 130;
            int y = 420;
            User32.SetForegroundWindow(scadMainWindow);
            IntPtr hTreeView = User32.WindowFromPoint(new System.Drawing.Point(x, y));

            IntPtr treeItem = IntPtr.Zero;
            //while (treeItem == IntPtr.Zero)
            //{
            //    hTreeView = User32.WindowFromPoint(new System.Drawing.Point(x, y));
            //    treeItem = User32.SendMessage(hTreeView, User32.TVM.TVM_GETNEXTITEM, (IntPtr)0x0, IntPtr.Zero);
            //    treeItem = User32.SendMessage(hTreeView, User32.TVM.TVM_GETNEXTITEM, (IntPtr)0x1, treeItem);
            //    treeItem = User32.SendMessage(hTreeView, User32.TVM.TVM_GETNEXTITEM, (IntPtr)0x4, treeItem);
            //    User32.SendMessage(hTreeView, User32.TVM.TVM_SELECTITEM, (IntPtr)0x9, treeItem);
            //}

            //RECT rct = new RECT();
            //while (rct.Left > 10000 || rct.Top > 10000 || rct.Left <= 0 || rct.Top <= 0)
            //{
            //    User32.GetWindowThreadProcessId(hTreeView, out int pid);
            //    IntPtr process = User32.OpenProcess(
            //        User32.ProcessAccessFlags.VirtualMemoryOperation |
            //        User32.ProcessAccessFlags.VirtualMemoryRead |
            //        User32.ProcessAccessFlags.VirtualMemoryWrite |
            //        User32.ProcessAccessFlags.QueryInformation,
            //        false, pid);
            //    IntPtr pBool = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(int)));
            //    Marshal.WriteInt32(pBool, 0, 1);
            //    int sizeRct = Marshal.SizeOf(typeof(RECT));
            //    IntPtr hRct = Marshal.AllocHGlobal(sizeRct);
            //    var hRecItem = User32.VirtualAllocEx(process, IntPtr.Zero, sizeRct, User32.AllocationType.Commit, User32.MemoryProtection.ReadWrite);
            //    User32.WriteProcessMemory(process, hRecItem, treeItem, sizeRct, out IntPtr written);
            //    User32.SendMessage(hTreeView, User32.TVM.TVM_GETITEMRECT, pBool, hRecItem);
            //    User32.ReadProcessMemory(process, hRecItem, hRct, sizeRct, out IntPtr read);
            //    rct = (RECT)Marshal.PtrToStructure(hRct, typeof(RECT));
            //    User32.CloseHandle(process);
            //}

            //User32.GetWindowRect(hTreeView, out RECT treeRect);
            //x = (rct.Left + rct.Right) / 2 + treeRect.Left;
            //y = (rct.Top + rct.Bottom) / 2 + treeRect.Top;
            //int w = y << 16 | x;

            //IntPtr hParams = IntPtr.Zero;
            //while (hParams == IntPtr.Zero)
            //{
            //    User32.PostMessage(hTreeView, Win32.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)w);
            //    User32.PostMessage(hTreeView, Win32.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)w);
            //    hParams = User32.FindWindow("#32770", "Параметры расчета");
            //    Thread.Sleep(2000);
            //}
            //User32.SetForegroundWindow(hParams);
            //User32.PostMessage(hParams, Win32.WM_KEYDOWN, (IntPtr)Win32.VK_RETURN, (IntPtr)0x0);

            //Process solverProcess = null;
            //while (solverProcess == null)
            //{
            //    var procs = Process.GetProcessesByName("SprSolverExe");
            //    if (procs.Any())
            //    {
            //        solverProcess = procs.First();
            //    }
            //    Thread.Sleep(2000);
            //}
            //solverProcess.WaitForExit();
            //Thread.Sleep(2000);
            //scadProcess.Kill();

            //scadProcess = Process.Start(scadInfo);
            //while (scadProcess.MainWindowHandle == IntPtr.Zero) { }
            //Thread.Sleep(2000);
            //scadMainWindow = scadProcess.MainWindowHandle;
            //if (!User32.SetForegroundWindow(scadMainWindow))
            //{
            //    return;
            //}

            //x = 130;
            //y = 420;
            //User32.SetForegroundWindow(scadMainWindow);
            //hTreeView = User32.WindowFromPoint(new System.Drawing.Point(x, y));

            //treeItem = IntPtr.Zero;
            while (treeItem == IntPtr.Zero)
            {
                hTreeView = User32.WindowFromPoint(new System.Drawing.Point(x, y));
                treeItem = User32.SendMessage(hTreeView, User32.TVM.TVM_GETNEXTITEM, (IntPtr)0x0, IntPtr.Zero);
                treeItem = User32.SendMessage(hTreeView, User32.TVM.TVM_GETNEXTITEM, (IntPtr)0x1, treeItem);
                treeItem = User32.SendMessage(hTreeView, User32.TVM.TVM_GETNEXTITEM, (IntPtr)0x1, treeItem);
                treeItem = User32.SendMessage(hTreeView, User32.TVM.TVM_GETNEXTITEM, (IntPtr)0x4, treeItem);
                treeItem = User32.SendMessage(hTreeView, User32.TVM.TVM_GETNEXTITEM, (IntPtr)0x1, treeItem);
                treeItem = User32.SendMessage(hTreeView, User32.TVM.TVM_GETNEXTITEM, (IntPtr)0x1, treeItem);
                User32.SendMessage(hTreeView, User32.TVM.TVM_SELECTITEM, (IntPtr)0x9, treeItem);
                Thread.Sleep(500);
            }

            RECT rct = new RECT();
            while (rct.Left > 10000 || rct.Top > 10000 || rct.Left <= 0 || rct.Top <= 0)
            {
                User32.GetWindowThreadProcessId(hTreeView, out int pid);
                IntPtr process = User32.OpenProcess(
                    User32.ProcessAccessFlags.VirtualMemoryOperation |
                    User32.ProcessAccessFlags.VirtualMemoryRead |
                    User32.ProcessAccessFlags.VirtualMemoryWrite |
                    User32.ProcessAccessFlags.QueryInformation,
                    false, pid);
                IntPtr pBool = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(int)));
                Marshal.WriteInt32(pBool, 0, 1);
                int sizeRct = Marshal.SizeOf(typeof(RECT));
                IntPtr hRct = Marshal.AllocHGlobal(sizeRct);
                var hRecItem = User32.VirtualAllocEx(process, IntPtr.Zero, sizeRct, User32.AllocationType.Commit, User32.MemoryProtection.ReadWrite);
                User32.WriteProcessMemory(process, hRecItem, treeItem, sizeRct, out IntPtr written);
                User32.SendMessage(hTreeView, User32.TVM.TVM_GETITEMRECT, pBool, hRecItem);
                User32.ReadProcessMemory(process, hRecItem, hRct, sizeRct, out IntPtr read);
                rct = (RECT)Marshal.PtrToStructure(hRct, typeof(RECT));
                User32.CloseHandle(process);
                Thread.Sleep(500);
            }

            User32.GetWindowRect(hTreeView, out RECT treeRect);
            x = (rct.Left + rct.Right) / 2 + treeRect.Left;
            y = (rct.Top + rct.Bottom) / 2 + treeRect.Top;
            int w = y << 16 | x;

            IntPtr hParams = IntPtr.Zero;
            while (hParams == IntPtr.Zero)
            {
                User32.PostMessage(hTreeView, Win32.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)w);
                User32.PostMessage(hTreeView, Win32.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)w);
                hParams = User32.FindWindow("#32770", "Результаты расчета");
                Thread.Sleep(2000);
            }

            IntPtr hParent = hParams;
            hParams = User32.FindWindowEx(hParams, IntPtr.Zero, "#32770", "Вывод результатов");
            var resSpreads = GetAllChildrenWindowHandles(hParams, "fpSpread80");
            var resButtons = GetAllChildrenWindowHandles(hParams, "Button");

            User32.SetForegroundWindow(resSpreads[1]);
            User32.PostMessage(resSpreads[1], Win32.WM_KEYDOWN, (IntPtr)Win32.VK_DOWN, IntPtr.Zero);
            User32.PostMessage(resSpreads[1], Win32.WM_KEYDOWN, (IntPtr)Win32.VK_DOWN, IntPtr.Zero);

            User32.GetWindowRect(resButtons[2], out RECT btnRect);
            x = (btnRect.Left + btnRect.Right) / 2;
            y = (btnRect.Top + btnRect.Bottom) / 2;
            w = y << 16 | x;
            IntPtr hwnd = User32.WindowFromPoint(new System.Drawing.Point(x, y));
            User32.SetForegroundWindow(hParams);
            IntPtr hChild = IntPtr.Zero;
            //User32.SetForegroundWindow(resButtons[2]);
            while (hChild == IntPtr.Zero)
            {
                User32.PostMessage(hwnd, Win32.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)w);
                User32.PostMessage(hwnd, Win32.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)w);
                User32.PostMessage(hwnd, Win32.WM_KEYDOWN, (IntPtr)Win32.VK_SPACE, (IntPtr)0x0);
                x = 465;
                y = 290;
                int yu = y << 16 | x;
                User32.PostMessage(hParams, Win32.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)yu);
                User32.PostMessage(hParams, Win32.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)yu);
                hChild = User32.FindWindow("#32770", "Величины перемещений");
                Thread.Sleep(500);
            }

            var okBtn = GetAllChildrenWindowHandles(hChild, "Button").First();
            User32.SetForegroundWindow(okBtn);
            User32.PostMessage(hChild, Win32.WM_KEYDOWN, (IntPtr)Win32.VK_RETURN, IntPtr.Zero);

            //User32.PostMessage(hwnd, Win32.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)w);
            //User32.PostMessage(hwnd, Win32.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)w);

            var createXlsBtn = GetAllChildrenWindowHandles(hParent, "Button").Last();
            User32.SetForegroundWindow(createXlsBtn);
            User32.GetWindowRect(createXlsBtn, out btnRect);
            x = (btnRect.Left + btnRect.Right) / 2;
            y = (btnRect.Top + btnRect.Bottom) / 2;
            w = y << 16 | x;
            User32.ShowWindow(scadMainWindow, User32.ShowWindowCommands.Minimize);
            User32.ShowWindow(scadMainWindow, User32.ShowWindowCommands.Maximize);

            IntPtr hSave = IntPtr.Zero;
            hwnd = User32.WindowFromPoint(new System.Drawing.Point(x, y));
            while (hSave == IntPtr.Zero)
            {
                User32.SetFocus(hParent);
                User32.SetForegroundWindow(createXlsBtn);
                User32.PostMessage(hwnd, Win32.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)w);
                User32.PostMessage(hwnd, Win32.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)w);
                User32.PostMessage(hwnd, Win32.WM_KEYDOWN, (IntPtr)Win32.VK_SPACE, IntPtr.Zero);
                hSave = User32.FindWindow("#32770", "Сохранение");
                Thread.Sleep(500);
            }

            Thread.Sleep(500);
            var cmbParent = GetAllChildrenWindowHandles(hSave, "ComboBoxEx32").First();
            var cmbName = GetAllChildrenWindowHandles(cmbParent, "ComboBox").First();
            var txbName = GetAllChildrenWindowHandles(cmbName, "Edit").First();
            User32.PostMessage(txbName, Win32.WM_KEYDOWN, (IntPtr)Win32.VK_1_KEY, IntPtr.Zero);

            var saveBtn = GetAllChildrenWindowHandles(hSave, "Button")[1];
            User32.SetForegroundWindow(saveBtn);
            User32.SetFocus(saveBtn);

            IntPtr hExec = IntPtr.Zero;
            while (hExec == IntPtr.Zero)
            {
                User32.PostMessage(saveBtn, Win32.WM_KEYDOWN, (IntPtr)Win32.VK_RETURN, IntPtr.Zero);
                hExec = User32.FindWindow("#32770", "Выполнение ...");
            }
            while (hExec != IntPtr.Zero)
            {
                hExec = User32.FindWindow("#32770", "Выполнение ...");
            }
            scadProcess.Kill();
            Process[] excelPcs = null;
            while (excelPcs == null || !excelPcs.Any())
            {
                excelPcs = Process.GetProcessesByName("EXCEL");
            }
            excelPcs.ToList().ForEach(p => p.Kill());
        }

        static List<IntPtr> GetAllChildrenWindowHandles(IntPtr hParent, string className, string windowName = null, int maxCount = 50)
        {
            List<IntPtr> result = new List<IntPtr>();
            int ct = 0;
            IntPtr prevChild = IntPtr.Zero;
            IntPtr currChild = IntPtr.Zero;
            while (true && ct < maxCount)
            {
                currChild = User32.FindWindowEx(hParent, prevChild, className, windowName);
                if (currChild == IntPtr.Zero) break;
                result.Add(currChild);
                prevChild = currChild;
                ++ct;
            }
            return result;
        }
    }
}
