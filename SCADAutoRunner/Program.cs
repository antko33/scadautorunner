﻿using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Security;
using System.Threading;
using System.Collections.Generic;
using System.IO;

namespace SCADAutoRunner
{
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
            IntPtr hTreeView = User32.WindowFromPoint(new System.Drawing.Point(x, y));
            User32.GetWindowThreadProcessId(hTreeView, out int pid);
            IntPtr process = User32.OpenProcess(
                User32.ProcessAccessFlags.VirtualMemoryOperation |
                User32.ProcessAccessFlags.VirtualMemoryRead |
                User32.ProcessAccessFlags.VirtualMemoryWrite |
                User32.ProcessAccessFlags.QueryInformation,
                false, pid);
            int tviSize = Marshal.SizeOf(typeof(TVItem));
            int textSize = 255;
            bool isUnicode = User32.IsWindowUnicode(hTreeView);
            if (isUnicode)
            {
                textSize *= 2;
            }
            IntPtr tviPtr = User32.VirtualAllocEx(process, IntPtr.Zero, tviSize + textSize, User32.AllocationType.Commit, User32.MemoryProtection.ReadWrite);

            IntPtr textPtr = IntPtr.Add(tviPtr, tviSize);

            IntPtr hItems = User32.SendMessage(hTreeView, User32.TVM.TVM_GETCOUNT, IntPtr.Zero, IntPtr.Zero);
            int citems = hItems.ToInt32();

            int w = y << 16 | x;

            IntPtr hParams = IntPtr.Zero;
            while (hParams == IntPtr.Zero)
            {
                User32.PostMessage(hTreeView, Win32.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)w);
                User32.PostMessage(hTreeView, Win32.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)w);
                hParams = User32.FindWindow("#32770", "Параметры расчета");
                Thread.Sleep(2000);
            }
            User32.SetForegroundWindow(hParams);
            User32.PostMessage(hParams, Win32.WM_KEYDOWN, (IntPtr)Win32.VK_RETURN, (IntPtr)0x0);

            Process solverProcess = null;
            while (solverProcess == null)
            {
                var procs = Process.GetProcessesByName("SprSolverExe");
                if (procs.Any())
                {
                    solverProcess = procs.First();
                }
                Thread.Sleep(2000);
            }
            solverProcess.WaitForExit();
            Thread.Sleep(2000);
            scadProcess.Kill();

            scadProcess = Process.Start(scadInfo);
            while (scadProcess.MainWindowHandle == IntPtr.Zero) { }
            Thread.Sleep(2000);
            scadMainWindow = scadProcess.MainWindowHandle;
            if (!User32.SetForegroundWindow(scadMainWindow))
            {
                return;
            }
            x = 145;
            y = 785;
            IntPtr hWnd = User32.WindowFromPoint(new System.Drawing.Point(x, y));
            w = y << 16 | x;

            hParams = IntPtr.Zero;
            while (hParams == IntPtr.Zero)
            {
                //User32.PostMessage(hWnd, Win32.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)w);
                //User32.PostMessage(hWnd, Win32.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)w);
                hParams = User32.FindWindow("#32770", "Результаты расчета");
                Thread.Sleep(2000);
            }
        }

        static List<IntPtr> GetAllChildrenWindowHandles(IntPtr hParent, string className, string windowName, int maxCount = 50)
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
