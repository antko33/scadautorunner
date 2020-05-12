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
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
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
            Input.LongDelay();
            IntPtr scadMainWindow = scadProcess.MainWindowHandle;
            if (!Win32API.SetForegroundWindow(scadMainWindow))
            {
                return;
            }

            int x = 130;
            int y = 420;

            var treeView = new TreeView();
            var linearCalcItem = treeView.GetLinearCalcuationsItem();
            var linearCalcPoint = treeView.GetItemPoint(linearCalcItem);
            IntPtr hParams = IntPtr.Zero;
            while (hParams == IntPtr.Zero)
            {
                Input.LeftMouseButtonClick(linearCalcPoint.X, linearCalcPoint.Y);
                hParams = Win32API.FindWindow("#32770", "Параметры расчета");
                Input.LongDelay();
            }
            Win32API.SetForegroundWindow(hParams);
            Input.KeyboardKeyPressed(Win32API.InputConstants.VK_RETURN, hParams);

            Process solverProcess = null;
            while (solverProcess == null)
            {
                var procs = Process.GetProcessesByName("SprSolverExe");
                if (procs.Any())
                {
                    solverProcess = procs.First();
                }
                Input.LongDelay();
            }
            solverProcess.WaitForExit();
            Input.LongDelay();
            scadProcess.Kill();

            scadProcess = Process.Start(scadInfo);
            while (scadProcess.MainWindowHandle == IntPtr.Zero) { }
            Input.LongDelay();
            scadMainWindow = scadProcess.MainWindowHandle;
            if (!Win32API.SetForegroundWindow(scadMainWindow))
            {
                return;
            }

            x = 130;
            y = 420;
            Win32API.SetForegroundWindow(scadMainWindow);

            treeView = new TreeView();
            var docResultItem = treeView.GetDocResultItem();
            var docResultPoint = treeView.GetItemPoint(docResultItem);

            hParams = IntPtr.Zero;
            while (hParams == IntPtr.Zero)
            {
                Input.LeftMouseButtonClick(docResultPoint.X, docResultPoint.Y);
                hParams = Win32API.FindWindow("#32770", "Результаты расчета");
                Input.LongDelay();
            }

            IntPtr hParent = hParams;
            hParams = Win32API.FindWindowEx(hParams, IntPtr.Zero, "#32770", "Вывод результатов");
            var resSpreads = GetAllChildrenWindowHandles(hParams, "fpSpread80");
            var resButtons = GetAllChildrenWindowHandles(hParams, "Button");

            Win32API.SetForegroundWindow(resSpreads[1]);
            Win32API.PostMessage(resSpreads[1], Win32API.InputConstants.WM_KEYDOWN, (IntPtr)Win32API.InputConstants.VK_DOWN, IntPtr.Zero);
            Win32API.PostMessage(resSpreads[1], Win32API.InputConstants.WM_KEYDOWN, (IntPtr)Win32API.InputConstants.VK_DOWN, IntPtr.Zero);

            Win32API.GetWindowRect(resButtons[2], out RECT btnRect);
            x = (btnRect.Left + btnRect.Right) / 2;
            y = (btnRect.Top + btnRect.Bottom) / 2;
            IntPtr hwnd = Win32API.WindowFromPoint(new System.Drawing.Point(x, y));
            Win32API.SetForegroundWindow(hParams);
            IntPtr hChild = IntPtr.Zero;
            while (hChild == IntPtr.Zero)
            {
                Input.LeftMouseButtonClick(x, y, hParams);
                Input.KeyboardKeyPressed(Win32API.InputConstants.VK_SPACE, hwnd);
                Input.LeftMouseButtonClick(465, 290, hParams); // TODO
                hChild = Win32API.FindWindow("#32770", "Величины перемещений");
                Input.ShortDelay();
            }

            var okBtn = GetAllChildrenWindowHandles(hChild, "Button").First();
            Win32API.SetForegroundWindow(okBtn);
            Win32API.PostMessage(hChild, Win32API.InputConstants.WM_KEYDOWN, (IntPtr)Win32API.InputConstants.VK_RETURN, IntPtr.Zero);

            var createXlsBtn = GetAllChildrenWindowHandles(hParent, "Button").Last();
            Win32API.SetForegroundWindow(createXlsBtn);
            Win32API.GetWindowRect(createXlsBtn, out btnRect);
            x = (btnRect.Left + btnRect.Right) / 2;
            y = (btnRect.Top + btnRect.Bottom) / 2;
            Win32API.ShowWindow(scadMainWindow, Win32API.ShowWindowCommands.Minimize);
            Win32API.ShowWindow(scadMainWindow, Win32API.ShowWindowCommands.Maximize);

            IntPtr hSave = IntPtr.Zero;
            hwnd = Win32API.WindowFromPoint(new System.Drawing.Point(x, y));
            while (hSave == IntPtr.Zero)
            {
                Win32API.SetFocus(hParent);
                Win32API.SetForegroundWindow(createXlsBtn);
                Input.LeftMouseButtonClick(x, y, hwnd);
                Input.KeyboardKeyPressed(Win32API.InputConstants.VK_SPACE, hwnd);
                hSave = Win32API.FindWindow("#32770", "Сохранение");
                Input.ShortDelay();
            }

            Input.ShortDelay();

            var cmbParent = GetAllChildrenWindowHandles(hSave, "ComboBoxEx32").First();
            var cmbName = GetAllChildrenWindowHandles(cmbParent, "ComboBox").First();
            var txbName = GetAllChildrenWindowHandles(cmbName, "Edit").First();
            Win32API.PostMessage(txbName, Win32API.InputConstants.WM_KEYDOWN, (IntPtr)Win32API.InputConstants.VK_1_KEY, IntPtr.Zero);

            var saveBtn = GetAllChildrenWindowHandles(hSave, "Button")[1];
            Win32API.SetForegroundWindow(saveBtn);
            Win32API.SetFocus(saveBtn);

            IntPtr hExec = IntPtr.Zero;
            while (hExec == IntPtr.Zero)
            {
                Win32API.PostMessage(saveBtn, Win32API.InputConstants.WM_KEYDOWN, (IntPtr)Win32API.InputConstants.VK_RETURN, IntPtr.Zero);
                hExec = Win32API.FindWindow("#32770", "Выполнение ...");
            }
            while (hExec != IntPtr.Zero)
            {
                hExec = Win32API.FindWindow("#32770", "Выполнение ...");
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
                currChild = Win32API.FindWindowEx(hParent, prevChild, className, windowName);
                if (currChild == IntPtr.Zero) break;
                result.Add(currChild);
                prevChild = currChild;
                ++ct;
            }
            return result;
        }
    }
}
