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
            Settings.Initialize();

            var param = "\"C:\\Users\\User\\Desktop\\4_26 — копия.SPR\"";

            var scadProject = new SCADProject(param);
            scadProject.Open();

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
            scadProject.Close();

            scadProject.Open();

            x = 130;
            y = 420;

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
            var resSpreads = Input.GetAllChildrenWindowHandles(hParams, "fpSpread80");
            var resButtons = Input.GetAllChildrenWindowHandles(hParams, "Button");

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

            var okBtn = Input.GetAllChildrenWindowHandles(hChild, "Button").First();
            Win32API.SetForegroundWindow(okBtn);
            Win32API.PostMessage(hChild, Win32API.InputConstants.WM_KEYDOWN, (IntPtr)Win32API.InputConstants.VK_RETURN, IntPtr.Zero);

            var createXlsBtn = Input.GetAllChildrenWindowHandles(hParent, "Button").Last();
            Win32API.SetForegroundWindow(createXlsBtn);
            Win32API.GetWindowRect(createXlsBtn, out btnRect);
            x = (btnRect.Left + btnRect.Right) / 2;
            y = (btnRect.Top + btnRect.Bottom) / 2;
            Win32API.ShowWindow(scadProject.mainWindow, Win32API.ShowWindowCommands.Minimize);
            Win32API.ShowWindow(scadProject.mainWindow, Win32API.ShowWindowCommands.Maximize);

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

            var cmbParent = Input.GetAllChildrenWindowHandles(hSave, "ComboBoxEx32").First();
            var cmbName = Input.GetAllChildrenWindowHandles(cmbParent, "ComboBox").First();
            var txbName = Input.GetAllChildrenWindowHandles(cmbName, "Edit").First();
            Win32API.PostMessage(txbName, Win32API.InputConstants.WM_KEYDOWN, (IntPtr)Win32API.InputConstants.VK_1_KEY, IntPtr.Zero);

            var saveBtn = Input.GetAllChildrenWindowHandles(hSave, "Button")[1];
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
            scadProject.Close();

            Process[] excelPcs = null;
            while (excelPcs == null || !excelPcs.Any())
            {
                excelPcs = Process.GetProcessesByName("EXCEL");
            }
            excelPcs.ToList().ForEach(p => p.Kill());
        }
    }
}
