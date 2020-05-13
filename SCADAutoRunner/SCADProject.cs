using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SCADAutoRunner
{
    class SCADProject
    {
        public IntPtr mainWindow;

        private string pathToProject;
        private Process scadProcess;

        private readonly ProcessStartInfo scadInfo;

        public SCADProject(string path)
        {
            pathToProject = path;
            scadInfo = new ProcessStartInfo()
            {
                FileName = Settings.PathToScad,
                Arguments = pathToProject,
                WindowStyle = ProcessWindowStyle.Maximized
            };
        }

        /// <summary>
        /// Открытие проекта в SCAD
        /// </summary>
        public void Open()
        {
            scadProcess = Process.Start(scadInfo);
            while (scadProcess.MainWindowHandle == IntPtr.Zero)
            {
                Input.ShortDelay();
            }
            Input.LongDelay();
            mainWindow = scadProcess.MainWindowHandle;
            if (!Win32API.SetForegroundWindow(mainWindow))
            {
                throw new Exception("Не могу найти главное окно SCAD++");
            }
        }

        /// <summary>
        /// Немедленно закрывает окно SCAD
        /// </summary>
        public void Close()
        {
            scadProcess.Kill();
        }
    }
}
