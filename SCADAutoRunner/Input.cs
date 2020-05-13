using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SCADAutoRunner
{
    class Input
    {
        /// <summary>
        /// Клик левой клавишей мыши в точкe с заданными координатами
        /// </summary>
        /// <param name="hWnd">Handle окна, обрабатывающего клик. По умолчанию - окно в точке (x, y)</param>
        public static void LeftMouseButtonClick(int x, int y, IntPtr? hWnd = null)
        {
            if (!hWnd.HasValue)
            {
                hWnd = Win32API.WindowFromPoint(new Point(x, y));
            }

            int w = y << 16 | x;
            Win32API.PostMessage(hWnd.Value, Win32API.InputConstants.WM_LBUTTONDOWN, (IntPtr)0x1, (IntPtr)w);
            Win32API.PostMessage(hWnd.Value, Win32API.InputConstants.WM_LBUTTONUP, (IntPtr)0x1, (IntPtr)w);
        }

        /// <summary>
        /// Клик левой клавишей мыши в заданной точке
        /// </summary>
        /// <param name="hWnd">Handle окна, обрабатывающего клик. По умолчанию - окно в точке (x, y)</param>
        public static void LeftMouseButtonClick(Point p, IntPtr? hWnd = null)
        {
            LeftMouseButtonClick(p.X, p.Y, hWnd);
        }

        /// <summary>
        /// Имитация ввода с клавиатуры
        /// </summary>
        /// <param name="keyCode">Код нажимаемой клавиши</param>
        /// <param name="hWnd">Handle окна, обрабатывающего ввод</param>
        public static void KeyboardKeyPressed(int keyCode, IntPtr hWnd)
        {
            Win32API.PostMessage(hWnd, Win32API.InputConstants.WM_KEYDOWN, (IntPtr)keyCode, (IntPtr)0x0);
        }

        public static void LongDelay()
        {
            Thread.Sleep(Settings.LongDelayTime);
        }

        public static void ShortDelay()
        {
            Thread.Sleep(Settings.ShortDelayTime);
        }

        /// <summary>
        /// Список дочерних элементов заданного окна, принадлежащих указанному классу
        /// </summary>
        /// <param name="hParent">Handle родительского окно</param>
        /// <param name="className">Класс</param>
        /// <param name="windowName">Заголовок</param>
        /// <param name="maxCount">Максимальное количество искомых элементов</param>
        /// <returns>Список handle'ов найденных элементов</returns>
        public static List<IntPtr> GetAllChildrenWindowHandles(IntPtr hParent, string className, string windowName = null, int maxCount = 50)
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
