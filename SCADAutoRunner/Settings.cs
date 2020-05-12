using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SCADAutoRunner
{
    /// <summary>
    /// Параметры проекта
    /// </summary>
    static class Settings
    {
        /// <summary>
        /// Полный путь к исполняемому файлу scadx.exe
        /// </summary>
        public static string PathToScad { get; set; }

        /// <summary>
        /// Каталог с проектами, для которых будет произведён расчёт
        /// </summary>
        public static string SourceFolder { get; set; }

        /// <summary>
        /// Каталог для хранения результатов расчётов
        /// </summary>
        public static string ResultFolder { get; set; }

        /// <summary>
        /// Продолжительность длительной задержки
        /// </summary>
        public static int LongDelayTime { get; set; }

        /// <summary>
        /// Продолжительность короткой задержки
        /// </summary>
        public static int ShortDelayTime { get; set; }

        /// <summary>
        /// X-координата точки, гарантированно принадлежащей TreeView в главном окне SCAD++
        /// </summary>
        public static int DefaultX { get; set; }

        /// <summary>
        /// Y-координата точки, гарантированно принадлежащей TreeView в главном окне SCAD++
        /// </summary>
        public static int DefaultY { get; set; }

        /// <summary>
        /// Максимальная величина координаты, при которой она считается корректной (метод Win32API.GetWindowRect)
        /// </summary>
        public static int MaxCorrectCoord { get; set; }
    }
}
