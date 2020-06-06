using System;

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

        /// <summary>
        /// Путь к папке SCADWORK
        /// </summary>
        public static string ScadWork { get; set; }

        /// <summary>
        /// Считывает настройки из .ini
        /// </summary>
        public static void Initialize()
        {
            Ini settings = new Ini(SettingsFileName);

            PathToScad = settings.GetValue("scad_path", CalcuationsSection);
            SourceFolder = settings.GetValue("workdir", SchemeSection);
            ResultFolder = settings.GetValue("resdir", CalcuationsSection);
            LongDelayTime = Convert.ToInt32(settings.GetValue("long_delay", CalcuationsSection));
            ShortDelayTime = Convert.ToInt32(settings.GetValue("short_delay", CalcuationsSection));
            DefaultX = Convert.ToInt32(settings.GetValue("default_x", CalcuationsSection));
            DefaultY = Convert.ToInt32(settings.GetValue("default_y", CalcuationsSection));
            MaxCorrectCoord = Convert.ToInt32(settings.GetValue("max_correct_coord", CalcuationsSection));
            ScadWork = settings.GetValue("scadwork", CalcuationsSection);
        }

        private const string SettingsFileName = "settings.ini";
        private const string CalcuationsSection = "calculations";
        private const string SchemeSection = "scheme";
    }
}
