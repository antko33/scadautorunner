using System;
using System.Collections.Generic;
using System.Linq;

namespace MatrixMaker
{
    static class Settings
    {
        /// <summary>
        /// Каталог с проектами, для которых будут получены результаты расчёта
        /// </summary>
        public static string SourceDir { get; set; }

        /// <summary>
        /// Имя итогового файла
        /// </summary>
        public static string ResFileName { get; set; }

        /// <summary>
        /// Узлы, перемещения в которых рассматриваются
        /// </summary>
        public static List<int> Nodes { get; set; }

        /// <summary>
        /// Расстояния от центра провала до опорных узлов
        /// </summary>
        public static List<double> Distances { get; set; }

        /// <summary>
        /// Номер строки, содержащей перемещения в узле с номером 1
        /// </summary>
        public static int FirstRow { get; set; }

        /// <summary>
        /// Номер (начиная с 1) столбца с перемещениями по оси Z
        /// </summary>
        public static int ZColumn { get; set; }

        /// <summary>
        /// Расстояние (в метрах), которому соответствует длина стороны ячейки сетки в расчётной схеме
        /// </summary>
        public static float MetersInCell { get; set; }

        /// <summary>
        /// Десятичный разделитель
        /// </summary>
        public static string Delimeter { get; set; }

        /// <summary>
        /// Считывает настройки из .ini
        /// </summary>
        public static void Initialize()
        {
            Ini settings = new Ini(SettingsFileName);

            Delimeter = System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;

            SourceDir = settings.GetValue("resdir", CalculationsSection);
            ResFileName = settings.GetValue("filename", ResultSection);
            Nodes = settings.GetValue("nodes", ResultSection)
                .Split(',')
                .Select(item => Convert.ToInt32(item))
                .ToList();
            Nodes.Sort();
            Distances = settings.GetValue("distances", ResultSection)
                .Split(',')
                .Select(item => Convert.ToDouble(item.Replace(".", Delimeter).Replace(",", Delimeter)))
                .ToList();
            Distances.Sort();
            FirstRow = Convert.ToInt32(settings.GetValue("first_row", ResultSection));
            ZColumn = Convert.ToInt32(settings.GetValue("z_column", ResultSection));
            MetersInCell = float.Parse(settings.GetValue("meters_in_cell", ResultSection)
                .Replace(".", Delimeter).Replace(",", Delimeter));
        }

        private const string SettingsFileName = "settings.ini";
        private const string CalculationsSection = "calculations";
        private const string ResultSection = "result";
    }
}
