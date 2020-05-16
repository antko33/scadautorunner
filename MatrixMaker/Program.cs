using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace MatrixMaker
{
    class Program
    {
        static string dir = @"C:\Users\User\source\repos\ConsoleApp1\MatrixMaker\RES";
        static string resFile = @"C:\Users\User\source\repos\ConsoleApp1\MatrixMaker\res.xls";

        static List<int> nodes = new List<int>() { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
        const int FIRST_ROW = 9;
        const int Z_COLUMN = 3;
        const double METERS_IN_CELL = 0.2;

        static void Main(string[] args)
        {
            var files = Directory.EnumerateFiles(dir).OrderBy(name => Convert.ToInt32(Path.GetFileNameWithoutExtension(name)));
            List<List<double>> result = new List<List<double>>();

            files.ToList().ForEach(file =>
            {
                Console.WriteLine($@"Begin reading {file}");

                Application xlApp = new Application();
                Workbook book = xlApp.Workbooks.Open(file);
                Worksheet sheet = book.Worksheets[1];
                Range srcRange = sheet.UsedRange;

                List<double> disps = new List<double>();
                double size = Convert.ToInt32(Path.GetFileNameWithoutExtension(file)) * METERS_IN_CELL;
                disps.Add(size * size);
                for (int i = 0; i < nodes.Count; i++)
                {
                    disps.Add(srcRange[FIRST_ROW + nodes[i] - 1, Z_COLUMN].Value);
                }
                result.Add(disps);

                Marshal.ReleaseComObject(srcRange);
                Marshal.ReleaseComObject(sheet);
                book.Close();
                Marshal.ReleaseComObject(book);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                Console.WriteLine($@"End reading {file}");
            });
        }
    }
}
