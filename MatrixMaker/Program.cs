using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace MatrixMaker
{
    class Program
    {
        static void Main(string[] args)
        {
            List<double> z0 = new List<double>();

            Settings.Initialize();

            var files = Directory.EnumerateFiles(Settings.SourceDir).OrderBy(name => Convert.ToInt32(Path.GetFileNameWithoutExtension(name)));
            List<List<double>> result = new List<List<double>>();

            files.ToList().ForEach(file =>
            {
                Console.WriteLine($@"Begin reading {file}");

                Application xlApp = new Application();
                Workbook book = xlApp.Workbooks.Open(file);
                Worksheet sheet = book.Worksheets[1];
                Range srcRange = sheet.UsedRange;

                if (file == files.First())
                {
                    for (int i = 0; i < Settings.Nodes.Count; i++)
                    {
                        z0.Add(srcRange[Settings.FirstRow + Settings.Nodes[i] - 1, Settings.ZColumn].Value);
                    }
                }
                else
                {
                    List<double> disps = new List<double>();
                    double size = Convert.ToInt32(Path.GetFileNameWithoutExtension(file)) * Settings.MetersInCell;
                    disps.Add(size * size);
                    for (int i = 0; i < Settings.Nodes.Count; i++)
                    {
                        disps.Add(Math.Abs(srcRange[Settings.FirstRow + Settings.Nodes[i] - 1, Settings.ZColumn].Value - z0[i]));
                    }
                    result.Add(disps);
                }

                Marshal.ReleaseComObject(srcRange);
                Marshal.ReleaseComObject(sheet);
                book.Close();
                Marshal.ReleaseComObject(book);

                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                Console.WriteLine($@"End reading {file}");
            });

            StreamWriter resFile = new StreamWriter(File.Create(Settings.ResFileName));
            resFile.Write(";");
            resFile.WriteLine(string.Join(";", Settings.Distances));
            foreach (var str in result)
            {
                resFile.WriteLine(string.Join(";", str));
            }
            resFile.Close();
        }
    }
}
