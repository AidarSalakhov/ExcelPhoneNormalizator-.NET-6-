using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Diagnostics;

namespace ExcelPhoneNormalizator
{
    internal class Program
    {
        public static List<OpenedFiles> listOpenedFiles = new List<OpenedFiles>();

        public static string[] files = new string[0];

        public static int index = 0;

        static void Main(string[] args)
        {

            try
            {
                Console.WriteLine("Запуск программы, подождите...");

                ExcelOperations helper = new ExcelOperations();

                files = Directory.GetFiles(Environment.CurrentDirectory, "*.csv");

                Stopwatch stopWatch = new Stopwatch();

                stopWatch.Start();

                for (int i = 0; i < files.Length; i++)
                {
                    OpenedFiles openedCsv = new OpenedFiles();

                    openedCsv._index = i + 1;

                    Program.index = i + 1;

                    openedCsv._fileName = Path.GetFileName(files[i]);

                    if (helper.OpenCSV(Path.Combine(Environment.CurrentDirectory, files[i])))
                    {
                        helper.SaveAsTXT(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                        helper.Dispose();

                        helper.OpenCSV(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                        openedCsv._projectName = helper.GetProjectName();

                        helper.DeleteColumn("B1:X1");

                        helper.RemoveDuplicatesFromColumn("A");

                        helper.Normalize();

                        helper.RemoveDuplicatesFromColumn("B");

                        helper.DeleteColumn("A1");

                        helper.DeleteRow("A1");

                        helper.DeleteColumn("B1");

                        helper.SetColumnWidth(1, 18);

                        helper.SaveAsXLSX(Path.Combine(Environment.CurrentDirectory, $"{files[i]}-{helper.GetLastRow()}.xlsx"));

                        openedCsv._leadsCont = helper.GetLastRow();

                        listOpenedFiles.Add(openedCsv);

                        File.Delete(Path.Combine(Environment.CurrentDirectory, "leads.txt"));

                        File.Delete(Path.Combine(Environment.CurrentDirectory, files[i]));

                        helper.Dispose();
                    }
                }

                stopWatch.Stop();

                TimeSpan ts = stopWatch.Elapsed;

                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}", ts.Hours, ts.Minutes, ts.Seconds);

                Console.Clear();

                Console.WriteLine($"Обработка прошла успешно. Затраченное время: {elapsedTime}\n");

                for (int i = 0; i < listOpenedFiles.Count; i++)
                {
                    OpenedFiles openedCsvNew = new OpenedFiles();
                    openedCsvNew = listOpenedFiles[i];
                    openedCsvNew.Print();

                    File.AppendAllText(Path.Combine(Environment.CurrentDirectory, "log.txt"), $"Файл: {openedCsvNew._fileName}\n");
                    File.AppendAllText(Path.Combine(Environment.CurrentDirectory, "log.txt"), $"Название проекта: {openedCsvNew._projectName}\n");
                    File.AppendAllText(Path.Combine(Environment.CurrentDirectory, "log.txt"), $"Количество чистых лидов: {openedCsvNew._leadsCont}\n\n");
                }

                Console.ReadLine();
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }
    }
}

