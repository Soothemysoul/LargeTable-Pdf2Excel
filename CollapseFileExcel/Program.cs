using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using App = Microsoft.Office.Interop.Excel.Application;

namespace CollapseFileExcel
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {

                string executableFile = Assembly.GetExecutingAssembly().Location;
                DirectoryInfo directoryInfo = new DirectoryInfo(executableFile);

                Console.WriteLine("Введите путь до папки с файлами: \n");
                string folderPath = Console.ReadLine().Trim('\"');
                Console.WriteLine();

                if (folderPath == "")
                {
                    Console.WriteLine("Ничего не введено!\n");
                }

                string[] files = Directory.GetFiles(folderPath, "*.xlsx");

                List<string>sortedPath = files.ToList();

                sortedPath = sortedPath.OrderBy(p => int.Parse(Path.GetFileNameWithoutExtension(p))).ToList();

                if (files.Count() == 0)
                {
                    Console.WriteLine("Папка пуста!\n");
                }
                else
                {
                    try
                    {
                        App app = new App();

                        Workbook wb = app.Workbooks.Add();

                        Worksheet ws = wb.Worksheets[1];

                        foreach (string file in sortedPath)
                        {
                            Workbook partWb = app.Workbooks.Open(file);

                            Console.WriteLine($"Обработка файла '{Path.GetFileName(file)}'");
                            Console.WriteLine();

                            AddedRange(partWb, ws);

                            partWb.Close(false);

                            File.Delete(file);

                        }

                        wb.SaveAs(folderPath + "\\collapsed.xlsx");
                        wb.Close();
                        

                    }
                    catch { }


                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine();
            }

            Console.WriteLine("Готово!\nНажмите любую клавишу для выхода.");
            Console.ReadKey();

        }

        private static void AddedRange(Workbook partWb, Worksheet ws)
        {
            long lastRow = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            Range targetRange = ws.Range["A" + (lastRow+1).ToString()];

            Worksheet sourceWs = partWb.Worksheets[1];
            long sourceLastRow = sourceWs.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row;

            for (long i = sourceLastRow; i > 0; i--)
            {

                bool isMerged;

                if (sourceWs.Cells[i, 1].MergeCells &&
                    sourceWs.Cells[i, 5].MergeCells &&
                    sourceWs.Cells[i, 9].MergeCells &&
                    sourceWs.Cells[i, 13].MergeCells &&
                    sourceWs.Cells[i, 17].MergeCells)
                {
                    isMerged = true;
                }
                else
                {
                    isMerged = false;
                }
                

                if (isMerged)
                {
                    Range mergedRange = sourceWs.Cells[i, 1].MergeArea;

                    mergedRange.ClearContents();

                    // получаем номера строк, которые попадают в объединенную ячейку
                    int startRow = mergedRange.Row;
                    int endRow = startRow + mergedRange.Rows.Count - 1;

                    // удаляем строки
                    sourceWs.Rows[startRow+":"+endRow].ClearContents();
                    sourceWs.Rows[endRow].Delete();

                }

            }

            Range usedSourceRange = sourceWs.UsedRange;

            usedSourceRange.Copy(targetRange);


        }
    }
}
