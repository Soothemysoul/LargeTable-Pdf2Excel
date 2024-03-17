using Acrobat;

using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using Path = System.IO.Path;


namespace PdfToExcel
{
    internal class Program
    {
        public static AcroApp AcrobatApp { get; private set; }

        static void Main(string[] args)
        {
            try
            {
                AcrobatApp = new AcroApp();


                string executableFile = Assembly.GetExecutingAssembly().Location;
                DirectoryInfo directoryInfo = new DirectoryInfo(executableFile);

                Console.WriteLine("Введите путь до pdf файла для экспорта: \n");
                string filePath = Console.ReadLine().Trim('\"');
                Console.WriteLine();

                if (filePath == "")
                {
                    Console.WriteLine("Ничего не введено!\n");
                }

                bool fileExist = File.Exists(filePath);
                if (!fileExist)
                {
                    Console.WriteLine("Файла не существует!\n");
                }
                else
                {
                    try
                    {

                        Console.WriteLine($"Обработка файла '{filePath}'");
                        Console.WriteLine();
                        Export(filePath);


                    }
                    catch { }


                }

                AcrobatApp.Maximize(1);
                AcrobatApp.Exit();


            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine();
            }

            Console.WriteLine("Готово!\nНажмите любую клавишу для выхода.");
            Console.ReadKey();

        }

        private static void Export(string pdfFile)
        {

            string tmpPdf = pdfFile.Replace(Path.GetFileName(pdfFile), $"tmp.pdf");

            int pagesRangeCount = 300;

            DirectoryInfo directoryInfo = new DirectoryInfo(pdfFile);

            CAcroAVDoc doc = new AcroAVDoc();

            int pagesCount = 0;

            if (doc.Open(pdfFile, null))
            {
                CAcroPDDoc pdfDoc = doc.GetPDDoc();
                pagesCount = pdfDoc.GetNumPages();

            }

            doc.Close(0);

            int hundreedsPage = pagesCount / pagesRangeCount;
            int balancePage = pagesCount % pagesRangeCount;

            for (int i = 0; i <= hundreedsPage; i++)
            {
                Console.WriteLine($"Итерация {i + 1}\n");

                string tmpExcel = pdfFile.Replace(Path.GetFileName(pdfFile), $"{i + 1}.xlsx");

                if (doc.Open(pdfFile, null))
                {

                    CAcroPDDoc pdfDoc = doc.GetPDDoc();

                    int[] pageRange = new int[2];
                    pageRange[0] = i * pagesRangeCount;

                    if (i * pagesRangeCount > pagesCount)
                    {
                        pageRange[1] = pagesCount - 1;
                    }
                    else
                    {
                        pageRange[1] = pagesRangeCount + i * pagesRangeCount;
                    }

                    Console.WriteLine($"Страниц до обрезки: {pdfDoc.GetNumPages()}\n");

                    if (pageRange[0] == 0)
                    {
                        pdfDoc.DeletePages(pageRange[1], pdfDoc.GetNumPages() - 1);
                    }
                    else if (pageRange[1] == pagesCount - 1)
                    {
                        pdfDoc.DeletePages(0, pageRange[0] - 1);
                    }
                    else
                    {
                        pdfDoc.DeletePages(pageRange[1], pdfDoc.GetNumPages() - 1);
                        pdfDoc.DeletePages(0, pageRange[0] - 1);
                    }

                    Console.WriteLine($"Страниц после обрезки: {pdfDoc.GetNumPages()}\n");

                    pdfDoc.Save(1, tmpPdf);
                    Object jsoObj = pdfDoc.GetJSObject();

                    Type jsType = jsoObj.GetType();
                    //have to use acrobat javascript api because, acrobat
                    object[] saveAsParam = { tmpExcel, "com.adobe.acrobat.xlsx", "", false, false};
                    jsType.InvokeMember("saveAs", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance, null, jsoObj, saveAsParam, CultureInfo.InvariantCulture);

                    pdfDoc.Close();


                }

                doc.Close(0);
                File.Delete(tmpPdf);
            }



        }
    }
}

