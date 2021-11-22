using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Statistics
{
    public class KPI : DocumentsForStatistics
    {

        //Автозагрузка файлов-----------------------------------------------------------------------------------------------

        public string AutoImport(string sourceFolder, string keyFolder1, string month, string keyFile1, string keyFile2, string fullPath)
        {
            string nextFolder = "";
            DirectoryInfo dir = new DirectoryInfo(sourceFolder);
            foreach (DirectoryInfo directory in dir.GetDirectories())
            {
                if (directory.Name.ToString().ToUpper().Contains(keyFolder1) && directory.Name.ToString().ToUpper().Contains(month.ToUpper()))
                {
                    nextFolder = directory.Name;

                    break;
                }
            }

            dir = new DirectoryInfo(sourceFolder + nextFolder);
            foreach (FileInfo files in dir.GetFiles())
            {
                if ( (files.Name.ToString().ToUpper().Contains(keyFile1) || files.Name.ToString().ToUpper().Contains(month.ToUpper()))
                    && !files.Name.ToString().Contains("$") )
                {
                    Path = sourceFolder + nextFolder + @"\" + files.Name;
                    fullPath = FillingPaths(fullPath, Path);

                    break;
                }
            }

            return fullPath;
        }

        //Возвращает полный путь файла для Drag&Drop------------------------------------------------------------------------

        public string DragDrop(string key, List<string> listMonth, string file, string fullPath)
        {
            if (file.ToUpper().Contains(key))
            {
                Path = file;
                fullPath = FillingPaths(fullPath, file);
            }
            else
            {
                for (int i = 0; i < listMonth.Count; i++)
                {
                    if (file.Contains(listMonth[i]))
                    {
                        Path = file;
                        FileName = GetFileName(file);
                        fullPath = FillingPaths(fullPath, file);

                        break;
                    }
                }
            }

            return fullPath;
        }

        //Подсчёт статистики за месяц---------------------------------------------------------------------------------------

        public string CalculateMonth(Statistics statistics, bool cancel)
        {
            statistics.Beshoz = 0;
            statistics.Abtb = 0;
            statistics.TrudRas = 0;
            string error = "";

            if (cancel == true) return error = "cancel";

            Excel.Application app = new Excel.Application();
            Excel.Workbook ObjWorkBook = null;
            try
            {
                ObjWorkBook = app.Workbooks.Open(Path);
            }
            catch
            {
                return error = "Не удалось открыть файл 'KPI'. Возможно, он сейчас используется.";
            }
            Excel.Worksheet ObjWorkSheet = null;
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.get_Item(1);

            for (int i = 2; i < 50; i++)
            {
                if (Contains(ObjWorkSheet, i, 2, "ИТОГО"))
                {
                    statistics.Beshoz = Convert.ToInt32(ToString(ObjWorkSheet, i, 4));
                    statistics.Abtb = Convert.ToInt32(ToString(ObjWorkSheet, i, 5));
                    statistics.TrudRas = Convert.ToInt32(ToString(ObjWorkSheet, i, 3));

                    break;
                }
            }

            try
            {
                ObjWorkBook.Close();
                app.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), ex.Source);
            }

            return error;
        }      
       
    }
}
