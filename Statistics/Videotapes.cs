using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Statistics
{
    public class Videotapes : DocumentsForStatistics
    {
        //Автозагрузка файлов-----------------------------------------------------------------------------------------------

        public string AutoImport(string sourceFolder, string keyFolder, string keyFile1, string keyFile2, string fullPath)
        {
            string nextFolder = "";
            DirectoryInfo dir = new DirectoryInfo(sourceFolder);
            foreach (DirectoryInfo directory in dir.GetDirectories())
            {
                if (directory.Name.ToString().ToUpper().Contains(keyFolder))
                {
                    nextFolder = directory.Name;

                    break;
                }
            }

            dir = new DirectoryInfo(sourceFolder + nextFolder);
            foreach (FileInfo files in dir.GetFiles())
            {
                if (files.Name.ToString().ToUpper().Contains(keyFile1) && (files.Name.ToString().Contains(keyFile2))
                      && !files.Name.ToString().Contains("$"))
                {
                    Path = sourceFolder + nextFolder + @"\" + files.Name;
                    fullPath = FillingPaths(fullPath, Path);

                    break;
                }
            }

            return fullPath;
        }

        //Подсчёт статистики за месяц---------------------------------------------------------------------------------------

        public string CalculateMonth(Statistics statistics, string monthNum, bool cancel)
        {
            statistics.Police = 0;
            statistics.AnotherOrg = 0;
            statistics.Airport = 0;
            statistics.Viewing = 0;
            statistics.Usb = 0;
            statistics.Dvd = 0;
            statistics.DiskN = 0;
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
                return error = "Не удалось открыть файл 'Запросы'. Возможно, он сейчас используется.";
            }
            Excel.Worksheet ObjWorkSheet = null;
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.get_Item(1);         

            for (int i = 2; i < 1000; i++)
            {
                if (getMonthNum(ObjWorkSheet, i, 6) == monthNum)
                {
                    CalculateBaseIssued(ObjWorkSheet, i, statistics);
                }
                if (getMonthNum(ObjWorkSheet, i, 5) == monthNum)
                {
                    CalculateBase(ObjWorkSheet, i, statistics);
                }                                              
                else if (ToString(ObjWorkSheet, i, 5) == "" && ToString(ObjWorkSheet, i + 1, 5) == "" && ToString(ObjWorkSheet, i + 2, 5) == "") break;
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

        //Подсчёт статистики за неделю--------------------------------------------------------------------------------------

        public string CalculateWeek(Statistics statistics, string day1, string day2, string monthNum1, string monthNum2, string year1, string year2, bool cancel)
        {
            statistics.Police = 0;
            statistics.AnotherOrg = 0;
            statistics.Airport = 0;
            statistics.Viewing = 0;
            statistics.Usb = 0;
            statistics.Dvd = 0;
            statistics.DiskN = 0;
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
                return error = "Не удалось открыть файл 'Запросы'. Возможно, он сейчас используется.";
            }
            Excel.Worksheet ObjWorkSheet = null;
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.get_Item(1);

            for (int i = 2; i < 1000; i++)
            {
               if ((monthNum1 != monthNum2 &&
                  ((getYear(ObjWorkSheet, i, 6) == year1 &&
                  getMonthNum(ObjWorkSheet, i, 6) == monthNum1 && Convert.ToInt32(getDay(ObjWorkSheet, i, 6)) >= Convert.ToInt32(day1)) ||
                  (getYear(ObjWorkSheet, i, 6) == year2 &&
                  getMonthNum(ObjWorkSheet, i, 6) == monthNum2 && Convert.ToInt32(getDay(ObjWorkSheet, i, 6)) <= Convert.ToInt32(day2))))
                  ||
                  (monthNum1 == monthNum2) &&
                  getMonthNum(ObjWorkSheet, i, 6) == monthNum1 &&
                  Convert.ToInt32(getDay(ObjWorkSheet, i, 6)) >= Convert.ToInt32(day1) &&
                  Convert.ToInt32(getDay(ObjWorkSheet, i, 6)) <= Convert.ToInt32(day2))
                {
                    CalculateBaseIssued(ObjWorkSheet, i, statistics);
                }
                if ( 
                    (
                     monthNum1 != monthNum2 &&
                     (                      
                       getYear(ObjWorkSheet, i, 5) == year1 &&
                       getMonthNum(ObjWorkSheet, i, 5) == monthNum1 && Convert.ToInt32(getDay(ObjWorkSheet, i, 5)) >= Convert.ToInt32(day1)                      
                       ||                      
                       getYear(ObjWorkSheet, i, 5) == year2 &&
                       getMonthNum(ObjWorkSheet, i, 5) == monthNum2 && Convert.ToInt32(getDay(ObjWorkSheet, i, 5)) <= Convert.ToInt32(day2)                     
                     )
                    )

                   ||

                   (monthNum1 == monthNum2 &&
                   getMonthNum(ObjWorkSheet, i, 5) == monthNum1 &&
                   Convert.ToInt32(getDay(ObjWorkSheet, i, 5)) >= Convert.ToInt32(day1) &&
                   Convert.ToInt32(getDay(ObjWorkSheet, i, 5)) <= Convert.ToInt32(day2)) )
                {
                    CalculateBase(ObjWorkSheet, i, statistics);
                }
                else if (ToString(ObjWorkSheet, i, 5) == "" && ToString(ObjWorkSheet, i + 1, 5) == "" && ToString(ObjWorkSheet, i + 2, 5) == "") break;
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

        //Общий алгоритм подсчёта статистики (подготовленные запросы и просмотр)--------------------------------------------

        private void CalculateBase(Excel.Worksheet ObjWorkSheet, int i, Statistics statistics)
        {
            if (Contains(ObjWorkSheet, i, 2, "З") && !Contains(ObjWorkSheet, i, 2, "ПОВТОР"))
            {
                if (Contains(ObjWorkSheet, i, 9, "МВД") || Contains(ObjWorkSheet, i, 9, "ЛОП")) statistics.Police++;
                else if (Contains(ObjWorkSheet, i, 9, "СТОРОН")) statistics.AnotherOrg++;
                else if (Contains(ObjWorkSheet, i, 9, "А/П")) statistics.Airport++;
            }
            if (Contains(ObjWorkSheet, i, 2, "П-") || Contains(ObjWorkSheet, i, 2, "П ")) statistics.Viewing++;
        }

        //Общий алгоритм подсчёта статистики (выдано) ----------------------------------------------------------------------

        private void CalculateBaseIssued(Excel.Worksheet ObjWorkSheet, int i, Statistics statistics)
        {
            if (!Contains(ObjWorkSheet, i, 7, "ПОВТОР") && !Contains(ObjWorkSheet, i, 8, "ПОВТОР"))
            {
                if (Contains(ObjWorkSheet, i, 7, "USB") || Contains(ObjWorkSheet, i, 8, "USB")) statistics.Usb++;
                else if (Contains(ObjWorkSheet, i, 7, "ДИСК N") || Contains(ObjWorkSheet, i, 8, "ДИСК N")) statistics.DiskN++;
                else if (!Contains(ObjWorkSheet, i, 7, "ДИСК N") && !Contains(ObjWorkSheet, i, 8, "ДИСК N") &&
                    (Contains(ObjWorkSheet, i, 7, "DVD") || Contains(ObjWorkSheet, i, 7, "CD") || Contains(ObjWorkSheet, i, 7, "ДИСК") ||
                     Contains(ObjWorkSheet, i, 8, "DVD") || Contains(ObjWorkSheet, i, 8, "CD") || Contains(ObjWorkSheet, i, 8, "ДИСК"))) statistics.Dvd++;
            }
        }


    }
}
