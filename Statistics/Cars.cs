﻿using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using System.Collections.Generic;

namespace Statistics
{
    public class Cars : DocumentsForStatistics
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
                if (files.Name.ToString().ToUpper().Contains(keyFile1) && files.Name.ToString().Contains(keyFile2) 
                    && !files.Name.ToString().Contains("$"))
                {
                    path = sourceFolder + nextFolder + @"\" + files.Name;
                    fullPath = FillingPaths(fullPath, path);

                    break;
                }
            }

            return fullPath;
        }

        //Возвращает полный путь файла для Drag&Drop------------------------------------------------------------------------

        public string DragDrop(string key1, string key2, string file, string fullPath)
        {
            if (file.ToUpper().Contains(key1) || file.ToUpper().Contains(key2))
            {
                path = file;
                fullPath = FillingPaths(fullPath, file);
            }

            return fullPath;
        }

        //Подсчёт статистики за месяц---------------------------------------------------------------------------------------

        public string CalculateMonth(Statistics statistics, string monthNum, bool cancel)
        {
            statistics.carsOneTime = 0;
            statistics.carsPermanent = 0;
            string error = "";

            if (cancel == true) return error = "cancel";

            Excel.Application app = new Excel.Application();
            Excel.Workbook ObjWorkBook = null;
            try
            {
                ObjWorkBook = app.Workbooks.Open(path);
            }
            catch
            {
                return error = "Не удалось открыть файл '30м'. Возможно, он сейчас используется.";
            }
            Excel.Worksheet ObjWorkSheet = null;
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.get_Item(1);       

            for (int i = 2; i < 1000; i++)
            {                            
                if (getMonthNum(ObjWorkSheet, i, 1) == monthNum)
                {
                    CalculateBaseMonth(ObjWorkSheet, i, statistics);
                }
                else if (ToString(ObjWorkSheet, i, 1) == "" && ToString(ObjWorkSheet, i + 1, 1) == "" && ToString(ObjWorkSheet, i + 2, 1) == "") break;
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
            statistics.listDate = new List<string>();
            statistics.listCarsOneTime = new List<int>();
            statistics.listCarsPermanent = new List<int>();
            string error = "";

            if (cancel == true) return error = "cancel";

            Excel.Application app = new Excel.Application();
            Excel.Workbook ObjWorkBook = null;
            try
            {
                ObjWorkBook = app.Workbooks.Open(path);
            }
            catch
            {
                return error = "Не удалось открыть файл '30м'. Возможно, он сейчас используется.";
            }
            Excel.Worksheet ObjWorkSheet = null;
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.get_Item(1);

            for (int i = 2; i < 1000; i++)
            {
                //if (getMonthNum(ObjWorkSheet, i, 1) == monthNum && Contains(ObjWorkSheet, i, 7, "ПОСТ")) statistics.carsPermanent++;
                //else if (getMonthNum(ObjWorkSheet, i, 1) == monthNum && Contains(ObjWorkSheet, i, 7, "РАЗ")) statistics.carsOneTime++;
                //else if (ToString(ObjWorkSheet, i, 1) == "" && ToString(ObjWorkSheet, i + 1, 1) == "") break;

                if ( (monthNum1 != monthNum2 &&
                      ((getYear(ObjWorkSheet, i, 1) == year1 &&
                      getMonthNum(ObjWorkSheet, i, 1) == monthNum1 && Convert.ToInt32(getDay(ObjWorkSheet, i, 1)) >= Convert.ToInt32(day1)) ||
                      (getYear(ObjWorkSheet, i, 1) == year2 &&
                      getMonthNum(ObjWorkSheet, i, 1) == monthNum2 && Convert.ToInt32(getDay(ObjWorkSheet, i, 1)) <= Convert.ToInt32(day2))))
                      ||
                      (monthNum1 == monthNum2) &&
                      getMonthNum(ObjWorkSheet, i, 1) == monthNum1 &&
                      Convert.ToInt32(getDay(ObjWorkSheet, i, 1)) >= Convert.ToInt32(day1) &&
                      Convert.ToInt32(getDay(ObjWorkSheet, i, 1)) <= Convert.ToInt32(day2) )

                {
                    CalculateBaseWeek(ObjWorkSheet, i, statistics);
                }
                else if (ToString(ObjWorkSheet, i, 1) == "" && ToString(ObjWorkSheet, i + 1, 1) == "" && ToString(ObjWorkSheet, i + 2, 1) == "") break;
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

        //Общий алгоритм подсчёта статистики за месяц-----------------------------------------------------------------------

        private void CalculateBaseMonth(Excel.Worksheet ObjWorkSheet, int i, Statistics statistics)
        {
            if (Contains(ObjWorkSheet, i, 7, "ПОСТ")) statistics.carsPermanent++;
            else if (Contains(ObjWorkSheet, i, 7, "РАЗ")) statistics.carsOneTime++;
        }

        //Общий алгоритм подсчёта статистики за неделю----------------------------------------------------------------------

        private void CalculateBaseWeek(Excel.Worksheet ObjWorkSheet, int i, Statistics statistics)
        {
            if ((statistics.listDate.Count > 0 && getDate(ObjWorkSheet, i, 1) != statistics.listDate.Last())
                || statistics.listDate.Count == 0)
            {
                statistics.listDate.Add(getDate(ObjWorkSheet, i, 1));
                statistics.listCarsPermanent.Add(0);
                statistics.listCarsOneTime.Add(0);
            }

            if (Contains(ObjWorkSheet, i, 7, "ПОСТ"))
                statistics.listCarsPermanent[statistics.listCarsPermanent.Count - 1] = statistics.listCarsPermanent.Last() + 1;
            else if (Contains(ObjWorkSheet, i, 7, "РАЗ"))
                statistics.listCarsOneTime[statistics.listCarsOneTime.Count - 1] = statistics.listCarsOneTime.Last() + 1;
        }

    }
}
