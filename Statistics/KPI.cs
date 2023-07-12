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
        public event EventHandler<string> CurrentRowChanged;

        //Автозагрузка файлов-----------------------------------------------------------------------------------------------

        public string AutoImport(string sourceFolder, string keyFolder, string month, string extention, string fullPath)
        {
            string nextFolder = "";
            DirectoryInfo dir = new DirectoryInfo(sourceFolder);
            foreach (DirectoryInfo directory in dir.GetDirectories())
            {
                if (directory.Name.ToString().ToUpper().Contains(keyFolder) && directory.Name.ToString().ToUpper().Contains(month.ToUpper()))
                {
                    nextFolder = directory.Name;

                    break;
                }
            }

            dir = new DirectoryInfo(sourceFolder + nextFolder);
            foreach (FileInfo files in dir.GetFiles())
            {
                if (files.Name.ToUpper().Contains(month.ToUpper()) && files.Name.ToUpper().Contains(extention) && !files.Name.Contains("$") )
                {
                    Path = files.FullName;
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

        public string CalculateMonth(Statistics statistics)
        {
            statistics.Beshoz = 0;
            statistics.Abtb = 0;
            statistics.TrudRas = 0;
            string error = "";        

            if (!OpenConnection())
            {
                isWorking = false;
                return "Не удалось открыть файл \"KPI\". Возможно, он сейчас используется.";
            }

            for (int i = 2; i < 50; i++)
            {
                if (!isWorking)
                {
                    CloseConnection();
                    return "cancel";
                }

                CurrentRowChanged?.Invoke(this, "Отладка: KPI. Подсчёт. Строка = " + i);

                if (Contains(i, 2, "ИТОГО"))
                {
                    statistics.Beshoz = Convert.ToInt32(ToString(i, 4));
                    statistics.Abtb = Convert.ToInt32(ToString(i, 5));
                    statistics.TrudRas = Convert.ToInt32(ToString(i, 3));

                    break;
                }
            }

            CloseConnection();
            return error;
        }            
    }
}
