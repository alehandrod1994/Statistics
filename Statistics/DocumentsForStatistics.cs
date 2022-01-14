﻿using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Statistics
{
    public abstract class DocumentsForStatistics
    {
        public DocumentsForStatistics() => Path = "";

        public string Path { get; set; }
        public string FileName { get; set; }

        //Загрузка файлов---------------------------------------------------------------------------------------------------

        public string Import(string str)                                                           
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.FileName = "";
            ofd.Filter = "Документ Excel (*.xls; *xlsx) | *.xls; *.xlsx";
            ofd.Title = "Выберите файл";

            if (ofd.ShowDialog() != DialogResult.Cancel)
            {
                try
                {
                    Path = ofd.FileName;
                    FileName = ofd.SafeFileName;

                    //Заполнение путей в listboxPath
                    str = FillingPaths(str, Path);

                    return str;
                }
                catch
                {
                    MessageBox.Show("Недопустимый формат файла");
                    return "";
                }          
            }
            else return str;
        }

        //Возвращает полный путь файла для Drag&Drop------------------------------------------------------------------------

        public string DragDrop(string key, string file, string fullPath)
        {
            if (file.ToUpper().Contains(key))
            {
                Path = file;
                FileName = GetFileName(file);
                fullPath = FillingPaths(fullPath, file);
            }

            return fullPath;
        }

        //Возвращает имя файла----------------------------------------------------------------------------------------------

        protected string GetFileName(string path)
        {
            int index = 0;
            int lastIndex = -1;

            while (index > -1)
            {
                index = path.IndexOf((@"\"), index + 1);
                if (index > -1)
                {
                    lastIndex = index;
                }
            }

            return FileName = path.Remove(0, lastIndex + 1);
        }

        //Возвращает день---------------------------------------------------------------------------------------------------

        public string GetDay(Excel.Worksheet ObjWorkSheet, int i, int j)
        {
            string str;

            str = ToString(ObjWorkSheet, i, j);
            
            if (str != "")
            {
                str = str.Substring(0, 2);
            }

            return str;
        }

        //Возвращает номер месяца-------------------------------------------------------------------------------------------

        public string GetMonthNum(Excel.Worksheet ObjWorkSheet, int i, int j)
        {
            string str;

            str = ToString(ObjWorkSheet, i, j);
            
            if (str != "")
            {
                str = str.Substring(3, 2);
            }

            return str;
        }

        //Возвращает год----------------------------------------------------------------------------------------------------

        public string GetYear(Excel.Worksheet ObjWorkSheet, int i, int j)
        {
            string str;
            
            str = ToString(ObjWorkSheet, i, j);
            
            if (str != "")
            {
                str = str.Substring(6, 4);
                if (str.Contains(" "))
                {
                    str = str.Remove(2);
                    str = "20" + str;
                }
            }

            return str;
        }

        //Возвращает дату----------------------------------------------------------------------------------------------------

        public string GetDate(Excel.Worksheet ObjWorkSheet, int i, int j)
        {
            string str;

            str = ToString(ObjWorkSheet, i, j);

            if (str != "")
            {
                str = str.Substring(0, 10);
                if (str.Contains(" "))
                {
                    str = str.Remove(8);
                    str = str.Insert(6, "20");
                }
            }

            return str;
        }

        //Поиск подстроки в строке------------------------------------------------------------------------------------------

        public bool Contains(Excel.Worksheet ObjWorkSheet, int i, int j, string key)
        {
            string str;

            str = ToString(ObjWorkSheet, i, j);
            if (str.Contains(key)) return true;
            else return false;
        }

        //Превращает в строку-----------------------------------------------------------------------------------------------

        public string ToString(Excel.Worksheet ObjWorkSheet, int i, int j)
        {
            Excel.Range rString;
            string str;

            rString = ObjWorkSheet.Cells[i, j] as Excel.Range;
            str = (rString.Value != null) ? rString.Value.ToString() : "";      
            return str.ToUpper();            
        }

        //Заполнение путей в listboxPath------------------------------------------------------------------------------------

        public string FillingPaths(string str, string value)
        {
            int index = str.IndexOf("|");
            str = str.Remove(index);
            str += "| " + value;

            return str;
        }

    }
}
