using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Statistics
{
    public abstract class DocumentsForStatistics
    {
        protected Excel.Application app;
        protected Excel.Workbook book;
        protected Excel.Worksheet sheet;

        public DocumentsForStatistics() => Path = "";

        public string Path { get; set; }
        public string FileName { get; set; }

        /// <summary>
		/// Открывает подключение к документу.
		/// </summary>
		/// <returns> True, если подключение прошло успешно; в противном случае - false. </returns>
        protected bool OpenConnection()
        {
            app = new Excel.Application();
            book = null;

            try
            {
                book = app.Workbooks.Open(Path);
            }
            catch
            {
                return false;
            }

            app.DisplayAlerts = false;
            sheet = book.Sheets[1];
            return true;
        }

        /// <summary>
        /// Закрывает подключение к документу.
        /// </summary>
        /// <returns> True, если закрытие подключения прошло успешно; в противном случае - false. </returns>
        protected bool CloseConnection()
        {
            try
            {
                book.Close();
                app.Quit();
                return true;
            }
            catch
            {
                return false;
            }
        }

        //Загрузка файлов---------------------------------------------------------------------------------------------------

        public string Import(string str)                                                           
        {
            OpenFileDialog ofd = new OpenFileDialog
            {
                FileName = "",
                Filter = "Документ Excel (*.xls; *xlsx) | *.xls; *.xlsx",
                Title = "Выберите файл"
            };

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

        public string GetDay(int i, int j)
        {
            string str;

            str = ToString(i, j);
            
            if (str != "")
            {
                str = str.Substring(0, 2);
            }

            return str;
        }

        //Возвращает номер месяца-------------------------------------------------------------------------------------------

        public string GetMonthNum(int i, int j)
        {
            string str;

            str = ToString(i, j);
            
            if (str != "")
            {
                str = str.Substring(3, 2);
            }

            return str;
        }

        //Возвращает год----------------------------------------------------------------------------------------------------

        public string GetYear(int i, int j)
        {
            string str;
            
            str = ToString(i, j);
            
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

        public string GetDate(int i, int j)
        {
            string str;

            str = ToString(i, j);

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

        public bool Contains(int i, int j, string key)
        {
            return ToString(i, j).ToUpper().Contains(key.ToUpper());
        }

        //Превращает в строку-----------------------------------------------------------------------------------------------

        public string ToString(int i, int j)
        {
            return ((Excel.Range)sheet.Cells[i, j]).Value?.ToString() ?? "";
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
