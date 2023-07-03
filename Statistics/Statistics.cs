using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Collections.Generic;

namespace Statistics
{
    public class Statistics : DocumentsForStatistics
    {
        public int Beshoz { get; set; }
        public int Abtb { get; set; }
        public int TrudRas { get; set; }
        public int CarsOneTime { get; set; }
        public int CarsPermanent { get; set; }
        public int Police { get; set; }
        public int AnotherOrg { get; set; }
        public int Airport { get; set; }
        public int Viewing { get; set; }
        public int Usb { get; set; }
        public int Dvd { get; set; }
        public int DiskN { get; set; }
        public List<CarDay> CarDays { get; set; }

        public string NewFileName { get; private set; }
        public string PathFolder { get; private set; }

        //Автозагрузка файлов-----------------------------------------------------------------------------------------------

        public string AutoImport(string sourceFolder, string keyFolder, string keyFile, string year, string fullPath)
        {
            string nextFolder = "";
            DirectoryInfo dir = new DirectoryInfo(sourceFolder);
            foreach (DirectoryInfo directory in dir.GetDirectories())
            {
                if (directory.Name.ToUpper().Contains(keyFolder) && directory.Name.ToUpper().Contains(year))
                {
                    nextFolder = directory.Name;

                    break;
                }
            }

            dir = new DirectoryInfo(sourceFolder + nextFolder);
            foreach (FileInfo file in dir.GetFiles())
            {
                if (file.Name.ToUpper() == keyFile + ".XLSX" && !file.Name.Contains("$"))
                {
                    Path = sourceFolder + nextFolder + @"\" + file.Name;
                    FileName = file.Name;
                    fullPath = FillingPaths(fullPath, Path);

                    break;
                }
            }

            return fullPath;
        }

        //Статистика за месяц-----------------------------------------------------------------------------------------------

        public string CalculateMonth(string monthNum, string month, string year, bool cancel)
        {
            string error = "";

            if (cancel == true) return error = "cancel";

            //Excel.Application app = new Excel.Application();
            //Excel.Workbook ObjWorkBook = null;
            //try
            //{
            //    ObjWorkBook = app.Workbooks.Open(Path);
            //}
            //catch
            //{
            //    return error = "Не удалось открыть файл 'Статистика'. Возможно, он сейчас используется.";
            //}
            //Excel.Worksheet ObjWorkSheet = null;
            ////Отключить отображение окон с сообщениями
            //app.DisplayAlerts = false;
            int row;

            //Возврат на лист "Общая"
            //ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.get_Item(1);

            if (!OpenConnection())
            {
                return "Не удалось открыть файл \"Статистика\". Возможно, он сейчас используется.";
            }

            //Поиск текущего года
            row = 0;
            for (int i = 1; i < 1000; i++)
            {
                if (Contains(i, 1, year))
                {
                    row = i;
                                     
                    break;
                }
            }
          
            //Если года нет, тогда создаём новую таблицу
            if (row == 0)
            {
                InsertEmptyRows(1, "A", 16, "A");
                InsertNewTable(17, "A", 30, "A", 1, "A", 3, "B", 14, "U");
                sheet.Cells[1, 1] = year;

                row = 1;
            }

                InsertData(row, Convert.ToInt32(monthNum) + 1, "БЕСХОЗ", Beshoz.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "РАЗОВ", CarsOneTime.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "ПОСТОЯН", CarsPermanent.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "АБ/ТБ", Abtb.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "МВД", Police.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "СТОРОН", AnotherOrg.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "А/П", Airport.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "ТРУД", TrudRas.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "ПРОСМОТР", Viewing.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "USB", Usb.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "DVD", Dvd.ToString());
                InsertData(row, Convert.ToInt32(monthNum) + 1, "ДИСК N", DiskN.ToString());

            //Сохранить                       
            string date = $"{DateTime.Now.Day}.{DateTime.Now.Month}.{DateTime.Now.Year}";
            PathFolder = GetPathFolder();
            NewFileName = "Статистика_" + month + "_" + year + ".xlsx";

            error = Save(error);

            ReplaceFiles(date);

            return error;
        }

        //Статистика за неделю----------------------------------------------------------------------------------------------

        public string CalculateWeek(string day1, string day2, string monthNum1, string monthNum2, string month, string year1, string year2, bool cancel)
        {
            string error = "";

            if (cancel == true) return error = "cancel";

            //Excel.Application app = new Excel.Application();
            //Excel.Workbook ObjWorkBook = null;
            //try
            //{
            //    ObjWorkBook = app.Workbooks.Open(Path);
            //}
            //catch
            //{
            //    return error = "Не удалось открыть файл 'Статистика'. Возможно, он сейчас используется.";
            //}
            //Excel.Worksheet ObjWorkSheet = null;
            ////Отключить отображение окон с сообщениями
            //app.DisplayAlerts = false;

            if (!OpenConnection())
            {
                return "Не удалось открыть файл \"Статистика\". Возможно, он сейчас используется.";
            }

            int row = 0;
            int rowCars = 0;
            int rowCopy = 0;
            int rowPaste = 0;

            //Меняем на лист "За неделю"
            sheet = SearchPage("За неделю");

            //Создание таблицы, если их нет
            if (ToString(1, 1) == "" && ToString(2, 1) == "" && ToString(3, 1) == "")
            {
                rowPaste = CreateNewTable();
            }

            else
            {
                //Вставить таблицу за неделю
                for (int i = 1; i < 10000; i++)
                {
                    if (ToString(i, 1) == "" && ToString(i + 1, 1) == "" && ToString(i + 2, 1) == "" &&
                        ToString(i + 3, 1) == "" && ToString(i + 4, 1) == "" && ToString(i + 5, 1) == "" &&
                        ToString(i + 6, 1) == "" && ToString(i + 7, 1) == "" && ToString(i + 8, 1) == "")
                    {
                        row = i + 1;

                        rowCars = row - 2;
                        while (ToString(rowCars, 9) != "")
                        {
                            rowCars++;
                        }

                        break;
                    }
                }

                rowCopy = row - 4;
                rowPaste = rowCars + 1;
                InsertNewTable(rowCopy, "A", rowCopy + 2, "A", rowPaste, "A", rowPaste + 2, "B", rowPaste + 2, "U");
            }

            //Заполнение таблицы
            sheet.Cells[rowPaste, 1] = day1 + "." + monthNum1 + "." + year1 + "-" + day2 + "." + monthNum2 + "." + year2;
            sheet.Cells[rowPaste + 2, 1] = month;

            InsertData(rowPaste, 2, "МВД", Police.ToString());
            InsertData(rowPaste, 2, "СТОРОН", AnotherOrg.ToString());
            InsertData(rowPaste, 2, "А/П", Airport.ToString());
            InsertData(rowPaste, 2, "ПРОСМОТР", Viewing.ToString());
            InsertData(rowPaste, 2, "USB", Usb.ToString());
            InsertData(rowPaste, 2, "DVD", Dvd.ToString());
            InsertData(rowPaste, 2, "ДИСК N", DiskN.ToString());
            InsertData(rowPaste, 2, "30");

            //Рамки 30м
            if (CarDays.Count > 1)
            {
                DrawBorders(rowPaste + 3, "I", rowPaste + CarDays.Count + 1, "I");
            }
   
            //Сохранить                       
            string date = $"{DateTime.Now.Day}.{DateTime.Now.Month}.{DateTime.Now.Year}";
            PathFolder = GetPathFolder();
            NewFileName = "Статистика_" + day1 + "." + monthNum1 + ".-" + day2 + "." + monthNum2 + "." + year2 + ".xlsx";

            error = Save(error);

            ReplaceFiles(date);

            return error;
        }

        //Замена файлов-----------------------------------------------------------------------------------------------------

        private void ReplaceFiles(string date)
        {
            bool existBackupFolder = Directory.Exists(PathFolder + "_backup");
            if (!existBackupFolder)
            {
                Directory.CreateDirectory(PathFolder + "_backup");
            }

            File.Replace(PathFolder + NewFileName, PathFolder + FileName, PathFolder + @"_backup\Статистика (от " + date + ").xlsx");
        }

        //Сохранить---------------------------------------------------------------------------------------------------------

        private string Save(string error)
        {
            try
            {
                book.SaveAs(PathFolder + NewFileName);
            }
            catch
            {
                error = "Не удалось сохранить файл 'Статистика'";
            }

            try
            {
                book.Close();
                app.Quit();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), ex.Source);
            }

            return error;
        }

        //Вернуть имя файла-------------------------------------------------------------------------------------------------

        private string GetFileName()
        {
            int lastIndex = TrimPathFolderAndFIle();

            return Path.Remove(0, lastIndex + 1);
        }

        //Разделить путь файла и папок--------------------------------------------------------------------------------------

        private int TrimPathFolderAndFIle()
        {
            int index = 0;
            int lastIndex = -1;

            while (index > -1)
            {
                index = Path.IndexOf((@"\"), index + 1);
                if (index > -1)
                {
                    lastIndex = index;
                }
            }

            return lastIndex;
        }

        //Вернуть путь папки------------------------------------------------------------------------------------------------

        private string GetPathFolder()
        {
            int lastIndex = TrimPathFolderAndFIle();

            return Path.Remove(lastIndex + 1);
        }

        //Вставка пустых строк----------------------------------------------------------------------------------------------

        private void InsertEmptyRows(int startX, string startY, int lastX, string lastY)
        {
            Excel.Range rowRange = sheet.Range[startY + startX + ":" + lastY + lastX];
            rowRange = rowRange.EntireRow;
            rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);         
        }

        //Создание новой таблицы--------------------------------------------------------------------------------------------

        private int CreateNewTable()
        {
            //Шапка
            sheet.Cells[3, 1] = "Месяц";
            sheet.Cells[3, 2] = "Предоставление видеоматериалов по запросам МВД";
            sheet.Cells[3, 3] = "Предоставление видеоматериалов по запросам сторонних организаций";
            sheet.Cells[3, 4] = "Предоставление видеоматериалов по запросам подразделений а/п";
            sheet.Cells[3, 5] = "Просмотр видеоархива";
            sheet.Cells[3, 6] = "Выдано видеомате-риалов на DVD";
            sheet.Cells[3, 7] = "Выдано видеома-териалов на USB";
            sheet.Cells[3, 8] = "Выдано видеома-териалов на диск N";
            sheet.Cells[3, 9] = "30 м. зона";

            //Высота строчки
            SetHeightRow(3, "A", 3, "I", 60);

            //Ширина столбцов
            SetWidthColumn(3, "A", 3, "A", 10);
            SetWidthColumn(3, "B", 3, "H", 20);
            SetWidthColumn(3, "I", 3, "I", 30);

            //Переносить по словам
            SetWrapText(3, "A", 3, "I", true);

            //Заливка
            FillBackground(2, "A", 2, "B", 255, 192, 0);
            FillBackground(3, "A", 3, "I", 112, 173, 71);
            FillBackground(4, "A", 4, "A", 255, 255, 0);

            //Жирный шрифт
            ApplyFontBold(3, "A", 3, "A", true);

            //Рамки
            DrawBorders(3, "A", 4, "I");

            //Выравнивание
            HorizontalAlignment(3, "A", 4, "I", Excel.XlVAlign.xlVAlignCenter);
            VerticalAlignment(3, "A", 3, "I", Excel.XlVAlign.xlVAlignBottom);

            int row = 2;
            return row;
        }

        //Вставка новой таблицы---------------------------------------------------------------------------------------------

        private void InsertNewTable(int startCopyX, string startCopyY, int lastCopyX, string lastCopyY,
            int startPasteX, string startPasteY, int startClearX, string startClearY, int lastClearX, string lastClearY)
        {         
            Excel.Range monthRange = sheet.Range[startCopyY + startCopyX + ":" + lastCopyY + lastCopyX];
            monthRange = monthRange.EntireRow;
            monthRange.Copy();
            sheet.Range[startPasteY + startPasteX].PasteSpecial();
            sheet.Range[startClearY + startClearX + ":" + lastClearY + lastClearX].ClearContents();
        }

        //Вставка данных----------------------------------------------------------------------------------------------------

        private void InsertData(int startTableIndex, int rowNum, string key, string value)
        {
            for (int j = 2; j < 20; j++)
            {
                if (Contains(startTableIndex + 1, j, key))
                {
                    sheet.Cells[startTableIndex + rowNum, j] = value;

                    break;
                }
            }
        }

        //Вставка данных 30м------------------------------------------------------------------------------------------------

        private void InsertData(int startTableIndex, int rowNum, string key)
        {
            for (int j = 2; j < 20; j++)
            {
                if (Contains(startTableIndex + 1, j, key))
                {
                    int rowCars = startTableIndex + rowNum;

                    if (CarDays.Count > 0)
                    {
                        for (int i = 0; i < CarDays.Count; i++)
                        {
                            if (CarDays[i].CarsOneTime > 0 && CarDays[i].CarsPermanent > 0)
                                sheet.Cells[i + rowCars, j] = CarDays[i].Date + " - " + CarDays[i].CarsOneTime + " (раз.), " + CarDays[i].CarsPermanent + " (пост.)";
                            else if (CarDays[i].CarsOneTime > 0)
                                sheet.Cells[i + rowCars, j] = CarDays[i].Date + " - " + CarDays[i].CarsOneTime + " (раз.)";
                            else if (CarDays[i].CarsPermanent > 0)
                                sheet.Cells[i + rowCars, j] = CarDays[i].Date + " - " + CarDays[i].CarsPermanent + " (пост.)";
                        }

                        // TODO: горизонтальное выравнивание по левому краю все списки

                    }
                    else sheet.Cells[rowCars, j] = 0;
                    // TODO: горизонтальное выравнивание по центру (занести в предыдущий блок где результат = 0)

                    break;
                }
            }
        }

        //Поиск листа-------------------------------------------------------------------------------------------------------

        private Excel.Worksheet SearchPage(string pageName)
        {
            foreach (Excel.Worksheet sh in book.Sheets)
            {
                if (sh.Name == pageName)
                {
                    sheet = sh;
                    return sheet;
                }
            }
            return null;
        }

        //Вставить новый лист-----------------------------------------------------------------------------------------------

        private Excel.Worksheet InsertNewSheet(string previousPage, string pageName)
        {
            var sh = book.Sheets;
            Excel.Worksheet sheetNew = (Excel.Worksheet)sh.Add(Type.Missing, sh[previousPage], Type.Missing, Type.Missing);
            sheetNew.Name = pageName;
            return sheetNew;
        }

        //Высота строчки----------------------------------------------------------------------------------------------------

        private void SetHeightRow(int startX, string startY, int lastX, string lastY, int height)
        {
            Excel.Range range;

            range = sheet.Range[startY + startX + ":" + lastY + lastX];
            range.RowHeight = height;
        }

        //Ширина столбцов---------------------------------------------------------------------------------------------------

        private void SetWidthColumn(int startX, string startY, int lastX, string lastY, int width)
        {
            Excel.Range range;

            range = sheet.Range[startY + startX + ":" + lastY + lastX];
            range.ColumnWidth = width;
        }

        //Автоширина столбцов-----------------------------------------------------------------------------------------------

        private void AutoFitColumn( int startX, string startY, int lastX, string lastY)
        {
            Excel.Range range;

            range = sheet.Range[startY + startX + ":" + lastY + lastX];
            range.EntireColumn.AutoFit();
        }

        //Перенос по словам-------------------------------------------------------------------------------------------------

        private void SetWrapText(int startX, string startY, int lastX, string lastY, bool wrap)
        {
            Excel.Range range;

            range = sheet.Range[startY + startX + ":" + lastY + lastX];
            range.WrapText = wrap;
        }

        //Заливка фона------------------------------------------------------------------------------------------------------

        private void FillBackground(int startX, string startY, int lastX, string lastY, int R, int G, int B)
        {
            Excel.Range range;

            range = sheet.Range[startY + startX + ":" + lastY + lastX];
            range.Interior.Color = Color.FromArgb(R, G, B);
        }

        //Жирный шрифт------------------------------------------------------------------------------------------------------

        private void ApplyFontBold(int startX, string startY, int lastX, string lastY, bool bold)
        {
            Excel.Range range;

            range = sheet.Range[startY + startX + ":" + lastY + lastX];
            range.Cells.Font.Bold = bold;
        }

        //Рамки-------------------------------------------------------------------------------------------------------------

        private void DrawBorders(int startX, string startY, int lastX, string lastY)
        {
            Excel.Range range;

            range = sheet.Range[startY + startX + ":" + lastY + lastX];
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;            
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;                     
        }

        //Горизонтальное выравнивание---------------------------------------------------------------------------------------

        private void HorizontalAlignment(int startX, string startY, int lastX, string lastY, Excel.XlVAlign valign)
        {
            Excel.Range range;

            range = sheet.Range[startY + startX + ":" + lastY + lastX];
            range.HorizontalAlignment = valign;
        }

        //Вертикальное выравнивание---------------------------------------------------------------------------------------

        private void VerticalAlignment(int startX, string startY, int lastX, string lastY, Excel.XlVAlign valign)
        {
            Excel.Range range;

            range = sheet.Range[startY + startX + ":" + lastY + lastX];
            range.VerticalAlignment = valign;
        }

    }
}
