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

        public List<string> ListDate { get; set; }
        public List<int> ListCarsOneTime { get; set; }
        public List<int> ListCarsPermanent { get; set; }

        public string NewFileName { get; private set; }
        public string PathFolder { get; private set; }

        //Автозагрузка файлов-----------------------------------------------------------------------------------------------

        public string AutoImport(string sourceFolder, string keyFolder, string keyFile, string year, string fullPath)
        {
            string nextFolder = "";
            DirectoryInfo dir = new DirectoryInfo(sourceFolder);
            foreach (DirectoryInfo directory in dir.GetDirectories())
            {
                if (directory.Name.ToString().ToUpper() == keyFolder || directory.Name.ToString().ToUpper() == keyFolder + " " + year ||
                    directory.Name.ToString().ToUpper() == keyFolder + "  " + year)
                {
                    nextFolder = directory.Name;

                    break;
                }
            }

            dir = new DirectoryInfo(sourceFolder + nextFolder);
            foreach (FileInfo files in dir.GetFiles())
            {
                if ( (files.Name.ToString().ToUpper() == keyFile + ".XLSX" || files.Name.ToString().ToUpper() == keyFile + " " + ".XLSX" ||
                      files.Name.ToString().ToUpper() == keyFile + " " + year + ".XLSX" || files.Name.ToString().ToUpper() == keyFile + "  " + year + ".XLSX")
                      && !files.Name.ToString().Contains("$") )
                {
                    Path = sourceFolder + nextFolder + @"\" + files.Name;
                    FileName = files.Name;
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

            Excel.Application app = new Excel.Application();
            Excel.Workbook ObjWorkBook = null;
            try
            {
                ObjWorkBook = app.Workbooks.Open(Path);
            }
            catch
            {
                return error = "Не удалось открыть файл 'Статистика'. Возможно, он сейчас используется.";
            }
            Excel.Worksheet ObjWorkSheet = null;
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
            int row;                

            //Возврат на лист "Общая"
            ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets.get_Item(1);           

            //Поиск текущего года
            row = 0;
            for (int i = 1; i < 1000; i++)
            {
                if (Contains(ObjWorkSheet, i, 1, year))
                {
                    row = i;
                                     
                    break;
                }
            }
          
            //Если года нет, тогда создаём новую таблицу
            if (row == 0)
            {
                InsertEmptyRows(ObjWorkSheet, 1, "A", 16, "A");
                InsertNewTable(ObjWorkSheet, 17, "A", 30, "A", 1, "A", 3, "B", 14, "U");
                ObjWorkSheet.Cells[1, 1] = year;

                row = 1;
            }

                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "БЕСХОЗ", Beshoz.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "РАЗОВ", CarsOneTime.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "ПОСТОЯН", CarsPermanent.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "АБ/ТБ", Abtb.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "МВД", Police.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "СТОРОН", AnotherOrg.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "А/П", Airport.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "ТРУД", TrudRas.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "ПРОСМОТР", Viewing.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "USB", Usb.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "DVD", Dvd.ToString());
                InsertData(ObjWorkSheet, row, Convert.ToInt32(monthNum) + 1, "ДИСК N", DiskN.ToString());

            //Сохранить                       
            string date = $"{DateTime.Now.Day}.{DateTime.Now.Month}.{DateTime.Now.Year}";           
            PathFolder = Directory.GetCurrentDirectory() + @"\";
            NewFileName = "Статистика_" + month + "_" + year + ".xlsx";

            error = Save(app, ObjWorkBook, error);

            // Создание папки _backup, если её нет.
            bool existBackupFolder = Directory.Exists("_backup");
            if (!existBackupFolder)
            {
                Directory.CreateDirectory("_backup");
            }

            File.Replace(NewFileName, FileName, @"_backup\Статистика (от " + date + ").xlsx");

            return error;
        }

        //Статистика за неделю----------------------------------------------------------------------------------------------

        public string CalculateWeek(string day1, string day2, string monthNum1, string monthNum2, string month, string year1, string year2, bool cancel)
        {
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
                return error = "Не удалось открыть файл 'Статистика'. Возможно, он сейчас используется.";
            }
            Excel.Worksheet ObjWorkSheet = null;
            //Отключить отображение окон с сообщениями
            app.DisplayAlerts = false;
            int row = 0;
            int rowCars = 0;
            int rowCopy = 0;
            int rowPaste = 0;

            //Меняем на лист "За неделю"
            ObjWorkSheet = SearchPage(ObjWorkBook, ObjWorkSheet, "За неделю");

            //Создание таблицы, если их нет
            if (ToString(ObjWorkSheet, 1, 1) == "" && ToString(ObjWorkSheet, 2, 1) == "" && ToString(ObjWorkSheet, 3, 1) == "")
            {
                rowPaste = CreateNewTable(ObjWorkSheet);
            }

            else
            {
                //Вставить таблицу за неделю
                for (int i = 1; i < 1000; i++)
                {
                    if (ToString(ObjWorkSheet, i, 1) == "" && ToString(ObjWorkSheet, i + 1, 1) == "" && ToString(ObjWorkSheet, i + 2, 1) == "" &&
                        ToString(ObjWorkSheet, i + 3, 1) == "" && ToString(ObjWorkSheet, i + 4, 1) == "" && ToString(ObjWorkSheet, i + 5, 1) == "" &&
                        ToString(ObjWorkSheet, i + 6, 1) == "" && ToString(ObjWorkSheet, i + 7, 1) == "" && ToString(ObjWorkSheet, i + 8, 1) == "")
                    {
                        row = i + 1;

                        rowCars = row - 2;
                        while (ToString(ObjWorkSheet, rowCars, 9) != "")
                        {
                            rowCars++;
                        }

                        break;
                    }
                }

                rowCopy = row - 4;
                rowPaste = rowCars + 1;
                InsertNewTable(ObjWorkSheet, rowCopy, "A", rowCopy + 2, "A", rowPaste, "A", rowPaste + 2, "B", rowPaste + 2, "U");
            }

            //Заполнение таблицы
            ObjWorkSheet.Cells[rowPaste, 1] = day1 + "." + monthNum1 + "." + year1 + "-" + day2 + "." + monthNum2 + "." + year2;
            ObjWorkSheet.Cells[rowPaste + 2, 1] = month;

            InsertData(ObjWorkSheet, rowPaste, 2, "МВД", Police.ToString());
            InsertData(ObjWorkSheet, rowPaste, 2, "СТОРОН", AnotherOrg.ToString());
            InsertData(ObjWorkSheet, rowPaste, 2, "А/П", Airport.ToString());
            InsertData(ObjWorkSheet, rowPaste, 2, "ПРОСМОТР", Viewing.ToString());
            InsertData(ObjWorkSheet, rowPaste, 2, "USB", Usb.ToString());
            InsertData(ObjWorkSheet, rowPaste, 2, "DVD", Dvd.ToString());
            InsertData(ObjWorkSheet, rowPaste, 2, "ДИСК N", DiskN.ToString());
            InsertData(ObjWorkSheet, rowPaste, 2, "30");

            //Рамки 30м
            if (ListDate.Count > 1)
            {
                DrawBorders(ObjWorkSheet, rowPaste + 3, "I", rowPaste + ListDate.Count + 1, "I");
            }
   
            //Сохранить                       
            string date = $"{DateTime.Now.Day}.{DateTime.Now.Month}.{DateTime.Now.Year}";
            PathFolder = Directory.GetCurrentDirectory() + @"\";
            NewFileName = "Статистика_" + day1 + "." + monthNum1 + ".-" + day2 + "." + monthNum2 + "." + year2 + ".xlsx";

            error = Save(app, ObjWorkBook, error);

            // Создание папки _backup, если её нет.
            bool existBackupFolder = Directory.Exists("_backup");
            if (!existBackupFolder)
            {
                Directory.CreateDirectory("_backup");
            }

            File.Replace(NewFileName, FileName, @"_backup\Статистика (от " + date + ").xlsx");

            return error;
        }

        //Сохранить---------------------------------------------------------------------------------------------------------

        private string Save(Excel.Application app, Excel.Workbook ObjWorkBook, string error)
        {
            //int index = 0;
            //int lastIndex = -1;

            //while (index > -1)
            //{
            //    index = path.IndexOf((@"\"), index + 1);
            //    if (index > -1)
            //    {
            //        lastIndex = index;
            //    }
            //}

            //pathFolder = path.Remove(lastIndex + 1);
            try
            {
                ObjWorkBook.SaveAs(PathFolder + NewFileName);
            }
            catch
            {
                error = "Не удалось сохранить файл 'Статистика'";
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

        //Вставка пустых строк----------------------------------------------------------------------------------------------

        private void InsertEmptyRows(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY)
        {
            Excel.Range rowRange = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            rowRange = rowRange.EntireRow;
            rowRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown, false);         
        }

        //Создание новой таблицы--------------------------------------------------------------------------------------------

        private int CreateNewTable(Excel.Worksheet ObjWorkSheet)
        {
            //Шапка
            ObjWorkSheet.Cells[3, 1] = "Месяц";
            ObjWorkSheet.Cells[3, 2] = "Предоставление видеоматериалов по запросам МВД";
            ObjWorkSheet.Cells[3, 3] = "Предоставление видеоматериалов по запросам сторонних организаций";
            ObjWorkSheet.Cells[3, 4] = "Предоставление видеоматериалов по запросам подразделений а/п";
            ObjWorkSheet.Cells[3, 5] = "Просмотр видеоархива";
            ObjWorkSheet.Cells[3, 6] = "Выдано видеомате-риалов на DVD";
            ObjWorkSheet.Cells[3, 7] = "Выдано видеома-териалов на USB";
            ObjWorkSheet.Cells[3, 8] = "Выдано видеома-териалов на диск N";
            ObjWorkSheet.Cells[3, 9] = "30 м. зона";

            //Высота строчки
            SetHeightRow(ObjWorkSheet, 3, "A", 3, "I", 60);

            //Ширина столбцов
            SetWidthColumn(ObjWorkSheet, 3, "A", 3, "A", 10);
            SetWidthColumn(ObjWorkSheet, 3, "B", 3, "H", 20);
            SetWidthColumn(ObjWorkSheet, 3, "I", 3, "I", 30);

            //Переносить по словам
            SetWrapText(ObjWorkSheet, 3, "A", 3, "I", true);

            //Заливка
            FillBackground(ObjWorkSheet, 2, "A", 2, "B", 255, 192, 0);
            FillBackground(ObjWorkSheet, 3, "A", 3, "I", 112, 173, 71);
            FillBackground(ObjWorkSheet, 4, "A", 4, "A", 255, 255, 0);

            //Жирный шрифт
            ApplyFontBold(ObjWorkSheet, 3, "A", 3, "A", true);

            //Рамки
            DrawBorders(ObjWorkSheet, 3, "A", 4, "I");

            //Выравнивание
            HorizontalAlignment(ObjWorkSheet, 3, "A", 4, "I", Excel.XlVAlign.xlVAlignCenter);
            VerticalAlignment(ObjWorkSheet, 3, "A", 3, "I", Excel.XlVAlign.xlVAlignBottom);

            int row = 2;
            return row;
        }

        //Вставка новой таблицы---------------------------------------------------------------------------------------------

        private void InsertNewTable(Excel.Worksheet ObjWorkSheet, int startCopyX, string startCopyY, int lastCopyX, string lastCopyY,
            int startPasteX, string startPasteY, int startClearX, string startClearY, int lastClearX, string lastClearY)
        {         
            Excel.Range monthRange = ObjWorkSheet.Range[startCopyY + startCopyX + ":" + lastCopyY + lastCopyX];
            monthRange = monthRange.EntireRow;
            monthRange.Copy();
            ObjWorkSheet.Range[startPasteY + startPasteX].PasteSpecial();
            ObjWorkSheet.Range[startClearY + startClearX + ":" + lastClearY + lastClearX].ClearContents();
        }

        //Вставка данных----------------------------------------------------------------------------------------------------

        private void InsertData(Excel.Worksheet ObjWorkSheet, int startTableIndex, int rowNum, string key, string value)
        {
            for (int j = 2; j < 20; j++)
            {
                if (Contains(ObjWorkSheet, startTableIndex + 1, j, key))
                {
                    ObjWorkSheet.Cells[startTableIndex + rowNum, j] = value;

                    break;
                }
            }
        }

        //Вставка данных 30м------------------------------------------------------------------------------------------------

        private void InsertData(Excel.Worksheet ObjWorkSheet, int startTableIndex, int rowNum, string key)
        {
            for (int j = 2; j < 20; j++)
            {
                if (Contains(ObjWorkSheet, startTableIndex + 1, j, key))
                {
                    int rowCars = startTableIndex + rowNum;

                    if (ListDate.Count > 0)
                    {
                        for (int i = 0; i < ListDate.Count; i++)
                        {
                            if (ListCarsOneTime[i] > 0 && ListCarsPermanent[i] > 0)
                                ObjWorkSheet.Cells[i + rowCars, j] = ListDate[i] + " - " + ListCarsOneTime[i] + " (раз.), " + ListCarsPermanent[i] + " (пост.)";
                            else if (ListCarsOneTime[i] > 0)
                                ObjWorkSheet.Cells[i + rowCars, j] = ListDate[i] + " - " + ListCarsOneTime[i] + " (раз.)";
                            else if (ListCarsPermanent[i] > 0)
                                ObjWorkSheet.Cells[i + rowCars, j] = ListDate[i] + " - " + ListCarsPermanent[i] + " (пост.)";
                        }
                        // TODO: горизонтальное выравнивание по левому краю все списки

                    }
                    else ObjWorkSheet.Cells[rowCars, j] = 0;
                    // TODO: горизонтальное выравнивание по центру (занести в предыдущий блок где результат = 0)

                    break;
                }
            }
        }

        //Поиск листа-------------------------------------------------------------------------------------------------------

        private Excel.Worksheet SearchPage(Excel.Workbook ObjWorkBook, Excel.Worksheet ObjWorkSheet, string pageName)
        {
            foreach (Excel.Worksheet sheet in ObjWorkBook.Sheets)
            {
                if (sheet.Name == pageName)
                {
                    ObjWorkSheet = sheet;
                    return ObjWorkSheet;
                }
            }
            return null;
        }

        //Вставить новый лист-----------------------------------------------------------------------------------------------

        private Excel.Worksheet InsertNewSheet(Excel.Workbook ObjWorkBook, string previousPage, string pageName)
        {
            var sh = ObjWorkBook.Sheets;
            Excel.Worksheet sheetNew = (Excel.Worksheet)sh.Add(Type.Missing, sh[previousPage], Type.Missing, Type.Missing);
            sheetNew.Name = pageName;
            return sheetNew;
        }

        //Высота строчки----------------------------------------------------------------------------------------------------

        private void SetHeightRow(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY, int height)
        {
            Excel.Range range;

            range = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            range.RowHeight = height;
        }

        //Ширина столбцов---------------------------------------------------------------------------------------------------

        private void SetWidthColumn(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY, int width)
        {
            Excel.Range range;

            range = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            range.ColumnWidth = width;
        }

        //Автоширина столбцов-----------------------------------------------------------------------------------------------

        private void AutoFitColumn(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY)
        {
            Excel.Range range;

            range = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            range.EntireColumn.AutoFit();
        }

        //Перенос по словам-------------------------------------------------------------------------------------------------

        private void SetWrapText(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY, bool wrap)
        {
            Excel.Range range;

            range = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            range.WrapText = wrap;
        }

        //Заливка фона------------------------------------------------------------------------------------------------------

        private void FillBackground(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY, int R, int G, int B)
        {
            Excel.Range range;

            range = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            range.Interior.Color = Color.FromArgb(R, G, B);
        }

        //Жирный шрифт------------------------------------------------------------------------------------------------------

        private void ApplyFontBold(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY, bool bold)
        {
            Excel.Range range;

            range = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            range.Cells.Font.Bold = bold;
        }

        //Рамки-------------------------------------------------------------------------------------------------------------

        private void DrawBorders(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY)
        {
            Excel.Range range;

            range = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.XlLineStyle.xlContinuous;            
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlInsideHorizontal).LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders.get_Item(Excel.XlBordersIndex.xlInsideVertical).LineStyle = Excel.XlLineStyle.xlContinuous;                     
        }

        //Горизонтальное выравнивание---------------------------------------------------------------------------------------

        private void HorizontalAlignment(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY, Excel.XlVAlign valign)
        {
            Excel.Range range;

            range = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            range.HorizontalAlignment = valign;
        }

        //Вертикальное выравнивание---------------------------------------------------------------------------------------

        private void VerticalAlignment(Excel.Worksheet ObjWorkSheet, int startX, string startY, int lastX, string lastY, Excel.XlVAlign valign)
        {
            Excel.Range range;

            range = ObjWorkSheet.Range[startY + startX + ":" + lastY + lastX];
            range.VerticalAlignment = valign;
        }

    }
}
