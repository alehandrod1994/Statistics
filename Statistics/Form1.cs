using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Statistics
{
    public partial class Form1 : Form
    {
        Cars cars;
        KPI kpi;
        Videotapes videotapes;
        Statistics statistics;
        Date date;
        Progress progress;

        public string newPath;
        public bool cancel;
        public string errorDescription;
        public string errorTitle;
        public int errorTitleHeight;

        public Form1()
        {
            InitializeComponent();

            cars = new Cars();
            kpi = new KPI();
            videotapes = new Videotapes();
            statistics = new Statistics();
            date = new Date();
                      
            DateAutoImport();
            DocumentsAutoImport();
        }

        //Автозагрузка дат--------------------------------------------------------------------------------------------------
        private void DateAutoImport()
        {
            date.AutoImport();                                                                     

            comboBoxMonth.Text = date.month;
            tbMonth.Text = date.year1;

            dTP1.Text = date.date1;
            dTP2.Text = date.date2;           
        }

        //Автозагрузка файлов-----------------------------------------------------------------------------------------------

        private void DocumentsAutoImport()
        {
            try
            {
                listBoxPath.Items[0] = cars.AutoImport(@"C:\PUBLIC_VS3\", "30", "ПАРКОВКА", date.year1, listBoxPath.Items[0].ToString());
            }
            catch { }

            try
            {
                listBoxPath.Items[2] = kpi.AutoImport(@"C:\PUBLIC_VS3\", "KPI", date.month, "KPI", date.month, listBoxPath.Items[2].ToString());
            }
            catch { }

            try
            {
                listBoxPath.Items[4] = videotapes.AutoImport(@"C:\PUBLIC_VS3\", "ЗАПРОС", "ЗАПРОС", date.year1, listBoxPath.Items[4].ToString());
            }
            catch { }

            try
            {               
                listBoxPath.Items[6] = statistics.AutoImport(@"C:\Users\VS 1\Desktop\", "СТАТИСТИКА", "СТАТИСТИКА", date.year1, listBoxPath.Items[6].ToString());
            }
            catch
            {
                try
                {
                    listBoxPath.Items[6] = statistics.AutoImport(@"C:\PUBLIC_VS3\", "СТАТИСТИКА", "СТАТИСТИКА", date.year1, listBoxPath.Items[6].ToString());
                }
                catch { }
            }
        }

        //Проверка на ошибки------------------------------------------------------------------------------------------------

        public bool CheckDateMonth()
        {
            try
            {
                Convert.ToInt32(tbMonth.Text);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public bool CheckingForErrors()
        {
            errorDescription = "";

            if ( (radioButtonWeek.Checked && dTP1.Value > dTP2.Value) || (radioButtonMonth.Checked && CheckDateMonth() == false) )
            {                 
                FormErrorShow("Неверно задана дата", 27, "");
                return false;               
            }

            //Если выбрана статистика за месяц
            if (radioButtonMonth.Checked)
            {                
                if (kpi.path == "") errorDescription += "KPI\n";
            }

            //Общее      
            if (cars.path == "") errorDescription += "30м\n";
            if (videotapes.path == "") errorDescription += "Запросы\n";
            if (statistics.path == "") errorDescription += "Статистика";

            if (errorDescription != "")
            {
                FormErrorShow("Не загружены файлы:", 27, errorDescription);
                return false;
            }

            return true;
        }

        public bool CheckCalculateErrors()
        {
            if (errorTitle != "")
            {
                CalculateStop();
                if (errorTitle != "cancel")
                    FormErrorShow(errorTitle, 67, "");
                return false;
            }
            else return true;
        }

        //Открыть форму для вывода ошибок-----------------------------------------------------------------------------------

        private void FormErrorShow(string _errorTitle, int _errorTitleHeight, string _errorDescription)
        {
            FormError FormError = new FormError();
            errorTitle = _errorTitle;
            errorTitleHeight = _errorTitleHeight;
            errorDescription = _errorDescription;
            FormError.Show(this);
        }

        //Движение полосы загрузки и текст----------------------------------------------------------------------------------

        private void ProgressMove(string action, string typeDoc)
        {
            progress.Move(action, typeDoc);
            labelProgress.Text = progress.text;
            imgProgressBar_100.Width = progress.value;
        }

        //Статистика за месяц-----------------------------------------------------------------------------------------------

        public async void CalculateMonthAsync()
        {
            CalculateStart(4);

            cancel = false;

            date.month = comboBoxMonth.Text;
            date.monthNum1 = (comboBoxMonth.SelectedIndex + 1).ToString();
            date.year1 = tbMonth.Text;
            if (date.monthNum1.Length == 1)
            {
                date.monthNum1 = "0" + date.monthNum1;
            }

            ProgressMove("Подсчёт", "'30м'");
            //errorTitle = cars.CalculateMonth(statistics, date.monthNum1, cancel);
            errorTitle = await Task.Run(() => cars.CalculateMonth(statistics, date.monthNum1, cancel));
            if (!CheckCalculateErrors()) return;
            Application.DoEvents();

            ProgressMove("Подсчёт", "'KPI'");
            errorTitle = await Task.Run(() => kpi.CalculateMonth(statistics, cancel));
            if (!CheckCalculateErrors()) return;
            Application.DoEvents();

            ProgressMove("Подсчёт", "'Запросов'");
            errorTitle = await Task.Run(() => videotapes.CalculateMonth(statistics, date.monthNum1, cancel));
            if (!CheckCalculateErrors()) return;
            Application.DoEvents();

            ProgressMove("Сохранение", "'Статистики'");
            errorTitle = await Task.Run(() => statistics.CalculateMonth(date.monthNum1, date.month, date.year1, cancel));
            if (!CheckCalculateErrors()) return;
            Application.DoEvents();

            CalculateStop();

            FormSuccessfullyShow();
        }     

        //Статистика за неделю----------------------------------------------------------------------------------------------

        public async void CalculateWeekAsync()
        {
            CalculateStart(3);

            date.day1 = dTP1.Text.Substring(0, 2);
            date.monthNum1 = dTP1.Text.Substring(3, 2);
            date.year1 = dTP1.Text.Substring(6, 4);

            date.day2 = dTP2.Text.Substring(0, 2);
            date.monthNum2 = dTP2.Text.Substring(3, 2);
            date.year2 = dTP2.Text.Substring(6, 4);

            date.month = comboBoxMonth.Items[Convert.ToInt32(date.monthNum2) - 1].ToString();

            ProgressMove("Подсчёт", "'30м'");
            errorTitle = await Task.Run(() => cars.CalculateWeek(statistics, date.day1, date.day2, date.monthNum1, date.monthNum2, date.year1, date.year2, cancel));
            if (!CheckCalculateErrors()) return;
            Application.DoEvents();

            ProgressMove("Подсчёт", "'Запросов'");
            errorTitle = await Task.Run(() => videotapes.CalculateWeek(statistics, date.day1, date.day2, date.monthNum1, date.monthNum2, date.year1, date.year2, cancel));
            if (!CheckCalculateErrors()) return;
            Application.DoEvents();

            ProgressMove("Сохранение", "'Статистики'");
            errorTitle = await Task.Run(() => statistics.CalculateWeek(date.day1, date.day2, date.monthNum1, date.monthNum2, date.month, date.year1, date.year2, cancel));
            if (!CheckCalculateErrors()) return;
            Application.DoEvents();

            CalculateStop();

            FormSuccessfullyShow();
        }

        //Начало подсчёта---------------------------------------------------------------------------------------------------

        private void CalculateStart(int numDoc)
        {
            progress.stepLast = numDoc;

            labelProgress.Visible = true;
            imgProgressBar_0.Visible = true;
            imgProgressBar_100.Width = 0;
            imgProgressBar_100.Visible = true;

            btn30m.Enabled = false;
            btnKPI.Enabled = false;
            btnVideotapes.Enabled = false;
            btnStatistics.Enabled = false;
            btnCancel.Visible = true;
            btnCalculate.Visible = false;
        }

        //Конец подсчёта----------------------------------------------------------------------------------------------------

        private void CalculateStop()
        {
            labelProgress.Visible = false;
            imgProgressBar_0.Visible = false;
            imgProgressBar_100.Visible = false;

            btn30m.Enabled = true;
            btnKPI.Enabled = true;
            btnVideotapes.Enabled = true;
            btnStatistics.Enabled = true;
            btnCalculate.Visible = true;
            btnCancel.Visible = false;
        }

        //Открыть форму для уведомлений-------------------------------------------------------------------------------------

        private void FormSuccessfullyShow()
        {
            FormSuccessfully FormSuccessfully = new FormSuccessfully();
            newPath = statistics.pathFolder + statistics.fileName;
            FormSuccessfully.Show(this);
        }

        //События при нажатии кнопок----------------------------------------------------------------------------------------

        private void btnCalculate_Click(object sender, EventArgs e)
        {
            if (!CheckingForErrors())
            {
                return;
            }

            progress = new Progress();           

            if (radioButtonMonth.Checked)
            {
                CalculateMonthAsync();
            }
            else CalculateWeekAsync();           
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            cancel = true;
        }

        private void btn30m_Click(object sender, EventArgs e)
        {
            listBoxPath.Items[0] = cars.Import(listBoxPath.Items[0].ToString());
        }

        private void btnKPI_Click(object sender, EventArgs e)
        {
            if (radioButtonMonth.Checked)
            {
                listBoxPath.Items[2] = kpi.Import(listBoxPath.Items[2].ToString());
            }
            else FormErrorShow("При выборе формата 'Статистика за неделю' загрузка данного файла не требуется", 67, "");
        }

        private void btnVideotapes_Click(object sender, EventArgs e)
        {          
            listBoxPath.Items[4] = videotapes.Import(listBoxPath.Items[4].ToString());
        }

        private void btnStatistics_Click(object sender, EventArgs e)
        {           
            listBoxPath.Items[6] = statistics.Import(listBoxPath.Items[6].ToString());
        }

        private void radioButtonMonth_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxMonth.Enabled = true;
            tbMonth.Enabled = true;
            dTP1.Enabled = false;
            dTP2.Enabled = false;

            listBoxPath.Items[2] = "KPI                    | " + kpi.path;
        }

        private void radioButtonWeek_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxMonth.Enabled = false;
            tbMonth.Enabled = false;
            dTP1.Enabled = true;
            dTP2.Enabled = true;

            listBoxPath.Items[2] = "";
        }

        private void listBoxPath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false) == true)
                e.Effect = DragDropEffects.All;
        }

        private void listBoxPath_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                if (radioButtonMonth.Checked)
                {                  
                    listBoxPath.Items[2] = kpi.DragDrop("KPI", date.listMonth, file, listBoxPath.Items[2].ToString());
                }
                listBoxPath.Items[0] = cars.DragDrop("30", "ПАРКОВКА", file, listBoxPath.Items[0].ToString());
                listBoxPath.Items[4] = videotapes.DragDrop("ЗАПРОС", file, listBoxPath.Items[4].ToString());
                listBoxPath.Items[6] = statistics.DragDrop("СТАТИСТИКА", file, listBoxPath.Items[6].ToString());                
            }            
        }

        //Анимация кнопок---------------------------------------------------------------------------------------------------

        private void btn30m_MouseEnter(object sender, EventArgs e)
        {
            btn30m.Image = Properties.Resources.btn_cars_enter;
        }

        private void btn30m_MouseLeave(object sender, EventArgs e)
        {
            btn30m.Image = Properties.Resources.btn_cars_normal;
        }

        private void btn30m_MouseDown(object sender, MouseEventArgs e)
        {
            btn30m.Image = Properties.Resources.btn_cars_down;
        }

        private void btnKPI_MouseEnter(object sender, EventArgs e)
        {
            btnKPI.Image = Properties.Resources.btn_kpi_enter;
        }

        private void btnKPI_MouseLeave(object sender, EventArgs e)
        {
            btnKPI.Image = Properties.Resources.btn_kpi_normal;
        }

        private void btnKPI_MouseDown(object sender, MouseEventArgs e)
        {
            btnKPI.Image = Properties.Resources.btn_kpi_down;
        }

        private void btnVideotapes_MouseEnter(object sender, EventArgs e)
        {
            btnVideotapes.Image = Properties.Resources.btn_videotapes_enter;
        }

        private void btnVideotapes_MouseLeave(object sender, EventArgs e)
        {
            btnVideotapes.Image = Properties.Resources.btn_videotapes_normal;
        }

        private void btnVideotapes_MouseDown(object sender, MouseEventArgs e)
        {
            btnVideotapes.Image = Properties.Resources.btn_videotapes_down;
        }

        private void btnStatistics_MouseEnter(object sender, EventArgs e)
        {
            btnStatistics.Image = Properties.Resources.btn_statistics_enter;
        }

        private void btnStatistics_MouseLeave(object sender, EventArgs e)
        {
            btnStatistics.Image = Properties.Resources.btn_statistics_normal;
        }

        private void btnStatistics_MouseDown(object sender, MouseEventArgs e)
        {
            btnStatistics.Image = Properties.Resources.btn_statistics_down;
        }

        private void btnCalculate_MouseEnter(object sender, EventArgs e)
        {
            btnCalculate.Image = Properties.Resources.btn_calculate_enter;
        }

        private void btnCalculate_MouseLeave(object sender, EventArgs e)
        {
            btnCalculate.Image = Properties.Resources.btn_calculate_normal;
        }

        private void btnCalculate_MouseDown(object sender, MouseEventArgs e)
        {
            btnCalculate.Image = Properties.Resources.btn_calculate_down;
        }

        private void btnCancel_MouseEnter(object sender, EventArgs e)
        {
            btnCancel.Image = Properties.Resources.btn_cancel_enter;
        }

        private void btnCancel_MouseLeave(object sender, EventArgs e)
        {
            btnCancel.Image = Properties.Resources.btn_cancel_normal;
        }

        private void btnCancel_MouseDown(object sender, MouseEventArgs e)
        {
            btnCancel.Image = Properties.Resources.btn_cancel_down;
        }
    }
}
