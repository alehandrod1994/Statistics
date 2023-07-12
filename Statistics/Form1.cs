using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Diagnostics;

namespace Statistics
{
    public partial class Form1 : Form
    {
        private Cars _cars;
        private KPI _kpi;
        private Videotapes _videotapes;
        private Statistics _statistics;
        private Date _date;
        private Progress _progress;
        private Settings _settings;

        public Form1()
        {
            InitializeComponent();

            CreateSourceData();
            AutoImportDate();
            AutoImportDocuments();
        }

        public string NewPath { get; private set; }
        public string ErrorDescription { get; private set; }
        public string ErrorTitle { get; private set; }
        public int ErrorTitleHeight { get; private set; }
        public event Action SetOnOtherUi;
        public event Action SetOffOtherUi;

        private void CreateSourceData()
        {
            _cars = new Cars();
            _kpi = new KPI();
            _videotapes = new Videotapes();
            _statistics = new Statistics();
            _date = new Date();
            _settings = new Settings();
        }

        //Автозагрузка дат--------------------------------------------------------------------------------------------------
        private void AutoImportDate()
        {
            _date.AutoImport();                                                                     

            comboBoxMonth.Text = _date.Month;
            tbMonth.Text = _date.Year1;

            dTP1.Text = _date.Date1;
            dTP2.Text = _date.Date2;

            int daysInMonth = DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month);
            int today = Convert.ToInt32(DateTime.Today.Day);
            if (today == daysInMonth || today < 3)
            {
                radioButtonMonth.Checked = true;
            }
            else
            {
                radioButtonWeek.Checked = true;
            }
        }

        //Автозагрузка файлов-----------------------------------------------------------------------------------------------

        private void AutoImportDocuments()
        {
            string currentDisk = Directory.Exists(@"Z:\PUBLIC_VS3\") ? @"Z:\PUBLIC_VS3\" : @"U:\PUBLIC_VS3\";

            try
            {
                listBoxPath.Items[0] = _cars.AutoImport(currentDisk, "30", "ПАРКОВКА", _date.Year1, listBoxPath.Items[0].ToString());
            }
            catch { }    
            
            try
            {
                listBoxPath.Items[2] = _kpi.AutoImport(currentDisk, "KPI", _date.Month, ".XLS", listBoxPath.Items[2].ToString());
            }
            catch { }

            try
            {
                listBoxPath.Items[4] = _videotapes.AutoImport(currentDisk, "ЗАПРОС", "ЗАПРОС", _date.Year1, listBoxPath.Items[4].ToString());
            }
            catch {}

            try
            {
                listBoxPath.Items[6] = _statistics.AutoImport(Application.StartupPath, "СТАТИСТИКА", listBoxPath.Items[6].ToString());
            }
            catch { }
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
            ErrorDescription = "";

            if ( (radioButtonWeek.Checked && dTP1.Value > dTP2.Value) || (radioButtonMonth.Checked && CheckDateMonth() == false) )
            {                 
                ShowFormError("Неверно задана дата", 27, "");
                return false;               
            }

            //Если выбрана статистика за месяц
            if (radioButtonMonth.Checked)
            {                
                if (_kpi.Path == "") ErrorDescription += "KPI\n";
            }

            //Общее      
            if (_cars.Path == "") ErrorDescription += "30м\n";
            if (_videotapes.Path == "") ErrorDescription += "Запросы\n";
            if (_statistics.Path == "") ErrorDescription += "Статистика";

            if (ErrorDescription != "")
            {
                ShowFormError("Не загружены файлы:", 27, ErrorDescription);
                return false;
            }

            return true;
        }

        public bool CheckCalculateErrors()
        {
            if (ErrorTitle != "")
            {
                SetUiForEndCalculate();
                if (ErrorTitle != "cancel")
                    ShowFormError(ErrorTitle, 67, "");
                return false;
            }
            else return true;
        }

        //Открыть форму для вывода ошибок-----------------------------------------------------------------------------------

        private void ShowFormError(string errorTitle, int errorTitleHeight, string errorDescription)
        {
            var form = new FormError();
            ErrorTitle = errorTitle;
            ErrorTitleHeight = errorTitleHeight;
            ErrorDescription = errorDescription;
            form.Show(this);
        }

        //Движение полосы загрузки и текст----------------------------------------------------------------------------------

        private void ProgressMove(string action, string typeDoc)
        {
            _progress.Move(action, typeDoc);
            labelProgress.Text = _progress.Text;
            imgProgressBar_100.Width = _progress.Value;
        }

        //Статистика за месяц-----------------------------------------------------------------------------------------------

        public async Task CalculateMonthAsync()
        {
            SetUiForStartCalculate(4);

            _date.Month = comboBoxMonth.Text;
            _date.MonthNum1 = (comboBoxMonth.SelectedIndex + 1).ToString();
            _date.Year1 = tbMonth.Text;
            if (_date.MonthNum1.Length == 1)
            {
                _date.MonthNum1 = "0" + _date.MonthNum1;
            }

            ProgressMove("Подсчёт", "30м");
            ErrorTitle = await Task.Run(() => _cars.CalculateMonth(_statistics, _date.MonthNum1));
            if (!CheckCalculateErrors()) return;

            ProgressMove("Подсчёт", "KPI");
            ErrorTitle = await Task.Run(() => _kpi.CalculateMonth(_statistics));
            if (!CheckCalculateErrors()) return;
 
            ProgressMove("Подсчёт", "Запросов");
            ErrorTitle = await Task.Run(() => _videotapes.CalculateMonth(_statistics, _date.MonthNum1));
            if (!CheckCalculateErrors()) return;

            ProgressMove("Сохранение", "Статистики");
            ErrorTitle = await Task.Run(() => _statistics.CalculateMonth(_date.MonthNum1, _date.Month, _date.Year1));
            if (!CheckCalculateErrors()) return;

            SetUiForEndCalculate();

            FormSuccessfullyShow();
        }     

        //Статистика за неделю----------------------------------------------------------------------------------------------

        public async Task CalculateWeekAsync()
        {
            SetUiForStartCalculate(3);

            _date.Day1 = dTP1.Text.Substring(0, 2);
            _date.MonthNum1 = dTP1.Text.Substring(3, 2);
            _date.Year1 = dTP1.Text.Substring(6, 4);

            _date.Day2 = dTP2.Text.Substring(0, 2);
            _date.MonthNum2 = dTP2.Text.Substring(3, 2);
            _date.Year2 = dTP2.Text.Substring(6, 4);

            _date.Month = comboBoxMonth.Items[Convert.ToInt32(_date.MonthNum2) - 1].ToString();

            ProgressMove("Подсчёт", "30м");
            ErrorTitle = await Task.Run(() => _cars.CalculateWeek(_statistics, _date.Day1, _date.Day2, _date.MonthNum1, _date.MonthNum2, _date.Year1, _date.Year2));
            if (!CheckCalculateErrors()) return;

            ProgressMove("Подсчёт", "Запросов");
            ErrorTitle = await Task.Run(() => _videotapes.CalculateWeek(_statistics, _date.Day1, _date.Day2, _date.MonthNum1, _date.MonthNum2, _date.Year1, _date.Year2));
            if (!CheckCalculateErrors()) return;

            ProgressMove("Сохранение", "Статистики");
            ErrorTitle = await Task.Run(() => _statistics.CalculateWeek(_date.Day1, _date.Day2, _date.MonthNum1, _date.MonthNum2, _date.Month, _date.Year1, _date.Year2));
            if (!CheckCalculateErrors()) return;

            SetUiForEndCalculate();

            FormSuccessfullyShow();
        }

        //Начало подсчёта---------------------------------------------------------------------------------------------------

        private void SetUiForStartCalculate(int countDoc)
        {
            _progress.StepLast = countDoc;

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
            menuStrip1.Enabled = false;

            SetOnOtherUi?.Invoke();
        }

        //Конец подсчёта----------------------------------------------------------------------------------------------------

        private void SetUiForEndCalculate()
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
            menuStrip1.Enabled = true;

            SetOffOtherUi?.Invoke();
        }

        //Открыть форму для уведомлений-------------------------------------------------------------------------------------

        private void FormSuccessfullyShow()
        {
            FormSuccessfully FormSuccessfully = new FormSuccessfully();
            NewPath = _statistics.PathFolder + _statistics.FileName;
            FormSuccessfully.Show(this);
        }

        private void SetOnDebugUi()
        {
            labelDebug.Text = "Отладка:";
            labelDebug.Visible = true;
        }

        private void SetOffDebugUi()
        {
            labelDebug.Visible = false;
        }

        private void Document_CurrentRowChanged(object sender, string e)
        {
            labelDebug.Invoke((Action)delegate { labelDebug.Text = e; });
        }

        //События при нажатии кнопок----------------------------------------------------------------------------------------

        private void BtnCalculate_Click(object sender, EventArgs e)
        {
            if (!CheckingForErrors())
            {
                return;
            }

            _progress = new Progress();           

            if (radioButtonMonth.Checked)
            {
                CalculateMonthAsync();
            }
            else CalculateWeekAsync();           
        }

        private void StopCalculate()
        {
            _cars.StopCalculate();
            _kpi.StopCalculate();
            _videotapes.StopCalculate();
            _statistics.StopCalculate();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            StopCalculate();
        }

        private void ImportCars()
        {
            listBoxPath.Items[0] = _cars.Import(listBoxPath.Items[0].ToString());
        }

        private void ImportKPI()
        {
            if (radioButtonMonth.Checked)
            {
                listBoxPath.Items[2] = _kpi.Import(listBoxPath.Items[2].ToString());
            }
            else ShowFormError("При выборе формата 'Статистика за неделю' загрузка данного файла не требуется", 67, "");
        }

        private void ImportVideotapes()
        {
            listBoxPath.Items[4] = _videotapes.Import(listBoxPath.Items[4].ToString());
        }

        private void ImportStatistics()
        {
            listBoxPath.Items[6] = _statistics.Import(listBoxPath.Items[6].ToString());
        }

        private void Btn30m_Click(object sender, EventArgs e)
        {
            ImportCars();
        }      

        private void BtnKPI_Click(object sender, EventArgs e)
        {
            ImportKPI();
        }

        private void BtnVideotapes_Click(object sender, EventArgs e)
        {
            ImportVideotapes();
        }

        private void BtnStatistics_Click(object sender, EventArgs e)
        {
            ImportStatistics();
        }

        private void radioButtonMonth_CheckedChanged(object sender, EventArgs e)
        {
            comboBoxMonth.Enabled = true;
            tbMonth.Enabled = true;
            dTP1.Enabled = false;
            dTP2.Enabled = false;

            listBoxPath.Items[2] = "KPI                    | " + _kpi.Path;
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
                    listBoxPath.Items[2] = _kpi.DragDrop("KPI", _date.ListMonth, file, listBoxPath.Items[2].ToString());
                }
                listBoxPath.Items[0] = _cars.DragDrop("30", "ПАРКОВКА", file, listBoxPath.Items[0].ToString());
                listBoxPath.Items[4] = _videotapes.DragDrop("ЗАПРОС", file, listBoxPath.Items[4].ToString());
                listBoxPath.Items[6] = _statistics.DragDrop("СТАТИСТИКА", file, listBoxPath.Items[6].ToString());                
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

        private void UpdateItem_Click(object sender, EventArgs e)
        {
            CreateSourceData();

            listBoxPath.Items[0] = _cars.FillingPaths(listBoxPath.Items[0].ToString(), "");
            if (radioButtonMonth.Checked)
            {
                listBoxPath.Items[2] = _kpi.FillingPaths(listBoxPath.Items[2].ToString(), "");
            }
            listBoxPath.Items[4] = _videotapes.FillingPaths(listBoxPath.Items[4].ToString(), "");
            listBoxPath.Items[6] = _statistics.FillingPaths(listBoxPath.Items[6].ToString(), "");
        }

        private void ImportCarsItem_Click(object sender, EventArgs e)
        {
            ImportCars();
        }

        private void ImportKPIItem_Click(object sender, EventArgs e)
        {
            ImportKPI();
        }

        private void ImportVideotapesItem_Click(object sender, EventArgs e)
        {
            ImportVideotapes();
        }

        private void ImportStatisticsItem_Click(object sender, EventArgs e)
        {
            ImportStatistics();
        }

        private void SettingsItem_Click(object sender, EventArgs e)
        {
            ShowSettingsForm();
        }

        private void ShowSettingsForm()
        {
            var form = new FormSettings(new Settings(_settings.Debug));
            if (form.ShowDialog() == DialogResult.OK)
            {
                if (_settings.Debug != form.Settings.Debug)
                {
                    if (form.Settings.Debug)
                    {
                        _cars.CurrentRowChanged += Document_CurrentRowChanged;
                        _kpi.CurrentRowChanged += Document_CurrentRowChanged;
                        _videotapes.CurrentRowChanged += Document_CurrentRowChanged;
                        _statistics.CurrentRowChanged += Document_CurrentRowChanged;
                        SetOnOtherUi += SetOnDebugUi;
                        SetOffOtherUi += SetOffDebugUi;
                    }
                    else
                    {
                        _cars.CurrentRowChanged -= Document_CurrentRowChanged;
                        _kpi.CurrentRowChanged -= Document_CurrentRowChanged;
                        _videotapes.CurrentRowChanged -= Document_CurrentRowChanged;
                        _statistics.CurrentRowChanged -= Document_CurrentRowChanged;
                        SetOnOtherUi -= SetOnDebugUi;
                        SetOffOtherUi -= SetOffDebugUi;
                    }
                }

                _settings = form.Settings;
            }
        }

        private void ExitItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void ManualItem_Click(object sender, EventArgs e)
        {
            string path = "Подсчёт статистики. Инструкция по эксплуатации для ГПУ ОТБ.docx";

            try
            {
                Process.Start(path);
            }
            catch
            {
                ShowFormError("Не удалось открыть инструкцию", 27, "");
            }

        }

        private void AboutItem_Click(object sender, EventArgs e)
        {
            var form = new FormAbout();
            form.Show();
        }
    }
}
