namespace Statistics
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.comboBoxMonth = new System.Windows.Forms.ComboBox();
            this.tbMonth = new System.Windows.Forms.TextBox();
            this.radioButtonMonth = new System.Windows.Forms.RadioButton();
            this.radioButtonWeek = new System.Windows.Forms.RadioButton();
            this.listBoxPath = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dTP1 = new System.Windows.Forms.DateTimePicker();
            this.dTP2 = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.labelProgress = new System.Windows.Forms.Label();
            this.imgProgressBar_100 = new System.Windows.Forms.PictureBox();
            this.btnCancel = new System.Windows.Forms.PictureBox();
            this.btnStatistics = new System.Windows.Forms.PictureBox();
            this.btnVideotapes = new System.Windows.Forms.PictureBox();
            this.btnKPI = new System.Windows.Forms.PictureBox();
            this.btn30m = new System.Windows.Forms.PictureBox();
            this.btnCalculate = new System.Windows.Forms.PictureBox();
            this.imgProgressBar_0 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.imgProgressBar_100)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnCancel)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnStatistics)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnVideotapes)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnKPI)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btn30m)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnCalculate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.imgProgressBar_0)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // comboBoxMonth
            // 
            this.comboBoxMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxMonth.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.comboBoxMonth.ForeColor = System.Drawing.SystemColors.ControlText;
            this.comboBoxMonth.FormattingEnabled = true;
            this.comboBoxMonth.Items.AddRange(new object[] {
            "Январь",
            "Февраль",
            "Март",
            "Апрель",
            "Май",
            "Июнь",
            "Июль",
            "Август",
            "Сентябрь",
            "Октябрь",
            "Ноябрь",
            "Декабрь"});
            this.comboBoxMonth.Location = new System.Drawing.Point(161, 62);
            this.comboBoxMonth.Name = "comboBoxMonth";
            this.comboBoxMonth.Size = new System.Drawing.Size(100, 24);
            this.comboBoxMonth.TabIndex = 0;
            // 
            // tbMonth
            // 
            this.tbMonth.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbMonth.ForeColor = System.Drawing.SystemColors.ControlText;
            this.tbMonth.Location = new System.Drawing.Point(267, 62);
            this.tbMonth.Name = "tbMonth";
            this.tbMonth.Size = new System.Drawing.Size(40, 22);
            this.tbMonth.TabIndex = 1;
            this.tbMonth.Text = "2020";
            // 
            // radioButtonMonth
            // 
            this.radioButtonMonth.AutoSize = true;
            this.radioButtonMonth.Checked = true;
            this.radioButtonMonth.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButtonMonth.Location = new System.Drawing.Point(161, 29);
            this.radioButtonMonth.Name = "radioButtonMonth";
            this.radioButtonMonth.Size = new System.Drawing.Size(163, 20);
            this.radioButtonMonth.TabIndex = 15;
            this.radioButtonMonth.TabStop = true;
            this.radioButtonMonth.Text = "Статистика за месяц";
            this.radioButtonMonth.UseVisualStyleBackColor = true;
            this.radioButtonMonth.CheckedChanged += new System.EventHandler(this.radioButtonMonth_CheckedChanged);
            // 
            // radioButtonWeek
            // 
            this.radioButtonWeek.AutoSize = true;
            this.radioButtonWeek.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.radioButtonWeek.Location = new System.Drawing.Point(564, 29);
            this.radioButtonWeek.Name = "radioButtonWeek";
            this.radioButtonWeek.Size = new System.Drawing.Size(174, 20);
            this.radioButtonWeek.TabIndex = 16;
            this.radioButtonWeek.Text = "Статистика за неделю";
            this.radioButtonWeek.UseVisualStyleBackColor = true;
            this.radioButtonWeek.CheckedChanged += new System.EventHandler(this.radioButtonWeek_CheckedChanged);
            // 
            // listBoxPath
            // 
            this.listBoxPath.AllowDrop = true;
            this.listBoxPath.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.listBoxPath.FormattingEnabled = true;
            this.listBoxPath.HorizontalScrollbar = true;
            this.listBoxPath.ItemHeight = 16;
            this.listBoxPath.Items.AddRange(new object[] {
            "30м                   |",
            "",
            "KPI                    |",
            "",
            "Запросы        |",
            "",
            "Статистика  |"});
            this.listBoxPath.Location = new System.Drawing.Point(25, 206);
            this.listBoxPath.Name = "listBoxPath";
            this.listBoxPath.SelectionMode = System.Windows.Forms.SelectionMode.None;
            this.listBoxPath.Size = new System.Drawing.Size(829, 148);
            this.listBoxPath.TabIndex = 27;
            this.listBoxPath.DragDrop += new System.Windows.Forms.DragEventHandler(this.listBoxPath_DragDrop);
            this.listBoxPath.DragEnter += new System.Windows.Forms.DragEventHandler(this.listBoxPath_DragEnter);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(25, 184);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(285, 16);
            this.label1.TabIndex = 29;
            this.label1.Text = "Выберите файлы или перетащите их сюда:";
            // 
            // dTP1
            // 
            this.dTP1.Enabled = false;
            this.dTP1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dTP1.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dTP1.Location = new System.Drawing.Point(589, 52);
            this.dTP1.Name = "dTP1";
            this.dTP1.Size = new System.Drawing.Size(107, 22);
            this.dTP1.TabIndex = 33;
            this.dTP1.Value = new System.DateTime(2020, 7, 6, 16, 38, 0, 0);
            // 
            // dTP2
            // 
            this.dTP2.Enabled = false;
            this.dTP2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.dTP2.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dTP2.Location = new System.Drawing.Point(589, 78);
            this.dTP2.Name = "dTP2";
            this.dTP2.Size = new System.Drawing.Size(107, 22);
            this.dTP2.TabIndex = 34;
            this.dTP2.Value = new System.DateTime(2020, 7, 12, 16, 44, 0, 0);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(564, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(15, 16);
            this.label2.TabIndex = 35;
            this.label2.Text = "с";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(564, 82);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(24, 16);
            this.label3.TabIndex = 36;
            this.label3.Text = "по";
            // 
            // labelProgress
            // 
            this.labelProgress.AutoSize = true;
            this.labelProgress.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelProgress.Location = new System.Drawing.Point(25, 365);
            this.labelProgress.Name = "labelProgress";
            this.labelProgress.Size = new System.Drawing.Size(285, 16);
            this.labelProgress.TabIndex = 38;
            this.labelProgress.Text = "Выберите файлы или перетащите их сюда:";
            this.labelProgress.Visible = false;
            // 
            // imgProgressBar_100
            // 
            this.imgProgressBar_100.Image = global::Statistics.Properties.Resources.progressBar_100_final;
            this.imgProgressBar_100.Location = new System.Drawing.Point(26, 390);
            this.imgProgressBar_100.Name = "imgProgressBar_100";
            this.imgProgressBar_100.Size = new System.Drawing.Size(827, 10);
            this.imgProgressBar_100.TabIndex = 56;
            this.imgProgressBar_100.TabStop = false;
            this.imgProgressBar_100.Visible = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Image = global::Statistics.Properties.Resources.btn_cancel_normal;
            this.btnCancel.Location = new System.Drawing.Point(342, 413);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(200, 57);
            this.btnCancel.TabIndex = 54;
            this.btnCancel.TabStop = false;
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            this.btnCancel.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnCancel_MouseDown);
            this.btnCancel.MouseEnter += new System.EventHandler(this.btnCancel_MouseEnter);
            this.btnCancel.MouseLeave += new System.EventHandler(this.btnCancel_MouseLeave);
            // 
            // btnStatistics
            // 
            this.btnStatistics.Image = global::Statistics.Properties.Resources.btn_statistics_normal;
            this.btnStatistics.Location = new System.Drawing.Point(654, 124);
            this.btnStatistics.Name = "btnStatistics";
            this.btnStatistics.Size = new System.Drawing.Size(200, 45);
            this.btnStatistics.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.btnStatistics.TabIndex = 53;
            this.btnStatistics.TabStop = false;
            this.btnStatistics.Click += new System.EventHandler(this.btnStatistics_Click);
            this.btnStatistics.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnStatistics_MouseDown);
            this.btnStatistics.MouseEnter += new System.EventHandler(this.btnStatistics_MouseEnter);
            this.btnStatistics.MouseLeave += new System.EventHandler(this.btnStatistics_MouseLeave);
            // 
            // btnVideotapes
            // 
            this.btnVideotapes.Image = global::Statistics.Properties.Resources.btn_videotapes_normal;
            this.btnVideotapes.Location = new System.Drawing.Point(444, 124);
            this.btnVideotapes.Name = "btnVideotapes";
            this.btnVideotapes.Size = new System.Drawing.Size(200, 45);
            this.btnVideotapes.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.btnVideotapes.TabIndex = 52;
            this.btnVideotapes.TabStop = false;
            this.btnVideotapes.Click += new System.EventHandler(this.btnVideotapes_Click);
            this.btnVideotapes.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnVideotapes_MouseDown);
            this.btnVideotapes.MouseEnter += new System.EventHandler(this.btnVideotapes_MouseEnter);
            this.btnVideotapes.MouseLeave += new System.EventHandler(this.btnVideotapes_MouseLeave);
            // 
            // btnKPI
            // 
            this.btnKPI.Image = global::Statistics.Properties.Resources.btn_kpi_normal;
            this.btnKPI.Location = new System.Drawing.Point(235, 124);
            this.btnKPI.Name = "btnKPI";
            this.btnKPI.Size = new System.Drawing.Size(200, 45);
            this.btnKPI.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.btnKPI.TabIndex = 51;
            this.btnKPI.TabStop = false;
            this.btnKPI.Click += new System.EventHandler(this.btnKPI_Click);
            this.btnKPI.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnKPI_MouseDown);
            this.btnKPI.MouseEnter += new System.EventHandler(this.btnKPI_MouseEnter);
            this.btnKPI.MouseLeave += new System.EventHandler(this.btnKPI_MouseLeave);
            // 
            // btn30m
            // 
            this.btn30m.Image = global::Statistics.Properties.Resources.btn_cars_normal;
            this.btn30m.Location = new System.Drawing.Point(25, 124);
            this.btn30m.Name = "btn30m";
            this.btn30m.Size = new System.Drawing.Size(200, 45);
            this.btn30m.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.btn30m.TabIndex = 50;
            this.btn30m.TabStop = false;
            this.btn30m.Click += new System.EventHandler(this.btn30m_Click);
            this.btn30m.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btn30m_MouseDown);
            this.btn30m.MouseEnter += new System.EventHandler(this.btn30m_MouseEnter);
            this.btn30m.MouseLeave += new System.EventHandler(this.btn30m_MouseLeave);
            // 
            // btnCalculate
            // 
            this.btnCalculate.Image = global::Statistics.Properties.Resources.btn_calculate_normal;
            this.btnCalculate.Location = new System.Drawing.Point(342, 413);
            this.btnCalculate.Name = "btnCalculate";
            this.btnCalculate.Size = new System.Drawing.Size(200, 57);
            this.btnCalculate.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.btnCalculate.TabIndex = 49;
            this.btnCalculate.TabStop = false;
            this.btnCalculate.Click += new System.EventHandler(this.btnCalculate_Click);
            this.btnCalculate.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnCalculate_MouseDown);
            this.btnCalculate.MouseEnter += new System.EventHandler(this.btnCalculate_MouseEnter);
            this.btnCalculate.MouseLeave += new System.EventHandler(this.btnCalculate_MouseLeave);
            // 
            // imgProgressBar_0
            // 
            this.imgProgressBar_0.Image = global::Statistics.Properties.Resources.progressBar_0;
            this.imgProgressBar_0.Location = new System.Drawing.Point(25, 389);
            this.imgProgressBar_0.Name = "imgProgressBar_0";
            this.imgProgressBar_0.Size = new System.Drawing.Size(829, 12);
            this.imgProgressBar_0.TabIndex = 55;
            this.imgProgressBar_0.TabStop = false;
            this.imgProgressBar_0.Visible = false;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Statistics.Properties.Resources.btn_statistics_normal;
            this.pictureBox1.Location = new System.Drawing.Point(334, 219);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(200, 45);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 57;
            this.pictureBox1.TabStop = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Green;
            this.ClientSize = new System.Drawing.Size(868, 482);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.imgProgressBar_100);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnStatistics);
            this.Controls.Add(this.btnVideotapes);
            this.Controls.Add(this.btnKPI);
            this.Controls.Add(this.btn30m);
            this.Controls.Add(this.btnCalculate);
            this.Controls.Add(this.labelProgress);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dTP2);
            this.Controls.Add(this.dTP1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBoxPath);
            this.Controls.Add(this.radioButtonWeek);
            this.Controls.Add(this.radioButtonMonth);
            this.Controls.Add(this.tbMonth);
            this.Controls.Add(this.comboBoxMonth);
            this.Controls.Add(this.imgProgressBar_0);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Подсчёт статистики";
            ((System.ComponentModel.ISupportInitialize)(this.imgProgressBar_100)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnCancel)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnStatistics)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnVideotapes)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnKPI)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btn30m)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnCalculate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.imgProgressBar_0)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox comboBoxMonth;
        private System.Windows.Forms.TextBox tbMonth;
        private System.Windows.Forms.RadioButton radioButtonMonth;
        private System.Windows.Forms.RadioButton radioButtonWeek;
        private System.Windows.Forms.ListBox listBoxPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dTP1;
        private System.Windows.Forms.DateTimePicker dTP2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label labelProgress;
        private System.Windows.Forms.PictureBox btnCalculate;
        private System.Windows.Forms.PictureBox btn30m;
        private System.Windows.Forms.PictureBox btnKPI;
        private System.Windows.Forms.PictureBox btnVideotapes;
        private System.Windows.Forms.PictureBox btnStatistics;
        private System.Windows.Forms.PictureBox btnCancel;
        private System.Windows.Forms.PictureBox imgProgressBar_0;
        private System.Windows.Forms.PictureBox imgProgressBar_100;
        private System.Windows.Forms.PictureBox pictureBox1;
    }
}

