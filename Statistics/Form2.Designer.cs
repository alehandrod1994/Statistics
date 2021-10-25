namespace Statistics
{
    partial class FormSuccessfully
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormSuccessfully));
            this.labelPath3 = new System.Windows.Forms.Label();
            this.labelPath1 = new System.Windows.Forms.Label();
            this.labelPath2 = new System.Windows.Forms.Label();
            this.btnClose = new System.Windows.Forms.PictureBox();
            this.btnOpenFolder = new System.Windows.Forms.PictureBox();
            this.btnOpen = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.btnClose)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnOpenFolder)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnOpen)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // labelPath3
            // 
            this.labelPath3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelPath3.Location = new System.Drawing.Point(81, 65);
            this.labelPath3.Name = "labelPath3";
            this.labelPath3.Size = new System.Drawing.Size(320, 73);
            this.labelPath3.TabIndex = 6;
            this.labelPath3.Text = "E:\\Диск C\\Пользователи\\Documents\\Visual Studio 2015\\Projects\\SectorCounting\\Secto" +
    "rCounting\\bin\\Debug";
            this.labelPath3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // labelPath1
            // 
            this.labelPath1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelPath1.Location = new System.Drawing.Point(81, 21);
            this.labelPath1.Name = "labelPath1";
            this.labelPath1.Size = new System.Drawing.Size(320, 27);
            this.labelPath1.TabIndex = 5;
            this.labelPath1.Text = "Статистика успешно подсчитана!";
            this.labelPath1.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // labelPath2
            // 
            this.labelPath2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.labelPath2.Location = new System.Drawing.Point(81, 48);
            this.labelPath2.Name = "labelPath2";
            this.labelPath2.Size = new System.Drawing.Size(320, 17);
            this.labelPath2.TabIndex = 10;
            this.labelPath2.Text = "Расположение файла:";
            this.labelPath2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // btnClose
            // 
            this.btnClose.Image = global::Statistics.Properties.Resources.btn_close_normal;
            this.btnClose.Location = new System.Drawing.Point(326, 130);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(118, 30);
            this.btnClose.TabIndex = 14;
            this.btnClose.TabStop = false;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            this.btnClose.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnClose_MouseDown);
            this.btnClose.MouseEnter += new System.EventHandler(this.btnClose_MouseEnter);
            this.btnClose.MouseLeave += new System.EventHandler(this.btnClose_MouseLeave);
            // 
            // btnOpenFolder
            // 
            this.btnOpenFolder.Image = ((System.Drawing.Image)(resources.GetObject("btnOpenFolder.Image")));
            this.btnOpenFolder.Location = new System.Drawing.Point(183, 130);
            this.btnOpenFolder.Name = "btnOpenFolder";
            this.btnOpenFolder.Size = new System.Drawing.Size(118, 30);
            this.btnOpenFolder.TabIndex = 13;
            this.btnOpenFolder.TabStop = false;
            this.btnOpenFolder.Click += new System.EventHandler(this.btnOpenFolder_Click);
            this.btnOpenFolder.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnOpenFolder_MouseDown);
            this.btnOpenFolder.MouseEnter += new System.EventHandler(this.btnOpenFolder_MouseEnter);
            this.btnOpenFolder.MouseLeave += new System.EventHandler(this.btnOpenFolder_MouseLeave);
            // 
            // btnOpen
            // 
            this.btnOpen.Image = ((System.Drawing.Image)(resources.GetObject("btnOpen.Image")));
            this.btnOpen.Location = new System.Drawing.Point(39, 130);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(118, 30);
            this.btnOpen.TabIndex = 12;
            this.btnOpen.TabStop = false;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            this.btnOpen.MouseDown += new System.Windows.Forms.MouseEventHandler(this.btnOpen_MouseDown);
            this.btnOpen.MouseEnter += new System.EventHandler(this.btnOpen_MouseEnter);
            this.btnOpen.MouseLeave += new System.EventHandler(this.btnOpen_MouseLeave);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Statistics.Properties.Resources.iconfinder_f_check_256_282474;
            this.pictureBox1.Location = new System.Drawing.Point(15, 31);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(64, 64);
            this.pictureBox1.TabIndex = 11;
            this.pictureBox1.TabStop = false;
            // 
            // FormSuccessfully
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.ClientSize = new System.Drawing.Size(483, 172);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOpenFolder);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.labelPath2);
            this.Controls.Add(this.labelPath3);
            this.Controls.Add(this.labelPath1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FormSuccessfully";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Уведомление";
            this.Load += new System.EventHandler(this.Form2_Load);
            ((System.ComponentModel.ISupportInitialize)(this.btnClose)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnOpenFolder)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.btnOpen)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label labelPath3;
        private System.Windows.Forms.Label labelPath1;
        private System.Windows.Forms.Label labelPath2;
        private System.Windows.Forms.PictureBox btnOpen;
        private System.Windows.Forms.PictureBox btnOpenFolder;
        private System.Windows.Forms.PictureBox btnClose;
    }
}