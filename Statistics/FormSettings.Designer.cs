
namespace Statistics
{
    partial class FormSettings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormSettings));
            this.label1 = new System.Windows.Forms.Label();
            this.rbDebugOn = new System.Windows.Forms.RadioButton();
            this.rbDebugOff = new System.Windows.Forms.RadioButton();
            this.btnOK = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.btnOK)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(35, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Отладчик";
            // 
            // rbDebugOn
            // 
            this.rbDebugOn.AutoSize = true;
            this.rbDebugOn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.rbDebugOn.Location = new System.Drawing.Point(146, 27);
            this.rbDebugOn.Name = "rbDebugOn";
            this.rbDebugOn.Size = new System.Drawing.Size(46, 19);
            this.rbDebugOn.TabIndex = 1;
            this.rbDebugOn.Text = "Вкл";
            this.rbDebugOn.UseVisualStyleBackColor = true;
            this.rbDebugOn.CheckedChanged += new System.EventHandler(this.RbDebugOn_CheckedChanged);
            // 
            // rbDebugOff
            // 
            this.rbDebugOff.AutoSize = true;
            this.rbDebugOff.Checked = true;
            this.rbDebugOff.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.rbDebugOff.Location = new System.Drawing.Point(213, 27);
            this.rbDebugOff.Name = "rbDebugOff";
            this.rbDebugOff.Size = new System.Drawing.Size(55, 19);
            this.rbDebugOff.TabIndex = 2;
            this.rbDebugOff.TabStop = true;
            this.rbDebugOff.Text = "Выкл";
            this.rbDebugOff.UseVisualStyleBackColor = true;
            this.rbDebugOff.CheckedChanged += new System.EventHandler(this.RbDebugOff_CheckedChanged);
            // 
            // btnOK
            // 
            this.btnOK.Image = global::Statistics.Properties.Resources.btn_save_normal;
            this.btnOK.Location = new System.Drawing.Point(100, 200);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(118, 30);
            this.btnOK.TabIndex = 20;
            this.btnOK.TabStop = false;
            this.btnOK.Click += new System.EventHandler(this.BtnOK_Click);
            this.btnOK.MouseDown += new System.Windows.Forms.MouseEventHandler(this.BtnOK_MouseDown);
            this.btnOK.MouseEnter += new System.EventHandler(this.BtnOK_MouseEnter);
            this.btnOK.MouseLeave += new System.EventHandler(this.BtnOK_MouseLeave);
            // 
            // FormSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(309, 244);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.rbDebugOff);
            this.Controls.Add(this.rbDebugOn);
            this.Controls.Add(this.label1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройки";
            ((System.ComponentModel.ISupportInitialize)(this.btnOK)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton rbDebugOn;
        private System.Windows.Forms.RadioButton rbDebugOff;
        private System.Windows.Forms.PictureBox btnOK;
    }
}