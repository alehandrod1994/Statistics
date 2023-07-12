using System;
using System.Windows.Forms;

namespace Statistics
{
    public partial class FormSettings : Form
    {
		public FormSettings(Settings settings)
		{
			InitializeComponent();

			Settings = settings ?? new Settings();

            if (Settings.Debug)
			{
				rbDebugOn.Checked = true;
			}
			else
			{
				rbDebugOff.Checked = true;
			}
		}

		public Settings Settings { get; private set; }

		private void RbDebugOn_CheckedChanged(object sender, EventArgs e)
		{
			Settings.Debug = true;
		}

		private void RbDebugOff_CheckedChanged(object sender, EventArgs e)
		{
			Settings.Debug = false;
		}

		private void BtnOK_Click(object sender, EventArgs e)
        {
			DialogResult = DialogResult.OK;
		}

		private void BtnOK_MouseEnter(object sender, EventArgs e)
		{
			btnOK.Image = Properties.Resources.btn_save_enter;
		}

		private void BtnOK_MouseLeave(object sender, EventArgs e)
		{
			btnOK.Image = Properties.Resources.btn_save_normal;
		}

		private void BtnOK_MouseDown(object sender, MouseEventArgs e)
		{
			btnOK.Image = Properties.Resources.btn_save_down;
		}       
    }
}
