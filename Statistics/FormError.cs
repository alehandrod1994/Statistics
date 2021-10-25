using System;
using System.Windows.Forms;

namespace Statistics
{
    public partial class FormError : Form
    {
        public FormError()
        {
            InitializeComponent();          
        }

        private void FormError_Load(object sender, EventArgs e)
        {
            Form1 Form1 = (Form1)Owner;
            labelErrorTitle.Text = Form1.errorTitle;
            labelErrorTitle.Height = Form1.errorTitleHeight;
            labelErrorDescription.Text = Form1.errorDescription;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnOK_MouseEnter(object sender, EventArgs e)
        {
            btnOK.Image = Properties.Resources.btn_ok_enter;
        }

        private void btnOK_MouseLeave(object sender, EventArgs e)
        {
            btnOK.Image = Properties.Resources.btn_ok_normal;
        }

        private void btnOK_MouseDown(object sender, MouseEventArgs e)
        {
            btnOK.Image = Properties.Resources.btn_ok_down;
        }
    }
}
