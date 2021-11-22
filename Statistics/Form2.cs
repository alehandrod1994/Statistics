using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace Statistics
{
    public partial class FormSuccessfully : Form
    {
        public FormSuccessfully()
        {
            InitializeComponent();         
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            Form1 Form1 = (Form1)Owner;
            labelPath3.Text = Form1.NewPath;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Hide();
            Form1 Form1 = new Form1();
            Form1.Show();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            string path = labelPath3.Text;
            try
            {
                if (File.Exists(path))
                    Process.Start(path);
                else
                    MessageBox.Show("Файл не найден");
            }
            catch
            {
                MessageBox.Show("Ошибка при открытии файла");
            }
           
        }

        private void btnOpenFolder_Click(object sender, EventArgs e)
        {
            Process PrFolder = new Process();
            ProcessStartInfo psi = new ProcessStartInfo();
            string file = labelPath3.Text;
            psi.CreateNoWindow = true;
            psi.WindowStyle = ProcessWindowStyle.Normal;
            psi.FileName = "explorer";
            psi.Arguments = @"/n, /select, " + file;
            PrFolder.StartInfo = psi;
            PrFolder.Start();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnOpen_MouseEnter(object sender, EventArgs e)
        {
            btnOpen.Image = Properties.Resources.btn_open_enter;
        }

        private void btnOpen_MouseLeave(object sender, EventArgs e)
        {
            btnOpen.Image = Properties.Resources.btn_open_normal;
        }

        private void btnOpen_MouseDown(object sender, MouseEventArgs e)
        {
            btnOpen.Image = Properties.Resources.btn_open_down;
        }

        private void btnOpenFolder_MouseEnter(object sender, EventArgs e)
        {
            btnOpenFolder.Image = Properties.Resources.btn_openFolder_enter;
        }

        private void btnOpenFolder_MouseLeave(object sender, EventArgs e)
        {
            btnOpenFolder.Image = Properties.Resources.btn_openFolder_normal;
        }

        private void btnOpenFolder_MouseDown(object sender, MouseEventArgs e)
        {
            btnOpenFolder.Image = Properties.Resources.btn_openFolder_down;
        }

        private void btnClose_MouseEnter(object sender, EventArgs e)
        {
            btnClose.Image = Properties.Resources.btn_close_enter;
        }

        private void btnClose_MouseLeave(object sender, EventArgs e)
        {
            btnClose.Image = Properties.Resources.btn_close_normal;
        }

        private void btnClose_MouseDown(object sender, MouseEventArgs e)
        {
            btnClose.Image = Properties.Resources.btn_close_down;
        }
    }
}
