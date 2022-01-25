using FontAwesome.Sharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tesalia_Redes_App
{
    public partial class Form4 : Form
    {
        public Form4()
        {
            InitializeComponent();
        }

        int pass = 0;
        private void VerPass_Click(object sender, EventArgs e)
        {
            if (pass == 0)
            {
                pass = 1;
                VerPass.IconChar = IconChar.Eye;
                StartPass.UseSystemPasswordChar = false;
            }
            else
            {
                pass = 0;
                VerPass.IconChar = IconChar.EyeSlash;
                StartPass.UseSystemPasswordChar = true;
            }
        }

        private void Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void MaxMin_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void Entrar_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.PassStart == StartPass.Text)
            {
                Hide();
                Form1 f1 = new Form1();
                f1.ShowDialog();
                Close();
            }
            else
            {
                MessageBox.Show("Contraseña incorrecta", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            Properties.Settings.Default.Upgrade();
            Properties.Settings.Default.Save();
            Properties.Settings.Default.Reload();

            if (Properties.Settings.Default.ActivarPassCC == "0")
            {
                this.Hide();
                Form1 f1 = new Form1();
                f1.ShowDialog();
                Close();
            }
            else if (string.IsNullOrEmpty(Properties.Settings.Default.PassStart))
            {
                this.Hide();
                Form1 f1 = new Form1();
                f1.ShowDialog();
                Close();
            }
        }

        private void StartPass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (Properties.Settings.Default.PassStart == StartPass.Text)
                {
                    this.Hide();
                    Form1 f1 = new Form1();
                    f1.ShowDialog();
                    Close();
                }
                else
                {
                    MessageBox.Show("Contraseña incorrecta", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void Form4_Shown(object sender, EventArgs e)
        {
            this.Focus();
            StartPass.Focus();
        }
    }
}
