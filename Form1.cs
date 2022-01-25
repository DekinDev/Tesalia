using FontAwesome.Sharp;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Point = System.Drawing.Point;
using Rectangle = System.Drawing.Rectangle;

namespace Tesalia_Redes_App
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            DateTime now = DateTime.Now;
            string nox = now.ToString("MMMM");
            Years1.Text = now.ToString("yyyy");
            Years2.Text = Years1.Text;
            char[] letters = nox.ToCharArray();
            letters[0] = char.ToUpper(letters[0]);
            string construct = "";
            foreach (char ch in letters)
            {
                construct += ch.ToString();
                Meses1.Text = construct;
                Meses2.Text = construct;
            }

            this.DoubleBuffered = true;
            this.SetStyle(ControlStyles.ResizeRedraw, true);

            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Resources") == false)
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Resources");
            }
            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Update") == false)
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Update");
            }
            if (Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Jornadas") == false)
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Jornadas");
            }
        }

        #region DECLARACIONES GENERALES
        int FormLoaded = 0;
        string MesPreSelected;
        int MesSelected = 0;
        int YearSelected = 0;
        readonly OpenFileDialog openFileDialog1 = new OpenFileDialog();
        string RutaFirmaS = "";
        string FileSend = "0";
        #endregion

        #region VENTANA
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();

        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hWnd, int wMsg, int wParam, int lParam);

        private void Titulo_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }
        private void PanelUP2_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void Close_Click(object sender, EventArgs e)
        {
            Close();
        }
        private void Minim_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void MaxMin_Click(object sender, EventArgs e)
        {
            if(this.WindowState == FormWindowState.Normal)
            {
                this.WindowState = FormWindowState.Maximized;
                MaxMin.IconChar = IconChar.WindowRestore;
            }
            else if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;
                
                MaxMin.IconChar = IconChar.WindowMaximize;
            }
        }
        private void PanelUP2_DoubleClick(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                this.WindowState = FormWindowState.Maximized;
                MaxMin.IconChar = IconChar.WindowRestore;
            }
            else if (this.WindowState == FormWindowState.Maximized)
            {
                this.WindowState = FormWindowState.Normal;

                MaxMin.IconChar = IconChar.WindowMaximize;
            }
        }

        private const int cGrip = 16;
        private const int cCaption = 32;

        protected override void OnPaint(PaintEventArgs e)
        {
            Rectangle rc = new Rectangle(this.ClientSize.Width - cGrip, this.ClientSize.Height - cGrip, cGrip, cGrip);
            ControlPaint.DrawSizeGrip(e.Graphics, this.BackColor, rc);
            rc = new Rectangle(0, 0, this.ClientSize.Width, cCaption);
            e.Graphics.FillRectangle(Brushes.DarkBlue, rc);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x84)
            {
                Point pos = new Point(m.LParam.ToInt32());
                pos = this.PointToClient(pos);
                if (pos.Y < cCaption)
                {
                    m.Result = (IntPtr)2;
                    return;
                }
                if (pos.X >= this.ClientSize.Width - cGrip && pos.Y >= this.ClientSize.Height - cGrip)
                {
                    m.Result = (IntPtr)17;
                    return;
                }
            }
            base.WndProc(ref m);
        }
        #endregion

        #region APARTADO REGISTRO DE JORNADA
        private void EZMode_Click(object sender, EventArgs e)
        {
            Pages.SelectedIndex = 3;
            SubPages1.SelectedIndex = 0;
        }
        private void HardMode_Click(object sender, EventArgs e)
        {
            Pages.SelectedIndex = 2;
            SubPages1.SelectedIndex = 1;
        }
        private void CrearJornada_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.Name != "")
            {
                if (Properties.Settings.Default.Documento != "")
                {
                    if (Properties.Settings.Default.SeguridadS != "")
                    {
                        if (Properties.Settings.Default.HorasdeJornada != "")
                        {
                            if (Properties.Settings.Default.RutaFirma != "")
                            {
                                Properties.Settings.Default.Installing = "1";
                                Properties.Settings.Default.Save();
                                Form2 fm2 = new Form2();
                                fm2.Show();
                                FileSend = "0";
                                WriteToExcel();
                            }
                            else
                            {
                                MessageBox.Show("No hay ninguna firma cargada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Es necesario indicar las horas de la jornada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se ha introducido documento de seguridad en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No se ha introducido el DNI en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("No se ha introducido el nombre completo en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void GuardarJornada_Click(object sender, EventArgs e)
        {
            SaveJornada();
        }
        private void BorrarJornada_Click(object sender, EventArgs e)
        {
            CleanJornada();
        }
        private void CrearEnviarJornada_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.Name != "")
            {
                if (Properties.Settings.Default.Documento != "")
                {
                    if (Properties.Settings.Default.SeguridadS != "")
                    {
                        if (Properties.Settings.Default.HorasdeJornada != "")
                        {
                            if (Properties.Settings.Default.RutaFirma != "")
                            {
                                Properties.Settings.Default.Installing = "1";
                                Properties.Settings.Default.Save();
                                Form2 fm2 = new Form2();
                                fm2.Show();
                                FileSend = "1";
                                WriteToExcel();

                                Properties.Settings.Default.CorreoType = "0";
                                Properties.Settings.Default.Save();
                                Form3 fm3 = new Form3("Registro de Jornada: " + MesPreSelected, "Registro de la Jornada de: " + Properties.Settings.Default.Name + " de " + MesPreSelected);
                                fm3.Show();
                            }
                            else
                            {
                                MessageBox.Show("No hay ninguna firma cargada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Es necesario indicar las horas de la jornada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se ha introducido documento de seguridad en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No se ha introducido el DNI en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("No se ha introducido el nombre completo en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region APARTADO REGISTRO MINIJORNADA
        private void SaveMiniJornada_Click(object sender, EventArgs e)
        {
            int IDDay = ComboDays.SelectedIndex;

            if (IDDay == 0)
            {
                EM1.Text = EM0.Text;
                SM1.Text = SM0.Text;
                ET1.Text = ET0.Text;
                ST1.Text = ST0.Text;
                Extra1.Checked = Extra0.Checked;
            }
            if (IDDay == 1)
            {
                EM2.Text = EM0.Text;
                SM2.Text = SM0.Text;
                ET2.Text = ET0.Text;
                ST2.Text = ST0.Text;
                Extra2.Checked = Extra0.Checked;
            }
            if (IDDay == 2)
            {
                EM3.Text = EM0.Text;
                SM3.Text = SM0.Text;
                ET3.Text = ET0.Text;
                ST3.Text = ST0.Text;
                Extra3.Checked = Extra0.Checked;
            }
            if (IDDay == 3)
            {
                EM4.Text = EM0.Text;
                SM4.Text = SM0.Text;
                ET4.Text = ET0.Text;
                ST4.Text = ST0.Text;
                Extra4.Checked = Extra0.Checked;
            }
            if (IDDay == 4)
            {
                EM5.Text = EM0.Text;
                SM5.Text = SM0.Text;
                ET5.Text = ET0.Text;
                ST5.Text = ST0.Text;
                Extra5.Checked = Extra0.Checked;
            }
            if (IDDay == 5)
            {
                EM6.Text = EM0.Text;
                SM6.Text = SM0.Text;
                ET6.Text = ET0.Text;
                ST6.Text = ST0.Text;
                Extra6.Checked = Extra0.Checked;
            }
            if (IDDay == 6)
            {
                EM7.Text = EM0.Text;
                SM7.Text = SM0.Text;
                ET7.Text = ET0.Text;
                ST7.Text = ST0.Text;
                Extra7.Checked = Extra0.Checked;
            }
            if (IDDay == 7)
            {
                EM8.Text = EM0.Text;
                SM8.Text = SM0.Text;
                ET8.Text = ET0.Text;
                ST8.Text = ST0.Text;
                Extra8.Checked = Extra0.Checked;
            }
            if (IDDay == 8)
            {
                EM9.Text = EM0.Text;
                SM9.Text = SM0.Text;
                ET9.Text = ET0.Text;
                ST9.Text = ST0.Text;
                Extra9.Checked = Extra0.Checked;
            }
            if (IDDay == 9)
            {
                EM10.Text = EM0.Text;
                SM10.Text = SM0.Text;
                ET10.Text = ET0.Text;
                ST10.Text = ST0.Text;
                Extra10.Checked = Extra0.Checked;
            }
            if (IDDay == 10)
            {
                EM11.Text = EM0.Text;
                SM11.Text = SM0.Text;
                ET11.Text = ET0.Text;
                ST11.Text = ST0.Text;
                Extra11.Checked = Extra0.Checked;
            }
            if (IDDay == 11)
            {
                EM12.Text = EM0.Text;
                SM12.Text = SM0.Text;
                ET12.Text = ET0.Text;
                ST12.Text = ST0.Text;
                Extra12.Checked = Extra0.Checked;
            }
            if (IDDay == 12)
            {
                EM13.Text = EM0.Text;
                SM13.Text = SM0.Text;
                ET13.Text = ET0.Text;
                ST13.Text = ST0.Text;
                Extra13.Checked = Extra0.Checked;
            }
            if (IDDay == 13)
            {
                EM14.Text = EM0.Text;
                SM14.Text = SM0.Text;
                ET14.Text = ET0.Text;
                ST14.Text = ST0.Text;
                Extra14.Checked = Extra0.Checked;
            }
            if (IDDay == 14)
            {
                EM15.Text = EM0.Text;
                SM15.Text = SM0.Text;
                ET15.Text = ET0.Text;
                ST15.Text = ST0.Text;
                Extra15.Checked = Extra0.Checked;
            }
            if (IDDay == 15)
            {
                EM16.Text = EM0.Text;
                SM16.Text = SM0.Text;
                ET16.Text = ET0.Text;
                ST16.Text = ST0.Text;
                Extra16.Checked = Extra0.Checked;
            }
            if (IDDay == 16)
            {
                EM17.Text = EM0.Text;
                SM17.Text = SM0.Text;
                ET17.Text = ET0.Text;
                ST17.Text = ST0.Text;
                Extra17.Checked = Extra0.Checked;
            }
            if (IDDay == 17)
            {
                EM18.Text = EM0.Text;
                SM18.Text = SM0.Text;
                ET18.Text = ET0.Text;
                ST18.Text = ST0.Text;
                Extra18.Checked = Extra0.Checked;
            }
            if (IDDay == 18)
            {
                EM19.Text = EM0.Text;
                SM19.Text = SM0.Text;
                ET19.Text = ET0.Text;
                ST19.Text = ST0.Text;
                Extra19.Checked = Extra0.Checked;
            }
            if (IDDay == 19)
            {
                EM20.Text = EM0.Text;
                SM20.Text = SM0.Text;
                ET20.Text = ET0.Text;
                ST20.Text = ST0.Text;
                Extra20.Checked = Extra0.Checked;
            }
            if (IDDay == 20)
            {
                EM21.Text = EM0.Text;
                SM21.Text = SM0.Text;
                ET21.Text = ET0.Text;
                ST21.Text = ST0.Text;
                Extra21.Checked = Extra0.Checked;
            }
            if (IDDay == 21)
            {
                EM22.Text = EM0.Text;
                SM22.Text = SM0.Text;
                ET22.Text = ET0.Text;
                ST22.Text = ST0.Text;
                Extra22.Checked = Extra0.Checked;
            }
            if (IDDay == 22)
            {
                EM23.Text = EM0.Text;
                SM23.Text = SM0.Text;
                ET23.Text = ET0.Text;
                ST23.Text = ST0.Text;
                Extra23.Checked = Extra0.Checked;
            }
            if (IDDay == 23)
            {
                EM24.Text = EM0.Text;
                SM24.Text = SM0.Text;
                ET24.Text = ET0.Text;
                ST24.Text = ST0.Text;
                Extra24.Checked = Extra0.Checked;
            }
            if (IDDay == 24)
            {
                EM25.Text = EM0.Text;
                SM25.Text = SM0.Text;
                ET25.Text = ET0.Text;
                ST25.Text = ST0.Text;
                Extra25.Checked = Extra0.Checked;
            }
            if (IDDay == 25)
            {
                EM26.Text = EM0.Text;
                SM26.Text = SM0.Text;
                ET26.Text = ET0.Text;
                ST26.Text = ST0.Text;
                Extra26.Checked = Extra0.Checked;
            }
            if (IDDay == 26)
            {
                EM27.Text = EM0.Text;
                SM27.Text = SM0.Text;
                ET27.Text = ET0.Text;
                ST27.Text = ST0.Text;
                Extra27.Checked = Extra0.Checked;
            }
            if (IDDay == 27)
            {
                EM28.Text = EM0.Text;
                SM28.Text = SM0.Text;
                ET28.Text = ET0.Text;
                ST28.Text = ST0.Text;
                Extra28.Checked = Extra0.Checked;
            }
            if (IDDay == 28)
            {
                EM29.Text = EM0.Text;
                SM29.Text = SM0.Text;
                ET29.Text = ET0.Text;
                ST29.Text = ST0.Text;
                Extra29.Checked = Extra0.Checked;
            }
            if (IDDay == 29)
            {
                EM30.Text = EM0.Text;
                SM30.Text = SM0.Text;
                ET30.Text = ET0.Text;
                ST30.Text = ST0.Text;
                Extra30.Checked = Extra0.Checked;
            }
            if (IDDay == 30)
            {
                EM31.Text = EM0.Text;
                SM31.Text = SM0.Text;
                ET31.Text = ET0.Text;
                ST31.Text = ST0.Text;
                Extra31.Checked = Extra0.Checked;
            }
            SaveJornada();
        }
        private void CleanMiniJornada_Click(object sender, EventArgs e)
        {
            EM0.Clear();
            SM0.Clear();
            ET0.Clear();
            ST0.Clear();
            Extra0.Checked = false;
        }
        private void CrearMiniJornada_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.Name != "")
            {
                if (Properties.Settings.Default.Documento != "")
                {
                    if (Properties.Settings.Default.SeguridadS != "")
                    {
                        if (Properties.Settings.Default.HorasdeJornada != "")
                        {
                            if (Properties.Settings.Default.RutaFirma != "")
                            {
                                Properties.Settings.Default.Installing = "1";
                                Properties.Settings.Default.Save();
                                Form2 fm2 = new Form2();
                                fm2.Show();
                                FileSend = "0";
                                WriteToExcel();
                            }
                            else
                            {
                                MessageBox.Show("No hay ninguna firma cargada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Es necesario indicar las horas de la jornada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se ha introducido documento de seguridad en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No se ha introducido el DNI en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("No se ha introducido el nombre completo en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void CrearEnviarMiniJornada_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.Name != "")
            {
                if (Properties.Settings.Default.Documento != "")
                {
                    if (Properties.Settings.Default.SeguridadS != "")
                    {
                        if (Properties.Settings.Default.HorasdeJornada != "")
                        {
                            if (Properties.Settings.Default.RutaFirma != "")
                            {
                                Properties.Settings.Default.Installing = "1";
                                Properties.Settings.Default.Save();
                                Form2 fm2 = new Form2();
                                fm2.Show();
                                FileSend = "1";
                                WriteToExcel();

                                Properties.Settings.Default.CorreoType = "0";
                                Properties.Settings.Default.Save();
                                Form3 fm3 = new Form3("Registro de Jornada de " + MesPreSelected, "Registro de la Jornada de: " + Properties.Settings.Default.Name + " de " + MesPreSelected);
                                fm3.Show();
                            }
                            else
                            {
                                MessageBox.Show("No hay ninguna firma cargada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Es necesario indicar las horas de la jornada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se ha introducido documento de seguridad en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No se ha introducido el DNI en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("No se ha introducido el nombre completo en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        #endregion

        #region MENU IZQUIERDO
        private void HomeBTN_Click(object sender, EventArgs e)
        {
            Pages.SelectedIndex = 0;
            Position.Location = new Point(HomeBTN.Location.X, HomeBTN.Location.Y);
        }

        private void AccountVTN_Click(object sender, EventArgs e)
        {
            Pages.SelectedIndex = 1;
            Position.Location = new Point(AccountVTN.Location.X, AccountVTN.Location.Y);
        }

        private void JornadaBTN_Click(object sender, EventArgs e)
        {
            Pages.SelectedIndex = 2;
            Position.Location = new Point(JornadaBTN.Location.X, JornadaBTN.Location.Y);
        }

        private void CorreoBTN_Click(object sender, EventArgs e)
        {
            if (Meses1.SelectedIndex != 9999999)
            {
                MessageBox.Show("Esta función no está disponible.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Pages.SelectedIndex = 4;
                Position.Location = new Point(CorreoBTN.Location.X, CorreoBTN.Location.Y);
            }
        }

        private void SettingsBTN_Click(object sender, EventArgs e)
        {
            Pages.SelectedIndex = 5;
            Position.Location = new Point(SettingsBTN.Location.X, SettingsBTN.Location.Y);
        }
        private void NominasBTN_Click(object sender, EventArgs e)
        {
            if(Meses1.SelectedIndex != 9999999)
            {
                MessageBox.Show("Esta función no está disponible.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                Pages.SelectedIndex = 6;
                Position.Location = new Point(NominasBTN.Location.X, NominasBTN.Location.Y);
            }
            
        }
        private void JornadasBTN_Click(object sender, EventArgs e)
        {
            Pages.SelectedIndex = 7;
            Position.Location = new Point(JornadaBTN.Location.X, JornadaBTN.Location.Y);
            DirectoryInfo di = new DirectoryInfo(Application.StartupPath + @"\Jornadas\");
            FileInfo[] files = di.GetFiles("*.xlsx");
            foreach (var file in files)
            {
                JornadaFiles.Items.Add(file);
            }
        }
        private void JornadaBTN_MouseEnter(object sender, EventArgs e)
        {
            timer1.Start();
            timer2.Stop();
        }

        private void JornadaBTN_MouseLeave(object sender, EventArgs e)
        {
            timer2.Start();
        }

        private void JornadasBTN_MouseEnter(object sender, EventArgs e)
        {
            timer2.Stop();
        }

        private void JornadasBTN_MouseLeave(object sender, EventArgs e)
        {
            timer1.Stop();
            timer2.Start();
        }
        private void SendHExtra_MouseEnter(object sender, EventArgs e)
        {
            timer2.Stop();
        }
        private void SendHExtra_MouseLeave(object sender, EventArgs e)
        {
            timer1.Stop();
            timer2.Start();
        }
        private void Timer1_Tick(object sender, EventArgs e)
        {
            JornadasBTN.Visible = true;
            SendHExtra.Visible = true;
            timer1.Stop();
        }
        private void Timer2_Tick(object sender, EventArgs e)
        {
            JornadasBTN.Visible = false;
            SendHExtra.Visible = false;
            timer2.Stop();
        }
        #endregion

        #region RELLENO DE LAS TABLAS FORM
        private void Meses1_SelectedIndexChanged(object sender, EventArgs e)
        {
            MesPreSelected = Meses1.Text;
            YearSelected = Convert.ToInt32(Years1.Text);
            Meses2.Text = Meses1.Text;
            Years2.Text = Years1.Text;
            ColorDays();
            if (FormLoaded == 1)
            {
                FechaConstruct();
                CleanJornada();
            }
        }
        private void Years1_SelectedIndexChanged(object sender, EventArgs e)
        {
            MesPreSelected = Meses1.Text;
            YearSelected = Convert.ToInt32(Years1.Text);
            Years2.Text = Years1.Text;
            Meses2.Text = Meses1.Text;
            ColorDays();
            if (FormLoaded == 1)
            {
                FechaConstruct();
                CleanJornada();
            }
        }
        private void Meses2_SelectedIndexChanged(object sender, EventArgs e)
        {
            MesPreSelected = Meses2.Text;
            Meses1.Text = Meses2.Text;
            Years1.Text = Years2.Text;
            YearSelected = Convert.ToInt32(Years1.Text);
            ColorDays();
            if (FormLoaded == 1)
            {
                FechaConstruct();
                CleanJornada();
            }
        }
        private void Years2_SelectedIndexChanged(object sender, EventArgs e)
        {
            MesPreSelected = Meses2.Text;
            Meses1.Text = Meses2.Text;
            Years1.Text = Years2.Text;
            YearSelected = Convert.ToInt32(Years1.Text);
            
            ColorDays();
            if (FormLoaded == 1)
            {
                FechaConstruct();
                CleanJornada();
            }
        }
        private void ComboDays_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadJornada();
            List<string> Chekeds = Properties.Settings.Default.Checks.Split(new char[] { ',' }).ToList();

            if (Chekeds[ComboDays.SelectedIndex] == "Checked")
            {
                Extra0.Checked = true;
            }
            else
            {
                Extra0.Checked = false;
            }
        }
        void ColorDays()
        {
            Day1.ForeColor = Color.White;
            Day2.ForeColor = Color.White;
            Day3.ForeColor = Color.White;
            Day4.ForeColor = Color.White;
            Day5.ForeColor = Color.White;
            Day6.ForeColor = Color.White;
            Day7.ForeColor = Color.White;
            Day8.ForeColor = Color.White;
            Day9.ForeColor = Color.White;
            Day10.ForeColor = Color.White;
            Day11.ForeColor = Color.White;
            Day12.ForeColor = Color.White;
            Day13.ForeColor = Color.White;
            Day14.ForeColor = Color.White;
            Day15.ForeColor = Color.White;
            Day16.ForeColor = Color.White;
            Day17.ForeColor = Color.White;
            Day18.ForeColor = Color.White;
            Day19.ForeColor = Color.White;
            Day20.ForeColor = Color.White;
            Day21.ForeColor = Color.White;
            Day22.ForeColor = Color.White;
            Day23.ForeColor = Color.White;
            Day24.ForeColor = Color.White;
            Day25.ForeColor = Color.White;
            Day26.ForeColor = Color.White;
            Day27.ForeColor = Color.White;
            Day28.ForeColor = Color.White;
            Day29.ForeColor = Color.White;
            Day30.ForeColor = Color.White;
            Day31.ForeColor = Color.White;
        }
        void FechaConstruct()
        {
            if (MesPreSelected == "Enero")
            {
                MesSelected = 1;
            }
            else if (MesPreSelected == "Febrero")
            {
                MesSelected = 2;
            }
            else if (MesPreSelected == "Marzo")
            {
                MesSelected = 3;
            }
            else if (MesPreSelected == "Abril")
            {
                MesSelected = 4;
            }
            else if (MesPreSelected == "Mayo")
            {
                MesSelected = 5;
            }
            else if (MesPreSelected == "Junio")
            {
                MesSelected = 6;
            }
            else if (MesPreSelected == "Julio")
            {
                MesSelected = 7;
            }
            else if (MesPreSelected == "Agosto")
            {
                MesSelected = 8;
            }
            else if (MesPreSelected == "Septiembre")
            {
                MesSelected = 9;
            }
            else if (MesPreSelected == "Octubre")
            {
                MesSelected = 10;
            }
            else if (MesPreSelected == "Noviembre")
            {
                MesSelected = 11;
            }
            else if (MesPreSelected == "Diciembre")
            {
                MesSelected = 12;
            }

            DayMonth26.Visible = true;
            DayMonth27.Visible = true;
            DayMonth28.Visible = true;
            DayMonth29.Visible = true;
            DayMonth30.Visible = true;
            DayMonth31.Visible = true;

            int Years = DateTime.DaysInMonth(Convert.ToInt32(YearSelected), MesSelected);
            if (Years == 25)
            {
                DayMonth26.Visible = false;
                DayMonth27.Visible = false;
                DayMonth28.Visible = false;
                DayMonth29.Visible = false;
                DayMonth30.Visible = false;
                DayMonth31.Visible = false;
            }
            else if (Years == 26)
            {
                DayMonth27.Visible = false;
                DayMonth28.Visible = false;
                DayMonth29.Visible = false;
                DayMonth30.Visible = false;
                DayMonth31.Visible = false;

            }
            else if (Years == 27)
            {
                DayMonth28.Visible = false;
                DayMonth29.Visible = false;
                DayMonth30.Visible = false;
                DayMonth31.Visible = false;

            }
            else if (Years == 28)
            {
                DayMonth29.Visible = false;
                DayMonth30.Visible = false;
                DayMonth31.Visible = false;

            }
            else if (Years == 29)
            {
                DayMonth30.Visible = false;
                DayMonth31.Visible = false;

            }
            else if (Years == 30)
            {
                DayMonth31.Visible = false;
            }


            CultureInfo ci = new CultureInfo("ES-ES");
            DateTime now = DateTime.Now;
            string Today = now.ToString("dddd", ci) + " - " + now.ToString("dd");
            ComboDays.Items.Clear();

            DateTime Days1 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 1);
            Day1.Text = Days1.ToString("dddd", ci) + " - " + Days1.ToString("dd");
            Day1.Text = FirstCharToUpper(Day1.Text);
            ComboDays.Items.Add(Day1.Text);
            if (Today == Days1.ToString("dddd", ci) + " - " + Days1.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days2 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 2);
            Day2.Text = Days2.ToString("dddd", ci) + " - " + Days2.ToString("dd");
            Day2.Text = FirstCharToUpper(Day2.Text);
            ComboDays.Items.Add(Day2.Text);
            if (Today == Days2.ToString("dddd", ci) + " - " + Days2.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days3 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 3);
            Day3.Text = Days3.ToString("dddd", ci) + " - " + Days3.ToString("dd");
            Day3.Text = FirstCharToUpper(Day3.Text);
            ComboDays.Items.Add(Day3.Text);
            if (Today == Days3.ToString("dddd", ci) + " - " + Days3.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days4 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 4);
            Day4.Text = Days4.ToString("dddd", ci) + " - " + Days4.ToString("dd");
            Day4.Text = FirstCharToUpper(Day4.Text);
            ComboDays.Items.Add(Day4.Text);
            if (Today == Days4.ToString("dddd", ci) + " - " + Days4.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days5 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 5);
            Day5.Text = Days5.ToString("dddd", ci) + " - " + Days5.ToString("dd");
            Day5.Text = FirstCharToUpper(Day5.Text);
            ComboDays.Items.Add(Day5.Text);
            if (Today == Days5.ToString("dddd", ci) + " - " + Days5.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days6 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 6);
            Day6.Text = Days6.ToString("dddd", ci) + " - " + Days6.ToString("dd");
            Day6.Text = FirstCharToUpper(Day6.Text);
            ComboDays.Items.Add(Day6.Text);
            if (Today == Days6.ToString("dddd", ci) + " - " + Days6.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days7 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 7);
            Day7.Text = Days7.ToString("dddd", ci) + " - " + Days7.ToString("dd");
            Day7.Text = FirstCharToUpper(Day7.Text);
            ComboDays.Items.Add(Day7.Text);
            if (Today == Days7.ToString("dddd", ci) + " - " + Days7.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days8 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 8);
            Day8.Text = Days8.ToString("dddd", ci) + " - " + Days8.ToString("dd");
            Day8.Text = FirstCharToUpper(Day8.Text);
            ComboDays.Items.Add(Day8.Text);
            if (Today == Days8.ToString("dddd", ci) + " - " + Days8.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days9 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 9);
            Day9.Text = Days9.ToString("dddd", ci) + " - " + Days9.ToString("dd");
            Day9.Text = FirstCharToUpper(Day9.Text);
            ComboDays.Items.Add(Day9.Text);
            if (Today == Days9.ToString("dddd", ci) + " - " + Days9.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days10 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 10);
            Day10.Text = Days10.ToString("dddd", ci) + " - " + Days10.ToString("dd");
            Day10.Text = FirstCharToUpper(Day10.Text);
            ComboDays.Items.Add(Day10.Text);
            if (Today == Days10.ToString("dddd", ci) + " - " + Days10.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days11 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 11);
            Day11.Text = Days11.ToString("dddd", ci) + " - " + Days11.ToString("dd");
            Day11.Text = FirstCharToUpper(Day11.Text);
            ComboDays.Items.Add(Day11.Text);
            if (Today == Days11.ToString("dddd", ci) + " - " + Days11.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days12 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 12);
            Day12.Text = Days12.ToString("dddd", ci) + " - " + Days12.ToString("dd");
            Day12.Text = FirstCharToUpper(Day12.Text);
            ComboDays.Items.Add(Day12.Text);
            if (Today == Days12.ToString("dddd", ci) + " - " + Days12.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days13 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 13);
            Day13.Text = Days13.ToString("dddd", ci) + " - " + Days13.ToString("dd");
            Day13.Text = FirstCharToUpper(Day13.Text);
            ComboDays.Items.Add(Day13.Text);
            if (Today == Days13.ToString("dddd", ci) + " - " + Days13.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days14 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 14);
            Day14.Text = Days14.ToString("dddd", ci) + " - " + Days14.ToString("dd");
            Day14.Text = FirstCharToUpper(Day14.Text);
            ComboDays.Items.Add(Day14.Text);
            if (Today == Days14.ToString("dddd", ci) + " - " + Days14.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days15 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 15);
            Day15.Text = Days15.ToString("dddd", ci) + " - " + Days15.ToString("dd");
            Day15.Text = FirstCharToUpper(Day15.Text);
            ComboDays.Items.Add(Day15.Text);
            if (Today == Days15.ToString("dddd", ci) + " - " + Days15.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days16 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 16);
            Day16.Text = Days16.ToString("dddd", ci) + " - " + Days16.ToString("dd");
            Day16.Text = FirstCharToUpper(Day16.Text);
            ComboDays.Items.Add(Day16.Text);
            if (Today == Days16.ToString("dddd", ci) + " - " + Days16.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days17 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 17);
            Day17.Text = Days17.ToString("dddd", ci) + " - " + Days17.ToString("dd");
            Day17.Text = FirstCharToUpper(Day17.Text);
            ComboDays.Items.Add(Day17.Text);
            if (Today == Days17.ToString("dddd", ci) + " - " + Days17.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days18 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 18);
            Day18.Text = Days18.ToString("dddd", ci) + " - " + Days18.ToString("dd");
            Day18.Text = FirstCharToUpper(Day18.Text);
            ComboDays.Items.Add(Day18.Text);
            if (Today == Days18.ToString("dddd", ci) + " - " + Days18.ToString("dd"))
            {
                ComboDays.Text = Today;
            }


            DateTime Days19 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 19);
            Day19.Text = Days19.ToString("dddd", ci) + " - " + Days19.ToString("dd");
            Day19.Text = FirstCharToUpper(Day19.Text);
            ComboDays.Items.Add(Day19.Text);
            if (Today == Days19.ToString("dddd", ci) + " - " + Days19.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days20 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 20);
            Day20.Text = Days20.ToString("dddd", ci) + " - " + Days20.ToString("dd");
            Day20.Text = FirstCharToUpper(Day20.Text);
            ComboDays.Items.Add(Day20.Text);
            if (Today == Days20.ToString("dddd", ci) + " - " + Days20.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days21 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 21);
            Day21.Text = Days21.ToString("dddd", ci) + " - " + Days21.ToString("dd");
            Day21.Text = FirstCharToUpper(Day21.Text);
            ComboDays.Items.Add(Day21.Text);
            if (Today == Days21.ToString("dddd", ci) + " - " + Days21.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days22 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 22);
            Day22.Text = Days22.ToString("dddd", ci) + " - " + Days22.ToString("dd");
            Day22.Text = FirstCharToUpper(Day22.Text);
            ComboDays.Items.Add(Day22.Text);
            if (Today == Days22.ToString("dddd", ci) + " - " + Days22.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days23 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 23);
            Day23.Text = Days23.ToString("dddd", ci) + " - " + Days23.ToString("dd");
            Day23.Text = FirstCharToUpper(Day23.Text);
            ComboDays.Items.Add(Day23.Text);
            if (Today == Days23.ToString("dddd", ci) + " - " + Days23.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            DateTime Days24 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 24);
            Day24.Text = Days24.ToString("dddd", ci) + " - " + Days24.ToString("dd");
            Day24.Text = FirstCharToUpper(Day24.Text);
            ComboDays.Items.Add(Day24.Text);
            if (Today == Days24.ToString("dddd", ci) + " - " + Days24.ToString("dd"))
            {
                ComboDays.Text = Today;
            }

            if (Convert.ToInt32(Years) == 25)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                Day25.Text = Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd");
                Day25.Text = FirstCharToUpper(Day25.Text);
                ComboDays.Items.Add(Day25.Text);
                if (Today == Days1.ToString("dddd", ci) + " - " + Days1.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }
            }

            if (Convert.ToInt32(Years) == 26)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                Day25.Text = Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd");
                Day25.Text = FirstCharToUpper(Day25.Text);
                ComboDays.Items.Add(Day25.Text);
                if (Today == Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                Day26.Text = Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd");
                Day26.Text = FirstCharToUpper(Day26.Text);
                ComboDays.Items.Add(Day26.Text);
                if (Today == Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }
            }

            if (Convert.ToInt32(Years) == 27)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                Day25.Text = Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd");
                Day25.Text = FirstCharToUpper(Day25.Text);
                ComboDays.Items.Add(Day25.Text);
                if (Today == Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                Day26.Text = Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd");
                Day26.Text = FirstCharToUpper(Day26.Text);
                ComboDays.Items.Add(Day26.Text);
                if (Today == Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                Day27.Text = Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd");
                Day27.Text = FirstCharToUpper(Day27.Text);
                ComboDays.Items.Add(Day27.Text);
                if (Today == Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }
            }

            if (Convert.ToInt32(Years) == 28)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                Day25.Text = Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd");
                Day25.Text = FirstCharToUpper(Day25.Text);
                ComboDays.Items.Add(Day25.Text);
                if (Today == Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                Day26.Text = Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd");
                Day26.Text = FirstCharToUpper(Day26.Text);
                ComboDays.Items.Add(Day26.Text);
                if (Today == Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                Day27.Text = Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd");
                Day27.Text = FirstCharToUpper(Day27.Text);
                ComboDays.Items.Add(Day27.Text);
                if (Today == Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days28 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 28);
                Day28.Text = Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd");
                Day28.Text = FirstCharToUpper(Day28.Text);
                ComboDays.Items.Add(Day28.Text);
                if (Today == Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }
            }

            if (Convert.ToInt32(Years) == 29)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                Day25.Text = Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd");
                Day25.Text = FirstCharToUpper(Day25.Text);
                ComboDays.Items.Add(Day25.Text);
                if (Today == Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                Day26.Text = Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd");
                Day26.Text = FirstCharToUpper(Day26.Text);
                ComboDays.Items.Add(Day26.Text);
                if (Today == Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                Day27.Text = Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd");
                Day27.Text = FirstCharToUpper(Day27.Text);
                ComboDays.Items.Add(Day27.Text);
                if (Today == Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days28 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 28);
                Day28.Text = Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd");
                Day28.Text = FirstCharToUpper(Day28.Text);
                ComboDays.Items.Add(Day28.Text);
                if (Today == Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days29 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 29);
                Day29.Text = Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd");
                Day29.Text = FirstCharToUpper(Day29.Text);
                ComboDays.Items.Add(Day29.Text);
                if (Today == Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }
            }

            if (Convert.ToInt32(Years) == 30)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                Day25.Text = Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd");
                Day25.Text = FirstCharToUpper(Day25.Text);
                ComboDays.Items.Add(Day25.Text);
                if (Today == Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                Day26.Text = Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd");
                Day26.Text = FirstCharToUpper(Day26.Text);
                ComboDays.Items.Add(Day26.Text);
                if (Today == Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                Day27.Text = Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd");
                Day27.Text = FirstCharToUpper(Day27.Text);
                ComboDays.Items.Add(Day27.Text);
                if (Today == Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days28 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 28);
                Day28.Text = Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd");
                Day28.Text = FirstCharToUpper(Day28.Text);
                ComboDays.Items.Add(Day28.Text);
                if (Today == Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days29 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 29);
                Day29.Text = Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd");
                Day29.Text = FirstCharToUpper(Day29.Text);
                ComboDays.Items.Add(Day29.Text);
                if (Today == Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days30 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 30);
                Day30.Text = Days30.ToString("dddd", ci) + " - " + Days30.ToString("dd");
                Day30.Text = FirstCharToUpper(Day30.Text);
                ComboDays.Items.Add(Day30.Text);
                if (Today == Days30.ToString("dddd", ci) + " - " + Days30.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }
            }

            if (Convert.ToInt32(Years) == 31)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                Day25.Text = Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd");
                Day25.Text = FirstCharToUpper(Day25.Text);
                ComboDays.Items.Add(Day25.Text);
                if (Today == Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                Day26.Text = Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd");
                Day26.Text = FirstCharToUpper(Day26.Text);
                ComboDays.Items.Add(Day26.Text);
                if (Today == Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                Day27.Text = Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd");
                Day27.Text = FirstCharToUpper(Day27.Text);
                ComboDays.Items.Add(Day27.Text);
                if (Today == Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days28 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 28);
                Day28.Text = Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd");
                Day28.Text = FirstCharToUpper(Day28.Text);
                ComboDays.Items.Add(Day28.Text);
                if (Today == Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days29 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 29);
                Day29.Text = Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd");
                Day29.Text = FirstCharToUpper(Day29.Text);
                ComboDays.Items.Add(Day29.Text);
                if (Today == Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days30 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 30);
                Day30.Text = Days30.ToString("dddd", ci) + " - " + Days30.ToString("dd");
                Day30.Text = FirstCharToUpper(Day30.Text);
                ComboDays.Items.Add(Day30.Text);
                if (Today == Days30.ToString("dddd", ci) + " - " + Days30.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }

                DateTime Days31 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 31);
                Day31.Text = Days31.ToString("dddd", ci) + " - " + Days31.ToString("dd");
                Day31.Text = FirstCharToUpper(Day31.Text);
                ComboDays.Items.Add(Day31.Text);
                if (Today == Days31.ToString("dddd", ci) + " - " + Days31.ToString("dd"))
                {
                    ComboDays.Text = Today;
                }
            }

            Festivos();
            LoadJornada();
        }
        public static string FirstCharToUpper(string input)
        {
            switch (input)
            {
                case null: throw new ArgumentNullException(nameof(input));
                case "": throw new ArgumentException($"{nameof(input)} cannot be empty", nameof(input));
                default: return input.First().ToString().ToUpper() + input.Substring(1);
            }
        }
        void Festivos()
        {
            if (Day1.Text.Contains("Sábado") == true)
            {
                Day1.ForeColor = Color.Red;
                EM1.Clear();
                SM1.Clear();
                ET1.Clear();
                ST1.Clear();
            }
            else if (Day1.Text.Contains("Domingo") == true)
            {
                Day1.ForeColor = Color.Red;
                EM1.Clear();
                SM1.Clear();
                ET1.Clear();
                ST1.Clear();
            }

            if (Day2.Text.Contains("Sábado") == true)
            {
                Day2.ForeColor = Color.Red;
                EM2.Clear();
                SM2.Clear();
                ET2.Clear();
                ST2.Clear();
            }
            else if (Day2.Text.Contains("Domingo") == true)
            {
                Day2.ForeColor = Color.Red;
                EM2.Clear();
                SM2.Clear();
                ET2.Clear();
                ST2.Clear();
            }

            if (Day3.Text.Contains("Sábado") == true)
            {
                Day3.ForeColor = Color.Red;
                EM3.Clear();
                SM3.Clear();
                ET3.Clear();
                ST3.Clear();
            }
            else if (Day3.Text.Contains("Domingo") == true)
            {
                Day3.ForeColor = Color.Red;
                EM3.Clear();
                SM3.Clear();
                ET3.Clear();
                ST3.Clear();
            }

            if (Day4.Text.Contains("Sábado") == true)
            {
                Day4.ForeColor = Color.Red;
                EM4.Clear();
                SM4.Clear();
                ET4.Clear();
                ST4.Clear();
            }
            else if (Day4.Text.Contains("Domingo") == true)
            {
                Day4.ForeColor = Color.Red;
                EM4.Clear();
                SM4.Clear();
                ET4.Clear();
                ST4.Clear();
            }

            if (Day5.Text.Contains("Sábado") == true)
            {
                Day5.ForeColor = Color.Red;
                EM5.Clear();
                SM5.Clear();
                ET5.Clear();
                ST5.Clear();
            }
            else if (Day5.Text.Contains("Domingo") == true)
            {
                Day5.ForeColor = Color.Red;
                EM5.Clear();
                SM5.Clear();
                ET5.Clear();
                ST5.Clear();
            }

            if (Day6.Text.Contains("Sábado") == true)
            {
                Day6.ForeColor = Color.Red;
                EM6.Clear();
                SM6.Clear();
                ET6.Clear();
                ST6.Clear();
            }
            else if (Day6.Text.Contains("Domingo") == true)
            {
                Day6.ForeColor = Color.Red;
                EM6.Clear();
                SM6.Clear();
                ET6.Clear();
                ST6.Clear();
            }

            if (Day7.Text.Contains("Sábado") == true)
            {
                Day7.ForeColor = Color.Red;
                EM7.Clear();
                SM7.Clear();
                ET7.Clear();
                ST7.Clear();
            }
            else if (Day7.Text.Contains("Domingo") == true)
            {
                Day7.ForeColor = Color.Red;
                EM7.Clear();
                SM7.Clear();
                ET7.Clear();
                ST7.Clear();
            }

            if (Day8.Text.Contains("Sábado") == true)
            {
                Day8.ForeColor = Color.Red;
                EM8.Clear();
                SM8.Clear();
                ET8.Clear();
                ST8.Clear();
            }
            else if (Day8.Text.Contains("Domingo") == true)
            {
                Day8.ForeColor = Color.Red;
                EM8.Clear();
                SM8.Clear();
                ET8.Clear();
                ST8.Clear();
            }

            if (Day9.Text.Contains("Sábado") == true)
            {
                Day9.ForeColor = Color.Red;
                EM9.Clear();
                SM9.Clear();
                ET9.Clear();
                ST9.Clear();
            }
            else if (Day9.Text.Contains("Domingo") == true)
            {
                Day9.ForeColor = Color.Red;
                EM9.Clear();
                SM9.Clear();
                ET9.Clear();
                ST9.Clear();
            }

            if (Day10.Text.Contains("Sábado") == true)
            {
                Day10.ForeColor = Color.Red;
                EM10.Clear();
                SM10.Clear();
                ET10.Clear();
                ST10.Clear();
            }
            else if (Day10.Text.Contains("Domingo") == true)
            {
                Day10.ForeColor = Color.Red;
                EM10.Clear();
                SM10.Clear();
                ET10.Clear();
                ST10.Clear();
            }

            if (Day11.Text.Contains("Sábado") == true)
            {
                Day11.ForeColor = Color.Red;
                EM11.Clear();
                SM11.Clear();
                ET11.Clear();
                ST11.Clear();
            }
            else if (Day11.Text.Contains("Domingo") == true)
            {
                Day11.ForeColor = Color.Red;
                EM11.Clear();
                SM11.Clear();
                ET11.Clear();
                ST11.Clear();
            }

            if (Day12.Text.Contains("Sábado") == true)
            {
                Day12.ForeColor = Color.Red;
                EM12.Clear();
                SM12.Clear();
                ET12.Clear();
                ST12.Clear();
            }
            else if (Day12.Text.Contains("Domingo") == true)
            {
                Day12.ForeColor = Color.Red;
                EM12.Clear();
                SM12.Clear();
                ET12.Clear();
                ST12.Clear();
            }

            if (Day13.Text.Contains("Sábado") == true)
            {
                Day13.ForeColor = Color.Red;
                EM13.Clear();
                SM13.Clear();
                ET13.Clear();
                ST13.Clear();
            }
            else if (Day13.Text.Contains("Domingo") == true)
            {
                Day13.ForeColor = Color.Red;
                EM13.Clear();
                SM13.Clear();
                ET13.Clear();
                ST13.Clear();
            }

            if (Day14.Text.Contains("Sábado") == true)
            {
                Day14.ForeColor = Color.Red;
                EM14.Clear();
                SM14.Clear();
                ET14.Clear();
                ST14.Clear();
            }
            else if (Day14.Text.Contains("Domingo") == true)
            {
                Day14.ForeColor = Color.Red;
                EM14.Clear();
                SM14.Clear();
                ET14.Clear();
                ST14.Clear();
            }

            if (Day15.Text.Contains("Sábado") == true)
            {
                Day15.ForeColor = Color.Red;
                EM15.Clear();
                SM15.Clear();
                ET15.Clear();
                ST15.Clear();
            }
            else if (Day15.Text.Contains("Domingo") == true)
            {
                Day15.ForeColor = Color.Red;
                EM15.Clear();
                SM15.Clear();
                ET15.Clear();
                ST15.Clear();
            }

            if (Day16.Text.Contains("Sábado") == true)
            {
                Day16.ForeColor = Color.Red;
                EM16.Clear();
                SM16.Clear();
                ET16.Clear();
                ST16.Clear();
            }
            else if (Day16.Text.Contains("Domingo") == true)
            {
                Day16.ForeColor = Color.Red;
                EM16.Clear();
                SM16.Clear();
                ET16.Clear();
                ST16.Clear();
            }

            if (Day17.Text.Contains("Sábado") == true)
            {
                Day17.ForeColor = Color.Red;
                EM17.Clear();
                SM17.Clear();
                ET17.Clear();
                ST17.Clear();
            }
            else if (Day17.Text.Contains("Domingo") == true)
            {
                Day17.ForeColor = Color.Red;
                EM17.Clear();
                SM17.Clear();
                ET17.Clear();
                ST17.Clear();
            }

            if (Day18.Text.Contains("Sábado") == true)
            {
                Day18.ForeColor = Color.Red;
                EM18.Clear();
                SM18.Clear();
                ET18.Clear();
                ST18.Clear();
            }
            else if (Day18.Text.Contains("Domingo") == true)
            {
                Day18.ForeColor = Color.Red;
                EM18.Clear();
                SM18.Clear();
                ET18.Clear();
                ST18.Clear();
            }

            if (Day19.Text.Contains("Sábado") == true)
            {
                Day19.ForeColor = Color.Red;
                EM19.Clear();
                SM19.Clear();
                ET19.Clear();
                ST19.Clear();
            }
            else if (Day19.Text.Contains("Domingo") == true)
            {
                Day19.ForeColor = Color.Red;
                EM19.Clear();
                SM19.Clear();
                ET19.Clear();
                ST19.Clear();
            }

            if (Day20.Text.Contains("Sábado") == true)
            {
                Day20.ForeColor = Color.Red;
                EM20.Clear();
                SM20.Clear();
                ET20.Clear();
                ST20.Clear();
            }
            else if (Day20.Text.Contains("Domingo") == true)
            {
                Day20.ForeColor = Color.Red;
                EM20.Clear();
                SM20.Clear();
                ET20.Clear();
                ST20.Clear();
            }

            if (Day21.Text.Contains("Sábado") == true)
            {
                Day21.ForeColor = Color.Red;
                EM21.Clear();
                SM21.Clear();
                ET21.Clear();
                ST21.Clear();
            }
            else if (Day21.Text.Contains("Domingo") == true)
            {
                Day21.ForeColor = Color.Red;
                EM21.Clear();
                SM21.Clear();
                ET21.Clear();
                ST21.Clear();
            }

            if (Day22.Text.Contains("Sábado") == true)
            {
                Day22.ForeColor = Color.Red;
                EM22.Clear();
                SM22.Clear();
                ET22.Clear();
                ST22.Clear();
            }
            else if (Day22.Text.Contains("Domingo") == true)
            {
                Day22.ForeColor = Color.Red;
                EM22.Clear();
                SM22.Clear();
                ET22.Clear();
                ST22.Clear();
            }

            if (Day23.Text.Contains("Sábado") == true)
            {
                Day23.ForeColor = Color.Red;
                EM23.Clear();
                SM23.Clear();
                ET23.Clear();
                ST23.Clear();
            }
            else if (Day23.Text.Contains("Domingo") == true)
            {
                Day23.ForeColor = Color.Red;
                EM23.Clear();
                SM23.Clear();
                ET23.Clear();
                ST23.Clear();
            }

            if (Day24.Text.Contains("Sábado") == true)
            {
                Day24.ForeColor = Color.Red;
                EM24.Clear();
                SM24.Clear();
                ET24.Clear();
                ST24.Clear();
            }
            else if (Day24.Text.Contains("Domingo") == true)
            {
                Day24.ForeColor = Color.Red;
                EM24.Clear();
                SM24.Clear();
                ET24.Clear();
                ST24.Clear();
            }

            if (Day25.Text.Contains("Sábado") == true)
            {
                Day25.ForeColor = Color.Red;
                EM25.Clear();
                SM25.Clear();
                ET25.Clear();
                ST25.Clear();
            }
            else if (Day25.Text.Contains("Domingo") == true)
            {
                Day25.ForeColor = Color.Red;
                EM25.Clear();
                SM25.Clear();
                ET25.Clear();
                ST25.Clear();
            }

            if (Day26.Text.Contains("Sábado") == true)
            {
                Day26.ForeColor = Color.Red;
                EM26.Clear();
                SM26.Clear();
                ET26.Clear();
                ST26.Clear();
            }
            else if (Day26.Text.Contains("Domingo") == true)
            {
                Day26.ForeColor = Color.Red;
                EM26.Clear();
                SM26.Clear();
                ET26.Clear();
                ST26.Clear();
            }

            if (Day27.Text.Contains("Sábado") == true)
            {
                Day27.ForeColor = Color.Red;
                EM27.Clear();
                SM27.Clear();
                ET27.Clear();
                ST27.Clear();
            }
            else if (Day27.Text.Contains("Domingo") == true)
            {
                Day27.ForeColor = Color.Red;
                EM27.Clear();
                SM27.Clear();
                ET27.Clear();
                ST27.Clear();
            }
            if (Day28.Text.Contains("Sábado") == true)
            {
                Day28.ForeColor = Color.Red;
                EM28.Clear();
                SM28.Clear();
                ET28.Clear();
                ST28.Clear();
            }
            else if (Day28.Text.Contains("Domingo") == true)
            {
                Day28.ForeColor = Color.Red;
                EM28.Clear();
                SM28.Clear();
                ET28.Clear();
                ST28.Clear();
            }

            if (Day29.Text.Contains("Sábado") == true)
            {
                Day29.ForeColor = Color.Red;
                EM29.Clear();
                SM29.Clear();
                ET29.Clear();
                ST29.Clear();
            }
            else if (Day29.Text.Contains("Domingo") == true)
            {
                Day29.ForeColor = Color.Red;
                EM29.Clear();
                SM29.Clear();
                ET29.Clear();
                ST29.Clear();
            }

            if (Day30.Text.Contains("Sábado") == true)
            {
                Day30.ForeColor = Color.Red;
                EM30.Clear();
                SM30.Clear();
                ET30.Clear();
                ST30.Clear();
            }
            else if (Day30.Text.Contains("Domingo") == true)
            {
                Day30.ForeColor = Color.Red;
                EM30.Clear();
                SM30.Clear();
                ET30.Clear();
                ST30.Clear();
            }

            if (Day31.Text.Contains("Sábado") == true)
            {
                Day31.ForeColor = Color.Red;
                EM31.Clear();
                SM31.Clear();
                ET31.Clear();
                ST31.Clear();
            }
            else if (Day31.Text.Contains("Domingo") == true)
            {
                Day31.ForeColor = Color.Red;
                EM31.Clear();
                SM31.Clear();
                ET31.Clear();
                ST31.Clear();
            }
        }
        void CleanJornada()
        {
            EM1.Clear();
            SM1.Clear();
            ET1.Clear();
            ST1.Clear();

            EM2.Clear();
            SM2.Clear();
            ET2.Clear();
            ST2.Clear();

            EM3.Clear();
            SM3.Clear();
            ET3.Clear();
            ST3.Clear();

            EM4.Clear();
            SM4.Clear();
            ET4.Clear();
            ST4.Clear();

            EM5.Clear();
            SM5.Clear();
            ET5.Clear();
            ST5.Clear();

            EM6.Clear();
            SM6.Clear();
            ET6.Clear();
            ST6.Clear();

            EM7.Clear();
            SM7.Clear();
            ET7.Clear();
            ST7.Clear();

            EM8.Clear();
            SM8.Clear();
            ET8.Clear();
            ST8.Clear();

            EM9.Clear();
            SM9.Clear();
            ET9.Clear();
            ST9.Clear();

            EM10.Clear();
            SM10.Clear();
            ET10.Clear();
            ST10.Clear();

            EM11.Clear();
            SM11.Clear();
            ET11.Clear();
            ST11.Clear();

            EM12.Clear();
            SM12.Clear();
            ET12.Clear();
            ST12.Clear();

            EM13.Clear();
            SM13.Clear();
            ET13.Clear();
            ST13.Clear();

            EM14.Clear();
            SM14.Clear();
            ET14.Clear();
            ST14.Clear();

            EM15.Clear();
            SM15.Clear();
            ET15.Clear();
            ST15.Clear();

            EM16.Clear();
            SM16.Clear();
            ET16.Clear();
            ST16.Clear();

            EM17.Clear();
            SM17.Clear();
            ET17.Clear();
            ST17.Clear();

            EM18.Clear();
            SM18.Clear();
            ET18.Clear();
            ST18.Clear();

            EM19.Clear();
            SM19.Clear();
            ET19.Clear();
            ST19.Clear();

            EM20.Clear();
            SM20.Clear();
            ET20.Clear();
            ST20.Clear();

            EM21.Clear();
            SM21.Clear();
            ET21.Clear();
            ST21.Clear();

            EM22.Clear();
            SM22.Clear();
            ET22.Clear();
            ST22.Clear();

            EM23.Clear();
            SM23.Clear();
            ET23.Clear();
            ST23.Clear();

            EM24.Clear();
            SM24.Clear();
            ET24.Clear();
            ST24.Clear();

            EM25.Clear();
            SM25.Clear();
            ET25.Clear();
            ST25.Clear();

            EM26.Clear();
            SM26.Clear();
            ET26.Clear();
            ST26.Clear();

            EM27.Clear();
            SM27.Clear();
            ET27.Clear();
            ST27.Clear();

            EM28.Clear();
            SM28.Clear();
            ET28.Clear();
            ST28.Clear();

            EM29.Clear();
            SM29.Clear();
            ET29.Clear();
            ST29.Clear();

            EM30.Clear();
            SM30.Clear();
            ET30.Clear();
            ST30.Clear();

            EM31.Clear();
            SM31.Clear();
            ET31.Clear();
            ST31.Clear();

            Extra1.Checked = false;
            Extra2.Checked = false;
            Extra3.Checked = false;
            Extra4.Checked = false;
            Extra5.Checked = false;
            Extra6.Checked = false;
            Extra7.Checked = false;
            Extra8.Checked = false;
            Extra9.Checked = false;
            Extra10.Checked = false;
            Extra11.Checked = false;
            Extra12.Checked = false;
            Extra13.Checked = false;
            Extra14.Checked = false;
            Extra15.Checked = false;
            Extra16.Checked = false;
            Extra17.Checked = false;
            Extra18.Checked = false;
            Extra19.Checked = false;
            Extra20.Checked = false;
            Extra21.Checked = false;
            Extra22.Checked = false;
            Extra23.Checked = false;
            Extra24.Checked = false;
            Extra25.Checked = false;
            Extra26.Checked = false;
            Extra27.Checked = false;
            Extra28.Checked = false;
            Extra29.Checked = false;
            Extra30.Checked = false;
            Extra31.Checked = false;

            EM0.Clear();
            SM0.Clear();
            ET0.Clear();
            ST0.Clear();
            Extra0.Checked = false;
        }
        void LoadJornada()
        {
            EM1.Text = Properties.Settings.Default.EM1;
            EM2.Text = Properties.Settings.Default.EM2;
            EM3.Text = Properties.Settings.Default.EM3;
            EM4.Text = Properties.Settings.Default.EM4;
            EM5.Text = Properties.Settings.Default.EM5;
            EM6.Text = Properties.Settings.Default.EM6;
            EM7.Text = Properties.Settings.Default.EM7;
            EM8.Text = Properties.Settings.Default.EM8;
            EM9.Text = Properties.Settings.Default.EM9;
            EM10.Text = Properties.Settings.Default.EM10;
            EM11.Text = Properties.Settings.Default.EM11;
            EM12.Text = Properties.Settings.Default.EM12;
            EM13.Text = Properties.Settings.Default.EM13;
            EM14.Text = Properties.Settings.Default.EM14;
            EM15.Text = Properties.Settings.Default.EM15;
            EM16.Text = Properties.Settings.Default.EM16;
            EM17.Text = Properties.Settings.Default.EM17;
            EM18.Text = Properties.Settings.Default.EM18;
            EM19.Text = Properties.Settings.Default.EM19;
            EM20.Text = Properties.Settings.Default.EM20;
            EM21.Text = Properties.Settings.Default.EM21;
            EM22.Text = Properties.Settings.Default.EM22;
            EM23.Text = Properties.Settings.Default.EM23;
            EM24.Text = Properties.Settings.Default.EM24;
            EM25.Text = Properties.Settings.Default.EM25;
            EM26.Text = Properties.Settings.Default.EM26;
            EM27.Text = Properties.Settings.Default.EM27;
            EM28.Text = Properties.Settings.Default.EM28;
            EM29.Text = Properties.Settings.Default.EM29;
            EM30.Text = Properties.Settings.Default.EM30;
            EM31.Text = Properties.Settings.Default.EM31;

            SM1.Text = Properties.Settings.Default.SM1;
            SM2.Text = Properties.Settings.Default.SM2;
            SM3.Text = Properties.Settings.Default.SM3;
            SM4.Text = Properties.Settings.Default.SM4;
            SM5.Text = Properties.Settings.Default.SM5;
            SM6.Text = Properties.Settings.Default.SM6;
            SM7.Text = Properties.Settings.Default.SM7;
            SM8.Text = Properties.Settings.Default.SM8;
            SM9.Text = Properties.Settings.Default.SM9;
            SM10.Text = Properties.Settings.Default.SM10;
            SM11.Text = Properties.Settings.Default.SM11;
            SM12.Text = Properties.Settings.Default.SM12;
            SM13.Text = Properties.Settings.Default.SM13;
            SM14.Text = Properties.Settings.Default.SM14;
            SM15.Text = Properties.Settings.Default.SM15;
            SM16.Text = Properties.Settings.Default.SM16;
            SM17.Text = Properties.Settings.Default.SM17;
            SM18.Text = Properties.Settings.Default.SM18;
            SM19.Text = Properties.Settings.Default.SM19;
            SM20.Text = Properties.Settings.Default.SM20;
            SM21.Text = Properties.Settings.Default.SM21;
            SM22.Text = Properties.Settings.Default.SM22;
            SM23.Text = Properties.Settings.Default.SM23;
            SM24.Text = Properties.Settings.Default.SM24;
            SM25.Text = Properties.Settings.Default.SM25;
            SM26.Text = Properties.Settings.Default.SM26;
            SM27.Text = Properties.Settings.Default.SM27;
            SM28.Text = Properties.Settings.Default.SM28;
            SM29.Text = Properties.Settings.Default.SM29;
            SM30.Text = Properties.Settings.Default.SM30;
            SM31.Text = Properties.Settings.Default.SM31;

            ET1.Text = Properties.Settings.Default.ET1;
            ET2.Text = Properties.Settings.Default.ET2;
            ET3.Text = Properties.Settings.Default.ET3;
            ET4.Text = Properties.Settings.Default.ET4;
            ET5.Text = Properties.Settings.Default.ET5;
            ET6.Text = Properties.Settings.Default.ET6;
            ET7.Text = Properties.Settings.Default.ET7;
            ET8.Text = Properties.Settings.Default.ET8;
            ET9.Text = Properties.Settings.Default.ET9;
            ET10.Text = Properties.Settings.Default.ET10;
            ET11.Text = Properties.Settings.Default.ET11;
            ET12.Text = Properties.Settings.Default.ET12;
            ET13.Text = Properties.Settings.Default.ET13;
            ET14.Text = Properties.Settings.Default.ET14;
            ET15.Text = Properties.Settings.Default.ET15;
            ET16.Text = Properties.Settings.Default.ET16;
            ET17.Text = Properties.Settings.Default.ET17;
            ET18.Text = Properties.Settings.Default.ET18;
            ET19.Text = Properties.Settings.Default.ET19;
            ET20.Text = Properties.Settings.Default.ET20;
            ET21.Text = Properties.Settings.Default.ET21;
            ET22.Text = Properties.Settings.Default.ET22;
            ET23.Text = Properties.Settings.Default.ET23;
            ET24.Text = Properties.Settings.Default.ET24;
            ET25.Text = Properties.Settings.Default.ET25;
            ET26.Text = Properties.Settings.Default.ET26;
            ET27.Text = Properties.Settings.Default.ET27;
            ET28.Text = Properties.Settings.Default.ET28;
            ET29.Text = Properties.Settings.Default.ET29;
            ET30.Text = Properties.Settings.Default.ET30;
            ET31.Text = Properties.Settings.Default.ET31;

            ST1.Text = Properties.Settings.Default.ST1;
            ST2.Text = Properties.Settings.Default.ST2;
            ST3.Text = Properties.Settings.Default.ST3;
            ST4.Text = Properties.Settings.Default.ST4;
            ST5.Text = Properties.Settings.Default.ST5;
            ST6.Text = Properties.Settings.Default.ST6;
            ST7.Text = Properties.Settings.Default.ST7;
            ST8.Text = Properties.Settings.Default.ST8;
            ST9.Text = Properties.Settings.Default.ST9;
            ST10.Text = Properties.Settings.Default.ST10;
            ST11.Text = Properties.Settings.Default.ST11;
            ST12.Text = Properties.Settings.Default.ST12;
            ST13.Text = Properties.Settings.Default.ST13;
            ST14.Text = Properties.Settings.Default.ST14;
            ST15.Text = Properties.Settings.Default.ST15;
            ST16.Text = Properties.Settings.Default.ST16;
            ST17.Text = Properties.Settings.Default.ST17;
            ST18.Text = Properties.Settings.Default.ST18;
            ST19.Text = Properties.Settings.Default.ST19;
            ST20.Text = Properties.Settings.Default.ST20;
            ST21.Text = Properties.Settings.Default.ST21;
            ST22.Text = Properties.Settings.Default.ST22;
            ST23.Text = Properties.Settings.Default.ST23;
            ST24.Text = Properties.Settings.Default.ST24;
            ST25.Text = Properties.Settings.Default.ST25;
            ST26.Text = Properties.Settings.Default.ST26;
            ST27.Text = Properties.Settings.Default.ST27;
            ST28.Text = Properties.Settings.Default.ST28;
            ST29.Text = Properties.Settings.Default.ST29;
            ST30.Text = Properties.Settings.Default.ST30;
            ST31.Text = Properties.Settings.Default.ST31;

            Auto1.Text = Properties.Settings.Default.CAuto1;
            Auto2.Text = Properties.Settings.Default.CAuto2;
            Auto3.Text = Properties.Settings.Default.CAuto3;
            Auto4.Text = Properties.Settings.Default.CAuto4;

            List<string> Chekeds = Properties.Settings.Default.Checks.Split(new char[] { ',' }).ToList();

            if (Chekeds[0] == "Checked")
            {
                Extra1.Checked = true;
            }
            if (Chekeds[1] == "Checked")
            {
                Extra2.Checked = true;
            }
            if (Chekeds[2] == "Checked")
            {
                Extra3.Checked = true;
            }
            if (Chekeds[3] == "Checked")
            {
                Extra4.Checked = true;
            }
            if (Chekeds[4] == "Checked")
            {
                Extra5.Checked = true;
            }
            if (Chekeds[5] == "Checked")
            {
                Extra6.Checked = true;
            }
            if (Chekeds[6] == "Checked")
            {
                Extra7.Checked = true;
            }
            if (Chekeds[7] == "Checked")
            {
                Extra8.Checked = true;
            }
            if (Chekeds[8] == "Checked")
            {
                Extra9.Checked = true;
            }
            if (Chekeds[9] == "Checked")
            {
                Extra10.Checked = true;
            }
            if (Chekeds[10] == "Checked")
            {
                Extra11.Checked = true;
            }
            if (Chekeds[11] == "Checked")
            {
                Extra12.Checked = true;
            }
            if (Chekeds[12] == "Checked")
            {
                Extra13.Checked = true;
            }
            if (Chekeds[13] == "Checked")
            {
                Extra14.Checked = true;
            }
            if (Chekeds[14] == "Checked")
            {
                Extra15.Checked = true;
            }
            if (Chekeds[15] == "Checked")
            {
                Extra16.Checked = true;
            }
            if (Chekeds[16] == "Checked")
            {
                Extra17.Checked = true;
            }
            if (Chekeds[17] == "Checked")
            {
                Extra18.Checked = true;
            }
            if (Chekeds[18] == "Checked")
            {
                Extra19.Checked = true;
            }
            if (Chekeds[19] == "Checked")
            {
                Extra20.Checked = true;
            }
            if (Chekeds[20] == "Checked")
            {
                Extra21.Checked = true;
            }
            if (Chekeds[21] == "Checked")
            {
                Extra22.Checked = true;
            }
            if (Chekeds[22] == "Checked")
            {
                Extra23.Checked = true;
            }
            if (Chekeds[23] == "Checked")
            {
                Extra24.Checked = true;
            }
            if (Chekeds[24] == "Checked")
            {
                Extra25.Checked = true;
            }
            if (Chekeds[25] == "Checked")
            {
                Extra26.Checked = true;
            }
            if (Chekeds[26] == "Checked")
            {
                Extra27.Checked = true;
            }
            if (Chekeds[27] == "Checked")
            {
                Extra28.Checked = true;
            }
            if (Chekeds[28] == "Checked")
            {
                Extra29.Checked = true;
            }
            if (Chekeds[29] == "Checked")
            {
                Extra30.Checked = true;
            }
            if (Chekeds[30] == "Checked")
            {
                Extra31.Checked = true;
            }

            CultureInfo ci = new CultureInfo("ES-ES");
            DateTime Days1 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 1);
            if (ComboDays.Text.ToLower() == (Days1.ToString("dddd", ci) + " - " + Days1.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM1;
                SM0.Text = Properties.Settings.Default.SM1;
                ET0.Text = Properties.Settings.Default.ET1;
                ST0.Text = Properties.Settings.Default.ST1;
            }
            DateTime Days2 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 2);
            if (ComboDays.Text.ToLower() == (Days2.ToString("dddd", ci) + " - " + Days2.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM2;
                SM0.Text = Properties.Settings.Default.SM2;
                ET0.Text = Properties.Settings.Default.ET2;
                ST0.Text = Properties.Settings.Default.ST2;
            }
            DateTime Days3 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 3);
            if (ComboDays.Text.ToLower() == (Days3.ToString("dddd", ci) + " - " + Days3.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM3;
                SM0.Text = Properties.Settings.Default.SM3;
                ET0.Text = Properties.Settings.Default.ET3;
                ST0.Text = Properties.Settings.Default.ST3;
            }
            DateTime Days4 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 4);
            if (ComboDays.Text.ToLower() == (Days4.ToString("dddd", ci) + " - " + Days4.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM4;
                SM0.Text = Properties.Settings.Default.SM4;
                ET0.Text = Properties.Settings.Default.ET4;
                ST0.Text = Properties.Settings.Default.ST4;
            }
            DateTime Days5 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 5);
            if (ComboDays.Text.ToLower() == (Days5.ToString("dddd", ci) + " - " + Days5.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM5;
                SM0.Text = Properties.Settings.Default.SM5;
                ET0.Text = Properties.Settings.Default.ET5;
                ST0.Text = Properties.Settings.Default.ST5;
            }
            DateTime Days6 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 6);
            if (ComboDays.Text.ToLower() == (Days6.ToString("dddd", ci) + " - " + Days6.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM6;
                SM0.Text = Properties.Settings.Default.SM6;
                ET0.Text = Properties.Settings.Default.ET6;
                ST0.Text = Properties.Settings.Default.ST6;
            }
            DateTime Days7 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 7);
            if (ComboDays.Text.ToLower() == (Days7.ToString("dddd", ci) + " - " + Days7.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM7;
                SM0.Text = Properties.Settings.Default.SM7;
                ET0.Text = Properties.Settings.Default.ET7;
                ST0.Text = Properties.Settings.Default.ST7;
            }
            DateTime Days8 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 8);
            if (ComboDays.Text.ToLower() == (Days8.ToString("dddd", ci) + " - " + Days8.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM8;
                SM0.Text = Properties.Settings.Default.SM8;
                ET0.Text = Properties.Settings.Default.ET8;
                ST0.Text = Properties.Settings.Default.ST8;
            }
            DateTime Days9 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 9);
            if (ComboDays.Text.ToLower() == (Days9.ToString("dddd", ci) + " - " + Days9.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM9;
                SM0.Text = Properties.Settings.Default.SM9;
                ET0.Text = Properties.Settings.Default.ET9;
                ST0.Text = Properties.Settings.Default.ST9;
            }
            DateTime Days10 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 10);
            if (ComboDays.Text.ToLower() == (Days10.ToString("dddd", ci) + " - " + Days10.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM10;
                SM0.Text = Properties.Settings.Default.SM10;
                ET0.Text = Properties.Settings.Default.ET10;
                ST0.Text = Properties.Settings.Default.ST10;
            }
            DateTime Days11 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 11);
            if (ComboDays.Text.ToLower() == (Days11.ToString("dddd", ci) + " - " + Days11.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM11;
                SM0.Text = Properties.Settings.Default.SM11;
                ET0.Text = Properties.Settings.Default.ET11;
                ST0.Text = Properties.Settings.Default.ST11;
            }
            DateTime Days12 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 12);
            if (ComboDays.Text.ToLower() == (Days12.ToString("dddd", ci) + " - " + Days12.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM12;
                SM0.Text = Properties.Settings.Default.SM12;
                ET0.Text = Properties.Settings.Default.ET12;
                ST0.Text = Properties.Settings.Default.ST12;
            }
            DateTime Days13 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 13);
            if (ComboDays.Text.ToLower() == (Days13.ToString("dddd", ci) + " - " + Days13.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM13;
                SM0.Text = Properties.Settings.Default.SM13;
                ET0.Text = Properties.Settings.Default.ET13;
                ST0.Text = Properties.Settings.Default.ST13;
            }
            DateTime Days14 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 14);
            if (ComboDays.Text.ToLower() == (Days14.ToString("dddd", ci) + " - " + Days14.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM14;
                SM0.Text = Properties.Settings.Default.SM14;
                ET0.Text = Properties.Settings.Default.ET14;
                ST0.Text = Properties.Settings.Default.ST14;
            }
            DateTime Days15 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 15);
            if (ComboDays.Text.ToLower() == (Days15.ToString("dddd", ci) + " - " + Days15.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM15;
                SM0.Text = Properties.Settings.Default.SM15;
                ET0.Text = Properties.Settings.Default.ET15;
                ST0.Text = Properties.Settings.Default.ST15;
            }
            DateTime Days16 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 16);
            if (ComboDays.Text.ToLower() == (Days16.ToString("dddd", ci) + " - " + Days16.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM16;
                SM0.Text = Properties.Settings.Default.SM16;
                ET0.Text = Properties.Settings.Default.ET16;
                ST0.Text = Properties.Settings.Default.ST16;
            }
            DateTime Days17 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 17);
            if (ComboDays.Text.ToLower() == (Days17.ToString("dddd", ci) + " - " + Days17.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM17;
                SM0.Text = Properties.Settings.Default.SM17;
                ET0.Text = Properties.Settings.Default.ET17;
                ST0.Text = Properties.Settings.Default.ST17;
            }
            DateTime Days18 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 18);
            if (ComboDays.Text.ToLower() == (Days18.ToString("dddd", ci) + " - " + Days18.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM18;
                SM0.Text = Properties.Settings.Default.SM18;
                ET0.Text = Properties.Settings.Default.ET18;
                ST0.Text = Properties.Settings.Default.ST18;
            }
            DateTime Days19 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 19);
            if (ComboDays.Text.ToLower() == (Days19.ToString("dddd", ci) + " - " + Days19.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM19;
                SM0.Text = Properties.Settings.Default.SM19;
                ET0.Text = Properties.Settings.Default.ET19;
                ST0.Text = Properties.Settings.Default.ST19;
            }
            DateTime Days20 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 20);
            if (ComboDays.Text.ToLower() == (Days20.ToString("dddd", ci) + " - " + Days20.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM20;
                SM0.Text = Properties.Settings.Default.SM20;
                ET0.Text = Properties.Settings.Default.ET20;
                ST0.Text = Properties.Settings.Default.ST20;
            }
            DateTime Days21 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 21);
            if (ComboDays.Text.ToLower() == (Days21.ToString("dddd", ci) + " - " + Days21.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM21;
                SM0.Text = Properties.Settings.Default.SM21;
                ET0.Text = Properties.Settings.Default.ET21;
                ST0.Text = Properties.Settings.Default.ST21;
            }
            DateTime Days22 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 22);
            if (ComboDays.Text.ToLower() == (Days22.ToString("dddd", ci) + " - " + Days22.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM22;
                SM0.Text = Properties.Settings.Default.SM22;
                ET0.Text = Properties.Settings.Default.ET22;
                ST0.Text = Properties.Settings.Default.ST22;
            }
            DateTime Days23 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 23);
            if (ComboDays.Text.ToLower() == (Days23.ToString("dddd", ci) + " - " + Days23.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM23;
                SM0.Text = Properties.Settings.Default.SM23;
                ET0.Text = Properties.Settings.Default.ET23;
                ST0.Text = Properties.Settings.Default.ST23;
            }
            DateTime Days24 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 24);
            if (ComboDays.Text.ToLower() == (Days24.ToString("dddd", ci) + " - " + Days24.ToString("dd")).ToLower())
            {
                EM0.Text = Properties.Settings.Default.EM24;
                SM0.Text = Properties.Settings.Default.SM24;
                ET0.Text = Properties.Settings.Default.ET24;
                ST0.Text = Properties.Settings.Default.ST24;
            }

            if (Convert.ToInt32(DateTime.DaysInMonth(Convert.ToInt32(YearSelected), MesSelected)) == 25)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                if (ComboDays.Text.ToLower() == (Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM25;
                    SM0.Text = Properties.Settings.Default.SM25;
                    ET0.Text = Properties.Settings.Default.ET25;
                    ST0.Text = Properties.Settings.Default.ST25;
                }
            }

            if (Convert.ToInt32(DateTime.DaysInMonth(Convert.ToInt32(YearSelected), MesSelected)) == 26)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                if (ComboDays.Text.ToLower() == (Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM25;
                    SM0.Text = Properties.Settings.Default.SM25;
                    ET0.Text = Properties.Settings.Default.ET25;
                    ST0.Text = Properties.Settings.Default.ST25;
                }
                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                if (ComboDays.Text.ToLower() == (Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM26;
                    SM0.Text = Properties.Settings.Default.SM26;
                    ET0.Text = Properties.Settings.Default.ET26;
                    ST0.Text = Properties.Settings.Default.ST26;
                }
            }

            if (Convert.ToInt32(DateTime.DaysInMonth(Convert.ToInt32(YearSelected), MesSelected)) == 27)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                if (ComboDays.Text.ToLower() == (Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM25;
                    SM0.Text = Properties.Settings.Default.SM25;
                    ET0.Text = Properties.Settings.Default.ET25;
                    ST0.Text = Properties.Settings.Default.ST25;
                }
                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                if (ComboDays.Text.ToLower() == (Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM26;
                    SM0.Text = Properties.Settings.Default.SM26;
                    ET0.Text = Properties.Settings.Default.ET26;
                    ST0.Text = Properties.Settings.Default.ST26;
                }
                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                if (ComboDays.Text.ToLower() == (Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM27;
                    SM0.Text = Properties.Settings.Default.SM27;
                    ET0.Text = Properties.Settings.Default.ET27;
                    ST0.Text = Properties.Settings.Default.ST27;
                }
            }

            if (Convert.ToInt32(DateTime.DaysInMonth(Convert.ToInt32(YearSelected), MesSelected)) == 28)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                if (ComboDays.Text.ToLower() == (Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM25;
                    SM0.Text = Properties.Settings.Default.SM25;
                    ET0.Text = Properties.Settings.Default.ET25;
                    ST0.Text = Properties.Settings.Default.ST25;
                }
                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                if (ComboDays.Text.ToLower() == (Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM26;
                    SM0.Text = Properties.Settings.Default.SM26;
                    ET0.Text = Properties.Settings.Default.ET26;
                    ST0.Text = Properties.Settings.Default.ST26;
                }
                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                if (ComboDays.Text.ToLower() == (Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM27;
                    SM0.Text = Properties.Settings.Default.SM27;
                    ET0.Text = Properties.Settings.Default.ET27;
                    ST0.Text = Properties.Settings.Default.ST27;
                }
                DateTime Days28 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 28);
                if (ComboDays.Text.ToLower() == (Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM28;
                    SM0.Text = Properties.Settings.Default.SM28;
                    ET0.Text = Properties.Settings.Default.ET28;
                    ST0.Text = Properties.Settings.Default.ST28;
                }
            }

            if (Convert.ToInt32(DateTime.DaysInMonth(Convert.ToInt32(YearSelected), MesSelected)) == 29)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                if (ComboDays.Text.ToLower() == (Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM25;
                    SM0.Text = Properties.Settings.Default.SM25;
                    ET0.Text = Properties.Settings.Default.ET25;
                    ST0.Text = Properties.Settings.Default.ST25;
                }
                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                if (ComboDays.Text.ToLower() == (Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM26;
                    SM0.Text = Properties.Settings.Default.SM26;
                    ET0.Text = Properties.Settings.Default.ET26;
                    ST0.Text = Properties.Settings.Default.ST26;
                }
                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                if (ComboDays.Text.ToLower() == (Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM27;
                    SM0.Text = Properties.Settings.Default.SM27;
                    ET0.Text = Properties.Settings.Default.ET27;
                    ST0.Text = Properties.Settings.Default.ST27;
                }
                DateTime Days28 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 28);
                if (ComboDays.Text.ToLower() == (Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM28;
                    SM0.Text = Properties.Settings.Default.SM28;
                    ET0.Text = Properties.Settings.Default.ET28;
                    ST0.Text = Properties.Settings.Default.ST28;
                }
                DateTime Days29 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 29);
                if (ComboDays.Text.ToLower() == (Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM29;
                    SM0.Text = Properties.Settings.Default.SM29;
                    ET0.Text = Properties.Settings.Default.ET29;
                    ST0.Text = Properties.Settings.Default.ST29;
                }
            }

            if (Convert.ToInt32(DateTime.DaysInMonth(Convert.ToInt32(YearSelected), MesSelected)) == 30)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                if (ComboDays.Text.ToLower() == (Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM25;
                    SM0.Text = Properties.Settings.Default.SM25;
                    ET0.Text = Properties.Settings.Default.ET25;
                    ST0.Text = Properties.Settings.Default.ST25;
                }
                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                if (ComboDays.Text.ToLower() == (Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM26;
                    SM0.Text = Properties.Settings.Default.SM26;
                    ET0.Text = Properties.Settings.Default.ET26;
                    ST0.Text = Properties.Settings.Default.ST26;
                }
                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                if (ComboDays.Text.ToLower() == (Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM27;
                    SM0.Text = Properties.Settings.Default.SM27;
                    ET0.Text = Properties.Settings.Default.ET27;
                    ST0.Text = Properties.Settings.Default.ST27;
                }
                DateTime Days28 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 28);
                if (ComboDays.Text.ToLower() == (Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM28;
                    SM0.Text = Properties.Settings.Default.SM28;
                    ET0.Text = Properties.Settings.Default.ET28;
                    ST0.Text = Properties.Settings.Default.ST28;
                }
                DateTime Days29 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 29);
                if (ComboDays.Text.ToLower() == (Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM29;
                    SM0.Text = Properties.Settings.Default.SM29;
                    ET0.Text = Properties.Settings.Default.ET29;
                    ST0.Text = Properties.Settings.Default.ST29;
                }
                DateTime Days30 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 30);
                if (ComboDays.Text.ToLower() == (Days30.ToString("dddd", ci) + " - " + Days30.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM30;
                    SM0.Text = Properties.Settings.Default.SM30;
                    ET0.Text = Properties.Settings.Default.ET30;
                    ST0.Text = Properties.Settings.Default.ST30;
                }
            }

            if (Convert.ToInt32(DateTime.DaysInMonth(Convert.ToInt32(YearSelected), MesSelected)) == 31)
            {
                DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
                if (ComboDays.Text.ToLower() == (Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM25;
                    SM0.Text = Properties.Settings.Default.SM25;
                    ET0.Text = Properties.Settings.Default.ET25;
                    ST0.Text = Properties.Settings.Default.ST25;
                }
                DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
                if (ComboDays.Text.ToLower() == (Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM26;
                    SM0.Text = Properties.Settings.Default.SM26;
                    ET0.Text = Properties.Settings.Default.ET26;
                    ST0.Text = Properties.Settings.Default.ST26;
                }
                DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
                if (ComboDays.Text.ToLower() == (Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM27;
                    SM0.Text = Properties.Settings.Default.SM27;
                    ET0.Text = Properties.Settings.Default.ET27;
                    ST0.Text = Properties.Settings.Default.ST27;
                }
                DateTime Days28 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 28);
                if (ComboDays.Text.ToLower() == (Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM28;
                    SM0.Text = Properties.Settings.Default.SM28;
                    ET0.Text = Properties.Settings.Default.ET28;
                    ST0.Text = Properties.Settings.Default.ST28;
                }
                DateTime Days29 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 29);
                if (ComboDays.Text.ToLower() == (Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM29;
                    SM0.Text = Properties.Settings.Default.SM29;
                    ET0.Text = Properties.Settings.Default.ET29;
                    ST0.Text = Properties.Settings.Default.ST29;
                }
                DateTime Days30 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 30);
                if (ComboDays.Text.ToLower() == (Days30.ToString("dddd", ci) + " - " + Days30.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM30;
                    SM0.Text = Properties.Settings.Default.SM30;
                    ET0.Text = Properties.Settings.Default.ET30;
                    ST0.Text = Properties.Settings.Default.ST30;
                }
                DateTime Days31 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 31);
                if (ComboDays.Text.ToLower() == (Days31.ToString("dddd", ci) + " - " + Days31.ToString("dd")).ToLower())
                {
                    EM0.Text = Properties.Settings.Default.EM31;
                    SM0.Text = Properties.Settings.Default.SM31;
                    ET0.Text = Properties.Settings.Default.ET31;
                    ST0.Text = Properties.Settings.Default.ST31;
                }
            }
        }
        void SaveJornada()
        {
            Properties.Settings.Default.EM1 = EM1.Text;
            Properties.Settings.Default.EM2 = EM2.Text;
            Properties.Settings.Default.EM3 = EM3.Text;
            Properties.Settings.Default.EM4 = EM4.Text;
            Properties.Settings.Default.EM5 = EM5.Text;
            Properties.Settings.Default.EM6 = EM6.Text;
            Properties.Settings.Default.EM7 = EM7.Text;
            Properties.Settings.Default.EM8 = EM8.Text;
            Properties.Settings.Default.EM9 = EM9.Text;
            Properties.Settings.Default.EM10 = EM10.Text;
            Properties.Settings.Default.EM11 = EM11.Text;
            Properties.Settings.Default.EM12 = EM12.Text;
            Properties.Settings.Default.EM13 = EM13.Text;
            Properties.Settings.Default.EM14 = EM14.Text;
            Properties.Settings.Default.EM15 = EM15.Text;
            Properties.Settings.Default.EM16 = EM16.Text;
            Properties.Settings.Default.EM17 = EM17.Text;
            Properties.Settings.Default.EM18 = EM18.Text;
            Properties.Settings.Default.EM19 = EM19.Text;
            Properties.Settings.Default.EM20 = EM20.Text;
            Properties.Settings.Default.EM21 = EM21.Text;
            Properties.Settings.Default.EM22 = EM22.Text;
            Properties.Settings.Default.EM23 = EM23.Text;
            Properties.Settings.Default.EM24 = EM24.Text;
            Properties.Settings.Default.EM25 = EM25.Text;
            Properties.Settings.Default.EM26 = EM26.Text;
            Properties.Settings.Default.EM27 = EM27.Text;
            Properties.Settings.Default.EM28 = EM28.Text;
            Properties.Settings.Default.EM29 = EM29.Text;
            Properties.Settings.Default.EM30 = EM30.Text;
            Properties.Settings.Default.EM31 = EM31.Text;

            Properties.Settings.Default.SM1 = SM1.Text;
            Properties.Settings.Default.SM2 = SM2.Text;
            Properties.Settings.Default.SM3 = SM3.Text;
            Properties.Settings.Default.SM4 = SM4.Text;
            Properties.Settings.Default.SM5 = SM5.Text;
            Properties.Settings.Default.SM6 = SM6.Text;
            Properties.Settings.Default.SM7 = SM7.Text;
            Properties.Settings.Default.SM8 = SM8.Text;
            Properties.Settings.Default.SM9 = SM9.Text;
            Properties.Settings.Default.SM10 = SM10.Text;
            Properties.Settings.Default.SM11 = SM11.Text;
            Properties.Settings.Default.SM12 = SM12.Text;
            Properties.Settings.Default.SM13 = SM13.Text;
            Properties.Settings.Default.SM14 = SM14.Text;
            Properties.Settings.Default.SM15 = SM15.Text;
            Properties.Settings.Default.SM16 = SM16.Text;
            Properties.Settings.Default.SM17 = SM17.Text;
            Properties.Settings.Default.SM18 = SM18.Text;
            Properties.Settings.Default.SM19 = SM19.Text;
            Properties.Settings.Default.SM20 = SM20.Text;
            Properties.Settings.Default.SM21 = SM21.Text;
            Properties.Settings.Default.SM22 = SM22.Text;
            Properties.Settings.Default.SM23 = SM23.Text;
            Properties.Settings.Default.SM24 = SM24.Text;
            Properties.Settings.Default.SM25 = SM25.Text;
            Properties.Settings.Default.SM26 = SM26.Text;
            Properties.Settings.Default.SM27 = SM27.Text;
            Properties.Settings.Default.SM28 = SM28.Text;
            Properties.Settings.Default.SM29 = SM29.Text;
            Properties.Settings.Default.SM30 = SM30.Text;
            Properties.Settings.Default.SM31 = SM31.Text;

            Properties.Settings.Default.ET1 = ET1.Text;
            Properties.Settings.Default.ET2 = ET2.Text;
            Properties.Settings.Default.ET3 = ET3.Text;
            Properties.Settings.Default.ET4 = ET4.Text;
            Properties.Settings.Default.ET5 = ET5.Text;
            Properties.Settings.Default.ET6 = ET6.Text;
            Properties.Settings.Default.ET7 = ET7.Text;
            Properties.Settings.Default.ET8 = ET8.Text;
            Properties.Settings.Default.ET9 = ET9.Text;
            Properties.Settings.Default.ET10 = ET10.Text;
            Properties.Settings.Default.ET11 = ET11.Text;
            Properties.Settings.Default.ET12 = ET12.Text;
            Properties.Settings.Default.ET13 = ET13.Text;
            Properties.Settings.Default.ET14 = ET14.Text;
            Properties.Settings.Default.ET15 = ET15.Text;
            Properties.Settings.Default.ET16 = ET16.Text;
            Properties.Settings.Default.ET17 = ET17.Text;
            Properties.Settings.Default.ET18 = ET18.Text;
            Properties.Settings.Default.ET19 = ET19.Text;
            Properties.Settings.Default.ET20 = ET20.Text;
            Properties.Settings.Default.ET21 = ET21.Text;
            Properties.Settings.Default.ET22 = ET22.Text;
            Properties.Settings.Default.ET23 = ET23.Text;
            Properties.Settings.Default.ET24 = ET24.Text;
            Properties.Settings.Default.ET25 = ET25.Text;
            Properties.Settings.Default.ET26 = ET26.Text;
            Properties.Settings.Default.ET27 = ET27.Text;
            Properties.Settings.Default.ET28 = ET28.Text;
            Properties.Settings.Default.ET29 = ET29.Text;
            Properties.Settings.Default.ET30 = ET30.Text;
            Properties.Settings.Default.ET31 = ET31.Text;

            Properties.Settings.Default.ST1 = ST1.Text;
            Properties.Settings.Default.ST2 = ST2.Text;
            Properties.Settings.Default.ST3 = ST3.Text;
            Properties.Settings.Default.ST4 = ST4.Text;
            Properties.Settings.Default.ST5 = ST5.Text;
            Properties.Settings.Default.ST6 = ST6.Text;
            Properties.Settings.Default.ST7 = ST7.Text;
            Properties.Settings.Default.ST8 = ST8.Text;
            Properties.Settings.Default.ST9 = ST9.Text;
            Properties.Settings.Default.ST10 = ST10.Text;
            Properties.Settings.Default.ST11 = ST11.Text;
            Properties.Settings.Default.ST12 = ST12.Text;
            Properties.Settings.Default.ST13 = ST13.Text;
            Properties.Settings.Default.ST14 = ST14.Text;
            Properties.Settings.Default.ST15 = ST15.Text;
            Properties.Settings.Default.ST16 = ST16.Text;
            Properties.Settings.Default.ST17 = ST17.Text;
            Properties.Settings.Default.ST18 = ST18.Text;
            Properties.Settings.Default.ST19 = ST19.Text;
            Properties.Settings.Default.ST20 = ST20.Text;
            Properties.Settings.Default.ST21 = ST21.Text;
            Properties.Settings.Default.ST22 = ST22.Text;
            Properties.Settings.Default.ST23 = ST23.Text;
            Properties.Settings.Default.ST24 = ST24.Text;
            Properties.Settings.Default.ST25 = ST25.Text;
            Properties.Settings.Default.ST26 = ST26.Text;
            Properties.Settings.Default.ST27 = ST27.Text;
            Properties.Settings.Default.ST28 = ST28.Text;
            Properties.Settings.Default.ST29 = ST29.Text;
            Properties.Settings.Default.ST30 = ST30.Text;
            Properties.Settings.Default.ST31 = ST31.Text;

            Properties.Settings.Default.CAuto1 = Auto1.Text;
            Properties.Settings.Default.CAuto2 = Auto2.Text;
            Properties.Settings.Default.CAuto3 = Auto3.Text;
            Properties.Settings.Default.CAuto4 = Auto4.Text;

            Properties.Settings.Default.Mes = Meses1.Text;
            Properties.Settings.Default.Year = Years1.Text;

            List<string> ListChecks = new List<string>
            {
                Extra1.CheckState.ToString(),
                Extra2.CheckState.ToString(),
                Extra3.CheckState.ToString(),
                Extra4.CheckState.ToString(),
                Extra5.CheckState.ToString(),
                Extra6.CheckState.ToString(),
                Extra7.CheckState.ToString(),
                Extra8.CheckState.ToString(),
                Extra9.CheckState.ToString(),
                Extra10.CheckState.ToString(),
                Extra11.CheckState.ToString(),
                Extra12.CheckState.ToString(),
                Extra13.CheckState.ToString(),
                Extra14.CheckState.ToString(),
                Extra15.CheckState.ToString(),
                Extra16.CheckState.ToString(),
                Extra17.CheckState.ToString(),
                Extra18.CheckState.ToString(),
                Extra19.CheckState.ToString(),
                Extra20.CheckState.ToString(),
                Extra21.CheckState.ToString(),
                Extra22.CheckState.ToString(),
                Extra23.CheckState.ToString(),
                Extra24.CheckState.ToString(),
                Extra25.CheckState.ToString(),
                Extra26.CheckState.ToString(),
                Extra27.CheckState.ToString(),
                Extra28.CheckState.ToString(),
                Extra29.CheckState.ToString(),
                Extra30.CheckState.ToString(),
                Extra31.CheckState.ToString()
            };
            Properties.Settings.Default.Checks = String.Join(",", ListChecks.ToArray());

            CultureInfo ci = new CultureInfo("ES-ES");
            DateTime Days1 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 1);
            if (ComboDays.Text.ToLower() == (Days1.ToString("dddd", ci) + " - " + Days1.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM1 = EM0.Text;
                Properties.Settings.Default.SM1 = SM0.Text;
                Properties.Settings.Default.ET1 = ET0.Text;
                Properties.Settings.Default.ST1 = ST0.Text;
            }
            DateTime Days2 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 2);
            if (ComboDays.Text.ToLower() == (Days2.ToString("dddd", ci) + " - " + Days2.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM2 = EM0.Text;
                Properties.Settings.Default.SM2 = SM0.Text;
                Properties.Settings.Default.ET2 = ET0.Text;
                Properties.Settings.Default.ST2 = ST0.Text;
            }
            DateTime Days3 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 3);
            if (ComboDays.Text.ToLower() == (Days3.ToString("dddd", ci) + " - " + Days3.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM3 = EM0.Text;
                Properties.Settings.Default.SM3 = SM0.Text;
                Properties.Settings.Default.ET3 = ET0.Text;
                Properties.Settings.Default.ST3 = ST0.Text;
            }
            DateTime Days4 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 4);
            if (ComboDays.Text.ToLower() == (Days4.ToString("dddd", ci) + " - " + Days4.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM4 = EM0.Text;
                Properties.Settings.Default.SM4 = SM0.Text;
                Properties.Settings.Default.ET4 = ET0.Text;
                Properties.Settings.Default.ST4 = ST0.Text;
            }
            DateTime Days5 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 5);
            if (ComboDays.Text.ToLower() == (Days5.ToString("dddd", ci) + " - " + Days5.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM5 = EM0.Text;
                Properties.Settings.Default.SM5 = SM0.Text;
                Properties.Settings.Default.ET5 = ET0.Text;
                Properties.Settings.Default.ST5 = ST0.Text;
            }
            DateTime Days6 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 6);
            if (ComboDays.Text.ToLower() == (Days6.ToString("dddd", ci) + " - " + Days6.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM6 = EM0.Text;
                Properties.Settings.Default.SM6 = SM0.Text;
                Properties.Settings.Default.ET6 = ET0.Text;
                Properties.Settings.Default.ST6 = ST0.Text;
            }
            DateTime Days7 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 7);
            if (ComboDays.Text.ToLower() == (Days7.ToString("dddd", ci) + " - " + Days7.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM7 = EM0.Text;
                Properties.Settings.Default.SM7 = SM0.Text;
                Properties.Settings.Default.ET7 = ET0.Text;
                Properties.Settings.Default.ST7 = ST0.Text;
            }
            DateTime Days8 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 8);
            if (ComboDays.Text.ToLower() == (Days8.ToString("dddd", ci) + " - " + Days8.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM8 = EM0.Text;
                Properties.Settings.Default.SM8 = SM0.Text;
                Properties.Settings.Default.ET8 = ET0.Text;
                Properties.Settings.Default.ST8 = ST0.Text;
            }
            DateTime Days9 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 9);
            if (ComboDays.Text.ToLower() == (Days9.ToString("dddd", ci) + " - " + Days9.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM9 = EM0.Text;
                Properties.Settings.Default.SM9 = SM0.Text;
                Properties.Settings.Default.ET9 = ET0.Text;
                Properties.Settings.Default.ST9 = ST0.Text;
            }
            DateTime Days10 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 10);
            if (ComboDays.Text.ToLower() == (Days10.ToString("dddd", ci) + " - " + Days10.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM10 = EM0.Text;
                Properties.Settings.Default.SM10 = SM0.Text;
                Properties.Settings.Default.ET10 = ET0.Text;
                Properties.Settings.Default.ST10 = ST0.Text;
            }
            DateTime Days11 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 11);
            if (ComboDays.Text.ToLower() == (Days11.ToString("dddd", ci) + " - " + Days11.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM11 = EM0.Text;
                Properties.Settings.Default.SM11 = SM0.Text;
                Properties.Settings.Default.ET11 = ET0.Text;
                Properties.Settings.Default.ST11 = ST0.Text;
            }
            DateTime Days12 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 12);
            if (ComboDays.Text.ToLower() == (Days12.ToString("dddd", ci) + " - " + Days12.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM12 = EM0.Text;
                Properties.Settings.Default.SM12 = SM0.Text;
                Properties.Settings.Default.ET12 = ET0.Text;
                Properties.Settings.Default.ST12 = ST0.Text;
            }
            DateTime Days13 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 13);
            if (ComboDays.Text.ToLower() == (Days13.ToString("dddd", ci) + " - " + Days13.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM13 = EM0.Text;
                Properties.Settings.Default.SM13 = SM0.Text;
                Properties.Settings.Default.ET13 = ET0.Text;
                Properties.Settings.Default.ST13 = ST0.Text;
            }
            DateTime Days14 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 14);
            if (ComboDays.Text.ToLower() == (Days14.ToString("dddd", ci) + " - " + Days14.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM14 = EM0.Text;
                Properties.Settings.Default.SM14 = SM0.Text;
                Properties.Settings.Default.ET14 = ET0.Text;
                Properties.Settings.Default.ST14 = ST0.Text;
            }
            DateTime Days15 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 15);
            if (ComboDays.Text.ToLower() == (Days15.ToString("dddd", ci) + " - " + Days15.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM15 = EM0.Text;
                Properties.Settings.Default.SM15 = SM0.Text;
                Properties.Settings.Default.ET15 = ET0.Text;
                Properties.Settings.Default.ST15 = ST0.Text;
            }
            DateTime Days16 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 16);
            if (ComboDays.Text.ToLower() == (Days16.ToString("dddd", ci) + " - " + Days16.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM16 = EM0.Text;
                Properties.Settings.Default.SM16 = SM0.Text;
                Properties.Settings.Default.ET16 = ET0.Text;
                Properties.Settings.Default.ST16 = ST0.Text;
            }
            DateTime Days17 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 17);
            if (ComboDays.Text.ToLower() == (Days17.ToString("dddd", ci) + " - " + Days17.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM17 = EM0.Text;
                Properties.Settings.Default.SM17 = SM0.Text;
                Properties.Settings.Default.ET17 = ET0.Text;
                Properties.Settings.Default.ST17 = ST0.Text;
            }
            DateTime Days18 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 18);
            if (ComboDays.Text.ToLower() == (Days18.ToString("dddd", ci) + " - " + Days18.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM18 = EM0.Text;
                Properties.Settings.Default.SM18 = SM0.Text;
                Properties.Settings.Default.ET18 = ET0.Text;
                Properties.Settings.Default.ST18 = ST0.Text;
            }
            DateTime Days19 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 19);
            if (ComboDays.Text.ToLower() == (Days19.ToString("dddd", ci) + " - " + Days19.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM19 = EM0.Text;
                Properties.Settings.Default.SM19 = SM0.Text;
                Properties.Settings.Default.ET19 = ET0.Text;
                Properties.Settings.Default.ST19 = ST0.Text;
            }
            DateTime Days20 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 20);
            if (ComboDays.Text.ToLower() == (Days20.ToString("dddd", ci) + " - " + Days20.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM20 = EM0.Text;
                Properties.Settings.Default.SM20 = SM0.Text;
                Properties.Settings.Default.ET20 = ET0.Text;
                Properties.Settings.Default.ST20 = ST0.Text;
            }
            DateTime Days21 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 21);
            if (ComboDays.Text.ToLower() == (Days21.ToString("dddd", ci) + " - " + Days21.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM21 = EM0.Text;
                Properties.Settings.Default.SM21 = SM0.Text;
                Properties.Settings.Default.ET21 = ET0.Text;
                Properties.Settings.Default.ST21 = ST0.Text;
            }
            DateTime Days22 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 22);
            if (ComboDays.Text.ToLower() == (Days22.ToString("dddd", ci) + " - " + Days22.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM22 = EM0.Text;
                Properties.Settings.Default.SM22 = SM0.Text;
                Properties.Settings.Default.ET22 = ET0.Text;
                Properties.Settings.Default.ST22 = ST0.Text;
            }
            DateTime Days23 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 23);
            if (ComboDays.Text.ToLower() == (Days23.ToString("dddd", ci) + " - " + Days23.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM23 = EM0.Text;
                Properties.Settings.Default.SM23 = SM0.Text;
                Properties.Settings.Default.ET23 = ET0.Text;
                Properties.Settings.Default.ST23 = ST0.Text;
            }
            DateTime Days24 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 24);
            if (ComboDays.Text.ToLower() == (Days24.ToString("dddd", ci) + " - " + Days24.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM24 = EM0.Text;
                Properties.Settings.Default.SM24 = SM0.Text;
                Properties.Settings.Default.ET24 = ET0.Text;
                Properties.Settings.Default.ST24 = ST0.Text;
            }
            DateTime Days25 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 25);
            if (ComboDays.Text.ToLower() == (Days25.ToString("dddd", ci) + " - " + Days25.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM25 = EM0.Text;
                Properties.Settings.Default.SM25 = SM0.Text;
                Properties.Settings.Default.ET25 = ET0.Text;
                Properties.Settings.Default.ST25 = ST0.Text;
            }
            DateTime Days26 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 26);
            if (ComboDays.Text.ToLower() == (Days26.ToString("dddd", ci) + " - " + Days26.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM26 = EM0.Text;
                Properties.Settings.Default.SM26 = SM0.Text;
                Properties.Settings.Default.ET26 = ET0.Text;
                Properties.Settings.Default.ST26 = ST0.Text;
            }
            DateTime Days27 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 27);
            if (ComboDays.Text.ToLower() == (Days27.ToString("dddd", ci) + " - " + Days27.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM27 = EM0.Text;
                Properties.Settings.Default.SM27 = SM0.Text;
                Properties.Settings.Default.ET27 = ET0.Text;
                Properties.Settings.Default.ST27 = ST0.Text;
            }
            DateTime Days28 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 28);
            if (ComboDays.Text.ToLower() == (Days28.ToString("dddd", ci) + " - " + Days28.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM28 = EM0.Text;
                Properties.Settings.Default.SM28 = SM0.Text;
                Properties.Settings.Default.ET28 = ET0.Text;
                Properties.Settings.Default.ST28 = ST0.Text;
            }
            DateTime Days29 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 29);
            if (ComboDays.Text.ToLower() == (Days29.ToString("dddd", ci) + " - " + Days29.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM29 = EM0.Text;
                Properties.Settings.Default.SM29 = SM0.Text;
                Properties.Settings.Default.ET29 = ET0.Text;
                Properties.Settings.Default.ST29 = ST0.Text;
            }
            DateTime Days30 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 30);
            if (ComboDays.Text.ToLower() == (Days30.ToString("dddd", ci) + " - " + Days30.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM30 = EM0.Text;
                Properties.Settings.Default.SM30 = SM0.Text;
                Properties.Settings.Default.ET30 = ET0.Text;
                Properties.Settings.Default.ST30 = ST0.Text;
            }
            DateTime Days31 = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 31);
            if (ComboDays.Text.ToLower() == (Days31.ToString("dddd", ci) + " - " + Days31.ToString("dd")).ToLower())
            {
                Properties.Settings.Default.EM31 = EM0.Text;
                Properties.Settings.Default.SM31 = SM0.Text;
                Properties.Settings.Default.ET31 = ET0.Text;
                Properties.Settings.Default.ST31 = ST0.Text;
            }

            Properties.Settings.Default.Save();
            Properties.Settings.Default.Reload();

            MessageBox.Show("Datos guardados!", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        #endregion

        #region MODIFICAR EXCEL
        void WriteToExcel()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();

                if (File.Exists(Application.StartupPath + @"\Plantilla.data"))
                {
                    File.Copy(Application.StartupPath + @"\Plantilla.data", Application.StartupPath + @"\Jornadas\Registro de Jornada de " + MesPreSelected + " del " + YearSelected.ToString() + ".xlsx", true);
                }
                Workbook workbook = Excel.Workbooks.Open(Application.StartupPath + @"\Jornadas\Registro de Jornada de " + MesPreSelected + " del " + YearSelected.ToString() + ".xlsx", ReadOnly: false, Editable: true);
                if (!(workbook.Worksheets.Item[1] is Worksheet worksheet))
                    return;

                DateTime oPrimerDiaDelMes = new DateTime(Convert.ToInt32(YearSelected), MesSelected, 1);
                DateTime oUltimoDiaDelMes = oPrimerDiaDelMes.AddMonths(1).AddDays(-1);

                worksheet.Cells[9, 8].Value = oUltimoDiaDelMes.ToString("dd/MM/yyyy");
                worksheet.Cells[51, 2].Value = oUltimoDiaDelMes.ToString("dd/MM/yyyy");
                worksheet.Cells[7, 3] = Properties.Settings.Default.Name;
                worksheet.Cells[8, 2] = Properties.Settings.Default.Documento;
                worksheet.Cells[8, 8] = Properties.Settings.Default.SeguridadS;
                worksheet.Cells[9, 4] = MesPreSelected;

                List<TimeSpan> SumaHorasJornada = new List<TimeSpan>();
                List<TimeSpan> SumaHoras = new List<TimeSpan>();
                List<TimeSpan> SumaHorasExtra = new List<TimeSpan>();

                //_1
                TimeSpan fecha1_1 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_1 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_1 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_1 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_1 = TimeSpan.Parse("0:00");
                if (EM1.Text != "")
                {
                    fecha1_1 = TimeSpan.Parse(EM1.Text);
                    worksheet.Cells[12, 2] = EM1.Text;
                }
                if (SM1.Text != "")
                {
                    fecha2_1 = TimeSpan.Parse(SM1.Text);
                    worksheet.Cells[12, 3] = SM1.Text;
                }
                if (ET1.Text != "")
                {
                    fecha3_1 = TimeSpan.Parse(ET1.Text);
                    worksheet.Cells[12, 4] = ET1.Text;
                }
                if (ST1.Text != "")
                {
                    fecha4_1 = TimeSpan.Parse(ST1.Text);
                    worksheet.Cells[12, 5] = ST1.Text;
                }
                TimeSpan restar_1 = fecha2_1 - fecha1_1;
                TimeSpan restax_1 = fecha4_1 - fecha3_1;
                TimeSpan suma_1 = restar_1 + restax_1;

                string resultadot_1 = (suma_1.ToString(@"h\:mm"));
                if (resultadot_1 != "0:00")
                {
                    worksheet.Cells[12, 6] = resultadot_1;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 371, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_1));
                    if (Extra1.Checked == true)
                    {
                        if (suma_1 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_1 = suma_1 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_1);
                            worksheet.Cells[12, 8] = RestaExtra_1.ToString(@"h\:mm");
                            worksheet.Cells[12, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[12, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[12, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_1);
                            worksheet.Cells[12, 8] = suma_1.ToString(@"h\:mm");
                            worksheet.Cells[12, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_1);
                        worksheet.Cells[12, 7] = suma_1.ToString(@"h\:mm");
                        worksheet.Cells[12, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_2
                TimeSpan fecha1_2 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_2 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_2 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_2 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_2 = TimeSpan.Parse("0:00");
                if (EM2.Text != "")
                {
                    fecha1_2 = TimeSpan.Parse(EM2.Text);
                    worksheet.Cells[13, 2] = EM2.Text;
                }
                if (SM2.Text != "")
                {
                    fecha2_2 = TimeSpan.Parse(SM2.Text);
                    worksheet.Cells[13, 3] = SM2.Text;
                }
                if (ET2.Text != "")
                {
                    fecha3_2 = TimeSpan.Parse(ET2.Text);
                    worksheet.Cells[13, 4] = ET2.Text;
                }
                if (ST2.Text != "")
                {
                    fecha4_2 = TimeSpan.Parse(ST2.Text);
                    worksheet.Cells[13, 5] = ST2.Text;
                }
                TimeSpan restar_2 = fecha2_2 - fecha1_2;
                TimeSpan restax_2 = fecha4_2 - fecha3_2;
                TimeSpan suma_2 = restar_2 + restax_2;

                string resultadot_2 = (suma_2.ToString(@"h\:mm"));
                if (resultadot_2 != "0:00")
                {
                    worksheet.Cells[13, 6] = resultadot_2;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 400, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_2));
                    if (Extra2.Checked == true)
                    {
                        if (suma_2 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_2 = suma_2 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_2);
                            worksheet.Cells[13, 8] = RestaExtra_2.ToString(@"h\:mm");
                            worksheet.Cells[13, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[13, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[13, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_2);
                            worksheet.Cells[13, 8] = suma_2.ToString(@"h\:mm");
                            worksheet.Cells[13, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_2);
                        worksheet.Cells[13, 7] = suma_2.ToString(@"h\:mm");
                        worksheet.Cells[13, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_3
                TimeSpan fecha1_3 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_3 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_3 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_3 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_3 = TimeSpan.Parse("0:00");
                if (EM3.Text != "")
                {
                    fecha1_3 = TimeSpan.Parse(EM3.Text);
                    worksheet.Cells[14, 2] = EM3.Text;
                }
                if (SM3.Text != "")
                {
                    fecha2_3 = TimeSpan.Parse(SM3.Text);
                    worksheet.Cells[14, 3] = SM3.Text;
                }
                if (ET3.Text != "")
                {
                    fecha3_3 = TimeSpan.Parse(ET3.Text);
                    worksheet.Cells[14, 4] = ET3.Text;
                }
                if (ST3.Text != "")
                {
                    fecha4_3 = TimeSpan.Parse(ST3.Text);
                    worksheet.Cells[14, 5] = ST3.Text;
                }
                TimeSpan restar_3 = fecha2_3 - fecha1_3;
                TimeSpan restax_3 = fecha4_3 - fecha3_3;
                TimeSpan suma_3 = restar_3 + restax_3;

                string resultadot_3 = (suma_3.ToString(@"h\:mm"));
                if (resultadot_3 != "0:00")
                {
                    worksheet.Cells[14, 6] = resultadot_3;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 428, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_3));
                    if (Extra3.Checked == true)
                    {
                        if (suma_3 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_3 = suma_3 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_3);
                            worksheet.Cells[14, 8] = RestaExtra_3.ToString(@"h\:mm");
                            worksheet.Cells[14, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[14, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[14, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_3);
                            worksheet.Cells[14, 8] = suma_3.ToString(@"h\:mm");
                            worksheet.Cells[14, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_3);
                        worksheet.Cells[14, 7] = suma_3.ToString(@"h\:mm");
                        worksheet.Cells[14, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_4
                TimeSpan fecha1_4 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_4 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_4 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_4 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_4 = TimeSpan.Parse("0:00");
                if (EM4.Text != "")
                {
                    fecha1_4 = TimeSpan.Parse(EM4.Text);
                    worksheet.Cells[15, 2] = EM4.Text;
                }
                if (SM4.Text != "")
                {
                    fecha2_4 = TimeSpan.Parse(SM4.Text);
                    worksheet.Cells[15, 3] = SM4.Text;
                }
                if (ET4.Text != "")
                {
                    fecha3_4 = TimeSpan.Parse(ET4.Text);
                    worksheet.Cells[15, 4] = ET4.Text;
                }
                if (ST4.Text != "")
                {
                    fecha4_4 = TimeSpan.Parse(ST4.Text);
                    worksheet.Cells[15, 5] = ST4.Text;
                }
                TimeSpan restar_4 = fecha2_4 - fecha1_4;
                TimeSpan restax_4 = fecha4_4 - fecha3_4;
                TimeSpan suma_4 = restar_4 + restax_4;

                string resultadot_4 = (suma_4.ToString(@"h\:mm"));
                if (resultadot_4 != "0:00")
                {
                    worksheet.Cells[15, 6] = resultadot_4;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 456, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_4));
                    if (Extra4.Checked == true)
                    {
                        if (suma_4 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_4 = suma_4 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_4);
                            worksheet.Cells[15, 8] = RestaExtra_4.ToString(@"h\:mm");
                            worksheet.Cells[15, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[15, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[15, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_4);
                            worksheet.Cells[15, 8] = suma_4.ToString(@"h\:mm");
                            worksheet.Cells[15, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_4);
                        worksheet.Cells[15, 7] = suma_4.ToString(@"h\:mm");
                        worksheet.Cells[15, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_5
                TimeSpan fecha1_5 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_5 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_5 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_5 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_5 = TimeSpan.Parse("0:00");
                if (EM5.Text != "")
                {
                    fecha1_5 = TimeSpan.Parse(EM5.Text);
                    worksheet.Cells[16, 2] = EM5.Text;
                }
                if (SM5.Text != "")
                {
                    fecha2_5 = TimeSpan.Parse(SM5.Text);
                    worksheet.Cells[16, 3] = SM5.Text;
                }
                if (ET5.Text != "")
                {
                    fecha3_5 = TimeSpan.Parse(ET5.Text);
                    worksheet.Cells[16, 4] = ET5.Text;
                }
                if (ST5.Text != "")
                {
                    fecha4_5 = TimeSpan.Parse(ST5.Text);
                    worksheet.Cells[16, 5] = ST5.Text;
                }
                TimeSpan restar_5 = fecha2_5 - fecha1_5;
                TimeSpan restax_5 = fecha4_5 - fecha3_5;
                TimeSpan suma_5 = restar_5 + restax_5;

                string resultadot_5 = (suma_5.ToString(@"h\:mm"));
                if (resultadot_5 != "0:00")
                {
                    worksheet.Cells[16, 6] = resultadot_5;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 484, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_5));
                    if (Extra5.Checked == true)
                    {
                        if (suma_5 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_5 = suma_5 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_5);
                            worksheet.Cells[16, 8] = RestaExtra_5.ToString(@"h\:mm");
                            worksheet.Cells[16, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[16, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[16, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_5);
                            worksheet.Cells[16, 8] = suma_5.ToString(@"h\:mm");
                            worksheet.Cells[16, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_5);
                        worksheet.Cells[16, 7] = suma_5.ToString(@"h\:mm");
                        worksheet.Cells[16, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_6
                TimeSpan fecha1_6 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_6 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_6 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_6 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_6 = TimeSpan.Parse("0:00");
                if (EM6.Text != "")
                {
                    fecha1_6 = TimeSpan.Parse(EM6.Text);
                    worksheet.Cells[17, 2] = EM6.Text;
                }
                if (SM6.Text != "")
                {
                    fecha2_6 = TimeSpan.Parse(SM6.Text);
                    worksheet.Cells[17, 3] = SM6.Text;
                }
                if (ET6.Text != "")
                {
                    fecha3_6 = TimeSpan.Parse(ET6.Text);
                    worksheet.Cells[17, 4] = ET6.Text;
                }
                if (ST6.Text != "")
                {
                    fecha4_6 = TimeSpan.Parse(ST6.Text);
                    worksheet.Cells[17, 5] = ST6.Text;
                }
                TimeSpan restar_6 = fecha2_6 - fecha1_6;
                TimeSpan restax_6 = fecha4_6 - fecha3_6;
                TimeSpan suma_6 = restar_6 + restax_6;

                string resultadot_6 = (suma_6.ToString(@"h\:mm"));
                if (resultadot_6 != "0:00")
                {
                    worksheet.Cells[17, 6] = resultadot_6;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 511, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_6));
                    if (Extra6.Checked == true)
                    {
                        if (suma_6 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_6 = suma_6 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_6);
                            worksheet.Cells[17, 8] = RestaExtra_6.ToString(@"h\:mm");
                            worksheet.Cells[17, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[17, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[17, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_6);
                            worksheet.Cells[17, 8] = suma_6.ToString(@"h\:mm");
                            worksheet.Cells[17, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_6);
                        worksheet.Cells[17, 7] = suma_6.ToString(@"h\:mm");
                        worksheet.Cells[17, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_7
                TimeSpan fecha1_7 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_7 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_7 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_7 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_7 = TimeSpan.Parse("0:00");
                if (EM7.Text != "")
                {
                    fecha1_7 = TimeSpan.Parse(EM7.Text);
                    worksheet.Cells[18, 2] = EM7.Text;
                }
                if (SM7.Text != "")
                {
                    fecha2_7 = TimeSpan.Parse(SM7.Text);
                    worksheet.Cells[18, 3] = SM7.Text;
                }
                if (ET7.Text != "")
                {
                    fecha3_7 = TimeSpan.Parse(ET7.Text);
                    worksheet.Cells[18, 4] = ET7.Text;
                }
                if (ST7.Text != "")
                {
                    fecha4_7 = TimeSpan.Parse(ST7.Text);
                    worksheet.Cells[18, 5] = ST7.Text;
                }
                TimeSpan restar_7 = fecha2_7 - fecha1_7;
                TimeSpan restax_7 = fecha4_7 - fecha3_7;
                TimeSpan suma_7 = restar_7 + restax_7;

                string resultadot_7 = (suma_7.ToString(@"h\:mm"));
                if (resultadot_7 != "0:00")
                {
                    worksheet.Cells[18, 6] = resultadot_7;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 538, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_7));
                    if (Extra7.Checked == true)
                    {
                        if (suma_7 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_7 = suma_7 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_7);
                            worksheet.Cells[18, 8] = RestaExtra_7.ToString(@"h\:mm");
                            worksheet.Cells[18, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[18, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[18, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_7);
                            worksheet.Cells[18, 8] = suma_7.ToString(@"h\:mm");
                            worksheet.Cells[18, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_7);
                        worksheet.Cells[18, 7] = suma_7.ToString(@"h\:mm");
                        worksheet.Cells[18, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_8
                TimeSpan fecha1_8 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_8 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_8 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_8 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_8 = TimeSpan.Parse("0:00");
                if (EM8.Text != "")
                {
                    fecha1_8 = TimeSpan.Parse(EM8.Text);
                    worksheet.Cells[19, 2] = EM8.Text;
                }
                if (SM8.Text != "")
                {
                    fecha2_8 = TimeSpan.Parse(SM8.Text);
                    worksheet.Cells[19, 3] = SM8.Text;
                }
                if (ET8.Text != "")
                {
                    fecha3_8 = TimeSpan.Parse(ET8.Text);
                    worksheet.Cells[19, 4] = ET8.Text;
                }
                if (ST8.Text != "")
                {
                    fecha4_8 = TimeSpan.Parse(ST8.Text);
                    worksheet.Cells[19, 5] = ST8.Text;
                }
                TimeSpan restar_8 = fecha2_8 - fecha1_8;
                TimeSpan restax_8 = fecha4_8 - fecha3_8;
                TimeSpan suma_8 = restar_8 + restax_8;

                string resultadot_8 = (suma_8.ToString(@"h\:mm"));
                if (resultadot_8 != "0:00")
                {
                    worksheet.Cells[19, 6] = resultadot_8;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 566, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_8));
                    if (Extra8.Checked == true)
                    {
                        if (suma_8 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_8 = suma_8 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_8);
                            worksheet.Cells[19, 8] = RestaExtra_8.ToString(@"h\:mm");
                            worksheet.Cells[19, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[19, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[19, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_8);
                            worksheet.Cells[19, 8] = suma_8.ToString(@"h\:mm");
                            worksheet.Cells[19, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_8);
                        worksheet.Cells[19, 7] = suma_8.ToString(@"h\:mm");
                        worksheet.Cells[19, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_9
                TimeSpan fecha1_9 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_9 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_9 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_9 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_9 = TimeSpan.Parse("0:00");
                if (EM9.Text != "")
                {
                    fecha1_9 = TimeSpan.Parse(EM9.Text);
                    worksheet.Cells[20, 2] = EM9.Text;
                }
                if (SM9.Text != "")
                {
                    fecha2_9 = TimeSpan.Parse(SM9.Text);
                    worksheet.Cells[20, 3] = SM9.Text;
                }
                if (ET9.Text != "")
                {
                    fecha3_9 = TimeSpan.Parse(ET9.Text);
                    worksheet.Cells[20, 4] = ET9.Text;
                }
                if (ST9.Text != "")
                {
                    fecha4_9 = TimeSpan.Parse(ST9.Text);
                    worksheet.Cells[20, 5] = ST9.Text;
                }
                TimeSpan restar_9 = fecha2_9 - fecha1_9;
                TimeSpan restax_9 = fecha4_9 - fecha3_9;
                TimeSpan suma_9 = restar_9 + restax_9;

                string resultadot_9 = (suma_9.ToString(@"h\:mm"));
                if (resultadot_9 != "0:00")
                {
                    worksheet.Cells[20, 6] = resultadot_9;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 593, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_9));
                    if (Extra9.Checked == true)
                    {
                        if (suma_9 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_9 = suma_9 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_9);
                            worksheet.Cells[20, 8] = RestaExtra_9.ToString(@"h\:mm");
                            worksheet.Cells[20, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[20, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[20, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_9);
                            worksheet.Cells[20, 8] = suma_9.ToString(@"h\:mm");
                            worksheet.Cells[20, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_9);
                        worksheet.Cells[20, 7] = suma_9.ToString(@"h\:mm");
                        worksheet.Cells[20, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_10
                TimeSpan fecha1_10 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_10 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_10 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_10 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_10 = TimeSpan.Parse("0:00");
                if (EM10.Text != "")
                {
                    fecha1_10 = TimeSpan.Parse(EM10.Text);
                    worksheet.Cells[21, 2] = EM10.Text;
                }
                if (SM10.Text != "")
                {
                    fecha2_10 = TimeSpan.Parse(SM10.Text);
                    worksheet.Cells[21, 3] = SM10.Text;
                }
                if (ET10.Text != "")
                {
                    fecha3_10 = TimeSpan.Parse(ET10.Text);
                    worksheet.Cells[21, 4] = ET10.Text;
                }
                if (ST10.Text != "")
                {
                    fecha4_10 = TimeSpan.Parse(ST10.Text);
                    worksheet.Cells[21, 5] = ST10.Text;
                }
                TimeSpan restar_10 = fecha2_10 - fecha1_10;
                TimeSpan restax_10 = fecha4_10 - fecha3_10;
                TimeSpan suma_10 = restar_10 + restax_10;

                string resultadot_10 = (suma_10.ToString(@"h\:mm"));
                if (resultadot_10 != "0:00")
                {
                    worksheet.Cells[21, 6] = resultadot_10;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 621, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_10));
                    if (Extra10.Checked == true)
                    {
                        if (suma_10 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_10 = suma_10 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_10);
                            worksheet.Cells[21, 8] = RestaExtra_10.ToString(@"h\:mm");
                            worksheet.Cells[21, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[21, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[21, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_10);
                            worksheet.Cells[21, 8] = suma_10.ToString(@"h\:mm");
                            worksheet.Cells[21, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_10);
                        worksheet.Cells[21, 7] = suma_10.ToString(@"h\:mm");
                        worksheet.Cells[21, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_11
                TimeSpan fecha1_11 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_11 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_11 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_11 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_11 = TimeSpan.Parse("0:00");
                if (EM11.Text != "")
                {
                    fecha1_11 = TimeSpan.Parse(EM11.Text);
                    worksheet.Cells[22, 2] = EM11.Text;
                }
                if (SM11.Text != "")
                {
                    fecha2_11 = TimeSpan.Parse(SM11.Text);
                    worksheet.Cells[22, 3] = SM11.Text;
                }
                if (ET11.Text != "")
                {
                    fecha3_11 = TimeSpan.Parse(ET11.Text);
                    worksheet.Cells[22, 4] = ET11.Text;
                }
                if (ST11.Text != "")
                {
                    fecha4_11 = TimeSpan.Parse(ST11.Text);
                    worksheet.Cells[22, 5] = ST11.Text;
                }
                TimeSpan restar_11 = fecha2_11 - fecha1_11;
                TimeSpan restax_11 = fecha4_11 - fecha3_11;
                TimeSpan suma_11 = restar_11 + restax_11;

                string resultadot_11 = (suma_11.ToString(@"h\:mm"));
                if (resultadot_11 != "0:00")
                {
                    worksheet.Cells[22, 6] = resultadot_11;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 649, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_11));
                    if (Extra11.Checked == true)
                    {
                        if (suma_11 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_11 = suma_11 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_11);
                            worksheet.Cells[22, 8] = RestaExtra_11.ToString(@"h\:mm");
                            worksheet.Cells[22, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[22, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[22, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_11);
                            worksheet.Cells[22, 8] = suma_11.ToString(@"h\:mm");
                            worksheet.Cells[22, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_11);
                        worksheet.Cells[22, 7] = suma_11.ToString(@"h\:mm");
                        worksheet.Cells[22, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_12
                TimeSpan fecha1_12 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_12 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_12 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_12 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_12 = TimeSpan.Parse("0:00");
                if (EM12.Text != "")
                {
                    fecha1_12 = TimeSpan.Parse(EM12.Text);
                    worksheet.Cells[23, 2] = EM12.Text;
                }
                if (SM12.Text != "")
                {
                    fecha2_12 = TimeSpan.Parse(SM12.Text);
                    worksheet.Cells[23, 3] = SM12.Text;
                }
                if (ET12.Text != "")
                {
                    fecha3_12 = TimeSpan.Parse(ET12.Text);
                    worksheet.Cells[23, 4] = ET12.Text;
                }
                if (ST12.Text != "")
                {
                    fecha4_12 = TimeSpan.Parse(ST12.Text);
                    worksheet.Cells[23, 5] = ST12.Text;
                }
                TimeSpan restar_12 = fecha2_12 - fecha1_12;
                TimeSpan restax_12 = fecha4_12 - fecha3_12;
                TimeSpan suma_12 = restar_12 + restax_12;

                string resultadot_12 = (suma_12.ToString(@"h\:mm"));
                if (resultadot_12 != "0:00")
                {
                    worksheet.Cells[23, 6] = resultadot_12;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 677, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_12));
                    if (Extra12.Checked == true)
                    {
                        if (suma_12 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_12 = suma_12 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_12);
                            worksheet.Cells[23, 8] = RestaExtra_12.ToString(@"h\:mm");
                            worksheet.Cells[23, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[23, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[23, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_12);
                            worksheet.Cells[23, 8] = suma_12.ToString(@"h\:mm");
                            worksheet.Cells[23, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_12);
                        worksheet.Cells[23, 7] = suma_12.ToString(@"h\:mm");
                        worksheet.Cells[23, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_13
                TimeSpan fecha1_13 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_13 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_13 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_13 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_13 = TimeSpan.Parse("0:00");
                if (EM13.Text != "")
                {
                    fecha1_13 = TimeSpan.Parse(EM13.Text);
                    worksheet.Cells[24, 2] = EM13.Text;
                }
                if (SM13.Text != "")
                {
                    fecha2_13 = TimeSpan.Parse(SM13.Text);
                    worksheet.Cells[24, 3] = SM13.Text;
                }
                if (ET13.Text != "")
                {
                    fecha3_13 = TimeSpan.Parse(ET13.Text);
                    worksheet.Cells[24, 4] = ET13.Text;
                }
                if (ST13.Text != "")
                {
                    fecha4_13 = TimeSpan.Parse(ST13.Text);
                    worksheet.Cells[24, 5] = ST13.Text;
                }
                TimeSpan restar_13 = fecha2_13 - fecha1_13;
                TimeSpan restax_13 = fecha4_13 - fecha3_13;
                TimeSpan suma_13 = restar_13 + restax_13;

                string resultadot_13 = (suma_13.ToString(@"h\:mm"));
                if (resultadot_13 != "0:00")
                {
                    worksheet.Cells[24, 6] = resultadot_13;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 705, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_13));
                    if (Extra13.Checked == true)
                    {
                        if (suma_13 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_13 = suma_13 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_13);
                            worksheet.Cells[24, 8] = RestaExtra_13.ToString(@"h\:mm");
                            worksheet.Cells[24, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[24, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[24, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_13);
                            worksheet.Cells[24, 8] = suma_13.ToString(@"h\:mm");
                            worksheet.Cells[24, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_13);
                        worksheet.Cells[24, 7] = suma_13.ToString(@"h\:mm");
                        worksheet.Cells[24, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_14
                TimeSpan fecha1_14 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_14 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_14 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_14 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_14 = TimeSpan.Parse("0:00");
                if (EM14.Text != "")
                {
                    fecha1_14 = TimeSpan.Parse(EM14.Text);
                    worksheet.Cells[25, 2] = EM14.Text;
                }
                if (SM14.Text != "")
                {
                    fecha2_14 = TimeSpan.Parse(SM14.Text);
                    worksheet.Cells[25, 3] = SM14.Text;
                }
                if (ET14.Text != "")
                {
                    fecha3_14 = TimeSpan.Parse(ET14.Text);
                    worksheet.Cells[25, 4] = ET14.Text;
                }
                if (ST14.Text != "")
                {
                    fecha4_14 = TimeSpan.Parse(ST14.Text);
                    worksheet.Cells[25, 5] = ST14.Text;
                }
                TimeSpan restar_14 = fecha2_14 - fecha1_14;
                TimeSpan restax_14 = fecha4_14 - fecha3_14;
                TimeSpan suma_14 = restar_14 + restax_14;

                string resultadot_14 = (suma_14.ToString(@"h\:mm"));
                if (resultadot_14 != "0:00")
                {
                    worksheet.Cells[25, 6] = resultadot_14;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 732, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_14));
                    if (Extra14.Checked == true)
                    {
                        if (suma_14 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_14 = suma_14 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_14);
                            worksheet.Cells[25, 8] = RestaExtra_14.ToString(@"h\:mm");
                            worksheet.Cells[25, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[25, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[25, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_14);
                            worksheet.Cells[25, 8] = suma_14.ToString(@"h\:mm");
                            worksheet.Cells[25, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_14);
                        worksheet.Cells[25, 7] = suma_14.ToString(@"h\:mm");
                        worksheet.Cells[25, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_15
                TimeSpan fecha1_15 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_15 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_15 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_15 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_15 = TimeSpan.Parse("0:00");
                if (EM15.Text != "")
                {
                    fecha1_15 = TimeSpan.Parse(EM15.Text);
                    worksheet.Cells[26, 2] = EM15.Text;
                }
                if (SM15.Text != "")
                {
                    fecha2_15 = TimeSpan.Parse(SM15.Text);
                    worksheet.Cells[26, 3] = SM15.Text;
                }
                if (ET15.Text != "")
                {
                    fecha3_15 = TimeSpan.Parse(ET15.Text);
                    worksheet.Cells[26, 4] = ET15.Text;
                }
                if (ST15.Text != "")
                {
                    fecha4_15 = TimeSpan.Parse(ST15.Text);
                    worksheet.Cells[26, 5] = ST15.Text;
                }
                TimeSpan restar_15 = fecha2_15 - fecha1_15;
                TimeSpan restax_15 = fecha4_15 - fecha3_15;
                TimeSpan suma_15 = restar_15 + restax_15;

                string resultadot_15 = (suma_15.ToString(@"h\:mm"));
                if (resultadot_15 != "0:00")
                {
                    worksheet.Cells[26, 6] = resultadot_15;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 760, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_15));
                    if (Extra15.Checked == true)
                    {
                        if (suma_15 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_15 = suma_15 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_15);
                            worksheet.Cells[26, 8] = RestaExtra_15.ToString(@"h\:mm");
                            worksheet.Cells[26, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[26, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[26, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_15);
                            worksheet.Cells[26, 8] = suma_15.ToString(@"h\:mm");
                            worksheet.Cells[26, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_15);
                        worksheet.Cells[26, 7] = suma_15.ToString(@"h\:mm");
                        worksheet.Cells[26, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_16
                TimeSpan fecha1_16 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_16 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_16 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_16 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_16 = TimeSpan.Parse("0:00");
                if (EM16.Text != "")
                {
                    fecha1_16 = TimeSpan.Parse(EM16.Text);
                    worksheet.Cells[27, 2] = EM16.Text;
                }
                if (SM16.Text != "")
                {
                    fecha2_16 = TimeSpan.Parse(SM16.Text);
                    worksheet.Cells[27, 3] = SM16.Text;
                }
                if (ET16.Text != "")
                {
                    fecha3_16 = TimeSpan.Parse(ET16.Text);
                    worksheet.Cells[27, 4] = ET16.Text;
                }
                if (ST16.Text != "")
                {
                    fecha4_16 = TimeSpan.Parse(ST16.Text);
                    worksheet.Cells[27, 5] = ST16.Text;
                }
                TimeSpan restar_16 = fecha2_16 - fecha1_16;
                TimeSpan restax_16 = fecha4_16 - fecha3_16;
                TimeSpan suma_16 = restar_16 + restax_16;

                string resultadot_16 = (suma_16.ToString(@"h\:mm"));
                if (resultadot_16 != "0:00")
                {
                    worksheet.Cells[27, 6] = resultadot_16;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 789, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_16));
                    if (Extra16.Checked == true)
                    {
                        if (suma_16 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_16 = suma_16 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_16);
                            worksheet.Cells[27, 8] = RestaExtra_16.ToString(@"h\:mm");
                            worksheet.Cells[27, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[27, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[27, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_16);
                            worksheet.Cells[27, 8] = suma_16.ToString(@"h\:mm");
                            worksheet.Cells[27, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_16);
                        worksheet.Cells[27, 7] = suma_16.ToString(@"h\:mm");
                        worksheet.Cells[27, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_17
                TimeSpan fecha1_17 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_17 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_17 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_17 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_17 = TimeSpan.Parse("0:00");
                if (EM17.Text != "")
                {
                    fecha1_17 = TimeSpan.Parse(EM17.Text);
                    worksheet.Cells[28, 2] = EM17.Text;
                }
                if (SM17.Text != "")
                {
                    fecha2_17 = TimeSpan.Parse(SM17.Text);
                    worksheet.Cells[28, 3] = SM17.Text;
                }
                if (ET17.Text != "")
                {
                    fecha3_17 = TimeSpan.Parse(ET17.Text);
                    worksheet.Cells[28, 4] = ET17.Text;
                }
                if (ST17.Text != "")
                {
                    fecha4_17 = TimeSpan.Parse(ST17.Text);
                    worksheet.Cells[28, 5] = ST17.Text;
                }
                TimeSpan restar_17 = fecha2_17 - fecha1_17;
                TimeSpan restax_17 = fecha4_17 - fecha3_17;
                TimeSpan suma_17 = restar_17 + restax_17;

                string resultadot_17 = (suma_17.ToString(@"h\:mm"));
                if (resultadot_17 != "0:00")
                {
                    worksheet.Cells[28, 6] = resultadot_17;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 816, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_17));
                    if (Extra17.Checked == true)
                    {
                        if (suma_17 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_17 = suma_17 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_17);
                            worksheet.Cells[28, 8] = RestaExtra_17.ToString(@"h\:mm");
                            worksheet.Cells[28, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[28, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[28, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_17);
                            worksheet.Cells[28, 8] = suma_17.ToString(@"h\:mm");
                            worksheet.Cells[28, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_17);
                        worksheet.Cells[28, 7] = suma_17.ToString(@"h\:mm");
                        worksheet.Cells[28, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_18
                TimeSpan fecha1_18 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_18 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_18 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_18 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_18 = TimeSpan.Parse("0:00");
                if (EM18.Text != "")
                {
                    fecha1_18 = TimeSpan.Parse(EM18.Text);
                    worksheet.Cells[29, 2] = EM18.Text;
                }
                if (SM18.Text != "")
                {
                    fecha2_18 = TimeSpan.Parse(SM18.Text);
                    worksheet.Cells[29, 3] = SM18.Text;
                }
                if (ET18.Text != "")
                {
                    fecha3_18 = TimeSpan.Parse(ET18.Text);
                    worksheet.Cells[29, 4] = ET18.Text;
                }
                if (ST18.Text != "")
                {
                    fecha4_18 = TimeSpan.Parse(ST18.Text);
                    worksheet.Cells[29, 5] = ST18.Text;
                }
                TimeSpan restar_18 = fecha2_18 - fecha1_18;
                TimeSpan restax_18 = fecha4_18 - fecha3_18;
                TimeSpan suma_18 = restar_18 + restax_18;

                string resultadot_18 = (suma_18.ToString(@"h\:mm"));
                if (resultadot_18 != "0:00")
                {
                    worksheet.Cells[29, 6] = resultadot_18;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 844, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_18));
                    if (Extra18.Checked == true)
                    {
                        if (suma_18 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_18 = suma_18 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_18);
                            worksheet.Cells[29, 8] = RestaExtra_18.ToString(@"h\:mm");
                            worksheet.Cells[29, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[29, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[29, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_18);
                            worksheet.Cells[29, 8] = suma_18.ToString(@"h\:mm");
                            worksheet.Cells[29, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_18);
                        worksheet.Cells[29, 7] = suma_18.ToString(@"h\:mm");
                        worksheet.Cells[29, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_19
                TimeSpan fecha1_19 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_19 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_19 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_19 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_19 = TimeSpan.Parse("0:00");
                if (EM19.Text != "")
                {
                    fecha1_19 = TimeSpan.Parse(EM19.Text);
                    worksheet.Cells[30, 2] = EM19.Text;
                }
                if (SM19.Text != "")
                {
                    fecha2_19 = TimeSpan.Parse(SM19.Text);
                    worksheet.Cells[30, 3] = SM19.Text;
                }
                if (ET19.Text != "")
                {
                    fecha3_19 = TimeSpan.Parse(ET19.Text);
                    worksheet.Cells[30, 4] = ET19.Text;
                }
                if (ST19.Text != "")
                {
                    fecha4_19 = TimeSpan.Parse(ST19.Text);
                    worksheet.Cells[30, 5] = ST19.Text;
                }
                TimeSpan restar_19 = fecha2_19 - fecha1_19;
                TimeSpan restax_19 = fecha4_19 - fecha3_19;
                TimeSpan suma_19 = restar_19 + restax_19;

                string resultadot_19 = (suma_19.ToString(@"h\:mm"));
                if (resultadot_19 != "0:00")
                {
                    worksheet.Cells[30, 6] = resultadot_19;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 871, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_19));
                    if (Extra19.Checked == true)
                    {
                        if (suma_19 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_19 = suma_19 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_19);
                            worksheet.Cells[30, 8] = RestaExtra_19.ToString(@"h\:mm");
                            worksheet.Cells[30, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[30, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[30, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_19);
                            worksheet.Cells[30, 8] = suma_19.ToString(@"h\:mm");
                            worksheet.Cells[30, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_19);
                        worksheet.Cells[30, 7] = suma_19.ToString(@"h\:mm");
                        worksheet.Cells[30, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_20
                TimeSpan fecha1_20 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_20 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_20 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_20 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_20 = TimeSpan.Parse("0:00");
                if (EM20.Text != "")
                {
                    fecha1_20 = TimeSpan.Parse(EM20.Text);
                    worksheet.Cells[31, 2] = EM20.Text;
                }
                if (SM20.Text != "")
                {
                    fecha2_20 = TimeSpan.Parse(SM20.Text);
                    worksheet.Cells[31, 3] = SM20.Text;
                }
                if (ET20.Text != "")
                {
                    fecha3_20 = TimeSpan.Parse(ET20.Text);
                    worksheet.Cells[31, 4] = ET20.Text;
                }
                if (ST20.Text != "")
                {
                    fecha4_20 = TimeSpan.Parse(ST20.Text);
                    worksheet.Cells[31, 5] = ST20.Text;
                }
                TimeSpan restar_20 = fecha2_20 - fecha1_20;
                TimeSpan restax_20 = fecha4_20 - fecha3_20;
                TimeSpan suma_20 = restar_20 + restax_20;

                string resultadot_20 = (suma_20.ToString(@"h\:mm"));
                if (resultadot_20 != "0:00")
                {
                    worksheet.Cells[31, 6] = resultadot_20;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 899, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_20));
                    if (Extra20.Checked == true)
                    {
                        if (suma_20 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_20 = suma_20 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_20);
                            worksheet.Cells[31, 8] = RestaExtra_20.ToString(@"h\:mm");
                            worksheet.Cells[31, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[31, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[31, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_20);
                            worksheet.Cells[31, 8] = suma_20.ToString(@"h\:mm");
                            worksheet.Cells[31, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_20);
                        worksheet.Cells[31, 7] = suma_20.ToString(@"h\:mm");
                        worksheet.Cells[31, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_21
                TimeSpan fecha1_21 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_21 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_21 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_21 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_21 = TimeSpan.Parse("0:00");
                if (EM21.Text != "")
                {
                    fecha1_21 = TimeSpan.Parse(EM21.Text);
                    worksheet.Cells[32, 2] = EM21.Text;
                }
                if (SM21.Text != "")
                {
                    fecha2_21 = TimeSpan.Parse(SM21.Text);
                    worksheet.Cells[32, 3] = SM21.Text;
                }
                if (ET21.Text != "")
                {
                    fecha3_21 = TimeSpan.Parse(ET21.Text);
                    worksheet.Cells[32, 4] = ET21.Text;
                }
                if (ST21.Text != "")
                {
                    fecha4_21 = TimeSpan.Parse(ST21.Text);
                    worksheet.Cells[32, 5] = ST21.Text;
                }
                TimeSpan restar_21 = fecha2_21 - fecha1_21;
                TimeSpan restax_21 = fecha4_21 - fecha3_21;
                TimeSpan suma_21 = restar_21 + restax_21;

                string resultadot_21 = (suma_21.ToString(@"h\:mm"));
                if (resultadot_21 != "0:00")
                {
                    worksheet.Cells[32, 6] = resultadot_21;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 926, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_21));
                    if (Extra21.Checked == true)
                    {
                        if (suma_21 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_21 = suma_21 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_21);
                            worksheet.Cells[32, 8] = RestaExtra_21.ToString(@"h\:mm");
                            worksheet.Cells[32, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[32, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[32, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_21);
                            worksheet.Cells[32, 8] = suma_21.ToString(@"h\:mm");
                            worksheet.Cells[32, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_21);
                        worksheet.Cells[32, 7] = suma_21.ToString(@"h\:mm");
                        worksheet.Cells[32, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_22
                TimeSpan fecha1_22 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_22 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_22 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_22 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_22 = TimeSpan.Parse("0:00");
                if (EM22.Text != "")
                {
                    fecha1_22 = TimeSpan.Parse(EM22.Text);
                    worksheet.Cells[33, 2] = EM22.Text;
                }
                if (SM22.Text != "")
                {
                    fecha2_22 = TimeSpan.Parse(SM22.Text);
                    worksheet.Cells[33, 3] = SM22.Text;
                }
                if (ET22.Text != "")
                {
                    fecha3_22 = TimeSpan.Parse(ET22.Text);
                    worksheet.Cells[33, 4] = ET22.Text;
                }
                if (ST22.Text != "")
                {
                    fecha4_22 = TimeSpan.Parse(ST22.Text);
                    worksheet.Cells[33, 5] = ST22.Text;
                }
                TimeSpan restar_22 = fecha2_22 - fecha1_22;
                TimeSpan restax_22 = fecha4_22 - fecha3_22;
                TimeSpan suma_22 = restar_22 + restax_22;

                string resultadot_22 = (suma_22.ToString(@"h\:mm"));
                if (resultadot_22 != "0:00")
                {
                    worksheet.Cells[33, 6] = resultadot_22;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 954, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_22));
                    if (Extra22.Checked == true)
                    {
                        if (suma_22 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_22 = suma_22 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_22);
                            worksheet.Cells[33, 8] = RestaExtra_22.ToString(@"h\:mm");
                            worksheet.Cells[33, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[33, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[33, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_22);
                            worksheet.Cells[33, 8] = suma_22.ToString(@"h\:mm");
                            worksheet.Cells[33, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_22);
                        worksheet.Cells[33, 7] = suma_22.ToString(@"h\:mm");
                        worksheet.Cells[33, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_23
                TimeSpan fecha1_23 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_23 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_23 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_23 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_23 = TimeSpan.Parse("0:00");
                if (EM23.Text != "")
                {
                    fecha1_23 = TimeSpan.Parse(EM23.Text);
                    worksheet.Cells[34, 2] = EM23.Text;
                }
                if (SM23.Text != "")
                {
                    fecha2_23 = TimeSpan.Parse(SM23.Text);
                    worksheet.Cells[34, 3] = SM23.Text;
                }
                if (ET23.Text != "")
                {
                    fecha3_23 = TimeSpan.Parse(ET23.Text);
                    worksheet.Cells[34, 4] = ET23.Text;
                }
                if (ST23.Text != "")
                {
                    fecha4_23 = TimeSpan.Parse(ST23.Text);
                    worksheet.Cells[34, 5] = ST23.Text;
                }
                TimeSpan restar_23 = fecha2_23 - fecha1_23;
                TimeSpan restax_23 = fecha4_23 - fecha3_23;
                TimeSpan suma_23 = restar_23 + restax_23;

                string resultadot_23 = (suma_23.ToString(@"h\:mm"));
                if (resultadot_23 != "0:00")
                {
                    worksheet.Cells[34, 6] = resultadot_23;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 982, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_23));
                    if (Extra23.Checked == true)
                    {
                        if (suma_23 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_23 = suma_23 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_23);
                            worksheet.Cells[34, 8] = RestaExtra_23.ToString(@"h\:mm");
                            worksheet.Cells[34, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[34, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[34, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_23);
                            worksheet.Cells[34, 8] = suma_23.ToString(@"h\:mm");
                            worksheet.Cells[34, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_23);
                        worksheet.Cells[34, 7] = suma_23.ToString(@"h\:mm");
                        worksheet.Cells[34, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_24
                TimeSpan fecha1_24 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_24 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_24 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_24 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_24 = TimeSpan.Parse("0:00");
                if (EM24.Text != "")
                {
                    fecha1_24 = TimeSpan.Parse(EM24.Text);
                    worksheet.Cells[35, 2] = EM24.Text;
                }
                if (SM24.Text != "")
                {
                    fecha2_24 = TimeSpan.Parse(SM24.Text);
                    worksheet.Cells[35, 3] = SM24.Text;
                }
                if (ET24.Text != "")
                {
                    fecha3_24 = TimeSpan.Parse(ET24.Text);
                    worksheet.Cells[35, 4] = ET24.Text;
                }
                if (ST24.Text != "")
                {
                    fecha4_24 = TimeSpan.Parse(ST24.Text);
                    worksheet.Cells[35, 5] = ST24.Text;
                }
                TimeSpan restar_24 = fecha2_24 - fecha1_24;
                TimeSpan restax_24 = fecha4_24 - fecha3_24;
                TimeSpan suma_24 = restar_24 + restax_24;

                string resultadot_24 = (suma_24.ToString(@"h\:mm"));
                if (resultadot_24 != "0:00")
                {
                    worksheet.Cells[35, 6] = resultadot_24;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 1009, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_24));
                    if (Extra24.Checked == true)
                    {
                        if (suma_24 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_24 = suma_24 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_24);
                            worksheet.Cells[35, 8] = RestaExtra_24.ToString(@"h\:mm");
                            worksheet.Cells[35, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[35, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[35, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_24);
                            worksheet.Cells[35, 8] = suma_24.ToString(@"h\:mm");
                            worksheet.Cells[35, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_24);
                        worksheet.Cells[35, 7] = suma_24.ToString(@"h\:mm");
                        worksheet.Cells[35, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_25
                TimeSpan fecha1_25 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_25 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_25 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_25 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_25 = TimeSpan.Parse("0:00");
                if (EM25.Text != "")
                {
                    fecha1_25 = TimeSpan.Parse(EM25.Text);
                    worksheet.Cells[36, 2] = EM25.Text;
                }
                if (SM25.Text != "")
                {
                    fecha2_25 = TimeSpan.Parse(SM25.Text);
                    worksheet.Cells[36, 3] = SM25.Text;
                }
                if (ET25.Text != "")
                {
                    fecha3_25 = TimeSpan.Parse(ET25.Text);
                    worksheet.Cells[36, 4] = ET25.Text;
                }
                if (ST25.Text != "")
                {
                    fecha4_25 = TimeSpan.Parse(ST25.Text);
                    worksheet.Cells[36, 5] = ST25.Text;
                }
                TimeSpan restar_25 = fecha2_25 - fecha1_25;
                TimeSpan restax_25 = fecha4_25 - fecha3_25;
                TimeSpan suma_25 = restar_25 + restax_25;

                string resultadot_25 = (suma_25.ToString(@"h\:mm"));
                if (resultadot_25 != "0:00")
                {
                    worksheet.Cells[36, 6] = resultadot_25;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 1037, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_25));
                    if (Extra25.Checked == true)
                    {
                        if (suma_25 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_25 = suma_25 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_25);
                            worksheet.Cells[36, 8] = RestaExtra_25.ToString(@"h\:mm");
                            worksheet.Cells[36, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[36, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[36, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_25);
                            worksheet.Cells[36, 8] = suma_25.ToString(@"h\:mm");
                            worksheet.Cells[36, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_25);
                        worksheet.Cells[36, 7] = suma_25.ToString(@"h\:mm");
                        worksheet.Cells[36, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_26
                TimeSpan fecha1_26 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_26 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_26 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_26 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_26 = TimeSpan.Parse("0:00");
                if (EM26.Text != "")
                {
                    fecha1_26 = TimeSpan.Parse(EM26.Text);
                    worksheet.Cells[37, 2] = EM26.Text;
                }
                if (SM26.Text != "")
                {
                    fecha2_26 = TimeSpan.Parse(SM26.Text);
                    worksheet.Cells[37, 3] = SM26.Text;
                }
                if (ET26.Text != "")
                {
                    fecha3_26 = TimeSpan.Parse(ET26.Text);
                    worksheet.Cells[37, 4] = ET26.Text;
                }
                if (ST26.Text != "")
                {
                    fecha4_26 = TimeSpan.Parse(ST26.Text);
                    worksheet.Cells[37, 5] = ST26.Text;
                }
                TimeSpan restar_26 = fecha2_26 - fecha1_26;
                TimeSpan restax_26 = fecha4_26 - fecha3_26;
                TimeSpan suma_26 = restar_26 + restax_26;

                string resultadot_26 = (suma_26.ToString(@"h\:mm"));
                if (resultadot_26 != "0:00")
                {
                    worksheet.Cells[37, 6] = resultadot_26;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 1066, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_26));
                    if (Extra26.Checked == true)
                    {
                        if (suma_26 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_26 = suma_26 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_26);
                            worksheet.Cells[37, 8] = RestaExtra_26.ToString(@"h\:mm");
                            worksheet.Cells[37, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[37, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[37, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_26);
                            worksheet.Cells[37, 8] = suma_26.ToString(@"h\:mm");
                            worksheet.Cells[37, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_26);
                        worksheet.Cells[37, 7] = suma_26.ToString(@"h\:mm");
                        worksheet.Cells[37, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_27
                TimeSpan fecha1_27 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_27 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_27 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_27 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_27 = TimeSpan.Parse("0:00");
                if (EM27.Text != "")
                {
                    fecha1_27 = TimeSpan.Parse(EM27.Text);
                    worksheet.Cells[38, 2] = EM27.Text;
                }
                if (SM27.Text != "")
                {
                    fecha2_27 = TimeSpan.Parse(SM27.Text);
                    worksheet.Cells[38, 3] = SM27.Text;
                }
                if (ET27.Text != "")
                {
                    fecha3_27 = TimeSpan.Parse(ET27.Text);
                    worksheet.Cells[38, 4] = ET27.Text;
                }
                if (ST27.Text != "")
                {
                    fecha4_27 = TimeSpan.Parse(ST27.Text);
                    worksheet.Cells[38, 5] = ST27.Text;
                }
                TimeSpan restar_27 = fecha2_27 - fecha1_27;
                TimeSpan restax_27 = fecha4_27 - fecha3_27;
                TimeSpan suma_27 = restar_27 + restax_27;

                string resultadot_27 = (suma_27.ToString(@"h\:mm"));
                if (resultadot_27 != "0:00")
                {
                    worksheet.Cells[38, 6] = resultadot_27;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 1094, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_27));
                    if (Extra27.Checked == true)
                    {
                        if (suma_27 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_27 = suma_27 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_27);
                            worksheet.Cells[38, 8] = RestaExtra_27.ToString(@"h\:mm");
                            worksheet.Cells[38, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[38, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[38, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_27);
                            worksheet.Cells[38, 8] = suma_27.ToString(@"h\:mm");
                            worksheet.Cells[38, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_27);
                        worksheet.Cells[38, 7] = suma_27.ToString(@"h\:mm");
                        worksheet.Cells[38, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_28
                TimeSpan fecha1_28 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_28 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_28 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_28 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_28 = TimeSpan.Parse("0:00");
                if (EM28.Text != "")
                {
                    fecha1_28 = TimeSpan.Parse(EM28.Text);
                    worksheet.Cells[39, 2] = EM28.Text;
                }
                if (SM28.Text != "")
                {
                    fecha2_28 = TimeSpan.Parse(SM28.Text);
                    worksheet.Cells[39, 3] = SM28.Text;
                }
                if (ET28.Text != "")
                {
                    fecha3_28 = TimeSpan.Parse(ET28.Text);
                    worksheet.Cells[39, 4] = ET28.Text;
                }
                if (ST28.Text != "")
                {
                    fecha4_28 = TimeSpan.Parse(ST28.Text);
                    worksheet.Cells[39, 5] = ST28.Text;
                }
                TimeSpan restar_28 = fecha2_28 - fecha1_28;
                TimeSpan restax_28 = fecha4_28 - fecha3_28;
                TimeSpan suma_28 = restar_28 + restax_28;

                string resultadot_28 = (suma_28.ToString(@"h\:mm"));
                if (resultadot_28 != "0:00")
                {
                    worksheet.Cells[39, 6] = resultadot_28;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 1121, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_28));
                    if (Extra28.Checked == true)
                    {
                        if (suma_28 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_28 = suma_28 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_28);
                            worksheet.Cells[39, 8] = RestaExtra_28.ToString(@"h\:mm");
                            worksheet.Cells[39, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[39, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[39, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_28);
                            worksheet.Cells[39, 8] = suma_28.ToString(@"h\:mm");
                            worksheet.Cells[39, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_28);
                        worksheet.Cells[39, 7] = suma_28.ToString(@"h\:mm");
                        worksheet.Cells[39, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_29
                TimeSpan fecha1_29 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_29 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_29 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_29 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_29 = TimeSpan.Parse("0:00");
                if (EM29.Text != "")
                {
                    fecha1_29 = TimeSpan.Parse(EM29.Text);
                    worksheet.Cells[40, 2] = EM29.Text;
                }
                if (SM29.Text != "")
                {
                    fecha2_29 = TimeSpan.Parse(SM29.Text);
                    worksheet.Cells[40, 3] = SM29.Text;
                }
                if (ET29.Text != "")
                {
                    fecha3_29 = TimeSpan.Parse(ET29.Text);
                    worksheet.Cells[40, 4] = ET29.Text;
                }
                if (ST29.Text != "")
                {
                    fecha4_29 = TimeSpan.Parse(ST29.Text);
                    worksheet.Cells[40, 5] = ST29.Text;
                }
                TimeSpan restar_29 = fecha2_29 - fecha1_29;
                TimeSpan restax_29 = fecha4_29 - fecha3_29;
                TimeSpan suma_29 = restar_29 + restax_29;

                string resultadot_29 = (suma_29.ToString(@"h\:mm"));
                if (resultadot_29 != "0:00")
                {
                    worksheet.Cells[40, 6] = resultadot_29;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 1149, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_29));
                    if (Extra29.Checked == true)
                    {
                        if (suma_29 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_29 = suma_29 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_29);
                            worksheet.Cells[40, 8] = RestaExtra_29.ToString(@"h\:mm");
                            worksheet.Cells[40, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[40, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[40, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_29);
                            worksheet.Cells[40, 8] = suma_29.ToString(@"h\:mm");
                            worksheet.Cells[40, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_29);
                        worksheet.Cells[40, 7] = suma_29.ToString(@"h\:mm");
                        worksheet.Cells[40, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_30
                TimeSpan fecha1_30 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_30 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_30 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_30 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_30 = TimeSpan.Parse("0:00");
                if (EM30.Text != "")
                {
                    fecha1_30 = TimeSpan.Parse(EM30.Text);
                    worksheet.Cells[41, 2] = EM30.Text;
                }
                if (SM30.Text != "")
                {
                    fecha2_30 = TimeSpan.Parse(SM30.Text);
                    worksheet.Cells[41, 3] = SM30.Text;
                }
                if (ET30.Text != "")
                {
                    fecha3_30 = TimeSpan.Parse(ET30.Text);
                    worksheet.Cells[41, 4] = ET30.Text;
                }
                if (ST30.Text != "")
                {
                    fecha4_30 = TimeSpan.Parse(ST30.Text);
                    worksheet.Cells[41, 5] = ST30.Text;
                }
                TimeSpan restar_30 = fecha2_30 - fecha1_30;
                TimeSpan restax_30 = fecha4_30 - fecha3_30;
                TimeSpan suma_30 = restar_30 + restax_30;

                string resultadot_30 = (suma_30.ToString(@"h\:mm"));
                if (resultadot_30 != "0:00")
                {
                    worksheet.Cells[41, 6] = resultadot_30;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 1176, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_30));
                    if (Extra30.Checked == true)
                    {
                        if (suma_30 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_30 = suma_30 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_30);
                            worksheet.Cells[41, 8] = RestaExtra_30.ToString(@"h\:mm");
                            worksheet.Cells[41, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[41, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[41, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_30);
                            worksheet.Cells[41, 8] = suma_30.ToString(@"h\:mm");
                            worksheet.Cells[41, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_30);
                        worksheet.Cells[41, 7] = suma_30.ToString(@"h\:mm");
                        worksheet.Cells[41, 7].NumberFormat = "[h]:mm";
                    }
                }

                //_31
                TimeSpan fecha1_31 = TimeSpan.Parse("0:00");
                TimeSpan fecha2_31 = TimeSpan.Parse("0:00");
                TimeSpan fecha3_31 = TimeSpan.Parse("0:00");
                TimeSpan fecha4_31 = TimeSpan.Parse("0:00");
                TimeSpan RestaExtra_31 = TimeSpan.Parse("0:00");
                if (EM31.Text != "")
                {
                    fecha1_31 = TimeSpan.Parse(EM31.Text);
                    worksheet.Cells[42, 2] = EM31.Text;
                }
                if (SM31.Text != "")
                {
                    fecha2_31 = TimeSpan.Parse(SM31.Text);
                    worksheet.Cells[42, 3] = SM31.Text;
                }
                if (ET31.Text != "")
                {
                    fecha3_31 = TimeSpan.Parse(ET31.Text);
                    worksheet.Cells[42, 4] = ET31.Text;
                }
                if (ST31.Text != "")
                {
                    fecha4_31 = TimeSpan.Parse(ST31.Text);
                    worksheet.Cells[42, 5] = ST31.Text;
                }
                TimeSpan restar_31 = fecha2_31 - fecha1_31;
                TimeSpan restax_31 = fecha4_31 - fecha3_31;
                TimeSpan suma_31 = restar_31 + restax_31;

                string resultadot_31 = (suma_31.ToString(@"h\:mm"));
                if (resultadot_31 != "0:00")
                {
                    worksheet.Cells[42, 6] = resultadot_31;
                    worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 584, 1204, 77, 27);
                    SumaHorasJornada.Add(TimeSpan.Parse(resultadot_31));
                    if (Extra31.Checked == true)
                    {
                        if (suma_31 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                        {
                            RestaExtra_31 = suma_31 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                            SumaHorasExtra.Add(RestaExtra_31);
                            worksheet.Cells[42, 8] = RestaExtra_31.ToString(@"h\:mm");
                            worksheet.Cells[42, 8].NumberFormat = "[h]:mm";

                            SumaHoras.Add(TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                            worksheet.Cells[42, 7] = Properties.Settings.Default.HorasdeJornada;
                            worksheet.Cells[42, 7].NumberFormat = "[h]:mm";
                        }
                        else
                        {
                            SumaHorasExtra.Add(suma_31);
                            worksheet.Cells[42, 8] = suma_31.ToString(@"h\:mm");
                            worksheet.Cells[42, 8].NumberFormat = "[h]:mm";
                        }
                    }
                    else
                    {
                        SumaHoras.Add(suma_31);
                        worksheet.Cells[42, 7] = suma_31.ToString(@"h\:mm");
                        worksheet.Cells[42, 7].NumberFormat = "[h]:mm";
                    }
                }

                //TOTAL HORAS JORNADA
                TimeSpan tiempol = new TimeSpan();
                for (int i = 0; i < SumaHorasJornada.Count; i++)
                {
                    tiempol += SumaHorasJornada[i];
                }

                var Minse = string.Format("{0:D2}", tiempol.Minutes);
                var Hourse = string.Format("{0:D2}", tiempol.Hours);
                var Dayse = string.Format("{0:D2}", tiempol.Days);
                int MulDaye = Convert.ToInt32(Dayse) * 24;
                int Calce = MulDaye + Convert.ToInt32(Hourse);
                string Finishe = Calce.ToString() + ":" + Minse;

                worksheet.Cells[43, 6] = Finishe;
                worksheet.Cells[43, 6].NumberFormat = "[h]:mm";

                //TOTAL HORAS ORDINARIAS
                TimeSpan tiempo = new TimeSpan();
                for (int i = 0; i < SumaHoras.Count; i++)
                {
                    tiempo += SumaHoras[i];
                }

                var Mins = string.Format("{0:D2}", tiempo.Minutes);
                var Hours = string.Format("{0:D2}", tiempo.Hours);
                var Days = string.Format("{0:D2}", tiempo.Days);
                int MulDay = Convert.ToInt32(Days) * 24;
                int Calc = MulDay + Convert.ToInt32(Hours);
                string Finish = Calc.ToString() + ":" + Mins;

                worksheet.Cells[43, 7] = Finish;
                worksheet.Cells[43, 7].NumberFormat = "[h]:mm";

                //TOTAL HORAS EXTRAS
                TimeSpan tiempu = new TimeSpan();
                for (int i = 0; i < SumaHorasExtra.Count; i++)
                {
                    tiempu += SumaHorasExtra[i];
                }

                string FinishExtra = "";
                var Minso = string.Format("{0:D2}", tiempu.Minutes);
                var Hourso = string.Format("{0:D2}", tiempu.Hours);
                if (Hourso != "00")
                {
                    string HourChange = Hourso.Replace("0", "") + ":00";
                    TimeSpan Finisho = TimeSpan.Parse(HourChange);// - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                    FinishExtra = Finisho.ToString().Replace(":", "").Replace("0", "").Replace("-", "") + ":" + Minso;
                }
                else if (Minso != "00")
                {
                    FinishExtra = "00:" + Minso;
                }

                worksheet.Cells[43, 8] = FinishExtra;
                worksheet.Cells[43, 8].NumberFormat = "[h]:mm";

                worksheet.Shapes.AddPicture(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 60, 1316, 115, 40);

                try
                {
                    Excel.Application.ActiveWorkbook.Save();
                    Excel.Application.Quit();
                    Excel.Quit();
                    Properties.Settings.Default.UltimoArchivoGenerado = System.Windows.Forms.Application.StartupPath + @"\Jornadas\Registro de Jornada de " + MesPreSelected + " del " + YearSelected.ToString() + ".xlsx";

                    Properties.Settings.Default.Save();
                    Properties.Settings.Default.Reload();
                    if(FileSend != "1")
                    {
                        MessageBox.Show("Documento creado y firmado.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        
                    }
                    Properties.Settings.Default.Installing = "0";
                }
                catch (Exception ex)
                {
                    MessageBox.Show("ERROR: " + ex, "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Properties.Settings.Default.Installing = "0";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex, "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show("Revisa si has introducido algún caracter incorrecto en algún campo.", "Tesalia Redes - Registro de jornada", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Properties.Settings.Default.Installing = "0";
            }
        }
        #endregion

        #region OPTIONS
        private void Form1_Load(object sender, EventArgs e)
        {
            Version.Text = "v" + Application.ProductVersion + " - Comprobar actualizaciones";
            
            FechaConstruct();
            FormLoaded = 1;

            if (Properties.Settings.Default.Mes != "0")
            {
                Meses1.Text = Properties.Settings.Default.Mes;
                Meses2.Text = Properties.Settings.Default.Mes;
            }
            if (Properties.Settings.Default.Year != "0")
            {
                Years1.Text = Properties.Settings.Default.Year;
                Years2.Text = Properties.Settings.Default.Year;
            }

            if (Properties.Settings.Default.EnviarMailCC == "1")
            {
                EnviarCopiaCC.Checked = true;
            }

            if (Properties.Settings.Default.StartPage == "0")
            {
                Pages.SelectedIndex = 0;
                StartPage.Text = StartPage.Items[0].ToString();
                Position.Location = new Point(HomeBTN.Location.X, HomeBTN.Location.Y);
            }
            else if (Properties.Settings.Default.StartPage == "1")
            {
                Pages.SelectedIndex = 1;
                StartPage.Text = StartPage.Items[1].ToString();
                Position.Location = new Point(AccountVTN.Location.X, AccountVTN.Location.Y);
            }
            else if (Properties.Settings.Default.StartPage == "2")
            {
                Pages.SelectedIndex = 3;
                StartPage.Text = StartPage.Items[2].ToString();
                Position.Location = new Point(JornadaBTN.Location.X, JornadaBTN.Location.Y);
            }
            else if (Properties.Settings.Default.StartPage == "3")
            {
                Pages.SelectedIndex = 2;
                StartPage.Text = StartPage.Items[3].ToString();
                Position.Location = new Point(JornadaBTN.Location.X, JornadaBTN.Location.Y);
            }
            else if (Properties.Settings.Default.StartPage == "4")
            {
                Pages.SelectedIndex = 4;
                StartPage.Text = StartPage.Items[4].ToString();
                Position.Location = new Point(CorreoBTN.Location.X, CorreoBTN.Location.Y);
            }
            else if (Properties.Settings.Default.StartPage == "6")
            {
                Pages.SelectedIndex = 6;
                StartPage.Text = StartPage.Items[5].ToString();
                Position.Location = new Point(NominasBTN.Location.X, NominasBTN.Location.Y);
            }

            NombreTXT.Text = Properties.Settings.Default.Name;
            DNITXT.Text = Properties.Settings.Default.Documento;
            SSTXT.Text = Properties.Settings.Default.SeguridadS;
            HJornada.Text = Properties.Settings.Default.HorasdeJornada;
            Workin.Text = Properties.Settings.Default.Trabajo;
            if (Properties.Settings.Default.RutaFirma != "")
            {
                Firma.BackgroundImage = Image.FromFile(Properties.Settings.Default.RutaFirma);
            }
            CorreoTXT.Text = Properties.Settings.Default.MailMail;
            PassTXT.Text = Properties.Settings.Default.MailPass;

            if (Properties.Settings.Default.ActivarPassCC == "1")
            {
                ActivarPass.Checked = true;
                SavePass.Enabled = true;
                StartPass.Enabled = true;
                StartPass.Text = Properties.Settings.Default.PassStart;
            }
        }
        private void AutoJornada_Click(object sender, EventArgs e)
        {
            if (Auto1.Text != "")
            {
                if (Auto2.Text != "")
                {
                    if (Auto3.Text != "")
                    {
                        if (Auto4.Text != "")
                        {
                            EM1.Text = Auto1.Text;
                            SM1.Text = Auto2.Text;
                            ET1.Text = Auto3.Text;
                            ST1.Text = Auto4.Text;

                            EM2.Text = Auto1.Text;
                            SM2.Text = Auto2.Text;
                            ET2.Text = Auto3.Text;
                            ST2.Text = Auto4.Text;

                            EM3.Text = Auto1.Text;
                            SM3.Text = Auto2.Text;
                            ET3.Text = Auto3.Text;
                            ST3.Text = Auto4.Text;

                            EM4.Text = Auto1.Text;
                            SM4.Text = Auto2.Text;
                            ET4.Text = Auto3.Text;
                            ST4.Text = Auto4.Text;

                            EM5.Text = Auto1.Text;
                            SM5.Text = Auto2.Text;
                            ET5.Text = Auto3.Text;
                            ST5.Text = Auto4.Text;

                            EM6.Text = Auto1.Text;
                            SM6.Text = Auto2.Text;
                            ET6.Text = Auto3.Text;
                            ST6.Text = Auto4.Text;

                            EM7.Text = Auto1.Text;
                            SM7.Text = Auto2.Text;
                            ET7.Text = Auto3.Text;
                            ST7.Text = Auto4.Text;

                            EM8.Text = Auto1.Text;
                            SM8.Text = Auto2.Text;
                            ET8.Text = Auto3.Text;
                            ST8.Text = Auto4.Text;

                            EM9.Text = Auto1.Text;
                            SM9.Text = Auto2.Text;
                            ET9.Text = Auto3.Text;
                            ST9.Text = Auto4.Text;

                            EM10.Text = Auto1.Text;
                            SM10.Text = Auto2.Text;
                            ET10.Text = Auto3.Text;
                            ST10.Text = Auto4.Text;

                            EM11.Text = Auto1.Text;
                            SM11.Text = Auto2.Text;
                            ET11.Text = Auto3.Text;
                            ST11.Text = Auto4.Text;

                            EM12.Text = Auto1.Text;
                            SM12.Text = Auto2.Text;
                            ET12.Text = Auto3.Text;
                            ST12.Text = Auto4.Text;

                            EM13.Text = Auto1.Text;
                            SM13.Text = Auto2.Text;
                            ET13.Text = Auto3.Text;
                            ST13.Text = Auto4.Text;

                            EM14.Text = Auto1.Text;
                            SM14.Text = Auto2.Text;
                            ET14.Text = Auto3.Text;
                            ST14.Text = Auto4.Text;

                            EM15.Text = Auto1.Text;
                            SM15.Text = Auto2.Text;
                            ET15.Text = Auto3.Text;
                            ST15.Text = Auto4.Text;

                            EM16.Text = Auto1.Text;
                            SM16.Text = Auto2.Text;
                            ET16.Text = Auto3.Text;
                            ST16.Text = Auto4.Text;

                            EM17.Text = Auto1.Text;
                            SM17.Text = Auto2.Text;
                            ET17.Text = Auto3.Text;
                            ST17.Text = Auto4.Text;

                            EM18.Text = Auto1.Text;
                            SM18.Text = Auto2.Text;
                            ET18.Text = Auto3.Text;
                            ST18.Text = Auto4.Text;

                            EM19.Text = Auto1.Text;
                            SM19.Text = Auto2.Text;
                            ET19.Text = Auto3.Text;
                            ST19.Text = Auto4.Text;

                            EM20.Text = Auto1.Text;
                            SM20.Text = Auto2.Text;
                            ET20.Text = Auto3.Text;
                            ST20.Text = Auto4.Text;

                            EM21.Text = Auto1.Text;
                            SM21.Text = Auto2.Text;
                            ET21.Text = Auto3.Text;
                            ST21.Text = Auto4.Text;

                            EM22.Text = Auto1.Text;
                            SM22.Text = Auto2.Text;
                            ET22.Text = Auto3.Text;
                            ST22.Text = Auto4.Text;

                            EM23.Text = Auto1.Text;
                            SM23.Text = Auto2.Text;
                            ET23.Text = Auto3.Text;
                            ST23.Text = Auto4.Text;

                            EM24.Text = Auto1.Text;
                            SM24.Text = Auto2.Text;
                            ET24.Text = Auto3.Text;
                            ST24.Text = Auto4.Text;

                            EM25.Text = Auto1.Text;
                            SM25.Text = Auto2.Text;
                            ET25.Text = Auto3.Text;
                            ST25.Text = Auto4.Text;

                            EM26.Text = Auto1.Text;
                            SM26.Text = Auto2.Text;
                            ET26.Text = Auto3.Text;
                            ST26.Text = Auto4.Text;

                            EM27.Text = Auto1.Text;
                            SM27.Text = Auto2.Text;
                            ET27.Text = Auto3.Text;
                            ST27.Text = Auto4.Text;

                            EM28.Text = Auto1.Text;
                            SM28.Text = Auto2.Text;
                            ET28.Text = Auto3.Text;
                            ST28.Text = Auto4.Text;

                            EM29.Text = Auto1.Text;
                            SM29.Text = Auto2.Text;
                            ET29.Text = Auto3.Text;
                            ST29.Text = Auto4.Text;

                            EM30.Text = Auto1.Text;
                            SM30.Text = Auto2.Text;
                            ET30.Text = Auto3.Text;
                            ST30.Text = Auto4.Text;

                            EM31.Text = Auto1.Text;
                            SM31.Text = Auto2.Text;
                            ET31.Text = Auto3.Text;
                            ST31.Text = Auto4.Text;
                        }
                    }
                }
            }

            Festivos();
        }
        #endregion

        #region MI CUENTA
        private void FirmaAdd_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            openFileDialog1.RestoreDirectory = false;
            openFileDialog1.Title = "Buscar Firma";
            openFileDialog1.Filter = "PNG|*.png";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.ShowDialog();
            try
            {
                if (openFileDialog1.FileName != "")
                {
                    File.Copy(openFileDialog1.FileName, System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png", true);
                    Firma.BackgroundImage = Image.FromFile(System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png");
                    RutaFirmaS = System.Windows.Forms.Application.StartupPath + @"\Resources\Firma.png";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERROR: " + ex, "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BorrarFirma_Click(object sender, EventArgs e)
        {
            Firma.BackgroundImage = null;
            RutaFirmaS = "";
        }
        private void SaveMiAccount_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Name = NombreTXT.Text;
            Properties.Settings.Default.Documento = DNITXT.Text;
            Properties.Settings.Default.SeguridadS = SSTXT.Text;
            Properties.Settings.Default.HorasdeJornada = HJornada.Text;
            Properties.Settings.Default.Trabajo = Workin.Text;
            Properties.Settings.Default.RutaFirma = RutaFirmaS;
            Properties.Settings.Default.MailMail = CorreoTXT.Text;
            Properties.Settings.Default.MailPass = PassTXT.Text;
            Properties.Settings.Default.Save();
            Properties.Settings.Default.Reload();
            MessageBox.Show("Datos guardados!", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void Deletemyaccount_Click(object sender, EventArgs e)
        {
            NombreTXT.Clear();
            DNITXT.Clear();
            SSTXT.Clear();
            HJornada.Text = null;
            CorreoTXT.Clear();
            PassTXT.Clear();
            Workin.Text = null;
        }
        #endregion

        #region AJUSTES
        private void STCheck1_Click(object sender, EventArgs e)
        {
            EnviarCopiaCC.Checked = true;
        }

        private void STCheck1_2_Click(object sender, EventArgs e)
        {
            EnviarCopiaCC.Checked = false;
        }

        private void EnviarCopiaCC_CheckedChanged(object sender, EventArgs e)
        {
            if (EnviarCopiaCC.Checked == true)
            {
                STCheck1.Visible = false;
                STCheck1_2.Visible = true;
                Properties.Settings.Default.EnviarMailCC = "1";
            }
            else
            {
                STCheck1.Visible = true;
                STCheck1_2.Visible = false;
                Properties.Settings.Default.EnviarMailCC = "0";
            }
            Properties.Settings.Default.Save();
        }

        private void StartPage_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (StartPage.SelectedIndex == 1)
            {
                Properties.Settings.Default.StartPage = "1";
            }
            else if (StartPage.SelectedIndex == 2)
            {
                Properties.Settings.Default.StartPage = "2";
            }
            else if (StartPage.SelectedIndex == 3)
            {
                Properties.Settings.Default.StartPage = "3";
            }
            else if (StartPage.SelectedIndex == 4)
            {
                if (Meses1.SelectedIndex != 9999999)
                {
                    MessageBox.Show("Esta función no está disponible.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    StartPage.SelectedIndex = 0;
                }
                else
                {
                    Properties.Settings.Default.StartPage = "6";
                }
            }
            else if (StartPage.SelectedIndex == 5)
            {
                if (Meses1.SelectedIndex != 9999999)
                {
                    MessageBox.Show("Esta función no está disponible.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    StartPage.SelectedIndex = 0;
                }
                else
                {
                    Properties.Settings.Default.StartPage = "4";
                }
            }
            Properties.Settings.Default.Save();
        }
        private void ActivarPass_CheckedChanged(object sender, EventArgs e)
        {
            if (ActivarPass.Checked == true)
            {
                STCheck2.Visible = false;
                STCheck2_2.Visible = true;
                Properties.Settings.Default.ActivarPassCC = "1";
                StartPass.Enabled = true;
                SavePass.Enabled = true;
            }
            else
            {
                STCheck2.Visible = true;
                STCheck2_2.Visible = false;
                Properties.Settings.Default.ActivarPassCC = "0";
                StartPass.Enabled = false;
                SavePass.Enabled = false;
            }
            Properties.Settings.Default.Save();
        }
        private void SavePass_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(StartPass.Text) == false)
            {
                Properties.Settings.Default.PassStart = StartPass.Text;
                Properties.Settings.Default.Save();
                MessageBox.Show("Contraseña guardada!", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void STCheck2_Click(object sender, EventArgs e)
        {
            ActivarPass.Checked = true;
        }
        private void STCheck2_2_Click(object sender, EventArgs e)
        {
            ActivarPass.Checked = false;
        }
        #endregion

        #region GESTION DE JORNADAS
        private void AbrirJornada_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(JornadaFiles.Text) == false)
            {
                Process.Start(System.Windows.Forms.Application.StartupPath + @"\Jornadas\" + JornadaFiles.Text);
            }
        }
        private void EnviarJornada_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(JornadaFiles.Text) == false)
            {
                if (Properties.Settings.Default.Name != "")
                {
                    if (Properties.Settings.Default.Documento != "")
                    {
                        if (Properties.Settings.Default.SeguridadS != "")
                        {
                            if (Properties.Settings.Default.HorasdeJornada != "")
                            {
                                if (Properties.Settings.Default.RutaFirma != "")
                                {
                                    Properties.Settings.Default.UltimoArchivoGenerado = System.Windows.Forms.Application.StartupPath + @"\Jornadas\" + JornadaFiles.Text;

                                    Form3 fm3 = new Form3("Registro de Jornada de " + MesPreSelected, "Registro de la Jornada de: " + Properties.Settings.Default.Name + " de " + MesPreSelected);
                                    fm3.Show();
                                }
                                else
                                {
                                    MessageBox.Show("No hay ninguna firma cargada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show("Es necesario indicar las horas de la jornada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("No se ha introducido documento de seguridad en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se ha introducido el DNI en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No se ha introducido el nombre completo en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void EliminarJornada_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(JornadaFiles.Text) == false)
            {
                try
                {
                    File.Delete(System.Windows.Forms.Application.StartupPath + @"\Jornadas\" + JornadaFiles.Text);
                    JornadaFiles.Items.Clear();
                    DirectoryInfo di = new DirectoryInfo(System.Windows.Forms.Application.StartupPath + @"\Jornadas\");
                    FileInfo[] files = di.GetFiles("*.xlsx");
                    foreach (var file in files)
                    {
                        JornadaFiles.Items.Add(file);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        #endregion

        private void SendHExtra_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.Name != "")
            {
                if (Properties.Settings.Default.HorasdeJornada != "")
                {
                    List<TimeSpan> SumaHorasExtra2 = new List<TimeSpan>();
                    List<string> SumaHorasExtraDia = new List<string>();

                    //_1
                    TimeSpan fecha1_1 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_1 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_1 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_1 = TimeSpan.Parse("0:00");
                    if (EM1.Text != "")
                    {
                        fecha1_1 = TimeSpan.Parse(EM1.Text);
                    }
                    if (SM1.Text != "")
                    {
                        fecha2_1 = TimeSpan.Parse(SM1.Text);
                    }
                    if (ET1.Text != "")
                    {
                        fecha3_1 = TimeSpan.Parse(ET1.Text);
                    }
                    if (ST1.Text != "")
                    {
                        fecha4_1 = TimeSpan.Parse(ST1.Text);
                    }
                    TimeSpan restar_1 = fecha2_1 - fecha1_1;
                    TimeSpan restax_1 = fecha4_1 - fecha3_1;
                    TimeSpan suma_1 = restar_1 + restax_1;

                    if (suma_1.ToString() != "0:00")
                    {
                        if (Extra1.Checked == true)
                        {
                            if (suma_1 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_1 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day1.Text + ": " + (suma_1 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_1);
                                SumaHorasExtraDia.Add(Day1.Text + ": " + suma_1.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_2
                    TimeSpan fecha1_2 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_2 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_2 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_2 = TimeSpan.Parse("0:00");
                    if (EM2.Text != "")
                    {
                        fecha1_2 = TimeSpan.Parse(EM2.Text);
                    }
                    if (SM2.Text != "")
                    {
                        fecha2_2 = TimeSpan.Parse(SM2.Text);
                    }
                    if (ET2.Text != "")
                    {
                        fecha3_2 = TimeSpan.Parse(ET2.Text);
                    }
                    if (ST2.Text != "")
                    {
                        fecha4_2 = TimeSpan.Parse(ST2.Text);
                    }
                    TimeSpan restar_2 = fecha2_2 - fecha1_2;
                    TimeSpan restax_2 = fecha4_2 - fecha3_2;
                    TimeSpan suma_2 = restar_2 + restax_2;

                    if (suma_2.ToString() != "0:00")
                    {
                        if (Extra2.Checked == true)
                        {
                            if (suma_2 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_2 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day2.Text + ": " + (suma_2 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_2);
                                SumaHorasExtraDia.Add(Day2.Text + ": " + suma_2.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_3
                    TimeSpan fecha1_3 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_3 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_3 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_3 = TimeSpan.Parse("0:00");
                    if (EM3.Text != "")
                    {
                        fecha1_3 = TimeSpan.Parse(EM3.Text);
                    }
                    if (SM3.Text != "")
                    {
                        fecha2_3 = TimeSpan.Parse(SM3.Text);
                    }
                    if (ET3.Text != "")
                    {
                        fecha3_3 = TimeSpan.Parse(ET3.Text);
                    }
                    if (ST3.Text != "")
                    {
                        fecha4_3 = TimeSpan.Parse(ST3.Text);
                    }
                    TimeSpan restar_3 = fecha2_3 - fecha1_3;
                    TimeSpan restax_3 = fecha4_3 - fecha3_3;
                    TimeSpan suma_3 = restar_3 + restax_3;

                    if (suma_3.ToString() != "0:00")
                    {
                        if (Extra3.Checked == true)
                        {
                            if (suma_3 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_3 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day3.Text + ": " + (suma_3 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_3);
                                SumaHorasExtraDia.Add(Day3.Text + ": " + suma_3.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_4
                    TimeSpan fecha1_4 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_4 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_4 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_4 = TimeSpan.Parse("0:00");
                    if (EM4.Text != "")
                    {
                        fecha1_4 = TimeSpan.Parse(EM4.Text);
                    }
                    if (SM4.Text != "")
                    {
                        fecha2_4 = TimeSpan.Parse(SM4.Text);
                    }
                    if (ET4.Text != "")
                    {
                        fecha3_4 = TimeSpan.Parse(ET4.Text);
                    }
                    if (ST4.Text != "")
                    {
                        fecha4_4 = TimeSpan.Parse(ST4.Text);
                    }
                    TimeSpan restar_4 = fecha2_4 - fecha1_4;
                    TimeSpan restax_4 = fecha4_4 - fecha3_4;
                    TimeSpan suma_4 = restar_4 + restax_4;

                    if (suma_4.ToString() != "0:00")
                    {
                        if (Extra4.Checked == true)
                        {
                            if (suma_4 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_4 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day4.Text + ": " + (suma_4 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_4);
                                SumaHorasExtraDia.Add(Day4.Text + ": " + suma_4.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_5
                    TimeSpan fecha1_5 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_5 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_5 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_5 = TimeSpan.Parse("0:00");
                    if (EM5.Text != "")
                    {
                        fecha1_5 = TimeSpan.Parse(EM5.Text);
                    }
                    if (SM5.Text != "")
                    {
                        fecha2_5 = TimeSpan.Parse(SM5.Text);
                    }
                    if (ET5.Text != "")
                    {
                        fecha3_5 = TimeSpan.Parse(ET5.Text);
                    }
                    if (ST5.Text != "")
                    {
                        fecha4_5 = TimeSpan.Parse(ST5.Text);
                    }
                    TimeSpan restar_5 = fecha2_5 - fecha1_5;
                    TimeSpan restax_5 = fecha4_5 - fecha3_5;
                    TimeSpan suma_5 = restar_5 + restax_5;

                    if (suma_5.ToString() != "0:00")
                    {
                        if (Extra5.Checked == true)
                        {
                            if (suma_5 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_5 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day5.Text + ": " + (suma_5 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_5);
                                SumaHorasExtraDia.Add(Day5.Text + ": " + suma_5.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_6
                    TimeSpan fecha1_6 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_6 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_6 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_6 = TimeSpan.Parse("0:00");
                    if (EM6.Text != "")
                    {
                        fecha1_6 = TimeSpan.Parse(EM6.Text);
                    }
                    if (SM6.Text != "")
                    {
                        fecha2_6 = TimeSpan.Parse(SM6.Text);
                    }
                    if (ET6.Text != "")
                    {
                        fecha3_6 = TimeSpan.Parse(ET6.Text);
                    }
                    if (ST6.Text != "")
                    {
                        fecha4_6 = TimeSpan.Parse(ST6.Text);
                    }
                    TimeSpan restar_6 = fecha2_6 - fecha1_6;
                    TimeSpan restax_6 = fecha4_6 - fecha3_6;
                    TimeSpan suma_6 = restar_6 + restax_6;

                    if (suma_6.ToString() != "0:00")
                    {
                        if (Extra6.Checked == true)
                        {
                            if (suma_6 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_6 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day6.Text + ": " + (suma_6 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_6);
                                SumaHorasExtraDia.Add(Day6.Text + ": " + suma_6.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_7
                    TimeSpan fecha1_7 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_7 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_7 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_7 = TimeSpan.Parse("0:00");
                    if (EM7.Text != "")
                    {
                        fecha1_7 = TimeSpan.Parse(EM7.Text);
                    }
                    if (SM7.Text != "")
                    {
                        fecha2_7 = TimeSpan.Parse(SM7.Text);
                    }
                    if (ET7.Text != "")
                    {
                        fecha3_7 = TimeSpan.Parse(ET7.Text);
                    }
                    if (ST7.Text != "")
                    {
                        fecha4_7 = TimeSpan.Parse(ST7.Text);
                    }
                    TimeSpan restar_7 = fecha2_7 - fecha1_7;
                    TimeSpan restax_7 = fecha4_7 - fecha3_7;
                    TimeSpan suma_7 = restar_7 + restax_7;

                    if (suma_7.ToString() != "0:00")
                    {
                        if (Extra7.Checked == true)
                        {
                            if (suma_7 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_7 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day7.Text + ": " + (suma_7 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_7);
                                SumaHorasExtraDia.Add(Day7.Text + ": " + suma_7.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_8
                    TimeSpan fecha1_8 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_8 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_8 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_8 = TimeSpan.Parse("0:00");
                    if (EM8.Text != "")
                    {
                        fecha1_8 = TimeSpan.Parse(EM8.Text);
                    }
                    if (SM8.Text != "")
                    {
                        fecha2_8 = TimeSpan.Parse(SM8.Text);
                    }
                    if (ET8.Text != "")
                    {
                        fecha3_8 = TimeSpan.Parse(ET8.Text);
                    }
                    if (ST8.Text != "")
                    {
                        fecha4_8 = TimeSpan.Parse(ST8.Text);
                    }
                    TimeSpan restar_8 = fecha2_8 - fecha1_8;
                    TimeSpan restax_8 = fecha4_8 - fecha3_8;
                    TimeSpan suma_8 = restar_8 + restax_8;

                    if (suma_8.ToString() != "0:00")
                    {
                        if (Extra8.Checked == true)
                        {
                            if (suma_8 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_8 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day8.Text + ": " + (suma_8 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_8);
                                SumaHorasExtraDia.Add(Day8.Text + ": " + suma_8.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_9
                    TimeSpan fecha1_9 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_9 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_9 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_9 = TimeSpan.Parse("0:00");
                    if (EM9.Text != "")
                    {
                        fecha1_9 = TimeSpan.Parse(EM9.Text);
                    }
                    if (SM9.Text != "")
                    {
                        fecha2_9 = TimeSpan.Parse(SM9.Text);
                    }
                    if (ET9.Text != "")
                    {
                        fecha3_9 = TimeSpan.Parse(ET9.Text);
                    }
                    if (ST9.Text != "")
                    {
                        fecha4_9 = TimeSpan.Parse(ST9.Text);
                    }
                    TimeSpan restar_9 = fecha2_9 - fecha1_9;
                    TimeSpan restax_9 = fecha4_9 - fecha3_9;
                    TimeSpan suma_9 = restar_9 + restax_9;

                    if (suma_9.ToString() != "0:00")
                    {
                        if (Extra9.Checked == true)
                        {
                            if (suma_9 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_9 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day9.Text + ": " + (suma_9 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_9);
                                SumaHorasExtraDia.Add(Day9.Text + ": " + suma_9.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_10
                    TimeSpan fecha1_10 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_10 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_10 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_10 = TimeSpan.Parse("0:00");
                    if (EM10.Text != "")
                    {
                        fecha1_10 = TimeSpan.Parse(EM10.Text);
                    }
                    if (SM10.Text != "")
                    {
                        fecha2_10 = TimeSpan.Parse(SM10.Text);
                    }
                    if (ET10.Text != "")
                    {
                        fecha3_10 = TimeSpan.Parse(ET10.Text);
                    }
                    if (ST10.Text != "")
                    {
                        fecha4_10 = TimeSpan.Parse(ST10.Text);
                    }
                    TimeSpan restar_10 = fecha2_10 - fecha1_10;
                    TimeSpan restax_10 = fecha4_10 - fecha3_10;
                    TimeSpan suma_10 = restar_10 + restax_10;

                    if (suma_10.ToString() != "0:00")
                    {
                        if (Extra10.Checked == true)
                        {
                            if (suma_10 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_10 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day10.Text + ": " + (suma_10 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_10);
                                SumaHorasExtraDia.Add(Day10.Text + ": " + suma_10.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_11
                    TimeSpan fecha1_11 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_11 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_11 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_11 = TimeSpan.Parse("0:00");
                    if (EM11.Text != "")
                    {
                        fecha1_11 = TimeSpan.Parse(EM11.Text);
                    }
                    if (SM11.Text != "")
                    {
                        fecha2_11 = TimeSpan.Parse(SM11.Text);
                    }
                    if (ET11.Text != "")
                    {
                        fecha3_11 = TimeSpan.Parse(ET11.Text);
                    }
                    if (ST11.Text != "")
                    {
                        fecha4_11 = TimeSpan.Parse(ST11.Text);
                    }
                    TimeSpan restar_11 = fecha2_11 - fecha1_11;
                    TimeSpan restax_11 = fecha4_11 - fecha3_11;
                    TimeSpan suma_11 = restar_11 + restax_11;

                    if (suma_11.ToString() != "0:00")
                    {
                        if (Extra11.Checked == true)
                        {
                            if (suma_11 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_11 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day11.Text + ": " + (suma_11 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_11);
                                SumaHorasExtraDia.Add(Day11.Text + ": " + suma_11.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_12
                    TimeSpan fecha1_12 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_12 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_12 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_12 = TimeSpan.Parse("0:00");
                    if (EM12.Text != "")
                    {
                        fecha1_12 = TimeSpan.Parse(EM12.Text);
                    }
                    if (SM12.Text != "")
                    {
                        fecha2_12 = TimeSpan.Parse(SM12.Text);
                    }
                    if (ET12.Text != "")
                    {
                        fecha3_12 = TimeSpan.Parse(ET12.Text);
                    }
                    if (ST12.Text != "")
                    {
                        fecha4_12 = TimeSpan.Parse(ST12.Text);
                    }
                    TimeSpan restar_12 = fecha2_12 - fecha1_12;
                    TimeSpan restax_12 = fecha4_12 - fecha3_12;
                    TimeSpan suma_12 = restar_12 + restax_12;

                    if (suma_12.ToString() != "0:00")
                    {
                        if (Extra12.Checked == true)
                        {
                            if (suma_12 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_12 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day12.Text + ": " + (suma_12 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_12);
                                SumaHorasExtraDia.Add(Day12.Text + ": " + suma_12.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_13
                    TimeSpan fecha1_13 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_13 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_13 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_13 = TimeSpan.Parse("0:00");
                    if (EM13.Text != "")
                    {
                        fecha1_13 = TimeSpan.Parse(EM13.Text);
                    }
                    if (SM13.Text != "")
                    {
                        fecha2_13 = TimeSpan.Parse(SM13.Text);
                    }
                    if (ET13.Text != "")
                    {
                        fecha3_13 = TimeSpan.Parse(ET13.Text);
                    }
                    if (ST13.Text != "")
                    {
                        fecha4_13 = TimeSpan.Parse(ST13.Text);
                    }
                    TimeSpan restar_13 = fecha2_13 - fecha1_13;
                    TimeSpan restax_13 = fecha4_13 - fecha3_13;
                    TimeSpan suma_13 = restar_13 + restax_13;

                    if (suma_13.ToString() != "0:00")
                    {
                        if (Extra13.Checked == true)
                        {
                            if (suma_13 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_13 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day13.Text + ": " + (suma_13 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_13);
                                SumaHorasExtraDia.Add(Day13.Text + ": " + suma_13.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_14
                    TimeSpan fecha1_14 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_14 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_14 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_14 = TimeSpan.Parse("0:00");
                    if (EM14.Text != "")
                    {
                        fecha1_14 = TimeSpan.Parse(EM14.Text);
                    }
                    if (SM14.Text != "")
                    {
                        fecha2_14 = TimeSpan.Parse(SM14.Text);
                    }
                    if (ET14.Text != "")
                    {
                        fecha3_14 = TimeSpan.Parse(ET14.Text);
                    }
                    if (ST14.Text != "")
                    {
                        fecha4_14 = TimeSpan.Parse(ST14.Text);
                    }
                    TimeSpan restar_14 = fecha2_14 - fecha1_14;
                    TimeSpan restax_14 = fecha4_14 - fecha3_14;
                    TimeSpan suma_14 = restar_14 + restax_14;

                    if (suma_14.ToString() != "0:00")
                    {
                        if (Extra14.Checked == true)
                        {
                            if (suma_14 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_14 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day14.Text + ": " + (suma_14 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_14);
                                SumaHorasExtraDia.Add(Day14.Text + ": " + suma_14.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_15
                    TimeSpan fecha1_15 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_15 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_15 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_15 = TimeSpan.Parse("0:00");
                    if (EM15.Text != "")
                    {
                        fecha1_15 = TimeSpan.Parse(EM15.Text);
                    }
                    if (SM15.Text != "")
                    {
                        fecha2_15 = TimeSpan.Parse(SM15.Text);
                    }
                    if (ET15.Text != "")
                    {
                        fecha3_15 = TimeSpan.Parse(ET15.Text);
                    }
                    if (ST15.Text != "")
                    {
                        fecha4_15 = TimeSpan.Parse(ST15.Text);
                    }
                    TimeSpan restar_15 = fecha2_15 - fecha1_15;
                    TimeSpan restax_15 = fecha4_15 - fecha3_15;
                    TimeSpan suma_15 = restar_15 + restax_15;

                    if (suma_15.ToString() != "0:00")
                    {
                        if (Extra15.Checked == true)
                        {
                            if (suma_15 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_15 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day15.Text + ": " + (suma_15 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_15);
                                SumaHorasExtraDia.Add(Day15.Text + ": " + suma_15.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_16
                    TimeSpan fecha1_16 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_16 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_16 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_16 = TimeSpan.Parse("0:00");
                    if (EM16.Text != "")
                    {
                        fecha1_16 = TimeSpan.Parse(EM16.Text);
                    }
                    if (SM16.Text != "")
                    {
                        fecha2_16 = TimeSpan.Parse(SM16.Text);
                    }
                    if (ET16.Text != "")
                    {
                        fecha3_16 = TimeSpan.Parse(ET16.Text);
                    }
                    if (ST16.Text != "")
                    {
                        fecha4_16 = TimeSpan.Parse(ST16.Text);
                    }
                    TimeSpan restar_16 = fecha2_16 - fecha1_16;
                    TimeSpan restax_16 = fecha4_16 - fecha3_16;
                    TimeSpan suma_16 = restar_16 + restax_16;

                    if (suma_16.ToString() != "0:00")
                    {
                        if (Extra16.Checked == true)
                        {
                            if (suma_16 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_16 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day16.Text + ": " + (suma_16 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_16);
                                SumaHorasExtraDia.Add(Day16.Text + ": " + suma_16.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_17
                    TimeSpan fecha1_17 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_17 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_17 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_17 = TimeSpan.Parse("0:00");
                    if (EM17.Text != "")
                    {
                        fecha1_17 = TimeSpan.Parse(EM17.Text);
                    }
                    if (SM17.Text != "")
                    {
                        fecha2_17 = TimeSpan.Parse(SM17.Text);
                    }
                    if (ET17.Text != "")
                    {
                        fecha3_17 = TimeSpan.Parse(ET17.Text);
                    }
                    if (ST17.Text != "")
                    {
                        fecha4_17 = TimeSpan.Parse(ST17.Text);
                    }
                    TimeSpan restar_17 = fecha2_17 - fecha1_17;
                    TimeSpan restax_17 = fecha4_17 - fecha3_17;
                    TimeSpan suma_17 = restar_17 + restax_17;

                    if (suma_17.ToString() != "0:00")
                    {
                        if (Extra17.Checked == true)
                        {
                            if (suma_17 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_17 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day17.Text + ": " + (suma_17 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_17);
                                SumaHorasExtraDia.Add(Day17.Text + ": " + suma_17.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_18
                    TimeSpan fecha1_18 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_18 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_18 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_18 = TimeSpan.Parse("0:00");
                    if (EM18.Text != "")
                    {
                        fecha1_18 = TimeSpan.Parse(EM18.Text);
                    }
                    if (SM18.Text != "")
                    {
                        fecha2_18 = TimeSpan.Parse(SM18.Text);
                    }
                    if (ET18.Text != "")
                    {
                        fecha3_18 = TimeSpan.Parse(ET18.Text);
                    }
                    if (ST18.Text != "")
                    {
                        fecha4_18 = TimeSpan.Parse(ST18.Text);
                    }
                    TimeSpan restar_18 = fecha2_18 - fecha1_18;
                    TimeSpan restax_18 = fecha4_18 - fecha3_18;
                    TimeSpan suma_18 = restar_18 + restax_18;

                    if (suma_18.ToString() != "0:00")
                    {
                        if (Extra18.Checked == true)
                        {
                            if (suma_18 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_18 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day18.Text + ": " + (suma_18 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_18);
                                SumaHorasExtraDia.Add(Day18.Text + ": " + suma_18.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_19
                    TimeSpan fecha1_19 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_19 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_19 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_19 = TimeSpan.Parse("0:00");
                    if (EM19.Text != "")
                    {
                        fecha1_19 = TimeSpan.Parse(EM19.Text);
                    }
                    if (SM19.Text != "")
                    {
                        fecha2_19 = TimeSpan.Parse(SM19.Text);
                    }
                    if (ET19.Text != "")
                    {
                        fecha3_19 = TimeSpan.Parse(ET19.Text);
                    }
                    if (ST19.Text != "")
                    {
                        fecha4_19 = TimeSpan.Parse(ST19.Text);
                    }
                    TimeSpan restar_19 = fecha2_19 - fecha1_19;
                    TimeSpan restax_19 = fecha4_19 - fecha3_19;
                    TimeSpan suma_19 = restar_19 + restax_19;

                    if (suma_19.ToString() != "0:00")
                    {
                        if (Extra19.Checked == true)
                        {
                            if (suma_19 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_19 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day19.Text + ": " + (suma_19 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_19);
                                SumaHorasExtraDia.Add(Day19.Text + ": " + suma_19.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_20
                    TimeSpan fecha1_20 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_20 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_20 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_20 = TimeSpan.Parse("0:00");
                    if (EM20.Text != "")
                    {
                        fecha1_20 = TimeSpan.Parse(EM20.Text);
                    }
                    if (SM20.Text != "")
                    {
                        fecha2_20 = TimeSpan.Parse(SM20.Text);
                    }
                    if (ET20.Text != "")
                    {
                        fecha3_20 = TimeSpan.Parse(ET20.Text);
                    }
                    if (ST20.Text != "")
                    {
                        fecha4_20 = TimeSpan.Parse(ST20.Text);
                    }
                    TimeSpan restar_20 = fecha2_20 - fecha1_20;
                    TimeSpan restax_20 = fecha4_20 - fecha3_20;
                    TimeSpan suma_20 = restar_20 + restax_20;

                    if (suma_20.ToString() != "0:00")
                    {
                        if (Extra20.Checked == true)
                        {
                            if (suma_20 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_20 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day20.Text + ": " + (suma_20 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_20);
                                SumaHorasExtraDia.Add(Day20.Text + ": " + suma_20.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_21
                    TimeSpan fecha1_21 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_21 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_21 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_21 = TimeSpan.Parse("0:00");
                    if (EM21.Text != "")
                    {
                        fecha1_21 = TimeSpan.Parse(EM21.Text);
                    }
                    if (SM21.Text != "")
                    {
                        fecha2_21 = TimeSpan.Parse(SM21.Text);
                    }
                    if (ET21.Text != "")
                    {
                        fecha3_21 = TimeSpan.Parse(ET21.Text);
                    }
                    if (ST21.Text != "")
                    {
                        fecha4_21 = TimeSpan.Parse(ST21.Text);
                    }
                    TimeSpan restar_21 = fecha2_21 - fecha1_21;
                    TimeSpan restax_21 = fecha4_21 - fecha3_21;
                    TimeSpan suma_21 = restar_21 + restax_21;

                    if (suma_21.ToString() != "0:00")
                    {
                        if (Extra21.Checked == true)
                        {
                            if (suma_21 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_21 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day21.Text + ": " + (suma_21 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_21);
                                SumaHorasExtraDia.Add(Day21.Text + ": " + suma_21.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_22
                    TimeSpan fecha1_22 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_22 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_22 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_22 = TimeSpan.Parse("0:00");
                    if (EM22.Text != "")
                    {
                        fecha1_22 = TimeSpan.Parse(EM22.Text);
                    }
                    if (SM22.Text != "")
                    {
                        fecha2_22 = TimeSpan.Parse(SM22.Text);
                    }
                    if (ET22.Text != "")
                    {
                        fecha3_22 = TimeSpan.Parse(ET22.Text);
                    }
                    if (ST22.Text != "")
                    {
                        fecha4_22 = TimeSpan.Parse(ST22.Text);
                    }
                    TimeSpan restar_22 = fecha2_22 - fecha1_22;
                    TimeSpan restax_22 = fecha4_22 - fecha3_22;
                    TimeSpan suma_22 = restar_22 + restax_22;

                    if (suma_22.ToString() != "0:00")
                    {
                        if (Extra22.Checked == true)
                        {
                            if (suma_22 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_22 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day22.Text + ": " + (suma_22 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_22);
                                SumaHorasExtraDia.Add(Day22.Text + ": " + suma_22.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_23
                    TimeSpan fecha1_23 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_23 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_23 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_23 = TimeSpan.Parse("0:00");
                    if (EM23.Text != "")
                    {
                        fecha1_23 = TimeSpan.Parse(EM23.Text);
                    }
                    if (SM23.Text != "")
                    {
                        fecha2_23 = TimeSpan.Parse(SM23.Text);
                    }
                    if (ET23.Text != "")
                    {
                        fecha3_23 = TimeSpan.Parse(ET23.Text);
                    }
                    if (ST23.Text != "")
                    {
                        fecha4_23 = TimeSpan.Parse(ST23.Text);
                    }
                    TimeSpan restar_23 = fecha2_23 - fecha1_23;
                    TimeSpan restax_23 = fecha4_23 - fecha3_23;
                    TimeSpan suma_23 = restar_23 + restax_23;

                    if (suma_23.ToString() != "0:00")
                    {
                        if (Extra23.Checked == true)
                        {
                            if (suma_23 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_23 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day23.Text + ": " + (suma_23 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_23);
                                SumaHorasExtraDia.Add(Day23.Text + ": " + suma_23.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_24
                    TimeSpan fecha1_24 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_24 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_24 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_24 = TimeSpan.Parse("0:00");
                    if (EM24.Text != "")
                    {
                        fecha1_24 = TimeSpan.Parse(EM24.Text);
                    }
                    if (SM24.Text != "")
                    {
                        fecha2_24 = TimeSpan.Parse(SM24.Text);
                    }
                    if (ET24.Text != "")
                    {
                        fecha3_24 = TimeSpan.Parse(ET24.Text);
                    }
                    if (ST24.Text != "")
                    {
                        fecha4_24 = TimeSpan.Parse(ST24.Text);
                    }
                    TimeSpan restar_24 = fecha2_24 - fecha1_24;
                    TimeSpan restax_24 = fecha4_24 - fecha3_24;
                    TimeSpan suma_24 = restar_24 + restax_24;

                    if (suma_24.ToString() != "0:00")
                    {
                        if (Extra24.Checked == true)
                        {
                            if (suma_24 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_24 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day24.Text + ": " + (suma_24 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_24);
                                SumaHorasExtraDia.Add(Day24.Text + ": " + suma_24.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_25
                    TimeSpan fecha1_25 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_25 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_25 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_25 = TimeSpan.Parse("0:00");
                    if (EM25.Text != "")
                    {
                        fecha1_25 = TimeSpan.Parse(EM25.Text);
                    }
                    if (SM25.Text != "")
                    {
                        fecha2_25 = TimeSpan.Parse(SM25.Text);
                    }
                    if (ET25.Text != "")
                    {
                        fecha3_25 = TimeSpan.Parse(ET25.Text);
                    }
                    if (ST25.Text != "")
                    {
                        fecha4_25 = TimeSpan.Parse(ST25.Text);
                    }
                    TimeSpan restar_25 = fecha2_25 - fecha1_25;
                    TimeSpan restax_25 = fecha4_25 - fecha3_25;
                    TimeSpan suma_25 = restar_25 + restax_25;

                    if (suma_25.ToString() != "0:00")
                    {
                        if (Extra25.Checked == true)
                        {
                            if (suma_25 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_25 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day25.Text + ": " + (suma_25 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_25);
                                SumaHorasExtraDia.Add(Day25.Text + ": " + suma_25.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_26
                    TimeSpan fecha1_26 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_26 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_26 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_26 = TimeSpan.Parse("0:00");
                    if (EM26.Text != "")
                    {
                        fecha1_26 = TimeSpan.Parse(EM26.Text);
                    }
                    if (SM26.Text != "")
                    {
                        fecha2_26 = TimeSpan.Parse(SM26.Text);
                    }
                    if (ET26.Text != "")
                    {
                        fecha3_26 = TimeSpan.Parse(ET26.Text);
                    }
                    if (ST26.Text != "")
                    {
                        fecha4_26 = TimeSpan.Parse(ST26.Text);
                    }
                    TimeSpan restar_26 = fecha2_26 - fecha1_26;
                    TimeSpan restax_26 = fecha4_26 - fecha3_26;
                    TimeSpan suma_26 = restar_26 + restax_26;

                    if (suma_26.ToString() != "0:00")
                    {
                        if (Extra26.Checked == true)
                        {
                            if (suma_26 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_26 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day26.Text + ": " + (suma_26 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_26);
                                SumaHorasExtraDia.Add(Day26.Text + ": " + suma_26.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_27
                    TimeSpan fecha1_27 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_27 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_27 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_27 = TimeSpan.Parse("0:00");
                    if (EM27.Text != "")
                    {
                        fecha1_27 = TimeSpan.Parse(EM27.Text);
                    }
                    if (SM27.Text != "")
                    {
                        fecha2_27 = TimeSpan.Parse(SM27.Text);
                    }
                    if (ET27.Text != "")
                    {
                        fecha3_27 = TimeSpan.Parse(ET27.Text);
                    }
                    if (ST27.Text != "")
                    {
                        fecha4_27 = TimeSpan.Parse(ST27.Text);
                    }
                    TimeSpan restar_27 = fecha2_27 - fecha1_27;
                    TimeSpan restax_27 = fecha4_27 - fecha3_27;
                    TimeSpan suma_27 = restar_27 + restax_27;

                    if (suma_27.ToString() != "0:00")
                    {
                        if (Extra27.Checked == true)
                        {
                            if (suma_27 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_27 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day27.Text + ": " + (suma_27 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_27);
                                SumaHorasExtraDia.Add(Day27.Text + ": " + suma_27.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_28
                    TimeSpan fecha1_28 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_28 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_28 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_28 = TimeSpan.Parse("0:00");
                    if (EM28.Text != "")
                    {
                        fecha1_28 = TimeSpan.Parse(EM28.Text);
                    }
                    if (SM28.Text != "")
                    {
                        fecha2_28 = TimeSpan.Parse(SM28.Text);
                    }
                    if (ET28.Text != "")
                    {
                        fecha3_28 = TimeSpan.Parse(ET28.Text);
                    }
                    if (ST28.Text != "")
                    {
                        fecha4_28 = TimeSpan.Parse(ST28.Text);
                    }
                    TimeSpan restar_28 = fecha2_28 - fecha1_28;
                    TimeSpan restax_28 = fecha4_28 - fecha3_28;
                    TimeSpan suma_28 = restar_28 + restax_28;

                    if (suma_28.ToString() != "0:00")
                    {
                        if (Extra28.Checked == true)
                        {
                            if (suma_28 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_28 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day28.Text + ": " + (suma_28 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_28);
                                SumaHorasExtraDia.Add(Day28.Text + ": " + suma_28.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_29
                    TimeSpan fecha1_29 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_29 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_29 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_29 = TimeSpan.Parse("0:00");
                    if (EM29.Text != "")
                    {
                        fecha1_29 = TimeSpan.Parse(EM29.Text);
                    }
                    if (SM29.Text != "")
                    {
                        fecha2_29 = TimeSpan.Parse(SM29.Text);
                    }
                    if (ET29.Text != "")
                    {
                        fecha3_29 = TimeSpan.Parse(ET29.Text);
                    }
                    if (ST29.Text != "")
                    {
                        fecha4_29 = TimeSpan.Parse(ST29.Text);
                    }
                    TimeSpan restar_29 = fecha2_29 - fecha1_29;
                    TimeSpan restax_29 = fecha4_29 - fecha3_29;
                    TimeSpan suma_29 = restar_29 + restax_29;

                    if (suma_29.ToString() != "0:00")
                    {
                        if (Extra29.Checked == true)
                        {
                            if (suma_29 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_29 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day29.Text + ": " + (suma_29 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_29);
                                SumaHorasExtraDia.Add(Day29.Text + ": " + suma_29.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_30
                    TimeSpan fecha1_30 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_30 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_30 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_30 = TimeSpan.Parse("0:00");
                    if (EM30.Text != "")
                    {
                        fecha1_30 = TimeSpan.Parse(EM30.Text);
                    }
                    if (SM30.Text != "")
                    {
                        fecha2_30 = TimeSpan.Parse(SM30.Text);
                    }
                    if (ET30.Text != "")
                    {
                        fecha3_30 = TimeSpan.Parse(ET30.Text);
                    }
                    if (ST30.Text != "")
                    {
                        fecha4_30 = TimeSpan.Parse(ST30.Text);
                    }
                    TimeSpan restar_30 = fecha2_30 - fecha1_30;
                    TimeSpan restax_30 = fecha4_30 - fecha3_30;
                    TimeSpan suma_30 = restar_30 + restax_30;

                    if (suma_30.ToString() != "0:00")
                    {
                        if (Extra30.Checked == true)
                        {
                            if (suma_30 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_30 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day30.Text + ": " + (suma_30 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_30);
                                SumaHorasExtraDia.Add(Day30.Text + ": " + suma_30.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //_31
                    TimeSpan fecha1_31 = TimeSpan.Parse("0:00");
                    TimeSpan fecha2_31 = TimeSpan.Parse("0:00");
                    TimeSpan fecha3_31 = TimeSpan.Parse("0:00");
                    TimeSpan fecha4_31 = TimeSpan.Parse("0:00");
                    if (EM31.Text != "")
                    {
                        fecha1_31 = TimeSpan.Parse(EM31.Text);
                    }
                    if (SM31.Text != "")
                    {
                        fecha2_31 = TimeSpan.Parse(SM31.Text);
                    }
                    if (ET31.Text != "")
                    {
                        fecha3_31 = TimeSpan.Parse(ET31.Text);
                    }
                    if (ST31.Text != "")
                    {
                        fecha4_31 = TimeSpan.Parse(ST31.Text);
                    }
                    TimeSpan restar_31 = fecha2_31 - fecha1_31;
                    TimeSpan restax_31 = fecha4_31 - fecha3_31;
                    TimeSpan suma_31 = restar_31 + restax_31;

                    if (suma_31.ToString() != "0:00")
                    {
                        if (Extra31.Checked == true)
                        {
                            if (suma_31 > TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada))
                            {
                                SumaHorasExtra2.Add(suma_31 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada));
                                SumaHorasExtraDia.Add(Day31.Text + ": " + (suma_31 - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada)).ToString());
                            }
                            else
                            {
                                SumaHorasExtra2.Add(suma_31);
                                SumaHorasExtraDia.Add(Day31.Text + ": " + suma_31.ToString(@"h\:mm"));
                            }
                        }
                    }

                    //TOTAL HORAS EXTRAS
                    TimeSpan tiempu = new TimeSpan();
                    for (int i = 0; i < SumaHorasExtra2.Count; i++)
                    {
                        tiempu += SumaHorasExtra2[i];
                    }

                    string FinishExtra = "";
                    var Minso = string.Format("{0:D2}", tiempu.Minutes);
                    var Hourso = string.Format("{0:D2}", tiempu.Hours);
                    if (Hourso != "00")
                    {
                        string HourChange = Hourso.Replace("0", "") + ":00";
                        TimeSpan Finisho = TimeSpan.Parse(HourChange);// - TimeSpan.Parse(Properties.Settings.Default.HorasdeJornada);
                        FinishExtra = Finisho.ToString().Replace(":", "").Replace("0", "").Replace("-", "") + ":" + Minso;
                    }
                    else if (Minso != "00")
                    {
                        FinishExtra = "00:" + Minso;
                    }

                    DateTime now = DateTime.Now;
                    string nox = now.ToString("dd");
                    DialogResult dialogResult = MessageBox.Show("Somos día " + nox + ". \rTienes " + SumaHorasExtra2.Count.ToString() + " días con " + FinishExtra.ToString() + " horas extra en total. \r\r¿Quieres envíar un correo informando de estas?", "Tesalia Redes", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    if (dialogResult == DialogResult.Yes)
                    {
                        string Texto = "";
                        for (int i = 0; i < SumaHorasExtra2.Count; i++)
                        {
                            Texto += SumaHorasExtraDia[i] + "h\r";
                        }

                        Properties.Settings.Default.CorreoType = "1";
                        Properties.Settings.Default.Save();
                        Form3 fm3 = new Form3("Horas extras de " + Properties.Settings.Default.Name + " de " + MesPreSelected, Texto);
                        fm3.Show();
                    }
                }
                else
                {
                    MessageBox.Show("Es necesario indicar las horas de la jornada.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("No se ha introducido el nombre completo en la información del trabajador.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        int stop = 0;
        private void timer3_Tick(object sender, EventArgs e)
        {
            if (stop == 0)
            {
                stop = 1;
                DateTime now = DateTime.Now;
                string nox = now.ToString("dd");
                if (nox == "22")
                {
                    MessageBox.Show("Somos día " + nox + ". \rRecuerda envíar tus horas extras antes de cobrar para que se te paguen este mes.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (nox == "23")
                {
                    MessageBox.Show("Somos día " + nox + ". \rRecuerda envíar tus horas extras antes de cobrar para que se te paguen este mes.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (nox == "24")
                {
                    MessageBox.Show("Somos día " + nox + ". \rRecuerda envíar tus horas extras antes de cobrar para que se te paguen este mes.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (nox == "25")
                {
                    MessageBox.Show("Somos día " + nox + ". \rRecuerda envíar tus horas extras antes de cobrar para que se te paguen este mes.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else if (nox == "26")
                {
                    MessageBox.Show("Somos día " + nox + ". \rRecuerda envíar tus horas extras antes de cobrar para que se te paguen este mes.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                timer3.Stop();
            }
        }

        private void Version_Click(object sender, EventArgs e)
        {
            WebClient mywebClient = new WebClient();
            mywebClient.DownloadFile("https://www.googleapis.com/drive/v3/files/1kqjOgMhQaA93QC6vcgtGF2hYphFniwCp?supportsAllDrives=true&supportsTeamDrives=true&key=AIzaSyC3etOnvvUrusPVzE03ZFOQJrDxMkHWzRU&alt=media", Application.StartupPath + @"\version.txt");
            WebClient mywebClient2 = new WebClient();
            mywebClient2.DownloadFile("https://www.googleapis.com/drive/v3/files/1NfGbih9LFxA0gJfA5uA6XjD8X36YZHtx?supportsAllDrives=true&supportsTeamDrives=true&key=AIzaSyC3etOnvvUrusPVzE03ZFOQJrDxMkHWzRU&alt=media", Application.StartupPath + @"\Updater.new");

            if (File.Exists(Application.StartupPath + @"\Updater.new"))
            {
                if (File.Exists(Application.StartupPath + @"\Updater.exe"))
                {
                    File.Delete(Application.StartupPath + @"\Updater.exe");
                    File.Move(Application.StartupPath + @"\Updater.new", Application.StartupPath + @"\Updater.exe");
                }
            }

            StreamReader sr = new StreamReader(Application.StartupPath + @"\version.txt");
            string line = sr.ReadLine();
            sr.Close();

            if (File.Exists(Application.StartupPath + @"\version.txt"))
            {
                File.Delete(Application.StartupPath + @"\version.txt");
            }

            if (line != Application.ProductVersion.ToString())
            {
                MessageBox.Show("Hay una nueva actualización del programa.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                MessageBox.Show("Se descargará y actualizará a continuación.", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Process.Start(Application.StartupPath + @"\Updater.exe");
                Close();
            }
        }
    }
}