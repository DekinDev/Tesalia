using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tesalia_Redes_App
{
    public partial class Form3 : Form
    {
        public Form3(string AsuntoC, string MensajeC)
        {
            InitializeComponent();

            if (Properties.Settings.Default.CorreoType == "0")
            {
                if (Properties.Settings.Default.Asunto == "")
                {
                    Asunto.Text = AsuntoC;
                }
                else
                {
                    Asunto.Text = Properties.Settings.Default.Asunto;
                }
                if (Properties.Settings.Default.Mensaje == "")
                {
                    Mensaje.Text = MensajeC;
                }
                else
                {
                    Mensaje.Text = Properties.Settings.Default.Mensaje;
                }
            }
            else if(Properties.Settings.Default.CorreoType == "1")
            {
                Asunto.Text = AsuntoC;
                Mensaje.Text = MensajeC;
            }
        }

        private void SaveMail_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Asunto = Asunto.Text;
            Properties.Settings.Default.Mensaje = Mensaje.Text;
            Properties.Settings.Default.Save();
            MessageBox.Show("Datos guardados!", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SendMail_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.MailMail != "")
            {
                if (Properties.Settings.Default.MailPass != "")
                {
                    if (Asunto.Text != "")
                    {
                        if (Mensaje.Text != "")
                        {
                            MailMessage correo = new MailMessage();
                            SmtpClient smtpclient = new SmtpClient("217.116.0.228");
                            smtpclient.Port = 25;
                            smtpclient.UseDefaultCredentials = false;
                            smtpclient.Credentials = new System.Net.NetworkCredential(Properties.Settings.Default.MailMail, Properties.Settings.Default.MailPass);
                            correo.From = new MailAddress(Properties.Settings.Default.MailMail);
                            correo.To.Add("dolores.pintor@tesaliaredes.com");
                            if(Properties.Settings.Default.EnviarMailCC == "1")
                            {
                                correo.CC.Add(Properties.Settings.Default.MailMail);
                            }
                            
                            correo.Subject = (Asunto.Text);
                            correo.Priority = MailPriority.High;
                            correo.IsBodyHtml = false;
                            correo.Body = Mensaje.Text + "\r\r(Enviado desde Tesalia Redes App)";

                            if (Properties.Settings.Default.CorreoType == "0")
                            {
                                Attachment data = new Attachment(Properties.Settings.Default.UltimoArchivoGenerado, MediaTypeNames.Application.Octet);
                                correo.Attachments.Add(data);
                            }
                            
                            try
                            {
                                smtpclient.Send(correo);
                                if (Properties.Settings.Default.CorreoType == "0")
                                {
                                    MessageBox.Show("Registro de Jornada enviado!", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                if (Properties.Settings.Default.CorreoType == "1")
                                {
                                    MessageBox.Show("Informe de horas extra enviado!", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                
                            }
                            catch (SmtpException ex)
                            {
                                MessageBox.Show("ERROR: " + ex.Message, "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Contraseña del Correo de Tesalia no se ha informado en Mi Cuenta", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Correo de Tesalia no informado en Mi cuenta", "Tesalia Redes", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }






        private void Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void MaxMin_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void PanelUP_MouseDown(object sender, MouseEventArgs e)
        {

        }

        private void Titulo_MouseDown(object sender, MouseEventArgs e)
        {

        }
    }
}
