namespace Tesalia_Redes_App
{
    partial class Form3
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form3));
            this.PanelUP = new System.Windows.Forms.Panel();
            this.Titulo = new System.Windows.Forms.Label();
            this.Close = new FontAwesome.Sharp.IconButton();
            this.MaxMin = new FontAwesome.Sharp.IconButton();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.Mensaje = new System.Windows.Forms.RichTextBox();
            this.TXT3 = new System.Windows.Forms.Label();
            this.SendMail = new FontAwesome.Sharp.IconButton();
            this.SaveMail = new FontAwesome.Sharp.IconButton();
            this.Asunto = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox3 = new System.Windows.Forms.PictureBox();
            this.pictureBox4 = new System.Windows.Forms.PictureBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.PanelUP.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // PanelUP
            // 
            this.PanelUP.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.PanelUP.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(78)))), ((int)(((byte)(60)))), ((int)(((byte)(86)))));
            this.PanelUP.Controls.Add(this.Titulo);
            this.PanelUP.Controls.Add(this.Close);
            this.PanelUP.Controls.Add(this.MaxMin);
            this.PanelUP.Location = new System.Drawing.Point(0, 0);
            this.PanelUP.Name = "PanelUP";
            this.PanelUP.Size = new System.Drawing.Size(559, 33);
            this.PanelUP.TabIndex = 7;
            this.PanelUP.MouseDown += new System.Windows.Forms.MouseEventHandler(this.PanelUP_MouseDown);
            // 
            // Titulo
            // 
            this.Titulo.BackColor = System.Drawing.Color.Transparent;
            this.Titulo.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Bold);
            this.Titulo.ForeColor = System.Drawing.Color.White;
            this.Titulo.Location = new System.Drawing.Point(3, 2);
            this.Titulo.Name = "Titulo";
            this.Titulo.Size = new System.Drawing.Size(317, 25);
            this.Titulo.TabIndex = 53;
            this.Titulo.Text = "Tesalia Redes - App";
            this.Titulo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.Titulo.MouseDown += new System.Windows.Forms.MouseEventHandler(this.Titulo_MouseDown);
            // 
            // Close
            // 
            this.Close.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Close.Cursor = System.Windows.Forms.Cursors.Hand;
            this.Close.FlatAppearance.BorderSize = 0;
            this.Close.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Close.Font = new System.Drawing.Font("Bahnschrift", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Close.ForeColor = System.Drawing.Color.White;
            this.Close.IconChar = FontAwesome.Sharp.IconChar.Times;
            this.Close.IconColor = System.Drawing.Color.White;
            this.Close.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.Close.IconSize = 22;
            this.Close.Location = new System.Drawing.Point(528, 2);
            this.Close.Name = "Close";
            this.Close.Size = new System.Drawing.Size(30, 30);
            this.Close.TabIndex = 10;
            this.Close.UseVisualStyleBackColor = true;
            this.Close.Click += new System.EventHandler(this.Close_Click);
            // 
            // MaxMin
            // 
            this.MaxMin.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.MaxMin.Cursor = System.Windows.Forms.Cursors.Hand;
            this.MaxMin.FlatAppearance.BorderSize = 0;
            this.MaxMin.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.MaxMin.Font = new System.Drawing.Font("Bahnschrift", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MaxMin.ForeColor = System.Drawing.Color.White;
            this.MaxMin.IconChar = FontAwesome.Sharp.IconChar.WindowMinimize;
            this.MaxMin.IconColor = System.Drawing.Color.White;
            this.MaxMin.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.MaxMin.IconSize = 20;
            this.MaxMin.Location = new System.Drawing.Point(497, 2);
            this.MaxMin.Name = "MaxMin";
            this.MaxMin.Size = new System.Drawing.Size(30, 30);
            this.MaxMin.TabIndex = 8;
            this.MaxMin.UseVisualStyleBackColor = true;
            this.MaxMin.Click += new System.EventHandler(this.MaxMin_Click);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Bold);
            this.label1.ForeColor = System.Drawing.Color.White;
            this.label1.Location = new System.Drawing.Point(52, 90);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(81, 27);
            this.label1.TabIndex = 208;
            this.label1.Text = "Mensaje:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(98)))), ((int)(((byte)(80)))), ((int)(((byte)(106)))));
            this.panel1.Controls.Add(this.Mensaje);
            this.panel1.Location = new System.Drawing.Point(138, 93);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(369, 138);
            this.panel1.TabIndex = 206;
            // 
            // Mensaje
            // 
            this.Mensaje.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Mensaje.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(36)))), ((int)(((byte)(60)))));
            this.Mensaje.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.Mensaje.Font = new System.Drawing.Font("Bahnschrift", 14F);
            this.Mensaje.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(200)))), ((int)(((byte)(230)))));
            this.Mensaje.Location = new System.Drawing.Point(1, 1);
            this.Mensaje.Margin = new System.Windows.Forms.Padding(2);
            this.Mensaje.Name = "Mensaje";
            this.Mensaje.Size = new System.Drawing.Size(367, 136);
            this.Mensaje.TabIndex = 155;
            this.Mensaje.Text = "";
            // 
            // TXT3
            // 
            this.TXT3.Font = new System.Drawing.Font("Bahnschrift", 12F, System.Drawing.FontStyle.Bold);
            this.TXT3.ForeColor = System.Drawing.Color.White;
            this.TXT3.Location = new System.Drawing.Point(51, 58);
            this.TXT3.Name = "TXT3";
            this.TXT3.Size = new System.Drawing.Size(81, 27);
            this.TXT3.TabIndex = 205;
            this.TXT3.Text = "Asunto:";
            this.TXT3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // SendMail
            // 
            this.SendMail.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(69)))), ((int)(((byte)(54)))), ((int)(((byte)(75)))));
            this.SendMail.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SendMail.FlatAppearance.BorderSize = 0;
            this.SendMail.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SendMail.Font = new System.Drawing.Font("Bahnschrift", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SendMail.ForeColor = System.Drawing.Color.White;
            this.SendMail.IconChar = FontAwesome.Sharp.IconChar.None;
            this.SendMail.IconColor = System.Drawing.Color.White;
            this.SendMail.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.SendMail.IconSize = 32;
            this.SendMail.Location = new System.Drawing.Point(283, 250);
            this.SendMail.Name = "SendMail";
            this.SendMail.Size = new System.Drawing.Size(224, 40);
            this.SendMail.TabIndex = 210;
            this.SendMail.Text = "Enviar";
            this.SendMail.UseVisualStyleBackColor = false;
            this.SendMail.Click += new System.EventHandler(this.SendMail_Click);
            // 
            // SaveMail
            // 
            this.SaveMail.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(69)))), ((int)(((byte)(54)))), ((int)(((byte)(75)))));
            this.SaveMail.Cursor = System.Windows.Forms.Cursors.Hand;
            this.SaveMail.FlatAppearance.BorderSize = 0;
            this.SaveMail.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.SaveMail.Font = new System.Drawing.Font("Bahnschrift", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SaveMail.ForeColor = System.Drawing.Color.White;
            this.SaveMail.IconChar = FontAwesome.Sharp.IconChar.None;
            this.SaveMail.IconColor = System.Drawing.Color.White;
            this.SaveMail.IconFont = FontAwesome.Sharp.IconFont.Auto;
            this.SaveMail.IconSize = 32;
            this.SaveMail.Location = new System.Drawing.Point(53, 250);
            this.SaveMail.Name = "SaveMail";
            this.SaveMail.Size = new System.Drawing.Size(224, 40);
            this.SaveMail.TabIndex = 209;
            this.SaveMail.Text = "Guardar";
            this.SaveMail.UseVisualStyleBackColor = false;
            this.SaveMail.Click += new System.EventHandler(this.SaveMail_Click);
            // 
            // Asunto
            // 
            this.Asunto.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(36)))), ((int)(((byte)(60)))));
            this.Asunto.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Asunto.Font = new System.Drawing.Font("Bahnschrift", 14F);
            this.Asunto.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(220)))), ((int)(((byte)(200)))), ((int)(((byte)(230)))));
            this.Asunto.Location = new System.Drawing.Point(0, 0);
            this.Asunto.Margin = new System.Windows.Forms.Padding(4);
            this.Asunto.Name = "Asunto";
            this.Asunto.Size = new System.Drawing.Size(369, 30);
            this.Asunto.TabIndex = 202;
            this.Asunto.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Location = new System.Drawing.Point(0, 0);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(369, 1);
            this.pictureBox1.TabIndex = 203;
            this.pictureBox1.TabStop = false;
            // 
            // pictureBox2
            // 
            this.pictureBox2.Location = new System.Drawing.Point(0, 29);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(369, 1);
            this.pictureBox2.TabIndex = 204;
            this.pictureBox2.TabStop = false;
            // 
            // pictureBox3
            // 
            this.pictureBox3.Location = new System.Drawing.Point(0, -3);
            this.pictureBox3.Name = "pictureBox3";
            this.pictureBox3.Size = new System.Drawing.Size(1, 40);
            this.pictureBox3.TabIndex = 205;
            this.pictureBox3.TabStop = false;
            // 
            // pictureBox4
            // 
            this.pictureBox4.Location = new System.Drawing.Point(368, -3);
            this.pictureBox4.Name = "pictureBox4";
            this.pictureBox4.Size = new System.Drawing.Size(1, 40);
            this.pictureBox4.TabIndex = 206;
            this.pictureBox4.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(98)))), ((int)(((byte)(80)))), ((int)(((byte)(106)))));
            this.panel2.Controls.Add(this.pictureBox4);
            this.panel2.Controls.Add(this.pictureBox3);
            this.panel2.Controls.Add(this.pictureBox2);
            this.panel2.Controls.Add(this.pictureBox1);
            this.panel2.Controls.Add(this.Asunto);
            this.panel2.Location = new System.Drawing.Point(138, 58);
            this.panel2.Margin = new System.Windows.Forms.Padding(2);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(369, 30);
            this.panel2.TabIndex = 207;
            // 
            // Form3
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(52)))), ((int)(((byte)(36)))), ((int)(((byte)(60)))));
            this.ClientSize = new System.Drawing.Size(558, 316);
            this.ControlBox = false;
            this.Controls.Add(this.SendMail);
            this.Controls.Add(this.SaveMail);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.TXT3);
            this.Controls.Add(this.PanelUP);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form3";
            this.Opacity = 0.99D;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Tesalia Redes";
            this.PanelUP.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox4)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel PanelUP;
        private System.Windows.Forms.Label Titulo;
        private new FontAwesome.Sharp.IconButton Close;
        private FontAwesome.Sharp.IconButton MaxMin;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RichTextBox Mensaje;
        private System.Windows.Forms.Label TXT3;
        private FontAwesome.Sharp.IconButton SendMail;
        private FontAwesome.Sharp.IconButton SaveMail;
        private System.Windows.Forms.TextBox Asunto;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox3;
        private System.Windows.Forms.PictureBox pictureBox4;
        private System.Windows.Forms.Panel panel2;
    }
}