using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tesalia_Redes_App
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Waiter_Tick(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.Installing == "0")
            {
                Close();
            }
        }
    }
}
