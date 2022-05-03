using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MessageTrackingLogs
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Haha");
        }

        private void Form2_Shown(object sender, EventArgs e)
        {
            Extension.WaitNSeconds(5);
            button1.PerformClick();
        }
    }
}
