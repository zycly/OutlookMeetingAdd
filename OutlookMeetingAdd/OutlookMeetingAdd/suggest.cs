using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookMeetingAdd
{
    public partial class suggest : Form
    {
        public int translation;
        public suggest(string parent_information)
        {
            InitializeComponent();
            this.textBox1.Text = parent_information;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            translation = 1;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            translation = 0;
        }
    }
}
