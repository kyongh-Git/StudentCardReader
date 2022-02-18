using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StudentCardReader
{
    public partial class Form2 : Form
    {
        public String fileDir = "";
        public Form2()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            label3.Text = openFileDialog1.FileName;
            button2.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (label3.Text == "")
            {
                label2.Visible = true;
            }
            else 
            {
                fileDir = label3.Text;
                this.Close();
            }
        }
    }
}
