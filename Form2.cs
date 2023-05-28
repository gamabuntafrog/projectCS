using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using System.Xml.Linq;

namespace WindowsFormsApp26
{
    public partial class Form2 : Form
    {
        public string Task { get; set; }
        public string Priority { get; set; }
        public string Term { get; set; }
        public string State { get; set; }
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Task = textBox1.Text;
            Priority = textBox2.Text;
            Term = textBox3.Text;
            State = textBox4.Text;
            //Якщо користувач нажимає окей то створюється новий рядок зі значеннями які він ввів і ця форма закривається.
            DialogResult = DialogResult.OK;
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
