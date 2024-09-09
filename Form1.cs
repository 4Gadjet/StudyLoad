using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;

namespace Program
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Нагрузка form2 = new Нагрузка();
            form2.ShowDialog();
        }

        private void toolStripLabel1_Click(object sender, EventArgs e)
        {
            Предмет form3 = new Предмет();
            form3.ShowDialog();
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Преподаватель form4 = new Преподаватель();
            form4.ShowDialog();
        }

        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            Отчет form5 = new Отчет();
            form5.ShowDialog();
        }
    }
}
