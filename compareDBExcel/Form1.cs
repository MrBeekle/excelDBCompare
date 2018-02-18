using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LinqToExcel;


namespace compareDBExcel
{
    public partial class Form1 : Form
    {
        string excelFilePath;


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                listBox2.Items.Add(openFileDialog1.FileName);

            excelFilePath = openFileDialog1.FileName;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

            ExcelDataRetriever dataFromExcel = new ExcelDataRetriever();
            dataFromExcel.GetExcelData(excelFilePath, ListBox1);

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
