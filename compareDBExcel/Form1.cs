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
        List<string> compareList = new List<string>();


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
            //Pobieranie danych z pliku excel
            string sheetName = "1.2 Cash";

            var excelFile = new ExcelQueryFactory(excelFilePath);
            var getThisData = from a in excelFile.Worksheet(sheetName) where (a["Is used"].Cast<int>() > 0 && (a["R"].Cast<int>() < 0 || a["R"].Cast<int>() > 0)) select a;
            /*
            *   a["Is used"].Cast<int>() > 0 pobiera wartosci z kolumny is used wieksze od zera czyli tam gdzie jest wartosc
            * 
            *   a["R"].Cast<int>() < 0 pobiera warosc z kolumny R tam gdzie jest znaczek funta(checkox nie zaznaczony) -> sprawdzic to 
            */

            foreach (var a in getThisData)
            {
                //dodanie do listy 'test' przefiltrowanych wedlug warunku z lini 19 danych z 3 kolumny arkusza
                compareList.Add(a[2]);
                listBox1.Items.Add(a[2]);
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
