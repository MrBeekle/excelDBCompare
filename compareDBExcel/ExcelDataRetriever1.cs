using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using System.Windows.Forms;

namespace compareDBExcel
{
    class ExcelDataRetriever : Form1
    {

        List<string> compareList = new List<string>();

        public void GetExcelData(string excelFilePath, ListBox ListBox1)
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
                ListBox1.Items.Add(a[2]);
            }

        }


    }
}
