using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel; 

namespace OpenAndEditExcel
{
    public partial class Mainfrm : Form
    {
        public Mainfrm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MainFunction();
        }
        /// <summary>
        /// 
        /// </summary>
        private void MainFunction()
        {
            Excel.Application excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                EditCells(OpenFile(excelApp));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                }
            }
        }

        private Excel.Worksheet OpenFile(Excel.Application excelApp)
        {
            MainOpenfile.FileName = "";
            MainOpenfile.ShowDialog();
            string fileName;
            if (MainOpenfile.FileName == "")
            {
                throw new Exception("a filename is required!");
            }

            fileName = MainOpenfile.FileName;

            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(fileName, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;

            if (excelSheets.Count == 0)
            {
                throw new Exception("At least one sheet is required !");
            }
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelSheets[1];

            return excelWorksheet;
        }

        private void EditCells(Excel.Worksheet excelWorksheet)
        {
            var col = excelWorksheet.UsedRange.Columns["H:H"];
            Random rnd = new Random();
            string tempsubstr = String.Empty;
            string substr = String.Empty;
            foreach (Excel.Range item in col.Cells)
            {
                tempsubstr = item.Value.ToString();
                if (item.Value != "CarPlate")
                {
                    switch (tempsubstr.Length)
                    {
                        case 6 :
                            substr=item.Value.Substring(0, 2);
                            break;
                        case 7 :
                            substr=item.Value.Substring(0, 3);
                            break;
                        default :
                            substr = item.Value.Substring(0, 2);
                            break;
                    }

                    int newNumber = rnd.Next(1, 9999);
                    item.Value = substr + newNumber.ToString().PadLeft(4, '0');
                }
            } 
        }
    }
}
