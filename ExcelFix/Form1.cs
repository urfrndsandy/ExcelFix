using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using excelap = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ExcelFix
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void readExcel()
        {
           excelap. Application excelApp = new excelap. Application();
            string filepath = @"\\SERVER1\Scan\Copy of (50 employees) Master of Form 16 Only Part  B for AY 2019-20.xlsm";
            Microsoft.Office.Interop.Excel.Workbook workBook = excelApp.Workbooks.Open(filepath);
           // var valx;
            foreach (Worksheet sheet in workBook.Worksheets)
            {
                //  string val = sheet.Rows[0][1];
                if (sheet.Name != "Deductor's Sheet" && sheet.Name!= "Salary Details")
                {
                  
                       
                      // sheet.Rows.Cells[11][13] = "2019-2020";

                        for (int i = 85; i <= 90; i++)

                        {
                        try
                        {
                            var valx = sheet.Rows.Cells[1][i].Value.ToString();


                            if (valx==("(a) Standard Deduction"))
                            {
                                sheet.Rows.Cells[1][i] = "     (a) Standard Deduction";
                            }
                           // MessageBox.Show(sheet.Name + ":" + valx);
                        }
                        catch { }

                        }
                    

                  
                }
             
            }
            workBook.Save();
          excelApp.Quit();
           // update(filepath);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            readExcel();
            MessageBox.Show("done");

        }


































        public void update(string filex)
        {
            System.IO.File.Copy(filex, "E:\\backups\\new.textl");
        }


    }
}
