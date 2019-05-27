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

            foreach (Worksheet sheet in workBook.Worksheets)
            {
                //  string val = sheet.Rows[0][1];
                if (sheet.Name != "Deductor's Sheet" || sheet.Name!= "Salary Details")
                {
                    try
                    {
                        int i;
                        int j;
                        int k;
                        int l;
                        int m;
                        int n;
                        int i1;
                        int i2;
                        int i3;
                        int i4;
                        int i5;
                        int i6;
                        ;
 
                        for (i = 65; i <= 66; i++)
                        {
                            for (j = 65; j <= 66; j++)
                            {
                                for (k = 65; k <= 66; k++)
                                {
                                    for (l = 65; l <= 66; l++)
                                    {
                                        for (m = 65; m <= 66; m++)
                                        {
                                            for (i1 = 65; i1 <= 66; i1++)
                                            {
                                                for (i2 = 65; i2 <= 66; i2++)
                                                {
                                                    for (i3 = 65; i3 <= 66; i3++)
                                                    {
                                                        for (i4 = 65; i4 <= 66; i4++)
                                                        {
                                                            for (i5 = 65; i5 <= 66; i5++)
                                                            {
                                                                for (i6 = 65; i6 <= 66; i6++)
                                                                {
                                                                    for (n = 32; n <= 126; n++)
                                                                     sheet.Unprotect ( i.ToString() + j.ToString() +k.ToString() + l.ToString() + m.ToString() + i1.ToString() + Strings.Chr(i2) + Strings.Chr(i3) + Strings.Chr(i4) + Strings.Chr(i5) + Strings.Chr(i6) + Strings.Chr(n));
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                       sheet.Rows.Cells[11][13] = "2019-2020";
                    }
                    catch { }
                }
             
            }
            workBook.Save();
            excelApp.Quit();
            update(filepath);
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
