using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml;
using System.Reflection;
using System.Threading;

namespace ExcelFile
{
    public partial class Form1 : Form
    {
        Excel.Application xlap = new Excel.Application();
         

        public Form1()
        {
            InitializeComponent();

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            if (xlap == null)
            {
                label1.Text = "Excel Library is not installed ";
                label1.ForeColor = Color.Red;
            }
            else
            {
                label1.Text = "Excel Library is installed ";
                label1.ForeColor = Color.Green;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            ProcessStartInfo procInfo = new ProcessStartInfo();

            procInfo.Arguments = "https://www.bowok.pp.ua/sitemap.xml?page=1";
            XmlDocument xDoc = new XmlDocument();
            xDoc.Load(procInfo.Arguments);

            // получим корневой элемент
            XmlElement xRoot = xDoc.DocumentElement;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


         
            xlWorkSheet.Cells[1, 1] = "Ref";
            xlWorkSheet.Cells[1, 2] = "Date";
            int number = 2;
            // обход всех узлов в корневом элементе
            foreach (XmlNode xnode in xRoot)
            {
               
                foreach (XmlNode childnode in xnode.ChildNodes)
                {
                    
                   
                    // если узел - loc
                    if (childnode.Name == "loc")
                    {
                       
                        xlWorkSheet.Cells[number, 1] = '"' + childnode.InnerText + '"';
                    }
                    //  если узел loc
                    if (childnode.Name == "lastmod")
                    {
                        
                        xlWorkSheet.Cells[number, 2] = '"' + childnode.InnerText + '"';
                    }
                    
                }
                number++;
            }


            xlWorkBook.SaveAs(@"C:\Users\user\Desktop\parsing.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
