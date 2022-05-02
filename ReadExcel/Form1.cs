using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ReadExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void read_Click(object sender, EventArgs e)
        {
            ReadExcelWithRangeAndUseWorkdayObj();
        }
        private void ReadExcel()
        {
            string filePath = @"C:\test.xls";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[7];
            // [sor, oszlop]
            MessageBox.Show(Convert.ToString(ws.Cells[8, 2].Value));

        }
        private void ReadExcel2()
        {
            string filePath = @"D:\test.xls";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[7];

            Range cell = ws.Range["B15"];
            
            string CellValue = cell.Value;
 
            MessageBox.Show(CellValue);
        }
        private void ReadExcelWithRange()
        {
            string filePath = @"C:\test.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[7];

            Range cell = ws.Range["B23:D23"];

            //string CellValue = cell.Value;
            foreach (int Result in cell.Value)
            {
                MessageBox.Show(Result.ToString());
            }
        }
        private void ReadExcelWithRangeAndUseWorkdayObj()
        {
            string filePath = @"C:\test.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[7];

            Range cell = ws.Range["B23:D23"];
            List<int> dates = new List<int>();

            foreach (int Result in cell.Value)
            {
                dates.Add(Result);
            }


            Workday workday = new Workday(dates);
            workday.WriteToConsole();
        }
    }
}
