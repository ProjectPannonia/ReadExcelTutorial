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
            //ReadExcelWithRangeAndUseWorkdayObj();
            ReadExcel();

            string[] fieldworkTime = ReadExcelGetFieldWorkTime();

        }
        private string[] ReadExcelGetFieldWorkTime()
        {
            string filePath = @"C:\test.xlsx";
            Microsoft.Office.Interop.Excel.Application excel =
                new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[7];

            double d = ws.Cells[23, 5].Value;
            DateTime conv = DateTime.FromOADate(d);
            string hour = conv.Hour.ToString();

            if (hour.Length == 1)
                hour = "0" + hour;

            string minute = conv.Minute.ToString();
            if (minute.Length == 1)
                minute += "0";

            return new string[] { hour, minute };
            //MessageBox.Show(hour + " : " + minute);
        }
        private void ReadExcel()
        {
            string filePath = @"C:\test.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = 
                new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[7];
            // [sor, oszlop]
            //MessageBox.Show(Convert.ToString(ws.Cells[23, 5].Value));
            // Kiszállás kezdete pozíció
            double d = ws.Cells[23, 5].Value;
            DateTime conv = DateTime.FromOADate(d);
            string hour = conv.Hour.ToString();
            if(hour.Length == 1) 
                hour = "0" + hour;
            string minute = conv.Minute.ToString();
            if (minute.Length == 1)
                minute += "0";
            //MessageBox.Show(conv.ToString());
            MessageBox.Show(hour + " : " + minute);
        }
        private void ReadExcel2()
        {
            string filePath = @"C:\test.xlsx";
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb;
            Worksheet ws;

            wb = excel.Workbooks.Open(filePath);
            ws = wb.Worksheets[7];

            Range cell = ws.Range["E23"];
            
            string val = cell.Value.ToString("HH:mm");
            //MessageBox.Show(cell.GetType());
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

            Range cell = ws.Range["B23:H23"];
            List<int> dates = new List<int>();
            MessageBox.Show("Number of cells: " + cell.Count.ToString());
            foreach (int Result in cell.Value)
            {
                MessageBox.Show(Result.ToString());
                dates.Add(Result);
            }

            Workday workday = new Workday(dates);
            workday.WriteToConsole();
        }
    }
}
