using Microsoft.Office.Interop.Excel;

namespace ReadExcel
{
    public class MyExcelReader
    {
        //@"C:\test.xlsx";
        private string excelPath;
        private Application Excel { get; set; }
        private Workbook wb;
        private Worksheet ws;

        public MyExcelReader(string pathToExcel)
        {
            Excel = new Application();
            excelPath = pathToExcel;
        }
        public Workday GetWorkdayData(int sheetNumber)
        {
            OpenExcel();
            OpenWorksheet(sheetNumber);


        }

        /*
         Inner functionality
         */
        private void OpenExcel()
        {
            wb = Excel.Workbooks.Open(excelPath);
        }
        private void OpenWorksheet(int sheetNumber)
        {
            ws = wb.Worksheets[sheetNumber];
        }

    }
}
