using System.Collections.Generic;
using System.Windows.Forms;

namespace ReadExcel
{
    public class Date
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public int Day { get; set; }

        public Date(int year, int month, int day)
        {
            Year = year;
            Month = month;
            Day = day;
        }
        public Date(List<int> dateList)
        {
            Year = dateList[0];
            Month = dateList[1];
            Day = dateList[2];
        }
        public void ShowDate()
        {
            MessageBox.Show($"Year: {Year}, Month: {Month}, Day: {Day}");
        }
    }
}