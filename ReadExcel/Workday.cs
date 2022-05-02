using System.Collections.Generic;
using System.Windows.Forms;

namespace ReadExcel
{
    public class Workday
    {
        public int Year { get; set; }
        public int Month { get; set; }
        public int Day { get; set; }

        public int FieldWorkStart { get; set; }
        public int FieldWorkEnd { get; set; }

        public int WorkStart { get; set; }
        public int WorkEnd { get; set; }

        public Workday(List<int> dates)
        {
            MessageBox.Show(dates.Count.ToString());
            Year = dates[0];
            Month = dates[1];
            Day = dates[2];
        }

        public void WriteToConsole()
        {
            MessageBox.Show($"Year: {Year}, Month: {Month}, Day: {Day}");
        }
    }
}
