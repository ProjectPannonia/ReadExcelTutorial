using System.Collections.Generic;
using System.Windows.Forms;

namespace ReadExcel
{
    public class FieldWork
    {
        public int FieldWorkStart { get; set; }
        public int FieldWorkEnd { get; set; }

        public int WorkStart { get; set; }
        public int WorkEnd { get; set; }

        public FieldWork(int fieldWorkStart, int fieldWorkEnd, int workStart, int workEnd)
        {
            FieldWorkStart = fieldWorkStart;
            FieldWorkEnd = fieldWorkEnd;
            WorkStart = workStart;
            WorkEnd = workEnd;
        }
        public FieldWork(List<int> workingHours)
        {
            FieldWorkStart = workingHours[0];
            FieldWorkEnd = workingHours[1];
            WorkStart = workingHours[2];
            WorkEnd = workingHours[3];
        }

        public void Show()
        {
            MessageBox.Show($"Travel started: {FieldWorkStart}, Travel ended: {FieldWorkEnd}, Work started: {WorkStart}, Work ended: {WorkEnd}");
        }
    }
}
