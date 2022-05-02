using System.Collections.Generic;

namespace ReadExcel
{
    public class Workday
    {
        public Date Date { get; set; }
        public FieldWork FieldWork { get; set; }
        private List<int> dateList;
        private List<int> workdayInfo;


        public Workday(List<int> dataList)
        {
            dateList = new List<int>();
            workdayInfo = new List<int>();
            SeparateLists(dataList);
            Date = new Date(dateList);
            FieldWork = new FieldWork(workdayInfo);
        }
        private void SeparateLists(List<int> dataList)
        {
            int counter = 0;

            foreach(int item in dataList)
            {
                if(counter < 3)
                    dateList.Add(item);
                else
                    workdayInfo.Add(item);
                counter++;
            }
        }
        public void WriteToConsole()
        {
            Date.ShowDate();
            FieldWork.Show();
        }
    }
}
