using System;
using System.Collections.Generic;

namespace Statistics
{
    public class Date
    {
        public Date()
        {
            ListMonth = new List<string>();
            ListMonth.Add("Январь");
            ListMonth.Add("Февраль");
            ListMonth.Add("Март");
            ListMonth.Add("Апрель");
            ListMonth.Add("Май");
            ListMonth.Add("Июнь");
            ListMonth.Add("Июль");
            ListMonth.Add("Август");
            ListMonth.Add("Сентябрь");
            ListMonth.Add("Октябрь");
            ListMonth.Add("Ноябрь");
            ListMonth.Add("Декабрь");
        }

        public string Day1 { get; set; }
        public string Day2 { get; set; }
        public string MonthNum1 { get; set; }
        public string MonthNum2 { get; set; }
        public string Month { get; set; }
        public string Year1 { get; set; }
        public string Year2 { get; set; }
        public string Date1 { get; set; }
        public string Date2 { get; set; }

        public List<string> ListMonth { get; }     

        public void AutoImport()
        {
            string dayOfWeek = DateTime.Today.DayOfWeek.ToString();
            int day = Convert.ToInt32(DateTime.Today.Day);
            string todayMonth = DateTime.Today.Month.ToString();


            if (day < 10 && todayMonth == "1")
            {
                Month = ListMonth[11];
                Year1 = (Convert.ToInt32(DateTime.Today.Year) - 1).ToString();
            }
            else if (day < 10)
            {
                Month = ListMonth[Convert.ToInt32(todayMonth) - 2];
                Year1 = DateTime.Today.Year.ToString();
            }
            else
            {
                Month = ListMonth[Convert.ToInt32(todayMonth) - 1];
                Year1 = DateTime.Today.Year.ToString();
            }


            if (dayOfWeek == "Friday")
            {
                Date1 = (DateTime.Now.AddDays(-4)).ToShortDateString();
                Date2 = DateTime.Today.ToString();
            }
            else if (dayOfWeek == "Saturday")
            {
                Date1 = (DateTime.Now.AddDays(-5)).ToShortDateString();
                Date2 = DateTime.Today.ToString();
            }
            else if (dayOfWeek == "Sunday")
            {
                Date1 = (DateTime.Now.AddDays(-6)).ToShortDateString();
                Date2 = DateTime.Today.ToString();
            }
            else if (dayOfWeek == "Monday")
            {
                Date1 = (DateTime.Now.AddDays(-7)).ToShortDateString();
                Date2 = (DateTime.Now.AddDays(-1)).ToShortDateString();
            }
            else if (dayOfWeek == "Tuesday")
            {
                Date1 = (DateTime.Now.AddDays(-8)).ToShortDateString();
                Date2 = (DateTime.Now.AddDays(-2)).ToShortDateString();
            }
            else if (dayOfWeek == "Wednesday")
            {
                Date1 = (DateTime.Now.AddDays(-9)).ToShortDateString();
                Date2 = (DateTime.Now.AddDays(-3)).ToShortDateString();
            }
        }

    }
}
