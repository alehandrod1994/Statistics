using System;
using System.Collections.Generic;

namespace Statistics
{
    public class Date
    {
        public string day1;
        public string day2;
        public string monthNum1;
        public string monthNum2;
        public string month;
        public string year1;
        public string year2;
        public string date1;
        public string date2;
        public List<string> listMonth;

        public Date()
        {
            listMonth = new List<string>();
            listMonth.Add("Январь");
            listMonth.Add("Февраль");
            listMonth.Add("Март");
            listMonth.Add("Апрель");
            listMonth.Add("Май");
            listMonth.Add("Июнь");
            listMonth.Add("Июль");
            listMonth.Add("Август");
            listMonth.Add("Сентябрь");
            listMonth.Add("Октябрь");
            listMonth.Add("Ноябрь");
            listMonth.Add("Декабрь");
        }

        public void AutoImport()
        {
            string dayOfWeek = DateTime.Today.DayOfWeek.ToString();
            int day = Convert.ToInt32(DateTime.Today.Day);
            string todayMonth = DateTime.Today.Month.ToString();


            if (day < 10 && todayMonth == "01")
            {
                month = listMonth[11];
                year1 = (Convert.ToInt32(DateTime.Today.Year) - 1).ToString();
            }
            else if (day < 10)
            {
                month = listMonth[Convert.ToInt32(todayMonth) - 2];
                year1 = DateTime.Today.Year.ToString();
            }
            else
            {
                month = listMonth[Convert.ToInt32(todayMonth) - 1];
                year1 = DateTime.Today.Year.ToString();
            }


            if (dayOfWeek == "Friday")
            {
                date1 = (DateTime.Now.AddDays(-4)).ToShortDateString();
                date2 = DateTime.Today.ToString();
            }
            else if (dayOfWeek == "Saturday")
            {
                date1 = (DateTime.Now.AddDays(-5)).ToShortDateString();
                date2 = DateTime.Today.ToString();
            }
            else if (dayOfWeek == "Sunday")
            {
                date1 = (DateTime.Now.AddDays(-6)).ToShortDateString();
                date2 = DateTime.Today.ToString();
            }
            else if (dayOfWeek == "Monday")
            {
                date1 = (DateTime.Now.AddDays(-7)).ToShortDateString();
                date2 = (DateTime.Now.AddDays(-1)).ToShortDateString();
            }
            else if (dayOfWeek == "Tuesday")
            {
                date1 = (DateTime.Now.AddDays(-8)).ToShortDateString();
                date2 = (DateTime.Now.AddDays(-2)).ToShortDateString();
            }
            else if (dayOfWeek == "Wednesday")
            {
                date1 = (DateTime.Now.AddDays(-9)).ToShortDateString();
                date2 = (DateTime.Now.AddDays(-3)).ToShortDateString();
            }
        }

    }
}
