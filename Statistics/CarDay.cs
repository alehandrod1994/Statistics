using System;


namespace Statistics
{
    public class CarDay
    {
        public CarDay() { }

        public CarDay(string date)
        {
            Date = date;
            CarsOneTime = 0;
            CarsPermanent = 0;
        }

        public string Date { get; set; }
        public int CarsOneTime { get; set; }
        public int CarsPermanent { get; set; }

    }
}
