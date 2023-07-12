using System;

namespace Statistics
{
    public abstract class Parser
    {
        public static int ParseInt(string value)
        {
            if (int.TryParse(value, out int result))
            {
                return result;
            }

            return default;
        }

        public static double ParseDouble(string value)
        {
            if (double.TryParse(value, out double result))
            {
                return result;
            }

            return default;
        }

        public static DateTime ParseDateTime(string value)
        {
            if (DateTime.TryParse(value, out DateTime result))
            {
                return result;
            }

            return default;
        }
    }
}
