using System.Collections.Generic;

namespace Statistics
{
    public class Progress
    {
        public int stepNow;
        public int stepLast;
        public string text;
        public int value;

        public Progress()
        {
            stepNow = 0;
        }

        public void Move(string action, string typeDoc)
        {
            stepNow++;
            text = "Шаг " + stepNow + " из " + stepLast + ": " + action + " " + typeDoc;
            value = 100 / stepLast * stepNow * 8;
        }
    }
}
