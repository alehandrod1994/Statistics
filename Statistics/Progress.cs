namespace Statistics
{
    public class Progress
    {
        public Progress() => StepNow = 0;

        public int StepNow { get; private set; }
        public int StepLast { get; set; }
        public string Text { get; private set; }
        public int Value { get; private set; }

        public void Move(string action, string typeDoc)
        {
            StepNow++;
            Text = $"Шаг {StepNow} из {StepLast}: {action} \"{typeDoc}\"";
            Value = 100 / StepLast * StepNow * 8;
        }
    }
}
