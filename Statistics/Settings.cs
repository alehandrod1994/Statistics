namespace Statistics
{
    public class Settings
    {
        public Settings() : this(false) { }

        public Settings(bool debug)
        {
            Debug = debug;
        }

        public bool Debug { get; set; }
    }
}
