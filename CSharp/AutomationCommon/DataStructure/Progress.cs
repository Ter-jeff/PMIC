namespace AutomationCommon.DataStructure
{
    public class ProgressStatus
    {
        public int Percentage { get; set; }
        public string Result { get; set; }
        public MessageLevel Level { get; set; }
        public bool EnableMsg { get; set; }

        public ProgressStatus()
        {
            Percentage = 0;
            Result = string.Empty;
            Level = MessageLevel.General;
            EnableMsg = true;
        }
    }
}
