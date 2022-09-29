using CommonLib.Enum;

namespace CommonLib.WriteMessage
{
    public class ProgressStatus
    {
        public ProgressStatus()
        {
            Percentage = 0;
            Message = string.Empty;
            Level = EnumMessageLevel.General;
            EnableMsg = true;
        }

        public string Message { get; set; }
        public EnumMessageLevel Level { get; set; }
        public int Percentage { get; set; }
        public bool EnableMsg { get; set; }
    }
}