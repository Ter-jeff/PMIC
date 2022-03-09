namespace AutomationCommon.DataStructure
{
    public enum MessageLevel
    {
        General,
        EndPoint,
        CheckPoint,
        Warning,
        Result,
        Error,
    };

    public class SettingStatus
    {
        public MessageLevel MsgLvl;
        public string Message;
        public int Percentage;
        public SettingStatus()
        {
            MsgLvl = MessageLevel.General;
            Message = "";
            Percentage = 0;
        }
    }
}