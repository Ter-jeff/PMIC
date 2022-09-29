namespace PmicAutogen.GenerateIgxl.Basic.GenPatSet
{
    public class PatSetStatus
    {
        public PatSetStatus()
        {
            Used = "NoCheck"; //1st check rule
            ValidTs = "NoCheck"; //2nd check rule
            Existed = "NoCheck"; //3rd check rule
            ContainSub = "NoCheck"; // 4th check rule
        }

        public string Used { get; set; }
        public string Existed { get; set; }
        public string ValidTs { get; set; }
        public string ContainSub { get; set; }
    }
}