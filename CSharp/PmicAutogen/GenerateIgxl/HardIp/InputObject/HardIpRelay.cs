namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    public class HardIpRelay
    {
        public HardIpRelay(string job, string name, RelayStatus status)
        {
            Job = job;
            Name = name;
            Status = status;
        }

        public string Job { get; set; }

        public string Name { get; set; }

        public RelayStatus Status { get; set; }
    }

    public enum RelayStatus
    {
        On = 0,
        Off
    }
}