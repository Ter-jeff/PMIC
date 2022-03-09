namespace PmicAutogen.Inputs.Setting.BinNumber
{
    public class BinNumberRuleRow
    {
        public BinNumberRuleRow()
        {
            Description = "";
            SoftBinStart = 0;
            SoftBinEnd = 0;
            SoftBinState = "Fail";
            CurrentSoftBin = 0;
            HardBin = "0";
            HardIpHlvBin = "";
            HardIpHvBin = "";
            HardIpLvBin = "";
            HardIpNvBin = "";
            CurrentBinLib = new SoftBinRangeRow();
            IsExceed = false;
        }

        public string Description { set; get; }
        public int SoftBinStart { set; get; }
        public int SoftBinEnd { set; get; }
        public string SoftBinState { set; get; }

        public int CurrentSoftBin { set; get; }

        public string HardBin { set; get; }
        public string HardIpHvBin { set; get; }
        public string HardIpHlvBin { set; get; }
        public string HardIpLvBin { set; get; }
        public string HardIpNvBin { set; get; }
        public SoftBinRangeRow CurrentBinLib { set; get; }

        public bool IsExceed { set; get; }
    }
}