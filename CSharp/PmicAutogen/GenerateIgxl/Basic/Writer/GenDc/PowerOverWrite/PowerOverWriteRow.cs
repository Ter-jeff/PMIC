using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDc.PowerOverWrite
{
    public class PowerOverWriteRow
    {
        private string _pinName;

        public PowerOverWriteRow()
        {
            PinName = "";
            Nv = "";
            NvValt = "";
            LvRatio = "";
            HvRatio = "";
            Ids = "";
            Ifold = "";
            Vil = "";
            Vih = "";
            Vol = "";
            Voh = "";
            Iol = "";
            Ioh = "";
            Vt = "";
            Vcl = "";
            Vch = "";
            DriveMode = "";
            Vicm = "";
            Vid = "";
            Vod = "";
            NeedRatio = false;
        }

        public string PinName
        {
            set
            {
                _pinName = value;
                if (Regex.IsMatch(_pinName, "^VDD", RegexOptions.IgnoreCase))
                    PinType = HardIpDcPinType.Power;
                else if (Regex.IsMatch(_pinName, "^Pins", RegexOptions.IgnoreCase))
                    PinType = HardIpDcPinType.LevelIo;
                else if (Regex.IsMatch(_pinName, "DIFF", RegexOptions.IgnoreCase))
                    PinType = HardIpDcPinType.IoDiff;
                else
                    PinType = HardIpDcPinType.IoSingle;
            }
            get { return _pinName; }
        }

        public string Nv { set; get; }
        public string NvValt { set; get; }
        public string LvRatio { set; get; }
        public string HvRatio { set; get; }
        public string Ids { set; get; }
        public string Ifold { set; get; }
        public string Vil { set; get; }
        public string Vih { set; get; }
        public string Vol { set; get; }
        public string Voh { set; get; }
        public string Iol { set; get; }
        public string Ioh { set; get; }
        public string Vt { set; get; }
        public string Vcl { set; get; }
        public string Vch { set; get; }
        public string DriveMode { set; get; }
        public string Vicm { set; get; }
        public string Vid { set; get; }
        public string Vod { set; get; }
        public string RowNum { set; get; }


        public HardIpDcPinType PinType { set; get; }

        public bool NeedRatio { set; get; }
    }
}