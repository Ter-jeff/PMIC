using System;
using System.Collections.Generic;
using System.Linq;

namespace ShmooLog.Base
{
    public class SelSramCondition
    {
        private const int SiteIndex = 1;
        private const int XAxisIndex = 2;
        private const int YAxisIndex = 3;
        private const int InstIndex = 4;
        private const int PowerSetIndex = 5;
        private const int DigSourceIndex = 6;
        private const int CompressIndex = 7;
        public string CompareCompressStr = "";
        public string InstanceName = "";
        public string OrgCompressStr = "";
        public string OrgDigSourecStr = "";
        public List<PowerSetting> PowerSetting = new List<PowerSetting>();
        public string PowerSettingStr = "";

        public string Site = "";
        public string Xaxis = "";
        public string Yaxis = "";

        public SelSramCondition(string data)
        {
            var array = data.ToUpper().Replace("[", "").Replace("]", "").Split(',');
            Site = array[SiteIndex];
            Xaxis = array[XAxisIndex];
            Yaxis = array[YAxisIndex];
            InstanceName = array[InstIndex];
            OrgDigSourecStr = array[DigSourceIndex];
            OrgCompressStr = array[CompressIndex];
            PowerSettingStr = array[PowerSetIndex];
            foreach (var powerSet in PowerSettingStr.Split(';'))
            {
                var powerPin = powerSet.Split(':').First();
                var value = Convert.ToDouble(powerSet.Split(':').Last());
                PowerSetting.Add(new PowerSetting(powerPin, value));
            }
        }

        public PowerSetting GetPowerPinItem(string powerPin)
        {
            return PowerSetting.FirstOrDefault(p =>
                p.PowerPinName.Equals(powerPin, StringComparison.InvariantCultureIgnoreCase));
        }

        public string GetXYDie()
        {
            return Xaxis + "," + Yaxis;
        }

        public bool Check()
        {
            return OrgCompressStr == CompareCompressStr;
        }
    }

    public class PowerSetting
    {
        public PowerSetting(string powerName, double value)
        {
            PowerPinName = powerName;
            Value = value;
        }

        public string PowerPinName { get; set; }
        public double Value { get; set; }
    }
}