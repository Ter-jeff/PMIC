using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDc.PowerOverWrite
{
    public class PowerOverWrite
    {
        public PowerOverWrite(string categoryName)
        {
            CategoryName = categoryName;
            DataRows = new List<PowerOverWriteRow>();
        }

        public string CategoryName { set; get; }
        public List<PowerOverWriteRow> DataRows { set; get; }
        public string LevelSheet { set; get; }
        public string DcCategory { set; get; }

        public string GetLevelName()
        {
            if (DataRows.Exists(p =>
                p.PinType.Equals(HardIpDcPinType.IoDiff) || p.PinType.Equals(HardIpDcPinType.IoSingle)))
                return "Levels_" + CategoryName;
            if (DataRows.Exists(p => p.PinType.Equals(HardIpDcPinType.Power) && p.Ifold != ""))
                return "Levels_" + CategoryName;
            if (DataRows.Exists(p =>
                p.PinType.Equals(HardIpDcPinType.LevelIo) &&
                (p.Iol != "" || p.Ioh != "" || p.Vcl != "" || p.Vch != "" || p.DriveMode != "")))
                return "Levels_" + CategoryName;
            return "Levels_" + CategoryName; //"Levels_HardIP";
        }

        public List<HardIpSpecValue> GetSpecValueFromDef(PowerOverWriteRow row)
        {
            var specValueList = new List<HardIpSpecValue>();
            var pinType = row.PinType;
            const bool hasRatio = true;

            var vmain = FilterDcSpecValue(row.Nv);
            var ifold = FilterDcSpecValue(row.Ifold);

            var vil = FilterDcSpecValue(row.Vil);
            var vih = FilterDcSpecValue(row.Vih);
            var vol = FilterDcSpecValue(row.Vol);
            var voh = FilterDcSpecValue(row.Voh);
            var iol = FilterDcSpecValue(row.Iol);
            var ioh = FilterDcSpecValue(row.Ioh);
            var vt = FilterDcSpecValue(row.Vt);
            var vcl = FilterDcSpecValue(row.Vcl);
            var vch = FilterDcSpecValue(row.Vch);

            var vid = FilterDcSpecValue(row.Vid);
            var vod = FilterDcSpecValue(row.Vod);
            var vicm = FilterDcSpecValue(row.Vicm);
            if (pinType.Equals(HardIpDcPinType.Power))
            {
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vmain", vmain, hasRatio, false, row.HvRatio,
                    row.LvRatio));
                if (!string.IsNullOrEmpty(row.Ifold))
                    specValueList.Add(new HardIpSpecValue(row.PinName, "iFoldLevel", ifold, false));
            }

            if (pinType.Equals(HardIpDcPinType.Dcvi))
            {
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vps", vmain, hasRatio, false, row.HvRatio,
                    row.LvRatio));
                if (!string.IsNullOrEmpty(row.Ifold))
                    specValueList.Add(new HardIpSpecValue(row.PinName, "Isc", ifold, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "tdelay", "0", false));
            }

            if (pinType.Equals(HardIpDcPinType.LevelIo))
            {
                specValueList.Add(new HardIpSpecValue(row.PinName, "vil", vil, hasRatio, false, row.HvRatio,
                    row.LvRatio));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vih", vih, hasRatio, false, row.HvRatio,
                    row.LvRatio));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vol", vol, hasRatio, false, row.HvRatio,
                    row.LvRatio));
                specValueList.Add(new HardIpSpecValue(row.PinName, "voh", voh, hasRatio, false, row.HvRatio,
                    row.LvRatio));
                specValueList.Add(new HardIpSpecValue(row.PinName, "iol", iol, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "ioh", ioh, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vt", vt, hasRatio, false, row.HvRatio,
                    row.LvRatio));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vcl", vcl, hasRatio, false, row.HvRatio,
                    row.LvRatio));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vch", vch, hasRatio, false, row.HvRatio,
                    row.LvRatio));
                specValueList.Add(new HardIpSpecValue(row.PinName, "driverMode", row.DriveMode, false));
            }

            if (pinType.Equals(HardIpDcPinType.IoSingle))
            {
                specValueList.Add(new HardIpSpecValue(row.PinName, "vil", vil, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vih", vih, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vol", vol, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "voh", voh, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "voh_alt1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "voh_alt2", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "iol", row.Iol, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "ioh", row.Ioh, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vt", vt, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vcl", vcl, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "vch", vch, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "voutLoTyp", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "voutHiTyp", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "driverMode", row.DriveMode, false));
            }

            if (pinType.Equals(HardIpDcPinType.IoDiff))
            {
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vicm", vicm, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vid", vid, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "dVid0", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "dVid1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "dVicm0", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "dVicm1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vod", vod, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vod_alt1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vod_alt2", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "dVod0", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "dVod1", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "Iol", iol, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "Ioh", ioh, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "VodTyp", vod, false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "VocmTyp", "0", false, true));
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vt", vt, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vcl", vcl, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "Vch", vch, false));
                specValueList.Add(new HardIpSpecValue(row.PinName, "DriverMode", row.DriveMode, false));
            }

            specValueList.RemoveAll(p => p.Value.Equals(""));
            return specValueList;
        }

        private string FilterDcSpecValue(string specValue)
        {
            var regexPattern = @"[a-zA-Z]\w+";
            if (!Regex.IsMatch(specValue, regexPattern)) return specValue;
            var resultValue = Regex.Replace(specValue, regexPattern, m => m.Value);
            return resultValue;
        }
    }

    public enum HardIpDcPinType
    {
        Power,
        Dcvi,
        LevelIo,
        IoSingle,
        IoDiff
    }

    public class HardIpSpecValue : LevelRow
    {
        public HardIpSpecValue(string pinName, string parameter, string value, bool has, bool fromDefault = false,
            string hv = "", string lv = "", string comment = "") : base(pinName, parameter, value, comment)
        {
            PinName = pinName;
            Parameter = parameter;
            Value = value;
            Comment = comment;
            HasRatio = has;
            FromDefault = fromDefault;
            Plus = hv;
            Minus = lv;

            if (Value == "")
            {
                //mark by Raze for removing default if the cell content is blank 2017/06/12
                if (Parameter.Equals("driverMode", StringComparison.OrdinalIgnoreCase))
                    Value = "HiZ";
                else if (Parameter.Equals("Vcl", StringComparison.OrdinalIgnoreCase))
                    Value = "-1";
                else if (Parameter.Equals("Vch", StringComparison.OrdinalIgnoreCase))
                    Value = "6";
                else
                    Value = "0";
                //2017/6/28 anderson add for get initial value 
                FromDefault = true;
            }
        }

        //if this spec have HV, LV ratio
        public bool
            FromDefault { set; get; } //this flag will decide if it is need to overwrite spec value in Level sheet

        public bool HasRatio { set; get; } //this flag will decide if it is need to change value in DC Spec
        public string Plus { set; get; }
        public string Minus { set; get; }
    }
}