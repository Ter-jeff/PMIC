using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using OfficeOpenXml;

namespace IgxlData.IgxlSheets
{
    public class PinMapConst
    {
        public const string TypeIo = "I/O";
        public const string TypeInput = "Input";
        public const string TypeOutput = "Output";
        public const string TypeAnalog = "Analog";
        public const string TypePower = "Power";
        public const string TypeGnd = "Gnd";
        public const string TypeUtility = "Utility";
        public const string TypeVoltage = "Voltage";
        public const string TypeCurrent = "Current";
        public const string TypeUnknown = "Unknown";
    }

    public class PinMapSheet : IgxlSheet
    {
        #region Property

        public List<Pin> PinList { get; set; }
        public List<PinGroup> GroupList { get; set; }

        #endregion

        #region Constructor

        public PinMapSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            PinList = new List<Pin>();
            GroupList = new List<PinGroup>();
            IgxlSheetName = IgxlSheetNameList.PinMap;
        }

        public PinMapSheet(string sheetName)
            : base(sheetName)
        {
            PinList = new List<Pin>();
            GroupList = new List<PinGroup>();
            IgxlSheetName = IgxlSheetNameList.PinMap;
        }

        #endregion

        #region Member Function

        protected override void WriteHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteColumnsHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteRows()
        {
            throw new NotImplementedException();
        }

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
                version = "2.1";
            if (version == "2.1")
                using (var sw = new StreamWriter(fileName, false))
                {
                    sw.WriteLine("DTPinMap,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tPin Map");
                    sw.WriteLine("\t\t\t\t\t");
                    sw.WriteLine("\tGroup Name\tPin Name\tType\tComment");

                    foreach (var pin in PinList)
                        sw.WriteLine("\t{0}\t{1}\t{2}\t{3}", "", pin.PinName.Replace("/", ""), pin.PinType,
                            pin.Comment);
                    foreach (var pinGroup in GroupList)
                    foreach (var pin in pinGroup.PinList)
                    {
                        if (string.IsNullOrEmpty(pin.PinType) && IsPinExist(pin.PinName))
                            pin.PinType = GetPin(pin.PinName).PinType;
                        sw.WriteLine("\t{0}\t{1}\t{2}\t{3}", pinGroup.PinName, pin.PinName.Replace("/", ""),
                            pin.PinType, pin.Comment);
                    }
                }
            else
                throw new Exception(string.Format("The PinMap sheet version:{0} is not supported!", version));
        }

        public void WriteOld(string fileName, string version = "2.1")
        {
            //if (version == "2.1")
            //{
            //    Action<string> validate = new Action<string>((a) => { });
            //    GenPinMap pinMapGen = new GenPinMap(fileName, validate, true);
            //    foreach (var pin in _pinList)
            //    {
            //        pinMapGen.AddPin(pin.PinName, pin.PinType, pin.Comment);
            //    }
            //    foreach (var pinGroup in _groupList)
            //    {
            //        pinMapGen.AddPinGroup(pinGroup.PinName, pinGroup.PinList, pinGroup.PinType, pinGroup.Comment);
            //    }
            //    pinMapGen.WriteSheet();
            //}
            //else
            //    throw new Exception(string.Format("The PinMap sheet version:{0} is not supported!", version));
        }

        public void AddRow(Pin pin)
        {
            AddPin(pin);
        }

        public void AddRow(PinGroup pinGroup)
        {
            AddPinGroup(pinGroup);
        }

        public string GetDiffGroupName(string[] pair)
        {
            if (pair.Length == 2)
            {
                var group = GroupList.FirstOrDefault(x => x.PinList.Count == 2 &&
                                                          x.PinList.Exists(p =>
                                                              p.PinName.Equals(pair[0],
                                                                  StringComparison.OrdinalIgnoreCase)) &&
                                                          x.PinList.Exists(p =>
                                                              p.PinName.Equals(pair[1],
                                                                  StringComparison.OrdinalIgnoreCase)));
                if (group != null)
                    return group.PinName;
            }

            return "";
        }

        public bool IsChannelType(string pinName, string type)
        {
            if (IsGroupExist(pinName))
            {
                var groupPins = GetGroup(pinName).PinList;
                var flag = groupPins.All(
                    x => GetChannelType(x.PinName).Equals(type, StringComparison.OrdinalIgnoreCase));
                return flag;
            }

            if (GetChannelType(pinName).Equals(type, StringComparison.OrdinalIgnoreCase))
                return true;
            return false;
        }

        public string GetChannelType(string pinName)
        {
            var type = "";
            if (PinList.Exists(x => x.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                type = PinList.Find(x => x.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)).ChannelType;
            return type;
        }

        public List<Pin> GetPowerPins()
        {
            var pinList = PinList.FindAll(p =>
                p.PinType.Equals(PinMapConst.TypePower, StringComparison.OrdinalIgnoreCase));
            return pinList;
        }

        public List<Pin> GetIoPins()
        {
            var pinList =
                PinList.FindAll(p => p.PinType.Equals(PinMapConst.TypeIo, StringComparison.OrdinalIgnoreCase));
            return pinList;
        }

        public List<Pin> GetIoContinuityPins()
        {
            var pinList = GetAllDigitalDisconnectContinuityPins();
            pinList = pinList.Where(x => !(x.PinName.StartsWith("REFCLK_", StringComparison.OrdinalIgnoreCase) ||
                                           x.PinName.EndsWith("_PA", StringComparison.OrdinalIgnoreCase)
                )).ToList();
            return pinList;
        }

        public List<Pin> GetAllDigitalDisconnectContinuityPins()
        {
            var pinList = PinList.FindAll(p =>
                p.ChannelType.Equals(PinMapConst.TypeIo, StringComparison.OrdinalIgnoreCase));
            pinList = pinList.Where(x =>
                !((x.PinName.StartsWith("VDD", StringComparison.OrdinalIgnoreCase) &&
                   x.PinName.ToUpper().Contains("SENSE")) ||
                  (x.PinName.StartsWith("VDD", StringComparison.OrdinalIgnoreCase) &&
                   x.PinName.ToUpper().Contains("MONITOR")) ||
                  (x.PinName.StartsWith("VSS", StringComparison.OrdinalIgnoreCase) &&
                   x.PinName.ToUpper().Contains("SENSE"))
                    )).ToList();
            return pinList;
        }

        #region pin

        private void AddPin(Pin pin)
        {
            if (!IsPinExist(pin.PinName)) PinList.Add(pin);
        }

        public bool IsPinExist(string pinName)
        {
            if (PinList.Exists(p => p.PinName.ToLower().Equals(pinName.ToLower())))
                return true;
            return false;
        }

        public Pin GetPin(string pinName)
        {
            var pin = PinList.Find(p => p.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase));
            return pin;
        }

        public string GetPinType(string pinName)
        {
            if (!PinList.Exists(p => p.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                if (!GroupList.Exists(p => p.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                    return "";
                GetPinType(GroupList.First().PinName);
            }

            var type = PinList.Find(p => p.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)).PinType;
            return type;
        }

        #endregion

        #region pinGroup

        public void AddPinGroup(PinGroup pinGroup)
        {
            if (!IsGroupExist(pinGroup.PinName))
            {
                if (pinGroup.PinList.Count == 0)
                    throw new Exception(string.Format("There isn't any pin in PinGroup:{0} . ", pinGroup.PinName));

                if (!Regex.IsMatch(pinGroup.PinName, @"^efuse|All_DiffPairs", RegexOptions.IgnoreCase))
                    if (pinGroup.PinType != "Differential")
                        pinGroup.PinList = pinGroup.PinList.OrderBy(x => x.PinName, new SemiNumericComparer()).ToList();

                GroupList.Add(pinGroup);
            }
        }

        public void AddPinGroups(List<PinGroup> pinGroups)
        {
            foreach (var pinGroup in pinGroups)
            {
                if (pinGroup.PinList == null || pinGroup.PinList.Count == 0)
                    continue;
                AddRow(pinGroup);
            }
        }

        public bool IsGroupExist(string groupName)
        {
            if (GroupList.Exists(p => p.PinName.ToLower().Equals(groupName.ToLower())))
                return true;
            return false;
        }

        public void SortPinGroup()
        {
            foreach (var pinGroup in GroupList)
                if (!Regex.IsMatch(pinGroup.PinName, @"^efuse|All_DiffPairs", RegexOptions.IgnoreCase))
                    if (pinGroup.PinType != "Differential")
                        pinGroup.PinList = pinGroup.PinList.OrderBy(x => x.PinName, new SemiNumericComparer()).ToList();
        }

        public PinGroup GetGroup(string groupName)
        {
            var group = GroupList.FirstOrDefault(p => p.PinName.ToLower().Equals(groupName.ToLower()));
            return group;
        }

        public List<Pin> GetPinsFromGroup(string pinGroup)
        {
            var resultPins = new List<Pin>();
            if (!GroupList.Exists(p => p.PinName.Equals(pinGroup, StringComparison.OrdinalIgnoreCase)))
                return new List<Pin>();
            var pinOrGroup = GroupList.Find(p => p.PinName.Equals(pinGroup, StringComparison.OrdinalIgnoreCase))
                .PinList;
            foreach (var pin in pinOrGroup)
            {
                if (pin.PinName.Equals(pinGroup, StringComparison.CurrentCultureIgnoreCase))
                {
                    resultPins.Add(pin);
                    continue;
                }

                if (IsGroupExist(pin.PinName))
                {
                    resultPins.AddRange(GetPinsFromGroup(pin.PinName));
                }
                else
                {
                    var singlePin =
                        PinList.Find(p => p.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase));
                    resultPins.Add(singlePin ?? new Pin(pin.PinName, "", "Does not exist in Pin map sheet"));
                }
            }

            return resultPins;
        }

        public List<PinGroup> GenPinGroupByInstrumentType()
        {
            var pinGroups = new List<PinGroup>();
            var groups = PinList.GroupBy(x => x.InstrumentType);
            foreach (var group1 in groups)
            {
                if (string.IsNullOrEmpty(group1.Key)) continue;

                foreach (var group2 in group1.GroupBy(x => x.PinType))
                {
                    var groupPinName = group1.GroupBy(x => x.PinType).Count() != 1
                        ? "All_" + group1.Key.Replace("/", "") + "_" + group2.Key.Replace("/", "")
                        : "All_" + group1.Key.Replace("/", "");
                    var pinGroup = new PinGroup(groupPinName);
                    pinGroup.AddPins(group2.ToList());
                    if (pinGroup.PinList.Count != 0)
                        pinGroups.Add(pinGroup);
                }
            }

            return pinGroups;
        }

        public List<PinGroup> GenDcviGroup()
        {
            var pinGroups = new List<PinGroup>();
            var group1 = PinList.Where(a =>
                Regex.IsMatch(a.ChannelType, @"^DCVI", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(a.PinType, PinMapConst.TypePower, RegexOptions.IgnoreCase)).ToList();
            var pinGroup1 = new PinGroup("All_DCVI");
            pinGroup1.AddPins(group1.ToList());
            if (pinGroup1.PinList.Count != 0) pinGroups.Add(pinGroup1);

            var group2 = PinList.Where(a =>
                Regex.IsMatch(a.ChannelType, @"^DCVI", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(a.PinType, PinMapConst.TypeAnalog, RegexOptions.IgnoreCase)).ToList();
            var pinGroup2 = new PinGroup("All_DCVI_" + PinMapConst.TypeAnalog);
            pinGroup2.AddPins(group2.ToList());
            if (pinGroup2.PinList.Count != 0) pinGroups.Add(pinGroup2);
            return pinGroups;
        }

        public void GenDcvsGroup()
        {
            var group = PinList.Where(a =>
                Regex.IsMatch(a.ChannelType, @"^DCVS", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(a.PinType, PinMapConst.TypePower, RegexOptions.IgnoreCase)).ToList();
            var pinGroup = new PinGroup("All_DCVS");
            pinGroup.AddPins(group.ToList());
            if (pinGroup.PinList.Count != 0)
                AddPinGroup(pinGroup);
        }

        #endregion

        #endregion
    }

    public class SemiNumericComparer : IComparer<string>
    {
        public int Compare(string s1, string s2)
        {
            int s1R, s2R;
            var s1N = int.TryParse(s1, out s1R);
            var s2N = int.TryParse(s2, out s2R);

            if (s1N && s2N) return s1R - s2R; //two number
            if (s1N) return -1; //one number + one string
            if (s2N) return 1; //one number + one string

            var num1 = Regex.Match(s1, @"\d+$");
            var num2 = Regex.Match(s2, @"\d+$");

            var onlyString1 = s1.Remove(num1.Index, num1.Length);
            var onlyString2 = s2.Remove(num2.Index, num2.Length);

            if (onlyString1 == onlyString2) // string + number
            {
                if (num1.Success && num2.Success) return Convert.ToInt32(num1.Value) - Convert.ToInt32(num2.Value);
                if (num1.Success) return 1;
                if (num2.Success) return -1;
            }

            return string.Compare(s1, s2, StringComparison.Ordinal);
        }
    }
}