using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlSheet
{
    public static class PinMapConst
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
        #region Constructor

        public PinMapSheet(string name)
            : base(name)
        {
        }

        #endregion

        #region Field

        #endregion

        #region Property

        public List<Pin> PinList
        {
            get;
        } = new List<Pin>();

        public List<PinGroup> GroupList
        {
            get;
        } = new List<PinGroup>();

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

        public override void Write(string file, string version)
        {
            if (version == "2.1")
            {
                using (StreamWriter sw = new StreamWriter(file, false))
                {
                    sw.WriteLine("DTPinMap,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tPin Map");
                    sw.WriteLine("\t\t\t\t\t");
                    sw.WriteLine("\tGroup Name\tPin Name\tType\tComment");

                    foreach (Pin pin in PinList)
                    {
                        sw.WriteLine("\t{0}\t{1}\t{2}\t{3}", "", pin.PinName, pin.PinType, pin.Comment);
                    }

                    foreach (PinGroup pinGroup in GroupList)
                    {
                        foreach (Pin pin in pinGroup.PinList)
                        {
                            if (string.IsNullOrEmpty(pin.PinType) && IsPinExist(pin.PinName))
                            {
                                pin.PinType = GetPin(pin.PinName).PinType;
                            }

                            sw.WriteLine("\t{0}\t{1}\t{2}\t{3}", pinGroup.PinName, pin.PinName, pin.PinType,
                                pin.Comment);
                        }
                    }
                }
            }
            else
            {
                throw new Exception($"The PinMap sheet version:{version} is not supported!");
            }
        }

        public void AddRow(PinBase pinBase)
        {
            if (pinBase is Pin)
            {
                AddPin((Pin)pinBase);
            }
            else
            {
                AddPinGroup((PinGroup)pinBase);
            }
        }

        public string GetDiffGroupName(string[] pair)
        {
            if (pair.Length == 2)
            {
                PinGroup group = GroupList.FirstOrDefault(x => x.PinList.Count == 2 &&
                                                               x.PinList.Exists(p =>
                                                                   p.PinName.Equals(pair[0],
                                                                       StringComparison.OrdinalIgnoreCase)) &&
                                                               x.PinList.Exists(p =>
                                                                   p.PinName.Equals(pair[1],
                                                                       StringComparison.OrdinalIgnoreCase)));
                if (group != null)
                {
                    return group.PinName;
                }
            }

            return "";
        }

        public bool IsChannelType(string pinName, string type)
        {
            if (IsGroupExist(pinName))
            {
                List<Pin> groupPins = GetGroup(pinName).PinList;
                bool flag = groupPins.All(x =>
                    GetChannelType(x.PinName).Equals(type, StringComparison.OrdinalIgnoreCase));
                return flag;
            }

            if (GetChannelType(pinName).Equals(type, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return false;
        }

        public string GetChannelType(string pinName)
        {
            string type = "";
            if (PinList.Exists(x => x.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                type = PinList.Find(x => x.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)).ChannelType;
            }

            return type;
        }

        public List<Pin> GetPowerPins()
        {
            List<Pin> pinList = PinList.FindAll(p =>
                p.PinType.Equals(PinMapConst.TypePower, StringComparison.OrdinalIgnoreCase));
            return pinList;
        }

        public List<Pin> GetIoPins()
        {
            List<Pin> pinList =
                PinList.FindAll(p => p.PinType.Equals(PinMapConst.TypeIo, StringComparison.OrdinalIgnoreCase));
            return pinList;
        }

        public List<Pin> GetIoContinuityPins()
        {
            List<Pin> pinList = GetAllDigitalDisconnectContinuityPins();
            pinList = pinList.Where(x => !(x.PinName.StartsWith("REFCLK_", StringComparison.OrdinalIgnoreCase) ||
                                           x.PinName.EndsWith("_PA", StringComparison.OrdinalIgnoreCase)
                )).ToList();
            return pinList;
        }

        public List<Pin> GetAllDigitalDisconnectContinuityPins()
        {
            List<Pin> pinList = PinList.FindAll(p =>
                p.ChannelType.Equals(PinMapConst.TypeIo, StringComparison.OrdinalIgnoreCase));
            pinList = pinList.Where(x =>
                !(x.PinName.StartsWith("VDD", StringComparison.OrdinalIgnoreCase) &&
                  x.PinName.ToUpper().Contains("SENSE") ||
                  x.PinName.StartsWith("VDD", StringComparison.OrdinalIgnoreCase) &&
                  x.PinName.ToUpper().Contains("MONITOR") ||
                  x.PinName.StartsWith("VSS", StringComparison.OrdinalIgnoreCase) &&
                  x.PinName.ToUpper().Contains("SENSE")
                    )).ToList();
            return pinList;
        }

        #region pin

        private void AddPin(Pin pin)
        {
            if (!IsPinExist(pin.PinName))
            {
                PinList.Add(pin);
            }
        }

        public bool IsPinExist(string pinName)
        {
            if (PinList.Exists(p => p.PinName.ToLower().Equals(pinName.ToLower())))
            {
                return true;
            }

            return false;
        }

        public Pin GetPin(string pinName)
        {
            Pin pin = PinList.Find(p => p.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase));
            return pin;
        }

        public string GetPinType(string pinName)
        {
            if (!PinList.Exists(p => p.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
            {
                return "";
            }

            string type = PinList.Find(p => p.PinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)).PinType;
            return type;
        }

        #endregion

        #region pinGroup

        public void AddPinGroup(PinGroup pinGroup)
        {
            if (!IsGroupExist(pinGroup.PinName))
            {
                if (pinGroup.PinList.Count == 0)
                {
                    throw new Exception($"There isn't any pin in PinGroup:{pinGroup.PinName} . ");
                }

                if (!Regex.IsMatch(pinGroup.PinName, @"^efuse|All_DiffPairs", RegexOptions.IgnoreCase))
                {
                    if (pinGroup.PinType != "Differential")
                    {
                        pinGroup.PinList = pinGroup.PinList.OrderBy(x => x.PinName, new SemiNumericComparer()).ToList();
                    }
                }

                GroupList.Add(pinGroup);
            }
        }

        public void AddPinGroups(List<PinGroup> pinGroups)
        {
            foreach (PinGroup pinGroup in pinGroups)
            {
                if (pinGroup.PinList == null || pinGroup.PinList.Count == 0)
                {
                    continue;
                }

                AddRow(pinGroup);
            }
        }

        public bool IsGroupExist(string groupName)
        {
            if (GroupList.Exists(p => p.PinName.ToLower().Equals(groupName.ToLower())))
            {
                return true;
            }

            return false;
        }

        public void SortPinGroup()
        {
            foreach (PinGroup pinGroup in GroupList)
            {
                if (!Regex.IsMatch(pinGroup.PinName, @"^efuse|All_DiffPairs", RegexOptions.IgnoreCase))
                {
                    if (pinGroup.PinType != "Differential")
                    {
                        pinGroup.PinList = pinGroup.PinList.OrderBy(x => x.PinName, new SemiNumericComparer()).ToList();
                    }
                }
            }
        }

        public PinGroup GetGroup(string groupName)
        {
            PinGroup group = GroupList.FirstOrDefault(p => p.PinName.ToLower().Equals(groupName.ToLower()));
            return group;
        }

        public List<Pin> GetPinsFromGroup(string groupName)
        {
            List<Pin> resultPins = new List<Pin>();
            if (!GroupList.Exists(p => p.PinName.Equals(groupName, StringComparison.OrdinalIgnoreCase)))
            {
                return new List<Pin>();
            }

            List<Pin> pinOrGroup = GroupList.Find(p => p.PinName.Equals(groupName, StringComparison.OrdinalIgnoreCase))
                .PinList;
            foreach (Pin pin in pinOrGroup)
            {
                if (IsGroupExist(pin.PinName))
                {
                    resultPins.AddRange(GetPinsFromGroup(pin.PinName));
                }
                else
                {
                    Pin singlePin =
                        PinList.Find(p => p.PinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase));
                    resultPins.Add(singlePin ?? new Pin(pin.PinName, "", "Does not exist in Pin map sheet"));
                }
            }

            return resultPins;
        }

        public List<PinGroup> GenPinGroup()
        {
            List<PinGroup> pinGroups = new List<PinGroup>();
            IEnumerable<IGrouping<string, Pin>> groups = PinList.GroupBy(x => x.InstrumentType);
            foreach (IGrouping<string, Pin> group1 in groups)
            {
                if (string.IsNullOrEmpty(group1.Key))
                {
                    continue;
                }

                foreach (IGrouping<string, Pin> group2 in group1.GroupBy(x => x.PinType))
                {
                    string groupPinName = group1.GroupBy(x => x.PinType).Count() != 1
                        ? "All_" + group1.Key.Replace("/", "") + "_" + group2.Key.Replace("/", "")
                        : "All_" + group1.Key.Replace("/", "");
                    PinGroup pinGroup = new PinGroup(groupPinName, group2.Key);
                    pinGroup.AddPins(group2.ToList());
                    if (pinGroup.PinList.Count != 0)
                    {
                        pinGroups.Add(pinGroup);
                    }
                }
            }

            return pinGroups;
        }

        public List<PinGroup> GenDcviGroup()
        {
            List<PinGroup> pinGroups = new List<PinGroup>();
            List<Pin> group1 = PinList.Where(a =>
                Regex.IsMatch(a.ChannelType, @"^DCVI", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(a.PinType, PinMapConst.TypePower, RegexOptions.IgnoreCase)).ToList();
            PinGroup pinGroup1 = new PinGroup("All_DCVI", PinMapConst.TypePower);
            pinGroup1.AddPins(group1.ToList());
            if (pinGroup1.PinList.Count != 0)
            {
                pinGroups.Add(pinGroup1);
            }

            List<Pin> group2 = PinList.Where(a =>
                Regex.IsMatch(a.ChannelType, @"^DCVI", RegexOptions.IgnoreCase) &&
                Regex.IsMatch(a.PinType, PinMapConst.TypeAnalog, RegexOptions.IgnoreCase)).ToList();
            PinGroup pinGroup2 = new PinGroup("All_DCVI_" + PinMapConst.TypeAnalog, PinMapConst.TypeAnalog);
            pinGroup2.AddPins(group2.ToList());
            if (pinGroup2.PinList.Count != 0)
            {
                pinGroups.Add(pinGroup2);
            }

            return pinGroups;
        }

        #endregion

        #endregion
    }

    public class SemiNumericComparer : IComparer<string>
    {
        public int Compare(string s1, string s2)
        {
            int s1R;
            int s2R;
            bool s1N = int.TryParse(s1, out s1R);
            bool s2N = int.TryParse(s2, out s2R);

            if (s1N && s2N)
            {
                return s1R - s2R; //two number
            }

            if (s1N)
            {
                return -1; //one number + one string
            }

            if (s2N)
            {
                return 1; //one number + one string
            }

            if (s1 != null && s2 != null)
            {
                Match num1 = Regex.Match(s1, @"\d+$");
                Match num2 = Regex.Match(s2, @"\d+$");

                string onlyString1 = s1.Remove(num1.Index, num1.Length);
                string onlyString2 = s2.Remove(num2.Index, num2.Length);

                if (onlyString1 == onlyString2) // string + number
                {
                    if (num1.Success && num2.Success)
                    {
                        return Convert.ToInt32(num1.Value) - Convert.ToInt32(num2.Value);
                    }

                    if (num1.Success)
                    {
                        return 1;
                    }

                    if (num2.Success)
                    {
                        return -1;
                    }
                }
            }

            return string.Compare(s1, s2, StringComparison.Ordinal);
        }
    }
}