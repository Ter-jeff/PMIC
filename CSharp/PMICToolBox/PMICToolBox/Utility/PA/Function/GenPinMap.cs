using PmicAutomation.MyControls;
using PmicAutomation.Utility.PA.Base;
using PmicAutomation.Utility.PA.Input;
using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlSheets;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.PA.Function
{
    public class GenPinMap : PaBase
    {
        public GenPinMap(string device, UflexConfig uflexConfig, MyForm.RichTextBoxAppend append)
            : base(device, uflexConfig, append)
        {
        }

        public PinMapSheet GetPinMapSheet(Dictionary<string, PaSheet> paSheets, PaSheet dgsReferenceSheet = null)
        {
            List<PaRow> paRows = GetDistinctPaRows(paSheets.SelectMany(x => x.Value.Rows).ToList());

            PinMapSheet pinMap = new PinMapSheet("PinMap_PA.txt");
            if (paSheets.Count == 0)
                return pinMap;
            SiteCnt = paRows.GroupBy(a => a.BumpName + "_" + a.PaType).Max(g => g.Select(x => x.Site).Distinct().Count());

            List<Pin> allPins = GetPinMap(paRows);

            allPins = allPins.OrderBy(x => x.PinName).ThenBy(x => x.PinType).ToList();

            //All Dgs pins should be at the bottom.
            if (Device.Equals("PMIC", StringComparison.OrdinalIgnoreCase))
                GenAllDgsPin(paRows, allPins, dgsReferenceSheet);

            foreach (Pin pin in allPins)
                pinMap.AddRow(pin);

            #region Gen pin grup
            GenPinGroup(pinMap);

            GenDcviGroup(pinMap);

            if (Device.Equals("PMIC", StringComparison.OrdinalIgnoreCase))
                GenDiffGroup(pinMap);
            #endregion

            return pinMap;
        }

        private void GenPinGroup(PinMapSheet pinMapSheet)
        {
            IEnumerable<IGrouping<string, Pin>> groups = pinMapSheet.PinList.GroupBy(x => x.InstrumentType);
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
                    pinGroup.AddPins(group2.Select(x => x.PinName).ToList());
                    if (pinGroup.PinList.Count != 0)
                    {
                        pinMapSheet.AddRow(pinGroup);
                    }
                }
            }
        }

        private void GenDiffGroup(PinMapSheet pinMapSheet)
        {
            for (int i = 0; i < pinMapSheet.PinList.Count; i++)
            {
                for (int j = i + 1; j < pinMapSheet.PinList.Count; j++)
                {
                    if (pinMapSheet.PinList[i].PinType
                        .Equals(pinMapSheet.PinList[j].PinType, StringComparison.CurrentCulture))
                    {
                        string diffName =
                            GetDiffPinName(pinMapSheet.PinList[i].PinName, pinMapSheet.PinList[j].PinName);
                        if (!string.IsNullOrEmpty(diffName))
                        {
                            PinGroup pinGroup = new PinGroup(diffName + "_Diff", "Differential");
                            pinGroup.AddPin(pinMapSheet.PinList[i].PinName, "Differential");
                            pinGroup.AddPin(pinMapSheet.PinList[j].PinName, "Differential");
                            pinMapSheet.AddRow(pinGroup);
                        }
                    }
                }
            }
        }

        private string GetDiffPinName(string pin1, string pin2)
        {
            string diffName = "";
            List<string> remains = new List<string>();
            if (pin1.Length != pin2.Length)
            {
                return "";
            }

            for (int i = 0; i < pin1.Length; i++)
            {
                if (pin1[i] == pin2[i])
                {
                    diffName += pin1[i];
                }
                else
                {
                    remains.Add(pin1[i].ToString());
                    remains.Add(pin2[i].ToString());
                }
            }

            if (remains.Count != 2)
            {
                return "";
            }

            if (remains[0].Equals("P", StringComparison.OrdinalIgnoreCase) &&
                remains[1].Equals("N", StringComparison.OrdinalIgnoreCase))
            {
                return diffName;
            }

            if (remains[0].Equals("N", StringComparison.OrdinalIgnoreCase) &&
                remains[1].Equals("P", StringComparison.OrdinalIgnoreCase))
            {
                return diffName;
            }

            return "";
        }

        private void GenDcviGroup(PinMapSheet pinMapSheet)
        {
            List<Pin> group1 = pinMapSheet.PinList.Where(a =>
                Regex.IsMatch(a.ChannelType, @"^DCVI", RegexOptions.IgnoreCase)
                && Regex.IsMatch(a.PinType, PinMapConst.TypePower, RegexOptions.IgnoreCase)).ToList();
            PinGroup pinGroup1 = new PinGroup(Dcvi, PinMapConst.TypePower);
            pinGroup1.AddPins(group1.Select(x => x.PinName).ToList());
            if (pinGroup1.PinList.Count != 0)
            {
                pinMapSheet.AddRow(pinGroup1);
            }

            List<Pin> group2 = pinMapSheet.PinList.Where(a =>
                Regex.IsMatch(a.ChannelType, @"^DCVI", RegexOptions.IgnoreCase)
                && Regex.IsMatch(a.PinType, PinMapConst.TypeAnalog, RegexOptions.IgnoreCase)).ToList();
            PinGroup pinGroup2 = new PinGroup(Dcvi + "_" + PinMapConst.TypeAnalog, PinMapConst.TypeAnalog);
            pinGroup2.AddPins(group2.Select(x => x.PinName).ToList());
            if (pinGroup2.PinList.Count != 0)
            {
                pinMapSheet.AddRow(pinGroup2);
            }
        }

        private List<Pin> GetPinMap(List<PaRow> paRows)
        {
            List<Pin> allPins = new List<Pin>();
            foreach (PaRow paItem in paRows)
            {
                string pinName = GetPinName(paRows, paItem);
                string pinType = GetPinMapType(paItem);
                paItem.PinMapType = pinType;
                Pin pin = new Pin(pinName, pinType)
                {
                    InstrumentType = paItem.InstrumentType,
                    ChannelType = paItem.PaType
                };
                if (!allPins.Exists(x => x.PinName.Equals(pin.PinName)))
                {
                    allPins.Add(pin);
                }

                //If device pin is DCVI pin, tool will automatically add [device_pin_name]_UVI80_DM & [device_pin_name]_UVI80_DT
                if (Device.Equals(Pmic, StringComparison.OrdinalIgnoreCase) &&
                    paItem.PaType.Equals("DCVI", StringComparison.OrdinalIgnoreCase))
                {
                    allPins.AddRange(GenPinMapSecondary(pin, paItem));
                }
            }

            return allPins;
        }

        private void GenAllDgsPin(List<PaRow> paRows, List<Pin> pins, PaSheet dgsReferenceSheet = null)
        {
            if (Device.Equals(Pmic, StringComparison.OrdinalIgnoreCase))
            {
                if (dgsReferenceSheet == null)
                {
                    pins.AddRange(GenDgsPin(paRows, "DC30", 4));
                    pins.AddRange(GenDgsPin(paRows, "UVI80", 8));
                }
                else
                {
                    DgsPool dgsPool = new DgsPool(dgsReferenceSheet, SiteCnt);
                    System.Text.StringBuilder sb = new System.Text.StringBuilder();
                    bool result = dgsPool.FormatPrecheck(ref sb);
                    if (!result)
                    {
                        pins.AddRange(GenDgsPin(paRows, "DC30", 4));
                        pins.AddRange(GenDgsPin(paRows, "UVI80", 8));
                    }
                    else
                    {
                        int DC30Cnt = dgsPool.GetDgsPinCount("DC30");
                        int UVI80Cnt = dgsPool.GetDgsPinCount("UVI80");
                        pins.AddRange(GenDgsPin(paRows, "DC30", DC30Cnt < 4 ? 4 : DC30Cnt));
                        pins.AddRange(GenDgsPin(paRows, "UVI80", UVI80Cnt < 8 ? 8 : UVI80Cnt));
                    }
                }
            }
        }

        private List<Pin> GenDgsPin(List<PaRow> paRows, string type, int pinNum)
        {
            List<Pin> pinList = new List<Pin>();
            bool flag = true;
            foreach (IGrouping<string, PaRow> siteList in paRows.GroupBy(x => x.Site))
            {
                IEnumerable<string> paList =
                    siteList.Where(x => x.InstrumentType == type).Select(y => GetChannel(y.Assignment));
                if (paList.Distinct().Count() != 1)
                {
                    flag = false;
                    break;
                }
            }

            if (flag)
            {
                for (int i = 1; i <= pinNum; i++)
                {
                    Pin pin = new Pin(type + "_DGS_" + i + "_DM", "Analog");
                    pinList.Add(pin);
                }
            }

            return pinList;
        }

        private List<Pin> GenPinMapSecondary(Pin pin, PaRow paRow)
        {
            List<Pin> newPinList = new List<Pin>();
            string instrumentType = paRow.InstrumentType;
            if (string.IsNullOrEmpty(instrumentType))
            {
                return newPinList;
            }

            if (!string.IsNullOrEmpty(instrumentType) && IsPinNameContainsToolType(pin.PinName, instrumentType))
            {
                pin.PinName = pin.PinName;
            }
            else
            {
                pin.PinName += "_" + instrumentType;
            }

            Pin dmPin = new Pin(pin.PinName + ConDm, "Analog");
            newPinList.Add(dmPin);
            Pin tmPin = new Pin(pin.PinName + ConDt, "Analog");
            newPinList.Add(tmPin);
            return newPinList;
        }

        private string GetPinMapType(PaRow paRow)
        {
            string pinType;
            if (Regex.IsMatch(paRow.PaType, @"Util", RegexOptions.IgnoreCase))
            {
                if (paRow.InstrumentType.Equals("I/O", StringComparison.OrdinalIgnoreCase))
                {
                    pinType = "I/O";
                }
                if (paRow.InstrumentType.Equals("CBIT", StringComparison.OrdinalIgnoreCase))
                {
                    pinType = "I/O";
                }
                else if (paRow.InstrumentType.Equals("Support", StringComparison.OrdinalIgnoreCase))
                {
                    pinType = "Utility";
                }
                else
                {
                    pinType = "";
                    Append("The channel " + paRow.Assignment + " of Utility pin " + paRow.BumpName + " is mismatch", Color.Red);
                }
            }
            else if (Regex.IsMatch(paRow.Ps, @"power", RegexOptions.IgnoreCase) ||
                     Regex.IsMatch(paRow.PaType, @"^DCVS", RegexOptions.IgnoreCase))
            {
                pinType = "power";
            }
            else if (Regex.IsMatch(paRow.PaType,
                @"^DCVI|^UltraCapture|^UltraSource|^MW|^GigaDigNeg|^GigaDigPos|^DCDiffMeter", RegexOptions.IgnoreCase))
            {
                pinType = "Analog";
            }
            else if (Regex.IsMatch(paRow.PaType, @"Gnd", RegexOptions.IgnoreCase))
            {
                pinType = "Gnd";
            }
            else if (Regex.IsMatch(paRow.PaType, @"N/C", RegexOptions.IgnoreCase))
            {
                pinType = "N/C";
            }
            else
            {
                pinType = "I/O";
            }

            // temp special case @20190131
            if (Regex.IsMatch(paRow.Assignment, @"\.GigaDig", RegexOptions.IgnoreCase))
            {
                pinType = "Analog";
            }
            else if (Regex.IsMatch(paRow.Assignment, @"\.drv|\.rcv", RegexOptions.IgnoreCase))
            {
                pinType = Regex.IsMatch(paRow.BumpName, @"RX", RegexOptions.IgnoreCase) ? "Input" : "Output";
            }

            return pinType;
        }

        private List<PaRow> GetDistinctPaRows(List<PaRow> paRows)
        {
            List<PaRow> distinctPaRows = new List<PaRow>();
            foreach (PaRow row in paRows)
            {
                if (!distinctPaRows.Exists(p => p.BumpName.Equals(row.BumpName, StringComparison.OrdinalIgnoreCase) &&
                                                p.Site.Equals(row.Site, StringComparison.OrdinalIgnoreCase) &&
                                                p.PaType.Equals(row.PaType, StringComparison.OrdinalIgnoreCase)))
                {
                    distinctPaRows.Add(row);
                }
            }

            return distinctPaRows;
        }
    }
}