using PmicAutomation.MyControls;
using PmicAutomation.Utility.PA.Base;
using PmicAutomation.Utility.PA.Input;
using IgxlData.IgxlBase;
using Library.Function.ErrorReport;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlSheets;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.PA.Function
{
    public class GenChannelMap : PaBase
    {
        public GenChannelMap(string device, UflexConfig uflexConfig, MyForm.RichTextBoxAppend append)
            : base(device, uflexConfig, append)
        {
        }

        public ChannelMapSheet GetChannelMapSheet(List<PaRow> paRows, List<string> hexPogoList = null,
            string subName = "", PaSheet dgsReferenceSheet = null)
        {
            SiteCnt = paRows.GroupBy(a => a.BumpName + "_" + a.PaType)
                .Max(g => g.Select(x => x.Site).Distinct().Count());

            ChannelMapSheet channelMap =
                new ChannelMapSheet("ChannelMap_" + SiteCnt + "_site" + subName + ".txt") { SiteNum = SiteCnt };

            List<ChannelMapRow> allChannels = GenChannelMapRows(paRows, hexPogoList);

            //GenAllDgsChannel(paRows, allChannels);

            channelMap.ChannelMapRows = allChannels.OrderBy(x => x.DiviceUnderTestPinName).ThenBy(y => y.Type).ToList();

            List<ChannelMapRow> allDgsChannels = this.GenAllDgsChannels(paRows, dgsReferenceSheet);
            //Dgs channel map items should be at the bottom.
            channelMap.ChannelMapRows.AddRange(allDgsChannels);

            ModifyTheSameChannel(channelMap.ChannelMapRows);

            return channelMap;
        }

        private List<ChannelMapRow> GenChannelMapRows(List<PaRow> paRows, List<string> hexPogos)
        {
            List<IGrouping<string, PaRow>> siteGroups = paRows.GroupBy(a => a.BumpName + "_" + a.PaType).ToList();
            IEnumerable<IGrouping<string, PaRow>> errorSiteCnt = siteGroups.Where(x => x.Count() != SiteCnt);
            foreach (IGrouping<string, PaRow> rows in errorSiteCnt)
            {
                string errMsg = "The count of pin " + rows.Key + " is not " + SiteCnt;
                ErrorManager.AddError(PaErrorType.Duplicated, rows.First().SourceSheetName, rows.First().RowNum,
                    errMsg);
            }

            List<ChannelMapRow> allChannels = new List<ChannelMapRow>();
            foreach (IGrouping<string, PaRow> siteGroup in siteGroups)
            {
                ChannelMapRow channelRow = new ChannelMapRow();
                List<PaRow> pinList = siteGroup.ToList();
                PaRow paItem = pinList[0];

                channelRow.DiviceUnderTestPinName = GetPinName(paRows, paItem);
                channelRow.Type = paItem.PaType.Equals("CBIT", StringComparison.CurrentCulture) ? "I/O" : paItem.PaType;
                List<string> sites = paRows.Select(x => x.Site).Distinct().ToList();
                foreach (string site in sites)
                {
                    PaRow pin = pinList.FirstOrDefault(a => a.Site == site.ToString());
                    if (pin != null && Regex.IsMatch(pin.Assignment,
                            @"\.ch|\.sense|\.util|.SrcPos|.SrcNeg|.cappos|.capneg|\.dgs|\.GigaDig|\.drv|\.rcv",
                            RegexOptions.IgnoreCase))
                    {
                        bool isHex = hexPogos != null && hexPogos.Contains(pin.Assignment.Split('.')[0]);
                        string channel;
                        int mergeNum = 0;

                        if (Regex.IsMatch(pin.PaType, @"Merged", RegexOptions.IgnoreCase))
                        {
                            mergeNum = Convert.ToInt32(Regex.Match(pin.PaType, @"(?<num>\d+)").Groups["num"]
                                .ToString());
                        }

                        if (Regex.IsMatch(pin.Assignment, @"\.GigaDig", RegexOptions.IgnoreCase))
                        {
                            int assNum = Convert.ToInt32(Regex.Match(pin.Assignment, @"\w+(?<assNum>\d+)")
                                .Groups["assNum"].ToString());
                            int slotNum = Convert.ToInt32(Regex.Match(pin.Assignment, @"(?<slotNum>\d+)\.\w+")
                                .Groups["slotNum"].ToString());
                            channelRow.Type = Regex.Match(pin.Assignment, @"\.(?<chName>\w+)\d+").Groups["chName"]
                                .ToString();
                            pin.Assignment = slotNum + ".cappos" + assNum;
                        }
                        else if (Regex.IsMatch(pin.Assignment, @"\.drv|\.rcv", RegexOptions.IgnoreCase))
                        {
                            int assNum = Convert.ToInt32(Regex.Match(pin.Assignment, @"\w+(?<assNum>\d+)")
                                .Groups["assNum"].ToString());
                            int slotNum = Convert.ToInt32(Regex.Match(pin.Assignment, @"(?<slotNum>\d+)\.\w+")
                                .Groups["slotNum"].ToString());
                            channelRow.Type = "Serial10G";
                            pin.Assignment = slotNum + "." + Regex.Match(pin.Assignment, @"\.(?<chName>\w+)\d+").Groups["chName"] + assNum;
                        }

                        PaRow siteInfo = pinList.FirstOrDefault(x =>
                            x.Assignment == pin.Assignment && Convert.ToInt32(x.Site) < Convert.ToInt32(pin.Site));
                        if (siteInfo != null)
                        {
                            channel = "site" + siteInfo.Site; // shared site
                        }
                        else
                        {
                            if (mergeNum > 2 && !Regex.IsMatch(pin.Assignment, @"hc$", RegexOptions.IgnoreCase) &&
                                !isHex)
                            {
                                channel = pin.Assignment + "hc";
                            }
                            else
                            {
                                channel = pin.Assignment;
                            }
                        }

                        var tokens = channel.Split(new char[] { '.' });
                        if (tokens.Count() > 2)
                        {
                            channel = tokens[0] + "." + tokens[1];
                        }
                        channelRow.Sites.Add(channel);
                    }
                    else
                    {
                        channelRow.Sites.Add("");
                    }
                }

                allChannels.Add(channelRow);

                //Note: If device pin is DCVI pin, Autogen tool will automatically add [device_pin_name]_UVI80_DM and [device_pin_name]_UVI80_DT
                if (Device.Equals("PMIC", StringComparison.OrdinalIgnoreCase) &&
                    paItem.PaType.Equals("DCVI", StringComparison.OrdinalIgnoreCase))
                {
                    allChannels.AddRange(GenChannelMapSecondary(channelRow));
                }
            }

            return allChannels;
        }


        private List<ChannelMapRow> GenAllDgsChannels(List<PaRow> paRows, PaSheet dgsReferenceSheet)
        {
            List<ChannelMapRow> allDgsChannels = new List<ChannelMapRow>();
            if (!Device.Equals(Pmic, StringComparison.OrdinalIgnoreCase))
            {
                return allDgsChannels;
            }
            if (dgsReferenceSheet != null)
            {
                DgsPool dgsPool = new DgsPool(dgsReferenceSheet, SiteCnt);
                System.Text.StringBuilder errorMessage = new System.Text.StringBuilder();
                bool result = dgsPool.FormatPrecheck(ref errorMessage);
                if (!result)
                {
                    this.GenAllDgsChannel(paRows, allDgsChannels);
                }
                else
                {
                    allDgsChannels.AddRange(dgsPool.GenAllDgsChannelMap());
                }
            }
            else
            {
                this.GenAllDgsChannel(paRows, allDgsChannels);
            }
            return allDgsChannels;
        }


        private void GenAllDgsChannel(List<PaRow> paRows, List<ChannelMapRow> channelMapRows)
        {
            //if (Device.Equals(Pmic, StringComparison.OrdinalIgnoreCase))
            //{
            channelMapRows.AddRange(GetDgsChannel(paRows, "DC30", 4));
            channelMapRows.AddRange(GetDgsChannel(paRows, "UVI80", 8));
            //}
        }

        private List<ChannelMapRow> GetDgsChannel(List<PaRow> paRows, string type, int pinNum)
        {
            List<ChannelMapRow> channelList = new List<ChannelMapRow>();
            List<string> channelNumList = new List<string>();
            bool flag = true;
            foreach (IGrouping<string, PaRow> siteList in paRows.GroupBy(x => x.Site))
            {
                List<string> paList = siteList.Where(x => x.InstrumentType == type)
                    .Select(y => GetChannel(y.Assignment)).ToList();
                if (paList.Distinct().Count() != 1)
                {
                    flag = false;
                    break;
                }

                channelNumList.Add(paList.Distinct().First());
            }

            if (flag)
            {
                for (int i = 1; i <= pinNum; i++)
                {
                    ChannelMapRow channelRow = new ChannelMapRow();
                    if (channelNumList.Distinct().Count() == 1)
                    {
                        string channel = channelNumList[0] + ".dgs" + i;
                        channelRow.Sites.Add(channel);
                        for (int j = 0; j < paRows.GroupBy(x => x.Site).Count() - 1; j++)
                        {
                            channel = "site0";
                            channelRow.Sites.Add(channel);
                        }
                    }
                    else if (channelNumList.Count == SiteCnt)
                    {
                        for (int j = 0; j < paRows.GroupBy(x => x.Site).Count(); j++)
                        {
                            string channel = channelNumList[j] + ".dgs" + i;
                            channelRow.Sites.Add(channel);
                        }
                    }
                    else
                    {
                        Append(@"The slot of PMIC Dgs pin should the same or equal site count", Color.Red);
                    }

                    channelRow.DiviceUnderTestPinName = type + "_DGS_" + i + "_DM";
                    channelRow.Type = ConDcDiffMeter;
                    channelList.Add(channelRow);
                }
            }
            return channelList;
        }

        private List<ChannelMapRow> GenChannelMapSecondary(ChannelMapRow channelMapRow)
        {
            List<ChannelMapRow> newChannelMapRow = new List<ChannelMapRow>();
            string toolType = GetToolTypeByConfig(channelMapRow.Sites.First());
            if (string.IsNullOrEmpty(toolType))
            {
                return newChannelMapRow;
            }

            if (!string.IsNullOrEmpty(toolType) &&
                IsPinNameContainsToolType(channelMapRow.DiviceUnderTestPinName, toolType))
            {
                channelMapRow.DiviceUnderTestPinName = channelMapRow.DiviceUnderTestPinName;
            }
            else
            {
                channelMapRow.DiviceUnderTestPinName += "_" + toolType;
            }

            ChannelMapRow newChDm = new ChannelMapRow
            {
                DiviceUnderTestPinName = channelMapRow.DiviceUnderTestPinName + ConDm,
                Type = ConDcDiffMeter
            };
            newChDm.Sites.AddRange(channelMapRow.Sites);
            newChannelMapRow.Add(newChDm);

            ChannelMapRow newChDt = new ChannelMapRow
            {
                DiviceUnderTestPinName = channelMapRow.DiviceUnderTestPinName + ConDt,
                Type = ConDcTime
            };
            newChDt.Sites.AddRange(channelMapRow.Sites);
            newChannelMapRow.Add(newChDt);
            return newChannelMapRow;
        }

        private void ModifyTheSameChannel(List<ChannelMapRow> channelMapRows)
        {
            for (int i = 0; i < channelMapRows.Count; i++)
                for (int j = 0; j < channelMapRows[i].Sites.Count; j++)
                {
                    if (i != 0)
                    {
                        if (!string.IsNullOrEmpty(channelMapRows[i].Sites[j])
                            && Regex.IsMatch(channelMapRows[i].Sites[j],
                                @"\.ch|\.sense|\.util|.SrcPos|.SrcNeg|.cappos|.capneg", RegexOptions.IgnoreCase))
                        {
                            List<ChannelMapRow> find = channelMapRows.GetRange(0, i).ToList();
                            ChannelMapRow findList = find.Find(x =>
                                x.Sites[j] == channelMapRows[i].Sites[j] && x.Type == channelMapRows[i].Type);
                            if (findList != null)
                            {
                                channelMapRows[i].Sites[j] = "S:" + findList.DiviceUnderTestPinName;
                            }
                        }
                    }
                }
        }

        private string GetToolTypeByConfig(string channelAssignment)
        {
            return UflexConfig.GetToolType(GetChannel(channelAssignment));
        }
    }
}