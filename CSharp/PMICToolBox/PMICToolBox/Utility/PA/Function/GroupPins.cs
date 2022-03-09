using PmicAutomation.Utility.PA.Base;
using PmicAutomation.Utility.PA.Input;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutomation.Utility.PA.Function
{
    public class GroupPins
    {
        public void GroupPinBySameChannel(Dictionary<string, PaSheet> paSheetDic,
            out List<PaGroup> paGroups, out List<PaGroup> missingPaGroup)
        {
            missingPaGroup = new List<PaGroup>();
            paGroups = new List<PaGroup>();
            foreach (KeyValuePair<string, PaSheet> sheet in paSheetDic)
            {
                List<PaRow> paRows = new List<PaRow>();
                List<IGrouping<string, PaRow>> group =
                    sheet.Value.Rows.GroupBy(x => x.Site + "_" + x.Assignment + "_" + x.PaType).ToList();


                foreach (IGrouping<string, PaRow> rows in group)
                {
                    if (rows.First().PaType.Equals("Utility", StringComparison.CurrentCultureIgnoreCase))
                    {
                        paRows.AddRange(rows);
                        continue;
                    }

                    if (rows.Count() > 1)
                    {
                        string newPinName = GetPinName(rows.Select(x => x.BumpName).ToList());
                        if (string.IsNullOrEmpty(newPinName))
                        {
                            paRows.AddRange(rows);
                            foreach (PaRow row in rows)
                            {
                                missingPaGroup.Add(new PaGroup {GroupName = row.Assignment, PinName = row.BumpName});
                            }
                        }
                        else
                        {
                            PaRow paRow = rows.First().DeepClone();
                            paRow.BumpName = GetPinName(rows.Select(x => x.BumpName).ToList());
                            paRows.Add(paRow);
                            foreach (PaRow row in rows)
                            {
                                paGroups.Add(new PaGroup {GroupName = paRow.BumpName, PinName = row.BumpName});
                            }
                        }
                    }
                    else
                    {
                        paRows.AddRange(rows);
                    }
                }

                sheet.Value.Rows = paRows;
            }
        }

        private string GetPinName(List<string> pins)
        {
            int length = pins.First().Length;
            foreach (string pin in pins)
            {
                if (pin.Length != length)
                {
                    return "";
                }
            }

            string pinName = "";
            for (int i = 0; i < pins.First().Length; i++)
            {
                bool flag = true;
                char checkChar = pins.First()[i];
                foreach (string pin in pins)
                {
                    if (pin[i] != checkChar)
                    {
                        flag = false;
                        break;
                    }
                }

                if (flag)
                {
                    pinName += checkChar;
                }
            }

            return pinName.Trim('_').Replace("__", "_");
        }
    }
}