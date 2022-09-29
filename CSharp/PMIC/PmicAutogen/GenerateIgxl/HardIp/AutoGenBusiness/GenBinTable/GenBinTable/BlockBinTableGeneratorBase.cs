using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTableRow;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTable
{
    public abstract class BlockBinTableGeneratorBase
    {
        protected BinTableRowGeneratorBase BinTableRowGenerator = null;
        protected List<string> DuplicateParameter;
        protected BinTableSheet HardIpBinTableSheet;
        protected List<HardIpPattern> PatternList;

        protected BlockBinTableGeneratorBase(BinTableSheet hardIpBinTableSheet, List<HardIpPattern> patternList,
            List<string> duplicateParameter)
        {
            HardIpBinTableSheet = hardIpBinTableSheet;
            PatternList = patternList;
            DuplicateParameter = duplicateParameter;
        }

        public virtual void GenerateBinTableRows()
        {
            foreach (var pattern in PatternList)
                if (!Regex.IsMatch(pattern.Pattern.GetLastPayload(), @"^cz_", RegexOptions.IgnoreCase) ||
                    HardIpDataMain.TestPlanData.CzBinOut)
                {
                    var binTableRow = BinTableRowGenerator.GenBinTableRow(pattern);
                    var binTableList = new BinTableRows {binTableRow};
                    if (!Regex.IsMatch(pattern.Pattern.GetLastPayload(), "instance:") &&
                        !Regex.IsMatch(pattern.SheetName, "DCTEST_IDS", RegexOptions.IgnoreCase) &&
                        !pattern.SheetName.Equals(PmicConst.PmicLeakage, StringComparison.OrdinalIgnoreCase))
                        binTableList = SplitBinTableRowByVoltage(binTableRow);
                    if (!binTableList.Exists(p => DuplicateParameter.Contains(p.Name)))
                        foreach (var item in binTableList)
                            if (item != null)
                            {
                                HardIpBinTableSheet.AddRow(item);
                                DuplicateParameter.Add(item.Name);
                            }
                }
        }

        private BinTableRows SplitBinTableRowByVoltage(BinTableRow binTableRow)
        {
            var volList = new List<string> {"HLV", "HV", "LV", "NV"};
            var result = new BinTableRows();

            foreach (var item in binTableRow.ExtraBinDictionary)
            {
                var newBinTableRow = binTableRow.CopyBinTableRow();
                newBinTableRow.Op = "AND";
                newBinTableRow.Name += "_" + item.Key;
                newBinTableRow.ItemList = SelectFailFlagByVoltage(newBinTableRow.ItemList, item.Key);
                newBinTableRow.Items = UpdateItemsWithItemList(newBinTableRow.ItemList);
                newBinTableRow.Sort += item.Value;
                result.Add(newBinTableRow);
                volList.Remove(item.Key);
            }

            volList.Remove("HLV");
            if (volList.Count > 0)
            {
                var flags = new List<string>();
                foreach (var vol in volList) flags.Add(SelectFailFlagByVoltage(binTableRow.ItemList, vol));
                binTableRow.ItemList = string.Join(",", flags);
                binTableRow.Items = UpdateItemsWithItemList(binTableRow.ItemList);
                result.Add(binTableRow);
            }

            return result;
        }

        private string SelectFailFlagByVoltage(string failFlag, string voltage)
        {
            var flagList = failFlag.Split(',').ToList();
            var result = new List<string>();
            switch (voltage)
            {
                case HardIpConstData.LabelHLv:
                    result.AddRange(flagList.Where(p => Regex.IsMatch(p,
                            "_[" + HardIpConstData.LabelHLv.Substring(0, 2) + "]" +
                            HardIpConstData.SuffixHardIpFailAction))
                        .ToList());
                    break;
                case HardIpConstData.LabelHv:
                    result.AddRange(flagList.Where(p => Regex.IsMatch(p,
                        "_[" + HardIpConstData.LabelHv[0] + "]" + HardIpConstData.SuffixHardIpFailAction)).ToList());
                    break;
                case HardIpConstData.LabelLv:
                    result.AddRange(flagList.Where(p => Regex.IsMatch(p,
                        "_[" + HardIpConstData.LabelLv[0] + "]" + HardIpConstData.SuffixHardIpFailAction)).ToList());
                    break;
                case HardIpConstData.LabelNv:
                    result.AddRange(flagList.Where(p => Regex.IsMatch(p,
                        "_[" + HardIpConstData.LabelNv[0] + "]" + HardIpConstData.SuffixHardIpFailAction)).ToList());
                    break;
                default:
                    result.Add(failFlag);
                    break;
            }

            return string.Join(",", result);
        }

        private List<string> UpdateItemsWithItemList(string itemList)
        {
            return Enumerable.Repeat("T", itemList.Split(',').Length).ToList();
        }
    }
}