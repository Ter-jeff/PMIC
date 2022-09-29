using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using IgxlData.IgxlBase;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.Basic.GenDcTest.Writer
{
    internal class DcTestBinTableWriter : DcTestWriter
    {
        public DcTestBinTableWriter(string sheetName, List<HardIpPattern> patternList) : base(sheetName, patternList)
        {
        }

        public BinTableRows GenBinTableRows()
        {
            var binTableRows = new BinTableRows();
            binTableRows.GenBlockBinTable(SheetName.SheetName2Block());
            foreach (var pattern in PatternList)
            {
                binTableRows.Add(GenBinTableRow(pattern, Hnlv, new List<string> { "H", "N", "L" }));
                binTableRows.Add(GenBinTableRow(pattern, Nlv, new List<string> { "N", "L" }));
                binTableRows.Add(GenBinTableRow(pattern, Hlv, new List<string> { "H", "L" }));
                binTableRows.Add(GenBinTableRow(pattern, Hnv, new List<string> { "H", "N" }));
                binTableRows.Add(GenBinTableRow(pattern, Hv, new List<string> { "H" }));
                binTableRows.Add(GenBinTableRow(pattern, Nv, new List<string> { "N" }));
                binTableRows.Add(GenBinTableRow(pattern, Lv, new List<string> { "L" }));
                if (LocalSpecs.HasUltraVoltageUHv)
                    binTableRows.Add(GenBinTableRow(pattern, UHv, new List<string> { "UH" }));
                if (LocalSpecs.HasUltraVoltageULv)
                    binTableRows.Add(GenBinTableRow(pattern, ULv, new List<string> { "UL" }));
            }

            binTableRows.GenSetError(SheetName);
            return binTableRows;
        }

        private BinTableRow GenBinTableRow(HardIpPattern pattern, string voltage, List<string> flagVoltages)
        {
            var binRow = new BinTableRow();
            binRow.Op = "AND";
            binRow.Name = CreateBinTableName(pattern, voltage);
            var itemList = new List<string>();
            foreach (var flagVoltage in flagVoltages)
                itemList.Add(CreateFailFlag(pattern, flagVoltage));
            binRow.ItemList = string.Join(",", itemList);
            binRow.Items = Enumerable.Repeat("T", flagVoltages.Count).ToList();
            binRow.Result = "Fail";
            var binLib = GetBin(pattern);
            binRow.Sort = binLib.CurrentSoftBin.ToString("G");
            binRow.Bin = binLib.HardBin;
            return binRow;
        }

        private BinNumberRuleRow GetBin(HardIpPattern pattern)
        {
            const string index5 = "HardIP_others";
            BinNumberRuleRow binRange;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.HardIp, index5);

            para.Condition = Regex.Replace(pattern.SheetName, "HardIP_|Wireless_|LCD_", "", RegexOptions.IgnoreCase)
                .Replace("_", "");
            var found = BinNumberSingleton.Instance().GetBinNumDefRow(para, out binRange);
            if (found)
                return binRange;
            para.Condition = index5;
            found = BinNumberSingleton.Instance().GetBinNumDefRow(para, out binRange);
            if (found)
                return binRange;
            const string errorMessage = "Missing bin number setting";
            ErrorManager.AddError(EnumErrorType.MissingBinNum, EnumErrorLevel.Error, pattern.SheetName,
                pattern.RowNum, errorMessage, para.Condition);
            return binRange;
        }
    }
}