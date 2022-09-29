using IgxlData.IgxlBase;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanBinTableWriter : ScanWriter
    {
        public BinTableRows WriteBinTableRows(IEnumerable<ProdCharRow> prodCharRows)
        {
            var binTableRows = new BinTableRows();
            binTableRows.GenBlockBinTable(BlockBinTableName);
            foreach (var prodCharRowScan in prodCharRows)
                binTableRows.AddRange(GenBinTable(prodCharRowScan.PayLoadName));
            var distinct = binTableRows.Distinct(new BinTableRowComparer()).ToList();
            binTableRows = new BinTableRows(distinct);
            binTableRows.GenSetError(Block);
            return binTableRows;
        }

        private List<BinTableRow> GenBinTable(string payload)
        {
            var binTableRows = new List<BinTableRow>();
            binTableRows.Add(GenPmicScanBinRow(payload, Hnlv, new List<string> { Hv, Nv, Lv }));
            binTableRows.Add(GenPmicScanBinRow(payload, Nlv, new List<string> { Nv, Lv }));
            binTableRows.Add(GenPmicScanBinRow(payload, Hlv, new List<string> { Hv, Lv }));
            binTableRows.Add(GenPmicScanBinRow(payload, Hnv, new List<string> { Hv, Nv }));
            binTableRows.Add(GenPmicScanBinRow(payload, Hv, new List<string> { Hv }));
            binTableRows.Add(GenPmicScanBinRow(payload, Nv, new List<string> { Nv }));
            binTableRows.Add(GenPmicScanBinRow(payload, Lv, new List<string> { Lv }));

            if (LocalSpecs.HasUltraVoltageUHv)
                binTableRows.Add(GenPmicScanBinRow(payload, UHv, new List<string> { UHv }));

            if (LocalSpecs.HasUltraVoltageULv)
                binTableRows.Add(GenPmicScanBinRow(payload, ULv, new List<string> { ULv }));
            return binTableRows;
        }

        private BinTableRow GenPmicScanBinRow(string payload, string voltage, List<string> flags)
        {
            var binName = GetBinTableName(payload, voltage);
            var para = new BinNumberRuleCondition(EnumBinNumberBlock, string.Format("PMIC_{0}_{1}", Block, voltage));
            BinNumberRuleRow bin;
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out bin);
            var binRow = new BinTableRow();
            binRow.Op = "AND";
            binRow.Name = binName;
            binRow.ItemList = string.Join(",", flags.Select(x => GetFlagName(payload, x)));
            binRow.Items = Enumerable.Repeat("T", flags.Count).ToList();
            binRow.Result = "Fail";
            binRow.Sort = bin.CurrentSoftBin.ToString();
            binRow.Bin = bin.HardBin;
            return binRow;
        }
    }
}