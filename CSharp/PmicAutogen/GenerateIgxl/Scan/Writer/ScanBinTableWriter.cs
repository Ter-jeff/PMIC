using System;
using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Local;
using AutomationCommon.Utility;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanBinTableWriter
    {
        protected const string Hnlv = "HNLV";
        protected const string Nlv = "NLV";
        protected const string Hlv = "HLV";
        protected const string Hnv = "HNV";
        protected const string Lv = "LV";
        protected const string Nv = "NV";
        protected const string Hv = "HV";
        protected const string ULv = "ULV";
        protected const string UHv = "UHV";

        protected string Block;

        public ScanBinTableWriter()
        {
            Block = "SCAN";
        }

        public virtual List<BinTableRow> WriteBinTable(List<ProdCharRowScan> prodCharRowScans)
        {
            var binTableRows = new List<BinTableRow>();
            foreach (var prodCharRowScan in prodCharRowScans)
                GenerateBinTable(binTableRows, prodCharRowScan.PayLoadName);
            return binTableRows;
        }

        protected void GenerateBinTable(List<BinTableRow> binTableRows, string payload)
        {
            GeneratePmicScanBinRow(binTableRows, payload, Hnlv, new List<string> {Hv, Nv, Lv});
            GeneratePmicScanBinRow(binTableRows, payload, Nlv, new List<string> {Nv, Lv});
            GeneratePmicScanBinRow(binTableRows, payload, Hlv, new List<string> {Hv, Lv});
            GeneratePmicScanBinRow(binTableRows, payload, Hnv, new List<string> {Hv, Nv});
            GeneratePmicScanBinRow(binTableRows, payload, Hv, new List<string> {Hv});
            GeneratePmicScanBinRow(binTableRows, payload, Nv, new List<string> {Nv});
            GeneratePmicScanBinRow(binTableRows, payload, Lv, new List<string> {Lv});
            if (LocalSpecs.HasUltraVoltageUHv)
            {
                GeneratePmicScanBinRow(binTableRows, payload, UHv, new List<string> { UHv });
            }

            if (LocalSpecs.HasUltraVoltageULv)
            {
                GeneratePmicScanBinRow(binTableRows, payload, ULv, new List<string> { ULv });
            }
        }

        private void GeneratePmicScanBinRow(List<BinTableRow> binTableRows, string payload, string voltage,
            List<string> flagList)
        {
            var binName = string.Format("Bin_{0}_{1}_{2}", Block, payload.GetSortPatNameForBinTable(), voltage);
            if (binTableRows.FindIndex(p => p.Name.Equals(binName)) != -1) return;
            BinNumberRuleCondition para = null;
            if (Block.Equals("Scan", StringComparison.CurrentCultureIgnoreCase))
                para = new BinNumberRuleCondition(EnumBinNumberBlock.Scan,
                    string.Format("PMIC_{0}_{1}", Block, voltage));
            else if (Block.Equals("MBIST", StringComparison.CurrentCultureIgnoreCase))
                para = new BinNumberRuleCondition(EnumBinNumberBlock.Mbist,
                    string.Format("PMIC_{0}_{1}", Block, voltage));
            BinNumberRuleRow bin;
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out bin);
            var binRow = new BinTableRow();
            binRow.Op = "AND";
            binRow.Name = binName;
            binRow.ItemList = string.Join(",", flagList.Select(x => string.Format("F_{0}_{1}", payload.GetSortPatNameForBinTable(), x)));
            binRow.Items = Enumerable.Repeat("T", flagList.Count).ToList();
            binRow.Result = "Fail";
            binRow.Sort = bin.CurrentSoftBin.ToString();
            binRow.Bin = bin.HardBin;
            binTableRows.Add(binRow);
        }
    }
}