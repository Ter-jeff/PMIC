using System.Collections.Generic;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTable;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local.Const;
using System;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenBinTable
{
    public class DcTestBinTableGenerator : HardIpBinTableGenerator
    {
        private const string Hnlv = "HNLV";
        private const string Nlv = "NLV";
        private const string Hlv = "HLV";
        private const string Hnv = "HNV";
        private const string Lv = "LV";
        private const string Nv = "NV";
        private const string Hv = "HV";
        private const string ULv = "ULV";
        private const string UHv = "UHV";

        private string SheetName;

        public DcTestBinTableGenerator(BinTableSheet hardIpBinTableSheet, string sheetName,
            List<HardIpPattern> patternList, List<string> duplicateParameter, List<string> errorBinNumbers) : base(
            hardIpBinTableSheet, sheetName, patternList, duplicateParameter, errorBinNumbers)
        {
            SheetName = sheetName;
            BinTableRowGenerator = new DcTestBinTableRowGenerator(sheetName, errorBinNumbers);
        }

        public override void GenerateBinTableRows()
        {
            foreach (var pattern in PatternList)
            {
                if (SheetName.Equals(PmicConst.FlowDcLeakage.Replace("Flow_", ""), StringComparison.OrdinalIgnoreCase))
                    HardIpBinTableSheet.AddRow(
                   BinTableRowGenerator.GeneratePmicBinRow(pattern, "", new List<string> {""}));
                else
                {
                    HardIpBinTableSheet.AddRow(
                        BinTableRowGenerator.GeneratePmicBinRow(pattern, Hnlv, new List<string> { "H", "N", "L" }));
                    HardIpBinTableSheet.AddRow(
                        BinTableRowGenerator.GeneratePmicBinRow(pattern, Nlv, new List<string> { "N", "L" }));
                    HardIpBinTableSheet.AddRow(
                        BinTableRowGenerator.GeneratePmicBinRow(pattern, Hlv, new List<string> { "H", "L" }));
                    HardIpBinTableSheet.AddRow(
                        BinTableRowGenerator.GeneratePmicBinRow(pattern, Hnv, new List<string> { "H", "N" }));
                    HardIpBinTableSheet.AddRow(
                        BinTableRowGenerator.GeneratePmicBinRow(pattern, Lv, new List<string> { "H" }));
                    HardIpBinTableSheet.AddRow(
                        BinTableRowGenerator.GeneratePmicBinRow(pattern, Nv, new List<string> { "N" }));
                    HardIpBinTableSheet.AddRow(
                        BinTableRowGenerator.GeneratePmicBinRow(pattern, Hv, new List<string> { "L" }));
                    if (LocalSpecs.HasUltraVoltageUHv)
                    {
                        HardIpBinTableSheet.AddRow(
                            BinTableRowGenerator.GeneratePmicBinRow(pattern, UHv, new List<string> { "UH" }));
                    }
                    if (LocalSpecs.HasUltraVoltageULv)
                    {
                        HardIpBinTableSheet.AddRow(
                            BinTableRowGenerator.GeneratePmicBinRow(pattern, ULv, new List<string> { "UL" }));
                    }
                }
            }
        }
    }
}