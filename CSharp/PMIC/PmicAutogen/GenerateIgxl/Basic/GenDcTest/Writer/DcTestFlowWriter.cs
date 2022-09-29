using CommonLib.Enum;
using CommonLib.Extension;
using CommonLib.WriteMessage;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Basic.GenDcTest.Writer
{
    internal class DcTestFlowWriter : DcTestWriter
    {
        public DcTestFlowWriter(string sheetName, List<HardIpPattern> patternList) : base(sheetName, patternList)
        {
        }

        public SubFlowSheet GenFlowSheet(string sheetName)
        {
            var subFlowSheet = new SubFlowSheet(sheetName);
            subFlowSheet.FlowRows.AddStartRows(subFlowSheet.SheetName, SubFlowSheet.Ttime);

            var voltages = new List<string>();
            voltages.Add("NV");
            voltages.Add("LV");
            voltages.Add("HV");

            if (LocalSpecs.HasUltraVoltageULv)
                voltages.Add("ULV");

            if (LocalSpecs.HasUltraVoltageUHv)
                voltages.Add("UHV");

            foreach (var pattern in PatternList)
                subFlowSheet.FlowRows.AddRange(GenFlowRowsByPattern(pattern, voltages));

            subFlowSheet.FlowRows.AddEndRows(subFlowSheet.SheetName, SubFlowSheet.Ttime, false);
            return subFlowSheet;
        }

        private FlowRows GenFlowRowsByPattern(HardIpPattern pattern, List<string> voltages)
        {
            var flowRows = new FlowRows();
            foreach (var voltage in voltages)
                try
                {
                    flowRows.AddRange(GenFlowTestRowsByVoltage(pattern, voltage));
                }
                catch (Exception e)
                {
                    Response.Report(e.ToString(), EnumMessageLevel.Error, 0);
                    throw new Exception("Error in Pattern : " + pattern.Pattern + " in RowNum: " + pattern.RowNum);
                }

            flowRows.Add_A_Enable_MP_SBIN(SheetName.SheetName2Block());
            flowRows.AddRange(WriteBinTableForPattern(pattern));

            return flowRows;
        }

        private FlowRows GenFlowTestRowsByVoltage(HardIpPattern pattern, string voltage)
        {
            var testRows = new FlowRows();
            var testRow = new FlowRow();
            testRow.Job = "";
            testRow.OpCode = CreateTestOpCode(pattern);
            testRow.Env = CreateTestEnv(pattern);
            testRow.Parameter = CreateTestName(pattern, voltage);
            testRow.Comment1 = pattern.SubBlockName;
            testRow.FailAction = CreateFailFlag(pattern, voltage).AddBlockFlag(SheetName.SheetName2Block());
            testRows.Add(testRow);
            return testRows;
        }

        private FlowRows WriteBinTableForPattern(HardIpPattern pattern)
        {
            var flowRows = new FlowRows();
            var voltages = new List<string> { Hnlv, Nlv, Hlv, Hnv, Hv, Nv, Lv };
            foreach (var voltage in voltages)
                flowRows.Add(GenBinTableRow(pattern, voltage));

            if (LocalSpecs.HasUltraVoltageUHv)
                flowRows.Add(GenBinTableRow(pattern, "UHV"));
            if (LocalSpecs.HasUltraVoltageULv)
                flowRows.Add(GenBinTableRow(pattern, "ULV"));
            return flowRows;
        }

        private string CreateTestOpCode(HardIpPattern pattern)
        {
            if (pattern.UseDeferLimit)
                return FlowRow.OpCodeTestDeferLimit;
            return FlowRow.OpCodeTest;
        }

        private string CreateTestEnv(HardIpPattern pattern)
        {
            return GetEnvFromPattern(pattern);
        }

        private FlowRow GenBinTableRow(HardIpPattern pattern, string voltage = "")
        {
            var rowBin = new FlowRow();
            rowBin.OpCode = FlowRow.OpCodeBinTable;
            rowBin.Parameter = CreateBinTableName(pattern, voltage);
            rowBin.Enable = "BinTable";
            return rowBin;
        }
    }
}