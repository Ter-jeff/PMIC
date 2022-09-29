using IgxlData.IgxlBase;
using IgxlData.VBT;
using PmicAutogen.Inputs.TestPlan.Reader.DcTestConti;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Data;

namespace PmicAutogen.GenerateIgxl.Basic.GenConti.Base
{
    public class DcContiTestOpenShortIoPmic : DcContiTestBase
    {
        public DcContiTestOpenShortIoPmic(DcTestContiRow row, DataTable relayTable, string subBlock = "") : base(row,
            relayTable, subBlock)
        {
        }

        public override FlowRows GenerateFlowRows(RelayStatus lastStatus)
        {
            var flowRows = new FlowRows();

            //Relay 
            if (!Relay.IsEqualStatus(lastStatus))
                flowRows.Add(CreateRelayRow("nop"));

            //AutoZ
            //if(_genAutoZOnly)
            //    GenAutoZFlowRows(flowRows);

            //WalkingZ
            var walkingZTestName = CreateWalkingZTestName();
            var walkingZRow = CreateTestFlowRow(walkingZTestName, "");
            walkingZRow.Env = "X";
            walkingZRow.Enable = "WalkingZ";
            walkingZRow.FailAction = DcContiConst.FlagNameAlarmFail;
            flowRows.Add(walkingZRow);

            //Serial
            var testName = CreateIoContinuitySerialTestName();
            var serialRow = CreateTestFlowRow(testName, "");
            serialRow.FailAction = DcContiConst.FlagNameAlarmFail;
            flowRows.Add(serialRow);

            return flowRows;
        }

        public override List<InstanceRow> GenerateInstanceRows()
        {
            var resultInstanceRows = new List<InstanceRow>();

            InstanceRow row;
            VbtFunctionBase vbt;

            //AutoZ Only
            //if (_genAutoZOnly)
            //{
            //    row = new InstanceRow();
            //    row.TestName = DcContiConst.InsNameAutoZOnly;
            //    row.Name = DcContiConst.VbtIoContinuitySerial;
            //    row.Type = "VBT";
            //    row.DcCategory = "Conti";
            //    row.DcSelector = "Typ";
            //    row.AcCategory = "Common";
            //    row.AcSelector = "Typ";
            //    row.PinLevels = "Levels_Conti";
            //    row.TimeSets = !string.IsNullOrEmpty(TimeSet) ? TimeSet : "";

            //    vbt = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtIoContinuitySerial);
            //    vbt.SetParamValue("digital_pins", TestPinGroup);
            //    vbt.SetParamValue("force_i", ForceCondition.ForceValue);
            //    vbt.SetParamValue("HiLimit_Short", TestLimits[0].HiLimitShort);
            //    vbt.SetParamValue("LowLimit_Short", TestLimits[0].LoLimitShort);
            //    vbt.SetParamValue("HiLimit_Open", TestLimits[0].HiLimitOpen);
            //    vbt.SetParamValue("LowLimit_Open", TestLimits[0].LoLimitOpen);
            //    row.ArgList = vbt.Parameters;
            //    row.Args = vbt.Args;
            //    resultInstanceRows.Add(row);
            //}

            //WalkingZ
            row = new InstanceRow();
            row.TestName = CreateWalkingZTestName();
            row.Name = DcContiConst.VbtContiWalkingZ;
            row.Type = "VBT";
            row.DcCategory = "Conti";
            row.DcSelector = "Typ";
            row.AcCategory = "Common";
            row.AcSelector = "Typ";
            row.TimeSets = !string.IsNullOrEmpty(TimeSet) ? TimeSet : "";
            row.PinLevels = ForceCondition.ForceValue.Contains("-") ? "Levels_WalkingZ_Neg" : "Levels_WalkingZ_Pos";

            vbt = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameFunctionalT);
            vbt.SetParamValue("Patterns", @".\PATTERN\WalkingZ\Continuity_Neg_WalkingZ.PAT");
            row.ArgList = vbt.Parameters;
            row.Args = vbt.Args;
            resultInstanceRows.Add(row);

            //IO_Continuity_Serial
            row = new InstanceRow();
            row.TestName = CreateIoContinuitySerialTestName();
            row.Name = DcContiConst.VbtIoContinuitySerial;
            row.Type = "VBT";
            row.DcCategory = "Conti";
            row.DcSelector = "Typ";
            row.AcCategory = "Common";
            row.AcSelector = "Typ";
            row.PinLevels = "Levels_Conti";
            row.TimeSets = !string.IsNullOrEmpty(TimeSet) ? TimeSet : "";

            vbt = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtIoContinuitySerial);
            vbt.SetParamValue("digital_pins", TestPinGroup);
            vbt.SetParamValue("force_i", ForceCondition.ForceValue);
            vbt.SetParamValue("HiLimit_Short", TestLimits[0].HiLimitShort);
            vbt.SetParamValue("LowLimit_Short", TestLimits[0].LoLimitShort);
            vbt.SetParamValue("HiLimit_Open", TestLimits[0].HiLimitOpen);
            vbt.SetParamValue("LowLimit_Open", TestLimits[0].LoLimitOpen);
            row.ArgList = vbt.Parameters;
            row.Args = vbt.Args;
            resultInstanceRows.Add(row);

            //IO_Continuity_Parallel
            //(digital_pins As PinList, force_i As Double, TestLimitMode As tlLimitForceResults, 
            //Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double, _
            //Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short", Optional connect_all_pins As PinList)
            row = new InstanceRow();
            row.TestName = CreateIoContinuityParallelTestName();
            row.Name = DcContiConst.VbtIoContinuityParallel;
            row.Type = "VBT";
            row.DcCategory = "Conti";
            row.DcSelector = "Typ";
            row.AcCategory = "Common";
            row.AcSelector = "Typ";
            row.PinLevels = "Levels_Conti";
            row.TimeSets = !string.IsNullOrEmpty(TimeSet) ? TimeSet : "";

            vbt = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtIoContinuitySerial);
            vbt.SetParamValue("digital_pins", TestPinGroup);
            vbt.SetParamValue("force_i", ForceCondition.ForceValue);
            vbt.SetParamValue("HiLimit_Short", TestLimits[0].HiLimitShort);
            vbt.SetParamValue("LowLimit_Short", TestLimits[0].LoLimitShort);
            vbt.SetParamValue("HiLimit_Open", TestLimits[0].HiLimitOpen);
            vbt.SetParamValue("LowLimit_Open", TestLimits[0].LoLimitOpen);
            row.ArgList = vbt.Parameters;
            row.Args = vbt.Args;
            resultInstanceRows.Add(row);

            return resultInstanceRows;
        }

        protected string CreateIoContinuitySerialTestName()
        {
            return GetTypeSubString("IO_Continuity_Serial_" + TestPinGroup);
        }

        protected string CreateIoContinuityParallelTestName()
        {
            return GetTypeSubString("IO_Continuity_Parallel_" + TestPinGroup);
        }
    }
}