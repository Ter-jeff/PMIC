using IgxlData.IgxlBase;
using PmicAutogen.Inputs.TestPlan.Reader.DcTestConti;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Data;

namespace PmicAutogen.GenerateIgxl.Basic.GenConti.Base
{
    public class DcContiTestOpenShortAnalogPmic : DcContiTestBase
    {
        public DcContiTestOpenShortAnalogPmic(DcTestContiRow row, DataTable relayTable, string subBlock = "") : base(
            row, relayTable, subBlock)
        {
        }

        public override FlowRows GenerateFlowRows(RelayStatus lastStatus)
        {
            var flowRows = new FlowRows();
            //Relay 
            if (!Relay.IsEqualStatus(lastStatus))
                flowRows.Add(CreateRelayRow("nop"));

            //Analog
            var powerTestName = CreateAnalogContinuitySerialTestName();
            var powerRow = CreateTestFlowRow(powerTestName, "");
            powerRow.Enable = "Analog";
            powerRow.FailAction = DcContiConst.FlagNameAlarmFail;
            flowRows.Add(powerRow);

            return flowRows;
        }

        public override List<InstanceRow> GenerateInstanceRows()
        {
            var resultInstanceRows = new List<InstanceRow>();
            var row = new InstanceRow();
            //Analog_Continuity_Parallel
            row.TestName = CreateAnalogContinuityParallelTestName();
            row.Name = DcContiConst.VbtAnalogContinuityParallel;
            row.Type = "VBT";
            row.DcCategory = "Conti";
            row.DcSelector = "Typ";
            row.AcCategory = "Common";
            row.AcSelector = "Typ";
            row.TimeSets = !string.IsNullOrEmpty(TimeSet) ? TimeSet : "";
            row.PinLevels = "Levels_Conti";
            var vbt = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtAnalogContinuitySerial);
            vbt.SetParamValue("analog_pins", TestPinGroup);
            vbt.SetParamValue("force_i", ForceCondition.ForceValue);
            vbt.SetParamValue("HiLimit_Short", TestLimits[0].HiLimitShort);
            vbt.SetParamValue("LowLimit_Short", TestLimits[0].LoLimitShort);
            vbt.SetParamValue("HiLimit_Open", TestLimits[0].HiLimitOpen);
            vbt.SetParamValue("LowLimit_Open", TestLimits[0].LoLimitOpen);
            row.ArgList = vbt.Parameters;
            row.Args = vbt.Args;
            resultInstanceRows.Add(row);

            //Analog_Continuity_Serial
            row = new InstanceRow();
            row.TestName = CreateAnalogContinuitySerialTestName();
            row.Name = DcContiConst.VbtAnalogContinuitySerial;
            row.Type = "VBT";
            row.DcCategory = "Conti";
            row.DcSelector = "Typ";
            row.AcCategory = "Common";
            row.AcSelector = "Typ";
            row.TimeSets = !string.IsNullOrEmpty(TimeSet) ? TimeSet : "";
            row.PinLevels = "Levels_Conti";

            vbt = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtAnalogContinuitySerial);
            vbt.SetParamValue("analog_pins", TestPinGroup);
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

        protected string CreateAnalogContinuitySerialTestName()
        {
            return GetTypeSubString("Analog_Continuity_Serial_" + TestPinGroup);
        }

        protected string CreateAnalogContinuityParallelTestName()
        {
            return GetTypeSubString("Analog_Continuity_Parallel_" + TestPinGroup);
        }
    }
}