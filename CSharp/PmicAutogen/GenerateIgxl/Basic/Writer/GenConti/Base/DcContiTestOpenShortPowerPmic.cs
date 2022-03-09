using System.Collections.Generic;
using System.Data;
using IgxlData.IgxlBase;
using PmicAutogen.Inputs.TestPlan.Reader.DcTestConti;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenConti.Base
{
    public class DcContiTestOpenShortPowerPmic : DcContiTestBase
    {
        public DcContiTestOpenShortPowerPmic(DcTestContiRow row, DataTable relayTable, string subBlock = "") : base(row,
            relayTable, subBlock)
        {
        }

        public override List<FlowRow> GenerateFlowRows(RelayStatus lastStatus)
        {
            var flowRows = new List<FlowRow>();

            //Relay 
            if (!Relay.IsEqualStatus(lastStatus))
                flowRows.Add(CreateRelayRow("nop"));

            //Power
            var powerTestName = CreatePowerContinuitySerialTestName();
            var powerRow = CreateTestFlowRow(powerTestName, "");
            powerRow.Enable = "Power";
            flowRows.Add(powerRow);

            return flowRows;
        }

        public override List<InstanceRow> GenerateInstanceRows()
        {
            //Power_Continuity_Parallel
            var resultInstanceRows = new List<InstanceRow>();
            var row = new InstanceRow();
            row.TestName = CreatePowerContinuityParallelTestName();
            row.Name = DcContiConst.VbtPowerContinuityParallel;
            row.Type = "VBT";
            row.DcCategory = "Conti";
            row.DcSelector = "Typ";
            row.AcCategory = "Common";
            row.AcSelector = "Typ";
            row.TimeSets = !string.IsNullOrEmpty(TimeSet) ? TimeSet : "";
            row.PinLevels = "Levels_Conti";
            var vbt = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtPowerContinuitySerial);
            vbt.SetParamValue("analog_pins", TestPinGroup);
            vbt.SetParamValue("force_i", ForceCondition.ForceValue);
            vbt.SetParamValue("HiLimit_Short", TestLimits[0].HiLimitShort);
            vbt.SetParamValue("LowLimit_Short", TestLimits[0].LoLimitShort);
            vbt.SetParamValue("HiLimit_Open", TestLimits[0].HiLimitOpen);
            vbt.SetParamValue("LowLimit_Open", TestLimits[0].LoLimitOpen);
            row.ArgList = vbt.Parameters;
            row.Args = vbt.Args;
            resultInstanceRows.Add(row);

            //Power_Continuity_Serial
            row = new InstanceRow();
            row.TestName = CreatePowerContinuitySerialTestName();
            row.Name = DcContiConst.VbtPowerContinuitySerial;
            row.Type = "VBT";
            row.DcCategory = "Conti";
            row.DcSelector = "Typ";
            row.AcCategory = "Common";
            row.AcSelector = "Typ";
            row.TimeSets = !string.IsNullOrEmpty(TimeSet) ? TimeSet : "";
            row.PinLevels = "Levels_Conti";


            vbt = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtPowerContinuitySerial);
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

        protected string CreatePowerContinuitySerialTestName()
        {
            return GetTypeSubString("Power_Continuity_Serial_" + TestPinGroup);
        }

        protected string CreatePowerContinuityParallelTestName()
        {
            return GetTypeSubString("Power_Continuity_Parallel_" + TestPinGroup);
        }
    }
}