using IgxlData.IgxlBase;
using PmicAutogen.Inputs.TestPlan.Reader.DcTestConti;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.Basic.GenConti.Base
{
    public abstract class DcContiTestBase
    {
        protected DcContiTestBase(DcTestContiRow row, DataTable relayTable, string subBlock)
        {
            TestLimits = new List<DcTestLimit>();
            TimeSet = row.TimeSet;
            SubBlock = subBlock;
            TestPinGroup = row.PinGroup;
            string condition;
            row.GetForceCondition(out condition);
            ForceCondition = row.TestType == ContiType.OpenShort
                ? new DcForceCondition("I", condition)
                : new DcForceCondition("V", condition);

            string hiLimitShort, loLimitShort, hiLimitOpen, loLimitOpen;
            row.GetTestLimit(out hiLimitShort, out loLimitShort, out hiLimitOpen, out loLimitOpen);
            TestLimits.Add(new DcTestLimit("25C", hiLimitShort, loLimitShort, hiLimitOpen, loLimitOpen));
            RelayTable = relayTable;
            Relay = GetRelayStatus();
        }

        protected DataTable RelayTable { set; get; }
        protected string TimeSet { set; get; }
        public RelayStatus Relay { set; get; }
        public List<DcTestLimit> TestLimits { set; get; }
        public DcForceCondition ForceCondition { set; get; }
        public string TestPinGroup { set; get; }
        public string SubBlock { get; set; }

        public InstanceRow GenerateRelayInstanceRow(RelayStatus lastStatus)
        {
            var relayInstance = new InstanceRow();
            relayInstance.TestName = CreateRelayTestName();
            relayInstance.Name = DcContiConst.VbtFuncNameRelayControl;
            relayInstance.Type = "VBT";
            var vbt = TestProgram.VbtFunctionLib.GetFunctionByName(DcContiConst.VbtFuncNameRelayControl);
            vbt.Args[0] = string.Join(",", Relay.OpenRelayList);
            vbt.Args[1] = string.Join(",", Relay.OffRelayList);
            vbt.Args[2] = DcContiConst.RelayWaitTime;

            relayInstance.ArgList = vbt.Parameters;
            relayInstance.Args = vbt.Args;
            return relayInstance;
        }

        public abstract FlowRows GenerateFlowRows(RelayStatus lastStatus);

        public abstract List<InstanceRow> GenerateInstanceRows();

        protected RelayStatus GetRelayStatus()
        {
            for (var i = 0; i < RelayTable.Rows.Count; i++)
            {
                var relayControlName = RelayTable.Rows[i][1].ToString();
                if (relayControlName.Equals(TestPinGroup, StringComparison.OrdinalIgnoreCase))
                {
                    var status = new RelayStatus();
                    for (var j = 2; j < RelayTable.Columns.Count; j++)
                    {
                        var value = RelayTable.Rows[i][j].ToString();
                        if (value.Equals("1"))
                            status.OpenRelayList.Add(RelayTable.Columns[j].ColumnName);
                        else
                            status.OffRelayList.Add(RelayTable.Columns[j].ColumnName);
                    }

                    return status;
                }
            }

            return new RelayStatus();
        }

        protected virtual string CreateContiTestName()
        {
            string testName;

            //OpenShort
            if (Regex.IsMatch(TestPinGroup, "^Continuity", RegexOptions.IgnoreCase))
            {
                if (!string.IsNullOrEmpty(SubBlock))
                    testName = "DC_" + SubBlock + "_" + TestPinGroup;
                else
                    testName = "DC_" + TestPinGroup;
            }
            else if (Regex.IsMatch(TestPinGroup, "^DC_Continuity", RegexOptions.IgnoreCase))
            {
                if (!string.IsNullOrEmpty(SubBlock))
                    testName = SubBlock + "_" + TestPinGroup;
                else
                    testName = TestPinGroup;
            }
            else
            {
                if (!string.IsNullOrEmpty(SubBlock))
                    testName = "DC_Continuity_" + SubBlock + "_" + TestPinGroup;
                else
                    testName = "DC_Continuity_" + TestPinGroup;
            }

            if (Regex.IsMatch(ForceCondition.ForceValue, @"[-].*"))
            {
                //Negative
                if (!Regex.IsMatch(testName, "_Neg$", RegexOptions.IgnoreCase)) testName = testName + "_Neg";
            }
            else
            {
                //Positive
                if (!Regex.IsMatch(testName, "_Pos$", RegexOptions.IgnoreCase)) testName = testName + "_Pos";
            }

            return testName;
        }

        protected string GetTypeSubString(string name)
        {
            if (Regex.IsMatch(ForceCondition.ForceValue, @"[-].*"))
            {
                if (!Regex.IsMatch(name, "_Neg$", RegexOptions.IgnoreCase))
                    name = name + "_Neg";
            }
            else
            {
                if (!Regex.IsMatch(name, "_Pos$", RegexOptions.IgnoreCase))
                    name = name + "_Pos";
            }

            return name;
        }

        protected string CreateWalkingZTestName()
        {
            return "WalkingZ_" + CreateContiTestName();
        }

        public string CreateWalkingZFlagName()
        {
            return "F_WalkingZ_" + CreateContiTestName();
        }

        protected string CreateRelayTestName()
        {
            return "Relay_On_" + CreateContiTestName();
        }

        protected FlowRow CreateTestFlowRow(string testName, string flagName)
        {
            var row = new FlowRow();
            row.OpCode = FlowRow.OpCodeTest;
            row.Parameter = testName;
            row.FailAction = flagName;
            return row;
        }

        protected FlowRow CreateRelayRow(string opCode = "")
        {
            var relayRow = new FlowRow();
            relayRow.OpCode = string.IsNullOrEmpty(opCode) ? FlowRow.OpCodeTest : opCode;
            relayRow.Parameter = CreateRelayTestName();
            return relayRow;
        }
    }
}