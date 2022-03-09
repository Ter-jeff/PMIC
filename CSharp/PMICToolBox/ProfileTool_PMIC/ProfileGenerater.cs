using CommonLib.EpplusErrorReport;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.VBT;
using ProfileTool_PMIC.Output;
using ProfileTool_PMIC.Reader;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace ProfileTool_PMIC
{
    public class ProfileGenerater
    {
        private const int MaxTime = 10;//Sec
        private readonly ProfileToolForm _profileToolForm;
        private readonly BasManager _basMain;

        public ProfileGenerater(ProfileToolForm profileToolForm, BasManager basMain)
        {
            _profileToolForm = profileToolForm;
            _basMain = basMain;
        }

        public SubFlowSheet WorkByFlow(SubFlowSheet sheet, List<ExecutionProfileRow> executionProfileRows, List<string> pins, ref InstanceSheet instanceSheet)
        {
            var profileRows = executionProfileRows;//.Where(x => x.Flowsheet.Equals(sheet.Name, StringComparison.CurrentCultureIgnoreCase)).ToList();
            var name = Regex.Replace(sheet.Name, "^Flow_", "", RegexOptions.IgnoreCase);
            var indexHaeder = profileRows.FindIndex(x => x.Flowstep.StartsWith(name + "_Header", StringComparison.CurrentCultureIgnoreCase));
            var indexFooter = indexHaeder != -1 ? profileRows.FindIndex(x => x.Flowstep.StartsWith(name + "_Footer", StringComparison.CurrentCultureIgnoreCase)) : -1;
            if ((indexHaeder != -1 && indexFooter != -1))
            {
                var rows = profileRows.GetRange(indexHaeder, indexFooter - indexHaeder);
                var testTimeFromExecProfile = rows.Sum(x => x.Total);

                if (testTimeFromExecProfile < MaxTime)
                {
                    if (_profileToolForm.checkBoxCurrent.Checked)
                    {
                        const string type = "Current";
                        const string whatToCapture = "I";
                        GenByFullFlow(sheet, instanceSheet, pins, type, whatToCapture, testTimeFromExecProfile);
                    }
                    if (_profileToolForm.checkBoxVoltage.Checked)
                    {
                        const string type = "Voltage";
                        const string whatToCapture = "V";
                        GenByFullFlow(sheet, instanceSheet, pins, type, whatToCapture, testTimeFromExecProfile);
                    }
                }
                else
                {
                    if (_profileToolForm.checkBoxCurrent.Checked)
                    {
                        const string type = "Current";
                        const string whatToCapture = "I";
                        GenByPartialFlow(sheet, instanceSheet, pins, type, whatToCapture, rows);
                    }
                    if (_profileToolForm.checkBoxVoltage.Checked)
                    {
                        const string type = "Voltage";
                        const string whatToCapture = "V";
                        GenByPartialFlow(sheet, instanceSheet, pins, type, whatToCapture, rows);
                    }
                }
            }
            else
            {
                //MessageBox.Show(string.Format("{0} don't have test time From ExecProfile !!!", sheet.Name), @"ERROR");
                _profileToolForm.AppendText(string.Format("{0} don't have test time From ExecProfile !!!", sheet.Name), Color.Red);
                EpplusErrorManager.AddError(BasicErrorType.Business.ToString(), sheet.Name, 1, 1, string.Format("{0} don't have test time From ExecProfile !!!", sheet.Name));
            }
            return sheet;
        }

        private void GenByPartialFlow(SubFlowSheet sheet, InstanceSheet instanceSheet, List<string> pins, string type, string whatToCapture, List<ExecutionProfileRow> rows)
        {
            var block = Regex.Replace(sheet.Name, "Flow_", "", RegexOptions.IgnoreCase);
            double testTimeFromExecProfile;
            var tempFlowRows = new List<Tuple<int, FlowRow>>();
            //Step1. Get indexList to match the test instance
            var indexDic = GetIndexDic(rows, sheet.FlowRows);

            //Step2 Get start and end index, and make sure the last row in executionProfileRows is existed in the flow sheet
            var groupCnt = 1;
            var startIndexExecutionProfile = 0;
            var endIndexExecutionProfile = 0;
            var startIndexFlow = 0;

            for (var i = 0; i < indexDic.Count(); i++)
            {
                if (indexDic[i].Item2 != -1)
                {
                    var expectedTestTimeFromExecProfile = indexDic.GetRange(startIndexExecutionProfile, i - startIndexExecutionProfile + 1).Sum(x => x.Item1.Total);
                    if (expectedTestTimeFromExecProfile > MaxTime)
                    {
                        testTimeFromExecProfile = indexDic.GetRange(startIndexExecutionProfile, endIndexExecutionProfile - startIndexExecutionProfile + 1).Sum(x => x.Item1.Total);
                        var enable = type + "Profile";
                        var profileStart = GenProfileFlowRow(block + "_" + type + "_Profile_Start_" + groupCnt, enable);
                        tempFlowRows.Add(new Tuple<int, FlowRow>(startIndexFlow, profileStart));
                        var profilePlot = GenProfileFlowRow(block + "_" + type + "_Profile_Plot_" + groupCnt, enable);
                        var indexAfterLimit = MoveAfterLimit(sheet, indexDic[endIndexExecutionProfile].Item2);
                        tempFlowRows.Add(new Tuple<int, FlowRow>(indexAfterLimit, profilePlot));
                        instanceSheet.AddRow(GetInstanceRowStart(true, whatToCapture, profileStart, pins, testTimeFromExecProfile, "{0:F6}"));
                        instanceSheet.AddRow(GetInstanceRowPlot(profilePlot, pins));

                        groupCnt++;
                        startIndexFlow = indexAfterLimit;
                        startIndexExecutionProfile = endIndexExecutionProfile + 1;
                    }
                    endIndexExecutionProfile = i;
                }
            }

            //Step3 Get final instance
            testTimeFromExecProfile = indexDic.GetRange(startIndexExecutionProfile, indexDic.Count - startIndexExecutionProfile).Sum(x => x.Item1.Total);
            var enable1 = type + "Profile";
            var finalprofileStart = GenProfileFlowRow(block + "_" + type + "_Profile_Start_" + groupCnt, enable1);
            tempFlowRows.Add(new Tuple<int, FlowRow>(startIndexFlow, finalprofileStart));
            var finalprofilePlot = GenProfileFlowRow(block + "_" + type + "_Profile_Plot_" + groupCnt, enable1);
            var index = sheet.FlowRows.FindIndex(x => x.Opcode.Equals("return", StringComparison.CurrentCultureIgnoreCase));
            var stopIndexFlow = index == -1 ? sheet.FlowRows.Count : index;
            tempFlowRows.Add(new Tuple<int, FlowRow>(stopIndexFlow, finalprofilePlot));

            for (var i = tempFlowRows.Count - 1; i > -1; i--)
            {
                var item = tempFlowRows.ElementAt(i);
                sheet.InsertRow(item.Item1, item.Item2);
            }

            instanceSheet.AddRow(GetInstanceRowStart(true, whatToCapture, finalprofileStart, pins, testTimeFromExecProfile, "{0:F6}"));
            instanceSheet.AddRow(GetInstanceRowPlot(finalprofilePlot, pins));
        }

        private List<Tuple<ExecutionProfileRow, int>> GetIndexDic(List<ExecutionProfileRow> rows, List<FlowRow> flowRows)
        {
            var dic = new List<Tuple<ExecutionProfileRow, int>>();
            var current = 0;
            for (var i = 0; i < rows.Count; i++)
            {
                var index = -1;
                for (var j = current; j < flowRows.Count; j++)
                {
                    if (rows[i].Flowstep.Equals(flowRows[j].Parameter, StringComparison.CurrentCultureIgnoreCase))
                    {
                        index = j;
                        current = j + 1;
                        break;
                    }
                }
                var item = new Tuple<ExecutionProfileRow, int>(rows[i], index);
                dic.Add(item);
            }
            return dic;
        }

        private void GenByFullFlow(SubFlowSheet sheet, InstanceSheet instanceSheet, List<string> pins, string type, string whatToCapture, double testTimeFromExecProfile)
        {
            var block = Regex.Replace(sheet.Name, "Flow_", "", RegexOptions.IgnoreCase);
            var enable = type + "Profile";
            var profileStart = GenProfileFlowRow(block + "_" + type + "_Profile_Start", enable);
            sheet.InsertRow(0, profileStart);
            var profilePlot = GenProfileFlowRow(block + "_" + type + "_Profile_Plot", enable);
            var index = sheet.FlowRows.FindIndex(x => x.Opcode.Equals("return", StringComparison.CurrentCultureIgnoreCase));
            sheet.InsertRow(index == -1 ? sheet.FlowRows.Count : index, profilePlot);
            instanceSheet.AddRow(GetInstanceRowStart(true, whatToCapture, profileStart, pins, testTimeFromExecProfile, "{0:F2}"));
            instanceSheet.AddRow(GetInstanceRowPlot(profilePlot, pins));
        }

        private int MoveAfterLimit(SubFlowSheet sheet, int indexFlow)
        {
            var cnt = 1;
            while (indexFlow + cnt < sheet.FlowRows.Count &&
                (sheet.FlowRows[indexFlow + cnt].Opcode.Equals("Use-Limit", StringComparison.CurrentCultureIgnoreCase) ||
                sheet.FlowRows[indexFlow + cnt].Opcode.Equals("Characterize", StringComparison.CurrentCultureIgnoreCase) ||
                sheet.FlowRows[indexFlow + cnt].Opcode.Equals("Test-defer-limits", StringComparison.CurrentCultureIgnoreCase)))
            {
                cnt++;
            }
            return indexFlow + cnt;
        }

        public SubFlowSheet WorkByInstance(SubFlowSheet sheet, ref InstanceSheet instanceSheet, List<string> pins, List<ExecutionProfileRow> executionProfileIgRows, List<InstanceRow> excludingInstanaceRows, List<PinInfoRow> pinInfoRows)
        {
            for (var index = sheet.FlowRows.Count - 1; index > 0; index--)
            {
                var row = sheet.FlowRows[index];
                if (!(row.Opcode.Equals("Test", StringComparison.CurrentCultureIgnoreCase)))
                    continue;
                if (!excludingInstanaceRows.Exists(x => x.TestName.Equals(row.Parameter, StringComparison.CurrentCultureIgnoreCase)))
                    continue;
                if (_profileToolForm.checkBox_Excluding_By_Job.Checked && !string.IsNullOrEmpty(row.Job))
                    continue;

                if (executionProfileIgRows.Exists(x => x.Flowstep.Equals(row.Parameter, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var testTimeFromExecProfile = executionProfileIgRows.Find(x => x.Flowstep.Equals(row.Parameter, StringComparison.CurrentCultureIgnoreCase)).Total;

                    if (_profileToolForm.checkBoxCurrent.Checked)
                    {
                        const string enable = "CurrentProfile";
                        var profilePlot = GenProfileFlowRow(row.Parameter + "_Current_Profile_Plot", enable);
                        var indexAfterLimit = MoveAfterLimit(sheet, index);
                        sheet.InsertRow(indexAfterLimit, profilePlot);
                        var instanceRowPlot = GetInstanceRowPlot(profilePlot, pins, row, pinInfoRows);
                        if (!instanceSheet.InstanceRows.Exists(x => x.TestName.Equals(instanceRowPlot.TestName, StringComparison.CurrentCultureIgnoreCase)))
                            instanceSheet.AddRow(instanceRowPlot);

                        var profileStart = GenProfileFlowRow(row.Parameter + "_Current_Profile_Start", enable);
                        sheet.InsertRow(index, profileStart);

                        var instanceRowStart = GetInstanceRowStart(false, "I", profileStart, pins, testTimeFromExecProfile, "{0:F6}", row, pinInfoRows);
                        if (!instanceSheet.InstanceRows.Exists(x => x.TestName.Equals(instanceRowStart.TestName, StringComparison.CurrentCultureIgnoreCase)))
                            instanceSheet.AddRow(instanceRowStart);
                    }

                    if (_profileToolForm.checkBoxVoltage.Checked)
                    {
                        const string enable = "VoltageProfile";
                        var profilePlot = GenProfileFlowRow(row.Parameter + "_Voltage_Profile_Plot", enable);
                        var indexAfterLimit = MoveAfterLimit(sheet, index);
                        sheet.InsertRow(indexAfterLimit, profilePlot);
                        var instanceRowPlot = GetInstanceRowPlot(profilePlot, pins, row, pinInfoRows);
                        if (!instanceSheet.InstanceRows.Exists(x => x.TestName.Equals(instanceRowPlot.TestName, StringComparison.CurrentCultureIgnoreCase)))
                            instanceSheet.AddRow(instanceRowPlot);

                        var voltageProfileStart = GenProfileFlowRow(row.Parameter + "_Voltage_Profile_Start", enable);
                        sheet.InsertRow(index, voltageProfileStart);

                        var instanceRowStart = GetInstanceRowStart(false, "V", voltageProfileStart, pins, testTimeFromExecProfile, "{0:F6}", row, pinInfoRows);
                        if (!instanceSheet.InstanceRows.Exists(x => x.TestName.Equals(instanceRowStart.TestName, StringComparison.CurrentCultureIgnoreCase)))
                            instanceSheet.AddRow(instanceRowStart);
                    }
                }
            }
            return sheet;
        }

        private InstanceRow GetInstanceRowStart(bool byFlow, string whatToCapture, FlowRow profilePlot, List<string> pins, double testTimeFromExecProfile, string format,
            FlowRow flowRow = null, List<PinInfoRow> pinInfoRows = null)
        {
            var instanceRowStart = new InstanceRow();
            instanceRowStart.TestName = profilePlot.Parameter;
            instanceRowStart.Type = "VBT";
            instanceRowStart.Name = "pwfm_VolCurr_trig";//"pwfm_current_trig";
            var vbtFunctionBase = _basMain.GetFunctionByName(instanceRowStart.Name);
            if (flowRow != null && pinInfoRows != null && !_profileToolForm.checkBox_PowerPinOnly.Checked)
                pins = AddNonCorepowerPins(pins, flowRow, pinInfoRows);
            vbtFunctionBase.SetParamValue("pins", string.Join(",", pins.Where(x => !string.IsNullOrEmpty(x))));
            vbtFunctionBase.SetParamValue("duration", string.Format(format, testTimeFromExecProfile));
            var trigType = whatToCapture.Equals("I", StringComparison.CurrentCultureIgnoreCase) ? "1" : "0";
            vbtFunctionBase.SetParamValue("TrigType", trigType);
            instanceRowStart.ArgList = vbtFunctionBase.Parameters;
            instanceRowStart.Args = vbtFunctionBase.Args;
            return instanceRowStart;
        }

        private InstanceRow GetInstanceRowPlot(FlowRow profilePlot, List<string> pins, FlowRow flowRow = null, List<PinInfoRow> pinInfoRows = null)
        {

            var instanceRowPlot = new InstanceRow();
            instanceRowPlot.TestName = profilePlot.Parameter;
            instanceRowPlot.Type = "VBT";
            instanceRowPlot.Name = "pwfm_fetch";
            var vbtFunctionBase = _basMain.GetFunctionByName(instanceRowPlot.Name);
            if (flowRow != null && pinInfoRows != null)
                pins = AddNonCorepowerPins(pins, flowRow, pinInfoRows);
            vbtFunctionBase.SetParamValue("pins", string.Join(",", pins.Where(x => !string.IsNullOrEmpty(x))));
            instanceRowPlot.ArgList = vbtFunctionBase.Parameters;
            instanceRowPlot.Args = vbtFunctionBase.Args;
            return instanceRowPlot;
        }

        private List<string> AddNonCorepowerPins(List<string> pins, FlowRow flowRow, List<PinInfoRow> pinInfoRows)
        {
            if (pinInfoRows.Exists(x => x.InstanceName.Equals(flowRow.Parameter, StringComparison.CurrentCultureIgnoreCase)))
            {
                var row = pinInfoRows.Find(x => x.InstanceName.Equals(flowRow.Parameter, StringComparison.CurrentCultureIgnoreCase));
                foreach (var pin in row.PinList)
                {
                    if (!pins.Exists(x => x.Equals(pin, StringComparison.CurrentCultureIgnoreCase)))
                        pins.Add(pin);
                }
            }
            return pins.Where(x => !string.IsNullOrEmpty(x)).ToList();
        }

        private FlowRow GenProfileFlowRow(string parameter, string enable)
        {
            var flowRow = new FlowRow();
            flowRow.Enable = enable;
            flowRow.Opcode = "Test";
            flowRow.Parameter = parameter;
            return flowRow;
        }
    }
}