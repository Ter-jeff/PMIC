using CommonLib.Enum;
using CommonLib.ErrorReport;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Basic.GenConti.Base;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Inputs.TestPlan.Reader.DcTestConti;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Basic.GenConti
{
    public class DcContiMain
    {
        private readonly DcTestContinuitySheet _dcTestContiSheet;

        public DcContiMain(DcTestContinuitySheet dcTestContiSheet)
        {
            _dcTestContiSheet = dcTestContiSheet;
        }

        public Dictionary<IgxlSheet, string> WorkFlow()
        {
            var contiTestList = CreateContiTest();

            var subFlowSheet = GenFlowSheet(contiTestList);

            var instanceSheet = GenInsSheet(contiTestList);

            var binTableRows = GenBinTableRows();
            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            binTable.AddRows(binTableRows);

            var igxlSheets = new Dictionary<IgxlSheet, string>();
            igxlSheets.Add(subFlowSheet, FolderStructure.DirDc);
            igxlSheets.Add(instanceSheet, FolderStructure.DirDc);
            return igxlSheets;
        }

        protected List<InstanceRow> GenerateCommonInstanceRows(List<DcContiTestBase> contiTestList)
        {
            var commonInstanceRows = new List<InstanceRow>();
            var status = new RelayStatus();
            foreach (var test in contiTestList)
            {
                if (status.IsEqualStatus(test.Relay)) continue;
                var commonInstRow = test.GenerateRelayInstanceRow(status);
                status = test.Relay;
                commonInstanceRows.Add(commonInstRow);
            }

            return commonInstanceRows;
        }

        protected List<DcContiTestBase> CreateContiTest()
        {
            var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;

            var relayTable = new DataTable();
            var contiTestList = new List<DcContiTestBase>();

            #region default pin group open short

            var firstIoPinGroup = true;
            foreach (var row in _dcTestContiSheet.Rows)
            {
                var pinList = pinMap.GetPinsFromGroup(row.PinGroup);
                var pinType = pinList.Select(x => x.PinType).Distinct().ToList();

                //Is single pin
                if (!pinType.Any() && pinMap.IsPinExist(row.PinGroup))
                    pinType.Add(pinMap.GetPinType(row.PinGroup));

                //if (pinType.Count == 1)
                //{

                //Support the item contains empty row, added by terry
                if (pinType.Count == 0) continue;

                if (pinType[0].Equals(PinMapConst.TypeIo, StringComparison.OrdinalIgnoreCase))
                {
                    if (firstIoPinGroup)
                    {
                        contiTestList.Add(new DcContiTestOpenShortIoPmic(row, relayTable));
                        firstIoPinGroup = false;
                    }
                    else
                    {
                        contiTestList.Add(new DcContiTestOpenShortIoPmic(row, relayTable));
                    }
                }
                else if (pinType[0].Equals(PinMapConst.TypePower, StringComparison.OrdinalIgnoreCase))
                {
                    contiTestList.Add(new DcContiTestOpenShortPowerPmic(row, relayTable));
                }
                else if (pinType[0].Equals(PinMapConst.TypeAnalog, StringComparison.OrdinalIgnoreCase))
                {
                    contiTestList.Add(new DcContiTestOpenShortAnalogPmic(row, relayTable));
                }

                //}
                if (pinType.Count > 1)
                {
                    var errorMessage = string.Format("The pin group {0} has more than two pin types !!!", row.PinGroup);
                    ErrorManager.AddError(EnumErrorType.MissingPinName, EnumErrorLevel.Error, PmicConst.DcTestContinuity,
                        row.RowNum, 1, errorMessage);
                }
            }

            #endregion

            return contiTestList;
        }

        private SubFlowSheet GenFlowSheet(List<DcContiTestBase> contiTestList)
        {
            var subFlowSheet = new SubFlowSheet(PmicConst.FlowDcConti);
            subFlowSheet.FlowRows.AddStartRows(subFlowSheet.SheetName, SubFlowSheet.Ttime);

            #region generate flag

            subFlowSheet.FlowRows.AddFlowRow("Flag-Clear", DcContiConst.FlagNameOpen);
            subFlowSheet.FlowRows.AddFlowRow("Flag-Clear", DcContiConst.FlagNameShort);
            subFlowSheet.FlowRows.AddFlowRow("Flag-Clear", DcContiConst.FlagNameVoltageClampCheck);

            subFlowSheet.FlowRows.AddFlowRow(FlowRow.OpCodeTest, DcContiConst.DgsRelayOn, "D_OpenSocket");

            var openShortTestList = contiTestList.FindAll(p => p is DcContiTestOpenShortIoPmic);
            //generate flag for PPMUOS with walking Z
            foreach (var contiTestBase in openShortTestList)
                subFlowSheet.FlowRows.AddFlowRow("Flag-true", contiTestBase.CreateWalkingZFlagName(), "PPMUOS");

            #endregion

            subFlowSheet.FlowRows.AddFlowRow("nop", "SetPower_Alarm");

            var status = new RelayStatus();
            foreach (var test in contiTestList)
            {
                var flowRows = test.GenerateFlowRows(status);
                status = test.Relay;
                subFlowSheet.FlowRows.AddRange(flowRows);
                subFlowSheet.FlowRows.AddFlowRow(FlowRow.OpCodeBinTable, DcContiConst.BinNameOpen);
                subFlowSheet.FlowRows.AddFlowRow(FlowRow.OpCodeBinTable, DcContiConst.BinNameShort);
                subFlowSheet.FlowRows.AddFlowRow(FlowRow.OpCodeBinTable, DcContiConst.BinNameAlarmFail);
            }

            subFlowSheet.FlowRows.AddFlowRow(FlowRow.OpCodeBinTable, DcContiConst.BinNameContiVoltageClampCheck);
            subFlowSheet.FlowRows.AddFlowRow(FlowRow.OpCodeTest, DcContiConst.DgsRelayOff, "D_OpenSocket");

            //AutoZ Only set-device row
            var row = new FlowRow();
            row.OpCode = "set-device";
            row.Enable = "AutoZOnly";
            row.BinFail = "9";
            row.SortFail = "9997";
            row.Result = "Fail";
            subFlowSheet.AddRow(row);

            subFlowSheet.FlowRows.AddFlowRow("nop", "PowerUp");
            subFlowSheet.FlowRows.AddEndRows(subFlowSheet.SheetName, SubFlowSheet.Ttime, false);
            return subFlowSheet;
        }

        private InstanceSheet GenInsSheet(List<DcContiTestBase> contiTestList)
        {
            var contiInstance = new InstanceSheet(PmicConst.TestInstDcConti);

            contiInstance.AddHeaderFooter();
            foreach (var test in contiTestList)
            {
                var instRows = test.GenerateInstanceRows();
                instRows.ForEach(x => x.PinLevels = "Levels_Func");
                contiInstance.InstanceRows.AddRange(instRows);
            }

            var commonInstanceRows = GenerateCommonInstanceRows(contiTestList);
            contiInstance.AddRows(commonInstanceRows);

            return contiInstance;
        }

        #region Bintable

        protected BinTableRows GenBinTableRows()
        {
            var binTableRows = new BinTableRows();

            GenOpen(binTableRows);

            GenShort(binTableRows);

            GenOpenShort(binTableRows);

            GenPowerShort(binTableRows);

            GenPowerOpen(binTableRows);

            GenAutoZCheck(binTableRows);

            GenAlarmFail(binTableRows);

            binTableRows.GenSetError("DC_Conti");

            return binTableRows;
        }

        private void GenAutoZCheck(BinTableRows binTableRows)
        {
            var binTableRow = new BinTableRow();
            binTableRow.Name = DcContiConst.BinAutoZCheck;
            binTableRow.ItemList = DcContiConst.FlagNameAutoZCheck;
            binTableRow.Op = "AND";
            binTableRow.Sort = "9997";
            binTableRow.Bin = "9";
            binTableRow.Result = "Fail-Stop";
            binTableRow.Items.Add("T");
            binTableRows.Add(binTableRow);
        }

        private void GenPowerOpen(BinTableRows binTableRows)
        {
            BinNumberRuleRow numDefRow;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.DcConti, "PowerOpen");
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out numDefRow);
            var binTableRow = new BinTableRow();
            binTableRow.Name = DcContiConst.BinNamePowerOpen;
            binTableRow.ItemList = DcContiConst.FlagNamePowerOpen;
            binTableRow.Op = "AND";
            binTableRow.Sort = numDefRow.CurrentSoftBin.ToString();
            binTableRow.Bin = numDefRow.HardBin;
            binTableRow.Result = numDefRow.SoftBinState;
            binTableRow.Items.Add("T");
            binTableRows.Add(binTableRow);
        }

        private void GenPowerShort(BinTableRows binTableRows)
        {
            BinNumberRuleRow numDefRow;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.DcConti, "PowerShort");
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out numDefRow);
            var binTableRow = new BinTableRow();
            binTableRow.Name = DcContiConst.BinNamePowerShort;
            binTableRow.ItemList = DcContiConst.FlagNamePowerShort;
            binTableRow.Op = "AND";
            binTableRow.Sort = numDefRow.CurrentSoftBin.ToString();
            binTableRow.Bin = numDefRow.HardBin;
            binTableRow.Result = numDefRow.SoftBinState;
            binTableRow.Items.Add("T");
            binTableRows.Add(binTableRow);
        }

        private void GenOpenShort(BinTableRows binTableRows)
        {
            BinNumberRuleRow numDefRow;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.DcConti, "OpenShort");
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out numDefRow);
            var binTableRow = new BinTableRow();
            binTableRow.Name = DcContiConst.BinNameOpenShort;
            binTableRow.ItemList = DcContiConst.FlagNameOpen + "," + DcContiConst.FlagNameShort;
            binTableRow.Op = "AND";
            binTableRow.Sort = numDefRow.CurrentSoftBin.ToString();
            binTableRow.Bin = numDefRow.HardBin;
            binTableRow.Result = numDefRow.SoftBinState;
            binTableRow.Items.AddRange(new[] { "T", "T" });
            binTableRows.Add(binTableRow);
        }

        private void GenShort(BinTableRows binTableRows)
        {
            BinNumberRuleRow numDefRow;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.DcConti, "Short");
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out numDefRow);
            var binTableRow = new BinTableRow();
            binTableRow.Name = DcContiConst.BinNameShort;
            binTableRow.ItemList = DcContiConst.FlagNameShort;
            binTableRow.Op = "AND";
            binTableRow.Sort = numDefRow.CurrentSoftBin.ToString();
            binTableRow.Bin = numDefRow.HardBin;
            binTableRow.Result = numDefRow.SoftBinState;
            binTableRow.Items.Add("T");
            binTableRows.Add(binTableRow);
        }

        private void GenOpen(BinTableRows binTableRows)
        {
            BinNumberRuleRow numDefRow;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.DcConti, "Open");
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out numDefRow);
            var binTableRow = new BinTableRow();
            binTableRow.Name = DcContiConst.BinNameOpen;
            binTableRow.ItemList = DcContiConst.FlagNameOpen;
            binTableRow.Op = "AND";
            binTableRow.Sort = numDefRow.CurrentSoftBin.ToString();
            binTableRow.Bin = numDefRow.HardBin;
            binTableRow.Result = numDefRow.SoftBinState;
            binTableRow.Items.Add("T");
            binTableRows.Add(binTableRow);
        }

        private void GenAlarmFail(BinTableRows binTableRows)
        {
            BinNumberRuleRow numDefRow;
            var para = new BinNumberRuleCondition(EnumBinNumberBlock.DcConti, "AlarmFail");
            BinNumberSingleton.Instance().GetBinNumDefRow(para, out numDefRow);
            var binTableRow = new BinTableRow();
            binTableRow.Name = DcContiConst.BinNameAlarmFail;
            binTableRow.ItemList = DcContiConst.FlagNameAlarmFail;
            binTableRow.Op = "AND";
            binTableRow.Sort = numDefRow.CurrentSoftBin.ToString();
            binTableRow.Bin = numDefRow.HardBin;
            binTableRow.Result = numDefRow.SoftBinState;
            binTableRow.Items.Add("T");
            binTableRows.Add(binTableRow);
        }

        #endregion
    }
}