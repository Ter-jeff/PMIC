using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using AutomationCommon.EpplusErrorReport;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenConti.Base;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Inputs.TestPlan.Reader.DcTestConti;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenConti
{
    public class DcContiMain
    {
        private readonly DcTestContinuitySheet _dcTestContiSheet;

        #region Constructor

        public DcContiMain(DcTestContinuitySheet dcTestContiSheet)
        {
            _dcTestContiSheet = dcTestContiSheet;
        }

        #endregion

        public Dictionary<IgxlSheet, string> WorkFlow()
        {
            var contiTestList = CreateContiTest();

            var contiFlow = GenerateContiFlow(contiTestList);

            var contiInstanceSheet = GenerateContiInstanceSheet(contiTestList);

            var commonInstanceRows = GenerateCommonInstanceRows(contiTestList);

            contiInstanceSheet.InstanceRows.AddRange(commonInstanceRows);
            //contiInstanceSheet.InstanceRows.AddRange(GenerateDgsRelayInstanceRows());

            var binTableRows = GenerateBinTableRows();

            var igxlSheets = new Dictionary<IgxlSheet, string>();
            igxlSheets.Add(contiFlow, FolderStructure.DirConti);
            igxlSheets.Add(contiInstanceSheet, FolderStructure.DirConti);
            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            binTable.AddRows(binTableRows);
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

        protected List<InstanceRow> GenerateDgsRelayInstanceRows()
        {
            var dgsRelayInstanceRows = new List<InstanceRow>();
            InstanceRow row = new InstanceRow();
            row.TestName = DcContiConst.DgsRelayOn;
            row.Name = DcContiConst.DgsRelayOn;
            row.Type = "VBT";
            row.DcCategory = "Conti";
            row.DcSelector = "Typ";
            row.AcCategory = "Common";
            row.AcSelector = "Typ";
            row.PinLevels = "Levels_Func";
            row.TimeSets = StaticTestPlan.DcTestContinuitySheet.Rows.First().TimeSet;
            dgsRelayInstanceRows.Add(row);

            row = new InstanceRow();
            row.TestName = DcContiConst.DgsRelayOff;
            row.Name = DcContiConst.DgsRelayOff;
            row.Type = "VBT";
            row.DcCategory = "Conti";
            row.DcSelector = "Typ";
            row.AcCategory = "Common";
            row.AcSelector = "Typ";
            row.PinLevels = "Levels_Func";
            row.TimeSets = StaticTestPlan.DcTestContinuitySheet.Rows.First().TimeSet;
            dgsRelayInstanceRows.Add(row);

            return dgsRelayInstanceRows;
        }

        protected List<DcContiTestBase> CreateContiTest()
        {
            var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;

            var relayTable = new DataTable();
            var contiTestList = new List<DcContiTestBase>();

            #region default pin group open short

            bool firstIoPinGroup = true;
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
                if (pinType.Count == 0)
                {
                    continue;
                }

                if (pinType[0].Equals(PinMapConst.TypeIo, StringComparison.OrdinalIgnoreCase))
                {
                    if (firstIoPinGroup)
                    {
                        contiTestList.Add(new DcContiTestOpenShortIoPmic(row, relayTable, "", true));
                        firstIoPinGroup = false;
                    }
                    else
                    {
                        contiTestList.Add(new DcContiTestOpenShortIoPmic(row, relayTable));
                    }
                }
                else if (pinType[0].Equals(PinMapConst.TypePower, StringComparison.OrdinalIgnoreCase))
                    contiTestList.Add(new DcContiTestOpenShortPowerPmic(row, relayTable));
                else if (pinType[0].Equals(PinMapConst.TypeAnalog, StringComparison.OrdinalIgnoreCase))
                    contiTestList.Add(new DcContiTestOpenShortAnalogPmic(row, relayTable));
                //}
                if (pinType.Count > 1)
                {
                    var errorMessage = string.Format("The pin group {0} has more than two pin types !!!", row.PinGroup);
                    EpplusErrorManager.AddError(BasicErrorType.Business, ErrorLevel.Error, PmicConst.DcTestContinuity,
                        row.RowNum, 1, errorMessage);
                }
            }

            #endregion

            return contiTestList;
        }

        private SubFlowSheet GenerateContiFlow(List<DcContiTestBase> contiTestList)
        {
            var contiFlow = new SubFlowSheet(PmicConst.FlowDcConti);
            contiFlow.AddStartRows(SubFlowSheet.Ttime);

            #region generate flag

            contiFlow.AddFlowRow("Flag-Clear", DcContiConst.FlagNameOpen);
            contiFlow.AddFlowRow("Flag-Clear", DcContiConst.FlagNameShort);
            contiFlow.AddFlowRow("Flag-Clear", DcContiConst.FlagNameVoltageClampCheck);

            contiFlow.AddFlowRow("Test", DcContiConst.DgsRelayOn, "D_OpenSocket");

            var openShortTestList = contiTestList.FindAll(p => p is DcContiTestOpenShortIoPmic);
            //generate flag for PPMUOS with walking Z
            foreach (var contiTestBase in openShortTestList)
                contiFlow.AddFlowRow("Flag-true", contiTestBase.CreateWalkingZFlagName(), "PPMUOS");
            #endregion

            contiFlow.AddFlowRow("nop", "SetPower_Alarm");

            var status = new RelayStatus();
            foreach (var test in contiTestList)
            {
                var flowRows = test.GenerateFlowRows(status);
                status = test.Relay;
                contiFlow.FlowRows.AddRange(flowRows);
                contiFlow.AddFlowRow(FlowRow.OpCodeBinTable, DcContiConst.BinNameOpen);
                contiFlow.AddFlowRow(FlowRow.OpCodeBinTable, DcContiConst.BinNameShort);
            }

            contiFlow.AddFlowRow(FlowRow.OpCodeBinTable, DcContiConst.BinNameContiVoltageClampCheck);
            contiFlow.AddFlowRow("Test", DcContiConst.DgsRelayOff, "D_OpenSocket");

            //AutoZ Only set-device row
            FlowRow row = new FlowRow();
            row.OpCode = "set-device";
            row.Enable = "AutoZOnly";
            row.BinFail = "9";
            row.SortFail = "9997";
            row.Result = "Fail";
            contiFlow.AddRow(row);

            contiFlow.AddFlowRow("nop", "PowerUp");
            contiFlow.AddEndRows(SubFlowSheet.Ttime,false);
            return contiFlow;
        }


        private InstanceSheet GenerateContiInstanceSheet(List<DcContiTestBase> contiTestList)
        {
            var contiInstance = new InstanceSheet(PmicConst.TestInstDcConti);

            contiInstance.AddHeaderFooter();
            foreach (var test in contiTestList)
            {
                var instRows = test.GenerateInstanceRows();
                instRows.ForEach(x => x.PinLevels = "Levels_Func");
                contiInstance.InstanceRows.AddRange(instRows);
            }

            return contiInstance;
        }

        #region Bintable

        protected List<BinTableRow> GenerateBinTableRows()
        {
            var binTableRows = new List<BinTableRow>();
            GenOpen(binTableRows);

            GenShort(binTableRows);

            GenOpenShort(binTableRows);

            GenPowerShort(binTableRows);

            GenPowerOpen(binTableRows);

            GenAutoZCheck(binTableRows);

            return binTableRows;
        }

        private void GenAutoZCheck(List<BinTableRow> binTableRows)
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

        private void GenPowerOpen(List<BinTableRow> binTableRows)
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

        private void GenPowerShort(List<BinTableRow> binTableRows)
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

        private void GenOpenShort(List<BinTableRow> binTableRows)
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
            binTableRow.Items.AddRange(new[] {"T", "T"});
            binTableRows.Add(binTableRow);
        }

        private void GenShort(List<BinTableRow> binTableRows)
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

        private void GenOpen(List<BinTableRow> binTableRows)
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

        #endregion
    }
}