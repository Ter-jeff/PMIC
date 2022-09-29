using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.GenerateIgxl.HardIp.SpecialSetting;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenFlow
{
    public abstract class FlowRowGeneratorBase
    {
        protected int ForNumber;

        public string LabelVoltage = string.Empty;
        protected HardIpPattern Pattern;
        protected string SheetName;

        protected FlowRowGeneratorBase(string sheetName)
        {
            SheetName = string.Empty;
            SheetName = sheetName;
        }

        public HardIpPattern Pat
        {
            set
            {
                Pattern = value;
                SetBasicInfoByPattern(value);
            }
            get { return Pattern; }
        }

        protected string Block { set; get; } = string.Empty;

   

        protected string IpName => CommonGenerator.GetIpName(Pat.MiscInfo);

        protected string SubBlockName => CommonGenerator.GetSubBlockName(Pat.Pattern.GetLastPayload(), Pat.MiscInfo,
            Block, Pat.ForceCondition.IsCz2InstName);

        protected string SubBlock2Name => CommonGenerator.GetSubBlock2Name(Pat.MiscInfo);

        protected string TimingAc => CommonGenerator.GetTimingAc(Pat.AcUsed);

        protected string InstNameSubStr => CommonGenerator.GetInstNameSubStr(Pat.MiscInfo);

        protected string EnableWord =>
            CommonGenerator.GenEnableWord(Pat.Pattern.GetLastPayload(), Pat.MiscInfo, LabelVoltage);

        protected bool NoPattern => CommonGenerator.NoPattern(Pat.Pattern.GetLastPayload());

        protected bool HasRetest => CommonGenerator.HasRetest(Pat.MiscInfo);

        protected bool HasSweepCode => CommonGenerator.HasSweepCode(Pat);

        protected bool HasSweepVoltage => CommonGenerator.HasSweepVoltage(Pat);

        public bool HasShmoo => CommonGenerator.HasShmoo(Pat);

        protected string ActualLabelVoltage => CommonGenerator.ActualLabelVoltage(LabelVoltage, Pat);

        public bool NoNeedToGen => CommonGenerator.PatFlowNoNeedToGen(LabelVoltage, Pat.MiscInfo);

        public int FlowControlFlag => Pat.FlowControlFlag;

        public bool IsFlowInsRepeat => Pat.IsFlowInsRepeat;

        public abstract FlowRows GenTestRows(bool isCz2Only = false);

        public abstract FlowRow GenBinTableRow(string voltage = "");

        public abstract FlowRows GenShmooRows(string labelVoltage = "");

        protected abstract void SetBasicInfoByPattern(HardIpPattern pattern);

        public FlowRows GenOpCodeRowsBefPat(List<string> opCodeList)
        {
            var flowRows = new FlowRows();
            flowRows.AddRange(OpCodeSettingMain.GenOpCodeSetting(opCodeList, Block, LabelVoltage, EnableWord));
            return flowRows;
        }

        public FlowRows GenOpCodeRowsAftPat()
        {
            var flowRows = new FlowRows();
            if (Regex.IsMatch(Pat.MiscInfo, HardIpConstData.RegOpCode, RegexOptions.IgnoreCase))
            {
                var opCodeList = SearchInfo.GetOpCode(Pat, "A");
                flowRows.AddRange(OpCodeSettingMain.GenOpCodeSetting(opCodeList, Block, LabelVoltage, EnableWord));
            }

            return flowRows;
        }

        public FlowRows GenSweepCodeForRow()
        {
            if (!HasSweepCode || FlowControlFlag == 1 && IsFlowInsRepeat)
                return null;
            var flowRows = new FlowRows();
            foreach (var sweepItems in Pat.SweepCodes)
            {
                var srcCodeIndex = "SrcCodeIndx" + sweepItems.Key;
                var flowRow = new FlowRow();
                flowRow.OpCode = "for";
                flowRow.Env = CreateTestEnv();
                flowRow.Enable = GenEnable(EnableWord, flowRow.Env);
                int steps;
                if (string.IsNullOrEmpty(sweepItems.Value[0].Misc))
                {
                    if ((sweepItems.Value[0].End - sweepItems.Value[0].Start) % sweepItems.Value[0].Step == 0)
                        steps = (sweepItems.Value[0].End - sweepItems.Value[0].Start) / sweepItems.Value[0].Step + 1;
                    else
                        steps = (sweepItems.Value[0].End - sweepItems.Value[0].Start) / sweepItems.Value[0].Step;
                }
                else
                {
                    steps = sweepItems.Value[0].End;
                }

                flowRow.Parameter = string.Format("{0} = 0; {0} < {1}; {0}++", srcCodeIndex, steps);
                flowRows.Add(flowRow);
                if (!TestProgram.IgxlWorkBk.FlowUsedInteger.Contains(srcCodeIndex))
                    TestProgram.IgxlWorkBk.FlowUsedInteger.Add(srcCodeIndex);
                ForNumber++;
            }

            return flowRows;
        }

        public FlowRows GenSweepVoltageForRow()
        {
            if (!HasSweepVoltage || (FlowControlFlag == -1 || FlowControlFlag == 1) && IsFlowInsRepeat)
                return null;
            var flowRows = new FlowRows();

            var svDataXy = _AnalyzeSweepVoltage();
            foreach (var svData in svDataXy)
            {
                var forRow = new FlowRow();
                forRow.Env = CreateTestEnv();
                forRow.Enable = GenEnable(EnableWord, forRow.Env);
                forRow.OpCode = "for";
                var srcCodeIndx = svData[0].Axis.Equals("X", StringComparison.CurrentCultureIgnoreCase)
                    ? "SrcCodeIndx"
                    : "SrcCodeIndxY";
                var count = (int)(Math.Abs(double.Parse(svData[0].Stop) - double.Parse(svData[0].Start)) /
                                   double.Parse(svData[0].Step));
                forRow.Parameter = string.Format("For {0} = 0; {0} < {1}; {0}++", srcCodeIndx, count + 1);
                flowRows.Add(forRow);
                ForNumber++;
            }

            return flowRows;
        }

        public FlowRows GenSweepCodeOrVoltageNextRow()
        {
            if (!HasSweepCode && !HasSweepVoltage || (FlowControlFlag == 0 || FlowControlFlag == -1) && IsFlowInsRepeat)
                return null;
            var flowRows = new FlowRows();
            if (ForNumber > 0)
            {
                for (var i = 0; i < ForNumber; i++)
                {
                    var flowRow = new FlowRow();
                    flowRow.OpCode = "next";
                    flowRows.Add(flowRow);
                }

                ForNumber = 0;
            }

            return flowRows;
        }

        public virtual FlowRows GenPreRetestRows()
        {
            return null;
        }

        public FlowRow GenRetestIfRow()
        {
            if (!HasRetest)
                return null;
            var ifRow = new FlowRow();
            ifRow.OpCode = "if";
            ifRow.Parameter = HardIpConstData.ReTestFlag;
            ifRow.Enable = "";
            return ifRow;
        }

        public FlowRow GenRetestEndIfRow()
        {
            if (!HasRetest)
                return null;
            var ifRow = new FlowRow();
            ifRow.OpCode = "endif";
            ifRow.Parameter = HardIpConstData.ReTestFlag;
            return ifRow;
        }

        protected FlowRow GenRetestClearFlagRow()
        {
            var flagClearRow = new FlowRow();
            flagClearRow.OpCode = HardIpConstData.FlagClear;
            flagClearRow.Parameter = HardIpConstData.ReTestFlag;
            flagClearRow.Enable = EnableWord;
            return flagClearRow;
        }

        public FlowRows GenTtrFlagClearRow(FlowRows flowBodyRows)
        {
            var list = flowBodyRows.Where(x => x.Env == HardIpConstData.EnvTtr).Select(y => y.FailAction).Distinct();
            var flagClearRowList = new FlowRows();
            foreach (var item in list)
            {
                var flagClearRow = new FlowRow();
                flagClearRow.OpCode = HardIpConstData.FlagClear;
                flagClearRow.Parameter = item;
                flagClearRowList.Add(flagClearRow);
            }

            return flagClearRowList;
        }

        public FlowRow GenNWireHardIpRow()
        {
            var returnRow = new FlowRow();
            returnRow.OpCode = "Call";
            returnRow.Parameter = "Flow_nWire_HARDIP";
            returnRow.Enable = EnableWord;
            return returnRow;
        }

        protected string CreateTestJob()
        {
            return ""; //SearchInfo.GetTtrEnable(Pat.TtrStr, LabelVoltage);
        }

        protected string CreateTestOpCode()
        {
            if (Pat.UseDeferLimit) return HardIpConstData.OpCodeTestDeferLimit;
            return HardIpConstData.OpCodeTest;
        }

        protected string CreateTestEnv()
        {
            var env = SearchInfo.GetEnvFromPattern(Pat);
            return env;
        }

        protected string CreateTestEnable(bool isCz2Only)
        {
            var value = isCz2Only ? "!ShmooOnly" : EnableWord;
            return value;
        }

        protected string CreateBinTableEnv()
        {
            return SearchInfo.GetEnvFromPattern(Pat, true);
        }

        protected string CreateBinTableOpCode()
        {
            return HardIpConstData.OpCodeBinTable;
        }

        protected string CreateBinTableEnable()
        {
            return HardIpConstData.HardipBinEnable;
        }

        private List<List<SweepVData>> _AnalyzeSweepVoltage()
        {
            var result = new List<List<SweepVData>>();
            foreach (var xyItem in Pat.SweepVoltage.OrderByDescending(p => p.Key))
            {
                var resultXy = new List<SweepVData>();
                foreach (var item in xyItem.Value)
                {
                    var data = new SweepVData(item);
                    data.Axis = xyItem.Key;
                    resultXy.Add(data);
                }

                result.Add(resultXy);
            }

            return result;
        }

        public string GenEnable(string enable, string env)
        {
            if (env != null && env.Equals("TTR", StringComparison.OrdinalIgnoreCase))
                enable = string.IsNullOrEmpty(enable) ? "HardIP_CZ" : enable + "||HardIP_CZ";
            return enable;
        }

        protected void SortFlowRows(FlowRows originFlowRows, FlowRows sortedFlowRows)
        {
            var testRowsDefault = new FlowRows();
            var testRowsCalcLimit = new FlowRows();
            var testRowsMeasC = new FlowRows();
            //if (LocalSpecs.Device != DeviceEnum.LCD && LocalSpecs.Device != DeviceEnum.RF)
            //{
            foreach (var row in originFlowRows)
                if (Regex.IsMatch(row.Comment, MeasType.MeasCalc, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(row.Comment, MeasType.MeasLimit, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(row.Comment, MeasType.WiMeas, RegexOptions.IgnoreCase))
                    testRowsCalcLimit.Add(row);
                else if (Regex.IsMatch(row.Comment, MeasType.MeasC, RegexOptions.IgnoreCase))
                    testRowsMeasC.Add(row);
                else
                    testRowsDefault.Add(row);
            sortedFlowRows.AddRange(testRowsDefault);
            sortedFlowRows.AddRange(testRowsMeasC);
            sortedFlowRows.AddRange(testRowsCalcLimit);
            //}
            //else
            //{
            //    sortedFlowRows.AddRange(originFlowRows);
            //}
        }

        //public FlowRows AddExtraBinRows(FlowRow binRow)
        //{
        //    var results = new FlowRows();
        //    results.Add(CopyBinRow(binRow, HardIpConstData.LabelNv));
        //    results.Add(CopyBinRow(binRow, HardIpConstData.LabelHv));
        //    results.Add(CopyBinRow(binRow, HardIpConstData.LabelLv));
        //    return results;
        //}

        //private FlowRow CopyBinRow(FlowRow refRow, string vol)
        //{
        //    var row = new FlowRow();
        //    row.Parameter = refRow.Parameter + "_" + vol;
        //    row.OpCode = FlowRow.OpCodeBinTable;
        //    row.Job = SearchInfo.GetTtrEnable(Pat.NoBinOutStr, vol);
        //    row.Enable = HardIpConstData.HardipBinEnable;
        //    foreach (var noBinFlag in Pat.NoBinOutStr.Split(';'))
        //    {
        //        if (noBinFlag.Contains(":")) continue;
        //        if (Regex.IsMatch(noBinFlag, vol, RegexOptions.IgnoreCase))
        //            row.Env = "NoBinOut";
        //    }

        //    return row;
        //}
    }
}