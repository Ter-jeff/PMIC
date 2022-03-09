using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenFlow;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenFlow
{
    public class DcTestFlowRowGenerator : FlowRowGeneratorBase
    {
        public DcTestFlowRowGenerator(string sheetName) : base(sheetName)
        {
        }

        protected override void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            BlockName = CommonGenerator.GetBlockName(pattern.MiscInfo, pattern.SheetName);
        }

        public override List<FlowRow> GenTestRows(bool isCz2Only = false)
        {
            var testRows = new List<FlowRow>();
            var testRow = new FlowRow();
            testRow.Job = ""; // CreateTestJob();
            testRow.OpCode = CreateTestOpCode();
            testRow.Env = CreateTestEnv();
            testRow.Enable = ""; //GenEnable(CreateTestEnable(isCz2Only), testRow.Env);
            testRow.Parameter = CreateTestParameter();
            testRow.Comment1 = SubBlockName;
            testRow.FailAction = CreateTestFailAction();
            testRows.Add(testRow);
            if(!SheetName.Equals(PmicConst.FlowDcLeakage.Replace("Flow_", ""), StringComparison.OrdinalIgnoreCase))
                SortFlowRows(GetLimitRows(testRow), testRows);
            return testRows;
        }

        protected List<FlowRow> GetLimitRows(FlowRow testRow)
        {
            //use limit rows
            var useLimitFailAction = CreateUseLimitFailAction();
            return GenUseLimitRows(testRow.Parameter, useLimitFailAction);
        }

        public override List<FlowRow> GenShmooRows(string labelVoltage = "")
        {
            var testRows = new List<FlowRow>();
            if (!HasShmoo)
                return testRows;
            var testRow = new FlowRow();
            testRow.Env = CreateTestEnv();
            testRow.Enable = "!TestOnly";
            testRow.OpCode = HardIpConstData.OpCodeChar;
            testRow.Parameter = Pat.Shmoo.IsSplitByVoltage
                ? CreateShmooParameter() + "_" + labelVoltage
                : CreateShmooParameter();
            testRow.Parameter = testRow.Parameter.Trim('_');
            testRow.Name = CreateShmooTName();
            testRows.Add(testRow);
            if (!SheetName.Equals(PmicConst.FlowDcLeakage.Replace("Flow_", ""), StringComparison.OrdinalIgnoreCase))
                SortFlowRows(GetLimitRows(testRow), testRows);
            return testRows;
        }

      

        public override List<FlowRow> GenPreRetestRows()
        {
            var preRestRows = new List<FlowRow>();
            if (!HasRetest)
                return preRestRows;
            // Clear Flag
            preRestRows.Add(GenRetestClearFlagRow());
            // Test Row 
            var testRow = new FlowRow();
            testRow.Job = CreateTestJob();
            testRow.OpCode = CreateTestOpCode();
            testRow.Enable = EnableWord;
            testRow.Env = CreatePreTestEnv();
            testRow.Parameter = CreatePreTestParameter();
            testRow.FailAction = CreatePreTestFailAction();
            preRestRows.Add(testRow);
            // Use Limit Rows
            preRestRows.AddRange(GenUseLimitRows(testRow.Parameter, testRow.FailAction));
            return preRestRows;
        }

        protected List<FlowRow> GenUseLimitRows(string parameter, string binFail)
        {
            var flowUseLimitRows = new List<FlowRow>();

            #region Generate Use-Limit rows

            var repeatStr = CommonGenerator.GetRepeatMapping(Pat.MiscInfo);
            List<MeasPin> useLimits;
            switch (ActualLabelVoltage)
            {
                case HardIpConstData.LabelHv:
                    useLimits = Pat.UseLimitsH;
                    break;
                case HardIpConstData.LabelLv:
                    useLimits = Pat.UseLimitsL;
                    break;
                case HardIpConstData.LabelNv:
                    useLimits = Pat.UseLimitsN;
                    break;
                default:
                    useLimits = Pat.UseLimitsN;
                    break;
            }

            //Gets the result that whether the pinGroup need to be decomposed for V,F,I or V,I,R
            var isSingleLimit = SearchInfo.GetFlagSingleLimit(Pat, ActualLabelVoltage);

            var dicUseLimits = useLimits.GroupBy(p => p.SequenceIndex).ToDictionary(p => p.Key, p => p.ToList());
            var storeNameList = SearchInfo.GetStoreName(Pat).Split('+').ToList();
            var i = -1;
            foreach (var item in dicUseLimits)
            {
                i++;
                var useLimitPin = new SortLimitPin();
                foreach (var pin in item.Value)
                {
                    if (pin.MeasType.Equals(MeasType.MeasC, StringComparison.OrdinalIgnoreCase) &&
                        pin.TestName.Equals("skip", StringComparison.OrdinalIgnoreCase) ||
                        pin.MeasType.Equals(MeasType.MeasN, StringComparison.OrdinalIgnoreCase))
                        continue;
                    var name = pin.PinName.Replace(" ", "");
                    if (name == "")
                        pin.PinName = "";
                    var nameList = new List<string>();
                    if (pin.MeasType.ToLower() == "measf" && name.Contains("::") ||
                        pin.MeasType.ToLower() == "measfdiff")
                    {
                        var groupName = SearchInfo.GenDiffGroupName(name, true);
                        nameList.Add(groupName);
                    }
                    else if (name.Contains("::") && !pin.MeasType.ToLower().Contains("diff") &&
                             pin.MeasType.ToLower() != "measvocm" && pin.MeasType.ToLower() != "measvdm")
                    {
                        nameList = Regex.Split(name, "::").ToList();
                    }
                    else
                    {
                        #region Change by laura at 2016/08/26

                        if (isSingleLimit)
                            nameList.Add(name);
                        else
                            nameList.AddRange(SearchInfo.DecomposeGroups(name));

                        #endregion
                    }

                    foreach (var pinName in nameList) useLimitPin.AddData(pinName, pin);
                }

                List<SortLimitPin> sortUseLimitPins;
                if (i < storeNameList.Count && storeNameList[i].Contains(':') || item.Value.Exists(p =>
                    p.MeasType.Equals(MeasType.MeasVdiff, StringComparison.OrdinalIgnoreCase))
                ) //|| (LocalSpecs.Device == DeviceEnum.RF))
                    sortUseLimitPins = useLimitPin.UseLimitPins;
                else
                    sortUseLimitPins = useLimitPin.UseLimitPins.OrderBy(x => x.PinName).ToList();

                foreach (var pin in sortUseLimitPins)
                foreach (var repeat in repeatStr.Split(','))
                {
                    var row = new FlowRow();
                    string lowUnit;
                    string lowScale;
                    string highUnit;
                    string highScale;
                    row.OpCode = HardIpConstData.OpCodeUseLimit;
                    row.Job = pin.MeasPinData.Job;
                    row.Parameter = parameter;
                    if (!string.IsNullOrEmpty(pin.MeasPinData.TestName))
                        row.Name = pin.MeasPinData.TestName;
                    var pinNameString = pin.PinName == "" ? "NoPinName" : pin.PinName;
                    row.Comment = Pat.Pattern.GetLastPayload() + "_" + pin.MeasPinData.MeasType + "_" + pinNameString;
                    row.LoLim = DataConvertor.ConvertUseLimit(pin.MeasPinData.LowLimit, out lowUnit, out lowScale);
                    if (repeat != "")
                    {
                        if (row.LoLim.Contains("="))
                            row.LoLim = row.LoLim.Replace("=", "=(") + ")" + "*" + repeat.Replace("x", "");
                        else
                            row.LoLim = CommonGenerator.CalculateLimit(row.LoLim, repeat);
                    }

                    row.HiLim = DataConvertor.ConvertUseLimit(pin.MeasPinData.HighLimit, out highUnit, out highScale);
                    if (repeat != "")
                    {
                        if (row.HiLim.Contains("="))
                            row.HiLim = row.HiLim.Replace("=", "=(") + ")" + "*" + repeat.Replace("x", "");
                        else
                            row.HiLim = CommonGenerator.CalculateLimit(row.HiLim, repeat);
                    }

                    row.Scale = lowScale == "" ? highScale : lowScale;

                    row.Units = lowUnit == "" ? highUnit : lowUnit;
                    row.FailAction = CommonGenerator.GetUseLimitFailFlag(pin.MeasPinData.MiscInfo) ? "" : binFail;
                    flowUseLimitRows.Add(row);
                }
            }

            #endregion

            return flowUseLimitRows;
        }

        protected virtual string CreateTestParameter()
        {
            if (!string.IsNullOrEmpty(Pat.TestName)) return Pat.TestName + "_" + ActualLabelVoltage;

            var patternName = Pattern.Pattern.GetPatternName();
            var para = CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, SubBlock2Name, IpName, patternName,
                Pat.PatternIndexFlag, TimingAc, Pat.DivideFlag, InstNameSubStr, ActualLabelVoltage, NoPattern,
                Pat.WirelessData.IsNeedPostBurn, true, Pat.WirelessData.IsDoMeasure);
            if (HasRetest)
                para = para.Replace(LabelVoltage, HardIpConstData.PrefixReTest + ActualLabelVoltage);
            return para;
        }

        protected string CreateShmooParameter()
        {
            var charName = "";
            var shmoo = Pat.Shmoo;
            if (shmoo != null)
                charName = shmoo.SetupName;
            return CreateTestParameter() + " " + CommonGenerator.GetSubBlockNameWithoutMinus(charName);
        }

        protected string CreateShmooTName()
        {
            var name = "";
            var shmooName = Pat.Shmoo;
            if (shmooName != null)
            {
                if (shmooName.TestNameInFlow.StartsWith("HAC"))
                    return name;
                name = shmooName.TestNameInFlow;
                var charArr = name.Split('_');
                if (!string.IsNullOrEmpty(LabelVoltage))
                {
                    charArr[2] = LabelVoltage.Substring(0, 1);
                    name = string.Join("_", charArr);
                }

                return name;
            }

            return name;
        }

        protected virtual string CreateTestFailAction()
        {
            if (CommonGenerator.GetPatternFailFlag(Pat.MiscInfo))
                return "";
            var patternName = Pattern.Pattern.GetPatternName();
            return CommonGenerator.GenHardIpFlowTestFailAction(SheetName, BlockName, SubBlockName, SubBlock2Name,
                patternName, TimingAc, InstNameSubStr, LabelVoltage, Pat.MiscInfo, NoPattern);
        }

        private string CreatePreTestEnv()
        {
            var env = string.Empty;
            return string.IsNullOrEmpty(CreateTestJob()) ? HardIpConstData.EnvTtr : env;
        }

        private string CreatePreTestParameter()
        {
            return CommonGenerator.GenHardIpInsTestName(BlockName, SubBlockName, SubBlock2Name, IpName,
                Pat.Pattern.GetLastPayload(), Pat.PatternIndexFlag, TimingAc, Pat.DivideFlag, InstNameSubStr,
                ActualLabelVoltage, NoPattern, Pat.WirelessData.IsNeedPostBurn, true, Pat.WirelessData.IsDoMeasure);
        }

        private string CreatePreTestFailAction()
        {
            return HardIpConstData.ReTestFlag;
        }

        protected string CreateUseLimitFailAction()
        {
            var patternName = Pattern.Pattern.GetPatternName();
            return CommonGenerator.GenHardIpFlowUseLimitFailAction(SheetName, BlockName, SubBlockName, SubBlock2Name,
                patternName, TimingAc, InstNameSubStr, LabelVoltage, Pat.MiscInfo, NoPattern);
        }

        protected virtual string CreateBinTableParameter()
        {
            var patternName = Pattern.Pattern.GetPatternName();
            if (LocalSpecs.CurrentProject.Equals("sicily", StringComparison.OrdinalIgnoreCase) ||
                LocalSpecs.CurrentProject.Equals("tonga", StringComparison.OrdinalIgnoreCase))
                return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName);
            return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName, patternName, TimingAc,
                InstNameSubStr, NoPattern);
        }

        public override FlowRow GenBinTableRow(string voltage = "")
        {
            var name = string.Format("{0}_{1}", CreateBinTableParameter(), voltage).Trim('_');
            return WriteBinTableItem(name);
        }

        private FlowRow WriteBinTableItem(string parameter)
        {
            var rowBin = new FlowRow();
            rowBin.OpCode = FlowRow.OpCodeBinTable;
            rowBin.Parameter = parameter;
            rowBin.Enable = FlowRow.OpCodeBinTable;
            return rowBin;
        }
    }
}