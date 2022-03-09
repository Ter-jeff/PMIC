using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.DataStructure;
using AutomationCommon.EpplusErrorReport;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.InputPackages;

namespace PmicAutogen.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    public class TestPlanPatParser
    {
        //eg: MeasC pin = JTAG(data):4 "testName" or  Limits "Cal_A" , (out1):4 or Limits IO1(out2)
        private const string RegTestNameExpression = @"([""])(?<testName>.+)([""])";

        //eg: MeasV Pin = PAD_MTR_ANALOG_TEST_N "Test 0"
        private const string RegPineNameExpression = @"^(?<pinName>[^)]+)[\s]*(([""])(?<testName>([^""])+)([""]))";

        //eg: label1(IO1:V:0;IO2:V:0.4)
        private const string RegForceConditionLabelDefine = @"(?<label>[\w]+)[\s]*([\(](?<value>[^)]*)[\)])";

        //eg: Calc "Add" sn1+sn2
        private const string RegCalcExpression = @"^Calc[\s]+(([""])(?<testName>.+)([""]))?[\s&,]*(?<expression>.*)$";

        //eg: Limits "Cal_A" , (out1):4 or Limits IO1(out2)
        private const string RegLimitsExpression =
            @"^Limits[\s]*(([""])(?<testName>.+)([""]))?[\w|\s|,]*([\(](?<cusStr>[^)]+)[\)])?[\s]*([\:](?<capBit>[\d]+))?[\s]*";

        //eg: MeasI pin = pin1,pin2               
        private const string RegMeasExpression =
            @"(?<MeasType>(Wi)*[(Meas)|(Src)]\S+)[\s]*(pin)?[\s]*=[\s]*(?<pin>(.*))";

        //eg: MeasC pin = JTAG(data):4 "testName"
        private const string RegMeasCPinValue =
            @"(?<pinName>[\w]+)[\s]*([\(](?<cusStr>[^)]+)[\)])?[\s]*([\:](?<capBit>[\d]+))?[\s]*(([""])(?<testName>([^""])+)([""]))?";

        //eg: calc (LPDP_TX1P(snV2)/LPDP_TX1P(sn2))(sn4)
        private const string RegCalc = @"[\)]\s*[\(](?<cusStr>[^)]+)[\)]";

        //(VDD_A,VDD_B)(sn3) or (VDD_A::VDD_B)(sn3)
        private const string RegPinName1 = @"^[\(](?<pinName>.+)[\)][\s]*([\(](?<storeName>[^)]+)[\)])";

        //VDD_B(sn3)
        private const string RegPinName2 = @"^(?<pinName>.+)[\s]*[\(](?<storeName>[^)]+)[\)]";

        //PinName:Vn:0.3,0.2
        private const string RegVDiffForceVn = @"^([\w]+)[:](Vn)[:]([.]|[\d])+[,]([.]|[\d])+";

        //PinName:Vp:0.3,0.2
        private const string RegVDiffForceVp = @"^([\w]+)[:](Vp)[:]([.]|[\d])+[,]([.]|[\d])+";

        //TBD,NA,N/A
        private const string RegNoLimit = @"^(TBD|NA|N[\/]A)$";

        //sweep(PinA:V:0.1) or sweepY(PinA:V:0.1)
        private const string RegSweepVoltage = @"sweep\w*\s*\((?<SweepStr>[^)]+)+\)";
        //XShmoo(PinName:Level:From,To,Step:ShmooType:ShmooAlgorithm,[jump_step]) or YShmoo(PinName:Level:From,To,Step:ShmooType:ShmooAlgorithm,[jump_step]) 

        private const string ErrorMsgWrongMeas = "Wrong measure data in 'Meas' column";
        private const string ErrorMsgWrongForceFormat = "Wrong format of force condition in Test Plan";
        private const string ErrorMsgWrongLimitValue = "Unrecognied limit value";

        private readonly TestPlanSheet _planSheet;
        private Dictionary<string, string> _forceConditionLabels;
        private HardIpPattern _pattern;

        public TestPlanPatParser(TestPlanSheet planSheet)
        {
            _planSheet = planSheet;
        }

        public void ConvertTpPatterns()
        {
            _forceConditionLabels = new Dictionary<string, string>();
            //Convert test plan raw data to HardIp patterns

            foreach (var patternRow in _planSheet.PatternRows)
                try
                {
                    //Replace force condition labels with actual values
                    ConvertForceLabelForPatternNew(patternRow);
                    ConvertForceLabelForPatternNew(patternRow.DdrExtraPat);
                    //
                    if (!string.IsNullOrEmpty(_planSheet.ForceStr))
                        patternRow.ForceCondition.ForceCondition += ";" + _planSheet.ForceStr;
                    //Convert test plan raw data to hardIp pattern data
                    _planSheet.PatternItems.Add(ConvertTpPattern(patternRow));
                }
                catch (Exception e)
                {
                    Response.Report(e.ToString(), MessageLevel.Error, 0);
                }


            //do some extra settings for hardip patterns
            for (var i = 0; i < _planSheet.PatternItems.Count; i++)
            {
                //Move force condition from MeasV to MeasE
                ConvertMeasEForce(_planSheet.PatternItems[i]);
                //Get relaySetting
                var lastPatRelaySetting = new Dictionary<string, string>();
                if (i > 0)
                    lastPatRelaySetting = _planSheet.PatternItems[i - 1].NewRelaySetting;
                SearchInfo.GetRelaySetting(_planSheet.PatternItems[i], lastPatRelaySetting);
            }
        }

        private HardIpPattern ConvertTpPattern(PatternRow patternRow)
        {
            if (patternRow == null)
                return null;

            _pattern = new HardIpPattern();
            _pattern.SheetName = _planSheet.SheetName;
            _pattern.RowNum = patternRow.RowNum;
            _pattern.ColumnNum = patternRow.PatternColumnNum;
            _pattern.Pattern = patternRow.Pattern;
            _pattern.TtrStr = patternRow.TtrStr;
            _pattern.NoBinOutStr = patternRow.NoBinOutStr;
            _pattern.DupIndex = patternRow.DupIndex;
            _pattern.MiscInfo = patternRow.MiscInfo;
            _pattern.ForceCondition = patternRow.ForceCondition;
            _pattern.ForceCondChar = patternRow.ForceConditionChar;
            _pattern.RegisterAssignment = patternRow.RegisterAssignment.Trim(';');

            //Read measure sequence info from test plan
            _pattern.TestPlanSequences = ReadSequenceInfoNew(patternRow);

            //Post test force condition
            _pattern.InterposePostTest = patternRow.PostPatForceCondition;

            #region Convert force condition str which assigned in the same row as pattern

            //Get Level:XXX
            _pattern.LevelUsed = patternRow.ForceCondition.GetLevelSetting();
            //Get AC setting
            _pattern.AcUsed = patternRow.ForceCondition.GetAcSetting();
            //Get AC Selector
            _pattern.AcSelectorUsed = patternRow.ForceCondition.GetAcSelector();
            //Get DC Category
            _pattern.DcCategory = patternRow.ForceCondition.GetDcCategory();
            //Get AC Category
            _pattern.AcCategory = patternRow.ForceCondition.GetAcCategory();
            //Get DC Selector
            _pattern.DcSelectorUsed = patternRow.ForceCondition.GetDcSelector();
            //Get MCG mode setting
            _pattern.TimeSetUsed.McgSetting = patternRow.ForceCondition.GetMcgSetting();
            //Get interpose pre pattern force condition
            _pattern.ForceConditionList =
                DivideForceCondition(patternRow.ForceCondition.GetPrePatForceCondition(), patternRow.RowNum);

            #endregion

            _pattern.TestPlanSequencesRf = ReadSequenceInfoRf(patternRow, _pattern.MeasPins);
            //Convert measure pins' information
            _pattern.OriMeasPins.AddRange(GetMeasPins(patternRow));
            _pattern.MeasPins.AddRange(_pattern.OriMeasPins);
            //CalcEqn
            _pattern.CalcEqn = GetPatternCalcEqn(_pattern.MeasPins);
            //RF Interpose
            _pattern.RfInterPose = patternRow.RfInterPose;

            //Loop Flow loop
            _pattern.SweepVoltage = _GetSweepVoltage(patternRow.MiscInfo);
            //Shmoo
            _pattern.Shmoo = patternRow.ForceCondition.GetShmoo(_pattern,
                !string.IsNullOrEmpty(patternRow.ForceConditionChar)
                    ? patternRow.ForceConditionChar
                    : patternRow.ForceCondition.ForceCondition,
                CommonGenerator.GetSubBlockName(_pattern.Pattern.GetLastPayload(), _pattern.MiscInfo, ""));

            //ADC Autogen Added
            _pattern.AnalogSrc = _AnalyzeAnalog_SrcInfo(patternRow.AnalogSetup);
            _pattern.AnalogCap = _AnalyzeAnalog_CapInfo(patternRow.MiscInfo);
            _pattern.UseDeferLimit = _pattern.AnalogCap != null;
            _pattern.TestName = patternRow.SpecifyTestName;

            //Extra Pattern for DDR_Eye use
            _pattern.ExtraPattern = ConvertTpPatternExtra(patternRow.DdrExtraPat);
            _pattern.IsNonHardIpBlock = _pattern.SheetName.StartsWith("plldebug_", StringComparison.OrdinalIgnoreCase);
            if (_pattern.ExtraPattern != null)
            {
                _pattern.MeasPins.AddRange(_pattern.ExtraPattern.MeasPins);
                _pattern.ExtraPattern.MiscInfo = patternRow.MiscInfo;
            }

            return _pattern;
        }

        private HardIpPattern ConvertTpPatternExtra(PatternRow patternRow)
        {
            if (patternRow == null)
                return null;
            var patternExtra = new HardIpPattern();
            patternExtra.SheetName = _planSheet.SheetName;
            patternExtra.RowNum = patternRow.RowNum;
            patternExtra.ColumnNum = patternRow.PatternColumnNum;
            patternExtra.Pattern = patternRow.Pattern;
            patternExtra.TtrStr = patternRow.TtrStr;
            patternExtra.NoBinOutStr = patternRow.NoBinOutStr;
            patternExtra.DupIndex = patternRow.DupIndex;
            patternExtra.MiscInfo = patternRow.MiscInfo;
            patternExtra.RegisterAssignment = patternRow.RegisterAssignment.Trim(';');

            //Read measure sequence info from test plan
            patternExtra.TestPlanSequences = ReadSequenceInfoNew(patternRow);
            //Post test force condition
            patternExtra.InterposePostTest = patternRow.PostPatForceCondition;

            #region Convert force condition str which assigned in the same row as pattern

            //Get Level:XXX
            patternExtra.LevelUsed = patternRow.ForceCondition.GetLevelSetting();
            //Get AC setting
            patternExtra.AcUsed = patternRow.ForceCondition.GetAcSetting();
            //Get MCG mode setting
            patternExtra.TimeSetUsed.McgSetting = patternRow.ForceCondition.GetMcgSetting();
            //Get interpose pre pattern force condition
            patternExtra.ForceConditionList =
                DivideForceCondition(patternRow.ForceCondition.GetPrePatForceCondition(), patternRow.RowNum);

            #endregion

            //Convert measure pins' information
            patternExtra.MeasPins.AddRange(GetMeasPins(patternRow));
            //CalcEqn
            patternExtra.CalcEqn = GetPatternCalcEqn(patternExtra.MeasPins); //_pattern.ExtraPattern.MeasPins

            #region Roger add for multiple init pat

            //_patternExtra.MultipleInit = patternRow.IsMultipleInitWithPatSet;
            //_patternExtra.InitList = patternRow.InitList;

            #endregion

            if (patternExtra.ExtraPattern != null)
                patternExtra.MeasPins.AddRange(_pattern.ExtraPattern.MeasPins);
            return patternExtra;
        }

        private void ConvertForceLabelForPatternNew(PatternRow patternRow)
        {
            if (patternRow == null)
                return;

            //Handle interpose pre pattern force condition
            patternRow.ForceCondition.ForceCondition =
                ConvertForceLabelToValue(patternRow.ForceCondition.ForceCondition);
            //Handle interpose pre measure force condition
            foreach (var patChildRow in patternRow.PatChildRows)
            foreach (var tpRow in ((PatSubChildRow) patChildRow).TpRows)
                tpRow.ForceCondition = ConvertForceLabelToValue(tpRow.ForceCondition);
            //Handle post pattern force condition
            patternRow.PostPatForceCondition = ConvertForceLabelToValue(patternRow.PostPatForceCondition);
        }

        private string ConvertForceLabelToValue(string forceStr)
        {
            //Excluding syntax like XShmoo(PinName:Level:From,To,Step:ShmooType:ShmooAlgorithm,[jump_step])
            if (Regex.IsMatch(forceStr, HardIpConstData.RegShmoo, RegexOptions.IgnoreCase))
                return forceStr;
            if (Regex.IsMatch(forceStr, RegSweepVoltage, RegexOptions.IgnoreCase))
                return forceStr;

            //find all the label definition, store the value and replace it. 
            //eg: label1(IO1:V:1;IO2:V:2), store "[label1, IO1:V:1;IO2:V:2]"
            //and replace label1(IO1:V:1;IO2:V:2) with "IO1:V:1;IO2:V:2"
            var matches = Regex.Matches(forceStr, RegForceConditionLabelDefine, RegexOptions.IgnoreCase);
            foreach (Match match in matches)
            {
                var labelName = match.Groups["label"].ToString();
                var labelValue = match.Groups["value"].ToString();
                if (!_forceConditionLabels.ContainsKey(labelName))
                    _forceConditionLabels.Add(labelName, labelValue);
                else
                    _forceConditionLabels[labelName] = labelValue;
                forceStr = forceStr.Replace(match.Value, labelValue);
            }

            forceStr = forceStr.Trim(';');

            //replace forceCondition label to actual value
            //eg: label1
            //should replace label1 with "IO1:V:1;IO2:V:2"
            var newStr = "";
            var splitStrArr = forceStr.Split(';');
            for (var i = 0; i < splitStrArr.Length; i++)
                if (_forceConditionLabels.ContainsKey(splitStrArr[i]))
                    newStr += _forceConditionLabels[splitStrArr[i]] + ";";
                else
                    newStr += splitStrArr[i] + ";";
            return newStr.Trim(';');
        }

        private List<ForceCondition> DivideForceCondition(string forceStr, int rowNum)
        {
            var forceConditionList = new List<ForceCondition>();
            //to support differential pin in force condition, we should replace "::" to "&" first
            var forceList = Regex.Replace(forceStr, "::", "&").Split(';').ToList();
            forceList.RemoveAll(string.IsNullOrEmpty);
            if (forceList.Count == 0)
                return forceConditionList;

            //Support new format for Vdiff2 pin's forceCondition
            ConvertVdiff2ForceFormat(forceList);
            //Check format
            CheckForceStrFormat(forceList, rowNum);
            //Convert force condition str to struct
            var forcePinList = ConvertForceStrToStruct(forceList);
            //Separate force condition, like USB_DP:V:5,6 to USB_DP:V:5 and USB_DP:V:6
            return SeparateForceCondition(forcePinList);
        }

        private List<ForceCondition> SeparateForceCondition(List<ForcePin> forcePinList)
        {
            var forceConditionList = new List<ForceCondition>();
            if (forcePinList.Count == 0)
                return forceConditionList;
            //find force condition count
            var count = 1;
            foreach (var forcePin in forcePinList)
                if (!Regex.IsMatch(forcePin.ForceType, "TERM|Sweep", RegexOptions.IgnoreCase))
                {
                    var forceCount = forcePin.ForceValue.Split(',').Length;
                    count = forceCount > count ? forceCount : count;
                    forcePin.ForceCnt = forceCount;
                }

            //separate to different force conditions
            if (count == 1)
            {
                var forceCondition = new ForceCondition();
                forceCondition.ForcePins.AddRange(forcePinList);
                //forceCondition.ForcePins.Sort((x, y) => string.Compare(x.PinName, y.PinName));
                forceConditionList.Add(forceCondition);
                return forceConditionList;
            }

            for (var i = 0; i < count; i++)
            {
                var forceCondition = new ForceCondition();
                foreach (var forcePin in forcePinList)
                {
                    var newForcePin = forcePin.DeepClone();
                    newForcePin.ForceValue = forcePin.ForceValue.Split(',').Length > i
                        ? forcePin.ForceValue.Split(',')[i]
                        : forcePin.ForceValue.Split(',')[0];
                    forceCondition.ForcePins.Add(newForcePin);
                }

                forceConditionList.Add(forceCondition);
                forceCondition.ForcePins.Sort((x, y) => string.CompareOrdinal(x.PinName, y.PinName));
            }

            return forceConditionList;
        }

        private void CheckForceStrFormat(List<string> forceList, int rowNum)
        {
            foreach (var forceStr in forceList)
            {
                if (Regex.IsMatch(forceStr, HardIpConstData.RegShmoo, RegexOptions.IgnoreCase)) continue;
                if (Regex.IsMatch(forceStr, RegSweepVoltage, RegexOptions.IgnoreCase)) continue;

                var forceArr = forceStr.Split(':').ToList();
                if (forceArr.Count < 2)
                {
                    EpplusErrorManager.AddError(HardIpErrorType.WrongForceCondition, ErrorLevel.Error,
                        _planSheet.SheetName, rowNum, _planSheet.ForceIndex, ErrorMsgWrongForceFormat, forceStr);
                    continue;
                }

                if (HardIpDataMain.TestPlanData.AllJobs.Exists(s =>
                    Regex.IsMatch(s, @"^" + forceArr[forceArr.Count - 1], RegexOptions.IgnoreCase)))
                    forceArr.RemoveAt(forceArr.Count - 1);

                if (!(forceArr.Count == 2 || forceArr.Count == 3))
                    EpplusErrorManager.AddError(HardIpErrorType.WrongForceCondition, ErrorLevel.Error,
                        _planSheet.SheetName, rowNum, _planSheet.ForceIndex, ErrorMsgWrongForceFormat, forceStr);
            }
        }

        private List<ForcePin> ConvertForceStrToStruct(List<string> forceList)
        {
            var forcePinList = new List<ForcePin>();
            if (forceList.Count == 0)
                return forcePinList;
            forceList = forceList.Distinct().ToList();
            foreach (var forceStr in forceList)
            {
                if (Regex.IsMatch(forceStr, HardIpConstData.RegShmoo, RegexOptions.IgnoreCase)) continue;
                if (Regex.IsMatch(forceStr, RegSweepVoltage, RegexOptions.IgnoreCase)) continue;
                var forceArr = forceStr.Split(':').ToList();
                if (forceArr.Count < 2) continue;
                var forcePin = new ForcePin();
                //Read job form force str
                if (HardIpDataMain.TestPlanData.AllJobs.Exists(s =>
                    Regex.IsMatch(s, @"^" + forceArr[forceArr.Count - 1], RegexOptions.IgnoreCase)))
                {
                    forcePin.ForceJob = forceArr[forceArr.Count - 1];
                    forceArr.RemoveAt(forceArr.Count - 1);
                }

                //Read force condition format type
                //pin:XXX:XXX => Normal, pin:XXX => Others
                if (!(forceArr.Count == 2 || forceArr.Count == 3))
                {
                    // continue; // even it's syntax error, it should put into the force condition as well. Should not be empty! Marx 2017/05/02
                }

                forcePin.Type = forceArr.Count == 3 ? ForceConditionType.Normal : ForceConditionType.Others;
                //Read pin name
                forcePin.PinName = Regex.Replace(forceArr[0], "&", "::").ToUpper();
                //Read force value
                if (forcePin.Type == ForceConditionType.Normal)
                {
                    forcePin.ForceType = forceArr[1].ToUpper();
                    forcePin.ForceValue = forceArr[2]; //.ToUpper(); 20170428 Roger modify,"mV" -> "MV"
                }
                else
                {
                    forcePin.ForceValue = forceArr[1];
                    if (forceArr.Count != 2)
                        forcePin.ForceValue = forceStr.Substring(forceArr[0].Length + 1,
                            forceStr.Length - forceArr[0].Length - 1);
                }

                //split pinName
                foreach (var pin in forcePin.PinName.Split(','))
                {
                    var newPin = forcePin.DeepClone();
                    newPin.PinName = pin;
                    forcePinList.Add(newPin);
                }
            }

            return forcePinList;
        }

        private void ConvertVdiff2ForceFormat(List<string> forceList)
        {
            foreach (var force in forceList.ToArray())
            {
                //if vdiff2 pin forceCondition like "PinName:Vn:0.3,0.2", change to "PinName:V1n:0.3;PinName:V2n:0.2"
                if (Regex.IsMatch(force, RegVDiffForceVn, RegexOptions.IgnoreCase))
                {
                    var strArray = force.Split(':');
                    forceList.Add(strArray[0] + ":V1n:" + strArray[2].Split(',')[0]);
                    forceList.Add(strArray[0] + ":V2n:" + strArray[2].Split(',')[1]);
                    forceList.Remove(force);
                    continue;
                }

                //if vdiff2 pin forceCondition like "PinName:Vp:0.3,0.2", change to "PinName:V1p:0.3;PinName:V2p:0.2"
                if (Regex.IsMatch(force, RegVDiffForceVp, RegexOptions.IgnoreCase))
                {
                    var strArray = force.Split(':');
                    forceList.Add(strArray[0] + ":V1p:" + strArray[2].Split(',')[0]);
                    forceList.Add(strArray[0] + ":V2p:" + strArray[2].Split(',')[1]);
                    forceList.Remove(force);
                }
            }
        }

        private List<MeasPin> GetMeasPins(PatternRow patternRow)
        {
            var measPins = GetPreMeasPins(patternRow);

            measPins = GetDivideMeasPins(patternRow, measPins);

            measPins = GetPostMeasPins(measPins);

            return measPins;
        }

        private List<MeasPin> GetPreMeasPins(PatternRow patternRow)
        {
            var measPins = new List<MeasPin>();
            foreach (var measRow in patternRow.PatChildRows)
            foreach (var tpMeasRow in ((PatSubChildRow) measRow).TpRows)
                measPins.Add(GetMeasPinInfo(tpMeasRow, patternRow));
            return measPins;
        }

        private MeasPin GetMeasPinInfo(TestPlanRow tpMeasRow, PatternRow patternRow)
        {
            var measPin = new MeasPin();
            measPin.MiscInfo = tpMeasRow.MiscInfo;
            measPin.RowNum = tpMeasRow.RowNum;
            var measStr = tpMeasRow.Meas.Replace("\n", "").Replace("\t", "").Trim();
            measPin.MeasType = GetMeasType(measStr, patternRow);
            if (measPin.MeasType != MeasType.MeasCalc && measPin.MeasType != MeasType.MeasLimit)
                measStr = Regex.Match(measStr, RegMeasExpression, RegexOptions.IgnoreCase).Groups["pin"].ToString()
                    .Trim(',').Trim();
            measPin.PinName = GetMeasPinName(measStr, measPin.MeasType);
            measPin.RfInstrumentSetup = tpMeasRow.RfInstrumentSetup;
            measPin.CusStr = GetCusStr(measStr, measPin.MeasType);
            if (measPin.MeasType == MeasType.MeasCalc && measPin.CusStr != "")
                measStr = measStr.Replace("(" + measPin.CusStr + ")",
                    ""); // if Calc_Eqn contain dictionary info, move to CusStr for further assignment 
            measPin.CapBit = GetCapBit(measStr, measPin.MeasType);
            measPin.TestName = tpMeasRow.TestName != "" ? tpMeasRow.TestName : GetTestName(measStr, measPin.MeasType);

            measPin.CalcEqn = GetCalcEqnForPin(measStr, tpMeasRow.MiscInfo, measPin.MeasType);
            measPin.RepeatCount = SearchInfo.GetRepeatLimitCount(tpMeasRow.MiscInfo);
            measPin.SkipUnit = GetSkipUnit(measStr, measPin.MeasType);
            //Convert limits to Hv,Lv,Nv limit
            var limits = GetLimits(tpMeasRow);
            measPin.MeasLimitsH = limits[0];
            measPin.MeasLimitsL = limits[1];
            measPin.MeasLimitsN = limits[2];
            //measPin.HighLimit = tpMeasRow.Limits.First().HiLimit;
            //measPin.LowLimit = tpMeasRow.Limits.First().LoLimit;

            #region Get measure pin's force condition and sequence index

            if (_pattern.TestPlanSequences.Count > 0)
            {
                var sequence =
                    _pattern.TestPlanSequences.Find(s =>
                        s.StartRow <= tpMeasRow.RowNum && s.EndRow >= tpMeasRow.RowNum);
                if (sequence != null && measPin.MeasType != MeasType.MeasC && measPin.MeasType != MeasType.MeasLimit &&
                    measPin.MeasType != MeasType.MeasCalc)
                {
                    measPin.ForceConditions =
                        DivideForceCondition(string.Join(";", sequence.ForceCondition), tpMeasRow.RowNum);
                    //default sequenceIndex=0, if test plan merged, will set the actual sequence index
                    measPin.SequenceIndex = sequence.SeqIndex;
                }
                //else if ((LocalSpecs.Device == DeviceEnum.LCD || LocalSpecs.Device ==DeviceEnum.RF) && measPin.MeasType == MeasType.MeasCalc)
                else if (measPin.MeasType == MeasType.MeasCalc)
                {
                    if (sequence != null)
                    {
                        measPin.ForceConditions =
                            DivideForceCondition(string.Join(";", sequence.ForceCondition), tpMeasRow.RowNum);
                        //default sequenceIndex=0, if test plan merged, will set the actual sequence index
                        measPin.SequenceIndex = sequence.SeqIndex;
                    }
                }
            }
            else if (_pattern.TestPlanSequencesRf.Count > 0)
            {
                var sequence = _pattern.TestPlanSequencesRf.Find(s =>
                    s.StartRow <= tpMeasRow.RowNum && s.EndRow >= tpMeasRow.RowNum);
                if (sequence != null && measPin.MeasType != MeasType.MeasC && measPin.MeasType != MeasType.MeasLimit &&
                    measPin.MeasType != MeasType.MeasCalc)
                {
                    measPin.ForceConditions =
                        DivideForceCondition(string.Join(";", sequence.ForceCondition), tpMeasRow.RowNum);
                    //default sequenceIndex=0, if test plan merged, will set the actual sequence index
                    measPin.SequenceIndex = sequence.SeqIndex;
                }

                //else if (LocalSpecs.Device == DeviceEnum.LCD && measPin.MeasType == MeasType.MeasCalc)
                //{
                //    measPin.ForceConditions = DivideForceCondition(string.Join(";", sequence.ForceCondition), tpMeasRow.RowNum);
                //    //default sequenceIndex=0, if test plan merged, will set the actual sequence index
                //    measPin.SequenceIndex = sequence.SeqIndex;
                //}
            }
            else
            {
                measPin.ForceConditions = DivideForceCondition(tpMeasRow.ForceCondition, tpMeasRow.RowNum);
            }

            #endregion

            //FW use
            measPin.InterPoseFunc = tpMeasRow.InterposeFunc;
            measPin.RfInterPose = tpMeasRow.RfInterpose;
            measPin.MeasWaitTime = GetMeasWaitTime(tpMeasRow);
            measPin.MeasRange = GetMeasIRange(tpMeasRow);
            measPin.RfInstrumentSetup = GetRfInstrumentSetup(tpMeasRow.RfInstrumentSetup);
            if (string.IsNullOrEmpty(measPin.RfInstrumentSetup))
                if (tpMeasRow.MergeRowNumForMeas != 0 && tpMeasRow.MergeRowNumForMeas != tpMeasRow.RowNum)
                {
                    measPin.RowNumForMergeMeas = tpMeasRow.MergeRowNumForMeas;
                    measPin.MeasType = MeasType.MeasLimit;
                    measPin.SequenceIndex = -1;
                }

            return measPin;
        }

        private List<List<MeasLimit>> GetLimits(TestPlanRow tpMeasRow)
        {
            var limits = new List<List<MeasLimit>>();
            var limitsH = new List<MeasLimit>();
            var limitsL = new List<MeasLimit>();
            var limitsN = new List<MeasLimit>();
            limits.Add(limitsH);
            limits.Add(limitsL);
            limits.Add(limitsN);

            foreach (var measLimit in tpMeasRow.Limits)
            {
                var limitH = new MeasLimit(measLimit.JobName);
                var limitL = new MeasLimit(measLimit.JobName);
                var limitN = new MeasLimit(measLimit.JobName);
                limitsH.Add(limitH);
                limitsL.Add(limitL);
                limitsN.Add(limitN);

                //limit value contains HV,LV,NV
                //user may assign unit "µ", equals to "u"
                var hi = measLimit.HiLimit.Replace("µ", "u");
                var lo = measLimit.LoLimit.Replace("µ", "u");
                //if assign TBD, NA or N/A, means no limit
                if (Regex.IsMatch(hi.Trim(), RegNoLimit, RegexOptions.IgnoreCase))
                    hi = "";
                if (Regex.IsMatch(lo.Trim(), RegNoLimit, RegexOptions.IgnoreCase))
                    lo = "";
                // Set HeaderIndex
                limitH.LoHeaderIndex = measLimit.LoHeaderIndex;
                limitH.HiHeaderIndex = measLimit.HiHeaderIndex;
                limitL.LoHeaderIndex = measLimit.LoHeaderIndex;
                limitL.HiHeaderIndex = measLimit.HiHeaderIndex;
                limitN.LoHeaderIndex = measLimit.LoHeaderIndex;
                limitN.HiHeaderIndex = measLimit.HiHeaderIndex;

                //Set hi limit for Hv,Lv,Nv
                var hiArr = hi.Split(',');
                if (hiArr.Length == 3)
                {
                    limitH.HiLimit = hiArr[0];
                    limitL.HiLimit = hiArr[1];
                    limitN.HiLimit = hiArr[2];
                }

                if (hiArr.Length == 2)
                {
                    limitH.HiLimit = hiArr[0];
                    limitL.HiLimit = hiArr[1];
                    limitN.HiLimit = "";
                }

                if (hiArr.Length == 1)
                {
                    limitH.HiLimit = hiArr[0];
                    limitL.HiLimit = hiArr[0];
                    limitN.HiLimit = hiArr[0];
                }

                //Set lo limit for Hv,Lv,Nv
                var loArr = lo.Split(',');
                if (loArr.Length == 3)
                {
                    limitH.LoLimit = loArr[0];
                    limitL.LoLimit = loArr[1];
                    limitN.LoLimit = loArr[2];
                }

                if (loArr.Length == 2)
                {
                    limitH.LoLimit = loArr[0];
                    limitL.LoLimit = loArr[1];
                    limitN.LoLimit = "";
                }

                if (loArr.Length == 1)
                {
                    limitH.LoLimit = loArr[0];
                    limitL.LoLimit = loArr[0];
                    limitN.LoLimit = loArr[0];
                }

                //Flag error if has valid format
                if (hiArr.Length > 3 || loArr.Length > 3)
                    EpplusErrorManager.AddError(HardIpErrorType.WrongLimitValue, ErrorLevel.Error, _planSheet.SheetName,
                        tpMeasRow.RowNum, ErrorMsgWrongLimitValue);
            }

            return limits;
        }

        private string GetRfInstrumentSetup(string rfInstrumentSetup)
        {
            var dic = new Dictionary<string, string>();
            foreach (var item in rfInstrumentSetup.Split(';'))
                if (item.Contains('='))
                {
                    var name = item.Split('=')[0];
                    var value = item.Split('=')[1];
                    if (!dic.ContainsKey(name)) dic.Add(name, value);
                }

            var list = new List<string>();
            foreach (var item in dic)
                list.Add(item.Key + "=" + item.Value);
            return string.Join(";", list);
        }

        private string GetCalcEqnForPin(string pinValue, string pinMiscInfo, string measType)
        {
            var calEqnList = new List<string>();

            //eg: Calc "Add", sn1+sn2, gets "C:Add:sn1+sn2"..update example by raze with comma 20170630
            //eg: Calc "test1", VDD_A(sn1)+VDD_B(sn2), and measure type is I, gets "I:test1:VDD_A(sn1)+VDD_B(sn2)" ..update example by raze with comma 20170630
            if (measType == MeasType.MeasCalc)
            {
                var calcExp = Regex.Match(pinValue, RegCalcExpression, RegexOptions.IgnoreCase).Groups["expression"]
                    .ToString();

                //remove first "(" and last")" if exist this condition for VBT parsing 2017/6/26 Raze Edited
                if (calcExp[0] == '(' && calcExp[calcExp.Length - 1] == ')')
                    calcExp = calcExp.Substring(1, calcExp.Length - 2);

                //string testName = Regex.Match(pinValue, RegCalcExpression, RegexOptions.IgnoreCase).Groups["testName"].ToString();
                var testName = Regex.Match(pinValue, RegCalcExpression, RegexOptions.IgnoreCase).Groups["testName"]
                    .ToString().Trim();
                if (testName.Contains(":"))
                    testName = testName.Substring(0, testName.IndexOf(":", StringComparison.Ordinal));
                calEqnList.Add(testName + ":" + calcExp);
                //return testName + ":" + calcExp;
            }
            //Eg: "Calc:Algorithm_A;CalcParameter:rd0,rd1" in miscInfo, gets "Alg::Algorithm_A(rd0,rd1)"

            if (!string.IsNullOrEmpty(SearchInfo.GetCalculation(pinMiscInfo)))
                calEqnList.Add(SearchInfo.GetCalculation(pinMiscInfo));
            return string.Join(";", calEqnList);
            //return "";
        }

        private List<string> JudgeCalcEqnByJob(string calcExp)
        {
            var result = new List<string> {calcExp};
            //var regCalcEqn = @"((?<Pin>\w+)(\(\w+\)))";
            //var matchPins = Regex.Matches(calcExp, regCalcEqn, RegexOptions.IgnoreCase);
            //var isNeedExtraCp = false;
            //var isNeedExtraFt = false;
            //foreach (var matchPin in matchPins)
            //{
            //    string cpNetName = "";
            //    string ftNetName = "";
            //    if (cpNetName != ftNetName)
            //    {
            //        if (cpNetName != "" && cpNetName != "N/A")
            //        {
            //            isNeedExtraCp = true;
            //        }
            //        if (ftNetName != "" && ftNetName != "N/A")
            //        {
            //            isNeedExtraFt = true;
            //        }
            //    }
            //}
            //if (isNeedExtraCp || isNeedExtraFt)
            //{
            //    result = new List<string>();
            //    if (isNeedExtraCp)
            //    {
            //        var eqn = calcExp;
            //        var isNeedGenEqn = true;
            //        foreach (var matchPin in matchPins)
            //        {
            //            string cpNetName = "";
            //            var pinName =
            //                Regex.Match(matchPin.ToString(), regCalcEqn, RegexOptions.IgnoreCase).Groups["Pin"].Value;
            //            //HardIpDataMain.TempResultData.PowerMerge.GetCpFtNetName(pinName, ref cpNetName, ref ftNetName);
            //            if (cpNetName != "" && cpNetName != "N/A")
            //            {
            //                var newMatchPin = matchPin.ToString().Replace(pinName, cpNetName);
            //                eqn = eqn.Replace(matchPin.ToString(), newMatchPin);
            //            }
            //            else
            //                isNeedGenEqn = false;

            //        }
            //        if (isNeedGenEqn)
            //            result.Add("CP=" + eqn);
            //    }
            //    if (isNeedExtraFt)
            //    {
            //        var eqn = calcExp;
            //        var isNeedGenEqn = true;
            //        foreach (var matchPin in matchPins)
            //        {
            //            string ftNetName = "";
            //            var pinName =
            //                Regex.Match(matchPin.ToString(), regCalcEqn, RegexOptions.IgnoreCase).Groups["Pin"].Value;
            //            //HardIpDataMain.TempResultData.PowerMerge.GetCpFtNetName(pinName, ref cpNetName, ref ftNetName);
            //            if (ftNetName != "" && ftNetName != "N/A")
            //            {
            //                var newMatchPin = matchPin.ToString().Replace(pinName, ftNetName);
            //                eqn = eqn.Replace(matchPin.ToString(), newMatchPin);
            //            }
            //            else
            //                isNeedGenEqn = false;
            //        }
            //        if (isNeedGenEqn)
            //            result.Add("FT=" + eqn);
            //    }
            //}


            return result;
        }

        private string GetTestName(string pinValue, string measType)
        {
            //eg: Limits "test1", test name is "test1"
            if (measType == MeasType.MeasLimit)
                return Regex.Match(pinValue.Replace(",", " "), RegLimitsExpression, RegexOptions.IgnoreCase)
                    .Groups["testName"].ToString().Trim();
            //eg: Calc "add" sn1+sn2, test name is "add"
            if (measType == MeasType.MeasCalc)
            {
                var testName = Regex.Match(pinValue.Replace(",", " "), RegCalcExpression, RegexOptions.IgnoreCase)
                    .Groups["testName"].ToString().Trim();
                if (testName.Contains(":"))
                    testName = testName.Substring(0, testName.IndexOf(":", StringComparison.Ordinal));
                return testName;
            }

            //Others => eg: MeasC pin = JTAG(data):4 "testName" or  Limits "Cal_A" , (out1):4 or Limits IO1(out2)
            return Regex.Match(pinValue, RegTestNameExpression, RegexOptions.IgnoreCase).Groups["testName"].ToString()
                .Trim();
        }

        private string GetSkipUnit(string pinValue, string measType)
        {
            var skipUnit = "";
            if (measType == MeasType.MeasCalc)
            {
                var testName = Regex.Match(pinValue.Replace(",", " "), RegCalcExpression, RegexOptions.IgnoreCase)
                    .Groups["testName"].ToString().Trim();
                if (testName.Contains(":"))
                {
                    var idx = testName.IndexOf(":", StringComparison.Ordinal);
                    skipUnit = testName.Substring(idx + 1, testName.Length - idx - 1);
                }
            }

            return skipUnit;
        }

        private string GetCapBit(string pinValue, string measType)
        {
            //eg: MeasC pin = JTAG(data):4 "testName", ccpBit is "4"
            if (measType == MeasType.MeasC)
                return Regex.Match(pinValue, RegMeasCPinValue, RegexOptions.IgnoreCase).Groups["capBit"].ToString()
                    .Trim();
            //eg: Limits "Cal_A" , (out1):4, 
            if (measType == MeasType.MeasLimit)
                return Regex.Match(pinValue, RegLimitsExpression, RegexOptions.IgnoreCase).Groups["capBit"].ToString()
                    .Trim();
            return "";
        }

        private string GetCusStr(string pinValue, string measType)
        {
            //eg: MeasC pin = JTAG(data):4 "testName", cusStr is "data"
            if (measType == MeasType.MeasC)
                return Regex.Match(pinValue, RegMeasCPinValue, RegexOptions.IgnoreCase).Groups["cusStr"].ToString()
                    .Trim();

            if (measType == MeasType.MeasCalc)
            {
                if (Regex.IsMatch(pinValue, RegCalc, RegexOptions.IgnoreCase))
                    return Regex.Match(pinValue, RegCalc, RegexOptions.IgnoreCase).Groups["cusStr"].ToString().Trim();
            }
            //eg: Limits "Cal_A" , (out1):4, cusStr is "out1" or eg:  Limits IO1(out2), cusStr is "out2"
            else if (measType == MeasType.MeasLimit)
            {
                if (Regex.IsMatch(pinValue, RegLimitsExpression, RegexOptions.IgnoreCase))
                    return Regex.Match(pinValue, RegLimitsExpression, RegexOptions.IgnoreCase).Groups["cusStr"]
                        .ToString().Trim();
            }
            //eg:MeasI pin=(VDDA,VDDB)(sn1), cusStr is "sn1",MeasF Pin = (IO1, IO3_P::IO3_N) (sn2)
            else
            {
                if (Regex.IsMatch(pinValue, RegPinName1, RegexOptions.IgnoreCase))
                    return Regex.Match(pinValue, RegPinName1, RegexOptions.IgnoreCase).Groups["storeName"].ToString()
                        .Trim();
                //eg:MeasI pin=VDDA(sn1), cusStr is "sn1"
                if (Regex.IsMatch(pinValue, RegPinName2, RegexOptions.IgnoreCase))
                    return Regex.Match(pinValue, RegPinName2, RegexOptions.IgnoreCase).Groups["storeName"].ToString()
                        .Trim();
            }

            return "";
        }

        private string GetMeasPinName(string pinValue, string measType)
        {
            if (string.IsNullOrEmpty(measType))
                return "";
            //Limits, Calc => eg: Limits "Cal_A" , (out1):4 or Limits IO1(out2)
            if (measType == MeasType.MeasCalc || measType == MeasType.MeasLimit) return HardIpConstData.FakePin;

            //eg: MeasC pin = JTAG(data):4 "testName", pin name is "JTAG"
            if (measType == MeasType.MeasC)
                return Regex.Match(pinValue, RegMeasCPinValue, RegexOptions.IgnoreCase).Groups["pinName"].ToString()
                    .Trim().ToUpper();
            //eg:MeasI pin=(VDDA,VDDB)(sn1), pin name is "VDDA,VDDB" or MeasFDIFF Pin =(DDR3_DQS_P0::DDR3_DQS_N0)(F1)
            if (Regex.IsMatch(pinValue, RegPinName1, RegexOptions.IgnoreCase))
                return Regex.Match(pinValue, RegPinName1, RegexOptions.IgnoreCase).Groups["pinName"].ToString().Trim()
                    .ToUpper();
            //eg:MeasI pin= VDDA(sn1), pin name is "VDDA"
            if (Regex.IsMatch(pinValue, RegPinName2, RegexOptions.IgnoreCase))
                return Regex.Match(pinValue, RegPinName2, RegexOptions.IgnoreCase).Groups["pinName"].ToString().Trim()
                    .ToUpper();
            //eg: MeasVdiff pin = pin_diff_Grp, this differential pin group exist in pin map, convert to "pin_P::pin_N"
            //because user can not write differential pin group name to patInfo, but do it in test plan.
            if (measType == MeasType.MeasVdiff || measType == MeasType.MeasIdiff || measType == MeasType.MeasVocm)
            {
                var name = Regex.Match(pinValue, RegPineNameExpression, RegexOptions.IgnoreCase).Groups["pinName"]
                    .ToString().Trim().ToUpper();
                if (name == "")
                    name = pinValue;
                return DataConvertor.ConvertDifferentialPinGroup(name);
            }

            if (Regex.IsMatch(pinValue, RegPineNameExpression, RegexOptions.IgnoreCase))
                return Regex.Match(pinValue, RegPineNameExpression, RegexOptions.IgnoreCase).Groups["pinName"]
                    .ToString().Trim().ToUpper();
            return pinValue.ToUpper();
        }

        private string GetMeasType(string measStr, PatternRow patternRow)
        {
            //Limit
            if (Regex.IsMatch(measStr.Replace(",", " "), RegLimitsExpression, RegexOptions.IgnoreCase))
                return MeasType.MeasLimit;
            //Calc
            if (Regex.IsMatch(measStr, RegCalcExpression, RegexOptions.IgnoreCase)) return MeasType.MeasCalc;
            //Meas
            if (Regex.IsMatch(measStr, RegMeasExpression, RegexOptions.IgnoreCase))
            {
                var type = Regex.Match(measStr, RegMeasExpression, RegexOptions.IgnoreCase).Groups["MeasType"]
                    .ToString();
                var realType =
                    HardIpDataMain.TestPlanData.MeasTypes.Find(s => s.Equals(type, StringComparison.OrdinalIgnoreCase));
                return realType ?? type;
            }

            if (Regex.IsMatch(measStr, MeasType.WiMeas, RegexOptions.IgnoreCase))
                return MeasType.WiMeas;
            if (Regex.IsMatch(measStr, MeasType.WiSrc, RegexOptions.IgnoreCase))
                return MeasType.WiSrc;
            if (Regex.IsMatch(measStr, MeasType.MeasWait, RegexOptions.IgnoreCase))
                return MeasType.MeasWait;
            if (Regex.IsMatch(measStr, MeasType.MeasN, RegexOptions.IgnoreCase))
                return MeasType.MeasN;

            if (string.IsNullOrEmpty(measStr))
                return "";

            EpplusErrorManager.AddError(HardIpErrorType.WrongMeasContent, ErrorLevel.Error, _planSheet.SheetName,
                patternRow.RowNum, ErrorMsgWrongMeas, patternRow.Pattern.TestPlanPatternName);
            return "";
        }

        private List<MeasPin> GetDivideMeasPins(PatternRow patternRow, List<MeasPin> measPins)
        {
            measPins = GetMeasPinsByName(measPins);

            #region repeatLimit

            var count = SearchInfo.GetRepeatLimitCount(patternRow.MiscInfo);
            if (measPins.Count > 0) measPins = GetMeasPinsByRepeat(measPins, count);

            #endregion

            measPins = GetMeasPinsByPowerMerge(measPins);

            return measPins;
        }

        private List<MeasPin> GetMeasPinsByName(List<MeasPin> measPins)
        {
            var measPinList = new List<MeasPin>();
            foreach (var measPin in measPins)
            foreach (var pinName in measPin.PinName.Split(','))
            {
                var newPin = measPin.DeepClone();
                newPin.PinName = pinName.Trim();
                newPin.PinCount = GetPinCount(pinName.Trim());
                newPin.VisitedTime = newPin.PinCount;
                measPinList.Add(newPin);
            }

            return measPinList;
        }

        private int GetPinCount(string pinNames)
        {
            var count = 0;
            foreach (var pin in pinNames.Split(','))
            {
                if (pin == "")
                    continue;
                if (pin.Contains("::"))
                {
                    count = count + 2;
                }
                else if (HardIpDataMain.TestPlanData.PinGroupList.ContainsKey(pin))
                {
                    var pinGroupList = SearchInfo.DecomposeGroups(pin);
                    count += pinGroupList.Count;
                }
                else
                {
                    count++;
                }
            }

            return count;
        }

        private List<MeasPin> GetMeasPinsByRepeat(List<MeasPin> measPins, int count)
        {
            var testPlanSeq = new List<TestPlanSequence>();
            var seqMax = measPins.Max(p => p.SequenceIndex);
            var newMeasPins = new List<MeasPin>();
            var j = 0;
            if (count > 0)
            {
                for (var i = 0; i <= count; i++)
                {
                    j = 0;
                    foreach (var measPin in measPins)
                    {
                        if (_pattern.TestPlanSequences.Count > 0)
                            testPlanSeq.Add(_pattern.TestPlanSequences[j]);
                        var newPin = measPin.DeepClone();
                        newPin.SequenceIndex += seqMax * i;
                        newPin.CusStr += string.Format("_{0}", i);
                        newMeasPins.Add(newPin);
                        j++;
                    }
                }

                _pattern.TestPlanSequences = testPlanSeq;
                return newMeasPins;
            }

            var subI = 0;
            foreach (var measPin in measPins)
            {
                if (_pattern.TestPlanSequences.Count > 0 && j < _pattern.TestPlanSequences.Count)
                    testPlanSeq.Add(_pattern.TestPlanSequences[j]);
                var newPin = measPin.DeepClone();
                newPin.SequenceIndex += subI;
                newMeasPins.Add(newPin);
                if (measPin.RepeatCount > 0)
                {
                    newPin.CusStr += string.Format("_{0}", 0);
                    for (var k = 1; k <= measPin.RepeatCount; k++)
                    {
                        if (_pattern.TestPlanSequences.Count > 0)
                            testPlanSeq.Add(_pattern.TestPlanSequences[j]);
                        newPin = measPin.DeepClone();
                        newPin.SequenceIndex += ++subI;
                        newPin.CusStr += string.Format("_{0}", k);
                        newMeasPins.Add(newPin);
                    }
                }

                j++;
            }

            _pattern.TestPlanSequences = testPlanSeq;
            return newMeasPins;
        }

        private List<MeasPin> GetMeasPinsByPowerMerge(List<MeasPin> measPins)
        {
            //if (HardIpDataMain.TempResultData.PowerMerge != null && measPins.Count != 0)
            //{
            //    var newMeasPins = new List<MeasPin>();
            //    foreach (var pin in measPins)
            //    {
            //        string cpNetName = "";
            //        string ftNetName = "";
            //        HardIpDataMain.TempResultData.PowerMerge.GetCpFtNetName(pin.PinName, ref cpNetName, ref ftNetName);
            //        if (cpNetName != ftNetName)
            //        {
            //            var cpPin = pin.DeepClone();
            //            cpPin.PinName = "CP=" + cpNetName.Replace(",", ",CP=");
            //            newMeasPins.Add(cpPin);
            //            var ftPin = pin.DeepClone();
            //            ftPin.PinName = "FT=" + ftNetName.Replace(",", ",FT=");
            //            newMeasPins.Add(ftPin);
            //        }
            //        else
            //            newMeasPins.Add(pin);
            //    }
            //    return newMeasPins;
            //}
            return measPins;
        }

        private List<MeasPin> GetPostMeasPins(List<MeasPin> measPins)
        {
            foreach (var measPin in measPins)
            {
                measPin.CurrentRangeList = measPin.GetCurrentRangeList();
                measPin.CurrentRangeListH = measPin.GetCurrentRangeListByVoltage(measPin.MeasLimitsH);
                measPin.CurrentRangeListL = measPin.GetCurrentRangeListByVoltage(measPin.MeasLimitsL);
                measPin.CurrentRangeListN = measPin.GetCurrentRangeListByVoltage(measPin.MeasLimitsN);
            }

            return measPins;
        }

        private List<TestPlanSequence> ReadSequenceInfoRf(PatternRow patternRow, List<MeasPin> measPins)
        {
            var nonSeqItems = new List<string> {MeasType.MeasLimit, MeasType.MeasCalc, MeasType.MeasC};
            var testPlanSequence = new List<TestPlanSequence>();
            var sequenceIndex = 1;
            string measStr;
            if (!measPins.Exists(p => p.MeasType.Equals(MeasType.WiMeas) || p.MeasType.Equals(MeasType.WiSrc)))
                return testPlanSequence;

            foreach (var childRow in patternRow.PatChildRows)
            {
                var startRow = ((PatSubChildRow) childRow).TpRows.First().RowNum;
                var endRow = ((PatSubChildRow) childRow).TpRows.Last().RowNum;
                var sequence = new TestPlanSequence(startRow, endRow, sequenceIndex++);
                var isSeqItems = false;
                foreach (var subChildRow in ((PatSubChildRow) childRow).TpRows)
                {
                    measStr = subChildRow.Meas;
                    var forceStr = subChildRow.ForceCondition;
                    if (nonSeqItems.All(p => Regex.IsMatch(measStr, p, RegexOptions.IgnoreCase)))
                        continue;
                    isSeqItems = true;
                    if (!string.IsNullOrEmpty(forceStr))
                        sequence.ForceCondition.AddRange(forceStr.Split(';').ToList());
                }

                if (isSeqItems)
                    testPlanSequence.Add(sequence);
            }

            return testPlanSequence;
        }

        private List<TestPlanSequence> ReadSequenceInfoNew(PatternRow patternRow)
        {
            var nonSeqItems = new List<string> {MeasType.MeasLimit, MeasType.MeasCalc, MeasType.MeasC};
            var testPlanSequence = new List<TestPlanSequence>();
            try
            {
                var sequenceIndex = 1;
                string measStr;

                foreach (var childRow in patternRow.PatChildRows)
                {
                    var startRow = ((PatSubChildRow) childRow).TpRows.First().RowNum;
                    var endRow = ((PatSubChildRow) childRow).TpRows.Last().RowNum;
                    var sequence = new TestPlanSequence(startRow, endRow, 0);
                    var isSeqItems = false;
                    foreach (var subChildRow in ((PatSubChildRow) childRow).TpRows)
                    {
                        measStr = subChildRow.Meas;
                        var forceStr = subChildRow.ForceCondition;
                        if (nonSeqItems.Any(
                            p => Regex.IsMatch(measStr.Trim().Split(' ')[0], p, RegexOptions.IgnoreCase)))
                            continue;
                        isSeqItems = true;
                        if (!string.IsNullOrEmpty(forceStr))
                            sequence.ForceCondition.AddRange(forceStr.Split(';').ToList());
                    }

                    if (isSeqItems)
                    {
                        sequence.SeqIndex = sequenceIndex++;
                        testPlanSequence.Add(sequence);
                    }
                }
            }
            catch (Exception e)
            {
                Response.Report(e.ToString(), MessageLevel.Error, 0);
            }

            return testPlanSequence;
        }

        private string GetPatternCalcEqn(List<MeasPin> patMeasPins)
        {
            var calcEqnList = new List<string>();

            for (var i = 0; i < patMeasPins.Count; i++)
            {
                var measPin = patMeasPins[i];
                if (measPin.MeasType == MeasType.MeasLimit && !string.IsNullOrEmpty(measPin.CalcEqn) &&
                    !calcEqnList.Contains(measPin.CalcEqn))
                    calcEqnList.Add(measPin.CalcEqn);
                if (measPin.MeasType == MeasType.MeasCalc && !string.IsNullOrEmpty(measPin.CalcEqn))
                {
                    var calcType = GetMeasTypeForCalEqn(patMeasPins.GetRange(0, i + 1), measPin.CalcEqn);

                    calcEqnList.AddRange(string.IsNullOrEmpty(measPin.SkipUnit)
                        ? JudgeCalcEqnByJob(calcType + ":" + measPin.CalcEqn + ":" + measPin.CusStr)
                        : JudgeCalcEqnByJob(calcType + "," + measPin.SkipUnit + ":" + measPin.CalcEqn + ":" +
                                            measPin.CusStr));
                }
            }

            var patternCalcEqn = GetCalcEqnForPattern(_pattern.MiscInfo);
            if (!calcEqnList.Contains(patternCalcEqn) && patternCalcEqn != "") calcEqnList.Add(patternCalcEqn);
            return string.Join(";", calcEqnList);
        }

        private string GetMeasTypeForCalEqn(List<MeasPin> measPins, string calcEqn)
        {
            var type = "";
            for (var i = measPins.Count - 1; i >= 0; i--)
                if (measPins[i].CusStr != "" && calcEqn.ToLower().Contains(measPins[i].CusStr.ToLower()))
                    if (measPins[i].MeasType != MeasType.MeasCalc && measPins[i].MeasType != MeasType.MeasLimit)
                    {
                        type = measPins[i].MeasType[4].ToString();
                        break;
                    }

            return type;
        }

        private void ConvertMeasEForce(HardIpPattern pattern)
        {
            //var newMeasPins = new List<MeasPin>();
            var forceConditions = new List<ForceCondition>();
            foreach (var pin in pattern.MeasPins)
                if (pin.ForceConditions.Count > 0 && pin.MeasType == MeasType.MeasV)
                {
                    forceConditions = pin.ForceConditions;
                }
                else if (pin.ForceConditions.Count == 0 && pin.MeasType == MeasType.MeasV)
                {
                }
                else if (pin.MeasType == MeasType.MeasE)
                {
                    foreach (var condition in forceConditions)
                    {
                        var newCondition = new ForceCondition();
                        foreach (var forcePin in condition.ForcePins)
                            if (forcePin.PinName == pin.PinName)
                                newCondition.ForcePins.Add(forcePin);
                        if (newCondition.ForcePins.Count > 0)
                            pin.ForceConditions.Add(newCondition);
                    }
                }
        }

        private string GetCalcEqnForPattern(string pinMiscInfo)
        {
            if (Regex.IsMatch(pinMiscInfo, HardIpConstData.Calc + ":", RegexOptions.IgnoreCase))
            {
                var calcEqnInPin = SearchInfo.GetCalculation(pinMiscInfo);
                return calcEqnInPin == "" ? SearchInfo.GetCalculation(_pattern.MiscInfo) : calcEqnInPin;
            }

            return "";
        }

        private Dictionary<string, List<string>> _GetSweepVoltage(string pinMiscInfo)
        {
            var result = new Dictionary<string, List<string>>();
            foreach (var miscInfo in pinMiscInfo.Split(';'))
                if (Regex.IsMatch(miscInfo, RegSweepVoltage, RegexOptions.IgnoreCase))
                {
                    var str = Regex.Match(miscInfo, RegSweepVoltage, RegexOptions.IgnoreCase).Groups["SweepStr"]
                        .ToString();
                    var xy = Regex.IsMatch(miscInfo, "sweepY", RegexOptions.IgnoreCase) ? "Y" : "X";
                    foreach (var sweepItem in str.Split('|'))
                        if (sweepItem.Split(':').Length == 2)
                        {
                            if (!result.ContainsKey(xy))
                                result.Add(xy, new List<string> {sweepItem});
                            else
                                result[xy].Add(sweepItem);
                        }
                }

            return result;
        }

        private AnalogSrc _AnalyzeAnalog_SrcInfo(string setup)
        {
            if (string.IsNullOrEmpty(setup)) return null;
            var srcInfo = new AnalogSrc();

            foreach (var setupInfo in setup.Split(';').ToList())
            {
                var item = setupInfo.Split(':')[0];
                //for src
                if (item.Equals("srcid", StringComparison.OrdinalIgnoreCase)) srcInfo.Id = setupInfo.Split(':')[1];
                if (item.Equals("srcpin", StringComparison.OrdinalIgnoreCase)) srcInfo.Pin = setupInfo.Split(':')[1];
                if (item.Equals("srcfs", StringComparison.OrdinalIgnoreCase))
                    srcInfo.Fs = DataConvertor.ConvertUnits(setupInfo.Split(':')[1]);
                if (item.Equals("srcfi", StringComparison.OrdinalIgnoreCase))
                    srcInfo.Fi = DataConvertor.ConvertUnits(setupInfo.Split(':')[1]);
                if (item.Equals("srcn", StringComparison.OrdinalIgnoreCase)) srcInfo.N = setupInfo.Split(':')[1];
                if (item.Equals("srcm", StringComparison.OrdinalIgnoreCase)) srcInfo.M = setupInfo.Split(':')[1];
                if (item.Equals("srcamp", StringComparison.OrdinalIgnoreCase)) srcInfo.Amp = setupInfo.Split(':')[1];
                if (item.Equals("srcvcm", StringComparison.OrdinalIgnoreCase)) srcInfo.Vcm = setupInfo.Split(':')[1];
                if (item.Equals("srcvcmenable", StringComparison.OrdinalIgnoreCase))
                    srcInfo.VcmEnable = setupInfo.Split(':')[1];
                if (item.Equals("srcbw", StringComparison.OrdinalIgnoreCase))
                    srcInfo.Bw = DataConvertor.ConvertUnits(setupInfo.Split(':')[1]);
                if (item.Equals("srcdiff", StringComparison.OrdinalIgnoreCase))
                    srcInfo.DiffMode = setupInfo.Split(':')[1];
                if (item.Equals("srcwave", StringComparison.OrdinalIgnoreCase))
                    srcInfo.UseWave = setupInfo.Split(':')[1];
            }

            return srcInfo;
        }

        private AnalogDigCap _AnalyzeAnalog_CapInfo(string miscInfo)
        {
            if (string.IsNullOrEmpty(miscInfo)) return null;
            if (!Regex.IsMatch(miscInfo, "Digcap_Fs")) return null;
            var capInfo = new AnalogDigCap();

            foreach (var setupInfo in miscInfo.Split(';').ToList())
            {
                var item = setupInfo.Split(':')[0];
                //for src
                if (item.Equals("Digcap_Fs", StringComparison.OrdinalIgnoreCase)) capInfo.Fs = setupInfo.Split(':')[1];
                if (item.Equals("Digcap_Coherent_N", StringComparison.OrdinalIgnoreCase))
                    capInfo.CoherentN = setupInfo.Split(':')[1];
                if (item.Equals("Digcap_M", StringComparison.OrdinalIgnoreCase)) capInfo.M = setupInfo.Split(':')[1];
                if (item.Equals("Digcap_Fr", StringComparison.OrdinalIgnoreCase)) capInfo.Fr = setupInfo.Split(':')[1];
                if (item.Equals("Digcap_Discard_Samples_Per_BLK", StringComparison.OrdinalIgnoreCase))
                    capInfo.DigcapDiscardSamplesPerBlk = setupInfo.Split(':')[1];
                if (item.Equals("Digcap_Discard_Samples_Per_ADC", StringComparison.OrdinalIgnoreCase))
                    capInfo.DigcapDiscardSamplesPerAdc = setupInfo.Split(':')[1];
                if (item.Equals("Digcap_PreProcess_Type", StringComparison.OrdinalIgnoreCase))
                    capInfo.PreProcessType = setupInfo.Split(':')[1];
                if (item.Equals("NumADC", StringComparison.OrdinalIgnoreCase)) capInfo.NumAdc = setupInfo.Split(':')[1];
                if (item.Equals("NumBlk", StringComparison.OrdinalIgnoreCase)) capInfo.NumBlk = setupInfo.Split(':')[1];
                if (item.Equals("ADC_FullScale", StringComparison.OrdinalIgnoreCase))
                    capInfo.AdcFullScale = setupInfo.Split(':')[1];
                if (item.Equals("Current_Blk", StringComparison.OrdinalIgnoreCase))
                    capInfo.CurrentBlk = setupInfo.Split(':')[1];
            }

            return capInfo;
        }

        private string GetMeasWaitTime(TestPlanRow tpRow)
        {
            foreach (var misc in tpRow.MiscInfo.Split(';'))
            {
                var miscItems = misc.Split(':');
                if (miscItems[0].Equals("MeasWait", StringComparison.OrdinalIgnoreCase) && miscItems.Length > 1)
                    return miscItems[1];
            }

            if (Regex.IsMatch(tpRow.Meas, MeasType.MeasWait, RegexOptions.IgnoreCase))
            {
                var regWait = @"MeasWait\s*(?<time>.*)";
                var time = Regex.Match(tpRow.Meas, regWait, RegexOptions.IgnoreCase).Groups["time"].ToString();
                return time;
            }

            return "";
        }


        private string GetMeasIRange(TestPlanRow tpRow)
        {
            foreach (var misc in tpRow.MiscInfo.Split(';'))
            {
                var miscItems = misc.Split(':');
                if (miscItems[0].Equals("MeasRange", StringComparison.OrdinalIgnoreCase) && miscItems.Length > 1)
                    return miscItems[1];
            }

            return "";
        }
    }
}