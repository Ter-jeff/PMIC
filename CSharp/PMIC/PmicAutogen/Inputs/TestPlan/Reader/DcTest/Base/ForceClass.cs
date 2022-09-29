using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    [Serializable]
    public class ForceClass
    {
        //Dc setting eg: Level:XXX
        private const string RegLevelSetting = @"(Level:|Levels:)(?<level>[\w]+)";

        //Ac setting eg: AC:XXX:XXX
        private const string RegAcSetting = @"AC:[\w|&]+:[\w]+";

        //Ac Category eg: AC:XXX
        private const string RegAcCategory = @"AC:[\w]+";

        //Ac selector eg: ACSelector:NV:XXX
        private const string RegAcSelector = @"ACSelector:[\w|&]+:[\w]+";

        //Ac Category eg: DC:XXX
        private const string RegDcCategory = @"DC:[\w]+";

        //Dc selector eg: DCSelector:NV:XXX
        private const string RegDcSelector = @"DCSelector:[\w|&]+:[\w]+";

        public ForceClass()
        {
            IsShmooInForce = false;
            IsShmooInProdInst = true;
            IsShmooInProdFlow = true;
            IsShmooInCharInst = false;
            IsShmooInCharFlow = false;
            IsCz2InstName = false;
            ForceCondition = "";
        }

        public bool IsShmooInForce { get; set; }
        public bool IsShmooInProdInst { get; set; }
        public bool IsShmooInProdFlow { get; set; }
        public bool IsShmooInCharInst { get; set; }
        public bool IsShmooInCharFlow { get; set; }
        public bool IsCz2InstName { get; set; }
        public string ForceCondition { get; set; }

        //public HardipCharSetup GetShmoo(HardIpPattern pattern, string pattForceCondition, string subBlockName)
        //{
        //    try
        //    {
        //        //subBlockName = subBlockName.Replace("-", "");
        //        var isSweep = false;
        //        var result = new HardipCharSetup();
        //        var originalParameterName = "";
        //        foreach (var forceCondition in pattForceCondition.Split(';'))
        //            //PowerTableRow fromByVoltage = null;
        //            //PowerTableRow toByVoltage = null;
        //            if (Regex.IsMatch(forceCondition, HardIpConstData.RegShmoo, RegexOptions.IgnoreCase))
        //            {
        //                var str = Regex.Match(forceCondition, HardIpConstData.RegShmoo, RegexOptions.IgnoreCase)
        //                    .Groups["ShmooStr"].ToString();
        //                //Eg. Sweep(VDD:vih:0.6,1.2,0.005:retest:jump,6)
        //                var strArr = str.Split(':');

        //                var oneShmooStep = new CharStep(pattern.Pattern.GetLastPayload(), pattern.PatternType);
        //                oneShmooStep.Mode = Regex.IsMatch(forceCondition, "yshmoo|sweepy", RegexOptions.IgnoreCase)
        //                    ? CharStepConst.ModeYShmoo
        //                    : CharStepConst.ModeXShmoo;

        //                if (strArr.Length < 3) return new HardipCharSetup();
        //                var pinName = strArr[0].Trim();

        //                string firstWord;
        //                if (strArr[1].Contains(","))
        //                {
        //                    var arr = strArr[1].Split(',');
        //                    firstWord = arr[0];
        //                }
        //                else
        //                {
        //                    firstWord = strArr[1];
        //                }

        //                originalParameterName = strArr[1];

        //                oneShmooStep.ApplyToPinExecMode = "Simultaneous";
        //                oneShmooStep.StepName = pinName.Replace(",", "_");
        //                oneShmooStep.ParameterName = HardipCharSetup.GetShmooParameterName(firstWord);
        //                oneShmooStep.ParameterType = CharSetupSingleton.Instance().GetShmooParameterType(firstWord);
        //                oneShmooStep.ApplyToTimeSets = HardipCharSetup.GetShmooTimeSets(strArr[1]);

        //                if (oneShmooStep.ParameterType == CharStepConst.ParameterTypeAcSpec ||
        //                    oneShmooStep.ParameterType == CharStepConst.ParameterTypeDcSpec ||
        //                    oneShmooStep.ParameterType == CharStepConst.ParameterTypeGlobalSpec)
        //                    oneShmooStep.ApplyToPins = "";
        //                else
        //                    oneShmooStep.ApplyToPins = strArr[0];

        //                oneShmooStep.StepName = strArr[0].Replace(",", "_") + "_" + oneShmooStep.ParameterName;

        //                var fromArr = strArr[2].Split(',');
        //                if (fromArr.Length == 3)
        //                {
        //                    oneShmooStep.RangeCalcField = CharStepConst.RangeCalcFieldSteps;
        //                    //if (HardIpDataMain.HardipDcDefault.PowRows.Exists(x => x.PinName.Equals(fromArr[0], StringComparison.CurrentCulture)))
        //                    //{
        //                    //    fromByVoltage = HardIpDataMain.HardipDcDefault.PowRows.Find(x => x.PinName.Equals(fromArr[0], StringComparison.CurrentCulture));
        //                    //    oneShmooStep.RangeFrom = fromByVoltage.Vmain;
        //                    //    oneShmooStep.VoltageType = "NV";
        //                    //    result.IsSplitByVoltage = true;
        //                    //}
        //                    //else
        //                    oneShmooStep.RangeFrom = DataConvertor.ConvertUnits(fromArr[0].Trim());

        //                    //if (HardIpDataMain.HardipDcDefault.PowRows.Exists(x => x.PinName.Equals(fromArr[1], StringComparison.CurrentCulture)))
        //                    //{
        //                    //    toByVoltage = HardIpDataMain.HardipDcDefault.PowRows.Find(x => x.PinName.Equals(fromArr[1], StringComparison.CurrentCulture));
        //                    //    oneShmooStep.RangeTo = toByVoltage.Vmain;
        //                    //    oneShmooStep.VoltageType = "NV";
        //                    //    result.IsSplitByVoltage = true;
        //                    //}
        //                    //else
        //                    oneShmooStep.RangeTo = DataConvertor.ConvertUnits(fromArr[1].Trim());
        //                    oneShmooStep.RangeStepSize = DataConvertor.ConvertUnits(fromArr[2].Trim());
        //                }

        //                if (strArr.Length == 3)
        //                {
        //                    if (oneShmooStep.ParameterType == CharStepConst.ParameterTypeAcSpec)
        //                    {
        //                        //Timing Sweep
        //                        result.TestMethod = CharSetupConst.TestMethodRetest;
        //                        oneShmooStep.AlgorithmName = CharStepConst.AlgorithmNameLinear;
        //                    }
        //                    else
        //                    {
        //                        if (Regex.IsMatch(forceCondition, "sweepy|yshmoo", RegexOptions.IgnoreCase))
        //                        {
        //                            result.TestMethod = CharSetupConst.TestMethodRetest;
        //                            oneShmooStep.AlgorithmName = CharStepConst.AlgorithmNameLinear;
        //                        }
        //                        else
        //                        {
        //                            //Default Sweep
        //                            result.TestMethod = CharSetupConst.TestMethodRetest;
        //                            oneShmooStep.AlgorithmName = CharStepConst.AlgorithmNameJump;
        //                            oneShmooStep.AlgorithmArgs = "6";
        //                        }
        //                    }
        //                }

        //                if (strArr.Length == 4)
        //                {
        //                    result.TestMethod = CharSetupConst.TestMethod.ContainsKey(strArr[3])
        //                        ? CharSetupConst.TestMethod[strArr[3]]
        //                        : "";
        //                }
        //                else if (strArr.Length == 5)
        //                {
        //                    result.TestMethod = CharSetupConst.TestMethod.ContainsKey(strArr[3])
        //                        ? CharSetupConst.TestMethod[strArr[3]]
        //                        : "";
        //                    var algorithmArr = strArr[4].Split(',');
        //                    if (algorithmArr.Length == 1)
        //                    {
        //                        oneShmooStep.AlgorithmName = algorithmArr[0];
        //                    }
        //                    else if (algorithmArr.Length == 2)
        //                    {
        //                        oneShmooStep.AlgorithmName = algorithmArr[0];
        //                        oneShmooStep.AlgorithmArgs = algorithmArr[1];
        //                    }
        //                }

        //                //if (HardIpDataMain.NwirePinsList.Any(x => x.OutClk.ToUpper().Contains(pinName.ToUpper()) || x.OutClkDiff.ToUpper().Contains(pinName.ToUpper())))
        //                //{
        //                //    var pin = HardIpDataMain.NwirePinsList.Where(x => x.OutClk.ToUpper().Contains(pinName.ToUpper()) || x.OutClkDiff.ToUpper().Contains(pinName.ToUpper())).ToList();
        //                //    var portName = pin.FirstOrDefault().CreatePortName();
        //                //    var freqVarName = pin.FirstOrDefault().CreateFreqVarName();
        //                //    oneShmooStep.PrePointFunction = "freerunclk_set_XY";
        //                //    oneShmooStep.PrePointArgs = oneShmooStep.Mode == CharStepConst.ModeYShmoo ? string.Format("Y," + portName + "," + freqVarName) : string.Format("X," + portName + "," + freqVarName);
        //                //    oneShmooStep.PostPointFunction = "freerunclk_stop";
        //                //    oneShmooStep.PostPointArgs = string.Format(portName);
        //                //    oneShmooStep.StepName = freqVarName;
        //                //    oneShmooStep.ParameterType = CharStepConst.ParameterTypeAcSpec;
        //                //    oneShmooStep.ParameterName = freqVarName;
        //                //    oneShmooStep.ApplyToPins = "";
        //                //}

        //                oneShmooStep.PostStepFunction = CharStepConst.PostStepFunctionPrintShmooInfo;
        //                oneShmooStep.PostStepArgs = "CorePower";

        //                if (Regex.IsMatch(forceCondition, "sweep|sweepy", RegexOptions.IgnoreCase))
        //                {
        //                    oneShmooStep.SuspendDataLog = "FALSE";
        //                    isSweep = true;
        //                }
        //                else if (Regex.IsMatch(forceCondition, "xshmoo|yshmoo", RegexOptions.IgnoreCase))
        //                {
        //                    oneShmooStep.SuspendDataLog = "TRUE";
        //                    isSweep = false;
        //                }

        //                oneShmooStep.OutputFormat = "Enhanced";
        //                oneShmooStep.OutputToTextFile = " Disable";
        //                oneShmooStep.OutputToSheet = "Disable";
        //                oneShmooStep.OutputToDataLog = "Enable";
        //                oneShmooStep.OutputToImmediateWin = "Disable";
        //                oneShmooStep.OutputToOutputWin = "Disable";

        //                if (result.IsSplitByVoltage)
        //                {
        //                    if (result.CharSteps.Exists(x => string.IsNullOrEmpty(x.VoltageType)))
        //                    {
        //                        var rows = result.CharSteps.Where(x => string.IsNullOrEmpty(x.VoltageType)).ToList();
        //                        for (var i = 0; i < rows.Count; i++)
        //                        {
        //                            var lvRow = rows[i].DeepClone();
        //                            var hvRow = rows[i].DeepClone();
        //                            lvRow.VoltageType = "LV";
        //                            hvRow.VoltageType = "HV";
        //                            rows[i].VoltageType = "NV";
        //                            result.AddStep(lvRow);
        //                            result.AddStep(hvRow);
        //                        }
        //                    }

        //                    var lvOneShmooStep = oneShmooStep.DeepClone();
        //                    var hvOneShmooStep = oneShmooStep.DeepClone();
        //                    //if (fromByVoltage != null)
        //                    //{
        //                    //    lvOneShmooStep.RangeFrom = fromByVoltage.OrgPowerTableStageItemsStageMainList[0].LowValue;
        //                    //    hvOneShmooStep.RangeFrom = fromByVoltage.OrgPowerTableStageItemsStageMainList[0].HighValue;
        //                    //}
        //                    //if (toByVoltage != null)
        //                    //{
        //                    //    lvOneShmooStep.RangeTo = toByVoltage.OrgPowerTableStageItemsStageMainList[0].LowValue;
        //                    //    hvOneShmooStep.RangeTo = toByVoltage.OrgPowerTableStageItemsStageMainList[0].HighValue;
        //                    //}
        //                    lvOneShmooStep.VoltageType = "LV";
        //                    hvOneShmooStep.VoltageType = "HV";
        //                    oneShmooStep.VoltageType = "NV";
        //                    result.AddStep(lvOneShmooStep);
        //                    result.AddStep(hvOneShmooStep);
        //                }

        //                result.AddStep(oneShmooStep);
        //            }

        //        if (result.CharSteps.Any())
        //        {
        //            var blockName = CommonGenerator.GetBlockNameFromSheetName(pattern.SheetName);
        //            if (string.IsNullOrEmpty(subBlockName)) subBlockName = "SubBlock";
        //            result.SetupName = HardIpConstData.PrefixShmooSetupName + blockName + "_" +
        //                               CommonGenerator.GetSubBlockNameWithoutMinus(subBlockName);
        //            result.TestNameInFlow =
        //                HardipCharSetup.GetShmooName(pattern, result, subBlockName, isSweep, originalParameterName);
        //        }

        //        return result;
        //    }
        //    catch (Exception e)
        //    {
        //        Response.Report(e.ToString(), MessageLevel.Error, 0);
        //        return null;
        //    }
        //}

        public string GetLevelSetting()
        {
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(force, RegLevelSetting, RegexOptions.IgnoreCase))
                    return force.Split(':')[1];
            return "";
        }

        public string GetAcSetting()
        {
            var acSettings = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase))
                    if (IsAcSpecPin(Regex.Replace(force, "::", "&").Split(':')[1]))
                        acSettings += force + ";";
            return acSettings.Trim(';');
        }

        private bool IsAcSpecPin(string pinName)
        {
            //return Regex.IsMatch(pinName, @"^(TCK|ShiftIn)$") ||
            //       NwireSingleton.Instance().SettingInfo.NwirePins.Find(s => s.OutClk.Equals(pinName, StringComparison.OrdinalIgnoreCase))
            //        != null;

            return Regex.IsMatch(pinName, @"^(TCK|ShiftIn)$");
        }

        public string GetAcSelector()
        {
            var acAcSelector = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSelector, RegexOptions.IgnoreCase))
                    acAcSelector += force + ";";
            return acAcSelector.Trim(';');
        }

        public string GetDcCategory()
        {
            var dcSettings = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcCategory, RegexOptions.IgnoreCase))
                    dcSettings += force + ";";
            return dcSettings.Trim(';');
        }

        public string GetAcCategory()
        {
            var acSettings = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcCategory, RegexOptions.IgnoreCase))
                    acSettings += force + ";";
            return acSettings.Trim(';');
        }

        public string GetDcSelector()
        {
            var dcAcSelector = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcSelector, RegexOptions.IgnoreCase))
                    dcAcSelector += force + ";";
            return dcAcSelector.Trim(';');
        }

        public string GetMcgSetting()
        {
            var mcgSettings = "";
            foreach (var force in ForceCondition.Split(';'))
                if (Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase))
                    if (!IsAcSpecPin(Regex.Replace(force, "::", "&").Split(':')[1]))
                        mcgSettings += force + ";";
            return mcgSettings.Trim(';');
        }

        public string GetPrePatForceCondition()
        {
            // Remove Dc setting, Ac setting, and Mcg setting
            var forceList = ForceCondition.Split(';').ToList();
            forceList.RemoveAll(string.IsNullOrEmpty);
            foreach (var force in forceList.ToArray())
                if (Regex.IsMatch(force, RegLevelSetting, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSetting, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcCategory, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcCategory, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegAcSelector, RegexOptions.IgnoreCase) ||
                    Regex.IsMatch(Regex.Replace(force, "::", "&"), RegDcSelector, RegexOptions.IgnoreCase))
                    forceList.Remove(force);
            return string.Join(";", forceList);
        }
    }
}