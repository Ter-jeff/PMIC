using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    [Serializable]
    public class HardipCharSetup : CharSetup
    {
        //public static string GetShmooName(HardIpPattern pattern, HardipCharSetup shmoo, string subBlockName,
        //    bool isSweep, string originalParameterName)
        //{
        //    string[] shmooNameArr = {"X", "X", "X", "X", "X", "X", "X", "X", "X", "X_"};
        //    var pinName1D = "";
        //    var pinName2D = "";
        //    var shmooStep = shmoo.CharSteps[0];
        //    var xShmoo = shmoo.CharSteps.Where(x => x.Mode == CharStepConst.ModeXShmoo)
        //        .FirstOrDefault(y => y.ApplyToPins != "");
        //    if (xShmoo != null) pinName1D = xShmoo.ApplyToPins;
        //    var yShmoo = shmoo.CharSteps.Where(x => x.Mode == CharStepConst.ModeYShmoo)
        //        .FirstOrDefault(y => y.ApplyToPins != "");
        //    if (yShmoo != null) pinName2D = yShmoo.ApplyToPins;

        //    if (isSweep) //HAC
        //    {
        //        shmooNameArr[0] = "HAC";
        //    }
        //    else //HFH,HFL,HIO
        //    {
        //        #region UD1

        //        var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
        //        if (pinMap != null)
        //        {
        //            var pinType = ""; //if any pin is power pin
        //            var allIoFlag = true;
        //            foreach (var charStep in shmoo.CharSteps)
        //            {
        //                var pinName = charStep.ApplyToPins;
        //                if (pinMap.IsPinExist(pinName))
        //                {
        //                    var type = pinMap.GetPin(pinName1D).PinType.ToUpper();
        //                    if (type == PinMapConst.TypePower || charStep.StepName.StartsWith("VDD"))
        //                        pinType = PinMapConst.TypePower;
        //                    if (type != PinMapConst.TypeIo) allIoFlag = false;
        //                }
        //            }

        //            if (allIoFlag)
        //            {
        //                shmooNameArr[0] = "HIO";
        //            }
        //            else if (pinType == PinMapConst.TypePower)
        //            {
        //                var from = DataConvertor.ConvertUnits(shmooStep.RangeFrom);
        //                var to = DataConvertor.ConvertUnits(shmooStep.RangeTo);
        //                double fromValue;
        //                var fromMatch = double.TryParse(from, out fromValue);
        //                double toValue;
        //                var toMatch = double.TryParse(to, out toValue);
        //                if (fromMatch && toMatch)
        //                {
        //                    shmooNameArr[0] = fromValue > toValue ? "HFL" : "HFH";
        //                }
        //                else
        //                {
        //                    var forceIndex = HardIpDataMain.TestPlanData.PlanHeaderIdx[pattern.SheetName]["forceIndex"];
        //                    var errorMessage = "Wrong ForceCondition for " + pinName1D;
        //                    EpplusErrorManager.AddError(HardIpErrorType.WrongForceCondition, ErrorLevel.Error,
        //                        pattern.SheetName, pattern.RowNum, forceIndex, errorMessage);
        //                }
        //            }
        //            else
        //            {
        //                shmooNameArr[0] = "UNKNOWN";
        //                var forceIndex = HardIpDataMain.TestPlanData.PlanHeaderIdx[pattern.SheetName]["forceIndex"];
        //                var errorMessage = "Wrong ForceCondition for " + pinType;
        //                EpplusErrorManager.AddError(HardIpErrorType.WrongForceCondition, ErrorLevel.Error,
        //                    pattern.SheetName, pattern.RowNum, forceIndex, errorMessage);
        //            }
        //        }

        //        #endregion

        //        #region UD2

        //        var measPin = pattern.MeasPins.Select(x => x.MeasType).FirstOrDefault();
        //        if (measPin != null)
        //            shmooNameArr[1] = measPin;

        //        #endregion

        //        #region UD4

        //        shmooNameArr[3] = subBlockName;

        //        #endregion

        //        #region UD5

        //        shmooNameArr[4] = CommonGenerator.GetBlockName(pattern.MiscInfo, pattern.SheetName);

        //        #endregion

        //        #region UD6

        //        if (shmoo.CharSteps.Count(y => y.Mode == CharStepConst.ModeXShmoo) > 1)
        //            shmooNameArr[5] = "MULTI";
        //        else
        //            shmooNameArr[5] = pinName1D.Replace(",", "");

        //        #endregion

        //        #region UD7

        //        #endregion

        //        #region UD8

        //        shmooNameArr[7] = pinName2D != ""
        //            ? pinName2D.Replace("_", "")
        //            : originalParameterName.Replace("_", "").Replace(",", "");

        //        #endregion
        //    }

        //    #region UD9

        //    //var measCPin = pattern.MeasPins.FirstOrDefault(x => x.MeasType == MeasType.MeasC);
        //    //if (measCPin != null) shmooNameArr[8] = measCPin.CusStr;
        //    var block2Name = CommonGenerator.GetSubBlock2Name(pattern.MiscInfo);
        //    if (block2Name != "") shmooNameArr[8] = block2Name;

        //    #endregion

        //    #region UD10

        //    //shmooNameArr[9]=Sequence counter , used to count number of similar measurement performed on same pin, useful to build sweep results for repeated results

        //    #endregion

        //    for (var i = 0; i < shmooNameArr.Length; i++)
        //        shmooNameArr[i] = shmooNameArr[i].Replace("_", "");

        //    return string.Join("_", shmooNameArr);
        //}

        //public static bool IsSameForceShmoo(string force, string forceChar)
        //{
        //    var flag = true;
        //    var forceList = force.Replace(" ", "").Split(';').ToList();
        //    var forceCharList = forceChar.Replace(" ", "").Split(';').ToList();
        //    foreach (var item in forceCharList)
        //        if ((item != "") & !Regex.IsMatch(item, HardIpConstData.RegShmoo, RegexOptions.IgnoreCase))
        //            if (forceList.Count(x => x == item) == 0)
        //                flag = false;
        //    foreach (var item in forceList)
        //        if ((item != "") & !Regex.IsMatch(item, HardIpConstData.RegShmoo, RegexOptions.IgnoreCase))
        //            if (forceCharList.Count(x => x == item) == 0)
        //                flag = false;

        //    return flag;
        //}

        //public static List<HardipCharSetup> GetShmoo(HardIpPattern pattern)
        //{
        //    var result = new List<HardipCharSetup>();
        //    if (pattern.Shmoo.IsSplitByVoltage)
        //    {
        //        //NV
        //        var newSetupNv = new HardipCharSetup();
        //        newSetupNv.SetupName =
        //            CommonGenerator.GetSubBlockNameWithoutMinus(pattern.Shmoo.SetupName) + "_" + "NV";
        //        newSetupNv.TestMethod = pattern.Shmoo.TestMethod;
        //        newSetupNv.CharSteps = pattern.Shmoo.CharSteps.Where(x => x.VoltageType == "NV").ToList();
        //        result.Add(newSetupNv);
        //        //LV
        //        var newSetupLv = new HardipCharSetup();
        //        newSetupLv.SetupName =
        //            CommonGenerator.GetSubBlockNameWithoutMinus(pattern.Shmoo.SetupName) + "_" + "LV";
        //        newSetupLv.TestMethod = pattern.Shmoo.TestMethod;
        //        newSetupLv.CharSteps = pattern.Shmoo.CharSteps.Where(x => x.VoltageType == "LV").ToList();
        //        result.Add(newSetupLv);
        //        //HV
        //        var newSetupHv = new HardipCharSetup();
        //        newSetupHv.SetupName =
        //            CommonGenerator.GetSubBlockNameWithoutMinus(pattern.Shmoo.SetupName) + "_" + "HV";
        //        newSetupHv.TestMethod = pattern.Shmoo.TestMethod;
        //        newSetupHv.CharSteps = pattern.Shmoo.CharSteps.Where(x => x.VoltageType == "HV").ToList();
        //        result.Add(newSetupHv);
        //    }
        //    else
        //    {
        //        var newSetup = new HardipCharSetup();
        //        var charName = CommonGenerator.GetSubBlockNameWithoutMinus(pattern.Shmoo.SetupName);
        //        newSetup.SetupName = charName;
        //        newSetup.TestMethod = pattern.Shmoo.TestMethod;
        //        newSetup.CharSteps = pattern.Shmoo.DeepClone();
        //        result.Add(newSetup);
        //    }

        //    return result;
        //}


        public static string GetShmooParameterName(string name)
        {
            var hardCodeDic = new Dictionary<string, string>
                {{"d0", "On"}, {"d1", "Data"}, {"d2", "Return"}, {"d3", "Off"}};
            if (CharStepConst.ParameterName.ContainsKey(name)) name = CharStepConst.ParameterName[name];

            if (hardCodeDic.ContainsKey(name.ToLower()))
                return hardCodeDic[name.ToLower()];

            return name;
        }

        public static string GetShmooTimeSets(string name)
        {
            if (name.Contains(","))
            {
                var arr = name.Split(',').ToList();
                arr.RemoveAt(0);
                return string.Join(",", arr);
            }

            return "";
        }

        #region Property

        public string TestNameInFlow { set; get; }
        public bool IsSplitByVoltage { set; get; }

        #endregion

        #region Constructor

        #endregion
    }
}