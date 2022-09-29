using System;
using System.Collections.Generic;
using System.Linq;
using IgxlData.VBT;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetVifValue : SetValueBase
    {
        public override void SetArgsListValue(HardIpPattern pattern, VbtFunctionBase function, string voltage)
        {
            if (pattern.RowNum == 833)
            {
            }

            var cPin = SearchInfo.GetMeasCPins(pattern);
            //patSet
            function.Args[0] = pattern.Pattern.GetInstancePatternName();
            var info = SearchInfo.GetHardIpInfo(pattern);
            //Cpu_flag_A
            function.SetParamValue("CPUA_Flag_In_Pat", SearchInfo.GetCpuFlag(info, pattern));

            #region Defalut value

            //TestLimitPerPin
            function.SetParamValue("TestLimitPerPin_VFI",
                string.Join("", SearchInfo.GetTestLimitPerMeasType(pattern).Values.ToList()));
            //DisableComparePins
            //function.SetParamValue("DisableComparePins", "-1");
            //MeasF_Interval
            if (pattern.IsNonHardIpBlock)
                function.SetParamValue("MeasF_Interval", "0.01");
            //MeasF_EventSourceWithTerminationMode
            function.SetParamValue("MeasF_EventSourceWithTerminationMode", pattern.IsNonHardIpBlock ? "2" : "0");
            //MeasF_ThresholdPercentage
            function.SetParamValue("MeasF_ThresholdPercentage", "0.5");
            //MeasF_WaitTime
            //function.SetParamValue("MeasF_WaitTime", "0.01");

            #endregion

            #region RegisterAssignment

            //DigSrc_Assignment: Use "Register Assignment" value in test plan
            //function.Args[25] = pattern.RegisterAssignment;

            function.SetParamValue("DigSrc_Assignment", pattern.DigSrcAssignment.Replace("[", ":").Replace("]", ""));
            //DigSrc_Equation: From patInfo file "Send Bit Name"
            function.SetParamValue("DigSrc_Equation", pattern.DigSrcEquation);
            //DigSrc_Sample_Size: Get from "Send Bit" in patInfo file, Like Send Bit: 160  ===> 160
            //function.Args[23] = info.SendBit.ToString();
            function.SetParamValue("DigSrc_Sample_Size", info.SendBit.ToString("G"));
            //DigSrc_DataWidth: Get from "Send Bit Str" in patInfo file. Like wdr0_16+wdr1_16+wdr2_16 ===> 16
            //function.Args[22] =
            function.SetParamValue("DigSrc_DataWidth", SearchInfo.GetDigDataWidth(info.SendBitStr, "0"));
            //DigSrc_Pin
            function.SetParamValue("DigSrc_Pin", SearchInfo.GetSrcPin(info));

            if (pattern.SweepCodes.SelectMany(p => p.Value).ToList().Find(p => p.IsGrayCode) != null)
                function.SetParamValue("CUS_Str_DigSrcData", "BinToGray");
            //}

            #endregion

            #region MeasC

            //if (cPin != "")//&& _capture)
            //{
            //DigCap_Pin: MeasC pin in Test Plan, Like "MeasC Pin = Pout" ===> Pout
            //function.Args[17] = cPin;
            function.SetParamValue("DigCap_Pin", cPin);
            //DigCap_DataWidth:  Get from "Cap Bit Str" in patInfo file. Like "wdr14_10+wdr23_10" ===> 10
            //function.Args[18] = Regex.Match(info.CapBitStr, @"^wdr\d+_(?<num>(\d+)).*").Groups["num"].ToString();
            function.SetParamValue("DigCap_DataWidth", SearchInfo.GetDigDataWidth(info.CapBitStr));
            //DigCap_Sample_Size: Get from "Cap Bit" in patInfo file
            //function.Args[19] = info.CapBit.ToString();
            function.SetParamValue("DigCap_Sample_Size", info.CapBit.ToString("G"));

            //CUS_Str_DigCapData
            //function.Args[31] = SearchInfo.GetCusStrDigCapData(pattern);
            function.SetParamValue("CUS_Str_DigCapData", SearchInfo.GetCusStrDigCapData(pattern));
            //}
            var storeName = SearchInfo.GetStoreName(pattern);
            function.SetParamValue("Meas_StoreName", storeName);

            #endregion

            #region Meas Pins

            var seqCount = info.NewInfo != null ? info.NewInfo.SeqInfo.Count :
                info.SeqInfo.Count == 0 ? pattern.TestPlanSequences.Count : info.SeqInfo.Count;
            if (seqCount > 0)
            {
                #region Gets measPins from TestSequence

                var measVPins = new List<string>();
                var measIPins = new List<string>();
                var measFPins = new List<string>();
                var measFdiffPins = new List<string>();

                for (var sequenceIndex = 1; sequenceIndex <= seqCount; sequenceIndex++)
                {
                    var seqPins = pattern.MeasPins.Where(p => p.SequenceIndex == sequenceIndex).ToList();
                    var selVType = seqPins.Exists(p => p.MeasType == MeasType.MeasVdiff)
                        ? MeasType.MeasVdiff
                        :
                        seqPins.Exists(p => p.MeasType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase))
                            ?
                            MeasType.MeasVdm
                            : MeasType.MeasV;
                    var measIInter = seqPins.Where(p =>
                        p.MeasType.Equals(MeasType.MeasI) || p.MeasType.Equals(MeasType.MeasR1) ||
                        p.MeasType.Equals(MeasType.MeasR2)).ToList();

                    var measVInter = seqPins.Where(p => p.MeasType.Equals(selVType, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    var measFInter = seqPins.Where(p => p.MeasType.Equals(MeasType.MeasF)).ToList();
                    var measFdiffInter = seqPins.Where(p => p.MeasType.Equals(MeasType.MeasFdiff)).ToList();
                    measVPins.Add(selVType.Equals(MeasType.MeasVdm, StringComparison.OrdinalIgnoreCase)
                        ? string.Join(",", measVInter.Select(p => p.PinName))
                        : string.Join(",", measVInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, false))));
                    measIPins.Add(string.Join(",",
                        measIInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, false))));
                    measFPins.Add(
                        string.Join(",", measFInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, true))));
                    measFdiffPins.Add(string.Join(",",
                        measFdiffInter.Select(p => SearchInfo.GenDiffGroupName(p.PinName, true))));
                }

                #endregion

                //MeasureV_PinS
                if (measVPins.Any(p => !string.IsNullOrEmpty(p)))
                    function.SetParamValue("MeasV_PinS",
                        DataConvertor.SortCpFtPin(SearchInfo.CheckInfoByStoreName(
                            DataConvertor.RemoveDummyPlusSign(string.Join("+", measVPins)), storeName, '+')));
                //MeasureF_PinS_SingleEnd
                if (measFPins.Any(p => !string.IsNullOrEmpty(p)))
                    function.SetParamValue("MeasF_PinS_SingleEnd",
                        DataConvertor.SortCpFtPin(SearchInfo.CheckInfoByStoreName(
                            DataConvertor.RemoveDummyPlusSign(string.Join("+", measFPins)), storeName, '+')));
                //MeasureF_PinS_Differential
                if (measFdiffPins.Any(p => !string.IsNullOrEmpty(p)))
                    function.SetParamValue("MeasF_PinS_Differential",
                        DataConvertor.SortCpFtPin(SearchInfo.CheckInfoByStoreName(
                            DataConvertor.RemoveDummyPlusSign(string.Join("+", measFdiffPins)), storeName, '+')));

                //MeasureI_pinS
                if (measIPins.Any(p => !string.IsNullOrEmpty(p)))
                {
                    function.SetParamValue("MeasI_pinS",
                        DataConvertor.SortCpFtPin(SearchInfo.CheckInfoByStoreName(
                            DataConvertor.RemoveDummyPlusSign(string.Join("+", measIPins)), storeName, '+')));
                    //MeasI_Range
                    function.SetParamValue("MeasI_Range",
                        DataConvertor.RedefineRange(pattern, info,
                            SearchInfo.CheckInfoByStoreName(SearchInfo.GetIRange(pattern, info, voltage), storeName,
                                '|')));
                }
            }

            #endregion

            var forceV = SearchInfo.GetForceV(pattern, info);
            function.SetParamValue("ForceV_Val", SearchInfo.CheckInfoByStoreName(forceV, storeName, '|'));
            function.SetParamValue("ForceI_Val",
                SearchInfo.CheckInfoByStoreName(SearchInfo.GetForceI(pattern, info), storeName, '|'));
            //Measure Sequence 
            var measSeq = SearchInfo.GetMeasSequence(pattern).ToUpper().Replace("FDIFF", "F").Replace("R1", "R")
                .Replace("VDIFF", "V").Replace("IDIFF", "I")
                .Replace("R2", "Z"); //FDIFF need to be changed to F in meas sequence //20170425 Roger add .ToUpper()
            function.SetParamValue("TestSequence", SearchInfo.CheckInfoByStoreName(measSeq, storeName, ',', true));

            //if MeasI2,set SpecialCalcValSetting=3
            if (pattern.SpecialMeasType.Equals(MeasType.MeasI2))
                function.SetParamValue("SpecialCalcValSetting", "3");
            else if (pattern.SpecialMeasType.Equals(MeasType.MeasVdiff))
                function.SetParamValue("SpecialCalcValSetting", "4");
            else if (pattern.SpecialMeasType.Equals(MeasType.MeasVocm))
                function.SetParamValue("SpecialCalcValSetting", "9");
            function.SetParamValue("Interpose_PrePat", SearchInfo.GetPrePat(pattern));
            function.SetParamValue("Interpose_PreMeas", SearchInfo.GetPreMeas(pattern));
            function.SetParamValue("Interpose_PostTest", pattern.InterposePostTest);

            //CalcEqn and storeName
            function.SetParamValue("Calc_Eqn", DataConvertor.ConvertValueSpec(pattern.CalcEqn));

            if (SearchInfo.GetFlagSingleLimit(pattern, voltage))
                function.SetParamValue("Flag_SingleLimit", "TRUE");
            //SweepCode
            function.SetParamValue("DigSrc_FlowForLoopIntegerName", SearchInfo.GetFlowForLoopIntegerName(pattern));
            function.SetParamValue("MSB_First_Flag", info.IsMsbFirst());
        }
    }
}