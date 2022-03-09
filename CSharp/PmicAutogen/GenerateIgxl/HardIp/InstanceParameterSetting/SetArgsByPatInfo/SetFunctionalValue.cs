using System.Text.RegularExpressions;
using IgxlData.VBT;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.DataConvertorUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetFunctionalValue : SetValueBase
    {
        public override void SetArgsListValue(HardIpPattern pattern, VbtFunctionBase function, string voltage)
        {
            function.CheckParam = false;

            #region Set value for Functional_T_Updated

            //Patterns
            function.Args[0] = pattern.Pattern.GetInstancePatternName();

            #region Default value

            //RelayMode
            function.SetParamValue("RelayMode", "1");
            function.SetParamValue("PatternTimeout", "30");

            #endregion

            if (pattern.IsNonHardIpBlock && Regex.IsMatch(pattern.Pattern.GetLastPayload(), @"_[D]SRA*M\w*DSSC",
                RegexOptions.IgnoreCase))
                function.SetParamValue("DigSource", "Test_AutoSwitch:JTAG_TDI");
            var interposePrePat = "";

            if (pattern.ForceConditionList.Count > 0)
            {
                var forceCondition = pattern.ForceConditionList[0];
                var termInfo = "";

                #region Check multiple force type in one force condition

                foreach (var pin in forceCondition.ForcePins)
                {
                    if (!Regex.IsMatch(pin.ForceType, "TERM", RegexOptions.IgnoreCase))
                    {
                        if (pin.Type == ForceConditionType.Normal)
                        {
                            if (pin.ForceJob == "")
                                interposePrePat += pin.PinName + ":" + pin.ForceType + ":" +
                                                   DataConvertor.ConvertForceValueToGlbSpec(pin) + ";";
                            else
                                interposePrePat += pin.PinName + ":" + pin.ForceType + ":" +
                                                   DataConvertor.ConvertForceValueToGlbSpec(pin) + ":" + pin.ForceJob +
                                                   ";";
                        }

                        interposePrePat = interposePrePat.Trim(',').Replace(",", ";");
                    }
                    else
                    {
                        termInfo = pin.PinName + ":" + pin.ForceType + ":" +
                                   DataConvertor.ConvertForceValueToGlbSpec(pin) + pin.ForceJob + ";";
                    }

                    if (pin.Type == ForceConditionType.Others)
                    {
                        interposePrePat += pin.PinName + ":" + pin.ForceValue + ";";
                        interposePrePat = interposePrePat.Trim(',').Replace(",", ";");
                    }
                }

                #endregion

                interposePrePat += termInfo;
            }

            function.SetParamValue("Interpose_PrePat", interposePrePat);
            function.SetParamValue("CharInputString", interposePrePat);

            #endregion
        }
    }
}