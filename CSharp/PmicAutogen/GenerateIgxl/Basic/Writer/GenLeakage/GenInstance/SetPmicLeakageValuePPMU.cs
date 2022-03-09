using IgxlData.VBT;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.GenerateIgxl.HardIp.InstanceParameterSetting;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenInstance
{
    public class SetPmicLeakageValuePpmu : SetValueBase
    {
        public override void SetArgsListValue(HardIpPattern pattern, VbtFunctionBase function, string voltage)
        {
            var info = SearchInfo.GetHardIpInfo(pattern);
            function.Args[0] = pattern.Pattern.GetInstancePatternName();
            function.SetParamValue("Measure_Pin_PPMU", SearchInfo.GetPpmuPin(pattern, info));
            var forceV = SearchInfo.GetForceV(pattern, info);
            function.SetParamValue("ForceV", forceV);
            function.SetParamValue("MeasureI_Range", SearchInfo.GetIRange(pattern, info, voltage));
        }
    }
}