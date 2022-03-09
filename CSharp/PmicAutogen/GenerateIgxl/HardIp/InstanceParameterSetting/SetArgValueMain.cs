using IgxlData.NonIgxlSheets;
using IgxlData.VBT;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenInstance;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo;

namespace PmicAutogen.GenerateIgxl.HardIp.InstanceParameterSetting
{
    public class SetArgValueMain
    {
        public void SetArgsValue(HardIpPattern pattern, VbtFunctionBase function, string voltage)
        {
            SetValueBase setValueFunction;
            switch (function.FunctionName.ToLower())
            {
                case VbtFunctionLib.FunctionalTUpdated:
                {
                    setValueFunction = new SetFunctionalValue();
                    break;
                }
                case VbtFunctionLib.PmicLeakageVbtName:
                {
                    setValueFunction = new SetPmicLeakageValuePpmu();
                    break;
                }
                case VbtFunctionLib.PmicLeakageDcviVbtName:
                {
                    setValueFunction = new SetPmicLeakageValueDcvi();
                    break;
                }
                case VbtFunctionLib.PmicLeakageDcvsVbtName:
                {
                    setValueFunction = new SetPmicLeakageValueDcvs();
                    break;
                }
                default:
                {
                    setValueFunction = new SetDefaultValue();
                    break;
                }
            }

            setValueFunction.SetArgsListValue(pattern, function, voltage);
            setValueFunction.SetValueByParamMapping(function, pattern);
            setValueFunction.CheckInstArgument(function, pattern);
        }
    }
}