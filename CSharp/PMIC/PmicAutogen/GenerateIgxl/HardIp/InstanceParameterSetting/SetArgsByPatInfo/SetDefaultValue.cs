using IgxlData.VBT;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.InstanceParameterSetting.SetArgsByPatInfo
{
    public class SetDefaultValue : SetValueBase
    {
        public override void SetArgsListValue(HardIpPattern pattern, VbtFunctionBase function, string voltage)
        {
            #region If the function contains recognised paramters, try to set value for them

            function.CheckParam = false;
            SetValueBase setVbt = new SetVifValue();
            setVbt.SetArgsListValue(pattern, function, voltage);

            #endregion
        }
    }
}