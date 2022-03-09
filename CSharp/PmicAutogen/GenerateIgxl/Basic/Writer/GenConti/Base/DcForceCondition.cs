namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenConti.Base
{
    public class DcForceCondition
    {
        public DcForceCondition(string type, string value)
        {
            ForceType = type;
            ForceValue = value;
        }

        public string ForceType { set; get; } //Force V or Force I
        public string ForceValue { set; get; } //The Current or Voltage which need be forced
    }
}