using System.Collections.Generic;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;

namespace PmicAutogen.GenerateIgxl.HardIp.SpecialSetting
{
    public class OpCodeSettingMain
    {
        public static List<FlowRow> GenOpCodeSetting(List<string> opSetting, string blockName = "", string voltage = "",
            string enable = "")
        {
            var reg = new Regex(@"\w+", RegexOptions.IgnoreCase);
            var voltageFlag = GenVoltageFlag(voltage);

            var flowRows = new List<FlowRow>();
            foreach (var setting in opSetting)
            {
                var opCodeRow = new FlowRow();
                opCodeRow.OpCode = setting.Split(':')[0];
                if (!Regex.IsMatch(opCodeRow.OpCode, "elseif|endif|if", RegexOptions.IgnoreCase))
                    opCodeRow.Enable = enable;
                opCodeRow.Parameter = setting.Split(':')[1];
                opCodeRow.Parameter = reg.Replace(opCodeRow.Parameter, delegate(Match m)
                {
                    if (Regex.IsMatch(m.Value, @"^pp_|^dd_", RegexOptions.IgnoreCase))
                        return "F_" + blockName + "_" + m.Value + voltageFlag;
                    return m.Value;
                });

                flowRows.Add(opCodeRow);
            }

            return flowRows;
        }

        private static string GenVoltageFlag(string labelVoltage)
        {
            const string flagN = "_N_Flag";
            const string flagL = "_L_Flag";
            const string flagH = "_H_Flag";
            if (string.IsNullOrEmpty(labelVoltage))
                return string.Empty;
            switch (labelVoltage)
            {
                case HardIpConstData.LabelNv:
                    return flagN;
                case HardIpConstData.LabelLv:
                    return flagL;
                case HardIpConstData.LabelHv:
                    return flagH;
                default:
                    return string.Empty;
            }
        }
    }
}