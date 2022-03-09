using System.Collections.Generic;
using PmicAutogen.Config.ProjectConfig;

namespace PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase
{
    public class ConfigData
    {
        public string InstanceNamingRule;
        public Dictionary<string, string> InstSpecialSetting;
        public string NameConflictUse;
        public Dictionary<string, PinRemapping> PinMappingDic;
        public Dictionary<string, string> VbtNameMapping;

        public ConfigData()
        {
            VbtNameMapping = new Dictionary<string, string>();
            InstSpecialSetting = new Dictionary<string, string>();
            PinMappingDic = new Dictionary<string, PinRemapping>();
            InstanceNamingRule =
                ProjectConfigSingleton.Instance().GetProjectConfigValue("HardIP", "InstanceNamingRule");
            NameConflictUse = ProjectConfigSingleton.Instance().GetProjectConfigValue("HardIP", "TNameConflictUse");
        }
    }

    public class PinRemapping
    {
        public string Cp = "";
        public string Ft = "";
    }
}