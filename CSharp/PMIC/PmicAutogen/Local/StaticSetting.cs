using PmicAutogen.Inputs.Setting;
using System.Collections.Generic;
using System.Data;

namespace PmicAutogen.Local
{
    public static class StaticSetting
    {
        public static Dictionary<string, List<string>> JobMap;
        public static DataTable PayloadTypeTable;

        public static void AddSheets(SettingManager settingManager)
        {
            JobMap = settingManager.JobMap;
            PayloadTypeTable = settingManager.PayloadTypeTable;
        }
    }
}