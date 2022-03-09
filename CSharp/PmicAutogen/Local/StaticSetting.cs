using System.Collections.Generic;
using System.Data;
using PmicAutogen.Inputs.Setting;

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