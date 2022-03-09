using System.Collections.Generic;
using PmicAutogen.Config.ProjectConfig;
using PmicAutogen.Inputs.TestPlan.Reader;

namespace PmicAutogen.Local
{
    public static class LocalSpecs
    {
        private static string _targetIgxlVersion;
        public static string TestPlanFileName { get; set; }
        public static string ScghFileName { get; set; }
        public static List<string> VbtGenToolFileName { get; set; }
        public static string TestPlanFileNameCopy { get; set; }
        public static string ScghFileNameCopy { get; set; }
        public static List<string> VbtGenToolFileNameCopy { get; set; }
        public static string PatListCsvFile { get; set; }
        public static string YamlFileName { get; set; }
        public static List<string> OtpFileName { get; set; }
        public static string SettingFile { get; set; }
        public static string ExtraPath { get; set; }
        public static string TarDir { get; set; }
        public static string TimeSetPath { get; set; }
        public static string PatternPath { get; set; }
        public static string BasLibraryPath { get; set; }
        public static string CurrentProject { get; set; }

        public static bool HasUltraVoltage { get; set; }
        public static bool HasUltraVoltageUHv { get; set; }
        public static bool HasUltraVoltageULv { get; set; }

        public static Dictionary<string,string> UltraVoltageCategory { get; set; }

        public static Dictionary<string, VddLevelsRow> VddRefInfoList { get; set; }

        public static string TargetIgxlVersion
        {
            set { _targetIgxlVersion = value; }
            get
            {
                _targetIgxlVersion =
                    ProjectConfigSingleton.Instance().GetProjectConfigValue("IGXL", "IgxlVersion") != ""
                        ? ProjectConfigSingleton.Instance().GetProjectConfigValue("IGXL", "IgxlVersion")
                        : "8.30";
                return _targetIgxlVersion;
            }
        }

        public static void Initialize()
        {
            VbtGenToolFileName = new List<string>();
            VbtGenToolFileNameCopy = new List<string>();
            OtpFileName = new List<string>();
            UltraVoltageCategory = new Dictionary<string, string>();
            VddRefInfoList = new Dictionary<string, VddLevelsRow>();
        }

        public static string GetUltraCategory(string category)
        {
            foreach (var baseCategory in UltraVoltageCategory.Keys)
            {
                if (baseCategory.Equals(category, System.StringComparison.InvariantCultureIgnoreCase))
                {
                    return UltraVoltageCategory[baseCategory];
                }
            }

            return UltraVoltageCategory["Common"];
        }
    }

    public enum PinSelector
    {
        NV,
        HV,
        LV,
        UHV,
        ULV
    }
}