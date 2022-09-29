using PmicAutogen.Config.ProjectConfig;
using PmicAutogen.Inputs.TestPlan.Reader;
using System;
using System.Collections.Generic;
using System.IO;

namespace PmicAutogen.Local
{
    public static class LocalSpecs
    {
        private static string _targetIgxlVersion;
        public static string TestPlanFileName { get; set; }
        public static string ScghFileName { get; set; }
        public static List<string> VbtGenToolFileNames { get; set; }
        public static string TestPlanFileNameCopy { get; set; }
        public static string ScghFileNameCopy { get; set; }
        public static List<string> VbtGenToolFileNameCopy { get; set; }
        public static string PatListCsvFile { get; set; }
        public static string YamlFileName { get; set; }
        public static List<string> OtpFileNames { get; set; }
        public static string SettingFile { get; set; }
        public static string ExtraPath { get; set; }
        private static string TargetDir { get; set; }

        public static string TarDir
        {
            get { return TargetDir; }
            set
            {
                TargetDir = value;
                if (Directory.Exists(FolderStructure.DirTrunk))
                    Directory.Delete(FolderStructure.DirTrunk, true);
            }
        }

        public static string TimeSetPath { get; set; }
        public static string PatternPath { get; set; }
        public static string BasLibraryPath { get; set; }
        public static string CurrentProject { get; set; }
        public static bool HasUltraVoltage { get; set; }
        public static bool HasUltraVoltageUHv { get; set; }
        public static bool HasUltraVoltageULv { get; set; }

        public static Dictionary<string, string> UltraVoltageCategory { get; set; }

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

        public static bool IsUnitTest { get; set; }

        public static void Initialize()
        {
            VbtGenToolFileNames = new List<string>();
            VbtGenToolFileNameCopy = new List<string>();
            OtpFileNames = new List<string>();
            UltraVoltageCategory = new Dictionary<string, string>();
            VddRefInfoList = new Dictionary<string, VddLevelsRow>();
        }

        public static string GetUltraCategory(string category)
        {
            foreach (var baseCategory in UltraVoltageCategory.Keys)
                if (baseCategory.Equals(category, StringComparison.OrdinalIgnoreCase))
                    return UltraVoltageCategory[baseCategory];
            return UltraVoltageCategory["Common"];
        }
    }

    public enum PinSelector
    {
        Nv,
        Hv,
        Lv,
        Uhv,
        Ulv
    }
}