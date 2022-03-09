using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using AutomationCommon.Utility;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.Config.ProjectConfig
{
    public class ProjectConfigSingleton
    {
        #region Constant

        public const string ConModuleName = "ModuleName";

        #endregion

        #region Initialize

        public void InitializeProjectConfigSetting()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
                if (resourceName.EndsWith(".Config.xlsx", StringComparison.CurrentCultureIgnoreCase))
                {
                    var inputExcel = new ExcelPackage(assembly.GetManifestResourceStream(resourceName));
                    InputFiles.ConfigWorkbook = inputExcel.Workbook;
                    var workSheet = InputFiles.ConfigWorkbook.Worksheets[PmicConst.ProjectConfigSetting];
                    if (workSheet != null)
                    {
                        var projectConfigSettingReader = new ProjectConfigSettingReader();
                        var sheet = projectConfigSettingReader.ReadSheet(workSheet);
                        _projectConfigSettingRows = sheet.Rows;
                    }

                    break;
                }
        }

        #endregion

        public void LoadProjectConfig()
        {
            var worksheet = InputFiles.InteropTestPlanWorkbook.GetSheet(PmicConst.ProjectConfig);
            if (worksheet != null)
            {
                var projectConfigReader = new ProjectConfigReader();
                var projectConfigSheet = projectConfigReader.ReadSheet(worksheet);
                _instance._projectConfigRows = projectConfigSheet.Rows;
            }
        }

        public void SaveProjectConfig()
        {
            var sheet = InputFiles.InteropTestPlanWorkbook.GetSheet(PmicConst.ProjectConfig);
            if (sheet == null)
            {
                InputFiles.InteropTestPlanWorkbook.AddSheet(PmicConst.ProjectConfig);
                InputFiles.InteropTestPlanWorkbook.Worksheets[PmicConst.ProjectConfig].Visible =
                    XlSheetVisibility.xlSheetHidden;
            }

            Range range = InputFiles.InteropTestPlanWorkbook.Worksheets[PmicConst.ProjectConfig].Cells[1, 1];
            range.LoadFromCollection(_projectConfigRows);
            InputFiles.InteropTestPlanWorkbook.Save();
        }

        public string ReplaceItemNameByConfigGroup(string group, string source)
        {
            var result = "";
            var values = _projectConfigRows.Where(t => t.GroupName == group).ToList();
            if (values.Count > 0)
            {
                for (var i = 0; i < values.Count; i++)
                    result = Regex.Replace(source, values[i].Name, values[i].Value, RegexOptions.IgnoreCase);
                return result;
            }

            return source;
        }

        public static void Initialize()
        {
            _instance = new ProjectConfigSingleton();
        }

        #region Singleton

        private static ProjectConfigSingleton _instance;

        private ProjectConfigSingleton()
        {
            InitializeProjectConfigSetting();
        }

        public static ProjectConfigSingleton Instance()
        {
            return _instance ?? (_instance = new ProjectConfigSingleton());
        }

        #endregion

        #region preivate field

        private List<ProjectConfigRow> _projectConfigRows = new List<ProjectConfigRow>();
        private List<ProjectConfigSettingRow> _projectConfigSettingRows;

        #endregion

        #region Get Value

        public List<ProjectConfigSettingRow> GetProjectConfigSetting()
        {
            return _projectConfigSettingRows;
        }

        public List<ProjectConfigRow> GetProjectConfigRow()
        {
            return _projectConfigRows;
        }

        public string GetProjectConfigValue(string groupName, string name)
        {
            var row = _projectConfigRows.Where(t => t.GroupName == groupName && t.Name == name).ToList();
            if (row.Count > 0)
                return row.First().Value;

            var defaultSetting = _projectConfigSettingRows.Find(t => t.Name == name);
            if (defaultSetting != null)
                return defaultSetting.Default;
            return "";
        }

        #endregion
    }
}