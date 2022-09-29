using OfficeOpenXml;
using PmicAutogen.Config.NamingRule;
using PmicAutogen.GenerateIgxl.PreAction.Reader;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Data;
using System.Reflection;

namespace PmicAutogen.Inputs.Setting
{
    public class SettingManager
    {
        public Dictionary<string, List<string>> JobMap;
        public DataTable PayloadTypeTable;

        public void CheckAll(ExcelWorkbook workbook)
        {
            #region Pre check

            var jobMappingSheet = workbook.Worksheets[PmicConst.JobMapping];
            if (jobMappingSheet != null)
            {
                var jobMapReader = new JobMapReader();
                JobMap = jobMapReader.ReadFlow(jobMappingSheet);
            }

            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
                if (resourceName.EndsWith("Scan_Config.xml", StringComparison.CurrentCultureIgnoreCase))
                {
                    var configReader = new ScanConfigFileReader();
                    PayloadTypeTable = configReader.ReadConfig(assembly.GetManifestResourceStream(resourceName));
                    break;
                }

            #endregion

            #region Post check

            #endregion
        }
    }
}