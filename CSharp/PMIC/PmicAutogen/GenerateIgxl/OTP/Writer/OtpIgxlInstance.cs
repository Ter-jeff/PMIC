using System.Collections.Generic;
using AutomationCommon.Utility;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.OTP.Writer
{
    public class OtpIgxlInstance
    {
        public List<InstanceSheet> GetSheets()
        {
            var instanceSheets = new List<InstanceSheet>();
            var sheets = FindInstancesSheets(InputFiles.SettingWorkbook);

            if (sheets.Count <= 0)
                return instanceSheets;

            for (var i = 0; i < sheets.Count; i++)
            {
                var sheet = InputFiles.SettingWorkbook.Worksheets[sheets[i]];
                var readInstanceSheet = new ReadInstanceSheet();
                instanceSheets.Add(readInstanceSheet.GetSheet(sheet));
            }

            return instanceSheets;
        }

        private List<string> FindInstancesSheets(ExcelWorkbook excelWorkbook)
        {
            var sheets = new List<string>();
            foreach (var worksheets in excelWorkbook.Worksheets)
                if (ComFunction.CompareString(worksheets.Name, OtpConst.InstOtpTable))
                    sheets.Add(worksheets.Name);
            return sheets;
        }
    }
}