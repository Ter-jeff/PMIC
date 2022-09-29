using System.Collections.Generic;
using AutomationCommon.Utility;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.OTP.Writer
{
    public class OtpIgxlWorkFlow
    {
        public List<SubFlowSheet> GetSheets()
        {
            var subFlowSheets = new List<SubFlowSheet>();
            var sheets = FindFlowSheets(InputFiles.SettingWorkbook);

            if (sheets.Count <= 0)
                return subFlowSheets;

            for (var i = 0; i < sheets.Count; i++)
            {
                var sheet = InputFiles.SettingWorkbook.Worksheets[sheets[i]];
                var readFlowSheet = new ReadFlowSheet();
                subFlowSheets.Add(readFlowSheet.GetSheet(sheet));
            }

            return subFlowSheets;
        }

        private List<string> FindFlowSheets(ExcelWorkbook excelWorkbook)
        {
            var sheets = new List<string>();
            for (var i = 1; i <= excelWorkbook.Worksheets.Count; i++)
                if (ComFunction.CompareString(excelWorkbook.Worksheets[i].Name, OtpConst.FlowOtpTable) || ComFunction.CompareString(excelWorkbook.Worksheets[i].Name, OtpConst.FlowEcid))
                    sheets.Add(excelWorkbook.Worksheets[i].Name);
            return sheets;
        }
    }
}