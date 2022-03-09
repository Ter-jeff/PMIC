using CommonLib.EpplusErrorReport;
using CommonLib.FormatCheck;
using IgxlData.IgxlBase;
using IgxlData.IgxlReader;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using LevelSheet = IgxlData.IgxlSheets.LevelSheet;

namespace IgxlData.Validation
{
    public class IgxlValidationManager
    {
        private ExcelPackage _excelPackage = new ExcelPackage();

        public IgxlValidationManager()
        {
            var assembly = Assembly.GetExecutingAssembly();

            SheetStructureManager.Initialize(assembly);
        }

        #region Member Function
        public void CheckAll(Workbook workbook)
        {
            ExcelWorksheet excelWorksheet = _excelPackage.Workbook.Worksheets.Add("Test");

            #region Read Sheet
            var igxlSheetReader = new IgxlSheetReader();
            var levelSheets = igxlSheetReader.GetIgxlSheets(workbook, SheetType.DTLevelSheet).OfType<LevelSheet>().ToList();
            #endregion

            #region Post check
            foreach (var levelSheet in levelSheets)
            {
               if (levelSheets.Exists(x => x.Name.Equals(levelSheet.Name + "_BinCut", StringComparison.CurrentCultureIgnoreCase)))
                {
                    var binCutSheet = levelSheets.Find(x => x.Name.Equals(levelSheet.Name + "_BinCut", StringComparison.CurrentCultureIgnoreCase));
                    foreach (var row in levelSheet.LevelRows)
                    {
                        if (binCutSheet.LevelRows.Exists(x => x.PinName == row.PinName && x.Parameter == row.Parameter))
                        {
                            var bincutRow = binCutSheet.LevelRows.Find(x => x.PinName == row.PinName && x.Parameter == row.Parameter);
                            string replace = "";
                            if (levelSheet.Name.EndsWith("_Scan", StringComparison.CurrentCultureIgnoreCase))
                                replace = "_VAR_SC";
                            else if (levelSheet.Name.EndsWith("_Mbist", StringComparison.CurrentCultureIgnoreCase))
                                replace = "_VAR_BI";
                            if (!string.IsNullOrEmpty(replace))
                            {
                                var newValue = Regex.Replace(row.Value, replace, "_VAR", RegexOptions.IgnoreCase);
                                if (bincutRow.Value != newValue)
                                {
                                    var errorMessage =
                                        string.Format("Please check value {0} vs {1} for pin {2} : {3} !!!",
                                            bincutRow.Value, row.Value, row.PinName, row.Parameter);
                                    EpplusErrorManager.AddError(BasicErrorType.FormatError.ToString(), ErrorLevel.Error,
                                        binCutSheet.Name, bincutRow.RowNum, 1, errorMessage);
                                }
                            }
                        }
                        else
                        {
                            var errorMessage = string.Format("This pin {0} : {1} can not be found !!!", row.PinName, row.Parameter);
                            EpplusErrorManager.AddError(BasicErrorType.FormatError.ToString(), ErrorLevel.Error, binCutSheet.Name, 1, 1, errorMessage);
                        }
                    }
                }
            }
            #endregion
        }

        private void Workbook2ExcelWorksheet(ref ExcelWorksheet excelWorksheet, Worksheet worksheet)
        {
            object[,] data = worksheet.UsedRange.Value2;
            for (int i = 1; i <= worksheet.UsedRange.Rows.Count; i++)
            {
                for (int j = 1; j <= worksheet.UsedRange.Columns.Count; j++)
                {
                    var value = data[i, j] == null ? "" : data[i, j].ToString();
                    excelWorksheet.Cells[i, j].Value = value;
                }
            }
        }
        #endregion
    }
}
