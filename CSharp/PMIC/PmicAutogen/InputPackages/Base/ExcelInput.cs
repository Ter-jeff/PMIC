using CommonLib.Enum;
using CommonLib.WriteMessage;
using OfficeOpenXml;
using PmicAutogen.InputPackages.Inputs;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace PmicAutogen.InputPackages.Base
{
    public abstract class ExcelInput : Input
    {
        public readonly List<string> SheetList;
        public List<string> SelectedSheetList;
        public List<ExcelWorksheet> SelectedSheets;
        public ExcelWorkbook Workbook;

        protected ExcelInput(FileInfo fileInfo, InputFileType inputFileType) : base(fileInfo, inputFileType)
        {
            SelectedSheetList = new List<string>();
            SelectedSheets = new List<ExcelWorksheet>();
            SheetList = new List<string>();
            if (inputFileType == InputFileType.OtpRegisterMap) return;
            var dummyExcel = Regex.Replace(fileInfo.FullName, fileInfo.Extension, "_DUM" + fileInfo.Extension);
            fileInfo.CopyTo(dummyExcel, true);
            var package = new ExcelPackage(new FileInfo(dummyExcel));
            File.Delete(dummyExcel);
            Workbook = package.Workbook;
            foreach (var sheet in Workbook.Worksheets)
            {
                SheetList.Add(sheet.Name);
                SelectedSheetList.Add(sheet.Name);
                SelectedSheets.Add(sheet);
            }
        }

        public void AnalyzeInput()
        {
            var sheetIndex = 1;
            SelectedSheets.Clear();
            SelectedSheetList.Clear();
            foreach (var sheet in Workbook.Worksheets)
            {
                if (IsValidSheet(sheet))
                {
                    SelectedSheets.Add(sheet);
                    SelectedSheetList.Add(sheet.Name);
                    Response.Report("Now Parsing Sheet: " + sheet.Name, EnumMessageLevel.General,
                        Convert.ToInt32(sheetIndex * 100 / Workbook.Worksheets.Count));

                    if (sheet.Dimension != null && sheet.Dimension.End.Row > 30000)
                        Response.Report("The number of rows in " + sheet.Name + " larger than 30000 !!!",
                            EnumMessageLevel.Warning, 0);
                }

                sheetIndex++;
            }

            Response.Report(string.Format("Parsing {0} Done", FileType), EnumMessageLevel.EndPoint, 100);
        }

        protected virtual bool IsValidSheet(ExcelWorksheet sheet)
        {
            return true;
        }
    }
}