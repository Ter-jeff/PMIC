using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CommonLib.Utility;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.VBA;
using Error = CommonLib.EpplusErrorReport.Error;

namespace CommonLib.EpplusErrorReport
{
    public static class EpplusErrorManager
    {
        private static readonly ErrorInstance ErrorInstance = ErrorInstance.Instance;

        public static void AddError(string errorType, string name, int rowNum, string message,
            params string[] comments)
        {
            Error errorNew = new Error
            {
                ErrorType = errorType,
                SheetName = name,
                RowNum = rowNum,
                Comments = comments.ToList(),
                Message = message,
                ErrorLevel = ErrorLevel.Error
            };
            AddError(errorNew);
        }

        public static void AddError(string errorType, string name, int rowNum, int colNum, string message,
            params string[] comments)
        {
            Error errorNew = new Error
            {
                ErrorType = errorType,
                SheetName = name,
                RowNum = rowNum,
                ColNum = colNum,
                Comments = comments.ToList(),
                Message = message,
                ErrorLevel = ErrorLevel.Error
            };
            AddError(errorNew);
        }

        public static void AddError(string errorType, ErrorLevel errorLevel, string name, int rowNum,
            string message, params string[] comments)
        {
            Error errorNew = new Error
            {
                ErrorType = errorType,
                SheetName = name,
                RowNum = rowNum,
                Comments = comments.ToList(),
                Message = message,
                ErrorLevel = errorLevel
            };
            AddError(errorNew);
        }

        public static void AddError(string errorType, ErrorLevel errorLevel, string name, int rowNum, int colNum,
            string message, params string[] comments)
        {
            Error errorNew = new Error
            {
                ErrorType = errorType,
                SheetName = name,
                RowNum = rowNum,
                ColNum = colNum,
                Comments = comments.ToList(),
                Message = message,
                ErrorLevel = errorLevel
            };
            AddError(errorNew);
        }

        private static void AddError(Error error)
        {
            ErrorInstance.AddError(error);
        }

        public static void ResetError()
        {
            ErrorInstance.Reset();
        }

        public static int GetErrorCountByType(string errorType)
        {
            return ErrorInstance.GetErrorCountByType(errorType);
        }

        public static int GetErrorCount()
        {
            return ErrorInstance.GetErrorCount();
        }

        public static void GenNewErrorReport(string outputFile, List<string> copyFiles)
        {
            ErrorInstance.GenErrorReport(outputFile, copyFiles);
        }

        public static void GenErrorReport(Workbook workbook, string errorReport)
        {
            var errors = ErrorInstance.GetErrorList();
            if (workbook.IsSheetExist(errorReport))
            {
                Application app = workbook.Parent;
                app.DisplayAlerts = false;
                workbook.Worksheets[errorReport].Delete();
                app.DisplayAlerts = true;
            }
            if (errors.Any())
            {
                Worksheet worksheet = workbook.AddSheet(errorReport);
                Range range = worksheet.Cells[1, 1];
                range.LoadFromCollection(errors);
                worksheet.Columns.AutoFit();
                worksheet.Select();
            }
        }

        public static void GenErrorReport(ExcelPackage excelPackage, List<string> copyFiles)
        {
            ErrorInstance.GenErrorReport(excelPackage, copyFiles);
        }

        public static void GenErrorReport(ExcelPackage excelPackage, List<string> copyFiles, string errorReprortName, string summaryReport = "SummaryReport")
        {
            ErrorInstance.GenErrorReport(excelPackage, copyFiles, errorReprortName, summaryReport);
        }

        public static bool AddMarcoFromBas(ExcelPackage excel)
        {
            bool flag = false;
            if (excel.Workbook.VbaProject == null)
                excel.Workbook.CreateVBAProject();

            const string libid = @"*\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\MSO.DLL#Microsoft Office 14.0 Object Library";
            if (!excel.Workbook.VbaProject.References.ToList().Exists(x => x.Libid.Equals(libid, StringComparison.CurrentCultureIgnoreCase)) && File.Exists(@"C:\Program Files (x86)\Common Files\Microsoft Shared\OFFICE14\MSO.DLL"))
            {
                ExcelVbaReference refer = new ExcelVbaReferenceProject();
                refer.Libid = libid;
                refer.Name = @"Office";
                excel.Workbook.VbaProject.References.Add(refer);
                flag = true;
            }

            if (excel.Workbook.VbaProject.References.ToList().Exists(x => x.Libid.Equals(libid, StringComparison.CurrentCultureIgnoreCase)))
            {
                const string moduleName = "LIB_General";
                ExcelVBAModule module = IsExistModule(excel, moduleName) ? excel.Workbook.VbaProject.Modules[moduleName] : excel.Workbook.VbaProject.Modules.AddModule(moduleName);
                List<string> lines = new List<string>();
                lines.Add("Sub GoBack()");
                lines.Add("Attribute GoBack.VB_ProcData.VB_Invoke_Func = \"q\\n14\"");
                lines.Add("    sheetName = ThisWorkbook.BuiltinDocumentProperties(\"subject\").Value");
                lines.Add("    If sheetName <> \"\" Then");
                lines.Add("    ThisWorkbook.Activate");
                lines.Add("    Sheets(sheetName).Select");
                lines.Add("    End If");
                lines.Add("End Sub");
                module.Code = string.Join("\r\n", lines);

                const string moduleName1 = "ThisWorkbook";
                ExcelVBAModule module1 = IsExistModule(excel, moduleName1) ? excel.Workbook.VbaProject.Modules[moduleName1] : excel.Workbook.VbaProject.Modules.AddModule(moduleName1);
                lines = new List<string>();
                lines.Add("Private Sub Workbook_Open()");
                lines.Add("");
                lines.Add("    Dim ContextMenu As CommandBar");
                lines.Add("    Dim ctrl As CommandBarControl");
                lines.Add("    Set ContextMenu = Application.CommandBars(\"Cell\")");
                lines.Add("    ");
                lines.Add("    ' Delete the controls first to avoid duplicates.");
                lines.Add("    For Each ctrl In ContextMenu.Controls");
                lines.Add("        If ctrl.Tag = \"My_Cell_Control_Tag\" Then");
                lines.Add("            ctrl.Delete");
                lines.Add("        End If");
                lines.Add("    Next ctrl");
                lines.Add("");
                lines.Add("    ' Add one custom button to the Cell context menu.");
                lines.Add("    With ContextMenu.Controls.Add(Type:=msoControlButton)");
                lines.Add("        .OnAction = \"'\" & ThisWorkbook.Name & \"'!\" & \"GoBack\"");
                lines.Add("        .FaceId = 59");
                lines.Add("        .Caption = \"GoBack (Ctrl+q)\"");
                lines.Add("        .Tag = \"My_Cell_Control_Tag\"");
                lines.Add("    End With");
                lines.Add("    ");
                lines.Add("    ' Add a separator");
                lines.Add("    ContextMenu.Controls(2).BeginGroup = True");
                lines.Add("    Application.MacroOptions Macro:=\"GoBack\"");
                lines.Add("End Sub");
                lines.Add("");
                lines.Add("Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)");
                lines.Add("");
                lines.Add("    Me.BuiltinDocumentProperties(\"subject\") = Sh.Name");
                lines.Add("");
                lines.Add("End Sub");
                module1.Code = string.Join("\r\n", lines);
            }
            return flag;
        }

        private static bool IsExistModule(ExcelPackage excel, string moduleName)
        {
            bool flag = false;
            foreach (var item in excel.Workbook.VbaProject.Modules)
                if (item.Name == moduleName) flag = true;
            return flag;
        }
    }
}