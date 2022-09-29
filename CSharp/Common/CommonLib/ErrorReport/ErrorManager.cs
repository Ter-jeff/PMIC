using CommonLib.Enum;
using CommonLib.Extension;
using Microsoft.Office.Interop.Excel;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CommonLib.ErrorReport
{
    public static class ErrorManager
    {
        private static readonly ErrorInstance Instance = ErrorInstance.Instance;

        public static void AddError(EnumErrorType errorType, EnumErrorLevel errorLevel, string name, int rowNum,
            string message, params string[] comments)
        {
            var errorNew = new Error
            {
                EnumErrorType = errorType,
                SheetName = name,
                RowNum = rowNum,
                Comments = comments.ToList(),
                Message = message,
                ErrorLevel = errorLevel
            };
            AddError(errorNew);
        }

        public static void AddError(EnumErrorType errorType, EnumErrorLevel errorLevel, string name, int rowNum, int colNum,
            string message, params string[] comments)
        {
            var errorNew = new Error
            {
                EnumErrorType = errorType,
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
            Instance.AddError(error);
        }

        public static void AddErrors(List<Error> errors)
        {
            Instance.AddErrors(errors);
        }

        public static void Initialize()
        {
            Instance.Initialize();
        }

        public static int GetErrorCountByType(string errorType)
        {
            return Instance.GetErrorCountByType(errorType);
        }

        public static List<Error> GetErrors()
        {
            return Instance.GetErrors();
        }

        public static int GetErrorCount()
        {
            return Instance.GetErrorCount();
        }

        public static void GenErrorReportVSTO(Workbook workbook, string errorReport)
        {
            var errors = Instance.GetErrors();
            if (workbook.IsSheetExist(errorReport))
            {
                Application app = workbook.Parent;
                app.DisplayAlerts = false;
                workbook.Worksheets[errorReport].Delete();
                app.DisplayAlerts = true;
            }

            if (errors.Any())
            {
                var worksheet = workbook.AddSheet(errorReport);
                Instance.WriteErrors(worksheet);
                worksheet.Columns.AutoFit();
                worksheet.Select();
                workbook.Save();
            }
        }

        public static void GenErrorReport(Workbook workbook, string errorReport)
        {
            var errors = Instance.GetErrors();
            if (workbook.IsSheetExist(errorReport))
            {
                Application app = workbook.Parent;
                app.DisplayAlerts = false;
                workbook.Worksheets[errorReport].Delete();
                app.DisplayAlerts = true;
            }

            if (errors.Any())
            {
                var worksheet = workbook.AddSheet(errorReport);
                Instance.WriteErrors(worksheet);
                worksheet.Columns.AutoFit();
                worksheet.Select();
                workbook.Save();
            }

            if (Path.GetExtension(workbook.FullName) == ".xlsm")
            {
                AddSetMacro(workbook);
                workbook.Save();
            }
            else
            {
                var oldFile = workbook.FullName;
                var newFile = Path.ChangeExtension(oldFile, ".xlsm");
                AddSetMacro(workbook);
                workbook.SaveAs(newFile, XlFileFormat.xlOpenXMLWorkbookMacroEnabled, AccessMode: XlSaveAsAccessMode.xlNoChange);
                File.Delete(oldFile);
            }
        }

        public static void GenErrorTxt(string outputPath)
        {
            using (StreamWriter sheetWriter = new StreamWriter(Path.Combine(outputPath, "ErrorReport.txt")))
            {
                var errors = Instance.GetErrors();
                string[] headers = { "ErrorType", "Level", "Link", "SheetName", "Row", "Col", "ErrorMessage" };
                sheetWriter.WriteLine(string.Join("\t", headers));
                foreach (var error in errors)
                {
                    var line = error.ErrorType + "\t" + error.Level + "\t" +
                        error.Link + "\t" + error.SheetName + "\t" +
                        error.RowNum + "\t" + error.ColNum + "\t" + error.Message;
                    sheetWriter.WriteLine(line);
                }
            }
        }

        public static void AddSetMacro(Workbook workbook)
        {
            const string vbtTemp = "VBT_Temp";
            var newStandardModule = workbook.GetVbComponents(vbtTemp);
            var codeModule1 = newStandardModule.CodeModule;
            codeModule1.DeleteLines(1, codeModule1.CountOfLines);
            codeModule1.Name = vbtTemp;
            codeModule1.InsertLines(1, SetGoBack());
            const string thisWorkbook = "ThisWorkbook";
            var module = workbook.GetVbComponents(thisWorkbook);
            var codeModule2 = module.CodeModule;
            codeModule2.DeleteLines(1, codeModule2.CountOfLines);
            codeModule2.Name = thisWorkbook;
            codeModule2.InsertLines(1, SetWorkbookOpen());
        }

        private static string SetGoBack()
        {
            var lines = new List<string>();
            lines.Add("Sub GoBack()");
            lines.Add("Attribute GoBack.VB_ProcData.VB_Invoke_Func = \"q\\n14\"");
            lines.Add("    sheetName = ThisWorkbook.BuiltinDocumentProperties(\"subject\").Value");
            lines.Add("    If sheetName <> \"\" Then");
            lines.Add("    ThisWorkbook.Activate");
            lines.Add("    Sheets(sheetName).Select");
            lines.Add("    End If");
            lines.Add("End Sub");
            return string.Join("\r\n", lines);
        }

        private static string SetWorkbookOpen()
        {
            var lines = new List<string>();
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
            return string.Join("\r\n", lines);
        }
    }
}