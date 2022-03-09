using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO;
using System.Drawing;
using System.Text.RegularExpressions;
using OfficeOpenXml.ConditionalFormatting;
using System.Xml;
using Library.DataStruct;

namespace Library.Output
{
    public class ReportWriter
    {
        public ReportWriter()
        {

        }

        public void GenerateDiffReport(List<InstanceCompareResult> compareResult, List<DiffResultLogRow> logDiffResultlst, string reportFilePath)
        {
            ExcelPackage epp = new ExcelPackage();
            ExcelWorkbook workbook = epp.Workbook;

            workbook.Worksheets.Add("Test Items");
            workbook.Worksheets.Add("Sub Items");
            workbook.Worksheets.Add("Compare TestName");
            ExcelWorksheet sheet1 = workbook.Worksheets["Test Items"];
            ExcelWorksheet sheet2 = workbook.Worksheets["Sub Items"];
            ExcelWorksheet sheet3 = workbook.Worksheets["Compare TestName"];

            //Edit Header  
            sheet1.Cells[1, 1].Value = "Instance Name";
            sheet1.Cells[1, 2].Value = "Device Number";
            sheet1.Cells[1, 3].Value = "Diff Type";
            sheet1.Cells[1, 4].Value = "Row(In Real DataLog)";
            sheet1.Cells[1, 5].Value = "Row(In Reference Datalog)";
            sheet1.Cells[1, 6].Value = ""; // blank column
            sheet1.Cells[1, 7].Value = "Diff Type Category";
            FormatHeader(sheet1, 1, 1, 7);

            sheet2.Cells[1, 1].Value = "Instance Name";
            sheet2.Cells[1, 2].Value = "Test Name";
            sheet2.Cells[1, 3].Value = "Site";
            sheet2.Cells[1, 4].Value = "Diff Type";
            sheet2.Cells[1, 5].Value = "Row in Base DataLog";
            sheet2.Cells[1, 6].Value = "Row In Compare Datalog";            
            sheet2.Cells[1, 7].Value = "Limit Low";
            sheet2.Cells[1, 8].Value = "Limit Hight";
            FormatHeader(sheet2, 1, 1, 8);

            sheet3.Cells[1, 1].Value = "Based Instance Name\n(" + Path.GetFileNameWithoutExtension(Common.CommonData.GetInstance().BaseTxtDatalogPath) + ")";
            sheet3.Cells[1, 2].Value = "Compared Instance Name\n(" + Path.GetFileNameWithoutExtension(Common.CommonData.GetInstance().CompareTxtDatalogPath) + ")";
            sheet3.Cells[1, 3].Value = "Test Name";
            sheet3.Cells[1, 4].Value = "Site";
            sheet3.Cells[1, 5].Value = "Diff Type";
            sheet3.Cells[1, 6].Value = "Row in Base DataLog";
            sheet3.Cells[1, 7].Value = "Row In Compare Datalog";
            sheet3.Cells[1, 8].Value = "Limit Low";
            sheet3.Cells[1, 9].Value = "Limit Hight";
            FormatHeader(sheet3, 1, 1, 9, true);

            //Edit Data
            int sheet1CurrentRow = 2;
            int sheet2CurrentRow = 2;
            int sheet3CurrentRow = 2;
            
            if (compareResult != null && compareResult.Count > 0)
            {
                foreach (InstanceCompareResult insCompResult in compareResult)
                {
                    //Edit Instance Row
                    sheet1.Cells[sheet1CurrentRow, 1].Value = insCompResult.InstanceName;
                    sheet1.Cells[sheet1CurrentRow, 2].Value = insCompResult.DeviceNumber;
                    sheet1.Cells[sheet1CurrentRow, 3].Value = insCompResult.Result.ToString();
                    sheet1.Cells[sheet1CurrentRow, 4].Value = insCompResult.Row;
                    sheet1.Cells[sheet1CurrentRow, 5].Value = insCompResult.RefLogFileRow;
                    FormatResultCell(sheet1, sheet1CurrentRow, 3, insCompResult.Result);                  
                    sheet1CurrentRow++;

                    if (insCompResult.LogDiffResultlst == null || insCompResult.LogDiffResultlst.Count == 0)
                        continue;
                    SetHypeLink(sheet1,sheet1CurrentRow - 1, 1, "Sub Items", "A" + sheet2CurrentRow);

                    //Edit Log Row
                    foreach (DiffResultLogRow logRow in insCompResult.LogDiffResultlst)
                    {
                        sheet2.Cells[sheet2CurrentRow, 1].Value = insCompResult.InstanceName;
                        sheet2.Cells[sheet2CurrentRow, 2].Value = logRow.TestName;
                        sheet2.Cells[sheet2CurrentRow, 3].Value = logRow.Site;
                        sheet2.Cells[sheet2CurrentRow, 4].Value = logRow.Result.ToString();
                        sheet2.Cells[sheet2CurrentRow, 5].Value = logRow.Row;
                        sheet2.Cells[sheet2CurrentRow, 6].Value = logRow.RefLogFileRow;            
                        sheet2.Cells[sheet2CurrentRow, 7].Value = logRow.LimitLow;
                        sheet2.Cells[sheet2CurrentRow, 8].Value = logRow.LimitHigh;
                        FormatResultCell(sheet2, sheet2CurrentRow, 4, logRow.Result);
                        logDiffResultlst.RemoveAll(s => s.TestName.Equals((logRow.TestName)));
                        //Add comment
                        if (logRow.Result == DiffResultType.Diff || logRow.Result == DiffResultType.LimitChange)
                        {                            
                            if (logRow.RefLimitLow != null)
                            {
                                EditDiffValue(sheet2, sheet2CurrentRow, 7, logRow.RefLimitLow);
                            }
                            if (logRow.RefLimitHigh != null)
                            {
                                EditDiffValue(sheet2, sheet2CurrentRow, 8, logRow.RefLimitHigh);
                            }
                        }
                        sheet2CurrentRow++;
                    }

                }

                // add Diff Type Category
                sheet1CurrentRow = 2;
                List<DiffResultType> resultList = compareResult.Select(p => p.Result).Distinct().ToList();
                foreach (var result in resultList)
                {
                    sheet1.Cells[sheet1CurrentRow, 7].Value = result.ToString();
                    FormatResultCell(sheet1, sheet1CurrentRow++, 7, result);
                }
            }

            foreach (DiffResultLogRow TestNameComResult in logDiffResultlst)
            {
                sheet3.Cells[sheet3CurrentRow, 1].Value = TestNameComResult.BasedInst;
                sheet3.Cells[sheet3CurrentRow, 2].Value = TestNameComResult.ComparedInst;
                sheet3.Cells[sheet3CurrentRow, 3].Value = TestNameComResult.TestName;
                sheet3.Cells[sheet3CurrentRow, 4].Value = TestNameComResult.Site;
                sheet3.Cells[sheet3CurrentRow, 5].Value = TestNameComResult.Result.ToString();
                sheet3.Cells[sheet3CurrentRow, 6].Value = TestNameComResult.Row;
                sheet3.Cells[sheet3CurrentRow, 7].Value = TestNameComResult.RefLogFileRow;
                sheet3.Cells[sheet3CurrentRow, 8].Value = TestNameComResult.LimitLow;
                sheet3.Cells[sheet3CurrentRow, 9].Value = TestNameComResult.LimitHigh;
                FormatResultCell(sheet3, sheet3CurrentRow, 3, TestNameComResult.Result);
                //Add comment
                if (TestNameComResult.Result == DiffResultType.Diff)
                {
                    if (TestNameComResult.RefLimitLow != null)
                    {
                        EditDiffValue(sheet3, sheet3CurrentRow, 8, TestNameComResult.RefLimitLow);
                    }
                    if (TestNameComResult.RefLimitHigh != null)
                    {
                        EditDiffValue(sheet3, sheet3CurrentRow, 9, TestNameComResult.RefLimitHigh);
                    }
                }
                sheet3CurrentRow++;
            }

            // set HyperLink for Compare TestName items, must be put after sheet3 calculated done
            int currentRow = 2;
            if (compareResult != null && compareResult.Count > 0)
            {
                foreach (InstanceCompareResult insCompResult in compareResult)
                {
                    ++currentRow;
                    if (insCompResult.LogDiffResultlst == null || insCompResult.LogDiffResultlst.Count == 0)
                    {
                        string columnStr = "";
                        int rowIndex = -1;
                        GetColumnIndexByInstanceName(sheet3, insCompResult, ref columnStr, ref rowIndex);
                        if (rowIndex != -1) SetHypeLink(sheet1, currentRow - 1, 1, "Compare TestName", columnStr + rowIndex);
                        continue;
                    }
                }
            }

            logDiffResultlst.Clear();

            var file = new FileInfo(reportFilePath);
            epp.SaveAs(file);
            epp.Dispose();
        }

        private void GetColumnIndexByInstanceName(ExcelWorksheet sheet, InstanceCompareResult insCompResult, ref string columnStr, ref int rowIndex)
        {
            int column = -1;
            if (insCompResult.Result == DiffResultType.OnlyInBaseDatalog)
            {
                columnStr = "A";
                column = 1;
            }
            else
            {
                columnStr = "B";
                column = 2;
            }

            for (int i = 2; i <= sheet.Dimension.End.Row; ++i)
            {
                if (sheet.Cells[i, column].Value.Equals(insCompResult.InstanceName))
                {
                    rowIndex = i;
                    break;
                }
            }
        }

        private void FormatHeader(ExcelWorksheet sheet, int row, int startColumn, int endColumn, bool wrap = false)
        {
            for (int column = startColumn; column <= endColumn; column++)
            {
                sheet.Cells[row, column].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Gray);
                sheet.Cells[row, column].Style.Font.Color.SetColor(Color.White);
                sheet.Cells[row, column].Style.Font.Size = 12;
                sheet.Cells[row, column].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
                sheet.Cells[row, column].AutoFitColumns();
                if (wrap) sheet.Cells[row, column].Style.WrapText = true;
            }
        }

        private void SetHypeLink(ExcelWorksheet sheet, int row, int column, string linksheetName, string linkaddress)
        {
            sheet.Cells[row, column].Hyperlink = new ExcelHyperLink((char)39 + linksheetName +
                        (char)39 + "!" + linkaddress, sheet.Cells[row, column].Value.ToString());
            sheet.Cells[row, column].Style.Font.UnderLine = true;
            sheet.Cells[row, column].Style.Font.Color.SetColor(Color.Blue);
        }

        private void EditDiffValue(ExcelWorksheet sheet, int row, int column, string comment)
        {
            var xlsComment = sheet.Cells[row, column].AddComment(comment, "AutoRun");
            sheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Red);
            try
            {
                int length = comment.Length;
                if (length <= 20)
                {
                    xlsComment.From.Column = 1;
                    xlsComment.To.Column = 2;
                    xlsComment.From.Row = 1;
                    xlsComment.To.Row = 4;
                }
                else if (length <= 100)
                {
                    xlsComment.From.Column = 1;
                    xlsComment.To.Column = 4;
                    xlsComment.From.Row = 1;
                    xlsComment.To.Row = 4;
                }
                else if (length <= 500)
                {
                    xlsComment.From.Column = 1;
                    xlsComment.To.Column = 6;
                    xlsComment.From.Row = 1;
                    xlsComment.To.Row = 10;
                }
                else
                {
                    xlsComment.From.Column = 1;
                    xlsComment.To.Column = 10;
                    xlsComment.From.Row = 1;
                    xlsComment.To.Row = 10;
                }
            }
            catch (Exception e) {
                ;
            }

            //xlsComment.AutoFit = true;
            //var xlsComment = sheet.Cells[row, column].AddComment(comment, "AutoRun");
            //xlsComment.Font.Bold = true;
            //var richTextComment = xlsComment.RichText.Add("Bruce extend comment");
            //richTextComment.Bold = true;
            //xlsComment.From.Column = 1;
            //xlsComment.To.Column = 2;
            //xlsComment.From.Row = 1;
            //xlsComment.To.Row = 2;
            //xlsComment.BackgroundColor = Color.Aqua;
            //xlsComment.RichText.Add("Bruce Setting Comment");
            //xlsComment.AutoFit = true;
           
        }

        private void FormatResultCell(ExcelWorksheet sheet, int row, int column, DiffResultType result)
        {
            sheet.Cells[row, column].Style.Fill.PatternType = ExcelFillStyle.Solid;
            if (result == DiffResultType.Diff)
               sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Red);
            if (result == DiffResultType.LimitChange)
                sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Red);
            if (result == DiffResultType.TestItemMismatch)
                sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80));
            if (result == DiffResultType.OnlyInCompareDatalog)
                sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
            if (result == DiffResultType.OnlyInBaseDatalog)
                sheet.Cells[row, column].Style.Fill.BackgroundColor.SetColor(Color.Orange);
        }
    }
}
