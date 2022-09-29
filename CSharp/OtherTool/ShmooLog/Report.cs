using OfficeOpenXml;
using OfficeOpenXml.Style;
using ShmooLog.Base;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ShmooLog
{
    public class Shmoo2DRowData
    {
        public Shmoo2DRowData()
        {
            AllShmoo = new List<ShmooContent>();
        }

        public List<ShmooContent> AllShmoo { get; set; }
    }

    public class ShmooContent
    {
        public bool IsSkip;

        public string Content { get; set; }
        public string DieInfo { get; set; }
    }

    public static class HandleExcel
    {
        public static void GenerateShmooReport(string strFilePath, ShmooSets shmooSets)
        {
            if (File.Exists(strFilePath)) File.Delete(strFilePath);

            var rgexSpecialSheet =
                new Regex(
                    @"Shmoo_Setup|Summary_1D|Summary_2D|ShmooHole_1D|AllPassOrFail|Summary_Abnormal|2D_LVCCHVCC|ShmooAlarm|SelSramDigSrcCheck",
                    RegexOptions.Compiled | RegexOptions.IgnoreCase);

            var ep = new ExcelPackage(); //在記憶體中建立一個Excel物件
            CreateDefaultNamedStyleInWorkBook(ref ep,
                "Shmoo"); //Style可以預先建立 然後在同一個Workbook內取用 但Conditional Formatting是跟著Sheet的

            //var listSpecialSheets = new List<string>();
            var dicHyperPoint = new Dictionary<string, int[]>(); //為了給HyperLink用的 test number當Key

            var currDtCount = 1;
            foreach (DataTable dt in
                     shmooSets.CurrShmooReport
                         .Tables) //Shmoo_Setup|Summary_1D|Summary_2D|ShmooHole_1D 以及各個Category(此時還沒有區分1D/2D)
            {
                if (dt.Rows.Count == 0) continue;

                var sheetName = "sheet1";
                if (dt.TableName != null || dt.TableName != string.Empty)
                {
                    if (dt.TableName != null && dt.TableName.Length < 32) sheetName = dt.TableName;
                    else if (dt.TableName != null) sheetName = dt.TableName.Substring(0, 32);
                }

                if (Regex.IsMatch(sheetName, @"1DShmooFreqSummary", RegexOptions.IgnoreCase))
                {
                    ep.Workbook.Worksheets.Add(sheetName); //加入一個Sheet
                    var currSheet = ep.Workbook.Worksheets[sheetName];

                    GenShmooFreqSumRpt(dt, currSheet);
                }
                else if (rgexSpecialSheet.IsMatch(sheetName)) //特殊Sheet
                {
                    //listSpecialSheets.Add(sheetName); //準備事後要再補印Test Num HyperLink

                    ep.Workbook.Worksheets.Add(sheetName); //加入一個Sheet
                    var currSheet = ep.Workbook.Worksheets[sheetName]; //取得剛剛加入的Sheet

                    for (var currColNum = 1; currColNum < dt.Columns.Count + 1; currColNum++) //印標題列
                    {
                        var HeaderName = dt.Columns[currColNum - 1].ColumnName;
                        currSheet.Cells[1, currColNum].Value = HeaderName; //加入標頭

                        if (HeaderName == "Pattern List")
                        {
                            currSheet.Column(currColNum).Style.WrapText = true;
                            currSheet.Column(currColNum).Width = 100;
                        }
                        else if (HeaderName == "LVCC List" || HeaderName == "HVCC List")
                        {
                            currSheet.Column(currColNum).Style.WrapText = true;
                            currSheet.Column(currColNum).Width = 50;
                        }
                        else if (HeaderName == "X,Y Axis")
                        {
                            currSheet.Column(currColNum).Style.WrapText = true;
                            currSheet.Column(currColNum).Width = 50;
                        }
                    }

                    currSheet.Cells[1, 1, 1, dt.Columns.Count].StyleName = "Title Row";
                    currSheet.Cells[1, 1, 1, dt.Columns.Count].Style.Border
                        .BorderAround(ExcelBorderStyle.Thin, Color.Black); //這一行沒辦法加到NamedStyle裡面...

                    for (var currRowNum = 2; currRowNum < dt.Rows.Count + 2; currRowNum++)
                        for (var currColNum = 1; currColNum < dt.Columns.Count + 1; currColNum++)
                            if (DBNull.Value != dt.Rows[currRowNum - 2][currColNum - 1])
                            {
                                currSheet.Cells[currRowNum, currColNum].Value = dt.Rows[currRowNum - 2][currColNum - 1];
                                currSheet.Cells[currRowNum, currColNum].Style.HorizontalAlignment =
                                    ExcelHorizontalAlignment.Center;
                                currSheet.Cells[currRowNum, currColNum].Style.VerticalAlignment =
                                    ExcelVerticalAlignment.Center;
                            }

                    var adr = new ExcelAddress(2, 1, dt.Rows.Count + 1, dt.Columns.Count);
                    if (sheetName != "Summary_Abnormal")
                    {
                        var ruleFail = currSheet.ConditionalFormatting.AddEqual(adr);
                        ruleFail.Formula = "0";
                        //rulePass.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ruleFail.Style.Fill.BackgroundColor.Color = Color.Red; //Color.MediumSpringGreen;
                        //rulePass.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        //rulePass.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        //rulePass.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        //rulePass.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    }

                    if (sheetName == "SelSramDigSrcCheck")
                    {
                        var cf = currSheet.ConditionalFormatting.AddContainsText(adr);
                        cf.Text = "(F)";
                        cf.Style.Fill.BackgroundColor.Color = Color.Red;

                        var cp = currSheet.ConditionalFormatting.AddContainsText(adr);
                        cp.Text = "(P)";
                        cp.Style.Fill.BackgroundColor.Color = Color.LightGreen;
                    }
                }
                else //一般1D 2D Shmoo圖
                {
                    //同一個Category有可能同時存在1D 2D
                    var shmooCount1D =
                        (from r in dt.AsEnumerable() where r.Field<string>("Shmoo Type") == "1D" select r).ToList();
                    if (shmooCount1D.Count > 0) ep.Workbook.Worksheets.Add(sheetName + "_1D"); //加入一個Sheet

                    var shmooCount2D =
                        (from r in dt.AsEnumerable() where r.Field<string>("Shmoo Type") == "2D" select r).ToList();
                    if (shmooCount2D.Count > 0) ep.Workbook.Worksheets.Add(sheetName + "_2D"); //加入一個Sheet

                    foreach (DataRow dtRow in dt.Rows)
                    {
                        var shmooType = (string)dtRow["Shmoo Type"];

                        var currSheet = ep.Workbook.Worksheets[sheetName + "_" + shmooType]; //取得剛剛加入的Sheet
                        //var dim = currSheet.Dimension;

                        if (shmooType == "1D")
                            PlotShmoo1D(dtRow, ref currSheet, ref dicHyperPoint); //所有繪圖工作統一由此處理
                        else
                            PlotShmoo2D(dtRow, shmooSets.MergeColorSetting, ref currSheet, ref dicHyperPoint);
                    }

                    if (shmooCount1D.Count > 0)
                    {
                        var currSheet = ep.Workbook.Worksheets[sheetName + "_1D"]; //取得剛剛加入的Sheet
                        currSheet.View.ShowGridLines = false;
                        currSheet.Column(1).Hidden = true;
                    }

                    if (shmooCount2D.Count > 0)
                    {
                        var currSheet = ep.Workbook.Worksheets[sheetName + "_2D"]; //取得剛剛加入的Sheet
                        currSheet.View.ShowGridLines = false;
                    }
                }

                //args.Percentage = Convert.ToInt16(currDtCount * 100 / shmooSets.CurrShmooReport.Tables.Count);
                //args.Result = string.Format("Plotting " + dt.TableName + " Done!!");
                //progress.Report(args);
                currDtCount++;
            } //Each DataTable

            //args.Percentage = 95;
            //args.Result = string.Format("Updating TestNum HyperLinks!!");
            //progress.Report(args);

            //回頭處理HyperLink
            foreach (var shtName in new List<string>
                     {
                         "Shmoo_Setup", "Summary_1D", "Summary_2D", "ShmooHole_1D", "AllPassOrFail", "Summary_Abnormal",
                         "2D_LVCCHVCC", "ShmooAlarm"
                     })
            {
                var sheet = ep.Workbook.Worksheets[shtName];
                if (sheet == null) continue;

                var table = shmooSets.CurrShmooReport.Tables[shtName];
                for (var currRowNum = 2; currRowNum < table.Rows.Count + 2; currRowNum++)
                {
                    var sheetName = (string)table.Rows[currRowNum - 2]["Category"] + "_" +
                                    (string)table.Rows[currRowNum - 2]["Type"];
                    var testNum = Convert.ToInt32(table.Rows[currRowNum - 2]["Test Num"]);
                    var testNumInstance = table.Rows[currRowNum - 2]["Test Num"] + ":" +
                                          table.Rows[currRowNum - 2]["Test Instance"];
                    var cellPos = ExcelCellBase.GetAddress(dicHyperPoint[testNumInstance][0],
                        dicHyperPoint[testNumInstance][1]);
                    sheet.Cells[currRowNum, 1].Hyperlink =
                        new ExcelHyperLink(sheetName + @"!" + cellPos, testNum.ToString());
                    sheet.Cells[currRowNum, 1].Style.Font.UnderLine = true;
                    sheet.Cells[currRowNum, 1].Style.Font.Color.SetColor(Color.Blue);
                }

                for (var currColNum = 1; currColNum < table.Columns.Count + 1; currColNum++) //印標題列
                {
                    var columnName = table.Columns[currColNum - 1].ColumnName;
                    if (columnName == "Pattern List")
                        sheet.Column(currColNum).Width = 100;
                    else if (columnName == "LVCC List" || columnName == "HVCC List")
                        sheet.Column(currColNum).Width = 30;
                    else if (columnName == "X,Y Axis")
                        sheet.Column(currColNum).Width = 20;
                    else sheet.Column(currColNum).AutoFit();
                }

                //sheet.Cells[sheet.Dimension.Address].AutoFitColumns(); //Column最佳化
                sheet.View.FreezePanes(2, 4);
                sheet.Cells[sheet.Dimension.Address].AutoFilter = true;
            }

            // Auto modified sheet 
            foreach (var shtName in new List<string> { "1DShmooFreqSummary", "SelSramDigSrcCheck" })
            {
                var sheet = ep.Workbook.Worksheets[shtName];
                if (sheet != null)
                {
                    sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
                    sheet.View.FreezePanes(2, 4);
                }
            }


            //args.Percentage = 100;
            //args.Result = string.Format("Generating Final Shmoo Report Xlsx!!");
            //progress.Report(args);

            var outputStream =
                new FileStream(strFilePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite); //建立檔案串流

            ep.SaveAs(outputStream); //把剛剛的Excel物件真實存進檔案裡

            outputStream.Close(); //關閉串流

            ep.Dispose(); //關閉資源
        }

        private static void GenShmooFreqSumRpt(DataTable dt, ExcelWorksheet currSheet)
        {
            var lvccCols = new List<int>();
            var hvccCols = new List<int>();

            for (var currColNum = 1; currColNum < dt.Columns.Count + 1; currColNum++) //印標題列
            {
                var HeaderName = dt.Columns[currColNum - 1].ColumnName;
                currSheet.Cells[1, currColNum].Value = HeaderName; //加入標頭

                if (Regex.IsMatch(HeaderName, "HVCC", RegexOptions.IgnoreCase)) hvccCols.Add(currColNum);

                if (Regex.IsMatch(HeaderName, "LVCC", RegexOptions.IgnoreCase)) lvccCols.Add(currColNum);


                if (HeaderName == "Pattern List")
                {
                    currSheet.Column(currColNum).Style.WrapText = true;
                    currSheet.Column(currColNum).Width = 100;
                }
                else if (HeaderName == "LVCC List" || HeaderName == "HVCC List")
                {
                    currSheet.Column(currColNum).Style.WrapText = true;
                    currSheet.Column(currColNum).Width = 50;
                }
                else if (HeaderName == "X,Y Axis")
                {
                    currSheet.Column(currColNum).Style.WrapText = true;
                    currSheet.Column(currColNum).Width = 50;
                }
            }

            currSheet.Cells[1, 1, 1, dt.Columns.Count].StyleName = "Title Row";
            currSheet.Cells[1, 1, 1, dt.Columns.Count].Style.Border
                .BorderAround(ExcelBorderStyle.Thin, Color.Black); //這一行沒辦法加到NamedStyle裡面...

            var adr = new ExcelAddress(2, lvccCols.Min(), dt.Rows.Count + 1, hvccCols.Max());

            var rule9999 = currSheet.ConditionalFormatting.AddEqual(adr);
            rule9999.Formula = "9999";
            rule9999.Style.Fill.BackgroundColor.Color = Color.White; //Color.MediumSpringGreen;

            var rulem9999 = currSheet.ConditionalFormatting.AddEqual(adr);
            rulem9999.Formula = "-9999";
            rulem9999.Style.Fill.BackgroundColor.Color = Color.White; //Color.MediumSpringGreen;

            var rule5555 = currSheet.ConditionalFormatting.AddEqual(adr);
            rule5555.Formula = "5555";
            //rulePass.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule5555.Style.Fill.BackgroundColor.Color = Color.White; //Color.MediumSpringGreen;

            var rulem5555 = currSheet.ConditionalFormatting.AddEqual(adr);
            rulem5555.Formula = "-5555";
            //rulePass.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rulem5555.Style.Fill.BackgroundColor.Color = Color.White; //Color.MediumSpringGreen;


            var lvccValueSet = new HashSet<double>();
            var hvccValueSet = new HashSet<double>();
            for (var currRowNum = 2; currRowNum < dt.Rows.Count + 2; currRowNum++)
            {
                for (var currColNum = 1; currColNum < dt.Columns.Count + 1; currColNum++)
                {
                    currSheet.Cells[currRowNum, currColNum].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    currSheet.Cells[currRowNum, currColNum].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    if (DBNull.Value != dt.Rows[currRowNum - 2][currColNum - 1])
                    {
                        currSheet.Cells[currRowNum, currColNum].Value = dt.Rows[currRowNum - 2][currColNum - 1];

                        if (lvccCols.Contains(currColNum))
                            lvccValueSet.Add((double)dt.Rows[currRowNum - 2][currColNum - 1]);
                        if (hvccCols.Contains(currColNum))
                            hvccValueSet.Add((double)dt.Rows[currRowNum - 2][currColNum - 1]);
                    }
                }

                var formatLvccRule =
                    currSheet.ConditionalFormatting.AddTwoColorScale(currSheet.Cells[currRowNum, lvccCols.Min(),
                        currRowNum, lvccCols.Max()]);
                var formatHvccRule =
                    currSheet.ConditionalFormatting.AddTwoColorScale(currSheet.Cells[currRowNum, hvccCols.Min(),
                        currRowNum, hvccCols.Max()]);

                if (lvccValueSet.Count() > 1)
                {
                    formatLvccRule.LowValue.Color = Color.FromArgb(255, 239, 156);
                    formatLvccRule.HighValue.Color = Color.FromArgb(99, 190, 123);
                }
                else
                {
                    formatLvccRule.HighValue.Color = Color.FromArgb(255, 239, 156);
                    formatLvccRule.LowValue.Color = Color.FromArgb(99, 190, 123);
                }


                if (hvccValueSet.Count() > 1)
                {
                    formatHvccRule.LowValue.Color = Color.FromArgb(255, 239, 156);
                    formatHvccRule.HighValue.Color = Color.FromArgb(99, 190, 123);
                }
                else
                {
                    formatHvccRule.HighValue.Color = Color.FromArgb(255, 239, 156);
                    formatHvccRule.LowValue.Color = Color.FromArgb(99, 190, 123);
                }


                hvccValueSet.Clear();
                lvccValueSet.Clear();
            }
        }

        private static void PlotShmoo1D(DataRow dtRow, ref ExcelWorksheet currSheet,
            ref Dictionary<string, int[]> dicHyperPoint)
        {
            var rgexComma = new Regex(@",", RegexOptions.Compiled);

            var dim = currSheet.Dimension;

            var resD = 0.0;

            //起始定位點
            var anchorRow = dim == null ? 1 : dim.End.Row + 3;
            var anchorCol = dim == null ? 1 : dim.Start.Column;
            var currRow = anchorRow;
            //var currCol = anchorCol;
            //*****************************************************
            // BinCut Spec infor
            var isBinCutSpec = false;
            var binCutSpec = new List<int>();
            var binCutSpecName = new List<string>();
            var binCutPmode = "";
            var binCutInfoRow = 0;
            var binCutFile = "";
            var binCutVersion = "";
            if ((string)dtRow["BinCutSpec"] != "N/A")
            {
                isBinCutSpec = true;
                binCutSpec = Array.ConvertAll(((string)dtRow["BinCutSpec"]).Split(','), int.Parse).ToList();
                binCutSpecName = ((string)dtRow["BinCutSpecName"]).Split(',').ToList();
                binCutPmode = (string)dtRow["BinCutPmode"];
                binCutFile = (string)dtRow["BinCutPlan"];
                binCutVersion = (string)dtRow["BinCutVersion"];
                binCutInfoRow = 1;
            }
            //**********************************


            //以下兩個加起來就是Max Column
            int shmooStep = Convert.ToInt16(dtRow["Shmoo Step"]);
            var infoHeaders = ((string)dtRow["Info Headers"]).Split(','); //那些Site Lot...
            var titles = ((string)dtRow["Titles"]).Split(':');
            //****************************************************************************************************************************
            //列印Test Number
            var testNum = Convert.ToInt32(titles[1]);
            var testNumInstance = titles[1] + ":" + titles[3];
            currSheet.Cells[currRow, anchorCol + 1].Value = "TestNum";
            currSheet.Cells[currRow, anchorCol + 2].Value = testNum;
            currSheet.Cells[currRow, anchorCol + 1, currRow, anchorCol + 2].StyleName = "Shmoo Axis Label";
            currSheet.Cells[currRow, anchorCol + 1, currRow, anchorCol + 2].Style.Font.Bold = true;
            dicHyperPoint[testNumInstance] = new int[2];
            //****************************************************************************************************************************
            var patterns = ((string)dtRow["Patterns"]).Split(','); //用群組印在Title之上
            currSheet.Cells[currRow, anchorCol + infoHeaders.Length].Value = "Pattern List";
            currSheet.Cells[currRow, anchorCol + infoHeaders.Length].StyleName = "Pattern List";
            currSheet.Cells[currRow, anchorCol + infoHeaders.Length, currRow,
                anchorCol + infoHeaders.Length + shmooStep - 1].Merge = true;
            //****************************************************************************************************************************
            //印Pattern List
            currRow = anchorRow + 1;
            for (var i = 0; i < patterns.Length; i++)
            {
                //currSheet.Cells[currRow + i, anchorCol + infoHeaders.Length - 1].Value = "#" + (i+1).ToString();
                currSheet.Cells[currRow + i, anchorCol + infoHeaders.Length].Value = patterns[i];
                currSheet.Row(currRow + i).OutlineLevel = 1;
                currSheet.Cells[currRow + i, anchorCol + infoHeaders.Length, currRow + i,
                    anchorCol + infoHeaders.Length + shmooStep - 1].Merge = true;
            }

            currSheet.Cells[currRow, anchorCol + infoHeaders.Length, currRow + patterns.Length,
                anchorCol + infoHeaders.Length + shmooStep].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            //****************************************************************************************************************************
            currRow = anchorRow + patterns.Length + 1;
            //FileName, Site, LotId, ....
            for (var i = 0; i < infoHeaders.Length; i++) currSheet.Cells[currRow, anchorCol + i].Value = infoHeaders[i];
            currSheet.Cells[currRow, anchorCol, currRow, anchorCol + infoHeaders.Length - 1].StyleName =
                "Shmoo Info Header";

            //列印Title
            var BinCutTitle = "";
            if (isBinCutSpec)
                BinCutTitle = "  BinCut Plan: " + binCutFile + "  Performance Mode: " + binCutPmode;
            currSheet.Cells[currRow, anchorCol + infoHeaders.Length].Value =
                "Test Instance: " + titles[3] + "   Setup Name: " + titles[5] + BinCutTitle;
            currSheet.Cells[currRow, anchorCol + infoHeaders.Length, currRow,
                anchorCol + infoHeaders.Length + shmooStep - 1].Merge = true;
            currSheet.Cells[currRow, anchorCol + infoHeaders.Length, currRow,
                anchorCol + infoHeaders.Length + shmooStep - 1].StyleName = "1D Shmoo Title";
            //****************************************************************************************************************************
            //印Per Die資訊
            currRow = anchorRow + patterns.Length + 2;

            var maxInfoLength = 0;
            var dieInfo = ((string)dtRow["Die Info"]).Split('#'); //Per Die
            for (var i = 0; i < dieInfo.Length; i++)
            {
                var infos = dieInfo[i].Split('|');
                if (infos.Length > maxInfoLength) maxInfoLength = infos.Length;
                for (var j = 0; j < infos.Length; j++)
                {
                    if (rgexComma.IsMatch(infos[j]))
                    {
                        currSheet.Cells[currRow, anchorCol + j].Value = infos[j];
                        continue;
                    }

                    if (double.TryParse(infos[j], out resD))
                        currSheet.Cells[currRow, anchorCol + j].Value = resD;
                    else
                        currSheet.Cells[currRow, anchorCol + j].Value = infos[j];
                }

                currRow++;
            }

            currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol, currRow, anchorCol + maxInfoLength].Style
                .HorizontalAlignment = ExcelHorizontalAlignment.Center;

            //****************************************************************************************************************************
            currRow = anchorRow + patterns.Length + 2;

            var allContents = ((string)dtRow["All Content"]).Split('#');

            for (var i = 0; i < allContents.Length; i++)
            {
                var shmooContent = allContents[i];

                if (rgexComma.IsMatch(shmooContent)) //代表有疊圖的資訊 
                {
                    var vv = rgexComma.Split(shmooContent);

                    var collectPassRate = new List<double>();

                    for (var j = 0; j < vv.Length; j++)
                    {
                        currSheet.Cells[currRow, anchorCol + maxInfoLength + j].Value = Convert.ToDouble(vv[j]);
                        collectPassRate.Add(Convert.ToDouble(vv[j]));
                    }

                    var ruleOverlay = currSheet.ConditionalFormatting.AddTwoColorScale(new ExcelAddress(currRow,
                        anchorCol + maxInfoLength, currRow, anchorCol + maxInfoLength + vv.Length - 1));

                    ruleOverlay.HighValue.Color = Color.LimeGreen;
                    //ruleOverlay.HighValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                    ruleOverlay.HighValue.Value = 100.0;

                    if (collectPassRate.Sum() < 0.0001) ruleOverlay.HighValue.Color = Color.Red;

                    ruleOverlay.LowValue.Color = Color.Red;
                    //ruleOverlay.LowValue.Type = eExcelConditionalFormattingValueObjectType.Percent;
                    ruleOverlay.LowValue.Value = 0.0;
                    if (collectPassRate.Sum() > 100.0 * collectPassRate.Count - 0.001)
                        ruleOverlay.HighValue.Color = Color.LimeGreen;
                }
                else
                {
                    for (var j = 0; j < shmooContent.Length; j++)
                    {
                        var v = shmooContent[j];
                        currSheet.Cells[currRow, anchorCol + maxInfoLength + j].Value = v.ToString();
                    }
                }

                currRow++;
            }
            //****************************************************************************************************************************

            currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength, currRow - 1,
                    anchorCol + maxInfoLength + shmooStep - 1].Style.Border.Bottom.Style =
                ExcelBorderStyle.Thin; //這一行沒辦法加到NamedStyle裡面...
            currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength, currRow - 1,
                anchorCol + maxInfoLength + shmooStep - 1].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength, currRow - 1,
                anchorCol + maxInfoLength + shmooStep - 1].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength, currRow - 1,
                anchorCol + maxInfoLength + shmooStep - 1].Style.Border.Right.Style = ExcelBorderStyle.Thin;

            var adr = new ExcelAddress(anchorRow + patterns.Length + 1, anchorCol + maxInfoLength, currRow,
                anchorCol + maxInfoLength + shmooStep - 1);
            var rulePass = currSheet.ConditionalFormatting.AddEqual(adr);
            rulePass.Formula = "\"P\"";
            rulePass.Style.Fill.BackgroundColor.Color = Color.LimeGreen;
            var ruleFail = currSheet.ConditionalFormatting.AddEqual(adr);
            ruleFail.Formula = "\"F\"";
            ruleFail.Style.Fill.BackgroundColor.Color = Color.Red;
            var ruleFail1 = currSheet.ConditionalFormatting.AddEqual(adr);
            ruleFail1.Formula = "\"S\"";
            ruleFail1.Style.Fill.BackgroundColor.Color = Color.Red;
            var ruleFail2 = currSheet.ConditionalFormatting.AddEqual(adr);
            ruleFail2.Formula = "\"B\"";
            ruleFail2.Style.Fill.BackgroundColor.Color = Color.Red;
            var ruleFail3 = currSheet.ConditionalFormatting.AddEqual(adr);
            ruleFail3.Formula = "\"C\"";
            ruleFail3.Style.Fill.BackgroundColor.Color = Color.Red;
            var ruleAssumedPass = currSheet.ConditionalFormatting.AddEqual(adr);
            ruleAssumedPass.Formula = "\"*\"";
            ruleAssumedPass.Style.Fill.BackgroundColor.Color = Color.PaleGreen;
            var ruleAssumedFail = currSheet.ConditionalFormatting.AddEqual(adr);
            ruleAssumedFail.Formula = "\"~\"";
            ruleAssumedFail.Style.Fill.BackgroundColor.Color = Color.Tomato;


            //*****************************************************


            //****************************************************************************************************************************
            //把Spec的點標上去!!
            var listSpecX = new List<int>();
            if ((string)dtRow["Spec X Point"] != "N/A")
                listSpecX = Array.ConvertAll(((string)dtRow["Spec X Point"]).Split(','), int.Parse).ToList();

            //for (var i = 0; i < listSpecX.Count; i++)
            //{
            //    var p = listSpecX[i];
            //    var specColor = i == 0 ? Color.Blue : Color.Yellow;
            //    //currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength + p, currRow - 1, anchorCol + maxInfoLength + p].AddComment(specInfoPoints[i], "");
            //    //currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength + p, currRow - 1, anchorCol + maxInfoLength + p].Style.Border.BorderAround(ExcelBorderStyle.MediumDashDot);
            //    //currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength + p, currRow - 1, anchorCol + maxInfoLength + p].Style.Border.Bottom.Color.SetColor(specColor);
            //    //currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength + p, currRow - 1, anchorCol + maxInfoLength + p].Style.Border.Top.Color.SetColor(specColor);
            //    //currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength + p, currRow - 1, anchorCol + maxInfoLength + p].Style.Border.Left.Color.SetColor(specColor);
            //    //currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength + p, currRow - 1, anchorCol + maxInfoLength + p].Style.Border.Right.Color.SetColor(specColor);
            //}
            //****************************************************************************************************************************

            currRow = anchorRow + patterns.Length + 2 + allContents.Length; //最下面印X Axis

            var xAxisAry = ((string)dtRow["X Axis"]).Split('#');
            var shmooStepX = Convert.ToInt16((string)dtRow["Shmoo Step"]);
            for (var i = 0; i < xAxisAry.Length; i++)
            {
                var xAxis = xAxisAry[i].Split(';'); //0;Label Name 1;Label Type 2; ... Each Step

                currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 2].RichText
                    .Add(xAxis[0] + "\r\n" + @"(" + xAxis[1] + @")");
                currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 2].StyleName = "Shmoo Axis Label";
                currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 2, currRow + i, anchorCol + maxInfoLength - 1]
                    .Merge = true;
                currSheet.Row(currRow + i).Height = 36;

                for (var x = 0; x < shmooStepX; x++)
                {
                    currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 1 + x + 2 - 1].Value =
                        Convert.ToDouble(xAxis[x + 2]); //填X軸的值

                    // plot bin cut dot line in x Axis
                    if (binCutSpec.Contains(x))
                        currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 1 + x + 2 - 1].Style.Border.Left
                            .Style = ExcelBorderStyle.Dotted;

                    if (listSpecX.Count == 0)
                        continue;

                    if (x == listSpecX[0]) //著色
                    {
                        currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 1 + x + 2 - 1].Style.Fill.PatternType =
                            ExcelFillStyle.Solid;
                        currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 1 + x + 2 - 1].Style.Fill
                            .BackgroundColor.SetColor(Color.Blue);
                        currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 1 + x + 2 - 1].Style.Font.Color
                            .SetColor(Color.WhiteSmoke);
                    }
                    else if (listSpecX.Contains(x))
                    {
                        currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 1 + x + 2 - 1].Style.Fill.PatternType =
                            ExcelFillStyle.Solid;
                        currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 1 + x + 2 - 1].Style.Fill
                            .BackgroundColor.SetColor(Color.Yellow);
                    }
                }

                // plot bin cut spec
                if (isBinCutSpec)
                    for (var bvIndex = 0; bvIndex < binCutSpec.Count; bvIndex++)
                    {
                        var specPoint = binCutSpec[bvIndex];
                        currSheet.Cells[currRow + xAxisAry.Length - 1 + binCutInfoRow,
                                anchorCol + maxInfoLength - 1 + specPoint + 2 - 1 - 1].RichText
                            .Add(binCutVersion + "" + binCutPmode + "\r\n" + binCutSpecName[bvIndex]);
                        currSheet.Cells[currRow + xAxisAry.Length - 1 + binCutInfoRow,
                            anchorCol + maxInfoLength - 1 + specPoint + 2 - 1 - 1,
                            currRow + xAxisAry.Length - 1 + binCutInfoRow,
                            anchorCol + maxInfoLength - 1 + specPoint + 2 - 1].StyleName = "Shmoo BinCut Label";
                        currSheet.Cells[currRow + xAxisAry.Length - 1 + binCutInfoRow,
                            anchorCol + maxInfoLength - 1 + specPoint + 2 - 1 - 1,
                            currRow + xAxisAry.Length - 1 + binCutInfoRow,
                            anchorCol + maxInfoLength - 1 + specPoint + 2 - 1].Merge = true;
                        currSheet.Row(currRow + xAxisAry.Length - 1 + binCutInfoRow).Height = 45;
                    }


                currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 1 + 1, currRow + i,
                    anchorCol + maxInfoLength - 1 + xAxis.Length].Style.TextRotation = 90;
                currSheet.Cells[currRow + i, anchorCol + maxInfoLength - 1 + 1, currRow + i,
                    anchorCol + maxInfoLength - 1 + xAxis.Length].Style.Numberformat.Format = @"0.000";
            }

            currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + maxInfoLength,
                    currRow + xAxisAry.Length + binCutInfoRow, anchorCol + maxInfoLength + shmooStep - 1].Style
                .HorizontalAlignment = ExcelHorizontalAlignment.Center;

            //****************************************************************************************************************************
            //四周畫線跟定位
            currSheet.Cells[anchorRow, anchorCol,
                    anchorRow + patterns.Length + 2 + allContents.Length + xAxisAry.Length - 1 + binCutInfoRow,
                    anchorCol + maxInfoLength + shmooStep - 1].Style.Border
                .BorderAround(ExcelBorderStyle.Thin, Color.Black); //這一行沒辦法加到NamedStyle裡面...

            //標定Mix Shmoo
            if (allContents.Length > 1)
                currSheet.Cells[anchorRow + patterns.Length + 2, anchorCol + 1, anchorRow + patterns.Length + 2,
                        anchorCol + maxInfoLength + shmooStep - 1].Style.Border
                    .BorderAround(ExcelBorderStyle.Thick, Color.Gold);

            dicHyperPoint[testNumInstance][0] = anchorRow + patterns.Length + 2 + allContents.Length + xAxisAry.Length -
                1 + binCutInfoRow;
            dicHyperPoint[testNumInstance][1] = anchorCol;
            //****************************************************************************************************************************

            //currSheet.Cells[anchorRow + patterns.Length + 1, anchorCol + maxInfoLength, currRow, anchorCol + maxInfoLength + shmooStep].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); //這一行沒辦法加到NamedStyle裡面...
        }

        private static void PlotShmoo2D(DataRow dtRow, ColorSetting colorSetting, ref ExcelWorksheet currSheet,
            ref Dictionary<string, int[]> dicHyperPoint)
        {
            var rgexComma = new Regex(@",", RegexOptions.Compiled);
            var rgexSepOverLay = new Regex(@"\|", RegexOptions.Compiled); // over lay percentage separate symbol
            var rgexSepMergePF = new Regex(@"\^", RegexOptions.Compiled); // over lay percentage separate symbol


            //shmooStepX + yAxisAry.Length + 1(X Lable) 加起來就是Max Column
            var shmooStep = ((string)dtRow["Shmoo Step"]).Split(':');
            var shmooStepX = Convert.ToInt16(shmooStep[0]);
            var shmooStepY = Convert.ToInt16(shmooStep[1]);

            var xAxisAry = ((string)dtRow["X Axis"]).Split('#');
            var yAxisAry = ((string)dtRow["Y Axis"]).Split('#');


            var colJumpStep = shmooStepX + yAxisAry.Length + 1 + 1; //橫向到下一個Shmoo要位移的距離 1(X Label) + 1(Gap) 

            //*******************************************
            // BinCut Spec infor
            var isBinCutSpec = false;
            var binCutSpec = new List<int>();
            var binCutSpecName = new List<string>();
            var binCutPmode = "";
            var binCutFile = "";
            var binCutInfoRow = 0;
            var binCutVersion = "";
            if ((string)dtRow["BinCutSpec"] != "N/A")
            {
                isBinCutSpec = true;
                binCutSpec = Array.ConvertAll(((string)dtRow["BinCutSpec"]).Split(','), int.Parse).ToList();
                binCutSpecName = ((string)dtRow["BinCutSpecName"]).Split(',').ToList();
                binCutPmode = (string)dtRow["BinCutPmode"];
                binCutVersion = (string)dtRow["BinCutVersion"];
                binCutFile = (string)dtRow["BinCutPlan"];
                binCutInfoRow = 1;
            }
            //******************************************


            var lAllContents = ((string)dtRow["All Content"]).Split('#').ToList(); //Length代表橫向有幾個Shmoo要畫!!
            var lDieInfo = ((string)dtRow["Die Info"]).Split('#').ToList(); //Per Die


            var shmoo2DRowDatas = SplitShmooContentToMultRowData(lAllContents, lDieInfo, colJumpStep);


            var titles =
                ((string)dtRow["Titles"])
                .Split(':'); //"Test Num:" + shmooSetup.TestNum.ToString() + ":Test Instance:" + shmooSetup.TestInstanceName + ":Setup:" + shmooSetup.SetupName;
            var testNum = Convert.ToInt32(titles[1]);
            var testNumInstance = titles[1] + ":" + titles[3];
            var patterns = ((string)dtRow["Patterns"]).Split(',');
            var testInstances = ((string)dtRow["TestInstances"]).Split(',');

            var isMergeItem = titles[3] == "MergedItem";


            var onlyPrintOverlay = true;

            dicHyperPoint[testNumInstance] = new int[2];

            var hyperOn1stData = true;


            foreach (var shm2DRowData in shmoo2DRowDatas)
            {
                var dim = currSheet.Dimension;

                //起始定位點
                var anchorRow = dim == null ? 1 : dim.End.Row + 2;
                var anchorCol = dim == null ? 1 : dim.Start.Column;

                var currRow = anchorRow;
                var patternRow = anchorRow + 1; // 1 代表 Test Num & Pattern List Title

                var titalLength = isMergeItem ? testInstances.Length : patterns.Length;

                var titleRow = anchorRow + 1 + titalLength; // 1 代表 Test Num & Pattern List Title

                var dieRow = anchorRow + 1 + titalLength + 2 + binCutInfoRow; // 2代表 Test Instance & Setup Name

                if (hyperOn1stData)
                {
                    //****************************************************************************************************************************
                    //標定位置
                    dicHyperPoint[testNumInstance][0] =
                        anchorRow + titalLength + 4 + shmooStepY + xAxisAry.Length - 1 + binCutInfoRow;
                    dicHyperPoint[testNumInstance][1] = anchorCol;
                    hyperOn1stData = false;
                }


                //var allContents = shm2DRowData.AllContents;
                //var dieInfo = shm2DRowData.DieInfos;

                for (var i = 0; i < shm2DRowData.AllShmoo.Count; i++)
                {
                    if (shm2DRowData.AllShmoo[i].IsSkip)
                        continue;

                    var infos = shm2DRowData.AllShmoo[i].DieInfo.Split('|');
                    var perDieInfo = "";
                    var testInstanceName = titles[3];
                    if (infos.Length > 5)
                    {
                        perDieInfo = "Site: " + infos[1] + "    LotID: " + infos[2] + "    Die XY: " + infos[3] +
                                     "    Soft Bin: " + infos[4];
                        if (isMergeItem) testInstanceName = infos[5];
                    }

                    else //Overlayed Shmoo Info
                    {
                        perDieInfo = shm2DRowData.AllShmoo[i].DieInfo;
                    }

                    if (anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep > 16384 ||
                        anchorCol + 1 + i * colJumpStep > 16384)
                    {
                        MessageBox.Show("Data is over the limit, can't report all the shmoo!");
                        break;
                    }

                    //****************************************************************************************************************************
                    //列印Test Number
                    currSheet.Cells[currRow, anchorCol + i * colJumpStep].Value = "TestNum";
                    currSheet.Cells[currRow, anchorCol + 1 + i * colJumpStep].Value = testNum;
                    currSheet.Cells[currRow, anchorCol + i * colJumpStep, currRow, anchorCol + 1 + i * colJumpStep]
                        .StyleName = "Shmoo Axis Label";
                    currSheet.Cells[currRow, anchorCol + i * colJumpStep, currRow, anchorCol + 1 + i * colJumpStep]
                        .Style
                        .Font.Bold = true;


                    //****************************************************************************************************************************
                    //用群組印在Title之上
                    //印Merge Instance Line
                    if (isMergeItem)
                    {
                        if (onlyPrintOverlay)
                        {
                            currSheet.Cells[currRow, anchorCol + 2 + i * colJumpStep].Value = "Merge Instances";
                            currSheet.Cells[currRow, anchorCol + 2 + i * colJumpStep].StyleName = "Pattern List";
                            currSheet.Cells[
                                currRow, anchorCol + 2 + i * colJumpStep, currRow,
                                anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;
                            for (var p = 0; p < testInstances.Length; p++)
                            {
                                currSheet.Cells[patternRow + p, anchorCol + 2 + i * colJumpStep].Value =
                                    testInstances[p];
                                currSheet.Row(patternRow + p).OutlineLevel = 1;
                                currSheet.Cells[
                                    patternRow + p, anchorCol + 2 + i * colJumpStep, patternRow + p,
                                    anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;
                            }

                            currSheet.Cells[
                                    patternRow, anchorCol + 2, patternRow + testInstances.Length,
                                    anchorCol + shmooStepX + yAxisAry.Length].Style.HorizontalAlignment =
                                ExcelHorizontalAlignment.Left;
                            onlyPrintOverlay = false;
                        }
                        else
                        {
                            if (patterns.Count() > 1)
                            {
                                currSheet.Cells[currRow, anchorCol + 2 + i * colJumpStep].Value = "Pattern List";
                                currSheet.Cells[currRow, anchorCol + 2 + i * colJumpStep].StyleName = "Pattern List";
                                currSheet.Cells[
                                    currRow, anchorCol + 2 + i * colJumpStep, currRow,
                                    anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;


                                currSheet.Cells[patternRow, anchorCol + 2 + i * colJumpStep].Value = patterns[i];
                                currSheet.Row(patternRow).OutlineLevel = 1;
                                currSheet.Cells[
                                    patternRow, anchorCol + 2 + i * colJumpStep, patternRow,
                                    anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;

                                currSheet.Cells[
                                        patternRow, anchorCol + 2, patternRow + testInstances.Length,
                                        anchorCol + shmooStepX + yAxisAry.Length].Style.HorizontalAlignment =
                                    ExcelHorizontalAlignment.Left;
                            }
                        }

                        //****************************************************************************************************************************
                    }
                    else
                    {
                        //****************************************************************************************************************************
                        //用群組印在Title之上
                        //印Pattern List
                        currSheet.Cells[currRow, anchorCol + 2 + i * colJumpStep].Value = "Pattern List";
                        currSheet.Cells[currRow, anchorCol + 2 + i * colJumpStep].StyleName = "Pattern List";
                        currSheet.Cells[
                            currRow, anchorCol + 2 + i * colJumpStep, currRow,
                            anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;
                        for (var p = 0; p < patterns.Length; p++)
                        {
                            currSheet.Cells[patternRow + p, anchorCol + 2 + i * colJumpStep].Value = patterns[p];
                            currSheet.Row(patternRow + p).OutlineLevel = 1;
                            currSheet.Cells[
                                patternRow + p, anchorCol + 2 + i * colJumpStep, patternRow + p,
                                anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;
                        }

                        currSheet.Cells[
                                patternRow, anchorCol + 2, patternRow + patterns.Length,
                                anchorCol + shmooStepX + yAxisAry.Length].Style.HorizontalAlignment =
                            ExcelHorizontalAlignment.Left;
                        //****************************************************************************************************************************
                    }


                    //列印固定Title 
                    currSheet.Cells[titleRow, anchorCol + i * colJumpStep].Value = "Test Instance: " + testInstanceName;

                    currSheet.Cells[
                        titleRow, anchorCol + i * colJumpStep, titleRow,
                        anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;
                    currSheet.Cells[titleRow + 1, anchorCol + i * colJumpStep].Value = " Setup Name: " + titles[5];
                    currSheet.Cells[
                        titleRow + 1, anchorCol + i * colJumpStep, titleRow + 1,
                        anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;
                    if (isBinCutSpec)
                    {
                        // add bincut plan info title
                        currSheet.Cells[titleRow + 1 + binCutInfoRow, anchorCol + i * colJumpStep].Value =
                            " BinCut Plan: " + binCutFile + "  Performance Mode: " + binCutPmode;
                        currSheet.Cells[
                            titleRow + 1 + binCutInfoRow, anchorCol + i * colJumpStep, titleRow + 1 + binCutInfoRow,
                            anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;
                    }

                    currSheet.Cells[
                        titleRow, anchorCol + i * colJumpStep, titleRow + 1 + binCutInfoRow,
                        anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].StyleName = "2D Shmoo Title";
                    //****************************************************************************************************************************
                    //****************************************************************************************************************************
                    //印Per Die資訊 //shmooId.SourceFileName + "|" + shmooId.Site + "|" + shmooId.LotId + "|" + shmooId.DieXY + "|" + shmooId.Sort + "#";


                    currSheet.Cells[dieRow, anchorCol + 1 + i * colJumpStep].Value = perDieInfo;
                    currSheet.Cells[dieRow, anchorCol + 1 + i * colJumpStep].StyleName = "Shmoo Sub Title";
                    currSheet.Cells[
                        dieRow, anchorCol + 1 + i * colJumpStep, dieRow,
                        anchorCol + shmooStepX + yAxisAry.Length + i * colJumpStep].Merge = true;
                    //****************************************************************************************************************************
                    //****************************************************************************************************************************
                }


                //列印Shmoo Content
                for (var i = 0; i < shm2DRowData.AllShmoo.Count; i++) //Per Die
                {
                    if (shm2DRowData.AllShmoo[i].IsSkip)
                        continue;


                    currRow = anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow; // 1代表 間隔

                    var shmooContent = shm2DRowData.AllShmoo[i].Content;
                    var perY = rgexComma.Split(shmooContent); //分解成一層層Y

                    if (rgexSepOverLay.IsMatch(shmooContent)) //代表有疊圖的資訊 |
                    {
                        var collectPassRate = new List<double>();

                        for (var y = 0; y < perY.Length; y++)
                        {
                            var perStep = rgexSepOverLay.Split(perY[y]);
                            for (var x = 0; x < perStep.Length; x++)
                            {
                                currSheet.Cells[currRow, anchorCol + 1 + x + i * colJumpStep].Value =
                                    Convert.ToDouble(perStep[x]);
                                collectPassRate.Add(Convert.ToDouble(perStep[x]));
                            }

                            currRow++;
                        }

                        var ruleOverlay = currSheet.ConditionalFormatting.AddTwoColorScale(
                            new ExcelAddress(anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow,
                                anchorCol + 1 + i * colJumpStep,
                                anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow + perY.Length - 1,
                                anchorCol + 1 + i * colJumpStep + shmooStepX - 1)
                        );

                        ruleOverlay.HighValue.Color = Color.LimeGreen;
                        ruleOverlay.HighValue.Value = 100.0;

                        if (collectPassRate.Sum() < 0.0001) ruleOverlay.HighValue.Color = Color.Red;

                        ruleOverlay.LowValue.Color = Color.Red;
                        ruleOverlay.LowValue.Value = 0.0;

                        if (collectPassRate.Sum() > 100.0 * collectPassRate.Count - 0.001)
                            ruleOverlay.HighValue.Color = Color.LimeGreen;
                    }
                    else if (rgexSepMergePF.IsMatch(shmooContent))
                    {
                        for (var y = 0; y < perY.Length; y++)
                        {
                            var perStep = rgexSepMergePF.Split(perY[y]);
                            for (var x = 0; x < perStep.Length; x++)
                                currSheet.Cells[currRow, anchorCol + 1 + x + i * colJumpStep].Value = perStep[x];
                            currRow++;
                        }

                        var adr = new ExcelAddress(anchorRow + 1 + titalLength + 2 + 1, anchorCol + 1 + i * colJumpStep,
                            anchorRow + 1 + titalLength + 2 + 1 + perY.Length - 1,
                            anchorCol + 1 + shmooStepX + i * colJumpStep - 1);

                        var rulePassPass = currSheet.ConditionalFormatting.AddEqual(adr);
                        rulePassPass.Formula = "\"PP\"";
                        rulePassPass.Style.Fill.BackgroundColor.Color = colorSetting.PPColor;

                        var rulePassFail = currSheet.ConditionalFormatting.AddEqual(adr);
                        rulePassFail.Formula = "\"PF\"";
                        rulePassFail.Style.Fill.BackgroundColor.Color = colorSetting.PFColor;

                        var ruleFailPass = currSheet.ConditionalFormatting.AddEqual(adr);
                        ruleFailPass.Formula = "\"FP\"";
                        ruleFailPass.Style.Fill.BackgroundColor.Color = colorSetting.FPColor;

                        var ruleFailFail = currSheet.ConditionalFormatting.AddEqual(adr);
                        ruleFailFail.Formula = "\"FF\"";
                        ruleFailFail.Style.Fill.BackgroundColor.Color = colorSetting.FFColor;
                    }
                    else
                    {
                        for (var y = 0; y < perY.Length; y++)
                        {
                            var v = perY[y];
                            for (var x = 0; x < v.Length; x++)
                                currSheet.Cells[currRow, anchorCol + 1 + x + i * colJumpStep].Value = v[x].ToString();
                            currRow++;
                        }

                        var adr = new ExcelAddress(anchorRow + 1 + titalLength + 2 + 1, anchorCol + 1 + i * colJumpStep,
                            anchorRow + 1 + titalLength + 2 + 1 + perY.Length - 1,
                            anchorCol + 1 + shmooStepX + i * colJumpStep - 1);
                        var rulePass = currSheet.ConditionalFormatting.AddEqual(adr);
                        rulePass.Formula = "\"P\"";
                        rulePass.Style.Fill.BackgroundColor.Color = Color.LightGreen;
                        var ruleFail = currSheet.ConditionalFormatting.AddEqual(adr);
                        ruleFail.Formula = "\"F\"";
                        ruleFail.Style.Fill.BackgroundColor.Color = Color.Red;
                        var ruleAssumedPass = currSheet.ConditionalFormatting.AddEqual(adr);
                        ruleAssumedPass.Formula = "\"*\"";
                        ruleAssumedPass.Style.Fill.BackgroundColor.Color = Color.PaleGreen;
                        var ruleAssumedFail = currSheet.ConditionalFormatting.AddEqual(adr);
                        ruleAssumedFail.Formula = "\"~\"";
                        ruleAssumedFail.Style.Fill.BackgroundColor.Color = Color.Tomato;
                    }

                    //這一行沒辦法加到NamedStyle裡面...
                    currSheet.Cells[
                            anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow,
                            anchorCol + 1 + i * colJumpStep,
                            anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow + perY.Length - 1,
                            anchorCol + 1 + shmooStepX + i * colJumpStep - 1].Style.Border.Bottom.Style =
                        ExcelBorderStyle.Thin;
                    currSheet.Cells[
                            anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow,
                            anchorCol + 1 + i * colJumpStep,
                            anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow + perY.Length - 1,
                            anchorCol + 1 + shmooStepX + i * colJumpStep - 1].Style.Border.Top.Style =
                        ExcelBorderStyle.Thin;
                    currSheet.Cells[
                            anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow,
                            anchorCol + 1 + i * colJumpStep,
                            anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow + perY.Length - 1,
                            anchorCol + 1 + shmooStepX + i * colJumpStep - 1].Style.Border.Left.Style =
                        ExcelBorderStyle.Thin;
                    currSheet.Cells[
                            anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow,
                            anchorCol + 1 + i * colJumpStep,
                            anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow + perY.Length - 1,
                            anchorCol + 1 + shmooStepX + i * colJumpStep - 1].Style.Border.Right.Style =
                        ExcelBorderStyle.Thin;
                    //currSheet.Cells[anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow, anchorCol + 1 + i * colJumpStep, anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow + perY.Length - 1, anchorCol + 1 + shmooStepX + i * colJumpStep - 1].Style.Border.BorderAround( ExcelBorderStyle.Thin); 
                }

                //****************************************************************************************************************************
                //列印Y軸
                var listSpecY = new List<int>();
                if ((string)dtRow["Spec Y Point"] != "N/A")
                    listSpecY = Array.ConvertAll(((string)dtRow["Spec Y Point"]).Split(','), int.Parse).ToList();
                for (var i = 0; i < shm2DRowData.AllShmoo.Count; i++) //Per Die
                {
                    if (shm2DRowData.AllShmoo[i].IsSkip)
                        continue;


                    currRow = anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow; // 1代表 間隔

                    for (var yx = 0; yx < yAxisAry.Length; yx++) //Per Y Axis
                    {
                        var yAxis = yAxisAry[yx].Split(':'); //0:Label Name 1:Label Type 2: ... Each Step
                        currSheet.Cells[currRow + shmooStepY, anchorCol + 1 + shmooStepX + i * colJumpStep + yx]
                            .RichText
                            .Add("Y Label\r\n" + yAxis[0] + "\r\n" + @"(" + yAxis[1] + @")");
                        currSheet.Cells[currRow + shmooStepY, anchorCol + 1 + shmooStepX + i * colJumpStep + yx]
                                .StyleName
                            = "Shmoo Axis Label";
                        for (var y = 0; y < shmooStepY; y++)
                        {
                            currSheet.Cells[
                                        currRow + shmooStepY - y - 1, anchorCol + 1 + shmooStepX + i * colJumpStep + yx]
                                    .Value =
                                Convert.ToDouble(yAxis[y + 2]);

                            if (listSpecY.Count == 0)
                                continue;

                            if (y == listSpecY[0]) //著色
                            {
                                currSheet.Cells[
                                        currRow + shmooStepY - y - 1, anchorCol + 1 + shmooStepX + i * colJumpStep + yx]
                                    .Style
                                    .Fill.PatternType = ExcelFillStyle.Solid;
                                currSheet.Cells[
                                        currRow + shmooStepY - y - 1, anchorCol + 1 + shmooStepX + i * colJumpStep + yx]
                                    .Style
                                    .Fill.BackgroundColor.SetColor(Color.Blue);
                                currSheet.Cells[
                                        currRow + shmooStepY - y - 1, anchorCol + 1 + shmooStepX + i * colJumpStep + yx]
                                    .Style
                                    .Font.Color.SetColor(Color.WhiteSmoke);
                            }
                            else if (listSpecY.Contains(y))
                            {
                                currSheet.Cells[
                                        currRow + shmooStepY - y - 1, anchorCol + 1 + shmooStepX + i * colJumpStep + yx]
                                    .Style
                                    .Fill.PatternType = ExcelFillStyle.Solid;
                                currSheet.Cells[
                                        currRow + shmooStepY - y - 1, anchorCol + 1 + shmooStepX + i * colJumpStep + yx]
                                    .Style
                                    .Fill.BackgroundColor.SetColor(Color.Yellow);
                            }
                        }
                    }

                    currSheet.Cells[currRow, anchorCol + 1 + shmooStepX + i * colJumpStep,
                            currRow + shmooStepY - 1, anchorCol + 1 + shmooStepX + i * colJumpStep + yAxisAry.Length]
                        .Style
                        .Numberformat.Format = @"0.000";
                }

                //****************************************************************************************************************************
                //列印X軸
                var listSpecX = new List<int>();
                if ((string)dtRow["Spec X Point"] != "N/A")
                    listSpecX = Array.ConvertAll(((string)dtRow["Spec X Point"]).Split(','), int.Parse).ToList();
                for (var i = 0; i < shm2DRowData.AllShmoo.Count; i++) //Per Die
                {
                    if (shm2DRowData.AllShmoo[i].IsSkip)
                        continue;


                    ///                                      tital + die info 
                    currRow = anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow + shmooStepY; // 1代表 間隔

                    for (var yx = 0; yx < xAxisAry.Length; yx++) //Per X Axis
                    {
                        var xAxis = xAxisAry[yx].Split(';'); //0;Label Name 1;Label Type 2; ... Each Step
                        currSheet.Cells[currRow + yx, anchorCol + i * colJumpStep].RichText.Add("X Label\r\n" +
                            xAxis[0] +
                            "\r\n" + @"(" + xAxis[1] +
                            @")");
                        currSheet.Cells[currRow + yx, anchorCol + i * colJumpStep].StyleName = "Shmoo Axis Label";

                        for (var x = 0; x < shmooStepX; x++)
                        {
                            currSheet.Cells[currRow + yx, anchorCol + 1 + x + i * colJumpStep].Value =
                                Convert.ToDouble(xAxis[x + 2]);

                            // plot bin cut dot line in x Axis
                            if (binCutSpec.Contains(x))
                                currSheet.Cells[currRow + yx, anchorCol + 1 + x + i * colJumpStep].Style.Border.Left
                                        .Style
                                    = ExcelBorderStyle.Dotted;

                            if (listSpecX.Count == 0) continue;

                            if (x == listSpecX[0]) //著色
                            {
                                currSheet.Cells[currRow + yx, anchorCol + 1 + x + i * colJumpStep].Style.Fill
                                        .PatternType
                                    = ExcelFillStyle.Solid;
                                currSheet.Cells[currRow + yx, anchorCol + 1 + x + i * colJumpStep].Style.Fill
                                    .BackgroundColor.SetColor(Color.Blue);
                                currSheet.Cells[currRow + yx, anchorCol + 1 + x + i * colJumpStep].Style.Font.Color
                                    .SetColor(Color.WhiteSmoke);
                            }
                            else if (listSpecX.Contains(x))
                            {
                                currSheet.Cells[currRow + yx, anchorCol + 1 + x + i * colJumpStep].Style.Fill
                                        .PatternType
                                    = ExcelFillStyle.Solid;
                                currSheet.Cells[currRow + yx, anchorCol + 1 + x + i * colJumpStep].Style.Fill
                                    .BackgroundColor.SetColor(Color.Yellow);
                            }
                        }
                    }

                    // plot bin cut spec
                    if (isBinCutSpec)
                        for (var bvIndex = 0; bvIndex < binCutSpec.Count; bvIndex++)
                        {
                            var specPoint = binCutSpec[bvIndex];
                            currSheet.Cells[
                                currRow + xAxisAry.Length - 1 + binCutInfoRow,
                                anchorCol + 1 + specPoint - 1 + i * colJumpStep].RichText.Add(binCutVersion + " " +
                                binCutPmode + "\r\n" + binCutSpecName[bvIndex]);
                            currSheet.Cells[
                                currRow + xAxisAry.Length - 1 + binCutInfoRow,
                                anchorCol + 1 + specPoint + i * colJumpStep - 1,
                                currRow + xAxisAry.Length - 1 + binCutInfoRow,
                                anchorCol + 1 + specPoint + i * colJumpStep
                            ].StyleName = "Shmoo BinCut Label";
                            currSheet.Cells[
                                currRow + xAxisAry.Length - 1 + binCutInfoRow,
                                anchorCol + 1 + specPoint + i * colJumpStep - 1,
                                currRow + xAxisAry.Length - 1 + binCutInfoRow,
                                anchorCol + 1 + specPoint + i * colJumpStep
                            ].Merge = true;
                            currSheet.Row(currRow + xAxisAry.Length - 1 + binCutInfoRow).Height = 45;
                        }


                    currSheet.Cells[
                        currRow, anchorCol + 1 + i * colJumpStep, currRow + xAxisAry.Length - 1,
                        anchorCol + 1 + shmooStepX + i * colJumpStep - 1].Style.TextRotation = 90;
                    currSheet.Cells[
                        currRow, anchorCol + 1 + i * colJumpStep, currRow + xAxisAry.Length - 1,
                        anchorCol + 1 + shmooStepX + i * colJumpStep - 1].Style.Numberformat.Format = @"0.000";
                    currSheet.Cells[
                            currRow, anchorCol + 1 + i * colJumpStep, currRow + xAxisAry.Length - 1,
                            anchorCol + 1 + shmooStepX + i * colJumpStep].Style.VerticalAlignment =
                        ExcelVerticalAlignment.Center;
                }

                //****************************************************************************************************************************
                //把Spec的範圍標上去!!
                for (var i = 0; i < shm2DRowData.AllShmoo.Count; i++) //Per Die
                {
                    if (shm2DRowData.AllShmoo[i].IsSkip)
                        continue;

                    var shmooContent = shm2DRowData.AllShmoo[i].Content;
                    var allY = rgexComma.Split(shmooContent).Length - 1; //分解成一層層Y

                    var listSpecPointsX = new List<int>();
                    var listSpecPointsY = new List<int>();
                    if ((string)dtRow["Spec X Point"] != "N/A")
                        listSpecPointsX =
                            Array.ConvertAll(((string)dtRow["Spec X Point"]).Split(','), int.Parse).ToList();
                    //第一點一定是Default值, Spec 對稱
                    else
                        break;
                    if ((string)dtRow["Spec Y Point"] != "N/A")
                        listSpecPointsY =
                            Array.ConvertAll(((string)dtRow["Spec Y Point"]).Split(','), int.Parse).ToList();
                    //第一點一定是Default值, Spec 對稱
                    else
                        break;

                    //2D與1D不同 必須考量畫線的方向
                    var stColSpecX = anchorCol + 1 + i * colJumpStep;
                    var stRowSpecY = anchorRow + 1 + titalLength + 2 + 1 + binCutInfoRow;

                    //Default Spec Point!!
                    if (listSpecPointsX.Count > 0 && listSpecPointsY.Count > 0)
                    {
                        // This Line could be troublesome for some 2D shmooes
                        if (stRowSpecY + allY - listSpecPointsY[0] > 0)
                            currSheet.Cells[stRowSpecY + allY - listSpecPointsY[0],
                                    stColSpecX + listSpecPointsX[0],
                                    stRowSpecY + allY - listSpecPointsY[0],
                                    stColSpecX + listSpecPointsX[0]].Style.Border
                                .BorderAround(ExcelBorderStyle.Thick, Color.Blue);

                        listSpecPointsX.RemoveAt(0);
                        listSpecPointsY.RemoveAt(0);
                    }

                    var loopCount = listSpecPointsX.Count / 2;
                    for (var l = 0; l < loopCount; l++)
                    {
                        var minSpecX = listSpecPointsX.Min();
                        var maxSpecX = listSpecPointsX.Max();
                        var minSpecY = listSpecPointsY.Min();
                        var maxSpecY = listSpecPointsY.Max();

                        // This Line could be troublesome for some 2D shmooes
                        if (stRowSpecY + allY - maxSpecY > 0)
                            currSheet.Cells[
                                stRowSpecY + allY - maxSpecY, stColSpecX + minSpecX, stRowSpecY + allY - minSpecY,
                                stColSpecX + maxSpecX].Style.Border.BorderAround(ExcelBorderStyle.MediumDashDot,
                                Color.Yellow);
                        listSpecPointsX.Remove(minSpecX);
                        listSpecPointsX.Remove(maxSpecX);
                        listSpecPointsY.Remove(minSpecY);
                        listSpecPointsY.Remove(maxSpecY);
                    }
                }
                //****************************************************************************************************************************

                for (var i = 0; i < shm2DRowData.AllShmoo.Count; i++) //Per Die
                {
                    if (shm2DRowData.AllShmoo[i].IsSkip)
                        continue;

                    //Text Alignment置中
                    currSheet.Cells[anchorRow + 1 + titalLength + 2 + 1, anchorCol + i * colJumpStep,
                                anchorRow + 1 + titalLength + 2 + 1 + shmooStepY + xAxisAry.Length - 1 + binCutInfoRow,
                                anchorCol + 1 + shmooStepX + yAxisAry.Length + i * colJumpStep - 1].Style
                            .HorizontalAlignment =
                        ExcelHorizontalAlignment.Center;

                    //周圍畫線
                    currSheet.Cells[anchorRow, anchorCol + i * colJumpStep,
                        anchorRow + 1 + titalLength + 2 + 1 + shmooStepY + xAxisAry.Length - 1 + binCutInfoRow * 2,
                        anchorCol + 1 + shmooStepX + yAxisAry.Length + i * colJumpStep - 1].Style.Border.BorderAround(
                        ExcelBorderStyle.Medium, Color.Black);
                }
                ////****************************************************************************************************************************

                //for (var i = 0; i < shm2DRowData.AllShmoo.Count; i++) //Per Die
                //{
                //    currSheet.Cells[anchorRow, anchorCol + i * colJumpStep,
                //        anchorRow + 1 + titalLength + 2 + 1 + shmooStepY + xAxisAry.Length - 1 + binCutInfoRow * 2,
                //        anchorCol + 1 + shmooStepX + yAxisAry.Length + i * colJumpStep - 1].Style.Border.BorderAround(
                //            ExcelBorderStyle.Medium, Color.Black);
                //}
            }
        }

        private static List<Shmoo2DRowData> SplitShmooContentToMultRowData(List<string> allContents,
            List<string> dieInfo, int colJumpStep)
        {
            var shm2DRowDatas = new List<Shmoo2DRowData>();

            var maxDie = 16384 / colJumpStep - 2;
            var cnt = 1;
            var shmoo2DRowData = new Shmoo2DRowData();
            for (var i = 0; i < allContents.Count; i++)
            {
                shmoo2DRowData.AllShmoo.Add(new ShmooContent { DieInfo = dieInfo[i], Content = allContents[i] });
                cnt += 1;

                if (cnt > maxDie)
                {
                    shm2DRowDatas.Add(shmoo2DRowData);
                    shmoo2DRowData = new Shmoo2DRowData();
                    cnt = 1;
                }
            }

            shm2DRowDatas.Add(shmoo2DRowData);


            //Add skip content for 2nd Rows
            var isFirst = true;
            foreach (var shm2DRowData in shm2DRowDatas)
            {
                if (isFirst)
                {
                    isFirst = false;
                    continue;
                }

                shm2DRowData.AllShmoo.Insert(0, new ShmooContent { IsSkip = true, DieInfo = "NA", Content = "NA" });
            }

            return shm2DRowDatas;
        }

        #region 共用Private Method

        public static string GetValueFromMergedCell(ExcelRangeBase range) // 官方提供的範例 包含取Merged Cell的方法 參考用
        {
            var objAry = range.Value as object[,]; //是Object[,] 不是一個值 有可能是[null, null, null]

            if (objAry == null) return string.Empty;

            var value = objAry[0, 0] as string; //理論上必須成立 就算是單一Cell應該不能設成Merged?

            if (string.IsNullOrEmpty(value)) //如果遇到 [null, null, null]
                value = string.Empty;

            return value;
        }

        public static Dictionary<int, Dictionary<int, string>>
            BuildDictionaryOfMergedColRowValue(ExcelWorksheet ws) // 原始Check Merged Cell的方法是線性, 處理大量Data建議先建Dictionary
        {
            var dicAddrValue = new Dictionary<int, Dictionary<int, string>>(); //Row -> Col -> Value

            foreach (var mergedCells in ws.MergedCells) //
            {
                var mergedAddress = new ExcelAddress(mergedCells);

                for (var row = mergedAddress.Start.Row; row <= mergedAddress.End.Row; row++)
                {
                    if (!dicAddrValue.ContainsKey(row)) dicAddrValue[row] = new Dictionary<int, string>();

                    for (var col = mergedAddress.Start.Column; col <= mergedAddress.End.Column; col++)
                        dicAddrValue[row][col] = ws.Cells[mergedCells].RichText.Text;
                }
            }

            return dicAddrValue;
        }

        public static string GetCellValueFromWorkSheetWithMergedCells(ExcelWorksheet ws, int row, int col,
            Dictionary<int, Dictionary<int, string>> dicAddrValue) //利用Dictionary來查詢Merged Cell
        {
            string value;

            if (ws.Cells[row, col].Merge)
            {
                if (dicAddrValue.ContainsKey(row) && dicAddrValue[row].ContainsKey(col))
                    value = dicAddrValue[row][col];
                else
                    value = string.Empty;
            }
            else
            {
                value = ws.Cells[row, col].Text.Trim();

                if (string.IsNullOrEmpty(value)) //如果是Null也要轉換成String.Empty
                    value = string.Empty;
            }

            return value;
        }

        private static void
            CreateDefaultNamedStyleInWorkBook(ref ExcelPackage ep, string epType) //預先對指定的Excel全Sheet建立Style
        {
            //Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black); //這一行沒辦法加到NamedStyle裡面...?
            var namedStyleTitleRow = ep.Workbook.Styles.CreateNamedStyle("Title Row");
            namedStyleTitleRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            namedStyleTitleRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            namedStyleTitleRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
            namedStyleTitleRow.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(0, 176, 240));
            namedStyleTitleRow.Style.Font.Color.SetColor(Color.White);
            //namedStyleTitleRow.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

            if (epType == "Shmoo")
            {
                var namedStyleInfoTitleShm = ep.Workbook.Styles.CreateNamedStyle("Shmoo Info Header");
                namedStyleInfoTitleShm.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleInfoTitleShm.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleInfoTitleShm.Style.Fill.BackgroundColor.SetColor(Color.Orange);
                namedStyleInfoTitleShm.Style.Font.Color.SetColor(Color.White);

                var namedStylePatListShm = ep.Workbook.Styles.CreateNamedStyle("Pattern List");
                namedStylePatListShm.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                namedStylePatListShm.Style.Font.Color.SetColor(Color.MediumPurple);

                var namedStyleShmTitleShm1D = ep.Workbook.Styles.CreateNamedStyle("1D Shmoo Title");
                namedStyleShmTitleShm1D.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                namedStyleShmTitleShm1D.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleShmTitleShm1D.Style.Fill.BackgroundColor.SetColor(Color.MediumSlateBlue);
                namedStyleShmTitleShm1D.Style.Font.Color.SetColor(Color.White);

                var namedStyleShmTitleShm2D = ep.Workbook.Styles.CreateNamedStyle("2D Shmoo Title");
                namedStyleShmTitleShm2D.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleShmTitleShm2D.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleShmTitleShm2D.Style.Fill.BackgroundColor.SetColor(Color.MediumSlateBlue);
                namedStyleShmTitleShm2D.Style.Font.Color.SetColor(Color.White);


                var namedStyleShmTitleEyeDiagram = ep.Workbook.Styles.CreateNamedStyle("EyeDiagram");
                namedStyleShmTitleEyeDiagram.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleShmTitleEyeDiagram.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleShmTitleEyeDiagram.Style.Fill.BackgroundColor.SetColor(Color.DimGray);
                namedStyleShmTitleEyeDiagram.Style.Font.Color.SetColor(Color.White);


                var namedStyleShmSubTitleShm = ep.Workbook.Styles.CreateNamedStyle("Shmoo Sub Title");
                namedStyleShmSubTitleShm.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleShmSubTitleShm.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleShmSubTitleShm.Style.Fill.BackgroundColor.SetColor(Color.LightSeaGreen);
                namedStyleShmSubTitleShm.Style.Font.Color.SetColor(Color.White);

                var namedStyleAxisShm1D = ep.Workbook.Styles.CreateNamedStyle("Shmoo Axis Label");
                namedStyleAxisShm1D.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleAxisShm1D.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleAxisShm1D.Style.WrapText = true;
                namedStyleAxisShm1D.Style.Fill.BackgroundColor.SetColor(Color.LemonChiffon);

                var namedStyleShmooBinCut = ep.Workbook.Styles.CreateNamedStyle("Shmoo BinCut Label");
                namedStyleShmooBinCut.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleShmooBinCut.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleShmooBinCut.Style.WrapText = true;
                namedStyleShmooBinCut.Style.Fill.BackgroundColor.SetColor(Color.LightBlue);
            }

            if (epType == "TestFlowProfile")
            {
                var namedStyleSubTitleRow = ep.Workbook.Styles.CreateNamedStyle("Sub Title Row");
                namedStyleSubTitleRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleSubTitleRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleSubTitleRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleSubTitleRow.Style.Fill.BackgroundColor.SetColor(Color.DeepSkyBlue);
                namedStyleSubTitleRow.Style.Font.Color.SetColor(Color.White);
                //namedStyleTitleRow.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);

                var namedStyleOddRow = ep.Workbook.Styles.CreateNamedStyle("Odd Row");
                namedStyleOddRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleOddRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleOddRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleOddRow.Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);

                var namedStyleEvenRow = ep.Workbook.Styles.CreateNamedStyle("Even Row");
                namedStyleEvenRow.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleEvenRow.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleEvenRow.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleEvenRow.Style.Fill.BackgroundColor.SetColor(Color.FloralWhite);

                var namedStyleTestSettingHeader = ep.Workbook.Styles.CreateNamedStyle("Test Setting Header");
                namedStyleTestSettingHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleTestSettingHeader.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleTestSettingHeader.Style.TextRotation = 180;

                var namedStyleHighlightCell = ep.Workbook.Styles.CreateNamedStyle("Highlight");
                namedStyleHighlightCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                namedStyleHighlightCell.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                namedStyleHighlightCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                namedStyleHighlightCell.Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
            }
        }

        public static bool CheckIfExcelIsOpened(string filePath)
        {
            if (!File.Exists(filePath)) return false;

            try
            {
                Stream s = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                s.Close();
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }

        #endregion
    }
}