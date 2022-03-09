using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.Others;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadChanMapSheet : IgxlSheetReader
    {
        #region public Function
        public ChannelMapSheet GetSheet(string fileName, TesterConfigManager testerConfigManager = null)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName), testerConfigManager);
        }

        public ChannelMapSheet GetSheet(Worksheet worksheet, TesterConfigManager testerConfigManager = null)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet), testerConfigManager);
        }

        public ChannelMapSheet GetSheet(ExcelWorksheet wSheet, TesterConfigManager testerConfigManager = null)
        {
            int endRow = wSheet.Dimension.End.Row;
            var channelMapSheet = new ChannelMapSheet(wSheet);
            List<int> siteIndex = GetSiteColumnIndex(wSheet);
            channelMapSheet.SiteNum = siteIndex.Count;

            channelMapSheet.IsPogo = GetViewMode(wSheet);
            const string regwalkroundChannel = @"(?<mainCh>\d+\.\d+)e\+(?<SubChan>\d+)";
            for (int i = 7; i <= endRow; i++)
            {
                var channelMapRow = new ChannelMapRow();
                channelMapRow.DiviceUnderTestPinName = EpplusOperation.GetCellValue(wSheet, i, 2);
                channelMapRow.DiviceUnderTestPackagePin = EpplusOperation.GetCellValue(wSheet, i, 3);
                channelMapRow.Type = GetPinType(EpplusOperation.GetCellValue(wSheet, i, 4));

                foreach (int columnIndex in siteIndex)
                {
                    string siteValue = EpplusOperation.GetCellValue(wSheet, i, columnIndex);
                    if (Regex.IsMatch(siteValue, regwalkroundChannel, RegexOptions.IgnoreCase))
                    {
                        var mainChan = Regex.Match(siteValue, regwalkroundChannel, RegexOptions.IgnoreCase).Groups["mainCh"].Value;
                        var SubChan = Regex.Match(siteValue, regwalkroundChannel, RegexOptions.IgnoreCase).Groups["SubChan"].Value;
                        double value;
                        double.TryParse(mainChan, out value);
                        siteValue = string.Format("{0}.e{1}", value.ToString().Replace(".", ""), int.Parse(SubChan) - 1);
                    }
                    channelMapRow.Sites.Add(siteValue);
                }

                if (string.IsNullOrEmpty(channelMapRow.DiviceUnderTestPinName) && string.IsNullOrEmpty(channelMapRow.Type))
                    break;

                channelMapSheet.AddRow(channelMapRow);
            }

            foreach (var row in channelMapSheet.ChannelMapRows)
            {
                // S:VDD_ANA_S_UVI80 => equals to VDD_ANA_S_UVI80
                List<string> sties = new List<string>();
                int cnt = 0;
                foreach (var site in row.Sites)
                {
                    sties.Add(Regex.IsMatch(site, @"^S\:", RegexOptions.IgnoreCase) ?
                        GetAliasSiteInfo(channelMapSheet, row.Sites, cnt) : site);
                    cnt++;
                }
                if (testerConfigManager != null)
                    row.InstrumentType = testerConfigManager.GetToolTypeByChannelAssignment(sties, wSheet.Name);
            }
            return channelMapSheet;
        }
        #endregion

        private List<int> GetSiteColumnIndex(ExcelWorksheet wSheet)
        {
            List<int> siteColumnIndex = new List<int>();
            int endColumn = wSheet.Dimension.End.Column;
            for (int col = 1; col <= endColumn; col++)
            {
                if (Regex.IsMatch(EpplusOperation.GetCellValue(wSheet, 6, col).ToLower(), @"^site\s*\d+"))
                    siteColumnIndex.Add(col);
            }
            return siteColumnIndex;
        }

        private bool GetViewMode(ExcelWorksheet wSheet)
        {
            for (int i = 1; i <= 7; i++)
            {
                for (int j = 1; j <= 10; j++)
                {
                    string text = EpplusOperation.GetCellValue(wSheet, i, j);
                    if (text.Equals("View Mode:", StringComparison.OrdinalIgnoreCase))
                    {
                        string mode = EpplusOperation.GetCellValue(wSheet, i, j + 1);
                        if (mode.Equals("Pogo", StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }
            }
            return false;
        }

        private string GetPinType(string type)
        {
            if (type.Equals("IO", StringComparison.OrdinalIgnoreCase))
                return "I/O";
            return type;
        }

        private string GetAliasSiteInfo(ChannelMapSheet channelMapSheet, List<string> rowSiteData, int siteIndex)
        {
            string site = rowSiteData[siteIndex].Split(':')[1];
            for (int i = 0; i < channelMapSheet.ChannelMapRows.Count; ++i)
            {
                ChannelMapRow mapRow = channelMapSheet.ChannelMapRows[i];
                if (mapRow.DiviceUnderTestPinName == site)
                {
                    return mapRow.Sites[siteIndex];
                }
            }
            return "";
        }
    }
}
