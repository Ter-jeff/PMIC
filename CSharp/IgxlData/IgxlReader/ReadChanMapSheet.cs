using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlReader
{
    public class ReadChanMapSheet : IgxlSheetReader
    {
        #region public Function
        public ChannelMapSheet ReadSheet(string fileName)
        {
            return ReadSheet(ConvertTxtToExcelSheet(fileName));
        }

        public ChannelMapSheet ReadSheet(ExcelWorksheet wSheet)
        {
            int endRow = wSheet.Dimension.End.Row;
            var channelMapSheet = new ChannelMapSheet(wSheet);
            List<int> siteIndex = GetSiteColumnIndex(wSheet);
            channelMapSheet.SiteNum = siteIndex.Count;

            channelMapSheet.IsPogo = GetViewMode(wSheet);
            const string regWalkRoundChannel = @"(?<mainCh>\d+\.\d+)e\+(?<SubChan>\d+)";
            for (int i = 7; i <= endRow; i++)
            {
                var channelMapRow = new ChannelMapRow();
                channelMapRow.RowNum = i;
                channelMapRow.DeviceUnderTestPinName = EpplusOperation.GetCellValue(wSheet, i, ChannelMapSheet.DeviceUnderTestPinName);
                channelMapRow.DeviceUnderTestPackagePin = EpplusOperation.GetCellValue(wSheet, i, ChannelMapSheet.DeviceUnderTestPackagePin);
                channelMapRow.Type = GetPinType(EpplusOperation.GetCellValue(wSheet, i, ChannelMapSheet.GetPinType));

                foreach (int columnIndex in siteIndex)
                {
                    string siteValue = EpplusOperation.GetCellValue(wSheet, i, columnIndex);
                    if (Regex.IsMatch(siteValue, regWalkRoundChannel, RegexOptions.IgnoreCase))
                    {
                        var mainChan = Regex.Match(siteValue, regWalkRoundChannel, RegexOptions.IgnoreCase).Groups["mainCh"].Value;
                        var subChan = Regex.Match(siteValue, regWalkRoundChannel, RegexOptions.IgnoreCase).Groups["SubChan"].Value;
                        siteValue = string.Format("{0}.e{1}", mainChan.Replace(".", ""), int.Parse(subChan) - 1);
                    }
                    channelMapRow.Sites.Add(siteValue);
                }

                if (string.IsNullOrEmpty(channelMapRow.DeviceUnderTestPinName) && string.IsNullOrEmpty(channelMapRow.Type))
                    break;

                channelMapSheet.AddRow(channelMapRow);
            }

            //foreach (var row in channelMapSheet.ChannelMapRows)
            //{
            //    // S:VDD_ANA_S_UVI80 => equals to VDD_ANA_S_UVI80
            //    List<string> sties = new List<string>();
            //    int cnt = 0;
            //    foreach (var site in row.Sites)
            //    {
            //        sties.Add(Regex.IsMatch(site, @"^S\:", RegexOptions.IgnoreCase) ?
            //            GetAliasSiteInfo(channelMapSheet, row.Sites, cnt) : site);
            //        cnt++;
            //    }
            //    if (testerCfg != null)
            //        row.InstrumentType = testerCfg.GetToolTypeByChannelAssignment(sties);
            //}
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
                if (mapRow.DeviceUnderTestPinName == site)
                {
                    return mapRow.Sites[siteIndex];
                }
            }
            return "";
        }
    }
}
