using CommonLib.Extension;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlReader
{
    public class ReadChanMapSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 6;

        private List<int> GetSiteColumnIndex(ExcelWorksheet excelWorksheet)
        {
            var siteColumnIndex = new List<int>();
            var endColumn = excelWorksheet.Dimension.End.Column;
            for (var col = 1; col <= endColumn; col++)
                if (Regex.IsMatch(excelWorksheet.GetCellValue(6, col).ToLower(), @"^site\s*\d+"))
                    siteColumnIndex.Add(col);
            return siteColumnIndex;
        }

        private bool GetViewMode(ExcelWorksheet excelWorksheet)
        {
            for (var i = 1; i <= 7; i++)
                for (var j = 1; j <= 10; j++)
                {
                    var text = excelWorksheet.GetCellValue(i, j);
                    if (text.Equals("View Mode:", StringComparison.OrdinalIgnoreCase))
                    {
                        var mode = excelWorksheet.GetCellValue(i, j + 1);
                        if (mode.Equals("Pogo", StringComparison.OrdinalIgnoreCase))
                            return true;
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
            var site = rowSiteData[siteIndex].Split(':')[1];
            for (var i = 0; i < channelMapSheet.ChannelMapRows.Count; ++i)
            {
                var mapRow = channelMapSheet.ChannelMapRows[i];
                if (mapRow.DeviceUnderTestPinName == site) return mapRow.Sites[siteIndex];
            }

            return "";
        }

        public ChannelMapSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public ChannelMapSheet GetSheet(ExcelWorksheet excelWorksheet)
        {
            var endRow = excelWorksheet.Dimension.End.Row;
            var channelMapSheet = new ChannelMapSheet(excelWorksheet);
            var siteIndex = GetSiteColumnIndex(excelWorksheet);
            channelMapSheet.SiteNum = siteIndex.Count;

            channelMapSheet.IsPogo = GetViewMode(excelWorksheet);
            const string regWalkRoundChannel = @"(?<mainCh>\d+\.\d+)e\+(?<SubChan>\d+)";
            for (var i = 7; i <= endRow; i++)
            {
                var channelMapRow = new ChannelMapRow();
                channelMapRow.RowNum = i;
                channelMapRow.DeviceUnderTestPinName =
                    excelWorksheet.GetCellValue(i, ChannelMapSheet.DeviceUnderTestPinName);
                channelMapRow.DeviceUnderTestPackagePin =
                    excelWorksheet.GetCellValue(i, ChannelMapSheet.DeviceUnderTestPackagePin);
                channelMapRow.Type = GetPinType(excelWorksheet.GetCellValue(i, ChannelMapSheet.GetPinType));

                foreach (var columnIndex in siteIndex)
                {
                    var siteValue = excelWorksheet.GetCellValue(i, columnIndex);
                    if (Regex.IsMatch(siteValue, regWalkRoundChannel, RegexOptions.IgnoreCase))
                    {
                        var mainChan = Regex.Match(siteValue, regWalkRoundChannel, RegexOptions.IgnoreCase)
                            .Groups["mainCh"].Value;
                        var subChan = Regex.Match(siteValue, regWalkRoundChannel, RegexOptions.IgnoreCase)
                            .Groups["SubChan"].Value;
                        siteValue = string.Format("{0}.e{1}", mainChan.Replace(".", ""), int.Parse(subChan) - 1);
                    }

                    channelMapRow.Sites.Add(siteValue);
                }

                if (string.IsNullOrEmpty(channelMapRow.DeviceUnderTestPinName) &&
                    string.IsNullOrEmpty(channelMapRow.Type))
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

        public List<string> GetSites(Stream stream, string sheetName)
        {
            var sites = new List<string>();
            var i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    var arr = line.Split('\t').ToList();
                    if (i == StartRowIndex)
                    {
                        for (int j = 0; j < arr.Count; j++)
                        {
                            if (Regex.IsMatch(arr.ElementAt(j), @"^site\s*\d+", RegexOptions.IgnoreCase))
                                sites.Add(arr.ElementAt(j));
                        }
                    }
                    i++;
                }
            }
            return sites;
        }
    }
}