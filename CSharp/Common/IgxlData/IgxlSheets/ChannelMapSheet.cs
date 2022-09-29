using IgxlData.IgxlBase;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    [Serializable]
    public class ChannelMapSheet : IgxlSheet
    {
        #region Field

        private const string SheetType = "DTChanMap";
        private List<ChannelMapRow> _channelData;
        private bool _isPogo;

        public const int DeviceUnderTestPinName = 2;
        public const int DeviceUnderTestPackagePin = 3;
        public const int GetPinType = 4;

        #endregion

        #region Property

        public List<ChannelMapRow> ChannelMapRows
        {
            get { return _channelData; }
            set { _channelData = value; }
        }

        public bool IsPogo
        {
            get { return _isPogo; }
            set { _isPogo = value; }
        }

        public int SiteNum { get; set; }

        #endregion

        #region Constructor

        public ChannelMapSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _channelData = new List<ChannelMapRow>();
            SiteNum = 1;
            IgxlSheetName = IgxlSheetNameList.ChannelMap;
        }

        public ChannelMapSheet(string sheetName)
            : base(sheetName)
        {
            _channelData = new List<ChannelMapRow>();
            SiteNum = 1;
            IgxlSheetName = IgxlSheetNameList.ChannelMap;
        }

        #endregion

        #region Member Function

        protected void WriteHeader()
        {
            IgxlWriter.WriteLine(
                "DTChanMap,version=2.6:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1:dataformat=signal\tChannel Map");
            IgxlWriter.WriteLine("");
        }

        protected void WriteColumnsHeader()
        {
            var viewMode = "";
            _isPogo = _channelData.Find(p => p.Sites.Exists(a => Regex.IsMatch(a, "ch", RegexOptions.IgnoreCase))) ==
                      null;

            IgxlWriter.WriteLine("\tDIB ID:\t\t\tView Mode:\t" + viewMode);
            IgxlWriter.WriteLine("\tUSL Tag:");
            IgxlWriter.WriteLine("\tDevice Under Test\t\tTester Channel");
            var siteNum = _channelData.Count != 0 ? _channelData[0].Sites.Count : 2;

            IgxlWriter.Write("\tPin Name\tPackage Pin\tType\t");
            for (var i = 0; i < siteNum; i++)
                IgxlWriter.Write("Site " + i + "\t");
            IgxlWriter.WriteLine("Comment");
        }

        protected void WriteRows()
        {
            foreach (var chanel in _channelData)
            {
                IgxlWriter.Write("\t" + chanel.DeviceUnderTestPinName + "\t" + chanel.DeviceUnderTestPackagePin + "\t" +
                                 chanel.Type + "\t");
                foreach (var site in chanel.Sites) IgxlWriter.Write(site + "\t");
                IgxlWriter.Write(chanel.Comment);
                IgxlWriter.WriteLine();
            }
        }

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
                version = "2.4";
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version == "2.6")
                {
                    Write2P6(fileName);
                }
                else if (version == "2.4")
                {
                    var igxlSheetsVersion = dic["2.4"];
                    Write2P4(fileName, version, igxlSheetsVersion);
                }
                else
                {
                    throw new Exception(string.Format("The ChannelMap sheet version:{0} is not supported!", version));
                }
            }
        }

        private void Write2P4(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_channelData.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var pinNameIndex = GetIndexFrom(igxlSheetsVersion, "Device Under Test", "Pin Name");
                var packagePinIndex = GetIndexFrom(igxlSheetsVersion, "Device Under Test", "Package Pin");
                var typeIndex = GetIndexFrom(igxlSheetsVersion, "Tester Channel", "Type");
                var siteIndex = GetIndexFrom(igxlSheetsVersion, "Site");
                var relativeColumnIndex = siteIndex + _channelData.First().Sites.Count;
                maxCount = Math.Max(maxCount, relativeColumnIndex + 1);

                #region headers

                var startRow = igxlSheetsVersion.Columns.RowCount;
                if (version == "2.4")
                    sw.WriteLine(SheetType + ",version=" + version +
                                 ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
                else
                    sw.WriteLine(SheetType + ",version=" + version +
                                 ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1:dataformat=signal\t" +
                                 IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    SetColumns(igxlSheetsVersion, i, arr);

                    #region Set Variant

                    if (igxlSheetsVersion.Columns.Variant != null)
                        foreach (var item in igxlSheetsVersion.Columns.Variant)
                        {
                            if (item.rowIndex == i)
                                arr[item.indexFrom] = item.columnName;

                            if (item.columnName == "Site" && item.rowIndex == i)
                                for (var index = 0; index < _channelData.First().Sites.Count; index++)
                                    arr[item.indexFrom + index] = "Site " + index;
                        }

                    #endregion

                    SetRelativeColumn(igxlSheetsVersion, i, arr, relativeColumnIndex);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data

                for (var j = 0; j < _channelData.Count; j++)
                {
                    var row = _channelData[j];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.DeviceUnderTestPinName))
                    {
                        arr[0] = row.ColumnA;
                        arr[pinNameIndex] = row.DeviceUnderTestPinName;
                        arr[packagePinIndex] = row.DeviceUnderTestPackagePin;
                        arr[typeIndex] = row.Type;
                        for (var i = 0; i < row.Sites.Count; i++)
                        {
                            var site = row.Sites[i];
                            arr[siteIndex + i] = site;
                        }

                        arr[relativeColumnIndex] = row.Comment;
                    }
                    else
                    {
                        arr = new[] { "\t" };
                    }

                    sw.WriteLine(string.Join("\t", arr));
                }

                #endregion
            }
        }

        private void Write2P6(string fileName)
        {
            GetStreamWriter(fileName);
            WriteHeader();
            WriteColumnsHeader();
            WriteRows();
            CloseStreamWriter();
        }

        public void AddRow(ChannelMapRow channelRow)
        {
            _channelData.Add(channelRow);
        }

        #endregion
    }
}