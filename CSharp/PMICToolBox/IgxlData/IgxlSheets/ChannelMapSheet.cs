using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlSheets
{
    [Serializable]
    public class ChannelMapSheet : IgxlSheet
    {
        private const string SheetType = "DTChanMap";

        #region Field
        private List<ChannelMapRow> _channelMapRows;
        private bool _isPogo;
        #endregion

        #region Property

        public List<ChannelMapRow> ChannelMapRows
        {
            get { return _channelMapRows; }
            set { _channelMapRows = value; }
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
            _channelMapRows = new List<ChannelMapRow>();
            SiteNum = 1;
            IgxlSheetName = IgxlSheetNameList.ChannelMap;
        }

        public ChannelMapSheet(string sheetName)
            : base(sheetName)
        {
            _channelMapRows = new List<ChannelMapRow>();
            SiteNum = 1;
            IgxlSheetName = IgxlSheetNameList.ChannelMap;
        }

        #endregion

        #region Member Function
        //protected override void WriteHeader()
        //{
        //    IgxlWriter.WriteLine(SheetType + ",version=2.6:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1:dataformat=signal\t" + _igxlSheetName);
        //    IgxlWriter.WriteLine("");
        //}

        //protected override void WriteColumnsHeader()
        //{
        //    const string viewMode = "";
        //    if (_channelMapRows.Find(p => p.Sites.Exists(a => Regex.IsMatch(a, "ch", RegexOptions.IgnoreCase))) != null)
        //        _isPogo = false;
        //    else
        //        _isPogo = true;

        //    IgxlWriter.WriteLine("\tDIB ID:\t\t\tView Mode:\t" + viewMode);
        //    IgxlWriter.WriteLine("\tUSL Tag:");
        //    IgxlWriter.WriteLine("\tDevice Under Test\t\tTester Channel");
        //    int siteNum;
        //    if (_channelMapRows.Count != 0)
        //        siteNum = _channelMapRows[0].Sites.Count;
        //    else
        //        siteNum = 2;

        //    IgxlWriter.Write("\tPin Name\tPackage Pin\tType\t");
        //    for (int i = 0; i < siteNum; i++)
        //        IgxlWriter.Write("Site " + i + "\t");
        //    IgxlWriter.WriteLine("Comment");
        //}

        //protected override void WriteRows()
        //{
        //    foreach (var chanel in _channelMapRows)
        //    {
        //        IgxlWriter.Write("\t" + chanel.DiviceUnderTestPinName + "\t" + chanel.DiviceUnderTestPackagePin + "\t" + chanel.Type + "\t");
        //        foreach (var site in chanel.Sites)
        //            IgxlWriter.Write(site + "\t");
        //        IgxlWriter.Write(chanel.Comment);
        //        IgxlWriter.WriteLine();
        //    }
        //}

        public override void Write(string fileName, string version = "2.4")
        {
            //if (version == "2.6")
            //    Write2p6(fileName, version);
            //else
            //    Write2p4(fileName, version);

            //Support 2.4 & 2.6
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                Directory.CreateDirectory(Path.GetDirectoryName(fileName));
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (dic.ContainsKey(version))
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (dic.ContainsKey("2.4"))
                {
                    var igxlSheetsVersion = dic["2.4"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_channelMapRows.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                int pinNameIndex = GetIndexFrom(igxlSheetsVersion,"Device Under Test", "Pin Name");
                int packagePinIndex = GetIndexFrom(igxlSheetsVersion,"Device Under Test", "Package Pin");
                int typeIndex = GetIndexFrom(igxlSheetsVersion, "Tester Channel","Type");
                int siteIndex = GetIndexFrom(igxlSheetsVersion, "Site");
                int relativeColumnIndex = siteIndex + _channelMapRows.First().Sites.Count;
                maxCount = Math.Max(maxCount, relativeColumnIndex + 1);

                #region headers
                var startRow = igxlSheetsVersion.Columns.RowCount;
                if (version == "2.4")
                    sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
                else
                    sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1:dataformat=signal\t" + IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    SetColumns(igxlSheetsVersion, i, arr);

                    #region Set Variant
                    if (igxlSheetsVersion.Columns.Variant != null)
                    {
                        foreach (var item in igxlSheetsVersion.Columns.Variant)
                        {
                            if (item.rowIndex == i)
                                arr[item.indexFrom] = item.columnName;

                            if (item.columnName == "Site" && item.rowIndex == i)
                            {
                                for (int index = 0; index < _channelMapRows.First().Sites.Count; index++)
                                    arr[item.indexFrom + index] = "Site " + index;
                            }
                        }
                    }
                    #endregion

                    SetRelativeColumn(igxlSheetsVersion, i, arr, relativeColumnIndex);

                    WriteHeader(arr, sw);
                }
                #endregion

                #region data
                for (var j = 0; j < _channelMapRows.Count; j++)
                {
                    var row = _channelMapRows[j];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.DiviceUnderTestPinName))
                    {
                        arr[0] = row.ColumnA;
                        arr[pinNameIndex] = row.DiviceUnderTestPinName;
                        arr[packagePinIndex] = row.DiviceUnderTestPackagePin;
                        arr[typeIndex] = row.Type;
                        for (int i = 0; i < row.Sites.Count; i++)
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


        //private void Write2p4(string fileName, string version)
        //{
        //    double versionDouble = Double.Parse(version);
        //    Action<string> validate = new Action<string>((a) => { });
        //    if (SiteNum == 0) return;
        //    GenChanMap genChanMap = new GenChanMap(fileName, validate, true, SiteNum, versionDouble);
        //    foreach (ChannelMapRow channel in _channelMapRows)
        //    {
        //        genChanMap.AddPin(channel.DiviceUnderTestPinName, channel.DiviceUnderTestPackagePin, channel.Type, channel.Sites.ToArray());
        //    }
        //    genChanMap.WriteSheet();
        //}

        //private void Write2p6(string fileName, string version)
        //{
        //    GetSreamWriter(fileName);
        //    WriteHeader();
        //    WriteColumnsHeader();
        //    WriteRows();
        //    CloseStreamWriter();
        //}

        public void AddRow(ChannelMapRow channelRow)
        {
            _channelMapRows.Add(channelRow);
        }
        #endregion
    }
}