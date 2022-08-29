using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using IgxlData.IgxlBase;
using OfficeOpenXml;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    [Serializable]
    public class TimeSetBasicSheet : IgxlSheet
    {
        #region Field

        private const string SheetType = "DTTimesetBasicSheet";

        #region Const Filed

        public const string TimeModeSingle = "Single";
        public const string TimeModeDual = "Dual";
        public const string TimeModeQuad = "Quad";

        #endregion

        public List<Tset> Tsets;

        #endregion

        #region Property

        public string TimingMode { get; set; }
        public string MasterTimeSet { get; set; }
        public string TimeDomain { get; set; }
        public string StrobeRefSetup { get; set; }

        public List<Tset> TimeSetsData
        {
            get { return Tsets; }
            set { Tsets = value; }
        }

        #endregion

        #region Constructor

        public TimeSetBasicSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            Tsets = new List<Tset>();
            IgxlSheetName = IgxlSheetNameList.TimeSetsBasic;
            TimingMode = "";
        }

        public TimeSetBasicSheet(string sheetName)
            : base(sheetName)
        {
            Tsets = new List<Tset>();
            IgxlSheetName = IgxlSheetNameList.TimeSetsBasic;
            TimingMode = "";
        }

        public TimeSetBasicSheet(ExcelWorksheet sheet, string timingMode)
            : base(sheet)
        {
            Tsets = new List<Tset>();
            TimingMode = timingMode;
            IgxlSheetName = IgxlSheetNameList.TimeSetsBasic;
        }

        public TimeSetBasicSheet(string sheetName, string timingMode)
            : base(sheetName)
        {
            Tsets = new List<Tset>();
            TimingMode = timingMode;
            IgxlSheetName = IgxlSheetNameList.TimeSetsBasic;
        }

        public TimeSetBasicSheet(ExcelWorksheet sheet, string timingMode, string masterTimeSet, string timeDomain,
            string strobeRefSetup)
            : base(sheet)
        {
            Tsets = new List<Tset>();
            TimingMode = timingMode;
            MasterTimeSet = masterTimeSet;
            TimeDomain = timeDomain;
            StrobeRefSetup = strobeRefSetup;
            IgxlSheetName = IgxlSheetNameList.TimeSetsBasic;
        }

        public TimeSetBasicSheet(string sheetName, string timingMode, string masterTimeSet, string timeDomain,
            string strobeRefSetup)
            : base(sheetName)
        {
            Tsets = new List<Tset>();
            TimingMode = timingMode;
            MasterTimeSet = masterTimeSet;
            TimeDomain = timeDomain;
            StrobeRefSetup = strobeRefSetup;
            IgxlSheetName = IgxlSheetNameList.TimeSetsBasic;
        }

        #endregion

        #region Member Function

        protected override void WriteHeader()
        {
            const string header =
                "DTTimesetBasicSheet,version=1.4:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tTime Sets (Basic)\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t";
            IgxlWriter.WriteLine(header);
            IgxlWriter.WriteLine("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
            IgxlWriter.WriteLine("\tTiming Mode:\t" + TimingMode + "\t\tMaster Timeset Name:\t" +
                                 MasterTimeSet + "\t\t\t\t\t\t\t\t\t\t\t\t");
            IgxlWriter.WriteLine("\tTime Domain:\t" + TimeDomain + "\t\tStrobe Ref Setup Name:\t" +
                                 StrobeRefSetup + "\t\t\t\t\t\t\t\t\t\t\t\t");
            IgxlWriter.WriteLine("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
        }

        protected override void WriteColumnsHeader()
        {
            const string columnsName =
                "\t\tCycle\tPin Group\t\t\tData\t\tDrive\t\t\t\tCompare\t\t\tEdge Resolution\t\t";
            IgxlWriter.WriteLine(columnsName);
            IgxlWriter.WriteLine(
                "\tTime Set\tPeriod\tName\tClock Period\tSetup\tSrc\tFmt\tOn\tData\tReturn\tOff\tMode\tOpen\tClose\tMode\tComment\t");
        }

        protected override void WriteRows()
        {
            foreach (var timeSets in Tsets)
            foreach (var timingRow in timeSets.TimingRows)
            {
                var row = new StringBuilder();
                row.Append("\t");
                row.Append(timeSets.Name);
                row.Append("\t");
                row.Append(timeSets.CyclePeriod);
                row.Append("\t");

                row.Append(timingRow.PinGrpName);
                row.Append("\t");
                row.Append(timingRow.PinGrpClockPeriod);
                row.Append("\t");
                row.Append(timingRow.PinGrpSetup);
                row.Append("\t");

                row.Append(timingRow.DataSrc);
                row.Append("\t");
                row.Append(timingRow.DataFmt);
                row.Append("\t");

                row.Append(timingRow.DriveOn);
                row.Append("\t");
                row.Append(timingRow.DriveData);
                row.Append("\t");
                row.Append(timingRow.DriveReturn);
                row.Append("\t");
                row.Append(timingRow.DriveOff);
                row.Append("\t");

                row.Append(timingRow.CompareMode);
                row.Append("\t");
                row.Append(timingRow.CompareOpen);
                row.Append("\t");
                row.Append(timingRow.CompareClose);
                row.Append("\t");
                //row.Append(timingRow.CompareClkOffset);
                //row.Append("\t");
                //row.Append(timingRow.CompareRefOffset);
                //row.Append("\t");

                row.Append(timingRow.EdgeMode);
                row.Append("\t");

                row.Append(timingRow.Comment);
                row.Append("\t");

                IgxlWriter.WriteLine(row.ToString());
            }
        }

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
                version = "2.3";
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                Directory.CreateDirectory(Path.GetDirectoryName(fileName));
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version == "1.4")
                {
                    GetStreamWriter(fileName);
                    WriteHeader();
                    WriteColumnsHeader();
                    WriteRows();
                    CloseStreamWriter();
                }
                else if (version == "2.3")
                {
                    var igxlSheetsVersion = dic["2.3"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else
                {
                    throw new Exception(string.Format("The timeSet version:{0} is not supported!", version));
                }
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (Tsets.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var timeSetIndex = GetIndexFrom(igxlSheetsVersion, "Time Set");
                var periodIndex = GetIndexFrom(igxlSheetsVersion, "Cycle", "Period");
                var nameIndex = GetIndexFrom(igxlSheetsVersion, "Pin/Group", "Name");
                var clockPeriodIndex = GetIndexFrom(igxlSheetsVersion, "Pin/Group", "Clock Period");
                var setupIndex = GetIndexFrom(igxlSheetsVersion, "Pin/Group", "Setup");
                var srcIndex = GetIndexFrom(igxlSheetsVersion, "Data", "Src");
                var fmtIndex = GetIndexFrom(igxlSheetsVersion, "Data", "Fmt");
                var onIndex = GetIndexFrom(igxlSheetsVersion, "Drive", "On");
                var dataIndex = GetIndexFrom(igxlSheetsVersion, "Drive", "Data");
                var returnIndex = GetIndexFrom(igxlSheetsVersion, "Drive", "Return");
                var offIndex = GetIndexFrom(igxlSheetsVersion, "Drive", "Off");
                var compareModeIndex = GetIndexFrom(igxlSheetsVersion, "Compare", "Mode");
                var compareOpenIndex = GetIndexFrom(igxlSheetsVersion, "Compare", "Open");
                var compareCloseIndex = GetIndexFrom(igxlSheetsVersion, "Compare", "Close");
                var compareRefOffsetIndex = GetIndexFrom(igxlSheetsVersion, "Compare", "Ref Offset");
                var edgeModeIndex = GetIndexFrom(igxlSheetsVersion, "Edge Resolution", "Mode");
                var commentIndex = GetIndexFrom(igxlSheetsVersion, "Comment");

                #region headers

                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" +
                             IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    if (igxlSheetsVersion.Field != null)
                        foreach (var item in igxlSheetsVersion.Field)
                            if (item.rowIndex == i)
                            {
                                arr[item.columnIndex] = item.fieldName;
                                if (item.fieldName.Equals("Timing Mode:", StringComparison.CurrentCultureIgnoreCase))
                                {
                                    if (string.IsNullOrEmpty(TimingMode))
                                        arr[item.columnIndex + 1] = "Single";
                                    else
                                        arr[item.columnIndex + 1] = TimingMode;
                                }

                                if (item.fieldName.Equals("Master Timeset Name:",
                                        StringComparison.CurrentCultureIgnoreCase))
                                    arr[item.columnIndex + 1] = MasterTimeSet;
                                if (item.fieldName.Equals("Time Domain:", StringComparison.CurrentCultureIgnoreCase))
                                    arr[item.columnIndex + 1] = TimeDomain;
                                if (item.fieldName.Equals("Strobe Ref Setup Name:",
                                        StringComparison.CurrentCultureIgnoreCase))
                                    arr[item.columnIndex + 1] = StrobeRefSetup;
                            }

                    SetColumns(igxlSheetsVersion, i, arr);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data

                for (var index = 0; index < Tsets.Count; index++)
                {
                    var tset = Tsets[index];
                    for (var i = 0; i < tset.TimingRows.Count; i++)
                    {
                        var row = tset.TimingRows[i];
                        var arr = Enumerable.Repeat("", maxCount).ToArray();
                        if (!string.IsNullOrEmpty(row.PinGrpName))
                        {
                            arr[0] = tset.ColumnA;
                            arr[timeSetIndex] = tset.Name;
                            arr[periodIndex] = tset.CyclePeriod;
                            arr[nameIndex] = row.PinGrpName;
                            arr[clockPeriodIndex] = row.PinGrpClockPeriod;
                            arr[setupIndex] = row.PinGrpSetup;
                            arr[srcIndex] = row.DataSrc;
                            arr[fmtIndex] = row.DataFmt;
                            arr[onIndex] = row.DriveOn;
                            arr[dataIndex] = row.DriveData;
                            arr[returnIndex] = row.DriveReturn;
                            arr[offIndex] = row.DriveOff;
                            arr[compareModeIndex] = row.CompareMode;
                            arr[compareOpenIndex] = row.CompareOpen;
                            arr[compareCloseIndex] = row.CompareClose;
                            arr[compareRefOffsetIndex] = row.CompareRefOffset;
                            arr[edgeModeIndex] = row.EdgeMode;
                            arr[commentIndex] = row.Comment;
                        }
                        else
                        {
                            arr = new[] {"\t"};
                        }

                        sw.WriteLine(string.Join("\t", arr));
                    }
                }

                #endregion
            }
        }

        public void AddTimeSet(Tset timeSet)
        {
            Tsets.Add(timeSet);
        }

        public TimeSetBasicSheet DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as TimeSetBasicSheet;
            }
        }

        #endregion
    }
}