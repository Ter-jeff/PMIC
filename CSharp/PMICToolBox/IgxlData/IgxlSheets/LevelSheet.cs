using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlSheets
{
    public class LevelSheet : IgxlSheet
    {
        private const string SheetType = "DTLevelSheet";

        #region Field
        private List<LevelRow> _levelData;
        #endregion

        #region Property
        public List<LevelRow> LevelRows
        {
            get { return _levelData; }
            set { _levelData = value; }
        }
        #endregion

        #region Constructor

        public LevelSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _levelData = new List<LevelRow>();
            IgxlSheetName = IgxlSheetNameList.PinLevel;
        }

        public LevelSheet(string sheetName)
            : base(sheetName)
        {
            _levelData = new List<LevelRow>();
            IgxlSheetName = IgxlSheetNameList.PinLevel;
        }


        #endregion

        #region Member Function

        public void AddPowerPinLevel(PowerLevel powerLevel)
        {
            LevelRow lLevelRow;

            lLevelRow = new LevelRow(powerLevel.PinName, "Vmain", powerLevel.vmain, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(powerLevel.PinName, "valt", powerLevel.valt, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(powerLevel.PinName, "iFoldLevel", powerLevel.iFoldLevel, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(powerLevel.PinName, "tdelay", powerLevel.tdelay, "");

            _levelData.Add(lLevelRow);
        }

        public void AddDcviPowerPinLevel(DcviPowerLevel powerLevel)
        {
            LevelRow lLevelRow;

            lLevelRow = new LevelRow(powerLevel.PinName, "Vps", powerLevel.vps, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(powerLevel.PinName, "Isc", powerLevel.isc, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(powerLevel.PinName, "tdelay", powerLevel.tdelay, "");

            _levelData.Add(lLevelRow);
        }

        public void AddIOPinLevel(IoLevel ioLevel)
        {
            LevelRow lLevelRow;

            lLevelRow = new LevelRow(ioLevel.PinName, "vil", ioLevel.vil, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "vih", ioLevel.vih, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "vol", ioLevel.vol, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "voh", ioLevel.voh, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "voh_alt1", ioLevel.voh_alt1, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "voh_alt2", ioLevel.voh_alt2, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "iol", ioLevel.iol, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "ioh", ioLevel.ioh, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "vt", ioLevel.vt, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "vcl", ioLevel.vcl, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "vch", ioLevel.vch, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "voutLoTyp", ioLevel.voutLoTyp, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "voutHiTyp", ioLevel.voutHiTyp, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(ioLevel.PinName, "driverMode", ioLevel.driverMode, "");

            _levelData.Add(lLevelRow);
        }

        public void AddDiffLevel(DiffLevel diffLevel)
        {
            LevelRow lLevelRow;

            lLevelRow = new LevelRow(diffLevel.PinName, "Vicm", diffLevel.Vicm, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "Vid", diffLevel.Vid, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "dVid0", diffLevel.DVid0, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "dVid1", diffLevel.DVid1, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "dVicm0", diffLevel.DVicm0, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "dVicm1", diffLevel.DVicm1, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "Vod", diffLevel.Vod, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "Vod_Alt1", diffLevel.Vod_Alt1, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "Vod_Alt2", diffLevel.Vod_Alt2, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "dVod0", diffLevel.DVod0, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "dVod1", diffLevel.DVod1, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "Iol", diffLevel.Iol, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "Ioh", diffLevel.Ioh, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "VodTyp", diffLevel.VodTyp, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "VocmTyp", diffLevel.VocmTyp, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "Vt", diffLevel.Vt, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "Vcl", diffLevel.Vcl, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "Vch", diffLevel.Vch, "");

            _levelData.Add(lLevelRow);

            lLevelRow = new LevelRow(diffLevel.PinName, "DriverMode", diffLevel.DriverMode, "");

            _levelData.Add(lLevelRow);
        }

        public void AddBaseLevel(LevelRow levelRow)
        {
            _levelData.Add(levelRow);
        }

        //protected override void WriteHeader()
        //{
        //    const string header = "DTLevelSheet,version=2.1:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tPin Levels";
        //    IgxlWriter.WriteLine(header);
        //    IgxlWriter.WriteLine();
        //}

        //protected override void WriteColumnsHeader()
        //{
        //    const string columnsName = "\tPin/Group\tSeq.\tParameter\tValue\tComment\t";
        //    IgxlWriter.WriteLine(columnsName);
        //}

        //protected override void WriteRows()
        //{
        //    foreach (var levelRow in _levelData)
        //    {

        //        var row = new StringBuilder();
        //        if (levelRow.IsBlankRow())
        //        {
        //            IgxlWriter.WriteLine("\n");
        //        }
        //        row.Append("\t");
        //        row.Append(levelRow.PinName);
        //        row.Append("\t");
        //        row.Append("");
        //        row.Append("\t");
        //        row.Append(levelRow.Parameter);
        //        row.Append("\t");
        //        if (levelRow.Value == "")
        //        {
        //            levelRow.Value = " ";
        //        }

        //        if (!levelRow.Value.Equals("") && !levelRow.Value.ToUpper().Equals("HIZ", StringComparison.OrdinalIgnoreCase) &&
        //            !levelRow.Value.ToUpper().Equals("VT", StringComparison.OrdinalIgnoreCase))
        //        {
        //            row.Append("=");
        //        }

        //        row.Append(levelRow.Value);
        //        row.Append("\t");
        //        row.Append(levelRow.Comment);

        //        IgxlWriter.WriteLine(row.ToString());
        //    }
        //}

        public override void Write(string fileName, string version)
        {
            //Action<string> validate = new Action<string>((a) => { });
            //GenLevelsSheet levelsSheetGen = new GenLevelsSheet(fileName, validate, true);
            //foreach (LevelRow levelRow in _levelData)
            //{
            //    levelsSheetGen.AddLevel(levelRow.PinName, levelRow.Parameter, levelRow.Value);
            //}
            //levelsSheetGen.WriteSheet();
            //if (version == "2.1")
            //{
            //    GetSreamWriter(fileName);
            //    WriteHeader();
            //    WriteColumnsHeader();
            //    WriteRows();
            //    CloseStreamWriter();
            //}
            //else
            //    throw new Exception(string.Format("The Level sheet version:{0} is not supported!", version));

            //Support 2.1
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
                else if (dic.ContainsKey("2.1"))
                {
                    var igxlSheetsVersion = dic["2.1"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (LevelRows.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var pinGroupIndex = GetIndexFrom(igxlSheetsVersion, "Pin/Group");
                var seqIndex = GetIndexFrom(igxlSheetsVersion, "Seq.");
                var parameterIndex = GetIndexFrom(igxlSheetsVersion, "Parameter");
                var valueIndex = GetIndexFrom(igxlSheetsVersion, "Value");
                var commentIndex = GetIndexFrom(igxlSheetsVersion, "Comment");

                #region headers
                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    SetColumns(igxlSheetsVersion, i, arr);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data
                for (var index = 0; index < LevelRows.Count; index++)
                {
                    var row = LevelRows[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.PinName))
                    {
                        arr[0] = row.ColumnA;
                        arr[pinGroupIndex] = row.PinName;
                        arr[seqIndex] = row.Seq;
                        arr[parameterIndex] = row.Parameter;

                        if (!row.Value.Equals("") && !row.Value.ToUpper().Equals("HIZ", StringComparison.OrdinalIgnoreCase) &&
                            !row.Value.ToUpper().Equals("VT", StringComparison.OrdinalIgnoreCase))
                            arr[valueIndex] = "=" + row.Value;
                        else
                            arr[valueIndex] = row.Value;
                        arr[commentIndex] = row.Comment;
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
        #endregion
    }
}