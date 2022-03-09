using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlSheets
{
    public class PatSetSubSheet : IgxlSheet
    {
        private const string SheetType = "DTPatternSubroutineSheet";

        #region Field

        private List<PatSetSubRow> _patSetSubData;

        #endregion

        #region Properity

        public List<PatSetSubRow> PatSetSubData
        {
            set { _patSetSubData = value; }
            get { return _patSetSubData; }
        }

        #endregion

        #region Constructor

        public PatSetSubSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _patSetSubData = new List<PatSetSubRow>();
            IgxlSheetName = IgxlSheetNameList.PatternSubroutine;
        }

        public PatSetSubSheet(string sheetName)
            : base(sheetName)
        {
            _patSetSubData = new List<PatSetSubRow>();
            IgxlSheetName = IgxlSheetNameList.PatternSubroutine;
        }

        #endregion

        #region Member Function

        public void AddRow(PatSetSubRow igxlItem)
        {
            _patSetSubData.Add(igxlItem);
        }

        public long GetPatSetSubCnt()
        {
            return _patSetSubData.Count;
        }

        //protected override void WriteHeader()
        //{
        //    const string header = "DTPatternSubroutineSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1	Pattern Subroutine";
        //    IgxlWriter.WriteLine(header);
        //    IgxlWriter.WriteLine();
        //    IgxlWriter.WriteLine();
        //}

        //protected override void WriteColumnsHeader()
        //{

        //}

        //protected override void WriteRows()
        //{
        //    foreach (var patSetSub in _patSetSubData)
        //    {
        //        var row = new StringBuilder();
        //        row.Append("\t");
        //        row.Append(patSetSub.PatternFileName);
        //        row.Append("\t");
        //        row.Append(patSetSub.Comment);
        //        IgxlWriter.WriteLine(row.ToString());
        //    }
        //}

        public override void Write(string fileName,string version ="2.0")
        {
            //if (version == "2.0")
            //{
            //    GetSreamWriter(fileName);
            //    WriteHeader();
            //    WriteColumnsHeader();
            //    WriteRows();
            //    CloseStreamWriter();
            //}
            //else
            //    throw new Exception(string.Format("The PatternSubroutine sheet version:{0} is not supported!"));

            //Support 2.0
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
                else if (dic.ContainsKey("2.0"))
                {
                    var igxlSheetsVersion = dic["2.0"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_patSetSubData.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var patternFilenameIndex = GetIndexFrom(igxlSheetsVersion, "Pattern Filename");
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
                for (var index = 0; index < _patSetSubData.Count; index++)
                {
                    var row = _patSetSubData[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.PatternFileName))
                    {
                        arr[0] = row.ColumnA;
                        arr[patternFilenameIndex] = row.PatternFileName;
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
