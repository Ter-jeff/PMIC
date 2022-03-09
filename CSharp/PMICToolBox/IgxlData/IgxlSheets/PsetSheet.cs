using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlSheets
{
    public class PSetSheet : IgxlSheet
    {
        private const string SheetType = "DTPsetsSheet";

        #region Field
        private List<PSet> _pSets;
        #endregion

        #region Property

        public List<PSet> PSets
        {
            get { return _pSets; }
            set { _pSets = value; }
        }
        #endregion

        #region Constructor
        public PSetSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _pSets = new List<PSet>();
            IgxlSheetName = IgxlSheetNameList.PSet;
        }

        #endregion

        #region Member Function

        public void AddRow(PSet pSet)
        {
            _pSets.Add(pSet);
        }

        public override void Write(string fileName, string version = "2.0")
        {
            //CreateFolder(Path.GetDirectoryName(fileName));
            //var writer = new StreamWriter(fileName);

            //const string firstline_1 = "DTPsetsSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1";
            //const string firstline_2 = "PSets";
            //writer.WriteLine(firstline_1 + '\t' + firstline_2);
            //writer.WriteLine();
            ////Name	Pin	Instrument Type	This large heading makes AutoFit enlarge the row height	par1

            //var subHeaders = new List<string> { "", "Name", "Pin", "Instrument Type", "This large heading makes AutoFit enlarge the row height" };

            //writer.WriteLine(string.Join("\t", subHeaders));
            //foreach (var pset in _pSets)
            //{
            //    var info = new List<string> { "", pset.Name, pset.Pin, pset.InstrumentType, "" };
            //    info.AddRange(pset.Parameters.Select(p => p.Value));
            //    writer.WriteLine(string.Join("\t", info));
            //}
            //writer.Close();

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
            if (_pSets.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var nameIndex = GetIndexFrom(igxlSheetsVersion, "Name");
                var pinIndex = GetIndexFrom(igxlSheetsVersion, "Pin");
                var instrumentTypeIndex = GetIndexFrom(igxlSheetsVersion, "Instrument Type");
                var thislargeheadingmakesAutoFitenlargetherowheightIndex =
                    GetIndexFrom(igxlSheetsVersion, "This large heading makes AutoFit enlarge the row height");
                var parameter0Index = GetIndexFrom(igxlSheetsVersion, "Parameter0");
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
                for (var index = 0; index < _pSets.Count; index++)
                {
                    var row = _pSets[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.Name))
                    {
                        arr[0] = row.ColumnA;
                        arr[nameIndex] = row.Name;
                        arr[pinIndex] = row.Pin;
                        arr[instrumentTypeIndex] = row.InstrumentType;
                        arr[thislargeheadingmakesAutoFitenlargetherowheightIndex] = row.ThislargeheadingmakesAutoFitenlargetherowheight;
                        for (int j = 0; j < row.Parameters.Count; j++)
                        {
                            if (j < 80)
                                arr[parameter0Index + j] = row.Parameters.ElementAt(j).Value;
                        }
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

        //private void CreateFolder(string pFolder)
        //{
        //    if (!Directory.Exists(pFolder))
        //    {
        //        Directory.CreateDirectory(pFolder);
        //    }
        //}
        #endregion
    }
}