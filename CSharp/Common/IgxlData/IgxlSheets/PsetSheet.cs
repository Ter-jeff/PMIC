using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using IgxlData.IgxlBase;
using OfficeOpenXml;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class PSetSheet : IgxlSheet
    {
        #region Field

        private const string SheetType = "DTPsetsSheet";

        #endregion

        #region Constructor

        public PSetSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            PSets = new List<PSet>();
            IgxlSheetName = IgxlSheetNameList.PSet;
        }

        #endregion

        //Name	Pin	Instrument Type

        #region Property

        public List<PSet> PSets { get; set; }

        #endregion

        #region Member Function

        public void AddRow(PSet pSet)
        {
            PSets.Add(pSet);
        }

        protected override void WriteHeader()
        {
        }

        protected override void WriteColumnsHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteRows()
        {
            throw new NotImplementedException();
        }

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
                version = "2.0";
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version == "2.0")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else
                {
                    throw new Exception(string.Format("The Pset sheet version:{0} is not supported!", version));
                }
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (PSets.Count == 0) return;

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
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" +
                             IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    SetColumns(igxlSheetsVersion, i, arr);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data

                for (var index = 0; index < PSets.Count; index++)
                {
                    var row = PSets[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.Name))
                    {
                        arr[0] = row.ColumnA;
                        arr[nameIndex] = row.Name;
                        arr[pinIndex] = row.Pin;
                        arr[instrumentTypeIndex] = row.InstrumentType;
                        arr[thislargeheadingmakesAutoFitenlargetherowheightIndex] =
                            row.ThislargeheadingmakesAutoFitenlargetherowheight;
                        for (var j = 0; j < row.Parameters.Count; j++)
                            if (j < 80)
                                arr[parameter0Index + j] = row.Parameters.ElementAt(j).Value;
                        arr[commentIndex] = row.Comment;
                    }
                    else
                    {
                        arr = new[] {"\t"};
                    }

                    sw.WriteLine(string.Join("\t", arr));
                }

                #endregion
            }
        }

        private void WritePset2_0(string fileName, string version)
        {
            CreateFolder(Path.GetDirectoryName(fileName));
            var writer = new StreamWriter(fileName);

            var firstLine1 = "DTPsetsSheet,version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1";
            var firstLine2 = "PSets";
            writer.WriteLine(firstLine1 + '\t' + firstLine2);
            writer.WriteLine();
            //Name	Pin	Instrument Type	This large heading makes AutoFit enlarge the row height	par1

            var subHeaders = new List<string>
                {"", "Name", "Pin", "Instrument Type", "This large heading makes AutoFit enlarge the row height"};

            writer.WriteLine(string.Join("\t", subHeaders));
            foreach (var pSet in PSets)
            {
                var info = new List<string> {"", pSet.Name, pSet.Pin, pSet.InstrumentType, ""};
                info.AddRange(pSet.Parameters.Select(p => p.Value));
                writer.WriteLine(string.Join("\t", info));
            }

            writer.Close();
        }

        private void CreateFolder(string pFolder)
        {
            if (!Directory.Exists(pFolder)) Directory.CreateDirectory(pFolder);
        }

        #endregion
    }
}