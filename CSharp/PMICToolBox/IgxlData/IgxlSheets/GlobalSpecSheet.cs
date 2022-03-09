using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlSheets
{
    public class GlobalSpecSheet : IgxlSheet
    {
        private const string SheetType = "DTGlobalSpecSheet";

        #region Field
        public List<GlobalSpec> Specs;
        #endregion

        #region Constructor
        public GlobalSpecSheet(ExcelWorksheet sheet, bool isAddDefault = true)
            : base(sheet)
        {
            Specs = new List<GlobalSpec>();
            if (isAddDefault)
                AddDefaultRow();
            IgxlSheetName = IgxlSheetNameList.GlobalSpec;
        }

        public GlobalSpecSheet(string sheetName, bool isAddDefault = true)
            : base(sheetName)
        {
            Specs = new List<GlobalSpec>();
            if (isAddDefault)
                AddDefaultRow();
            IgxlSheetName = IgxlSheetNameList.GlobalSpec;
        }
        #endregion

        #region Member Function
        public void AddRow(GlobalSpec globalSpec)
        {
            Specs.Add(globalSpec);
        }

        public void AddRange(List<GlobalSpec> globalSpecsList)
        {
            Specs.AddRange(globalSpecsList);
        }

        public void InsertRow(GlobalSpec globalSpec)
        {
            var foundGlobalSpecs = FindRowList(globalSpec.Symbol, globalSpec.Job);
            if (foundGlobalSpecs.Count > 0)
            {
                return;
            }
            else
            {
                foundGlobalSpecs = FindRowList(globalSpec.Symbol);
                if (foundGlobalSpecs.Count > 0)
                {
                    var rowIndex = Specs.LastIndexOf(foundGlobalSpecs[foundGlobalSpecs.Count - 1]);
                    //Edward
                    Specs.Insert(rowIndex + 1, globalSpec);
                }
                else
                {
                    AddRow(globalSpec);
                }
            }
        }

        public List<GlobalSpec> FindRowList(string globalSpecSymbol, string globalSpecJob)
        {
            var foundGlobalSpecs = Specs.FindAll(g => g.Symbol.Equals(globalSpecSymbol, StringComparison.CurrentCultureIgnoreCase) &&
                                                             g.Job.Equals(globalSpecJob, StringComparison.CurrentCultureIgnoreCase));
            return foundGlobalSpecs;
        }

        public List<GlobalSpec> FindRowList(string globalSpecSymbol)
        {
            var foundGlobalSpecs = Specs.FindAll(g => g.Symbol.Equals(globalSpecSymbol, StringComparison.CurrentCultureIgnoreCase));
            return foundGlobalSpecs;
        }

        private void AddDefaultRow()
        {
            var vclDefault = new GlobalSpec("Vcl_default", "=-1", "", "Detector clamp voltage low");
            AddRow(vclDefault);
            var vchDefault = new GlobalSpec("Vch_default", "=6", "", "Detector clamp voltage high");
            AddRow(vchDefault);
            var vphDefault = new GlobalSpec("Vph_default", "=5", "", "Hi-V pin voltage high");
            AddRow(vphDefault);
        }

        //protected override void WriteHeader()
        //{
        //    const string headerLine = "DTGlobalSpecSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tGlobal Specs\t\t\t\t";
        //    IgxlWriter.WriteLine(headerLine);
        //    IgxlWriter.WriteLine("\t\t\t\t\t");
        //}

        //protected override void WriteColumnsHeader()
        //{
        //    const string columnsName = "\tSymbol\tJob\tValue\tComment\t";
        //    IgxlWriter.WriteLine(columnsName);
        //}

        //protected override void WriteRows()
        //{
        //    foreach (var globalSpec in Specs)
        //    {
        //        var globalRow = new StringBuilder();
        //        globalRow.Append("\t");
        //        globalRow.Append(globalSpec.Symbol);
        //        globalRow.Append("\t");
        //        globalRow.Append(globalSpec.Job);
        //        globalRow.Append("\t");
        //        globalRow.Append(globalSpec.Value);
        //        globalRow.Append("\t");
        //        globalRow.Append(globalSpec.Comment);
        //        IgxlWriter.WriteLine(globalRow.ToString());
        //    }
        //}

        public override void Write(string fileName, string version)
        {
            //if (version == "2.0")
            //{
            //    var validate = new Action<string>((a) => { });
            //    var glbSpecGenerator = new GenGlobalSpecSheet(fileName, validate, true);
            //    foreach (var globalSpec in Specs)
            //    {
            //        glbSpecGenerator.AddGlobalSpec(globalSpec.Symbol, globalSpec.Job, globalSpec.Value, globalSpec.Comment);
            //    }
            //    glbSpecGenerator.WriteSheet();
            //}
            //else
            //    throw new Exception(string.Format("The GlobalSpec sheet version:{0}", version));

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
            if (Specs.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var symbolIndex = GetIndexFrom(igxlSheetsVersion, "Symbol");
                var jobIndex = GetIndexFrom(igxlSheetsVersion, "Job");
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
                for (var index = 0; index < Specs.Count; index++)
                {
                    var row = Specs[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.Symbol))
                    {
                        arr[0] = row.ColumnA;
                        arr[symbolIndex] = row.Symbol;
                        arr[jobIndex] = row.Job;
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