using IgxlData.IgxlBase;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class GlobalSpecSheet : IgxlSheet
    {
        public List<GlobalSpec> GetGlobalSpecs()
        {
            return GlobalSpecsRows;
        }

        public bool IsExist(string symbol)
        {
            return GlobalSpecsRows.Exists(x => x.Symbol.Equals(symbol, StringComparison.CurrentCultureIgnoreCase));
        }

        private const string SheetType = "DTGlobalSpecSheet";
        public List<GlobalSpec> GlobalSpecsRows;

        public GlobalSpecSheet(ExcelWorksheet sheet, bool isAddDefault = true)
            : base(sheet)
        {
            GlobalSpecsRows = new List<GlobalSpec>();
            if (isAddDefault)
                AddDefaultRow();
            IgxlSheetName = IgxlSheetNameList.GlobalSpec;
        }

        public GlobalSpecSheet(string sheetName, bool isAddDefault = true)
            : base(sheetName)
        {
            GlobalSpecsRows = new List<GlobalSpec>();
            if (isAddDefault)
                AddDefaultRow();
            IgxlSheetName = IgxlSheetNameList.GlobalSpec;
        }

        public void AddRow(GlobalSpec globalSpec)
        {
            GlobalSpecsRows.Add(globalSpec);
        }

        public void AddRows(List<GlobalSpec> globalSpecsList)
        {
            GlobalSpecsRows.AddRange(globalSpecsList);
        }

        public void InsertRow(GlobalSpec globalSpec)
        {
            var foundGlobalSpecs = FindRowList(globalSpec.Symbol, globalSpec.Job);
            if (foundGlobalSpecs.Count > 0)
            {
            }
            else
            {
                foundGlobalSpecs = FindRowList(globalSpec.Symbol);
                if (foundGlobalSpecs.Count > 0)
                {
                    var rowIndex = GlobalSpecsRows.LastIndexOf(foundGlobalSpecs[foundGlobalSpecs.Count - 1]);
                    //Edward
                    GlobalSpecsRows.Insert(rowIndex + 1, globalSpec);
                }
                else
                {
                    AddRow(globalSpec);
                }
            }
        }

        public List<GlobalSpec> FindRowList(string globalSpecSymbol, string globalSpecJob)
        {
            var foundGlobalSpecs = GlobalSpecsRows.FindAll(g =>
                g.Symbol.Equals(globalSpecSymbol, StringComparison.CurrentCultureIgnoreCase) &&
                g.Job.Equals(globalSpecJob, StringComparison.CurrentCultureIgnoreCase));
            return foundGlobalSpecs;
        }

        public List<GlobalSpec> FindRowList(string globalSpecSymbol)
        {
            var foundGlobalSpecs = GlobalSpecsRows.FindAll(g =>
                g.Symbol.Equals(globalSpecSymbol, StringComparison.CurrentCultureIgnoreCase));
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

        protected void WriteHeader()
        {
            const string headerLine =
                "DTGlobalSpecSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tGlobal Specs\t\t\t\t";
            IgxlWriter.WriteLine(headerLine);
            IgxlWriter.WriteLine("\t\t\t\t\t");
        }

        protected void WriteColumnsHeader()
        {
            const string columnsName = "\tSymbol\tJob\tValue\tComment\t";
            IgxlWriter.WriteLine(columnsName);
        }

        protected void WriteRows()
        {
            foreach (var globalSpec in GlobalSpecsRows)
            {
                var globalRow = new StringBuilder();
                globalRow.Append("\t");
                globalRow.Append(globalSpec.Symbol);
                globalRow.Append("\t");
                globalRow.Append(globalSpec.Job);
                globalRow.Append("\t");
                globalRow.Append(globalSpec.Value);
                globalRow.Append("\t");
                globalRow.Append(globalSpec.Comment);
                IgxlWriter.WriteLine(globalRow.ToString());
            }
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
                    throw new Exception(string.Format("The GlobalSpec sheet version:{0}", version));
                }
            }
        }


        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (GlobalSpecsRows.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var symbolIndex = GetIndexFrom(igxlSheetsVersion, "Symbol");
                var jobIndex = GetIndexFrom(igxlSheetsVersion, "Job");
                var valueIndex = GetIndexFrom(igxlSheetsVersion, "Value");
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

                for (var index = 0; index < GlobalSpecsRows.Count; index++)
                {
                    var row = GlobalSpecsRows[index];
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
    }
}