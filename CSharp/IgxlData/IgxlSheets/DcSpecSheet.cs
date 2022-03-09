using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using Teradyne.Oasis.IGLinkBase.ProgramGeneration;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class DcSpecSheet : IgxlSheet
    {
        #region Field
        private readonly List<DcSpecs> _dcSpecs;
        private readonly List<string> _selectorNameList;
        private const string SheetType = "DTDCSpecSheet";
        #endregion

        #region Property

        public List<string> CategoryList { get; set; }

        public List<DcSpecs> GetDcSpecsData()
        {
            return _dcSpecs;
        }

        #endregion

        #region Constructor


        public DcSpecSheet(string sheetName, List<string> catList)
            : base(sheetName)
        {
            CategoryList = catList;
            _dcSpecs = new List<DcSpecs>();
            _selectorNameList = GetSelectorList().Select(x=>x.SelectorName).ToList();
            IgxlSheetName = IgxlSheetNameList.DcSpec;
        }

        public DcSpecSheet(ExcelWorksheet sheet, List<string> catList, List<string> selectorNameList)
            : base(sheet)
        {
            CategoryList = catList;
            _dcSpecs = new List<DcSpecs>();
            _selectorNameList = selectorNameList;
            IgxlSheetName = IgxlSheetNameList.DcSpec;
        }

        public DcSpecSheet(string sheetName, List<string> catList, List<string> selectorNameList)
            : base(sheetName)
        {
            CategoryList = catList;
            _dcSpecs = new List<DcSpecs>();
            _selectorNameList = selectorNameList;
            IgxlSheetName = IgxlSheetNameList.DcSpec;
        }
        #endregion

        #region Member Function

        public List<Selector> CreateSelectorList()
        {
            List<Selector> selectorList = new List<Selector>();
            foreach (var name in _selectorNameList)
            {
                selectorList.Add(new Selector(name, ""));
            }
            return selectorList;
        }

        public void AddRow(DcSpecs dcSpecs)
        {
            if (!ExistDcSpecs(dcSpecs.Symbol))
                _dcSpecs.Add(dcSpecs);
        }

        public void AddRows(List<DcSpecs> dcSpecsList)
        {
            foreach (var dcSpecs in dcSpecsList)
            {
                if (!ExistDcSpecs(dcSpecs.Symbol))
                    _dcSpecs.Add(dcSpecs);
            }
        }

        protected override void WriteHeader()
        {
            const string header = "DTDCSpecSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tDC Specs";
            IgxlWriter.WriteLine(header);
            IgxlWriter.WriteLine();
            IgxlWriter.WriteLine();
        }

        protected override void WriteColumnsHeader()
        {
            var firstColumnsName = new StringBuilder();
            var secondColumnsName = new StringBuilder();
            firstColumnsName.Append("\t\t\tSelector\t\t");
            secondColumnsName.Append("\tSymbol\tValue\tName\tVal\tTyp\t");
            foreach (var category in _dcSpecs[0].CategoryList)
            {
                firstColumnsName.Append(category.Name);
                firstColumnsName.Append("\t");
                secondColumnsName.Append("Min\tMax\tTyp\t");
            }
            IgxlWriter.WriteLine(firstColumnsName.ToString());
            IgxlWriter.WriteLine(secondColumnsName.ToString());
        }

        protected override void WriteRows()
        {
            foreach (var dcSpecs in _dcSpecs)
            {
                foreach (var selector in dcSpecs.SelectorList)
                {
                    var dcSpecRow = new StringBuilder();
                    dcSpecRow.Append(dcSpecs.SpecialComment);
                    dcSpecRow.Append("\t");
                    dcSpecRow.Append(dcSpecs.Symbol);
                    dcSpecRow.Append("\t");
                    dcSpecRow.Append(dcSpecs.Value);
                    dcSpecRow.Append("\t");
                    dcSpecRow.Append(selector.SelectorName);
                    dcSpecRow.Append("\t");
                    dcSpecRow.Append(selector.SelectorValue);
                    dcSpecRow.Append("\t");
                    foreach (var category in dcSpecs.CategoryList)
                    {
                        // Write category
                        dcSpecRow.Append(category.Typ);
                        dcSpecRow.Append("\t");
                        dcSpecRow.Append(category.Min);
                        dcSpecRow.Append("\t");
                        dcSpecRow.Append(category.Max);
                        dcSpecRow.Append("\t");
                    }
                    dcSpecRow.Append(dcSpecs.Comment);
                    IgxlWriter.WriteLine(dcSpecRow.ToString());
                }

            }
        }

        public override void Write(string fileName, string version = "2.0")
        {
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version == "2.0")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet2P0(fileName, version, igxlSheetsVersion);
                }
                else if (version=="3.0")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet3P0(fileName, version, igxlSheetsVersion);
                }
                else
                {
                    throw new Exception(string.Format("The DCSpec sheet version:{0} is not supported!", version));
                }
            }
        }

        private void WriteSheet2P0(string fileName,string version,SheetInfo igxlSheetsVersion)
        {
            if (_dcSpecs.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var symbolIndex = GetIndexFrom(igxlSheetsVersion, "Symbol");
                var valueIndex = GetIndexFrom(igxlSheetsVersion, "Value");
                var selectorsIndex = GetIndexFrom(igxlSheetsVersion, "Selector");
                int categoryValuesIndex = selectorsIndex + 2;
                int relativeColumnIndex = selectorsIndex + _selectorNameList.Count + CategoryList.Count * 3;
                maxCount = Math.Max(maxCount, relativeColumnIndex + 1);

                #region headers
                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
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
                            if (item.indexFrom == item.indexTo && item.rowIndex == i)
                                arr[item.indexFrom] = item.columnName;

                            if ((item.columnName == "Selector") && item.rowIndex + 1 == i)
                            {
                                for (int index = 0; index < _selectorNameList.Count; index++)
                                {
                                    var selectorName = _selectorNameList[index];
                                    arr[selectorsIndex + index] = selectorName;
                                }
                            }

                            if (item.variantName == "CategoryValues")
                            {
                                int categoryCount = 0;
                                foreach (var category in CategoryList)
                                {
                                    foreach (var column in item.Column1)
                                    {
                                        if (column.indexFrom == 5 && item.rowIndex == i)
                                            arr[categoryValuesIndex - 1 + categoryCount * 3 + column.indexFrom - 4] = category;
                                        if (column.rowIndex == i)
                                            arr[categoryValuesIndex - 1 + categoryCount * 3 + column.indexFrom - 4] = column.columnName;
                                    }
                                    categoryCount++;
                                }
                            }
                        }
                    }
                    #endregion

                    SetRelativeColumn(igxlSheetsVersion, i, arr, relativeColumnIndex);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data
                for (var index = 0; index < _dcSpecs.Count; index++)
                {
                    var row = _dcSpecs[index];
                    foreach (var selector in row.SelectorList)
                    {
                        var arr = Enumerable.Repeat("", maxCount).ToArray();
                        if (!string.IsNullOrEmpty(row.Symbol))
                        {
                            arr[0] = row.ColumnA;
                            arr[symbolIndex] = row.Symbol;
                            arr[valueIndex] = "=#N/A";

                            for (int i = 0; i < _selectorNameList.Count; i++)
                            {
                                var selectorName =
                                    row.SelectorList.Find(
                                        x =>
                                            x.SelectorName.Equals(_selectorNameList[i],
                                                StringComparison.CurrentCultureIgnoreCase));
                                if (selectorName != null)
                                    arr[selectorsIndex + i] = selectorName.SelectorValue;
                                else
                                    arr[selectorsIndex + i] = "";
                            }

                            arr[selectorsIndex] = selector.SelectorValue;
                            arr[selectorsIndex + 1] = selector.SelectorValue;
                            int categoryCount = 0;
                            foreach (var category in row.CategoryList)
                            {
                                arr[categoryValuesIndex + categoryCount * 3] = category.Typ;
                                arr[categoryValuesIndex + categoryCount * 3 + 1] = category.Min;
                                arr[categoryValuesIndex + categoryCount * 3 + 2] = category.Max;
                                categoryCount++;
                            }
                            arr[relativeColumnIndex] = row.Comment;
                        }
                        else
                        {
                            arr = new[] { "\t" };
                        }
                        sw.WriteLine(string.Join("\t", arr));
                    }
                }
                #endregion
            }

            var testSetting = _dcSpecs.Find(x => !string.IsNullOrEmpty(x.SpecialComment));
            if (testSetting != null)
                AddVersion(fileName, testSetting.SpecialComment);
        }

        private void WriteSheet3P0(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_dcSpecs.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var symbolIndex = GetIndexFrom(igxlSheetsVersion, "Symbol");
                var valueIndex = GetIndexFrom(igxlSheetsVersion, "Value");
                var selectorsIndex = GetIndexFrom(igxlSheetsVersion, "Selectors");
                int categoryValuesIndex = selectorsIndex + _selectorNameList.Count;
                int relativeColumnIndex = selectorsIndex + _selectorNameList.Count + CategoryList.Count * 3;
                maxCount = Math.Max(maxCount, relativeColumnIndex + 1);

                #region headers
                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    if (igxlSheetsVersion.Columns.Column != null)
                    {
                        foreach (var item in igxlSheetsVersion.Columns.Column)
                        {
                            if (item.rowIndex == i)
                                arr[item.indexFrom] = item.columnName;
                            if (item.Column1 != null)
                            {
                                foreach (var column1 in item.Column1)
                                {
                                    if (column1.rowIndex == i)
                                        arr[column1.indexFrom] = column1.columnName;
                                }
                            }
                        }
                    }

                    #region Set Variant
                    if (igxlSheetsVersion.Columns.Variant != null)
                    {
                        foreach (var item in igxlSheetsVersion.Columns.Variant)
                        {
                            if (item.indexFrom == item.indexTo && item.rowIndex == i)
                                arr[item.indexFrom] = item.columnName;

                            if ((item.columnName == "Selectors") && item.rowIndex + 1 == i)
                            {
                                for (int index = 0; index < _selectorNameList.Count; index++)
                                {
                                    var selectorName = _selectorNameList[index];
                                    arr[selectorsIndex + index] = selectorName;
                                }
                            }

                            if (item.variantName == "CategoryValues")
                            {
                                int categoryCount = 0;
                                foreach (var category in CategoryList)
                                {
                                    foreach (var column in item.Column1)
                                    {
                                        if (column.indexFrom == 1 && item.rowIndex == i)
                                            arr[categoryValuesIndex - 1 + categoryCount * 3 + column.indexFrom] = category;
                                        if (column.rowIndex == i)
                                            arr[categoryValuesIndex - 1 + categoryCount * 3 + column.indexFrom] = column.columnName;
                                    }
                                    categoryCount++;
                                }
                            }
                        }
                    }
                    #endregion

                    SetRelativeColumn(igxlSheetsVersion, i, arr, relativeColumnIndex);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data
                for (var index = 0; index < _dcSpecs.Count; index++)
                {
                    var row = _dcSpecs[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.Symbol))
                    {
                        arr[0] = row.ColumnA;
                        arr[symbolIndex] = row.Symbol;
                        arr[valueIndex] = "=#N/A";
                        for (int i = 0; i < _selectorNameList.Count; i++)
                        {
                            var selectorName = row.SelectorList.Find(x => x.SelectorName.Equals(_selectorNameList[i], StringComparison.CurrentCultureIgnoreCase));
                            if (selectorName != null)
                                arr[selectorsIndex + i] = selectorName.SelectorValue;
                            else
                                arr[selectorsIndex + i] = "";
                        }
                        int categoryCount = 0;
                        foreach (var category in row.CategoryList)
                        {
                            arr[categoryValuesIndex + categoryCount * 3] = category.Typ;
                            arr[categoryValuesIndex + categoryCount * 3 + 1] = category.Min;
                            arr[categoryValuesIndex + categoryCount * 3 + 2] = category.Max;
                            categoryCount++;
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

            var testSetting = _dcSpecs.Find(x => !string.IsNullOrEmpty(x.SpecialComment));
            if (testSetting != null)
                AddVersion(fileName, testSetting.SpecialComment);
        }

        private void AddVersion(string filename, string version)
        {
            string[] allLine = File.ReadAllLines(filename);
            if (allLine.Length >= 5)
            {
                allLine[4] = allLine[4] + version;
                File.WriteAllLines(filename, allLine);
            }
        }

        public DcSpecs FindDcSpecs(string symbol)
        {
            return _dcSpecs.Find(dc => dc.Symbol.Equals(symbol));
        }

        public bool ExistDcSpecs(string symbol)
        {
            return _dcSpecs.Exists(dc => dc.Symbol.Equals(symbol));
        }

        public bool FindValue(string dcCategory, string dcSelector, ref string frequencyName)
        {
            if (Regex.IsMatch(frequencyName, "^_"))
                frequencyName = Regex.Replace(frequencyName, "^_", "");
            foreach (var dcSpec in _dcSpecs)
            {
                if (dcSpec.Symbol.Equals(frequencyName, StringComparison.CurrentCultureIgnoreCase))
                {
                    if (dcSpec.CategoryList.Exists(x => x.Name.Equals(dcCategory, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var row = dcSpec.CategoryList.Find(x => x.Name.Equals(dcCategory, StringComparison.CurrentCultureIgnoreCase));
                        if (dcSelector.Equals("Typ", StringComparison.CurrentCultureIgnoreCase))
                        {
                            frequencyName = row.Typ;
                            return true;
                        }
                        if (dcSelector.Equals("Max", StringComparison.CurrentCultureIgnoreCase))
                        {
                            frequencyName = row.Max;
                            return true;
                        }
                        if (dcSelector.Equals("Min", StringComparison.CurrentCultureIgnoreCase))
                        {
                            frequencyName = row.Min;
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private List<Selector> GetSelectorList()
        {
            var selectorList = new List<Selector>();
            selectorList.Add(new Selector("Min", "Min"));
            selectorList.Add(new Selector("Typ", "Typ"));
            selectorList.Add(new Selector("Max", "Max"));
            return selectorList;
        }

        public int FindCategoryIndex(string category)
        {
            if (!CategoryList.Exists(x => x.Equals(category, StringComparison.OrdinalIgnoreCase))
            ) return -1;
            return CategoryList.FindIndex(x => x.Equals(category, StringComparison.OrdinalIgnoreCase));
        }


        //public DcSpecs FindDcSpecs(string symbol)
        //{
        //    return DcSpecSheet.FindDcSpecs(symbol);
        //}
        #endregion


    }
}