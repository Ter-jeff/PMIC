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
    public class AcSpecSheet : IgxlSheet
    {
        public string GetAcByTimeSet(string timeset)
        {
            foreach (var item in CategoryTimeSetDic)
                if (item.Value.Contains(timeset, StringComparer.CurrentCultureIgnoreCase))
                    return item.Key;
            return "";
        }

        #region Field

        private const string SheetType = "DTACSpecSheet";
        private readonly List<string> _selectorNameList;

        #endregion

        #region Property

        public List<string> CategoryList { get; set; }

        public Dictionary<string, List<string>> CategoryTimeSetDic = new Dictionary<string, List<string>>();

        public List<AcSpec> AcSpecs { get; set; }

        #endregion

        #region Constructor

        public AcSpecSheet(ExcelWorksheet sheet, List<string> categoryList, List<string> selectorNameList)
            : base(sheet)
        {
            AcSpecs = new List<AcSpec>();
            CategoryList = categoryList;
            _selectorNameList = selectorNameList;
            IgxlSheetName = IgxlSheetNameList.AcSpec;
        }

        public AcSpecSheet(string sheetName, List<string> categoryList, List<string> selectorNameList)
            : base(sheetName)
        {
            AcSpecs = new List<AcSpec>();
            CategoryList = categoryList;
            _selectorNameList = selectorNameList;
            IgxlSheetName = IgxlSheetNameList.AcSpec;
        }

        #endregion

        #region Member Function

        public void AddRow(AcSpec acSpecRow)
        {
            AcSpecs.Add(acSpecRow);
        }

        protected override void WriteHeader()
        {
            throw new NotImplementedException();
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
                version = "3.0";
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version == "2.0")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet2P0(fileName, version, igxlSheetsVersion);
                }
                else if (version == "3.0")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet3P0(fileName, version, igxlSheetsVersion);
                }
                else
                {
                    throw new Exception(string.Format("The AC Spec sheet version:{0} is not supported!", version));
                }
            }
        }

        private void WriteSheet2P0(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (AcSpecs.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var symbolIndex = GetIndexFrom(igxlSheetsVersion, "Symbol");
                var valueIndex = GetIndexFrom(igxlSheetsVersion, "Value");
                var selectorsIndex = GetIndexFrom(igxlSheetsVersion, "Selector");
                var categoryValuesIndex = selectorsIndex + 2;
                var relativeColumnIndex = selectorsIndex + _selectorNameList.Count + CategoryList.Count * 3;
                maxCount = Math.Max(maxCount, relativeColumnIndex + 1);

                #region headers

                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" +
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
                            if (item.indexFrom == item.indexTo && item.rowIndex == i)
                                arr[item.indexFrom] = item.columnName;

                            if (item.columnName == "Selector" && item.rowIndex + 1 == i)
                                for (var index = 0; index < _selectorNameList.Count; index++)
                                {
                                    var selectorName = _selectorNameList[index];
                                    arr[selectorsIndex + index] = selectorName;
                                }

                            if (item.variantName == "CategoryValues")
                            {
                                var categoryCount = 0;
                                foreach (var category in CategoryList)
                                {
                                    foreach (var column in item.Column1)
                                    {
                                        if (column.indexFrom == 5 && item.rowIndex == i)
                                            arr[categoryValuesIndex - 1 + categoryCount * 3 + column.indexFrom - 4] =
                                                category;
                                        if (column.rowIndex == i)
                                            arr[categoryValuesIndex - 1 + categoryCount * 3 + column.indexFrom - 4] =
                                                column.columnName;
                                    }

                                    categoryCount++;
                                }
                            }
                        }

                    #endregion

                    SetRelativeColumn(igxlSheetsVersion, i, arr, relativeColumnIndex);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data

                for (var index = 0; index < AcSpecs.Count; index++)
                {
                    var row = AcSpecs[index];
                    foreach (var selector in row.SelectorList)
                    {
                        var arr = Enumerable.Repeat("", maxCount).ToArray();
                        if (!string.IsNullOrEmpty(row.Symbol))
                        {
                            arr[0] = row.ColumnA;
                            arr[symbolIndex] = row.Symbol;
                            arr[valueIndex] = "=#N/A";

                            for (var i = 0; i < _selectorNameList.Count; i++)
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
                            var categoryCount = 0;
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
        }

        private void WriteSheet3P0(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (AcSpecs.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var symbolIndex = GetIndexFrom(igxlSheetsVersion, "Symbol");
                var valueIndex = GetIndexFrom(igxlSheetsVersion, "Value");
                var selectorsIndex = GetIndexFrom(igxlSheetsVersion, "Selectors");
                if (selectorsIndex == -1) selectorsIndex = GetIndexFrom(igxlSheetsVersion, "Selector");
                var categoryValuesIndex = selectorsIndex + _selectorNameList.Count;
                var relativeColumnIndex = selectorsIndex + _selectorNameList.Count + CategoryList.Count * 3;
                maxCount = Math.Max(maxCount, relativeColumnIndex + 1);

                #region headers

                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" +
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
                            if (item.indexFrom == item.indexTo && item.rowIndex == i)
                                arr[item.indexFrom] = item.columnName;

                            if (item.columnName == "Selectors" && item.rowIndex + 1 == i)
                                for (var index = 0; index < _selectorNameList.Count; index++)
                                {
                                    var selectorName = _selectorNameList[index];
                                    arr[selectorsIndex + index] = selectorName;
                                }

                            if (item.variantName == "CategoryValues")
                            {
                                var categoryCount = 0;
                                foreach (var category in CategoryList)
                                {
                                    foreach (var column in item.Column1)
                                    {
                                        if (column.indexFrom == 1 && item.rowIndex == i)
                                            arr[categoryValuesIndex - 1 + categoryCount * 3 + column.indexFrom] =
                                                category;
                                        if (column.rowIndex == i)
                                            arr[categoryValuesIndex - 1 + categoryCount * 3 + column.indexFrom] =
                                                column.columnName;
                                    }

                                    categoryCount++;
                                }
                            }
                        }

                    #endregion

                    SetRelativeColumn(igxlSheetsVersion, i, arr, relativeColumnIndex);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data

                for (var index = 0; index < AcSpecs.Count; index++)
                {
                    var row = AcSpecs[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.Symbol))
                    {
                        arr[0] = row.ColumnA;
                        arr[symbolIndex] = row.Symbol;
                        arr[valueIndex] = "=#N/A";
                        for (var i = 0; i < _selectorNameList.Count; i++)
                        {
                            var selectorName = row.SelectorList.Find(x =>
                                x.SelectorName.Equals(_selectorNameList[i], StringComparison.CurrentCultureIgnoreCase));
                            if (selectorName != null)
                                arr[selectorsIndex + i] = selectorName.SelectorValue;
                            else
                                arr[selectorsIndex + i] = "";
                        }

                        var categoryCount = 0;
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
        }

        public bool FindValue(string acCategory, string acSelector, ref string frequencyName)
        {
            if (Regex.IsMatch(frequencyName, "^_"))
                frequencyName = Regex.Replace(frequencyName, "^_", "");
            foreach (var acSpec in AcSpecs)
                if (acSpec.Symbol.Equals(frequencyName, StringComparison.CurrentCultureIgnoreCase))
                    if (acSpec.CategoryList.Exists(x =>
                            x.Name.Equals(acCategory, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var row = acSpec.CategoryList.Find(x =>
                            x.Name.Equals(acCategory, StringComparison.CurrentCultureIgnoreCase));
                        if (acSelector.Equals("Typ", StringComparison.CurrentCultureIgnoreCase))
                        {
                            frequencyName = row.Typ;
                            return true;
                        }

                        if (acSelector.Equals("Max", StringComparison.CurrentCultureIgnoreCase))
                        {
                            frequencyName = row.Max;
                            return true;
                        }

                        if (acSelector.Equals("Min", StringComparison.CurrentCultureIgnoreCase))
                        {
                            frequencyName = row.Min;
                            return true;
                        }
                    }

            return false;
        }

        public AcSpec AddAcSpecs(string symbol, string value, string typ, string min, string max)
        {
            //Write basic data
            var acSpecs = new AcSpec(symbol, GetSelectorList(), value);
            //Write Category
            foreach (var category in CategoryList)
            {
                var categoryInSpec = new CategoryInSpec(category, typ, min, max);
                acSpecs.AddCategory(categoryInSpec);
            }

            AddRow(acSpecs);
            return acSpecs;
        }

        private List<Selector> GetSelectorList()
        {
            var selectorList = new List<Selector>();
            selectorList.Add(new Selector("Typ", "Typ"));
            selectorList.Add(new Selector("Min", "Min"));
            selectorList.Add(new Selector("Max", "Max"));
            return selectorList;
        }

        public bool IsSymbolExist(string name)
        {
            foreach (var acSpecs in AcSpecs)
            {
                if (acSpecs.Symbol.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    return true;
            }
            return false;
        }

        #endregion
    }
}