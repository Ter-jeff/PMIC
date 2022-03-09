using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlSheets
{
    public class AcSpecSheet : IgxlSheet
    {
        private const string SheetType = "DTACSpecSheet";

        #region Field
        private List<AcSpecs> _acSpecs;
        private List<string> _categoryList;
        private readonly List<string> _selectorNameList;
        #endregion

        #region Property
        public List<string> CategoryList
        {
            get { return _categoryList; }
            set { _categoryList = value; }
        }

        public List<AcSpecs> AcSpecs
        {
            get { return _acSpecs; }
            set { _acSpecs = value; }
        }
        #endregion

        #region Constructor
        public AcSpecSheet(ExcelWorksheet sheet, List<string> categoryList, List<string> selectorNameList)
            : base(sheet)
        {
            _acSpecs = new List<AcSpecs>();
            _categoryList = categoryList;
            _selectorNameList = selectorNameList;
            IgxlSheetName = IgxlSheetNameList.AcSpec;
        }

        public AcSpecSheet(string sheetName, List<string> categoryList, List<string> selectorNameList)
            : base(sheetName)
        {
            _acSpecs = new List<AcSpecs>();
            _categoryList = categoryList;
            _selectorNameList = selectorNameList;
            IgxlSheetName = IgxlSheetNameList.AcSpec;
        }
        #endregion

        #region Member Function
        public List<Selector> CreateSelectorList()
        {
            var selectorList = new List<Selector>();
            return selectorList;
        }

        public AcSpecs GetAcSpecs(string name)
        {
            return _acSpecs.Find(x => x.Symbol.Equals(name, StringComparison.CurrentCulture));

        }

        public void AddRow(AcSpecs acSpecRow)
        {
            _acSpecs.Add(acSpecRow);
        }

        public override void Write(string fileName, string version = "3.0")
        {
            //if (version == "2.0")
            //    WriteSheet2P0(fileName);
            //else if (version == "3.0")
            //    WriteSheet3P0(fileName);
            //else
            //    throw new Exception(string.Format("The AC Spec sheet version:{0} is not supported!", version));

            //Support 3.0 & 2.0
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                Directory.CreateDirectory(Path.GetDirectoryName(fileName));
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version == "2.0")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet2P0(fileName, version, igxlSheetsVersion);
                }
                else if (dic.ContainsKey(version))
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (dic.ContainsKey("3.0"))
                {
                    var igxlSheetsVersion = dic["3.0"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        private void WriteSheet2P0(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_acSpecs.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var symbolIndex = GetIndexFrom(igxlSheetsVersion, "Symbol");
                var valueIndex = GetIndexFrom(igxlSheetsVersion, "Value");
                var selectorsIndex = GetIndexFrom(igxlSheetsVersion, "Selector");
                int categoryValuesIndex = selectorsIndex + 2;
                int relativeColumnIndex = selectorsIndex + _selectorNameList.Count + _acSpecs.First().CategoryList.Count * 3;
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
                                foreach (var category in _categoryList)
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
                for (var index = 0; index < _acSpecs.Count; index++)
                {
                    var row = _acSpecs[index];
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
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_acSpecs.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var symbolIndex = GetIndexFrom(igxlSheetsVersion, "Symbol");
                var valueIndex = GetIndexFrom(igxlSheetsVersion, "Value");
                var selectorsIndex = GetIndexFrom(igxlSheetsVersion, "Selectors");
                if (selectorsIndex == -1) selectorsIndex = GetIndexFrom(igxlSheetsVersion, "Selector");
                int categoryValuesIndex = selectorsIndex + _selectorNameList.Count;
                int relativeColumnIndex = selectorsIndex + _selectorNameList.Count + _acSpecs.First().CategoryList.Count * 3;
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

                            if (item.columnName == "Selectors" && item.rowIndex + 1 == i)
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
                                foreach (var category in _categoryList)
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
                for (var index = 0; index < _acSpecs.Count; index++)
                {
                    var row = _acSpecs[index];
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
        }

        //private void WriteSheet2P0(string fileName)
        //{
        //    var validate = new Action<string>((a) => { });
        //    if (_categoryList.Count > 0)
        //    {
        //        var acSpecSheetGen = new GenACSpecSheet(fileName, validate, true, _categoryList);
        //        foreach (var acSpecs in _acSpecs)
        //        {
        //            var typValueList = new List<string>();
        //            var minValueList = new List<string>();
        //            var maxValueList = new List<string>();
        //            foreach (var categroyItem in acSpecs.CategoryList)
        //            {
        //                typValueList.Add(categroyItem.Typ);
        //                minValueList.Add(categroyItem.Min);
        //                maxValueList.Add(categroyItem.Max);
        //            }
        //            if (acSpecs.Symbol == "")
        //            {
        //                acSpecSheetGen.AddBlankLine();
        //            }
        //            else
        //            {
        //                foreach (var selecor in acSpecs.SelectorList)
        //                {
        //                    acSpecSheetGen.AddSpecSelectorValues(acSpecs.Symbol, selecor.SelectorName,
        //                       selecor.SelectorValue, typValueList.ToArray(), minValueList.ToArray(),
        //                        maxValueList.ToArray());
        //                }
        //            }
        //        }
        //        acSpecSheetGen.WriteSheet();
        //    }
        //}

        //private void WriteSheet3P0(string fileName)
        //{
        //    var validate = new Action<string>((a) => { });
        //    var acSpecSheetGen = new GenACSpecSheetVer30(fileName, validate, true, _selectorNameList, _categoryList);

        //    foreach (var acSpec in _acSpecs)
        //    {
        //        var typValueList = new List<string>();
        //        var minValueList = new List<string>();
        //        var maxValueList = new List<string>();
        //        foreach (var categroyItem in acSpec.CategoryList)
        //        {
        //            typValueList.Add(categroyItem.Typ);
        //            minValueList.Add(categroyItem.Min);
        //            maxValueList.Add(categroyItem.Max);
        //        }
        //        if (acSpec.Symbol == "")
        //        {
        //            acSpecSheetGen.AddBlankLine();
        //        }
        //        else
        //        {
        //            string[] selectorVal;
        //            if (acSpec.SelectorList.Count >= _selectorNameList.Count)
        //            {
        //                selectorVal = acSpec.SelectorList.Select(p => p.SelectorValue).ToArray();
        //            }
        //            else
        //            {
        //                selectorVal = Enumerable.Repeat("Typ", _selectorNameList.Count).ToArray();
        //            }

        //            acSpecSheetGen.AddSpec(acSpec.Symbol, selectorVal, typValueList.ToArray(), minValueList.ToArray(),
        //                            maxValueList.ToArray());
        //        }
        //    }
        //    acSpecSheetGen.WriteSheet();
        //}

        public bool FindValue(string acCategory, string acSelector, ref string frequencyName)
        {
            if (Regex.IsMatch(frequencyName, "^_"))
                frequencyName = Regex.Replace(frequencyName, "^_", "");
            foreach (var acSpec in AcSpecs)
            {
                if (acSpec.Symbol.Equals(frequencyName, StringComparison.CurrentCultureIgnoreCase))
                {
                    if (acSpec.CategoryList.Exists(x => x.Name.Equals(acCategory, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var row = acSpec.CategoryList.Find(x => x.Name.Equals(acCategory, StringComparison.CurrentCultureIgnoreCase));
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
                }
            }
            return false;
        }

        public bool IsSame(string categoryName, AcSpecSheet targetAcSheet, string targetCategoryName)
        {
            if (!_categoryList.Exists(x => x.Equals(categoryName, StringComparison.CurrentCultureIgnoreCase)))
                return false;
            if (!targetAcSheet.CategoryList.Exists(x => x.Equals(targetCategoryName, StringComparison.CurrentCultureIgnoreCase)))
                return false;

            var pins = AcSpecs.Select(x => x.Symbol).Distinct().ToList();
            foreach (var pin in pins)
            {
                var oldAcSymbol = AcSpecs.FindAll(a => a.Symbol.ToUpper().Equals(pin.ToUpper()));
                var newAcSymbol = targetAcSheet.AcSpecs.FindAll(a => a.Symbol.ToUpper().Equals(pin.ToUpper()));
                foreach (var newAcSymbolData in oldAcSymbol)
                {
                    foreach (var oldAcSymbolData in newAcSymbol)
                    {
                        if (newAcSymbolData.ContainsCategory(categoryName) && oldAcSymbolData.ContainsCategory(targetCategoryName))
                        {
                            var newCategoryItem = newAcSymbolData.GetCategoryItem(categoryName);
                            var oldCategoryItem = oldAcSymbolData.GetCategoryItem(targetCategoryName);
                            if (!newCategoryItem.Min.Equals(oldCategoryItem.Min, StringComparison.CurrentCultureIgnoreCase))
                                return false;
                            if (!newCategoryItem.Max.Equals(oldCategoryItem.Max, StringComparison.CurrentCultureIgnoreCase))
                                return false;
                            if (!newCategoryItem.Typ.Equals(oldCategoryItem.Typ, StringComparison.CurrentCultureIgnoreCase))
                                return false;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
            return true;

        }
        #endregion
    }
}