using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.Others;
using IgxlData.Others.MultiTimeSet;
using PmicAutogen.Inputs.PatternList;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenAC
{
    public class AcSpecsMain
    {
        private const string AcSpecDefault = "-1";

        public AcSpecSheet WorkFlow(List<ComTimeSetBasicSheet> comTimeSetBasicSheets)
        {
            var categoryList = InitialAcCatList();
            var selectorNameList = new List<string>();
            selectorNameList.Add("Typ");
            selectorNameList.Add("Min");
            selectorNameList.Add("Max");
            var acSpecSheet = new AcSpecSheet(PmicConst.AcSpecs, categoryList, selectorNameList);

            InitialAcSymbols(acSpecSheet);

            AcSpecSheetUpdateEquation(acSpecSheet, comTimeSetBasicSheets);

            AddComment(acSpecSheet);

            return acSpecSheet;
        }

        protected List<string> InitialAcCatList()
        {
            var categoryList = new List<string>();
            categoryList.Add(Category.Common);
            categoryList.Add(Category.Scan);
            categoryList.Add(Category.Mbist);
            categoryList.Add(Category.BScan);
            categoryList.Add(Category.JTag);
            return categoryList;
        }

        public void InitialAcSymbols(AcSpecSheet acSpecSheet)
        {
            var freq24Mhz = 24e6.ToString(CultureInfo.InvariantCulture);
            acSpecSheet.AddAcSpecs(SpecFormat.GenAcSpecSymbol(TimeSetConst.TckFreqVar), freq24Mhz, freq24Mhz, freq24Mhz,
                freq24Mhz);
            acSpecSheet.AddAcSpecs(SpecFormat.GenAcSpecSymbol(TimeSetConst.ShiftInFreqVar), AcSpecDefault,
                AcSpecDefault, AcSpecDefault, AcSpecDefault);
        }

        private void AcSpecSheetUpdateEquation(AcSpecSheet acSpecSheet,
            List<ComTimeSetBasicSheet> comTimeSetBasicSheets)
        {
            if (!comTimeSetBasicSheets.Any()) return;

            foreach (var timeSetBasicSheet in comTimeSetBasicSheets)
            {
                var dic = new Dictionary<string, double>();
                foreach (var tsetEqnVarMap in timeSetBasicSheet.AllTsetEqnVariable)
                foreach (var variable in tsetEqnVarMap.DictVariable)
                    if (!dic.ContainsKey(variable.Key.ToUpper()))
                        dic.Add(variable.Key.ToUpper(), variable.Value);

                //TimeSet Sheet default block name, ex: SocMbist/GfxScan
                var categoryName = GetTimeSetCategory(timeSetBasicSheet.SheetName);
                //only could be one of BlockType.Mbist/BlockType.Scan/Block.HardIp
                var blockTypeName = GetTimeSetBlockType(timeSetBasicSheet.SheetName);

                if (dic.Count > 0)
                {
                    UpdateSymbolSpecs(acSpecSheet, dic);
                    categoryName = TsetMapCategory(acSpecSheet, dic, categoryName);
                }

                if (!InputFiles.PatternListMap.Contains(timeSetBasicSheet.SheetName))
                    InputFiles.PatternListMap.SetRow(timeSetBasicSheet.SheetName, blockTypeName, categoryName);
            }
        }

        private string TsetMapCategory(AcSpecSheet acSpecSheet, Dictionary<string, double> tSetVarValueDict,
            string categoryName)
        {
            var regex = new Regex(categoryName, RegexOptions.IgnoreCase | RegexOptions.Compiled);
            var targetCategoryList = acSpecSheet.CategoryList.FindAll(regex.IsMatch);
            var checkValueDic = new Dictionary<string, Dictionary<string, bool>>();
            foreach (var targetCategory in targetCategoryList)
            {
                var dic = new Dictionary<string, bool>();
                foreach (var tset in tSetVarValueDict)
                {
                    var acSpecsList = acSpecSheet.AcSpecs.FindAll(a => a.Symbol.ToUpper().Equals(tset.Key.ToUpper()));

                    foreach (var acSpecs in acSpecsList)
                        if (acSpecs.ContainsCategory(targetCategory))
                        {
                            var categoryItem = acSpecs.GetCategoryItem(targetCategory);
                            if (categoryItem.Min.Equals(tset.Value.ToString(CultureInfo.InvariantCulture)))
                            {
                                if (!dic.ContainsKey(tset.Key))
                                    dic.Add(tset.Key, true);
                            }
                            else if (categoryItem.Min == AcSpecDefault) // update value
                            {
                                categoryItem.Min = tset.Value.ToString(CultureInfo.InvariantCulture);
                                categoryItem.Typ = tset.Value.ToString(CultureInfo.InvariantCulture);
                                categoryItem.Max = tset.Value.ToString(CultureInfo.InvariantCulture);
                                if (!dic.ContainsKey(tset.Key))
                                    dic.Add(tset.Key, true);
                            }
                            else if (!categoryItem.Min.Equals(tset.Value.ToString(CultureInfo.InvariantCulture)))
                            {
                                if (!dic.ContainsKey(tset.Key))
                                    dic.Add(tset.Key, false);
                            }
                        }
                }

                checkValueDic.Add(targetCategory, dic);
            }

            string newCategoryName;
            foreach (var item in checkValueDic)
                if (item.Value.All(p => p.Value)) // all true no need create category
                {
                    newCategoryName = item.Key;
                    return newCategoryName;
                }

            newCategoryName = categoryName + "_" + targetCategoryList.Count;
            AddAcCategory(acSpecSheet, tSetVarValueDict, newCategoryName);

            return newCategoryName;
        }

        private void AddAcCategory(AcSpecSheet acSpecSheet, Dictionary<string, double> tSetVarValueDict,
            string categoryName)
        {
            if (acSpecSheet.CategoryList.Contains(categoryName))
                return;

            acSpecSheet.CategoryList.Add(categoryName);
            foreach (var specs in acSpecSheet.AcSpecs)
            {
                var value = !tSetVarValueDict.ContainsKey(specs.Symbol.ToUpper())
                    ? specs.CategoryList.First().Typ
                    : tSetVarValueDict[specs.Symbol.ToUpper()].ToString(CultureInfo.InvariantCulture);
                var newCategory = new CategoryInSpec(categoryName, value, value, value);
                specs.AddCategory(newCategory);
            }
        }

        private void UpdateSymbolSpecs(AcSpecSheet acSpecSheet, Dictionary<string, double> tSetVarValueDict)
        {
            foreach (var timing in tSetVarValueDict)
            {
                if (acSpecSheet.AcSpecs.Exists(x =>
                    x.Symbol.Equals(timing.Key, StringComparison.CurrentCultureIgnoreCase)))
                    continue;
                acSpecSheet.AddAcSpecs(timing.Key, AcSpecDefault, AcSpecDefault, AcSpecDefault, AcSpecDefault);
            }
        }

        private void AddComment(AcSpecSheet acSpecSheet)
        {
            var categoryTimingSheet = new List<Tuple<string, string>>();
            foreach (var category in acSpecSheet.CategoryList)
            {
                var timing = InputFiles.PatternListMap.GetCategoryUsageTsetSheetName(category);
                categoryTimingSheet.Add(new Tuple<string, string>(category, timing));
            }

            var selectorList = new List<Selector>();
            selectorList.Add(new Selector("", ""));
            acSpecSheet.AddRow(new AcSpecs("", selectorList));
            acSpecSheet.AddRow(new AcSpecs("", selectorList));
            var commentRow = new AcSpecs("Using TSet", selectorList);
            foreach (var timing in categoryTimingSheet)
            {
                var item = new CategoryInSpec(timing.Item1, timing.Item2, "", "");
                commentRow.AddCategory(item);
            }

            acSpecSheet.AddRow(commentRow);
        }

        protected string GetTimeSetCategory(string sheetName)
        {
            var blockName = "TBD";
            var tokens = sheetName.Split(new[] {'_'}, StringSplitOptions.RemoveEmptyEntries);
            var subBlock = Category.Common; //Scan/Mbist/Other/HardIp/Common
            if (tokens.Length >= 5)
            {
                if (tokens[3].Equals("SC", StringComparison.CurrentCultureIgnoreCase))
                    subBlock = Category.Scan;
                else if (tokens[3].Equals("BI", StringComparison.CurrentCultureIgnoreCase))
                    subBlock = Category.Mbist;
                else if (tokens[3].Equals("JT", StringComparison.CurrentCultureIgnoreCase))
                    subBlock = Category.Common;
                else if (tokens[3].Equals("IO", StringComparison.CurrentCultureIgnoreCase))
                    subBlock = Category.BScan;
                blockName = subBlock;
            }

            return blockName;
        }

        protected BlockType GetTimeSetBlockType(string sheetName)
        {
            var blockType = BlockType.HardIp;
            var tokens = sheetName.Split(new[] {'_'}, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length >= 5)
            {
                if (tokens[3].Equals("SC", StringComparison.CurrentCultureIgnoreCase))
                    blockType = BlockType.Scan;
                else if (tokens[3].Equals("BI", StringComparison.CurrentCultureIgnoreCase))
                    blockType = BlockType.Mbist;
                else if (tokens[3].Equals("JT", StringComparison.CurrentCultureIgnoreCase))
                    blockType = BlockType.Common;
                else if (tokens[3].Equals("IO", StringComparison.CurrentCultureIgnoreCase))
                    blockType = BlockType.BScan;
            }

            return blockType;
        }
    }
}