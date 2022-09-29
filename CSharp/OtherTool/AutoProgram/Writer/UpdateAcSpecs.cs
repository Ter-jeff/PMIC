using AutoProgram.Base;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.Others.MultiTimeSet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace AutoProgram.Writer
{
    public class UpdateAcSpecs
    {
        private const string AcSpecDefault = "-1";

        public AcSpecSheet Work(AcSpecSheet acSpecSheet, List<ComTimeSetBasicSheet> comTimeSetBasicSheets)
        {
            AcSpecSheetUpdateEquation(acSpecSheet, comTimeSetBasicSheets);

            //AddComment(acSpecSheet);

            return acSpecSheet;
        }

        private void AcSpecSheetUpdateEquation(AcSpecSheet acSpecSheet, List<ComTimeSetBasicSheet> comTimeSetBasicSheets)
        {
            if (!comTimeSetBasicSheets.Any()) return;

            foreach (var timeSetBasicSheet in comTimeSetBasicSheets)
            {
                var dic = new Dictionary<string, double>();
                foreach (var tsetEqnVarMap in timeSetBasicSheet.AllTsetEqnVariable)
                    foreach (var variable in tsetEqnVarMap.DictVariable)
                        if (!dic.ContainsKey(variable.Key.ToUpper()))
                            dic.Add(variable.Key.ToUpper(), variable.Value);

                //var blockTypeName = GetTimeSetBlockType(timeSetBasicSheet.SheetName);
                ////TimeSet Sheet default block name, ex: SocMbist/GfxScan
                //var categoryName = GetTimeSetCategory(timeSetBasicSheet.SheetName);
                var tokens = timeSetBasicSheet.SheetName.Split(new[] { '_' },
                    StringSplitOptions.RemoveEmptyEntries).ToList();
                var categoryName = "AC_" + string.Join("_", tokens.GetRange(1, tokens.Count - 1));

                if (dic.Count > 0)
                    UpdateSymbolSpecs(acSpecSheet, dic);
                if (!acSpecSheet.CategoryList.Exists(x =>
                        x.Equals(categoryName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    AddAcCategory(acSpecSheet, dic, categoryName);
                    if (acSpecSheet.CategoryTimeSetDic.ContainsKey(categoryName))
                        acSpecSheet.CategoryTimeSetDic[categoryName].Add(timeSetBasicSheet.SheetName);
                    else
                        acSpecSheet.CategoryTimeSetDic.Add(categoryName,
                            new List<string> { timeSetBasicSheet.SheetName });
                }
            }
        }

        protected BlockType GetTimeSetBlockType(string sheetName)
        {
            var blockType = BlockType.HardIp;
            var tokens = sheetName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
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

        private void AddAcCategory(AcSpecSheet acSpecSheet, Dictionary<string, double> varDic, string categoryName
            , string copyCategoryName = "")
        {
            if (acSpecSheet.CategoryList.Contains(categoryName))
                return;

            acSpecSheet.CategoryList.Add(categoryName);
            foreach (var specs in acSpecSheet.AcSpecs)
            {
                var categoryInSpec = specs.CategoryList.First();
                if (!string.IsNullOrEmpty(copyCategoryName))
                    if (specs.CategoryList.Exists(x =>
                            x.Name.Equals(copyCategoryName, StringComparison.CurrentCultureIgnoreCase)))
                        categoryInSpec = specs.CategoryList.Find(x =>
                            x.Name.Equals(copyCategoryName, StringComparison.CurrentCultureIgnoreCase));
                var value = !varDic.ContainsKey(specs.Symbol.ToUpper())
                    ? categoryInSpec.Typ
                    : varDic[specs.Symbol.ToUpper()].ToString(CultureInfo.InvariantCulture);
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
            var selectorList = new List<Selector>();
            selectorList.Add(new Selector("", ""));
            acSpecSheet.AddRow(new AcSpec("", selectorList));
            acSpecSheet.AddRow(new AcSpec("", selectorList));
            var commentRow = new AcSpec("Using TSet", selectorList);
            foreach (var category in acSpecSheet.CategoryList)
            {
                var timeSets = "";
                if (acSpecSheet.CategoryTimeSetDic.ContainsKey(category))
                    timeSets = string.Join(",", acSpecSheet.CategoryTimeSetDic[category].Distinct());
                var categoryInSpec = new CategoryInSpec(category, timeSets, "", "");
                commentRow.AddCategory(categoryInSpec);
            }

            acSpecSheet.AddRow(commentRow);
        }

        protected string GetTimeSetCategory(string sheetName)
        {
            var tokens = sheetName.Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (tokens.Count >= 5)
            {
                if (tokens[3].Equals("SC", StringComparison.CurrentCultureIgnoreCase))
                    return Category.Scan;
                if (tokens[3].Equals("BI", StringComparison.CurrentCultureIgnoreCase))
                    return Category.Mbist;
                if (tokens[3].Equals("JT", StringComparison.CurrentCultureIgnoreCase))
                    return Category.Common;
                if (tokens[3].Equals("IO", StringComparison.CurrentCultureIgnoreCase))
                    return Category.BScan;
            }
            return "AC_" + string.Join("_", tokens.GetRange(1, tokens.Count - 1));
        }
    }
}