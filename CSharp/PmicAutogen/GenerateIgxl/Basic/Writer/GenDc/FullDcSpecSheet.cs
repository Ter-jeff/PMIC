using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.Others;
using PmicAutogen.Inputs.TestPlan.Reader;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenGlobalDc
{
    public class FullDcSpecSheet
    {
        #region Porperty
        public List<DcSpecs> IoDcSpecsList { get; set; }
        public DcSpecSheet DcSpecSheet { get; set; }
        #endregion

        #region Constructor

        public FullDcSpecSheet(string sheetName, List<string> categoryList = null)
        {
            DcSpecSheet = new DcSpecSheet(sheetName, categoryList, GetSelectorStringList());
            IoDcSpecsList = new List<DcSpecs>();
        }

        #endregion

        private List<string> GetSelectorStringList()
        {
            return GetSelectorList().Select(p => p.SelectorName).ToList();
        }

        private List<Selector> GetSelectorList()
        {
            var selectorList = new List<Selector>();
            selectorList.Add(new Selector("Min", "Min"));
            selectorList.Add(new Selector("Typ", "Typ"));
            selectorList.Add(new Selector("Max", "Max"));
            return selectorList;
        }

        public CategoryInSpec GetCategoryItem(string category, string glbSymbol)
        {
            var categoryItem = new CategoryInSpec(category);
            var isFormula = false;
            if (glbSymbol.Contains('+'))
                isFormula = true;
            else if (glbSymbol.Contains('-'))
                isFormula = true;
            else if (glbSymbol.Contains('*'))
                isFormula = true;
            else if (glbSymbol.Contains('/'))
                isFormula = true;
            categoryItem.Typ = isFormula ? "=(" + glbSymbol + ")" : "=" + glbSymbol;
            categoryItem.Max = categoryItem.Typ + "*_IO_Pins_GLB_Plus";
            categoryItem.Min = categoryItem.Typ + "*_IO_Pins_GLB_Minus";
            return categoryItem;
        }

        public void AddDcSpec(DcSpecs dcSpecs)
        {
            dcSpecs.SelectorList = GetSelectorList();
            DcSpecSheet.AddRow(dcSpecs);
        }

        public void AddSpecsRange(List<DcSpecs> dcSpecsList)
        {
            DcSpecSheet.AddRows(dcSpecsList);
        }

        public int FindCategoryIndex(string category)
        {
            if (!DcSpecSheet.CategoryList.Exists(x => x.Equals(category, StringComparison.OrdinalIgnoreCase))
            ) return -1;
            return DcSpecSheet.CategoryList.FindIndex(x => x.Equals(category, StringComparison.OrdinalIgnoreCase));
        }

        public bool ExistDcSpecs(string symbol)
        {
            return DcSpecSheet.ExistDcSpecs(symbol);
        }

        public DcSpecs FindDcSpecs(string symbol)
        {
            return DcSpecSheet.FindDcSpecs(symbol);
        }
    }
}