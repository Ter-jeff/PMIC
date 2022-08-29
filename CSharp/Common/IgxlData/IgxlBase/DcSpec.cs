using System;
using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class DcSpec : Spec
    {
        #region Field

        #endregion

        #region Property

        public List<CategoryInSpec> CategoryList { get; set; }

        public List<Selector> SelectorList { get; set; }

        public string SpecialComment { get; set; }

        public int CategoryCount
        {
            get { return CategoryList.Count; }
            set { throw new NotImplementedException(); }
        }

        #endregion

        #region Constructor

        public DcSpec()
        {
            CategoryList = new List<CategoryInSpec>();
        }

        public DcSpec(string dcSpecSymbol, string value = "", string comment = "")
            : base(dcSpecSymbol, value, comment)
        {
            CategoryList = new List<CategoryInSpec>();
            SelectorList = new List<Selector>();
        }

        public DcSpec(string dcSpecSymbol, List<Selector> selectorList, string value = "", string comment = "")
            : base(dcSpecSymbol, value, comment)
        {
            CategoryList = new List<CategoryInSpec>();
            SelectorList = selectorList;
        }

        #endregion

        #region Member Function

        public void AddCategory(CategoryInSpec categoryInSpec)
        {
            CategoryList.Add(categoryInSpec);
        }

        public bool ContainsCategory(string categoryName)
        {
            return CategoryList.Exists(p => p.Name.Equals(categoryName, StringComparison.OrdinalIgnoreCase));
        }

        public void SetCategory(string categoryName, CategoryInSpec categoryInSpec)
        {
            for (var i = 0; i < CategoryList.Count; i++)
                if (CategoryList[i].Name.Equals(categoryName, StringComparison.OrdinalIgnoreCase))
                {
                    CategoryList[i] = categoryInSpec;
                    break;
                }
        }

        public void InsertCategory(int index, CategoryInSpec categoryInSpec)
        {
            CategoryList.Insert(index, categoryInSpec);
        }

        #endregion
    }
}