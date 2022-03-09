using System;
using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class DcSpecs : Spec
    {
        #region Field
        private List<CategoryInSpec> _categoryList;
        private List<Selector> _selectorList;
        #endregion

        #region Property
        public List<CategoryInSpec> CategoryList
        {
            get { return _categoryList; }
            set { _categoryList = value; }
        }

        public List<Selector> SelectorList
        {
            get { return _selectorList; }
            set { _selectorList = value; }
        }

        public string SpecialComment { get; set; }

        public int CategoryCount
        {
            get { return _categoryList.Count; }
        }
        #endregion

        #region Constructor
        public DcSpecs()
            : base()
        {
            _categoryList = new List<CategoryInSpec>();
        }

        public DcSpecs(string dcSpecSymbol, string value = "", string comment = "")
            : base(dcSpecSymbol, value, comment)
        {
            _categoryList = new List<CategoryInSpec>();
            _selectorList = new List<Selector>();
        }

        public DcSpecs(string dcSpecSymbol, List<Selector> selectorList, string value = "", string comment = "")
            : base(dcSpecSymbol, value, comment)
        {
            _categoryList = new List<CategoryInSpec>();
            _selectorList = selectorList;
        }
        #endregion

        #region  Member Function

        public void AddCategory(CategoryInSpec categroyItem)
        {
            _categoryList.Add(categroyItem);
        }

        /// <summary>
        /// Judge whether already contains given category
        /// </summary>
        /// <param name="categoryName"></param>
        /// <returns></returns>
        public bool ContainsCategory(string categoryName)
        {
            return _categoryList.Exists(p => { return p.Name.Equals(categoryName, StringComparison.OrdinalIgnoreCase); });
        }

        public void SetCategory(string categoryName, CategoryInSpec categroyItem)
        {
            for (int i = 0; i < _categoryList.Count; i++)
            {
                if (_categoryList[i].Name.Equals(categoryName, StringComparison.OrdinalIgnoreCase))
                {
                    _categoryList[i] = categroyItem;
                    break;
                }
            }
        }

        public void InsertCategory(int index, CategoryInSpec categroyItem)
        {
            _categoryList.Insert(index, categroyItem);
        }

        public string GetDCvalue(string dcSpec)
        {
            if (_categoryList.Exists(x => x.Name.Equals(dcSpec, StringComparison.OrdinalIgnoreCase)))
            {
                int index = _categoryList.FindIndex(x => x.Name.Equals(dcSpec, StringComparison.OrdinalIgnoreCase));
                //if (GetDcSpecsData().Exists(x => x.SelectorList.Exists(y => y.Equals(selector, StringComparison.OrdinalIgnoreCase))))
                //{
                //    var pin = GetDcSpecsData().First(x => x.Name.StartsWith(pinName, StringComparison.OrdinalIgnoreCase)).SelectorList.First(y => y.Name.Equals(selector, StringComparison.OrdinalIgnoreCase));
                //    if (selector.Equals("Max", StringComparison.OrdinalIgnoreCase))
                //        value = pin.CategoryValues[index].Max;
                //    else if (selector.Equals("Typ", StringComparison.OrdinalIgnoreCase))
                //        value = pin.CategoryValues[index].Typ;
                //    else if (selector.Equals("Min", StringComparison.OrdinalIgnoreCase))
                //        value = pin.CategoryValues[index].Min;
                //    return value;
                //}
            }
            return "";
        }
        #endregion

    }
}