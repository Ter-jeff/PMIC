using System.Collections.Generic;
using System.Linq;

namespace IgxlData.IgxlBase
{
    public class AcSpec : Spec
    {
        public List<CategoryInSpec> CategoryList { get; set; }

        public List<Selector> SelectorList { get; set; }

        public AcSpec()
        {
            CategoryList = new List<CategoryInSpec>();
            SelectorList = new List<Selector>();
        }

        public AcSpec(string acSpecSymbol, List<Selector> selectorList, string value = "", string comment = "")
            : base(acSpecSymbol, value, comment)
        {
            CategoryList = new List<CategoryInSpec>();
            SelectorList = selectorList.Select(p => new Selector(p.SelectorName, p.SelectorValue)).ToList();
        }

        public void InsertCategory(int index, CategoryInSpec categoryInSpec)
        {
            CategoryList.Insert(index, categoryInSpec);
        }

        public void AddCategory(CategoryInSpec categoryInSpec)
        {
            CategoryList.Add(categoryInSpec);
        }

        public bool ContainsCategory(string categoryName)
        {
            return CategoryList.Exists(p => p.Name.Equals(categoryName));
        }

        public CategoryInSpec GetCategoryItem(string categoryName)
        {
            foreach (var categoryInSpec in CategoryList)
                if (categoryInSpec.Name.Equals(categoryName))
                    return categoryInSpec;
            return null;
        }
    }
}