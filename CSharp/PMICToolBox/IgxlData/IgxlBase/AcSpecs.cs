using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public class AcSpecs : Spec
    {
        #region Field
        private List<CategoryInSpec> _categoryList;
        private List<Selector> _selectorList;
        #endregion

        #region Property
        public List<CategoryInSpec> CategoryList
        {
            get { return _categoryList; }
        }
        public List<Selector> SelectorList
        {
            get { return _selectorList; }
        }
        #endregion

        #region Constructor

        public AcSpecs()
        {
            _categoryList = new List<CategoryInSpec>();
            _selectorList = new List<Selector>();
        }

        public AcSpecs(string acSpecSymbol, List<Selector> selectorList, string value = "", string comment = "")
            : base(acSpecSymbol, value, comment)
        {
            _categoryList = new List<CategoryInSpec>();
            _selectorList = selectorList.Select(p => new Selector(p.SelectorName, p.SelectorValue)).ToList();
        }
        #endregion

        #region Member function
        public void InsertCategory(int index, CategoryInSpec categroyItem)
        {
            _categoryList.Insert(index, categroyItem);
        }

        public void AddCategory(CategoryInSpec categroyItem)
        {
            _categoryList.Add(categroyItem);
        }

        public bool ContainsCategory(string categoryName)
        {
            return _categoryList.Exists(p => { return p.Name.Equals(categoryName); });
        }

        public CategoryInSpec GetCategoryItem(string categoryName)
        {
            foreach (CategoryInSpec catgoryItem in _categoryList)
            {
                if (catgoryItem.Name.Equals(categoryName))
                {
                    return catgoryItem;
                }
            }
            return null;
        }

        public AcSpecs DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as AcSpecs;
            }
        }
        #endregion

    }
}