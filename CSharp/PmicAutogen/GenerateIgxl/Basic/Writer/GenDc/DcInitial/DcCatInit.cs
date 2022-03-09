using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDc.DcInitial
{
    public class DcCatInit
    {
        public List<DcCategory> InitFlow(PowerOverWriteSheet powerOverWriteSheet, IoLevelsSheet ioLevelsSheet)
        {
            var categoryList = new List<DcCategory>();
            categoryList.Add(new DcCategory(Category.Common, "", Category.Common, DcCategoryType.Default));
            categoryList.Add(new DcCategory(Category.Ids, "", Category.Ids, DcCategoryType.Ids));
            categoryList.Add(new DcCategory(Category.Conti, "", Category.Conti, DcCategoryType.Conti));
            categoryList.Add(new DcCategory(Category.Leakage, "", Category.Leakage, DcCategoryType.Default));
            categoryList.Add(new DcCategory(Category.Nwire, "", Category.Nwire, DcCategoryType.Default));
            categoryList.Add(new DcCategory(Category.Scan, "", Category.Scan, DcCategoryType.Default));
            categoryList.Add(new DcCategory(Category.BScan, "", Category.BScan, DcCategoryType.Default));
            if (ioLevelsSheet != null)
                categoryList.AddRange(ioLevelsSheet.GenBScanDcCategory(""));
            categoryList.Add(new DcCategory(Category.Mbist, "", Category.Mbist, DcCategoryType.Default));
            categoryList.Add(new DcCategory(Category.Analog, "", Category.Analog, DcCategoryType.Default));

            if (ioLevelsSheet != null)
                ioLevelsSheet.UpdateDcCategory(categoryList, "");

            if (powerOverWriteSheet != null)
                foreach (var catDef in powerOverWriteSheet.PowerOverWrite)
                    if (!categoryList.Exists(x =>
                        x.CategoryName.Equals(catDef.CategoryName, StringComparison.OrdinalIgnoreCase)))
                    {
                        categoryList.Add(new DcCategory(catDef.CategoryName, "", catDef.CategoryName,
                            DcCategoryType.Pmic));
                    }
                    else
                    {
                        var index = categoryList.FindIndex(x =>
                            x.CategoryName.Equals(catDef.CategoryName, StringComparison.OrdinalIgnoreCase));
                        categoryList[index].Type = DcCategoryType.Pmic;
                    }

            //return GenExtraDcCategoryList(categoryList);
            return categoryList;
        }

        //private List<DcCategory> GenExtraDcCategoryList(List<DcCategory> dcCategoryList)
        //{
        //    var group = StaticTestPlan.VddLevelsSheet.Rows.SelectMany(x => x.ExtraSelectors.Keys).Distinct().ToList();
        //    var newDcCategoryList = new List<DcCategory>();
        //    newDcCategoryList.AddRange(dcCategoryList);
        //    foreach (var item in group)
        //    foreach (var category in dcCategoryList)
        //    {
        //        var newDcCategory = new DcCategory(category.CategoryName + "_" + item, category.Block,
        //            category.SubCategory, category.Type, category.DcSpecSheet);
        //        newDcCategoryList.Add(newDcCategory);
        //    }

        //    return newDcCategoryList;
        //}
    }
}