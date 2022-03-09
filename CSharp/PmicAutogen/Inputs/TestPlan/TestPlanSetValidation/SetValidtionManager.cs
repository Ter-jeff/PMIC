using System.Collections.Generic;
using System.Text.RegularExpressions;
using AutomationCommon.Utility;
using Microsoft.Office.Interop.Excel;

namespace PmicAutogen.Inputs.TestPlan.TestPlanSetValidation
{
    public class SetValidationManager
    {
        private readonly List<EnumColumn> _dynamicTypes = new List<EnumColumn> {EnumColumn.Pin};

        #region Member Function

        public void SetValidations(Workbook workbook, bool isDynamic = false)
        {
            foreach (var sheetConfig in SheetStructureManager.SheetConfigs)
            {
                if (sheetConfig.Type == EnumColumn.None)
                    continue;

                if (isDynamic && !_dynamicTypes.Exists(x => x == sheetConfig.Type))
                    continue;

                var sheetName = sheetConfig.SheetName;
                var worksheets = GetSheet(workbook, sheetName);
                foreach (var worksheet in worksheets)
                    if (worksheet != null)
                        SheetStructureManager.SetInitialColumns(workbook, worksheet, sheetConfig.FirstHeaderName,
                            sheetConfig.HeaderName, sheetConfig.Type);
            }
        }


        public List<Worksheet> GetSheet(Workbook workbook, string name)
        {
            if (name.Contains("*"))
            {
                var worksheets = new List<Worksheet>();
                if (name.EndsWith("*"))
                    foreach (Worksheet worksheet in workbook.Worksheets)
                        if (Regex.IsMatch(worksheet.Name, "^" + name.TrimEnd('*'), RegexOptions.IgnoreCase))
                        {
                            worksheets.Add(worksheet);
                            return worksheets;
                        }
            }
            else
            {
                return new List<Worksheet> {workbook.GetSheet(name)};
            }

            return null;
        }

        #endregion
    }
}