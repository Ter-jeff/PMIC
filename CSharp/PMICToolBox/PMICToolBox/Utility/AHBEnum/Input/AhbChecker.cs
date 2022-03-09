using Library.Function.ErrorReport;
using System;
using System.Collections.Generic;

namespace PmicAutomation.Utility.AHBEnum.Input
{
    public class AhbChecker
    {
        public void CheckDuplicateAhbEnum(AhbRegisterMapSheet ahbRegisterMapSheet, bool fieldNameType = true)
        {
            Dictionary<string, int> fieldNames = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
            foreach (AhbRegisterMapRow row in ahbRegisterMapSheet.Rows)
            {
                string fieldName =
                    fieldNameType && !row.FieldName.StartsWith(row.RegName, StringComparison.CurrentCulture)
                        ? row.RegName + "_" + row.FieldName
                        : row.FieldName;
                if (fieldNames.ContainsKey(fieldName))
                {
                    string errMsg = "Ahb name of row " + fieldNames[fieldName] + " & " + row.RowNum + " are duplicate - " + fieldName + " !!!";
                    ErrorManager.AddError(AhbErrorType.Duplicated, ahbRegisterMapSheet.Name, row.RowNum,
                        ahbRegisterMapSheet.HeaderIndexDic["field name"], errMsg);
                }
                else
                {
                    fieldNames.Add(fieldName, row.RowNum);
                }
            }
        }
    }
}