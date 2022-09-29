using CommonLib.Enum;
using CommonLib.ErrorReport;
using System.Collections.Generic;
using System.Linq;

namespace CommonReaderLib
{
    public abstract class MySheet
    {
        public List<Error> Errors = new List<Error>();
        public string SheetName;

        public void AddError(EnumErrorType errorType, EnumErrorLevel errorLevel, string sheetName, int rowNum, int colNum,
            string message, params string[] comments)
        {
            var error = new Error
            {
                EnumErrorType = errorType,
                ErrorLevel = errorLevel,
                SheetName = sheetName,
                RowNum = rowNum,
                ColNum = colNum,
                Message = message,
                Comments = comments.ToList()
            };
            Errors.Add(error);
        }

        public void AddDimensionError()
        {
            AddError(EnumErrorType.FormatError, EnumErrorLevel.Error, SheetName, 0, 0, "No data in this sheet !!!");
        }

        public void AddFirstHeaderError(string firstHeader)
        {
            AddError(EnumErrorType.FormatError, EnumErrorLevel.Error, SheetName, 0, 0,
                string.Format("Can't find first header {0}!!!", firstHeader));
        }
    }
}