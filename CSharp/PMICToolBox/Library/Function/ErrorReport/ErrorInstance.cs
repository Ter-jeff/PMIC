using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Library.Function.ErrorReport
{
    public class ErrorInstance
    {
        private readonly List<Error> _errors;

        private ErrorInstance()
        {
            _errors = new List<Error>();
        }

        public static ErrorInstance Instance { get; } = new ErrorInstance();

        public List<Error> GetErrorList()
        {
            return _errors;
        }

        public void AddError(Error error)
        {
            _errors.Add(error);
        }

        public void Reset()
        {
            _errors.Clear();
        }

        public int GetErrorCount()
        {
            return _errors.Count;
        }

        public int GetErrorCountByType(Type type)
        {
            return GetErrorsByType(type).Count;
        }

        private List<Error> GetErrorsByType(Type type)
        {
            List<Error> targetList = _errors.Where(a => a.ErrorType.GetType() == type).ToList();

            return targetList;
        }

        private List<Type> GetErrorTypeList()
        {
            List<Type> typeList = _errors.GroupBy(p => p.ErrorType.GetType()).Select(p => p.Key).ToList();
            return typeList;
        }

        public void GenErrorReport(string outputFile, List<string> copyFiles)
        {
            if (File.Exists(outputFile))
            {
                File.Delete(outputFile);
            }

            using (ExcelPackage package = new ExcelPackage(new FileInfo(outputFile)))
            {
                ExcelWorkbook workbook = package.Workbook;
                workbook.CopyWorkSheets(copyFiles);

                List<ExcelWorkbook> workbooks = new List<ExcelWorkbook> {package.Workbook};
                if (GetErrorCount() > 0)
                {
                    List<Type> typeList = GetErrorTypeList();
                    foreach (Type type in typeList)
                    {
                        List<Error> errorList = GetErrorsByType(type);
                        ErrorReport errorReport = new ErrorReport(errorList);
                        errorReport.WriteReport(workbooks);
                    }
                }

                package.Save();
            }
        }
    }
}