using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CommonLib.Utility;
using OfficeOpenXml;

namespace CommonLib.EpplusErrorReport
{
    public class ErrorInstance
    {
        private readonly List<Error> _errors;

        private ErrorInstance()
        {
            _errors = new List<Error>();
        }

        public static ErrorInstance Instance = new ErrorInstance();

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

        public int GetErrorCountByType(string type)
        {
            return GetErrorsByType(type).Count;
        }

        private List<Error> GetErrorsByType(string type)
        {
            List<Error> targetList = _errors.Where(a => a.ErrorType.Equals(type, StringComparison.CurrentCulture)).ToList();
            return targetList;
        }

        private List<string> GetErrorTypeList()
        {
            List<string> typeList = _errors.GroupBy(p => p.ErrorType).Select(p => p.Key).ToList();
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

                List<ExcelWorkbook> workbooks = new List<ExcelWorkbook> { package.Workbook };
                if (GetErrorCount() > 0)
                {
                    List<string> typeList = GetErrorTypeList();
                    foreach (var type in typeList)
                    {
                        List<Error> errorList = GetErrorsByType(type);
                        ErrorReport errorReport = new ErrorReport(errorList);
                        errorReport.WriteReport(workbooks);
                    }
                }

                package.Save();
            }
        }

        public void GenErrorReport(ExcelPackage excelPackage, List<string> copyFiles)
        {
            ExcelWorkbook workbook = excelPackage.Workbook;
            workbook.CopyWorkSheets(copyFiles);

            List<ExcelWorkbook> workbooks = new List<ExcelWorkbook> { excelPackage.Workbook };
            if (GetErrorCount() > 0)
            {
                List<string> typeList = GetErrorTypeList();
                foreach (var type in typeList)
                {
                    List<Error> errorList = GetErrorsByType(type);
                    ErrorReport errorReport = new ErrorReport(errorList);
                    errorReport.WriteReport(workbooks);
                }
            }
        }

        public void GenErrorReport(ExcelPackage excelPackage, List<string> copyFiles, string errorReprortName, string summaryReport = "SummaryReport")
        {
            ExcelWorkbook workbook = excelPackage.Workbook;
            workbook.CopyWorkSheets(copyFiles);

            List<ExcelWorkbook> workbooks = new List<ExcelWorkbook> { excelPackage.Workbook };
            if (GetErrorCount() > 0)
            {
                List<Error> errorList = GetErrorList();
                ErrorReport errorReport = new ErrorReport(errorList);
                errorReport.WriteReport(workbooks, errorReprortName, summaryReport);
            }
        }
    }
}