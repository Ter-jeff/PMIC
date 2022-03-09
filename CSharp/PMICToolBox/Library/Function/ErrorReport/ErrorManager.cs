using System;
using System.Collections.Generic;
using System.Linq;

namespace Library.Function.ErrorReport
{
    public static class ErrorManager
    {
        private static readonly ErrorInstance ErrorInstance = ErrorInstance.Instance;

        public static void AddError(object errorType, string name, int rowNum, string message,
            params string[] comments)
        {
            Error errorNew = new Error
            {
                ErrorType = errorType,
                SheetName = name,
                RowNum = rowNum,
                Comments = comments.ToList(),
                Message = message,
                ErrorLevel = ErrorLevel.Error
            };
            AddError(errorNew);
        }

        public static void AddError(object errorType, string name, int rowNum, int colNum, string message,
            params string[] comments)
        {
            Error errorNew = new Error
            {
                ErrorType = errorType,
                SheetName = name,
                RowNum = rowNum,
                ColNum = colNum,
                Comments = comments.ToList(),
                Message = message,
                ErrorLevel = ErrorLevel.Error
            };
            AddError(errorNew);
        }

        public static void AddError(object errorType, ErrorLevel errorLevel, string name, int rowNum,
            string message, params string[] comments)
        {
            Error errorNew = new Error
            {
                ErrorType = errorType,
                SheetName = name,
                RowNum = rowNum,
                Comments = comments.ToList(),
                Message = message,
                ErrorLevel = errorLevel
            };
            AddError(errorNew);
        }

        public static void AddError(object errorType, ErrorLevel errorLevel, string name, int rowNum, int colNum,
            string message, params string[] comments)
        {
            Error errorNew = new Error
            {
                ErrorType = errorType,
                SheetName = name,
                RowNum = rowNum,
                ColNum = colNum,
                Comments = comments.ToList(),
                Message = message,
                ErrorLevel = errorLevel
            };
            AddError(errorNew);
        }

        private static void AddError(Error error)
        {
            ErrorInstance.AddError(error);
        }

        public static void ResetError()
        {
            ErrorInstance.Reset();
        }

        public static int GetErrorCountByType(object errorType)
        {
            Type type = errorType.GetType();
            return ErrorInstance.GetErrorCountByType(type);
        }

        public static int GetErrorCount()
        {
            return ErrorInstance.GetErrorCount();
        }

        public static void GenErrorReport(string outputFile, List<string> copyFiles)
        {
            ErrorInstance.GenErrorReport(outputFile, copyFiles);
        }
    }
}