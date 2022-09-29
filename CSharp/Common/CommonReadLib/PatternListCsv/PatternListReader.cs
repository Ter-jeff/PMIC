using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace CommonReaderLib.PatternListCsv
{
    public class PatternListReader : MySheetReader
    {
        private const string ConHeaderNumber = "#";
        private const string ConHeaderPattern = "Pattern";
        private const string ConHeaderUseNotUse = "USE/No Use";
        private const string ConHeaderTimesetLatest = "Timeset Latest";
        private const string ConHeaderFileVersions = "File Versions";
        private int _indexFileVersions = -1;

        private int _indexNumber = -1;
        private int _indexPattern = -1;
        private int _indexTimesetLatest = -1;
        private int _indexUseNotUse = -1;

        public PatternListSheet ReadSheet(ExcelWorksheet worksheet)
        {
            var sheetName = worksheet.Name;

            var sheet = new PatternListSheet(sheetName);

            ExcelWorksheet = worksheet;

            if (!GetDimensions())
            {
                sheet.AddDimensionError();
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                sheet.AddFirstHeaderError(ConHeaderNumber);
                return null;
            }

            GetHeaderIndex();

            sheet = ReadSheet(sheetName);

            return sheet;
        }

        private PatternListSheet ReadSheet(string sheetName)
        {
            var sheet = new PatternListSheet(sheetName);
            for (var i = StartRowNumber + 1; i <= EndRowNumber; i++)
            {
                var row = new PatternListRow(sheetName);
                row.RowNum = i;
                if (_indexNumber != -1)
                    row.Number = ExcelWorksheet.GetMergedCellValue(i, _indexNumber).Trim();
                if (_indexPattern != -1)
                    row.Pattern = ExcelWorksheet.GetMergedCellValue(i, _indexPattern).Trim();
                if (_indexUseNotUse != -1)
                    row.UseNotUse = ExcelWorksheet.GetMergedCellValue(i, _indexUseNotUse).Trim();
                if (_indexTimesetLatest != -1)
                    row.TimeSetLatest = ExcelWorksheet.GetMergedCellValue(i, _indexTimesetLatest).Trim();
                if (_indexFileVersions != -1)
                    row.FileVersions = ExcelWorksheet.GetMergedCellValue(i, _indexFileVersions).Trim();
                if (!string.IsNullOrEmpty(row.Pattern))
                    sheet.Rows.Add(row);
            }

            sheet.IndexNumber = _indexNumber;
            sheet.IndexPattern = _indexPattern;
            sheet.IndexUseNotUse = _indexUseNotUse;
            sheet.IndexTimesetLatest = _indexTimesetLatest;
            sheet.IndexFileVersions = _indexFileVersions;

            return sheet;
        }

        private void GetHeaderIndex()
        {
            for (var i = StartColNumber; i <= EndColNumber; i++)
            {
                var header = ExcelWorksheet.GetCellValue(StartRowNumber, i).Trim();
                if (header.Equals(ConHeaderNumber, StringComparison.OrdinalIgnoreCase))
                {
                    _indexNumber = i;
                    continue;
                }

                if (header.Equals(ConHeaderPattern, StringComparison.OrdinalIgnoreCase))
                {
                    _indexPattern = i;
                    continue;
                }

                if (header.Equals(ConHeaderUseNotUse, StringComparison.OrdinalIgnoreCase))
                {
                    _indexUseNotUse = i;
                    continue;
                }

                if (header.Equals(ConHeaderTimesetLatest, StringComparison.OrdinalIgnoreCase))
                {
                    _indexTimesetLatest = i;
                    continue;
                }

                if (header.Equals(ConHeaderFileVersions, StringComparison.OrdinalIgnoreCase))
                {
                    _indexFileVersions = i;
                }
            }
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = EndRowNumber > 10 ? 10 : EndRowNumber;
            var colNum = EndColNumber > 10 ? 10 : EndColNumber;
            for (var i = 1; i <= rowNum; i++)
                for (var j = 1; j <= colNum; j++)
                    if (ExcelWorksheet.GetCellValue(i, j).Trim()
                        .Equals(ConHeaderNumber, StringComparison.OrdinalIgnoreCase))
                    {
                        StartRowNumber = i;
                        return true;
                    }

            return false;
        }
    }

    public class PatternListSheet : MySheet
    {
        public int IndexFileVersions = -1;
        public int IndexNumber = -1;
        public int IndexPattern = -1;
        public int IndexTimesetLatest = -1;
        public int IndexUseNotUse = -1;

        #region Constructor

        public PatternListSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<PatternListRow>();
        }

        #endregion

        public List<PatternListRow> Rows { set; get; }

        public void CheckPatternTimeSet(string timeFolder, List<string> patterns)
        {
            foreach (var row in Rows)
            {
                var timeSet = Path.Combine(timeFolder, row.TimeSet);
                if (!row.TimeSet.Equals("NA", StringComparison.CurrentCultureIgnoreCase) &&
                    !string.IsNullOrEmpty(row.TimeSet) &&
                    !File.Exists(timeSet))
                    Errors.Add(new Error
                    {
                        EnumErrorType = EnumErrorType.Missing,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = SheetName,
                        RowNum = row.RowNum,
                        ColNum = IndexTimesetLatest,
                        Message = string.Format("[Timeset Latest] The timeset \"{0}\" in pattern dash board is not existed in pattern folder !!!", timeSet)
                    });

                if (!patterns.Exists(x => x.Equals(row.PatternDate, StringComparison.CurrentCultureIgnoreCase)))
                    Errors.Add(new Error
                    {
                        EnumErrorType = EnumErrorType.Missing,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = SheetName,
                        RowNum = row.RowNum,
                        ColNum = IndexFileVersions,
                        Message = string.Format("[File Versions] The pattern \"{0}\" in pattern dash board is not existed in pattern folder !!!", row.PatternDate)
                    });
            }
        }
    }

    public class PatternListRow : MyRow
    {
        #region Property

        public string Number { set; get; }
        public string Pattern { set; get; }
        public string UseNotUse { set; get; }
        public string TimeSetLatest { set; get; }
        public string FileVersions { set; get; }

        public string TimeSet
        {
            get { return Path.GetFileName(TimeSetLatest); }
        }

        public string PatternDate
        {
            get
            {
                var patternVersion = Path.GetFileName(Regex.Replace(Regex.Replace(FileVersions, ".gz$", "",
                    RegexOptions.IgnoreCase), ".atp$", "", RegexOptions.IgnoreCase));
                patternVersion = Regex.Replace(patternVersion, ".Pat$", "", RegexOptions.IgnoreCase);
                return patternVersion;
            }
        }

        #endregion

        #region Constructor

        public PatternListRow()
        {
        }

        public PatternListRow(string sourceSheetName)
        {
            SheetName = sourceSheetName;
        }

        #endregion
    }
}