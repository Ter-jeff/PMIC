using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace PmicAutogen.Config.ProjectConfig
{
    public class ProjectConfigRow
    {
        #region Constructor

        public ProjectConfigRow()
        {
            Name = "";
            GroupName = "";
            Value = "";
        }

        #endregion

        #region Property

        public string Name { set; get; }
        public string GroupName { set; get; }
        public string Value { set; get; }

        #endregion
    }

    public class ProjectConfigSheet
    {
        #region Constructor

        public ProjectConfigSheet()
        {
            Rows = new List<ProjectConfigRow>();
        }

        #endregion

        #region Property

        public List<ProjectConfigRow> Rows { get; set; }

        #endregion
    }

    public class ProjectConfigReader
    {
        private const string HeaderName = "Name";
        private const string HeaderGroupName = "GroupName";
        private const string HeaderValue = "Value";
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _groupNameIndex = -1;
        private int _nameIndex = -1;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _valueIndex = -1;
        private Worksheet _worksheet;

        public ProjectConfigSheet ReadSheet(Worksheet worksheet)
        {
            if (!ReadHeaders(worksheet)) return new ProjectConfigSheet();

            return ReadRows();
        }

        private bool ReadHeaders(Worksheet worksheet)
        {
            if (worksheet == null) return false;

            _worksheet = worksheet;

            Reset();

            if (!GetDimensions()) return false;

            if (!GetFirstHeaderPosition()) return false;

            GetHeaderIndex();

            return true;
        }

        private ProjectConfigSheet ReadRows()
        {
            var projectConfigSheet = new ProjectConfigSheet();
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new ProjectConfigRow();
                if (_nameIndex != -1)
                    row.Name = GetMergeCellValue2(_worksheet.Cells[i, _nameIndex]).Trim();
                if (_groupNameIndex != -1)
                    row.GroupName = GetMergeCellValue2(_worksheet.Cells[i, _groupNameIndex]).Trim();
                if (_valueIndex != -1)
                    row.Value = GetMergeCellValue2(_worksheet.Cells[i, _valueIndex]).Trim();
                projectConfigSheet.Rows.Add(row);
            }

            return projectConfigSheet;
        }

        public string GetMergeCellValue2(Range range)
        {
            if (!range.MergeCells)
                return range.Value2 != null ? range.Value2.ToString() : string.Empty;

            return range.MergeArea.Cells[1.1].Value2 != null
                ? range.MergeArea.Cells[1.1].Value2.ToString()
                : string.Empty;
        }

        private void GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = GetMergeCellValue2(_worksheet.Cells[_startRowNumber, i]).Trim();
                if (lStrHeader.Equals(HeaderName, StringComparison.OrdinalIgnoreCase))
                {
                    _nameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderGroupName, StringComparison.OrdinalIgnoreCase))
                {
                    _groupNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderValue, StringComparison.OrdinalIgnoreCase)) _valueIndex = i;
            }
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i <= rowNum; i++)
            for (var j = 1; j <= colNum; j++)
                if (GetMergeCellValue2(_worksheet.Cells[i, j]).Trim()
                    .Equals(HeaderName, StringComparison.OrdinalIgnoreCase))
                {
                    _startRowNumber = i;
                    return true;
                }

            return false;
        }

        private bool GetDimensions()
        {
            if (_worksheet.UsedRange != null)
            {
                _startColNumber = _worksheet.UsedRange.Column;
                _startRowNumber = _worksheet.UsedRange.Row;
                _endColNumber = _worksheet.UsedRange.Column + _worksheet.UsedRange.Columns.Count - 1;
                _endRowNumber = _worksheet.UsedRange.Row + _worksheet.UsedRange.Rows.Count - 1;
                return true;
            }

            return false;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _nameIndex = -1;
            _groupNameIndex = -1;
            _valueIndex = -1;
        }
    }
}