using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace PmicAutomation.Utility.Relay.Input
{
    public class LinkedNodeRuleRow
    {
        #region Properity

        public string SourceSheetName;
        public int RowNum;
        public string Node { set; get; }
        public string LinkedNode { set; get; }

        #endregion

        #region Constructor

        public LinkedNodeRuleRow()
        {
            Node = "";
            LinkedNode = "";
        }

        public LinkedNodeRuleRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            Node = "";
            LinkedNode = "";
        }

        #endregion
    }

    public class LinkedNodeRuleSheet
    {
        #region Constructor

        public LinkedNodeRuleSheet(string name)
        {
            Name = name;
            Rows = new List<LinkedNodeRuleRow>();
        }

        #endregion

        #region Field

        #endregion

        #region Properity

        public string Name { get; set; }
        public List<LinkedNodeRuleRow> Rows { get; set; }
        public Dictionary<string, int> HeaderIndexDic  = new Dictionary<string, int>();

        #endregion
    }

    public class LinkedNodeRuleReader
    {
        private const string HeaderNode = "Node";
        private const string HeaderLinkedNode = "LinkedNode";

        private readonly Dictionary<string, bool> _headerOptional = new Dictionary<string, bool>
        {
            {"Node", true}, {"LinkedNode", true}
        };

        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _linkedNodeIndex = -1;
        private LinkedNodeRuleSheet _linkedNodeRuleSheet;
        private int _nodeIndex = -1;
        private string _name;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;

        public LinkedNodeRuleSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            _name = worksheet.Name;

            _linkedNodeRuleSheet = new LinkedNodeRuleSheet(_name);

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            if (!GetFirstHeaderPosition())
            {
                return null;
            }

            if (!GetHeaderIndex())
            {
                return null;
            }

            _linkedNodeRuleSheet = ReadSheetData();

            return _linkedNodeRuleSheet;
        }

        private LinkedNodeRuleSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                LinkedNodeRuleRow row = new LinkedNodeRuleRow(_name) {RowNum = i};
                if (_nodeIndex != -1)
                {
                    row.Node = _excelWorksheet.GetMergeCellValue(i, _nodeIndex).Trim();
                }

                if (_linkedNodeIndex != -1)
                {
                    row.LinkedNode = _excelWorksheet.GetMergeCellValue(i, _linkedNodeIndex).Trim();
                }

                _linkedNodeRuleSheet.Rows.Add(row);
            }

            return _linkedNodeRuleSheet;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = _excelWorksheet.GetMergeCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderNode, StringComparison.OrdinalIgnoreCase))
                {
                    _nodeIndex = i;
                    _linkedNodeRuleSheet.HeaderIndexDic.Add(HeaderNode, i);
                    continue;
                }

                if (lStrHeader.Equals(HeaderLinkedNode, StringComparison.OrdinalIgnoreCase))
                {
                    _linkedNodeIndex = i;
                    _linkedNodeRuleSheet.HeaderIndexDic.Add(HeaderLinkedNode, i);
                }
            }

            foreach (KeyValuePair<string, int> header in _linkedNodeRuleSheet.HeaderIndexDic)
            {
                if (header.Value == -1 && _headerOptional.ContainsKey(header.Key) && _headerOptional[header.Key])
                {
                    return false;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
            for (int j = 1; j <= colNum; j++)
            {
                if (_excelWorksheet.GetMergeCellValue(i, j).Trim()
                    .Equals(HeaderNode, StringComparison.OrdinalIgnoreCase))
                {
                    _startRowNumber = i;
                    return true;
                }
            }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
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
            _nodeIndex = -1;
            _linkedNodeIndex = -1;
        }

        public List<Dictionary<string, string>> GenMappingDictionary()
        {
            List<Dictionary<string, string>> dictionary = new List<Dictionary<string, string>>();
            foreach (LinkedNodeRuleRow row in _linkedNodeRuleSheet.Rows)
            {
                Dictionary<string, string> dic = new Dictionary<string, string>
                {
                    {"Node", row.Node}, {"LinkedNode", row.LinkedNode}
                };
                dictionary.Add(dic);
            }

            return dictionary;
        }
    }
}