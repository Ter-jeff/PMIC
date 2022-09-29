using CommonLib.Extension;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows;

namespace PmicAutogen.GenerateIgxl.Basic.GenContiVbt
{
    public class VbtDictionaryReader
    {
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;

        public List<Dictionary<string, string>> ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            Reset();

            if (!GetDimensions()) return null;

            if (!CheckHeader()) return null;

            var dicList = new List<Dictionary<string, string>>();
            for (var i = 2; i <= _endRowNumber; i++)
            {
                var dic = new Dictionary<string, string>();
                for (var j = 1; j <= _endColNumber; j++)
                {
                    var header = _excelWorksheet.GetMergedCellValue(1, j).Replace(" ", "");
                    var content = _excelWorksheet.GetMergedCellValue(i, j);
                    dic.Add(header, content);
                }

                dicList.Add(dic);
            }

            return dicList;
        }

        private bool CheckHeader()
        {
            var i = 1;
            for (var j = 1; j <= _endColNumber; j++)
            {
                var content = _excelWorksheet.GetCellValue(i, j);
                if (string.IsNullOrEmpty(content))
                {
                    MessageBox.Show(string.Format("The cell {0} can't be empty !!!",
                        _excelWorksheet.Cells[i, j].Address));
                    return true;
                }

                if (Regex.IsMatch(content, @"\W"))
                {
                    MessageBox.Show(string.Format("The cell {0} can't contains any special char !!!",
                        _excelWorksheet.Cells[i, j].Address));
                    return true;
                }
            }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }

            return false;
        }

        private void Reset()
        {
            _endColNumber = -1;
            _endRowNumber = -1;
        }
    }
}