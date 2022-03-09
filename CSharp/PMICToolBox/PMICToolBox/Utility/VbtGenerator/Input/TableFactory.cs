using Library.Function;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PmicAutomation.Utility.VbtGenerator.Input
{
    public abstract class TableFactory
    {
        protected int _endColNumber = -1;
        protected int _endRowNumber = -1;
        protected ExcelWorksheet _excelWorksheet;

        public static TableFactory CreateFactory(bool isPMIC_IDS_TP)
        {
            if (isPMIC_IDS_TP)
                return new PmicIdsTableReader();
            else
                return new CommonTableReader();
        }

        public abstract List<TableSheet> ReadSheet(ExcelWorksheet worksheet);

        protected bool IsAllPinSetting(Dictionary<string, string> pairs)
        {
            foreach (KeyValuePair<string, string> pair in pairs)
            {
                if (pair.Value != null &&
                    pair.Value.StartsWith("All Pin Setting", StringComparison.CurrentCultureIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        protected int IsHeaderRow(int i)
        {
            int cnt = 0;
            for (int j = 1; j <= _endColNumber; j++)
            {
                string content = _excelWorksheet.GetMergeCellValue(i, j);
                if (!string.IsNullOrEmpty(content))
                {
                    cnt++;
                }
            }

            return cnt;
        }

        protected bool CheckHeader()
        {
            //var i = 1;
            //for (var j = 1; j <= _endColNumber; j++)
            //{
            //    var content = _excelWorksheet.GetMergeCellValue(i, j);
            //    if (string.IsNullOrEmpty(content))
            //    {
            //        _append($@"The cell {_excelWorksheet.Cells[i, j].Address} can't be empty !!!", Color.Red);
            //        return true;
            //    }
            //}
            return true;
        }

        protected bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }

            return false;
        }

        protected void Reset()
        {
            _endColNumber = -1;
            _endRowNumber = -1;
        }
    }
}
