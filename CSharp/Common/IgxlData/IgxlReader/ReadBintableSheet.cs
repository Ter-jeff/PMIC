using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.Collections.Generic;

namespace IgxlData.IgxlReader
{
    public class ReadBinTableSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;
        private readonly List<string> _headList = new List<string>();

        #region Private Function

        private BinTableRow GetBinTableRow(ExcelWorksheet sheet, int row)
        {
            var binRow = new BinTableRow();
            var index = 2;
            string lStrContent;
            binRow.RowNum = row;
            lStrContent = GetCellText(sheet, row, index);
            binRow.Name = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            binRow.ItemList = lStrContent;

            index++;
            lStrContent = GetCellText(sheet, row, index);
            binRow.Op = lStrContent;

            index++;
            lStrContent = GetCellText(sheet, row, index);
            binRow.Sort = lStrContent;

            index++;
            lStrContent = GetCellText(sheet, row, index);
            binRow.Bin = lStrContent;

            index++;
            lStrContent = GetCellText(sheet, row, index);
            binRow.Result = lStrContent;
            index++;

            var items = new List<string>();
            for (var i = index; i < sheet.Dimension.Columns; i++)
                if (!string.IsNullOrEmpty(GetCellText(sheet, row, i)))
                    items.Add(GetCellText(sheet, row, i));
            binRow.Items = items;

            return binRow;
        }

        #endregion

        #region public Function

        public BinTableSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public BinTableSheet GetSheet(ExcelWorksheet sheet)
        {
            var subFlowSheet = new BinTableSheet(sheet);
            var maxRowCount = sheet.Dimension.End.Row;
            var maxColumnCount = sheet.Dimension.End.Column;

            // Set Head Index By Source Sheet
            for (var i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                var lStrValue = GetMergeCellValue(sheet, StartRowIndex, i);
                //string lStrValue2 = GetCellText(sheet, StartRowIndex + 1, i);
                var lStrHead = lStrValue.Trim(); // + "_" + lStrValue2.Trim();
                _headList.Add(lStrHead);
            }

            // Set Row
            for (var i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                var binRow = GetBinTableRow(sheet, i);
                if (!string.IsNullOrEmpty(binRow.Name))
                    subFlowSheet.AddRow(binRow);
            }

            return subFlowSheet;
        }

        #endregion
    }
}