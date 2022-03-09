using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadBintableSheet : IgxlSheetReader
    {
        private List<string> _headList = new List<string>();
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function
        public BinTableSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public BinTableSheet GetSheet(ExcelWorksheet sheet)
        {
            BinTableSheet subFlowSheet = new BinTableSheet(sheet);
            int maxRowCount = sheet.Dimension.End.Row;
            int maxColumnCount = sheet.Dimension.End.Column;

            // Set Head Index By Source Sheet
            for (int i = StartColumnIndex; i <= maxColumnCount; i++)
            {
                string lStrValue = GetMergeCellValue(sheet, StartRowIndex, i);
                //string lStrValue2 = GetCellText(sheet, StartRowIndex + 1, i);
                string lStrHead = lStrValue.Trim();// + "_" + lStrValue2.Trim();
                _headList.Add(lStrHead);
            }

            // Set Row
            for (int i = StartRowIndex + 1; i <= maxRowCount; i++)
            {
                BinTableRow binRow = GetBinTableRow(sheet, i);
                if (!string.IsNullOrEmpty(binRow.Name))
                    subFlowSheet.AddRow(binRow);
            }
            return subFlowSheet;
        }
        #endregion

        #region Private Function
        private BinTableRow GetBinTableRow(ExcelWorksheet sheet, int row)
        {
            BinTableRow binRow = new BinTableRow();
            int index = 2;
            string lStrContent;
            binRow.LinNum = row;
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
            for (int i = index; i < sheet.Dimension.Columns; i++)
            {
                if (!string.IsNullOrEmpty(GetCellText(sheet, row, i)))
                    items.Add(GetCellText(sheet, row, i));
            }
            binRow.Items = items;

            return binRow;
        }
        #endregion
    }
}
