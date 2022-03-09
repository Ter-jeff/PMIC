using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Collections.Generic;

namespace IgxlData.IgxlReader
{
    public class ReadBinTableSheet : IgxlSheetReader
    {
        private readonly List<string> _headList = new List<string>();
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        #region public Function
        public BinTableSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public BinTableSheet GetSheet(Worksheet worksheet)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet));
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
                BinTableRow BinRow = GetBinTableRow(sheet, i);
                subFlowSheet.AddRow(BinRow);
            }
            return subFlowSheet;
        }
        #endregion

        #region Private Function
        private BinTableRow GetBinTableRow(ExcelWorksheet sheet, int row)
        {



            BinTableRow BinRow = new BinTableRow();
            int index = 2;
            string lStrContent;
            BinRow.LinNum = row;
            lStrContent = GetCellText(sheet, row, index);
            BinRow.Name = lStrContent;
            index++;
            lStrContent = GetCellText(sheet, row, index);
            BinRow.ItemList = lStrContent;

            index++;
            lStrContent = GetCellText(sheet, row, index);
            BinRow.Op = lStrContent;

            index++;
            lStrContent = GetCellText(sheet, row, index);
            BinRow.Sort = lStrContent;

            index++;
            lStrContent = GetCellText(sheet, row, index);
            BinRow.Bin = lStrContent;

            index++;
            lStrContent = GetCellText(sheet, row, index);
            BinRow.Result = lStrContent;
            index++;

            var items = new List<string>();
            for (int i = index; i < sheet.Dimension.Columns; i++)
            {
                if (!string.IsNullOrEmpty(GetCellText(sheet, row, i)))
                    items.Add(GetCellText(sheet, row, i));
            }

            BinRow.Items = items;



            for (int i = index; i < sheet.Dimension.Columns; i++)
            {
                BinRow.ItemsWithIndex.Add(i - index, GetCellText(sheet, row, i));
            }

            //var flags = BinRow.ItemList.Split(',').ToArray();
            //for (int i = 0; i < flags.Count(); i++)
            //{
            //    BinRow.FlagEnableMap.Add(flags[i], GetCellText(sheet, row, index + i));
            //}

            return BinRow;
        }
        #endregion
    }
}
