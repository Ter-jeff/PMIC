using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Data;
using DataTable = System.Data.DataTable;

namespace IgxlData.IgxlReader
{
    public class ReadNonIgxlSheet : IgxlSheetReader
    {

        #region public Function

        public DataTable GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }


        public DataTable GetSheet(Worksheet worksheet)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet));
        }


        private DataTable GetSheet(ExcelWorksheet excelWorksheet)
        {
            var hasHeader = true;
            var tbl = new DataTable();
            foreach (var firstRowCell in excelWorksheet.Cells[1, 1, 1, excelWorksheet.Dimension.End.Column])
            {
                tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
            }
            var startRow = hasHeader ? 2 : 1;
            for (int rowNum = startRow; rowNum <= excelWorksheet.Dimension.End.Row; rowNum++)
            {
                var wsRow = excelWorksheet.Cells[rowNum, 1, rowNum, excelWorksheet.Dimension.End.Column];
                DataRow row = tbl.Rows.Add();
                foreach (var cell in wsRow)
                {
                    row[cell.Start.Column - 1] = cell.Text;
                }
            }
            return tbl;
        }

        #endregion

    }
}