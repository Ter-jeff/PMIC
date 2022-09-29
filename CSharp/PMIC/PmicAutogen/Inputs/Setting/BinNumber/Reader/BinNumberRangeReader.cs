using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;

namespace PmicAutogen.Inputs.Setting.BinNumber.Reader
{
    public class BinNumberRangeReader
    {
        private const string Description = "Description";
        private const string RangeStart = "SW BIN Range Start";
        private const string RangeEnd = "SW BIN Range End";
        private const string RangeState = "good or bad";
        private const string RangeBin = "Bin #";
        private const string Block = "Block";
        private const string Condition = "Condition";
        private const string Hv = "HV";
        private const string Lv = "LV";
        private const string Nv = "NV";
        private const string Hlv = "HLV";
        private const string HardBin = "Hard Bin";

        public List<SoftBinRangeRow> ReadSheet(ExcelWorksheet excelWorksheet)
        {
            var softBinRanges = new List<SoftBinRangeRow>();
            var startRow = 0;
            var columnDescription = 0;
            var columnRangeStart = 0;
            var columnRangeEnd = 0;
            var columnRangeState = 0;
            var columnRangeBin = 0;
            var columnBlock = 0;
            var columnCondition = 0;
            var columnHardHvBin = 0;
            var columnHardLvBin = 0;
            var columnHardNvBin = 0;
            var columnHardHlvBin = 0;
            var columnHardBin = 0;
            for (var i = 1; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                for (var j = 1; j <= excelWorksheet.Dimension.End.Column; j++)
                    if (excelWorksheet.GetCellValue(i, j)
                        .Equals(Description, StringComparison.OrdinalIgnoreCase))
                    {
                        startRow = i;
                        columnDescription = j;
                        break;
                    }

                if (startRow != 0) break;
            }

            if (startRow == 0 || columnDescription == 0)
                throw new Exception("Error occur when Initial Soft Bin Range ");

            for (var i = columnDescription; i <= excelWorksheet.Dimension.End.Column; i++)
                if (excelWorksheet.GetCellValue(startRow, i)
                    .Equals(RangeStart, StringComparison.OrdinalIgnoreCase))
                    columnRangeStart = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(RangeEnd, StringComparison.OrdinalIgnoreCase))
                    columnRangeEnd = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(RangeState, StringComparison.OrdinalIgnoreCase))
                    columnRangeState = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(RangeBin, StringComparison.OrdinalIgnoreCase))
                    columnRangeBin = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(Block, StringComparison.OrdinalIgnoreCase))
                    columnBlock = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(Condition, StringComparison.OrdinalIgnoreCase))
                    columnCondition = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(HardBin, StringComparison.OrdinalIgnoreCase))
                    columnHardBin = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(Hlv, StringComparison.OrdinalIgnoreCase))
                    columnHardHlvBin = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(Lv, StringComparison.OrdinalIgnoreCase))
                    columnHardLvBin = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(Nv, StringComparison.OrdinalIgnoreCase))
                    columnHardNvBin = i;
                else if (excelWorksheet.GetCellValue(startRow, i)
                         .Equals(Hv, StringComparison.OrdinalIgnoreCase)) columnHardHvBin = i;

            if (columnRangeBin == 0 || columnRangeEnd == 0 || columnRangeStart == 0 || columnRangeState == 0 ||
                columnHardBin == 0 || columnBlock == 0 ||
                columnCondition == 0) throw new Exception("Error occur when Initial Soft Bin Range ");

            for (var i = startRow + 1; i <= excelWorksheet.Dimension.End.Row; i++)
            {
                var data = new SoftBinRangeRow();
                data.Description = excelWorksheet.GetCellValue(i, columnDescription);
                data.Start = excelWorksheet.GetCellValue(i, columnRangeStart);
                data.End = excelWorksheet.GetCellValue(i, columnRangeEnd);
                data.State = excelWorksheet.GetCellValue(i, columnRangeState);
                data.Block = excelWorksheet.GetCellValue(i, columnBlock).Replace("_", "").Replace(" ", "");
                data.Condition = excelWorksheet.GetCellValue(i, columnCondition).Replace("_", "")
                    .Replace(" ", "");
                data.HardBin = excelWorksheet.GetCellValue(i, columnHardBin);
                if (columnHardHlvBin != 0)
                    data.HardHlvBin = excelWorksheet.GetCellValue(i, columnHardHlvBin);
                if (columnHardHvBin != 0)
                    data.HardHvBin = excelWorksheet.GetCellValue(i, columnHardHvBin);
                if (columnHardLvBin != 0)
                    data.HardLvBin = excelWorksheet.GetCellValue(i, columnHardLvBin);
                if (columnHardNvBin != 0)
                    data.HardNvBin = excelWorksheet.GetCellValue(i, columnHardNvBin);
                if (!string.IsNullOrEmpty(data.Block) && !string.IsNullOrEmpty(data.Condition))
                    softBinRanges.Add(data);
            }

            return softBinRanges;
        }
    }
}