using System;
using System.Collections.Generic;
using AutomationCommon.Utility;
using OfficeOpenXml;

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

        public List<SoftBinRangeRow> ReadSheet(ExcelWorksheet worksheet)
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
            for (var i = 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (var j = 1; j <= worksheet.Dimension.End.Column; j++)
                    if (EpplusOperation.GetCellValue(worksheet, i, j)
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

            for (var i = columnDescription; i <= worksheet.Dimension.End.Column; i++)
                if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(RangeStart, StringComparison.OrdinalIgnoreCase))
                    columnRangeStart = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(RangeEnd, StringComparison.OrdinalIgnoreCase))
                    columnRangeEnd = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(RangeState, StringComparison.OrdinalIgnoreCase))
                    columnRangeState = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(RangeBin, StringComparison.OrdinalIgnoreCase))
                    columnRangeBin = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(Block, StringComparison.OrdinalIgnoreCase))
                    columnBlock = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(Condition, StringComparison.OrdinalIgnoreCase))
                    columnCondition = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(HardBin, StringComparison.OrdinalIgnoreCase))
                    columnHardBin = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(Hlv, StringComparison.OrdinalIgnoreCase))
                    columnHardHlvBin = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(Lv, StringComparison.OrdinalIgnoreCase))
                    columnHardLvBin = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(Nv, StringComparison.OrdinalIgnoreCase))
                    columnHardNvBin = i;
                else if (EpplusOperation.GetCellValue(worksheet, startRow, i)
                    .Equals(Hv, StringComparison.OrdinalIgnoreCase)) columnHardHvBin = i;

            if (columnRangeBin == 0 || columnRangeEnd == 0 || columnRangeStart == 0 || columnRangeState == 0 ||
                columnHardBin == 0 || columnBlock == 0 ||
                columnCondition == 0) throw new Exception("Error occur when Initial Soft Bin Range ");

            for (var i = startRow + 1; i <= worksheet.Dimension.End.Row; i++)
            {
                var data = new SoftBinRangeRow();
                data.Description = EpplusOperation.GetCellValue(worksheet, i, columnDescription);
                data.Start = EpplusOperation.GetCellValue(worksheet, i, columnRangeStart);
                data.End = EpplusOperation.GetCellValue(worksheet, i, columnRangeEnd);
                data.State = EpplusOperation.GetCellValue(worksheet, i, columnRangeState);
                data.Block = EpplusOperation.GetCellValue(worksheet, i, columnBlock).Replace("_", "").Replace(" ", "");
                data.Condition = EpplusOperation.GetCellValue(worksheet, i, columnCondition).Replace("_", "")
                    .Replace(" ", "");
                data.HardBin = EpplusOperation.GetCellValue(worksheet, i, columnHardBin);
                if (columnHardHlvBin != 0)
                    data.HardHlvBin = EpplusOperation.GetCellValue(worksheet, i, columnHardHlvBin);
                if (columnHardHvBin != 0)
                    data.HardHvBin = EpplusOperation.GetCellValue(worksheet, i, columnHardHvBin);
                if (columnHardLvBin != 0)
                    data.HardLvBin = EpplusOperation.GetCellValue(worksheet, i, columnHardLvBin);
                if (columnHardNvBin != 0)
                    data.HardNvBin = EpplusOperation.GetCellValue(worksheet, i, columnHardNvBin);
                if (!string.IsNullOrEmpty(data.Block) && !string.IsNullOrEmpty(data.Condition))
                    softBinRanges.Add(data);
            }

            return softBinRanges;
        }
    }
}