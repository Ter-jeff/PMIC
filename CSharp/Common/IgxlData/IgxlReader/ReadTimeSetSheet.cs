using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlReader
{
    public class ReadTimeSetSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 2;
        private const int EndRowIndex = 7;
        private const int StartColumnIndex = 4;
        private readonly List<string> _headList = new List<string>();

        public TimeSetBasicSheet GetSheet(Stream stream, string sheetName)
        {
            var timeSetBasicSheet = new TimeSetBasicSheet(sheetName);
            var maxColumnCount = 5;
            var isBackup = false;
            var i = 1;
            var tempName = "";
            var tSet = new TSet();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > EndRowIndex)
                    {
                        var arr = line.Split('\t');
                        var timingRow = GetTimingRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(timingRow.PinGrpName))
                        {
                            isBackup = true;
                            continue;
                        }

                        timingRow.IsBackup = isBackup;
                        if (tempName != arr[1] && tSet.TimingRows.Any())
                        {
                            timeSetBasicSheet.AddTimeSet(tSet);
                            tSet = new TSet();
                        }
                        tSet.Name = arr[1];
                        tSet.CyclePeriod = arr[2];
                        tSet.AddTimingRow(timingRow);
                        tempName = arr[1];
                    }
                    else
                    {
                        #region Timing Mode & Master Timeset Name
                        var arr = line.Split('\t');
                        for (var col = 0; col <= maxColumnCount; col++)
                        {
                            var value = arr[col];
                            if (IsLiked(value, "Timing Mode:"))
                            {
                                value = arr[col + 1];
                                timeSetBasicSheet.TimingMode = value;
                            }

                            if (IsLiked(value, "Master Timeset Name:"))
                            {
                                value = arr[col + 1];
                                timeSetBasicSheet.MasterTimeSet = value;
                            }
                        }
                        #endregion

                        #region Time Domain & Strobe Ref Setup Name
                        for (var col = 0; col <= maxColumnCount; col++)
                        {
                            var value = arr[col];
                            if (IsLiked(value, "Time Domain:"))
                            {
                                value = arr[col + 1];
                                timeSetBasicSheet.TimeDomain = value;
                            }

                            if (IsLiked(value, "Strobe Ref Setup Name:"))
                            {
                                value = arr[col + 1];
                                timeSetBasicSheet.StrobeRefSetup = value;
                            }
                        }
                        #endregion
                    }
                    i++;
                }
            }
            timeSetBasicSheet.AddTimeSet(tSet);
            return timeSetBasicSheet;
        }

        private TimingRow GetTimingRow(string line, string sheetName, int row)
        {
            var arr = line.Split('\t');
            var timingRow = new TimingRow();
            timingRow.RowNum = row;
            timingRow.SheetName = sheetName;
            var index = StartColumnIndex - 1;
            var content = GetCellText(arr, 0);
            timingRow.ColumnA = content;
            content = GetCellText(arr, index);
            timingRow.PinGrpName = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.PinGrpClockPeriod = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.PinGrpSetup = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.DataSrc = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.DataFmt = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.DriveOn = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.DriveData = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.DriveReturn = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.DriveOff = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.CompareMode = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.CompareOpen = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.CompareClose = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.CompareRefOffset = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.EdgeMode = content;
            index++;
            content = GetCellText(arr, index);
            timingRow.Comment = content;
            return timingRow;
        }

        private void SetItemContent(ExcelWorksheet peExcelWorksheet, int pIRowIndex, out TSet pDataRow)
        {
            pDataRow = new TSet();
            var timingRow = new TimingRow();

            for (var i = StartRowIndex; i < StartRowIndex + _headList.Count; i++)
            {
                var lStrHead = _headList[i - StartRowIndex];
                var content = GetCellText(peExcelWorksheet, pIRowIndex, i);
                switch (FormatStringForCompare(lStrHead))
                {
                    case "_TIME_SET":
                        pDataRow.Name = content;
                        break;
                    case "CYCLE_PERIOD":
                        pDataRow.CyclePeriod = content;
                        break;
                    case "PIN/GROUP_NAME":
                        timingRow.PinGrpName = content;
                        break;
                    case "_CLOCK_PERIOD":
                        timingRow.PinGrpClockPeriod = content;
                        break;
                    case "_SETUP":
                        timingRow.PinGrpSetup = content;
                        break;
                    case "DATA_SRC":
                        timingRow.DataSrc = content;
                        break;
                    case "_FMT":
                        timingRow.DataFmt = content;
                        break;
                    case "DRIVE_ON":
                        timingRow.DriveOn = content;
                        break;
                    case "_DATA":
                        timingRow.DriveData = content;
                        break;
                    case "_RETURN":
                        timingRow.DriveReturn = content;
                        break;
                    case "_OFF":
                        timingRow.DriveOff = content;
                        break;
                    case "COMPARE_MODE":
                        timingRow.CompareMode = content;
                        break;
                    case "_OPEN":
                        timingRow.CompareOpen = content;
                        break;
                    case "_CLOSE":
                        timingRow.CompareClose = content;
                        break;
                    case "_REF_OFFSET":
                        timingRow.CompareRefOffset = content;
                        break;
                    case "EDGE_RESOLUTION_MODE":
                        timingRow.EdgeMode = content;
                        break;
                    case "_COMMENT":
                        timingRow.Comment = content;
                        break;
                }
            }

            pDataRow.TimingRows.Add(timingRow);
        }

        public TimeSetBasicSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public TimeSetBasicSheet GetSheet(ExcelWorksheet sheet)
        {
            var assignmentSheet = sheet;

            var lResultIgxlSheet = new TimeSetBasicSheet(assignmentSheet);

            string lStrValue;
            TSet tSet = null;

            // Get Source Sheet
            var lObjSheetAssignment = assignmentSheet;


            var lIHeadRowIndex = 3 + 3;

            var lINowRowIndex = 3;

            // Get Max Row Count
            var lIMaxRowCount = sheet.Dimension.End.Row;

            // Get Max Column
            var lIMaxColumnCount = sheet.Dimension.End.Column;

            // Set Basic Info

            #region Timing Mode & Master Timeset Name

            for (var i = StartRowIndex; i <= lIMaxColumnCount; i++)
            {
                lStrValue = GetMergeCellValue(lObjSheetAssignment, lINowRowIndex, i);
                if (IsLiked(lStrValue, "Timing Mode:"))
                {
                    lStrValue = GetMergeCellValue(lObjSheetAssignment, lINowRowIndex, i + 1);
                    lResultIgxlSheet.TimingMode = lStrValue;
                }

                if (IsLiked(lStrValue, "Master Timeset Name:"))
                {
                    lStrValue = GetMergeCellValue(lObjSheetAssignment, lINowRowIndex, i + 1);
                    lResultIgxlSheet.MasterTimeSet = lStrValue;
                }
            }

            #endregion

            #region Time Domain & Strobe Ref Setup Name

            lINowRowIndex++;
            for (var i = StartRowIndex; i <= lIMaxColumnCount; i++)
            {
                lStrValue = GetMergeCellValue(lObjSheetAssignment, lINowRowIndex, i);
                if (IsLiked(lStrValue, "Time Domain:"))
                {
                    lStrValue = GetMergeCellValue(lObjSheetAssignment, lINowRowIndex, i + 1);
                    lResultIgxlSheet.TimeDomain = lStrValue;
                }

                if (IsLiked(lStrValue, "Strobe Ref Setup Name:"))
                {
                    lStrValue = GetMergeCellValue(lObjSheetAssignment, lINowRowIndex, i + 1);
                    lResultIgxlSheet.StrobeRefSetup = lStrValue;
                }
            }

            #endregion

            lINowRowIndex = lINowRowIndex + 2;
            for (var i = StartRowIndex; i <= lIMaxColumnCount; i++)
            {
                lStrValue = GetMergeCellValue(lObjSheetAssignment, lIHeadRowIndex, i);
                var lStrValue2 = GetCellText(lObjSheetAssignment, lIHeadRowIndex + 1, i);

                var lStrHead = lStrValue.Trim() + "_" + lStrValue2.Trim();

                _headList.Add(lStrHead);
            }

            lINowRowIndex = lINowRowIndex + 2;

            // Set Row
            for (var i = lINowRowIndex; i <= lIMaxRowCount; i++)
            {
                for (var j = 1; j <= lIMaxColumnCount; j++)
                    SetItemContent(lObjSheetAssignment, i, out tSet);

                lResultIgxlSheet.AddTimeSet(tSet);
            }

            return lResultIgxlSheet;
        }
    }
}