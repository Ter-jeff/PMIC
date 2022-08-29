using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadTimeSetSheet : IgxlSheetReader
    {
        #region Private Function

        private void SetItemContent(ExcelWorksheet peExcelWorksheet, int pIRowIndex, out Tset pDataRow)
        {
            pDataRow = new Tset();
            var timingRow = new TimingRow();

            for (var i = _iStartColumnIndex; i < _iStartColumnIndex + _headList.Count; i++)
            {
                var lStrHead = _headList[i - _iStartColumnIndex];
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

        #endregion

        #region private variable

        private readonly List<string> _headList = new List<string>();
        private int _iStartRowIndex;
        private int _iStartColumnIndex;

        #endregion

        #region public Function

        public TimeSetBasicSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public TimeSetBasicSheet GetSheet(ExcelWorksheet sheet)
        {
            var assignmentSheet = sheet;

            var lResultIgxlSheet = new TimeSetBasicSheet(assignmentSheet);

            string lStrValue;
            Tset lDataRow = null;

            // Get Source Sheet
            var lObjSheetAssignment = assignmentSheet;

            _iStartRowIndex = 3;
            _iStartColumnIndex = 2;

            var lIHeadRowIndex = _iStartRowIndex + 3;

            var lINowRowIndex = _iStartRowIndex;

            // Get Max Row Count
            var lIMaxRowCount = sheet.Dimension.End.Row;

            // Get Max Column
            var lIMaxColumnCount = sheet.Dimension.End.Column;

            // Set Basic Info

            #region Timing Mode & Master Timeset Name

            for (var i = _iStartColumnIndex; i <= lIMaxColumnCount; i++)
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
            for (var i = _iStartColumnIndex; i <= lIMaxColumnCount; i++)
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
            for (var i = _iStartColumnIndex; i <= lIMaxColumnCount; i++)
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
                for (var j = _iStartColumnIndex; j <= lIMaxColumnCount; j++)
                    SetItemContent(lObjSheetAssignment, i, out lDataRow);

                lResultIgxlSheet.AddTimeSet(lDataRow);
            }

            return lResultIgxlSheet;
        }

        #endregion
    }
}