using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadTimeSetSheet : IgxlSheetReader
    {
        #region private variable
        private List<string> _headList;
        private int _iStartRowIndex = 0;
        private int _iStartColumnIndex = 0;
        #endregion

        #region public Function
        public TimeSetBasicSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public TimeSetBasicSheet GetSheet(Worksheet worksheet)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet));
        }

        public TimeSetBasicSheet GetSheet(ExcelWorksheet sheet)
        {
            ExcelWorksheet assignmentSheet = sheet;

            ExcelWorksheet lObjSheet_Assignment = null;
            TimeSetBasicSheet lResultIgxlSheet = new TimeSetBasicSheet(assignmentSheet);

            int lIMaxRowCount = 0;
            int lIMaxColumCount = 0;
            int lINowRowIndex = 1;
            string lStrValue = "";
            string lStrValue2 = "";
            string lStrHead = "";
            int lIHeadRowIndex = 0;
            Tset lDataRow = null;
            _headList = new List<string>();
            // Get Source Sheet
            lObjSheet_Assignment = assignmentSheet;

            _iStartRowIndex = 3;
            _iStartColumnIndex = 2;

            lIHeadRowIndex = _iStartRowIndex + 3;

            lINowRowIndex = _iStartRowIndex;

            // Get Max Row Count
            lIMaxRowCount = sheet.Dimension.End.Row;

            // Get Max Colum
            lIMaxColumCount = sheet.Dimension.End.Column;

            // Set Basic Info
            #region Timing Mode & Master Timeset Name
            for (int i = _iStartColumnIndex; i <= lIMaxColumCount; i++)
            {
                lStrValue = GetMergeCellValue(lObjSheet_Assignment, lINowRowIndex, i);
                if (IsLiked(lStrValue, "Timing Mode:") == true)
                {
                    lStrValue = GetMergeCellValue(lObjSheet_Assignment, lINowRowIndex, i + 1);
                    lResultIgxlSheet.TimingMode = lStrValue;
                }

                if (IsLiked(lStrValue, "Master Timeset Name:") == true)
                {
                    lStrValue = GetMergeCellValue(lObjSheet_Assignment, lINowRowIndex, i + 1);
                    lResultIgxlSheet.MasterTimeSet = lStrValue;
                }
            }
            #endregion

            #region Time Domain & Strobe Ref Setup Name
            lINowRowIndex++;
            for (int i = _iStartColumnIndex; i <= lIMaxColumCount; i++)
            {
                lStrValue = GetMergeCellValue(lObjSheet_Assignment, lINowRowIndex, i);
                if (IsLiked(lStrValue, "Time Domain:") == true)
                {
                    lStrValue = GetMergeCellValue(lObjSheet_Assignment, lINowRowIndex, i + 1);
                    lResultIgxlSheet.TimeDomain = lStrValue;
                }

                if (IsLiked(lStrValue, "Strobe Ref Setup Name:") == true)
                {
                    lStrValue = GetMergeCellValue(lObjSheet_Assignment, lINowRowIndex, i + 1);
                    lResultIgxlSheet.StrobeRefSetup = lStrValue;
                }
            }
            #endregion

            // Set Head Index By Souce Sheet
            lINowRowIndex = lINowRowIndex + 2;
            for (int i = _iStartColumnIndex; i <= lIMaxColumCount; i++)
            {
                lStrValue = GetMergeCellValue(lObjSheet_Assignment, lIHeadRowIndex, i);
                lStrValue2 = GetCellText(lObjSheet_Assignment, lIHeadRowIndex + 1, i);

                lStrHead = lStrValue.Trim() + "_" + lStrValue2.Trim();

                _headList.Add(lStrHead);
            }

            lINowRowIndex = lINowRowIndex + 2;

            // Set Row
            for (int i = lINowRowIndex; i <= lIMaxRowCount; i++)
            {
                for (int j = _iStartColumnIndex; j <= lIMaxColumCount; j++)
                {
                    SetItemContent(lObjSheet_Assignment, i, ref lDataRow);
                }

                lResultIgxlSheet.AddTimeSet(lDataRow);
            }

            return lResultIgxlSheet;
        }
        #endregion

        #region Private Function
        private void SetItemContent(ExcelWorksheet peExcelWorksheet, int pIRowIndex, ref Tset pDataRow)
        {
            string lStrHead;
            string lStrContent;

            pDataRow = new Tset();
            TimingRow timingRow = new TimingRow();

            for (int i = _iStartColumnIndex; i < _iStartColumnIndex + _headList.Count; i++)
            {
                lStrHead = _headList[i - _iStartColumnIndex];
                lStrContent = GetCellText(peExcelWorksheet, pIRowIndex, i);
                switch (FormatStringForCompare(lStrHead))
                {
                    case "_TIME_SET":
                        pDataRow.Name = lStrContent;
                        break;
                    case "CYCLE_PERIOD":
                        pDataRow.CyclePeriod = lStrContent;
                        break;
                    case "PIN/GROUP_NAME":
                        timingRow.PinGrpName = lStrContent;
                        break;
                    case "_CLOCK_PERIOD":
                        timingRow.PinGrpClockPeriod = lStrContent;
                        break;
                    case "_SETUP":
                        timingRow.PinGrpSetup = lStrContent;
                        break;
                    case "DATA_SRC":
                        timingRow.DataSrc = lStrContent;
                        break;
                    case "_FMT":
                        timingRow.DataFmt = lStrContent;
                        break;
                    case "DRIVE_ON":
                        timingRow.DriveOn = lStrContent;
                        break;
                    case "_DATA":
                        timingRow.DriveData = lStrContent;
                        break;
                    case "_RETURN":
                        timingRow.DriveReturn = lStrContent;
                        break;
                    case "_OFF":
                        timingRow.DriveOff = lStrContent;
                        break;
                    case "COMPARE_MODE":
                        timingRow.CompareMode = lStrContent;
                        break;
                    case "_OPEN":
                        timingRow.CompareOpen = lStrContent;
                        break;
                    case "_CLOSE":
                        timingRow.CompareClose = lStrContent;
                        break;
                    case "_REF_OFFSET":
                        timingRow.CompareRefOffset = lStrContent;
                        break;
                    case "EDGE_RESOLUTION_MODE":
                        timingRow.EdgeMode = lStrContent;
                        break;
                    case "_COMMENT":
                        timingRow.Comment = lStrContent;
                        break;
                }
            }

            pDataRow.TimingRows.Add(timingRow);
        }
        #endregion
    }
}
