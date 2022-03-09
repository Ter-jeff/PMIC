
﻿using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.Collections.Generic;


namespace IgxlData.IgxlReader
{
    public class ReadPinMapSheet : IgxlSheetReader
    {

        #region public Function
        public PinMapSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public PinMapSheet GetSheet(Worksheet worksheet)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet));
        }

        public PinMapSheet GetSheet(ExcelWorksheet pinMapSheet)
        {
            if (pinMapSheet.Dimension == null)
                return null;
            int endRow = pinMapSheet.Dimension.End.Row;
            var mapSheet = new PinMapSheet(pinMapSheet);
            var groupList = new Dictionary<string, PinGroup>();
            string pinType = "";
            for (int i = 4; i <= endRow; i++)
            {
                string groupName = "";
                string pinName = "";
                groupName = EpplusOperation.GetCellValue(pinMapSheet, i, 2);
                pinName = EpplusOperation.GetCellValue(pinMapSheet, i, 3).ToUpper();
                pinType = string.IsNullOrEmpty(EpplusOperation.GetCellValue(pinMapSheet, i, 4)) ? pinType : EpplusOperation.GetCellValue(pinMapSheet, i, 4);
                if (groupName == "" && pinName != "")
                {
                    Pin pin = new Pin(pinName, pinType, "");
                    mapSheet.AddRow(pin);
                }
                else if (groupName != "")
                {
                    if (groupList.ContainsKey(groupName) == false)
                    {
                        PinGroup pinGroup = new PinGroup(groupName, pinType);
                        List<string> pinList = new List<string>();
                        pinList.Add(pinName);
                        pinGroup.AddPins(pinList, pinType);
                        mapSheet.AddRow(pinGroup);
                        groupList.Add(groupName, pinGroup);
                    }
                    else
                    {
                        groupList[groupName].AddPin(pinName, pinType);
                    }
                }
                else
                {
                    break;
                }
            }
            mapSheet.SortPinGroup();
            return mapSheet;
        }
        #endregion
    }
}
