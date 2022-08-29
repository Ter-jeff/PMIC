using System.Collections.Generic;
using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadPinMapSheet : IgxlSheetReader
    {
        internal IgxlSheet GetSheet(ExcelWorksheet sheet)
        {
            if (sheet.Dimension == null) return null;
            var endRow = sheet.Dimension.End.Row;
            var pinMapSheet = new PinMapSheet(sheet.Name);
            var groupList = new Dictionary<string, PinGroup>();
            var pinType = "";
            for (var i = 4; i <= endRow; i++)
            {
                var groupName = EpplusOperation.GetCellValue(sheet, i, 2);
                var pinName = EpplusOperation.GetCellValue(sheet, i, 3).ToUpper();
                pinType = string.IsNullOrEmpty(EpplusOperation.GetCellValue(sheet, i, 4))
                    ? pinType
                    : EpplusOperation.GetCellValue(sheet, i, 4);
                if (groupName == "" && pinName != "")
                {
                    var pin = new Pin(pinName, pinType);
                    pinMapSheet.AddRow(pin);
                }
                else if (groupName != "")
                {
                    if (groupList.ContainsKey(groupName) == false)
                    {
                        var pinGroup = new PinGroup(groupName);
                        var pinList = new List<string>();
                        pinList.Add(pinName);
                        pinGroup.AddPins(pinList, pinType);
                        pinMapSheet.AddRow(pinGroup);
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

            pinMapSheet.SortPinGroup();
            return pinMapSheet;
        }
    }
}