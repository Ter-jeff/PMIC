using System.Collections.Generic;
using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Local.Const;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class IoPinMapReader
    {
        public PinMapSheet ReadSheet(ExcelWorksheet sheet)
        {
            if (sheet.Dimension == null) return null;
            var endRow = sheet.Dimension.End.Row;
            var pinMapSheet = new PinMapSheet(PmicConst.PinMap);
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