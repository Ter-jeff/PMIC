using CommonLib.Extension;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.Local.Const;
using System.Collections.Generic;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class IoPinMapReader
    {
        public PinMapSheet ReadSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet.Dimension == null) return null;
            var endRow = excelWorksheet.Dimension.End.Row;
            var pinMapSheet = new PinMapSheet(PmicConst.PinMap);
            var groupList = new Dictionary<string, PinGroup>();
            var pinType = "";
            for (var i = 4; i <= endRow; i++)
            {
                var groupName = excelWorksheet.GetCellValue(i, 2);
                var pinName = excelWorksheet.GetCellValue(i, 3).ToUpper();
                pinType = string.IsNullOrEmpty(excelWorksheet.GetCellValue(i, 4))
                    ? pinType
                    : excelWorksheet.GetCellValue(i, 4);
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