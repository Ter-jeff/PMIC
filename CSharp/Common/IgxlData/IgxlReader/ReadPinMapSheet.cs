using CommonLib.Extension;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;

namespace IgxlData.IgxlReader
{
    public class ReadPinMapSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 3;
        private const int StartColumnIndex = 2;

        public PinMapSheet GetSheet(Stream stream, string sheetName)
        {
            var pinMapSheet = new PinMapSheet(sheetName);
            var isBackup = false;
            var i = 1;
            var groups = new Dictionary<string, PinGroup>();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > StartRowIndex)
                    {
                        var arr = line.Split('\t');
                        var groupName = arr[1];
                        var pinName = arr[2];
                        var pinType = arr[3];
                        var comment = arr[4];
                        if (string.IsNullOrEmpty(pinName))
                        {
                            isBackup = true;
                            continue;
                        }

                        var pin = new Pin(pinName, pinType);
                        pin.RowNum = i;
                        pin.SheetName = sheetName;
                        pin.IsBackup = isBackup;
                        pin.Comment = comment;
                        if (groupName == "" && pinName != "")
                        {
                            pinMapSheet.AddRow(pin);
                        }
                        else if (groupName != "")
                        {
                            if (!groups.ContainsKey(groupName))
                            {
                                var pinGroup = new PinGroup(groupName);
                                pinGroup.RowNum = i;
                                var pins = new List<Pin>();
                                pins.Add(pin);
                                pinGroup.AddPins(pins, pinType);
                                pinMapSheet.AddRow(pinGroup);
                                groups.Add(groupName, pinGroup);
                            }
                            else
                            {
                                groups[groupName].AddPin(pin);
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                    i++;
                }
            }
            pinMapSheet.SortPinGroup();
            return pinMapSheet;
        }

        public PinMapSheet GetSheet(ExcelWorksheet excelWorksheet)
        {
            if (excelWorksheet.Dimension == null) return null;
            var endRow = excelWorksheet.Dimension.End.Row;
            var pinMapSheet = new PinMapSheet(excelWorksheet.Name);
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