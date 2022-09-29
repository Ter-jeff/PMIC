using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlReader
{
    public class ReadPortMapSheet : IgxlSheetReader
    {
        private const int StartRowIndex = 4;
        private const int StartColumnIndex = 2;
        private readonly SheetObjMap _sheetObjMap;

        public ReadPortMapSheet()
        {
        }

        public ReadPortMapSheet(SheetObjMap sheetObjMap)
        {
            _sheetObjMap = sheetObjMap;
        }

        public PortMapSheet GetSheet(Stream stream, string sheetName)
        {
            var portMapSheet = new PortMapSheet(sheetName);
            var isBackup = false;
            var i = 1;
            var tempName = "";
            var portSet = new PortSet();
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > StartRowIndex)
                    {
                        var arr = line.Split('\t');
                        var portRow = GetPortRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(portRow.PortName))
                        {
                            isBackup = true;
                            continue;
                        }

                        portRow.IsBackup = isBackup;
                        if (tempName != arr[1] && portSet.PortRows.Any())
                        {
                            portMapSheet.AddPortSet(portSet);
                            portSet = new PortSet();
                        }
                        portSet.PortName = arr[1];
                        portSet.AddPortRow(portRow);
                        tempName = arr[1];
                    }
                    i++;
                }
            }
            portMapSheet.AddPortSet(portSet);
            return portMapSheet;
        }

        private PortRow GetPortRow(string line, string sheetName, int row)
        {
            var arr = line.Split('\t');
            var portRow = new PortRow();
            portRow.RowNum = row;
            portRow.SheetName = sheetName;
            var index = StartColumnIndex - 1;
            var content = GetCellText(arr, 0);
            portRow.ColumnA = content;
            content = GetCellText(arr, index);
            portRow.PortName = content;
            index++;
            content = GetCellText(arr, index);
            portRow.ProtocolFamily = content;
            index++;
            content = GetCellText(arr, index);
            portRow.ProtocolType = content;
            index++;
            content = GetCellText(arr, index);
            portRow.ProtocolSettings = content;
            index++;
            for (int i = 0; i < 10; i++)
            {
                content = GetCellText(arr, index);
                portRow.ProtocolSettingValues.Add(content);
                index++;
            }
            content = GetCellText(arr, index);
            portRow.FunctionName = content;
            index++;
            content = GetCellText(arr, index);
            portRow.FunctionPin = content;
            index++;
            content = GetCellText(arr, index);
            portRow.FunctionProperties = content;
            index++;
            for (int i = 0; i < 10; i++)
            {
                content = GetCellText(arr, index);
                portRow.FunctionPropertyValues.Add(content);
                index++;
            }
            index++;
            content = GetCellText(arr, index);
            portRow.Comment = content;
            return portRow;
        }

        public PortMapSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public PortMapSheet GetSheet(ExcelWorksheet sheet)
        {
            var portMapSheet = new PortMapSheet(sheet);
            var stopRow = sheet.Dimension.End.Row;
            var portRows = new List<PortRow>();
            for (var i = StartRowIndex + 1; i <= stopRow; i++)
            {
                var portRow = new PortRow();
                portRow.RowNum = i;
                foreach (var innerObj in _sheetObjMap.InnerObj)
                {
                    foreach (var property in innerObj.Property)
                    {
                        var value = GetCellText(sheet, i, property.indexInSheet + StartColumnIndex - 1);
                        var propertyInfo = portRow.GetType().GetProperty(property.name);
                        if (value != null && propertyInfo != null)
                            propertyInfo.SetValue(portRow, value, null);
                    }

                    const int startIndex = 0;
                    foreach (var innerObjComplex in innerObj.InnerObjComplex)
                        GetValueByInnerObjComplex(sheet, innerObjComplex, startIndex, i, portRow);
                }
                portRows.Add(portRow);
            }

            for (var i = 0; i < portRows.Count; i++)
            {
                var portRow = portRows[i];
                var portSet = new PortSet(portRow.PortName);
                portSet.AddPortRow(portRow);
                for (var j = i + 1; j < portRows.Count; j++)
                    if (portRow.PortName.Equals(portRows[j].PortName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        portSet.AddPortRow(portRows[j]);
                    }
                    else
                    {
                        i = j - 1;
                        break;
                    }

                portMapSheet.AddPortSet(portSet);
            }

            return portMapSheet;
        }

        public PortMapSheet GetSheet(Worksheet sheet)
        {
            var portMapSheet = new PortMapSheet(sheet);
            var stopRow = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1;
            var portRows = new List<PortRow>();
            for (var i = StartRowIndex + 1; i <= stopRow; i++)
            {
                var portRow = new PortRow();
                portRow.RowNum = i;
                foreach (var innerObj in _sheetObjMap.InnerObj)
                {
                    foreach (var property in innerObj.Property)
                    {
                        var value = GetCellText(sheet, i, property.indexInSheet + StartColumnIndex - 1);
                        var propertyInfo = portRow.GetType().GetProperty(property.name);
                        if (value != null && propertyInfo != null)
                            propertyInfo.SetValue(portRow, value, null);
                    }

                    const int startIndex = 0;
                    foreach (var innerObjComplex in innerObj.InnerObjComplex)
                        GetValueByInnerObjComplex(sheet, innerObjComplex, startIndex, i, portRow);
                }

                portRows.Add(portRow);
            }

            for (var i = 0; i < portRows.Count; i++)
            {
                var portRow = portRows[i];
                var portSet = new PortSet(portRow.PortName);
                portSet.AddPortRow(portRow);
                for (var j = i + 1; j < portRows.Count; j++)
                    if (portRow.PortName.Equals(portRows[j].PortName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        portSet.AddPortRow(portRows[j]);
                    }
                    else
                    {
                        i = j - 1;
                        break;
                    }

                portMapSheet.AddPortSet(portSet);
            }

            return portMapSheet;
        }

        private void GetValueByInnerObjComplex(ExcelWorksheet sheet, ClassInnerObj innerObjComplex, int startIndex,
            int i, PortRow portRow)
        {
            foreach (var property in innerObjComplex.Property)
            {
                if (property.name.Equals("start", StringComparison.CurrentCulture))
                    startIndex = property.indexInSheet;

                var value = GetCellText(sheet, i, property.indexInSheet + StartColumnIndex - 1);
                var propertyInfo = portRow.GetType().GetProperty(property.name);
                if (value != null && propertyInfo != null)
                    propertyInfo.SetValue(portRow, value, null);

                if (property.name.Equals("end", StringComparison.CurrentCulture))
                {
                    var list = portRow.GetType().GetProperty(innerObjComplex.name);
                    var listValue = new List<string>();
                    for (var j = startIndex + StartColumnIndex - 1; j <= property.indexInSheet + StartColumnIndex - 1; j++)
                        listValue.Add(GetCellText(sheet, i, j));
                    if (value != null && list != null)
                        list.SetValue(portRow, listValue, null);
                }
            }

            if (innerObjComplex.InnerObjComplex != null)
                foreach (var childInnerObjComplex in innerObjComplex.InnerObjComplex)
                    GetValueByInnerObjComplex(sheet, childInnerObjComplex, startIndex, i, portRow);
        }

        private void GetValueByInnerObjComplex(Worksheet sheet, ClassInnerObj innerObjComplex, int startIndex, int i,
            PortRow portRow)
        {
            foreach (var property in innerObjComplex.Property)
            {
                if (property.name.Equals("start", StringComparison.CurrentCulture))
                    startIndex = property.indexInSheet;

                var value = GetCellText(sheet, i, property.indexInSheet + StartColumnIndex - 1);
                var propertyInfo = portRow.GetType().GetProperty(property.name);
                if (value != null && propertyInfo != null)
                    propertyInfo.SetValue(portRow, value, null);

                if (property.name.Equals("end", StringComparison.CurrentCulture))
                {
                    var list = portRow.GetType().GetProperty(innerObjComplex.name);
                    var listValue = new List<string>();
                    for (var j = startIndex + StartColumnIndex - 1; j <= property.indexInSheet + StartColumnIndex - 1; j++)
                        listValue.Add(GetCellText(sheet, i, j));
                    if (value != null && list != null)
                        list.SetValue(portRow, listValue, null);
                }
            }

            if (innerObjComplex.InnerObjComplex != null)
                foreach (var childInnerObjComplex in innerObjComplex.InnerObjComplex)
                    GetValueByInnerObjComplex(sheet, childInnerObjComplex, startIndex, i, portRow);
        }
    }
}