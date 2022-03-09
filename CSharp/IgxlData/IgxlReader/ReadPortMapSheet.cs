using System;
using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlReader
{
    public class ReadPortMapSheet : IgxlSheetReader
    {
        private readonly SheetObjMap _sheetObjMap;

        #region private variable
        private const int StartRow = 5;
        private const int StartCol = 1;

        public ReadPortMapSheet(SheetObjMap sheetObjMap)
        {
            _sheetObjMap = sheetObjMap;
        }
        #endregion

        #region public Function
        public PortMapSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public PortMapSheet GetSheet(ExcelWorksheet sheet)
        {
            PortMapSheet portMapSheet = new PortMapSheet(sheet);
            int stopRow = sheet.Dimension.End.Row;
            List<PortRow> portRows = new List<PortRow>();
            for (int i = StartRow; i <= stopRow; i++)
            {
                PortRow portRow = new PortRow();
                portRow.RowNum = i;
                foreach (var innerObj in _sheetObjMap.InnerObj)
                {
                    foreach (var property in innerObj.Property)
                    {
                        var value = GetCellText(sheet, i, property.indexInSheet + StartCol);
                        var propertyInfo = portRow.GetType().GetProperty(property.name);
                        if (value != null && propertyInfo != null)
                            propertyInfo.SetValue(portRow, value, null);
                    }

                    const int startIndex = 0;
                    foreach (var innerObjComplex in innerObj.InnerObjComplex)
                    {
                        GetValueByInnerObjComplex(sheet, innerObjComplex, startIndex, i, portRow);
                    }
                }
                portRows.Add(portRow);
            }

            for (int i = 0; i < portRows.Count; i++)
            {
                var portRow = portRows[i];
                PortSet portSet = new PortSet(portRow.PortName);
                portSet.AddPortRow(portRow);
                for (int j = i + 1; j < portRows.Count; j++)
                {
                    if (portRow.PortName.Equals(portRows[j].PortName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        portSet.AddPortRow(portRows[j]);
                    }
                    else
                    {
                        i = j - 1;
                        break;
                    }
                }
                portMapSheet.AddPortSet(portSet);
            }
            return portMapSheet;
        }

        public PortMapSheet GetSheet(Worksheet sheet)
        {
            PortMapSheet portMapSheet = new PortMapSheet(sheet);
            int stopRow = sheet.UsedRange.Row + sheet.UsedRange.Rows.Count - 1;
            List<PortRow> portRows = new List<PortRow>();
            for (int i = StartRow; i <= stopRow; i++)
            {
                PortRow portRow = new PortRow();
                portRow.RowNum = i;
                foreach (var innerObj in _sheetObjMap.InnerObj)
                {
                    foreach (var property in innerObj.Property)
                    {
                        var value = GetCellText(sheet, i, property.indexInSheet + StartCol);
                        var propertyInfo = portRow.GetType().GetProperty(property.name);
                        if (value != null && propertyInfo != null)
                            propertyInfo.SetValue(portRow, value, null);
                    }

                    const int startIndex = 0;
                    foreach (var innerObjComplex in innerObj.InnerObjComplex)
                    {
                        GetValueByInnerObjComplex(sheet, innerObjComplex, startIndex, i, portRow);
                    }
                }
                portRows.Add(portRow);
            }

            for (int i = 0; i < portRows.Count; i++)
            {
                var portRow = portRows[i];
                PortSet portSet = new PortSet(portRow.PortName);
                portSet.AddPortRow(portRow);
                for (int j = i + 1; j < portRows.Count; j++)
                {
                    if (portRow.PortName.Equals(portRows[j].PortName, StringComparison.CurrentCultureIgnoreCase))
                    {
                        portSet.AddPortRow(portRows[j]);
                    }
                    else
                    {
                        i = j - 1;
                        break;
                    }
                }
                portMapSheet.AddPortSet(portSet);
            }
            return portMapSheet;
        }

        private void GetValueByInnerObjComplex(ExcelWorksheet sheet, ClassInnerObj innerObjComplex, int startIndex, int i, PortRow portRow)
        {
            foreach (var property in innerObjComplex.Property)
            {
                if (property.name.Equals("start", StringComparison.CurrentCulture))
                    startIndex = property.indexInSheet;

                var value = GetCellText(sheet, i, property.indexInSheet + StartCol);
                var propertyInfo = portRow.GetType().GetProperty(property.name);
                if (value != null && propertyInfo != null)
                    propertyInfo.SetValue(portRow, value, null);

                if (property.name.Equals("end", StringComparison.CurrentCulture))
                {
                    var list = portRow.GetType().GetProperty(innerObjComplex.name);
                    List<string> listValue = new List<string>();
                    for (int j = startIndex + StartCol; j <= property.indexInSheet + StartCol; j++)
                        listValue.Add(GetCellText(sheet, i, j));
                    if (value != null && list != null)
                        list.SetValue(portRow, listValue, null);
                }
            }
            if (innerObjComplex.InnerObjComplex != null)
                foreach (var childInnerObjComplex in innerObjComplex.InnerObjComplex)
                    GetValueByInnerObjComplex(sheet, childInnerObjComplex, startIndex, i, portRow);
        }

        private void GetValueByInnerObjComplex(Worksheet sheet, ClassInnerObj innerObjComplex, int startIndex, int i, PortRow portRow)
        {
            foreach (var property in innerObjComplex.Property)
            {
                if (property.name.Equals("start", StringComparison.CurrentCulture))
                    startIndex = property.indexInSheet;

                var value = GetCellText(sheet, i, property.indexInSheet + StartCol);
                var propertyInfo = portRow.GetType().GetProperty(property.name);
                if (value != null && propertyInfo != null)
                    propertyInfo.SetValue(portRow, value, null);

                if (property.name.Equals("end", StringComparison.CurrentCulture))
                {
                    var list = portRow.GetType().GetProperty(innerObjComplex.name);
                    List<string> listValue = new List<string>();
                    for (int j = startIndex + StartCol; j <= property.indexInSheet + StartCol; j++)
                        listValue.Add(GetCellText(sheet, i, j));
                    if (value != null && list != null)
                        list.SetValue(portRow, listValue, null);
                }
            }
            if (innerObjComplex.InnerObjComplex != null)
                foreach (var childInnerObjComplex in innerObjComplex.InnerObjComplex)
                    GetValueByInnerObjComplex(sheet, childInnerObjComplex, startIndex, i, portRow);
        }

        #endregion
    }
}
