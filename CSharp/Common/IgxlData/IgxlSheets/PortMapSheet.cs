using IgxlData.IgxlBase;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class PortMapSheet : IgxlSheet
    {
        private const string SheetType = "DTPortMapSheet";
        public List<PortSet> PortSets { get; set; }

        public PortMapSheet(Worksheet sheet) : base(sheet)
        {
            PortSets = new List<PortSet>();
            IgxlSheetName = IgxlSheetNameList.PortMap;
        }

        public PortMapSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            PortSets = new List<PortSet>();
            IgxlSheetName = IgxlSheetNameList.PortMap;
        }

        public PortMapSheet(string sheetName)
            : base(sheetName)
        {
            PortSets = new List<PortSet>();
            IgxlSheetName = IgxlSheetNameList.PortMap;
        }

        protected void WriteHeader()
        {
            const string header =
                "DTPortMapSheet,version=2.0:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tPort Map\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t";
            IgxlWriter.WriteLine(header);
            IgxlWriter.WriteLine("\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t\t");
        }

        protected void WriteColumnsHeader()
        {
            const string columnsName = "\t\tProtocol\t\t\t\t\t\t\t\t\t\t\t\t\tFunction\t\t\t\t\t\t\t\t\t\t\t\t\t\t";
            IgxlWriter.WriteLine(columnsName);
            IgxlWriter.WriteLine(
                "\tPort Name\tFamily\tType\tSettings\tSetting0\tSetting1\tSetting2\tSetting3\tSetting4\tSetting5\tSetting6\tSetting7\tSetting8\tSetting9\tName\tPin\tProperties\tProperty0\tProperty1\tProperty2\tProperty3\tProperty4\tProperty5\tProperty6\tProperty7\tProperty8\tProperty9\tComment\t");
        }

        protected void WriteRows()
        {
            foreach (var portSet in PortSets)
                foreach (var port in portSet.PortRows)
                {
                    var row = new StringBuilder();
                    row.Append("\t");
                    row.Append(portSet.PortName);
                    row.Append("\t");
                    row.Append(port.ProtocolFamily);
                    row.Append("\t");
                    row.Append(port.ProtocolType);
                    row.Append("\t");
                    row.Append(port.ProtocolSettings);
                    row.Append("\t");
                    //foreach (var portSetting in port.SettingList)
                    //{
                    //    row.Append(portSetting);
                    //    row.Append("\t");
                    //}
                    for (var i = 0; i < PortRow.ConSettingNumber; i++)
                    {
                        row.Append(port.ProtocolSettingValues.Count > i ? port.ProtocolSettingValues[i] : "");
                        row.Append("\t");
                    }

                    row.Append(port.FunctionName);
                    row.Append("\t");
                    row.Append(port.FunctionPin);
                    row.Append("\t");
                    row.Append(port.FunctionProperties);
                    row.Append("\t");
                    //foreach (var portProperty in port.PropertyList)
                    //{
                    //    row.Append(portProperty);
                    //    row.Append("\t");
                    //}
                    for (var i = 0; i < PortRow.ConPropertyNumber; i++)
                    {
                        row.Append(port.FunctionPropertyValues.Count > i ? port.FunctionPropertyValues[i] : "");
                        row.Append("\t");
                    }

                    row.Append(port.Comment);
                    row.Append("\t");
                    IgxlWriter.WriteLine(row);
                }
        }

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
                version = "2.0";
            if (version == "2.0")
            {
                GetStreamWriter(fileName);
                WriteHeader();
                WriteColumnsHeader();
                WriteRows();
                CloseStreamWriter();
            }
            else
            {
                throw new Exception(string.Format("The PortMap sheet version:{0} is not supported!", version));
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (PortSets.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var portNameIndex = GetIndexFrom(igxlSheetsVersion, "Port Name");
                var familyIndex = GetIndexFrom(igxlSheetsVersion, "Protocol", "Family");
                var typeIndex = GetIndexFrom(igxlSheetsVersion, "Protocol", "Type");
                var settingsIndex = GetIndexFrom(igxlSheetsVersion, "Protocol", "Settings");
                var setting0Index = GetIndexFrom(igxlSheetsVersion, "Protocol", "Setting0");
                var nameIndex = GetIndexFrom(igxlSheetsVersion, "Function", "Name");
                var pinIndex = GetIndexFrom(igxlSheetsVersion, "Function", "Pin");
                var propertiesIndex = GetIndexFrom(igxlSheetsVersion, "Function", "Properties");
                var property0Index = GetIndexFrom(igxlSheetsVersion, "Function", "Property0");
                var commentIndex = GetIndexFrom(igxlSheetsVersion, "Comment");

                #region headers

                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" +
                             IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    SetColumns(igxlSheetsVersion, i, arr);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data

                for (var index = 0; index < PortSets.Count; index++)
                {
                    var portSet = PortSets[index];
                    for (var i = 0; i < portSet.PortRows.Count; i++)
                    {
                        var row = portSet.PortRows[i];
                        var arr = Enumerable.Repeat("", maxCount).ToArray();
                        if (!string.IsNullOrEmpty(portSet.PortName))
                        {
                            //arr[0] = row.ColumnA;
                            arr[portNameIndex] = portSet.PortName;
                            arr[familyIndex] = row.ProtocolFamily;
                            arr[typeIndex] = row.ProtocolType;
                            arr[settingsIndex] = row.ProtocolSettings;
                            for (var j = 0; j < row.ProtocolSettingValues.Count; j++)
                                if (j < 10)
                                    arr[setting0Index + j] = row.ProtocolSettingValues[j];
                            arr[nameIndex] = row.FunctionName;
                            arr[pinIndex] = row.FunctionPin;
                            arr[propertiesIndex] = row.FunctionProperties;
                            for (var j = 0; j < row.FunctionPropertyValues.Count; j++)
                                if (j < 10)
                                    arr[property0Index + j] = row.FunctionPropertyValues[j];
                            arr[commentIndex] = row.Comment;
                        }
                        else
                        {
                            arr = new[] { "\t" };
                        }

                        sw.WriteLine(string.Join("\t", arr));
                    }
                }

                #endregion
            }
        }

        public void AddPortSet(PortSet portSet)
        {
            PortSets.Add(portSet);
        }
    }
}