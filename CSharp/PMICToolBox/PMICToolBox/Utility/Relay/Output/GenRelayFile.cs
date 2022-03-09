using PmicAutomation.MyControls;
using PmicAutomation.Utility.Relay.Base;
using PmicAutomation.Utility.Relay.Input;
using Library.Function;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using PmicAutomation.MyControls;
using System.Text.RegularExpressions;

namespace PmicAutomation.Utility.Relay.Output
{
    public class GenRelayFile
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly string _outputPath;
        private const string Delimiter = "_To_";

        public GenRelayFile(MyForm.RichTextBoxAppend appendText, string outputPath)
        {
            _appendText = appendText;
            _outputPath = outputPath;
        }

        public void GenTraceBasedRelayControl(string fileName, List<string> resourcePins, List<string> devicePins,
            List<ComPinRow> comPinRows, List<AdgMatrix> adgMatrices, List<RelayItem> relayItems)
        {
            string file = Path.Combine(_outputPath, fileName);
            if (File.Exists(file))
            {
                File.Delete(file);
            }

            using (ExcelPackage package = new ExcelPackage(new FileInfo(file)))
            {
                ExcelWorkbook workBook = package.Workbook;
                PrintFilterPins(workBook, resourcePins, devicePins);
                PrintLinkedRows(workBook, comPinRows);
                if (adgMatrices.Any())
                    PrintAdg414Matrix(workBook, adgMatrices, relayItems);
                PrintPathSummary(workBook, relayItems);
                package.Save();
            }
        }

        private void PrintFilterPins(ExcelWorkbook workBook, List<string> resourcePins, List<string> devicePins)
        {
            _appendText.Invoke("Starting to generate FilterPins sheet ...", Color.Black);
            ExcelWorksheet sheet = workBook.AddSheet("FilterPins_AutoGen");
            sheet.Cells[1, 1].PrintExcelCol(new[] { "Resource Pin(Tree)" });
            sheet.Cells[2, 1].PrintExcelCol(resourcePins.ToArray());

            sheet.Cells[1, 2].PrintExcelCol(new[] { "Device Pin(Node)" });
            sheet.Cells[2, 2].PrintExcelCol(devicePins.ToArray());

            sheet.Column(1).AutoFit();
            sheet.Column(2).AutoFit();
        }

        private void PrintPathSummary(ExcelWorkbook workBook, List<RelayItem> relayItems)
        {
            _appendText.Invoke("Starting to generate PathSummary sheet ...", Color.Black);
            ExcelWorksheet sheet = workBook.AddSheet("PathSummary");

            object[,] array = new object[1, 2];
            array[0, 0] = "Resource Pin";
            array[0, 1] = "Device Pin";
            int cnt = 2;
            for (int i = 0; i < relayItems.Count(); i++)
            {
                for (int j = 0; j < relayItems[i].Paths.Count(); j++)
                {
                    List<string> list = new List<string>();
                    list.Add(relayItems[i].ResourcePin);
                    list.Add(relayItems[i].DevicePin);
                    List<string> path = relayItems[i].Paths[j];
                    path.Reverse();
                    list.AddRange(path);
                    sheet.Cells[cnt, 1].PrintExcelRow(list.ToArray());
                    cnt++;
                }
            }

            sheet.Cells[1, 1].PrintExcelRange(array);
            sheet.Cells["A:D"].AutoFitColumns();
            sheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        }

        private void PrintAdg414Matrix(ExcelWorkbook workBook, List<AdgMatrix> adgMatrices, List<RelayItem> relayItems)
        {
            List<ComPinRow> devices = adgMatrices.First().DevicePins.OrderByDescending(x => x.PinName).ToList();
            foreach (AdgMatrix adgMatrix in adgMatrices)
            {
                adgMatrix.DevicePins = adgMatrix.DevicePins.OrderByDescending(x => x.PinName).ToList();
            }

            _appendText.Invoke("Starting to generate ADG1414_Matrix sheet ...", Color.Black);
            ExcelWorksheet sheet = workBook.AddSheet("ADG1414_Matrix");

            const int rowNum = 11;
            object[,] array = new object[4 + rowNum * adgMatrices.Count, 4 + adgMatrices.First().ResourcePins.Count];
            array[0, 0] = "ADG_Matrix";
            array[1, 3] = "ADG IN";
            for (int j = 0; j < devices.Count; j++)
            {
                array[1, 4 + j] = devices[j].PinName;
            }

            for (int i = 0; i < adgMatrices.Count; i++)
            {
                array[2 + rowNum * i, 0] = "IN";
                for (int j = 0; j < adgMatrices[i].DevicePins.Count; j++)
                {
                    array[2 + rowNum * i, 4 + j] = adgMatrices[i].DevicePins[j].NetName.Length < 4 ?
                        adgMatrices[i].DevicePins[j].NetName :
                    RelayItem.GetPinName(adgMatrices[i].DevicePins[j].NetName);
                    array[3 + rowNum * i, 4 + j] = Math.Pow(2, adgMatrices[i].DevicePins.Count - j - 1);
                }

                array[4 + rowNum * i, 0] = i + 1;
                array[4 + rowNum * i, 1] = RelayItem.GetPinName(adgMatrices[i].Name);
                for (int j = 0; j < adgMatrices[i].ResourcePins.Count; j++)
                {
                    array[4 + rowNum * i + j, 2] = adgMatrices[i].ResourcePins[j].PinName;
                }

                for (int j = 0; j < adgMatrices[i].ResourcePins.Count; j++)
                {
                    array[4 + rowNum * i + j, 3] = adgMatrices[i].ResourcePins[j].NetName.Length < 4 ?
                        adgMatrices[i].ResourcePins[j].NetName + "_" +
                        adgMatrices[i].ResourcePins[j].PinName :
                        RelayItem.GetPinName(adgMatrices[i].ResourcePins[j].NetName) + "_" +
                                                   adgMatrices[i].ResourcePins[j].PinName;
                }

                object[,] relayArray = new object[adgMatrices[i].ResourcePins.Count, adgMatrices[i].DevicePins.Count];
                for (int j = 0; j < adgMatrices[i].ResourcePins.Count; j++)
                {
                    int k = adgMatrices[i].ResourcePins.Count - j - 1;
                    if (relayItems.Any(x =>
                        x.ResourcePin.Equals(adgMatrices[i].ResourcePins[j].NetName,
                            StringComparison.CurrentCultureIgnoreCase) &&
                        x.DevicePin.Equals(adgMatrices[i].DevicePins[k].NetName,
                            StringComparison.CurrentCultureIgnoreCase)))
                    {
                        List<string> relay = relayItems.Find(x =>
                            x.ResourcePin.Equals(adgMatrices[i].ResourcePins[j].NetName,
                                StringComparison.CurrentCultureIgnoreCase) &&
                            x.DevicePin.Equals(adgMatrices[i].DevicePins[k].NetName,
                                StringComparison.CurrentCultureIgnoreCase)).Relays;
                        if (relay.Any())
                            relayArray[j, k] = "v (" + string.Join(",", relay) + ")";
                    }
                }

                sheet.Cells[5 + rowNum * i, 5].PrintExcelRange(relayArray);
            }

            sheet.Cells[1, 1].PrintExcelRange(array);
            sheet.Cells["A:D"].AutoFitColumns();
            sheet.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        }

        private void PrintLinkedRows(ExcelWorkbook workBook, List<ComPinRow> comPinRows)
        {
            _appendText.Invoke("Starting to generate LinkedRows sheet ...", Color.Black);
            ExcelWorksheet sheet = workBook.AddSheet("LinkedRows");
            sheet.Cells[1, 1].LoadFromCollection(comPinRows, true);
        }

        public void GenIniFile(Dictionary<string, List<RelayPathRecord>> printDataListDic)
        {
            string outputPath = Path.Combine(_outputPath, "PinToPin");
            if (!Directory.Exists(outputPath))
            {
                Directory.CreateDirectory(outputPath);
            }

            foreach (KeyValuePair<string, List<RelayPathRecord>> printDataList in printDataListDic)
            {
                //printDataList.Value.Remove(printDataList.Value.Last());
                List<string> lines = new List<string>();
                for (int index = 0; index < printDataList.Value.Count; index++)
                {
                    lines.Add("[" + index + "]");
                    lines.Add("bOutput_empty=" + printDataList.Value[index].BOutputEmpty);
                    lines.Add("in=" + printDataList.Value[index].Input);
                    lines.Add("inType=" + printDataList.Value[index].InType);
                    lines.Add("output=" + printDataList.Value[index].Output);
                    lines.Add("outType=" + printDataList.Value[index].OutType);
                    lines.Add("SourceIndex=" + printDataList.Value[index].SourceIndex);
                }

                File.WriteAllLines(Path.Combine(outputPath, printDataList.Key + ".ini"), lines);
            }
        }

        public List<string> GenBasFile(List<RelayItem> relayItems, List<string> resourcePins, List<string> devicePins, List<AdgMatrix> adgMatrices)
        {
            List<string> lines = new List<string>();
            lines.AddRange(GenEnumToStringLines(relayItems));
            lines.AddRange(GenCaseLines(relayItems));
            lines.AddRange(GenEnumLines(relayItems));
            lines.AddRange(GenSinExtractLines(adgMatrices));
            return lines;
        }

        public void GenEnumToString(List<RelayItem> relayItems)
        {
            var lines = GenEnumToStringLines(relayItems);
            File.WriteAllLines(Path.Combine(_outputPath, "RelaysOnTrace_EnumToString.txt"), lines);
        }

        private static List<string> GenEnumToStringLines(List<RelayItem> relayItems)
        {
            List<string> lines = new List<string>
            {
                "Public Function TraceEnumToString(EnumInput As Trace) As String", "\tSelect Case EnumInput"
            };
            foreach (RelayItem data in relayItems)
            {
                var names = data.GetNames();
                if (names.Count() == 1)
                {
                    string enumInput = names[0];
                    lines.Add("\t\tCase\t" + enumInput + ":TraceEnumToString = \"" + enumInput + "\"");
                }
                else
                {
                    for (var j = 0; j < names.Count; j++)
                    {
                        string enumInput = data.GetResourcePin() + Delimiter + names[j] +
                                           Delimiter + data.GetDevicePin();
                        lines.Add("\t\tCase\t" + enumInput + ":TraceEnumToString = \"" + enumInput + "\"");
                    }
                }
            }

            lines.Add("\tEnd Select");
            lines.Add("Exit Function");
            lines.Add("End Function");
            return lines;
        }

        public void GenCase(List<RelayItem> relayItems)
        {
            var lines = GenCaseLines(relayItems);
            File.WriteAllLines(Path.Combine(_outputPath, "RelaysOnTrace_Case.txt"), lines);
        }

        private static List<string> GenCaseLines(List<RelayItem> relayItems)
        {
            List<string> lines = new List<string>
            {
                "Public Function Relay_Info(Trace As Double) As String",
                "Dim Relay As String",
                "Dim Concat As String",
                "Dim Which_ADG As String",
                "Dim SIN As String",
                "\tSelect Case Trace"
            };
            foreach (RelayItem data in relayItems)
            {
                var names = data.GetNames();
                if (names.Count() == 1)
                {
                    string enumInput = names[0];
                   
                    if (data.Adgs.Count != 0)
                    {

                        lines.Add("\t\tCase " + enumInput);
                        lines.Add("\t\t\tWhich_ADG = \"" + string.Join(",", data.Adgs.
                                      Select(x => x.Substring(0, x.LastIndexOf("_"))).Distinct()) + "\"");
                        lines.Add("\t\t\tSIN = \"" + string.Join(",", data.Adgs.Distinct()) + "\"");
                    }

                    var name = string.Join(",", data.Relays.Distinct());
                    if (!string.IsNullOrEmpty(name))
                    {
                        lines.Add("\t\tCase " + enumInput);
                        lines.Add("\t\t\t Relay=\"" + name + "\"");
                    }
                }
                else
                {
                    for (var j = 0; j < names.Count; j++)
                    {
                        string enumInput = data.GetResourcePin() + Delimiter + names[j] +
                                           Delimiter + data.GetDevicePin();

                        if (data.Adgs.Count != 0)
                        {
                            lines.Add("\t\tCase " + enumInput);
                            lines.Add("\t\t\tWhich_ADG = \"" + data.Adgs[j].Substring(0, data.Adgs[j].LastIndexOf("_")) + "\"");
                            lines.Add("\t\t\tSIN = \"" + data.Adgs[j] + "\"");

                        }

                        if (!string.IsNullOrEmpty(data.Relays[j]))
                        {
                            lines.Add("\t\tCase " + enumInput);
                            lines.Add("\t\t\t Relay=\"" + data.Relays[j] + "\"");
                        }
                    }
                }
            }

            lines.Add("\tEnd Select");
            lines.Add("Concat = Relay &\"&\"& Which_ADG &\"&\"& SIN");
            lines.Add("Relay_Info = Concat");
            lines.Add("Exit Function");
            lines.Add("End Function");
            return lines;
        }

        public void GenEnum(List<RelayItem> relayItems)
        {
            List<string> lines = GenEnumLines(relayItems);
            File.WriteAllLines(Path.Combine(_outputPath, "RelaysOnTrace_Enum.txt"), lines);
        }

        private List<string> GenEnumLines(List<RelayItem> relayItems)
        {
            List<string> lines = new List<string> { "Public Enum Trace" };
            foreach (var item in relayItems)
            {
                lines.Add("\t" + item.GetResourcePin() + Delimiter + item.GetDevicePin());
            }

            lines.Add("End Enum");
            return lines;
        }

        public void GenSinExtract(List<AdgMatrix> adgMatrices)
        {
            var lines = GenSinExtractLines(adgMatrices);
            File.WriteAllLines(Path.Combine(_outputPath, "SIN_Extract.txt"), lines);
        }

        private static List<string> GenSinExtractLines(List<AdgMatrix> adgMatrices)
        {
            List<string> lines = new List<string>
            {
                "Public Function SIN_Extract(Trace As Double) As String", "Select Case Trace"
            };
            foreach (AdgMatrix adgMatrix in adgMatrices)
            {
                lines.Add(@"'___" + RelayItem.GetPinName(adgMatrix.Name));
                for (int j = 0; j < adgMatrix.ResourcePins.Count; j++)
                {
                    if (!string.IsNullOrEmpty(adgMatrix.ResourcePins[j].NetName) &&
                        !string.IsNullOrEmpty(adgMatrix.DevicePins[j].NetName))
                    {
                        string path = RelayItem.GetPinName(adgMatrix.ResourcePins[j].NetName) + Delimiter +
                                      RelayItem.GetPinName(adgMatrix.DevicePins[j].NetName);
                        var arr = adgMatrix.Name.Split('_').ToList();
                        arr.RemoveAt(0);
                        path = string.Join("_", arr) + "_S" + (j + 1);
                        lines.Add("\tCase " + path + ":SIN_Extract = \"" +
                                  RelayItem.GetPinName(adgMatrix.ResourcePins[j].NetName) + "_" +
                                  adgMatrix.ResourcePins[j].PinName + "\"");
                    }
                }
            }

            lines.Add("End Select");
            lines.Add("Exit Function");
            lines.Add("End Function");
            return lines;
        }
    }
}