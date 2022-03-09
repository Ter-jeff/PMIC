using CommonLib.EpplusErrorReport;
using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.Others.BinCutELB
{
    public class InstanceMappingRow
    {
        public List<string> BinCutInstanceNames = new List<string>();
        public string HardipInstanceName;
        public bool IsFound;
        public InstanceRow BinCutRow;
    }

    public class BinCutElbMain
    {
        public void WorkFlow(string trunkPath, string ouputFolderNew, string ouputFolderOld, InstanceSheet testInstSheet, List<SubFlowSheet> flowSheets,
            ref Dictionary<string, SubFlowSheet> subFlowSheets, ref Dictionary<string, InstanceSheet> instanceSheets, ref List<InstanceSheet> instanceHadripSheets)
        {
            var instanceMappingRows = GetInstanceMappings(testInstSheet);
            var igxlSheetReader = new IgxlSheetReader();
            var instanceSheetNames = igxlSheetReader.GetSheetsByType(trunkPath, SheetType.DTTestInstancesSheet);
            foreach (var instanceSheetName in instanceSheetNames)
            {
                var readInstanceSheet = new ReadInstanceSheet();
                var instanceSheet = readInstanceSheet.GetSheet(instanceSheetName);
                var oldSheet = instanceSheet.DeepClone();
                var flag = false;
                foreach (var row in instanceSheet.InstanceRows)
                {
                    var instanceName = row.TestName.ToUpper();
                    if (instanceMappingRows.Exists(x => x.HardipInstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var instanceMappingRow = instanceMappingRows.Find(x => x.HardipInstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase));
                        row.Overlay = GetOverlay(row, instanceMappingRow.HardipInstanceName);
                        flag = true;
                    }
                }
                if (flag)
                {
                    var name = Path.GetFileNameWithoutExtension(instanceSheetName);
                    instanceHadripSheets.Add(instanceSheet);
                    instanceSheets.Add(Path.Combine(ouputFolderNew, name + ".txt"), instanceSheet);
                    instanceSheets.Add(Path.Combine(ouputFolderOld, name + ".txt"), oldSheet);
                }
            }

            if (instanceMappingRows.Count == 0)
            {
                return;
            }

            foreach (var instanceMappingRow in instanceMappingRows)
            {
                if (instanceMappingRow.IsFound == false)
                {
                    var errMsg = string.Format("The ELB instance {0} can not be found !!!", instanceMappingRow.HardipInstanceName);
                    EpplusErrorManager.AddError(BasicErrorType.FormatWarning.ToString(), testInstSheet.Name, instanceMappingRow.BinCutRow.RowNum, errMsg);
                }
            }

            #region Get flow sheet
            var flowSheetFiles = igxlSheetReader.GetSheetsByType(trunkPath, SheetType.DTFlowtableSheet);
            var hardIpFlowSheets = new List<SubFlowSheet>();
            foreach (var flowsheetFile in flowSheetFiles)
            {
                var name = Path.GetFileNameWithoutExtension(flowsheetFile);
                if (name.StartsWith("Flow_HARDIP_", StringComparison.CurrentCulture))
                {
                    var readFlowSheet = new ReadFlowSheet();
                    var subFlowSheet = readFlowSheet.GetSheet(flowsheetFile);
                    hardIpFlowSheets.Add(subFlowSheet);
                }
            }

            var binCutFlowSheets = flowSheets.Where(x => x.Name.EndsWith("_TD_Mbist_BV", StringComparison.CurrentCulture) ||
                x.Name.StartsWith("Flow_VddBinning", StringComparison.CurrentCulture)).ToList();
            #endregion

            var limitDic = GetHardipLimit(instanceMappingRows, hardIpFlowSheets);

            subFlowSheets = ModifyBinCutFlow(ouputFolderNew, ouputFolderOld, binCutFlowSheets, instanceMappingRows, limitDic);
        }

        public void WorkFlow(Workbook workbook, string ouputFolderNew, string ouputFolderOld)
        {
            Application app = workbook.Parent;
            foreach (_Worksheet wroksheet in workbook.Worksheets)
            {
                if (wroksheet.Name.StartsWith("TestInst_", StringComparison.CurrentCultureIgnoreCase) ||
                    wroksheet.Name.StartsWith("Flow_", StringComparison.CurrentCultureIgnoreCase))
                {
                    app.StatusBar = string.Format("{0} activate ...", wroksheet.Name);
                    wroksheet.Activate();
                }
            }

            var sheet = workbook.GetSheet("TestInst_Vddbinning");
            if (sheet != null)
            {
                var instanceMappingRows = GetInstanceMappings(app, sheet);
                var igxlSheetReader = new IgxlSheetReader();
                var instanceSheetNames = igxlSheetReader.GetSheetsByType(workbook, SheetType.DTTestInstancesSheet);
                foreach (var instanceSheetName in instanceSheetNames)
                {
                    app.StatusBar = string.Format("Reading {0} ...", instanceSheetName);
                    var readInstanceSheet = new ReadInstanceSheet();
                    InstanceSheet instanceSheet = readInstanceSheet.GetSheet(workbook.Worksheets[instanceSheetName]);
                    var old = instanceSheet.DeepClone();
                    var flag = false;
                    foreach (var row in instanceSheet.InstanceRows)
                    {
                        var instanceName = row.TestName.ToUpper();
                        if (instanceMappingRows.Exists(x => x.HardipInstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase)))
                        {
                            var instanceMappingRow = instanceMappingRows.Find(x => x.HardipInstanceName.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase));
                            row.Overlay = GetOverlay(row, instanceMappingRow.HardipInstanceName);
                            flag = true;
                        }
                    }
                    if (flag)
                    {
                        instanceSheet.Write(Path.Combine(ouputFolderNew, instanceSheetName + ".txt"));
                        old.Write(Path.Combine(ouputFolderOld, instanceSheetName + ".txt"));
                    }
                }


                if (instanceMappingRows.Count == 0)
                {
                    return;
                }

                foreach (var instanceMappingRow in instanceMappingRows)
                {
                    if (instanceMappingRow.IsFound == false)
                    {
                        var errMsg = string.Format("The ELB instance {0} can not be found !!!", instanceMappingRow.HardipInstanceName);
                        EpplusErrorManager.AddError(BasicErrorType.FormatWarning.ToString(), sheet.Name, instanceMappingRow.BinCutRow.RowNum, errMsg);
                    }
                }

                #region Get flow sheet
                var allFlowSheets = igxlSheetReader.GetSheetsByType(workbook, SheetType.DTFlowtableSheet);
                var hadipFlowSheets = new List<SubFlowSheet>();
                var binCutFlowSheets = new List<SubFlowSheet>();
                foreach (var flowsheet in allFlowSheets)
                {
                    app.StatusBar = string.Format("Reading {0} ...", flowsheet);
                    if (flowsheet.StartsWith("Flow_HARDIP_", StringComparison.CurrentCulture))
                    {
                        var readFlowSheet = new ReadFlowSheet();
                        SubFlowSheet subFlowSheet = readFlowSheet.GetSheet(workbook.Worksheets[flowsheet]);
                        hadipFlowSheets.Add(subFlowSheet);
                    }
                    else if (flowsheet.EndsWith("_TD_Mbist_BV", StringComparison.CurrentCulture))
                    {
                        var readFlowSheet = new ReadFlowSheet();
                        workbook.Worksheets[flowsheet].Activate();
                        workbook.Worksheets[flowsheet].Select();
                        SubFlowSheet subFlowSheet = readFlowSheet.GetSheet(workbook.Worksheets[flowsheet]);
                        binCutFlowSheets.Add(subFlowSheet);
                    }
                    else if (flowsheet.StartsWith("Flow_VddBinning", StringComparison.CurrentCulture))
                    {
                        var readFlowSheet = new ReadFlowSheet();
                        workbook.Worksheets[flowsheet].Activate();
                        workbook.Worksheets[flowsheet].Select();
                        SubFlowSheet subFlowSheet = readFlowSheet.GetSheet(workbook.Worksheets[flowsheet]);
                        binCutFlowSheets.Add(subFlowSheet);
                    }
                }

                #endregion

                var limitDic = GetHardipLimit(instanceMappingRows, hadipFlowSheets);

                var subFlowSheets = ModifyBinCutFlow(ouputFolderNew, ouputFolderOld, binCutFlowSheets, instanceMappingRows, limitDic);

                foreach (var subFlowSheet in subFlowSheets)
                {
                    var version = subFlowSheet.Value.GetVersion();
                    subFlowSheet.Value.Write(subFlowSheet.Key, version);
                }
            }
        }

        private string GetOverlay(InstanceRow row, string hardipInstanceName)
        {
            var overlay = row.Overlay;
            var newOverlay = "Overlay_BV_" + hardipInstanceName;
            if (string.IsNullOrEmpty(overlay))
                overlay = newOverlay;
            else
            {
                var arr = overlay.Split(',').ToList();
                if (!arr.Exists(x => x.Equals(newOverlay, StringComparison.CurrentCultureIgnoreCase)))
                {
                    arr.Add(newOverlay);
                    overlay = string.Join(",", arr);
                }
            }
            return overlay;
        }

        private Dictionary<string, SubFlowSheet> ModifyBinCutFlow(string ouputFolderNew, string ouputFolderOld, List<SubFlowSheet> binCutSheets, List<InstanceMappingRow> instanceMappingRows, Dictionary<string, List<FlowRow>> limitDic)
        {
            var subFlowSheets = new Dictionary<string, SubFlowSheet>();
            foreach (var binCutSheet in binCutSheets)
            {
                var oldSheet = binCutSheet.DeepClone();
                var flag = false;
                for (var index = 0; index < binCutSheet.FlowRows.Count; index++)
                {
                    var row = binCutSheet.FlowRows[index];
                    if (row.Opcode.Equals("Test", StringComparison.CurrentCultureIgnoreCase) ||
                        row.Opcode.Equals("Nop", StringComparison.CurrentCultureIgnoreCase))
                    {
                        if (row.Parameter.ToUpper().Contains("_CALLINST"))
                        {

                            if (string.IsNullOrEmpty(row.FailAction))
                                row.FailAction = "F_BV_CALLINST";
                            else
                            {
                                var arr = row.FailAction.Split(',').ToList();
                                if (!arr.Exists(x => x.Equals("F_BV_CALLINST", StringComparison.CurrentCultureIgnoreCase)))
                                    row.FailAction = row.FailAction + ",F_BV_CALLINST";
                            }
                        }
                        if (instanceMappingRows.Exists(x => x.BinCutInstanceNames.Exists(y => y.Equals(row.Parameter, StringComparison.CurrentCultureIgnoreCase))))
                        {
                            var hardipInstance = instanceMappingRows.Find(x => x.BinCutInstanceNames.Exists(y => y.Equals(row.Parameter, StringComparison.CurrentCultureIgnoreCase)));
                            if (limitDic.ContainsKey(hardipInstance.HardipInstanceName))
                            {
                                var copyRows = new List<FlowRow>();
                                foreach (var limitRow in limitDic[hardipInstance.HardipInstanceName])
                                {
                                    var copyRow = limitRow.DeepClone();
                                    copyRow.Parameter = row.Parameter;
                                    copyRow.FailAction = "F_BV_CALLINST";
                                    copyRows.Add(copyRow);
                                }
                                binCutSheet.RemoveLimitRows(index);
                                binCutSheet.InsertRow(index + 1, copyRows);
                                flag = true;
                            }
                        }
                    }
                }

                if (flag)
                {
                    subFlowSheets.Add(Path.Combine(ouputFolderNew, binCutSheet.Name + ".txt"), binCutSheet);
                    subFlowSheets.Add(Path.Combine(ouputFolderOld, binCutSheet.Name + ".txt"), oldSheet);
                }
            }
            return subFlowSheets;
        }

        private Dictionary<string, List<FlowRow>> GetHardipLimit(List<InstanceMappingRow> instanceMappingRows, List<SubFlowSheet> hadipSheets)
        {
            var limitDic = new Dictionary<string, List<FlowRow>>();
            var instanceNames = instanceMappingRows.Select(x => x.HardipInstanceName).Distinct().ToList();
            foreach (var instanceName in instanceNames)
            {
                foreach (var hadipSheet in hadipSheets)
                {
                    if (hadipSheet.FlowRows.Exists(x => x.Parameter.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var index = hadipSheet.FlowRows.FindIndex(x => x.Parameter.Equals(instanceName, StringComparison.CurrentCultureIgnoreCase) &&
                                     (x.Opcode.Equals("Test", StringComparison.CurrentCultureIgnoreCase) ||
                                      x.Opcode.Equals("Nop", StringComparison.CurrentCultureIgnoreCase)));
                        if (index != -1)
                        {
                            if (!limitDic.ContainsKey(instanceName))
                            {
                                var flowRows = hadipSheet.GetLimitRows(index);
                                limitDic.Add(instanceName, flowRows);
                            }
                            else
                            {
                                var errMsg = string.Format("The limit of ELB instance are duplicated -{0} !!!", instanceName);
                                EpplusErrorManager.AddError(BasicErrorType.FormatWarning.ToString(), hadipSheet.Name,
                                    int.Parse(hadipSheet.FlowRows[index].LineNum), errMsg);
                            }
                        }
                    }
                }
            }
            return limitDic;
        }

        private List<InstanceMappingRow> GetInstanceMappings(InstanceSheet instanceSheet)
        {
            var instanceRows = new List<InstanceRow>();
            foreach (var row in instanceSheet.InstanceRows)
            {
                if (row.Name.Equals("GradeSearch_CallInstance_VT", StringComparison.CurrentCultureIgnoreCase) ||
                    row.Name.Equals("GradeSearch_HVCC_CallInstance_VT", StringComparison.CurrentCultureIgnoreCase))
                    instanceRows.Add(row);
            }

            var instanceMappings = new List<InstanceMappingRow>();
            foreach (var instanceRow in instanceRows)
            {
                if (instanceRow.Args.Count > 5)
                {
                    var hardipInstanceName = instanceRow.Args[4].ToUpper();
                    if (string.IsNullOrEmpty(hardipInstanceName))
                        continue;

                    if (instanceMappings.Exists(x => x.HardipInstanceName.Equals(hardipInstanceName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var instanceMappingRow = instanceMappings.Find(x => x.HardipInstanceName.Equals(hardipInstanceName, StringComparison.CurrentCultureIgnoreCase));
                        instanceMappingRow.BinCutInstanceNames.Add(instanceRow.TestName);
                        instanceMappingRow.BinCutRow = instanceRow;
                        instanceMappingRow.IsFound = true;
                    }
                    else
                    {
                        var instanceMappingRow = new InstanceMappingRow();
                        instanceMappingRow.HardipInstanceName = hardipInstanceName;
                        instanceMappingRow.BinCutInstanceNames.Add(instanceRow.TestName);
                        instanceMappingRow.BinCutRow = instanceRow;
                        instanceMappingRow.IsFound = true;
                        instanceMappings.Add(instanceMappingRow);
                    }
                }
            }
            return instanceMappings;
        }

        private List<InstanceMappingRow> GetInstanceMappings(Application app, Worksheet sheet)
        {
            app.StatusBar = "Reading TestInst_Vddbinning";
            var readInstanceSheet = new ReadInstanceSheet();
            var instanceSheet = readInstanceSheet.GetSheet(sheet);
            return GetInstanceMappings(instanceSheet);
        }
    }
}