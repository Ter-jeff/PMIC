using System.Linq;
using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using Teradyne.Oasis.IGLinkBase.ProgramGeneration;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class SubFlowSheet : FlowSheet
    {
        private const string SheetType = "DTFlowtableSheet";
        public const string Ttime = "TTIME";
        #region Constructor
        public SubFlowSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
        }

        public SubFlowSheet(string sheetName)
            : base(sheetName)
        {
        }

        #endregion

        #region Member Function
        public List<FlowRow> GetFlowTestGroup(int flowIndex)
        {
            List<FlowRow> flowRows = new List<FlowRow>();
            bool getFirst = false;
            int groupEnd = 0;
            for (int idx = flowIndex; idx < FlowRows.Count; idx++)
            {
                if (FlowRows[idx].OpCode.Equals("Test"))
                {
                    if (getFirst)
                    {
                        groupEnd = idx;
                        break;
                    }
                    getFirst = true;
                }

                // the last row
                if ((idx == FlowRows.Count - 1) && getFirst)
                {
                    groupEnd = FlowRows.Count;
                    break;
                }
            }

            for (int i = flowIndex; i < groupEnd; i++)
            {
                flowRows.Add(FlowRows[i]);
            }
            return flowRows;
        }
        public override void Write(string fileName, string version)
        {
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version == "2.3")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (version == "3.0")
                {
                    var igxlSheetsVersion = dic["3.0"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }


        }

        protected override void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (FlowRows.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var labelIndex = GetIndexFrom(igxlSheetsVersion, "Label");
                var enableIndex = GetIndexFrom(igxlSheetsVersion, "Enable");
                var jobIndex = GetIndexFrom(igxlSheetsVersion, "Gate", "Job");
                var partIndex = GetIndexFrom(igxlSheetsVersion, "Gate", "Part");
                var envIndex = GetIndexFrom(igxlSheetsVersion, "Gate", "Env");
                var opcodeIndex = GetIndexFrom(igxlSheetsVersion, "Command", "Opcode");
                var parameterIndex = GetIndexFrom(igxlSheetsVersion, "Command", "Parameter");
                var tNameIndex = GetIndexFrom(igxlSheetsVersion, "TName");
                var tNumIndex = GetIndexFrom(igxlSheetsVersion, "TNum");
                var loLimIndex = GetIndexFrom(igxlSheetsVersion, "Limits", "LoLim");
                var hiLimIndex = GetIndexFrom(igxlSheetsVersion, "Limits", "HiLim");
                var scaleIndex = GetIndexFrom(igxlSheetsVersion, "Datalog Display Results", "Scale");
                var unitsIndex = GetIndexFrom(igxlSheetsVersion, "Datalog Display Results", "Units");
                var formatIndex = GetIndexFrom(igxlSheetsVersion, "Datalog Display Results", "Format");
                var binNumberPassIndex = GetIndexFrom(igxlSheetsVersion, "Bin Number", "Pass");
                var binNumberFailIndex = GetIndexFrom(igxlSheetsVersion, "Bin Number", "Fail");
                var sortNumberPassIndex = GetIndexFrom(igxlSheetsVersion, "Sort Number", "Pass");
                var sortNumberFailIndex = GetIndexFrom(igxlSheetsVersion, "Sort Number", "Fail");
                var resultIndex = GetIndexFrom(igxlSheetsVersion, "Result");
                var actionPassIndex = GetIndexFrom(igxlSheetsVersion, "Action", "Pass");
                var actionFailIndex = GetIndexFrom(igxlSheetsVersion, "Action", "Fail");
                var stateIndex = GetIndexFrom(igxlSheetsVersion, "State");
                var groupSpecifierIndex = GetIndexFrom(igxlSheetsVersion, "Group", "Specifier");
                var groupSenseIndex = GetIndexFrom(igxlSheetsVersion, "Group", "Sense");
                var groupConditionIndex = GetIndexFrom(igxlSheetsVersion, "Group", "Condition");
                var groupNameIndex = GetIndexFrom(igxlSheetsVersion, "Group", "Name");
                var deviceSenseIndex = GetIndexFrom(igxlSheetsVersion, "Device", "Sense");
                var deviceConditionIndex = GetIndexFrom(igxlSheetsVersion, "Device", "Condition");
                var deviceNameIndex = GetIndexFrom(igxlSheetsVersion, "Device", "Name");
                var debugAssumeIndex = GetIndexFrom(igxlSheetsVersion, "Debug", "Assume");
                var debugSitesIndex = GetIndexFrom(igxlSheetsVersion, "Debug", "Sites");
                var elapsedTimeIndex = GetIndexFrom(igxlSheetsVersion, "CT Profile Data", "Elapsed Time (s)");
                var backgroundTypeIndex = GetIndexFrom(igxlSheetsVersion, "CT Profile Data", "Background Type");
                var serializeIndex = GetIndexFrom(igxlSheetsVersion, "CT Profile Data", "Serialize");
                var resourceLockIndex = GetIndexFrom(igxlSheetsVersion, "CT Profile Data", "Resource Lock");
                var flowStepLockedIndex = GetIndexFrom(igxlSheetsVersion, "CT Profile Data", "Flow Step Locked");
                var commentIndex = GetIndexFrom(igxlSheetsVersion, "Comment");


                #region headers
                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    SetColumns(igxlSheetsVersion, i, arr);

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data
                for (var index = 0; index < FlowRows.Count; index++)
                {
                    var row = FlowRows[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    arr[0] = row.ColumnA;
                    arr[labelIndex] = row.Label;
                    arr[enableIndex] = row.Enable;
                    arr[jobIndex] = row.Job;
                    arr[partIndex] = row.Part;
                    arr[envIndex] = row.Env;
                    arr[opcodeIndex] = row.OpCode;
                    arr[parameterIndex] = row.Parameter;
                    arr[tNameIndex] = row.Name;
                    arr[tNumIndex] = row.Num;
                    arr[loLimIndex] = row.LoLim;
                    arr[hiLimIndex] = row.HiLim;
                    arr[scaleIndex] = row.Scale;
                    arr[unitsIndex] = row.Units;
                    arr[formatIndex] = row.Format;
                    arr[binNumberPassIndex] = row.BinPass;
                    arr[binNumberFailIndex] = row.BinFail;
                    arr[sortNumberPassIndex] = row.SortPass;
                    arr[sortNumberFailIndex] = row.SortFail;
                    arr[resultIndex] = row.Result;
                    arr[actionPassIndex] = row.PassAction;
                    arr[actionFailIndex] = row.FailAction;
                    arr[stateIndex] = row.State;
                    arr[groupSpecifierIndex] = row.GroupSpecifier;
                    arr[groupSenseIndex] = row.GroupSense;
                    arr[groupConditionIndex] = row.GroupCondition;
                    arr[groupNameIndex] = row.GroupName;
                    arr[deviceSenseIndex] = row.DeviceSense;
                    arr[deviceConditionIndex] = row.DeviceCondition;
                    arr[deviceNameIndex] = row.DeviceName;
                    arr[debugAssumeIndex] = row.DebugAsume;
                    arr[debugSitesIndex] = row.DebugSites;
                    arr[elapsedTimeIndex] = row.CtProfileDataElapsedTime;
                    arr[backgroundTypeIndex] = row.CtProfileDataBackgroundType;
                    arr[serializeIndex] = row.CtProfileDataSerialize;
                    arr[resourceLockIndex] = row.CtProfileDataResourceLock;
                    arr[flowStepLockedIndex] = row.CtProfileDataFlowStepLocked;
                    arr[commentIndex] = row.Comment;

                    if (string.IsNullOrEmpty(row.Comment1))
                        sw.WriteLine(string.Join("\t", arr));
                    else
                        sw.WriteLine(string.Join("\t", arr) + "\t" + row.Comment1);
                }
                #endregion
            }
        }

        public void WriteOld(string fileName, string version)
        {
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                return;
            using (var sw = new StreamWriter(fileName, false))
            {
                sw.WriteLine("DTFlowtableSheet,version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tFlow Table");
                sw.WriteLine("\t\t\t\t\t\tFlow Domain:");
                sw.WriteLine("\t\t\tGate\t\t\tCommand\t\t\t\tLimits\t\tDatalog Display Results\t\t\tBin Number\t\tSort Number\t\t\tAction\t\t\tGroup\t\t\t\tDevice\t\t\tDebug\t\tCT Profile Data");
                sw.WriteLine("\tLabel\tEnable\tJob\tPart\tEnv\tOpcode\tParameter\tTName\tTNum\tLoLim\tHiLim\tScale\tUnits\tFormat\tPass\tFail\tPass\tFail\tResult\tPass\tFail\tState\tSpecifier\tSense\tCondition\tName\tSense\tCondition\tName\tAssume\tSites\tElapsed Time (s)\tBackground Type\tSerialize\tResource Lock\tFlow Step Locked\tComment");

                string[] arr = new string[38];
                foreach (FlowRow fr in FlowRows)
                {
                    arr[0] = fr.ColumnA;
                    arr[1] = fr.Label;
                    arr[2] = fr.Enable;
                    arr[3] = fr.Job;
                    arr[4] = fr.Part;
                    arr[5] = fr.Env;
                    arr[6] = fr.OpCode;
                    arr[7] = fr.Parameter;
                    arr[8] = fr.Name;
                    arr[9] = fr.Num;
                    arr[10] = fr.LoLim;
                    arr[11] = fr.HiLim;
                    arr[12] = fr.Scale;
                    arr[13] = fr.Units;
                    arr[14] = fr.Format;
                    arr[15] = fr.BinPass;
                    arr[16] = fr.BinFail;
                    arr[17] = fr.SortPass;
                    arr[18] = fr.SortFail;
                    arr[19] = fr.Result;
                    arr[20] = fr.PassAction;
                    arr[21] = fr.FailAction;
                    arr[22] = fr.State;
                    arr[23] = fr.GroupSpecifier;
                    arr[24] = fr.GroupSense;
                    arr[25] = fr.GroupCondition;
                    arr[26] = fr.GroupName;
                    arr[27] = fr.DeviceSense;
                    arr[28] = fr.DeviceCondition;
                    arr[29] = fr.DeviceName;
                    arr[30] = fr.DebugAsume;
                    arr[31] = fr.DebugSites;
                    arr[32] = fr.CtProfileDataElapsedTime;
                    arr[33] = fr.CtProfileDataBackgroundType;
                    arr[34] = fr.CtProfileDataSerialize;
                    arr[35] = fr.CtProfileDataResourceLock;
                    arr[36] = fr.CtProfileDataFlowStepLocked;
                    arr[37] = fr.Comment;
                    if (string.IsNullOrEmpty(fr.Comment1))
                        sw.WriteLine(string.Join("\t", arr));
                    else
                        sw.WriteLine(string.Join("\t", arr) + "\t" + fr.Comment1);
                }
            }
        }

        public void ReplaceParameter(string oldFile, string newFile, Dictionary<string, string> instanceReplaceDictionary)
        {
            using (var sr = new StreamReader(oldFile))
            using (var sw = new StreamWriter(newFile, false))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    string[] spt = Regex.Split(line, @"\t");
                    if (spt.Length < 8)
                        sw.WriteLine(line);
                    else
                    {
                        if (instanceReplaceDictionary.ContainsKey(spt[7].ToUpper()))
                            spt[7] = instanceReplaceDictionary[spt[7].ToUpper()];
                        string newText = string.Join("\t", spt);
                        sw.WriteLine(newText);
                    }
                }
            }
        }

        public void ReplaceOpCode(string oldFile, string newFile, List<string> nopInstances)
        {
            if (oldFile == newFile)
            {
                List<string> strings = new List<string>();
                using (var sr = new StreamReader(oldFile))
                {
                    while (!sr.EndOfStream)
                    {
                        var line = sr.ReadLine();
                        string[] spt = Regex.Split(line, @"\t");
                        if (spt.Length < 8)
                        {
                            strings.Add(line);
                        }
                        else
                        {
                            if (nopInstances.Exists(x => x.Equals(spt[7], StringComparison.OrdinalIgnoreCase)))
                                spt[6] = "nop";
                            string newText = string.Join("\t", spt);
                            strings.Add(newText);
                        }
                    }
                }
                using (var sw = new StreamWriter(newFile, false))
                {
                    foreach (var data in strings)
                    {
                        sw.WriteLine(data);
                    }
                }
            }
            else
            {
                using (var sr = new StreamReader(oldFile))
                using (var sw = new StreamWriter(newFile, false))
                {
                    while (!sr.EndOfStream)
                    {
                        var line = sr.ReadLine();
                        string[] spt = Regex.Split(line, @"\t");
                        if (spt.Length < 8)
                            sw.WriteLine(line);
                        else
                        {
                            if (nopInstances.Exists(x => x.Equals(spt[7], StringComparison.OrdinalIgnoreCase)))
                                spt[6] = "nop";
                            string newText = string.Join("\t", spt);
                            sw.WriteLine(newText);
                        }
                    }
                }
            }
        }

        public void AddDebugPrint()
        {
            FlowRows.Add(new FlowRow
            {
                OpCode = "Test",
                Parameter = "Debug_Print_Instrument_Status_Check_End",
                Enable = "B_DebugPrint_Instrument_Status"
            });
        }

        public void AddReturnRow()
        {
            FlowRows.Add(new FlowRow
            {
                OpCode = "Return"
            });
        }

        public void AddHeaderRow(string sheetName, string enable)
        {
            FlowRows.Add(new FlowRow
            {
                OpCode = "Test",
                Parameter = sheetName + "_Header",
                Enable = enable
            });
        }

        public void AddFooterRow(string sheetName, string enable)
        {
            FlowRows.Add(new FlowRow
            {
                OpCode = "Test",
                Parameter = sheetName + "_Footer",
                Enable = enable
            });
        }

        public void AddPrintStartRow(string sheetName)
        {
            FlowRows.Add(new FlowRow
            {
                OpCode = "Print",
                Parameter = "\"" + sheetName + " Start\""
            });
        }

        public void AddPrintEndRow(string sheetName)
        {
            FlowRows.Add(new FlowRow
            {
                OpCode = "Print",
                Parameter = "\"" + sheetName + " End\""
            });
        }

        public void AddStartRows(string enable = "")
        {
            var arr = SheetName.Split('_').ToList();
            arr.RemoveAt(0);
            var sheetName = string.Join("_", arr);

            //Set Error Bin
            if (FlowRows.Count > 0 && FlowRows[0].OpCode.Equals("set-error-bin", StringComparison.CurrentCultureIgnoreCase))
            {
                //pass
            }
            else
                AddSetErrorBinRow();
            //Print
            AddPrintStartRow(sheetName);
            //Header
            AddHeaderRow(sheetName, enable);
        }

        public void AddEndRows(string enable = "", bool isPrintDebug = true)
        {
            var arr = SheetName.Split('_').ToList();
            arr.RemoveAt(0);
            var sheetName = string.Join("_", arr);

            //Footer
            AddFooterRow(sheetName, enable);
            //Print
            AddPrintEndRow(sheetName);
            //Debug print
            if(isPrintDebug)
                AddDebugPrint();
            //Return
            AddReturnRow();
        }

        public SubFlowSheet ReplaceFlowName(Dictionary<string, string> instanceReplaceDictionary, List<string> nopInstances)
        {
            bool replaceFlag = false;
            foreach (var row in FlowRows)
            {
                if (instanceReplaceDictionary.ContainsKey(row.Parameter.ToUpper()))
                {
                    row.Parameter = instanceReplaceDictionary[row.Parameter.ToUpper()];
                    replaceFlag = true;
                }

                if (nopInstances.Exists(x => x.Equals(row.Parameter, StringComparison.OrdinalIgnoreCase)))
                {
                    row.OpCode = "nop";
                    replaceFlag = true;
                }
            }
            if (replaceFlag)
                return this;
            return null;
        }

        public void AddClearFlags()
        {
            var flagClears = FlowRows.Where(x => !string.IsNullOrEmpty(x.FailAction))
                    .Select(x => x.FailAction.Split(',')).SelectMany(x => x).Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            List<FlowRow> flowRows = new List<FlowRow>();
            foreach (var flagClear in flagClears)
            {
                FlowRow flowRow = new FlowRow();
                flowRow.OpCode = "flag-clear";
                flowRow.Parameter = flagClear;
                flowRows.Add(flowRow);
            }
            InsertRow(0, flowRows);
        }

        public void AddFlowRow(string opCode, string parameter, string enableWord = "")
        {
            FlowRow row = new FlowRow();
            row.OpCode = opCode;
            row.Parameter = parameter;
            row.Enable = enableWord;
            FlowRows.Add(row);
        }

        public void AddSetErrorBinRow(string binFail = "999", string sortFail = "999")
        {
            FlowRows.Add(new FlowRow
            {
                OpCode = "set-error-bin",
                BinFail = binFail,
                SortFail = sortFail
            });
        }
        #endregion
    }

}
