using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;


namespace IgxlData.IgxlSheets
{
    [Serializable]
    public class SubFlowSheet : IgxlSheet
    {
        private const string SheetType = "DTFlowtableSheet";

        #region Field
        private List<FlowRow> _flowRows;
        public bool IsCalledFromMainFlow;
        #endregion

        #region Property
        public List<FlowRow> FlowRows
        {
            get { return _flowRows; }
            set { _flowRows = value; }
        }
        public string JobName { get; set; }

        #endregion

        #region Contructer
        public SubFlowSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _flowRows = new List<FlowRow>();
            IgxlSheetName = IgxlSheetNameList.FlowTable;
            IsCalledFromMainFlow = true;
        }

        public SubFlowSheet(string sheetName)
            : base(sheetName)
        {
            _flowRows = new List<FlowRow>();
            IgxlSheetName = IgxlSheetNameList.FlowTable;
            IsCalledFromMainFlow = true;
        }
        #endregion

        #region Member Function
        public List<FlowRow> GetFlowTestGroup(int flowIndex)
        {
            List<FlowRow> flowgroup = new List<FlowRow>();
            bool getFirst = false;
            int groupEnd = 0;
            for (int idx = flowIndex; idx < FlowRows.Count; idx++)
            {
                if (FlowRows[idx].Opcode.Equals("Test", StringComparison.CurrentCultureIgnoreCase))
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
                flowgroup.Add(FlowRows[i]);
            }
            return flowgroup;
        }

        public override void Write(string fileName, string version)
        {
            //if (!Directory.Exists(Path.GetDirectoryName(fileName)))
            //    return;
            //using (var sw = new StreamWriter(fileName, false))
            //{
            //    sw.WriteLine("DTFlowtableSheet,version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tFlow Table");
            //    sw.WriteLine("\t\t\t\t\t\tFlow Domain:");
            //    sw.WriteLine("\t\t\tGate\t\t\tCommand\t\t\t\tLimits\t\tDatalog Display Results\t\t\tBin Number\t\tSort Number\t\t\tAction\t\t\tGroup\t\t\t\tDevice\t\t\tDebug\t\tCT Profile Data");
            //    sw.WriteLine("\tLabel\tEnable\tJob\tPart\tEnv\tOpcode\tParameter\tTName\tTNum\tLoLim\tHiLim\tScale\tUnits\tFormat\tPass\tFail\tPass\tFail\tResult\tPass\tFail\tState\tSpecifier\tSense\tCondition\tName\tSense\tCondition\tName\tAssume\tSites\tElapsed Time (s)\tBackground Type\tSerialize\tResource Lock\tFlow Step Locked\tComment");

            //    string[] arr = new string[38];
            //    foreach (FlowRow fr in FlowRows)
            //    {
            //        arr[0] = fr.ColumnA;
            //        arr[1] = fr.Label;
            //        arr[2] = fr.Enable;
            //        arr[3] = fr.Job;
            //        arr[4] = fr.Part;
            //        arr[5] = fr.Env;
            //        arr[6] = fr.Opcode;
            //        arr[7] = fr.Parameter;
            //        arr[8] = fr.TName;
            //        arr[9] = fr.TNum;
            //        arr[10] = fr.LoLim;
            //        arr[11] = fr.HiLim;
            //        arr[12] = fr.Scale;
            //        arr[13] = fr.Units;
            //        arr[14] = fr.Format;
            //        arr[15] = fr.BinPass;
            //        arr[16] = fr.BinFail;
            //        arr[17] = fr.SortPass;
            //        arr[18] = fr.SortFail;
            //        arr[19] = fr.Result;
            //        arr[20] = fr.PassAction;
            //        arr[21] = fr.FailAction;
            //        arr[22] = fr.State;
            //        arr[23] = fr.GroupSpecifier;
            //        arr[24] = fr.GroupSense;
            //        arr[25] = fr.GroupCondition;
            //        arr[26] = fr.GroupName;
            //        arr[27] = fr.DeviceSense;
            //        arr[28] = fr.DeviceCondition;
            //        arr[29] = fr.DeviceName;
            //        arr[30] = fr.DebugAsume;
            //        arr[31] = fr.DebugSites;
            //        arr[32] = fr.CtProfileDataElapsedTime;
            //        arr[33] = fr.CtProfileDataBackgroundType;
            //        arr[34] = fr.CtProfileDataSerialize;
            //        arr[35] = fr.CtProfileDataResourceLock;
            //        arr[36] = fr.CtProfileDataFlowStepLocked;
            //        arr[37] = fr.Comment;
            //        if (string.IsNullOrEmpty(fr.Comment1))
            //            sw.WriteLine(string.Join("\t", arr));
            //        else
            //            sw.WriteLine(string.Join("\t", arr) + "\t" + fr.Comment1);
            //    }
            //}

            //Support 3.0 & 2.3
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                Directory.CreateDirectory(Path.GetDirectoryName(fileName));
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (dic.ContainsKey(version))
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (dic.ContainsKey("3.0"))
                {
                    var igxlSheetsVersion = dic["3.0"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        //public void WriteOld(string fileName, string version)
        //{
        //    double versionDouble = Double.Parse(version);
        //    Action<string> validate = new Action<string>((a) => { });
        //    GenSubFlowSheet flowSheetGenerator = new GenSubFlowSheet(fileName, validate, true, versionDouble);
        //    foreach (FlowRow fr in FlowRows)
        //    {
        //        FlowStepArgs flowStep = new FlowStepArgs();
        //        flowStep.Label = fr.Label;
        //        flowStep.EnableWords = fr.Enable;
        //        flowStep.GateJob = fr.Job;
        //        flowStep.GateEnv = fr.Env;
        //        flowStep.Opcode = fr.Opcode;
        //        flowStep.Parameter = fr.Parameter;
        //        flowStep.TName = fr.TName;
        //        flowStep.TNum = fr.TNum;
        //        flowStep.LoLim = fr.LoLim;
        //        flowStep.HiLim = fr.HiLim;
        //        flowStep.DatalogScale = fr.Scale;
        //        flowStep.DatalogUnits = fr.Units;
        //        flowStep.DatalogFormat = fr.Format;
        //        flowStep.HardBinPass = fr.BinPass;
        //        flowStep.HardBinFail = fr.BinFail;
        //        flowStep.SoftBinPass = fr.SortPass;
        //        flowStep.SoftBinFail = fr.SortFail;
        //        flowStep.Result = fr.Result;
        //        flowStep.FlagPass = fr.PassAction;
        //        flowStep.FlagFail = fr.FailAction;
        //        flowStep.State = fr.State;
        //        flowStep.GroupSpecifier = fr.GroupSpecifier;
        //        flowStep.GroupSense = fr.GroupSense;
        //        flowStep.GroupCondition = fr.GroupCondition;
        //        flowStep.GroupName = fr.GroupName;
        //        flowStep.DeviceSense = fr.DeviceSense;
        //        flowStep.DeviceCondition = fr.DeviceCondition;
        //        flowStep.DeviceName = fr.DeviceName;
        //        flowStep.DebugAssume = fr.DebugAsume;
        //        flowStep.DebugSites = fr.DebugSites;
        //        flowStep.CtProfileDataElapsedTime = fr.CtProfileDataElapsedTime;
        //        flowStep.CtProfileDataBackgroundType = fr.CtProfileDataBackgroundType;
        //        flowStep.CtProfileDataSerialize = fr.CtProfileDataSerialize;
        //        flowStep.CtProfileDataResourceLock = fr.CtProfileDataResourceLock;
        //        flowStep.Comment = fr.Comment;
        //        flowSheetGenerator.AddFlowStep(flowStep);
        //    }
        //    flowSheetGenerator.WriteSheet();
        //}

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

        public void ReplaceOpcode(string oldFile, string newFile, List<string> nopInstances)
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

        public void AddReturnRow()
        {
            FlowRows.Add(new FlowRow
            {
                Opcode = "Return"
            });
        }

        public void AddHeaderRow(string sheetName, string enbale)
        {
            FlowRows.Add(new FlowRow
            {
                Opcode = "Test",
                Parameter = sheetName + "_Header",
                Enable = enbale
            });
        }

        public void AddFooterRow(string sheetName, string enbale)
        {
            FlowRows.Add(new FlowRow
            {
                Opcode = "Test",
                Parameter = sheetName + "_Footer",
                Enable = enbale
            });
        }

        public void AddPrintStartRow(string sheetName)
        {
            FlowRows.Add(new FlowRow
            {
                Opcode = "Print",
                Parameter = "\"" + sheetName + " Start\""
            });
        }

        public void AddPrintEndRow(string sheetName)
        {
            FlowRows.Add(new FlowRow
            {
                Opcode = "Print",
                Parameter = "\"" + sheetName + " End\""
            });
        }

        public void AddStartRows(string enable = "")
        {
            var arr = Name.Split('_').ToList();
            arr.RemoveAt(0);
            var sheetName = string.Join("_", arr);

            //Print
            AddPrintStartRow(sheetName);
            //Header
            AddHeaderRow(sheetName, enable);
        }

        public void AddEndRows(string enable = "")
        {
            var arr = Name.Split('_').ToList();
            arr.RemoveAt(0);
            var sheetName = string.Join("_", arr);

            //Footer
            AddFooterRow(sheetName, enable);
            //Print
            AddPrintEndRow(sheetName);
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
                    row.Opcode = "nop";
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
                flowRow.Opcode = "flag-clear";
                flowRow.Parameter = flagClear;
                flowRows.Add(flowRow);
            }
            InsertRow(0, flowRows);
        }
        #endregion

        public List<FlowRow> GetLimitRows(int index)
        {
            List<FlowRow> flowRows = new List<FlowRow>();
            for (int i = index + 1; i < FlowRows.Count(); i++)
            {
                var row = FlowRows[i];
                if (!row.Opcode.Equals("Use-Limit", StringComparison.CurrentCultureIgnoreCase) &&
                    !row.Opcode.Equals("characterize", StringComparison.CurrentCultureIgnoreCase))
                    break;
                row.FailAction = "F_BV_CALLINST";
                flowRows.Add(row);
            }
            return flowRows;
        }

        public void RemoveLimitRows(int index)
        {
            for (int i = index + 1; i < FlowRows.Count(); i++)
            {
                var row = FlowRows[i];
                if (row.Opcode.Equals("Use-Limit", StringComparison.CurrentCultureIgnoreCase) ||
                    row.Opcode.Equals("characterize", StringComparison.CurrentCultureIgnoreCase))
                {
                    FlowRows.RemoveAt(i);
                    i--;
                }
                else
                    break;
            }
        }

        public SubFlowSheet DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as SubFlowSheet;
            }
        }

        public bool IsSame(SubFlowSheet flowSheet)
        {
            if (FlowRows.Count != flowSheet.FlowRows.Count)
                return false;

            Type type = typeof(FlowRow);
            const BindingFlags memberFlags = BindingFlags.Public | BindingFlags.Instance ;
            var members = new List<MemberInfo>();
            members.AddRange(type.GetFields(memberFlags).Where(p => p.DeclaringType == typeof(FlowRow)).ToList());

            for (int index = 0; index < FlowRows.Count; index++)
            {
                var sourceRow = FlowRows[index];
                if (sourceRow.Opcode.Equals("print",StringComparison.CurrentCultureIgnoreCase))
                    continue;
                if (sourceRow.Parameter.ToUpper().Contains("_HEADER"))
                    continue;
                if (sourceRow.Parameter.ToUpper().Contains("_FOOTER"))
                    continue;

                var targetRow = flowSheet.FlowRows[index];
                foreach (MemberInfo t in members)
                {
                    if (t is FieldInfo)
                    {
                        FieldInfo fieldInfo = (FieldInfo)t;
                        if (fieldInfo.GetValue(sourceRow).ToString() != fieldInfo.GetValue(targetRow).ToString())
                            return false;
                    }
                }
            }
            return true;
        }

        public void InsertRow(int index, List<FlowRow> flowRows)
        {
            _flowRows.InsertRange(index, flowRows);
        }

        // For Relay
        public void InsertRow(int index, FlowRow flowRow)
        {
            _flowRows.Insert(index, flowRow);
        }

        public void AddRow(FlowRow igxlItem)
        {
            _flowRows.Add(igxlItem);
        }

        public void AddRows(List<FlowRow> igxlItemList)
        {
            _flowRows.AddRange(igxlItemList);
        }

        protected void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_flowRows.Count == 0) return;

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
                for (var index = 0; index < _flowRows.Count; index++)
                {
                    var row = _flowRows[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    arr[0] = row.ColumnA;
                    arr[labelIndex] = row.Label;
                    arr[enableIndex] = row.Enable;
                    arr[jobIndex] = row.Job;
                    arr[partIndex] = row.Part;
                    arr[envIndex] = row.Env;
                    arr[opcodeIndex] = row.Opcode;
                    arr[parameterIndex] = row.Parameter;
                    arr[tNameIndex] = row.TName;
                    arr[tNumIndex] = row.TNum;
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
    }

}