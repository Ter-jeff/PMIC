using System.Collections.Generic;
using System.IO;
using IgxlData.IgxlBase;
using OfficeOpenXml;
using Teradyne.Oasis.IGLinkBase.ProgramGeneration;
using System;
using Teradyne.Oasis.IGData;
using Teradyne.Oasis.IGData.Utilities;
using System.Linq;

namespace IgxlData.IgxlSheets
{
    public class FlowSheet : IgxlSheet
    {
        #region Field
        private const string SheetType = "DTFlowtableSheet";
        #endregion

        #region Property

        public List<FlowRow> FlowRows { get; set; }

        #endregion

        #region Contructer

        public FlowSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            FlowRows = new List<FlowRow>();
            IgxlSheetName = IgxlSheetNameList.FlowTable;
        }

        public FlowSheet(string sheetName)
            : base(sheetName)
        {
            FlowRows = new List<FlowRow>();
            IgxlSheetName = IgxlSheetNameList.FlowTable;
        }
        #endregion

        #region Member function
        public void InsertRow(int index, List<FlowRow> flowRows)
        {
            FlowRows.InsertRange(index, flowRows);
        }

        // For Relay
        public void InsertRow(int index, FlowRow flowRow)
        {
            FlowRows.Insert(index, flowRow);
        }

        public void AddRow(FlowRow igxlItem)
        {
            //throw new System.NotImplementedException();
            FlowRows.Add(igxlItem);
        }

        public void AddRows(List<FlowRow> igxlItemList)
        {
            FlowRows.AddRange(igxlItemList);
        }

        protected override void WriteHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteColumnsHeader()
        {
            throw new NotImplementedException();
        }

        protected override void WriteRows()
        {
            throw new NotImplementedException();
        }

        public override void Write(string fileName, string version)
        {
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version=="2.3")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (version=="3.0")
                {
                    var igxlSheetsVersion = dic["3.0"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        protected virtual void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
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

        #endregion
    }
}