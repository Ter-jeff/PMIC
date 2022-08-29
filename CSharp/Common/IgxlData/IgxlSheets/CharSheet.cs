using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using IgxlData.IgxlBase;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class CharSheet : IgxlSheet
    {
        private const string SheetType = "DTCharacterizationSheet";

        #region Constructor

        public CharSheet(string sheetName)
            : base(sheetName)
        {
            CharSetups = new List<CharSetup>();
            IgxlSheetName = IgxlSheetNameList.Characterization;
        }

        #endregion

        public List<CharSetup> CharSetups { get; set; }

        public void AddRow(CharSetup charSetup)
        {
            CharSetups.Add(charSetup);
        }

        public void AddRows(List<CharSetup> charSetups)
        {
            CharSetups.AddRange(charSetups);
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

        public override void Write(string fileName, string version = "")
        {
            if (string.IsNullOrEmpty(version))
                version = "2.6";
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version == "2.5")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet2_5(fileName, version, igxlSheetsVersion);
                }
                else if (version == "2.6")
                {
                    var igxlSheetsVersion = dic["2.6"];
                    WriteSheet2_6(fileName, version, igxlSheetsVersion);
                }
                else
                {
                    throw new Exception(string.Format("The Characterization sheet version:{0} is not supported!",
                        version));
                }
            }
        }

        private void WriteSheet2_5(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (CharSetups.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var setupNameIndex = GetIndexFrom(igxlSheetsVersion, "Setup Name");
                var testMethodIndex = GetIndexFrom(igxlSheetsVersion, "Test Method");
                var stepNameIndex = GetIndexFrom(igxlSheetsVersion, "Step Name");
                var modeIndex = GetIndexFrom(igxlSheetsVersion, "Mode");
                var parameterTypeIndex = GetIndexFrom(igxlSheetsVersion, "Parameter", "Type");
                var parameterNameIndex = GetIndexFrom(igxlSheetsVersion, "Parameter", "Name");
                var rangeCalcFieldIndex = GetIndexFrom(igxlSheetsVersion, "Range", "Calc Field");
                var rangeFromIndex = GetIndexFrom(igxlSheetsVersion, "Range", "From");
                var rangeToIndex = GetIndexFrom(igxlSheetsVersion, "Range", "To");
                var rangeStepsIndex = GetIndexFrom(igxlSheetsVersion, "Range", "Steps");
                var rangeStepSizeIndex = GetIndexFrom(igxlSheetsVersion, "Range", "Step Size");
                var performTestIndex = GetIndexFrom(igxlSheetsVersion, "Perform Test");
                var testLimitsLowIndex = GetIndexFrom(igxlSheetsVersion, "Test Limits", "Low");
                var testLimitsHighIndex = GetIndexFrom(igxlSheetsVersion, "Test Limits", "High");
                var algorithmNameIndex = GetIndexFrom(igxlSheetsVersion, "Algorithm", "Name");
                var algorithmArgumentsIndex = GetIndexFrom(igxlSheetsVersion, "Algorithm", "Arguments");
                var algorithmResultsCheckIndex = GetIndexFrom(igxlSheetsVersion, "Algorithm", "Results Check");
                var algorithmTransitionIndex = GetIndexFrom(igxlSheetsVersion, "Algorithm", "Transition");
                var applyToPinsIndex = GetIndexFrom(igxlSheetsVersion, "Apply To", "Pins");
                var applyToPinExecModeIndex = GetIndexFrom(igxlSheetsVersion, "Apply To", "Pin Exec Mode");
                var applyToTimeSetsIndex = GetIndexFrom(igxlSheetsVersion, "Apply To", "Time Sets");
                var deviceMarginContextsIndex = GetIndexFrom(igxlSheetsVersion, "Device Margin", "Contexts");
                var deviceMarginPatternsIndex = GetIndexFrom(igxlSheetsVersion, "Device Margin", "Patterns");
                var deviceMarginInstancesIndex = GetIndexFrom(igxlSheetsVersion, "Device Margin", "Instances");
                var adjustBackoffIndex = GetIndexFrom(igxlSheetsVersion, "Adjust", "Backoff");
                var adjustSpecNameIndex = GetIndexFrom(igxlSheetsVersion, "Adjust", "Spec Name");
                var adjustFromSetupIndex = GetIndexFrom(igxlSheetsVersion, "Adjust", "From Setup");
                var axisExecutionOrderIndex = GetIndexFrom(igxlSheetsVersion, "Shmoo", "Axis Execution Order");
                var functionIndex = GetIndexFrom(igxlSheetsVersion, "Function");
                var argumentsIndex = GetIndexFrom(igxlSheetsVersion, "Arguments");
                //var interposeFunctionsIndex = GetIndexFrom(igxlSheetsVersion, "Interpose Functions");
                var preSetupIndex = GetIndexFrom(igxlSheetsVersion, "Interpose Functions", "Pre Setup");
                var preStepIndex = GetIndexFrom(igxlSheetsVersion, "Interpose Functions", "Pre Step");
                var prePointIndex = GetIndexFrom(igxlSheetsVersion, "Interpose Functions", "Pre Point");
                var postPointIndex = GetIndexFrom(igxlSheetsVersion, "Interpose Functions", "Post Point");
                var postStepIndex = GetIndexFrom(igxlSheetsVersion, "Interpose Functions", "Post Step");
                var postSetupIndex = GetIndexFrom(igxlSheetsVersion, "Interpose Functions", "Post Setup");
                var outputFormatIndex = GetIndexFrom(igxlSheetsVersion, "Output", "Format");
                var outputTextFileIndex = GetIndexFrom(igxlSheetsVersion, "Output", "Text File");
                var outputSheetIndex = GetIndexFrom(igxlSheetsVersion, "Output", "Sheet");
                var outputSuspendDatalogIndex = GetIndexFrom(igxlSheetsVersion, "Output", "Suspend Datalog");
                var outputDestinationsTextFileIndex =
                    GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Text File");
                var outputDestinationsSheetIndex = GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Sheet");
                var outputDestinationsDatalogIndex = GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Datalog");
                var outputDestinationsImmediateWinIndex =
                    GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Immediate Win");
                var outputDestinationsOutputWinIndex =
                    GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Output Win");
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

                for (var index = 0; index < CharSetups.Count; index++)
                {
                    var row = CharSetups[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    foreach (var charStep in row.CharSteps)
                    {
                        if (!string.IsNullOrEmpty(row.SetupName))
                        {
                            arr[0] = row.ColumnA;
                            arr[setupNameIndex] = row.SetupName;
                            arr[testMethodIndex] = row.TestMethod;
                            arr[stepNameIndex] = charStep.StepName;
                            arr[modeIndex] = charStep.Mode;
                            arr[parameterTypeIndex] = charStep.ParameterType;
                            arr[parameterNameIndex] = charStep.ParameterName;
                            arr[rangeCalcFieldIndex] = charStep.RangeCalcField;
                            arr[rangeFromIndex] = charStep.RangeFrom;
                            arr[rangeToIndex] = charStep.RangeTo;
                            arr[rangeStepsIndex] = charStep.RangeSteps;
                            arr[rangeStepSizeIndex] = charStep.RangeStepSize;
                            arr[performTestIndex] = charStep.PerformTest;
                            arr[testLimitsLowIndex] = charStep.TestLimitLow;
                            arr[testLimitsHighIndex] = charStep.TestLimitHigh;
                            arr[algorithmNameIndex] = charStep.AlgorithmName;
                            arr[algorithmArgumentsIndex] = charStep.AlgorithmArgs;
                            arr[algorithmResultsCheckIndex] = charStep.AlgorithmResultsCheck;
                            arr[algorithmTransitionIndex] = charStep.AlgorithmTransition;
                            arr[applyToPinsIndex] = charStep.ApplyToPins;
                            arr[applyToPinExecModeIndex] = charStep.ApplyToPinExecMode;
                            arr[applyToTimeSetsIndex] = charStep.ApplyToTimeSets;
                            arr[deviceMarginContextsIndex] = charStep.DeviceMarginContexts;
                            arr[deviceMarginPatternsIndex] = charStep.DeviceMarginPatterns;
                            arr[deviceMarginInstancesIndex] = charStep.DeviceMarginInstances;
                            arr[adjustBackoffIndex] = charStep.AdjustBackOff;
                            arr[adjustSpecNameIndex] = charStep.AdjustSpecName;
                            arr[adjustFromSetupIndex] = charStep.AdjustFromSetup;
                            if (axisExecutionOrderIndex != -1)
                                arr[axisExecutionOrderIndex] = charStep.AxisExecutionOrder;
                            arr[functionIndex] = charStep.Function;
                            arr[argumentsIndex] = charStep.Arguments;
                            //arr[interposeFunctionsIndex] = charStep.InterposeFunctions;
                            arr[preSetupIndex] = charStep.PreSetupFunction;
                            arr[preSetupIndex + 1] = charStep.PreSetupArgs;
                            arr[preStepIndex] = charStep.PreStepFunction;
                            arr[preStepIndex + 1] = charStep.PreStepArgs;
                            arr[prePointIndex] = charStep.PrePointFunction;
                            arr[prePointIndex + 1] = charStep.PrePointArgs;
                            arr[postPointIndex] = charStep.PostPointFunction;
                            arr[postPointIndex + 1] = charStep.PostPointArgs;
                            arr[postStepIndex] = charStep.PostStepFunction;
                            arr[postStepIndex + 1] = charStep.PostStepArgs;
                            arr[postSetupIndex] = charStep.PostSetupFunction;
                            arr[postSetupIndex + 1] = charStep.PostSetupArgs;
                            arr[outputFormatIndex] = charStep.OutputFormat;
                            arr[outputTextFileIndex] = charStep.OutputTextFile;
                            arr[outputSheetIndex] = charStep.OutputSheet;
                            if (outputSuspendDatalogIndex != -1)
                                arr[outputSuspendDatalogIndex] = charStep.SuspendDataLog;
                            arr[outputDestinationsTextFileIndex] = charStep.OutputToTextFile;
                            arr[outputDestinationsSheetIndex] = charStep.OutputToSheet;
                            arr[outputDestinationsDatalogIndex] = charStep.OutputToDataLog;
                            arr[outputDestinationsImmediateWinIndex] = charStep.OutputToImmediateWin;
                            arr[outputDestinationsOutputWinIndex] = charStep.OutputToOutputWin;
                            arr[commentIndex] = charStep.Comment;
                        }
                        else
                        {
                            arr = new[] {"\t"};
                        }

                        sw.WriteLine(string.Join("\t", arr));
                    }
                }

                #endregion
            }
        }

        private void WriteSheet2_6(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            WriteSheet2_5(fileName, version, igxlSheetsVersion);
        }
    }
}