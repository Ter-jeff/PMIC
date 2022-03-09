using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlSheets
{
    public class CharSheet : IgxlSheet
    {
        private const string SheetType = "DTCharacterizationSheet";

        private List<CharSetup> _charSetups;

        #region Properity
        public List<CharSetup> CharSetups
        {
            get { return _charSetups; }
            set { _charSetups = value; }
        }
        #endregion

        #region Constructor
        public CharSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            _charSetups = new List<CharSetup>();
            IgxlSheetName = IgxlSheetNameList.Characterization;
        }

        public CharSheet(string sheetName)
            : base(sheetName)
        {
            _charSetups = new List<CharSetup>();
            IgxlSheetName = IgxlSheetNameList.Characterization;
        }

        #endregion

        public void AddRow(CharSetup setup)
        {
            _charSetups.Add(setup);
        }

        public override void Write(string fileName, string version = "2.5")
        {
            //double versionDouble = Double.Parse(version);
            //Action<string> validate = new Action<string>((a) => { });
            //GenCZSheet characterizationSheetGenerator = new GenCZSheet(fileName, validate, "", true, versionDouble);
            //foreach (CharSetup setup in CharSetups)
            //{
            //    foreach (CharStep charStep in setup.CharSteps)
            //    {
            //        characterizationSheetGenerator.AddRow(setup.SetupName, setup.TestMethod, charStep.StepName, charStep.Mode, charStep.ParameterType, charStep.ParameterName,
            //            charStep.RangeCalcField, charStep.RangeFrom, charStep.RangeTo, charStep.RangeSteps, charStep.RangeStepSize, charStep.PerformTest, charStep.TestLimitLow,
            //            charStep.TestLimitHigh, charStep.AlgorithmName, charStep.AlgorithmArguments, charStep.AlgorithmResultsCheck, charStep.AlgorithmTransition,
            //            charStep.ApplyToPins, charStep.ApplyToPinExecMode, charStep.ApplyToTimeSets, charStep.DeviceMarginContexts, charStep.DeviceMarginPatterns, charStep.DeviceMarginInstances,
            //            charStep.AdjustBackoff, charStep.AdjustSpecName, charStep.AdjustFromSetup, charStep.Function, charStep.Arguments, charStep.PreSetup,
            //            charStep.PreSetupArguments, charStep.PreStep, charStep.PreStepArguments, charStep.PrePoint, charStep.PrePointArguments, charStep.PostPoint, charStep.PostPointArguments,
            //            charStep.PostStep, charStep.PostStepArguments, charStep.PostSetup, charStep.PostSetupArguments, charStep.OutputFormat, charStep.OutputTextFile, charStep.OutputSheet,
            //            charStep.OutputDestinationsTextFile, charStep.OutputDestinationsSheet, charStep.OutputDestinationsDatalog, charStep.OutputDestinationsImmediateWin, charStep.OutputDestinationsOutputWin, charStep.Comment, charStep.AxisExecutionOrder, charStep.OutputSuspendDatalog);
            //    }
            //}
            //characterizationSheetGenerator.WriteSheet();

            //Support 2.5 & 2.6
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
                else if (dic.ContainsKey("2.6"))
                {
                    var igxlSheetsVersion = dic["2.6"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (_charSetups.Count == 0) return;

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
                var interposeFunctionsIndex = GetIndexFrom(igxlSheetsVersion, "Interpose Functions");
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
                var outputDestinationsTextFileIndex = GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Text File");
                var outputDestinationsSheetIndex = GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Sheet");
                var outputDestinationsDatalogIndex = GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Datalog");
                var outputDestinationsImmediateWinIndex = GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Immediate Win");
                var outputDestinationsOutputWinIndex = GetIndexFrom(igxlSheetsVersion, "Output Destinations", "Output Win");
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
                for (var index = 0; index < _charSetups.Count; index++)
                {
                    var row = _charSetups[index];
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
                            arr[algorithmArgumentsIndex] = charStep.AlgorithmArguments;
                            arr[algorithmResultsCheckIndex] = charStep.AlgorithmResultsCheck;
                            arr[algorithmTransitionIndex] = charStep.AlgorithmTransition;
                            arr[applyToPinsIndex] = charStep.ApplyToPins;
                            arr[applyToPinExecModeIndex] = charStep.ApplyToPinExecMode;
                            arr[applyToTimeSetsIndex] = charStep.ApplyToTimeSets;
                            arr[deviceMarginContextsIndex] = charStep.DeviceMarginContexts;
                            arr[deviceMarginPatternsIndex] = charStep.DeviceMarginPatterns;
                            arr[deviceMarginInstancesIndex] = charStep.DeviceMarginInstances;
                            arr[adjustBackoffIndex] = charStep.AdjustBackoff;
                            arr[adjustSpecNameIndex] = charStep.AdjustSpecName;
                            arr[adjustFromSetupIndex] = charStep.AdjustFromSetup;
                            if (axisExecutionOrderIndex != -1)
                                arr[axisExecutionOrderIndex] = charStep.AxisExecutionOrder;
                            arr[functionIndex] = charStep.Function;
                            arr[argumentsIndex] = charStep.Arguments;
                            arr[interposeFunctionsIndex] = charStep.InterposeFunctions;
                            arr[preSetupIndex] = charStep.PreSetup;
                            arr[preSetupIndex + 1] = charStep.PreSetupArguments;
                            arr[preStepIndex] = charStep.PreStep;
                            arr[preStepIndex + 1] = charStep.PreStepArguments;
                            arr[prePointIndex] = charStep.PrePoint;
                            arr[prePointIndex + 1] = charStep.PrePointArguments;
                            arr[postPointIndex] = charStep.PostPoint;
                            arr[postPointIndex + 1] = charStep.PostPointArguments;
                            arr[postStepIndex] = charStep.PostStep;
                            arr[postStepIndex + 1] = charStep.PostStepArguments;
                            arr[postSetupIndex] = charStep.PostSetup;
                            arr[postSetupIndex + 1] = charStep.PostSetupArguments;
                            arr[outputFormatIndex] = charStep.OutputFormat;
                            arr[outputTextFileIndex] = charStep.OutputTextFile;
                            arr[outputSheetIndex] = charStep.OutputSheet;
                            if (outputSuspendDatalogIndex != -1)
                                arr[outputSuspendDatalogIndex] = charStep.OutputSuspendDatalog;
                            arr[outputDestinationsTextFileIndex] = charStep.OutputDestinationsTextFile;
                            arr[outputDestinationsSheetIndex] = charStep.OutputDestinationsSheet;
                            arr[outputDestinationsDatalogIndex] = charStep.OutputDestinationsDatalog;
                            arr[outputDestinationsImmediateWinIndex] = charStep.OutputDestinationsImmediateWin;
                            arr[outputDestinationsOutputWinIndex] = charStep.OutputDestinationsOutputWin;
                            arr[commentIndex] = charStep.Comment;
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
    }
}