using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Extension;
using CommonReaderLib.Input;
using CommonReaderLib.PatternListCsv;
using IgxlData.IgxlBase;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using IgxlData.NonIgxlSheets;
using IgxlData.VBT;
using OfficeOpenXml;
using PatInfoCmdLib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace CommonReaderLib.DebugPlan
{
    public class DebugPlanMain
    {
        private const string Cz = "CZ";

        private const string ConPatternDashboard = "pattern_dashboard";
        private const string ConProcessCondition = "Process_Condition";

        private const string Debug_LVCC_VminBoundary = "Debug_LVCC_VminBoundary";
        private const string Enable_DFTLHFC_Debug = "Enable_DFTLHFC_Debug";
        private const string Enable_Faillog_Debug = "Enable_Faillog_Debug";

        public List<AiTestPlanSheet> AiTestPlanSheets = new List<AiTestPlanSheet>();

        public List<Error> Errors = new List<Error>();
        public string InputFile;

        public PatternListSheet PatternListSheet;
        public ProcessConditionSheet ProcessCondition;

        public DebugPlanMain(string inputFile)
        {
            InputFile = inputFile;
        }

        public void Read()
        {
            using (var package = new ExcelPackage(new FileInfo(InputFile)))
            {
                var dashboard = package.Workbook.Worksheets[ConPatternDashboard];
                if (dashboard == null)
                {
                    var error = new Error
                    {
                        EnumErrorType = EnumErrorType.MissingSheet,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = ConPatternDashboard,
                        Message = string.Format("Can not find {0} sheet in input file !!!", ConPatternDashboard)
                    };
                }
                else
                {
                    var patternListReader = new PatternListReader();
                    PatternListSheet = patternListReader.ReadSheet(dashboard);
                }

                var processCondition = package.Workbook.Worksheets[ConProcessCondition];
                if (processCondition == null)
                {
                    var error = new Error
                    {
                        EnumErrorType = EnumErrorType.MissingSheet,
                        ErrorLevel = EnumErrorLevel.Error,
                        SheetName = ConProcessCondition,
                        Message = string.Format("Can not find {0} sheet in input file !!!", ConProcessCondition)
                    };
                }
                else
                {
                    var processConditionReader = new ProcessConditionReader();
                    ProcessCondition = processConditionReader.ReadSheet(processCondition);
                }

                foreach (var sheet in package.Workbook.Worksheets)
                {
                    var sheetType = sheet.Cells[1, 1].Value.ToString().Split(';').First();
                    if (sheetType.Equals("AITestPlanSheet", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var aITestPlanReader = new AiTestPlanReader();
                        AiTestPlanSheets.Add(aITestPlanReader.ReadSheet(sheet));
                    }
                }

                SetUniquePatSetName(AiTestPlanSheets);
            }
        }

        public OtherSheet GenDfcList()
        {
            var OtherSheet = new OtherSheet("DFC_List");
            OtherSheet.Lines.Add("Test Instance");
            var lines = AiTestPlanSheets.SelectMany(x => x.Rows)
                .Where(x => x.EnumDataLoggingSettingType == EnumDataLoggingSettingType.DFC).Select(x => x.TestName);
            OtherSheet.Lines.AddRange(lines);
            return OtherSheet;
        }

        public bool CheckAll(string patternFolder, string timeFolder, string testProgram)
        {
            Errors.Clear();

            #region pre action
            var timeSets = new List<string>();
            var Dcspecs = new List<string>();
            var patterns = new List<string>();
            var pins = new List<string>();
            var acSymbols = new List<string>();
            var checkTimeSet = false;
            var checkDcSpec = false;
            var checkPattern = false;
            var checkPin = false;
            if (Directory.Exists(patternFolder))
            {
                var gzs = new HashSet<string>(Directory.GetFiles(patternFolder, "*.gz", SearchOption.AllDirectories));
                foreach (var gz in gzs)
                {
                    var name = Regex.Replace(gz, ".gz$", "", RegexOptions.IgnoreCase);
                    name = Regex.Replace(name, ".atp$", "", RegexOptions.IgnoreCase);
                    name = Regex.Replace(name, ".pat$", "", RegexOptions.IgnoreCase);
                    name = Regex.Replace(name, ".patx$", "", RegexOptions.IgnoreCase);
                    patterns.Add(Path.GetFileName(name));
                }
                checkPattern = true;
            }
            if (Directory.Exists(timeFolder))
            {
                var timeSetFiles = new List<string>(Directory.GetFiles(timeFolder, "*.txt", SearchOption.AllDirectories));
                timeSets.AddRange(timeSetFiles.Select(x => Path.GetFileNameWithoutExtension(x)).ToList());
                checkTimeSet = true;
            }
            if (File.Exists(testProgram))
            {
                var igxlData = new IgxlDataReader(testProgram);
                timeSets.AddRange(igxlData.TimeSetBasicSheets);
                Dcspecs.AddRange(igxlData.DcSpecSheets.SelectMany(x => x.CategoryList).Distinct().ToList());
                pins = igxlData.PinMapSheets.SelectMany(x => x.GetAllPins()).Distinct().ToList();
                acSymbols = igxlData.AcSpecSheets.SelectMany(x => x.AcSpecs).Where(x => !x.IsBackup).Select(x => x.Symbol).Distinct().ToList();
                checkTimeSet = true;
                checkDcSpec = true;
                checkPin = true;
            }
            #endregion

            #region check AiTestPlanSheets
            foreach (var aiTestPlanSheet in AiTestPlanSheets)
            {
                aiTestPlanSheet.Chcek();
                if (checkTimeSet)
                    aiTestPlanSheet.ChcekTimeSet(timeSets.Distinct().ToList());
                if (checkDcSpec)
                    aiTestPlanSheet.ChcekDcSpec(Dcspecs);
                if (checkPattern && PatternListSheet != null)
                    aiTestPlanSheet.ChcekPattern(patterns, PatternListSheet);
                if (checkPin)
                    aiTestPlanSheet.ChcekPins(pins, acSymbols);
                Errors.AddRange(aiTestPlanSheet.Errors);
            }
            #endregion

            #region check pattern_dashboard
            if (PatternListSheet != null && !string.IsNullOrEmpty(timeFolder))
            {
                if (checkTimeSet & checkPattern)
                    PatternListSheet.CheckPatternTimeSet(timeFolder, patterns);
                Errors.AddRange(PatternListSheet.Errors);

            }
            #endregion

            return Errors.Where(x => x.ErrorLevel == EnumErrorLevel.Error).Any();
        }

        public List<PatSetSubRow> GenPatSetSubRows(string patternFolder)
        {
            var patSetSubRows = new List<PatSetSubRow>();
            var usedPatterns = AiTestPlanSheets.SelectMany(x => x.Rows).SelectMany(x => x.Patterns)
                .Select(x => x.OriName.ToUpper()).Distinct().ToList();
            var patterns = new List<string>(Directory.GetFiles(patternFolder, "*.gz", SearchOption.AllDirectories)).ToList();
            foreach (var usedPattern in usedPatterns)
            {
                var patName = usedPattern + ".PAT.GZ";
                if (PatternListSheet.Rows.Exists(x => x.Pattern.Equals(usedPattern, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var name = PatternListSheet.Rows.Find(x =>
                        x.Pattern.Equals(usedPattern, StringComparison.CurrentCultureIgnoreCase)).PatternDate;
                    patName = name + ".PAT.GZ";
                }

                if (patterns.Exists(x => Path.GetFileName(x).Equals(patName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var find = patterns.Find(x => Path.GetFileName(x).Equals(patName, StringComparison.CurrentCultureIgnoreCase));
                    var patternFolderName = Path.GetFileName(patternFolder);
                    var fileValue = find.Replace(patternFolder, @".\" + patternFolderName);
                    fileValue = fileValue.ToUpper().Replace(@".ATP.GZ", "").Replace(".PAT.GZ", "");

                    var moduleName = "";
                    var atpContent = "";
                    var patInfoReader = new PatPatInfoReader();
                    if (new PatInfoCmd().ConvertByArgs(find, ref atpContent, "-hdr -switches"))
                    {
                        var vmVectorName = patInfoReader.GetModuleNames(atpContent.Split('\n').ToList());
                        if (vmVectorName.Split(',').Count() == 2)
                            moduleName = vmVectorName.Split(',').Last();
                    }
                    if (string.IsNullOrEmpty(moduleName))
                        continue;

                    var patSetSubRow = new PatSetSubRow
                    {
                        Comment = "New for CZ",
                        PatternFileName = fileValue + ".PAT:" + moduleName.ToUpper()
                    };
                    patSetSubRows.Add(patSetSubRow);
                }
            }

            return patSetSubRows;
        }

        private void SetUniquePatSetName(List<AiTestPlanSheet> aiTestPlanSheets)
        {
            var patSetNames = new List<string>();
            foreach (var aiTestPlanSheet in aiTestPlanSheets)
                foreach (var row in aiTestPlanSheet.Rows)
                    if (patSetNames.Contains(row.TestInstanceName, StringComparer.CurrentCultureIgnoreCase))
                    {
                        row.PatSetName = row.SheetName + "_Row" + row.RowNum + "_" + row.TestInstanceName;
                    }
                    else
                    {
                        row.PatSetName = row.TestInstanceName;
                        patSetNames.Add(row.TestInstanceName);
                    }
        }

        public PatSetSheet GenPatSetAllSheet(string patternFolder)
        {
            var patterns = new List<string>(Directory.GetFiles(patternFolder, "*.gz", SearchOption.AllDirectories))
                .ToList();
            const string patSetsAllCz = "PatSets_All_" + Cz;
            var patSetSheet = new PatSetSheet(patSetsAllCz);
            var rows = AiTestPlanSheets.SelectMany(x => x.Rows).Where(x => Regex.IsMatch(x.UseNotUse, "^use", RegexOptions.IgnoreCase)).ToList();
            var usedPatterns = rows.SelectMany(x => x.Patterns).Select(x => x.OriName).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
            foreach (var usedPattern in usedPatterns)
            {
                var patternDate = new PatternDate(usedPattern);
                var patGz = patternDate.OriName + ".PAT.GZ";
                if (PatternListSheet.Rows.Exists(x => x.Pattern.Equals(patternDate.OriName, StringComparison.CurrentCultureIgnoreCase)))
                {
                    patGz = PatternListSheet.Rows
                       .Find(x => x.Pattern.Equals(patternDate.OriName, StringComparison.CurrentCultureIgnoreCase))
                       .PatternDate + ".PAT.GZ";
                }

                var fileValue = "";
                if (patterns.Exists(x => Path.GetFileName(x).Equals(patGz, StringComparison.CurrentCultureIgnoreCase)))
                {
                    var find = patterns.Find(x =>
                        Path.GetFileName(x).Equals(patGz, StringComparison.CurrentCultureIgnoreCase));
                    var patternFolderName = Path.GetFileName(patternFolder);
                    fileValue = find.Replace(patternFolder, @".\" + patternFolderName);
                }

                var patSet = new PatSet();
                patSet.PatSetName = patternDate.Name;
                var patSetRow = new PatSetRow();
                patSetRow.PatternSet = patternDate.Name;
                patSetRow.File = fileValue;
                patSetRow.Burst = "No";
                patSet.AddRow(patSetRow);
                patSetSheet.AddPatSet(patSet);
            }
            return patSetSheet;
        }

        public PatSetSheet GenPatSetSheet(string patternFolder)
        {
            const string patSetsAllCz = "PatSets_" + Cz;
            var patSetSheet = new PatSetSheet(patSetsAllCz);
            foreach (var aiTestPlanSheet in AiTestPlanSheets)
                foreach (var row in aiTestPlanSheet.Rows)
                {
                    if (!Regex.IsMatch(row.UseNotUse, "^use", RegexOptions.IgnoreCase))
                        continue;
                    var patSet = new PatSet();
                    patSet.PatSetName = row.PatSetName;
                    foreach (var pattern in row.Patterns)
                    {
                        var patSetRow = new PatSetRow();
                        patSetRow.PatternSet = row.PatSetName;
                        patSetRow.File = pattern.Name;
                        patSetRow.Burst = "No";
                        patSet.AddRow(patSetRow);
                    }

                    patSetSheet.AddPatSet(patSet);
                }

            return patSetSheet;
        }

        public List<BasFile> GenBas(string execEnableWord, string testProgram, string tempFolder)
        {
            var basFiles = new List<BasFile>();
            var sheetName = "VBT_LIB_PV.bas";
            var basFile = new BasFile(sheetName);
            var execEnableWords = execEnableWord.Split(',').ToList();
            var igxlSheetReader = new IgxlSheetReader();
            var totalEnableWords = igxlSheetReader.GetEnables(testProgram);
            basFile.Lines.Add("Attribute VB_Name = \"" + Path.GetFileNameWithoutExtension(sheetName) + "\"");
            basFile.Lines.Add(SetEnableWords(execEnableWords, totalEnableWords));
            basFile.Lines.Add(PrintEnableWords(totalEnableWords));
            basFiles.Add(basFile);
            return basFiles;
        }

        private string SetEnableWords(List<string> execEnableWords, List<string> totalEnableWords)
        {
            var codeText = "Public Sub SetEnableWords()" + "\r\n";

            foreach (var enableWord in totalEnableWords)
            {
                var flag = execEnableWords.Exists(
                    x => x.Equals(enableWord, StringComparison.CurrentCultureIgnoreCase));
                codeText += string.Format("  tl_ExecSetEnableWord \"{0}\", {1}", enableWord, flag) + "\r\n";
            }

            codeText += "End Sub\r\n";
            return codeText;
        }

        private string PrintEnableWords(List<string> totalEnableWords)
        {
            var codeText = "Public Sub PrintEnableWords()" + "\r\n";
            foreach (var enableWord in totalEnableWords)
                codeText += string.Format(
                    "  If (tl_ExecGetEnableWord(\"{0}\")) Then TheExec.Datalog.WriteComment \"{0}:\" + CStr(tl_ExecGetEnableWord(\"{0}\"))",
                    enableWord) + "\r\n";
            codeText += "End Sub\r\n";
            return codeText;
        }

        public InstanceSheet GenInstSheet(VbtFunctionLib vbtFunctionLib, AcSpecSheet acSpecSheet)
        {
            const string instCz = "Inst_" + Cz;
            const string vbtName = VbtFunctionLib.FunctionalCharName;
            var instanceSheet = new InstanceSheet(instCz);
            foreach (var aiTestPlanSheet in AiTestPlanSheets)
            {
                instanceSheet.AddHeaderFooter(Cz);
                instanceSheet.AddHeaderFooter("PrintEnableWords");
                instanceSheet.AddRow(new InstanceRow
                { TestName = "PrintEnableWords", Name = "PrintEnableWords", Type = "VBT" });
                foreach (var row in aiTestPlanSheet.Rows)
                {
                    if (!Regex.IsMatch(row.UseNotUse, "^use", RegexOptions.IgnoreCase))
                        continue;
                    var instanceRow = new InstanceRow();
                    var vbtFunction = vbtFunctionLib.GetFunctionByName(vbtName);
                    instanceRow.ColumnA = row.SetColumnA();
                    instanceRow.TestName = row.TestName;
                    instanceRow.Name = vbtFunction.FunctionName;
                    instanceRow.Type = "VBT";
                    instanceRow.ArgList = vbtFunction.Parameters;
                    instanceRow.DcCategory = row.DcCategory;
                    instanceRow.DcSelector = row.DcSelector;
                    instanceRow.TimeSets = GetTimeSet(row);
                    instanceRow.AcCategory = acSpecSheet.GetAcByTimeSet(instanceRow.TimeSets);
                    instanceRow.AcSelector = "Typ";
                    var block = row.GetBlock();
                    instanceRow.PinLevels = "Levels_" + block;
                    instanceRow.Args = vbtFunction.Args;
                    instanceSheet.AddRow(instanceRow);
                    vbtFunction.SetParamValue("PMode", instanceRow.DcCategory + ":" + instanceRow.DcSelector);
                    SetVbtParameters(vbtFunction, row);
                }
            }

            return instanceSheet;
        }

        private string GetTimeSet(AiTestPlanRow row)
        {
            if (!string.IsNullOrEmpty(row.Timeset))
                return row.Timeset;

            var timeSets = row.GetTimeSetsByPayloads(PatternListSheet);
            if (timeSets.Any())
                return string.Join(",", timeSets.Distinct());

            return "";
        }

        private void SetVbtParameters(VbtFunctionBase vbtFunction, AiTestPlanRow aiTestPlanRow)
        {
            //vbtFunction.SetParamValue("Interpose_Prepat", ConvertCharCondition(""));
            for (var index = 0; index < aiTestPlanRow.Inits.Count; index++)
            {
                var init = aiTestPlanRow.Inits[index];
                vbtFunction.SetParamValue("Init_Patt" + (index + 1), init.Name);
            }

            for (var index = 0; index < aiTestPlanRow.Payloads.Count; index++)
            {
                var payload = aiTestPlanRow.Payloads[index];
                vbtFunction.SetParamValue("PayLoad_Patt" + (index + 1), payload.Name);
            }

            vbtFunction.SetParamValue("Power_Run_Scenario", "init_NV_pl_Sweep");
            vbtFunction.SetParamValue("Wait", GetWaitTime(""));
            vbtFunction.SetParamValue("BlockType", "");

            vbtFunction.SetParamValue("PatternTimeout", "30");
            vbtFunction.SetParamValue("SELSRAM_DSSC", aiTestPlanRow.SelsramDssc);
            vbtFunction.SetParamValue("Vbump", "True");
            //vbtFunction.SetParamValue("Vbump",
            //    Regex.IsMatch(_GetInfoFromTestName(planItem.TestInstanceName, 10), "SelSr[a]*m",
            //        RegexOptions.IgnoreCase)
            //        ? "True"
            //        : "False");
        }

        private string GetWaitTime(string retention)
        {
            if (retention == "")
                return ",,,,,,,,,,,,,,";

            if (Regex.Matches(retention, ",").Count == 4)
                return ",,,,,,,,,," + retention;

            double time;
            if (double.TryParse(retention, out time))
                return ",,,,,,,,,," + time + ",,,,";

            return retention;
        }

        public SubFlowSheet GenFlowSheet()
        {
            const string flowCz = "Flow_" + Cz;
            var flowSheet = new SubFlowSheet(flowCz);
            flowSheet.FlowRows.AddHeaderRow(Cz, "");
            flowSheet.AddRow(new FlowRow { OpCode = FlowRow.OpCodeTest, Parameter = "PrintEnableWords_Header" });
            flowSheet.AddRow(new FlowRow { OpCode = FlowRow.OpCodeTest, Parameter = "PrintEnableWords" });
            flowSheet.AddRow(new FlowRow { OpCode = FlowRow.OpCodeTest, Parameter = "PrintEnableWords_Footer" });
            flowSheet.AddRow(new FlowRow { OpCode = FlowRow.OpCodeNop, Enable = Enable_Faillog_Debug });
            flowSheet.AddRow(new FlowRow { OpCode = FlowRow.OpCodeNop, Enable = "Enable_DFTLHFC_Debug" });

            var currentDataLoggingSettingType = EnumDataLoggingSettingType.NA;
            foreach (var aiTestPlanSheet in AiTestPlanSheets)
                foreach (var row in aiTestPlanSheet.Rows)
                {
                    if (!Regex.IsMatch(row.UseNotUse, "^use", RegexOptions.IgnoreCase))
                        continue;

                    var type = row.EnumDataLoggingSettingType;
                    if (type != currentDataLoggingSettingType)
                    {
                        if (type == EnumDataLoggingSettingType.NA)
                            flowSheet.AddRows(GenEnableOfTest());
                        else if (type == EnumDataLoggingSettingType.FC)
                            flowSheet.AddRows(GenEnableOfFc());
                        else if (type == EnumDataLoggingSettingType.DFC) flowSheet.AddRows(GenEnableOfDFC());
                    }

                    var flowRow = new FlowRow();
                    flowRow.ColumnA = row.SetColumnA();
                    flowRow.OpCode = FlowRow.OpCodeTest;
                    if (row.EnumAiType == EnumAiType.Shmoo_1D ||
                        row.EnumAiType == EnumAiType.Shmoo_2D)
                        flowRow.OpCode = FlowRow.OpCodeCharacterize;
                    flowRow.Parameter = row.Parameter;
                    flowSheet.AddRow(flowRow);

                    currentDataLoggingSettingType = type;
                }

            flowSheet.FlowRows.AddFooterRow(Cz, "");
            flowSheet.FlowRows.AddReturnRow();
            return flowSheet;
        }

        private List<FlowRow> GenEnableOfTest()
        {
            var flowRows = new List<FlowRow>();
            flowRows.Add(new FlowRow { OpCode = "disable-flow-word", Parameter = Debug_LVCC_VminBoundary });
            flowRows.Add(new FlowRow { OpCode = "disable-flow-word", Parameter = Enable_DFTLHFC_Debug });
            flowRows.Add(new FlowRow { OpCode = "disable-flow-word", Parameter = Enable_Faillog_Debug });
            return flowRows;
        }

        private List<FlowRow> GenEnableOfFc()
        {
            var flowRows = new List<FlowRow>();
            flowRows.Add(new FlowRow { OpCode = "enable-flow-word", Parameter = Debug_LVCC_VminBoundary });
            flowRows.Add(new FlowRow { OpCode = "disable-flow-word", Parameter = Enable_DFTLHFC_Debug });
            flowRows.Add(new FlowRow { OpCode = "disable-flow-word", Parameter = Enable_Faillog_Debug });
            return flowRows;
        }

        private List<FlowRow> GenEnableOfDFC()
        {
            var flowRows = new List<FlowRow>();
            flowRows.Add(new FlowRow { OpCode = "enable-flow-word", Parameter = Debug_LVCC_VminBoundary });
            flowRows.Add(new FlowRow { OpCode = "enable-flow-word", Parameter = Enable_DFTLHFC_Debug });
            flowRows.Add(new FlowRow { OpCode = "enable-flow-word", Parameter = Enable_Faillog_Debug });
            return flowRows;
        }

        public CharSheet GenCharSheet(PinMapSheet currentPinMapSheet, AcSpecSheet currentAcSpecSheet)
        {
            const string charCz = "Char_" + Cz;
            var charSheet = new CharSheet(charCz);
            foreach (var aiTestPlanSheet in AiTestPlanSheets)
                foreach (var row in aiTestPlanSheet.Rows)
                {
                    if (!Regex.IsMatch(row.UseNotUse, "^use", RegexOptions.IgnoreCase))
                        continue;

                    if (row.EnumAiType == EnumAiType.Shmoo_1D)
                    {
                        var charSetup = new CharSetup();
                        charSetup.ColumnA = row.SetColumnA();
                        charSetup.SetupName = row.CharName;
                        charSetup.TestMethod = CharSetupConst.TestMethodRetest;
                        var pin = row.Pins.First(x => x.IsSearch);
                        var testMethod = row.Search.Split(';').First();
                        charSetup.CharSteps.Add(CreateShmoo(row.CharName, pin,
                            testMethod, CharStepConst.ModeXShmoo,
                            currentPinMapSheet, currentAcSpecSheet));
                        charSheet.AddRow(charSetup);
                    }
                    else if (row.EnumAiType == EnumAiType.Shmoo_2D)
                    {
                        var charSetup = new CharSetup();
                        charSetup.ColumnA = row.SetColumnA();
                        charSetup.SetupName = row.CharName;
                        charSetup.TestMethod = CharSetupConst.TestMethodRetest;
                        var pins = row.Pins.Where(x => x.IsSearch).ToList();
                        if (!string.IsNullOrEmpty(row.Order))
                        {
                            var pinList = new List<Pin>();
                            foreach (var pin in row.Order.Split(','))
                                if (pins.Exists(x => x.Name.Equals(pin, StringComparison.CurrentCultureIgnoreCase)))
                                    pinList.Add(
                                        pins.Find(x => x.Name.Equals(pin, StringComparison.CurrentCultureIgnoreCase)));

                            if (pinList.Count == 2 && pins.Count == 2)
                                pins = pinList;
                        }

                        var testMethods = row.Search.Split(';');
                        var testMethodsX = testMethods.First();
                        var testMethodsY = testMethods.First();
                        var pinX = pins.ElementAt(0);
                        var pinY = pins.ElementAt(0);
                        if (testMethods.Length == 2)
                            testMethodsY = testMethods.ElementAt(1);
                        if (pins.Count == 2)
                            pinY = pins.ElementAt(1);
                        charSetup.CharSteps.Add(CreateShmoo(row.CharName, pinX,
                            testMethodsX, CharStepConst.ModeXShmoo,
                            currentPinMapSheet, currentAcSpecSheet));

                        charSetup.CharSteps.Add(CreateShmoo(row.CharName, pinY,
                            testMethodsY, CharStepConst.ModeYShmoo,
                            currentPinMapSheet, currentAcSpecSheet));
                        charSheet.AddRow(charSetup);
                    }
                }

            return charSheet;
        }

        private CharStep CreateShmoo(string charName, Pin pin, string testMethod, string method,
            PinMapSheet currentPinMapSheet, AcSpecSheet currentAcSpecSheet)
        {
            var stepName = pin.Name + "_" + method.Replace(" ", "_");
            var setup = new CharStep(charName, stepName);
            setup.Mode = method;
            var forceType = GetForceType(pin.Name, currentPinMapSheet, currentAcSpecSheet);
            setup.ParameterType = GetParameterTypeGlobalSpec(forceType);
            if (setup.ParameterType == CharStepConst.ParameterTypeAcSpec)
                setup.ParameterName = pin.Name;
            else
                setup.ParameterName = pin.ShmooName;

            setup.RangeCalcField = CharStepConst.RangeCalcFieldSteps;
            if (forceType == EnumForceType.Frequency)
            {
                var start = "";
                pin.Start.TryConvertToFreq(out start);
                setup.RangeFrom = start;
                var stop = "";
                pin.Stop.TryConvertToFreq(out stop);
                setup.RangeTo = stop;
                var step = "";
                pin.Step.TryConvertToFreq(out step);
                setup.RangeStepSize = step;
            }
            else
            {
                var start = "";
                pin.Start.TryConvertToVolt(out start);
                setup.RangeFrom = start;
                var stop = "";
                pin.Stop.TryConvertToVolt(out stop);
                setup.RangeTo = stop;
                var step = "";
                pin.Step.TryConvertToVolt(out step);
                setup.RangeStepSize = step;
            }
            //int stepSize = 0;
            //pin.TryParseRangeStepSize(out stepSize);
            //setup.RangeStepSize = stepSize.ToString();

            setup.AlgorithmName = CharStepConst.AlgorithmNameLinear;
            var arr = testMethod.Split(' ');
            if (arr.First().Equals("Jump", StringComparison.CurrentCultureIgnoreCase))
            {
                setup.AlgorithmName = arr.First();
                if (arr.Length == 2)
                    setup.AlgorithmArgs = arr.Last();
            }

            setup.PostStepArgs = "CorePower," + pin.Name;
            setup.PostStepFunction = CharStepConst.PostStepFunctionPrintShmooInfo;
            if (setup.ParameterType != CharStepConst.ParameterTypeAcSpec)
            {
                setup.ApplyToPins = pin.Name;
                setup.ApplyToPinExecMode = "Simultaneous";
            }
            setup.AxisExecutionOrder = "X-Y[-Z]";
            setup.OutputFormat = "Enhanced";
            setup.OutputTextFile = "Disable";
            setup.OutputSheet = "Disable";
            setup.SuspendDataLog = "TRUE";

            setup.OutputToDataLog = "Enable";
            setup.OutputToImmediateWin = "Disable";
            setup.OutputToOutputWin = "Disable";
            return setup;
        }

        private string GetParameterTypeGlobalSpec(EnumForceType forceType)
        {
            if (forceType == EnumForceType.Voltage)
                return CharStepConst.ParameterTypeGlobalSpec;
            if (forceType == EnumForceType.Frequency)
                return CharStepConst.ParameterTypeAcSpec;

            return CharStepConst.ParameterTypeGlobalSpec;
        }

        private EnumForceType GetForceType(string name, PinMapSheet currentPinMapSheet, AcSpecSheet currentAcSpecSheet)
        {
            if (currentPinMapSheet != null && currentPinMapSheet.IsPinExist(name))
                return EnumForceType.Voltage;
            if (currentAcSpecSheet != null && currentAcSpecSheet.IsSymbolExist(name))
                return EnumForceType.Frequency;

            return EnumForceType.Voltage;
        }


    }

    internal enum EnumForceType
    {
        Voltage,
        Current,
        Frequency
    }
}