using AutomationCommon.DataStructure;
using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenAC;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenConti;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenContiVbt;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDc.DcInitial;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenLevel;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenPatSet;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenTimeSet;
using PmicAutogen.GenerateIgxl.HardIp;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.InputReader;
using PmicAutogen.GenerateIgxl.PreAction.Reader.ReadBasLib;
using PmicAutogen.InputPackages;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using PmicAutogen.Local.Version;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.Basic
{
    public class BasicMain : MainBase
    {
        public void WorkFlow()
        {
            try
            {
                Initialize();

                GenLevel();

                GenGlobalSpec();

                GenDcSpec();

                GetDcEnum();

                GenPatternSet();

                GenTimeSetAcSpecs();

                GenContinuity();

                GetPmicIds();

                GenPmicLeakage();

                AddIgxlSheets(IgxlSheets);

                GenDcTest();

                GenPowerPinStatusVBT();

                ExportVDDLevels();

                Response.Report("Basic Completed!", MessageLevel.General, 100);
            }
            catch (Exception e)
            {
                Response.Report("Meet an error in basic part of autogen. " + e.Message, MessageLevel.Error, 0);
            }
        }

        private void GenPowerPinStatusVBT()
        {
            var VDDLevels = InputFiles.TestPlanWorkbook.Worksheets[PmicConst.VddLevels];
            if (VDDLevels != null)
            {
                VddLevelsReader reader = new VddLevelsReader();
                VddLevelsSheet sheet = reader.ReadSheet(VDDLevels);
                List<VddLevelsRow> rows = sheet.Rows;
                //IEnumerable<string> vddPins = rows.FindAll(y => y.WsBumpName.StartsWith("VDD", StringComparison.OrdinalIgnoreCase)).Select(x => x.WsBumpName);

                //modified by terry
                List<string> vddPins = new List<string>();
                foreach (VddLevelsRow l_Row in rows)
                {
                    vddPins.Add(l_Row.WsBumpName);
                }

                if (!Directory.Exists(FolderStructure.DirCommonSheets))
                    Directory.CreateDirectory(FolderStructure.DirCommonSheets);
                string basPath = Path.Combine(FolderStructure.DirCommonSheets, "VBT_LIB_PowerLevelStatus_Read.bas");
                using (FileStream fileStream = new FileStream(basPath, FileMode.Create))
                {
                    using (StreamWriter writer = new StreamWriter(fileStream))
                    {
                        writer.WriteLine("Attribute VB_Name = \"VBT_LIB_PowerLevelStatus_Read\"");
                        writer.WriteLine("Option Explicit");
                        writer.WriteLine("Public g_sDatalog_VDDSupply As String");
                        foreach (string vddPin in vddPins)
                        {
                            writer.WriteLine("Public g_s" + vddPin + " As String");
                        }
                        //Print e_VDDPin_All block
                        writer.WriteLine("");
                        writer.WriteLine("Public Enum e_VDDPin_All");
                        int i = 0;
                        foreach (string vddPin in vddPins)
                        {
                            writer.WriteLine("    e" + vddPin + " = " + i.ToString());
                            i++;
                        }
                        writer.WriteLine("End Enum");

                        //Print e_BootUpSeq block
                        writer.WriteLine("");
                        writer.WriteLine("Public Enum e_BootUpSeq");
                        int l_intSeqNumber = getMaxSeqNumber(sheet);
                        for (int j = 0; j < l_intSeqNumber; j++)
                        {
                            int l_intIndex = j + 1;
                            writer.WriteLine("    eSEQ" + l_intIndex.ToString() + " = " + j.ToString());
                        }

                        writer.WriteLine("End Enum");

                        writer.WriteLine("");
                        writer.WriteLine("Public Function VBTPOPGen_PowerPin_LevelStatus_Read()");
                        writer.WriteLine("    On Error GoTo ErrHandler");
                        writer.WriteLine("    Dim sFuncName As String:: sFuncName = \"VBTPOPGen_PowerPin_LevelStatus_Read\"");
                        writer.WriteLine("");
                        writer.WriteLine("    Static bParsingDone As Boolean");
                        writer.WriteLine("");
                        writer.WriteLine("    If bParsingDone = False Or TheExec.Datalog.Setup.LotSetup.TestMode = Engineeringmode Then");
                        writer.WriteLine("        Call VDD_Parsing_Levels_Information");
                        writer.WriteLine("        bParsingDone = True");
                        writer.WriteLine("    End If");
                        writer.WriteLine("");
                        writer.WriteLine("    g_sDatalog_VDDSupply = \"VDD\" & g_Voltage_Corner");
                        writer.WriteLine("");
                        writer.WriteLine(@"    Call VDD_Checking_Levels_Information");
                        writer.WriteLine("");
                        foreach (string vddPin in vddPins)
                        {
                            writer.WriteLine("    g_s" + vddPin + " = \"VDD\" & Format(TheHdw.DCVI.Pins(\"" + vddPin + "\").Voltage, \"0.00\") & \"V\"");
                        }
                        writer.WriteLine("");
                        writer.WriteLine("    Exit Function");
                        writer.WriteLine("ErrHandler:");
                        writer.WriteLine("    TheExec.AddOutput \"<Error>\" + sFuncName + \":: Please check it out.\"");
                        writer.WriteLine("    TheExec.Datalog.WriteComment \"<Error>\" + sFuncName + \":: Please check it out.\"");
                        writer.WriteLine("    If AbortTest Then Exit Function Else Resume Next");
                        writer.WriteLine("");
                        writer.WriteLine("End Function");
                    }
                }

                if (!Directory.Exists(FolderStructure.DirLibPowerup))
                    Directory.CreateDirectory(FolderStructure.DirLibPowerup);
                File.Move(basPath, Path.Combine(FolderStructure.DirLibPowerup, "VBT_LIB_PowerLevelStatus_Read.bas"));
            }
        }

        /// <summary>
        /// get Max Sequence number
        /// </summary>
        /// <param name="p_sheet"></param>
        /// <returns></returns>
        private int getMaxSeqNumber(VddLevelsSheet p_sheet)
        {
            int l_intRtn = 0;
            string l_strRegex = @"^[a-zA-Z]*(?<SeqNumber>\d+)";
            Regex l_Regex = new Regex(l_strRegex, RegexOptions.IgnoreCase);

            foreach (VddLevelsRow l_Row in p_sheet.Rows)
            {
                string l_strSeq = l_Row.Seq;
                int l_intSeq = 0;
                if (int.TryParse(l_strSeq, out l_intSeq))
                {
                    if (l_intSeq > l_intRtn)
                    {
                        l_intRtn = l_intSeq;
                    }
                    else
                    {
                        //do nothing
                    }
                }
                else if (l_Regex.IsMatch(l_strSeq))
                {
                    Match l_Match = l_Regex.Match(l_strSeq);
                    string l_strtSeqNumber = l_Match.Groups["SeqNumber"].Value;
                    int l_intSeqNumber = 0;
                    if (int.TryParse(l_strtSeqNumber, out l_intSeqNumber))
                    {
                        if (l_intSeqNumber > l_intRtn)
                        {
                            l_intRtn = l_intSeqNumber;
                        }
                        else
                        {
                            //do nothing
                        }
                    }
                    else
                    {
                        //do nothing
                    }
                }
                else
                {
                    //do nothing
                }
            }

            return l_intRtn;
        }

        private void ExportVDDLevels()
        {
            var VDDLevels = InputFiles.TestPlanWorkbook.Worksheets[PmicConst.VddLevels];
            if (VDDLevels != null)
            {

                //add fixed column for VDD_Levels
                //2021.7.21 Add two fixed columns by Ze
                //2021.8.27 #task #143                
                string[,] l_FixedColumn = GetFixedColumnAfterSEQ(VDDLevels);
                List<int> skipColumnList = new List<int>();
                VddLevelsSheet vddLvlSheet = StaticTestPlan.VddLevelsSheet;
                if (vddLvlSheet.ULvHasNA)
                {
                    skipColumnList.Add(vddLvlSheet.ULvIndex - 1);
                }
                if (vddLvlSheet.UHvHasNA)
                {
                    skipColumnList.Add(vddLvlSheet.UHvIndex - 1);
                }
                InputFiles.TestPlanWorkbook.Worksheets[PmicConst.VddLevels].VddLevelExportToTxt(Path.Combine(FolderStructure.DirCommonSheets, "VDD_Levels_Information.txt"), vddLvlSheet.SeqIndex, "\t", l_FixedColumn, skipColumnList);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets, "VDD_Levels_Information");
            }
            else
            {
                //do nothing
            }
        }

        private void GetDcEnum()
        {
            var spLeakPinsCond = InputFiles.TestPlanWorkbook.Worksheets[PmicConst.SpLeakPinsCond];
            if (spLeakPinsCond != null)
            {
                var vbtTemplatePath = Directory.GetCurrentDirectory() + "\\Config\\VbtTemplate\\Conti_template.tmp";
                var vbtParser = new VbtParser(vbtTemplatePath);
                var vbtTableSheetReader = new VbtDictionaryReader();
                var dcEnumFile = Path.Combine(FolderStructure.DirVbtGenTool, spLeakPinsCond.Name + ".bas");
                var mergeFileName = Path.Combine(FolderStructure.DirLib, "PMIC", "VBT_LIB_DC_Leak_PMIC.bas");
                if (File.Exists(mergeFileName))
                {
                    vbtParser.GenVbt(vbtTableSheetReader.ReadSheet(spLeakPinsCond), spLeakPinsCond.Name, dcEnumFile);
                    var basMain = new BasMain(VersionControl.SrcInfoRows);
                    basMain.MergeBasFile(dcEnumFile, mergeFileName, mergeFileName);
                    if (File.Exists(dcEnumFile)) File.Delete(dcEnumFile);
                }
                else
                {
                    Response.Report("VBT_LIB_DC_Leak_PMIC.bas can not be found !!!", MessageLevel.Error, 100);
                }
            }

            var spContiPinsCond = InputFiles.TestPlanWorkbook.Worksheets[PmicConst.SpContiPinsCond];
            if (spContiPinsCond != null)
            {
                var vbtTemplatePath = Directory.GetCurrentDirectory() + "\\Config\\VbtTemplate\\Leak_template.tmp";
                var vbtParser = new VbtParser(vbtTemplatePath);
                var vbtTableSheetReader = new VbtDictionaryReader();
                var dcEnumFile = Path.Combine(FolderStructure.DirVbtGenTool, spContiPinsCond.Name + ".bas");
                var mergeFileName = Path.Combine(FolderStructure.DirLib, "PMIC", "VBT_LIB_DC_Conti_PMIC.bas");
                if (File.Exists(mergeFileName))
                {
                    vbtParser.GenVbt(vbtTableSheetReader.ReadSheet(spContiPinsCond), spContiPinsCond.Name, dcEnumFile);
                    var basMain = new BasMain(VersionControl.SrcInfoRows);
                    basMain.MergeBasFile(dcEnumFile, mergeFileName, mergeFileName);
                    if (File.Exists(dcEnumFile)) File.Delete(dcEnumFile);
                }
                else
                {
                    Response.Report("VBT_LIB_DC_Conti_PMIC.bas can not be found !!!", MessageLevel.Error, 100);
                }
            }
        }

        protected void GenGlobalSpec()
        {
            var globalSpecSheet = new GlobalSpecSheet(PmicConst.GlobalSpecs);
            Response.Report("Generating Global_SPEC ...", MessageLevel.General, 35);
            globalSpecSheet.AddRange(StaticTestPlan.VddLevelsSheet.GenGlbSymbol(StaticTestPlan.IfoldPowerTableSheet));
            globalSpecSheet.AddRow(new GlobalSpec("IO_Pins_GLB_Plus", "=1"));
            globalSpecSheet.AddRow(new GlobalSpec("IO_Pins_GLB_Minus", "=1"));
            globalSpecSheet.AddRange(StaticTestPlan.IoLevelsSheet.GenGlbSymbol());
            globalSpecSheet.AddRow(new GlobalSpec("SBC_Freq_Glb", "=0"));
            IgxlSheets.Add(globalSpecSheet, FolderStructure.DirGlbSpec);
        }

        protected void GenDcSpec()
        {
            Response.Report("Generating DC_SPEC ...", MessageLevel.General, 50);
            var dcInitial = new DcCatInit();
            var dcCategoryList = dcInitial.InitFlow(StaticTestPlan.PowerOverWriteSheet, StaticTestPlan.IoLevelsSheet);
            var dcSpecSheet = new DcSpecSheet(PmicConst.DcSpecs, dcCategoryList.Select(x => x.CategoryName).ToList());
            dcSpecSheet.AddRows(StaticTestPlan.VddLevelsSheet.GenDcSymbol(dcCategoryList));
            if (StaticTestPlan.IoLevelsSheet != null)
                dcSpecSheet.AddRows(StaticTestPlan.IoLevelsSheet.GenDcSpecForIoPins(dcSpecSheet.CategoryList));
            if (StaticTestPlan.PowerOverWriteSheet != null)
                StaticTestPlan.PowerOverWriteSheet.SetPowerOverWrite(dcSpecSheet);
            IgxlSheets.Add(dcSpecSheet, FolderStructure.DirDcSpec);
        }


        private void GenDcTest()
        {
            var dcTestSheets = new List<ExcelWorksheet>();
            foreach (var worksheet in InputFiles.TestPlanWorkbook.Worksheets)
                if (Regex.IsMatch(worksheet.Name, @"DCTEST_", RegexOptions.IgnoreCase) &&
                    !worksheet.Name.Equals(PmicConst.DcTestContinuity, StringComparison.CurrentCultureIgnoreCase))
                    dcTestSheets.Add(worksheet);

            foreach (var sheet in dcTestSheets)
            {
                ResetHardipData();

                var hardIpReader = new TestPlanConverter();
                var testPlanDic = hardIpReader.ReadHardipSheet(sheet);
                var block = sheet.Name.ToUpper().Replace(HardIpConstData.PrefixDctest, "").Replace(" ", "")
                    .Replace("_", "");
                var autoGenMain = new HardIpAutoGenMain();

                var instance = autoGenMain.GenInst(testPlanDic);
                var instanceSheet = new InstanceSheet("TestInst_" + HardIpConstData.PrefixDctest + block);
                instanceSheet.AddHeaderFooter();
                instanceSheet.AddRows(instance.SelectMany(x => x.InstanceRows).ToList());
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirHardIp, instanceSheet);

                var flows = autoGenMain.GenFlow(testPlanDic);
                var flowSheet = new SubFlowSheet("Flow_" + HardIpConstData.PrefixDctest + block);
                flowSheet.AddRows(flows.SelectMany(x => x.FlowRows).ToList());
                TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirHardIp, flowSheet);

                var binTableSheet = autoGenMain.GenBinTable(testPlanDic);
                var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
                foreach (var binTableRow in binTableSheet.BinTableRows.ToList())
                    binTable.AddRow(binTableRow);
            }
        }

        private void GetPmicIds()
        {
            var sheet = StaticTestPlan.PmicIdsSheet;
            if (sheet != null)
            {
                var instanceRows = sheet.GenInstance();
                var instanceSheetName = "TestInst_" + PmicConst.PmicIds;
                var instanceSheet = new InstanceSheet(instanceSheetName);
                instanceSheet.AddHeaderFooter(PmicConst.PmicIds + "_NV");
                instanceSheet.AddHeaderFooter(PmicConst.PmicIds + "_LV");
                instanceSheet.AddHeaderFooter(PmicConst.PmicIds + "_HV");
                instanceSheet.AddRows(instanceRows);
                TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirConti, instanceSheet);

                var flows = sheet.GenSubFlowSheets(PmicConst.PmicIds);
                foreach (var flow in flows)
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirConti, flow);

                var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
                var binTableRows = sheet.GenBinTableRows();
                foreach (var binTableRow in binTableRows)
                    binTable.AddRow(binTableRow);

                var fileNameTxt = Path.Combine(FolderStructure.DirConti, PmicConst.PmicIds + ".txt");
                Dictionary<int, List<string>> notMatchedPins = GetNotMatchedPinsWithVDDLevels(sheet);
                InputFiles.TestPlanWorkbook.Worksheets[PmicConst.PmicIds].ExportToTxt(fileNameTxt, extraRows: notMatchedPins);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirConti, PmicConst.PmicIds);
            }
        }

        private void GenPmicLeakage()
        {
            PmicLeakageSheet sheet = StaticTestPlan.PmicLeakageSheet;
            if (sheet == null)
                return;

            var instanceRows = sheet.GenInstance();
            var instanceSheet = new InstanceSheet(PmicConst.TestInstDcLeakage);
            instanceSheet.AddHeaderFooter(PmicConst.DcLeakage);
            instanceSheet.AddRows(instanceRows);
            TestProgram.IgxlWorkBk.AddInsSheet(FolderStructure.DirConti, instanceSheet);

            var flow = sheet.GenSubFlowSheet(PmicConst.FlowDcLeakage);
            TestProgram.IgxlWorkBk.AddSubFlowSheet(FolderStructure.DirConti, flow);

            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            var binTableRows = sheet.GenBinTableRows();
            foreach (var binTableRow in binTableRows)
                binTable.AddRow(binTableRow);

            var fileNameTxt = Path.Combine(FolderStructure.DirConti, PmicConst.PmicLeakage + ".txt");
            List<int> skipColumns = new List<int>();
            skipColumns.Add(14);
            Dictionary<int, List<string>> notMatchedPins = GetNotTestedPinsFromPinmap(sheet);
            InputFiles.TestPlanWorkbook.Worksheets[PmicConst.PmicLeakage].ExportToTxt(fileNameTxt, "\t", null, skipColumns, extraRows: notMatchedPins);
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirConti, PmicConst.PmicLeakage);
        }
        private void ResetHardipData()
        {
            var powerOverWriteSheet = new PowerOverWriteSheet();
            if (InputFiles.TestPlanWorkbook.Worksheets[PmicConst.PowerOverWrite] != null)
            {
                var dcReader = new PowerOverWriteReader();
                powerOverWriteSheet =
                    dcReader.ReadFlowMain(InputFiles.TestPlanWorkbook.Worksheets[PmicConst.PowerOverWrite]);
            }

            var testPlanData = new TestPlanData();

            HardIpDataMain.Initialize();
            HardIpDataMain.PowerOverWriteSheet = powerOverWriteSheet;
            HardIpDataMain.TestPlanData = testPlanData;
            HardIpDataMain.ReadPinMap();
        }

        protected void GenLevel()
        {
            if (StaticTestPlan.VddLevelsSheet != null && StaticTestPlan.VddLevelsSheet != null)
            {
                var levelSheets = new List<LevelSheet>();
                var levelMain = new LevelMain(StaticTestPlan.VddLevelsSheet, StaticTestPlan.IoLevelsSheet);
                levelMain.GenLevelSheets(ref levelSheets);
                levelMain.GenPinGroup(StaticTestPlan.IoLevelsSheet);

                if (StaticTestPlan.PowerOverWriteSheet != null)
                {
                    var overWriteSheets = new List<LevelSheet>();
                    foreach (var catDef in StaticTestPlan.PowerOverWriteSheet.PowerOverWrite)
                    {
                        var categoryName = catDef.CategoryName.StartsWith("Levels_", StringComparison.OrdinalIgnoreCase)
                            ? catDef.CategoryName
                            : "Levels_" + catDef.CategoryName;

                        if (levelSheets.Exists(
                            x => x.SheetName.Equals(categoryName, StringComparison.OrdinalIgnoreCase)))
                        {
                            var levelSheet = levelSheets.Find(x =>
                                x.SheetName.Equals(categoryName, StringComparison.OrdinalIgnoreCase));
                            levelMain.OverrideLevels(ref levelSheet, catDef);
                        }
                        else
                        {
                            if (levelSheets.Exists(x =>
                                x.SheetName.Equals(PmicConst.LevelsFunc, StringComparison.OrdinalIgnoreCase)))
                            {
                                var funcSheet = levelSheets.Find(x =>
                                    x.SheetName.Equals(PmicConst.LevelsFunc, StringComparison.OrdinalIgnoreCase));
                                var levelSheet = funcSheet.DeepClone();
                                levelSheet.SheetName = categoryName;
                                levelMain.OverrideLevels(ref levelSheet, catDef);
                                overWriteSheets.Add(levelSheet);
                            }
                        }
                    }

                    levelSheets.AddRange(overWriteSheets);
                }

                foreach (var levelSheet in levelSheets)
                    IgxlSheets.Add(levelSheet, FolderStructure.DirLevel);
            }
        }

        protected void GenContinuity()
        {
            Response.Report("Generating Continuity test ...", MessageLevel.General, 90);
            var dcContiMain = new DcContiMain(StaticTestPlan.DcTestContinuitySheet);
            var sheets = dcContiMain.WorkFlow();
            foreach (var sheet in sheets)
                IgxlSheets.Add(sheet.Key, sheet.Value);
        }

        protected void GenTimeSetAcSpecs()
        {
            Response.Report(string.Format("Copying Timing Set from path {0} ...", LocalSpecs.TimeSetPath),
                MessageLevel.General, 72);
            try
            {
                var timeSetGenerator = new TimeSetGenerator();
                var comTimeSetBasicSheets = timeSetGenerator.GenerateFlow(InputFiles.PatternListMap.PatternListCsvRows,
                    LocalSpecs.TimeSetPath, FolderStructure.DirTimings);

                if (!Directory.Exists(LocalSpecs.TimeSetPath))
                    Response.Report("TimeSet Path is not existed !!!", MessageLevel.Warning, 75);
                else if (comTimeSetBasicSheets.Count == 0)
                    Response.Report("Found no TimeSet File !!!", MessageLevel.Warning, 75);

                var checker = new TimeSetChecker();
                checker.CheckTimeSet(comTimeSetBasicSheets);

                foreach (var comTimeSetBasicSheet in comTimeSetBasicSheets)
                    IgxlSheets.Add(comTimeSetBasicSheet, FolderStructure.DirTimings);

                Response.Report("Generate AC Specs sheet ...", MessageLevel.General, 45);
                var acGenerator = new AcSpecsMain();
                var acSpecSheet = acGenerator.WorkFlow(comTimeSetBasicSheets);
                if (acSpecSheet != null)
                    IgxlSheets.Add(acSpecSheet, FolderStructure.DirAcSpec);
            }
            catch (Exception ex)
            {
                Response.Report("Generating TimeSet/AcSpecs failed! " + ex.Message, MessageLevel.Warning, 90);
            }
        }

        protected void GenPatternSet()
        {
            Response.Report("Generating PatSet_All ...", MessageLevel.General, 45);
            var patSetGenerator = new PatSetGenerator();
            var patternData = InputFiles.PatternListMap.PatternListCsvRows;
            patSetGenerator.GenerateFlow(patternData, LocalSpecs.PatternPath);
            var patSetAll = patSetGenerator.PatSetSheetAll;
            if (patSetAll != null)
                IgxlSheets.Add(patSetAll, FolderStructure.DirPatSetsAll);
            var patSetSub = patSetGenerator.PatSubSheetAll;
            if (patSetSub != null)
                IgxlSheets.Add(patSetSub, FolderStructure.DirPatSetsAll);
        }

        private string[,] GetFixedColumnAfterSEQ(ExcelWorksheet vddLvlWorkSheet)
        {
            VddLevelsSheet vddLvlSheet = StaticTestPlan.VddLevelsSheet;
            if (vddLvlSheet == null)
            {
                vddLvlSheet = new VddLevelsReader().ReadSheet(vddLvlWorkSheet);
            }

            int seqColumnIndex = vddLvlSheet.SeqIndex;
            if (seqColumnIndex < 0) seqColumnIndex = 0;

            if (vddLvlSheet.FinalSeqIndex < 0 && vddLvlSheet.ReferenceLevelIndex < 0)
            {
                string[,] fixedColumns = new string[,] { { "Isc", "1", (seqColumnIndex).ToString() },
                { "BW_LowCap", "0", (seqColumnIndex+1).ToString() }, { "BW_HighCap", "0", (seqColumnIndex+2).ToString() },
                { "CPBorad_BW_LowCap", "0", (seqColumnIndex+3).ToString() }, { "CPBorad_BW_HighCap", "0", (seqColumnIndex+4).ToString() },
                { "Reference_Level", "", (seqColumnIndex+5).ToString() }, { "Final_SEQ", "", (seqColumnIndex+6).ToString() } };
                return fixedColumns;
            }
            else
            {
                string[,] fixedColumns = new string[,] { { "Isc", "1", (seqColumnIndex).ToString() },
                { "BW_LowCap", "0", (seqColumnIndex+1).ToString() }, { "BW_HighCap", "0", (seqColumnIndex+2).ToString() },
                { "CPBorad_BW_LowCap", "0", (seqColumnIndex+3).ToString() }, { "CPBorad_BW_HighCap", "0", (seqColumnIndex+4).ToString() }
                };
                return fixedColumns;
            }
        }

        private Dictionary<int, List<string>> GetNotMatchedPinsWithVDDLevels(PmicIdsSheet sheet)
        {
            Dictionary<int, List<string>> notMatchedPins = new Dictionary<int, List<string>>();
            List<string> idsPins = new List<string>();
            List<string> vddLevelPins = new List<string>();
            notMatchedPins.Add(0, vddLevelPins);
            VddLevelsSheet vddLevelsSheet = StaticTestPlan.VddLevelsSheet;
            List<string> measurePins = sheet.GetMeasurePins();
            List<string> wsBumpNames = vddLevelsSheet.Rows.Select(o => o.WsBumpName.Trim()).Distinct().ToList();
            foreach (var wsBumpName in wsBumpNames)
            {
                if (!measurePins.Contains(wsBumpName))
                {
                    vddLevelPins.Add(wsBumpName);
                }
            }

            foreach (var measurePin in measurePins)
            {
                if (!wsBumpNames.Contains(measurePin))
                {
                    idsPins.Add(measurePin);
                }
            }
            if (idsPins.Any())
            {
                string pinNames = string.Join(",", idsPins.Distinct());
                string pinNameStr = string.Format(@"[{0}] not exist in VDD_Levels sheet.", pinNames);
                Response.Report("IDS sheet measure pins not match with VDD_Levels sheet." + Environment.NewLine + pinNameStr, MessageLevel.Warning, 90);
            }
            return notMatchedPins;
        }

        private Dictionary<int, List<string>> GetNotTestedPinsFromPinmap(PmicLeakageSheet sheet)
        {
            Dictionary<int, List<string>> notTestedPins = new Dictionary<int, List<string>>();
            List<string> notTestIOPins = new List<string>();
            notTestedPins.Add(0, notTestIOPins);
            List<string> notTestAnalogPins = new List<string>();
            notTestedPins.Add(1, notTestAnalogPins);

            var pinMapSheet = StaticTestPlan.IoPinMapSheet;
            List<string> allIOPins = pinMapSheet.PinList.FindAll(o => o.PinType.Equals("I/O", StringComparison.CurrentCultureIgnoreCase)).Select(o => o.PinName).ToList();
            List<string> allAnalogPins = pinMapSheet.PinList.FindAll(o => o.PinType.Equals("Analog", StringComparison.CurrentCultureIgnoreCase)
                                                                        && !o.PinName.EndsWith("_DM")
                                                                        && !o.PinName.EndsWith("_DT"))
                                                            .Select(o => o.PinName).ToList();

            List<string> testedPins = new List<string>();
            foreach (var leakageRow in sheet.Rows)
            {
                var measurePins = leakageRow.MeasurePin.Split(',').ToList();
                foreach (var measurePin in measurePins)
                {
                    testedPins.AddRange(GetAllMeasurePins(measurePin, pinMapSheet));
                }
            }

            notTestIOPins.AddRange(allIOPins.Except(testedPins.Distinct()));
            notTestAnalogPins.AddRange(allAnalogPins.Except(testedPins.Distinct()));
            return notTestedPins;
        }

        private List<string> GetAllMeasurePins(string measurePin, PinMapSheet pinMapSheet)
        {
            List<string> allMeasurePins = new List<string>();
            if (pinMapSheet.IsPinExist(measurePin))
            {
                allMeasurePins.Add(measurePin);
                return allMeasurePins;
            }
            else if (pinMapSheet.IsGroupExist(measurePin))
            {
                return pinMapSheet.GetPinsFromGroup(measurePin).Select(o => o.PinName).ToList();
            }
            return allMeasurePins;
        }
    }
}