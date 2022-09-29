using CommonLib.Enum;
using CommonLib.Extension;
using CommonLib.WriteMessage;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.Basic.GenAC;
using PmicAutogen.GenerateIgxl.Basic.GenConti;
using PmicAutogen.GenerateIgxl.Basic.GenContiVbt;
using PmicAutogen.GenerateIgxl.Basic.GenDc.DcInitial;
using PmicAutogen.GenerateIgxl.Basic.GenDcTest;
using PmicAutogen.GenerateIgxl.Basic.GenIds;
using PmicAutogen.GenerateIgxl.Basic.GenLeakage;
using PmicAutogen.GenerateIgxl.Basic.GenLevel;
using PmicAutogen.GenerateIgxl.Basic.GenPatSet;
using PmicAutogen.GenerateIgxl.Basic.GenTimeSet;
using PmicAutogen.GenerateIgxl.PreAction.Reader.ReadBasLib;
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
                GenLevel();

                GenGlobalSpec();

                GenDcSpec();

                GetDcEnum();

                GenPatternSet();

                GenTimeSetAcSpecs();

                GenContinuity();

                GetIds();

                GenPmicLeakage();

                GenDcTest();

                GenPowerPinStatusVbt();

                ExportVddLevels();

                TestProgram.IgxlWorkBk.AddIgxlSheets(IgxlSheets);

                Response.Report("Basic Completed!", EnumMessageLevel.General, 100);
            }
            catch (Exception e)
            {
                Response.Report("Meet an error in basic part of autogen. " + e.Message, EnumMessageLevel.Error, 0);
            }
        }

        private void GenPowerPinStatusVbt()
        {
            var vddLevels = InputFiles.TestPlanWorkbook.Worksheets[PmicConst.VddLevels];
            if (vddLevels != null)
            {
                var reader = new VddLevelsReader();
                var sheet = reader.ReadSheet(vddLevels);
                var rows = sheet.Rows;
                //IEnumerable<string> vddPins = rows.FindAll(y => y.WsBumpName.StartsWith("VDD", StringComparison.OrdinalIgnoreCase)).Select(x => x.WsBumpName);

                //modified by terry
                var vddPins = new List<string>();
                foreach (var lRow in rows) vddPins.Add(lRow.WsBumpName);

                if (!Directory.Exists(FolderStructure.DirPowerUpDown))
                    Directory.CreateDirectory(FolderStructure.DirPowerUpDown);
                var basPath = Path.Combine(FolderStructure.DirPowerUpDown, "VBT_LIB_PowerLevelStatus_Read.bas");
                using (var fileStream = new FileStream(basPath, FileMode.Create))
                {
                    using (var writer = new StreamWriter(fileStream))
                    {
                        writer.WriteLine("Attribute VB_Name = \"VBT_LIB_PowerLevelStatus_Read\"");
                        writer.WriteLine("Option Explicit");
                        writer.WriteLine("Public g_sDatalog_VDDSupply As String");
                        foreach (var vddPin in vddPins) writer.WriteLine("Public g_s" + vddPin + " As String");
                        //Print e_VDDPin_All block
                        writer.WriteLine("");
                        writer.WriteLine("Public Enum e_VDDPin_All");
                        var i = 0;
                        foreach (var vddPin in vddPins)
                        {
                            writer.WriteLine("    e" + vddPin + " = " + i);
                            i++;
                        }

                        writer.WriteLine("End Enum");

                        //Print e_BootUpSeq block
                        writer.WriteLine("");
                        writer.WriteLine("Public Enum e_BootUpSeq");
                        var lIntSeqNumber = GetMaxSeqNumber(sheet);
                        for (var j = 0; j < lIntSeqNumber; j++)
                        {
                            var lIntIndex = j + 1;
                            writer.WriteLine("    eSEQ" + lIntIndex + " = " + j);
                        }

                        writer.WriteLine("End Enum");

                        writer.WriteLine("");
                        writer.WriteLine("Public Function VBTPOPGen_PowerPin_LevelStatus_Read()");
                        writer.WriteLine("    On Error GoTo ErrHandler");
                        writer.WriteLine(
                            "    Dim sFuncName As String:: sFuncName = \"VBTPOPGen_PowerPin_LevelStatus_Read\"");
                        writer.WriteLine("");
                        writer.WriteLine("    Static bParsingDone As Boolean");
                        writer.WriteLine("");
                        writer.WriteLine(
                            "    If bParsingDone = False Or TheExec.Datalog.Setup.LotSetup.TestMode = Engineeringmode Then");
                        writer.WriteLine("        Call VDD_Parsing_Levels_Information");
                        writer.WriteLine("        bParsingDone = True");
                        writer.WriteLine("    End If");
                        writer.WriteLine("");
                        writer.WriteLine("    g_sDatalog_VDDSupply = \"VDD\" & GetVoltageCorner");
                        //writer.WriteLine("    g_sDatalog_VDDSupply = \"VDD\" & g_Voltage_Corner");
                        writer.WriteLine("");
                        writer.WriteLine(@"    Call VDD_Checking_Levels_Information");
                        writer.WriteLine("");
                        foreach (var vddPin in vddPins)
                            writer.WriteLine("    g_s" + vddPin + " = \"VDD\" & Format(TheHdw.DCVI.Pins(\"" + vddPin +
                                             "\").Voltage, \"0.00\") & \"V\"");
                        writer.WriteLine("");
                        writer.WriteLine("    Exit Function");
                        writer.WriteLine("ErrHandler:");
                        writer.WriteLine("    TheExec.AddOutput \"<Error>\" + sFuncName + \":: Please check it out.\"");
                        writer.WriteLine(
                            "    TheExec.Datalog.WriteComment \"<Error>\" + sFuncName + \":: Please check it out.\"");
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

        private int GetMaxSeqNumber(VddLevelsSheet pSheet)
        {
            var lIntRtn = 0;
            var lStrRegex = @"^[a-zA-Z]*(?<SeqNumber>\d+)";
            var lRegex = new Regex(lStrRegex, RegexOptions.IgnoreCase);

            foreach (var lRow in pSheet.Rows)
            {
                var lStrSeq = lRow.Seq;
                int lIntSeq;
                if (int.TryParse(lStrSeq, out lIntSeq))
                {
                    if (lIntSeq > lIntRtn) lIntRtn = lIntSeq;
                }
                else if (lRegex.IsMatch(lStrSeq))
                {
                    var lMatch = lRegex.Match(lStrSeq);
                    var lStrtSeqNumber = lMatch.Groups["SeqNumber"].Value;
                    int lIntSeqNumber;
                    if (int.TryParse(lStrtSeqNumber, out lIntSeqNumber))
                        if (lIntSeqNumber > lIntRtn)
                            lIntRtn = lIntSeqNumber;
                }
            }

            return lIntRtn;
        }

        private void ExportVddLevels()
        {
            var vddLevels = InputFiles.TestPlanWorkbook.Worksheets[PmicConst.VddLevels];
            if (vddLevels != null)
            {
                //add fixed column for VDD_Levels
                //2021.7.21 Add two fixed columns by Ze
                //2021.8.27 #task #143                
                var lFixedColumn = GetFixedColumnAfterSeq(vddLevels);
                //List<int> skipColumnList = new List<int>();
                var vddLvlSheet = StaticTestPlan.VddLevelsSheet;
                //if (vddLvlSheet.ULvAllNA)
                //{
                //    skipColumnList.Add(vddLvlSheet.ULvIndex - 1);
                //}
                //if (vddLvlSheet.UHvAllNA)
                //{
                //    skipColumnList.Add(vddLvlSheet.UHvIndex - 1);
                //}
                InputFiles.TestPlanWorkbook.Worksheets[PmicConst.VddLevels].VddLevelExportToTxt(Path.Combine(
                    FolderStructure.DirPowerUpDown,
                    "VDD_Levels_Information.txt"), vddLvlSheet.SeqIndex, "\t", lFixedColumn);
                TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirPowerUpDown, "VDD_Levels_Information");
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
                var dcEnumFile = Path.Combine(FolderStructure.DirOtp, spLeakPinsCond.Name + ".bas");
                var mergeFileName = Path.Combine(FolderStructure.DirOtherWaitForClassify, "PMIC",
                    "VBT_LIB_DC_Leak_PMIC.bas");
                if (File.Exists(mergeFileName))
                {
                    vbtParser.GenVbt(vbtTableSheetReader.ReadSheet(spLeakPinsCond), spLeakPinsCond.Name, dcEnumFile);
                    var basMain = new BasMain(VersionControl.SrcInfoRows);
                    basMain.MergeBasFile(dcEnumFile, mergeFileName, mergeFileName);
                    if (File.Exists(dcEnumFile)) File.Delete(dcEnumFile);
                }
                else
                {
                    Response.Report("VBT_LIB_DC_Leak_PMIC.bas can not be found !!!", EnumMessageLevel.Error, 100);
                }
            }

            var spContiPinsCond = InputFiles.TestPlanWorkbook.Worksheets[PmicConst.SpContiPinsCond];
            if (spContiPinsCond != null)
            {
                var vbtTemplatePath = Directory.GetCurrentDirectory() + "\\Config\\VbtTemplate\\Leak_template.tmp";
                var vbtParser = new VbtParser(vbtTemplatePath);
                var vbtTableSheetReader = new VbtDictionaryReader();
                var dcEnumFile = Path.Combine(FolderStructure.DirOtp, spContiPinsCond.Name + ".bas");
                var mergeFileName = Path.Combine(FolderStructure.DirOtherWaitForClassify, "PMIC",
                    "VBT_LIB_DC_Conti_PMIC.bas");
                if (File.Exists(mergeFileName))
                {
                    vbtParser.GenVbt(vbtTableSheetReader.ReadSheet(spContiPinsCond), spContiPinsCond.Name, dcEnumFile);
                    var basMain = new BasMain(VersionControl.SrcInfoRows);
                    basMain.MergeBasFile(dcEnumFile, mergeFileName, mergeFileName);
                    if (File.Exists(dcEnumFile)) File.Delete(dcEnumFile);
                }
                else
                {
                    Response.Report("VBT_LIB_DC_Conti_PMIC.bas can not be found !!!", EnumMessageLevel.Error, 100);
                }
            }
        }

        protected void GenGlobalSpec()
        {
            var globalSpecSheet = new GlobalSpecSheet(PmicConst.GlobalSpecs);
            Response.Report("Generating Global_SPEC ...", EnumMessageLevel.General, 35);
            globalSpecSheet.AddRows(StaticTestPlan.VddLevelsSheet.GenGlbSymbol(StaticTestPlan.IfoldPowerTableSheet));
            globalSpecSheet.AddRow(new GlobalSpec("IO_Pins_GLB_Plus", "=1"));
            globalSpecSheet.AddRow(new GlobalSpec("IO_Pins_GLB_Minus", "=1"));
            globalSpecSheet.AddRows(StaticTestPlan.IoLevelsSheet.GenGlbSymbol());
            globalSpecSheet.AddRow(new GlobalSpec("SBC_Freq_Glb", "=0"));
            IgxlSheets.Add(globalSpecSheet, FolderStructure.DirGlbSpec);
        }

        protected void GenDcSpec()
        {
            Response.Report("Generating DC_SPEC ...", EnumMessageLevel.General, 50);
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
            foreach (var dcTestSheet in StaticTestPlan.DcTestSheets)
            {
                var dcTestMain = new DcTestMain(dcTestSheet);
                var igxlSheets = dcTestMain.Workflow();
                foreach (var igxlSheet in igxlSheets)
                    IgxlSheets.Add(igxlSheet.Key, igxlSheet.Value);
            }
        }

        private void GetIds()
        {
            if (StaticTestPlan.PmicIdsSheet == null)
                return;
            var idsMain = new IdsMain(StaticTestPlan.PmicIdsSheet);
            var igxlSheets = idsMain.Workflow();
            foreach (var igxlSheet in igxlSheets)
                IgxlSheets.Add(igxlSheet.Key, igxlSheet.Value);
        }

        private void GenPmicLeakage()
        {
            if (StaticTestPlan.PmicLeakageSheet == null)
                return;
            var leakageMain = new LeakageMain(StaticTestPlan.PmicLeakageSheet);
            var igxlSheets = leakageMain.Workflow();
            foreach (var igxlSheet in igxlSheets)
                IgxlSheets.Add(igxlSheet.Key, igxlSheet.Value);
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
            Response.Report("Generating Continuity test ...", EnumMessageLevel.General, 90);
            var dcContiMain = new DcContiMain(StaticTestPlan.DcTestContinuitySheet);
            var igxlSheets = dcContiMain.WorkFlow();
            foreach (var igxlSheet in igxlSheets)
                IgxlSheets.Add(igxlSheet.Key, igxlSheet.Value);
        }

        protected void GenTimeSetAcSpecs()
        {
            Response.Report(string.Format("Copying Timing Set from path {0} ...", LocalSpecs.TimeSetPath),
                EnumMessageLevel.General, 72);
            try
            {
                var timeSetGenerator = new TimeSetGenerator();
                var comTimeSetBasicSheets = timeSetGenerator.GenerateFlow(InputFiles.PatternListMap.PatternListCsvRows,
                    LocalSpecs.TimeSetPath, FolderStructure.DirTimings);

                if (!Directory.Exists(LocalSpecs.TimeSetPath))
                    Response.Report("TimeSet Path is not existed !!!", EnumMessageLevel.Warning, 75);
                else if (comTimeSetBasicSheets.Count == 0)
                    Response.Report("Found no TimeSet File !!!", EnumMessageLevel.Warning, 75);

                var checker = new TimeSetChecker();
                checker.CheckTimeSet(comTimeSetBasicSheets);

                foreach (var comTimeSetBasicSheet in comTimeSetBasicSheets)
                    IgxlSheets.Add(comTimeSetBasicSheet, FolderStructure.DirTimings);

                Response.Report("Generate AC Specs sheet ...", EnumMessageLevel.General, 45);
                var acGenerator = new AcSpecsMain();
                var acSpecSheet = acGenerator.WorkFlow(comTimeSetBasicSheets);
                if (acSpecSheet != null)
                    IgxlSheets.Add(acSpecSheet, FolderStructure.DirAcSpec);
            }
            catch (Exception ex)
            {
                Response.Report("Generating TimeSet/AcSpecs failed! " + ex.Message, EnumMessageLevel.Warning, 90);
            }
        }

        protected void GenPatternSet()
        {
            Response.Report("Generating PatSet_All ...", EnumMessageLevel.General, 45);
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

        private string[,] GetFixedColumnAfterSeq(ExcelWorksheet vddLvlWorkSheet)
        {
            var vddLvlSheet = StaticTestPlan.VddLevelsSheet;
            if (vddLvlSheet == null) vddLvlSheet = new VddLevelsReader().ReadSheet(vddLvlWorkSheet);

            var seqColumnIndex = vddLvlSheet.SeqIndex;
            if (seqColumnIndex < 0) seqColumnIndex = 0;

            if (vddLvlSheet.FinalSeqIndex < 0 && vddLvlSheet.ReferenceLevelIndex < 0)
            {
                string[,] fixedColumns =
                {
                    {"Isc", "1", seqColumnIndex.ToString()},
                    {"BW_LowCap", "0", (seqColumnIndex + 1).ToString()},
                    {"BW_HighCap", "0", (seqColumnIndex + 2).ToString()},
                    {"CPBorad_BW_LowCap", "0", (seqColumnIndex + 3).ToString()},
                    {"CPBorad_BW_HighCap", "0", (seqColumnIndex + 4).ToString()},
                    {"Reference_Level", "", (seqColumnIndex + 5).ToString()},
                    {"Final_SEQ", "", (seqColumnIndex + 6).ToString()}
                };
                return fixedColumns;
            }
            else
            {
                string[,] fixedColumns =
                {
                    {"Isc", "1", seqColumnIndex.ToString()},
                    {"BW_LowCap", "0", (seqColumnIndex + 1).ToString()},
                    {"BW_HighCap", "0", (seqColumnIndex + 2).ToString()},
                    {"CPBorad_BW_LowCap", "0", (seqColumnIndex + 3).ToString()},
                    {"CPBorad_BW_HighCap", "0", (seqColumnIndex + 4).ToString()}
                };
                return fixedColumns;
            }
        }
    }
}