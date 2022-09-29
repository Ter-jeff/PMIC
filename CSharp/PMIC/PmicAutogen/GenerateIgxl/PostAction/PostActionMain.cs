using CommonLib.Enum;
using CommonLib.WriteMessage;
using IgxlData.IgxlReader;
using IgxlData.Others;
using PmicAutogen.GenerateIgxl.PostAction.GenGlobal;
using PmicAutogen.GenerateIgxl.PostAction.GenJob;
using PmicAutogen.GenerateIgxl.PostAction.GenMainFlow;
using PmicAutogen.GenerateIgxl.PostAction.GenTestNumber;
using PmicAutogen.GenerateIgxl.PostAction.ModifyChannelMap;
using PmicAutogen.GenerateIgxl.PostAction.PostCheck;
using PmicAutogen.GenerateIgxl.PostAction.SPMI;
using PmicAutogen.Inputs.CopyLib;
using PmicAutogen.Local;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.PostAction
{
    public class PostActionMain : MainBase
    {
        public void WorkFlow()
        {
            try
            {
                //Add N/C in ChannelMap
                var channelMapMain = new ChannelMapPostMain();
                channelMapMain.WorkFlow();

                //Move pinGroup type is TimeDomain to the last item
                var sortPin = new SortPinMap.SortPinMap();
                sortPin.Sort(TestProgram.IgxlWorkBk.PinMapPair.Value);

                //Gen main flow
                var mainFlowGen = new MainFlowMain();
                var igxlSheets = mainFlowGen.WorkFlow();
                foreach (var igxlSheet in igxlSheets)
                    IgxlSheets.Add(igxlSheet.Key, igxlSheet.Value);

                //Gen JobList
                Response.Report("Generating JobList Sheets ...", EnumMessageLevel.General, 40);
                var jobListGen = new JobListMain(igxlSheets);
                var jobList = jobListGen.WorkFlow();
                IgxlSheets.Add(jobList.Key, jobList.Value);

                var mainSpmi = new SpmiMain();
                mainSpmi.WorkFlow();

                //Add GlobalSpec for Characterization
                var genCharacterizationGlobalSpec = new GenCharacterizationGlobalSpec();
                genCharacterizationGlobalSpec.ExtendGlobalSpec();

                //Add GlobalSpec for VddLevels Seq='x' 
                var genVddLevelsXSeqGlobalSpec = new GenVddLevelsXSeqGlobalSpec();
                genVddLevelsXSeqGlobalSpec.ExtendGlobalSpec();

                //Add test Number
                var testNumber = new TestNumberMain();
                var nonTestNumberSheets = testNumber.WorkFlow();
                if (nonTestNumberSheets.Count > 0)
                    Response.Report(string.Format("SubFlow sheets didn't have specify mapping for test number:\n" +
                                                  string.Join("\n", nonTestNumberSheets)), EnumMessageLevel.Warning, 60);

                //Divide flow when row > 100000
                var divideFlow = new DivideFlowMain();
                var subFlowSheets = divideFlow.WorkFlow(TestProgram.IgxlWorkBk.SubFlowSheets);
                foreach (var sheet in subFlowSheets)
                    IgxlSheets.Add(sheet.Key, sheet.Value);

                var duplicateInstanceChecker = new DuplicateInstanceChecker();
                duplicateInstanceChecker.WorkFlow();

                CheckTimeSetLength();

                TestProgram.IgxlWorkBk.AddIgxlSheets(IgxlSheets);
            }
            catch (Exception e)
            {
                Response.Report("Meet an error in post-action of autogen. " + e.Message, EnumMessageLevel.Error, 0);
            }
        }

        private void CheckTimeSetLength()
        {
            var ignoreTimeList = TestProgram.IgxlWorkBk.TimeSetSheets.Keys.Select(a => a)
                .Where(x => Path.GetFileName(x).ToString().Length >= 32).ToList();
            if (ignoreTimeList.Count > 0)
                Response.Report("The Timing set list below will not be included in test program!", EnumMessageLevel.Error,
                    0);
            foreach (var item in ignoreTimeList)
            {
                Response.Report(item + ": The length of Timing set name exceeds the maximum 31 chars ",
                    EnumMessageLevel.Error, 0);
                TestProgram.IgxlWorkBk.AllIgxlSheets.Remove(item);
            }
        }

        public void PrintFlow()
        {
            TestProgram.IgxlWorkBk.PrintAllSheets(LocalSpecs.TargetIgxlVersion);

            CopyFromExtraSheets(LocalSpecs.ExtraPath);

            TestProgram.IgxlWorkBk.PrintBinTable(LocalSpecs.TargetIgxlVersion);

            var copyLib = new CopyLibFiles();
            copyLib.Work();
        }

        private void CopyFromExtraSheets(string extraFolder)
        {
            if (Directory.Exists(extraFolder))
            {
                var igxlConfigFolder = Path.Combine(extraFolder, "IGXLConfig");
                ////Copy all the igxl config files to output folder
                //if (Directory.Exists(igxlConfigFolder))
                //{
                //    FileInfo[] files = new DirectoryInfo(igxlConfigFolder).GetFiles();
                //    foreach(FileInfo file in files)
                //    {
                //        string destFile = Path.Combine(FolderStructure.DirIgLink, file.Name);
                //        File.Copy(file.FullName, destFile,true);
                //    }
                //    DirectoryInfo[] dirs = new DirectoryInfo(igxlConfigFolder).GetDirectories();
                //    foreach (DirectoryInfo directory in dirs)
                //    {               
                //        string destDir = Path.Combine(FolderStructure.DirIgLink, directory.Name);
                //        CopyDir(directory.FullName, destDir);                       
                //    }
                //}

                var dir = new DirectoryInfo(extraFolder);
                var extraTxts = dir.GetFiles("*.txt", SearchOption.AllDirectories);
                foreach (var extraTxt in extraTxts)
                {
                    if (Regex.IsMatch(extraTxt.Name, "^Channel*", RegexOptions.IgnoreCase))
                        continue;

                    var type = new IgxlSheetReader().GetIgxlSheetTypeByFile(extraTxt.FullName);
                    if (type == SheetTypes.DTBintablesSheet)
                    {
                        var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
                        var rows = new ReadBinTableSheet().GetSheet(extraTxt.FullName).BinTableRows;
                        foreach (var row in rows)
                        {
                            var toolGeneratedRow = binTable.BinTableRows.Find(s =>
                                s.Name.Equals(row.Name, StringComparison.OrdinalIgnoreCase));
                            if (toolGeneratedRow != null)
                                binTable.RemoveRow(toolGeneratedRow);
                            binTable.AddRow(row);
                        }
                    }
                    else
                    {
                        var flag = false;
                        foreach (var igxlSheetPair in TestProgram.IgxlWorkBk.AllIgxlSheets)
                        {
                            var fileName = Path.GetFileName(igxlSheetPair.Key + ".txt");
                            if (fileName.Equals(extraTxt.Name, StringComparison.OrdinalIgnoreCase))
                            {
                                flag = true;
                                File.Copy(extraTxt.FullName, igxlSheetPair.Key + ".txt", true);
                                var outString =
                                    string.Format("ExtraSheet:{0} Conflict with generate output... Please Check",
                                        extraTxt.Name);
                                Response.Report(outString, EnumMessageLevel.Warning, 0);
                                break;
                            }
                        }

                        if (!flag)
                        {
                            if (extraTxt.Name.Equals("Flow_Init_EnableWd.txt",
                                    StringComparison.CurrentCultureIgnoreCase))
                                ExtraFileCopy(extraTxt, FolderStructure.DirMainFlow);
                            else if (extraTxt.Name.Equals("QQ_LimitSheet.txt",
                                         StringComparison.CurrentCultureIgnoreCase))
                                ExtraFileCopy(extraTxt, FolderStructure.DirLimitSheet);
                            else if (extraTxt.Name.Equals("References.txt", StringComparison.CurrentCultureIgnoreCase))
                                ExtraFileCopy(extraTxt, FolderStructure.DirReference);
                            else if (extraTxt.Name.Equals("TIMESET_PMIC_Dummy.txt",
                                         StringComparison.CurrentCultureIgnoreCase))
                                ExtraFileCopy(extraTxt, FolderStructure.DirTimings);
                            else if (extraTxt.Name.Equals("TestInst_Common.txt",
                                         StringComparison.CurrentCultureIgnoreCase))
                                ExtraFileCopy(extraTxt, FolderStructure.DirTestInstance);
                            else if (extraTxt.Name.Contains("OTP") ||
                                     extraTxt.Name.ToLowerInvariant().Contains("version"))
                                ExtraFileCopy(extraTxt, FolderStructure.DirOtp);
                            else
                                ExtraFileCopy(extraTxt, FolderStructure.DirOtherWaitForClassify);
                        }
                    }
                }
            }
        }

        private void CopyDir(string srcPath, string aimPath)
        {
            try
            {
                if (aimPath[aimPath.Length - 1] != Path.DirectorySeparatorChar)
                    aimPath += Path.DirectorySeparatorChar;
                if (!Directory.Exists(aimPath)) Directory.CreateDirectory(aimPath);
                var fileList = Directory.GetFileSystemEntries(srcPath);
                foreach (var file in fileList)
                    if (Directory.Exists(file))
                        CopyDir(file, aimPath + Path.GetFileName(file));
                    else
                        File.Copy(file, aimPath + Path.GetFileName(file), true);
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        private void ExtraFileCopy(FileInfo extraTxt, string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            File.Copy(extraTxt.FullName, Path.Combine(path, extraTxt.Name),
                true);
            TestProgram.NonIgxlSheetsList.Add(path, extraTxt.Name.Replace(".txt", ""));
        }
    }
}