using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.DataStructure;
using AutomationCommon.EpplusErrorReport;
using IgxlData.IgxlReader;
using IgxlData.Others;
using PmicAutogen.GenerateIgxl.PostAction.GenGlobal;
using PmicAutogen.GenerateIgxl.PostAction.GenJob;
using PmicAutogen.GenerateIgxl.PostAction.GenMainFlow;
using PmicAutogen.GenerateIgxl.PostAction.GenTestNumber;
using PmicAutogen.GenerateIgxl.PostAction.ModifyChannelMap;
using PmicAutogen.GenerateIgxl.PostAction.PostCheck;
using PmicAutogen.GenerateIgxl.PostAction.SPMI;
using PmicAutogen.InputPackages;
using PmicAutogen.Local;
using PmicAutogen.Local.Version;
using Teradyne.Oasis.IGData;

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
                mainFlowGen.WorkFlow();

                //Gen JobList
                Response.Report("Generating JobList Sheets ...", MessageLevel.General, 40);
                var jobListGen = new JobListMain();
                jobListGen.WorkFlow();

                //task #144 and #145
                var mainSPMI = new SPMIMain();
                mainSPMI.WorkFlow();

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
                                      string.Join("\n", nonTestNumberSheets)), MessageLevel.Warning, 60);


                //Divide flow when row > 100000
                var divideFlow = new DivideFlowMain();
                var subFlowSheets = divideFlow.WorkFlow(TestProgram.IgxlWorkBk.SubFlowSheets);
                foreach (var sheet in subFlowSheets)
                    TestProgram.IgxlWorkBk.AddSubFlowSheet(sheet);

                var duplicateInstanceChecker = new DuplicateInstanceChecker();
                duplicateInstanceChecker.WorkFlow();

                var copyLib = new CopyLibFiles();
                copyLib.Work();

                CheckTimeSetLength();

                //Generate Error Report
                var rtErrorTable = EpplusErrorManager.GetErrorInfo();
                foreach (DataRow row in rtErrorTable.Rows)
                {
                    var infoRow = new SrcInfoRow(row[0].ToString(), row[1].ToString());
                    TestProgram.SourceInfoSheet.AddSrcInfo(infoRow);
                }
            }
            catch (Exception e)
            {
                Response.Report("Meet an error in post-action of autogen. " + e.Message, MessageLevel.Error, 0);
            }
        }

        private void CheckTimeSetLength()
        {
            var ignoreTimeList = TestProgram.IgxlWorkBk.TimeSetSheets.Keys.Select(a => a)
                .Where(x => Path.GetFileName(x).ToString().Length >= 32).ToList();
            if (ignoreTimeList.Count > 0)
                Response.Report("The Timing set list below will not be included in test program!", MessageLevel.Error,
                    0);
            foreach (var item in ignoreTimeList)
            {
                Response.Report(item + ": The length of Timing set name exceeds the maximum 31 chars ",
                    MessageLevel.Error, 0);
                TestProgram.IgxlWorkBk.AllIgxlSheets.Remove(item);
            }
        }

        public void PrintFlow()
        {
            TestProgram.IgxlWorkBk.PrintAllSheets(LocalSpecs.TargetIgxlVersion);

            CopyFromExtraSheets(LocalSpecs.ExtraPath);

            TestProgram.IgxlWorkBk.PrintBinTable(LocalSpecs.TargetIgxlVersion);

            //TestProgram.SourceInfoSheet.Print(FolderStructure.DirCommonSheets);
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirCommonSheets,
                Path.GetFileNameWithoutExtension(TestProgram.SourceInfoSheet.FileName));

            CopyLibFiles copyLibFiles = new CopyLibFiles();
            copyLibFiles.FileStructurePostAction();
        }

        private void CopyFromExtraSheets(string extraFolder)
        {
            if (Directory.Exists(extraFolder))
            {
                string igxlConfigFolder = Path.Combine(extraFolder, "IGXLConfig");
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
                var mExtraList = dir.GetFiles("*.txt", SearchOption.AllDirectories);
                foreach (var extraTxt in mExtraList)
                {
                    //if (extraTxt.FullName.StartsWith(igxlConfigFolder))
                    //    continue;
                    if (Regex.IsMatch(extraTxt.Name, "^Channel*", RegexOptions.IgnoreCase))
                        continue;

                    var type = new IgxlSheetReader().GetIgxlSheetTypeByFile(extraTxt.FullName);
                    if (type == Sheet.SheetTypes.DTBintablesSheet)
                    {
                        IgxlData.IgxlSheets.BinTableSheet binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
                        var rows = new ReadBintableSheet().GetSheet(extraTxt.FullName).BinTableRows;
                        foreach(IgxlData.IgxlBase.BinTableRow row in rows)
                        {
                            IgxlData.IgxlBase.BinTableRow toolGeneratedRow = binTable.BinTableRows.Find(s => s.Name.Equals(row.Name, StringComparison.OrdinalIgnoreCase));
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
                                var outString = string.Format("ExtraSheet:{0} Conflict with generate output... Please Check", extraTxt.Name);
                                Response.Report(outString, MessageLevel.Warning, 0);
                                break;
                            }
                        }

                        if (!flag)
                        {
                            if (extraTxt.Name.Equals("Flow_Init_EnableWd.txt", StringComparison.CurrentCultureIgnoreCase))
                            {
                                ExtraFileCopy(extraTxt, FolderStructure.DirMain);
                            }
                            else if (extraTxt.Name.Equals("QQ_LimitSheet.txt", StringComparison.CurrentCultureIgnoreCase))
                            {
                                ExtraFileCopy(extraTxt, FolderStructure.DirLimitSheet);
                            }
                            else if (extraTxt.Name.Equals("References.txt", StringComparison.CurrentCultureIgnoreCase))
                            {
                                ExtraFileCopy(extraTxt, FolderStructure.DirReference);
                            }
                            else if (extraTxt.Name.Equals("TIMESET_PMIC_Dummy.txt", StringComparison.CurrentCultureIgnoreCase))
                            {
                                ExtraFileCopy(extraTxt, FolderStructure.DirTimings);
                            }
                            else if (extraTxt.Name.Equals("TestInst_Common.txt", StringComparison.CurrentCultureIgnoreCase))
                            {
                                ExtraFileCopy(extraTxt, FolderStructure.DirTestInstance);
                            }
                            else if (extraTxt.Name.Contains("OTP") || extraTxt.Name.ToLowerInvariant().Contains("version"))
                            {
                                ExtraFileCopy(extraTxt, FolderStructure.DirOtp);
                            }
                            else
                            {
                                ExtraFileCopy(extraTxt, FolderStructure.DirOtherWaitForClassify);
                            }
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
                {
                    aimPath += System.IO.Path.DirectorySeparatorChar;
                }
                if (!Directory.Exists(aimPath))
                {
                    Directory.CreateDirectory(aimPath); 
                }               
                string[] fileList = Directory.GetFileSystemEntries(srcPath);
                foreach (string file in fileList)
                {
                    if (Directory.Exists(file))
                    {
                        CopyDir(file, aimPath + Path.GetFileName(file));
                    }
                    else
                    {
                        File.Copy(file, aimPath + System.IO.Path.GetFileName(file), true);
                    }
                }
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
