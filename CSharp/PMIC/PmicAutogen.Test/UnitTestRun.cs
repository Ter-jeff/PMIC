using CommonLib.Enum;
using CommonLib.ErrorReport;
using CommonLib.Utility;
using CommonLib.WriteMessage;
using IgxlData.IgxlManager;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using PmicAutogen.GenerateIgxl;
using PmicAutogen.InputPackages;
using PmicAutogen.InputPackages.Base;
using PmicAutogen.InputPackages.Inputs;
using PmicAutogen.Inputs.Setting.BinNumber;
using PmicAutogen.Inputs.TestPlan;
using PmicAutogen.Local;
using PmicAutogen.Local.Version;
using PmicAutogen.Singleton;
using PmicAutogen.Test.FileDiff;
using PmicAutogen.ViewModel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace PmicAutogen.Test
{
    [TestClass]
    public class UnitTestRun
    {
        private const string DemoPath = "Demo";
        private const string InputPath = DemoPath + @"\" + "Input";
        private const string OutputPath = DemoPath + @"\" + "Output";
        private const string ExpectedPath = DemoPath + @"\" + "Expected";
        private const string KDiff3 = "KDiff3";
        private const string IgDataXml = "IGDataXML";
        private static readonly string DemoDir = Path.Combine(Directory.GetCurrentDirectory(), DemoPath);

        [TestMethod]
        public void TestMethodForUser()
        {
            TestMethod();
            var expected = DemoDir + @"\Expected";
            new FileComparisonReport().Main(LocalSpecs.TarDir, expected, true);
        }

        [TestMethod]
        //[DeploymentItem(InputPath, InputPath)]
        //[DeploymentItem(ExpectedPath, ExpectedPath)]
        //[DeploymentItem(KDiff3, KDiff3)]
        //[DeploymentItem(IgDataXml, IgDataXml)]
        public void TestMethodForCi()
        {
            TestMethod();

            var expected = DemoPath + @"\Expected";
            var fail = new FileComparisonReport().Main(LocalSpecs.TarDir, expected, false);
            if (fail)
                Assert.Fail("Mismatch !!!");
        }

        public void TestMethod()
        {
            var timeMock = new Mock<TimeProvider>();
            timeMock.SetupGet(tp => tp.Now).Returns(new DateTime(2010, 3, 11));
            TimeProvider.Current = timeMock.Object;

            Response.Progress = WriteMessage();

            LocalSpecs.IsUnitTest = true;
            LocalSpecs.Initialize();
            LocalSpecs.CurrentProject = "ProjectName";
            LocalSpecs.TestPlanFileName = DemoDir + @"\Input\ProjectName_A1_TestPlan_20220226.xlsx";
            LocalSpecs.ScghFileName = DemoDir + @"\Input\ProjectName_A0_scgh_file#1_20200207.xlsx";
            LocalSpecs.VbtGenToolFileNames = new List<string>
                {DemoDir + @"\Input\ProjectName_A0_VBTPOP_Gen_tool_MP10P_BuckSW_UVI80_DiffMeter_20200430.xlsm"};
            LocalSpecs.PatListCsvFile = DemoDir + @"\Input\ProjectName_A0_Pattern_List_Ext_20190823.csv";
            LocalSpecs.OtpFileNames = new List<string> { DemoDir + @"\Input\ProjectName_A0_otp_AVA.otp" };
            LocalSpecs.YamlFileName = DemoDir + @"\Input\ProjectName_A0_OTP_register_map.yaml";

            LocalSpecs.TarDir = DemoDir + @"\Output";
            LocalSpecs.PatternPath = DemoDir + @"\Input\K\ProjectName";
            LocalSpecs.TimeSetPath = DemoDir + @"\Input\K\ProjectName\TimeSet";
            LocalSpecs.BasLibraryPath = DemoDir + @"\Input\LibBas_PMIC\PMIC";
            LocalSpecs.SettingFile = @"Default";
            LocalSpecs.ExtraPath = DemoDir + @"\Input\ExtraSheets";

            if (Directory.Exists(LocalSpecs.TarDir))
            {
                DirectoryInfo di = new DirectoryInfo(LocalSpecs.TarDir);
                foreach (FileInfo file in di.GetFiles())
                    file.Delete();
                foreach (DirectoryInfo dir in di.GetDirectories())
                    dir.Delete(true);
            }

            TestProgram.Initialize();
            InputFiles.Initialize();
            ErrorManager.Initialize();
            BinNumberSingleton.Initialize();
            SheetStructureManager.Initialize();
            VersionControl.Initialize();
            CharSetupSingleton.Initialize();

            if (!FileCheck())
            {
                Debug.Fail("Check input files !!!!");
                return;
            }

            //Workbook _workbook = null;
            if (!File.Exists(LocalSpecs.TestPlanFileName))
            {
                MessageBox.Show(LocalSpecs.TestPlanFileName);
                Debug.Fail("No file !!!!");
            }

            var app = new Application();
            var workbook = app.Workbooks.Open(LocalSpecs.TestPlanFileName);
            try
            {
                var inputFiles = new List<Input>();
                inputFiles.Add(new InputTestPlan(new FileInfo(LocalSpecs.TestPlanFileName))
                { FullName = LocalSpecs.TestPlanFileName, FileType = InputFileType.TestPlan });
                foreach (var otpFileName in LocalSpecs.OtpFileNames)
                    inputFiles.Add(new InputOtpRegisterMap(new FileInfo(otpFileName))
                    { FullName = otpFileName, FileType = InputFileType.OtpRegisterMap });
                inputFiles.Add(new InputOtpRegisterMap(new FileInfo(LocalSpecs.YamlFileName))
                { FullName = LocalSpecs.YamlFileName, FileType = InputFileType.OtpRegisterMap });
                inputFiles.Add(new InputPatternListCsv(new FileInfo(LocalSpecs.PatListCsvFile))
                { FullName = LocalSpecs.PatListCsvFile, FileType = InputFileType.PatternListCsv });
                inputFiles.Add(new InputScgh(new FileInfo(LocalSpecs.ScghFileName))
                { FullName = LocalSpecs.ScghFileName, FileType = InputFileType.ScghPatternList });
                foreach (var vbtGenToolFileName in LocalSpecs.VbtGenToolFileNames)
                    inputFiles.Add(new InputVbtGenTool(new FileInfo(vbtGenToolFileName))
                    { FullName = vbtGenToolFileName, FileType = InputFileType.VbtGenTool });
                var inputPackage = new InputPackage();
                inputPackage.InputFiles = inputFiles;
                ViewModelMain.Instance().SetButtonStatusTrue();

                var pmicGenerator = new PmicGenerator(inputFiles);
                pmicGenerator.Run(workbook);
                var igxlItems = pmicGenerator.GetAllIgxlItems();

                var exportMain = new IgxlManagerMain();
                exportMain.GenIgxlProgram(igxlItems, LocalSpecs.TarDir, LocalSpecs.CurrentProject, TestProgram.IgxlWorkBk,
                    Response.Report, LocalSpecs.TargetIgxlVersion);

                if (ErrorManager.GetErrorCount() > 0)
                    ErrorManager.GenErrorTxt(LocalSpecs.TarDir);
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                workbook.Close(Type.Missing, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(workbook);
                Marshal.FinalReleaseComObject(app.Workbooks);
                app.Quit();
                Marshal.FinalReleaseComObject(app);
            }
        }

        private bool FileCheck()
        {
            if (!File.Exists(LocalSpecs.TestPlanFileName)) return false;
            if (!File.Exists(LocalSpecs.ScghFileName)) return false;
            foreach (var vbtGenToolFileName in LocalSpecs.VbtGenToolFileNames)
                if (!File.Exists(vbtGenToolFileName))
                    return false;
            if (!File.Exists(LocalSpecs.PatListCsvFile)) return false;
            if (!Directory.Exists(LocalSpecs.TarDir))
                Directory.Exists(LocalSpecs.TarDir);
            if (!Directory.Exists(LocalSpecs.PatternPath)) return false;
            if (!Directory.Exists(LocalSpecs.TimeSetPath)) return false;
            if (!Directory.Exists(LocalSpecs.BasLibraryPath)) return false;
            if (!Directory.Exists(LocalSpecs.ExtraPath)) return false;
            foreach (var otpFileName in LocalSpecs.OtpFileNames)
                if (!File.Exists(otpFileName))
                    return false;
            if (!File.Exists(LocalSpecs.YamlFileName)) return false;

            return true;
        }

        private Progress<ProgressStatus> WriteMessage()
        {
            var progress = new Progress<ProgressStatus>();
            progress.ProgressChanged += (o, info) =>
            {
                WriteMessage(info.Message, info.Level, info.Percentage);
            };
            return progress;
        }

        private void WriteMessage(string msg, EnumMessageLevel level = EnumMessageLevel.General, int percentage = -1)
        {
            Debug.WriteLine(msg);
        }
    }
}