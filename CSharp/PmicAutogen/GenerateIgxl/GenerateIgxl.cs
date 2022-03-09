using System;
using System.Linq;
using AutomationCommon.DataStructure;
using Microsoft.Office.Interop.Excel;
using PmicAutogen.GenerateIgxl.Basic;
using PmicAutogen.GenerateIgxl.Mbist;
using PmicAutogen.GenerateIgxl.OTP;
using PmicAutogen.GenerateIgxl.PostAction;
using PmicAutogen.GenerateIgxl.PreAction;
using PmicAutogen.GenerateIgxl.Scan;
using PmicAutogen.GenerateIgxl.VbtGenTool;
using PmicAutogen.GenerateIgxl.VbtGenTool.Checker;
using PmicAutogen.InputPackages;
using PmicAutogen.Inputs.CopyXml;
using PmicAutogen.Inputs.OtpFiles;
using PmicAutogen.Inputs.ScghFile;
using PmicAutogen.Inputs.Setting;
using PmicAutogen.Inputs.TestPlan;
using PmicAutogen.Inputs.VbtGenTool;
using PmicAutogen.Local;
using System.Collections.Generic;
using PmicAutogen.Inputs.TestPlan.Reader;

namespace PmicAutogen.GenerateIgxl
{
    public class PmicGenerator
    {
        public void Run(Workbook workbook, InputPackage inputPackageAutomation)
        {
            try
            {
                if (InputFiles.SettingWorkbook != null)
                {
                    Response.Report("Reading Setting ...", MessageLevel.General, 0);
                    var settingManager = new SettingManager();
                    settingManager.CheckAll(InputFiles.SettingWorkbook);
                    StaticSetting.AddSheets(settingManager);
                }

                if (InputFiles.TestPlanWorkbook != null)
                {
                    Response.Report("Reading TestPlan ...", MessageLevel.General, 0);
                    var testPlanManager = new TestPlanManager();
                    testPlanManager.ReadAll(InputFiles.TestPlanWorkbook);
                    StaticTestPlan.AddSheets(testPlanManager);
                }

                if (InputFiles.ScghWorkbook != null)
                {
                    Response.Report("Reading SCGH file ...", MessageLevel.General, 0);
                    var scghFileManager = new ScghFileManager();
                    scghFileManager.CheckAll(InputFiles.ScghWorkbook);
                    StaticScgh.AddSheets(scghFileManager);
                }

                if (!string.IsNullOrEmpty(LocalSpecs.YamlFileName))
                {
                    Response.Report("Reading OTP Files ...", MessageLevel.General, 0);
                    var otpManager = new OtpManager();
                    otpManager.CheckAll();
                    StaticOtp.AddSheets(otpManager);
                }

                if (InputFiles.VbtGenToolWorkbooks != null && InputFiles.VbtGenToolWorkbooks.Any())
                {
                    Response.Report("Reading VbtGenTool Files ...", MessageLevel.General, 0);
                    var vbtGenToolManager = new VbtGenToolManager();
                    vbtGenToolManager.CheckAll(InputFiles.VbtGenToolWorkbooks);
                    StaticVbtGenTool.AddSheets(vbtGenToolManager);
                }

                Dictionary<string, string> pinList = new Dictionary<string, string>();
                Dictionary<string, VddLevelsRow> vDDPinList = new Dictionary<string, VddLevelsRow>();

                foreach (var ioPin in StaticTestPlan.IoLevelsSheet.Rows)
                {
                    if (pinList.ContainsKey(ioPin.Domain))
                    {
                        continue;
                    }
                    pinList.Add(ioPin.Domain, ioPin.IoLevelDate[0].Vdd);
                }

                foreach (var vddPin in StaticTestPlan.VddLevelsSheet.Rows)
                {
                    vDDPinList.Add(vddPin.WsBumpName, vddPin);
                }

                VDDRefForm vddRefForm = new VDDRefForm(pinList, vDDPinList);

                if (vddRefForm.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    foreach (var pin in vddRefForm.RefVddPins)
                    {
                        LocalSpecs.VddRefInfoList.Add(pin.Key, pin.Value);
                    }
                }

                var copyNwire = new CopyXmlFiles();
                copyNwire.Work();
                var copyConfig = new CopyIGXLConfigFiles();
                copyConfig.Work();
            }
            catch (Exception e)
            {
                Response.Report("Meet an error when reading files. " + e.Message, MessageLevel.Error, 0);
            }

            #region Post Check

            try
            {
                new PinNameChecker().Check(StaticVbtGenTool.VbtGenTestPlanSheets);

                new BitFieldChecker().Check(StaticVbtGenTool.VbtGenTestPlanSheets, StaticTestPlan.AhbRegisterMapSheet);

                new AhbRegisterChecker().Check(StaticVbtGenTool.VbtGenTestPlanSheets,
                    StaticTestPlan.AhbRegisterMapSheet);

                new PatternChecker().Check(StaticVbtGenTool.VbtGenTestPlanSheets);
            }
            catch (Exception e)
            {
                Response.Report("Meet an error in post check. " + e.Message, MessageLevel.Error, 0);
            }

            #endregion

            #region PreAction

            Response.Report("Running Pre-Action ...", MessageLevel.CheckPoint, 0);
            using (var preActionMain = new PreActionMain())
            {
                preActionMain.WorkFlow();
            }

            Response.Report("Pre-Action is Completed", MessageLevel.EndPoint, 0);

            #endregion

            #region Basic

            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Basic).Down)
            {
                Response.Report("Running Basic ...", MessageLevel.CheckPoint, 0);
                using (var basicMain = new BasicMain())
                {
                    basicMain.WorkFlow();
                }

                Response.Report("Basic is Completed", MessageLevel.EndPoint, 0);
            }

            #endregion

            #region Scan

            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Scan).Down)
            {
                Response.Report("Running Scan ...", MessageLevel.CheckPoint, 0);
                using (var scanMain = new ScanMain())
                {
                    scanMain.WorkFlow();
                }

                Response.Report("Scan is Completed", MessageLevel.EndPoint, 0);
            }

            #endregion

            #region Mbist

            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Mbist).Down)
            {
                Response.Report("Running Mbist ...", MessageLevel.CheckPoint, 0);
                using (var mbistMain = new MbistMain())
                {
                    mbistMain.WorkFlow();
                }

                Response.Report("Mbist is Completed", MessageLevel.EndPoint, 0);
            }

            #endregion

            #region OTP

            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Otp).Down)
            {
                Response.Report("Running OTP ...", MessageLevel.CheckPoint, 0);
                using (var otpMain = new OtpMain())
                {
                    otpMain.WorkFlow();
                }

                Response.Report("OTP is Completed", MessageLevel.EndPoint, 0);
            }

            #endregion

            #region VBT

            if (BlockStatus.GetAutomationBlockStatus(BlockStatus.Vbt).Down)
                foreach (var testParameterSheet in StaticVbtGenTool.TestParameterSheets)
                {
                    Response.Report(string.Format("Running {0} ...", testParameterSheet.Block), MessageLevel.CheckPoint,
                        0);
                    using (var main = new VbtGenMain(testParameterSheet))
                    {
                        main.WorkFlow();
                    }

                    Response.Report(testParameterSheet.Block + " Generation Completed", MessageLevel.EndPoint, 0);
                }

            #endregion

            #region PostAction

            Response.Report("Running Post-Action ...", MessageLevel.CheckPoint, 0);
            using (var postActionMain = new PostActionMain())
            {
                postActionMain.WorkFlow();
                Response.Report("Exporting ASCII txt files ...", MessageLevel.General, 100);
                postActionMain.PrintFlow();
            }

            Response.Report("Post-Action is Completed", MessageLevel.EndPoint, 0);

            #endregion
        }
    }
}