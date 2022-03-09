using CommonLib.EpplusErrorReport;
using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlManager;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using IgxlData.Others;
using IgxlData.VBT;
using OfficeOpenXml;
using ProfileTool_PMIC.Output;
using ProfileTool_PMIC.Reader;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using ChannelMapSheet = IgxlData.IgxlSheets.ChannelMapSheet;

namespace ProfileTool_PMIC
{
    public class ProfileGeneraterMain
    {
        private string _executionProfileFile;
        private string _testProgramFile;
        private string _outputPath;
        private string _tempPath;
        private string _channelMapSheet;

        private ProfileToolForm _profileToolForm;
        private bool _isradioButtonByFlowChecked;
        private List<PinInfoRow> _pinInfoRows;

        public ProfileGeneraterMain(ProfileToolForm profileToolForm, string tempPath)
        {
            _profileToolForm = profileToolForm;
            _executionProfileFile = profileToolForm.FileOpen_ExecutionProfile.ButtonTextBox.Text;
            _testProgramFile = profileToolForm.FileOpen_TestProgram.ButtonTextBox.Text;
            _channelMapSheet = profileToolForm.ComboBox_ChanMap.Text;
            _outputPath = profileToolForm.FileOpen_OutputPath1.ButtonTextBox.Text;
            _isradioButtonByFlowChecked = profileToolForm.radioButtonByFlow.Checked;
            _tempPath = tempPath;
        }

        public void WorkFlow()
        {
            try
            {
                EpplusErrorManager.ResetError();

                _profileToolForm.AppendText(string.Format("Reading executionProfile ..."), Color.Blue);
                var executionProfileReader = new ExecutionProfileReader(_profileToolForm);
                var executionProfile = executionProfileReader.ReadIgxl90(_executionProfileFile);
                if (!_executionProfileFile.Any())
                    _profileToolForm.AppendText(string.Format("Please check format of executionProfile ..."), Color.Red);

                var exportMain = new IgxlManagerMain();
                var version = exportMain.GetVersion(_testProgramFile);

                _profileToolForm.AppendText(string.Format("Generating flow and instance ..."), Color.Blue);
                GenIgxlSheet(_tempPath, executionProfile, version);

                using (var errorReport = new ExcelPackage(new FileInfo(Path.Combine(_outputPath, "Error.xlsx"))))
                {
                    EpplusErrorManager.GenErrorReport(errorReport, null);
                    if (errorReport.Workbook.Worksheets.Count > 0)
                        errorReport.Save();
                }

                _profileToolForm.AppendText(string.Format("=> Generating the final test program ..."), Color.Blue);
                var newTestProgram = Path.Combine(_outputPath, Path.GetFileName(_testProgramFile.Replace(".", "_copy_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".")));
                exportMain.GenTestProgramByTxt(_tempPath, newTestProgram, version);
                _profileToolForm.AppendText("Done", Color.Blue);
            }
            catch(Exception ex)
            {
                _profileToolForm.AppendText("Error: " + ex.ToString(), Color.Red);
            }
        }

        private void GenIgxlSheet(string tempPath, List<ExecutionProfileRow> executionProfile, double version)
        {
            var excelIgxlReader = new IgxlSheetReader();
            _profileToolForm.AppendText(string.Format("Reading channelMap ..."), Color.Blue);
            var channelMaps = excelIgxlReader.GetIgxlSheets(tempPath, SheetType.DTChanMap).OfType<ChannelMapSheet>().ToList();
            var channelMap = channelMaps.Find(x => x.Name.Equals(_channelMapSheet, StringComparison.CurrentCultureIgnoreCase));

            _profileToolForm.AppendText(string.Format("Reading flow sheet ..."), Color.Blue);
            var flowSheets = excelIgxlReader.GetIgxlSheets(tempPath, SheetType.DTFlowtableSheet).OfType<SubFlowSheet>().ToList();
            _profileToolForm.AppendText(string.Format("Reading instance sheet ..."), Color.Blue);
            var instanaceSheets = excelIgxlReader.GetIgxlSheets(tempPath, SheetType.DTTestInstancesSheet).OfType<InstanceSheet>().ToList();
            var instanaceRows = GetValidInstanceRows(instanaceSheets);

            var allPowerPins = GetAllPowerPins(channelMap);
            _pinInfoRows = GetPinInfoRows(instanaceRows, allPowerPins);
            var dic = GenPinByFlow(_pinInfoRows, flowSheets);

            var corePowerPins = !string.IsNullOrEmpty(_profileToolForm.FileOpen_CorePowerPins.ButtonTextBox.Text)
                ? File.ReadAllLines(_profileToolForm.FileOpen_CorePowerPins.ButtonTextBox.Text).ToList() :
                allPowerPins;

            var joblistSheet = excelIgxlReader.GetIgxlSheets(tempPath, SheetType.DTJobListSheet).OfType<JoblistSheet>().First();
            var basMain = new BasManager(_tempPath);
            var profileGenerater = new ProfileGenerater(_profileToolForm, basMain);
            var instanceSheet = new InstanceSheet("TestInst_Profile");
            var newflowSheets = new List<SubFlowSheet>();
            foreach (var sheet in flowSheets)
            {
                var configFile = Path.Combine(Directory.GetCurrentDirectory(), @"Config\Excluding_Flow.txt");
                var _excludingFlows = File.ReadAllLines(configFile).ToList();
                if (_excludingFlows.Exists(x => x.Equals(sheet.Name, StringComparison.CurrentCultureIgnoreCase)))
                    continue;

                _profileToolForm.AppendText(string.Format("Generating sheet {0} ...", sheet.Name), Color.Blue);
                SubFlowSheet newflowSheet;
                if (_isradioButtonByFlowChecked)
                {
                    var pins = !_profileToolForm.checkBox_PowerPinOnly.Checked
                        ? GetCorePinsByFlowSheet(corePowerPins, dic, sheet) : corePowerPins;
                    newflowSheet = profileGenerater.WorkByFlow(sheet, executionProfile, pins, ref instanceSheet);
                }
                else
                {
                    newflowSheet = profileGenerater.WorkByInstance(sheet, ref instanceSheet, corePowerPins, executionProfile, instanaceRows, _pinInfoRows);
                }
                newflowSheets.Add(newflowSheet);
            }

            var flowSheetDic = new Dictionary<string, SubFlowSheet>();
            foreach (var sheet in newflowSheets)
                flowSheetDic.Add(sheet.Name, sheet);
            var divideFlowMain = new DivideFlowMain();
            var newSheets = divideFlowMain.WorkFlow(flowSheetDic);
            foreach (var sheet in newSheets)
                flowSheetDic.Add(sheet.Key, sheet.Value);

            foreach (var sheet in flowSheetDic)
                sheet.Value.Write(Path.Combine(_tempPath, sheet.Value.Name + ".txt"), version < 9.0 ? "2.3" : "3.0");
            instanceSheet.InstanceRows.Reverse();
            instanceSheet.WriteNew(Path.Combine(_tempPath, instanceSheet.Name + ".txt"));

            foreach (var row in joblistSheet.JobRows)
            {
                if (!row.TestInstances.Contains("TestInst_Profile"))
                    row.TestInstances += ",TestInst_Profile";
            }
            joblistSheet.Write(Path.Combine(_tempPath, joblistSheet.Name + ".txt"), joblistSheet.GetVersion());

            GenPinSummaryReport(_pinInfoRows);
            GenPinSummaryReportByFlow(dic);
        }

        private static List<string> GetCorePinsByFlowSheet(List<string> corePowerPins, Dictionary<string, string> dic, SubFlowSheet sheet)
        {
            var pins = new List<string>();
            pins.AddRange(corePowerPins);
            if (dic.ContainsKey(sheet.Name))
            {
                foreach (var pin in dic[sheet.Name].Split(','))
                {
                    if (!pins.Exists(x => x.Equals(pin, StringComparison.CurrentCultureIgnoreCase)))
                        pins.Add(pin);
                }
            }
            return pins;
        }

        private List<string> GetAllPowerPins(ChannelMapSheet channelMap)
        {
            var currentChannelMap = new CurrentChannelReader();
            currentChannelMap.ReadFile(Path.Combine(Directory.GetCurrentDirectory(), @"Config\ChannelMapping.txt"));
            var allpins = new List<string>();
            var uvspins = currentChannelMap.GetUvsPinList(channelMap);
            var vsmpins = currentChannelMap.GetVsmPinList(channelMap);
            var hexpins = currentChannelMap.GetHexVsPinList(channelMap);
            allpins.AddRange(uvspins);
            allpins.AddRange(vsmpins);
            allpins.AddRange(hexpins);
            return allpins;
        }

        private List<PinInfoRow> GetPinInfoRows(List<InstanceRow> instanaceRows, List<string> allpins)
        {
            var pinInfoRows = new List<PinInfoRow>();
            foreach (var row in instanaceRows)
            {
                var pinInfoRow = new PinInfoRow();
                pinInfoRow.InstanceName = row.TestName;
                var pins = row.GetHardipMeasurePin();
                if (pins != null)
                {
                    pinInfoRow.PinList = pins.Where(x => allpins.Exists(y => y.Equals(x, StringComparison.CurrentCultureIgnoreCase))).Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
                    pinInfoRow.Pins = pinInfoRow.PinList == null ? null : string.Join(",", pinInfoRow.PinList);
                    if (!string.IsNullOrEmpty(pinInfoRow.Pins))
                        pinInfoRows.Add(pinInfoRow);
                }
            }
            return pinInfoRows;
        }

        private void GenPinSummaryReport(List<PinInfoRow> pinInfoRows)
        {

            using (var excel = new ExcelPackage(new FileInfo(Path.Combine(_outputPath, "Summary.xlsx"))))
            {
                var wroksheet = excel.Workbook.AddSheet("PinSummary");
                wroksheet.Cells[1, 1].LoadFromCollection(pinInfoRows, true);
                excel.Save();
            }
        }

        private Dictionary<string, string> GenPinByFlow(List<PinInfoRow> pinInfoRows, List<SubFlowSheet> flowSheets)
        {
            var dic = new Dictionary<string, string>();
            foreach (var flowSheet in flowSheets)
            {
                var pins = new List<string>();
                foreach (var flowRow in flowSheet.FlowRows)
                {
                    if (pinInfoRows.Exists(x => x.InstanceName.Equals(flowRow.Parameter, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var row = pinInfoRows.Find(x => x.InstanceName.Equals(flowRow.Parameter, StringComparison.CurrentCultureIgnoreCase));
                        foreach (var pin in row.PinList)
                            pins.Add(pin);
                    }
                }
                if (!dic.ContainsKey(flowSheet.Name))
                    dic.Add(flowSheet.Name, string.Join(",", pins.Where(x => !string.IsNullOrEmpty(x)).Distinct()));
            }
            return dic;
        }

        private void GenPinSummaryReportByFlow(Dictionary<string, string> dic)
        {
            using (var excel = new ExcelPackage(new FileInfo(Path.Combine(_outputPath, "Summary.xlsx"))))
            {
                var wroksheet = excel.Workbook.AddSheet("PinSummaryByFlow");
                wroksheet.Cells[1, 1].LoadFromCollection(dic, true);
                excel.Save();
            }
        }

        private List<InstanceRow> GetValidInstanceRows(List<InstanceSheet> instanaceSheets)
        {
            var instanaceRows = new List<InstanceRow>();
            var configFile = Path.Combine(Directory.GetCurrentDirectory(), @"Config\Excluding_VBT.txt");
            var excludingVbts = File.ReadAllLines(configFile).ToList();
            foreach (var sheet in instanaceSheets)
            {
                foreach (var row in sheet.InstanceRows)
                {
                    if (!excludingVbts.Exists(x => x.Equals(row.Name, StringComparison.CurrentCultureIgnoreCase)))
                        instanaceRows.Add(row);
                }
            }
            return instanaceRows;
        }
    }
}