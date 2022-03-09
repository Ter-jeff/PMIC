using PmicAutomation.MyControls;
using PmicAutomation.Utility.PA.Base;
using PmicAutomation.Utility.PA.Check;
using PmicAutomation.Utility.PA.Function;
using PmicAutomation.Utility.PA.Input;
using Library.Function;
using Library.Function.ErrorReport;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using IgxlData.IgxlSheets;
using PmicAutomation.MyControls;
using System.Reflection;
using System.Xml.Serialization;

namespace PmicAutomation.Utility.PA
{
    public class PaMain
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly List<string> _files;
        private readonly UflexConfig _uflexConfig;
        private readonly string _outputPath;
        private readonly string _DGSReferenceFile;
        private readonly List<string> _hexVs;
        private readonly string _device;
        private string _igxlVersion;

        private readonly Dictionary<string, PaSheet> _paSheets = new Dictionary<string, PaSheet>();
        private List<PaGroup> _missingPaGroup = new List<PaGroup>();
        private List<PaGroup> _paGroups = new List<PaGroup>();
        private bool sharePinEnable;
        private PaSheet _DGSReferenceSheet;

        public PaMain(List<string> files, string uflexConfigPath, string outputPath, List<string> hexVs, string device)
        {
            _files = files;
            _outputPath = outputPath;
            _uflexConfig = UflexConfigReader.GetXml(uflexConfigPath);
            _hexVs = hexVs;
            _device = device;
        }

        public PaMain(PaParser paParser, UflexConfig uflexConfig)
        {
            _appendText = paParser.AppendText;
            _files = paParser.FileOpen_PA.ButtonTextBox.Text.Split(',').ToList();
            _outputPath = paParser.FileOpen_OutputPath.ButtonTextBox.Text;
            _DGSReferenceFile = paParser.FileOpen_DGSReference.ButtonTextBox.Text.Trim();
            _device = paParser.comboBox_Device.Text;
            _hexVs = paParser.textBox_HexVs.Text.Split(',').ToList();
            _uflexConfig = uflexConfig;
            _igxlVersion = paParser.IGXL_Version;
            sharePinEnable = paParser.checkBox_SharePinEnable.Checked;
            ErrorManager.ResetError();
        }

        public void WorkFlow()
        {
            ReadFiles();

            Check();

            GenFiles();

            _appendText.Invoke("All processes were completed !!!", Color.Black);
        }

        private void ReadFiles()
        {
            _appendText.Invoke("Starting to read PA files ...", Color.Black);
            foreach (string file in _files)
            {
                List<PaSheet> paSheets = GetPaSheet(file);
                bool hasChannelPA = false;
                foreach (PaSheet sheet in paSheets)
                {
                    if (sheet.Name.Equals("ChannelPA", StringComparison.CurrentCultureIgnoreCase)
                        && !_paSheets.ContainsKey(file))
                    {
                        _paSheets.Add(file, sheet);
                        hasChannelPA = true;
                    }
                }
                if (!hasChannelPA)
                {
                    string errorMsg = "The PA file \"" + Path.GetFileName(file) + "\" does not contain a subsheet named \"ChannelPA\".";
                    _appendText.Invoke(errorMsg, Color.Red);
                }
            }


            if (!string.IsNullOrEmpty(_DGSReferenceFile))
            {
                _appendText.Invoke("Starting to read DGS Reference files ...", Color.Black);
                List<PaSheet> paSheets = GetPaSheet(_DGSReferenceFile);
                foreach (PaSheet sheet in paSheets)
                {
                    if (sheet.Name.ToLowerInvariant().Contains("_dgspool"))
                    {
                        _DGSReferenceSheet = sheet;
                        break;
                    }
                }

                if (_DGSReferenceSheet == null)
                {
                    string errorMsg = "The DGS Reference file \"" + Path.GetFileName(_DGSReferenceFile) + "\" does not contain a subsheet which contains keywords \"_DgsPool\".";
                    _appendText.Invoke(errorMsg, Color.Red);
                }
            }


            //if (_device.Equals("PMIC", StringComparison.OrdinalIgnoreCase))
            if (sharePinEnable)
            {
                new GroupPins().GroupPinBySameChannel(_paSheets, out _paGroups, out _missingPaGroup);
            }
        }

        private List<PaSheet> GetPaSheet(string file)
        {
            List<PaSheet> paSheets = new List<PaSheet>();
            if (file != null && Path.GetExtension(file).Equals(".csv", StringComparison.OrdinalIgnoreCase))
            {
                paSheets.Add(new PaCsvReader(_uflexConfig).Read(file));
            }
            else if (file != null &&
                     (Path.GetExtension(file).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                      Path.GetExtension(file).Equals(".xlsm", StringComparison.OrdinalIgnoreCase)))
            {
                using (ExcelPackage inputExcel = new ExcelPackage(new FileInfo(file)))
                {
                    foreach (ExcelWorksheet sheet in inputExcel.Workbook.Worksheets)
                    {
                        PaSheet paSheet = new PaExcelReader(_uflexConfig).ReadSheet(sheet);
                        if (paSheet != null)
                        {
                            paSheets.Add(paSheet);
                        }
                    }
                }
            }

            return paSheets;
        }

        private void Check()
        {
            new PaChecker().CheckPaFile(_paSheets);
        }

        private void GenFiles()
        {
            _appendText.Invoke("Starting to generate Pin Map ...", Color.Black);
            GenPinMap genPinMap = new GenPinMap(_device, _uflexConfig, _appendText);
            PinMapSheet pinMap = genPinMap.GetPinMapSheet(_paSheets, _DGSReferenceSheet);
            string pinMapVersion = GetIgxlSheetVersion(pinMap.IgxlSheetName);
            pinMap.Write(Path.Combine(_outputPath, pinMap.Name), pinMapVersion);

            _appendText.Invoke("Starting to generate Channel Map ...", Color.Black);
            foreach (KeyValuePair<string, PaSheet> sheet in _paSheets)
            {
                GenChannelMap genChannelMap = new GenChannelMap(_device, _uflexConfig, _appendText);
                ChannelMapSheet channelMap =
                    genChannelMap.GetChannelMapSheet(sheet.Value.Rows, _hexVs, "_" + Path.GetFileName(sheet.Key), _DGSReferenceSheet);
                string channelMapVersion = GetIgxlSheetVersion(channelMap.IgxlSheetName);
                channelMap.Write(Path.Combine(_outputPath, channelMap.Name), channelMapVersion);
            }

            string outputFile = Path.Combine(_outputPath, "Error.xlsx");
            if (ErrorManager.GetErrorCount() > 0)
            {
                _appendText.Invoke("Starting to print error report ...", Color.Red);
                ErrorManager.GenErrorReport(outputFile, _files);
            }

            if (_paGroups.Count > 0 || _missingPaGroup.Count > 0)
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(outputFile)))
                {
                    ExcelWorkbook workbook = package.Workbook;

                    #region Pa group

                    if (_missingPaGroup.Any())
                    {
                        ExcelWorksheet missingSheet = workbook.AddSheet("MissingPaGroup");
                        missingSheet.Cells[1, 1].LoadFromCollection(_missingPaGroup, true);
                        missingSheet.MergeColumn(1);
                        missingSheet.Cells.AutoFitColumns();
                    }

                    if (_paGroups.Any())
                    {
                        ExcelWorksheet sheet = workbook.AddSheet("PaGroup");
                        sheet.Cells[1, 1].LoadFromCollection(_paGroups, true);
                        sheet.MergeColumn(1);
                        sheet.Cells.AutoFitColumns();
                    }

                    #endregion

                    package.Save();
                }
            }

            GenTimeSet genTimeSet = new GenTimeSet();
            genTimeSet.Write(_outputPath, _paSheets.SelectMany(x => x.Value.Rows)
                .Where(x => x.PaType.Equals("I/O", StringComparison.CurrentCulture) ||
                x.PaType.Equals("IO", StringComparison.CurrentCulture))
                .Select(y => y.BumpName).Distinct().ToList());

            GenPattern genPattern = new GenPattern();
            genPattern.Write(_outputPath, _paSheets);
        }

        private string GetIgxlSheetVersion(string sheetName)
        {
            Dictionary<string, List<SheetInfo>> igxlConfigDic = GetConfigDic();
            List<SheetInfo> currentVersionSheetInfos = igxlConfigDic[_igxlVersion];
            SheetInfo neededSheetInfo = currentVersionSheetInfos.Find(x => x.sheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            if (neededSheetInfo == null)
                return string.Empty;

            return neededSheetInfo.sheetVersion;
        }

        private Dictionary<string, List<SheetInfo>> GetConfigDic()
        {
            var assembly = Assembly.GetAssembly(typeof(IgxlSheet));
            var resourceNames = assembly.GetManifestResourceNames();
            var igxlConfigDic = new Dictionary<string, List<SheetInfo>>();
            foreach (var resourceName in resourceNames)
            {
                if (resourceName.Contains(".IGXLSheetsVersion."))
                {
                    var xs = new XmlSerializer(typeof(IGXLVersion));
                    var igxlConfig = (IGXLVersion)xs.Deserialize(assembly.GetManifestResourceStream(resourceName));
                    var igxlVersion = igxlConfig.igxlVersion;
                    if (!igxlConfigDic.ContainsKey(igxlVersion))
                    {
                        igxlConfigDic.Add(igxlVersion, igxlConfig.Sheets.ToList());
                    }
                }
            }
            return igxlConfigDic;
        }
    }
}