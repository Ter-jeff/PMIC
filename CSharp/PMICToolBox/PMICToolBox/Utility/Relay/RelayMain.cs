using PmicAutomation.Utility.Relay.Base;
using PmicAutomation.Utility.Relay.Function;
using PmicAutomation.Utility.Relay.Input;
using PmicAutomation.Utility.Relay.Output;
using Library.Function;
using Library.Function.ErrorReport;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.Relay
{
    public class RelayMain
    {
        private readonly MyForm.RichTextBoxAppend _appendText;
        private readonly string _comPin;
        private readonly string _relayConfig;
        private readonly string _outputPath;

        private const string FileName = "Trace_Based_Relay_Control.xlsx";

        private AdgMatrixSheet _adgMatrixSheet;
        private ComPinSheet _comPinSheet;
        private LinkedNodeRuleSheet _linkedNodeRuleSheet;
        private PinFilterSheet _pinFilterSheet;
        private bool _ADG1414Mode;

        public RelayMain(string comPin, string relayConfig, string outputPath,bool ADG1414Mode = false)
        {
            _comPin = comPin;
            _relayConfig = relayConfig;
            _outputPath = outputPath;
            _ADG1414Mode = ADG1414Mode;
        }

        public RelayMain(Relay relay)
        {
            _appendText = relay.AppendText;
            _comPin = relay.FileOpen_ComPin.ButtonTextBox.Text;
            _relayConfig = relay.FileOpen_RelayConfig.ButtonTextBox.Text;
            _outputPath = relay.FileOpen_OutputPath.ButtonTextBox.Text;
            _ADG1414Mode = relay.chkboxAdg1414.Checked;
            ErrorManager.ResetError();
        }

        public void WorkFlow()
        {
            ReadFiles();

            Check();

            if (!_ADG1414Mode)
            {
                Dictionary<string, List<string>> filterPins = _comPinSheet.GetFilterPins(_pinFilterSheet);
                List<string> resourcePins = _comPinSheet.GetResourcePins(filterPins);
                List<string> devicePins = _comPinSheet.GetDevicePins(filterPins);
                List<ComPinRow> comPinRows = _comPinSheet.FilterRows(_linkedNodeRuleSheet, filterPins);
                List<AdgMatrix> adgMatrixList = _comPinSheet.GetAdgMatrix(_comPinSheet.GetAdgMatrixSequence(_adgMatrixSheet), comPinRows);
                SearchRelay searchRelay = new SearchRelay(_appendText);
                searchRelay.SetNodes(comPinRows, _linkedNodeRuleSheet);
                _appendText.Invoke("Starting to analyze PinToPin ...", Color.Black);
                Dictionary<string, List<RelayPathRecord>> relayPathsDic = searchRelay.GenPinToPinFiles(filterPins);
                List<RelayItem> relayItems = searchRelay.GenRelayList(filterPins, adgMatrixList);
                GenFiles(resourcePins, devicePins, relayItems, adgMatrixList, comPinRows, relayPathsDic);
                _appendText.Invoke("All processes were completed !!!", Color.Black);
            }
            else
            {
                List<ComPinRow> comPinRows = _comPinSheet.FilterRows(null, null);
                Adg1414Reader _adg1414Reader = new Adg1414Reader();
                List<ADG1414Group> adg1414Groups = _adg1414Reader.ReadData(comPinRows);
                GenAdg1414Files(adg1414Groups);
                _appendText.Invoke("All processes were completed !!!", Color.Black);
            }
        }

        private void GenFiles(List<string> resourcePins, List<string> devicePins, List<RelayItem> relayItems,
            List<AdgMatrix> adgMatrices, List<ComPinRow> comPinRows,
            Dictionary<string, List<RelayPathRecord>> relayPathsDic)
        {
            GenRelayFile genRelayFile = new GenRelayFile(_appendText, _outputPath);
            _appendText.Invoke("Starting to generate PinToPin ...", Color.Black);
            genRelayFile.GenIniFile(relayPathsDic);
            _appendText.Invoke("Starting to generate Enum ...", Color.Black);
            genRelayFile.GenEnum(relayItems);
            _appendText.Invoke("Starting to generate EnumToString ...", Color.Black);
            genRelayFile.GenEnumToString(relayItems);
            _appendText.Invoke("Starting to generate Case ...", Color.Black);
            genRelayFile.GenCase(relayItems);
            _appendText.Invoke("Starting to generate SinExtract ...", Color.Black);
            genRelayFile.GenSinExtract(adgMatrices);
            _appendText.Invoke("Starting to generate Trace_Based_Relay_Control.xlsx ...", Color.Black);
            genRelayFile.GenTraceBasedRelayControl(FileName, resourcePins, devicePins, comPinRows, adgMatrices,
                relayItems);
            _appendText.Invoke("Starting to generate Relay.Bas ...", Color.Black);
            BasFileManage.GenBasFile(Path.Combine(_outputPath, "Relay.bas"), genRelayFile.GenBasFile(relayItems, resourcePins, devicePins, adgMatrices));
        }

        private void GenAdg1414Files(List<ADG1414Group> adg1414Rows)
        {
            GenADG1414File genAdg1414File = new Output.GenADG1414File(_appendText, _outputPath);
            _appendText.Invoke("Starting to generate ADG1414_Matrix ...", Color.Black);
            genAdg1414File.GenADG1414Matrix(adg1414Rows);
            _appendText.Invoke("Starting to generate ADG1414_CONTROL ...", Color.Black);
            genAdg1414File.GenADG1414CONTROL(adg1414Rows);
            _appendText.Invoke("Starting to generate SIN_Extract ...", Color.Black);
            genAdg1414File.GenSINExtract(adg1414Rows);
        }

        private void ReadFiles()
        {
            using (ExcelPackage inputExcel = new ExcelPackage(new FileInfo(_comPin)))
            {
                foreach (ExcelWorksheet sheet in inputExcel.Workbook.Worksheets)
                {
                    _appendText.Invoke("Starting to read ComPin sheet " + sheet.Name + " ...", Color.Black);
                    ComPinReader comPinReader = new ComPinReader();
                    _comPinSheet = comPinReader.ReadSheet(sheet);
                }
            }

            if (!_ADG1414Mode)
            {
                using (ExcelPackage inputExcel =
                        new ExcelPackage(new FileInfo(_relayConfig)))
                {
                    foreach (ExcelWorksheet sheet in inputExcel.Workbook.Worksheets)
                    {
                        _appendText.Invoke("Starting to read relay config sheet " + sheet.Name + " ...", Color.Black);
                        if (sheet.Name.Equals("PinFilter", StringComparison.CurrentCultureIgnoreCase))
                        {
                            PinFilterReader pinFilterReader = new PinFilterReader();
                            _pinFilterSheet = pinFilterReader.ReadSheet(sheet);
                        }

                        if (sheet.Name.Equals("LinkedNodeRule", StringComparison.CurrentCultureIgnoreCase))
                        {
                            LinkedNodeRuleReader linkedNodeRuleReader = new LinkedNodeRuleReader();
                            _linkedNodeRuleSheet = linkedNodeRuleReader.ReadSheet(sheet);
                        }

                        if (sheet.Name.Equals("ADGMatrix", StringComparison.CurrentCultureIgnoreCase))
                        {
                            AdgMatrixReader adgMatrixReader = new AdgMatrixReader();
                            _adgMatrixSheet = adgMatrixReader.ReadSheet(sheet);
                        }
                    }
                }
            }
        }

        private void Check()
        {
            if (_comPinSheet == null)
            {
                _appendText.Invoke("Please check if there is Component Pin Report !!!", Color.Red);
                throw new NullReferenceException();
            }

            if (!_ADG1414Mode)
            {
                if (_pinFilterSheet == null)
                {
                    _appendText.Invoke("Please check if there is PinFilter sheet in relay config !!!", Color.Red);
                    throw new NullReferenceException();
                }

                if (_linkedNodeRuleSheet == null)
                {
                    _appendText.Invoke("Please check if there is LinkedNodeRule sheet in relay config !!!", Color.Red);
                    throw new NullReferenceException();
                }

                if (_adgMatrixSheet == null)
                {
                    _appendText.Invoke("Please check if there is ADGMatrix sheet in relay config !!!", Color.Red);
                    throw new NullReferenceException();
                }
                string file = Path.Combine(_outputPath, FileName);
                if (Epplus.IsExcelOpened(file))
                {
                    _appendText.Invoke("Please close file " + file + " !!!", Color.Red);
                    throw new AccessViolationException();
                }
            }
        }
    }
}