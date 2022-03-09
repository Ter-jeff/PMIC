using System.Reflection;
using IgxlData.IgxlSheets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using AcSpecSheet = IgxlData.IgxlSheets.AcSpecSheet;
using BinTableSheet = IgxlData.IgxlSheets.BinTableSheet;
using ChannelMapSheet = IgxlData.IgxlSheets.ChannelMapSheet;
using DcSpecSheet = IgxlData.IgxlSheets.DcSpecSheet;
using GlobalSpecSheet = IgxlData.IgxlSheets.GlobalSpecSheet;
using LevelSheet = IgxlData.IgxlSheets.LevelSheet;
using PinMapSheet = IgxlData.IgxlSheets.PinMapSheet;

namespace IgxlData.IgxlWorkBooks
{
    public class IgxlWorkBook
    {
        #region Const Field
        public const string MainBinTblSheetName = "Bin_Table";
        public const string AcSpecsSheetName = "AC_Specs";
        public const string FlowTableMainInitEnableWd = "Flow_Table_Main_Init_EnableWd";
        #endregion

        private Dictionary<string, SubFlowSheet> _subFlowSheets;
        private Dictionary<string, SubFlowSheet> _mainFlowSheets;
        private Dictionary<string, InstanceSheet> _insSheets;
        private Dictionary<string, DcSpecSheet> _dcSpecSheets;
        private Dictionary<string, AcSpecSheet> _acSpecSheets;
        private Dictionary<string, LevelSheet> _levelSheets;
        private Dictionary<string, TimeSetBasicSheet> _timeSetSheets;
        private KeyValuePair<string, PinMapSheet> _pinMapPair;
        private KeyValuePair<string, GlobalSpecSheet> _glbSpecSheetPair;
        private Dictionary<string, PatSetSheet> _patSetSheets;
        private Dictionary<string, PatSetSubSheet> _patSetSubSheets;
        private Dictionary<string, BinTableSheet> _binTblSheets;
        private Dictionary<string, JoblistSheet> _joblistSheets;
        private Dictionary<string, ChannelMapSheet> _channelMapSheets;
        private Dictionary<string, PortMapSheet> _portMapSheets;
        private Dictionary<string, CharSheet> _charSheets;
        private Dictionary<string, WaveDefinitionSheet> _wavedefSheets;
        private Dictionary<string, MixedSignalSheet> _mixsigSheets;
        private Dictionary<string, PSetSheet> _psetSheets;
        private Dictionary<string, IgxlSheet> _allIgxlSheets;
        private IGXL _igxlConfig;
        private List<string> _flowUsedInteger;

        #region Constructor
        public IgxlWorkBook()
        {
            _pinMapPair = new KeyValuePair<string, PinMapSheet>();
            _allIgxlSheets = new Dictionary<string, IgxlSheet>(StringComparer.CurrentCultureIgnoreCase);
            _subFlowSheets = new Dictionary<string, SubFlowSheet>(StringComparer.CurrentCultureIgnoreCase);
            _mainFlowSheets = new Dictionary<string, SubFlowSheet>(StringComparer.CurrentCultureIgnoreCase);
            _insSheets = new Dictionary<string, InstanceSheet>(StringComparer.CurrentCultureIgnoreCase);
            _dcSpecSheets = new Dictionary<string, DcSpecSheet>(StringComparer.CurrentCultureIgnoreCase);
            _acSpecSheets = new Dictionary<string, AcSpecSheet>(StringComparer.CurrentCultureIgnoreCase);
            _levelSheets = new Dictionary<string, LevelSheet>(StringComparer.CurrentCultureIgnoreCase);
            _timeSetSheets = new Dictionary<string, TimeSetBasicSheet>(StringComparer.CurrentCultureIgnoreCase);
            _patSetSheets = new Dictionary<string, PatSetSheet>(StringComparer.CurrentCultureIgnoreCase);
            _patSetSubSheets = new Dictionary<string, PatSetSubSheet>(StringComparer.CurrentCultureIgnoreCase);
            _binTblSheets = new Dictionary<string, BinTableSheet>(StringComparer.CurrentCultureIgnoreCase);
            _joblistSheets = new Dictionary<string, JoblistSheet>(StringComparer.CurrentCultureIgnoreCase);
            _channelMapSheets = new Dictionary<string, ChannelMapSheet>(StringComparer.CurrentCultureIgnoreCase);
            _portMapSheets = new Dictionary<string, PortMapSheet>(StringComparer.CurrentCultureIgnoreCase);
            _wavedefSheets = new Dictionary<string, WaveDefinitionSheet>(StringComparer.CurrentCultureIgnoreCase);
            _mixsigSheets = new Dictionary<string, MixedSignalSheet>(StringComparer.CurrentCultureIgnoreCase);
            _psetSheets = new Dictionary<string, PSetSheet>(StringComparer.CurrentCultureIgnoreCase);
            _charSheets = new Dictionary<string, CharSheet>(StringComparer.CurrentCultureIgnoreCase);
            _flowUsedInteger = new List<string>();
        }
        #endregion

        #region Property
        public Dictionary<string, SubFlowSheet> SubFlowSheets
        {
            get { return _subFlowSheets; }
        }

        public Dictionary<string, SubFlowSheet> MainFlowSheets
        {
            get { return _mainFlowSheets; }
        }

        public Dictionary<string, InstanceSheet> InsSheets
        {
            get { return _insSheets; }
        }

        public Dictionary<string, DcSpecSheet> DcSpecSheets
        {
            get { return _dcSpecSheets; }
        }

        public Dictionary<string, AcSpecSheet> AcSpecSheets
        {
            get { return _acSpecSheets; }
        }

        public Dictionary<string, LevelSheet> LevelSheets
        {
            get { return _levelSheets; }
        }

        public Dictionary<string, TimeSetBasicSheet> TimeSetSheets
        {
            get { return _timeSetSheets; }
        }

        public Dictionary<string, PatSetSheet> PatSetSheets
        {
            get { return _patSetSheets; }
        }

        public Dictionary<string, PatSetSubSheet> PatSetSubSheets
        {
            get { return _patSetSubSheets; }
        }

        public Dictionary<string, CharSheet> CharSheets
        {
            get { return _charSheets; }
        }

        public Dictionary<string, WaveDefinitionSheet> WaveDefSheets
        {
            get { return _wavedefSheets; }
        }

        public Dictionary<string, MixedSignalSheet> MixedSignalSheets
        {
            get { return _mixsigSheets; }
        }

        public List<string> FlowUsedInteger
        {
            set { _flowUsedInteger = value; }
            get { return _flowUsedInteger; }
        }

        public Dictionary<string, BinTableSheet> BinTblSheets
        {
            get { return _binTblSheets; }
        }

        public Dictionary<string, ChannelMapSheet> ChannelMapSheets
        {
            get { return _channelMapSheets; }
        }

        public Dictionary<string, PortMapSheet> PortMapSheets
        {
            get { return _portMapSheets; }
        }

        public KeyValuePair<string, GlobalSpecSheet> GlbSpecSheetPair
        {
            get { return _glbSpecSheetPair; }
            set
            {
                if (_glbSpecSheetPair.Value == null)
                {
                    var subFileSubName = Contact(value.Key, value.Value.Name);
                    _glbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(subFileSubName, value.Value);
                    _allIgxlSheets.Add(subFileSubName, _glbSpecSheetPair.Value);
                }
                else
                {
                    throw new Exception("Can't duplicate set igxl global Spec Sheet.");
                }
            }
        }

        public KeyValuePair<string, PinMapSheet> PinMapPair
        {
            get
            {
                return _pinMapPair;
            }
            set
            {
                if (_pinMapPair.Value == null)
                {
                    var subFileSubName = Contact(value.Key, value.Value.Name);
                    _pinMapPair = new KeyValuePair<string, PinMapSheet>(subFileSubName, value.Value);
                    _allIgxlSheets.Add(subFileSubName, _pinMapPair.Value);
                }
                else
                {
                    throw new Exception("Can't duplicate set igxl pin map.");
                }
            }
        }

        public Dictionary<string, JoblistSheet> JoblistSheets
        {
            get { return _joblistSheets; }
        }

        public Dictionary<string, IgxlSheet> AllIgxlSheets
        {
            get { return _allIgxlSheets; }
        }
        #endregion

        #region Member Function
        public InstanceSheet GetInstanceSheet(string name)
        {
            foreach (var insSheet in _insSheets)
            {
                if (insSheet.Value.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    return insSheet.Value;
            }
            return null;
        }

        public AcSpecSheet GetAcSpecsSheet(string name = AcSpecsSheetName)
        {
            foreach (var acSpecSheet in _acSpecSheets)
            {
                if (acSpecSheet.Value.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    return acSpecSheet.Value;
            }
            return null;
        }

        public SubFlowSheet GetFlowTableMainInitEnableWdSheet(string name = FlowTableMainInitEnableWd)
        {
            foreach (var mainFlow in _subFlowSheets)
            {
                if (mainFlow.Value.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    return mainFlow.Value;
            }
            return null;
        }

        public BinTableSheet GetBinTblSheet(string bintableName)
        {
            foreach (var binTblPair in _binTblSheets)
            {
                if (binTblPair.Value.Name.Equals(bintableName, StringComparison.CurrentCultureIgnoreCase))
                    return binTblPair.Value;
            }
            return null;
        }


        public BinTableSheet GetMainBinTblSheet(string bintableName = MainBinTblSheetName)
        {
            BinTableSheet binTableSheet = null;
            foreach (var binTblPair in _binTblSheets)
            {
                if (binTblPair.Value.Name.Equals(bintableName, StringComparison.CurrentCultureIgnoreCase))
                {
                    binTableSheet = binTblPair.Value;
                    break;
                }
            }
            if (binTableSheet == null)
            {
                binTableSheet = new BinTableSheet(MainBinTblSheetName);
                //throw new Exception(string.Format("Can't find mian bin table sheet, which sheet name is: {0} in IgxlWorkBook. ", bintableName));
            }

            return binTableSheet;
        }


        public IgxlWorkBook MergrIgxlWorkBook(IgxlWorkBook igxlWorkWorkBooks)
        {
            foreach (var sheet in igxlWorkWorkBooks.SubFlowSheets)
                if (!_subFlowSheets.ContainsKey(sheet.Key)) _subFlowSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.MainFlowSheets)
                if (!_mainFlowSheets.ContainsKey(sheet.Key)) _mainFlowSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.InsSheets)
                if (!_subFlowSheets.ContainsKey(sheet.Key)) _insSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.DcSpecSheets)
                if (!_dcSpecSheets.ContainsKey(sheet.Key)) _dcSpecSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.AcSpecSheets)
                if (!_acSpecSheets.ContainsKey(sheet.Key)) _acSpecSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.LevelSheets)
                if (!_levelSheets.ContainsKey(sheet.Key)) _levelSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.TimeSetSheets)
                if (!_timeSetSheets.ContainsKey(sheet.Key)) _timeSetSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.PatSetSheets)
                if (!_patSetSheets.ContainsKey(sheet.Key)) _patSetSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.PatSetSubSheets)
                if (!_patSetSubSheets.ContainsKey(sheet.Key)) _patSetSubSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.BinTblSheets)
                if (!_binTblSheets.ContainsKey(sheet.Key)) _binTblSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.JoblistSheets)
                if (!_joblistSheets.ContainsKey(sheet.Key)) _joblistSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.ChannelMapSheets)
                if (!_channelMapSheets.ContainsKey(sheet.Key)) _channelMapSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.PortMapSheets)
                if (!_portMapSheets.ContainsKey(sheet.Key)) _portMapSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.CharSheets)
                if (!_charSheets.ContainsKey(sheet.Key)) _charSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.WaveDefSheets)
                if (!_wavedefSheets.ContainsKey(sheet.Key)) _wavedefSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.MixedSignalSheets)
                if (!_mixsigSheets.ContainsKey(sheet.Key)) _mixsigSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.AllIgxlSheets)
                if (!_allIgxlSheets.ContainsKey(sheet.Key)) _allIgxlSheets.Add(sheet.Key, sheet.Value);

            if (igxlWorkWorkBooks.PinMapPair.Value != null)
                _pinMapPair.Value.PinList.AddRange(igxlWorkWorkBooks.PinMapPair.Value.PinList);
            if (igxlWorkWorkBooks.PinMapPair.Value != null)
                _pinMapPair.Value.GroupList.AddRange(igxlWorkWorkBooks.PinMapPair.Value.GroupList);
            if (igxlWorkWorkBooks.GlbSpecSheetPair.Value != null)
                _glbSpecSheetPair.Value.AddRange(igxlWorkWorkBooks.GlbSpecSheetPair.Value.Specs);

            return this;
        }

        public void RemoveSheet(string key)
        {
            _allIgxlSheets.Remove(key);
        }

        public void AddSubFlowSheet(KeyValuePair<string, SubFlowSheet> sheet)
        {
            if (!_subFlowSheets.ContainsKey(sheet.Key))
            {
                _subFlowSheets.Add(sheet.Key, sheet.Value);
                _allIgxlSheets.Add(sheet.Key, sheet.Value);
            }
            else
            {
                _subFlowSheets[sheet.Key] = sheet.Value;
                _allIgxlSheets[sheet.Key] = sheet.Value;
            }
        }

        public void AddSubFlowSheet(string fileSubPath, SubFlowSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_subFlowSheets.ContainsKey(subFileSubName))
            {
                _subFlowSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _subFlowSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddMainFlowSheet(string fileSubPath, SubFlowSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_mainFlowSheets.ContainsKey(subFileSubName))
            {
                _mainFlowSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _mainFlowSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }

        }

        public void AddInsSheet(KeyValuePair<string, InstanceSheet> sheet)
        {
            if (!_subFlowSheets.ContainsKey(sheet.Key))
            {
                _insSheets.Add(sheet.Key, sheet.Value);
                _allIgxlSheets.Add(sheet.Key, sheet.Value);
            }
            else
            {
                _insSheets[sheet.Key] = sheet.Value;
                _allIgxlSheets[sheet.Key] = sheet.Value;
            }
        }

        public void AddInsSheet(string fileSubPath, InstanceSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_insSheets.ContainsKey(subFileSubName))
            {
                _insSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _insSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddDcSpecSheet(string fileSubPath, DcSpecSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_dcSpecSheets.ContainsKey(subFileSubName))
            {
                _dcSpecSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _dcSpecSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddAcSpecSheet(string fileSubPath, AcSpecSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_acSpecSheets.ContainsKey(subFileSubName))
            {
                _acSpecSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _acSpecSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddLevelSheet(string fileSubPath, LevelSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_levelSheets.ContainsKey(subFileSubName))
            {
                _levelSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _levelSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddTimeSetSheet(string fileSubPath, TimeSetBasicSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_timeSetSheets.ContainsKey(subFileSubName))
            {
                _timeSetSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _timeSetSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddPatSetSheet(string fileSubPath, PatSetSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_patSetSheets.ContainsKey(subFileSubName))
            {
                _patSetSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _patSetSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddPatSetSubSheet(string fileSubPath, PatSetSubSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_patSetSubSheets.ContainsKey(subFileSubName))
            {
                _patSetSubSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _patSetSubSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddBinTblSheet(string fileSubPath, BinTableSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_binTblSheets.ContainsKey(subFileSubName))
            {
                _binTblSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _binTblSheets[subFileSubName].BinTableRows.AddRange(sheet.BinTableRows);
            }
        }

        public void AddJobListSheet(string fileSubPath, JoblistSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_joblistSheets.ContainsKey(subFileSubName))
            {
                _joblistSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _joblistSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddChannelMapSheet(string fileSubPath, ChannelMapSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_channelMapSheets.ContainsKey(subFileSubName))
            {
                _channelMapSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _channelMapSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddPortMapSheet(string fileSubPath, PortMapSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_portMapSheets.ContainsKey(subFileSubName))
            {
                _portMapSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _portMapSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddWaveDefSheet(string fileSubPath, WaveDefinitionSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_wavedefSheets.ContainsKey(subFileSubName))
            {
                _wavedefSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _wavedefSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddMixedSignalSheet(string fileSubPath, MixedSignalSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_mixsigSheets.ContainsKey(subFileSubName))
            {
                _mixsigSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _mixsigSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddPsetSheet(string fileSubPath, PSetSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_mixsigSheets.ContainsKey(subFileSubName))
            {
                _psetSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _psetSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public void AddCharSheet(string fileSubPath, CharSheet sheet)
        {
            var subFileSubName = GetSubFileSubName(fileSubPath, sheet.Name);
            if (!_charSheets.ContainsKey(subFileSubName))
            {
                _charSheets.Add(subFileSubName, sheet);
                _allIgxlSheets.Add(subFileSubName, sheet);
            }
            else
            {
                _charSheets[subFileSubName] = sheet;
                _allIgxlSheets[subFileSubName] = sheet;
            }
        }

        public List<string> GetFlowSheetNameList()
        {
            var flowSheetNameList = new List<string>();
            foreach (var flowsheet in _subFlowSheets)
                flowSheetNameList.Add(flowsheet.Value.Name);
            return flowSheetNameList;
        }

        private string GetSubFileSubName(string fileSubPath, string sheetName)
        {
            return Contact(fileSubPath, sheetName);
        }

        public void PrintAllSheets(string igxlVersion)
        {
            foreach (var igxlSheetPair in _allIgxlSheets)
            {
                var sheetVersion = GetSheetVersion(igxlSheetPair.Value, igxlVersion);
                igxlSheetPair.Value.Write(igxlSheetPair.Key + ".txt", sheetVersion);
            }
        }

        public void Clear()
        {
            _allIgxlSheets.Clear();
            _subFlowSheets.Clear();
            _mainFlowSheets.Clear();
            _insSheets.Clear();
            _dcSpecSheets.Clear();
            _acSpecSheets.Clear();
            _levelSheets.Clear();
            _timeSetSheets.Clear();
            _patSetSheets.Clear();
            _patSetSubSheets.Clear();
            _binTblSheets.Clear();
            _joblistSheets.Clear();
            _channelMapSheets.Clear();
            _charSheets.Clear();
            _wavedefSheets.Clear();
            _mixsigSheets.Clear();
            _glbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(null, null);
            _pinMapPair = new KeyValuePair<string, PinMapSheet>(null, null);
            _portMapSheets.Clear();
        }

        public string GetSheetVersion(IgxlSheet sheet, string igxlVersion)
        {
            if (_igxlConfig == null)
            {
                var assembly = Assembly.GetExecutingAssembly();
                var resourceNames = assembly.GetManifestResourceNames();
                foreach (var resourceName in resourceNames)
                {
                    if (resourceName.Contains(".SheetClassMapping.v" +igxlVersion + "_ultraflex.xml"))
                    {
                        _igxlConfig = LoadConfig(assembly.GetManifestResourceStream(resourceName));
                        //var configFileName = Directory.GetCurrentDirectory() + "\\IGDataXML\\SheetClassMapping\\" + "v" +igxlVersion + "_ultraflex.xml";
                        //_igxlConfig = LoadConfig(configFileName);
                    }
                }
            }
            if (_igxlConfig != null)
                return _igxlConfig.SheetItemClass.First(p => p.sheetname.Equals(sheet.IgxlSheetName)).sheetversion;

            return "";
        }

        private IGXL LoadConfig(Stream sr)
        {
            IGXL result;
            try
            {
                var xs = new XmlSerializer(typeof(IGXL));
                IGXL sysData = (IGXL)xs.Deserialize(sr);
                sr.Close();
                result = sysData;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return result;
        }

        private IGXL LoadConfig(string fileName)
        {
            IGXL result;
            try
            {
                var xs = new XmlSerializer(typeof(IGXL));
                var sr = new StreamReader(fileName);
                var sysData = (IGXL)xs.Deserialize(sr);
                sr.Close();
                result = sysData;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }
            return result;
        }

        public string Contact(string mainDir, string subDir)
        {
            if (mainDir != null && subDir != null)
            {
                mainDir = mainDir.Replace('/', '\\');
                subDir = subDir.Replace('/', '\\');
                mainDir = mainDir.TrimEnd('\\');
                subDir = subDir.TrimStart('\\');
                return mainDir + "\\" + subDir;
            }
            throw new Exception(string.Format("Directory object: {0} or {1} is null. ", mainDir, subDir));
        }

        #endregion
    }
}