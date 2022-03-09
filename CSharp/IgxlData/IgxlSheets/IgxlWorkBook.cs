using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Serialization;
using Teradyne.Oasis.IGData.Utilities;

namespace IgxlData.IgxlSheets
{
    public class IgxlWorkBook
    {
        #region Field

        #region Const Field
        public const string MainBinTableName = "Bin_Table";
        #endregion

        private KeyValuePair<string, PinMapSheet> _pinMapPair;
        private KeyValuePair<string, GlobalSpecSheet> _glbSpecSheetPair;
        private IGXL _igxlConfig;

        #endregion

        #region Constructor
        public IgxlWorkBook()
        {
            _pinMapPair = new KeyValuePair<string, PinMapSheet>();
            AllIgxlSheets = new Dictionary<string, IgxlSheet>(StringComparer.CurrentCultureIgnoreCase);
            SubFlowSheets = new Dictionary<string, SubFlowSheet>(StringComparer.CurrentCultureIgnoreCase);
            MainFlowSheets = new Dictionary<string, FlowSheet>(StringComparer.CurrentCultureIgnoreCase);
            InsSheets = new Dictionary<string, InstanceSheet>(StringComparer.CurrentCultureIgnoreCase);
            DcSpecSheets = new Dictionary<string, DcSpecSheet>(StringComparer.CurrentCultureIgnoreCase);
            AcSpecSheets = new Dictionary<string, AcSpecSheet>(StringComparer.CurrentCultureIgnoreCase);
            LevelSheets = new Dictionary<string, LevelSheet>(StringComparer.CurrentCultureIgnoreCase);
            TimeSetSheets = new Dictionary<string, TimeSetBasicSheet>(StringComparer.CurrentCultureIgnoreCase);
            PatSetSheets = new Dictionary<string, PatSetSheet>(StringComparer.CurrentCultureIgnoreCase);
            PatSetSubSheets = new Dictionary<string, PatSetSubSheet>(StringComparer.CurrentCultureIgnoreCase);
            BinTblSheets = new Dictionary<string, BinTableSheet>(StringComparer.CurrentCultureIgnoreCase);
            JobListSheets = new Dictionary<string, JobListSheet>(StringComparer.CurrentCultureIgnoreCase);
            ChannelMapSheets = new Dictionary<string, ChannelMapSheet>(StringComparer.CurrentCultureIgnoreCase);
            PortMapSheets = new Dictionary<string, PortMapSheet>(StringComparer.CurrentCultureIgnoreCase);
            WaveDefSheets = new Dictionary<string, WaveDefinitionSheet>(StringComparer.CurrentCultureIgnoreCase);
            MixedSignalSheets = new Dictionary<string, MixedSignalSheet>(StringComparer.CurrentCultureIgnoreCase);
            CharSheets = new Dictionary<string, CharSheet>(StringComparer.CurrentCultureIgnoreCase);
            FlowUsedInteger = new List<string>();
        }
        #endregion

        #region Property
        public Dictionary<string, SubFlowSheet> SubFlowSheets { get; }

        public Dictionary<string, FlowSheet> MainFlowSheets { get; }

        public Dictionary<string, InstanceSheet> InsSheets { get; }

        public Dictionary<string, DcSpecSheet> DcSpecSheets { get; }

        public Dictionary<string, AcSpecSheet> AcSpecSheets { get; }

        public Dictionary<string, LevelSheet> LevelSheets { get; }

        public Dictionary<string, TimeSetBasicSheet> TimeSetSheets { get; }

        public Dictionary<string, PatSetSheet> PatSetSheets { get; }

        public Dictionary<string, PatSetSubSheet> PatSetSubSheets { get; }

        public Dictionary<string, CharSheet> CharSheets { get; }

        public Dictionary<string, WaveDefinitionSheet> WaveDefSheets { get; }

        public Dictionary<string, MixedSignalSheet> MixedSignalSheets { get; }

        public List<string> FlowUsedInteger { set; get; }

        public BinTableSheet GetMainBinTblSheet()
        {
            BinTableSheet mainBinTblSheet = null;
            foreach (KeyValuePair<string, BinTableSheet> binTblPair in BinTblSheets)
            {
                if (binTblPair.Value.SheetName.Equals(MainBinTableName))
                {
                    mainBinTblSheet = binTblPair.Value;
                    break;
                }
            }
            if (mainBinTblSheet == null)
            {
                throw new Exception(string.Format("Can't find main bin table sheet, which sheet name is: {0} in IgxlWorkBook. ", MainBinTableName));
            }
            return mainBinTblSheet;
        }

        public BinTableSheet GetMainBinTblSheet(string binTableName)
        {
            BinTableSheet mainBinTblSheet = null;
            foreach (KeyValuePair<string, BinTableSheet> binTblPair in BinTblSheets)
            {
                if (binTblPair.Value.SheetName.Equals(binTableName))
                {
                    mainBinTblSheet = binTblPair.Value;
                    break;
                }
            }
            if (mainBinTblSheet == null)
            {
                return null;
            }
            return mainBinTblSheet;
        }

        public Dictionary<string, BinTableSheet> BinTblSheets { get; }

        public Dictionary<string, ChannelMapSheet> ChannelMapSheets { get; }

        public Dictionary<string, PortMapSheet> PortMapSheets { get; }

        public KeyValuePair<string, GlobalSpecSheet> GlbSpecSheetPair
        {
            get { return _glbSpecSheetPair; }
            set
            {
                if (_glbSpecSheetPair.Value == null)
                {
                    string subFileSubName = Contact(value.Key, value.Value.SheetName);
                    _glbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(subFileSubName, value.Value);
                    AllIgxlSheets.Add(subFileSubName, _glbSpecSheetPair.Value);
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
                    string subFileSubName = Contact(value.Key, value.Value.SheetName);
                    _pinMapPair = new KeyValuePair<string, PinMapSheet>(subFileSubName, value.Value);
                    AllIgxlSheets.Add(subFileSubName, _pinMapPair.Value);
                }
                else
                {
                    throw new Exception("Can't duplicate set igxl pin map.");
                }
            }
        }

        public Dictionary<string, JobListSheet> JobListSheets { get; }

        public Dictionary<string, IgxlSheet> AllIgxlSheets { get; }

        #endregion

        #region Member Function

        public CharSheet GetCharSheet(string sheetName)
        {
            foreach (var charSheet in CharSheets)
            {
                if (charSheet.Value.SheetName.Equals(sheetName, StringComparison.CurrentCultureIgnoreCase))
                    return charSheet.Value;
            }
            return new CharSheet(sheetName);
        }

        public IgxlWorkBook MergeIgxlWorkBook(IgxlWorkBook igxlWorkWorkBooks)
        {
            foreach (var sheet in igxlWorkWorkBooks.SubFlowSheets)
                if (!SubFlowSheets.ContainsKey(sheet.Key)) SubFlowSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.MainFlowSheets)
                if (!MainFlowSheets.ContainsKey(sheet.Key)) MainFlowSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.InsSheets)
                if (!SubFlowSheets.ContainsKey(sheet.Key)) InsSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.DcSpecSheets)
                if (!DcSpecSheets.ContainsKey(sheet.Key)) DcSpecSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.AcSpecSheets)
                if (!AcSpecSheets.ContainsKey(sheet.Key)) AcSpecSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.LevelSheets)
                if (!LevelSheets.ContainsKey(sheet.Key)) LevelSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.TimeSetSheets)
                if (!TimeSetSheets.ContainsKey(sheet.Key)) TimeSetSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.PatSetSheets)
                if (!PatSetSheets.ContainsKey(sheet.Key)) PatSetSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.PatSetSubSheets)
                if (!PatSetSubSheets.ContainsKey(sheet.Key)) PatSetSubSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.BinTblSheets)
                if (!BinTblSheets.ContainsKey(sheet.Key)) BinTblSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.JobListSheets)
                if (!JobListSheets.ContainsKey(sheet.Key)) JobListSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.ChannelMapSheets)
                if (!ChannelMapSheets.ContainsKey(sheet.Key)) ChannelMapSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.PortMapSheets)
                if (!PortMapSheets.ContainsKey(sheet.Key)) PortMapSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.CharSheets)
                if (!CharSheets.ContainsKey(sheet.Key)) CharSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.WaveDefSheets)
                if (!WaveDefSheets.ContainsKey(sheet.Key)) WaveDefSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.MixedSignalSheets)
                if (!MixedSignalSheets.ContainsKey(sheet.Key)) MixedSignalSheets.Add(sheet.Key, sheet.Value);
            foreach (var sheet in igxlWorkWorkBooks.AllIgxlSheets)
                if (!AllIgxlSheets.ContainsKey(sheet.Key)) AllIgxlSheets.Add(sheet.Key, sheet.Value);

            if (igxlWorkWorkBooks.PinMapPair.Value != null)
                _pinMapPair.Value.PinList.AddRange(igxlWorkWorkBooks.PinMapPair.Value.PinList);
            if (igxlWorkWorkBooks.PinMapPair.Value != null)
                _pinMapPair.Value.GroupList.AddRange(igxlWorkWorkBooks.PinMapPair.Value.GroupList);
            if (igxlWorkWorkBooks.GlbSpecSheetPair.Value != null)
                _glbSpecSheetPair.Value.AddRange(igxlWorkWorkBooks.GlbSpecSheetPair.Value.GetGlobalSpecs());

            return this;
        }

        public void AddSubFlowSheet(KeyValuePair<string, SubFlowSheet> flowSheet)
        {
            if (!SubFlowSheets.ContainsKey(flowSheet.Key))
            {
                SubFlowSheets.Add(flowSheet.Key, flowSheet.Value);
                AllIgxlSheets.Add(flowSheet.Key, flowSheet.Value);
            }
        }

        public void AddSubFlowSheet(string fileSubPath, SubFlowSheet flowSheet)
        {
            if (flowSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, flowSheet.SheetName);
            if (!SubFlowSheets.ContainsKey(subFileSubName))
            {
                SubFlowSheets.Add(subFileSubName, flowSheet);
                AllIgxlSheets.Add(subFileSubName, flowSheet);
            }
        }

        public void AddMainFlowSheet(string fileSubPath, FlowSheet flowSheet)
        {
            if (flowSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, flowSheet.SheetName);
            if (!MainFlowSheets.ContainsKey(subFileSubName))
            {
                MainFlowSheets.Add(subFileSubName, flowSheet);
                AllIgxlSheets.Add(subFileSubName, flowSheet);
            }
        }

        public void AddInsSheet(string fileSubPath, InstanceSheet insSheet)
        {
            if (insSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, insSheet.SheetName);
            if (!InsSheets.ContainsKey(subFileSubName))
            {
                InsSheets.Add(subFileSubName, insSheet);
                AllIgxlSheets.Add(subFileSubName, insSheet);
            }
        }

        public void AddDcSpecSheet(string fileSubPath, DcSpecSheet dcSpecSheet)
        {
            if (dcSpecSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, dcSpecSheet.SheetName);
            if (!DcSpecSheets.ContainsKey(subFileSubName))
            {
                DcSpecSheets.Add(subFileSubName, dcSpecSheet);
                AllIgxlSheets.Add(subFileSubName, dcSpecSheet);
            }
        }

        public void AddAcSpecSheet(string fileSubPath, AcSpecSheet acSpecSheet)
        {
            if (acSpecSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, acSpecSheet.SheetName);
            if (!AcSpecSheets.ContainsKey(subFileSubName))
            {
                AcSpecSheets.Add(subFileSubName, acSpecSheet);
                AllIgxlSheets.Add(subFileSubName, acSpecSheet);
            }
        }

        public void AddLevelSheet(string fileSubPath, LevelSheet levelSheet)
        {
            if (levelSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, levelSheet.SheetName);
            if (!LevelSheets.ContainsKey(subFileSubName))
            {
                LevelSheets.Add(subFileSubName, levelSheet);
                AllIgxlSheets.Add(subFileSubName, levelSheet);
            }
        }

        public void AddTimeSetSheet(string fileSubPath, TimeSetBasicSheet timeSetSheet)
        {
            if (timeSetSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, timeSetSheet.SheetName);
            if (!TimeSetSheets.ContainsKey(subFileSubName))
            {
                TimeSetSheets.Add(subFileSubName, timeSetSheet);
                AllIgxlSheets.Add(subFileSubName, timeSetSheet);
            }
        }

        public void AddPatSetSheet(string fileSubPath, PatSetSheet patSetSheet)
        {
            if (patSetSheet == null) return;
            if (!patSetSheet.PatSetRows.Any()) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, patSetSheet.SheetName);
            if (!PatSetSheets.ContainsKey(subFileSubName))
            {
                PatSetSheets.Add(subFileSubName, patSetSheet);
                AllIgxlSheets.Add(subFileSubName, patSetSheet);
            }
        }

        public void AddPatSetSubSheet(string fileSubPath, PatSetSubSheet patSetSubSheet)
        {
            if (patSetSubSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, patSetSubSheet.SheetName);
            if (!PatSetSubSheets.ContainsKey(subFileSubName))
            {
                PatSetSubSheets.Add(subFileSubName, patSetSubSheet);
                AllIgxlSheets.Add(subFileSubName, patSetSubSheet);
            }
        }

        public void AddBinTblSheet(string fileSubPath, BinTableSheet binTableSheet)
        {
            if (binTableSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, binTableSheet.SheetName);
            if (!BinTblSheets.ContainsKey(subFileSubName))
            {
                BinTblSheets.Add(subFileSubName, binTableSheet);
                AllIgxlSheets.Add(subFileSubName, binTableSheet);
            }
            else
            {
                BinTblSheets[subFileSubName].BinTableRows.AddRange(binTableSheet.BinTableRows);
            }
        }

        public void AddJobListSheet(string fileSubPath, JobListSheet jobListSheet)
        {
            if (jobListSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, jobListSheet.SheetName);
            if (!JobListSheets.ContainsKey(subFileSubName))
            {
                JobListSheets.Add(subFileSubName, jobListSheet);
                AllIgxlSheets.Add(subFileSubName, jobListSheet);
            }
        }

        public void AddChannelMapSheet(string fileSubPath, ChannelMapSheet channelMapSheet)
        {
            if (channelMapSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, channelMapSheet.SheetName);
            if (!ChannelMapSheets.ContainsKey(subFileSubName))
            {
                ChannelMapSheets.Add(subFileSubName, channelMapSheet);
                AllIgxlSheets.Add(subFileSubName, channelMapSheet);
            }
        }

        public void AddPortMapSheet(string fileSubPath, PortMapSheet portMapSheet)
        {
            if (portMapSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, portMapSheet.SheetName);
            if (!PortMapSheets.ContainsKey(subFileSubName))
            {
                PortMapSheets.Add(subFileSubName, portMapSheet);
                AllIgxlSheets.Add(subFileSubName, portMapSheet);
            }
        }

        public void AddWaveDefSheet(string fileSubPath, WaveDefinitionSheet waveDefinitionSheet)
        {
            if (waveDefinitionSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, waveDefinitionSheet.SheetName);
            if (!WaveDefSheets.ContainsKey(subFileSubName))
            {
                WaveDefSheets.Add(subFileSubName, waveDefinitionSheet);
                AllIgxlSheets.Add(subFileSubName, waveDefinitionSheet);
            }
        }

        public void AddMixedSignalSheet(string fileSubPath, MixedSignalSheet mixedSignalSheet)
        {
            if (mixedSignalSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, mixedSignalSheet.SheetName);
            if (!MixedSignalSheets.ContainsKey(subFileSubName))
            {
                MixedSignalSheets.Add(subFileSubName, mixedSignalSheet);
                AllIgxlSheets.Add(subFileSubName, mixedSignalSheet);
            }
        }

        public void AddPSetSheet(string fileSubPath, PSetSheet pSetSheet)
        {
            if (pSetSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, pSetSheet.SheetName);
            if (!MixedSignalSheets.ContainsKey(subFileSubName))
            {
                AllIgxlSheets.Add(subFileSubName, pSetSheet);
            }
        }

        public void AddCharSheet(string fileSubPath, CharSheet charSheet)
        {
            if (charSheet == null) return;
            string subFileSubName = GetSubFileSubName(fileSubPath, charSheet.SheetName);
            if (!CharSheets.ContainsKey(subFileSubName))
            {
                CharSheets.Add(subFileSubName, charSheet);
                AllIgxlSheets.Add(subFileSubName, charSheet);
            }
        }

        public List<string> GetFlowSheetNameList()
        {
            List<string> flowSheetNameList = new List<string>();
            foreach (var flowSheet in SubFlowSheets)
            {
                flowSheetNameList.Add(flowSheet.Value.SheetName);
            }
            return flowSheetNameList;
        }

        private string GetSubFileSubName(string fileSubPath, string sheetName)
        {
            return Contact(fileSubPath, sheetName);
        }

        public void PrintAllSheets(string igxlVersion)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
            {
                if (resourceName.EndsWith("SheetClassMapping." + igxlVersion + "_ultraflex.xml",
                    StringComparison.CurrentCultureIgnoreCase))
                {
                    _igxlConfig = LoadConfig(assembly.GetManifestResourceStream(resourceName));
                    break;
                }
            }

            string sheetVersion = "";
            foreach (KeyValuePair<string, IgxlSheet> igxlSheetPair in AllIgxlSheets)
            {
                try
                {
                    sheetVersion = GetSheetVersion(igxlSheetPair.Value);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
                var filePath = Path.GetDirectoryName(igxlSheetPair.Key);
                if (!Directory.Exists(filePath))
                    Directory.CreateDirectory(filePath);
                igxlSheetPair.Value.Write(igxlSheetPair.Key + ".txt", sheetVersion);
            }
        }

        public void PrintBinTable(string igxlVersion)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            foreach (var resourceName in resourceNames)
            {
                if (resourceName.EndsWith("SheetClassMapping." + igxlVersion + "_ultraflex.xml",
                    StringComparison.CurrentCultureIgnoreCase))
                {
                    _igxlConfig = LoadConfig(assembly.GetManifestResourceStream(resourceName));
                    break;
                }
            }

            string sheetVersion = "";
            foreach (KeyValuePair<string, IgxlSheet> igxlSheetPair in AllIgxlSheets)
            {
                if (!igxlSheetPair.Value.SheetName.Equals(MainBinTableName))
                    continue;

                try
                {
                    sheetVersion = GetSheetVersion(igxlSheetPair.Value);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
                var filePath = Path.GetDirectoryName(igxlSheetPair.Key);
                if (!Directory.Exists(filePath))
                    Directory.CreateDirectory(filePath);
                igxlSheetPair.Value.Write(igxlSheetPair.Key + ".txt", sheetVersion);
            }
        }

        public void Clear()
        {
            AllIgxlSheets.Clear();
            SubFlowSheets.Clear();
            MainFlowSheets.Clear();
            InsSheets.Clear();
            DcSpecSheets.Clear();
            AcSpecSheets.Clear();
            LevelSheets.Clear();
            TimeSetSheets.Clear();
            PatSetSheets.Clear();
            PatSetSubSheets.Clear();
            BinTblSheets.Clear();
            JobListSheets.Clear();
            ChannelMapSheets.Clear();
            CharSheets.Clear();
            WaveDefSheets.Clear();
            MixedSignalSheets.Clear();
            _glbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(null, null);
            _pinMapPair = new KeyValuePair<string, PinMapSheet>(null, null);
            PortMapSheets.Clear();
        }

        private string GetSheetVersion(IgxlSheet sheet)
        {
            return _igxlConfig.SheetItemClass.FirstOrDefault(p => p.sheetname.Equals(sheet.IgxlSheetName)).sheetversion;
        }

        private IGXL LoadConfig(Stream sr)
        {
            IGXL result;
            try
            {
                XmlSerializer xs = new XmlSerializer(typeof(IGXL));
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