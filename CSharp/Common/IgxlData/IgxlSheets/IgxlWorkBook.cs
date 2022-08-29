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
        public IgxlWorkBook()
        {
            _pinMapPair = new KeyValuePair<string, PinMapSheet>();
            AllIgxlSheets = new Dictionary<string, IgxlSheet>(StringComparer.CurrentCultureIgnoreCase);
            SubFlowSheets = new Dictionary<string, SubFlowSheet>(StringComparer.CurrentCultureIgnoreCase);
            MainFlowSheets = new Dictionary<string, MainFlowSheet>(StringComparer.CurrentCultureIgnoreCase);
            InsSheets = new Dictionary<string, InstanceSheet>(StringComparer.CurrentCultureIgnoreCase);
            DcSpecSheets = new Dictionary<string, DcSpecSheet>(StringComparer.CurrentCultureIgnoreCase);
            AcSpecSheets = new Dictionary<string, AcSpecSheet>(StringComparer.CurrentCultureIgnoreCase);
            LevelSheets = new Dictionary<string, LevelSheet>(StringComparer.CurrentCultureIgnoreCase);
            TimeSetSheets = new Dictionary<string, TimeSetBasicSheet>(StringComparer.CurrentCultureIgnoreCase);
            PatSetSheets = new Dictionary<string, PatSetSheet>(StringComparer.CurrentCultureIgnoreCase);
            PatSetSubSheets = new Dictionary<string, PatSetSubSheet>(StringComparer.CurrentCultureIgnoreCase);
            BinTableSheets = new Dictionary<string, BinTableSheet>(StringComparer.CurrentCultureIgnoreCase);
            JobListSheets = new Dictionary<string, JobListSheet>(StringComparer.CurrentCultureIgnoreCase);
            ChannelMapSheets = new Dictionary<string, ChannelMapSheet>(StringComparer.CurrentCultureIgnoreCase);
            PortMapSheets = new Dictionary<string, PortMapSheet>(StringComparer.CurrentCultureIgnoreCase);
            WaveDefSheets = new Dictionary<string, WaveDefinitionSheet>(StringComparer.CurrentCultureIgnoreCase);
            MixedSignalSheets = new Dictionary<string, MixedSignalSheet>(StringComparer.CurrentCultureIgnoreCase);
            CharSheets = new Dictionary<string, CharSheet>(StringComparer.CurrentCultureIgnoreCase);
            FlowUsedInteger = new List<string>();
        }

        #region Field

        public const string MainBinTableName = "Bin_Table";
        private KeyValuePair<string, PinMapSheet> _pinMapPair;
        private KeyValuePair<string, GlobalSpecSheet> _glbSpecSheetPair;
        private IGXL _igxlConfig;

        #endregion

        #region Property

        public Dictionary<string, SubFlowSheet> SubFlowSheets { get; }

        public Dictionary<string, MainFlowSheet> MainFlowSheets { get; }

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
            BinTableSheet binTable = null;
            foreach (var binTblPair in BinTableSheets)
                if (binTblPair.Value.SheetName.Equals(MainBinTableName))
                {
                    binTable = binTblPair.Value;
                    break;
                }

            if (binTable == null)
                throw new Exception(string.Format(
                    "Can't find main bin table sheet, which sheet name is: {0} in IgxlWorkBook. ", MainBinTableName));
            return binTable;
        }

        public Dictionary<string, BinTableSheet> BinTableSheets { get; }

        public Dictionary<string, ChannelMapSheet> ChannelMapSheets { get; }

        public Dictionary<string, PortMapSheet> PortMapSheets { get; }

        public KeyValuePair<string, GlobalSpecSheet> GlbSpecSheetPair
        {
            get { return _glbSpecSheetPair; }
            set
            {
                if (_glbSpecSheetPair.Value == null)
                {
                    var subFileSubName = Contact(value.Key, value.Value.SheetName);
                    _glbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(subFileSubName, value.Value);
                    AllIgxlSheets.Add(subFileSubName, _glbSpecSheetPair.Value);
                }
                else
                {
                    _glbSpecSheetPair.Value.AddRows(value.Value.GlobalSpecsRows);
                }
            }
        }

        public KeyValuePair<string, PinMapSheet> PinMapPair
        {
            get { return _pinMapPair; }
            set
            {
                if (_pinMapPair.Value == null)
                {
                    var subFileSubName = Contact(value.Key, value.Value.SheetName);
                    _pinMapPair = new KeyValuePair<string, PinMapSheet>(subFileSubName, value.Value);
                    AllIgxlSheets.Add(subFileSubName, _pinMapPair.Value);
                }
                else
                {
                    _pinMapPair.Value.GroupList.AddRange(value.Value.GroupList);
                    _pinMapPair.Value.PinList.AddRange(value.Value.PinList);
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
                if (charSheet.Value.SheetName.Equals(sheetName, StringComparison.CurrentCultureIgnoreCase))
                    return charSheet.Value;
            return new CharSheet(sheetName);
        }

        private void AddSubFlowSheet(SubFlowSheet flowSheet, string fileSubPath)
        {
            if (flowSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, flowSheet.SheetName);
            if (!SubFlowSheets.ContainsKey(subFileSubName))
            {
                SubFlowSheets.Add(subFileSubName, flowSheet);
                AllIgxlSheets.Add(subFileSubName, flowSheet);
            }
        }

        private void AddMainFlowSheet(MainFlowSheet flowSheet, string fileSubPath)
        {
            if (flowSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, flowSheet.SheetName);
            if (!MainFlowSheets.ContainsKey(subFileSubName))
            {
                MainFlowSheets.Add(subFileSubName, flowSheet);
                AllIgxlSheets.Add(subFileSubName, flowSheet);
            }
        }

        private void AddInsSheet(InstanceSheet insSheet, string fileSubPath)
        {
            if (insSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, insSheet.SheetName);
            if (!InsSheets.ContainsKey(subFileSubName))
            {
                InsSheets.Add(subFileSubName, insSheet);
                AllIgxlSheets.Add(subFileSubName, insSheet);
            }
        }

        private void AddDcSpecSheet(DcSpecSheet dcSpecSheet, string fileSubPath)
        {
            if (dcSpecSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, dcSpecSheet.SheetName);
            if (!DcSpecSheets.ContainsKey(subFileSubName))
            {
                DcSpecSheets.Add(subFileSubName, dcSpecSheet);
                AllIgxlSheets.Add(subFileSubName, dcSpecSheet);
            }
        }

        private void AddAcSpecSheet(AcSpecSheet acSpecSheet, string fileSubPath)
        {
            if (acSpecSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, acSpecSheet.SheetName);
            if (!AcSpecSheets.ContainsKey(subFileSubName))
            {
                AcSpecSheets.Add(subFileSubName, acSpecSheet);
                AllIgxlSheets.Add(subFileSubName, acSpecSheet);
            }
        }

        private void AddLevelSheet(LevelSheet levelSheet, string fileSubPath)
        {
            if (levelSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, levelSheet.SheetName);
            if (!LevelSheets.ContainsKey(subFileSubName))
            {
                LevelSheets.Add(subFileSubName, levelSheet);
                AllIgxlSheets.Add(subFileSubName, levelSheet);
            }
        }

        private void AddTimeSetSheet(TimeSetBasicSheet timeSetSheet, string fileSubPath)
        {
            if (timeSetSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, timeSetSheet.SheetName);
            if (!TimeSetSheets.ContainsKey(subFileSubName))
            {
                TimeSetSheets.Add(subFileSubName, timeSetSheet);
                AllIgxlSheets.Add(subFileSubName, timeSetSheet);
            }
        }

        private void AddPatSetSheet(PatSetSheet patSetSheet, string fileSubPath)
        {
            if (patSetSheet == null) return;
            if (!patSetSheet.PatSetRows.Any()) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, patSetSheet.SheetName);
            if (!PatSetSheets.ContainsKey(subFileSubName))
            {
                PatSetSheets.Add(subFileSubName, patSetSheet);
                AllIgxlSheets.Add(subFileSubName, patSetSheet);
            }
        }

        private void AddPatSetSubSheet(PatSetSubSheet patSetSubSheet, string fileSubPath)
        {
            if (patSetSubSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, patSetSubSheet.SheetName);
            if (!PatSetSubSheets.ContainsKey(subFileSubName))
            {
                PatSetSubSheets.Add(subFileSubName, patSetSubSheet);
                AllIgxlSheets.Add(subFileSubName, patSetSubSheet);
            }
        }

        private void AddBinTblSheet(BinTableSheet binTableSheet, string fileSubPath)
        {
            if (binTableSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, binTableSheet.SheetName);
            if (!BinTableSheets.ContainsKey(subFileSubName))
            {
                BinTableSheets.Add(subFileSubName, binTableSheet);
                AllIgxlSheets.Add(subFileSubName, binTableSheet);
            }
            else
            {
                BinTableSheets[subFileSubName].BinTableRows.AddRange(binTableSheet.BinTableRows);
            }
        }

        private void AddJobListSheet(JobListSheet jobListSheet, string fileSubPath)
        {
            if (jobListSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, jobListSheet.SheetName);
            if (!JobListSheets.ContainsKey(subFileSubName))
            {
                JobListSheets.Add(subFileSubName, jobListSheet);
                AllIgxlSheets.Add(subFileSubName, jobListSheet);
            }
        }

        private void AddChannelMapSheet(ChannelMapSheet channelMapSheet, string fileSubPath)
        {
            if (channelMapSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, channelMapSheet.SheetName);
            if (!ChannelMapSheets.ContainsKey(subFileSubName))
            {
                ChannelMapSheets.Add(subFileSubName, channelMapSheet);
                AllIgxlSheets.Add(subFileSubName, channelMapSheet);
            }
        }

        private void AddPortMapSheet(PortMapSheet portMapSheet, string fileSubPath)
        {
            if (portMapSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, portMapSheet.SheetName);
            if (!PortMapSheets.ContainsKey(subFileSubName))
            {
                PortMapSheets.Add(subFileSubName, portMapSheet);
                AllIgxlSheets.Add(subFileSubName, portMapSheet);
            }
        }

        private void AddCharSheet(CharSheet charSheet, string fileSubPath)
        {
            if (charSheet == null) return;
            var subFileSubName = GetSubFileSubName(fileSubPath, charSheet.SheetName);
            if (!CharSheets.ContainsKey(subFileSubName))
            {
                CharSheets.Add(subFileSubName, charSheet);
                AllIgxlSheets.Add(subFileSubName, charSheet);
            }
        }

        private string GetSubFileSubName(string fileSubPath, string sheetName)
        {
            return Contact(fileSubPath, sheetName);
        }

        private void GetIgxlConfig(string igxlVersion)
        {
            //var assembly = Assembly.GetExecutingAssembly();
            //var resourceNames = assembly.GetManifestResourceNames();
            //foreach (var resourceName in resourceNames)
            //{
            //    if (resourceName.EndsWith("SheetClassMapping." + igxlVersion + "_ultraflex.xml",
            //        StringComparison.CurrentCultureIgnoreCase))
            //    {
            //        _igxlConfig = LoadConfig(assembly.GetManifestResourceStream(resourceName));
            //        break;
            //    }
            //}

            var exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().GetName().CodeBase)
                .Replace("file:\\", "");
            var files = Directory.GetFiles(Path.Combine(exePath, "IGDataXML\\SheetClassMapping"));
            foreach (var file in files)
                if (file.EndsWith(igxlVersion + "_ultraflex.xml", StringComparison.CurrentCultureIgnoreCase))
                {
                    _igxlConfig = LoadConfig(File.OpenRead(file));
                    break;
                }
        }

        public void PrintAllSheets(string igxlVersion)
        {
            GetIgxlConfig(igxlVersion);

            var sheetVersion = "";
            foreach (var igxlSheetPair in AllIgxlSheets)
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
            GetIgxlConfig(igxlVersion);

            var sheetVersion = "";
            foreach (var igxlSheetPair in AllIgxlSheets)
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
            BinTableSheets.Clear();
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
                var xs = new XmlSerializer(typeof(IGXL));
                var sysData = (IGXL) xs.Deserialize(sr);
                sr.Close();
                result = sysData;
            }
            catch (Exception e)
            {
                throw new Exception(e.Message);
            }

            return result;
        }

        private string Contact(string mainDir, string subDir)
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

        public void AddIgxlSheets(Dictionary<IgxlSheet, string> igxlSheets)
        {
            foreach (var igxlSheet in igxlSheets)
                if (igxlSheet.Key is MainFlowSheet)
                    AddMainFlowSheet((MainFlowSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is ChannelMapSheet)
                    AddChannelMapSheet((ChannelMapSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is PinMapSheet)
                    PinMapPair = new KeyValuePair<string, PinMapSheet>(igxlSheet.Value, (PinMapSheet) igxlSheet.Key);
                else if (igxlSheet.Key is PortMapSheet)
                    AddPortMapSheet((PortMapSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is SubFlowSheet)
                    AddSubFlowSheet((SubFlowSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is BinTableSheet)
                    AddBinTblSheet((BinTableSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is InstanceSheet)
                    AddInsSheet((InstanceSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is PatSetSheet)
                    AddPatSetSheet((PatSetSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is GlobalSpecSheet)
                    GlbSpecSheetPair =
                        new KeyValuePair<string, GlobalSpecSheet>(igxlSheet.Value, (GlobalSpecSheet) igxlSheet.Key);
                else if (igxlSheet.Key is DcSpecSheet)
                    AddDcSpecSheet((DcSpecSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is AcSpecSheet)
                    AddAcSpecSheet((AcSpecSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is LevelSheet)
                    AddLevelSheet((LevelSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is TimeSetBasicSheet)
                    AddTimeSetSheet((TimeSetBasicSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is PatSetSubSheet)
                    AddPatSetSubSheet((PatSetSubSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is CharSheet)
                    AddCharSheet((CharSheet) igxlSheet.Key, igxlSheet.Value);
                else if (igxlSheet.Key is JobListSheet)
                    AddJobListSheet((JobListSheet) igxlSheet.Key, igxlSheet.Value);
        }

        //public void Add(string path, IgxlSheet igxlSheet)
        //{
        //    if (igxlSheet is ChannelMapSheet)
        //        AddChannelMapSheet(path, (ChannelMapSheet)igxlSheet);
        //    else if (igxlSheet is PinMapSheet)
        //        PinMapPair = new KeyValuePair<string, PinMapSheet>(path, (PinMapSheet)igxlSheet);
        //    else if (igxlSheet is PortMapSheet)
        //        AddPortMapSheet(path, (PortMapSheet)igxlSheet);
        //    else if (igxlSheet is SubFlowSheet)
        //        AddSubFlowSheet(path, (SubFlowSheet)igxlSheet);
        //    else if (igxlSheet is BinTableSheet)
        //        AddBinTblSheet(path, (BinTableSheet)igxlSheet);
        //    else if (igxlSheet is InstanceSheet)
        //        AddInsSheet(path, (InstanceSheet)igxlSheet);
        //    else if (igxlSheet is PatSetSheet)
        //        AddPatSetSheet(path, (PatSetSheet)igxlSheet);
        //    else if (igxlSheet is GlobalSpecSheet)
        //        GlbSpecSheetPair = new KeyValuePair<string, GlobalSpecSheet>(path, (GlobalSpecSheet)igxlSheet);
        //    else if (igxlSheet is DcSpecSheet)
        //        AddDcSpecSheet(path, (DcSpecSheet)igxlSheet);
        //    else if (igxlSheet is AcSpecSheet)
        //        AddAcSpecSheet(path, (AcSpecSheet)igxlSheet);
        //    else if (igxlSheet is LevelSheet)
        //        AddLevelSheet(path, (LevelSheet)igxlSheet);
        //    else if (igxlSheet is TimeSetBasicSheet)
        //        AddTimeSetSheet(path, (TimeSetBasicSheet)igxlSheet);
        //    else if (igxlSheet is PatSetSubSheet)
        //        AddPatSetSubSheet(path, (PatSetSubSheet)igxlSheet);
        //    else if (igxlSheet is CharSheet)
        //        AddCharSheet(path, (CharSheet)igxlSheet);
        //}

        #endregion
    }
}