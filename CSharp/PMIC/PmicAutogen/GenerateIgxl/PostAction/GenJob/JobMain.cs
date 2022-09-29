using IgxlData.IgxlBase;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.PostAction.GenJob
{
    public class JobListMain
    {
        private readonly Dictionary<IgxlSheet, string> _igxlSheets;

        public JobListMain(Dictionary<IgxlSheet, string> igxlSheets)
        {
            _igxlSheets = igxlSheets;
        }

        #region Member Function

        public KeyValuePair<IgxlSheet, string> WorkFlow()
        {
            if (StaticSetting.JobMap.Count == 0)
                return default(KeyValuePair<IgxlSheet, string>);

            var insSheets = TestProgram.IgxlWorkBk.InsSheets.Select(x => x.Value.SheetName).ToList();
            var patSetsSheets = TestProgram.IgxlWorkBk.PatSetSheets.Select(x => x.Value.SheetName).ToList();
            var portMapSheets = TestProgram.IgxlWorkBk.PortMapSheets.Select(x => x.Value.SheetName).ToList();
            AddInstanceSheetByExtraSheet(LocalSpecs.ExtraPath, ref insSheets);
            AddPatSetsSheetByExtraSheet(LocalSpecs.ExtraPath, ref patSetsSheets);
            AddPortMapSheetByExtraSheet(LocalSpecs.ExtraPath, ref portMapSheets);
            var acSpecs = string.Join(",",
                TestProgram.IgxlWorkBk.AcSpecSheets.Where(x => x.Value.CategoryList.Count > 0)
                    .Select(x => x.Value.SheetName));
            var dcSpecs = string.Join(",",
                TestProgram.IgxlWorkBk.DcSpecSheets.Where(x => x.Value.CategoryList.Count > 0)
                    .Select(x => x.Value.SheetName));
            var binTable = string.Join(",", TestProgram.IgxlWorkBk.BinTableSheets.Select(x => x.Value.SheetName));
            //var portMap = string.Join(",", TestProgram.IgxlWorkBk.PortMapSheets.Select(x => x.Value.SheetName));
            var characterization = string.Join(",", TestProgram.IgxlWorkBk.CharSheets.Select(x => x.Value.SheetName));
            var mixedSignalTiming =
                string.Join(",", TestProgram.IgxlWorkBk.MixedSignalSheets.Select(x => x.Value.SheetName));
            var waveDefinition = string.Join(",", TestProgram.IgxlWorkBk.WaveDefSheets.Select(x => x.Value.SheetName));

            var jobListSheet = new JobListSheet(PmicConst.JobList);
            foreach (var stageList in StaticSetting.JobMap)
            {
                var jobList = stageList.Value;
                foreach (var jobName in jobList)
                {
                    var jobRow = new JobRow();
                    jobRow.JobName = jobName;
                    if (TestProgram.IgxlWorkBk.PinMapPair.Value != null)
                        jobRow.PinMap = TestProgram.IgxlWorkBk.PinMapPair.Value.SheetName;
                    jobRow.TestInstance = string.Join(",", insSheets);
                    jobRow.FlowTable = FindJobMainFlow(jobName);
                    jobRow.AcSpecs = acSpecs;
                    jobRow.DcSpecs = dcSpecs;
                    jobRow.PatternSets = string.Join(",", patSetsSheets);
                    jobRow.PatternGroups = "";
                    jobRow.BinTable = binTable;
                    jobRow.Characterization = characterization;
                    jobRow.TestProcedures = "";
                    jobRow.MixedSignalTiming = mixedSignalTiming;
                    jobRow.WaveDefinition = waveDefinition;
                    jobRow.Signals = "";
                    jobRow.PortMap = string.Join(",", portMapSheets);
                    jobRow.FractionalBus = "";
                    jobRow.ConcurrentSequence = "";
                    jobRow.Comment = "";
                    jobListSheet.AddRow(jobRow);
                }
            }

            return new KeyValuePair<IgxlSheet, string>(jobListSheet, FolderStructure.DirJob);
        }

        private void AddInstanceSheetByExtraSheet(string extraFolder, ref List<string> insSheets)
        {
            var extraInstances = new List<string>();
            if (Directory.Exists(extraFolder))
            {
                var igxlManagerMain = new IgxlSheetReader();
                extraInstances = igxlManagerMain.GetSheetByType(extraFolder, SheetTypes.DTTestInstancesSheet);
            }

            insSheets.AddRange(extraInstances.Select(Path.GetFileNameWithoutExtension));
            insSheets = insSheets.Distinct().ToList();
        }

        private void AddPatSetsSheetByExtraSheet(string extraFolder, ref List<string> patSetsSheets)
        {
            var extraPatSets = new List<string>();
            if (Directory.Exists(extraFolder))
            {
                var igxlManagerMain = new IgxlSheetReader();
                extraPatSets = igxlManagerMain.GetSheetByType(extraFolder, SheetTypes.DTPatternSetSheet);
            }

            patSetsSheets.AddRange(extraPatSets.Select(Path.GetFileNameWithoutExtension));
            patSetsSheets = patSetsSheets.Distinct().ToList();
        }

        private void AddPortMapSheetByExtraSheet(string extraFolder, ref List<string> portMapSheets)
        {
            var extraPortMaps = new List<string>();
            if (Directory.Exists(extraFolder))
            {
                var igxlManagerMain = new IgxlSheetReader();
                extraPortMaps = igxlManagerMain.GetSheetByType(extraFolder, SheetTypes.DTPortMapSheet);
            }

            portMapSheets.AddRange(extraPortMaps.Select(Path.GetFileNameWithoutExtension));
            portMapSheets = portMapSheets.Distinct().ToList();
        }

        private string FindJobMainFlow(string job)
        {
            foreach (var mainFlow in _igxlSheets)
            {
                var sheetName = mainFlow.Key.SheetName;
                if (sheetName.ToUpper().Contains(job.ToUpper()))
                    return sheetName;
            }
            return _igxlSheets.First().Key.SheetName;
        }

        #endregion
    }
}