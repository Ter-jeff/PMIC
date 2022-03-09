using IgxlData.IgxlBase;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using BinTableSheet = IgxlData.IgxlSheets.BinTableSheet;
using FlowSheet = IgxlData.IgxlSheets.SubFlowSheet;
using GlobalSpecSheet = IgxlData.IgxlSheets.GlobalSpecSheet;
using PinMapSheet = IgxlData.IgxlSheets.PinMapSheet;

namespace BinOutCheck
{
    public class IgxlProgram
    {
        /* properities */
        //private readonly string _selectJob;

        public readonly string TpPath;

        //private readonly IgxlProgramLoader _loader;

        public string Name;


        //private readonly IgxlTestProgram _testProgram;

        private readonly List<FlowRow> _allFlowSteps = new List<FlowRow>();

        public List<BinTableSheet> BintableSheets = new List<BinTableSheet>();

        public List<FlowSheet> FlowSheets = new List<FlowSheet>();

        public List<string> Modules = new List<string>();

        //public List<TimeSetBasicSheet> TimeSetSheets = new List<TimeSetBasicSheet>();

        //public List<PatSetSheet> PatSetSheets = new List<PatSetSheet>();

        //public List<InstanceSheet> InstanceSheets = new List<InstanceSheet>();

        public JoblistSheet JoblistSheet;

        public List<PinMapSheet> PinMapSheets;

        public List<GlobalSpecSheet> GlobalSpecSheets;

        //public List<Pin> PinList;

        //public List<PinGroup> PinGroups;

        // public List<ChannelMapSheet> ChannelMapSheets = new List<ChannelMapSheet>();

        //public IgxlProgram(string programName, string job)
        //{
        //    _selectJob = job;
        //    Name = programName;
        //}

        public void LoadJobSheet(string path)
        {
            JoblistSheet =
               new ReadJobListSheet().GetIgxlSheets(path, SheetType.DTJobListSheet)
                   .Cast<JoblistSheet>()
                   .ToList()[0];
        }



        public void LoadIgxlProgramAsync(string path)
        {
            var readBinTable = new Task(() =>
            {
                BintableSheets =
                    new ReadBinTableSheet().GetIgxlSheets(path, SheetType.DTBintablesSheet)
                    .Cast<BinTableSheet>()
                    .ToList();
            });

            var readVBT = new Task(() =>
            {
                Modules = GetModules(path);
            });

            //var readInstance = new Task(() =>
            //{
            //    InstanceSheets =
            //        new ReadInstanceSheet().GetIgxlSheets(path, SheetType.DTTestInstancesSheet)
            //            .Cast<InstanceSheet>()
            //            .ToList();
            //});

            var readFlow = new Task(() =>
            {
                FlowSheets =
                    new ReadFlowSheet().GetIgxlSheets(path, SheetType.DTFlowtableSheet)
                        .Cast<FlowSheet>()
                        .ToList();
            });
            
            //var readTimeset = new Task(() =>
            //{
            //    TimeSetSheets =
            //        new ReadTimeSetSheet().GetIgxlSheets(path, SheetType.DTTimesetBasicSheet)
            //            .Cast<TimeSetBasicSheet>()
            //            .ToList();
            //});

            //var readPattern = new Task(() =>
            //{
            //    PatSetSheets =
            //        new ReadPatSetSheet().GetIgxlSheets(path, SheetType.DTPatternSetSheet)
            //            .Cast<PatSetSheet>()
            //            .ToList();
            //});

            //var readOthers = new Task(() =>
            //{
            //    JoblistSheet =
            //        new ReadJobListSheet().GetIgxlSheets(path, SheetType.DTJobListSheet)
            //            .Cast<JoblistSheet>()
            //            .ToList()[0];


            //    PinMapSheets = new ReadPinMapSheet().GetIgxlSheets(path, SheetType.DTPinMap)
            //        .Cast<PinMapSheet>().ToList();


                //{
                //    var pinMaps = new IgxlSheetReader().GetSheetByType(path, SheetType.DTPinMap);
                //    var pinmapreader = new ReadPinMapSheet();

                //    var pinMapSheet = pinmapreader.GetSheet(pinMaps[0]);
                //    PinList = pinMapSheet.PinList;
                //    PinGroups = pinMapSheet.GroupList;
                //}

                //GlobalSpecSheets = new ReadGlobalSpecSheet().GetIgxlSheets(path, SheetType.DTGlobalSpecSheet)
                //                .Cast<GlobalSpecSheet>().ToList();

                //ChannelMapSheets =
                //    new ReadChanMapSheet().GetIgxlSheets(path, SheetType.DTChanMap)
                //        .Cast<ChannelMapSheet>()
                //        .ToList();



            //});
            readBinTable.Start();
            //readInstance.Start();
            readFlow.Start();
            readVBT.Start();
            //readTimeset.Start();
            //readPattern.Start();
            //readOthers.Start();
            Task.WaitAll(readBinTable, readFlow, readVBT);
        }

        private List<string> GetModules(string exportFolder)
        {
            DirectoryInfo dir = new DirectoryInfo(exportFolder);

            FileInfo[] allbas = dir.GetFiles("*.bas");
            FileInfo[] allcls = dir.GetFiles("*.cls");
            List<string> result = new List<string>();
            for (int i = 0; i < allbas.Length; i++)
            {
                result.Add(allbas[i].FullName);
            }

            for (int i = 0; i < allcls.Length; i++)
            {
                result.Add(allcls[i].FullName);
            }

            return result;
        }

        public List<InstanceSheet> LoadInstanceSheet(string exportFolder, string jobName, string sheetKeyWord = ".*")
        {
            JobRow jobRow = GetSelectedJobRow(jobName);
            var enableInstanceSheetName = jobRow.TestInstances.Split(',').ToList().Where(ist => Regex.IsMatch(ist, sheetKeyWord, RegexOptions.IgnoreCase)).ToList();

            var instanceSheetPath = new List<string>();

            foreach (var sheetName in enableInstanceSheetName)
            {
                instanceSheetPath.Add(Path.Combine(exportFolder, sheetName + ".txt"));
            }



            var instanceSheets =
                    new ReadInstanceSheet().GetIgxlSheets(instanceSheetPath, SheetType.DTTestInstancesSheet)
                        .Cast<InstanceSheet>()
                        .ToList();


            return instanceSheets;
        }




        //public List<InstanceSheet> InstList
        //{
        //    get
        //    {
        //        var instList = new List<InstanceSheet>();

        //        // get job objcet
        //        //var job = _testProgram.Jobs.FirstOrDefault(a => a.Name.Equals(_selectJob, StringComparison.OrdinalIgnoreCase));
        //        if (Joblist.JobRows.All
        //            (p => !p.JobName.Equals(_selectJob, StringComparison.OrdinalIgnoreCase))) return instList;

        //        // get inst from inst sheet specified in the job sheet
        //        //instList.AddRange(
        //        //    job.GetSheetsByType(SheetType.DTTestInstancesSheet)
        //        //    .Where(sh => sh != null)
        //        //    .Cast<TestInstanceSheet>()
        //        //    .SelectMany(instSheet => instSheet.TestInstances.Where(inst => inst != null  && !string.IsNullOrEmpty(inst.Name))));

        //        return InstanceSheets;
        //    }
        //}

        public Dictionary<string, List<string>> ReadPatSet(string patSetFilePath)
        {
            var patSetDict = new Dictionary<string, List<string>>();

            try
            {
                // Read the file and display it line by line.  
                var file = new StreamReader(patSetFilePath);
                var isHeader = true;
                var subPatCol = 0;
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    var tokens = line.Split('\t');
                    if (tokens.Count() <= 3) continue;

                    if (isHeader)
                    {
                        if (tokens[1] != "Pattern Set") continue;

                        for (var i = 2; i < tokens.Count(); i++)
                        {
                            if (tokens[i] != "File/Group Name") continue;
                            subPatCol = i;
                            isHeader = false;
                            break;
                        }
                    }
                    else
                    {
                        if (tokens.Count() < subPatCol) continue;
                        var patName = tokens[1].ToUpper();
                        var subPat = tokens[subPatCol].ToUpper();
                        if (subPat.Contains("\\"))
                            patSetDict[patName] = new List<string> { patName };
                        else if (patSetDict.ContainsKey(patName))
                            patSetDict[patName].Add(subPat);
                        else
                            patSetDict[patName] = new List<string> { subPat };
                    }
                }

                file.Close();
            }
            catch (IOException e)
            {
                Console.WriteLine(e.Message);
            }

            return patSetDict;
        }

        public string TestProgramVersion
        {
            get
            {
                var reTpName = new Regex(@"^(?<Partid>(\w{2}\d{2}))_(?<Sortstage>(((CP)|(FT)|(QA)|(WLFT))\d*))_(?<Dut>(X\d{1,2}))_(?<Efusecode>(E\d{2}))_(?<Daycode>(\d{6}))_(?<EngProdVersion>(V\d{2}\w)).*", RegexOptions.IgnoreCase | RegexOptions.Compiled);
                var tpbasename = Path.GetFileNameWithoutExtension(TpPath);
                if (tpbasename == null) return "";
                var matchTPname = reTpName.Match(tpbasename);
                return !matchTPname.Success ? tpbasename : matchTPname.Groups["EngProdVersion"].ToString();
            }
        }

        /* constructor */
        public IgxlProgram(string tpNamePath)
        {
            TpPath = tpNamePath;
            //_selectJob = jobName;
            //_loader = loader;
            //_testProgram = loader.Load();
            //_allFlowSteps = _WalkFlowSheets();
        }

        /* methods */
        public List<string> GetJobList()
        {
            return JoblistSheet.JobRows.Select(p => p.JobName).ToList();
        }

        public JobRow GetSelectedJobRow(string job)
        {
            var jobRow =
              JoblistSheet.JobRows.FirstOrDefault(
                  a => a.JobName.Equals(job, StringComparison.OrdinalIgnoreCase));

            return jobRow;

        }



        public List<FlowRow> GetAllFlowSteps(string job)
        {
            var _allFlowSteps = new List<FlowRow>();
            var mainFlowSheet = GetMainFlowSheet(job);
            if (mainFlowSheet == null) return _allFlowSteps;
            ReadFlowSheet(mainFlowSheet, _allFlowSteps);
            return _allFlowSteps;
        }

        public List<FlowRow> GetInstancesInFlow()
        {
            return _allFlowSteps.Where(flowstep => flowstep.Opcode.ToLower().Trim() == "test").ToList();
        }

        public FlowSheet GetMainFlowSheet(string job)
        {
            SubFlowSheet sheet = null;
            var UsedJob = JoblistSheet.JobRows.FirstOrDefault
                (p => !p.JobName.Equals(job, StringComparison.OrdinalIgnoreCase));
            if (UsedJob == null) return sheet;

            return GetFlowSheet(UsedJob.FlowTable);
            //return FlowSheets.FirstOrDefault(p => p.Name.Equals(UsedJob.FlowTable, StringComparison.OrdinalIgnoreCase));
        }

        public FlowSheet GetFlowSheet(string flowName)
        {
            return FlowSheets.FirstOrDefault(p => p.Name.Equals(flowName, StringComparison.OrdinalIgnoreCase));
        }


        private void ReadFlowSheet(FlowSheet flowSheet, ICollection<FlowRow> flowSteps)
        {
            if (flowSheet == null)
                return;

            foreach (var flowstep in flowSheet.FlowRows)
            {
                if (flowstep.Opcode.ToLower().Trim() == "return")
                    return;

                if (flowstep.Opcode.ToLower().Trim() == "call")
                    ReadFlowSheet(
                        FlowSheets.FirstOrDefault
                        (p => p.Name.Equals(flowstep.Parameter, StringComparison.OrdinalIgnoreCase))
                        , flowSteps);
                else
                    flowSteps.Add(flowstep);
            }
        }
    }
}