using IgxlData.IgxlBase;
using IgxlData.IgxlWorkBooks;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlSheets
{
    public class JoblistSheet : IgxlSheet
    {
        private const string SheetType = "DTJobListSheet";

        public List<JobRow> JobRows;
        private readonly Dictionary<string, int> _headerIndex = new Dictionary<string, int>();

        public Dictionary<string, int> HeaderIndex { get { return _headerIndex; } }

        public JoblistSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            JobRows = new List<JobRow>();
            IgxlSheetName = IgxlSheetNameList.JobList;
        }

        public JoblistSheet(string sheetName)
            : base(sheetName)
        {
            JobRows = new List<JobRow>();
            IgxlSheetName = IgxlSheetNameList.JobList;
        }

        //protected override void WriteHeader()
        //{
        //    const string header = "DTJobListSheet,version=2.5:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1	Job List";
        //    IgxlWriter.WriteLine(header);
        //    IgxlWriter.WriteLine();
        //}

        //protected override void WriteColumnsHeader()
        //{
        //    var firstRow = new StringBuilder();
        //    var secondRow = new StringBuilder();
        //    firstRow.Append("\t\tSheet Parameters\t\t");
        //    secondRow.Append("\tJob Name\tPin Map\tTest Instances\tFlow Table\tAC Specs\tDC Specs\tPattern Sets\tPattern Groups\tBin Table\tCharacterization\tTest Procedures\tMixed Signal Timing\tWave Definitions\tPsets\tSignals\tPort Map\tFractional Bus\tConcurrent Sequence\tComment\t");
        //    IgxlWriter.WriteLine(firstRow.ToString());
        //    IgxlWriter.WriteLine(secondRow.ToString());
        //}

        //protected override void WriteRows()
        //{
        //    foreach (var job in JobRows)
        //    {
        //        var jobRow = new StringBuilder();
        //        jobRow.Append("\t");
        //        jobRow.Append(job.JobName);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.PinMap);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.TestInstances);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.FlowTable);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.AcSpecs);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.DcSpecs);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.PatternSets);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.PatternGroups);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.BinTable);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.Characterization);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.TestProcedures);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.MixedSignalTiming);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.WaveDefinitions);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.Psets);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.Signals);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.PortMap);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.FractionalBus);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.ConcurrentSequence);
        //        jobRow.Append("\t");
        //        jobRow.Append(job.Comment);
        //        IgxlWriter.WriteLine(jobRow.ToString());
        //    }
        //}

        public override void Write(string fileName, string version = "3.1")
        {
            //if (version == "2.5")
            //{
            //    GetSreamWriter(fileName);
            //    WriteHeader();
            //    WriteColumnsHeader();
            //    WriteRows();
            //    CloseStreamWriter();
            //}
            //else
            //{
            //    var versionDouble = Double.Parse(version);
            //    if (versionDouble > 3.1)
            //        versionDouble = 3.1;
            //    var validate = new Action<string>((a) => { });
            //    var genJobList = new GenJobSheet(fileName, validate, true, versionDouble);
            //    foreach (var row in JobRows)
            //    {

            //        var entryPara = new JobEntryParms();
            //        entryPara.ACSheetNames = row.AcSpecs;
            //        entryPara.BinTableSheetNames = row.BinTable;
            //        entryPara.CharacterizationSheetNames = row.Characterization;
            //        entryPara.Comment = row.Comment;
            //        entryPara.ConcurrentSequenceSheetNames = row.ConcurrentSequence;
            //        entryPara.DCSheetNames = row.DcSpecs;
            //        entryPara.FlowTableSheetName = row.FlowTable;
            //        entryPara.FractionalBusSheetNames = row.FractionalBus;
            //        entryPara.MixedSignalSheetNames = row.MixedSignalTiming;
            //        entryPara.Name = row.JobName;
            //        entryPara.PatternGroupSheetNames = row.PatternGroups;
            //        entryPara.PatternSetSheetNames = row.PatternSets;
            //        entryPara.PinMapSheetName = row.PinMap;
            //        entryPara.PortMapSheetNames = row.PortMap;
            //        entryPara.PsetsSheetNames = row.Psets;
            //        entryPara.SignalsSheetNames = row.Signals;
            //        entryPara.TestInstanceSheetNames = row.TestInstances;
            //        entryPara.TestProcedureSheetNames = row.TestProcedures;
            //        entryPara.WaveDefinitionSheetNames = row.WaveDefinitions;

            //        var entry = new UltraFlexJobSheetEntry(entryPara, row.JobName);
            //        genJobList.AddJob(row.JobName, entry);
            //    }
            //    genJobList.WriteSheet();
            //}

            //Support 2.5 & 3.1 & 3.2
            if (!Directory.Exists(Path.GetDirectoryName(fileName)))
                Directory.CreateDirectory(Path.GetDirectoryName(fileName));
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (dic.ContainsKey(version))
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (dic.ContainsKey("3.1"))
                {
                    var igxlSheetsVersion = dic["3.1"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
            }
        }

        private void WriteSheet(string fileName, string version, SheetInfo igxlSheetsVersion)
        {
            if (JobRows.Count == 0) return;

            using (var sw = new StreamWriter(fileName, false))
            {
                var maxCount = GetMaxCount(igxlSheetsVersion);
                var jobNameIndex = GetIndexFrom(igxlSheetsVersion, "Job Name");
                var pinMapIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Pin Map");
                var testInstancesIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Test Instances");
                var flowTableIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Flow Table");
                var aCSpecsIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "AC Specs");
                var dCSpecsIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "DC Specs");
                var patternSetsIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Pattern Sets");
                var binTableIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Bin Table");
                var characterizationIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Characterization");
                var mixedSignalTimingIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Mixed Signal Timing");
                var waveDefinitionsIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Wave Definitions");
                var psetsIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Psets");
                var signalsIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Signals");
                var portMapIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Port Map");
                var fractionalBusIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Fractional Bus");
                var concurrentSequenceIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "Concurrent Sequence");
                var spikeCheckConfigIndex = GetIndexFrom(igxlSheetsVersion, "Sheet Parameters", "SpikeCheck Config");
                var commentIndex = GetIndexFrom(igxlSheetsVersion, "Comment");

                #region headers
                var startRow = igxlSheetsVersion.Columns.RowCount;
                sw.WriteLine(SheetType + ",version=" + version + ":platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\t" + IgxlSheetName);
                for (var i = 1; i < startRow; i++)
                {
                    var arr = Enumerable.Repeat("", maxCount).ToArray();

                    SetField(igxlSheetsVersion, i, arr);

                    if (igxlSheetsVersion.Columns.Column != null)
                    {
                        foreach (var item in igxlSheetsVersion.Columns.Column)
                        {
                            if (item.rowIndex == i)
                                arr[item.indexFrom] = item.columnName;
                            if (item.Column1 != null)
                            {
                                foreach (var column1 in item.Column1)
                                {
                                    if (column1.rowIndex == i)
                                    {
                                        if (version=="3.1")
                                            arr[column1.indexFrom] = column1.columnName.Replace("AC Specs", "AC Spec");
                                        else
                                            arr[column1.indexFrom] = column1.columnName;
                                    }
                                }
                            }
                        }
                    }

                    WriteHeader(arr, sw);
                }

                #endregion

                #region data
                for (var index = 0; index < JobRows.Count; index++)
                {
                    var row = JobRows[index];
                    var arr = Enumerable.Repeat("", maxCount).ToArray();
                    if (!string.IsNullOrEmpty(row.JobName))
                    {
                        arr[0] = row.ColumnA;
                        arr[jobNameIndex] = row.JobName;
                        arr[pinMapIndex] = row.PinMap;
                        arr[testInstancesIndex] = row.TestInstances;
                        arr[flowTableIndex] = row.FlowTable;
                        arr[aCSpecsIndex] = row.AcSpecs;
                        arr[dCSpecsIndex] = row.DcSpecs;
                        arr[patternSetsIndex] = row.PatternSets;
                        arr[binTableIndex] = row.BinTable;
                        arr[characterizationIndex] = row.Characterization;
                        arr[mixedSignalTimingIndex] = row.MixedSignalTiming;
                        arr[waveDefinitionsIndex] = row.WaveDefinitions;
                        arr[psetsIndex] = row.Psets;
                        arr[signalsIndex] = row.Signals;
                        arr[portMapIndex] = row.PortMap;
                        arr[fractionalBusIndex] = row.FractionalBus;
                        arr[concurrentSequenceIndex] = row.ConcurrentSequence;
                        if (spikeCheckConfigIndex != -1)
                            arr[spikeCheckConfigIndex] = row.SpikeCheckConfig;
                        arr[commentIndex] = row.Comment;

                    }
                    else
                    {
                        arr = new[] { "\t" };
                    }
                    sw.WriteLine(string.Join("\t", arr));
                }
                #endregion
            }
        }

        public void AddRow(JobRow job)
        {
            JobRows.Add(job);
        }

        public JobRow GetRow(string job)
        {
            return JobRows.Find(x => x.JobName.Equals(job, StringComparison.OrdinalIgnoreCase));
        }
    }
}