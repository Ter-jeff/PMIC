using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using Teradyne.Oasis.IGData;
using Teradyne.Oasis.IGData.UltraFlex;
using Teradyne.Oasis.IGLinkBase.ProgramGeneration;
using Teradyne.Oasis.IGData.Utilities;
using System.Linq;
using System.IO;

namespace IgxlData.IgxlSheets
{
    public class JobListSheet : IgxlSheet
    {
        private const string SheetType = "DTJobListSheet";
        public List<JobRow> JobRows;

        public Dictionary<string, int> HeaderIndex { get; set; } = new Dictionary<string, int>();

        public JobListSheet(ExcelWorksheet sheet)
            : base(sheet)
        {
            JobRows = new List<JobRow>();
            IgxlSheetName = IgxlSheetNameList.JobList;
        }

        public JobListSheet(string sheetName)
            : base(sheetName)
        {
            JobRows = new List<JobRow>();
            IgxlSheetName = IgxlSheetNameList.JobList;
        }

        protected override void WriteHeader()
        {
            const string header = "DTJobListSheet,version=2.5:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1	Job List";
            IgxlWriter.WriteLine(header);
            IgxlWriter.WriteLine();
            IgxlWriter.WriteLine();
            IgxlWriter.WriteLine();
        }

        protected override void WriteColumnsHeader()
        {
            var firstRow = new StringBuilder();
            var secondRow = new StringBuilder();
            firstRow.Append("\t\tSheet Parameters\t\t");
            secondRow.Append("\tJob Name\tPin Map\tTest Instances\tFlow Table\tAC Specs\tDC Specs\tPattern Sets\tPattern Groups\tBin Table\tCharacterization\tTest Procedures\tMixed Signal Timing\tWave Definitions\tPsets\tSignals\tPort Map\tFractional Bus\tConcurrent Sequence\tComment\t");
            IgxlWriter.WriteLine(firstRow.ToString());
            IgxlWriter.WriteLine(secondRow.ToString());
        }

        protected override void WriteRows()
        {
            foreach (var job in JobRows)
            {
                var jobRow = new StringBuilder();
                jobRow.Append("\t");
                jobRow.Append(job.JobName);
                jobRow.Append("\t");
                jobRow.Append(job.PinMap);
                jobRow.Append("\t");
                jobRow.Append(job.TestInstance);
                jobRow.Append("\t");
                jobRow.Append(job.FlowTable);
                jobRow.Append("\t");
                jobRow.Append(job.AcSpecs);
                jobRow.Append("\t");
                jobRow.Append(job.DcSpecs);
                jobRow.Append("\t");
                jobRow.Append(job.PatternSets);
                jobRow.Append("\t");
                jobRow.Append(job.PatternGroups);
                jobRow.Append("\t");
                jobRow.Append(job.BinTable);
                jobRow.Append("\t");
                jobRow.Append(job.Characterization);
                jobRow.Append("\t");
                jobRow.Append(job.TestProcedures);
                jobRow.Append("\t");
                jobRow.Append(job.MixedSignalTiming);
                jobRow.Append("\t");
                jobRow.Append(job.WaveDefinition);
                jobRow.Append("\t");
                jobRow.Append(job.PSets);
                jobRow.Append("\t");
                jobRow.Append(job.Signals);
                jobRow.Append("\t");
                jobRow.Append(job.PortMap);
                jobRow.Append("\t");
                jobRow.Append(job.FractionalBus);
                jobRow.Append("\t");
                jobRow.Append(job.ConcurrentSequence);
                jobRow.Append("\t");
                jobRow.Append(job.Comment);
                IgxlWriter.WriteLine(jobRow.ToString());
            }
        }

        public override void Write(string fileName, string version)
        {
            var sheetClassMapping = GetIgxlSheetsVersion();
            if (sheetClassMapping.ContainsKey(IgxlSheetName))
            {
                var dic = sheetClassMapping[IgxlSheetName];
                if (version=="2.5")
                {
                    var igxlSheetsVersion = dic[version];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if (dic.ContainsKey("3.1"))
                {
                    var igxlSheetsVersion = dic["3.1"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else if(version=="3.2")
                {
                    var igxlSheetsVersion = dic["3.2"];
                    WriteSheet(fileName, version, igxlSheetsVersion);
                }
                else
                {
                    throw new Exception(string.Format("The JobList sheet version:{0} is not supported!", version));
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
                                        if (version == "3.1")
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
                        arr[testInstancesIndex] = row.TestInstance;
                        arr[flowTableIndex] = row.FlowTable;
                        arr[aCSpecsIndex] = row.AcSpecs;
                        arr[dCSpecsIndex] = row.DcSpecs;
                        arr[patternSetsIndex] = row.PatternSets;
                        arr[binTableIndex] = row.BinTable;
                        arr[characterizationIndex] = row.Characterization;
                        arr[mixedSignalTimingIndex] = row.MixedSignalTiming;
                        arr[waveDefinitionsIndex] = row.WaveDefinition;
                        arr[psetsIndex] = row.PSets;
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
