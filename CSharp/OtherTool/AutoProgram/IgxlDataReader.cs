using IgxlData;
using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using IgxlData.NonIgxlSheets;
using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoProgram
{
    public class IgxlDataReader
    {
        private readonly string _jobName;
        public List<AcSpecSheet> AcSpecSheets = new List<AcSpecSheet>();
        public List<DcSpecSheet> DcSpecSheets = new List<DcSpecSheet>();
        public List<SubFlowSheet> FlowSheets = new List<SubFlowSheet>();
        public List<InstanceSheet> InstanceSheets = new List<InstanceSheet>();
        public PatSetSheet PatSetsAll;
        public PatSetSubSheet PatSetSubSheet = new PatSetSubSheet("Pattern_Subroutine");
        public List<PinMapSheet> PinMapSheets = new List<PinMapSheet>();
        public List<TimeSetBasicSheet> TimeSetBasicSheets = new List<TimeSetBasicSheet>();
        public VbtFunctionLib VbtFunctionLib = new VbtFunctionLib();

        public IgxlDataReader(string testProgram, string jobName)
        {
            VbtFunctionLib.Read(testProgram);

            var igxlSheetReader = new IgxlSheetReader();
            using (var zip = new ZipFile(testProgram))
            {
                var zipArchiveEntries = zip.Entries.ToList();
                foreach (var zipArchiveEntry in zipArchiveEntries)
                {
                    var sheetName = Path.GetFileNameWithoutExtension(zipArchiveEntry.FileName);
                    var stream = zipArchiveEntry.OpenReader();
                    string firstLine;
                    using (var sr = new StreamReader(stream))
                        firstLine = sr.ReadLine();

                    var sheetType = igxlSheetReader.GetIgxlSheetType(firstLine);
                    if (sheetType == SheetTypes.DTFlowtableSheet)
                        FlowSheets.Add(new ReadFlowSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName));
                    if (sheetType == SheetTypes.DTTestInstancesSheet)
                        InstanceSheets.Add(
                            new ReadInstanceSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName));
                    else if (sheetType == SheetTypes.DTJobListSheet)
                        JobListSheet = new ReadJobListSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName);
                    else if (sheetType == SheetTypes.DTGlobalSpecSheet)
                        GlobalSpecSheet = new ReadGlobalSpecSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName);
                    else if (sheetType == SheetTypes.DTTimesetBasicSheet)
                        TimeSetBasicSheets.Add(new ReadTimeSetSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName));
                    else if (sheetType == SheetTypes.DTPortMapSheet)
                        PortMapSheet = new ReadPortMapSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName);
                    else if (sheetType == SheetTypes.DTACSpecSheet)
                        AcSpecSheets.Add(new ReadAcSpecSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName));
                    else if (sheetType == SheetTypes.DTDCSpecSheet)
                        DcSpecSheets.Add(new ReadDcSpecSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName));
                    else if (sheetType == SheetTypes.DTPinMap)
                        PinMapSheets.Add(new ReadPinMapSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName));
                    else if (sheetType == SheetTypes.DTPatternSetSheet &&
                             sheetName.Equals("PatSets_All", StringComparison.CurrentCultureIgnoreCase))
                        PatSetsAll = new ReadPatSetSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName);
                    else if (sheetType == SheetTypes.DTPatternSubroutineSheet &&
                             sheetName.Equals("Pattern_Subroutine", StringComparison.CurrentCultureIgnoreCase))
                        PatSetSubSheet = new ReadPatSubroutineSheet().GetSheet(zipArchiveEntry.OpenReader(), sheetName);
                }
            }
            _jobName = jobName;
        }

        public AcSpecSheet CurrentAcSpecSheet
        {
            get
            {
                if (JobListSheet != null)
                {
                    var jobRow = JobListSheet.GetRow(_jobName);
                    if (jobRow != null)
                        if (AcSpecSheets.Exists(x =>
                                x.SheetName.Equals(jobRow.AcSpecs, StringComparison.CurrentCultureIgnoreCase)))
                            return AcSpecSheets.Find(x =>
                                x.SheetName.Equals(jobRow.AcSpecs, StringComparison.CurrentCultureIgnoreCase));
                }

                if (AcSpecSheets.Count > 0)
                    return AcSpecSheets.First();
                return null;
            }
        }

        public DcSpecSheet CurrentDcSpecSheet
        {
            get
            {
                if (JobListSheet != null)
                {
                    var jobRow = JobListSheet.GetRow(_jobName);
                    if (jobRow != null)
                        if (DcSpecSheets.Exists(x =>
                                x.SheetName.Equals(jobRow.DcSpecs, StringComparison.CurrentCultureIgnoreCase)))
                            return DcSpecSheets.Find(x =>
                                x.SheetName.Equals(jobRow.DcSpecs, StringComparison.CurrentCultureIgnoreCase));
                }

                if (DcSpecSheets.Count > 0)
                    return DcSpecSheets.First();
                return null;
            }
        }

        public PinMapSheet CurrentPinMapSheet
        {
            get
            {
                if (JobListSheet != null)
                {
                    var jobRow = JobListSheet.GetRow(_jobName);
                    if (jobRow != null)
                        if (PinMapSheets.Exists(x =>
                                x.SheetName.Equals(jobRow.PinMap, StringComparison.CurrentCultureIgnoreCase)))
                            return PinMapSheets.Find(x =>
                                x.SheetName.Equals(jobRow.PinMap, StringComparison.CurrentCultureIgnoreCase));
                }

                if (PinMapSheets.Count > 0)
                    return PinMapSheets.First();
                return null;
            }
        }

        public JobListSheet JobListSheet { get; set; }
        public GlobalSpecSheet GlobalSpecSheet { get; set; }
        public PortMapSheet PortMapSheet { get; set; }
    }
}