using CommonLib.Extension;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace IgxlData.IgxlReader
{
    public class ReadJobListSheet : IgxlSheetReader
    {
        private const string ConHeaderJobName = "Job Name";
        private const string ConHeaderPinMap = "Pin Map";
        private const string ConHeaderTestIns = "Test Instances";
        private const string ConHeaderFlowTable = "Flow Table";
        private const string ConHeaderAcSpecs = "AC Specs";
        private const string ConHeaderAcSpec = "AC Spec";
        private const string ConHeaderDcSpec = "DC Specs";
        private const string ConHeaderPatternSets = "Pattern Sets";
        private const string ConHeaderBinTable = "Bin Table";
        private const string ConHeaderCharacterization = "Characterization";
        private const string ConHeaderMixSignalTiming = "Mixed Signal Timing";
        private const string ConHeaderWaveDef = "Wave Definitions";
        private const string ConHeaderPSets = "Psets";
        private const string ConHeaderSignals = "Signals";
        private const string ConHeaderPortMap = "Port Map";
        private const string ConHeaderFraction = "Fractional Bus";
        private const string ConHeaderConcurrent = "Concurrent Sequence";
        private const string ConHeaderComment = "Comment";
        private int _acSpecIndex = -1;
        private int _binTableIndex = -1;
        private int _characterizationIndex = -1;
        private int _commentIndex = -1;
        private int _concurrentSeqIndex = -1;
        private int _dcSpecIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _flowTableIndex = -1;
        private int _fractionalIndex = -1;
        private JobListSheet _jobListSheet;
        private int _jobNameIndex = -1;
        private int _mixSignalIndex = -1;
        private int _patternSetsIndex = -1;
        private int _pinMapIndex = -1;
        private int _portMapIndex = -1;
        private int _pSetsIndex = -1;
        private int _signalIndex = -1;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _testInsIndex = -1;
        private int _waveDefineIndex = -1;

        private const int StartRowIndex = 4;
        private const int StartColumnIndex = 2;

        public JobListSheet GetSheet(Stream stream, string sheetName)
        {
            var jobListSheet = new JobListSheet(sheetName);
            var isBackup = false;
            var i = 1;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (i > StartRowIndex)
                    {
                        var jobRow = GetJobRow(line, sheetName, i);
                        if (string.IsNullOrEmpty(jobRow.JobName))
                        {
                            isBackup = true;
                            continue;
                        }

                        jobRow.IsBackup = isBackup;
                        jobListSheet.AddRow(jobRow);
                    }
                    i++;
                }
            }
            return jobListSheet;
        }

        private JobRow GetJobRow(string line, string sheetName, int row)
        {
            var arr = line.Split('\t');
            var jobRow = new JobRow();
            jobRow.RowNum = row;
            jobRow.SheetName = sheetName;
            var index = StartColumnIndex - 1;
            var content = GetCellText(arr, 0);
            jobRow.ColumnA = content;
            content = GetCellText(arr, index);
            jobRow.JobName = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.PinMap = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.TestInstance = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.FlowTable = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.AcSpecs = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.DcSpecs = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.PatternSets = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.BinTable = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.Characterization = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.MixedSignalTiming = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.WaveDefinition = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.PSets = content;
            index++;
            jobRow.Signals = string.Empty;
            index++;
            content = GetCellText(arr, index);
            jobRow.PortMap = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.FractionalBus = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.ConcurrentSequence = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.SpikeCheckConfig = content;
            index++;
            content = GetCellText(arr, index);
            jobRow.Comment = content;
            return jobRow;
        }

        public List<string> GetJobs(Stream stream, string sheetName)
        {
            var jobs = new List<string>();
            var flag = false;
            using (var sr = new StreamReader(stream))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();
                    if (line != null)
                    {
                        var arr = line.Split(new[] { '\t' }, StringSplitOptions.None);
                        if (arr.Count() > 3)
                        {
                            var job = arr[1];
                            if (!string.IsNullOrEmpty(job) && flag)
                                jobs.Add(job);
                            if (job.Equals("Job Name", StringComparison.CurrentCultureIgnoreCase))
                                flag = true;
                        }
                    }
                }
            }

            return jobs.Distinct().ToList();
        }

        public JobListSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public JobListSheet GetSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _jobListSheet = new JobListSheet(worksheet);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _jobListSheet = ReadSheetData();

            return _jobListSheet;
        }

        private JobListSheet ReadSheetData()
        {
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = GetJobRow(i);
                _jobListSheet.AddRow(row);
            }

            return _jobListSheet;
        }

        private JobRow GetJobRow(int i)
        {
            var row = new JobRow();
            row.RowNum = i;
            if (_jobNameIndex != -1)
                row.JobName = _excelWorksheet.GetMergedCellValue(i, _jobNameIndex).Trim();
            if (_pinMapIndex != -1)
                row.PinMap = _excelWorksheet.GetMergedCellValue(i, _pinMapIndex).Trim();
            if (_testInsIndex != -1)
                row.TestInstance = _excelWorksheet.GetMergedCellValue(i, _testInsIndex).Trim();
            if (_flowTableIndex != -1)
                row.FlowTable = _excelWorksheet.GetMergedCellValue(i, _flowTableIndex).Trim();
            if (_acSpecIndex != -1)
                row.AcSpecs = _excelWorksheet.GetMergedCellValue(i, _acSpecIndex).Trim();
            if (_dcSpecIndex != -1)
                row.DcSpecs = _excelWorksheet.GetMergedCellValue(i, _dcSpecIndex).Trim();
            if (_patternSetsIndex != -1)
                row.PatternSets = _excelWorksheet.GetMergedCellValue(i, _patternSetsIndex).Trim();
            if (_binTableIndex != -1)
                row.BinTable = _excelWorksheet.GetMergedCellValue(i, _binTableIndex).Trim();
            if (_characterizationIndex != -1)
                row.Characterization = _excelWorksheet.GetMergedCellValue(i, _characterizationIndex).Trim();
            if (_mixSignalIndex != -1)
                row.MixedSignalTiming =
                    _excelWorksheet.GetMergedCellValue(i, _mixSignalIndex).Trim();
            if (_waveDefineIndex != -1)
                row.WaveDefinition =
                    _excelWorksheet.GetMergedCellValue(i, _waveDefineIndex).Trim();
            if (_pSetsIndex != -1)
                row.PSets = _excelWorksheet.GetMergedCellValue(i, _pSetsIndex).Trim();
            if (_signalIndex != -1)
                row.Signals = _excelWorksheet.GetMergedCellValue(i, _signalIndex).Trim();
            if (_portMapIndex != -1)
                row.PortMap = _excelWorksheet.GetMergedCellValue(i, _portMapIndex).Trim();
            if (_fractionalIndex != -1)
                row.FractionalBus = _excelWorksheet.GetMergedCellValue(i, _fractionalIndex).Trim();
            if (_concurrentSeqIndex != -1)
                row.ConcurrentSequence = _excelWorksheet.GetMergedCellValue(i, _concurrentSeqIndex)
                    .Trim();
            if (_commentIndex != -1)
                row.Comment = _excelWorksheet.GetMergedCellValue(i, _commentIndex).Trim();
            return row;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _jobNameIndex = -1;
            _pinMapIndex = -1;
            _testInsIndex = -1;
            _flowTableIndex = -1;
            _acSpecIndex = -1;
            _dcSpecIndex = -1;
            _patternSetsIndex = -1;
            _binTableIndex = -1;
            _characterizationIndex = -1;
            _mixSignalIndex = -1;
            _waveDefineIndex = -1;
            _pSetsIndex = -1;
            _signalIndex = -1;
            _portMapIndex = -1;
            _fractionalIndex = -1;
            _concurrentSeqIndex = -1;
            _commentIndex = -1;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = _excelWorksheet.GetCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderJobName, StringComparison.OrdinalIgnoreCase))
                {
                    _jobNameIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderJobName, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderPinMap, StringComparison.OrdinalIgnoreCase))
                {
                    _pinMapIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderPinMap, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderTestIns, StringComparison.OrdinalIgnoreCase))
                {
                    _testInsIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderTestIns, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderFlowTable, StringComparison.OrdinalIgnoreCase))
                {
                    _flowTableIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderFlowTable, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderAcSpec, StringComparison.OrdinalIgnoreCase) ||
                    lStrHeader.Equals(ConHeaderAcSpecs, StringComparison.OrdinalIgnoreCase))
                {
                    _acSpecIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderAcSpec, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderDcSpec, StringComparison.OrdinalIgnoreCase))
                {
                    _dcSpecIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderDcSpec, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderPatternSets, StringComparison.OrdinalIgnoreCase))
                {
                    _patternSetsIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderPatternSets, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderBinTable, StringComparison.OrdinalIgnoreCase))
                {
                    _binTableIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderBinTable, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderCharacterization, StringComparison.OrdinalIgnoreCase))
                {
                    _characterizationIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderCharacterization, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderMixSignalTiming, StringComparison.OrdinalIgnoreCase))
                {
                    _mixSignalIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderMixSignalTiming, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderWaveDef, StringComparison.OrdinalIgnoreCase))
                {
                    _waveDefineIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderWaveDef, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderPSets, StringComparison.OrdinalIgnoreCase))
                {
                    _pSetsIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderPSets, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderSignals, StringComparison.OrdinalIgnoreCase))
                {
                    _signalIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderSignals, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderPortMap, StringComparison.OrdinalIgnoreCase))
                {
                    _portMapIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderPortMap, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderFraction, StringComparison.OrdinalIgnoreCase))
                {
                    _fractionalIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderFraction, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderConcurrent, StringComparison.OrdinalIgnoreCase))
                {
                    _concurrentSeqIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderConcurrent, i);
                    continue;
                }

                if (lStrHeader.Equals(ConHeaderComment, StringComparison.OrdinalIgnoreCase))
                {
                    _commentIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderComment, i);
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            var rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            var colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (var i = 1; i <= rowNum; i++)
                for (var j = 1; j <= colNum; j++)
                    if (_excelWorksheet.GetCellValue(i, j).Trim()
                        .Equals(ConHeaderJobName, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
                    }

            return false;
        }

        private bool GetDimensions()
        {
            if (_excelWorksheet.Dimension != null)
            {
                _startColNumber = _excelWorksheet.Dimension.Start.Column;
                _startRowNumber = _excelWorksheet.Dimension.Start.Row;
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
                return true;
            }

            return false;
        }
    }
}