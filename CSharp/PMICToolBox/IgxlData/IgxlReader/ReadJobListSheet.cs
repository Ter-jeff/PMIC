using System;
using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace IgxlData.IgxlReader
{
    public class ReadJobListSheet : IgxlSheetReader
    {
        private ExcelWorksheet _excelWorksheet;
        private static JoblistSheet _jobListSheet;
        private string _sheetName;
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
        private const string ConHeaderPsets = "Psets";
        private const string ConHeaderSignals = "Signals";
        private const string ConHeaderPortMap = "Port Map";
        private const string ConHeaderFraction = "Fractional Bus";
        private const string ConHeaderConcurrent = "Concurrent Sequence";
        private const string ConHeaderComment = "Comment";

        private int _startColNumber = -1;
        private int _startRowNumber = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private int _jobnameIndex = -1;
        private int _pinmapIndex = -1;
        private int _testinsIndex = -1;
        private int _flowtabIndex = -1;
        private int _acspecIndex = -1;
        private int _dcspecIndex = -1;
        private int _patternsetsIndex = -1;
        private int _bintableIndex = -1;
        private int _characterizationIndex = -1;
        private int _mixsignalIndex = -1;
        private int _wavedefineIndex = -1;
        private int _psetsIndex = -1;
        private int _signalIndex = -1;
        private int _portmapIndex = -1;
        private int _fractionalIndex = -1;
        private int _concurrentseqIndex = -1;
        private int _commentIndex = -1;

        #region public Function
        public JoblistSheet GetSheet(Worksheet worksheet)
        {
            return GetSheet(ConvertWorksheetToExcelSheet(worksheet));
        }

        public JoblistSheet GetSheet(string fileName)
        {
            return GetSheet(ConvertTxtToExcelSheet(fileName));
        }

        public JoblistSheet GetSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _jobListSheet = new JoblistSheet(worksheet);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _jobListSheet = ReadSheetData();

            return _jobListSheet;
        }

        public static string GetDcSpec(string job)
        {
            JobRow rowdata = _jobListSheet.GetRow(job);
            return rowdata.DcSpecs;
        }
        #endregion

        #region Private Function
        private JoblistSheet ReadSheetData()
        {
            for (int i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                JobRow row = new JobRow();
                row.LineNum = i.ToString();
                if (_jobnameIndex != -1)
                    row.JobName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _jobnameIndex).Trim();
                if (_pinmapIndex != -1)
                    row.PinMap = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _pinmapIndex).Trim();
                if (_testinsIndex != -1)
                    row.TestInstances = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _testinsIndex).Trim();
                if (_flowtabIndex != -1)
                    row.FlowTable = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _flowtabIndex).Trim();
                if (_acspecIndex != -1)
                    row.AcSpecs = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _acspecIndex).Trim();
                if (_dcspecIndex != -1)
                    row.DcSpecs = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _dcspecIndex).Trim();
                if (_patternsetsIndex != -1)
                    row.PatternSets = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _patternsetsIndex).Trim();
                if (_bintableIndex != -1)
                    row.BinTable = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _bintableIndex).Trim();
                if (_characterizationIndex != -1)
                    row.Characterization = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _characterizationIndex).Trim();
                if (_mixsignalIndex != -1)
                    row.MixedSignalTiming = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _mixsignalIndex).Trim();
                if (_wavedefineIndex != -1)
                    row.WaveDefinitions = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _wavedefineIndex).Trim();
                if (_psetsIndex != -1)
                    row.Psets = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _psetsIndex).Trim();
                if (_signalIndex != -1)
                    row.Signals = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _signalIndex).Trim();
                if (_portmapIndex != -1)
                    row.PortMap = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _portmapIndex).Trim();
                if (_fractionalIndex != -1)
                    row.FractionalBus = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _fractionalIndex).Trim();
                if (_concurrentseqIndex != -1)
                    row.ConcurrentSequence = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _concurrentseqIndex).Trim();
                if (_commentIndex != -1)
                    row.Comment = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _commentIndex).Trim();
                _jobListSheet.AddRow(row);
            }
            return _jobListSheet;
        }

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _jobnameIndex = -1;
            _pinmapIndex = -1;
            _testinsIndex = -1;
            _flowtabIndex = -1;
            _acspecIndex = -1;
            _dcspecIndex = -1;
            _patternsetsIndex = -1;
            _bintableIndex = -1;
            _characterizationIndex = -1;
            _mixsignalIndex = -1;
            _wavedefineIndex = -1;
            _psetsIndex = -1;
            _signalIndex = -1;
            _portmapIndex = -1;
            _fractionalIndex = -1;
            _concurrentseqIndex = -1;
            _commentIndex = -1;
        }

        private bool GetHeaderIndex()
        {
            for (int i = _startColNumber; i <= _endColNumber; i++)
            {
                string lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
                if (lStrHeader.Equals(ConHeaderJobName, StringComparison.OrdinalIgnoreCase))
                {
                    _jobnameIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderJobName, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderPinMap, StringComparison.OrdinalIgnoreCase))
                {
                    _pinmapIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderPinMap, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderTestIns, StringComparison.OrdinalIgnoreCase))
                {
                    _testinsIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderTestIns, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderFlowTable, StringComparison.OrdinalIgnoreCase))
                {
                    _flowtabIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderFlowTable, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderAcSpec, StringComparison.OrdinalIgnoreCase)||
                    lStrHeader.Equals(ConHeaderAcSpecs, StringComparison.OrdinalIgnoreCase))
                {
                    _acspecIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderAcSpec, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderDcSpec, StringComparison.OrdinalIgnoreCase))
                {
                    _dcspecIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderDcSpec, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderPatternSets, StringComparison.OrdinalIgnoreCase))
                {
                    _patternsetsIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderPatternSets, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderBinTable, StringComparison.OrdinalIgnoreCase))
                {
                    _bintableIndex = i;
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
                    _mixsignalIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderMixSignalTiming, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderWaveDef, StringComparison.OrdinalIgnoreCase))
                {
                    _wavedefineIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderWaveDef, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderPsets, StringComparison.OrdinalIgnoreCase))
                {
                    _psetsIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderPsets, i);
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
                    _portmapIndex = i;
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
                    _concurrentseqIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderConcurrent, i);
                    continue;
                }
                if (lStrHeader.Equals(ConHeaderComment, StringComparison.OrdinalIgnoreCase))
                {
                    _commentIndex = i;
                    _jobListSheet.HeaderIndex.Add(ConHeaderComment, i);
                    continue;
                }
            }

            return true;
        }

        private bool GetFirstHeaderPosition()
        {
            int rowNum = _endRowNumber > 10 ? 10 : _endRowNumber;
            int colNum = _endColNumber > 10 ? 10 : _endColNumber;
            for (int i = 1; i <= rowNum; i++)
                for (int j = 1; j <= colNum; j++)
                {
                    if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim().Equals(ConHeaderJobName, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRowNumber = i;
                        return true;
                    }
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
        #endregion
    }
}