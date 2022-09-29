using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class BscanCharRow
    {
        #region Constructor

        public BscanCharRow(string sourceSheetName)
        {
            SourceSheetName = sourceSheetName;
            BscanPatternName = "";
            LevelDomain = "";
            GpiOxSupply = "";
            GpiOxDrv = "";
            GpiOxIoconfig = "";
            GpiOxPuPd = "";
        }

        #endregion

        #region Property

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string BscanPatternName { get; set; }
        public string LevelDomain { get; set; }
        public decimal LevelIohAndIol { get; set; }
        public string GpiOxSupply { get; set; }
        public string GpiOxDrv { get; set; }
        public string GpiOxIoconfig { get; set; }
        public string GpiOxPuPd { get; set; }

        #endregion
    }

    public class BscanCharSheet
    {
        #region Constructor

        public BscanCharSheet(string sheetName)
        {
            SheetName = sheetName;
            Rows = new List<BscanCharRow>();
            _domainCurrentDic = new Dictionary<string, List<decimal>>();
        }

        #endregion

        #region Property

        public string SheetName { get; set; }
        public List<BscanCharRow> Rows { get; set; }
        public int BscanPatternNameIndex = -1;
        public int LevelDomainIndex = -1;
        public int LevelIohAndIolIndex = -1;
        public int GpiOxSupplyIndex = -1;
        public int GpiOxDrvIndex = -1;
        public int GpiOxIoconfigIndex = -1;
        public int GpiOxPuPdIndex = -1;
        private readonly Dictionary<string, List<decimal>> _domainCurrentDic;

        #endregion

        #region public method

        public Dictionary<string, List<decimal>> GetDomainCurrentMapping()
        {
            if (_domainCurrentDic.Count > 0)
                return _domainCurrentDic;

            foreach (var bscanCharRow in Rows)
            {
                if (bscanCharRow.LevelIohAndIol == 0)
                    continue;

                if (_domainCurrentDic.ContainsKey(bscanCharRow.LevelDomain)
                    && !_domainCurrentDic[bscanCharRow.LevelDomain].Contains(bscanCharRow.LevelIohAndIol))
                    _domainCurrentDic[bscanCharRow.LevelDomain].Add(bscanCharRow.LevelIohAndIol);
                else if (!_domainCurrentDic.ContainsKey(bscanCharRow.LevelDomain))
                    _domainCurrentDic.Add(bscanCharRow.LevelDomain, new List<decimal> { bscanCharRow.LevelIohAndIol });
            }

            return _domainCurrentDic;
        }


        public List<decimal> GetDomainCurrents()
        {
            var currents = Rows.Where(o => o.LevelIohAndIol > 0).Select(o => o.LevelIohAndIol).Distinct().ToList();
            //Descending sort
            currents.Sort((a, b) => b.CompareTo(a));
            return currents;
        }

        #endregion
    }

    public class BscanCharReader
    {
        private const string HeaderBscanPatternName = "pattern name";
        private const string HeaderLevelDomain = "Domain";
        private const string HeaderLevelIohAndIol = "IOH&IOL (mA)";
        private const string HeaderGpiOxSupply = "GPIOx_SUPPLY";
        private const string HeaderGpiOxDrv = "GPIOx_DRV";
        private const string HeaderGpiOxIoconfig = "GPIOx_IOCONFIG";
        private const string HeaderGpiOxPuPd = "GPIOx_PU_PD";
        private BscanCharSheet _bscanCharSheet;

        private int _bscanPatternNameIndex = -1;
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private int _gpiOxDrvIndex = -1;
        private int _gpiOxIoconfigIndex = -1;
        private int _gpiOxPuPdIndex = -1;
        private int _gpiOxSupplyIndex = -1;
        private int _levelDomainIndex = -1;
        private int _levelIohAndIolIndex = -1;
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;

        public BscanCharSheet ReadSheet(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return null;

            _excelWorksheet = worksheet;

            _sheetName = worksheet.Name;

            _bscanCharSheet = new BscanCharSheet(_sheetName);

            Reset();

            if (!GetDimensions()) return null;

            if (!GetFirstHeaderPosition()) return null;

            if (!GetHeaderIndex()) return null;

            _bscanCharSheet = ReadSheetData();

            return _bscanCharSheet;
        }

        private BscanCharSheet ReadSheetData()
        {
            var bscanCharSheet = new BscanCharSheet(_sheetName);
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new BscanCharRow(_sheetName);
                row.RowNum = i;
                if (_bscanPatternNameIndex != -1)
                    row.BscanPatternName = _excelWorksheet.GetMergedCellValue(i, _bscanPatternNameIndex).Trim();
                if (_levelDomainIndex != -1)
                    row.LevelDomain = _excelWorksheet.GetMergedCellValue(i, _levelDomainIndex).Trim();
                if (_levelIohAndIolIndex != -1)
                {
                    decimal levelIohAndIol = 0;
                    decimal.TryParse(
                        _excelWorksheet.GetMergedCellValue(i, _levelIohAndIolIndex).Trim(),
                        out levelIohAndIol);
                    row.LevelIohAndIol = levelIohAndIol;
                }

                if (_gpiOxSupplyIndex != -1)
                    row.GpiOxSupply = _excelWorksheet.GetMergedCellValue(i, _gpiOxSupplyIndex).Trim();
                if (_gpiOxDrvIndex != -1)
                    row.GpiOxDrv = _excelWorksheet.GetMergedCellValue(i, _gpiOxDrvIndex).Trim();
                if (_gpiOxIoconfigIndex != -1)
                    row.GpiOxIoconfig = _excelWorksheet.GetMergedCellValue(i, _gpiOxIoconfigIndex)
                        .Trim();
                if (_gpiOxPuPdIndex != -1)
                    row.GpiOxPuPd = _excelWorksheet.GetMergedCellValue(i, _gpiOxPuPdIndex).Trim();
                bscanCharSheet.Rows.Add(row);
            }

            bscanCharSheet.BscanPatternNameIndex = _bscanPatternNameIndex;
            bscanCharSheet.LevelDomainIndex = _levelDomainIndex;
            bscanCharSheet.LevelIohAndIolIndex = _levelIohAndIolIndex;
            bscanCharSheet.GpiOxSupplyIndex = _gpiOxSupplyIndex;
            bscanCharSheet.GpiOxDrvIndex = _gpiOxDrvIndex;
            bscanCharSheet.GpiOxIoconfigIndex = _gpiOxIoconfigIndex;
            bscanCharSheet.GpiOxPuPdIndex = _gpiOxPuPdIndex;

            return bscanCharSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = _excelWorksheet.GetCellValue(_startRowNumber, i).Trim();
                if (lStrHeader.Equals(HeaderBscanPatternName, StringComparison.OrdinalIgnoreCase))
                {
                    _bscanPatternNameIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderLevelDomain, StringComparison.OrdinalIgnoreCase))
                {
                    _levelDomainIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderLevelIohAndIol, StringComparison.OrdinalIgnoreCase))
                {
                    _levelIohAndIolIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderGpiOxSupply, StringComparison.OrdinalIgnoreCase))
                {
                    _gpiOxSupplyIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderGpiOxDrv, StringComparison.OrdinalIgnoreCase))
                {
                    _gpiOxDrvIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderGpiOxIoconfig, StringComparison.OrdinalIgnoreCase))
                {
                    _gpiOxIoconfigIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderGpiOxPuPd, StringComparison.OrdinalIgnoreCase)) _gpiOxPuPdIndex = i;
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
                        .Equals(HeaderBscanPatternName, StringComparison.OrdinalIgnoreCase))
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

        private void Reset()
        {
            _startColNumber = -1;
            _startRowNumber = -1;
            _endColNumber = -1;
            _endRowNumber = -1;
            _bscanPatternNameIndex = -1;
            _levelDomainIndex = -1;
            _levelIohAndIolIndex = -1;
            _gpiOxSupplyIndex = -1;
            _gpiOxDrvIndex = -1;
            _gpiOxIoconfigIndex = -1;
            _gpiOxPuPdIndex = -1;
        }
    }
}