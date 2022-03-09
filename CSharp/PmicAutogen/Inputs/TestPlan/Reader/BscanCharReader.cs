using AutomationCommon.Utility;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
            GPIOx_SUPPLY = "";
            GPIOx_DRV = "";
            GPIOx_IOCONFIG = "";
            GPIOx_PU_PD = "";
        }

        #endregion

        #region Property

        public string SourceSheetName { get; set; }
        public int RowNum { get; set; }
        public string BscanPatternName { get; set; }
        public string LevelDomain { get; set; }
        public decimal LevelIOHAndIOL { get; set; }
        public string GPIOx_SUPPLY { get; set; }
        public string GPIOx_DRV { get; set; }
        public string GPIOx_IOCONFIG { get; set; }
        public string GPIOx_PU_PD { get; set; }

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
        public int LevelIOHAndIOLIndex = -1;
        public int GPIOx_SUPPLY_Index = -1;
        public int GPIOx_DRV_Index = -1;
        public int GPIOx_IOCONFIG_Index = -1;
        public int GPIOx_PU_PD_Index = -1;
        private Dictionary<string, List<decimal>> _domainCurrentDic;
        #endregion

        #region public method
        public Dictionary<string,List<decimal>> GetDomainCurrentMapping()
        {
            if (_domainCurrentDic.Count > 0)
                return _domainCurrentDic;

            foreach(var bscanCharRow in Rows)
            {
                if (bscanCharRow.LevelIOHAndIOL == 0)
                    continue;

                if(_domainCurrentDic.ContainsKey(bscanCharRow.LevelDomain) 
                    && !_domainCurrentDic[bscanCharRow.LevelDomain].Contains(bscanCharRow.LevelIOHAndIOL))
                {
                    _domainCurrentDic[bscanCharRow.LevelDomain].Add(bscanCharRow.LevelIOHAndIOL);
                }
                else if(!_domainCurrentDic.ContainsKey(bscanCharRow.LevelDomain))
                {
                    _domainCurrentDic.Add(bscanCharRow.LevelDomain, new List<decimal>() { bscanCharRow.LevelIOHAndIOL });
                }
            }

            return _domainCurrentDic;
        }


        public List<decimal> GetDomainCurrents()
        {
            var currents= Rows.Where(o => o.LevelIOHAndIOL > 0).Select(o => o.LevelIOHAndIOL).Distinct().ToList();
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
        private const string HeaderLevelIOHAndIOL = "IOH&IOL (mA)";
        private const string HeaderGPIOx_SUPPLY = "GPIOx_SUPPLY";
        private const string HeaderGPIOx_DRV = "GPIOx_DRV";
        private const string HeaderGPIOx_IOCONFIG = "GPIOx_IOCONFIG";
        private const string HeaderGPIOx_PU_PD = "GPIOx_PU_PD";
        private int _endColNumber = -1;
        private int _endRowNumber = -1;
        private ExcelWorksheet _excelWorksheet;
        private BscanCharSheet _bscanCharSheet;      
        private string _sheetName;

        private int _startColNumber = -1;
        private int _startRowNumber = -1;

        private int _bscanPatternNameIndex = -1;
        private int _levelDomainIndex = -1;
        private int _levelIOHAndIOLIndex = -1;
        private int _GPIOx_SUPPLYIndex = -1;
        private int _GPIOx_DRVIndex = -1;
        private int _GPIOx_IOCONFIGIndex = -1;
        private int _GPIOx_PU_PDIndex = -1;

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
                    row.BscanPatternName = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _bscanPatternNameIndex).Trim();
                if (_levelDomainIndex != -1)
                    row.LevelDomain = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _levelDomainIndex).Trim();
                if (_levelIOHAndIOLIndex != -1)
                {
                    decimal levelIOHAndIOL = 0;
                    decimal.TryParse(EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _levelIOHAndIOLIndex).Trim(), out levelIOHAndIOL);
                    row.LevelIOHAndIOL = levelIOHAndIOL;
                }
                if (_GPIOx_SUPPLYIndex != -1)
                    row.GPIOx_SUPPLY = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _GPIOx_SUPPLYIndex).Trim();
                if (_GPIOx_DRVIndex != -1)
                    row.GPIOx_DRV = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _GPIOx_DRVIndex).Trim();
                if (_GPIOx_IOCONFIGIndex != -1)
                    row.GPIOx_IOCONFIG = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _GPIOx_IOCONFIGIndex).Trim();
                if (_GPIOx_PU_PDIndex != -1)
                    row.GPIOx_PU_PD = EpplusOperation.GetMergedCellValue(_excelWorksheet, i, _GPIOx_PU_PDIndex).Trim();
                bscanCharSheet.Rows.Add(row);
            }

            bscanCharSheet.BscanPatternNameIndex = _bscanPatternNameIndex;
            bscanCharSheet.LevelDomainIndex = _levelDomainIndex;
            bscanCharSheet.LevelIOHAndIOLIndex = _levelIOHAndIOLIndex;
            bscanCharSheet.GPIOx_SUPPLY_Index = _GPIOx_SUPPLYIndex;
            bscanCharSheet.GPIOx_DRV_Index = _GPIOx_DRVIndex;
            bscanCharSheet.GPIOx_IOCONFIG_Index = _GPIOx_IOCONFIGIndex;
            bscanCharSheet.GPIOx_PU_PD_Index = _GPIOx_PU_PDIndex;

            return bscanCharSheet;
        }

        private bool GetHeaderIndex()
        {
            for (var i = _startColNumber; i <= _endColNumber; i++)
            {
                var lStrHeader = EpplusOperation.GetCellValue(_excelWorksheet, _startRowNumber, i).Trim();
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

                if (lStrHeader.Equals(HeaderLevelIOHAndIOL, StringComparison.OrdinalIgnoreCase))
                {
                    _levelIOHAndIOLIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderGPIOx_SUPPLY, StringComparison.OrdinalIgnoreCase))
                {
                    _GPIOx_SUPPLYIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderGPIOx_DRV, StringComparison.OrdinalIgnoreCase))
                {
                    _GPIOx_DRVIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderGPIOx_IOCONFIG, StringComparison.OrdinalIgnoreCase))
                {
                    _GPIOx_IOCONFIGIndex = i;
                    continue;
                }

                if (lStrHeader.Equals(HeaderGPIOx_PU_PD, StringComparison.OrdinalIgnoreCase))
                {
                    _GPIOx_PU_PDIndex = i;
                    continue;
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
                    if (EpplusOperation.GetCellValue(_excelWorksheet, i, j).Trim()
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
            _levelIOHAndIOLIndex = -1;
            _GPIOx_SUPPLYIndex = -1;
            _GPIOx_DRVIndex = -1;
            _GPIOx_IOCONFIGIndex = -1;
            _GPIOx_PU_PDIndex = -1;
        }
    }
}
