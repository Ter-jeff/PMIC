using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.ScghFile.Reader
{
    public class ProdCharSheetReader
    {
        #region Constructor

        public ProdCharSheetReader()
        {
            _startColumnNumber = 1;
            _startRowNumber = 1;
            _endColNumber = 1;
            _endRowNumber = 1;
            _initCnt = 0;
            _initStartColNumber = 0;
            _payloadStartColNumber = 0;
            _supplyVoltageColNumber = 0;
            _commentsColNumber = 0;
            _sVoltageColNumber = 0;
            _pVoltageColNumber = 0;
            _enableWordColumnNumber = -1;
            _levelHVorLv = -1;
            _usageColNumber = -1;
            _applicationColNumber = -1;
        }

        #endregion

        public ProdCharSheet ReadScghSheet(ExcelWorksheet sheetScgh)
        {
            _excelWorksheet = sheetScgh;
            _sheetName = sheetScgh.Name;

            ReadHeader();

            var sheetScgScan = ReadAllData();

            sheetScgScan.SheetName = sheetScgh.Name;
            return sheetScgScan;
        }

        private void ReadHeader()
        {
            ResetValue();

            GetDimensions();

            GetFirstHeaderPosition();

            for (var i = _startColumnNumber; i <= _endColNumber; i++)
            {
                var header = GetCellValue(_startRowNumber, i).ToUpper().Trim();

                if (IsLiked(header, HeaderBlock))
                {
                    _blockColNumber = i;
                    continue;
                }

                if (IsLiked(header, HeaderMode))
                {
                    _modeColNumber = i;
                    continue;
                }

                if (IsLiked(header, HeaderItem))
                {
                    _itemColNumber = i;
                    continue;
                }

                if (IsLiked(header, HeaderApplication))
                {
                    _applicationColNumber = i;
                    continue;
                }

                if (IsLiked(header, HeaderSupplyVoltage))
                {
                    _supplyVoltageColNumber = i;
                    continue;
                }

                if (Regex.IsMatch(header, HeaderPayload, RegexOptions.IgnoreCase))
                {
                    if (_payloadStartColNumber == 0) _payloadStartColNumber = i;
                    _payloadCnt++;
                    continue;
                }

                if (Regex.IsMatch(header, HeaderInit, RegexOptions.IgnoreCase))
                {
                    if (_initStartColNumber == 0) _initStartColNumber = i;
                    _initCnt++;
                    continue;
                }

                if (IsLiked(header, HeaderEnable))
                {
                    _enableWordColumnNumber = i;
                    continue;
                }

                if (IsLiked(header, HeaderLevelHVorLv))
                {
                    _levelHVorLv = i;
                    continue;
                }

                if (IsLiked(header, HeaderUsage))
                {
                    _usageColNumber = i;
                    continue;
                }

                if (IsLiked(header, HeaderComment))
                {
                    _commentsColNumber = i;
                    continue;
                }

                if (IsLiked(header, PeripheralVoltage))
                {
                    _pVoltageColNumber = i;
                    continue;
                }

                if (IsLiked(header, SramVoltage)) _sVoltageColNumber = i;
            }
        }

        private void ResetValue()
        {
            _startColumnNumber = 1;
            _startRowNumber = 1;
            _endColNumber = 1;
            _endRowNumber = 1;
            _initCnt = 0;
            _initStartColNumber = 0;

            _enableWordColumnNumber = -1;
            _levelHVorLv = -1;

            _usageColNumber = -1;
            _applicationColNumber = -1;
        }

        private void GetFirstHeaderPosition()
        {
            for (var i = 1; i < 10; i++)
                for (var j = 1; j < 10; j++)
                    if (GetCellValue(i, j).Trim().ToUpper() == HeaderMode)
                    {
                        _startRowNumber = i;
                        break;
                    }
        }

        private void GetDimensions()
        {
            if (_excelWorksheet != null)
            {
                _endColNumber = _excelWorksheet.Dimension.End.Column;
                _endRowNumber = _excelWorksheet.Dimension.End.Row;
            }
        }

        private List<string> ReadInitList(int rowNumber)
        {
            var initList = new List<string>();
            for (var i = 0; i < _initCnt; i++)
                initList.Add(_excelWorksheet.GetMergedCellValue(rowNumber, _initStartColNumber + i).Trim());
            return initList;
        }

        private List<string> ReadPayloadList(int rowNumber)
        {
            var payloadList = new List<string>();
            for (var i = 0; i < _payloadCnt; i++)
            {
                var payload = _excelWorksheet.GetMergedCellValue(rowNumber, _payloadStartColNumber + i).Trim();
                if (payload != "" && payload != "NA")
                    payloadList.Add(payload);
            }

            return payloadList;
        }

        private ProdCharSheet ReadAllData()
        {
            var sheetScg = new ProdCharSheet();
            for (var i = _startRowNumber + 1; i <= _endRowNumber; i++)
            {
                var row = new ProdCharSheetRow(_sheetName);
                if (_payloadStartColNumber == 0)
                {
                }

                if (GetCellValue(i, _payloadStartColNumber) != "")
                {
                    row.RowNum = i;
                    row.Block = _blockColNumber == 0
                        ? ""
                        : _excelWorksheet.GetMergedCellValue(i, _blockColNumber).Trim();
                    row.Mode = _excelWorksheet.GetMergedCellValue(i, _modeColNumber).Trim();
                    row.Item = _excelWorksheet.GetMergedCellValue(i, _itemColNumber).Trim();

                    if (_payloadStartColNumber != 0)
                    {
                        var payloads = ReadPayloadList(i);
                        row.PayloadList.AddRange(payloads);
                        row.PayloadAliasList.AddRange(payloads);
                    }

                    if (_initStartColNumber != 0)
                    {
                        var inits = ReadInitList(i);
                        row.InitList.AddRange(inits);
                        row.InitAliasList.AddRange(inits);
                    }

                    row.Application = _applicationColNumber != -1
                        ? _excelWorksheet.GetMergedCellValue(i, _applicationColNumber).Trim()
                        : GetApplicationField(row.PayloadValue);

                    row.Usage = _usageColNumber != -1 ? GetCellValue(i, _usageColNumber).Trim() : "1";

                    if (_supplyVoltageColNumber != 0)
                    {
                        row.SupplyVoltage = GetCellValue(i, _supplyVoltageColNumber).Trim();
                        row.SupplyVoltage = Regex.Replace(row.SupplyVoltage, @"\n", "");
                    }

                    if (_commentsColNumber != 0) row.Comments = GetCellValue(i, _commentsColNumber).Trim();

                    if (_sVoltageColNumber != 0 && _pVoltageColNumber != 0)
                    {
                        row.SramVoltage = GetCellValue(i, _sVoltageColNumber).Trim();
                        row.SramVoltage = Regex.Replace(row.SramVoltage, @"\n", "");
                        row.SramVoltage = row.SramVoltage.Equals("N/A", StringComparison.OrdinalIgnoreCase)
                            ? ""
                            : row.SramVoltage;
                        row.SramVoltage = row.SramVoltage.Equals("NA", StringComparison.OrdinalIgnoreCase)
                            ? ""
                            : row.SramVoltage;

                        row.PeripheralVoltage = GetCellValue(i, _pVoltageColNumber).Trim();
                        row.PeripheralVoltage = Regex.Replace(row.PeripheralVoltage, @"\n", "");
                        row.PeripheralVoltage = row.PeripheralVoltage.Equals("N/A", StringComparison.OrdinalIgnoreCase)
                            ? ""
                            : row.PeripheralVoltage;
                        row.PeripheralVoltage = row.PeripheralVoltage.Equals("NA", StringComparison.OrdinalIgnoreCase)
                            ? ""
                            : row.PeripheralVoltage;
                    }

                    if (_enableWordColumnNumber != -1) row.EnableWord = GetCellValue(i, _enableWordColumnNumber).Trim();

                    if (_levelHVorLv != -1) row.LevelHVorLv = GetCellValue(i, _levelHVorLv).Trim();

                    sheetScg.RowList.Add(row);
                }
            }

            return sheetScg;
        }

        private string GetApplicationField(string payLoadName)
        {
            var appName = "";

            if (payLoadName == "")
                return appName;

            var payloadTok = payLoadName.ToUpper().Split(new[] { '_' }, StringSplitOptions.RemoveEmptyEntries);
            if (payloadTok.Length != 0)
            {
                if (payloadTok[0] == "PP")
                    appName = "Production";
                else if (payloadTok[0] == "CZ")
                    appName = "Characterization";
                else if (payloadTok[0] == "DD")
                    appName = "Debug";
                else if (payloadTok[0] == "HT")
                    appName = "HTOL";
                else
                    appName = payloadTok[0];
            }

            return appName;
        }

        private string GetCellValue(int rowNumber, int columnNumber)
        {
            var value = _excelWorksheet.Cells[rowNumber, columnNumber].Value;
            if (value != null) return value.ToString();

            return "";
        }


        private string ReplaceDoubleBlank(string pString)
        {
            var lStrResult = pString;
            do
            {
                lStrResult = lStrResult.Replace("  ", " ");
            } while (lStrResult.IndexOf("  ", StringComparison.Ordinal) >= 0);

            return lStrResult;
        }

        private string FormatStringForCompare(string text)
        {
            var result = text.Trim();

            result = ReplaceDoubleBlank(result);

            result = result.Replace(" ", "_");

            result = result.ToUpper();

            return result;
        }

        private bool IsLiked(string pStrInput, string pStrPatten)
        {
            if (pStrPatten.IndexOf(@".*", StringComparison.Ordinal) >= 0 ||
                pStrPatten.IndexOf(@".+", StringComparison.Ordinal) >= 0)
            {
                var value = Regex.IsMatch(FormatStringForCompare(pStrInput), FormatStringForCompare(pStrPatten));
                return value;
            }
            else
            {
                var value = FormatStringForCompare(pStrInput) == FormatStringForCompare(pStrPatten);
                return value;
            }
        }

        #region Field

        private int _blockColNumber;
        private int _modeColNumber;
        private int _itemColNumber;
        private int _applicationColNumber;
        private int _initStartColNumber;
        private int _payloadStartColNumber;
        private int _usageColNumber;
        private int _supplyVoltageColNumber;
        private int _commentsColNumber;
        private int _sVoltageColNumber;
        private int _pVoltageColNumber;

        private const string HeaderBlock = "Block";
        private const string HeaderMode = "MODE";
        private const string HeaderItem = "ITEM";
        private const string HeaderApplication = "APPLICATION";
        private const string HeaderPayload = @"PAYLOAD\d*";
        private const string HeaderInit = @"INIT\d+";
        private const string HeaderUsage = "USAGE.*";
        private const string HeaderComment = "COMMENT.*";
        private const string HeaderSupplyVoltage = "SUPPLY VOLTAGE";
        private const string PeripheralVoltage = "Peripheral Voltage";
        private const string SramVoltage = "SRAM Voltage";
        private const string HeaderEnable = "Enable";
        private const string HeaderLevelHVorLv = "Level.*";

        private ExcelWorksheet _excelWorksheet;
        private int _startColumnNumber;
        private int _startRowNumber;
        private int _endColNumber;
        private int _endRowNumber;
        private int _initCnt;
        private int _payloadCnt;
        private string _sheetName;
        private int _enableWordColumnNumber;
        private int _levelHVorLv;

        #endregion
    }
}