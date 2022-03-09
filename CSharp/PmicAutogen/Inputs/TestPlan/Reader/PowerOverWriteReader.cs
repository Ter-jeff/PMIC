using System;
using System.Collections.Generic;
using AutomationCommon.Utility;
using IgxlData.IgxlSheets;
using IgxlData.Others;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDc.PowerOverWrite;

namespace PmicAutogen.Inputs.TestPlan.Reader
{
    public class PowerOverWriteSheet
    {
        public PowerOverWriteSheet()
        {
            PowerOverWrite = new List<PowerOverWrite>();
        }

        public List<PowerOverWrite> PowerOverWrite { set; get; }


        public void SetPowerOverWrite(DcSpecSheet dcSpecSheet)
        {
            foreach (var catDef in PowerOverWrite)
            {
                var index = dcSpecSheet.FindCategoryIndex(catDef.CategoryName);
                if (index != -1)
                    foreach (var row in catDef.DataRows)
                    {
                        var dcSpec = dcSpecSheet.FindDcSpecs(SpecFormat.GenDcSpecSymbol(row.PinName));
                        if (dcSpec != null)
                            if (!string.IsNullOrEmpty(row.Nv))
                            {
                                if (!string.IsNullOrEmpty(row.HvRatio))
                                    dcSpec.CategoryList[index].Max = "=" + row.Nv + "*" + row.HvRatio;
                                if (!string.IsNullOrEmpty(row.LvRatio))
                                    dcSpec.CategoryList[index].Min = "=" + row.Nv + "*" + row.LvRatio;
                                dcSpec.CategoryList[index].Typ = "=" + row.Nv;
                            }
                    }
            }
        }
    }

    public class PowerOverWriteReader
    {
        #region Constuctor

        public PowerOverWriteReader()
        {
            _startColumn = 0;
            _startRow = 0;
            _headers = new Dictionary<string, int>();
        }

        #endregion

        #region Read Flow

        public PowerOverWriteSheet ReadFlowMain(ExcelWorksheet worksheet)
        {
            _inputWorksheet = worksheet;

            ReadHeader();

            return ReadData();
        }

        #endregion

        #region Filed

        private ExcelWorksheet _inputWorksheet;
        private const int MaxSearchColumn = 10;
        private const int MaxSearchRow = 10;
        public const string HeaderNv = "NV";
        public const string HeaderNvValt = "NV(Valt)";
        public const string HeaderHvRatio = "HV_Ratio";
        public const string HeaderLvRatio = "LV_Ratio";
        public const string HeaderIfold = "Ifold";
        public const string HeaderVil = "Vil";
        public const string HeaderVih = "Vih";
        public const string HeaderVol = "Vol";
        public const string HeaderVoh = "Voh";
        public const string HeaderIol = "Iol";
        public const string HeaderIoh = "Ioh";
        public const string HeaderVt = "Vt";
        public const string HeaderVcl = "Vcl";
        public const string HeaderVch = "Vch";
        public const string HeaderDriveMode = "DriverMode";
        public const string HeaderVicm = "Vicm";
        public const string HeaderVid = "Vid";
        public const string HeaderVod = "Vod";

        private int _startColumn;
        private int _startRow;
        private readonly Dictionary<string, int> _headers;

        #endregion

        #region Member function

        private void ReadHeader()
        {
            string cellContent;
            var hasFind = false;
            //locate the first header
            for (var i = 1; i <= MaxSearchRow; i++)
            {
                for (var j = 1; j <= MaxSearchColumn; j++)
                {
                    cellContent = EpplusOperation.GetCellValue(_inputWorksheet, i, j);
                    if (cellContent.Equals(HeaderNv, StringComparison.OrdinalIgnoreCase))
                    {
                        _startRow = i;
                        _startColumn = j;
                        hasFind = true;
                        break;
                    }
                }

                if (hasFind) break;
            }

            //Get All header columns
            for (var i = _startColumn; i <= _inputWorksheet.Dimension.End.Column; i++)
            {
                cellContent = EpplusOperation.GetCellValue(_inputWorksheet, _startRow, i);
                if (cellContent.Equals("")) break;
                _headers.Add(cellContent, i);
            }
        }

        private PowerOverWriteSheet ReadData()
        {
            var powerOverWriteSheet = new PowerOverWriteSheet();
            var nvColumn = _headers.ContainsKey(HeaderNv) ? _headers[HeaderNv] : -1;
            var nvValtColumn = _headers.ContainsKey(HeaderNvValt) ? _headers[HeaderNvValt] : -1;
            var hvRatioColumn = _headers.ContainsKey(HeaderHvRatio) ? _headers[HeaderHvRatio] : -1;
            var lvRatioColumn = _headers.ContainsKey(HeaderLvRatio) ? _headers[HeaderLvRatio] : -1;
            var ifoldColumn = _headers.ContainsKey(HeaderIfold) ? _headers[HeaderIfold] : -1;
            var vilColumn = _headers.ContainsKey(HeaderVil) ? _headers[HeaderVil] : -1;
            var vihColumn = _headers.ContainsKey(HeaderVih) ? _headers[HeaderVih] : -1;
            var volColumn = _headers.ContainsKey(HeaderVol) ? _headers[HeaderVol] : -1;
            var vohColumn = _headers.ContainsKey(HeaderVoh) ? _headers[HeaderVoh] : -1;
            var iolColumn = _headers.ContainsKey(HeaderIol) ? _headers[HeaderIol] : -1;
            var iohColumn = _headers.ContainsKey(HeaderIoh) ? _headers[HeaderIoh] : -1;
            var vtColumn = _headers.ContainsKey(HeaderVt) ? _headers[HeaderVt] : -1;
            var vclColumn = _headers.ContainsKey(HeaderVcl) ? _headers[HeaderVcl] : -1;
            var vchColumn = _headers.ContainsKey(HeaderVch) ? _headers[HeaderVch] : -1;
            var driverModeColumn = _headers.ContainsKey(HeaderDriveMode) ? _headers[HeaderDriveMode] : -1;
            var vicmColumn = _headers.ContainsKey(HeaderVicm) ? _headers[HeaderVicm] : -1;
            var vidColumn = _headers.ContainsKey(HeaderVid) ? _headers[HeaderVid] : -1;
            var vodColumn = _headers.ContainsKey(HeaderVod) ? _headers[HeaderVod] : -1;

            var categoryName = EpplusOperation.GetCellValue(_inputWorksheet, _startRow, nvColumn - 1);
            var powerOverWrite = new PowerOverWrite(categoryName);
            for (var i = _startRow + 1; i <= _inputWorksheet.Dimension.End.Row; i++)
            {
                var cellContent = EpplusOperation.GetCellValue(_inputWorksheet, i, nvColumn);
                if (cellContent.Equals(HeaderNv, StringComparison.OrdinalIgnoreCase))
                {
                    powerOverWrite.DcCategory = GetDcCategoryName(categoryName);
                    powerOverWrite.LevelSheet = powerOverWrite.GetLevelName();
                    powerOverWriteSheet.PowerOverWrite.Add(powerOverWrite);
                    //Start a new Category
                    categoryName = EpplusOperation.GetCellValue(_inputWorksheet, i, nvColumn - 1);
                    powerOverWrite = new PowerOverWrite(categoryName);
                }
                else
                {
                    //add pin definitions
                    var row = new PowerOverWriteRow();
                    row.PinName = EpplusOperation.GetCellValue(_inputWorksheet, i, nvColumn - 1);
                    row.Nv = EpplusOperation.GetCellValue(_inputWorksheet, i, nvColumn);
                    if (nvValtColumn != -1)
                        row.NvValt = EpplusOperation.GetCellValue(_inputWorksheet, i, nvValtColumn);
                    row.HvRatio = EpplusOperation.GetCellValue(_inputWorksheet, i, hvRatioColumn);
                    row.LvRatio = EpplusOperation.GetCellValue(_inputWorksheet, i, lvRatioColumn);
                    row.Ifold = EpplusOperation.GetCellValue(_inputWorksheet, i, ifoldColumn);
                    row.Vil = EpplusOperation.GetCellValue(_inputWorksheet, i, vilColumn);
                    row.Vih = EpplusOperation.GetCellValue(_inputWorksheet, i, vihColumn);
                    row.Vol = EpplusOperation.GetCellValue(_inputWorksheet, i, volColumn);
                    row.Voh = EpplusOperation.GetCellValue(_inputWorksheet, i, vohColumn);
                    row.Iol = EpplusOperation.GetCellValue(_inputWorksheet, i, iolColumn);
                    row.Ioh = EpplusOperation.GetCellValue(_inputWorksheet, i, iohColumn);
                    row.Vt = EpplusOperation.GetCellValue(_inputWorksheet, i, vtColumn);
                    row.Vcl = EpplusOperation.GetCellValue(_inputWorksheet, i, vclColumn);
                    row.Vch = EpplusOperation.GetCellValue(_inputWorksheet, i, vchColumn);
                    row.DriveMode = EpplusOperation.GetCellValue(_inputWorksheet, i, driverModeColumn);
                    row.Vicm = EpplusOperation.GetCellValue(_inputWorksheet, i, vicmColumn);
                    row.Vid = EpplusOperation.GetCellValue(_inputWorksheet, i, vidColumn);
                    row.Vod = EpplusOperation.GetCellValue(_inputWorksheet, i, vodColumn);
                    row.RowNum = i.ToString();
                    if (!row.PinName.Equals("")) powerOverWrite.DataRows.Add(row);
                }
            }

            powerOverWrite.DcCategory = GetDcCategoryName(categoryName);
            powerOverWrite.LevelSheet = powerOverWrite.GetLevelName();
            if (!string.IsNullOrEmpty(categoryName))
                powerOverWriteSheet.PowerOverWrite.Add(powerOverWrite);
            return powerOverWriteSheet;
        }

        private string GetDcCategoryName(string categoryName)
        {
            return categoryName.Equals("Scan", StringComparison.OrdinalIgnoreCase) ||
                   categoryName.Equals("Mbist", StringComparison.OrdinalIgnoreCase)
                ? categoryName
                : "HardIP_" + categoryName;
        }

        #endregion
    }
}