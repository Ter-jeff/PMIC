using System;
using AutomationCommon.EpplusErrorReport;
using AutomationCommon.Utility;
using OfficeOpenXml;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Base;

namespace PmicAutogen.Inputs.TestPlan.TestPlanPreCheck.Sheet
{
    public class PreCheckIoLevels : PreCheckBase
    {
        public PreCheckIoLevels(ExcelWorkbook workbook, string sheetName) : base(workbook, sheetName)
        {
        }

        protected override bool CheckBusiness()
        {
            return CheckBusiness(Worksheet);
        }


        public bool CheckBusiness(ExcelWorksheet worksheet)
        {
            if (worksheet == null) return false;

            return PreCheck();
        }

        private bool PreCheck()
        {
            var flag = true;
            var ioLevelsItem = new IoLevelsItem();
            for (var i = StartRow + 1; i <= StopRow; i++)
            {
                var row = new IoLevelsRow(SheetName);
                row.RowNum = i;

                var cnt = 0;
                var domainIndex = 0;
                var hasVdd = false;
                var hasVih = false;
                var hasVil = false;
                var hasVoh = false;
                var hasVol = false;
                for (var j = StartColumn + 1; j <= StopColumn; j++)
                {
                    var levelName = EpplusOperation.GetMergedCellValue(Worksheet, StartRow - 1, j).Trim();
                    var headerName = EpplusOperation.GetMergedCellValue(Worksheet, StartRow, j).Trim();

                    if (!string.IsNullOrEmpty(levelName) &&
                        headerName.Equals("Domain", StringComparison.OrdinalIgnoreCase))
                    {
                        if (cnt != 0)
                        {
                            row.IoLevelDate.Add(ioLevelsItem);
                            flag = CheckMissingColumn(hasVdd, hasVih, hasVil, hasVoh, hasVol, ioLevelsItem, domainIndex,
                                flag);
                        }

                        ioLevelsItem = new IoLevelsItem(levelName);
                        domainIndex = j;
                        hasVdd = false;
                        hasVih = false;
                        hasVil = false;
                        hasVoh = false;
                        hasVol = false;
                        cnt++;
                    }

                    var value = EpplusOperation.GetMergedCellValue(Worksheet, i, j).Trim();
                    if (headerName.Equals("Domain", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!CheckDomain(i, j)) flag = false;
                        ioLevelsItem.Domain = value;
                    }
                    else if (headerName.Equals("Vdd", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!CheckVdd(i, j)) flag = false;
                        hasVdd = true;
                    }
                    else if (headerName.Equals("Vih", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!CheckLevel(i, j, ioLevelsItem.Domain)) flag = false;
                        hasVih = true;
                    }
                    else if (headerName.Equals("Vil", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!CheckLevel(i, j, ioLevelsItem.Domain)) flag = false;
                        hasVil = true;
                    }
                    else if (headerName.Equals("Voh", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!CheckLevel(i, j, ioLevelsItem.Domain)) flag = false;
                        hasVoh = true;
                    }
                    else if (headerName.Equals("Vol", StringComparison.OrdinalIgnoreCase))
                    {
                        if (!CheckLevel(i, j, ioLevelsItem.Domain)) flag = false;
                        hasVol = true;
                    }

                    if (j == StopColumn)
                    {
                        if (cnt != 0)
                        {
                            row.IoLevelDate.Add(ioLevelsItem);
                            flag = CheckMissingColumn(hasVdd, hasVih, hasVil, hasVoh, hasVol, ioLevelsItem, domainIndex,
                                flag);
                        }

                        ioLevelsItem = new IoLevelsItem(levelName);
                        cnt++;
                    }
                }
            }

            return flag;
        }

        private bool CheckMissingColumn(bool hasVdd, bool hasVih, bool hasVil, bool hasVoh, bool hasVol,
            IoLevelsItem ioLevelsItem, int domainIndex, bool flag)
        {
            if (!hasVdd || !hasVih || !hasVil || !hasVoh || !hasVol)
            {
                var errorMessage = string.Format("The Vdd/Vil/Vih/Vol/Voh columns are missing for {0} !",
                    ioLevelsItem.Domain);
                EpplusErrorManager.AddError(BasicErrorType.FormatError, ErrorLevel.Error, SheetName, StartRow,
                    domainIndex, errorMessage);
                flag = false;
            }

            return flag;
        }

        private bool CheckDomain(int i, int j)
        {
            var value = EpplusOperation.GetMergedCellValue(Worksheet, i, j).Trim().Replace(" ", "");
            if (string.IsNullOrEmpty(value))
            {
                var errorMessage = string.Format("The Domain : {0} can not be empty !", value);
                EpplusErrorManager.AddError(BasicErrorType.FormatError, ErrorLevel.Error, SheetName, i, j,
                    errorMessage);
                return false;
            }

            return true;
        }

        private bool CheckVdd(int i, int j)
        {
            var value = EpplusOperation.GetMergedCellValue(Worksheet, i, j).Trim().Replace(" ", "");
            if (!string.IsNullOrEmpty(value))
            {
                double result;
                if (double.TryParse(value, out result))
                    return true;

                var errorMessage = string.Format("\"{0}\" should be a number !", value);
                EpplusErrorManager.AddError(BasicErrorType.FormatError, ErrorLevel.Error, SheetName, i, j,
                    errorMessage);
                return false;
            }

            return true;
        }

        private bool CheckLevel(int i, int j, string domain)
        {
            var value = EpplusOperation.GetMergedCellValue(Worksheet, i, j).Trim().Replace(" ", "");
            var flag = SheetStructureManager.IsFormula(value, domain);
            if (!flag)
            {
                var errorMessage = string.Format("\"{0}\" should be a number or formula !!!", value);
                EpplusErrorManager.AddError(BasicErrorType.FormatError, ErrorLevel.Error, SheetName, i, j,
                    errorMessage);
                return false;
            }

            return true;
        }
    }
}