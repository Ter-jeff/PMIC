using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using Library.Function;

namespace PmicAutomation.Utility.VbtGenerator.Input
{
    public class PmicIdsTableReader : TableFactory
    {
        public override List<TableSheet> ReadSheet(ExcelWorksheet worksheet)
        {
            int measPinPos = -1;
            int descriptionPos = -1;

            if (worksheet == null)
            {
                return null;
            }

            _excelWorksheet = worksheet;

            Reset();

            if (!GetDimensions())
            {
                return null;
            }

            //if (!CheckHeader())
            //{
            //    return null;
            //}

            List<TableSheet> tableSheets = new List<TableSheet>();
            TableSheet tableSheet = null;
            List<Dictionary<string, string>> table = new List<Dictionary<string, string>>();
            Dictionary<string, string> allPinSettingDic = new Dictionary<string, string>();
            int headerRow = 1;
            string lastInstance = string.Empty;
            for (int i = 1; i <= _endRowNumber; i++)
            {
                if (IsHeaderRow(i) == 1)
                {
                    string block = _excelWorksheet.GetMergeCellValue(i, 1);
                    tableSheet = new TableSheet { Name = worksheet.Name, Block = block };
                    headerRow = i + 1;
                }
                else if (IsHeaderRow(i) == 0)
                {
                    if (tableSheet != null)
                    {
                        tableSheet.Table = table;
                        tableSheet.AllPinSettingDic = allPinSettingDic;
                        tableSheets.Add(tableSheet.DeepClone());
                    }

                    tableSheet = null;
                }
                else if (i != headerRow)
                {
                    if (tableSheet == null)
                    {
                        tableSheet = new TableSheet { Name = worksheet.Name };
                    }

                    string currentInstance = GetInstanceField(headerRow, i, ref descriptionPos);
                    if (!string.IsNullOrEmpty(currentInstance))
                        lastInstance = currentInstance;

                    if (isMeasPinEmpty(headerRow, i, ref measPinPos))
                        continue;

                    Dictionary<string, string> pairs = new Dictionary<string, string>();
                    for (int j = 1; j <= _endColNumber; j++)
                    {
                        string header = _excelWorksheet.GetMergeCellValue(headerRow, j).Replace(" ", "");
                        string content = _excelWorksheet.GetMergeCellValue(i, j);
                        if (header.ToLower().Equals("instance") && string.IsNullOrEmpty(content))
                            content = lastInstance;
                        content = Epplus.ConvertUnit(content);
                        pairs.Add(header, content);
                    }

                    if (IsAllPinSetting(pairs))
                    {
                        allPinSettingDic = pairs;
                    }
                    else
                    {
                        table.Add(pairs);
                    }
                }
            }

            if (tableSheet != null)
            {
                tableSheet.Table = table;
                tableSheet.AllPinSettingDic = allPinSettingDic;
                tableSheets.Add(tableSheet.DeepClone());
            }

            return tableSheets;
        }

        private string GetInstanceField(int headerRow, int currentRow, ref int instancePos)
        {
            if (instancePos == -1)
            {
                for (int c = 1; c <= _endColNumber; ++c)
                {
                    string header = _excelWorksheet.GetMergeCellValue(headerRow, c);
                    if (header.ToUpper().Equals("INSTANCE"))
                    {
                        instancePos = c;
                        break;
                    }
                }
                if (instancePos == -1)
                    return string.Empty;
            }
            return _excelWorksheet.GetMergeCellValue(currentRow, instancePos) ?? string.Empty;
        }

        private bool isMeasPinEmpty(int headerRow, int currentRow, ref int measPinPos)
        {
            if (measPinPos == -1)
            {
                for (int c = 1; c <= _endColNumber; ++c)
                {
                    string header = _excelWorksheet.GetMergeCellValue(headerRow, c);
                    if (header.ToUpper().Equals("MEASUREPINS"))
                    {
                        measPinPos = c;
                        break;
                    }
                }
                if (measPinPos == -1)
                    return true;
            }

            if (string.IsNullOrEmpty(_excelWorksheet.GetMergeCellValue(currentRow, measPinPos)))
                return true;
            else
                return false;
        }
    }
}
