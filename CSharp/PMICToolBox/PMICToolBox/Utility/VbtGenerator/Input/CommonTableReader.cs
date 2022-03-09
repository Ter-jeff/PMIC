using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Library.Function;
using OfficeOpenXml;

namespace PmicAutomation.Utility.VbtGenerator.Input
{
    public class CommonTableReader : TableFactory
    {
        public override List<TableSheet> ReadSheet(ExcelWorksheet worksheet)
        {
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

            if (!CheckHeader())
            {
                return null;
            }

            List<TableSheet> tableSheets = new List<TableSheet>();
            TableSheet tableSheet = null;
            List<Dictionary<string, string>> table = new List<Dictionary<string, string>>();
            Dictionary<string, string> allPinSettingDic = new Dictionary<string, string>();
            int headerRow = 1;
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

                    Dictionary<string, string> pairs = new Dictionary<string, string>();
                    for (int j = 1; j <= _endColNumber; j++)
                    {
                        string header = _excelWorksheet.GetMergeCellValue(headerRow, j).Replace(" ", "");
                        string content = _excelWorksheet.GetMergeCellValue(i, j);
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
    }
}
