using CommonLib.Extension;
using OfficeOpenXml;
using System;
using System.Linq;

namespace CommonReaderLib.DebugPlan
{
    public class ProcessConditionReader : MySheetReader
    {
        public ProcessConditionSheet ReadSheet(ExcelWorksheet worksheet)
        {
            var sheetName = worksheet.Name;

            var sheet = new ProcessConditionSheet(sheetName);

            ExcelWorksheet = worksheet;

            if (!GetDimensions())
            {
                sheet.AddDimensionError();
                return null;
            }

            for (var i = StartRowNumber; i <= EndRowNumber; i++)
            {
                var text = ExcelWorksheet.GetCellValue(i, 1).Trim();
                var key = text.Split(':').First();
                if (key.Equals("Efuse Enable word", StringComparison.OrdinalIgnoreCase))
                    sheet.EfuseEnableWord = text.Split(':').Last();
                else if (key.Equals("Tester", StringComparison.OrdinalIgnoreCase))
                    sheet.Tester = text.Split(':').Last();
            }

            return sheet;
        }
    }

    public class ProcessConditionSheet : MySheet
    {
        #region Constructor

        public ProcessConditionSheet(string sheetName)
        {
            SheetName = sheetName;
        }

        #endregion

        public string EfuseEnableWord { get; set; }
        public string Tester { get; set; }
    }
}