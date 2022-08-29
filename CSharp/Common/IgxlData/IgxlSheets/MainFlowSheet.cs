using OfficeOpenXml;

namespace IgxlData.IgxlSheets
{
    public class MainFlowSheet : SubFlowSheet
    {
        public MainFlowSheet(ExcelWorksheet sheet) : base(sheet)
        {
        }

        public MainFlowSheet(string sheetName) : base(sheetName)
        {
        }
    }
}