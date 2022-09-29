using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using PmicAutogen.Inputs.PatternList;
using System.Collections.Generic;

namespace PmicAutogen.Local
{
    public static class InputFiles
    {
        public static Workbook InteropTestPlanWorkbook;

        public static PatternListMap PatternListMap;
        public static ExcelWorkbook TestPlanWorkbook { get; set; }
        public static ExcelWorkbook ScghWorkbook { get; set; }
        public static List<ExcelWorkbook> VbtGenToolWorkbooks { get; set; }
        public static ExcelWorkbook SettingWorkbook { get; set; }
        public static ExcelWorkbook ConfigWorkbook { get; set; }
        public static ExcelPackage TestPlanExcelPackage { get; set; }
        public static ExcelPackage ScghPackage { get; set; }
        public static List<ExcelPackage> VbtGenToolPackage { get; set; }

        public static void Initialize()
        {
            VbtGenToolWorkbooks = new List<ExcelWorkbook>();
            VbtGenToolPackage = new List<ExcelPackage>();
            PatternListMap = new PatternListMap();
        }
    }
}