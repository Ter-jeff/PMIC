using System.Collections.Generic;
using OfficeOpenXml;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.HardIp.InputReader
{
    public class TestPlanConverter
    {
        public Dictionary<string, List<HardIpPattern>> ReadHardipSheet(ExcelWorksheet sheet)
        {
            var planDic = new Dictionary<string, List<HardIpPattern>>();
            var sheetName = sheet.Name;

            var reader = new TestPlanReader();
            var planSheet = reader.ReadSheet(sheet);
            //planSheet.ConvertRealPatternName(scgh);
            planSheet.DividePatternRow();
            ParsePlanSheet(planSheet);
            planDic.Add(sheetName, planSheet.PatternItems);           
           
            return planDic;
        }

        protected void ParsePlanSheet(TestPlanSheet planSheet)
        {
            var tpPreProcess = new TestPlanSheetPatPreprocess(planSheet);
            tpPreProcess.UpdateSheetPattern();

            var testPlanPatParser = new TestPlanPatParser(planSheet);
            testPlanPatParser.ConvertTpPatterns();
        }
    }
}