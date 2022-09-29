using CommonLib.Enum;
using CommonLib.WriteMessage;
using OfficeOpenXml;
using PmicAutogen.Inputs.ScghFile.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.ScghFile
{
    public class ScghFileManager
    {
        public ProdCharSheet ScghMbistSheet;
        public ProdCharSheet ScghScanSheet;

        #region Member Function

        public void CheckAll(ExcelWorkbook workbook)
        {
            #region Pre check

            var scghScan = workbook.Worksheets.FirstOrDefault(s => s.Name.ToUpper().EndsWith(PmicConst.ScghScan));
            if (scghScan != null)
            {
                var sheetReader = new ProdCharSheetReader();
                ScghScanSheet = sheetReader.ReadScghSheet(scghScan);
            }

            var scghMbist = workbook.Worksheets.FirstOrDefault(s => s.Name.ToUpper().EndsWith(PmicConst.ScghMbist));
            if (scghMbist != null)
            {
                var sheetReader = new ProdCharSheetReader();
                ScghMbistSheet = sheetReader.ReadScghSheet(scghMbist);
            }

            #endregion

            var lIsPassCheck = true;

            #region Post check

            for (var i = 0; i < ScghScanSheet.RowList.Count; i++)
                if (!IsValidPinName(ScghScanSheet.RowList[i].SupplyVoltage))
                {
                    lIsPassCheck = false;
                    Response.Report("In Workbook: [" + LocalSpecs.TestPlanFileName + "] In Sheet: [" +
                                    ScghScanSheet.SheetName +
                                    "] value is invalid or contains illegal character! Position: Row " + (i + 2)
                                    + " Column: [Supply Voltage]", EnumMessageLevel.Warning, 0);
                }

            for (var i = 0; i < ScghMbistSheet.RowList.Count; i++)
                if (!IsValidPinName(ScghMbistSheet.RowList[i].PeripheralVoltage))
                {
                    lIsPassCheck = false;
                    Response.Report("In Workbook: [" + LocalSpecs.TestPlanFileName + "] In Sheet: [" +
                                    ScghMbistSheet.SheetName +
                                    "] value is invalid or contains illegal character! Position: Row " + (i + 2)
                                    + " Column: [Peripheral Voltage]", EnumMessageLevel.Warning, 0);
                }

            if (lIsPassCheck == false)
                Response.Report("The output skeleton program, " +
                                "will include the characterize sheet, " +
                                "But program might validate fail due to value is invalid " +
                                "or contains illegal character!", EnumMessageLevel.Warning, 0);

            #endregion
        }

        /// <summary>
        ///     Pin name validation
        /// </summary>
        /// <param name="p_strPinName">Pin Name</param>
        /// <returns>Pass return true, else return false</returns>
        private bool IsValidPinName(string pStrPinName)
        {
            var lRtn = true;

            if (string.IsNullOrEmpty(pStrPinName)) return lRtn;

            var lRegexStartWithLetterOrUnderscore = new Regex(@"^[a-zA-Z_]");
            var lRegexNoIllegalCharacter = new Regex(@"^[a-zA-Z0-9_]+$");

            if (!lRegexStartWithLetterOrUnderscore.IsMatch(pStrPinName)) lRtn = false;

            if (!lRegexNoIllegalCharacter.IsMatch(pStrPinName)) lRtn = false;

            return lRtn;
        }

        #endregion
    }
}