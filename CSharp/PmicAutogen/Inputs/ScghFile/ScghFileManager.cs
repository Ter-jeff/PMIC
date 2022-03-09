using AutomationCommon.DataStructure;
using OfficeOpenXml;
using PmicAutogen.InputPackages;
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

            var scghScan = workbook.Worksheets.FirstOrDefault(s=>s.Name.ToUpper().EndsWith(PmicConst.ScghScan));
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

            bool l_IsPassCheck = true;
            #region Post check
            for (int i=0;i<ScghScanSheet.RowList.Count;i++)
            {
                if (!IsValidPinName(ScghScanSheet.RowList[i].SupplyVoltage))
                {
                    l_IsPassCheck = false;
                    Response.Report("In Workbook: ["+ LocalSpecs.TestPlanFileName + "] In Sheet: [" + ScghScanSheet.SheetName + "] value is invalid or contains illegal character! Position: Row " + (i+2).ToString()
                                + " Column: [Supply Voltage]", MessageLevel.Warning, 0);
                }
                else
                { 
                    //do nothing
                }
            }

            for (int i = 0; i < ScghMbistSheet.RowList.Count; i++)
            {
                if (!IsValidPinName(ScghMbistSheet.RowList[i].PeripheralVoltage))
                {
                    l_IsPassCheck = false;
                    Response.Report("In Workbook: [" + LocalSpecs.TestPlanFileName + "] In Sheet: [" + ScghMbistSheet.SheetName + "] value is invalid or contains illegal character! Position: Row " + (i + 2).ToString()
                                + " Column: [Peripheral Voltage]", MessageLevel.Warning, 0);
                }
                else
                {
                    //do nothing
                }
            }

            if (l_IsPassCheck == false)
            {
                Response.Report("The output skeleton program, " +
                    "will include the characterize sheet, " +
                    "But program might validate fail due to value is invalid " +
                    "or contains illegal character!", MessageLevel.Warning, 0);
            }
            else
            { 
                //do nothing
            }

            #endregion
        }

        /// <summary>
        /// Pin name validation
        /// </summary>
        /// <param name="p_strPinName">Pin Name</param>
        /// <returns>Pass return true, else return false</returns>
        private bool IsValidPinName(string p_strPinName)
        {
            bool l_Rtn = true;

            if (string.IsNullOrEmpty(p_strPinName))
            {
                return l_Rtn;
            }

            Regex l_RegexStartWithLetterOrUnderscore = new Regex(@"^[a-zA-Z_]");
            Regex l_RegexNoIllegalCharacter = new Regex(@"^[a-zA-Z0-9_]+$");

            if (!l_RegexStartWithLetterOrUnderscore.IsMatch(p_strPinName))
            {
                l_Rtn = false;
            }
            else
            {
                //do nothing
            }

            if (!l_RegexNoIllegalCharacter.IsMatch(p_strPinName))
            {
                l_Rtn = false;
            }
            else
            {
                //do nothing
            }

            return l_Rtn;
        }

        #endregion
    }
}