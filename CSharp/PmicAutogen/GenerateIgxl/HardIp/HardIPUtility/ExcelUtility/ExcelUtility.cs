using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.EpplusErrorReport;

namespace PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.ExcelUtility
{
    public class ExcelUtility
    {
        public static int GetHeaderIndex(string sheetName, Dictionary<string, int> headerOrder, string header,
            bool optionalFlag = true)
        {
            var headerIndex = headerOrder.FirstOrDefault(a =>
                Regex.IsMatch(a.Key, header, RegexOptions.IgnoreCase) &&
                !Regex.IsMatch(a.Key, "Pattern Release Status")).Value;
            if (headerIndex > 0)
                return headerIndex;
            if (optionalFlag)
            {
                header = header.Replace(@"\s*", " ").Replace(@"\s", " ").Replace(@".*", "");
                var errorMessage = "Missing header " + header + " in sheet " + sheetName;
                EpplusErrorManager.AddError(HardIpErrorType.MissingHeader, ErrorLevel.Error, sheetName, 1, errorMessage,
                    header);
            }

            return 1;
        }
    }
}