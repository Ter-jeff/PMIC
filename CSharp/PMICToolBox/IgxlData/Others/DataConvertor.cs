using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.IgxlSheets;

namespace IgxlData.Others
{
    public class DataConvertor
    {

        public static string ConvertUnits(string limitStr)
        {
            if (limitStr.Contains("10^"))
                limitStr = limitStr.Replace("*10^", "E");
            if (limitStr == "" || limitStr.Contains("E") || Regex.IsMatch(limitStr, @"^(\d|\.|-)+$"))//Limit value may be 1.2E-5
                return limitStr;
            if (Regex.IsMatch(limitStr, @"^(\d|\.|-)+(\w)*$"))
            {
                string limitNum = Regex.Match(limitStr, @"(?<num>((\d|\.|-)+))[^\*]*").Groups["num"].ToString();
                if (limitNum == "0")
                    return limitNum;
                string limitUnit = limitStr.Replace(limitNum, "").Trim();
                double rate = 1;
                if (Regex.IsMatch(limitUnit, "^m.*"))
                    rate = 1 / (double)1000;
                else if (Regex.IsMatch(limitUnit, "^u.*"))
                    rate = 1 / (double)1000000;
                else if (Regex.IsMatch(limitUnit, "^n.*"))
                    rate = 1 / (double)1000000000;
                else if (Regex.IsMatch(limitUnit.ToLower(), "^k.*"))
                    rate = 1000;
                else if (Regex.IsMatch(limitUnit, "^M.*"))
                    rate = 1000000;
                else if (Regex.IsMatch(limitUnit, "^G.*"))
                    rate = 1000000000;
                double value;
                if (double.TryParse(limitNum, out value))
                {
                    return (value * rate).ToString("G");
                }
            }
            return limitStr;
        }
    }
}
