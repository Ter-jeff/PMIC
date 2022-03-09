
using System.Collections.Generic;
using System.Linq;

namespace AutomationCommon.Utility
{
    public class ComCombine
    {
        public static string CombineByUnderLine(string str1, string str2)
        {
            str1 = str1.Trim();
            str2 = str2.Trim();
            if (string.IsNullOrEmpty(str1))
                return str2;
            if (string.IsNullOrEmpty(str2))
                return str1;
            return str1 + "_" + str2;
        }

        public static string CombineEnableWord(string str1, string str2)
        {
            str1 = str1.Trim();
            str2 = str2.Trim();
            if (string.IsNullOrEmpty(str1))
                return str2;
            if (string.IsNullOrEmpty(str2))
                return str1;
            return "(" + str1 + ") && " + str2;
        }

        public static string ConnectStringListByUnderLine(List<string> pStrList)
        {
            string result = "";
            if (!pStrList.Any())
            {
                return result;
            }

            for (int i = 0; i < pStrList.Count; i++)
            {
                string temp = pStrList[i].Trim();
                if (result == "" && temp != "")
                {
                    result = temp;
                }
                else if (temp != "")
                {
                    result = result + "_" + temp;
                }
            }
            return result;
        }

    }
}