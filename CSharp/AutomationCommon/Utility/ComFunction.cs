using System;
using System.Text.RegularExpressions;

namespace AutomationCommon.Utility
{
    public class ComFunction
    {

        public static bool CompareString(string source, string compare, bool ignoreUnderLine = false)
        {
            if (source == null) return false;

            foreach (var pCompareSplit in compare.Split('|'))
            {
                var str1 = Normalization(source);
                var str2 = Normalization(pCompareSplit);


                if (ignoreUnderLine)
                {
                    str1 = str1.Replace(' ', '_');
                    str2 = str2.Replace(' ', '_');
                }

                str2 = ReplaceSpecialCase(str2);

                if (str2.IndexOf("\\*", StringComparison.Ordinal) >= 0 || str2.IndexOf("\\?", StringComparison.Ordinal) >= 0)
                {
                    str2 = str2.Replace("\\*", ".*");
                    str2 = str2.Replace("\\?", ".+");
                }

                str2 = "^" + str2 + "$";
                if (Regex.IsMatch(str1, str2, RegexOptions.IgnoreCase)) return true;
            }

            return false;
        }

        private static string Normalization(string text)
        {
            var result = text.Trim();

            result = ReplaceEnter(result);

            result = ReplaceDoubleBlank(result);

            return result;
        }

        private static string ReplaceSpecialCase(string text)
        {
            return Regex.Escape(text);
        }

        private static string ReplaceDoubleBlank(string text)
        {
            var lStrResult = text;
            do
            {
                lStrResult = lStrResult.Replace("  ", " ");
            } while (lStrResult.IndexOf("  ", StringComparison.Ordinal) >= 0);

            return lStrResult;
        }

        private static string ReplaceEnter(string text)
        {
            return text.Replace("\n", " ");
        }
    }
}