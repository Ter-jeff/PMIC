using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace CommonLib.Extension
{
    public static class StringExtensions
    {
        public static bool IsOpened(this string filePath)
        {
            if (!File.Exists(filePath)) return false;
            try
            {
                Stream s = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.None);
                s.Close();
                return false;
            }
            catch (Exception)
            {
                return true;
            }
        }


        public static string TrimSpace(this string input)
        {
            return input.Replace("\n", "").Replace(" ", "").Replace("\t", "").Replace("\r", "").Trim();
        }

        public static string SheetName2Block(this string name)
        {
            if (name.Equals("DCTEST_Func", StringComparison.OrdinalIgnoreCase))
                return "IO";
            if (name.Equals("DCTEST_IDCODE", StringComparison.OrdinalIgnoreCase))
                return "JTAG";
            return name;
        }

        public static string AddBlockFlag(this string source, string name)
        {
            if (string.IsNullOrEmpty(source))
                return "F_" + name;
            return source + ",F_" + name;
        }

        public static string PadBoth(this string source, int length)
        {
            var spaces = length - source.Length;
            var padLeft = spaces / 2 + source.Length;
            return source.PadLeft(padLeft).PadRight(length);
        }

        public static string GetSortPatNameForBinTable(this string patName)
        {
            if (string.IsNullOrEmpty(patName.Trim()))
                return "";

            var items = patName.Split('_');
            if (items.Length < 11)
                return patName;

            // pp_rtca0_c_fulp_io_xxxx_bsr_jtg_uns_allfrv_si_vih
            // Bin_FUNC_RTCA0_IO_JTG_UNS_VIH” + ”HV/LV/NV…”
            var resultList = new List<string>();
            resultList.Add(items[1]);
            resultList.Add(items[4]);
            resultList.Add(items[6]);
            resultList.Add(items[7]);
            resultList.Add(items[8]);
            if (items.Length >= 12) resultList.Add(items[11]);
            if (items.Length >= 13) resultList.Add(items[12]);
            var result = string.Join("_", resultList);
            return result.ToUpper();
        }

        public static bool IsInit(this string source)
        {
            var tokens = source.Split('_');
            if (tokens.Length > 4)
                if (tokens.Length > 4)
                    return Regex.IsMatch(tokens[3], @"IN\d\d", RegexOptions.IgnoreCase);
            return false;
        }
    }
}