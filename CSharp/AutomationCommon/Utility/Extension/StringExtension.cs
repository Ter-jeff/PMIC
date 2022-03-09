using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace AutomationCommon.Utility
{
    public static class StringExtension
    {
        public static string PadBoth(this string source, int length)
        {
            int spaces = length - source.Length;
            int padLeft = spaces / 2 + source.Length;
            return source.PadLeft(padLeft).PadRight(length);
        }

        public static string GetSortPatNameForBinTable(this string patname)
        {
            if (string.IsNullOrEmpty(patname.Trim()))
                return "";

            string[] items = patname.Split('_');
            if (items.Length < 11)
                return patname;

            // pp_rtca0_c_fulp_io_xxxx_bsr_jtg_uns_allfrv_si_vih
            // Bin_FUNC_RTCA0_IO_JTG_UNS_VIH” + ”HV/LV/NV…”
            List<string> resultList = new List<string>();
            resultList.Add(items[1]);
            resultList.Add(items[4]);
            resultList.Add(items[7]);
            resultList.Add(items[8]);
            if (items.Length >= 12) resultList.Add(items[11]);
            if (items.Length >= 13) resultList.Add(items[12]);
            string result = string.Join("_", resultList);
            return result.ToUpper();
        }
    }
}