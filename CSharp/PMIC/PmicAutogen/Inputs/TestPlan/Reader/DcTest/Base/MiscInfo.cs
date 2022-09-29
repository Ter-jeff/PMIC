using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base
{
    public static class MiscInfoStringExtension
    {
        public static bool IsLimit(this string miscInfo)
        {
            return miscInfo.Equals(Limit, StringComparison.CurrentCultureIgnoreCase);
        }

        public static bool IsNvOnly(this string miscInfo)
        {
            return Regex.IsMatch(miscInfo, NvOnly, RegexOptions.IgnoreCase);
        }

        public static bool IsHvOnly(this string miscInfo)
        {
            return Regex.IsMatch(miscInfo, HvOnly, RegexOptions.IgnoreCase);
        }

        public static bool IsLvOnly(this string miscInfo)
        {
            return Regex.IsMatch(miscInfo, LvOnly, RegexOptions.IgnoreCase);
        }

        public static bool IsNoBin(this string miscInfo)
        {
            return Regex.IsMatch(miscInfo, NoBin, RegexOptions.IgnoreCase);
        }

        public static string GetInstNameSubStr(this string miscInfo)
        {
            var instNameSubStr = "";
            foreach (var info in miscInfo.Split(';'))
                if (info.ToLower().Contains(InstNameSubStr.ToLower()) && info.Contains(":"))
                {
                    var misArr = info.Split(':');
                    if (misArr.Length == 2 &&
                        misArr[0].Equals(InstNameSubStr, StringComparison.OrdinalIgnoreCase))
                    {
                        instNameSubStr = misArr[1];
                        break;
                    }
                }

            return instNameSubStr;
        }

        public static string GetBlockName(this string miscInfo, string sheetName)
        {
            var blockName = GetIpName(miscInfo);

            if (string.IsNullOrEmpty(blockName))
                blockName = GetBlockNameFromSheetName(sheetName);
            return blockName;
        }

        private static string GetBlockNameFromSheetName(string sheetName)
        {
            var arr = sheetName.Split('_').ToList();
            if (arr.Count > 1)
                arr.RemoveAt(0);
            return string.Join("", arr).Replace(" ", "").ToUpper();
        }

        public static string GetIpName(this string miscInfo)
        {
            var blockName = "";
            foreach (var assign in miscInfo.Split(';'))
            {
                var assignArr = assign.Split(':');
                if (assignArr.Length == 2 &&
                    assignArr[0].Equals(IpName, StringComparison.OrdinalIgnoreCase))
                    blockName = assignArr[1].Replace("_", "");
            }

            return blockName;
        }

        public static string GetSubBlockName(this string miscInfo, string patternName, string blockName,
            bool isShmooInChar = false)
        {
            var subBlockName = "";
            foreach (var assign in miscInfo.Split(';'))
            {
                var assignArr = assign.Split(':');
                if (assignArr.Length == 2 && assignArr[0].Trim()
                        .Equals(SubBlockName, StringComparison.OrdinalIgnoreCase))
                {
                    subBlockName = assignArr[1].Replace("_", "");
                    break;
                }
            }

            if (string.IsNullOrEmpty(subBlockName))
                subBlockName = GetSubBlockNameByPattern(patternName, blockName);

            if (isShmooInChar)
                subBlockName += "CZ";
            return subBlockName;
        }

        public static string GetSubBlockNameByPattern(string patternName, string blockName, bool isCheckScghItem = true)
        {
            var subBlocks = new List<string>();
            var patternSeg = patternName.Split('_').ToList();
            var siDmIndex = patternSeg.FindLastIndex(p =>
                p.Equals("SI", StringComparison.OrdinalIgnoreCase) ||
                p.Equals("DM", StringComparison.OrdinalIgnoreCase));
            if (siDmIndex != -1 && siDmIndex != patternSeg.Count - 1)
            {
                var subBlockSegments = patternSeg.GetRange(siDmIndex + 1, patternSeg.Count - siDmIndex - 1);
                foreach (var subBlockSeg in subBlockSegments)
                    if (!subBlockSeg.Equals(blockName, StringComparison.CurrentCultureIgnoreCase) && isCheckScghItem)
                        subBlocks.Add(subBlockSeg);
                    else
                        subBlocks.Add(subBlockSeg);
            }

            return string.Join("_", subBlocks);
        }

        public static string GetSubBlock2Name(this string miscInfo)
        {
            var subBlock2Name = "";
            foreach (var assign in miscInfo.Split(';'))
            {
                var assignArr = assign.Split(':');
                if (assignArr.Length == 2 &&
                    assignArr[0].Equals(SubBlock2Name, StringComparison.OrdinalIgnoreCase))
                {
                    subBlock2Name = assignArr[1].Replace("_", "");
                    break;
                }
            }

            return subBlock2Name;
        }

        #region Misc info Standard name

        private const string Limit = "Limit";
        private const string VbtKey = "VBT";
        private const string Timing = "Timing";
        private const string Vbt = "Func|Trim|Meas";
        private const string OpCode = "opcode";
        private const string RegOpCode = @"opcode\s*:";
        private const string HighTemp = "High_Temp_Only";
        private const string RoomTemp = "Room_Temp_Only";
        private const string InstanceName = "InstanceName";
        private const string RelayOn = "RelayOn";
        private const string RelayOff = "RelayOff";

        private const string NvOnly = "NV_Only_For_HLN_Flow|NvOnly";

        private const string
            HvOnly = "HV_Only_For_HLN_Flow|HvOnly"; //Added to support HV_Only_For_HLN_Flow on 2016/6/27

        private const string
            LvOnly = "LV_Only_For_HLN_Flow|LvOnly"; //Added to support LV_Only_For_HLN_Flow on 2016/6/27

        private const string RemoveNv = @"Run_NV_Flow_Only|RemoveNv"; //Change to Run_NV_Flow_Only on 2016/6/28
        private const string RemoveLv = @"Run_LV_Flow_Only|RemoveLv"; //Change to Run_LV_Flow_Only on 2016/6/28
        private const string RemoveHv = @"Run_HV_Flow_Only|RemoveHv"; //Change to Run_HV_Flow_Only on 2016/6/28
        private const string ReTest = @"Fail_Retest|ReTest"; //Change from ReTest to Fail_Retest on 2016/6/23
        private const string PreNwireEnaOrDis = "FreerunClk";
        private const string FreeRunClkEnableWord = PreNwireEnaOrDis + "Enable";
        private const string FreeRunClkDisableWord = PreNwireEnaOrDis + "Disable";
        private const string NoBin = "No_Fail_Flag_All|NoBin"; //Change from NoBin to No_Fail_Flag_All on 2016/6/24

        private const string NoBinUseLimit = "No_Fail_Flag_UseLimit|NoBinUseLimit";
        private const string RemovePattern = @"Do_Not_Generate|RemovePattern";
        private const string Manual = @"Generate_But_Manually_Modify|Manual";
        private const string SkipCheck = "Skip_Pre_Check";
        private const string InstNameSubStr = "InstNameSubStr";
        private const string Calc = "Calc";
        private const string CalcParameter = "CalcArg";
        private const string IgnorePatInfo = "Ignore_Patt_Comment";
        private const string IgnorePatMeasC = "Ignore_Patt_MeasC";
        private const string RepeatLimit = "^Repeat_Limit";
        private const string NoPattern = "No_patt";
        private const string RegInsInPattern = @"^Instance[\s]*:[\s]*(?<InsName>[\w]+)";
        private const string Block = "Block";
        private const string IpName = "IP";
        private const string SubBlockName = "SubBlock";
        private const string SubBlock2Name = "SubBlock2";
        private const string SubBlockCzName = "SubBlockCZ";
        private const string ExecCond = "ExecCond";

        private const string CallExtraFlow = "Call";
        private const string KeepDsscOut = "Disable_MeasC_Split";
        private const string Cz2Only = "CZ2_Only";

        private const string Usl = "USL";
        private const string Lsl = "LSL";

        #endregion
    }
}