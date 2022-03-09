using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.Others.PatternListCsvFile
{
    public class CompilePatReader
    {
        private string _compileFile;
        //private string _oriPatListCsvFile;
        //private string _twPatListCsvFile;
        //private string _timeSetFolder;
        private Dictionary<string, CompileITem> _compileList = new Dictionary<string, CompileITem>();
        //private Dictionary<string, OriPatListItem> _oriPatCsvList = new Dictionary<string, OriPatListItem>();

        public Dictionary<string, CompileITem> ReadCompileFile(string filePath)
        {
            _compileFile = filePath;

            int idx_product = 0;
            int idx_version = 0;
            int idx_tpCategory = 0;
            int idx_atpName = 0;
            int idx_opCode = 0;
            int idx_scanMode = 0;
            int idx_halt = 0;
            int idx_compilation = 0;
            int idx_md5 = 0;
            int idx_hLV = 0;
            int idx_ScanSetupTSet = 0;
            var reader = new StreamReader(File.OpenRead(_compileFile));
            bool dataStart = false;
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');
                if (dataStart)
                {
                    if (values[idx_atpName].Length > 10) //key
                    {
                        var compItem = new CompileITem();
                        compItem.Product = values[idx_product];
                        compItem.Version = values[idx_version];
                        compItem.TpCategory = values[idx_tpCategory];
                        compItem.AtpName = values[idx_atpName];
                        compItem.OpCode = values[idx_opCode];
                        compItem.ScanMode = values[idx_scanMode];
                        compItem.Halt = values[idx_halt];
                        compItem.Compilation = values[idx_compilation];
                        compItem.Md5 = values[idx_md5];
                        compItem.HLv = values[idx_hLV];
                        compItem.ScanSetupTSet = values[idx_ScanSetupTSet];
                        if (!_compileList.ContainsKey(values[idx_atpName]))
                            _compileList.Add(values[idx_atpName], compItem);
                    }
                }
                if (Regex.IsMatch(line, ".*Product.*Compilation.*")) //the 1st line
                {
                    for (int item = 0; item < values.Count(); item++)
                    {
                        if (values[item] == "Product") idx_product = item;
                        if (values[item] == "Version") idx_version = item;
                        if (values[item] == "T/P Category") idx_tpCategory = item;
                        if (values[item] == "AtpName") idx_atpName = item;
                        if (values[item] == "OpCode") idx_opCode = item;
                        if (values[item] == "ScanMode") idx_scanMode = item;
                        if (values[item] == "Halt") idx_halt = item;
                        if (values[item] == "Compilation") idx_compilation = item;
                        if (values[item] == "MD5") idx_md5 = item;
                        if (values[item] == "HLV") idx_hLV = item;
                        if (values[item] == "ScanSetupTSet") idx_ScanSetupTSet = item;
                    }
                    dataStart = true;
                }
            }
            reader.Close();
            return _compileList;
        }

        public Dictionary<string, string> GetLatestPatDict()
        {
            var lGenericPatternGroup = new Dictionary<string, List<string>>();
            foreach (var compiledItem in _compileList)
            {
                var patGeneric = new PatternNameInfo(compiledItem.Key);

                if (!lGenericPatternGroup.ContainsKey(patGeneric.GenericName))
                {
                    lGenericPatternGroup.Add(patGeneric.GenericName, new List<string>());
                }
                lGenericPatternGroup[patGeneric.GenericName].Add(patGeneric.PatternVersion + "_" + patGeneric.SiliconVersion + "_" + patGeneric.TimeStamp);
            }

            var lLatestPatDict = new Dictionary<string, string>();
            foreach (var patGroup in lGenericPatternGroup)
            {
                var allVersions = patGroup.Value;
                string latestVersion = "";
                foreach (string patternVersion in allVersions)
                {
                    if (latestVersion == "")
                    {
                        latestVersion = patternVersion;
                        continue;
                    }
                    latestVersion = CompareVersion(latestVersion, patternVersion);
                }
                lLatestPatDict.Add(patGroup.Key, latestVersion);
            }
            return lLatestPatDict;
        }

        private static string CompareVersion(string version1, string version2)
        {
            // Version1:  "1_A0_1510070021"  =>1
            // Version2:  "2_A0_1510070021"  =>2
            // 2 > 1 return Version 2

            int versionnumber1 = Convert.ToInt16(version1.Split('_')[0]);
            int versionnumber2 = Convert.ToInt16(version2.Split('_')[0]);
            if (versionnumber1 > versionnumber2)
                return version1;
            else
                return version2;
        }
    }
}