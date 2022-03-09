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

            int idxProduct = 0;
            int idxVersion = 0;
            int idxTpCategory = 0;
            int idxAtpName = 0;
            int idxOpCode = 0;
            int idxScanMode = 0;
            int idxHalt = 0;
            int idxCompilation = 0;
            int idxMd5 = 0;
            int idxHLv = 0;
            int idxScanSetupTSet = 0;
            var reader = new StreamReader(File.OpenRead(_compileFile));
            bool dataStart = false;
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine();
                var values = line.Split(',');
                if (dataStart)
                {
                    if (values[idxAtpName].Length > 10) //key
                    {
                        var compItem = new CompileITem();
                        compItem.Product = values[idxProduct];
                        compItem.Version = values[idxVersion];
                        compItem.TpCategory = values[idxTpCategory];
                        compItem.AtpName = values[idxAtpName];
                        compItem.OpCode = values[idxOpCode];
                        compItem.ScanMode = values[idxScanMode];
                        compItem.Halt = values[idxHalt];
                        compItem.Compilation = values[idxCompilation];
                        compItem.Md5 = values[idxMd5];
                        compItem.HLv = values[idxHLv];
                        compItem.ScanSetupTSet = values[idxScanSetupTSet];
                        if (!_compileList.ContainsKey(values[idxAtpName]))
                            _compileList.Add(values[idxAtpName], compItem);
                    }
                }
                if (Regex.IsMatch(line, ".*Product.*Compilation.*")) //the 1st line
                {
                    for (int item = 0; item < values.Count(); item++)
                    {
                        if (values[item] == "Product") idxProduct = item;
                        if (values[item] == "Version") idxVersion = item;
                        if (values[item] == "T/P Category") idxTpCategory = item;
                        if (values[item] == "AtpName") idxAtpName = item;
                        if (values[item] == "OpCode") idxOpCode = item;
                        if (values[item] == "ScanMode") idxScanMode = item;
                        if (values[item] == "Halt") idxHalt = item;
                        if (values[item] == "Compilation") idxCompilation = item;
                        if (values[item] == "MD5") idxMd5 = item;
                        if (values[item] == "HLV") idxHLv = item;
                        if (values[item] == "ScanSetupTSet") idxScanSetupTSet = item;
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