using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.EMMA;

namespace Automation.Util.PatternListCsvFile
{
    public class LatestPatFromCompileItems
    {


        private Dictionary<string,List<string>> _genericPatternGroup  = new Dictionary<string, List<string>>();

        public LatestPatFromCompileItems(Dictionary<string, CompileITem> compileITems)
        {
            GroupPattern(compileITems);
        }

        public Dictionary<string, string> GetLatestPat()
        {
            var lLatestPatDict = new Dictionary<string, string>();
            foreach (var patGroup in _genericPatternGroup)
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

        private void GroupPattern(Dictionary<string, CompileITem> compileITems)
        {
            foreach (var compiledItem in compileITems)
            {
                var patGeneric = new Library.Pattern(compiledItem.Key);

                if (!_genericPatternGroup.ContainsKey(patGeneric.GenericName))
                {
                    _genericPatternGroup.Add(patGeneric.GenericName, new List<string>());
                }
                _genericPatternGroup[patGeneric.GenericName].Add(patGeneric.PatternVersion + "_" + patGeneric.SiliconVersion + "_" + patGeneric.TimeStamp);
            }
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