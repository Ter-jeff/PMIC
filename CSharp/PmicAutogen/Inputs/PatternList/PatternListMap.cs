using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using IgxlData.Others.PatternListCsvFile;

namespace PmicAutogen.Inputs.PatternList
{
    public class PatternListMap
    {
        private readonly List<string> _patternListInUserPath = new List<string>();
        private readonly List<TimeSetBlock2Category> _timeSetBlock2Categories = new List<TimeSetBlock2Category>();
        private readonly Dictionary<string, TimeSetItem> _timeSetVersionDic;
        public readonly List<PatternListCsvRow> PatternListCsvRows = new List<PatternListCsvRow>();

        public static PatternListMap Initialize(string patListCsvFile, string timeSetPath, string patternPath)
        {
            return new PatternListMap(patListCsvFile, timeSetPath, patternPath);
        }

        public string GetTimeSet(string patternName)
        {
            if (PatternListCsvRows.Exists(x =>
                x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase)))
            {
                var pattern = PatternListCsvRows.Find(x =>
                    x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase));
                return pattern.ActualTimeSetVersion;
                //var timeSet = pattern.TimeSetVersion;              
                //if (_timeSetVersionDic.ContainsKey(timeSet))
                //    return timeSet + "_" + _timeSetVersionDic[timeSet].Version;
                //return "";
                
            }

            return "TBD";
        }

        public bool GetStatusInPatternList(string patternName, out string status)
        {
            status = "";
            if (!PatternListCsvRows.Exists(x =>
                x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase)))
            {
                status = "MissPattInPattList";
                return false;
            }

            var pattern = PatternListCsvRows.Find(x =>
                x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase));
            var name = Regex.Replace(
                Regex.Replace(Path.GetFileNameWithoutExtension(pattern.FileVersion), ".atp$", "",
                    RegexOptions.IgnoreCase), ".pat$", "", RegexOptions.IgnoreCase);
            if (!_patternListInUserPath.Exists(x => x.Equals(name, StringComparison.CurrentCultureIgnoreCase)))
            {
                status = "MissPattInPatternFolder";
                return false;
            }

            if (pattern.TimeSetVersion.ToLower() == "na")
            {
                status = "MissTimesetInPattList";
                return false;
            }

            if (pattern.FileVersion == "NA")
            {
                status = "MissFileVersionInPattList";
                return false;
            }

            return true;
        }

        public bool Contains(string tSetSheet)
        {
            var isFoundSameRow = false;
            foreach (var oneRow in _timeSetBlock2Categories)
                if (oneRow.TimeSetSheet.Equals(tSetSheet, StringComparison.CurrentCultureIgnoreCase))
                {
                    isFoundSameRow = true;
                    break;
                }

            return isFoundSameRow;
        }

        public bool SetRow(string tSetSheet, BlockType block, string category)
        {
            var isFoundSameRow = false;
            foreach (var oneRow in _timeSetBlock2Categories)
                if (oneRow.TimeSetSheet == tSetSheet && oneRow.Block == block && oneRow.Category == category)
                {
                    isFoundSameRow = true;
                    break;
                }

            if (!isFoundSameRow)
            {
                var oneRow = new TimeSetBlock2Category();
                oneRow.TimeSetSheet = tSetSheet;
                oneRow.Block = block;
                oneRow.Category = category;
                _timeSetBlock2Categories.Add(oneRow);
            }

            return !isFoundSameRow;
        }

        public string GetCategoryUsageTsetSheetName(string category)
        {
            var tSetSheetList = new List<string>();
            var blockCategoryList = _timeSetBlock2Categories.FindAll(t => t.Category.Equals(category)).ToArray();
            foreach (var lItem in blockCategoryList)
                if (!tSetSheetList.Contains(lItem.TimeSetSheet))
                    tSetSheetList.Add(lItem.TimeSetSheet);
            if (tSetSheetList.Count > 0)
                return string.Join(", ", tSetSheetList.ToArray());

            return "";
        }

        #region Constructor

        public PatternListMap()
        {
        }

        public PatternListMap(string patListCsvFile, string timeSetPath, string patternPath)
        {
            if (File.Exists(patListCsvFile))
            {
                var patternListReader = new PatternListReader();
                PatternListCsvRows = patternListReader.ReadPatList(patListCsvFile);
            }

            if (!string.IsNullOrEmpty(patternPath) && Directory.Exists(patternPath))
                _patternListInUserPath = Directory.GetFiles(patternPath, "*.pat.gz", SearchOption.AllDirectories)
                    .Select(x =>
                        Regex.Replace(Path.GetFileNameWithoutExtension(x), ".atp$", "", RegexOptions.IgnoreCase))
                    .Select(x => Regex.Replace(x, ".pat$", "", RegexOptions.IgnoreCase)).ToList();

            #region CheckTiming Set

            var timeSetFiles = Directory.GetFiles(timeSetPath, "TIMESET*.txt", SearchOption.TopDirectoryOnly);
            _timeSetVersionDic = new Dictionary<string, TimeSetItem>();
            foreach (var file in timeSetFiles)
            {
                var setName = file.Split('\\').Last();
                if (Regex.IsMatch(setName, @"_\d+.TXT$", RegexOptions.IgnoreCase))
                {
                    var timeSet = Regex.Match(setName, @"(?<str>.*)_\d+.TXT$", RegexOptions.IgnoreCase).Groups["str"]
                        .ToString().ToUpper();
                    var paraVer = Convert.ToInt32(Regex.Match(setName, @".*_(?<ver>\d+).TXT$", RegexOptions.IgnoreCase)
                        .Groups["ver"].ToString());
                    var timeItem = new TimeSetItem();
                    timeItem.Version = paraVer;
                    var srTimeSet = new StreamReader(file);
                    do
                    {
                        var line = srTimeSet.ReadLine();
                        if ((line == null) | (line == string.Empty)) continue;

                        if (Regex.IsMatch(line, "Timing Mode"))
                        {
                            var tmpAry = Regex.Split(line.Trim(), @",|\t");
                            timeItem.TimeMod = tmpAry[1].ToUpper();
                            break;
                        }
                    } while (srTimeSet.Peek() != -1);

                    srTimeSet.Close();

                    if (!_timeSetVersionDic.ContainsKey(timeSet))
                    {
                        _timeSetVersionDic.Add(timeSet, timeItem);
                    }
                    else
                    {
                        if (_timeSetVersionDic[timeSet].Version < timeItem.Version)
                        {
                            _timeSetVersionDic[timeSet].Version = timeItem.Version;
                            _timeSetVersionDic[timeSet].TimeMod = timeItem.TimeMod;
                        }
                    }
                }
            }

            foreach (var pattern in PatternListCsvRows)
            {
                var timeSet = pattern.TimeSetVersion;
                if (_timeSetVersionDic.ContainsKey(timeSet))
                    pattern.ActualTimeSetVersion = timeSet + "_" + _timeSetVersionDic[timeSet].Version;
                else
                    pattern.ActualTimeSetVersion = timeSet;
            }

            #endregion
        }

        #endregion
    }

    public class TimeSetBlock2Category
    {
        public BlockType Block;
        public string Category;
        public string TimeSetSheet;
    }

    public enum BlockType
    {
        Common,
        Scan,
        Mbist,
        HardIp,
        BScan
    }
}