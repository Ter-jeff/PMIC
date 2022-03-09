using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace IgxlData.Others.PatternListCsvFile
{
    public class NonTwCsvFileReader
    {
        public Dictionary<string, OriPatListItem> ReadNonTwCsvFile(string tempOriPatListCsvFile, string timeSetFolder)
        {
            Dictionary<string, OriPatListItem> oriPatCsvList = new Dictionary<string, OriPatListItem>();

            #region CheckTiming Set
            string[] timeSetFiles = Directory.GetFiles(timeSetFolder, "TIMESET*.txt", SearchOption.TopDirectoryOnly);
            Dictionary<string, TimeSetItem> dicTimeSetVersion = new Dictionary<string, TimeSetItem>();
            foreach (var file in timeSetFiles)
            {
                var setName = file.Split('\\').Last();
                if (Regex.IsMatch(setName, @"_\d+.TXT$", RegexOptions.IgnoreCase)) //              paramStr = Regex.Match(line, @"\((?<str>.*)\)").Groups["str"].ToString();
                {
                    var timeset = Regex.Match(setName, @"(?<str>.*)_\d+.TXT$", RegexOptions.IgnoreCase).Groups["str"].ToString().ToUpper();
                    var paraVer = Convert.ToInt32(Regex.Match(setName, @".*_(?<ver>\d+).TXT$", RegexOptions.IgnoreCase).Groups["ver"].ToString());
                    var timeItem = new TimeSetItem();

                    timeItem.Version = paraVer;

                    var srTimeset = new StreamReader(file);
                    do
                    {
                        var line = srTimeset.ReadLine(); //讀取每一行
                        if (line == null | line == String.Empty) continue;

                        if (Regex.IsMatch(line, "Timing Mode"))
                        {
                            var tmpAry = Regex.Split(line.Trim(), @",|\t");
                            timeItem.TimeMod = tmpAry[1].ToUpper();
                            break;
                        }

                    } while (srTimeset.Peek() != -1);
                    srTimeset.Close();

                    if (!dicTimeSetVersion.ContainsKey(timeset))
                    {
                        dicTimeSetVersion.Add(timeset, timeItem);
                    }
                    else
                    {
                        if (dicTimeSetVersion[timeset].Version < timeItem.Version)
                        {
                            dicTimeSetVersion[timeset].Version = timeItem.Version;
                            dicTimeSetVersion[timeset].TimeMod = timeItem.TimeMod;
                        }
                    }
                }
            }
            #endregion

            #region Pattern List Reader of Customer's format
            int idx_idx = -1;
            int idx_pattern = -1;
            int idx_latestVersion = -1;
            int idx_releaseDate = -1;
            int idx_useNoUse = -1;
            int idx_dRI = -1;
            int idx_releaseNote = -1;
            int idx_radarNum = -1;
            int idx_org = -1;
            int idx_typeSpec = -1;
            int idx_timesetLatest = -1;
            int idx_fileVersions = -1;
            int idx_opCode = -1;
            int idx_scanMode = -1;
            int idx_halt = -1;
            int idx_compilation = -1;
            int idx_tpCategory = -1;
            int idx_hLV = -1;
            int headerCnt = 0;

            var reader = new StreamReader(File.OpenRead(tempOriPatListCsvFile));
            bool dataStart = false;
            var regexHeader = @".*Pattern.*Timeset.*File\s+Versions.*";
            while (!reader.EndOfStream)
            {
                var line = reader.ReadLine().Replace("\"", "");
                int valueCnt = Regex.Matches(line, ",").Count + 1;
                if (headerCnt > 0)
                    for (int iChr = 0; iChr < (headerCnt - valueCnt); iChr++)
                    {
                        line += ",";
                    }
                var values = line.Split(',');
                if (dataStart == false && Regex.IsMatch(line, regexHeader, RegexOptions.IgnoreCase)) //the 1st line, Header
                {
                    headerCnt = values.Count();
                    for (int item = 0; item < values.Count(); item++)
                    {
                        if (Regex.IsMatch(values[item], @"^\s*\#\s*$", RegexOptions.IgnoreCase)) idx_idx = item;
                        if (Regex.IsMatch(values[item], @"^Pattern$", RegexOptions.IgnoreCase)) idx_pattern = item;
                        if (Regex.IsMatch(values[item], @"Latest\s+Version", RegexOptions.IgnoreCase)) idx_latestVersion = item;
                        if (Regex.IsMatch(values[item], @"USE.*No.*Use", RegexOptions.IgnoreCase)) idx_useNoUse = item;
                        if (Regex.IsMatch(values[item], @"DRI", RegexOptions.IgnoreCase)) idx_dRI = item;
                        if (Regex.IsMatch(values[item], @"Release\s+Date", RegexOptions.IgnoreCase)) idx_releaseDate = item;
                        if (Regex.IsMatch(values[item], @"Release\s+Note", RegexOptions.IgnoreCase)) idx_releaseNote = item;
                        if (Regex.IsMatch(values[item], @"Radar", RegexOptions.IgnoreCase)) idx_radarNum = item;
                        if (Regex.IsMatch(values[item], @"Org", RegexOptions.IgnoreCase)) idx_org = item;
                        if (Regex.IsMatch(values[item], @"Type\s+Spec", RegexOptions.IgnoreCase)) idx_typeSpec = item;
                        if (Regex.IsMatch(values[item], @"Timeset\s+Latest", RegexOptions.IgnoreCase)) idx_timesetLatest = item;
                        if (Regex.IsMatch(values[item], @"File\s+Versions", RegexOptions.IgnoreCase)) idx_fileVersions = item;

                        if (values[item] == "OpCode") idx_opCode = item;
                        if (values[item] == "ScanMode") idx_scanMode = item;
                        if (values[item] == "Halt") idx_halt = item;
                        if (values[item] == "Compilation") idx_compilation = item;
                        if (values[item] == "T/P Category") idx_tpCategory = item;
                        if (values[item] == "HLV") idx_hLV = item;
                    }
                    dataStart = true;
                    continue;
                }
                if (dataStart)
                {
                    if (values[idx_pattern].Length > 10) //key
                    {
                        var oriPatCsvItem = new OriPatListItem();
                        if (idx_idx >= 0) oriPatCsvItem.Idx = values[idx_idx];
                        if (idx_pattern >= 0) oriPatCsvItem.Pattern = values[idx_pattern];
                        if (idx_latestVersion >= 0) oriPatCsvItem.LatestVersion = values[idx_latestVersion];
                        if (idx_releaseDate >= 0) oriPatCsvItem.ReleaseDate = values[idx_releaseDate];
                        if (idx_useNoUse >= 0) oriPatCsvItem.UseNoUse = values[idx_useNoUse];
                        if (idx_dRI >= 0) oriPatCsvItem.DRi = values[idx_dRI];
                        if (idx_releaseNote >= 0) oriPatCsvItem.ReleaseNote = values[idx_releaseNote];
                        if (idx_radarNum >= 0) oriPatCsvItem.RadarNum = values[idx_radarNum];
                        if (idx_org >= 0) oriPatCsvItem.Org = values[idx_org];
                        if (idx_typeSpec >= 0) oriPatCsvItem.TypeSpec = values[idx_typeSpec];
                        if (idx_timesetLatest >= 0)
                            oriPatCsvItem.TimesetLatest = values[idx_timesetLatest].ToUpper() == "N/A" ? "NA" : (values[idx_timesetLatest].Split('/').Last().ToUpper());
                        if (dicTimeSetVersion.ContainsKey(oriPatCsvItem.TimesetLatest.Replace(".TXT", "")))
                        {
                            var key = oriPatCsvItem.TimesetLatest.Replace(".TXT", "");
                            oriPatCsvItem.TimesetLatest = oriPatCsvItem.TimesetLatest.Replace(".TXT", "") + "_" +
                                                           dicTimeSetVersion[key].Version + ".TXT";
                            oriPatCsvItem.OriTimeMod = dicTimeSetVersion[key].TimeMod.ToUpper();
                        }

                        if (idx_fileVersions >= 0) oriPatCsvItem.FileVersions = values[idx_fileVersions];
                        if (idx_opCode >= 0) oriPatCsvItem.OpCode = values[idx_opCode];
                        if (idx_scanMode >= 0) oriPatCsvItem.ScanMode = values[idx_scanMode];
                        if (idx_halt >= 0) oriPatCsvItem.Halt = values[idx_halt];
                        if (idx_compilation >= 0) oriPatCsvItem.Compilation = values[idx_compilation];
                        if (idx_tpCategory >= 0) oriPatCsvItem.TpCategory = values[idx_tpCategory];
                        if (idx_hLV >= 0) oriPatCsvItem.HLv = values[idx_hLV];
                        string patKeyName = values[idx_pattern];//+ "#" + values[idx_fileVersions].Split('/').Last().ToUpper().Replace(".ATP", "").Replace(".GZ", "");
                        //string patKeyName = values[idx_pattern] + "#" + values[idx_fileVersions].Split('/').Last().ToUpper().Replace(".ATP", "").Replace(".GZ", "");
                        if (!oriPatCsvList.ContainsKey(patKeyName))
                        {
                            oriPatCsvList.Add(patKeyName, oriPatCsvItem);
                        }
                        else
                        {
                            MessageBox.Show(@"Duplicated Pattern in List -> " + values[idx_pattern], @"Pattern Duplicated", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            reader.Close();
            #endregion

            return oriPatCsvList;
        }
    }
}