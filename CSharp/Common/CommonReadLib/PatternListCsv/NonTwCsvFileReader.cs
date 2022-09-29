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
            int idxIdx = -1;
            int idxPattern = -1;
            int idxLatestVersion = -1;
            int idxReleaseDate = -1;
            int idxUseNoUse = -1;
            int idxDRi = -1;
            int idxReleaseNote = -1;
            int idxRadarNum = -1;
            int idxOrg = -1;
            int idxTypeSpec = -1;
            int idxTimesetLatest = -1;
            int idxFileVersions = -1;
            int idxOpCode = -1;
            int idxScanMode = -1;
            int idxHalt = -1;
            int idxCompilation = -1;
            int idxTpCategory = -1;
            int idxHLv = -1;
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
                        if (Regex.IsMatch(values[item], @"^\s*\#\s*$", RegexOptions.IgnoreCase)) idxIdx = item;
                        if (Regex.IsMatch(values[item], @"^Pattern$", RegexOptions.IgnoreCase)) idxPattern = item;
                        if (Regex.IsMatch(values[item], @"Latest\s+Version", RegexOptions.IgnoreCase)) idxLatestVersion = item;
                        if (Regex.IsMatch(values[item], @"USE.*No.*Use", RegexOptions.IgnoreCase)) idxUseNoUse = item;
                        if (Regex.IsMatch(values[item], @"DRI", RegexOptions.IgnoreCase)) idxDRi = item;
                        if (Regex.IsMatch(values[item], @"Release\s+Date", RegexOptions.IgnoreCase)) idxReleaseDate = item;
                        if (Regex.IsMatch(values[item], @"Release\s+Note", RegexOptions.IgnoreCase)) idxReleaseNote = item;
                        if (Regex.IsMatch(values[item], @"Radar", RegexOptions.IgnoreCase)) idxRadarNum = item;
                        if (Regex.IsMatch(values[item], @"Org", RegexOptions.IgnoreCase)) idxOrg = item;
                        if (Regex.IsMatch(values[item], @"Type\s+Spec", RegexOptions.IgnoreCase)) idxTypeSpec = item;
                        if (Regex.IsMatch(values[item], @"Timeset\s+Latest", RegexOptions.IgnoreCase)) idxTimesetLatest = item;
                        if (Regex.IsMatch(values[item], @"File\s+Versions", RegexOptions.IgnoreCase)) idxFileVersions = item;

                        if (values[item] == "OpCode") idxOpCode = item;
                        if (values[item] == "ScanMode") idxScanMode = item;
                        if (values[item] == "Halt") idxHalt = item;
                        if (values[item] == "Compilation") idxCompilation = item;
                        if (values[item] == "T/P Category") idxTpCategory = item;
                        if (values[item] == "HLV") idxHLv = item;
                    }
                    dataStart = true;
                    continue;
                }
                if (dataStart)
                {
                    if (values[idxPattern].Length > 10) //key
                    {
                        var oriPatCsvItem = new OriPatListItem();
                        if (idxIdx >= 0) oriPatCsvItem.Idx = values[idxIdx];
                        if (idxPattern >= 0) oriPatCsvItem.Pattern = values[idxPattern];
                        if (idxLatestVersion >= 0) oriPatCsvItem.LatestVersion = values[idxLatestVersion];
                        if (idxReleaseDate >= 0) oriPatCsvItem.ReleaseDate = values[idxReleaseDate];
                        if (idxUseNoUse >= 0) oriPatCsvItem.UseNoUse = values[idxUseNoUse];
                        if (idxDRi >= 0) oriPatCsvItem.DRi = values[idxDRi];
                        if (idxReleaseNote >= 0) oriPatCsvItem.ReleaseNote = values[idxReleaseNote];
                        if (idxRadarNum >= 0) oriPatCsvItem.RadarNum = values[idxRadarNum];
                        if (idxOrg >= 0) oriPatCsvItem.Org = values[idxOrg];
                        if (idxTypeSpec >= 0) oriPatCsvItem.TypeSpec = values[idxTypeSpec];
                        if (idxTimesetLatest >= 0)
                            oriPatCsvItem.TimesetLatest = values[idxTimesetLatest].ToUpper() == "N/A" ? "NA" : (values[idxTimesetLatest].Split('/').Last().ToUpper());
                        if (dicTimeSetVersion.ContainsKey(oriPatCsvItem.TimesetLatest.Replace(".TXT", "")))
                        {
                            var key = oriPatCsvItem.TimesetLatest.Replace(".TXT", "");
                            oriPatCsvItem.TimesetLatest = oriPatCsvItem.TimesetLatest.Replace(".TXT", "") + "_" +
                                                           dicTimeSetVersion[key].Version + ".TXT";
                            oriPatCsvItem.OriTimeMod = dicTimeSetVersion[key].TimeMod.ToUpper();
                        }

                        if (idxFileVersions >= 0) oriPatCsvItem.FileVersions = values[idxFileVersions];
                        if (idxOpCode >= 0) oriPatCsvItem.OpCode = values[idxOpCode];
                        if (idxScanMode >= 0) oriPatCsvItem.ScanMode = values[idxScanMode];
                        if (idxHalt >= 0) oriPatCsvItem.Halt = values[idxHalt];
                        if (idxCompilation >= 0) oriPatCsvItem.Compilation = values[idxCompilation];
                        if (idxTpCategory >= 0) oriPatCsvItem.TpCategory = values[idxTpCategory];
                        if (idxHLv >= 0) oriPatCsvItem.HLv = values[idxHLv];
                        string patKeyName = values[idxPattern];//+ "#" + values[idx_fileVersions].Split('/').Last().ToUpper().Replace(".ATP", "").Replace(".GZ", "");
                        //string patKeyName = values[idx_pattern] + "#" + values[idx_fileVersions].Split('/').Last().ToUpper().Replace(".ATP", "").Replace(".GZ", "");
                        if (!oriPatCsvList.ContainsKey(patKeyName))
                        {
                            oriPatCsvList.Add(patKeyName, oriPatCsvItem);
                        }
                        else
                        {
                            MessageBox.Show(@"Duplicated Pattern in List -> " + values[idxPattern], @"Pattern Duplicated", MessageBoxButtons.OK, MessageBoxIcon.Error);
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