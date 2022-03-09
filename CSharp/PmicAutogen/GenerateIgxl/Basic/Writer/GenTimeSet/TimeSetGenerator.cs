using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using AutomationCommon.EpplusErrorReport;
using IgxlData.Others;
using IgxlData.Others.MultiTimeSet;
using IgxlData.Others.PatternListCsvFile;
using PmicAutogen.Local.Const;
using LocalTimeSetBasic = IgxlData.Others.MultiTimeSet.ComTimeSetBasic;
using LocalTimingRow = IgxlData.IgxlBase.TimingRow;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenTimeSet
{
    public class TimeSetGenerator
    {
        public List<ComTimeSetBasicSheet> GenerateFlow(List<PatternListCsvRow> patternListCsvRows, string timeSetPath,
            string dirTimings)
        {
            _srcTimeSetsPathList = new List<string>();

            var tSetFileList = GetTSetFileList(patternListCsvRows);

            GenTSetFile(tSetFileList, timeSetPath, dirTimings);

            return _comTimeSetBasicSheets;
        }

        private List<string> GetTSetFileList(List<PatternListCsvRow> patternListCsvRows)
        {
            var tSetFileList = new List<string>();
            foreach (var pattern in patternListCsvRows)
            {
                var tSet = pattern.ActualTimeSetVersion;
                if (!tSetFileList.Contains(tSet) && !IgnoredFileName(tSet))
                    tSetFileList.Add(tSet);
            }

            return tSetFileList;
        }

        private void GenTSetFile(List<string> tSetFileList, string timeSetPath, string dirTimings)
        {
            try
            {
                if (Directory.Exists(timeSetPath))
                {
                    foreach (var fileName in tSetFileList)
                    {
                        if (fileName == string.Empty)
                            continue;
                        var completeFileName = Path.ChangeExtension(fileName, @".txt");
                        if (completeFileName != null)
                        {
                            var fileFullPath = Path.Combine(timeSetPath, completeFileName);
                            if (!Directory.Exists(dirTimings))
                                Directory.CreateDirectory(dirTimings);
                            if (File.Exists(fileFullPath))
                            {
                                var tarFile = Path.Combine(dirTimings, completeFileName);
                                CopyFile(fileFullPath, tarFile, true);
                            }
                            else
                            {
                                MissFileMessage(MissingFile, fileFullPath);
                            }
                        }
                    }

                    _comTimeSetBasicSheets = ReadTimeSetTxt1P4(_srcTimeSetsPathList);
                }
            }
            catch (Exception e)
            {
                throw new Exception("Error occurred in generating time set " + e.Message);
            }
        }

        private bool IgnoredFileName(string fileName)
        {
            if (fileName.ToUpper().Equals(Na)) return true;

            return false;
        }

        private void CopyFile(string srcFile, string tarFile, bool overwrite)
        {
            File.Copy(srcFile, tarFile, overwrite);
            _srcTimeSetsPathList.Add(tarFile);
        }

        public List<ComTimeSetBasicSheet> ReadTimeSetTxt1P4(List<string> timeSetPathList)
        {
            var comTimeSetBasicSheets = new List<ComTimeSetBasicSheet>();
            const string lStrTimeModePattern = @"Timing Mode:[\t](?<str>\w*)";
            const string lStrMasterTsPattern = @"Master TimeSet Name:[\t](?<str>\w*)";
            const string lStrTimeDomainPattern = @"Time Domain:[\t](?<str>\w*)";
            const string lStrHeader = @"Time Set[\t]Period";

            foreach (var timeSetPath in timeSetPathList)
            {
                if (!File.Exists(timeSetPath))
                    continue;

                var lines = File.ReadAllLines(timeSetPath);

                var lStrStrobe = "";

                var lStrTimeMode = Regex.Match(lines[2], lStrTimeModePattern, RegexOptions.IgnoreCase).Groups["str"]
                    .ToString();
                var lStrMasterTs = Regex.Match(lines[2], lStrMasterTsPattern, RegexOptions.IgnoreCase).Groups["str"]
                    .ToString();
                var lStrTimeDomain = Regex.Match(lines[3], lStrTimeDomainPattern, RegexOptions.IgnoreCase).Groups["str"]
                    .ToString();

                var timeRowConverter = Converter(lines[0]);
                var sheetName = Regex.Replace(Path.GetFileName(timeSetPath), ".txt", "", RegexOptions.IgnoreCase);
                var timeSetBasicSheet =
                    new ComTimeSetBasicSheet(sheetName, lStrTimeMode, lStrMasterTs, lStrTimeDomain, lStrStrobe);
                var timeSets = new Dictionary<string, LocalTimeSetBasic>();
                var lIStartRowNum = 4;
                for (var i = lIStartRowNum; i < lines.Length; i++)
                    if (Regex.IsMatch(lines[i], lStrHeader))
                    {
                        lIStartRowNum = i + 1;
                        break;
                    }

                for (var i = lIStartRowNum; i < lines.Length; i++)
                {
                    LocalTimingRow timingRow;
                    var tokens = lines[i].Split('\t');

                    if (tokens.Length < 5)
                    {
                        if (!lines[i].Contains("=") && (lines[i].Contains("*") || lines[i].Contains(@"/"))
                        ) //ignore comment line
                            continue;
                        if (!lines[i].Split(new[] {'\t'}, StringSplitOptions.RemoveEmptyEntries).Any())
                            continue;
                        if (lines[i].Contains("="))
                        {
                            var spt = lines[i].Split('=');
                            var varTok = spt[0].Trim(); // ^_ need replace 
                            varTok = Regex.Replace(varTok, @"^_", "", RegexOptions.IgnoreCase);
                            var valueTok = spt[1].Trim();

                            //HardIpUtilityMain.ResetUtilities();
                            var value = DataConvertor.ConvertUnits(valueTok);
                            //Use HardIp Function
                            double dValue;
                            var isNumOk = double.TryParse(value, out dValue);
                            if (isNumOk)
                                foreach (var subTsb in timeSets)
                                    if (subTsb.Value.SubContextVariable.Contains(varTok))
                                        //which means the variable appear on the comment is use by this time set
                                        if (!subTsb.Value.SubCommentVariable.ContainsKey(varTok))
                                            subTsb.Value.SubCommentVariable.Add(varTok, dValue);
                        }
                        else //just write variable name, and no equal char (ex: AAA, BBB )
                        {
                            var varTok = Regex.Replace(lines[i], @"\s+", "");
                            var dValue = -1e9;
                            foreach (var subTsb in timeSets)
                                if (subTsb.Value.SubContextVariable.Contains(varTok))
                                    //which means the variable appear on the comment is use by this time set
                                    if (!subTsb.Value.SubCommentVariable.ContainsKey(varTok))
                                        subTsb.Value.SubCommentVariable.Add(varTok, dValue);
                        }

                        continue;
                    }

                    var lStrTimeSet = tokens[1];
                    var lStrClockPeriod = tokens[2];
                    List<string> contextVar;
                    Dictionary<string, double> shiftFreqVar;
                    //ReadTimeRow1P4() add argument _contextVar for read equation base variable
                    if (!ReadTimeRow(tokens, timeRowConverter, out timingRow, out contextVar, out shiftFreqVar)) break;

                    if (timeSets.ContainsKey(lStrTimeSet))
                    {
                        timeSets[lStrTimeSet].AddTimingRow(timingRow);
                        foreach (var varTmp in contextVar)
                            if (!timeSets[lStrTimeSet].SubContextVariable.Contains(varTmp))
                                timeSets[lStrTimeSet].SubContextVariable.Add(varTmp);

                        foreach (var dicPair in shiftFreqVar)
                            if (!timeSets[lStrTimeSet].ShiftInReserve.ContainsKey(dicPair.Key))
                                timeSets[lStrTimeSet].ShiftInReserve.Add(dicPair.Key, dicPair.Value);
                    }
                    else
                    {
                        var timeSetBasic = new LocalTimeSetBasic();
                        timeSetBasic.Name = lStrTimeSet;
                        timeSetBasic.CyclePeriod = lStrClockPeriod;
                        timeSetBasic.AddTimingRow(timingRow);

                        foreach (var varTmp in contextVar)
                            timeSetBasic.SubContextVariable.Add(varTmp);

                        foreach (var dicPair in shiftFreqVar)
                            timeSetBasic.ShiftInReserve.Add(dicPair.Key, dicPair.Value);

                        timeSets.Add(lStrTimeSet, timeSetBasic);
                    }
                }

                CheckMissingEquationBaseVar(sheetName, timeSets);

                foreach (var keyValuePair in timeSets) timeSetBasicSheet.AddTimeSet(keyValuePair.Value);

                comTimeSetBasicSheets.Add(timeSetBasicSheet);
            }

            return comTimeSetBasicSheets;
        }

        private bool ReadTimeRow(string[] line, ITimeRowConverter converter, out LocalTimingRow row,
            out List<string> subContextVar, out Dictionary<string, double> shiftInReserveVar)
        {
            row = new LocalTimingRow();
            subContextVar = new List<string>();
            shiftInReserveVar = new Dictionary<string, double>();

            if (line.Length <= 15) // < 17)
                return false;

            row = converter.ConvertTimeRow(line);
            GetContextVariable(row.PinGrpClockPeriod, ref subContextVar);
            GetContextVariable(row.DriveOn, ref subContextVar);
            GetContextVariable(row.DriveData, ref subContextVar);
            GetContextVariable(row.DriveReturn, ref subContextVar);
            GetContextVariable(row.DriveOff, ref subContextVar);
            GetContextVariable(row.CompareOpen, ref subContextVar);
            GetContextVariable(row.CompareClose, ref subContextVar);

            var cyclePeriod = line[2];
            if (cyclePeriod != "")
            {
                if (IsContextVariable(cyclePeriod))
                {
                    GetContextVariable(cyclePeriod, ref subContextVar);
                }
                else
                {
                    var periodVal = GetEgValueInDecimal(cyclePeriod);

                    if (periodVal != (decimal) 0.0 && row.DriveData != "") //check D1
                    {
                        var d1Val = GetEgValueInDecimal(row.DriveData);
                        var tRatio = d1Val / periodVal;
                        //string tRoundedRatio = Math.Round(tRatio, 2).ToString();
                        tRatio = Math.Round(tRatio, 2);

                        if (Regex.IsMatch(row.DataFmt, "RL", RegexOptions.IgnoreCase)) //if (tRatio != (decimal)0)
                            if (!shiftInReserveVar.ContainsKey(SpecFormat.GenAcSpecSymbol(TimeSetConst.ClockS)))
                                shiftInReserveVar.Add(SpecFormat.GenAcSpecSymbol(TimeSetConst.ClockS), (double) tRatio);
                    }

                    if (periodVal != (decimal) 0.0 && row.CompareOpen != "") //check C0
                    {
                        var c0Val = GetEgValueInDecimal(row.CompareOpen);
                        var tRatio = c0Val / periodVal;
                        //string tRoundedRatio = Math.Round(tRatio, 2).ToString();
                        tRatio = Math.Round(tRatio, 2);

                        if (tRatio != 0)
                            if (!shiftInReserveVar.ContainsKey(SpecFormat.GenAcSpecSymbol(TimeSetConst.Strobe)))
                                shiftInReserveVar.Add(SpecFormat.GenAcSpecSymbol(TimeSetConst.Strobe), (double) tRatio);
                    }
                }
            }

            return true;
        }

        private void GetContextVariable(string cell, ref List<string> subContextVar)
        {
            //cell context example:
            //=_RT_CLK32768_Freq_GLB 
            //=(1/_TCK_Freq_VAR)
            //=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR+_Strobe_VAR
            //=_Cycle_S_VAR+0.7/_ShiftIn_Freq_VAR+_Clock_E_VAR

            var matches = Regex.Matches(cell, @"_(?<var>[\d|\w]+)");
            foreach (Match match in matches)
            {
                var contextVar = match.Groups["var"].ToString();
                if (contextVar != "" && !subContextVar.Contains(contextVar))
                    subContextVar.Add(contextVar);
            }
        }

        private bool IsContextVariable(string cell)
        {
            //cell context example:
            //=_RT_CLK32768_Freq_GLB 
            //=(1/_TCK_Freq_VAR)
            //=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR+_Strobe_VAR
            //=_Cycle_S_VAR+0.7/_ShiftIn_Freq_VAR+_Clock_E_VAR    
            return Regex.IsMatch(cell, @"_(?<var>[\d|\w]+)");
        }

        private void MissFileMessage(string pErrorType, string fileFullPath)
        {
            var message = string.Format("Can't find this file: {0}. ", fileFullPath);
            var caption = pErrorType;
            var buttons = MessageBoxButtons.OK;
            MessageBox.Show(message, caption, buttons);
        }

        private void CheckMissingEquationBaseVar(string sheetName, Dictionary<string, LocalTimeSetBasic> timeSets)
        {
            foreach (var tSetDataPair in timeSets)
            {
                var contextVars = tSetDataPair.Value.SubContextVariable;
                var commentVarsDict = tSetDataPair.Value.SubCommentVariable;

                //Check Rule1. if comment doesn't contains value
                foreach (var commentPair in commentVarsDict)
                    if (commentPair.Value == EqnBaseInitValue
                    ) //ConEqnBaseInitValue: -1e9, represent var not assigned value
                    {
                        var errMsg =
                            string.Format(
                                "Equation base variable '{0}' used in Time Set file {1} is not assigned an initial value",
                                commentPair.Key, sheetName);
                        EpplusErrorManager.AddError(BasicErrorType.FormatWarning.ToString(), ErrorLevel.Error,
                            tSetDataPair.Value.Name, 1, errMsg, commentPair.Key);
                        //Response.Report(string.Format("Equation base variable '{0}' used in Time Set file {1} is not assigned an initial value", _commentPair.Key, sheetName), 73, MessageLevel.Error);
                    }

                //Check Rule2. use context vars as base, check if comment vars are not equal to context
                var commentVars = commentVarsDict.Keys.ToList();
                foreach (var contextVar in contextVars)
                    if (!commentVars.Contains(contextVar))
                    {
                        var errMsg =
                            string.Format(
                                "Equation base variable '{0}' used in the context of Time Set file {1} is not assigned value in comment",
                                contextVar, sheetName);
                        EpplusErrorManager.AddError(BasicErrorType.FormatWarning.ToString(), ErrorLevel.Error,
                            tSetDataPair.Value.Name, 1, errMsg, contextVar);
                    }
            }
        }

        private ITimeRowConverter Converter(string header)
        {
            if (Regex.IsMatch(header, @"DTTimesetBasicSheet,version=2.3", RegexOptions.IgnoreCase))
                return new TimeRow2P3();
            if (Regex.IsMatch(header, @"DFF\s1.4\sTime\sSets\s[(]Basic[)]", RegexOptions.IgnoreCase))
                return new TimeRow1P4();

            return new TimeRow1P4();
        }

        #region Field

        private const double EqnBaseInitValue = -1e9;
        private const string MissingFile = "Missing File:";
        private const string Na = "NA";
        private List<string> _srcTimeSetsPathList;
        private List<ComTimeSetBasicSheet> _comTimeSetBasicSheets = new List<ComTimeSetBasicSheet>();

        #endregion


        #region Change Scan Time Set as equation based

        public bool IsMatch(string source, string patten, bool ignoreUnderLine = false)
        {
            var str1 = Normalization(source);
            var str2 = patten;

            if (ignoreUnderLine)
            {
                str1 = str1.Replace(' ', '_');
                str2 = str2.Replace(' ', '_');
            }

            return Regex.IsMatch(str1, str2, RegexOptions.IgnoreCase);
        }

        public string Normalization(string source)
        {
            var result = source.Trim();

            result = ReplaceEnter(result);

            result = ReplaceDoubleBlank(result);

            return result;
        }

        public string ReplaceDoubleBlank(string result)
        {
            var lStrResult = result;
            do
            {
                lStrResult = lStrResult.Replace("  ", " ");
            } while (lStrResult.IndexOf("  ", StringComparison.Ordinal) >= 0);

            return lStrResult;
        }

        public string ReplaceEnter(string result)
        {
            return result.Replace("\n", " ");
        }

        private decimal GetEgValueInDecimal(string inStr)
        {
            var dVal = inStr.Contains("E")
                ? Convert.ToDecimal(decimal.Parse(inStr, NumberStyles.Float))
                : decimal.Parse(inStr);
            return dVal;
        }

        #endregion
    }
}