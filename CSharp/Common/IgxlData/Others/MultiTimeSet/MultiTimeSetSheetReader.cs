using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using AutomationCommon.EpplusErrorReport;
using IgxlData.IgxlBase;

namespace IgxlData.Others.MultiTimeSet
{
    public class MultiTimeSetSheetReader
    {
        public const string CycleS = "Cycle_S";
        public const string ClockS = "Clock_S";
        public const string ClockE = "Clock_E";
        public const string Strobe = "Strobe";


        public MultiTimeSetSheets ReadTimeSetTxt1P4(List<string> timeSetPathList, bool isRemoveBackup = false)
        {
            var multiTimeSetSheets = new MultiTimeSetSheets();
            const string lStrTimeModePattern = @"Timing Mode:[\t](?<str>\w*)";
            const string lStrMasterTsPattern = @"Master Timeset Name:[\t](?<str>\w*)";
            const string lStrTimeDomainPattern = @"Time Domain:[\t](?<str>\w*)";
            const string lStrHeader = @"Time Set[\t]Period";
            try
            {
                foreach (string timeSetPath in timeSetPathList)
                {
                    if (!File.Exists(timeSetPath))
                        continue;

                    string[] lines = File.ReadAllLines(timeSetPath);



                    string lStrStrobe = "";

                    string lStrTimeMode = Regex.Match(lines[2], lStrTimeModePattern, RegexOptions.IgnoreCase).Groups["str"].ToString();
                    string lStrMasterTs = Regex.Match(lines[2], lStrMasterTsPattern, RegexOptions.IgnoreCase).Groups["str"].ToString();
                    string lStrTimeDomain = Regex.Match(lines[3], lStrTimeDomainPattern, RegexOptions.IgnoreCase).Groups["str"].ToString();

                    // support 1.4 ,2.3 timing row conveter 20180613 by JN
                    var timeRowConverter = Converter(lines[0]);
                    var sheetName = Regex.Replace(Path.GetFileName(timeSetPath), ".txt", "", RegexOptions.IgnoreCase);
                    var timeSetBasicSheet = new ComTimeSetBasicSheet(sheetName, lStrTimeMode, lStrMasterTs, lStrTimeDomain, lStrStrobe);
                    var timeSetDatas = new Dictionary<string, ComTimeSetBasic>();
                    int lIStartRowNum = 4;
                    for (int i = lIStartRowNum; i < lines.Length; i++)
                    {
                        if (Regex.IsMatch(lines[i], lStrHeader))
                        {
                            lIStartRowNum = i + 1;
                            break;
                        }
                    }

                    string lastTimeSet = "";
                    for (int i = lIStartRowNum; i < lines.Length; i++)
                    {
                        string lStrTimeSet = "";
                        string lStrClockPeriod = "";
                        TimingRow timingRow;
                        string[] datas = lines[i].Split('\t');
                        //2017/5/5 anderson add for TSB support equation base
                        //         original code is if (datas.Length < 2) -> break, now try to catch equation base var in this section

                        if (datas.Length < 5)
                        {
                            if (!lines[i].Contains("=") && (lines[i].Contains("*") || lines[i].Contains(@"/"))) //ignore comment line
                                continue;
                            if (lines[i].Split(new char[] { '\t' }, StringSplitOptions.RemoveEmptyEntries).Count() == 0)
                                continue;
                            if (lines[i].Contains("=")) //	HTOL_Freq_VAR	=1000000, tset1Per	=1000.000*ns
                            {
                                string[] spt = lines[i].Split('=');
                                string varTok = spt[0].Trim(); // ^_ need replace 
                                varTok = Regex.Replace(varTok, @"^_", "", RegexOptions.IgnoreCase);
                                string valueTok = spt[1].Trim();

                                //HardIpUtilityMain.ResetUtilities();
                                string valueSter = DataConvertor.ConvertUnits(valueTok);
                                //Use HardIp Function
                                double dValue;
                                bool isNumOk = double.TryParse(valueSter, out dValue);
                                if (isNumOk)
                                {
                                    //timeSetDatas[_lastTimeSet].SubCommentVariable.Add(varTok, dValue);
                                    //timeSetDatas[_lastTimeSet].IsEqnBase = true;
                                    foreach (KeyValuePair<string, ComTimeSetBasic> subTsb in timeSetDatas)
                                    {
                                        if (subTsb.Value.SubContextVariable.Contains(varTok))
                                        //which means the variable appear on the comment is use by this time set
                                        {
                                            if (!subTsb.Value.SubCommentVariable.ContainsKey(varTok))
                                                subTsb.Value.SubCommentVariable.Add(varTok, dValue);
                                        }
                                    }
                                }
                                else
                                {
                                    //Add error report...  under construct
                                }
                            }
                            else //just write variable name, and no equal char (ex: AAA, BBB )
                            {
                                string varTok = Regex.Replace(lines[i], @"\s+", "");
                                double dValue = -1e9;
                                //timeSetDatas[_lastTimeSet].SubCommentVariable.Add(varTok, dValue);
                                //timeSetDatas[_lastTimeSet].IsEqnBase = true;
                                foreach (KeyValuePair<string, ComTimeSetBasic> subTsb in timeSetDatas)
                                {
                                    if (subTsb.Value.SubContextVariable.Contains(varTok))
                                    //which means the variable appear on the comment is use by this time set
                                    {
                                        if (!subTsb.Value.SubCommentVariable.ContainsKey(varTok))
                                            subTsb.Value.SubCommentVariable.Add(varTok, dValue);
                                    }
                                }

                                //Add error report...  under construct
                                //-> for no initial value
                            }

                            continue;
                        }

                        lStrTimeSet = datas[1];
                        lStrClockPeriod = datas[2];
                        lastTimeSet = lStrTimeSet;
                        List<string> contextVar;
                        Dictionary<string, double> shiftFreqVar;
                        //ReadTimeRow1P4() add argument _contextVar for read equation base variable
                        if (!ReadTimeRow(datas, timeRowConverter, out timingRow, out contextVar, out shiftFreqVar))
                        {
                            break;
                        }

                        if (timeSetDatas.ContainsKey(lStrTimeSet))
                        {
                            timeSetDatas[lStrTimeSet].AddTimingRow(timingRow);
                            foreach (string varTmp in contextVar)
                            {
                                if (!timeSetDatas[lStrTimeSet].SubContextVariable.Contains(varTmp))
                                    timeSetDatas[lStrTimeSet].SubContextVariable.Add(varTmp);
                            }

                            foreach (var dicPair in shiftFreqVar)
                            {
                                if (!timeSetDatas[lStrTimeSet].ShiftInReserve.ContainsKey(dicPair.Key))
                                    timeSetDatas[lStrTimeSet].ShiftInReserve.Add(dicPair.Key, dicPair.Value);
                            }
                        }
                        else
                        {
                            ComTimeSetBasic timeSetBasic = new ComTimeSetBasic();
                            timeSetBasic.Name = lStrTimeSet;
                            timeSetBasic.CyclePeriod = lStrClockPeriod;
                            timeSetBasic.AddTimingRow(timingRow);

                            foreach (string varTmp in contextVar)
                                timeSetBasic.SubContextVariable.Add(varTmp);

                            foreach (var dicPair in shiftFreqVar)
                                timeSetBasic.ShiftInReserve.Add(dicPair.Key, dicPair.Value);

                            timeSetDatas.Add(lStrTimeSet, timeSetBasic);
                        }
                    }

                    //2017/5/17 anderson add: check all timesets contain in one timeset sheet, if comment var count != context var, report error
                    CheckMissingEquationBaseVar(sheetName, timeSetDatas);

                    foreach (KeyValuePair<string, ComTimeSetBasic> keyValuePair in timeSetDatas)
                    {
                        if (string.IsNullOrEmpty(keyValuePair.Value.Name) && isRemoveBackup)
                            break;

                        if (!string.IsNullOrEmpty(keyValuePair.Value.Name))
                            timeSetBasicSheet.AddTimeSet(keyValuePair.Value);

                    }

                    multiTimeSetSheets.AddTimeSetSheet(timeSetBasicSheet);
                }
            }
            catch (Exception e)
            {
            }
            return multiTimeSetSheets;
        }

        private void CheckMissingEquationBaseVar(string sheetName, Dictionary<string, ComTimeSetBasic> timeSetDatas)
        {
            foreach (KeyValuePair<string, ComTimeSetBasic> tSetDataPair in timeSetDatas)
            {
                List<string> contextVars = tSetDataPair.Value.SubContextVariable;
                Dictionary<string, double> commentVarsDict = tSetDataPair.Value.SubCommentVariable;

                //Check Rule1. if comment doesn't contains value
                foreach (KeyValuePair<string, double> commentPair in commentVarsDict)
                {
                    if (commentPair.Value == -1e9)  //ConEqnBaseInitValue: -1e9, represent var not assigned value
                    {
                        string errMsg = string.Format("Equation base variable '{0}' used in Time Set file {1} is not assigned an initial value", commentPair.Key, sheetName);
                        EpplusErrorManager.AddError(BasicErrorType.FormatWarning.ToString(),ErrorLevel.Error, tSetDataPair.Value.Name, 1, errMsg, commentPair.Key);
                    }
                }

                //Check Rule2. use context vars as base, check if comment vars are not equal to context
                List<string> commentVars = commentVarsDict.Keys.ToList();
                foreach (string contextVar in contextVars)
                {
                    if (!commentVars.Contains(contextVar))
                    {
                        string errMsg = string.Format("Equation base variable '{0}' used in the context of Time Set file {1} is not assgined value in comment", contextVar, sheetName);
                        EpplusErrorManager.AddError(BasicErrorType.FormatWarning.ToString(),ErrorLevel.Error, tSetDataPair.Value.Name, 1, errMsg, contextVar);
                    }
                }
            }
        }

        private bool ReadTimeRow(string[] line, ITimeRowConverter converter, out TimingRow row, out List<string> subContextVar, out Dictionary<string, double> shiftInReserveVar)
        {

            row = new TimingRow();
            subContextVar = new List<string>();
            shiftInReserveVar = new Dictionary<string, double>();

            if (line.Length <= 15)// < 17)
            {
                return false;
            }

            row = converter.ConvertTimeRow(line);

            //check var exsist in context
            GetContextVairable(row.PinGrpClockPeriod, ref subContextVar);
            GetContextVairable(row.DriveOn, ref subContextVar);
            GetContextVairable(row.DriveData, ref subContextVar);
            GetContextVairable(row.DriveReturn, ref subContextVar);
            GetContextVairable(row.DriveOff, ref subContextVar);
            GetContextVairable(row.CompareOpen, ref subContextVar);
            GetContextVairable(row.CompareClose, ref subContextVar);

            //get all variable in percetage format for scan equation base used
            try
            {
                string cyclePeriod = line[2];
                if (cyclePeriod != "")
                {
                    if (IsContextVairable(cyclePeriod))
                        GetContextVairable(cyclePeriod, ref subContextVar);
                    else
                    {
                        decimal periodVal = (decimal)-1.0; //get period value
                        periodVal = GetEgValueInDecimal(cyclePeriod);

                        decimal d1Val = (decimal)-1.0;
                        decimal c0Val = (decimal)-1.0;
                        if (periodVal != (decimal)0.0 && row.DriveData != "") //check D1
                        {
                            d1Val = GetEgValueInDecimal(row.DriveData);
                            decimal tRatio = d1Val / periodVal;
                            //string tRoundedRatio = Math.Round(tRatio, 2).ToString();
                            tRatio = Math.Round(tRatio, 2);

                            if (Regex.IsMatch(row.DataFmt, "RL", RegexOptions.IgnoreCase)) //if (tRatio != (decimal)0)
                            {
                                if (!shiftInReserveVar.ContainsKey(SpecFormat.GenAcSpecSymbol(CycleS)))
                                    shiftInReserveVar.Add(SpecFormat.GenAcSpecSymbol(CycleS),
                                        (double)tRatio);
                            }
                        }

                        if (periodVal != (decimal)0.0 && row.CompareOpen != "") //check C0
                        {
                            c0Val = GetEgValueInDecimal(row.CompareOpen);
                            decimal tRatio = c0Val / periodVal;
                            //string tRoundedRatio = Math.Round(tRatio, 2).ToString();
                            tRatio = Math.Round(tRatio, 2);

                            if (tRatio != (decimal)0)
                            {
                                if (!shiftInReserveVar.ContainsKey(SpecFormat.GenAcSpecSymbol(Strobe)))
                                    shiftInReserveVar.Add(SpecFormat.GenAcSpecSymbol(Strobe),
                                        (double)tRatio);
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
            }

            return true;
        }

        private decimal GetEgValueInDecimal(string inStr)
        {
            decimal dVal = (decimal)-1.0;

            if (inStr.Contains("E"))
            {
                dVal = Convert.ToDecimal(Decimal.Parse(inStr, NumberStyles.Float));
            }
            else
            {
                dVal = decimal.Parse(inStr);
            }

            return dVal;
        }

        private void GetContextVairable(string cell, ref List<string> subContextVar)
        {
            //cell context example:
            //=_RT_CLK32768_Freq_GLB 
            //=(1/_TCK_Freq_VAR)
            //=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR+_Strobe_VAR
            //=_Cycle_S_VAR+0.7/_ShiftIn_Freq_VAR+_Clock_E_VAR

            var matches = Regex.Matches(cell, @"_(?<var>[\d|\w]+)");
            foreach (Match match in matches)
            {
                string contextVar = match.Groups["var"].ToString();
                if (contextVar != "" && !subContextVar.Contains(contextVar))
                    subContextVar.Add(contextVar);
            }
        }

        private bool IsContextVairable(string cell)
        {
            //cell context example:
            //=_RT_CLK32768_Freq_GLB 
            //=(1/_TCK_Freq_VAR)
            //=_Cycle_S_VAR+0.1/_ShiftIn_Freq_VAR+_Strobe_VAR
            //=_Cycle_S_VAR+0.7/_ShiftIn_Freq_VAR+_Clock_E_VAR    
            return Regex.IsMatch(cell, @"_(?<var>[\d|\w]+)");
        }

        public static ITimeRowConverter Converter(string header)
        {
            if (Regex.IsMatch(header, @"DTTimesetBasicSheet,version=2.3", RegexOptions.IgnoreCase))
                return new TimeRow2P3();
            if (Regex.IsMatch(header, @"DFF\s1.4\sTime\sSets\s[(]Basic[)]", RegexOptions.IgnoreCase))
                return new TimeRow1P4();

            return new TimeRow1P4();
        }
    }

}
