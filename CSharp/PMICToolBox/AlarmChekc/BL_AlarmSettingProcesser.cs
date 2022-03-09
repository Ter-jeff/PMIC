//------------------------------------------------------------------------------
// Copyright (C) 2021 Teradyne, Inc. All rights reserved.
//
// This document contains proprietary and confidential information of Teradyne,
// Inc. and is tendered subject to the condition that the information (a) be
// retained in confidence and (b) not be used or incorporated in any product
// except with the express written consent of Teradyne, Inc.
//
// <File description paragraph>
//
// Revision History:
// (Place the most recent revision history at the top.)
// Date        Name           Task#           Notes
// 2022-2-23  Terry Zhang     #312       Initial creation
//------------------------------------------------------------------------------ 

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace AlarmChekc
{
    public class BL_AlarmSettingProcesser
    {
        private List<string> _lstmodules = null;
        private StreamReader _modulereader = null;
        private StreamWriter _modulewritter = null;

        public BL_AlarmSettingProcesser(List<string> p_lstModules)
        {
            this._lstmodules = p_lstModules;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_strOutputPath"></param>
        public void GenerateNewSetting(string p_strOutputPath)
        {
            string l_strLine = string.Empty;
            foreach (string l_strModule in this._lstmodules)
            {
                this._modulereader = new StreamReader(l_strModule);
                string l_strFileName = Path.GetFileName(l_strModule);
                string l_strOutputFileName = Path.Combine(p_strOutputPath, l_strFileName);
                List<string> l_lstNewSetting = new List<string>();
                bool l_boolIsNewSetting = false;
                while ((l_strLine = this._modulereader.ReadLine()) != null)
                {

                    if (RegexManager.RegexFirst.IsMatch(l_strLine))
                    {
                        //replace and write to file
                        string l_strInstrument = RegexManager.RegexFirst.Match(l_strLine).Groups["Instrument"].Value;
                        string l_strAlarmCategory = RegexManager.RegexFirst.Match(l_strLine).Groups["AlarmCategory"].Value;
                        string l_strAlarmBehavior = RegexManager.RegexFirst.Match(l_strLine).Groups["AlarmBehavior"].Value;
                        l_strAlarmBehavior = Regex.Replace(l_strAlarmBehavior, @"'.*", "");

                        l_lstNewSetting.Add("'" + l_strLine);
                        l_lstNewSetting.Add("Library_AlarmControl \"" + l_strInstrument + "\"" + ", ," + l_strAlarmCategory + "," + l_strAlarmBehavior);
                        l_boolIsNewSetting = true;
                    }
                    else if (RegexManager.RegexSecond.IsMatch(l_strLine))
                    {
                        //replace and write to file
                        string l_strInstrument = RegexManager.RegexSecond.Match(l_strLine).Groups["Instrument"].Value;
                        string l_strPins = RegexManager.RegexSecond.Match(l_strLine).Groups["Pins"].Value;
                        string l_strAlarmCategory = RegexManager.RegexSecond.Match(l_strLine).Groups["AlarmCategory"].Value;
                        string l_strAlarmBehavior = RegexManager.RegexSecond.Match(l_strLine).Groups["AlarmBehavior"].Value;
                        l_strAlarmBehavior = Regex.Replace(l_strAlarmBehavior, @"'.*", "");

                        l_lstNewSetting.Add("'" + l_strLine);
                        l_lstNewSetting.Add("Library_AlarmControl \"" + l_strInstrument + "\"" + "," + l_strPins + "," + l_strAlarmCategory + "," + l_strAlarmBehavior);
                        l_boolIsNewSetting = true;
                    }
                    else if (RegexManager.RegexThird.IsMatch(l_strLine))
                    {
                        //replace and write to file
                        string l_strInstrument = RegexManager.RegexThird.Match(l_strLine).Groups["Instrument"].Value;
                        string l_strPins = RegexManager.RegexThird.Match(l_strLine).Groups["Pins"].Value;
                        string l_strAlarmCategory = RegexManager.RegexThird.Match(l_strLine).Groups["AlarmCategory"].Value;
                        string l_strAlarmBehavior = RegexManager.RegexThird.Match(l_strLine).Groups["AlarmBehavior"].Value;
                        l_strAlarmBehavior = Regex.Replace(l_strAlarmBehavior, @"'.*", "");

                        l_lstNewSetting.Add("'" + l_strLine);
                        l_lstNewSetting.Add("Library_AlarmControl \"" + l_strInstrument + "\"" + "," + l_strPins + "," + l_strAlarmCategory + "," + l_strAlarmBehavior);
                        l_boolIsNewSetting = true;
                    }
                    else if (RegexManager.RegexFourth.IsMatch(l_strLine))
                    {
                        string l_strInstrument = RegexManager.RegexFourth.Match(l_strLine).Groups["Instrument"].Value;
                        string l_strPins = RegexManager.RegexFourth.Match(l_strLine).Groups["Pins"].Value;
                        string l_strAlarmCategory =string.Empty;
                        string l_strAlarmBehavior =string.Empty;
                        string l_strEndLine = string.Empty;
                        l_lstNewSetting.Add(l_strLine);
                        while ((l_strEndLine=this._modulereader.ReadLine())!=null)
                        {
                            if (RegexManager.RegexAlarmCatagory.IsMatch(l_strEndLine))
                            {
                                l_strAlarmCategory = RegexManager.RegexAlarmCatagory.Match(l_strEndLine).Groups["AlarmCategory"].Value;
                                l_strAlarmBehavior = RegexManager.RegexAlarmCatagory.Match(l_strEndLine).Groups["AlarmBehavior"].Value;
                                l_strAlarmBehavior = Regex.Replace(l_strAlarmBehavior, @"'.*", "");
                                l_lstNewSetting.Add("'" + l_strEndLine);
                                l_boolIsNewSetting = true;
                            }
                            else if (Regex.IsMatch(l_strEndLine, "End With"))
                            {
                                l_lstNewSetting.Add(l_strEndLine);
                                if (l_boolIsNewSetting == true)
                                {
                                    l_lstNewSetting.Add("Library_AlarmControl \"" + l_strInstrument + "\"" + "," + l_strPins + "," + l_strAlarmCategory + "," + l_strAlarmBehavior);
                                }
                                else
                                {
                                    //do nothing
                                }
                                break;
                            }
                            else
                            {
                                l_lstNewSetting.Add(l_strEndLine);
                            }
                        }

                    }
                    else if (RegexManager.RegexFifth.IsMatch(l_strLine))
                    {
                        string l_strInstrument = RegexManager.RegexFifth.Match(l_strLine).Groups["Instrument"].Value;
                        string l_strPins = RegexManager.RegexFifth.Match(l_strLine).Groups["Pins"].Value;
                        string l_strAlarmCategory = string.Empty;
                        string l_strAlarmBehavior = string.Empty;
                        string l_strEndLine = string.Empty;
                        l_lstNewSetting.Add(l_strLine);
                        while ((l_strEndLine = this._modulereader.ReadLine()) != null)
                        {
                            if (RegexManager.RegexAlarmCatagory.IsMatch(l_strEndLine))
                            {
                                l_strAlarmCategory = RegexManager.RegexAlarmCatagory.Match(l_strEndLine).Groups["AlarmCategory"].Value;
                                l_strAlarmBehavior = RegexManager.RegexAlarmCatagory.Match(l_strEndLine).Groups["AlarmBehavior"].Value;
                                l_strAlarmBehavior = Regex.Replace(l_strAlarmBehavior, @"'.*", "");
                                l_lstNewSetting.Add("'" + l_strEndLine);
                                l_boolIsNewSetting = true;
                            }
                            else if (Regex.IsMatch(l_strEndLine, "End With"))
                            {
                                l_lstNewSetting.Add(l_strEndLine);
                                if (!string.IsNullOrEmpty(l_strAlarmBehavior))
                                {
                                    l_lstNewSetting.Add("Library_AlarmControl \"" + l_strInstrument + "\"" + "," + l_strPins + "," + l_strAlarmCategory + "," + l_strAlarmBehavior);
                                }
                                else
                                {
                                    //do nothing
                                }
                                break;
                            }
                            else
                            {
                                l_lstNewSetting.Add(l_strEndLine);
                            }
                        }
                    }
                    else
                    {
                        l_lstNewSetting.Add(l_strLine);
                    }
                }
                this._modulereader.Close();

                if (l_boolIsNewSetting == true)
                {
                    this._modulewritter = new StreamWriter(l_strOutputFileName);
                    foreach (string l_strContent in l_lstNewSetting)
                    {
                        this._modulewritter.WriteLine(l_strContent);
                    }
                    this._modulewritter.Close();
                }
                else
                {
                    //do nothing
                }
            }
        }
    }
}
