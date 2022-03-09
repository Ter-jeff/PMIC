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
// 2022-2-14  Terry Zhang     #312       Initial creation
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
    public class BL_ModuleParser
    {

        private List<string> _lstmodules = null;
        private StreamReader _modulereader = null;

        public BL_ModuleParser(List<string> p_lstModules)
        {
            this._lstmodules = p_lstModules;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public List<ReportItem> getReportItem()
        {
            List<ReportItem> l_lstRtn = new List<ReportItem>();

            foreach (string l_strModule in this._lstmodules)
            {
                this._modulereader = new StreamReader(l_strModule);

                string l_strFileName = Path.GetFileName(l_strModule);
                FunctionItem l_FunctionItem = null;
                while((l_FunctionItem = this.getOneFunctionContent())!=null)
                {
                    List<ReportItem> l_lstReportItem = this.getReportItemByContent(l_FunctionItem);

                    foreach(ReportItem l_ReportItem in l_lstReportItem)
                    {
                        l_ReportItem.FileName = l_strFileName;
                        l_ReportItem.FunctionName = l_FunctionItem.FunctionName;

                        if (l_ReportItem != null)
                        {
                            if (!Regex.IsMatch(l_ReportItem.Pins, @"""") && !string.IsNullOrEmpty(l_ReportItem.Pins))
                            {
                                l_ReportItem.Comment += @"Pin Name is a variable, not a string.";
                            }

                            ReportItem l_TempItem = this.isExistSameAlarmSetting(l_ReportItem, l_lstRtn);
                            if (l_TempItem == null)
                            {
                                l_lstRtn.Add(l_ReportItem);
                            }
                            else
                            {
                                l_lstRtn.Add(l_TempItem);
                            }
                        }
                        else
                        {
                            //do nothing
                        }
                    }
                    
                }

                this._modulereader.Close();
            }

            return l_lstRtn;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private FunctionItem getOneFunctionContent()
        {
            FunctionItem l_Rtn = null;

            string l_strLine = string.Empty;

            while((l_strLine=this._modulereader.ReadLine())!=null)
            {
                if(RegexManager.CommentRegex.IsMatch(l_strLine))
                {
                    continue;
                }
                else
                {
                    //do nothing
                }

                if (RegexManager.SubStartRegex.IsMatch(l_strLine))
                {
                    l_Rtn = this.getSubFunction(l_strLine);
                    return l_Rtn;
                }
                else if(RegexManager.FunctionStartRegex.IsMatch(l_strLine))
                {
                    l_Rtn = this.getFunction(l_strLine);
                    return l_Rtn;
                }
                else
                {
                    //do nothing
                }
            }
            //throw new Exception("can not find function start, please check it.");

            return l_Rtn;
        }

        private FunctionItem getSubFunction(string p_strLine)
        {
            FunctionItem l_Rtn = null;

            string l_strFunctionName = RegexManager.SubStartRegex.Match(p_strLine).Groups["FunctionName"].Value;
            l_Rtn = new FunctionItem();
            l_Rtn.FunctionName = l_strFunctionName;

            while ((p_strLine = this._modulereader.ReadLine()) != null)
            {
                if (RegexManager.SubEndRegex.IsMatch(p_strLine))
                {
                    return l_Rtn;
                }
                else
                {
                    l_Rtn.lstFunctionContent.Add(p_strLine);
                }
            }
            throw new Exception("can not find function end, please check it.");
        }

        private FunctionItem getFunction(string p_strLine)
        {
            FunctionItem l_Rtn = null;

            string l_strFunctionName = RegexManager.FunctionStartRegex.Match(p_strLine).Groups["FunctionName"].Value;
            l_Rtn = new FunctionItem();
            l_Rtn.FunctionName = l_strFunctionName;

            while ((p_strLine = this._modulereader.ReadLine()) != null)
            {
                if (RegexManager.FunctionEndRegex.IsMatch(p_strLine))
                {
                    return l_Rtn;
                }
                else
                {
                    l_Rtn.lstFunctionContent.Add(p_strLine);
                }
            }
            throw new Exception("can not find function end, please check it.");
        }

        private ReportItem isExistSameAlarmSetting(ReportItem p_ReportItem,List<ReportItem> p_lstReportItem)
        {
            ReportItem l_Rtn = null;

            foreach(ReportItem l_Item in p_lstReportItem)
            {
                if(l_Item.Equals(p_ReportItem))
                {
                    l_Rtn = l_Item;
                    l_Rtn.Comment += "Same control condition appear in same function, please check any \"If / Case\" control rule.";
                    return l_Rtn;
                }
                else
                {
                    //do nothing
                }
            }

            return l_Rtn;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p_lstContent"></param>
        /// <returns></returns>
        private List<ReportItem> getReportItemByContent(FunctionItem p_FunctionItem)
        {

            List<ReportItem> l_lstRtn = new List<ReportItem>();

            ReportItem l_ReportItem = null;
            foreach (string l_strLine in p_FunctionItem.lstFunctionContent)
            {
                string l_strNewLine = Regex.Replace(l_strLine, "'.*", "");

                if (RegexManager.RegexFirst.IsMatch(l_strNewLine))
                {
                    string l_strInstrument = RegexManager.RegexFirst.Match(l_strNewLine).Groups["Instrument"].Value;
                    string l_strAlarmCategory = RegexManager.RegexFirst.Match(l_strNewLine).Groups["AlarmCategory"].Value;
                    string l_strAlarmBehavior = RegexManager.RegexFirst.Match(l_strNewLine).Groups["AlarmBehavior"].Value;
                    l_ReportItem = new ReportItem();

                    l_ReportItem.Instrument = l_strInstrument;
                    l_ReportItem.AlarmCategory = l_strAlarmCategory;
                    l_ReportItem.AlarmBehavior = l_strAlarmBehavior;

                    l_lstRtn.Add(l_ReportItem);
                    l_ReportItem = null;
                }
                else if (RegexManager.RegexSecond.IsMatch(l_strNewLine))
                {
                    string l_strInstrument = RegexManager.RegexSecond.Match(l_strNewLine).Groups["Instrument"].Value;
                    string l_strPins = RegexManager.RegexSecond.Match(l_strNewLine).Groups["Pins"].Value;
                    string l_strAlarmCategory = RegexManager.RegexSecond.Match(l_strNewLine).Groups["AlarmCategory"].Value;
                    string l_strAlarmBehavior = RegexManager.RegexSecond.Match(l_strNewLine).Groups["AlarmBehavior"].Value;
                    l_ReportItem = new ReportItem();
                    l_ReportItem.Pins = l_strPins;
                    l_ReportItem.Instrument = l_strInstrument;
                    l_ReportItem.AlarmCategory = l_strAlarmCategory;
                    l_ReportItem.AlarmBehavior = l_strAlarmBehavior;

                    l_lstRtn.Add(l_ReportItem);
                    l_ReportItem = null;
                }
                else if (RegexManager.RegexThird.IsMatch(l_strNewLine))
                {
                    string l_strInstrument = RegexManager.RegexThird.Match(l_strNewLine).Groups["Instrument"].Value;
                    string l_strPins = RegexManager.RegexThird.Match(l_strNewLine).Groups["Pins"].Value;
                    string l_strAlarmCategory = RegexManager.RegexThird.Match(l_strNewLine).Groups["AlarmCategory"].Value;
                    string l_strAlarmBehavior = RegexManager.RegexThird.Match(l_strNewLine).Groups["AlarmBehavior"].Value;
                    l_ReportItem = new ReportItem();
                    l_ReportItem.Pins = l_strPins;
                    l_ReportItem.Instrument = l_strInstrument;
                    l_ReportItem.AlarmCategory = l_strAlarmCategory;
                    l_ReportItem.AlarmBehavior = l_strAlarmBehavior;

                    l_lstRtn.Add(l_ReportItem);
                    l_ReportItem = null;
                }
                else if (RegexManager.RegexFourth.IsMatch(l_strNewLine))
                {
                    string l_strInstrument = RegexManager.RegexFourth.Match(l_strNewLine).Groups["Instrument"].Value;
                    string l_strPins = RegexManager.RegexFourth.Match(l_strNewLine).Groups["Pins"].Value;

                    l_ReportItem = new ReportItem();
                    l_ReportItem.Pins = l_strPins;
                    l_ReportItem.Instrument = l_strInstrument;

                }
                else if (RegexManager.RegexFifth.IsMatch(l_strNewLine))
                {
                    string l_strInstrument = RegexManager.RegexFifth.Match(l_strNewLine).Groups["Instrument"].Value;
                    string l_strPins = RegexManager.RegexFifth.Match(l_strNewLine).Groups["Pins"].Value;

                    l_ReportItem = new ReportItem();
                    l_ReportItem.Pins = l_strPins;
                    l_ReportItem.Instrument = l_strInstrument;
                }
                else if(l_ReportItem!=null)
                {
                    if(RegexManager.RegexAlarmCatagory.IsMatch(l_strNewLine))
                    {
                        string l_strAlarmCategory = RegexManager.RegexAlarmCatagory.Match(l_strNewLine).Groups["AlarmCategory"].Value;
                        string l_strAlarmBehavior = RegexManager.RegexAlarmCatagory.Match(l_strNewLine).Groups["AlarmBehavior"].Value;

                        l_ReportItem.AlarmCategory = l_strAlarmCategory;
                        l_ReportItem.AlarmBehavior = l_strAlarmBehavior;
                    }
                    else if(Regex.IsMatch(l_strNewLine, "End With"))
                    {
                        if (!string.IsNullOrEmpty(l_ReportItem.AlarmBehavior))
                        {
                            l_lstRtn.Add(l_ReportItem);
                        }
                        else
                        {
                            //do nothing, no alarm setting
                        }    
                        l_ReportItem = null;
                    }
                    else
                    {
                        //do nothing
                    }
                }
                else
                {
                    //do nothing
                }
            }


            return l_lstRtn;
        }
    }
}
