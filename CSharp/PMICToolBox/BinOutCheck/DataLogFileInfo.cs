//------------------------------------------------------------------------------
// Copyright (C) 2019 Teradyne, Inc. All rights reserved.
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
//
// 2022-02-21  Steven Chen    #320            Support No FlowHeader in Binoutcheck
// 2022-01-10  Bruce          #295	          Support key name '0-9'
// 2022-01-10  Bruce          #294	          Support end key word 'Stop'
// 2021-12-21  Bruce          #270	          Compare the datalog to check NoBinOut
//------------------------------------------------------------------------------ 

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BinOutCheck
{
    public class DataLogFileInfo
    {
        public Dictionary<string, List<string>> TestFlows { get; set; } = new Dictionary<string, List<string>>();
        // 2022-01-10  Bruce          #295	          Support key name '0-9' chg start
        //private Regex _RegFlowStart = new Regex(@"^\s*(?<flowname>[a-z_]+)\s+(Start)\s*$", RegexOptions.IgnoreCase);
        //// 2022-01-10  Bruce          #294	          Support end key word 'Stop' chg start
        ////private Regex _RegFlowEnd = new Regex(@"^\s*(?<flowname>[a-z_]+)\s+(End)\s*$", RegexOptions.IgnoreCase);
        //private Regex _RegFlowEnd = new Regex(@"^\s*(?<flowname>[a-z_]+)\s+((End)|(Stop))\s*$", RegexOptions.IgnoreCase);
        //// 2022-01-10  Bruce          #294	          Support end key word 'Stop' chg end
        //private Regex _RegTestItem = new Regex(@"^\s*\<(?<testname>[a-z_]+)\>s*$", RegexOptions.IgnoreCase);
        private Regex _RegFlowStart = new Regex(@"^\s*(?<flowname>[a-z0-9_]+)\s+(Start)\s*$", RegexOptions.IgnoreCase);
        private Regex _RegFlowEnd = new Regex(@"^\s*(?<flowname>[a-z0-9_]+)\s+((End)|(Stop))\s*$", RegexOptions.IgnoreCase);
        private Regex _RegTestItem = new Regex(@"^\s*\<(?<testname>[a-z0-9_]+)\>s*$", RegexOptions.IgnoreCase);
        // 2022-01-10  Bruce          #295	          Support key name '0-9' chg end
        // 2022-02-21  Steven Chen    #320	          Support No FlowHeader in Binoutcheck add start
        // key:No flow test item or test flow name  value: whether is test flow name
        public Dictionary<string, bool> FlowSteps { get; set; } = new Dictionary<string, bool>();
        // 2022-02-21  Steven Chen    #320	          Support No FlowHeader in Binoutcheck add end
        public string PreCheckDataLog(string datalogFile)
        {
            FileInfo datalog = new FileInfo(datalogFile);
            if (!datalog.Exists) return "DataLog File Not Exists";

            TestFlows.Clear();
            // 2022-02-21  Steven Chen    #320	          Support No FlowHeader in Binoutcheck add start
            FlowSteps.Clear();
            // 2022-02-21  Steven Chen    #320	          Support No FlowHeader in Binoutcheck add end

            string res = ReadDataLog(datalogFile);

            if (res != "")
            {
                return string.Format("DataLog File Format Error. Flow {0} End not found after Flow Start.", res);
            }

            return "";
        }
        public string ReadDataLog(string datalogFile)
        {
            string[] allLines = File.ReadAllLines(datalogFile);
            string flow = "";
            List<string> testItems = null;
            Match match = null;

            string StartEndPairError = "DataLog File Format Error. Flow {0} End not found after Flow Start.";
            string DuplicateFlowError = "DataLog File Format Error. Flow {0} is Duplicated. We found the same flow in file";

            for (int i = 0; i < allLines.Length; i++)
            {
                if (flow == "")
                {
                    match = _RegFlowStart.Match(allLines[i]);
                    if (match.Success)
                    {
                        flow = match.Groups["flowname"].ToString();
                        if (TestFlows.ContainsKey(flow))
                        {
                            return string.Format(DuplicateFlowError, flow);
                        }

                        testItems = new List<string>();
                        TestFlows.Add(flow, testItems);
                        // 2022-02-21  Steven Chen    #320	          Support No FlowHeader in Binoutcheck add start
                        if (!FlowSteps.Keys.Contains(flow))
                            FlowSteps.Add(flow, true);
                        // 2022-02-21  Steven Chen    #320	          Support No FlowHeader in Binoutcheck add end
                    }
                    // 2022-02-21  Steven Chen    #320	          Support No FlowHeader in Binoutcheck add start
                    else
                    {
                        match = _RegTestItem.Match(allLines[i]);
                        if (match.Success)
                        {
                            if (!FlowSteps.Keys.Contains(match.Groups["testname"].ToString()))
                                FlowSteps.Add(match.Groups["testname"].ToString(), false);
                        }
                    }
                    // 2022-02-21  Steven Chen    #320	          Support No FlowHeader in Binoutcheck add end
                }
                else
                {
                    match = _RegFlowEnd.Match(allLines[i]);
                    if (match.Success)
                    {
                        if (flow == match.Groups["flowname"].ToString())
                        {
                            flow = "";
                        }
                        else
                        {
                            return string.Format(StartEndPairError, flow);
                        }
                    }
                    else
                    {
                        match = _RegTestItem.Match(allLines[i]);
                        if (match.Success)
                        {
                            testItems.Add(match.Groups["testname"].ToString());
                        }
                    }
                }
            }

            if (flow != "")
            {
                return string.Format(DuplicateFlowError, flow);
            }

            return "";
        }
    }
}
