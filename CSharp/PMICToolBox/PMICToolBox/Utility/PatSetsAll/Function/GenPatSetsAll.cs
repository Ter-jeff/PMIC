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
// 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz"
// 2021-08-27  Ze Chen        #148            Tool box , Pattern Set all. When we gen pattern set all the pattern will appear twice when we select "include all".
//
//------------------------------------------------------------------------------ 
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutomation.Utility.PatSetsAll.Function
{
    public class GenPatSetsAll
    {

        // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg Start
        //public void Print(string inputPath, string outputPath, bool absolutePath, IGXLVersionEnum igxlVersion)
        public void Print(string inputPath, string outputPath, bool absolutePath, IGXLVersionEnum igxlVersion, bool gzOnly, bool patOnly, bool all)
        // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg End
        {
            string outputFile = Path.Combine(outputPath, "PatSets_All.txt");

            // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg Start
            //var files = Directory.GetFiles(inputPath, "*", SearchOption.AllDirectories)
            //    .Where(s => s.EndsWith(".gz", StringComparison.OrdinalIgnoreCase) || s.EndsWith(".pat.gz", StringComparison.OrdinalIgnoreCase));
            var files = Directory.GetFiles(inputPath, "*", SearchOption.AllDirectories)
                .Where(s => IsTargetFile(s, gzOnly, patOnly, all));
            // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg End

            List<string> listPatternName = new List<string>();
            foreach (var file in files.ToList())
            {
                string dayCodePattern = Path.GetFileNameWithoutExtension(file).ToUpper().Replace(".PAT", "");
                string pattern = Regex.Match(dayCodePattern, @"(?<pattern>\w+)_\d_", RegexOptions.IgnoreCase).Groups["pattern"].ToString();
                if (!string.IsNullOrEmpty(pattern) && !listPatternName.Contains(pattern))
                {
                    listPatternName.Add(pattern);
                }
            }

            listPatternName.Sort();

            List<string> targets = new List<string>();
            List<string> orderList = files.ToList().OrderByDescending(p => p).ToList();
            listPatternName.ForEach(p =>
            {
                targets.Add(orderList.Find(s => Regex.IsMatch(s, @p + @"_\d_", RegexOptions.IgnoreCase)));
            });

            switch (igxlVersion)
            {
                case IGXLVersionEnum.v9_xx_ultraflex:
                    WritePatSetsSheet_v9_ultraFlex(outputFile, targets, absolutePath, inputPath);
                    break;
                case IGXLVersionEnum.v10_xx_ultraflex:
                    WritePatSetsSheet_v10_ultraFlex(outputFile, targets, absolutePath, inputPath);
                    break;
            }
        }
        private void WritePatSetsSheet_v9_ultraFlex(string outputFile, List<string> targets, bool absolutePath, string inputPath)
        {
            if (inputPath.EndsWith("\\"))
            {
                inputPath = inputPath.TrimEnd('\\');
            }
            string path = Path.GetFileName(inputPath);
            List<string> lines = new List<string>();
            lines.Add("DTPatternSetSheet,version=2.2:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tPattern Sets");
            lines.Add("");
            lines.Add("\tPattern Set\tTD Group\tTime Domain	Enable\tFile/Group Name\tBurst\tStart Label\tStop Label\tComment");

            foreach (var file in targets)
            {
                string dayCodePattern = Path.GetFileNameWithoutExtension(file).ToUpper().Replace(".PAT", "");
                string pattern = Regex.Match(dayCodePattern, @"(?<pattern>\w+)_\d_", RegexOptions.IgnoreCase).Groups["pattern"].ToString();
                // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg Start
                //var patSet = absolutePath
                //    ? "\t" + pattern + "\t\t\t\t" + Path.ChangeExtension(file, null) + "\tno"
                //    : "\t" + pattern + "\t\t\t\t" +
                //      Path.ChangeExtension(file.Replace(outputPath, ".\\" + path), null) + "\tno";
                var patfile = file.EndsWith(".gz", StringComparison.InvariantCultureIgnoreCase) ? Path.ChangeExtension(file, null) : file;
                var patSet = absolutePath
                        ? "\t" + pattern + "\t\t\t\t" + patfile + "\tno"
                        : "\t" + pattern + "\t\t\t\t" +
                          file.Replace(inputPath, ".\\" + path).TrimEnd(".gz".ToArray()) + "\tno";
                // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg End

                lines.Add(patSet);
            }
            File.WriteAllLines(outputFile, lines);

        }
        private void WritePatSetsSheet_v10_ultraFlex(string outputFile, List<string> targets, bool absolutePath, string inputPath)
        {
            if (inputPath.EndsWith("\\"))
            {
                inputPath = inputPath.TrimEnd('\\');
            }
            string path = Path.GetFileName(inputPath);
            List<string> lines = new List<string>();
            lines.Add("DTPatternSetSheet,version=2.3:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1\tPattern Sets");
            lines.Add("");
            lines.Add("\tPattern Set\tTime Domain\tEnable\tFile/Group Name\tBurst\tStart Label\tStop Label\tComment");


            foreach (var file in targets)
            {
                string dayCodePattern = Path.GetFileNameWithoutExtension(file).ToUpper().Replace(".PAT", "");
                string pattern = Regex.Match(dayCodePattern, @"(?<pattern>\w+)_\d_", RegexOptions.IgnoreCase).Groups["pattern"].ToString();


                // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg Start
                //var patSet = absolutePath
                //    ? "\t" + pattern + "\t\t\t" + Path.ChangeExtension(file, null) + "\tno"
                //    : "\t" + pattern + "\t\t\t" +
                //      Path.ChangeExtension(file.Replace(outputPath, ".\\" + path), null) + "\tno";
                var patfile = file.EndsWith(".gz", StringComparison.InvariantCultureIgnoreCase) ? Path.ChangeExtension(file, null) : file;
                var patSet = absolutePath
                        ? "\t" + pattern + "\t\t\t" + patfile + "\tno"
                        : "\t" + pattern + "\t\t\t" +
                            file.Replace(inputPath, ".\\" + path).TrimEnd(".gz".ToArray()) + "\tno";
                // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg End
                lines.Add(patSet);
            }
            File.WriteAllLines(outputFile, lines);
        }

        // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg Start
        private bool IsTargetFile(string fileName, bool gzOnly, bool patOnly, bool all)
        {
            if (gzOnly)
            {
                return fileName.EndsWith(".gz", StringComparison.OrdinalIgnoreCase) || fileName.EndsWith(".pat.gz", StringComparison.OrdinalIgnoreCase);
            }
            else if (patOnly)
            {
                return fileName.EndsWith(".pat", StringComparison.OrdinalIgnoreCase);
            }
            else
            {
                return fileName.EndsWith(".gz", StringComparison.OrdinalIgnoreCase) ||
                    fileName.EndsWith(".pat.gz", StringComparison.OrdinalIgnoreCase) ||
                    fileName.EndsWith(".pat", StringComparison.OrdinalIgnoreCase);
            }
        }
        // 2021-07-06  Bruce Qian     #86             Tool Box , Gen pattern set all , the rule ".pat.gz" Chg End
    }
}