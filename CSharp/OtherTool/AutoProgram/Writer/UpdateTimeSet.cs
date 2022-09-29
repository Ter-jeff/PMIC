using AutoProgram.Reader;
using CommonReaderLib.DebugPlan;
using IgxlData.IgxlSheets;
using IgxlData.Others.MultiTimeSet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace AutoProgram.Writer
{
    public class UpdateTimeSet
    {
        public List<ComTimeSetBasicSheet> Work(DebugPlanMain debugTestPlan, List<TimeSetBasicSheet> timeSetBasicSheets,
            PortMapSheet portMapSheet,
            string patternFolder)
        {
            var timeSets = debugTestPlan.AiTestPlanSheets.SelectMany(x => x.Rows)
                .Where(x => !string.IsNullOrEmpty(x.Timeset)).Select(x => x.Timeset + ".txt").ToList();

            var emptyTimeSets = new List<string>();
            var timeSetEmptyRows = debugTestPlan.AiTestPlanSheets.SelectMany(x => x.Rows)
                .Where(x => string.IsNullOrEmpty(x.Timeset)).ToList();
            foreach (var timeSetEmptyRow in timeSetEmptyRows)
            {
                foreach (var pattern in timeSetEmptyRow.Payloads)
                {
                    if (debugTestPlan.PatternListSheet.Rows.Exists(x =>
                            x.Pattern.Equals(pattern.OriName, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        var timeset = debugTestPlan.PatternListSheet.Rows.Find(x =>
                            x.Pattern.Equals(pattern.OriName, StringComparison.CurrentCultureIgnoreCase)).TimeSet;
                        emptyTimeSets.Add(timeset);
                    }
                }
            }

            timeSets.AddRange(emptyTimeSets);
            timeSets = timeSets.Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            var timeSetPaths = timeSets.Except(timeSetBasicSheets.Select(x => x.SheetName))
                .Select(x => Path.Combine(patternFolder, @"Timeset\" + x)).ToList();

            var comTimeSetBasicSheets = new TimesetReader().ReadTimeSetTxt1P4(timeSetPaths);

            #region add port set
            var allTsets = timeSetBasicSheets.SelectMany(x => x.Tsets).ToList();
            List<ComTimeSetBasic> tsets = new List<ComTimeSetBasic>();
            foreach (var portSets in portMapSheet.PortSets)
            {
                foreach (var tset in allTsets)
                {
                    if (portSets.PortName.Equals(tset.Name, StringComparison.CurrentCultureIgnoreCase))
                    {
                        var comTimeSetBasic = new ComTimeSetBasic();
                        comTimeSetBasic.Name = tset.Name;
                        comTimeSetBasic.CyclePeriod = tset.CyclePeriod;
                        comTimeSetBasic.AddTimingRows(tset.TimingRows);
                        tsets.Add(comTimeSetBasic);
                        break;
                    }
                }
            }
            if (tsets.Count > 0)
            {
                foreach (var comTimeSetBasicSheet in comTimeSetBasicSheets)
                {
                    foreach (var tset in tsets)
                    {
                        if (!comTimeSetBasicSheet.Tsets.Exists(x => x.Name.Equals(tset.Name, StringComparison.CurrentCultureIgnoreCase)))
                            comTimeSetBasicSheet.Tsets.Add(tset);
                    }
                }
            }
            #endregion

            return comTimeSetBasicSheets;
        }
    }
}