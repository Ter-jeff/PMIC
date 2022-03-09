using System;
using System.Collections.Generic;
using System.Linq;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase
{
    public class TestPlanData
    {
        private static readonly List<string> DefaultJobs = new List<string>
        {
            "CP1",
            "CP2",
            "FT1",
            "FT2",
            "FT3"
        };

        public bool CzBinOut;
        public bool CzHvEnable;
        public bool CzLvEnable;
        public bool CzNvEnable;
        public bool ExtendInstance;
        public bool ExtendLimits;
        public bool HvEnable;
        public bool LvEnable;
        public bool NvEnable;

        public List<string> PerformanceModeList;
        public bool SpecialSetting;
        public bool SplitCzFlow;

        public TestPlanData()
        {
            PinList = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            PinGroupList = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            PlanHeaderIdx = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
            var jobList = new List<string>();
            if (StaticSetting.JobMap != null)
                foreach (var jobMap in StaticSetting.JobMap)
                    jobList.AddRange(jobMap.Value);
            JobMappingDic = InitialJobs(jobList);

            PerformanceModeList = new List<string>();
            SpecialSetting = false;
            ExtendLimits = false;
            HvEnable = true;
            LvEnable = true;
            NvEnable = true;
            CzHvEnable = true;
            CzLvEnable = true;
            CzNvEnable = true;
            ExtendInstance = false;
            SplitCzFlow = false;
            CzBinOut = false;
        }

        public Dictionary<string, string> PinList { get; set; }
        public Dictionary<string, List<string>> PinGroupList { get; set; }

        public Dictionary<string, Dictionary<string, int>> PlanHeaderIdx { get; set; }
        public List<string> AllJobs => JobMappingDic.Keys.ToList();
        public Dictionary<string, string> JobMappingDic { get; }

        public List<string> MeasTypes { get; } = new List<string>
        {
            MeasType.MeasV,
            MeasType.MeasE,
            MeasType.MeasI,
            MeasType.MeasC,
            MeasType.MeasF,
            MeasType.MeasIdiff,
            MeasType.MeasVdiff,
            MeasType.MeasVdiff2,
            MeasType.MeasFdiff,
            MeasType.MeasVocm,
            MeasType.MeasR1,
            MeasType.MeasR2,
            MeasType.MeasI2,
            MeasType.MeasCalc,
            MeasType.MeasLimit,
            MeasType.MeasCalcLimit
        };

        public static Dictionary<string, string> InitialJobs(List<string> jobList)
        {
            var jobMap = new Dictionary<string, string>();
            if (jobList.Count == 0)
                foreach (var job in DefaultJobs)
                    jobMap.Add(job, job);
            else
                for (var i = 0; i < jobList.Count; i++)
                {
                    if (i >= DefaultJobs.Count) break;
                    jobMap.Add(DefaultJobs[i], jobList[i]);
                }

            return jobMap;
        }
    }
}