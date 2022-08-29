using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlBase
{
    public class FlowRow : IgxlItem
    {
        public const string OpCodeTest = "Test";
        public const string OpCodeBinTable = "BinTable";
        public const string OpCodeNop = "Nop";
        public const string OpCodeCharacterize = "Characterize";
        public const string OpCodeUseLimit = "Use-Limit";
        public const string OpCodeTestDeferLimit = "Test-defer-limits";

        public FlowRow()
        {
            Label = "";
        }

        public List<string> GetEnables()
        {
            var enables = Regex.Split(Enable, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim())
                .ToList();
            return enables;
        }

        public List<string> GetJobs()
        {
            var jobs = Regex.Split(Job, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return jobs;
        }

        public List<string> GetEnvs()
        {
            var envs = Regex.Split(Env, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return envs;
        }

        public bool IsMatchEnable(List<string> enableWords)
        {
            if (string.IsNullOrEmpty(Enable)) return true;
            var enables = Regex.Split(Enable, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim())
                .ToList();
            return enables.Exists(x => enableWords.Exists(y => y.Equals(x, StringComparison.CurrentCultureIgnoreCase)));
        }

        public bool IsMatchJob(string job)
        {
            if (string.IsNullOrEmpty(Job)) return true;
            var jobs = Regex.Split(Job, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return jobs.Exists(x => x.Equals(job, StringComparison.CurrentCultureIgnoreCase));
        }

        public bool IsMatchEnv(string env)
        {
            if (string.IsNullOrEmpty(Env)) return true;
            var envs = Regex.Split(Env, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return envs.Exists(x => x.Equals(env, StringComparison.CurrentCultureIgnoreCase));
        }

        #region Property

        public string SheetName { get; set; }
        public string LineNum { get; set; }
        public string Label { get; set; }
        public string Enable { get; set; }
        public string Job { get; set; }
        public string Part { get; set; }
        public string Env { get; set; }
        public string OpCode { get; set; }
        public string Parameter { get; set; }
        public string Name { get; set; }
        public string Num { get; set; }
        public string LoLim { get; set; }
        public string HiLim { get; set; }
        public string Scale { get; set; }
        public string Units { get; set; }
        public string Format { get; set; }
        public string BinPass { get; set; }
        public string BinFail { get; set; }
        public string SortPass { get; set; }
        public string SortFail { get; set; }
        public string Result { get; set; }
        public string PassAction { get; set; }
        public string FailAction { get; set; }
        public string State { get; set; }
        public string GroupSpecifier { get; set; }
        public string GroupSense { get; set; }
        public string GroupCondition { get; set; }
        public string GroupName { get; set; }
        public string DeviceSense { get; set; }
        public string DeviceCondition { get; set; }
        public string DeviceName { get; set; }
        public string DebugAsume { get; set; }
        public string DebugSites { get; set; }
        public string CtProfileDataElapsedTime { get; set; }
        public string CtProfileDataBackgroundType { get; set; }
        public string CtProfileDataSerialize { get; set; }
        public string CtProfileDataResourceLock { get; set; }
        public string CtProfileDataFlowStepLocked { get; set; }
        public string Comment { get; set; }
        public string Comment1 { get; set; }
        public int RowNumber { get; set; }

        #endregion
    }
}