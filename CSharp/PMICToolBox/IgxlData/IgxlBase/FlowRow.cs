using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public class FlowRow : IgxlRow
    {
        #region Property
        public string SheetName = "";
        public string LineNum = "";
        public string Label = "";
        public string Enable = "";
        public string Job = "";
        public string Part = "";
        public string Env = "";
        public string Opcode = "";
        public string Parameter = "";
        public string TName = "";
        public string TNum = "";
        public string LoLim = "";
        public string HiLim = "";
        public string Scale = "";
        public string Units = "";
        public string Format = "";
        public string BinPass = "";
        public string BinFail = "";
        public string SortPass = "";
        public string SortFail = "";
        public string Result = "";
        public string PassAction = "";
        public string FailAction = "";
        public string State = "";
        public string GroupSpecifier = "";
        public string GroupSense = "";
        public string GroupCondition = "";
        public string GroupName = "";
        public string DeviceSense = "";
        public string DeviceCondition = "";
        public string DeviceName = "";
        public string DebugAsume = "";
        public string DebugSites = "";
        public string CtProfileDataElapsedTime = "";
        public string CtProfileDataBackgroundType = "";
        public string CtProfileDataSerialize = "";
        public string CtProfileDataResourceLock = "";
        public string CtProfileDataFlowStepLocked = "";
        public string Comment = "";
        public string Comment1 = "";
        public int Rownumber;
        public const string OpCodeTest = "Test";

        #endregion
        public List<string> GetEnables()
        {
            List<string> enables = Regex.Split(Enable, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return enables;
        }

        public List<string> GetJobs()
        {
            List<string> jobs = Regex.Split(Job, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return jobs;
        }

        public List<string> GetEnvs()
        {
            List<string> envs = Regex.Split(Env, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return envs;
        }

        public bool IsMatchEnable(List<string> enableWords)
        {
            if (string.IsNullOrEmpty(Enable)) return true;
            List<string> enables = Regex.Split(Enable, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return enables.Exists(x => enableWords.Exists(y => y.Equals(x, StringComparison.CurrentCultureIgnoreCase)));
        }

        public bool IsMatchJob(string job)
        {
            if (string.IsNullOrEmpty(Job)) return true;
            List<string> jobs = Regex.Split(Job, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return jobs.Exists(x => x.Equals(job, StringComparison.CurrentCultureIgnoreCase));
        }

        public bool IsMatchEnv(string env)
        {

            if (string.IsNullOrEmpty(Env)) return true;
            List<string> envs = Regex.Split(Env, @"[^\w]").Where(x => Regex.IsMatch(x, @"\w")).Select(x => x.Trim()).ToList();
            return envs.Exists(x => x.Equals(env, StringComparison.CurrentCultureIgnoreCase));
        }

        public FlowRow DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as FlowRow;
            }
        }
    }
}