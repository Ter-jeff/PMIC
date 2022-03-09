using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;

namespace PmicAutogen.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess
{
    [Serializable]
    public class PatternClass
    {
        private const string Pattern = "([,&;])";
        public List<List<string>> InitList = new List<List<string>>();
        public List<string> InstanceInitName = new List<string>();
        public List<string> InstancePatternName = new List<string>();
        public List<string> InstancePayloadName = new List<string>();
        public List<List<string>> PatternList = new List<List<string>>();
        public List<List<string>> PayloadList = new List<List<string>>();
        public string RealPatternName;
        public string TestPlanPatternName;

        public PatternClass(string patternName)
        {
            TestPlanPatternName = patternName.ToLower();
            RealPatternName = patternName.ToLower();

            foreach (var pat in Regex.Split(TestPlanPatternName, "[,&;]").ToList())
                InstancePatternName.Add(pat);

            if (InstancePatternName.Count > 1)
            {
            }

            foreach (var name in TestPlanPatternName.Split(';'))
            foreach (var seq in name.Split('&'))
                PatternList.Add(seq.Split(',').ToList());

            if (TestPlanPatternName.Contains(';') && TestPlanPatternName.Split(';').Length == 2)
            {
                var arr = TestPlanPatternName.Split(';').ToList();
                foreach (var seq in arr[0].Split('&'))
                    InitList.Add(seq.Split(',').ToList());
                foreach (var seq in arr[1].Split('&'))
                    PayloadList.Add(seq.Split(',').ToList());
            }
            else
            {
                foreach (var name in TestPlanPatternName.Split(';'))
                foreach (var seq in name.Split('&'))
                    PayloadList.Add(seq.Split(',').ToList());
            }
        }

        public string GetPatternName()
        {
            if (IsMultiple())
                return "Multiple_" + GetLastPayload();
            return GetLastPayload();
        }

        public bool IsMultiple()
        {
            return RealPatternName.Contains(",") || RealPatternName.Contains("&") || RealPatternName.Contains(";") ||
                   TestPlanPatternName.Contains(",") || TestPlanPatternName.Contains("&") ||
                   TestPlanPatternName.Contains(";");
        }

        public List<string> GetAliasPatternList()
        {
            return Regex.Split(TestPlanPatternName, Pattern).ToList();
        }

        public List<string> GetRealPatternList()
        {
            return Regex.Split(RealPatternName, Pattern).ToList();
        }

        public string GetLastPayload()
        {
            if (PatternList.Count == 0)
                return "";
            if (PatternList.Last().Count == 0)
                return "";
            return PatternList.Last().Last();
        }

        public string GetInstancePatternName(bool isRealPattern = false)
        {
            if (!Regex.IsMatch(TestPlanPatternName, HardIpConstData.NoPattern, RegexOptions.IgnoreCase) &&
                !Regex.IsMatch(TestPlanPatternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase))
            {
                if (isRealPattern)
                    return string.Join(",", PatternList.SelectMany(x => x));
                return string.Join(";", InstancePatternName);
            }

            return "";
        }

        public string GetInstanceInitName(bool isRealPattern = false)
        {
            if (!Regex.IsMatch(TestPlanPatternName, HardIpConstData.NoPattern, RegexOptions.IgnoreCase) &&
                !Regex.IsMatch(TestPlanPatternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase))
            {
                if (isRealPattern)
                    return string.Join(",", InitList.SelectMany(x => x));
                return string.Join(";", InstanceInitName);
            }

            return "";
        }

        public string GetInstancePayloadName(bool isRealPattern = false)
        {
            if (!Regex.IsMatch(TestPlanPatternName, HardIpConstData.NoPattern, RegexOptions.IgnoreCase) &&
                !Regex.IsMatch(TestPlanPatternName, HardIpConstData.RegInsInPattern, RegexOptions.IgnoreCase))
            {
                if (isRealPattern)
                    return string.Join(",", PayloadList.SelectMany(x => x));
                return string.Join(";", InstancePayloadName);
            }

            return "";
        }
    }
}