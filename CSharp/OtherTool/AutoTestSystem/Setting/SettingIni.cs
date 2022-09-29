using System;
using System.Collections.Generic;
using System.IO;

namespace AutoTestSystem.Setting
{
    public class SettingIni
    {
        private const string Cautogen = "CAutogen";

        private const string ConstJobName = "JobName"; //for CAutogen
        private const string ConstTestProgram = "TestProgram"; //for CAutogen  
        private const string ConstTestProgramDlex = "TestProgram.dlex(OPT)"; //for CAutogen 
        private const string ConstPatternFolder = "PatternFolder"; //for CAutogen
        private const string ConstPatternSync = "PatternSync"; //for CAutogen
        private const string ConstOutputFolder = "OutputFolder"; //for CAutogen

        private const string ConstMailTo = "MailTo";
        private const string ConstLotId = "LotId"; //for igxl testing 
        private const string ConstWaferId = "WaferID"; //for igxl testing
        private const string ConstSetXy = "SetXY"; //for igxl testing  
        private const string ConstEnableWords = "EnableWords"; //for igxl testing
        private const string ConstSites = "Sites"; //for igxl testing
        private const string ConstDoAll = "DoAll"; //for igxl testing
        private const string ConstOverrideFailStop = "OverrideFailStop"; //for igxl testing   

        public List<IniRow> Rows = new List<IniRow>();

        public string JobName
        {
            get { return GetCondition(ConstJobName); }
        }

        public string TestProgram
        {
            get { return GetCondition(ConstTestProgram); }
        }

        public string TestProgramDlex
        {
            get { return GetCondition(ConstTestProgramDlex); }
        }

        public string PatternFolder
        {
            get { return GetCondition(ConstPatternFolder); }
        }

        public string PatternSync
        {
            get { return GetCondition(ConstPatternSync); }
        }

        public string MailTo
        {
            get { return GetCondition(ConstMailTo); }
        }
        public string WaferId
        {
            get { return GetCondition(ConstWaferId); }
        }

        public string LotId
        {
            get { return GetCondition(ConstLotId); }
        }
        public string SetXy
        {
            get { return GetCondition(ConstSetXy); }
        }

        public string EnableWords
        {
            get { return GetCondition(ConstEnableWords); }
        }

        public string Sites
        {
            get { return GetCondition(ConstSites); }
        }

        public string DoAll
        {
            get { return GetCondition(ConstDoAll); }
        }

        public string OverrideFailStop
        {
            get { return GetCondition(ConstOverrideFailStop); }
        }

        public string TimeFolder
        {
            get
            {
                if (!string.IsNullOrEmpty(PatternFolder))
                    return Path.Combine(PatternFolder, "TimeSet");
                return "";
            }
        }

        private string GetCondition(string name)
        {
            foreach (var item in Rows)
                if (item.Name.Equals(name, StringComparison.CurrentCultureIgnoreCase))
                    return item.Value;
            return "";
        }

        public List<IniRow> GetGenProgramCondition()
        {
            var iniRows = new List<IniRow>();
            foreach (var item in Rows)
                if (item.Name.Equals(ConstJobName, StringComparison.CurrentCultureIgnoreCase) ||
                    item.Name.Equals(ConstTestProgram, StringComparison.CurrentCultureIgnoreCase) ||
                    item.Name.Equals(ConstTestProgramDlex, StringComparison.CurrentCultureIgnoreCase) ||
                    item.Name.Equals(ConstPatternFolder, StringComparison.CurrentCultureIgnoreCase) ||
                    item.Name.Equals(ConstOutputFolder, StringComparison.CurrentCultureIgnoreCase))
                    iniRows.Add(item);
            return iniRows;
        }

        public void Read(string iniFile)
        {
            var comIni = new ComIni();
            var allDic = comIni.Read(iniFile);
            if (allDic.ContainsKey(Cautogen))
                foreach (var item in allDic[Cautogen])
                    Rows.Add(item);
        }
    }
}