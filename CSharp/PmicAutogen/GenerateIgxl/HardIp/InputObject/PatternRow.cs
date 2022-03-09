using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using PmicAutogen.GenerateIgxl.HardIp.InputReader.TestPlanPreprocess;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    public class PatternRow
    {
        public PatternRow(string patternName = "")
        {
            PatChildRows = new List<PatChildRow>();
            RowNum = 0;
            PatternColumnNum = 0;
            TtrStr = "";
            NoBinOutStr = "";
            Description = "";
            ForceCondition = new ForceClass();
            PostPatForceCondition = "";
            RegisterAssignment = "";
            MiscInfo = "";
            RfInterPose = "";
            DupIndex = 0;
            DdrExtraPat = null;
            Pattern = new PatternClass(patternName);
        }

        public string SheetName { get; set; }
        public int RowNum { get; set; }
        public int PatternColumnNum { get; set; }
        public string TtrStr { get; set; }
        public string NoBinOutStr { get; set; }
        public string Description { get; set; }
        public PatternClass Pattern { get; set; }
        public ForceClass ForceCondition { get; set; }
        public string ForceConditionChar { get; set; }
        public string AnalogSetup { get; set; }
        public string SpecifyTestName { get; set; }
        public string RfInterPose { get; set; }
        public string PostPatForceCondition { get; set; }
        public string RegisterAssignment { get; set; }
        public string MiscInfo { get; set; }
        public List<PatChildRow> PatChildRows { get; set; }
        public int DupIndex { get; set; }
        public PatternRow DdrExtraPat { get; set; }

        public PatternRow DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as PatternRow;
            }
        }

        public List<TestPlanRow> GetTestPlanRows()
        {
            var testPlanRows = new List<TestPlanRow>();

            foreach (var measRow in PatChildRows)
            {
                var tpRows = ((PatSubChildRow) measRow).TpRows;
                testPlanRows.AddRange(tpRows);
            }

            return testPlanRows;
        }
    }
}