using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    [Serializable]
    public class TestPlanRow
    {
        public TestPlanRow()
        {
            Limits = new List<MeasLimit>();
            TestName = "";
            RfInstrumentSetup = "";
            InterposeFunc = "";
        }

        public int RowNum { get; set; }
        public string Description { get; set; }
        public string ForceCondition { get; set; }
        public string ForceConditionChar { get; set; }
        public string RegisterAssignment { get; set; }
        public string MiscInfo { get; set; }
        public string Meas { get; set; }
        public List<MeasLimit> Limits { get; set; }
        public int MergeRowNumForMeas { get; set; }
        public string TestName { get; set; }
        public string InterposeFunc { get; set; }
        public string RfInterpose { get; set; }
        public string RfInstrumentSetup { get; set; }
    }
}