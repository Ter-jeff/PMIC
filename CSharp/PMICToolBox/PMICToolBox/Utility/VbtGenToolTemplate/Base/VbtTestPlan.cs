using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PmicAutomation.Utility.VbtGenToolTemplate.Base
{
    public class VbtTestPlanRow
    {
        public int RowNum;
        public string TopList { set; get; }
        public string Command { set; get; }
        public string FunctionName { set; get; }
        public string RegisterMacroName { set; get; }
        public string BitfieldName { set; get; }
        public string Values { set; get; }
        public string Pin { set; get; }
        public string DatalogVariable { set; get; }
        public string Unit { set; get; }
        public string LowLimit { set; get; }
        public string HighLimit { set; get; }
        public string CallbackFunction { set; get; }
        public string VRangeL { set; get; }
        public string RangeL { set; get; }
        public string VoltageL { set; get; }
        public string CurrentL { set; get; }
        public string VRangeM { set; get; }
        public string RangeM { set; get; }
        public string VoltageM { set; get; }
        public string CurrentM { set; get; }
        public string Frequency { set; get; }
        public string SampleSize { set; get; }
        public string Comment { set; get; }

        public string WriteLine()
        {
            string line = "";
            line += string.IsNullOrEmpty(TopList) ? "\t" : TopList + "\t";
            line += string.IsNullOrEmpty(Command) ? "\t" : Command + "\t";
            line += string.IsNullOrEmpty(FunctionName) ? "\t" : FunctionName + "\t";
            line += string.IsNullOrEmpty(RegisterMacroName) ? "\t" : RegisterMacroName + "\t";
            line += string.IsNullOrEmpty(BitfieldName) ? "\t" : BitfieldName + "\t";
            line += string.IsNullOrEmpty(Values) ? "\t" : Values + "\t";
            line += string.IsNullOrEmpty(Pin) ? "\t" : Pin + "\t";
            line += string.IsNullOrEmpty(DatalogVariable) ? "\t" : DatalogVariable + "\t";
            line += string.IsNullOrEmpty(Unit) ? "\t" : Unit + "\t";
            line += string.IsNullOrEmpty(LowLimit) ? "\t" : LowLimit + "\t";
            line += string.IsNullOrEmpty(HighLimit) ? "\t" : HighLimit + "\t";
            line += string.IsNullOrEmpty(CallbackFunction) ? "\t" : CallbackFunction + "\t";
            line += string.IsNullOrEmpty(VRangeL) ? "\t" : VRangeL + "\t";
            line += string.IsNullOrEmpty(RangeL) ? "\t" : RangeL + "\t";
            line += string.IsNullOrEmpty(VoltageL) ? "\t" : VoltageL + "\t";
            line += string.IsNullOrEmpty(CurrentL) ? "\t" : CurrentL + "\t";
            line += string.IsNullOrEmpty(VRangeM) ? "\t" : VRangeM + "\t";
            line += string.IsNullOrEmpty(RangeM) ? "\t" : RangeM + "\t";
            line += string.IsNullOrEmpty(VoltageM) ? "\t" : VoltageM + "\t";
            line += string.IsNullOrEmpty(CurrentM) ? "\t" : CurrentM + "\t";
            line += string.IsNullOrEmpty(Frequency) ? "\t" : Frequency + "\t";
            line += string.IsNullOrEmpty(SampleSize) ? "\t" : SampleSize + "\t";
            line += string.IsNullOrEmpty(Comment) ? "\t" : Comment + "\t";
            return line;
        }
    }

    public class VbtTestPlanSheet
    {
        public readonly List<VbtTestPlanRow> Rows = new List<VbtTestPlanRow>();
        public Dictionary<string, int> HeaderIndexDic = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        public string Name;
        public string Title;

        public void GenTxt(string file)
        {
            if (Rows.Count == 0)
            {
                return;
            }

            File.WriteAllLines(file, Rows.Select(x => x.WriteLine()).ToList());
        }
    }
}