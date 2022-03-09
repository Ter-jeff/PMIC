using System.Collections.Generic;
using System.IO;

namespace PmicAutomation.Utility.PA.Function
{
    public class GenTimeSet
    {
        public void Write(string outputFile, List<string> pins)
        {
            List<string> lines = new List<string>();
            List<string> pinList = new List<string>();
            lines.Add("DTTimesetBasicSheet,version = 2.3:platform = Jaguar:toprow = -1:leftcol = -1:rightcol = -1:tabcolor = 16777215\tTime Sets (Basic)");
            lines.Add("");
            lines.Add("\tTiming Mode:\tSingle\t\tMaster Timeset Name:");
            lines.Add("\tTime Domain:\t\t\tStrobe Ref Setup Name:");
            lines.Add("");
            lines.Add("\t\tCycle\tPin / Group\t\t\tData\t\tDrive\t\t\t\tCompare\t\t\t\tEdge Resolution");
            lines.Add("\tTime Set\tPeriod\tName\tClock Period\tSetup\tSrc\tFmt\tOn\tData\tReturn\tOff\tMode\tOpen\tClose\tRef Offset\tMode\tComment");
            foreach (var pin in pins)
                lines.Add("\ttsetAHB\t= _JTAG_Period\t" + pin + "\t\ti/ o\tPAT\tNR\t0\t0\t\t0\tEdge\t= _JTAG_Period * 0.76923\t\t\tAuto");

            File.WriteAllLines(Path.Combine(outputFile, "TIMESET_Dummy.txt"), lines);
        }
    }
}