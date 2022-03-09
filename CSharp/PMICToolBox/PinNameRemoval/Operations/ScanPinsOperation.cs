using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace PinNameRemoval.Operations
{
    public class ScanPinsOperation : Operation
    {
        public static new Regex RegexScript = new Regex(@"scan_pins", RegexOptions.IgnoreCase);
        public ScanPinsOperation(string line) : base(line) { }

        public override List<string> RemovePins(ref int readLineIndex, StreamReader sr)
        {
            readLineIndex++;
            List<string> ret = new List<string>();
            foreach (string pinName in Ctrl.PinsToDelete)
                currentLine = currentLine.Replace(pinName + ",", "");
            //currentLine = currentLine.Replace(Ctrl.PinsToDelete[0] + ",", "").Replace(Ctrl.PinsToDelete[1] + ",", "");
            ret.Add(currentLine);
            while (!currentLine.Contains("}"))
            {
                currentLine = sr.ReadLine();
                if (string.IsNullOrEmpty(currentLine)) break;
                //currentLine = currentLine.Replace(Ctrl.PinsToDelete[0] + ",", "").Replace(Ctrl.PinsToDelete[1] + ",", "");
                foreach (string pinName in Ctrl.PinsToDelete)
                    currentLine = currentLine.Replace(pinName + ",", "");
                ret.Add(currentLine);
                readLineIndex++;
            }
            return ret;
        }
    }
}
