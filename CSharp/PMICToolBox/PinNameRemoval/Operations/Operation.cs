using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace PinNameRemoval.Operations
{
    public class Operation
    {
        protected static Regex RegexScript = new Regex(@"^\s*//", RegexOptions.IgnoreCase);
        protected string currentLine;

        public Operation(string line)
        {
            currentLine = line;
        }

        public static Operation CreateOperation(string line)
        {
            Operation op = new Operation(line);

            if (RegexScript.Match(line).Success)
                op = new Operation(line);
            else if (InstPinsOperation.RegexScript.Match(line).Success)
                op = new InstPinsOperation(line);
            else if (ScanPinsOperation.RegexScript.Match(line).Success)
                op = new ScanPinsOperation(line);
            else if (PinNamesOperation.RegexScript.Match(line).Success)
                op = new PinNamesOperation(line);
            else if (ScanPinDataOperation.RegexScript.Match(line).Success)
                op = new ScanPinDataOperation(line);
            else if (PinDataOperation.RegexScript.Match(line).Success)
                op = new PinDataOperation(line);

            return op;
        }

        public virtual List<string> RemovePins(ref int readLineIndex, StreamReader sr)
        {
            readLineIndex++;
            List<string> ret = new List<string>();
            ret.Add(currentLine);
            return ret;
        }
    }
}
