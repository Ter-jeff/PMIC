using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace PinNameRemoval.Operations
{
    public class ScanPinDataOperation : Operation
    {
        public static new Regex RegexScript = new Regex(@"^\s*[(]", RegexOptions.IgnoreCase);
        public ScanPinDataOperation(string line) : base(line) { }

        public override List<string> RemovePins(ref int readLineIndex, StreamReader sr)
        {
            readLineIndex++;
            List<string> ret = new List<string>();
            // To Add: delete whole content in "()" if found pin name to be deleted
            Regex regex = new Regex(string.Join("|", Ctrl.PinsToDelete), RegexOptions.IgnoreCase);
            if (!regex.Match(currentLine).Success)
            {
                ret.Add(currentLine);
                return ret;
            }
            //ret.Add(currentLine);
            while (!currentLine.Contains(")"))
            {
                readLineIndex++;
                currentLine = sr.ReadLine();
                if (string.IsNullOrEmpty(currentLine)) break;
                //ret.Add(currentLine);
            }
            return ret;
        }
    }
}
