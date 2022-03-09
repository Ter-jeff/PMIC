using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace PinNameRemoval.Operations
{
    class PinNamesOperation : Operation
    {
        public static new Regex RegexScript = new Regex(@"vm_vector", RegexOptions.IgnoreCase);
        public PinNamesOperation(string line) : base(line) { }

        public override List<string> RemovePins(ref int readLineIndex, StreamReader sr)
        {
            readLineIndex++;
            List<string> ret = new List<string>();
            Regex tset = new Regex("$tset", RegexOptions.IgnoreCase);
            PinNameLine pnl;
            if (tset.Match(currentLine).Success)
            {
                //Ctrl.GetPinsIndex(currentLine);
                Ctrl.PinIndexList = Ctrl.GetPinsIndexList(currentLine);
                //foreach (string pinName in Ctrl.PinsToDelete)
                //    currentLine = currentLine.Replace(pinName + ",", "");
                pnl = new PinNameLine(currentLine);
                pnl.Delete(Ctrl.PinIndexList);
                //ret.Add(currentLine);
                ret.Add(pnl.ToString());
            }
            else
            {
                readLineIndex++;
                ret.Add(currentLine);
                currentLine = sr.ReadLine();
                if (string.IsNullOrEmpty(currentLine)) return ret;
                pnl = new PinNameLine(currentLine);
                //Ctrl.GetPinsIndex(currentLine);
                //foreach (string pinName in Ctrl.PinsToDelete)
                //    currentLine = currentLine.Replace(pinName + ",", "");
                //ret.Add(currentLine);
                Ctrl.PinIndexList = Ctrl.GetPinsIndexList(currentLine);
                pnl.Delete(Ctrl.PinIndexList);
                ret.Add(pnl.ToString());
            }
            return ret;
        }
    }

    public class PinNameLine
    {
        string prefix;
        string postfix;
        List<List<string>> pinNames;

        public PinNameLine(string input)
        {
            pinNames = Ctrl.GetPinList(input);
            int startIndex = input.IndexOf(",");
            prefix = input.Substring(0, startIndex + 1);
            int endIndex = input.LastIndexOf(")");
            postfix = input.Substring(endIndex);
        }

        public void Delete(List<List<int>> pinIndexList)
        {
            for (int i = pinIndexList.Count - 1; i >= 0; i--)
            {
                for (int j = pinIndexList[i].Count - 1; j >= 0; j--)
                {
                    pinNames[i].RemoveAt(pinIndexList[i][j]);
                }
            }
        }

        public override string ToString()
        {
            string output = prefix;
            List<string> datas = new List<string>();
            foreach (List<string> obj in pinNames)
            {
                if (obj.Count > 1)
                    datas.Add("(" + string.Join(", ", obj) + ")");
                else if (obj.Count == 1)
                    datas.Add(obj[0]);
            }
            return prefix + string.Join(", ", datas) + postfix;
        }
    }
}
