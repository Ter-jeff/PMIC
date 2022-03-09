using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PinNameRemoval.Operations
{
    class PinDataOperation : Operation
    {
        public static new Regex RegexScript = new Regex(@">", RegexOptions.IgnoreCase);
        public PinDataOperation(string line) : base(line) { }

        public override List<string> RemovePins(ref int readLineIndex, StreamReader sr)
        {
            readLineIndex++;
            List<string> ret = new List<string>();
            //List<string> pinData = Regex.Split(currentLine, @"\s+").ToList();
            PinData pd = new PinData(currentLine);
            //int startIndex = pinData.FindIndex(s => s.Contains(">"));
            //Ctrl.PinsIndex.ForEach(index => pinData.RemoveAt(index + startIndex + 2));
            //ret.Add(string.Join(" ", pinData));
            pd.Delete(Ctrl.PinIndexList);
            ret.Add(pd.ToString());
            return ret;
        }
    }

    public class PinData
    {
        string prefix;
        string postfix;
        List<List<string>> value;

        public PinData(string input)
        {
            int startIndex = input.IndexOf('>');
            startIndex = input.IndexOf(' ', startIndex + 2);
            prefix = input.Substring(0, startIndex + 1 + 1);
            int endIndex = input.IndexOf(';');
            postfix = input.Substring(endIndex - 1);

            value = new List<List<string>>();
            string pinData = input.Replace(prefix, "").Replace(postfix, "").Trim();
            foreach (string s in Regex.Split(pinData, @"\s+").ToList())
            {
                if (s.Length == 1)
                    value.Add(new List<string>() { s });
                else
                    value.Add(s.Select(c => c.ToString()).ToList());
            }
        }

        public void Delete(List<List<int>> pinIndexList)
        {
            for (int i = pinIndexList.Count - 1; i >= 0; i--)
            {
                for (int j = pinIndexList[i].Count - 1; j >= 0; j--)
                {
                    value[i].RemoveAt(pinIndexList[i][j]);
                }
            }
        }

        public override string ToString()
        {
            string output = prefix;
            List<string> datas = new List<string>();
            foreach (List<string> obj in value)
            {
                if (obj.Count > 1)
                    datas.Add(string.Join("", obj));
                else if (obj.Count == 1)
                    datas.Add(obj[0]);
            }
            return prefix + string.Join(" ", datas) + postfix;
        }
    }
}
