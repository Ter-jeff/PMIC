using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace PinNameRemoval.Operations
{
    class InstPinsOperation : Operation
    {
        public static new Regex RegexScript = new Regex(@"instruments\s*=\s*{", RegexOptions.IgnoreCase);
        public InstPinsOperation(string line) : base(line) { }

        public override List<string> RemovePins(ref int readLineIndex, StreamReader sr)
        {
            readLineIndex++;
            List<string> ret = new List<string>();
            Regex format = new Regex(@"\(.*,.*\)", RegexOptions.IgnoreCase);
            Regex delPins = new Regex(string.Join("|", Ctrl.PinsToDelete), RegexOptions.IgnoreCase);
            PinNameLine pnl;
            if ((format.Match(currentLine).Success) && (delPins.Match(currentLine).Success))
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
                if ((format.Match(currentLine).Success) && (delPins.Match(currentLine).Success))
                {
                    string tempLine = currentLine;
                    //Remove the specify pins and changed the DigCap or DigSrc pin's count.
                    //(GPIO9, GPIO10, GPIO11, GPIO12, GPIO13, GPIO14, GPIO15, GPIO16, GPIO17, GPIO18, GPIO19, GPIO20, GPIO21, GPIO22, GPIO23, GPIO24, GPIO25, GPIO26):DigCap 18:auto_trig_enable;
                    Regex l_RegexAll = new Regex(@"\((?<PinList>.*?)\)(?<NonPinList>\s*:\s*((DigCap)|(DigSrc))?\s*(?<Count>\d*)\s*:?.*)?", RegexOptions.IgnoreCase);

                    string l_strPinlist = l_RegexAll.Match(tempLine).Groups["PinList"].Value;
                    string l_strCount = l_RegexAll.Match(tempLine).Groups["Count"].Value;
                    string l_strNonPinStr= l_RegexAll.Match(tempLine).Groups["NonPinList"].Value;

                    string[] l_AryPinList = l_strPinlist.Replace(" ","").Split(new string[] {","},System.StringSplitOptions.RemoveEmptyEntries);
                    List<string> l_ChangedPinList = new List<string>(l_AryPinList);

                    foreach (string pinName in Ctrl.PinsToDelete)
                    {
                        if(l_ChangedPinList.Contains(pinName))
                        {
                            l_ChangedPinList.Remove(pinName);
                        }
                    }
                    
                    if (l_ChangedPinList.Count < 1 )
                    {
                        return ret;
                    }
                    else
                    {
                        string pinString = string.Join(",", l_ChangedPinList);
                        string defString = l_strNonPinStr;
                        if (string.IsNullOrEmpty(l_strCount) == false)
                        {

                            if (Regex.IsMatch(defString, @"(?<NonPinList>\s*:\s*((DigCap)|(DigSrc))?\s*(?<Count>\d*)\s*:?.*)", RegexOptions.IgnoreCase))
                            {
                                Regex l_RegexCount = new Regex(@"\d*");
                                string l_ChangedCount = Convert.ToString(l_ChangedPinList.Count);
                                defString = Regex.Replace(defString, l_strCount, l_ChangedCount);
                            }
                        }
                        string retString = "(" + pinString + ")" + (defString);
                        ret.Add(retString);
                        return ret;
                    }
                }

                if (!(delPins.Match(currentLine).Success))
                {
                    ret.Add(currentLine);
                    return ret;
                }
            }
            return ret;
        }

        private string ReplaceCount(string l_strNonPinStr, int l_strCount, int l_ChangedCount)
        {
            string ret = l_strNonPinStr.Replace(char.Parse(l_strCount.ToString()), char.Parse(l_ChangedCount.ToString()));
            return ret;
        }
    }
}
