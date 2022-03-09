using PmicAutomation.Utility.PA.Input;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PmicAutomation.Utility.PA.Function
{
    public class GenPattern
    {
        public void Write(string outputFile, Dictionary<string, PaSheet> PaSheets)
        {
            var pins = PaSheets.SelectMany(x => x.Value.Rows).Where(x => x.PaType.Equals("I/O", StringComparison.CurrentCulture) ||
            x.PaType.Equals("IO", StringComparison.CurrentCulture))
                .Where(x => !string.IsNullOrEmpty(x.GenPattern)).GroupBy(x => x.BumpName).Select(x => x.First()).ToList();

            var OtherPins = PaSheets.SelectMany(x => x.Value.Rows).Where(x => x.PaType.Equals("I/O", StringComparison.CurrentCulture) ||
            x.PaType.Equals("IO", StringComparison.CurrentCulture))
                .Where(x => string.IsNullOrEmpty(x.GenPattern)).GroupBy(x => x.BumpName).Select(x => x.First()).ToList();

            List<string> lines = new List<string>();
            List<string> pinList = new List<string>();
            string patternName = "Dummy";
            string test = "tsetBSCAN";
            lines.Add("import tset " + test + ";");
            foreach (var item in pins.GroupBy(x => x.GenPattern))
            {
                var subPins = item.Select(x => x.BumpName).ToList();
                if (subPins.Count == 1)
                    pinList.Add(subPins[0]);
                else
                    pinList.Add("(" + string.Join(",", subPins) + ")");
            }
            foreach (var item in OtherPins)
                pinList.Add(item.BumpName);

            lines.Add("vm_vector " + patternName + " ( $tset    ," + string.Join(",", pinList) + ")");
            lines.Add("{");
            lines.Add("//");
            int maxPin = pins.Count == 0 ? 0 : pins.Max(x => x.BumpName.Length);
            int maxOther = OtherPins.Count == 0 ? 0 : OtherPins.Max(x => x.BumpName.Length);
            int max = Math.Max(maxPin, maxOther);
            for (int i = 0; i < max; i++)
            {
                string data = "//           ";
                foreach (var item in pins.GroupBy(x => x.GenPattern))
                {
                    foreach (var pin in item)
                    {
                        if (i < pin.BumpName.Length)
                            data += pin.BumpName[i];
                        else
                            data += " ";
                    }
                    data += " ";
                }

                foreach (var item in OtherPins)
                {
                    if (i < item.BumpName.Length)
                        data += item.BumpName[i];
                    else
                        data += " ";
                    data += " ";
                }
                lines.Add(data);
            }
            lines.Add("//");

            lines.Add(patternName + ":");

            string data1 = "";
            foreach (var item in pins.GroupBy(x => x.GenPattern))
            {
                foreach (var pin in item)
                {
                    if (!string.IsNullOrEmpty(pin.PinType) && pin.PinType.ToUpper().Contains("CLOCK"))
                        data1 += "1";
                    else
                        data1 += "X";
                }
                data1 += " ";
            }

            foreach (var pin in OtherPins)
            {
                if (!string.IsNullOrEmpty(pin.PinType) && pin.PinType.ToUpper().Contains("CLOCK"))
                    data1 += "1 ";
                else
                    data1 += "X ";
            }

            for (int i = 1; i < 129; i++)
            {
                lines.Add(" > " + test + " " + data1 + " ; // 0 V:" + i + " C:" + i);
            }
            lines.Add("}");

            File.WriteAllLines(Path.Combine(outputFile, "Dummy.atp"), lines);
        }
    }
}