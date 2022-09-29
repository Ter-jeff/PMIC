using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{PatSetName}")]
    public class PatSet : IgxlRow
    {
        public string PatSetName { get; set; }
        public List<PatSetRow> PatSetRows { get; set; }

        public PatSet()
        {
            PatSetRows = new List<PatSetRow>();
        }

        public void AddRow(PatSetRow row)
        {
            PatSetRows.Add(row);
        }

        public string GetNewPatSetNameWithX(List<string> patterns)
        {
            if (patterns.Count == 0)
                return "";

            var arr = patterns.First().ToCharArray();
            foreach (var pattern in patterns)
                for (var index = 0; index < patterns.First().Length; index++)
                    if (index >= pattern.Length)
                    {
                        arr[index] = 'X';
                    }
                    else
                    {
                        if (!pattern[index].ToString().Equals(patterns.First()[index].ToString(),
                                StringComparison.CurrentCultureIgnoreCase)) arr[index] = 'X';
                    }

            return string.Join("", arr);
        }

        public string GetNewPatSetName(List<string> patterns)
        {
            if (patterns.Count == 0)
                return "";

            patterns = patterns.Where(x => !Regex.IsMatch(x, @"_IN\w{2}_", RegexOptions.IgnoreCase)).ToList();
            var max = patterns.Max(x => x.Split('_').Length);
            var arr = new bool[max];
            var first = patterns.First().Split('_');
            foreach (var pattern in patterns)
            {
                var items = pattern.Split('_');
                for (var index = 0; index < max; index++)
                    if (index < items.Length && index < first.Length)
                    {
                        if (first[index] != items[index])
                            arr[index] = true;
                    }
                    else
                    {
                        arr[index] = true;
                    }
            }

            var final = new string[max];
            for (var index = 0; index < max; index++)
            {
                var one = new List<string>();
                foreach (var pattern in patterns)
                {
                    var items = pattern.Split('_');
                    if (index < items.Length)
                    {
                        if (arr[index])
                            one.Add(items[index]);

                        if (!one.Any())
                            one.Add(items[index]);
                    }

                    final[index] = string.Join("_", one);
                }
            }
            return string.Join("_", final);
        }
    }
}