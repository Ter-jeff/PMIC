using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PatInfoCmdLib
{
    public class PatPatInfoReader
    {
        public string GetModuleNames(List<string> list)
        {
            foreach (var line in list)
            {
                var context = line.Trim('\r').Trim();

                if (context == "") continue;

                if (Regex.IsMatch(context, @"Module names:", RegexOptions.IgnoreCase))
                {
                    var moduleNameList = context.Split(':')[1].Split(',').Select(x => x.Trim()).Where(x => !string.IsNullOrEmpty(x)).ToList();
                    return string.Join(",", moduleNameList);
                }
            }
            return "";
        }
    }
}