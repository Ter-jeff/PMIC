using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Library.DataStruct
{
    public class Settings
    {
        public List<ItemInfo> ItemDefine = new List<ItemInfo>();
        public List<HeaderPattern> HeaderPatterns = new List<HeaderPattern>();
        public List<LogRowTypePattern> LogRowTypePatterns = new List<LogRowTypePattern>();
        public List<IgnoredItemPattern> IgnoredItemPatterns = new List<IgnoredItemPattern>();

        public HeaderPattern GetHeaderPatternByHeader(string line)
        {
            return HeaderPatterns.FirstOrDefault(pat => pat.HeaderReg.IsMatch(line));
        }

        public HeaderPattern GetHeaderPatternByName(string logType)
        {
            return HeaderPatterns.FirstOrDefault(pat => pat.Name.Equals(logType, StringComparison.OrdinalIgnoreCase));
        }

        public LogRowTypePattern GetRowTypeTypePatternByName(string name)
        {
            return LogRowTypePatterns.FirstOrDefault(pat => pat.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        public IgnoredItemPattern GetIgnoredItemPatternByName(string name)
        {
            return IgnoredItemPatterns.FirstOrDefault(pat => pat.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }
    }
}
