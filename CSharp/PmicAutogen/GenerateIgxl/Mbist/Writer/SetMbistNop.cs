using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class SetMbistNop
    {
        public void SetNop(List<ProdCharRowMbist> instanceList)
        {
            const string matchPattern = "^TIMESET";
            foreach (var instance in instanceList)
            {
                if (instance.Nop)
                    continue;

                //Check if all the patterns are exist in patSet
                var containPattern = true;
                string status;
                foreach (var pattern in instance.PayloadList)
                    if (!InputFiles.PatternListMap.GetStatusInPatternList(pattern.PatternName, out status))
                        containPattern = false;
                foreach (var init in instance.InitList.Values)
                    if (init.PatternName != "" &&
                        !InputFiles.PatternListMap.GetStatusInPatternList(init.PatternName, out status))
                        containPattern = false;
                if (!containPattern)
                {
                    instance.Nop = true;
                    instance.NopType = NopType.WrongTimeSet;
                    continue;
                }


                //Check timeSet
                foreach (var pattern in instance.PayloadList)
                {
                    var patternListCsvRow = InputFiles.PatternListMap.PatternListCsvRows.Find(x =>
                        x.PatternName.Equals(pattern.PatternName, StringComparison.CurrentCultureIgnoreCase));
                    var timeSet = patternListCsvRow.TimeSetVersion;
                    if (!Regex.IsMatch(timeSet, matchPattern, RegexOptions.IgnoreCase))
                    {
                        instance.Nop = true;
                        instance.NopType = NopType.WrongTimeSet;
                    }
                }

                //Check if the pattern in patSet
                var noUse = false;
                foreach (var pattern in instance.PayloadList)
                {
                    var patternListCsvRow = InputFiles.PatternListMap.PatternListCsvRows.Find(x =>
                        x.PatternName.Equals(pattern.PatternName, StringComparison.CurrentCultureIgnoreCase));
                    var use = patternListCsvRow.Use;
                    if (Regex.IsMatch(use, "dont_use", RegexOptions.IgnoreCase)) noUse = true;
                }

                foreach (var init in instance.InitList.Values)
                {
                    if (init.PatternName == "")
                        continue;

                    var patternListCsvRow = InputFiles.PatternListMap.PatternListCsvRows.Find(x =>
                        x.PatternName.Equals(init.PatternName, StringComparison.CurrentCultureIgnoreCase));
                    var use = patternListCsvRow.Use;
                    if (Regex.IsMatch(use, "dont_use", RegexOptions.IgnoreCase)) noUse = true;
                }

                if (noUse)
                {
                    instance.Nop = true;
                    instance.NopType = NopType.NoUse;
                }
            }
        }
    }
}