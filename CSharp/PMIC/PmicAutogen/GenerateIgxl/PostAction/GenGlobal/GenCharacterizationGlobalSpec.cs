using IgxlData.IgxlBase;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.PostAction.GenGlobal
{
    public class GenCharacterizationGlobalSpec
    {
        public bool ExtendGlobalSpec()
        {
            var lGlobalSpecsList = new List<GlobalSpec>();

            lGlobalSpecsList.Add(new GlobalSpec(""));

            if (TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value == null) return false;

            var lGlobalSpecsListSource = TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.GetGlobalSpecs();

            for (var i = 0; i < lGlobalSpecsListSource.Count; i++)
            {
                var glbSpec = lGlobalSpecsListSource[i];
                if (glbSpec.Job != "")
                {
                    var lGlobalSpec = new GlobalSpec(glbSpec.Symbol, glbSpec.Value, glbSpec.Job, glbSpec.Comment);
                    lGlobalSpecsList.Add(lGlobalSpec);
                    if (IsCharacterizationSpec(glbSpec.Job))
                    {
                        lGlobalSpec = new GlobalSpec(glbSpec.Symbol, glbSpec.Value, glbSpec.Job + "_CHAR",
                            glbSpec.Comment);
                        lGlobalSpecsList.Add(lGlobalSpec);
                    }
                }
            }

            if (lGlobalSpecsList.Count <= 1) return false;

            lGlobalSpecsListSource.AddRange(lGlobalSpecsList);
            return true;
        }

        private bool IsCharacterizationSpec(string pJob)
        {
            return StaticSetting.JobMap.Any(job => job.Value[0] == pJob);
        }
    }
}