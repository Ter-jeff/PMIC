using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using PmicAutogen.Local;
using PmicAutogen.Inputs.TestPlan.Reader;
using IgxlData.Others;

namespace PmicAutogen.GenerateIgxl.PostAction.GenGlobal
{
    public class GenVddLevelsXSeqGlobalSpec
    {
        public bool ExtendGlobalSpec()
        {
            var lGlobalSpecsList = new List<GlobalSpec>();

            if (TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value == null) return false;

            var lGlobalSpecsListSource = TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.GetGlobalSpecs();
            VddLevelsSheet vddLvlSheet = StaticTestPlan.VddLevelsSheet;
            foreach(VddLevelsRow xRow in vddLvlSheet.xRows)
            {
                lGlobalSpecsList.AddRange(GenGlobalSpec(xRow,vddLvlSheet.UHvHasNA,vddLvlSheet.ULvHasNA));
            }

            if (lGlobalSpecsList.Count <= 0) return false;

            lGlobalSpecsListSource.AddRange(lGlobalSpecsList);
            return true;
        }

        private List<GlobalSpec> GenGlobalSpec(VddLevelsRow row,bool UHvHasNA,bool ULvHasNA)
        {
            List<GlobalSpec> glbSpecList = new List<GlobalSpec>();
            var prefix = string.IsNullOrEmpty(row.Nv) ? "VIN_0v_" : "VIN_" + row.Nv.Replace(".", "p") + "v_";

            string specName= prefix+ SpecFormat.GenGlbSpecSymbol(row.WsBumpName);

            GlobalSpec nvGblSpec = new GlobalSpec(specName, SpecFormat.GenSpecValueSingleValue(row.Nv));
            glbSpecList.Add(nvGblSpec);

            GlobalSpec lvGblSpec = new GlobalSpec(specName + "_LV", SpecFormat.GenSpecValueSingleValue(row.Lv));
            glbSpecList.Add(lvGblSpec);

            GlobalSpec hvGblSpec = new GlobalSpec(specName + "_HV", SpecFormat.GenSpecValueSingleValue(row.Hv));
            glbSpecList.Add(hvGblSpec);

            if(!UHvHasNA)
            {
                GlobalSpec uhvGblSpec = new GlobalSpec(specName + "_UHV", SpecFormat.GenSpecValueSingleValue(row.UHv));
                glbSpecList.Add(uhvGblSpec);
            }

            if(!ULvHasNA)
            {
                GlobalSpec ulvGblSpec = new GlobalSpec(specName + "_ULV", SpecFormat.GenSpecValueSingleValue(row.ULv));
                glbSpecList.Add(ulvGblSpec);
            }

            return glbSpecList;
        }
    }
}
