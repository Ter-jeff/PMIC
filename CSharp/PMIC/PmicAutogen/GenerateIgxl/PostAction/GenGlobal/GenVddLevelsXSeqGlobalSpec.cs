using IgxlData.IgxlBase;
using IgxlData.Others;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.PostAction.GenGlobal
{
    public class GenVddLevelsXSeqGlobalSpec
    {
        public bool ExtendGlobalSpec()
        {
            var lGlobalSpecsList = new List<GlobalSpec>();

            if (TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value == null) return false;

            var lGlobalSpecsListSource = TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.GetGlobalSpecs();
            var vddLvlSheet = StaticTestPlan.VddLevelsSheet;
            foreach (var xRow in vddLvlSheet.XRows)
                lGlobalSpecsList.AddRange(GenGlobalSpec(xRow, vddLvlSheet.UHvAllNa, vddLvlSheet.ULvAllNa));

            if (lGlobalSpecsList.Count <= 0) return false;

            lGlobalSpecsListSource.AddRange(lGlobalSpecsList);
            return true;
        }

        private List<GlobalSpec> GenGlobalSpec(VddLevelsRow row, bool uHvHasNa, bool uLvHasNa)
        {
            var glbSpecList = new List<GlobalSpec>();
            var prefix = string.IsNullOrEmpty(row.Nv) ? "VIN_0v_" : "VIN_" + row.Nv.Replace(".", "p") + "v_";

            var specName = prefix + SpecFormat.GenGlbSpecSymbol(row.WsBumpName);

            var nvGblSpec = new GlobalSpec(specName, SpecFormat.GenSpecValueSingleValue(row.Nv));
            glbSpecList.Add(nvGblSpec);

            var lvGblSpec = new GlobalSpec(specName + "_LV", SpecFormat.GenSpecValueSingleValue(row.Lv));
            glbSpecList.Add(lvGblSpec);

            var hvGblSpec = new GlobalSpec(specName + "_HV", SpecFormat.GenSpecValueSingleValue(row.Hv));
            glbSpecList.Add(hvGblSpec);

            if (!uHvHasNa)
            {
                var uhvGblSpec = new GlobalSpec(specName + "_UHV", SpecFormat.GenSpecValueSingleValue(row.UHv));
                glbSpecList.Add(uhvGblSpec);
            }

            if (!uLvHasNa)
            {
                var ulvGblSpec = new GlobalSpec(specName + "_ULV", SpecFormat.GenSpecValueSingleValue(row.ULv));
                glbSpecList.Add(ulvGblSpec);
            }

            return glbSpecList;
        }
    }
}