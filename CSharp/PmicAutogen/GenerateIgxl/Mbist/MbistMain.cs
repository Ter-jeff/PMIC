using AutomationCommon.DataStructure;
using PmicAutogen.GenerateIgxl.Mbist.Writer;
using PmicAutogen.InputPackages;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Inputs.ScghFile.ProChar.Business;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Mbist
{
    public class MbistMain : MainBase
    {
        public void WorkFlow()
        {
            try
            {
                Initialize();

                Response.Report("Generating Mbist Files ...", MessageLevel.General, 60);
                MbistAutoGen();

                AddIgxlSheets(IgxlSheets);
                Response.Report("Mbist Completed!", MessageLevel.General, 100);
            }
            catch (Exception e)
            {
                Response.Report("Meet an Error in Mbist: " + e.Message, MessageLevel.Error, 100);
            }
        }

        protected void MbistAutoGen()
        {
            if (StaticScgh.ScghMbistSheet != null)
            {
                var rowList = StaticScgh.ScghMbistSheet.RowList;
                var mbistManagerData = new MbistPatSetConstructor(rowList);
                var prodCharRowMbists = mbistManagerData.WorkFlow();

                GenAllMbistResult(prodCharRowMbists);
            }
        }

        protected void GenFlow(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var writeFlowTableMbist = new MbistFlowTableWriter();
            var subFlowSheet = writeFlowTableMbist.WriteFlow(prodCharRowMbists);
            IgxlSheets.Add(subFlowSheet, FolderStructure.DirMbist);
        }

        protected void GenBinTableRows(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var binTableWriter = new MbistBinTableWriter();
            var binTableRows = binTableWriter.WriteBinTable(prodCharRowMbists);
            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            binTable.AddRows(binTableRows);
        }

        protected void GenInstance(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var writeInstanceMbist = new MbistInstanceWriter();
            var instanceSheet = writeInstanceMbist.WriteInstance(prodCharRowMbists);
            IgxlSheets.Add(instanceSheet, FolderStructure.DirMbist);
        }

        protected virtual void GenPatSet(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var writePatternSet = new MbistPatSetWriter();
            if (prodCharRowMbists.Count != 0)
            {
                var patSetSheet = writePatternSet.WritePatSet(prodCharRowMbists);
                IgxlSheets.Add(patSetSheet, FolderStructure.DirMbist);
            }
        }

        public void GenCharacterization(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var charSheet = TestProgram.IgxlWorkBk.GetCharSheet(PmicConst.CharSetUpPmic);
            var characterization = new MbistCharacterization();
            var charSetups = characterization.WorkFlow(prodCharRowMbists);
            foreach (var charSetup in charSetups)
                if (!charSheet.CharSetups.Exists(p =>
                    p.SetupName.Equals(charSetup.SetupName, StringComparison.CurrentCultureIgnoreCase)))
                    charSheet.AddRow(charSetup);

            IgxlSheets.Add(charSheet, FolderStructure.DirDevChar);
        }

        public void GenAllMbistResult(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var setMbistNop = new SetMbistNop();
            setMbistNop.SetNop(prodCharRowMbists);

            GenFlow(prodCharRowMbists);

            GenInstance(prodCharRowMbists);

            GenBinTableRows(prodCharRowMbists);

            GenPatSet(prodCharRowMbists);

            GenCharacterization(prodCharRowMbists);
        }
    }
}