using CommonLib.Enum;
using CommonLib.WriteMessage;
using PmicAutogen.GenerateIgxl.Mbist.Writer;
using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Inputs.ScghFile.ProChar.Business;
using PmicAutogen.Inputs.ScghFile.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Mbist
{
    public class MbistMain : MainBase
    {
        public List<ProdCharRowMbist> ProdCharRowMbists { get; set; }

        public void WorkFlow()
        {
            try
            {
                Response.Report("Generating Mbist Files ...", EnumMessageLevel.General, 0);

                if (StaticScgh.ScghMbistSheet == null)
                    return;

                ReadFiles(StaticScgh.ScghMbistSheet);

                new SetNopMain().SetNop(ProdCharRowMbists);

                GenFlow(ProdCharRowMbists);

                GenInstance(ProdCharRowMbists);

                GenBinTableRows(ProdCharRowMbists);

                GenPatSet(ProdCharRowMbists);

                GenCharacterization(ProdCharRowMbists);

                TestProgram.IgxlWorkBk.AddIgxlSheets(IgxlSheets);

                Response.Report("Mbist Completed!", EnumMessageLevel.General, 100);
            }
            catch (Exception e)
            {
                Response.Report("Meet an Error in Mbist: " + e.Message, EnumMessageLevel.Error, 100);
            }
        }

        private void ReadFiles(ProdCharSheet scgWorksheet)
        {
            var rowList = StaticScgh.ScghMbistSheet.RowList;
            var mbistManagerData = new MbistPatSetConstructor(rowList);
            ProdCharRowMbists = mbistManagerData.WorkFlow();
        }

        private void GenFlow(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var writeFlowTableMbist = new MbistFlowTableWriter();
            var subFlowSheet = writeFlowTableMbist.GenFlowSheet(prodCharRowMbists);
            IgxlSheets.Add(subFlowSheet, FolderStructure.DirMbist);
        }

        private void GenBinTableRows(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var binTableWriter = new MbistBinTableWriter();
            var binTableRows = binTableWriter.WriteBinTableRows(prodCharRowMbists);
            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            binTable.AddRows(binTableRows);
        }

        private void GenInstance(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var writeInstanceMbist = new MbistInstanceWriter();
            var instanceSheet = writeInstanceMbist.WriteInstance(prodCharRowMbists);
            IgxlSheets.Add(instanceSheet, FolderStructure.DirMbist);
        }

        private void GenPatSet(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var writePatternSet = new MbistPatSetWriter();
            if (prodCharRowMbists.Count != 0)
            {
                var patSetSheet = writePatternSet.WritePatSet(prodCharRowMbists);
                IgxlSheets.Add(patSetSheet, FolderStructure.DirMbist);
            }
        }

        private void GenCharacterization(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var charSheet = TestProgram.IgxlWorkBk.GetCharSheet(PmicConst.CharSetUpPmic);
            var characterization = new MbistCharacterization();
            var charSetups = characterization.WorkFlow(prodCharRowMbists);
            foreach (var charSetup in charSetups)
                if (!charSheet.CharSetups.Exists(p =>
                        p.SetupName.Equals(charSetup.SetupName, StringComparison.CurrentCultureIgnoreCase)))
                    charSheet.AddRow(charSetup);

            IgxlSheets.Add(charSheet, FolderStructure.DirDigitalScanMbist);
        }
    }
}