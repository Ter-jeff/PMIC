using CommonLib.Enum;
using CommonLib.WriteMessage;
using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Inputs.ScghFile.ProChar.Business;
using PmicAutogen.Inputs.ScghFile.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Scan
{
    public class ScanMain : MainBase
    {
        public List<ProdCharRowScan> ProdCharRowScans { get; set; }

        public void WorkFlow()
        {
            try
            {
                Response.Report("Generating Scan Files ...", EnumMessageLevel.General, 0);

                if (StaticScgh.ScghScanSheet == null)
                    return;

                ReadFiles(StaticScgh.ScghScanSheet);

                new SetNopMain().SetNop(ProdCharRowScans);

                GenFlow(ProdCharRowScans);

                GenInstance(ProdCharRowScans);

                GenBinTableRows(ProdCharRowScans);

                GenPatSet(ProdCharRowScans);

                GenCharacterization(ProdCharRowScans);

                TestProgram.IgxlWorkBk.AddIgxlSheets(IgxlSheets);

                Response.Report("Scan Completed!", EnumMessageLevel.General, 100);
            }
            catch (Exception e)
            {
                var message = "Scan AutoGen Failed: " + e.Message;
                Response.Report(message, EnumMessageLevel.Error, 100);
            }
        }

        private void ReadFiles(ProdCharSheet scgWorksheet)
        {
            var rowList = StaticScgh.ScghScanSheet.RowList;
            var scanManagerData = new ScanPatSetConstructor(rowList);
            ProdCharRowScans = scanManagerData.WorkFlow();
        }

        private void GenFlow(List<ProdCharRowScan> prodCharRowScans)
        {
            var writeFlowTableScan = new ScanFlowTableWriter();
            var subFlowSheet = writeFlowTableScan.GenFlowSheet(prodCharRowScans);
            IgxlSheets.Add(subFlowSheet, FolderStructure.DirScan);
        }

        private void GenBinTableRows(List<ProdCharRowScan> prodCharRowScans)
        {
            var binTableWriter = new ScanBinTableWriter();
            var binTableRows = binTableWriter.WriteBinTableRows(prodCharRowScans);
            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            binTable.AddRows(binTableRows);
        }

        private void GenInstance(List<ProdCharRowScan> prodCharRowScans)
        {
            var writeInstanceScan = new ScanInstanceWriter();
            var instanceSheet = writeInstanceScan.WriteInstance(prodCharRowScans);
            IgxlSheets.Add(instanceSheet, FolderStructure.DirScan);
        }

        private void GenPatSet(List<ProdCharRowScan> prodCharRowScans)
        {
            var writePatternSet = new ScanPatSetWriter();
            if (prodCharRowScans.Count != 0)
            {
                var patSetSheet = writePatternSet.WritePatSet(prodCharRowScans);
                IgxlSheets.Add(patSetSheet, FolderStructure.DirScan);
            }
        }

        private void GenCharacterization(List<ProdCharRowScan> prodCharRowScans)
        {
            var charSheet = TestProgram.IgxlWorkBk.GetCharSheet(PmicConst.CharSetUpPmic);
            var characterization = new ScanCharacterization();
            var charSetups = characterization.WorkFlow(prodCharRowScans);
            foreach (var charSetup in charSetups)
                if (!charSheet.CharSetups.Exists(p =>
                        p.SetupName.Equals(charSetup.SetupName, StringComparison.CurrentCultureIgnoreCase)))
                    charSheet.AddRow(charSetup);

            IgxlSheets.Add(charSheet, FolderStructure.DirDigitalScanMbist);
        }
    }
}