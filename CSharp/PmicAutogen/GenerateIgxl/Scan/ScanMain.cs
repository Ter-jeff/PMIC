using AutomationCommon.DataStructure;
using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.InputPackages;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Inputs.ScghFile.ProChar.Business;
using PmicAutogen.Inputs.ScghFile.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Scan
{
    public class ScanMain : MainBase
    {
        public void WorkFlow()
        {
            try
            {
                Initialize();

                Response.Report("Generating Scan Files ...", MessageLevel.General, 60);
                ScanAutoGen();

                AddIgxlSheets(IgxlSheets);
                Response.Report("Scan Completed!", MessageLevel.General, 100);
            }
            catch (Exception e)
            {
                var message = "Scan AutoGen Failed: " + e.Message;
                Response.Report(message, MessageLevel.Error, 100);
            }
        }

        protected void ScanAutoGen()
        {
            var scgWorksheet = InputFiles.ScghWorkbook.Worksheets.FirstOrDefault(s => s.Name.ToUpper().EndsWith(PmicConst.ScghScan));
            if (scgWorksheet != null)
            {
                var sheetReader = new ProdCharSheetReader();
                var prodCharSheet = sheetReader.ReadScghSheet(scgWorksheet);

                var rowList = prodCharSheet.RowList;
                var scanManagerData = new ScanPatSetConstructor(rowList);
                var prodCharRowScans = scanManagerData.WorkFlow();

                GenAllScanResult(prodCharRowScans);
            }
        }

        protected void GenFlow(List<ProdCharRowScan> prodCharRowScans)
        {
            var writeFlowTableScan = new ScanFlowTableWriter();
            var subFlowSheet = writeFlowTableScan.WriteFlow(prodCharRowScans);
            IgxlSheets.Add(subFlowSheet, FolderStructure.DirScan);
        }

        protected void GenBinTableRows(List<ProdCharRowScan> prodCharRowScans)
        {
            var binTableWriter = new ScanBinTableWriter();
            var binTableRows = binTableWriter.WriteBinTable(prodCharRowScans);
            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            binTable.AddRows(binTableRows);
        }

        protected void GenInstance(List<ProdCharRowScan> prodCharRowScans)
        {
            var writeInstanceScan = new ScanInstanceWriter();
            var instanceSheet = writeInstanceScan.WriteInstance(prodCharRowScans);
            IgxlSheets.Add(instanceSheet, FolderStructure.DirScan);
        }

        protected virtual void GenPatSet(List<ProdCharRowScan> prodCharRowScans)
        {
            var writePatternSet = new ScanPatSetWriter();
            if (prodCharRowScans.Count != 0)
            {
                var patSetSheet = writePatternSet.WritePatSet(prodCharRowScans);
                IgxlSheets.Add(patSetSheet, FolderStructure.DirScan);
            }
        }

        public void GenCharacterization(List<ProdCharRowScan> prodCharRowScans)
        {
            var charSheet = TestProgram.IgxlWorkBk.GetCharSheet(PmicConst.CharSetUpPmic);
            var characterization = new ScanCharacterization();
            var charSetups = characterization.WorkFlow(prodCharRowScans);
            foreach (var charSetup in charSetups)
                if (!charSheet.CharSetups.Exists(p =>
                    p.SetupName.Equals(charSetup.SetupName, StringComparison.CurrentCultureIgnoreCase)))
                    charSheet.AddRow(charSetup);

            IgxlSheets.Add(charSheet, FolderStructure.DirDevChar);
        }

        public void GenAllScanResult(List<ProdCharRowScan> prodCharRowScans)
        {
            var setScanNop = new SetScanNop();
            setScanNop.SetNop(prodCharRowScans);

            GenFlow(prodCharRowScans);

            GenInstance(prodCharRowScans);

            GenBinTableRows(prodCharRowScans);

            GenPatSet(prodCharRowScans);

            GenCharacterization(prodCharRowScans);
        }

    }
}