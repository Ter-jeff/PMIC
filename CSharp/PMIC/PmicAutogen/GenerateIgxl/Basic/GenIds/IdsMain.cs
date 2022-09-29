using CommonLib.Enum;
using CommonLib.Extension;
using CommonLib.WriteMessage;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Basic.GenIds
{
    internal class IdsMain
    {
        public IdsMain(PmicIdsSheet pmicIdsSheet)
        {
            PmicIdsSheet = pmicIdsSheet;
        }

        private PmicIdsSheet PmicIdsSheet { get; }

        internal Dictionary<IgxlSheet, string> Workflow()
        {
            var instanceSheet = PmicIdsSheet.GenInsSheet("TestInst_" + PmicConst.PmicIds);

            var subFlowSheets = PmicIdsSheet.GenSubFlowSheets("Flow_" + PmicConst.PmicIds);

            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            var binTableRows = PmicIdsSheet.GenBinTableRows();
            binTable.AddRows(binTableRows);

            var fileNameTxt = Path.Combine(FolderStructure.DirDc, PmicConst.PmicIds + ".txt");
            var notMatchedPins = GetNotMatchedPinsWithVddLevels(PmicIdsSheet);
            InputFiles.TestPlanWorkbook.Worksheets[PmicConst.PmicIds]
                .ExportToTxt(fileNameTxt, extraRows: notMatchedPins);
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirDc, PmicConst.PmicIds);

            var igxlSheets = new Dictionary<IgxlSheet, string>();
            igxlSheets.Add(instanceSheet, FolderStructure.DirDc);
            foreach (var flow in subFlowSheets)
                igxlSheets.Add(flow, FolderStructure.DirDc);
            return igxlSheets;
        }

        private Dictionary<int, List<string>> GetNotMatchedPinsWithVddLevels(PmicIdsSheet sheet)
        {
            var notMatchedPins = new Dictionary<int, List<string>>();
            var idsPins = new List<string>();
            var vddLevelPins = new List<string>();
            notMatchedPins.Add(0, vddLevelPins);
            var vddLevelsSheet = StaticTestPlan.VddLevelsSheet;
            var measurePins = sheet.GetMeasurePins();
            var wsBumpNames = vddLevelsSheet.Rows.Select(o => o.WsBumpName.Trim()).Distinct().ToList();
            foreach (var wsBumpName in wsBumpNames)
                if (!measurePins.Contains(wsBumpName))
                    vddLevelPins.Add(wsBumpName);

            foreach (var measurePin in measurePins)
                if (!wsBumpNames.Contains(measurePin))
                    idsPins.Add(measurePin);
            if (idsPins.Any())
            {
                var pinNames = string.Join(",", idsPins.Distinct());
                var pinNameStr = string.Format(@"[{0}] not exist in VDD_Levels sheet.", pinNames);
                Response.Report(
                    "IDS sheet measure pins not match with VDD_Levels sheet." + Environment.NewLine + pinNameStr,
                    EnumMessageLevel.Warning, 90);
            }

            return notMatchedPins;
        }
    }
}