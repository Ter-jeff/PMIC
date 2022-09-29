using CommonLib.Extension;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Basic.GenLeakage
{
    internal class LeakageMain
    {
        public LeakageMain(PmicLeakageSheet pmicLeakageSheet)
        {
            PmicLeakageSheet = pmicLeakageSheet;
        }

        private PmicLeakageSheet PmicLeakageSheet { get; }

        internal Dictionary<IgxlSheet, string> Workflow()
        {
            var sheetName = PmicLeakageSheet.SheetName;

            var instanceSheet = PmicLeakageSheet.GenInsSheet(PmicConst.TestInstDcLeakage);

            var subFlowSheet = PmicLeakageSheet.GenFlowSheet(PmicConst.FlowDcLeakage);

            var binTable = TestProgram.IgxlWorkBk.GetMainBinTblSheet();
            var binTableRows = PmicLeakageSheet.GenBinTableRows();
            binTable.AddRows(binTableRows);

            var fileNameTxt = Path.Combine(FolderStructure.DirDc, sheetName + ".txt");
            var skipColumns = new List<int>();
            skipColumns.Add(14);
            var notMatchedPins = GetNotTestedPinsFromPinmap(PmicLeakageSheet, StaticTestPlan.IoPinMapSheet);
            InputFiles.TestPlanWorkbook.Worksheets[sheetName]
                .ExportToTxt(fileNameTxt, "\t", null, skipColumns, notMatchedPins);
            TestProgram.NonIgxlSheetsList.Add(FolderStructure.DirDc, sheetName);

            var igxlSheets = new Dictionary<IgxlSheet, string>();
            igxlSheets.Add(instanceSheet, FolderStructure.DirDc);
            igxlSheets.Add(subFlowSheet, FolderStructure.DirDc);
            return igxlSheets;
        }

        private Dictionary<int, List<string>> GetNotTestedPinsFromPinmap(PmicLeakageSheet sheet,
            PinMapSheet ioPinMapSheet)
        {
            var notTestedPins = new Dictionary<int, List<string>>();
            var notTestIoPins = new List<string>();
            notTestedPins.Add(0, notTestIoPins);
            var notTestAnalogPins = new List<string>();
            notTestedPins.Add(1, notTestAnalogPins);

            var allIoPins = ioPinMapSheet.PinList.FindAll(
                    o => o.PinType.Equals("I/O", StringComparison.CurrentCultureIgnoreCase)).Select(o => o.PinName)
                .ToList();
            var allAnalogPins = ioPinMapSheet.PinList.FindAll(
                    o => o.PinType.Equals("Analog", StringComparison.CurrentCultureIgnoreCase)
                         && !o.PinName.EndsWith("_DM")
                         && !o.PinName.EndsWith("_DT"))
                .Select(o => o.PinName).ToList();

            var testedPins = new List<string>();
            foreach (var leakageRow in sheet.Rows)
            {
                var measurePins = leakageRow.MeasurePin.Split(',').ToList();
                foreach (var measurePin in measurePins)
                    testedPins.AddRange(GetAllMeasurePins(measurePin, ioPinMapSheet));
            }

            notTestIoPins.AddRange(allIoPins.Except(testedPins.Distinct()));
            notTestAnalogPins.AddRange(allAnalogPins.Except(testedPins.Distinct()));
            return notTestedPins;
        }

        private List<string> GetAllMeasurePins(string measurePin, PinMapSheet pinMapSheet)
        {
            var allMeasurePins = new List<string>();
            if (pinMapSheet.IsPinExist(measurePin))
            {
                allMeasurePins.Add(measurePin);
                return allMeasurePins;
            }

            if (pinMapSheet.IsGroupExist(measurePin))
                return pinMapSheet.GetPinsFromGroup(measurePin).Select(o => o.PinName).ToList();

            return allMeasurePins;
        }
    }
}