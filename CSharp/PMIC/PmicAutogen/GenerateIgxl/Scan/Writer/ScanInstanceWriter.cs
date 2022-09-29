using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanInstanceWriter : ScanWriter
    {
        protected string SheetName;

        public ScanInstanceWriter()
        {
            SheetName = "TestInst_Scan";
        }

        public InstanceSheet WriteInstance(IEnumerable<ProdCharRow> prodCharRows)
        {
            var sheetInstance = new InstanceSheet(SheetName);

            sheetInstance.AddHeaderFooter();

            var voltages = new List<string> { Nv };
            voltages.Add(Hv);
            if (LocalSpecs.HasUltraVoltageUHv)
                voltages.Add(UHv);
            voltages.Add(Lv);
            if (LocalSpecs.HasUltraVoltageULv)
                voltages.Add(ULv);

            foreach (var testInstance in prodCharRows)
                foreach (var voltage in voltages)
                {
                    var row = GetInstanceRow(testInstance, sheetInstance.SheetName, voltage);
                    if (row != null)
                    {
                        sheetInstance.AddRow(row);
                        testInstance.InstanceRow = row;
                    }
                }

            return sheetInstance;
        }

        private InstanceRow GetInstanceRow(ProdCharRow prodCharRow, string sheetName, string voltage)
        {
            if (!prodCharRow.PayloadList.Any()) return null;
            var instanceRow = new InstanceRow();
            instanceRow.SheetName = sheetName;
            instanceRow.Type = "VBT";

            instanceRow.TestName = Combine.CombineByUnderLine(prodCharRow.InstanceName, voltage);
            instanceRow.PinLevels = GetLevelFromPerformanceMode();
            if (prodCharRow.PayloadList.Any())
                instanceRow.TimeSets = GetTimeSetName(prodCharRow.InitList, prodCharRow.PayloadList);

            WriteCategory(ref instanceRow, voltage);
            WriteSelector(ref instanceRow, voltage);
            WriteArgsAndArgs(ref instanceRow, prodCharRow.PatSetName);

            instanceRow.InitList.AddRange(prodCharRow.InitList.Values.Select(x => x.PatternName));
            instanceRow.PayloadList.AddRange(prodCharRow.PayloadList.Select(x => x.PatternName));
            return instanceRow;
        }

        private void WriteSelector(ref InstanceRow row, string voltage)
        {
            row.AcSelector = "Typ";
            switch (voltage)
            {
                case Hv:
                case UHv:
                    row.DcSelector = "Max";
                    break;
                case Lv:
                case ULv:
                    row.DcSelector = "Min";
                    break;
                case Nv:
                    row.DcSelector = "Typ";
                    break;
            }
        }

        private void WriteCategory(ref InstanceRow row, string voltage)
        {
            row.AcCategory = Block;
            if (voltage.Equals(ULv, StringComparison.CurrentCultureIgnoreCase) ||
                voltage.Equals(UHv, StringComparison.CurrentCultureIgnoreCase))
                row.DcCategory = LocalSpecs.GetUltraCategory(Block);
            else
                row.DcCategory = Block;
        }

        private string GetLevelFromPerformanceMode()
        {
            return "Levels_Func";
        }

        private void WriteArgsAndArgs(ref InstanceRow instanceRow, string instanceName)
        {
            var vbt = TestProgram.VbtFunctionLib.GetFunctionByName("Functional_T_updated");
            instanceRow.Name = vbt.FunctionName;
            instanceRow.ArgList = vbt.Parameters;
            instanceRow.Args = vbt.Args;
            vbt.SetParamValue("Patterns", instanceName);
            vbt.SetParamValue("ResultMode", "0");
            vbt.SetParamValue("RelayMode", "1");
        }

        private string GetTimeSetName(Dictionary<int, PatternWithMode> initList, List<PatternWithMode> payloads)
        {
            var timeSets = new List<string>();
            foreach (var init in initList)
                timeSets.Add(InputFiles.PatternListMap.GetTimeSet(init.Value.PatternName));
            foreach (var payload in payloads)
                timeSets.Add(InputFiles.PatternListMap.GetTimeSet(payload.PatternName));

            var timeSet = string.Join(",", timeSets.Distinct(StringComparer.CurrentCultureIgnoreCase));
            return string.IsNullOrEmpty(timeSet) ? "NA" : timeSet;
        }
    }
}