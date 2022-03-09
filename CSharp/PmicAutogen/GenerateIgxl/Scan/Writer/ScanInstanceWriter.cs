using System;
using System.Collections.Generic;
using System.Linq;
using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanInstanceWriter
    {
        protected string Block;
        protected string SheetName;

        public ScanInstanceWriter()
        {
            SheetName = "TestInst_Scan";
            Block = "Scan";
        }

        public virtual InstanceSheet WriteInstance(List<ProdCharRowScan> prodCharRowScans)
        {
            var sheetInstance = new InstanceSheet(SheetName);
            sheetInstance.AddHeaderFooter();

            foreach (var testInstance in prodCharRowScans)
            {
                var nv = GetInstanceRow(testInstance, sheetInstance.SheetName, "NV");
                if (nv != null)
                {
                    sheetInstance.AddRow(nv);
                    testInstance.InstanceRow = nv;
                }

                var hv = GetInstanceRow(testInstance, sheetInstance.SheetName, "HV");
                if (hv != null)
                {
                    sheetInstance.AddRow(hv);
                    if (LocalSpecs.HasUltraVoltageUHv)
                    {
                        hv = GetUltraInstanceRow(hv, testInstance, "HV");
                        sheetInstance.AddRow(hv);
                    }
                }
                var lv = GetInstanceRow(testInstance, sheetInstance.SheetName, "LV");
                if (lv != null)
                {
                    sheetInstance.AddRow(lv);
                    if (LocalSpecs.HasUltraVoltageULv)
                    {
                        lv = GetUltraInstanceRow(lv, testInstance, "LV");
                        sheetInstance.AddRow(lv);
                    }
                }
            }

            return sheetInstance;
        }

        protected void WriteItemCategory(ref InstanceRow row)
        {
            row.AcCategory = Block;
            row.DcCategory = Block;
        }

        protected string GetLevelFromPerformanceMode()
        {
            return "Levels_Func";
        }

        protected virtual InstanceRow GetInstanceRow(ProdCharRowScan prodCharRowScan, string sheetName,
            string selectorName)
        {
            if (!prodCharRowScan.PayloadList.Any()) return null;
            var instanceRow = new InstanceRow();
            instanceRow.SheetName = sheetName;
            instanceRow.Type = "VBT";
            WriteItemSelector(ref instanceRow, selectorName);
            instanceRow.TestName = ComCombine.CombineByUnderLine(prodCharRowScan.InstanceName, selectorName);
            instanceRow.PinLevels = GetLevelFromPerformanceMode();
            if (prodCharRowScan.PayloadList.Any())
                instanceRow.TimeSets = GetTimeSetName(prodCharRowScan.InitList, prodCharRowScan.PayloadList);
            WriteItemCategory(ref instanceRow);
            WriteArgsAndArgs(ref instanceRow, prodCharRowScan.PatSetName);


            instanceRow.InitList.AddRange(prodCharRowScan.InitList.Values.Select(x => x.PatternName));
            instanceRow.PayloadList.AddRange(prodCharRowScan.PayloadList.Select(x => x.PatternName));
            return instanceRow;
        }

        protected virtual InstanceRow GetUltraInstanceRow(InstanceRow baserow, ProdCharRowScan prodCharRowScan, 
                                                    string selectorName)
        {
            InstanceRow ultraRow = new InstanceRow();
            ultraRow.SheetName = baserow.SheetName;
            ultraRow.Type = baserow.Type;
            ultraRow.AcSelector = baserow.AcSelector;
            ultraRow.DcSelector = baserow.DcSelector;

            if (selectorName.Equals("Hv",StringComparison.InvariantCultureIgnoreCase))
            {
                ultraRow.TestName = ComCombine.CombineByUnderLine(prodCharRowScan.InstanceName, "UHV");
            }
            else
            {
                ultraRow.TestName = ComCombine.CombineByUnderLine(prodCharRowScan.InstanceName, "ULV");
            }
            ultraRow.PinLevels = baserow.PinLevels;
            ultraRow.TimeSets = baserow.TimeSets;
            ultraRow.AcCategory = baserow.AcCategory;
            ultraRow.DcCategory = LocalSpecs.GetUltraCategory(Block);

            ultraRow.Name = baserow.Name;
            ultraRow.ArgList = baserow.ArgList;
            for (int i = 0; i < baserow.Args.Count; i++)
            {
                ultraRow.Args.Add(baserow.Args[i]);
            }

            for (int i = 0; i < baserow.InitList.Count; i++)
            {
                ultraRow.InitList.Add(baserow.InitList[i]);
            }

            for (int i = 0; i < baserow.PayloadList.Count; i++)
            {
                ultraRow.PayloadList.Add(baserow.PayloadList[i]);
            }
            return ultraRow;
        }

        protected void WriteItemSelector(ref InstanceRow row, string selectorName)
        {
            row.AcSelector = "Typ";
            switch (selectorName)
            {
                case "HV":
                    row.DcSelector = "Max";

                    break;
                case "LV":
                    row.DcSelector = "Min";

                    break;
                case "NV":
                    row.DcSelector = "Typ";

                    break;
            }
        }

        protected void WriteArgsAndArgs(ref InstanceRow instanceRow, string instanceName)
        {
            var vbt = TestProgram.VbtFunctionLib.GetFunctionByName("Functional_T_updated");
            instanceRow.Name = vbt.FunctionName;
            instanceRow.ArgList = vbt.Parameters;
            instanceRow.Args = vbt.Args;
            vbt.SetParamValue("Patterns", instanceName);
            vbt.SetParamValue("ResultMode", "0");
            vbt.SetParamValue("RelayMode", "1");
        }

        protected string GetTimeSetName(Dictionary<int, PatternWithMode> initList, List<PatternWithMode> payloadList)
        {
            var timeSets = new List<string>();
            foreach (var init in initList)
                timeSets.Add(InputFiles.PatternListMap.GetTimeSet(init.Value.PatternName));
            foreach (var payload in payloadList)
                timeSets.Add(InputFiles.PatternListMap.GetTimeSet(payload.PatternName));

            var timeSet = string.Join(",", timeSets.Distinct(StringComparer.CurrentCultureIgnoreCase));
            return string.IsNullOrEmpty(timeSet) ? "NA" : timeSet;
        }
    }
}