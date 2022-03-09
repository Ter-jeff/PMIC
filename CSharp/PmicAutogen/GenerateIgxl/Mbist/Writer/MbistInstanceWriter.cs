using System.Collections.Generic;
using System.Linq;
using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class MbistInstanceWriter : ScanInstanceWriter
    {
        public MbistInstanceWriter()
        {
            SheetName = "TestInst_Mbist";
            Block = "Mbist";
        }

        public InstanceSheet WriteInstance(List<ProdCharRowMbist> instancesNameList)
        {
            var sheetInstance = new InstanceSheet(SheetName);
            sheetInstance.AddHeaderFooter();

            foreach (var testInstance in instancesNameList)
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

        protected InstanceRow GetInstanceRow(ProdCharRowMbist instance, string sheetName, string selectorName)
        {
            if (!instance.PayloadList.Any()) return null;
            var instanceRow = new InstanceRow();
            instanceRow.SheetName = sheetName;
            instanceRow.Type = "VBT";
            WriteItemSelector(ref instanceRow, selectorName);
            instanceRow.TestName = ComCombine.CombineByUnderLine(instance.InstanceName, selectorName);
            instanceRow.PinLevels = GetLevelFromPerformanceMode();
            if (instance.PayloadList.Any())
                instanceRow.TimeSets = GetTimeSetName(instance.InitList, instance.PayloadList);
            WriteItemCategory(ref instanceRow);
            WriteArgsAndArgs(ref instanceRow, instance.PatSetName);
            instanceRow.InitList.AddRange(instance.InitList.Values.Select(x => x.PatternName));
            instanceRow.PayloadList.AddRange(instance.PayloadList.Select(x => x.PatternName));
            return instanceRow;
        }

        protected InstanceRow GetUltraInstanceRow(InstanceRow baserow, ProdCharRowMbist instance, string selectorName)
        {
            InstanceRow ultraRow = new InstanceRow();
            ultraRow.SheetName = baserow.SheetName;
            ultraRow.Type = baserow.Type;
            ultraRow.AcSelector = baserow.AcSelector;
            ultraRow.DcSelector = baserow.DcSelector;

            if (selectorName.Equals("Hv",System.StringComparison.InvariantCultureIgnoreCase))
            {
                ultraRow.TestName = ComCombine.CombineByUnderLine(instance.InstanceName, "UHV");
            }
            else
            {
                ultraRow.TestName = ComCombine.CombineByUnderLine(instance.InstanceName, "ULV");
            }
            ultraRow.PinLevels = baserow.PinLevels;
            ultraRow.TimeSets = baserow.TimeSets;
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
    }
}