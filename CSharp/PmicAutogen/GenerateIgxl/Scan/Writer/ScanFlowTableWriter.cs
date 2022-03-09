using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanFlowTableWriter
    {
        private const string SheetName = "Flow_Scan";
        private const string Block = "SCAN";

        public SubFlowSheet WriteFlow(List<ProdCharRowScan> prodCharRowScans)
        {
            var subFlowSheet = new SubFlowSheet(SheetName);
            WriteSheet(prodCharRowScans, ref subFlowSheet);
            return subFlowSheet;
        }

        private void WriteSheet(List<ProdCharRowScan> prodCharRowScans, ref SubFlowSheet flowSheet)
        {
            if (!prodCharRowScans.Any())
                return;

            flowSheet.AddStartRows(SubFlowSheet.Ttime);

            foreach (var instance in prodCharRowScans)
                WriteItems(new List<ProdCharRowScan> { instance }, ref flowSheet);

            flowSheet.AddEndRows(SubFlowSheet.Ttime, false);
        }

        private void WriteItems(List<ProdCharRowScan> prodCharRowScans, ref SubFlowSheet sheetFlow)
        {
            var instances = prodCharRowScans.FindAll(p =>
                p.Nop == false || p.NopType == NopType.BlankInit || p.NopType == NopType.WrongTimeSet ||
                p.NopType == NopType.NonUsage).ToList();
            if (instances.Count < 1) return;

            foreach (var instance in instances)
            {
                WriteFlowRow(ref sheetFlow, instance, "HV");

                WriteFlowRow(ref sheetFlow, instance, "NV");

                WriteFlowRow(ref sheetFlow, instance, "LV");

                if (LocalSpecs.HasUltraVoltageUHv)
                {
                    WriteFlowRow(ref sheetFlow, instance, "UHV");
                }

                if (LocalSpecs.HasUltraVoltageULv)
                {
                    WriteFlowRow(ref sheetFlow, instance, "ULV");
                }

                if (!(instance.SupplyVoltage.Equals("") || instance.SupplyVoltage.Equals("N/A")))
                {
                    var singlePins = instance.GetSinglePins(instance.SupplyVoltage);

                    foreach (var pin in singlePins)
                    {
                        var setupName = instance.Get1DCharNamePeriod(pin);
                        Write1DCharItem(ref sheetFlow, instance, "NV", setupName);
                    }

                    var trackingGroups = instance.GetTrackingGroup(instance.SupplyVoltage);
                    foreach (var pins in trackingGroups)
                        foreach (var pin in pins)
                        {
                            var setupName = instance.Get2DCharNamePeriod(pin, "SCAN");
                            Write2DCharItem(ref sheetFlow, instance, "NV", setupName);
                        }
                }

                WriteBinTable(ref sheetFlow, instances.First());
            }
        }

        private void WriteFlowRow(ref SubFlowSheet flowTable, ProdCharRowScan prodCharRowScan, string binType)
        {
            var itemInstanceName = prodCharRowScan.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.OpCode = "Test";
            rowTemp.Parameter = ComCombine.CombineByUnderLine(itemInstanceName, binType);
            var binPayLoadName = prodCharRowScan.PayLoadName.GetSortPatNameForBinTable();
            rowTemp.FailAction = string.Format("F_{0}_{1}", binPayLoadName, binType);
            flowTable.AddRow(rowTemp);
        }

        private void WriteBinTable(ref SubFlowSheet flowTable, ProdCharRowScan prodCharRowScan)
        {
            var binPayLoadName = prodCharRowScan.PayLoadName.GetSortPatNameForBinTable();
            var binHnlvName = string.Format("Bin_{0}_{1}_{2}", Block, binPayLoadName, "HNLV");
            var binNlvName = string.Format("Bin_{0}_{1}_{2}", Block, binPayLoadName, "NLV");
            var binHlvName = string.Format("Bin_{0}_{1}_{2}", Block, binPayLoadName, "HLV");
            var binHnvName = string.Format("Bin_{0}_{1}_{2}", Block, binPayLoadName, "HNV");
            var binHvName = string.Format("Bin_{0}_{1}_{2}", Block, binPayLoadName, "HV");
            var binNvName = string.Format("Bin_{0}_{1}_{2}", Block, binPayLoadName, "NV");
            var binLvName = string.Format("Bin_{0}_{1}_{2}", Block, binPayLoadName, "LV");

            WriteBinTableItem(ref flowTable, binHnlvName);
            WriteBinTableItem(ref flowTable, binNlvName);
            WriteBinTableItem(ref flowTable, binHlvName);
            WriteBinTableItem(ref flowTable, binHnvName);
            WriteBinTableItem(ref flowTable, binHvName);
            WriteBinTableItem(ref flowTable, binNvName);
            WriteBinTableItem(ref flowTable, binLvName);

            if (LocalSpecs.HasUltraVoltageUHv)
            {
                var binUHvName = string.Format("Bin_{0}_{1}_{2}", Block, binPayLoadName, "UHV");
                WriteBinTableItem(ref flowTable, binUHvName);
            }

            if (LocalSpecs.HasUltraVoltageULv)
            {
                var binULvName = string.Format("Bin_{0}_{1}_{2}", Block, binPayLoadName, "ULV");
                WriteBinTableItem(ref flowTable, binULvName);
            }
        }

        private void Write1DCharItem(ref SubFlowSheet flowTable, ProdCharRowScan prodCharRowScan, string binType,
            string setupName)
        {
            if (string.IsNullOrEmpty(prodCharRowScan.SupplyVoltage))
                return;
            var itemInstanceName = prodCharRowScan.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.Enable = "B_Debug_1D_SHMOO";
            rowTemp.OpCode = FlowRow.OpCodeCharacterize;
            rowTemp.Parameter = ComCombine.CombineByUnderLine(itemInstanceName, binType) + " " + setupName;
            flowTable.AddRow(rowTemp);
        }

        private void Write2DCharItem(ref SubFlowSheet flowTable, ProdCharRowScan pInstance, string pBinType,
            string setupName)
        {
            if (string.IsNullOrEmpty(pInstance.SupplyVoltage))
                return;
            var itemInstanceName = pInstance.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.Enable = "B_Debug_2D_SHMOO";
            rowTemp.OpCode = FlowRow.OpCodeCharacterize;
            rowTemp.Parameter = ComCombine.CombineByUnderLine(itemInstanceName, pBinType) + " " + setupName;
            flowTable.AddRow(rowTemp);
        }

        private void WriteBinTableItem(ref SubFlowSheet flowTable, string parameter)
        {
            var flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeBinTable;
            flowRow.Parameter = parameter;
            flowRow.Enable = FlowRow.OpCodeBinTable;
            flowTable.AddRow(flowRow);
        }
    }
}