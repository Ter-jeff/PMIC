using AutomationCommon.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class MbistFlowTableWriter
    {
        private const string SheetName = "Flow_Mbist";
        private const string Block = "MBIST";

        public SubFlowSheet WriteFlow(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var subFlowSheet = new SubFlowSheet(SheetName);
            WriteSheet(prodCharRowMbists, ref subFlowSheet);
            return subFlowSheet;
        }

        private void WriteSheet(List<ProdCharRowMbist> prodCharRowMbists, ref SubFlowSheet flowSheet)
        {
            if (!prodCharRowMbists.Any())
                return;

            flowSheet.AddStartRows(SubFlowSheet.Ttime);

            foreach (var instance in prodCharRowMbists)
                WriteItems(new List<ProdCharRowMbist> {instance}, ref flowSheet);

            flowSheet.AddEndRows(SubFlowSheet.Ttime,false);
        }

        private void WriteItems(List<ProdCharRowMbist> prodCharRowMbists, ref SubFlowSheet sheetFlow)
        {
            var instances = prodCharRowMbists.FindAll(p =>
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

                if (!(instance.PeripheralVoltage.Equals("") || instance.PeripheralVoltage.Equals("N/A")))
                {
                    var singlePins = instance.GetSinglePins(instance.PeripheralVoltage);

                    foreach (var pin in singlePins)
                    {
                        var setupName = instance.Get1DCharNamePeriod(pin);
                        Write1DCharItem(ref sheetFlow, instance, "NV", setupName);
                    }

                    var trackingGroups = instance.GetTrackingGroup(instance.PeripheralVoltage);
                    foreach (var pins in trackingGroups)
                    foreach (var pin in pins)
                    {
                        var setupName = instance.Get2DCharNamePeriod(pin, "MBIST");
                        Write2DCharItem(ref sheetFlow, instance, "NV", setupName);
                    }
                }

                WriteBinTable(ref sheetFlow, instances.First());
            }
        }

        private void WriteFlowRow(ref SubFlowSheet flowTable, ProdCharRowMbist pInstance, string pBinType)
        {
            var itemInstanceName = pInstance.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.OpCode = "Test";
            rowTemp.Parameter = ComCombine.CombineByUnderLine(itemInstanceName, pBinType);
            var binPayLoadName = pInstance.PayLoadName.GetSortPatNameForBinTable();
            rowTemp.FailAction = string.Format("F_{0}_{1}", binPayLoadName, pBinType);
            flowTable.AddRow(rowTemp);
        }

        private void WriteBinTable(ref SubFlowSheet flowTable, ProdCharRowMbist pInstance)
        {
            var binPayLoadName = pInstance.PayLoadName.GetSortPatNameForBinTable();
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

        private void Write1DCharItem(ref SubFlowSheet flowTable, ProdCharRowMbist pInstance, string pBinType,
            string setupName)
        {
            if (string.IsNullOrEmpty(pInstance.PeripheralVoltage))
                return;
            var itemInstanceName = pInstance.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.Enable = "B_Debug_1D_SHMOO";
            rowTemp.OpCode = FlowRow.OpCodeCharacterize;
            rowTemp.Parameter = ComCombine.CombineByUnderLine(itemInstanceName, pBinType) + " " + setupName;
            flowTable.AddRow(rowTemp);
        }

        private void Write2DCharItem(ref SubFlowSheet flowTable, ProdCharRowMbist pInstance, string pBinType,
            string setupName)
        {
            if (string.IsNullOrEmpty(pInstance.PeripheralVoltage))
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
            var rowBin = new FlowRow();
            rowBin.OpCode = FlowRow.OpCodeBinTable;
            rowBin.Parameter = parameter;
            rowBin.Enable = FlowRow.OpCodeBinTable;
            flowTable.AddRow(rowBin);
        }
    }
}