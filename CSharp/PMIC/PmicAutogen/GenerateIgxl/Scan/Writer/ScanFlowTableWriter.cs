using CommonLib.Extension;
using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanFlowTableWriter : ScanWriter
    {
        public ScanFlowTableWriter()
        {
            SheetName = "Flow_Scan";
            Block = "SCAN";
        }

        protected string SheetName { get; set; }

        public SubFlowSheet GenFlowSheet(IEnumerable<ProdCharRow> prodCharRows)
        {
            var subFlowSheet = new SubFlowSheet(SheetName);
            var charRows = prodCharRows.ToList();
            if (!charRows.Any())
                return subFlowSheet;

            subFlowSheet.FlowRows.AddStartRows(subFlowSheet.SheetName, SubFlowSheet.Ttime);

            foreach (var instance in charRows)
                WriteBody(new List<ProdCharRow> { instance }, ref subFlowSheet);

            subFlowSheet.FlowRows.AddEndRows(subFlowSheet.SheetName, SubFlowSheet.Ttime, false);

            return subFlowSheet;
        }

        protected virtual void WriteBody(List<ProdCharRow> prodCharRows, ref SubFlowSheet subFlowSheet)
        {
            var instances = prodCharRows.FindAll(p =>
                p.Nop == false || p.NopType == NopType.BlankInit || p.NopType == NopType.WrongTimeSet ||
                p.NopType == NopType.NonUsage).ToList();
            if (instances.Count < 1) return;

            foreach (var prodCharRow in instances)
            {
                var instance = (ProdCharRowScan)prodCharRow;
                WriteFlowRow(ref subFlowSheet, instance, Hv);

                WriteFlowRow(ref subFlowSheet, instance, Nv);

                WriteFlowRow(ref subFlowSheet, instance, Lv);

                if (LocalSpecs.HasUltraVoltageUHv) WriteFlowRow(ref subFlowSheet, instance, UHv);

                if (LocalSpecs.HasUltraVoltageULv) WriteFlowRow(ref subFlowSheet, instance, ULv);

                if (!(instance.SupplyVoltage.Equals("") || instance.SupplyVoltage.Equals("N/A")))
                {
                    var singlePins = instance.GetSinglePins(instance.SupplyVoltage);

                    foreach (var pin in singlePins)
                    {
                        var setupName = instance.Get1DCharNamePeriod(pin);
                        Write1DChar(ref subFlowSheet, instance, "NV", setupName);
                    }

                    var trackingGroups = instance.GetTrackingGroup(instance.SupplyVoltage);
                    foreach (var pins in trackingGroups)
                        foreach (var pin in pins)
                        {
                            var setupName = instance.Get2DCharNamePeriod(pin, "SCAN");
                            Write2DChar(ref subFlowSheet, instance, "NV", setupName);
                        }
                }

                subFlowSheet.FlowRows.Add_A_Enable_MP_SBIN(BlockBinTableName);
                WriteBinTables(ref subFlowSheet, instances.First());
            }
        }

        protected void WriteFlowRow(ref SubFlowSheet subFlowSheet, ProdCharRow prodCharRow, string binType)
        {
            var itemInstanceName = prodCharRow.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.OpCode = FlowRow.OpCodeTest;
            rowTemp.Parameter = Combine.CombineByUnderLine(itemInstanceName, binType);
            var binPayLoadName = prodCharRow.PayLoadName.GetSortPatNameForBinTable();
            rowTemp.FailAction = GetFlagName(binPayLoadName, binType).AddBlockFlag(BlockBinTableName);
            subFlowSheet.AddRow(rowTemp);
        }

        protected void WriteBinTables(ref SubFlowSheet subFlowSheet, ProdCharRow prodCharRow)
        {
            var binPayLoadName = prodCharRow.PayLoadName.GetSortPatNameForBinTable();
            var binHnlvName = GetBinTableName(binPayLoadName, Hnlv);
            var binNlvName = GetBinTableName(binPayLoadName, Nlv);
            var binHlvName = GetBinTableName(binPayLoadName, Hlv);
            var binHnvName = GetBinTableName(binPayLoadName, Hnv);
            var binHvName = GetBinTableName(binPayLoadName, Hv);
            var binNvName = GetBinTableName(binPayLoadName, Nv);
            var binLvName = GetBinTableName(binPayLoadName, Lv);

            WriteBinTable(ref subFlowSheet, binHnlvName);
            WriteBinTable(ref subFlowSheet, binNlvName);
            WriteBinTable(ref subFlowSheet, binHlvName);
            WriteBinTable(ref subFlowSheet, binHnvName);
            WriteBinTable(ref subFlowSheet, binHvName);
            WriteBinTable(ref subFlowSheet, binNvName);
            WriteBinTable(ref subFlowSheet, binLvName);

            if (LocalSpecs.HasUltraVoltageUHv)
            {
                var binUHvName = GetBinTableName(binPayLoadName, UHv);
                WriteBinTable(ref subFlowSheet, binUHvName);
            }

            if (LocalSpecs.HasUltraVoltageULv)
            {
                var binULvName = GetBinTableName(binPayLoadName, ULv);
                WriteBinTable(ref subFlowSheet, binULvName);
            }
        }

        protected virtual void Write1DChar(ref SubFlowSheet subFlowSheet, ProdCharRow prodCharRow, string binType,
            string setupName)
        {
            var row = (ProdCharRowScan)prodCharRow;
            if (string.IsNullOrEmpty(row.SupplyVoltage))
                return;
            var itemInstanceName = row.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.Enable = "B_Debug_1D_SHMOO";
            rowTemp.OpCode = FlowRow.OpCodeCharacterize;
            rowTemp.Parameter = Combine.CombineByUnderLine(itemInstanceName, binType) + " " + setupName;
            subFlowSheet.AddRow(rowTemp);
        }

        protected virtual void Write2DChar(ref SubFlowSheet subFlowSheet, ProdCharRow prodCharRow, string binType,
            string setupName)
        {
            var row = (ProdCharRowScan)prodCharRow;
            if (string.IsNullOrEmpty(row.SupplyVoltage))
                return;
            var itemInstanceName = row.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.Enable = "B_Debug_2D_SHMOO";
            rowTemp.OpCode = FlowRow.OpCodeCharacterize;
            rowTemp.Parameter = Combine.CombineByUnderLine(itemInstanceName, binType) + " " + setupName;
            subFlowSheet.AddRow(rowTemp);
        }

        private void WriteBinTable(ref SubFlowSheet subFlowSheet, string parameter)
        {
            var flowRow = new FlowRow();
            flowRow.OpCode = FlowRow.OpCodeBinTable;
            flowRow.Parameter = parameter;
            flowRow.Enable = "BinTable";
            subFlowSheet.AddRow(flowRow);
        }
    }
}