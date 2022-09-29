using CommonLib.Utility;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class MbistFlowTableWriter : ScanFlowTableWriter
    {
        public MbistFlowTableWriter()
        {
            SheetName = "Flow_Mbist";
            Block = "MBIST";
            BlockBinTableName = "MBIST";
        }

        protected override void WriteBody(List<ProdCharRow> prodCharRows, ref SubFlowSheet subFlowSheet)
        {
            var instances = prodCharRows.FindAll(p =>
                p.Nop == false || p.NopType == NopType.BlankInit || p.NopType == NopType.WrongTimeSet ||
                p.NopType == NopType.NonUsage).ToList();
            if (instances.Count < 1) return;

            foreach (var prodCharRow in instances)
            {
                var instance = (ProdCharRowMbist)prodCharRow;
                WriteFlowRow(ref subFlowSheet, instance, "HV");

                WriteFlowRow(ref subFlowSheet, instance, "NV");

                WriteFlowRow(ref subFlowSheet, instance, "LV");

                if (LocalSpecs.HasUltraVoltageUHv) WriteFlowRow(ref subFlowSheet, instance, "UHV");

                if (LocalSpecs.HasUltraVoltageULv) WriteFlowRow(ref subFlowSheet, instance, "ULV");

                if (!(instance.PeripheralVoltage.Equals("") || instance.PeripheralVoltage.Equals("N/A")))
                {
                    var singlePins = instance.GetSinglePins(instance.PeripheralVoltage);

                    foreach (var pin in singlePins)
                    {
                        var setupName = instance.Get1DCharNamePeriod(pin);
                        Write1DChar(ref subFlowSheet, instance, "NV", setupName);
                    }

                    var trackingGroups = instance.GetTrackingGroup(instance.PeripheralVoltage);
                    foreach (var pins in trackingGroups)
                        foreach (var pin in pins)
                        {
                            var setupName = instance.Get2DCharNamePeriod(pin, "MBIST");
                            Write2DChar(ref subFlowSheet, instance, "NV", setupName);
                        }
                }

                subFlowSheet.FlowRows.Add_A_Enable_MP_SBIN("MBIST");
                WriteBinTables(ref subFlowSheet, instances.First());
            }
        }

        protected override void Write1DChar(ref SubFlowSheet subFlowSheet, ProdCharRow prodCharRow, string pBinType,
            string setupName)
        {
            var row = (ProdCharRowMbist)prodCharRow;
            if (string.IsNullOrEmpty(row.PeripheralVoltage))
                return;
            var itemInstanceName = row.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.Enable = "B_Debug_1D_SHMOO";
            rowTemp.OpCode = FlowRow.OpCodeCharacterize;
            rowTemp.Parameter = Combine.CombineByUnderLine(itemInstanceName, pBinType) + " " + setupName;
            subFlowSheet.AddRow(rowTemp);
        }

        protected override void Write2DChar(ref SubFlowSheet subFlowSheet, ProdCharRow prodCharRow, string pBinType,
            string setupName)
        {
            var row = (ProdCharRowMbist)prodCharRow;
            if (string.IsNullOrEmpty(row.PeripheralVoltage))
                return;
            var itemInstanceName = row.InstanceName;
            var rowTemp = new FlowRow();
            rowTemp.Enable = "B_Debug_2D_SHMOO";
            rowTemp.OpCode = FlowRow.OpCodeCharacterize;
            rowTemp.Parameter = Combine.CombineByUnderLine(itemInstanceName, pBinType) + " " + setupName;
            subFlowSheet.AddRow(rowTemp);
        }
    }
}