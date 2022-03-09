using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenFlow;
using PmicAutogen.GenerateIgxl.HardIp.DividerManager.FlowDividerManager;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenFlow
{
    public class LeakageFlowSheetGenerator : DcTestFlowSheetGenerator
    {
        public LeakageFlowSheetGenerator(string sheetName, List<HardIpPattern> patternList) : base(sheetName,
            patternList)
        {
            FlowRowGenerator = new LeakageFlowRowGenerator(SheetName);
        }

        protected override List<HardIpPattern> DividePatterns()
        {
            return FlowLimitDivider.DivideUseLimit(PatternList);
        }

        protected override List<FlowRow> GenFlowBodyRows(bool shmooFlag = false)
        {
            var flowBodyRows = new List<FlowRow>();
            flowBodyRows.AddRange(GenFlowTestRowsByVoltage());
            //flowBodyRows.AddRange(GenResetRelayRows());
            flowBodyRows.AddRange(FlowRowGenerator.GenTtrFlagClearRow(flowBodyRows));
            return flowBodyRows;
        }

        private List<FlowRow> GenFlowTestRowsByVoltage(string labelVoltage = "")
        {
            var flowTestRows = new List<FlowRow>();
            FlowRowGenerator.LabelVoltage = labelVoltage;
            foreach (var pattern in ExtendedPatList)
            {
                FlowRowGenerator.Pat = pattern;
                var sweepCodeForRow = FlowRowGenerator.GenSweepCodeForRow();
                var sweepCodeNextRow = FlowRowGenerator.GenSweepCodeOrVoltageNextRow();
                //flowTestRows.AddRange(FlowRowGenerator.GenRelayRows());
                if (sweepCodeForRow != null)
                    flowTestRows.AddRange(sweepCodeForRow);
                flowTestRows.AddRange(FlowRowGenerator.GenTestRows());
                if (sweepCodeNextRow != null)
                    flowTestRows.AddRange(sweepCodeNextRow);
            }

            if (ExtendedPatList.Where(p => p.UseDeferLimit).ToList().Count > 0
            ) // if exist test defer limit=> generate limits all
                flowTestRows.Add(new FlowRow {OpCode = "limits-all"});
            return flowTestRows;
        }
    }
}