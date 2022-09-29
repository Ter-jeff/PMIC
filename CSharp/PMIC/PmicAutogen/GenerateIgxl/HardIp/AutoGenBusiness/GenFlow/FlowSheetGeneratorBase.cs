using System.Collections.Generic;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenFlow
{
    public abstract class FlowSheetGeneratorBase
    {
        private const string PrefixFlowSheet = "Flow_";

        //Pattern List after divided
        protected List<HardIpPattern> ExtendedPatList;
        protected FlowRowGeneratorBase FlowRowGenerator = null;

        protected string FlowSheetName = string.Empty;

        //Pattern List from test plan 
        protected List<HardIpPattern> PatternList;
        protected string SheetName;

        protected FlowSheetGeneratorBase(string sheetName, List<HardIpPattern> patternList)
        {
            SheetName = CommonGenerator.GetHardipSheetName(sheetName).ToUpper();
            PatternList = patternList;
        }

        public List<SubFlowSheet> GenerateFlowSheet()
        {
            var flowSheets = new List<SubFlowSheet>();
            FlowSheetName = PrefixFlowSheet + SheetName;
            var flowSheet = new SubFlowSheet(FlowSheetName);
            ExtendedPatList = DividePatterns();
            flowSheet.AddStartRows(SubFlowSheet.Ttime);
            flowSheet.AddRows(GenFlowBodyRows());
            flowSheet.AddEndRows(SubFlowSheet.Ttime, false);
            flowSheets.Add(flowSheet);
            return flowSheets;
        }

        protected abstract List<HardIpPattern> DividePatterns();

        protected abstract FlowRows GenFlowBodyRows(bool shmooFlag = false);

        //protected virtual FlowRows GenEndRows()
        //{
        //    var flowRows = new FlowRows();
        //    if (Regex.IsMatch(FlowSheetName, HardIpConstData.PrefixHardIp + "|" + HardIpConstData.GpioBlockName, RegexOptions.IgnoreCase) && 
        //        !Regex.IsMatch(FlowSheetName, "init|nWire", RegexOptions.IgnoreCase))
        //        flowRows.Add(FlowRowGenerator.GenHardIpDatalogFlow("Disable"));
        //    flowRows.Add(FlowRowGenerator.GenPrintStopRow());
        //    flowRows.Add(FlowRowGeneratorBase.GenReturnRow());
        //    return flowRows;
        //}

        //protected FlowRows GenResetRelayRows(string labelVoltage = "")
        //{
        //    var flowRows = new FlowRows();
        //    var lastSetting = new Dictionary<string, string>();
        //    var lastPlanItem = ExtendedPatList.LastOrDefault(a => a.IsInTestPlan);
        //    if (lastPlanItem != null)
        //        lastSetting = lastPlanItem.NewRelaySetting;

        //    if (lastSetting.Count > 0)
        //    {
        //        string lastEnable = !string.IsNullOrEmpty(labelVoltage) &&
        //                        !Regex.IsMatch(ExtendedPatList[ExtendedPatList.Count - 1].MiscInfo, HardIpConstData.RemoveNv,
        //                            RegexOptions.IgnoreCase)
        //        ? HardIpConstData.PrefixHardIp + labelVoltage
        //        : "";
        //        foreach (var item in lastSetting)
        //        {
        //            string setting = item.Value;
        //            flowRows.AddRange(RelaySettingMain.GenRelaySettingInJob(null, SearchInfo.ReverseRelaySetting(setting), item.Key, lastEnable));
        //        }
        //    }
        //    return flowRows;
        //}
    }
}