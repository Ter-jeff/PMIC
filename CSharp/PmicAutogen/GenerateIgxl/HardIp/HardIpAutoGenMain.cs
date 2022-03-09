using System.Collections.Generic;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenFlow;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenInstance;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp
{
    public class HardIpAutoGenMain
    {
        public List<SubFlowSheet> GenFlow(Dictionary<string, List<HardIpPattern>> planDic)
        {
            var flowGenerator = new FlowGenerator();
            return flowGenerator.GenFlowSheets(planDic);
        }

        public List<InstanceSheet> GenInst(Dictionary<string, List<HardIpPattern>> planDic)
        {
            var instanceGenerator = new InstanceGenerator();
            return instanceGenerator.GenInstanceSheets(planDic);
        }

        public BinTableSheet GenBinTable(Dictionary<string, List<HardIpPattern>> planDic)
        {
            var binTableSheetGenerator = new BinTableSheetGenerator();
            return binTableSheetGenerator.GenBinTableSheet(planDic);
        }

        //public CharSheet GenDcTestCharSheet(Dictionary<string, List<HardIpPattern>> planDic)
        //{
        //    var charSheetGenerator = new CharSheetGenerator();
        //    return charSheetGenerator.GenCharSheet(planDic);
        //}

        //public void GenDcTestAcCategory(Dictionary<string, List<HardIpPattern>> planDic, List<InstanceSheet> instSheets)
        //{
        //    var acCategoryGenerator = new AcCategoryGenerator();
        //    acCategoryGenerator.AddNwireAcCategory(planDic, instSheets);
        //}

        //public List<TimeSetBasicSheet> GenDcTestTimeSet(Dictionary<string, List<HardIpPattern>> planDic)
        //{
        //    var timeSetGenerator = new TimeSetSheetGenerator();
        //    return timeSetGenerator.GenerateTimeSetSheets(planDic);
        //}
    }
}