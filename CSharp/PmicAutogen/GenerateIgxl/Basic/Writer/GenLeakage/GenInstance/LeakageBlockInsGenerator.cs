using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlSheets;
using IgxlData.NonIgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenInstance;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenInstance
{
    public class LeakageBlockInsGenerator : BlockInstanceGenerator
    {
        public LeakageBlockInsGenerator(string sheetName, List<HardIpPattern> pattenList) : base(sheetName, pattenList)
        {
            InstanceRowGenerator = new LeakageInsRowGenerator(sheetName);
        }

        public override List<InstanceSheet> GenBlockInsRows()
        {
            var instanceSheetList = new List<InstanceSheet>();
            var leakageInsSheet = new InstanceSheet(PmicConst.TestInstDcLeakage);
            foreach (var hardIpPattern in HardIpPatterns)
            {
                hardIpPattern.FunctionName = GetLeakageFunctionName(hardIpPattern);

                InstanceRowGenerator.LabelVoltage = "";
                InstanceRowGenerator.Pat = hardIpPattern;
                var insRowList = InstanceRowGenerator.GenInsRows();
                foreach (var insRow in insRowList)
                    leakageInsSheet.AddRow(insRow);
            }

            if (leakageInsSheet.InstanceRows.Count != 0) instanceSheetList.Add(leakageInsSheet);
            return instanceSheetList;
        }

        private string GetLeakageFunctionName(HardIpPattern hardIpPattern)
        {
            var functionName = SearchInfo.GetVbtNameByPattern(hardIpPattern);
            if (string.IsNullOrEmpty(functionName))
            {
                functionName = VbtFunctionLib.PmicLeakageVbtName;
                var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
                if (pinMap != null)
                {
                    var measPins = hardIpPattern.MeasPins.Select(x => x.PinName).ToList();
                    if (measPins.All(x => pinMap.IsChannelType(x, "DCVI")))
                        functionName = VbtFunctionLib.PmicLeakageDcviVbtName;

                    if (measPins.All(x => pinMap.IsChannelType(x, "DCVS")))
                        functionName = VbtFunctionLib.PmicLeakageDcvsVbtName;
                }
            }

            return functionName;
        }
    }
}