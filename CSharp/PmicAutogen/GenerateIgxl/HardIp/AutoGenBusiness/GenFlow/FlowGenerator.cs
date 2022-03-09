using System;
using System.Collections.Generic;
using AutomationCommon.DataStructure;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenFlow;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenFlow;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.InputPackages;
using PmicAutogen.Local.Const;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenFlow
{
    public class FlowGenerator
    {
        public List<SubFlowSheet> GenFlowSheets(Dictionary<string, List<HardIpPattern>> planDic)
        {
            var flowSheets = new List<SubFlowSheet>();
            foreach (var sheetName in planDic.Keys)
                try
                {
                    var flowSheetGenerator = new DcTestFlowSheetGenerator(sheetName, planDic[sheetName]);

                    if (sheetName.Equals(PmicConst.PmicLeakage, StringComparison.OrdinalIgnoreCase))
                        flowSheetGenerator = new LeakageFlowSheetGenerator(sheetName, planDic[sheetName]);

                    flowSheets.AddRange(flowSheetGenerator.GenerateFlowSheet());
                }
                catch (Exception ex)
                {
                    Response.Report("Generating Flow " + sheetName + " failed " + ex.Message, MessageLevel.Error, 0);
                }

            return flowSheets;
        }
    }
}