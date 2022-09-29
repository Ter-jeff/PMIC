using AutomationCommon.DataStructure;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenFlow;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.InputPackages;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;

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
                    var flowSheetGenerator = new DcTestFlowWriter(sheetName, planDic[sheetName]);

                    //if (sheetName.EndsWith("_" + PmicConst.Leakage, StringComparison.CurrentCultureIgnoreCase))
                    //    flowSheetGenerator = new LeakageFlowSheetGenerator(sheetName, planDic[sheetName]);

                    //flowSheets.AddRange(flowSheetGenerator.GenerateFlowRows());
                }
                catch (Exception ex)
                {
                    Response.Report("Generating Flow " + sheetName + " failed " + ex.Message, MessageLevel.Error, 0);
                }

            return flowSheets;
        }
    }
}