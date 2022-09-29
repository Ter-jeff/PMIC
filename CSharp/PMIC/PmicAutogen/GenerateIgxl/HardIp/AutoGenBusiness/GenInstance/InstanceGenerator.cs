using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenInstance;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenInstance
{
    public class InstanceGenerator
    {
        public List<InstanceSheet> GenInstanceSheets(Dictionary<string, List<HardIpPattern>> planDic)
        {
            var instSheets = new List<InstanceSheet>();
            foreach (var sheetName in planDic.Keys)
            {
                var blockInstanceGenerator = HardipInstanceFactory(planDic, sheetName);
                instSheets.AddRange(blockInstanceGenerator.GenInsRows());
            }

            return instSheets;
        }

        private BlockInstanceGenerator HardipInstanceFactory(Dictionary<string, List<HardIpPattern>> planDic,
            string sheetName)
        {
            BlockInstanceGenerator blockInstanceGenerator = null;
            //Leakage
            //if (sheetName.EndsWith("_" + PmicConst.Leakage, StringComparison.CurrentCultureIgnoreCase))
            //    blockInstanceGenerator = new LeakageBlockInsGenerator(sheetName, planDic[sheetName]);
            //else
            //DCTest
            //blockInstanceGenerator = new DcTestInstanceWriter(sheetName, planDic[sheetName]);

            return blockInstanceGenerator;
        }
    }
}