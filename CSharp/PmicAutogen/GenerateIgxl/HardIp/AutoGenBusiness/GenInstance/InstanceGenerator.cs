using System;
using System.Collections.Generic;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenInstance;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenInstance;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local.Const;

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
                instSheets.AddRange(blockInstanceGenerator.GenBlockInsRows());
            }

            return instSheets;
        }

        private BlockInstanceGenerator HardipInstanceFactory(Dictionary<string, List<HardIpPattern>> planDic,
            string sheetName)
        {
            BlockInstanceGenerator blockInstanceGenerator;
            //Leakage
            if (sheetName.Equals(PmicConst.PmicLeakage, StringComparison.OrdinalIgnoreCase))
                blockInstanceGenerator = new LeakageBlockInsGenerator(sheetName, planDic[sheetName]);
            else
                //DCTest
                blockInstanceGenerator = new DcTestBlockInsGenerator(sheetName, planDic[sheetName]);

            return blockInstanceGenerator;
        }
    }
}