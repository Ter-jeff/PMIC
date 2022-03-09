using System.Collections.Generic;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenInstance
{
    public abstract class BlockInstanceGenerator
    {
        protected List<HardIpPattern> HardIpPatterns;
        protected InsRowGenerator InstanceRowGenerator = null;
        protected string SheetName;

        protected BlockInstanceGenerator(string sheetName, List<HardIpPattern> pattenList)
        {
            SheetName = sheetName;
            HardIpPatterns = pattenList;
        }

        public abstract List<InstanceSheet> GenBlockInsRows();
    }
}