using System.Collections.Generic;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTable;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenBinTable
{
    public class LeakageBinTableGenerator : HardIpBinTableGenerator
    {
        public LeakageBinTableGenerator(BinTableSheet hardIpBinTableSheet, string sheetName,
            List<HardIpPattern> patternList, List<string> duplicateParameter, List<string> errorBinNumbers) : base(
            hardIpBinTableSheet, sheetName, patternList, duplicateParameter, errorBinNumbers)
        {
            BinTableRowGenerator = new LeakageBinTableRowGenerator(sheetName, errorBinNumbers);
        }
    }
}