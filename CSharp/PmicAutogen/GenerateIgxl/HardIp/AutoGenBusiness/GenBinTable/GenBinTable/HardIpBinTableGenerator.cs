using System.Collections.Generic;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTableRow;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTable
{
    public class HardIpBinTableGenerator : BlockBinTableGeneratorBase
    {
        public HardIpBinTableGenerator(BinTableSheet hardIpBinTableSheet, string sheetName,
            List<HardIpPattern> patternList, List<string> duplicateParameter, List<string> errorBinNumbers) : base(
            hardIpBinTableSheet, patternList, duplicateParameter)
        {
            BinTableRowGenerator = new HardIpBinTableRowGenerator(sheetName, errorBinNumbers);
        }
    }
}