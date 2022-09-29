using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenBinTable;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTable;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable
{
    public class BinTableSheetGenerator
    {
        public BinTableSheet GenBinTableSheet(Dictionary<string, List<HardIpPattern>> planDic)
        {
            var binTables = new BinTableSheet(IgxlWorkBook.MainBinTableName);
            var errorBinNumbers = new List<string>();
            var duplicateParameter = new List<string>();
            foreach (var sheet in planDic)
            {
                //BlockBinTableGeneratorBase blockBinGenerator = new DcTestBinTableGenerator(binTables, sheet.Key,
                //    planDic[sheet.Key], duplicateParameter, errorBinNumbers);
                //if (sheet.Key.EndsWith("_" + PmicConst.Leakage, StringComparison.CurrentCultureIgnoreCase))
                //    blockBinGenerator = new LeakageBinTableWriter(binTables, sheet.Key, planDic[sheet.Key],
                //        duplicateParameter, errorBinNumbers);
                //blockBinGenerator.GenerateBinTableRows();
            }

            return binTables;
        }
    }
}