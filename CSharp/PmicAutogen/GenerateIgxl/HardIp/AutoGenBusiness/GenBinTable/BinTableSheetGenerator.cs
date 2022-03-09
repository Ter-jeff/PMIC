using System;
using System.Collections.Generic;
using IgxlData.IgxlSheets;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenBinTable;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenBinTable;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTable;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local.Const;

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
                BlockBinTableGeneratorBase blockBinGenerator = new DcTestBinTableGenerator(binTables, sheet.Key,
                    planDic[sheet.Key], duplicateParameter, errorBinNumbers);
                if (sheet.Key.Equals(PmicConst.PmicLeakage, StringComparison.OrdinalIgnoreCase))
                    blockBinGenerator = new LeakageBinTableGenerator(binTables, sheet.Key, planDic[sheet.Key],
                        duplicateParameter, errorBinNumbers);
                blockBinGenerator.GenerateBinTableRows();
            }

            return binTables;
        }
    }
}