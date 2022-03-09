using System.Collections.Generic;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTableRow;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenLeakage.GenBinTable
{
    public class LeakageBinTableRowGenerator : HardIpBinTableRowGenerator
    {
        public LeakageBinTableRowGenerator(string sheetName, List<string> errorBinNumbers) : base(sheetName,
            errorBinNumbers)
        {
        }

        protected override BinTableRow GenBinTableRowForPattern(string voltage = "")
        {
            var binTableRow = new BinTableRow();
            binTableRow.Name = CreateHardIpName();
            binTableRow.ItemList = CreateHardIpItemList();
            binTableRow.Op = "AND";
            binTableRow.Sort = CreateSortBin();
            binTableRow.Bin = CreateHardBin();
            binTableRow.Result = CreateResult();
            binTableRow.Items = CreateLeakageItems();
            binTableRow.ExtraBinDictionary = CreateSortBinExtraBinDic();
            CheckErrorBinNum();
            return binTableRow;
        }

        private List<string> CreateLeakageItems()
        {
            return new List<string> {"T"};
        }

        protected override string CreateHardIpItemList(string voltage = "")
        {
            var patternName = Pattern.Pattern.GetPatternName();
            return CommonGenerator.GenHardIpFlowTestFailAction(SheetName, BlockName, SubBlockName, SubBlock2Name,
                patternName, TimingAc, InstNameSubStr, "", Pattern.MiscInfo, NoPattern);
        }
    }
}