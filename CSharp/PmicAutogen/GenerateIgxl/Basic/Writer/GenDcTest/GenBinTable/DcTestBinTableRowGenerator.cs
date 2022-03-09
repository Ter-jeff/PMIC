using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTableRow;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenDcTest.GenBinTable
{
    public class DcTestBinTableRowGenerator : HardIpBinTableRowGenerator
    {
        public DcTestBinTableRowGenerator(string sheetName, List<string> errorBinNumbers) : base(sheetName,
            errorBinNumbers)
        {
        }

        public override BinTableRow GeneratePmicBinRow(HardIpPattern pattern, string voltage, List<string> flagList)
        {
            SetPattern(pattern);
            var binRow = new BinTableRow();
            binRow.Op = "AND";
            binRow.Name = (CreateHardIpName() + "_" + voltage).Trim('_');
            var itemList = new List<string>();
            foreach (var item in flagList) itemList.Add(CreateHardIpItemList(item));
            binRow.ItemList = string.Join(",", itemList);
            binRow.Items = Enumerable.Repeat("T", flagList.Count).ToList();
            binRow.Result = "Fail";
            binRow.Sort = CreateSortBin();
            binRow.Bin = CreateHardBin();
            return binRow;
        }

        protected override string CreateHardIpItemList(string voltage = "")
        {
            var patternName = Pattern.Pattern.GetPatternName();
            return CommonGenerator.GenHardIpFlowTestFailAction(SheetName, BlockName, SubBlockName, SubBlock2Name,
                patternName, TimingAc, InstNameSubStr, voltage, Pattern.MiscInfo, NoPattern);
        }
    }
}