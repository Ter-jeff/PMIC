using System;
using System.Collections.Generic;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTableRow
{
    public class HardIpBinTableRowGenerator : BinTableRowGeneratorBase
    {
        private const string OpCode = "OR";

        public HardIpBinTableRowGenerator(string sheetName, List<string> errorBinNumbers) : base(sheetName,
            errorBinNumbers)
        {
        }

        protected override BinTableRow GenBinTableRowForPattern(string voltage = "")
        {
            var binTableRow = new BinTableRow();
            binTableRow.Name = CreateHardIpName();
            binTableRow.ItemList = CreateHardIpItemList();
            binTableRow.Op = CreateHardIpOpCode();
            binTableRow.Items = CreateHardIpItems(voltage);
            binTableRow.Sort = CreateSortBin();
            binTableRow.Bin = CreateHardBin();
            binTableRow.Result = CreateResult();
            binTableRow.ExtraBinDictionary = CreateSortBinExtraBinDic();
            CheckErrorBinNum();
            return binTableRow;
        }

        protected override void SetPattern(HardIpPattern pattern)
        {
            SetBasicInfoByPattern(pattern);
        }

        protected virtual string CreateHardIpItemList(string voltage = "")
        {
            var patternName = Pattern.Pattern.GetPatternName();

            var failFlagN = CommonGenerator.GenHardIpFlowFailFlag(SheetName, BlockName, SubBlockName, SubBlock2Name,
                patternName, TimingAc, InstNameSubStr, HardIpConstData.LabelNv, NoPattern);
            var failFlagH = CommonGenerator.GenHardIpFlowFailFlag(SheetName, BlockName, SubBlockName, SubBlock2Name,
                patternName, TimingAc, InstNameSubStr, HardIpConstData.LabelHv, NoPattern);
            var failFlagL = CommonGenerator.GenHardIpFlowFailFlag(SheetName, BlockName, SubBlockName, SubBlock2Name,
                patternName, TimingAc, InstNameSubStr, HardIpConstData.LabelLv, NoPattern);
            return failFlagN + "," + failFlagH + "," + failFlagL;
        }

        protected string CreateHardIpName()
        {
            var patternName = Pattern.Pattern.GetPatternName();
            if (LocalSpecs.CurrentProject.Equals("sicily", StringComparison.OrdinalIgnoreCase) ||
                LocalSpecs.CurrentProject.Equals("tonga", StringComparison.OrdinalIgnoreCase))
                return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName);
            return CommonGenerator.GenHardIpFlowBinParameter(SheetName, BlockName, SubBlockName, patternName, TimingAc,
                InstNameSubStr, NoPattern);
        }

        private string CreateHardIpOpCode()
        {
            return OpCode;
        }

        private List<string> CreateHardIpItems(string voltage = "")
        {
            switch (voltage)
            {
                case HardIpConstData.LabelHv:
                    return new List<string> {"", "T", ""};
                case HardIpConstData.LabelLv:
                    return new List<string> {"", "", "T"};
                case HardIpConstData.LabelNv:
                    return new List<string> {"T", "", ""};
                default:
                    return new List<string> {"T", "T", "T"};
            }
        }
    }
}