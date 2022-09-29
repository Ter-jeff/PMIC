using PmicAutogen.Inputs.Setting.BinNumber.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.Inputs.Setting.BinNumber
{
    public class BinNumberSingleton
    {
        #region Initialization

        private void InitialSoftBinRange()
        {
            var rangeWorksheet = InputFiles.SettingWorkbook.Worksheets[PmicConst.BinNumberRule];
            if (rangeWorksheet == null) return;
            var reader = new BinNumberRangeReader();
            _softBinRanges = reader.ReadSheet(rangeWorksheet);
        }

        #endregion

        public static void Initialize()
        {
            _instance = null;
        }

        #region Singleton

        private static BinNumberSingleton _instance;
        private static List<SoftBinRangeRow> _softBinRanges;

        private BinNumberSingleton()
        {
            if (InputFiles.SettingWorkbook.Worksheets.Count == 0)
                throw new Exception("Can not find Bin number config file, please check if it has existed!");

            InitialSoftBinRange();
        }

        public static BinNumberSingleton Instance()
        {
            return _instance ?? (_instance = new BinNumberSingleton());
        }

        #endregion

        #region Get Bin number

        public bool GetBinNumDefRow(BinNumberRuleCondition binNumDefPara, out BinNumberRuleRow defOut)
        {
            defOut = new BinNumberRuleRow();
            var targetItem = SearchSoftBinRangeData(binNumDefPara);
            if (targetItem != null)
            {
                defOut.Description = targetItem.Description;
                defOut.CurrentSoftBin = targetItem.GetSoftBinNumber();
                defOut.HardBin = targetItem.HardBin;
                defOut.SoftBinStart = targetItem.GetSoftBinStart();
                defOut.SoftBinEnd = targetItem.GetSoftBinEnd();
                defOut.SoftBinState = targetItem.GetStatus();
                defOut.IsExceed = targetItem.CheckExceed();
                defOut.CurrentBinLib = targetItem;
                defOut.HardIpHlvBin = targetItem.HardHlvBin;
                defOut.HardIpHvBin = targetItem.HardHvBin;
                defOut.HardIpLvBin = targetItem.HardLvBin;
                defOut.HardIpNvBin = targetItem.HardNvBin;
                return true;
            }

            return false;
        }

        public SoftBinRangeRow SearchSoftBinRangeData(BinNumberRuleCondition binNumDefPara)
        {
            var binNumForBlock = _softBinRanges
                .Where(p => p.Block.Equals(binNumDefPara.Block, StringComparison.OrdinalIgnoreCase)).Select(a => a)
                .ToList();

            SoftBinRangeRow softBinRangeRow = null;
            if (binNumForBlock.Exists(p => p.Match(binNumDefPara.Condition)))
                softBinRangeRow = binNumForBlock.Find(p => p.Match(binNumDefPara.Condition));

            if (softBinRangeRow == null)
            {
                binNumForBlock = _softBinRanges
                    .Where(p => p.Block.Equals(EnumBinNumberBlock.Default.ToString(),
                        StringComparison.OrdinalIgnoreCase)).Select(a => a).ToList();
                softBinRangeRow = binNumForBlock.Find(p => p.Match("Default"));
            }

            return softBinRangeRow;
        }

        #endregion
    }
}