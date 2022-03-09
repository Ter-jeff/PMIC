using System.Collections.Generic;
using System.Text.RegularExpressions;
using AutomationCommon.EpplusErrorReport;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.Common;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.Inputs.Setting.BinNumber;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenBinTable.GenBinTableRow
{
    public abstract class BinTableRowGeneratorBase
    {
        private const string BinErrorMessage = "SortBin number out of range in config";

        #region Constructor

        protected BinTableRowGeneratorBase(string sheetName, List<string> errorBinNumbers)
        {
            SheetName = sheetName;
            ErrorBinNumbers = errorBinNumbers;
        }

        #endregion

        #region Main Methods

        public BinTableRow GenBinTableRow(HardIpPattern pattern, string voltage = "")
        {
            SetPattern(pattern);
            return GenBinTableRowForPattern(voltage);
        }

        #endregion

        #region Fileds

        protected string SheetName;
        protected string BlockName = string.Empty;
        protected string SubBlockName = string.Empty;
        protected string SubBlock2Name = string.Empty;
        protected List<string> ErrorBinNumbers;
        protected string IpName = string.Empty;
        protected string TimingAc = string.Empty;
        protected string InstNameSubStr = string.Empty;
        protected bool NoPattern;
        protected BinNumberRuleRow BinLib;
        protected HardIpPattern Pattern;

        #endregion

        #region Abstract Methods

        protected abstract BinTableRow GenBinTableRowForPattern(string voltage = "");

        protected abstract void SetPattern(HardIpPattern pattern);

        #endregion

        #region Create columns creater methods

        protected string CreateSortBin()
        {
            return BinLib.CurrentSoftBin.ToString("G");
        }

        protected string CreateHardBin()
        {
            return BinLib.HardBin;
        }

        protected string CreateResult()
        {
            return BinLib.SoftBinState;
        }

        protected Dictionary<string, string> CreateSortBinExtraBinDic()
        {
            var extraBinDic = new Dictionary<string, string>();
            if (BinLib.HardIpHlvBin != "") extraBinDic.Add("HLV", BinLib.HardIpHlvBin);
            if (BinLib.HardIpHvBin != "") extraBinDic.Add("HV", BinLib.HardIpHvBin);
            if (BinLib.HardIpLvBin != "") extraBinDic.Add("LV", BinLib.HardIpLvBin);
            if (BinLib.HardIpNvBin != "") extraBinDic.Add("NV", BinLib.HardIpNvBin);

            return extraBinDic;
        }

        #endregion

        #region Other Methods

        protected void CheckErrorBinNum()
        {
            if (BinLib.CurrentSoftBin == BinLib.SoftBinEnd && BinLib.IsExceed)
                if (!ErrorBinNumbers.Contains(BinLib.Description))
                {
                    ErrorBinNumbers.Add(BinLib.Description);
                    EpplusErrorManager.AddError(HardIpErrorType.MissingBinNum.ToString(), ErrorLevel.Warning, "", 0,
                        BinErrorMessage, BinLib.Description);
                }
        }

        protected void SetBasicInfoByPattern(HardIpPattern pattern)
        {
            Pattern = pattern;
            BlockName = Regex.Replace(CommonGenerator.GetBlockName(pattern.MiscInfo, pattern.SheetName),
                "wireless_|lcd_", "", RegexOptions.IgnoreCase);
            SubBlockName =
                CommonGenerator.GetSubBlockNameWithoutMinus(pattern.Pattern.GetLastPayload(), pattern.MiscInfo,
                    BlockName);
            SubBlock2Name = CommonGenerator.GetSubBlock2Name(pattern.MiscInfo);
            IpName = CommonGenerator.GetIpName(pattern.MiscInfo);
            TimingAc = CommonGenerator.GetTimingAc(Pattern.AcUsed);
            InstNameSubStr = CommonGenerator.GetInstNameSubStr(Pattern.MiscInfo);
            NoPattern = CommonGenerator.NoPattern(pattern.Pattern.RealPatternName);
            BinLib = SearchInfo.GetHardIpBin(pattern);
        }

        public virtual BinTableRow GeneratePmicBinRow(HardIpPattern pattern, string voltage, List<string> flagList)
        {
            //Do nothing
            return new BinTableRow();
        }

        #endregion
    }
}