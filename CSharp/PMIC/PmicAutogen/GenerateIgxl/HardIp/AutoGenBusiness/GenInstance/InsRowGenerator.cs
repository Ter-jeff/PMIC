using System;
using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using IgxlData.VBT;
using PmicAutogen.GenerateIgxl.Basic.Writer.GenDc.PowerOverWrite;
using PmicAutogen.GenerateIgxl.HardIp.HardIpData.DataBase;
using PmicAutogen.GenerateIgxl.HardIp.HardIPUtility.SearchInfoUtility;
using PmicAutogen.GenerateIgxl.HardIp.InputObject;
using PmicAutogen.GenerateIgxl.HardIp.InstanceParameterSetting;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.HardIp.AutoGenBusiness.GenInstance
{
    public abstract class InsRowGenerator
    {
        protected string BlockName = string.Empty;
        protected string InstNameSubStr = string.Empty;
        protected string IpName = string.Empty;
        public string LabelVoltage = string.Empty;
        protected bool NoPattern = false;
        protected HardIpPattern Pattern;
        public PowerOverWrite PowerOverWrite = null;

        protected string SheetName;
        protected string SubBlock2Name = string.Empty;
        protected string SubBlockName = string.Empty;
        protected string TimingAc = string.Empty;
        protected VbtFunctionBase VbtFunction = null;

        #region Constructor

        protected InsRowGenerator(string sheetName)
        {
            SheetName = sheetName;
        }

        #endregion

        public HardIpPattern Pat
        {
            set
            {
                Pattern = value;
                SetBasicInfoByPattern(value);
            }
            get { return Pattern; }
        }

        public abstract List<InstanceRow> GenInsRows();

        protected abstract void SetBasicInfoByPattern(HardIpPattern pattern);

        protected string CreateType()
        {
            return HardIpConstData.InstanceTypeDefault;
        }

        protected string CreateVbtName()
        {
            if (string.IsNullOrEmpty(Pat.CustomVbName))
                return VbtFunction.FunctionName;
            return Pat.CustomVbName;
        }

        protected string CreateArgList()
        {
            return VbtFunction.Parameters;
        }

        protected List<string> CreateArgs()
        {
            var setArgValueMain = new SetArgValueMain();
            if (VbtFunction.Parameters != "") setArgValueMain.SetArgsValue(Pattern, VbtFunction, LabelVoltage);
            return VbtFunction.Args.Select(p => p).ToList();
        }

        protected string CreateAcSelector()
        {
            var patternName = Pattern.Pattern.GetLastPayload();

            if (InputFiles.PatternListMap == null)
                return "";

            if (!InputFiles.PatternListMap.PatternListCsvRows.Exists(x =>
                x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase))) return "";

            var acSelector = HardIpConstData.SelectTyp;
            if (!string.IsNullOrEmpty(Pattern.AcSelectorUsed))
            {
                var newAcSelector = SearchInfo.GetAcSelector(LabelVoltage, Pattern.AcSelectorUsed);
                if (!string.IsNullOrEmpty(newAcSelector))
                    acSelector = newAcSelector;
            }

            return acSelector;
        }

        protected string CreateDcSelector()
        {
            string dcSelector;
            switch (LabelVoltage)
            {
                case HardIpConstData.LabelHv:
                case HardIpConstData.LabelUHv:
                    dcSelector = HardIpConstData.SelectMax;
                    break;
                case HardIpConstData.LabelLv:
                case HardIpConstData.LabelULv:
                    dcSelector = HardIpConstData.SelectMin;
                    break;
                case HardIpConstData.LabelNv:
                    dcSelector = HardIpConstData.SelectTyp;
                    break;
                default:
                    dcSelector = HardIpConstData.SelectTyp;
                    break;
            }

            if (!string.IsNullOrEmpty(Pattern.DcSelectorUsed))
            {
                var newDcSelector = SearchInfo.GetDcSelector(LabelVoltage, Pattern.DcSelectorUsed);
                if (!string.IsNullOrEmpty(newDcSelector))
                    dcSelector = newDcSelector;
            }

            return dcSelector;
        }
    }
}