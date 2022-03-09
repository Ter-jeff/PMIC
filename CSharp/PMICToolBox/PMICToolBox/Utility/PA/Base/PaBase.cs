using PmicAutomation.MyControls;
using PmicAutomation.Utility.PA.Input;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using PmicAutomation.MyControls;

namespace PmicAutomation.Utility.PA.Base
{
    public class PaBase
    {
        protected const string Pmic = "PMIC";
        protected const string Dcvi = "DCVI";
        protected const string ConDm = "_DM";
        protected const string ConDt = "_DT";
        protected const string ConDcTime = "DCTime";
        protected const string ConDcDiffMeter = "DCDiffMeter";

        protected readonly MyForm.RichTextBoxAppend Append;

        protected readonly string Device;

        protected readonly List<string> SuffixPmicTypeList = new List<string>
        {
            "UVI80",
            "DC30",
            "HEX",
            "UPAC_SRC",
            "UPAC_CAP"
        };

        protected readonly UflexConfig UflexConfig;
        protected int SiteCnt = 0;

        protected List<string> TypeList = new List<string>
        {
            "I/O",
            "DCVS",
            "DCVSMerged2",
            "DCVSMerged4",
            "DCVSMerged6",
            "DCVSMerged8",
            "DCVI",
            "Utility",
            "UltraSource",
            "UltraCapture",
            "GigaDigNeg",
            "GigaDigPos",
            "MW",
            "MWSource",
            "Gnd",
            "N/C"
        };

        public PaBase(string device, UflexConfig uflexConfig, MyForm.RichTextBoxAppend append)
        {
            Device = device;
            UflexConfig = uflexConfig;
            Append = append;
        }

        protected string GetChannel(string channelAssignment)
        {
            if (channelAssignment != null && Regex.IsMatch(channelAssignment,
                    @"\.ch|\.sense|\.util|.SrcPos|.SrcNeg|.cappos|.capneg", RegexOptions.IgnoreCase))
            {
                return channelAssignment.Split('.')[0];
            }

            return "";
        }

        protected string GetPinName(List<PaRow> paRows, PaRow paRow)
        {
            List<IGrouping<string, PaRow>> paItems = paRows
                .Where(x => x.BumpName == paRow.BumpName && x.Site == paRow.Site).GroupBy(y => y.PaType).ToList();
            string pinName;
            string toolType = paRow.InstrumentType;

            if (!string.IsNullOrEmpty(toolType) && IsPinNameContainsToolType(paRow.BumpName, toolType))
            {
                return paRow.BumpName;
            }

            if (paRow.IsPower() && Device.Equals(Pmic, StringComparison.OrdinalIgnoreCase))
            {
                pinName = paRow.BumpName + "_" + toolType;
            }
            else if (paItems.Count > 1 &&
                     paItems.First().Any(x => Regex.IsMatch(x.Ps, "power", RegexOptions.IgnoreCase)))
            {
                pinName = paRow.BumpName + "_" + toolType;
            }
            else if (paItems.Count > 1 && paRow.PaType.Equals("DCVI", StringComparison.OrdinalIgnoreCase))
            {
                pinName = paRow.BumpName;
            }
            else if (paItems.Count > 1)
            {
                pinName = paRow.BumpName + "_" + toolType;
            }
            else
            {
                pinName = paRow.BumpName;
            }

            return pinName;
        }

        protected bool IsPinNameContainsToolType(string pinName, string toolType)
        {
            return pinName.ToUpper().Contains("_" + toolType.ToUpper()) ||
                   pinName.ToUpper().Contains("_" + toolType.ToUpper() + "_");
        }
    }
}