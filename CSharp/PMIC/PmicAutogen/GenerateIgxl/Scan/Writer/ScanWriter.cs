using CommonLib.Extension;
using PmicAutogen.Inputs.Setting.BinNumber;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanWriter
    {
        protected const string Hnlv = "HNLV";
        protected const string Nlv = "NLV";
        protected const string Hlv = "HLV";
        protected const string Hnv = "HNV";
        protected const string Lv = "LV";
        protected const string Nv = "NV";
        protected const string Hv = "HV";
        protected const string ULv = "ULV";
        protected const string UHv = "UHV";
        protected string Block;

        public ScanWriter()
        {
            Block = "SCAN";
            BlockBinTableName = "SCAN";
            EnumBinNumberBlock = EnumBinNumberBlock.Scan;
        }

        protected string BlockBinTableName { get; set; }
        protected EnumBinNumberBlock EnumBinNumberBlock { get; set; }

        protected string GetBinTableName(string payload, string voltage)
        {
            return string.Format("Bin_{0}_{1}_{2}", Block, payload.GetSortPatNameForBinTable(), voltage);
        }

        protected string GetFlagName(string payload, string x)
        {
            return string.Format("F_{0}_{1}", payload.GetSortPatNameForBinTable(), x);
        }
    }
}