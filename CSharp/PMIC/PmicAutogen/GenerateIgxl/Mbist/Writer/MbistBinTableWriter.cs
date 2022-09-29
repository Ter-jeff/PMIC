using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.Inputs.Setting.BinNumber;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class MbistBinTableWriter : ScanBinTableWriter
    {
        public MbistBinTableWriter()
        {
            Block = "MBIST";
            EnumBinNumberBlock = EnumBinNumberBlock.Mbist;
            BlockBinTableName = "MBIST";
        }
    }
}