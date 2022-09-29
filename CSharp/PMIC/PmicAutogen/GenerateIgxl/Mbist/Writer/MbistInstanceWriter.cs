using PmicAutogen.GenerateIgxl.Scan.Writer;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class MbistInstanceWriter : ScanInstanceWriter
    {
        public MbistInstanceWriter()
        {
            SheetName = "TestInst_Mbist";
            Block = "Mbist";
            BlockBinTableName = "MBIST";
        }
    }
}