using System.Collections.Generic;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class MbistBinTableWriter : ScanBinTableWriter
    {
        public MbistBinTableWriter()
        {
            Block = "MBIST";
        }

        public List<BinTableRow> WriteBinTable(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var binTableRows = new List<BinTableRow>();
            foreach (var prodCharRowMbist in prodCharRowMbists)
                GenerateBinTable(binTableRows, prodCharRowMbist.PayLoadName);
            return binTableRows;
        }
    }
}