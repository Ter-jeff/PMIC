using PmicAutogen.Inputs.ScghFile.Reader;
using System;

namespace PmicAutogen.Inputs.ScghFile.ProChar.Base
{
    public class ProdCharRowMbist : ProdCharRow
    {
        public ProdCharRowMbist(IProdCharSheetRow prodCharRow) : base(prodCharRow)
        {
        }

        public string PerformanceMode { set; get; }

        public string PeripheralVoltage
        {
            get
            {
                var prodCharRow = (ProdCharSheetRow)ProdCharItem;
                return prodCharRow.PeripheralVoltage;
            }
            set { throw new NotImplementedException(); }
        }
    }
}