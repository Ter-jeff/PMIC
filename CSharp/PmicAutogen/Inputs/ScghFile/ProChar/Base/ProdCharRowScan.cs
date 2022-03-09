using System;
using PmicAutogen.Inputs.ScghFile.Reader;

namespace PmicAutogen.Inputs.ScghFile.ProChar.Base
{
    public class ProdCharRowScan : ProdCharRow
    {
        public ProdCharRowScan(IProdCharSheetRow prodCharRow) : base(prodCharRow)
        {
        }

        public string PerformanceMode { set; get; }

        public string SupplyVoltage
        {
            get
            {
                var prodCharRow = (ProdCharSheetRow) ProdCharItem;
                return prodCharRow.SupplyVoltage;
            }
            set { throw new NotImplementedException(); }
        }
    }
}