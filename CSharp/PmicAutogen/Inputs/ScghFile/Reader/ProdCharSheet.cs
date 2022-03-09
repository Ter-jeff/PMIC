using System.Collections.Generic;

namespace PmicAutogen.Inputs.ScghFile.Reader
{
    public class ProdCharSheet
    {
        #region Constructor

        public ProdCharSheet()
        {
            RowList = new List<ProdCharSheetRow>();
        }

        #endregion

        #region Property

        public string SheetName { get; set; }
        public List<ProdCharSheetRow> RowList { get; set; }

        #endregion
    }
}