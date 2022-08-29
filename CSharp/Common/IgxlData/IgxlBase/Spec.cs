using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{Symbol}")]
    public abstract class Spec : IgxlItem
    {
        #region Property

        public string Symbol { get; set; }
        public string Value { get; set; }
        public string Comment { get; set; }

        #endregion

        #region Constructor

        protected Spec()
        {
        }

        protected Spec(string specSym)
        {
            Symbol = specSym;
            Value = "";
            Comment = "";
        }

        protected Spec(string specSym, string specVal = "", string specComm = "")
        {
            Symbol = specSym;
            Value = specVal;
            Comment = specComm;
        }

        #endregion
    }
}