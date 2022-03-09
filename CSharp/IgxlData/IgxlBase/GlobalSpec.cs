namespace IgxlData.IgxlBase
{
    public class GlobalSpec : Spec
    {
        #region Field
        #endregion

        #region Property

        public string Job { get; set; }

        #endregion

        #region Constructor

        public GlobalSpec(string glbSym) : base(glbSym)
        {
            Job = "";
        }

        public GlobalSpec(string glbSym, string glbValue = "", string glbJob = "", string glbComm = "")
            : base(glbSym, glbValue, glbComm)
        {
            Job = glbJob;
        }

        #endregion
    }
}