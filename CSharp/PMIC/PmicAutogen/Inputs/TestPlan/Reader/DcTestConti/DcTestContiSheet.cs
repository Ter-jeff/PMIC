using System.Collections.Generic;

namespace PmicAutogen.Inputs.TestPlan.Reader.DcTestConti
{
    public class DcTestContinuitySheet
    {
        #region Constructor

        public DcTestContinuitySheet(string sheetName)
        {
            Rows = new List<DcTestContiRow>();
            DicCategory = new Dictionary<string, List<DcTestContiRow>>();
            SheetName = sheetName;
        }

        #endregion

        #region Member Function

        public void AddRow(DcTestContiRow dcTestContiRow)
        {
            Rows.Add(dcTestContiRow);
        }

        #endregion

        #region Property

        public Dictionary<string, List<DcTestContiRow>> DicCategory;
        public List<DcTestContiRow> Rows { get; set; }
        public string SheetName { get; set; }

        public int CategoryIndex;
        public int PinGroupIndex;
        public int TimeSetIndex;
        public int ConditionIndex;
        public int LimitIndex;

        #endregion
    }
}