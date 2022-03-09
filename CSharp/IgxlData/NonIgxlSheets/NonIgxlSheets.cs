using System.Collections.Generic;
using System.IO;

namespace IgxlData.NonIgxlSheets
{
    public class NonIgxlSheets
    {
        #region Field

        #endregion

        #region Property
        public List<string> SheetList { get; set; }

        #endregion

        #region Constructor
        public NonIgxlSheets()
        {
            SheetList = new List<string>();
        }
        #endregion

        #region Member Function
        public void Add(string dir, string fileName)
        {
            string fileFullName = Path.Combine(dir, fileName);
            SheetList.Add(fileFullName);
        }

        public void Clear()
        {
            SheetList.Clear();
        }
        #endregion
    }
}