using System.Collections.Generic;
using System.IO;

namespace IgxlData.NonIgxlSheets
{
    public class NonIgxlSheets
    {
        #region Constructor

        public NonIgxlSheets()
        {
            SheetList = new List<string>();
        }

        #endregion

        #region Property

        public List<string> SheetList { get; set; }

        #endregion

        #region Field

        #endregion

        #region Member Function

        public void Add(string dir, string fileName)
        {
            var fileFullName = Path.Combine(dir, fileName);
            SheetList.Add(fileFullName);
        }

        public void Clear()
        {
            SheetList.Clear();
        }

        #endregion
    }
}