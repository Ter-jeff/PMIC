using System;
using System.Collections.Generic;
using System.IO;

namespace IgxlData.Others.MultiTimeSet
{
    public class MultiTimeSetSheets
    {
        #region Field

        private List<ComTimeSetBasicSheet> _timeSetBasicSheetsList;

        #endregion

        #region Propety
        public List<ComTimeSetBasicSheet> TimeSetBasicSheetsList
        {
            get { return _timeSetBasicSheetsList; }
        }
        #endregion

        #region Constructor

        public MultiTimeSetSheets()
        {
            _timeSetBasicSheetsList = new List<ComTimeSetBasicSheet>();
        }

        #endregion

        #region Member Function
        public void AddTimeSetSheet(ComTimeSetBasicSheet tsetSheet)
        {
            _timeSetBasicSheetsList.Add(tsetSheet);
        }

        public ComTimeSetBasicSheet FindTimeSetSheet(string name)
        {
            name = Path.GetFileNameWithoutExtension(name);
            return _timeSetBasicSheetsList.Find(x => x.Name.Equals(name,StringComparison.CurrentCultureIgnoreCase));
        }
        #endregion
    }
}