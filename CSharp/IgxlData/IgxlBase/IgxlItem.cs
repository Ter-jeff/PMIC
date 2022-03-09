using System;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public abstract class IgxlItem : ICloneable
    {
        public string ColumnA { get; set; }
        #region Member Function
        public object Clone()
        {
            return MemberwiseClone();
        }
        #endregion
    }
}