using IgxlData.IgxlBase;
using System;
using System.IO;

namespace IgxlData.IgxlSheet
{
    [Serializable]
    public abstract class IgxlSheet : IgxlItem
    {
        #region Field

        protected StreamWriter IgxlWriter;

        #endregion

        #region Constructer

        protected IgxlSheet(string name)
        {
            Name = name;
        }

        #endregion

        #region Property

        public string Name { get; set; }

        #endregion

        #region Member function

        protected void GetStreamWriter(string file)
        {
            IgxlWriter = new StreamWriter(file);
        }

        protected void CloseStreamWriter()
        {
            IgxlWriter?.Close();
        }

        protected abstract void WriteHeader();

        protected abstract void WriteColumnsHeader();

        protected abstract void WriteRows();

        public abstract void Write(string file, string version);

        #endregion
    }
}