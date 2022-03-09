using IgxlData.IgxlBase;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace IgxlData.IgxlSheet
{
    [Serializable]
    public class ChannelMapSheet : IgxlSheet
    {
        #region Constructor

        public ChannelMapSheet(string name)
            : base(name)
        {
            Name = name;
        }

        #endregion

        #region Field

        private List<ChannelMapRow> _channelData = new List<ChannelMapRow>();
        private bool _isPogo;

        #endregion

        #region Property

        public List<ChannelMapRow> ChannelMapRows
        {
            get { return _channelData; }
            set { _channelData = value; }
        }

        public bool IsPogo
        {
            get { return _isPogo; }
            set { _isPogo = value; }

        }

        public int SiteNum { get; set; }

        #endregion

        #region Member Function

        protected override void WriteHeader()
        {
            IgxlWriter.WriteLine(
                "DTChanMap,version=2.6:platform=Jaguar:toprow=-1:leftcol=-1:rightcol=-1:dataformat=signal\tChannel Map");
            IgxlWriter.WriteLine("");
        }

        protected override void WriteColumnsHeader()
        {
            string viewMode = "";
            _isPogo = _channelData.Find(p => p.Sites.Exists(a => Regex.IsMatch(a, "ch", RegexOptions.IgnoreCase))) ==
                      null;

            IgxlWriter.WriteLine("\tDIB ID:\t\t\tView Mode:\t" + viewMode);
            IgxlWriter.WriteLine("\tUSL Tag:");
            IgxlWriter.WriteLine("\tDevice Under Test\t\tTester Channel");
            int siteNum = _channelData.Count != 0 ? _channelData[0].Sites.Count : 2;

            IgxlWriter.Write("\tPin Name\tPackage Pin\tType\t");
            for (int i = 0; i < siteNum; i++)
            {
                IgxlWriter.Write("Site " + i + "\t");
            }

            IgxlWriter.WriteLine("Comment");
        }

        protected override void WriteRows()
        {
            foreach (ChannelMapRow chanel in _channelData)
            {
                IgxlWriter.Write("\t" + chanel.DiviceUnderTestPinName + "\t" + chanel.DiviceUnderTestPackagePin + "\t" +
                                 chanel.Type + "\t");
                foreach (string site in chanel.Sites)
                {
                    IgxlWriter.Write(site + "\t");
                }

                IgxlWriter.Write(chanel.Comment);
                IgxlWriter.WriteLine();
            }
        }

        public override void Write(string file, string version)
        {
            Write2P6(file);
        }

        private void Write2P6(string file)
        {
            GetStreamWriter(file);
            WriteHeader();
            WriteColumnsHeader();
            WriteRows();
            CloseStreamWriter();
        }

        public void AddRow(ChannelMapRow channelRow)
        {
            _channelData.Add(channelRow);
        }

        #endregion
    }
}