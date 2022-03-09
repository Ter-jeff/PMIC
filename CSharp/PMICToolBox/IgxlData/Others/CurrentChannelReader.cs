using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Windows.Forms;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;

namespace IgxlData.Others
{
    [Serializable]
    public class CurrentChannelMapRow : IgxlRow
    {
        #region Properity
        public string TesterChannel { get; set; }
        public string DibChannel { get; set; }
        public string SignalName { get; set; }
        #endregion

        #region Constructor
        public CurrentChannelMapRow()
        {
            this.TesterChannel = string.Empty;
            this.DibChannel = string.Empty; ;
            this.SignalName = string.Empty;
        }
        #endregion

        public CurrentChannelMapRow DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as CurrentChannelMapRow;
            }
        }
    }

    public class CurrentChannelReader
    {
        #region Field
        private Dictionary<string, List<CurrentChannelMapRow>> _currentChannel;
        #endregion

        #region Properity
        public Dictionary<string, List<CurrentChannelMapRow>> CurrentChannel
        {
            get { return _currentChannel; }
            set { _currentChannel = value; }
        }
        #endregion

        #region Constructor
        public CurrentChannelReader()
        {
            _currentChannel = new Dictionary<string, List<CurrentChannelMapRow>>();
        }
        #endregion

        public void ReadFile(string path)
        {
            try
            {
                if (!File.Exists(path)) return;

                string line;
                bool startflag = false;
                string instrumentName = "";
                _currentChannel = new Dictionary<string, List<CurrentChannelMapRow>>();
                List<CurrentChannelMapRow> rows = new List<CurrentChannelMapRow>();
                using (StreamReader sr = new StreamReader(path))
                {
                    while (!sr.EndOfStream)
                    {
                        line = sr.ReadLine();

                        if (!startflag)
                        {
                            if (line.StartsWith("TesterChannel	DibChannel	SignalName", StringComparison.OrdinalIgnoreCase))
                            {
                                startflag = true;
                                sr.ReadLine(); //TesterChannel	DibChannel	SignalName
                                line = sr.ReadLine(); //empty line
                            }
                        }

                        bool nextflag = false;
                        if (startflag)
                        {
                            string[] lineSpt = line.Split('\t');
                            if (lineSpt.Length == 1 && !string.IsNullOrEmpty(lineSpt[0]))
                                instrumentName = lineSpt[0];

                            CurrentChannelMapRow row = new CurrentChannelMapRow();
                            if (lineSpt.Length == 5)
                            {
                                row.TesterChannel = lineSpt[0];
                                row.DibChannel = lineSpt[2];
                                row.SignalName = lineSpt[4];
                                rows.Add(row);
                            }

                            if (string.IsNullOrEmpty(line))
                                nextflag = true;
                        }

                        if (nextflag)
                        {
                            if (!string.IsNullOrEmpty(instrumentName) && !_currentChannel.ContainsKey(instrumentName))
                            {
                                var newRow = rows.Select(x => x.DeepClone()).ToList();
                                _currentChannel.Add(instrumentName, newRow);
                                rows.Clear();
                            }
                        }
                    }
                    var Row = rows.Select(x => x.DeepClone()).ToList();
                    _currentChannel.Add(instrumentName, Row);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format(ex.ToString()));
            }
        }

        public Dictionary<string, string> GetPogomapping(string group)
        {
            var currentChannelMap = _currentChannel[group];
            return currentChannelMap.ToDictionary(row => row.DibChannel.Split('.')[1], row => row.SignalName.Split('.')[1]);
        }

        public List<string> GetPinList(ChannelMapSheet sheet, string type)
        {
            List<string> pins = new List<string>();
            if (_currentChannel.ContainsKey(type))
            {
                var currentChannelMap = _currentChannel[type];
                foreach (var row in sheet.ChannelMapRows)
                {
                    foreach (var site in row.Sites)
                    {
                        if (site.Contains('.'))
                        {
                            var chan = site.Split('.').Last();
                            foreach (var item in currentChannelMap)
                            {
                                var DibChannel = item.DibChannel.Split('.').Last();
                                if (chan.Equals(DibChannel, StringComparison.CurrentCultureIgnoreCase))
                                {
                                    pins.Add(row.DiviceUnderTestPinName);
                                    break;
                                }
                                var SignalName = item.SignalName.Split('.').Last();
                                if (chan.Equals(SignalName, StringComparison.CurrentCultureIgnoreCase))
                                {
                                    pins.Add(row.DiviceUnderTestPinName);
                                    break;
                                }
                                var TesterChannel = item.TesterChannel.Split('.').Last();
                                if (chan.Equals(TesterChannel, StringComparison.CurrentCultureIgnoreCase))
                                {
                                    pins.Add(row.DiviceUnderTestPinName);
                                    break;
                                }
                            }
                        }
                    }
                }
            }
            return pins.Distinct().ToList();
        }

        public List<string> GetVsmPinList(ChannelMapSheet sheet)
        {
            List<string> pins = new List<string>();
            List<CurrentChannelMapRow> currentChannelMapRows = new List<CurrentChannelMapRow>();
            if (_currentChannel.ContainsKey("VSM, VSM"))
                currentChannelMapRows.AddRange(_currentChannel["VSM, VSM"]);

            GetPins(sheet, currentChannelMapRows, pins);

            return pins.Distinct().ToList();
        }

        public List<string> GetUvsPinList(ChannelMapSheet sheet)
        {
            List<string> pins = new List<string>();
            List<CurrentChannelMapRow> currentChannelMapRows = new List<CurrentChannelMapRow>();
            if (_currentChannel.ContainsKey("VHDVS"))
                currentChannelMapRows.AddRange(_currentChannel["VHDVS"]);

            GetPins(sheet, currentChannelMapRows, pins);
            return pins.Distinct().ToList();
        }

        public List<string> GetHexVsPinList(ChannelMapSheet sheet)
        {
            List<string> pins = new List<string>();
            List<CurrentChannelMapRow> currentChannelMapRows = new List<CurrentChannelMapRow>();
            if (_currentChannel.ContainsKey("HexVS"))
                currentChannelMapRows.AddRange(_currentChannel["HexVS"]);

            GetPins(sheet, currentChannelMapRows, pins);
            return pins.Distinct().ToList();
        }

        private static void GetPins(ChannelMapSheet sheet, List<CurrentChannelMapRow> currentChannelMapRows, List<string> pins)
        {

            foreach (var row in sheet.ChannelMapRows)
            {
                if (!row.Type.StartsWith("DCVS", StringComparison.CurrentCultureIgnoreCase))
                    continue;

                foreach (var site in row.Sites)
                {
                    if (site.Contains('.'))
                    {
                        var chan = site.Split('.').Last();
                        foreach (var item in currentChannelMapRows)
                        {
                            var DibChannel = item.DibChannel.Split('.').Last();
                            if (chan.Equals(DibChannel, StringComparison.CurrentCultureIgnoreCase))
                            {
                                pins.Add(row.DiviceUnderTestPinName);
                                break;
                            }
                            var SignalName = item.SignalName.Split('.').Last();
                            if (chan.Equals(SignalName, StringComparison.CurrentCultureIgnoreCase))
                            {
                                pins.Add(row.DiviceUnderTestPinName);
                                break;
                            }
                            var TesterChannel = item.TesterChannel.Split('.').Last();
                            if (chan.Equals(TesterChannel, StringComparison.CurrentCultureIgnoreCase))
                            {
                                pins.Add(row.DiviceUnderTestPinName);
                                break;
                            }
                        }
                    }
                }
            }
        }

     
    }
}