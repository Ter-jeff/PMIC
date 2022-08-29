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
    public class CurrentChannelMapRow : IgxlItem
    {
        #region Constructor

        public CurrentChannelMapRow()
        {
            TesterChannel = string.Empty;
            DibChannel = string.Empty;
            SignalName = string.Empty;
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

        #region Property

        public string TesterChannel { get; set; }
        public string DibChannel { get; set; }
        public string SignalName { get; set; }

        #endregion
    }

    public class CurrentChannelReader
    {
        #region Constructor

        public CurrentChannelReader()
        {
            CurrentChannel = new Dictionary<string, List<CurrentChannelMapRow>>();
        }

        #endregion

        #region Property

        public Dictionary<string, List<CurrentChannelMapRow>> CurrentChannel { get; set; }

        #endregion

        public void ReadFile(string path)
        {
            try
            {
                if (!File.Exists(path)) return;

                var startFlag = false;
                var instrumentName = "";
                CurrentChannel = new Dictionary<string, List<CurrentChannelMapRow>>();
                var rows = new List<CurrentChannelMapRow>();
                using (var sr = new StreamReader(path))
                {
                    while (!sr.EndOfStream)
                    {
                        var line = sr.ReadLine();

                        if (!startFlag)
                            if (line.StartsWith("TesterChannel	DibChannel	SignalName",
                                    StringComparison.OrdinalIgnoreCase))
                            {
                                startFlag = true;
                                sr.ReadLine(); //TesterChannel	DibChannel	SignalName
                                line = sr.ReadLine(); //empty line
                            }

                        var nextFlag = false;
                        if (startFlag)
                        {
                            var lineSpt = line.Split('\t');
                            if (lineSpt.Length == 1 && !string.IsNullOrEmpty(lineSpt[0]))
                                instrumentName = lineSpt[0];

                            var currentChannelMapRow = new CurrentChannelMapRow();
                            if (lineSpt.Length == 5)
                            {
                                currentChannelMapRow.TesterChannel = lineSpt[0];
                                currentChannelMapRow.DibChannel = lineSpt[2];
                                currentChannelMapRow.SignalName = lineSpt[4];
                                rows.Add(currentChannelMapRow);
                            }

                            if (string.IsNullOrEmpty(line))
                                nextFlag = true;
                        }

                        if (nextFlag)
                            if (!string.IsNullOrEmpty(instrumentName) && !CurrentChannel.ContainsKey(instrumentName))
                            {
                                var newRow = rows.Select(x => x.DeepClone()).ToList();
                                CurrentChannel.Add(instrumentName, newRow);
                                rows.Clear();
                            }
                    }

                    var row = rows.Select(x => x.DeepClone()).ToList();
                    CurrentChannel.Add(instrumentName, row);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        public Dictionary<string, string> GetPogoMapping(string group)
        {
            var currentChannelMap = CurrentChannel[group];
            return currentChannelMap.ToDictionary(row => row.DibChannel.Split('.')[1],
                row => row.SignalName.Split('.')[1]);
        }

        public List<string> GetPinList(ChannelMapSheet sheet, string type)
        {
            var pins = new List<string>();
            if (CurrentChannel.ContainsKey(type))
            {
                var currentChannelMap = CurrentChannel[type];
                foreach (var row in sheet.ChannelMapRows)
                foreach (var site in row.Sites)
                    if (site.Contains('.'))
                    {
                        var chan = site.Split('.').Last();
                        foreach (var item in currentChannelMap)
                        {
                            var dibChannel = item.DibChannel.Split('.').Last();
                            if (chan.Equals(dibChannel, StringComparison.CurrentCultureIgnoreCase))
                            {
                                pins.Add(row.DeviceUnderTestPinName);
                                break;
                            }

                            var signalName = item.SignalName.Split('.').Last();
                            if (chan.Equals(signalName, StringComparison.CurrentCultureIgnoreCase))
                            {
                                pins.Add(row.DeviceUnderTestPinName);
                                break;
                            }

                            var testerChannel = item.TesterChannel.Split('.').Last();
                            if (chan.Equals(testerChannel, StringComparison.CurrentCultureIgnoreCase))
                            {
                                pins.Add(row.DeviceUnderTestPinName);
                                break;
                            }
                        }
                    }
            }

            return pins.Distinct().ToList();
        }

        public List<string> GetVsmPinList(ChannelMapSheet sheet)
        {
            var pins = new List<string>();
            var currentChannelMapRows = new List<CurrentChannelMapRow>();
            if (CurrentChannel.ContainsKey("VSM, VSM"))
                currentChannelMapRows.AddRange(CurrentChannel["VSM, VSM"]);

            GetPins(sheet, currentChannelMapRows, pins);

            return pins.Distinct().ToList();
        }

        public List<string> GetUvsPinList(ChannelMapSheet sheet)
        {
            var pins = new List<string>();
            var currentChannelMapRows = new List<CurrentChannelMapRow>();
            if (CurrentChannel.ContainsKey("VHDVS"))
                currentChannelMapRows.AddRange(CurrentChannel["VHDVS"]);

            GetPins(sheet, currentChannelMapRows, pins);
            return pins.Distinct().ToList();
        }

        public List<string> GetHexVsPinList(ChannelMapSheet sheet)
        {
            var pins = new List<string>();
            var currentChannelMapRows = new List<CurrentChannelMapRow>();
            if (CurrentChannel.ContainsKey("HexVS"))
                currentChannelMapRows.AddRange(CurrentChannel["HexVS"]);

            GetPins(sheet, currentChannelMapRows, pins);
            return pins.Distinct().ToList();
        }

        private void GetPins(ChannelMapSheet sheet, List<CurrentChannelMapRow> currentChannelMapRows, List<string> pins)
        {
            foreach (var row in sheet.ChannelMapRows)
            {
                if (!row.Type.StartsWith("DCVS", StringComparison.CurrentCultureIgnoreCase))
                    continue;

                foreach (var site in row.Sites)
                    if (site.Contains('.'))
                    {
                        var chan = site.Split('.').Last();
                        foreach (var item in currentChannelMapRows)
                        {
                            var dibChannel = item.DibChannel.Split('.').Last();
                            if (chan.Equals(dibChannel, StringComparison.CurrentCultureIgnoreCase))
                            {
                                pins.Add(row.DeviceUnderTestPinName);
                                break;
                            }

                            var signalName = item.SignalName.Split('.').Last();
                            if (chan.Equals(signalName, StringComparison.CurrentCultureIgnoreCase))
                            {
                                pins.Add(row.DeviceUnderTestPinName);
                                break;
                            }

                            var testerChannel = item.TesterChannel.Split('.').Last();
                            if (chan.Equals(testerChannel, StringComparison.CurrentCultureIgnoreCase))
                            {
                                pins.Add(row.DeviceUnderTestPinName);
                                break;
                            }
                        }
                    }
            }
        }

        #region Field

        #endregion
    }
}