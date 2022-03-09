using System.Linq;
using IgxlData.IgxlBase;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.PostAction.ModifyChannelMap
{
    public class ChannelMapPostMain
    {
        public void WorkFlow()
        {
            var channelMapSheets = TestProgram.IgxlWorkBk.ChannelMapSheets;
            var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMap != null)
                foreach (var channelMapSheet in channelMapSheets.Keys)
                {
                    var chData = channelMapSheets[channelMapSheet].ChannelMapRows;
                    var channelMapPins = chData.Select(channel => channel.DeviceUnderTestPinName.ToUpper()).ToList();
                    var channelMapNonExistPins = (from pin in pinMap.PinList
                        where !channelMapPins.Contains(pin.PinName.ToUpper())
                        select pin.PinName).ToList();
                    if (channelMapNonExistPins.Count > 0)
                        foreach (var pin in channelMapNonExistPins)
                        {
                            var channelMapRow = new ChannelMapRow();
                            channelMapRow.DeviceUnderTestPinName = pin;
                            channelMapRow.Type = "N/C";
                            channelMapSheets[channelMapSheet].AddRow(channelMapRow);
                        }
                }
        }
    }
}