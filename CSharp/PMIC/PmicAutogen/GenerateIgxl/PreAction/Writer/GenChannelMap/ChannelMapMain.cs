using IgxlData.IgxlReader;
using IgxlData.IgxlSheets;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.PreAction.Writer.GenChannelMap
{
    public class ChannelMapMain
    {
        public KeyValuePair<IgxlSheet, string> WorkFlow(string sheet)
        {
            var readChanMapSheet = new ReadChanMapSheet();
            var channelMapSheet = readChanMapSheet.GetSheet(sheet);

            ChannelMapPostAction();

            var igxlSheet = new KeyValuePair<IgxlSheet, string>(channelMapSheet, FolderStructure.DirChannelMap);
            return igxlSheet;
        }

        private void ChannelMapPostAction()
        {
            ModifyPinMapByChannelMap();
        }

        private void ModifyPinMapByChannelMap()
        {
            if (TestProgram.IgxlWorkBk.PinMapPair.Value != null)
            {
                var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
                foreach (var pin in pinMap.PinList)
                    if (TestProgram.IgxlWorkBk.ChannelMapSheets != null)
                        if (TestProgram.IgxlWorkBk.ChannelMapSheets.SelectMany(x => x.Value.ChannelMapRows).ToList()
                            .Exists(y =>
                                y.DeviceUnderTestPinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase)))
                        {
                            var channelMapRow = TestProgram.IgxlWorkBk.ChannelMapSheets
                                .SelectMany(x => x.Value.ChannelMapRows).ToList().Find(y =>
                                    y.DeviceUnderTestPinName.Equals(pin.PinName, StringComparison.OrdinalIgnoreCase));
                            pin.ChannelType = channelMapRow.Type;
                            pin.InstrumentType = channelMapRow.InstrumentType;
                        }
            }
        }
    }
}