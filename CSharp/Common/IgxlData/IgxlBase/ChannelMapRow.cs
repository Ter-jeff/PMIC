﻿using System.Collections.Generic;
using System.Diagnostics;

namespace IgxlData.IgxlBase
{
    [DebuggerDisplay("{DeviceUnderTestPinName}")]
    public class ChannelMapRow : IgxlRow
    {
        public ChannelMapRow()
        {
            Sites = new List<string>();
            DeviceUnderTestPinName = "";
            DeviceUnderTestPackagePin = "";
            Type = "";
            Comment = "";
            InstrumentType = "";
        }

        public string DeviceUnderTestPinName { get; set; }
        public string DeviceUnderTestPackagePin { get; set; }
        public string Type { get; set; }
        public List<string> Sites { get; set; }
        public string Comment { get; set; }
        public string InstrumentType { get; set; }

    }
}