using System;

namespace IgxlData.IgxlBase
{
    [Serializable]
    public class TimingRow
    {
        #region Property

        public string PinGrpName { get; set; }
        public string PinGrpClockPeriod { get; set; }
        public string PinGrpSetup { get; set; }
        public string DataSrc { get; set; }
        public string DataFmt { get; set; }
        public string DriveOn { get; set; }
        public string DriveData { get; set; }
        public string DriveReturn { get; set; }
        public string DriveOff { get; set; }
        public string CompareMode { get; set; }
        public string CompareOpen { get; set; }
        public string CompareClose { get; set; }
        public string CompareClkOffset { get; set; }
        public string CompareRefOffset { get; set; }
        public string EdgeMode { get; set; }
        public string Comment { get; set; }

        #endregion

        #region Constructor

        public TimingRow()
        {
        }

        public TimingRow(string pinGrpName, string pinGrpClockPeriod, string pinGrpSetup, string dataSrc,
            string dataFmt,
            string driveOn, string driveData, string driveReturn, string driveOff, string compareMode,
            string compareOpen, string compareClose, string compareRefOffset, string edgeMode, string comment)
        {
            PinGrpName = pinGrpName;
            PinGrpClockPeriod = pinGrpClockPeriod;
            PinGrpSetup = pinGrpSetup;
            DataSrc = dataSrc;
            DataFmt = dataFmt;
            DriveOn = driveOn;
            DriveData = driveData;
            DriveReturn = driveReturn;
            DriveOff = driveOff;
            CompareMode = compareMode;
            CompareOpen = compareOpen;
            CompareClose = compareClose;
            CompareRefOffset = compareRefOffset;
            EdgeMode = edgeMode;
            Comment = comment;
        }

        #endregion
    }
}