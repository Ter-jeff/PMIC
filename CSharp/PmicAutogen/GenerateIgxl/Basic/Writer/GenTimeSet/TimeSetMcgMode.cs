using System;
using System.Collections.Generic;
using System.Linq;
using IgxlData.IgxlBase;
using IgxlData.Others.MultiTimeSet;
using IgxlData.Others.PatternListCsvFile;
using PmicAutogen.Singleton;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenTimeSet
{
    public class TimeSetMcgMode
    {
        #region Constructor

        public TimeSetMcgMode(List<ProtocolAwarePin> nWirePins)
        {
            _nWirePins = nWirePins;
        }

        #endregion

        public void ConvertFlow(List<ComTimeSetBasicSheet> timeSetSheets, List<PatternListCsvRow> patternListCsvRows)
        {
            foreach (var timeSetSheet in timeSetSheets)
            {
                var row = patternListCsvRows.Find(p => p.TimeSetVersion.Equals(timeSetSheet.SheetName));
                if (row != null && !JudgeHardIpPattern(row.PatternName)) continue;
                foreach (var tset in timeSetSheet.TimeSetsData)
                foreach (var wirePin in _nWirePins)
                    if (tset.TimingRows.Exists(p => p.PinGrpName.Equals(wirePin.OutClk, StringComparison.OrdinalIgnoreCase)))
                        ModifyMcgMode(tset, wirePin);
            }
        }

        private bool JudgeHardIpPattern(string pattern)
        {
            var subName = pattern.Split('_').ToList();
            if (subName.Count > 5)
            {
                //if (_patternDigital2TypeDic.ContainsKey(subName[2].ToUpper()) && 
                //    _patternDigital2TypeDic[subName[2].ToUpper()] == HardIp)
                //{
                //    return true;
                //}

                //if (_patternDigital4TypeDic.ContainsKey(subName[4].ToUpper()) &&
                //    _patternDigital4TypeDic[subName[4].ToUpper()] == HardIp)
                //{
                //    return true;
                //}   
            }

            return false;
        }

        private void ModifyMcgMode(Tset tset, ProtocolAwarePin nWirePin)
        {
            var period = "= 1/_" + nWirePin.CreateFreqVarName(); //CreateFreqSpecName();
            var driveData = "= 1/(2*_" + nWirePin.CreateFreqVarName() + ")"; //CreateFreqSpecName();

            if (nWirePin.Freq < _digitalChannelMinFreq)
                //If the frequency exceed the Low limit, do not need to change time set
                return;

            var timingRow = tset.TimingRows.Find(p => p.PinGrpName.Equals(nWirePin.OutClk, StringComparison.OrdinalIgnoreCase));
            timingRow.PinGrpSetup = SetUpClock;

            if (nWirePin.Freq > _digitalChannelMaxFreq)
            {
                //If the frequency exceed the high limit, change the Mode as "clock_2X"
                period = "= 2/_" + nWirePin.CreateFreqVarName(); //CreateFreqSpecName();
                driveData = "= 1/_" + nWirePin.CreateFreqVarName(); //CreateFreqSpecName();
                timingRow.PinGrpSetup = SetUpClock2X;
            }

            /*
             t0t1       t2           t3t0t1     t2          t3
                          ____________           ____________ 
             _____________            ___________
             */
            timingRow.PinGrpClockPeriod = period;
            timingRow.DataSrc = SrcAllHi;
            timingRow.DataFmt = FmtRl;
            timingRow.DriveData = driveData;
            timingRow.DriveReturn = period;
            timingRow.CompareMode = Off;

            if (nWirePin.PinType == IoPinType.Diff)
            {
                timingRow = tset.TimingRows.Find(p => p.PinGrpName.Equals(nWirePin.OutClkDiff, StringComparison.OrdinalIgnoreCase));
                if (timingRow == null)
                {
                    timingRow = new TimingRow();
                    timingRow.PinGrpName = nWirePin.OutClkDiff;
                    tset.AddTimingRow(timingRow);
                }

                /*
             t0t1       t2           t3t0t1     t2          t3
             _____________            ____________            
                           ___________            ____________
             */
                timingRow.PinGrpSetup = SetUpClock;
                timingRow.PinGrpClockPeriod = period;
                timingRow.DataSrc = SrcAllLo;
                timingRow.DataFmt = FmtRh;
                timingRow.DriveData = driveData;
                timingRow.DriveReturn = period;
                timingRow.CompareMode = Off;
            }
        }

        #region Field

        private const string SetUpClock = "clock";
        private const string SetUpClock2X = "clock_2X";
        private const string SrcAllHi = "ALLHI";
        private const string SrcAllLo = "ALLLO";
        private const string FmtRl = "RL";
        private const string FmtRh = "RH";
        private const string Off = "Off";
        private readonly List<ProtocolAwarePin> _nWirePins;

        private readonly double _digitalChannelMaxFreq = 550e6; //550MHZ
        private readonly double _digitalChannelMinFreq = 700e3; //700KHZ

        #endregion
    }
}