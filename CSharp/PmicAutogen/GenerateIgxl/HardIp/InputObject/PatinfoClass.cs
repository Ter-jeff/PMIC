using System;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.HardIp.InputObject
{
    public class HardIpInfoNew
    {
        public List<HardIpSeqInfoNew> SeqInfo;
        public List<string> TrimLoopHLimit;
        public List<string> TrimLoopLLimit;

        public HardIpInfoNew()
        {
            MeasWait = "";
            RfSetup = "";
            LLimit = "";
            HLimit = "";
            ExpectValue = "";
            ForceValue = "";
            ForcePin = "";
            ForceType = "";
            MeasStoreName = "";
            CalcStoreName = "";
            CalcEquation = "";
            MeasName = "";
            MeasPin = "";
            MeasSeq = "";
            TrimLoopLLimit = new List<string>();
            TrimLoopHLimit = new List<string>();
        }

        public string MeasSeq { get; set; }

        public string MeasPin { get; set; }

        public string MeasName { get; set; }

        public string CalcEquation { get; set; }

        public string CalcStoreName { get; set; }

        public string MeasStoreName { get; set; }

        public string ForceType { get; set; }

        public string ForcePin { get; set; }

        public string ForceValue { get; set; }

        public string ExpectValue { get; set; }

        public string HLimit { get; set; }

        public string LLimit { get; set; }

        public string RfSetup { get; set; }

        public string MeasWait { get; set; }
    }

    public class HardIpSeqInfo
    {
        public List<MeasPin> MeasPins = new List<MeasPin>();
    }

    public class HardIpReference
    {
        public string CapStoreName;
        public string DigSrcAssignment;
        public List<string> ForceIPinList;
        public List<string> ForcePrePatRaw;
        public List<string> ForceVPinList;

        public bool IllegalChar;
        public string IsCapLsbFirstInfo;
        public string IsSendLsbFirstInfo;
        public List<string> MeasFdiffPinList;
        public List<string> MeasFPinList;
        public List<string> MeasIdiffPinList;
        public List<string> MeasIioPinList;

        public List<string> MeasIPowerPinList;
        public string MeasName;
        public List<string> MeasR1PinList;
        public List<string> MeasR2PinList;
        public List<string> MeasVdiff2PinList;
        public List<string> MeasVdiffPinList;
        public List<string> MeasVioPinList;
        public List<string> MeasVocmPinList;
        public List<string> MeasVPowerPinList;
        public List<string> MiscInfo;
        public HardIpInfoNew NewInfo;
        public PowerOnFly PowerOnFlyItem;
        public List<HardIpSeqInfo> SeqInfo;
        public List<string> SpecialLimits;

        public HardIpReference()
        {
            IsCapLsbFirstInfo = "";
            IsSendLsbFirstInfo = "";
            PowerOnFlyItem = null;
            CapStoreName = "";
            MeasName = "";
            ForceIPinList = new List<string>();
            ForceVPinList = new List<string>();
            IllegalChar = false;
            DigSrcAssignment = "";
            MiscInfo = new List<string>();
            ForcePrePatRaw = new List<string>();
            SpecialLimits = new List<string>();
            SeqInfo = new List<HardIpSeqInfo>();
            MeasR2PinList = new List<string>();
            MeasR1PinList = new List<string>();
            MeasVocmPinList = new List<string>();
            MeasFdiffPinList = new List<string>();
            MeasVdiff2PinList = new List<string>();
            MeasIdiffPinList = new List<string>();
            MeasVdiffPinList = new List<string>();
            MeasFPinList = new List<string>();
            MeasVioPinList = new List<string>();
            MeasIioPinList = new List<string>();
            MeasVPowerPinList = new List<string>();
            MeasIPowerPinList = new List<string>();
            TestName = "";
            DigSrcDataWidth = "";
            DigSrcAssign = "";
            MeasSeqStr = "";
            DsscOut = "";
            SendPinName = "";
            CapPinName = "";
            SendBitName = "";
            CapBitName = "";
            SendBitStr = "";
            CapBitStr = "";
            SubRoutine = "";
            TimeSet = "";
            Payload = "";
        }

        public string Payload { get; set; }
        public string TimeSet { get; set; }
        public string SubRoutine { get; set; }
        public int CapBit { get; set; }
        public int SendBit { get; set; }
        public string CapBitStr { get; set; }
        public string SendBitStr { get; set; }
        public string CapBitName { get; set; }
        public string SendBitName { get; set; }
        public string CapPinName { get; set; }
        public string SendPinName { get; set; }
        public string DsscOut { get; set; }
        public string MeasSeqStr { get; set; }
        public string DigSrcAssign { get; set; }
        public string DigSrcDataWidth { get; set; }
        public string TestName { get; set; }
        public bool IsIgnoreComment { get; set; }

        public object Clone()
        {
            return MemberwiseClone();
        }

        public string IsMsbFirst()
        {
            if (string.IsNullOrEmpty(IsSendLsbFirstInfo))
                return "";
            foreach (var isLsb in IsSendLsbFirstInfo.Split('+'))
                if (isLsb.Equals("false", StringComparison.OrdinalIgnoreCase))
                    return "TRUE";
            foreach (var isLsb in IsCapLsbFirstInfo.Split('+'))
                if (isLsb.Equals("false", StringComparison.OrdinalIgnoreCase))
                    return "TRUE";
            return "";
        }
    }

    public class PowerOnFly
    {
        public List<string> VAlt = new List<string>();
    }

    public class HardIpSeqInfoNew
    {
        public List<MeasPin> Calc = new List<MeasPin>();
        public List<MeasPin> MeasPins = new List<MeasPin>();
    }
}