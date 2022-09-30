using System;
using System.Collections.Generic;

namespace IgxlData.IgxlBase
{
    public class CharStepConst : IgxlRow
    {
        public const string ModeXShmoo = "X Shmoo";
        public const string ModeYShmoo = "Y Shmoo";
        public const string ModeZShmoo = "Z Shmoo";
        public const string ModeAdjust = "Adjust";
        public const string ModeAdjustFrom = "Adjust From";
        public const string ModeMargin = "Margin";
        public const string ModeMeasure = "Measure";
        public const string ModeNone = "None";

        public const string ParameterTypeAcSpec = "AC Spec";
        public const string ParameterTypeDcSpec = "DC Spec";
        public const string ParameterTypeEdge = "Edge";
        public const string ParameterTypeGlobalSpec = "Global Spec";
        public const string ParameterTypeLevel = "Level";
        public const string ParameterTypePeriod = "Period";
        public const string ParameterTypeProtocolAware = "Protocol Aware";
        public const string ParameterTypeSerialTiming = "Serial Timing";

        public const string RangeCalcFieldStepSize = "Step Size";
        public const string RangeCalcFieldSteps = "Steps";
        public const string RangeCalcFieldTo = "To";
        public const string RangeCalcFieldFrom = "From";
        public const string RangeCalcFieldFromTo = "FromTo";

        public const string AlgorithmNameBinary = "Binary";
        public const string AlgorithmNameEdge = "Edge";
        public const string AlgorithmNameJump = "Jump";
        public const string AlgorithmNameLinear = "Linear";
        public const string AlgorithmNameList = "List";

        public const string PostStepFunctionPrintShmooInfo = "PrintShmooInfo";

        public static readonly List<string> AlgorithmName = new List<string>
            {"Binary", "Edge", "Jump", "Linear", "List"};

        public static readonly Dictionary<string, string> ParameterType =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"AC Spec", "AC Spec"}, {"DCSpec", "DC Spec"}, {"Edge", "Edge"}, {"GlobalSpec", "Global Spec"},
                {"Level", "Level"}, {"Period", "Period"}, {"Protocol Aware", "Protocol Aware"},
                {"SerialTiming", "Serial Timing"}
            };

        public static readonly Dictionary<string, string> ParameterName =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"RefOffset", "Ref Offset"}, {"DriveDelay", "Drive Delay"}, {"ReceiveDelay", "Receive Delay"},
                {"HiZDelay", "HiZ Delay"}, {"ReferenceOffset", "Reference Offset"}
            };
    }
}