using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;

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

    [Serializable]
    public class CharStep : IgxlRow
    {
        #region Constructor

        public CharStep(string setupName, string stepName)
        {
            VoltageType = string.Empty;
            SetupName = setupName;
            StepName = stepName;
            AdjustBackOff = string.Empty;
            AdjustFromSetup = string.Empty;
            AdjustSpecName = string.Empty;
            AlgorithmArgs = string.Empty;
            AlgorithmName = string.Empty;
            AlgorithmResultsCheck = string.Empty;
            AlgorithmTransition = string.Empty;
            ApplyToPinExecMode = string.Empty;
            ApplyToPins = string.Empty;
            ApplyToTimeSets = string.Empty;
            Arguments = string.Empty;
            DeviceMarginContexts = string.Empty;
            DeviceMarginInstances = string.Empty;
            DeviceMarginPatterns = string.Empty;
            Function = string.Empty;
            Mode = string.Empty;
            OutputFormat = string.Empty;
            OutputSheet = string.Empty;
            OutputTextFile = string.Empty;
            OutputToDataLog = string.Empty;
            OutputToImmediateWin = string.Empty;
            OutputToOutputWin = string.Empty;
            OutputToSheet = string.Empty;
            OutputToTextFile = string.Empty;
            ParameterName = string.Empty;
            ParameterType = string.Empty;
            PerformTest = string.Empty;
            PostPointArgs = string.Empty;
            PostPointFunction = string.Empty;
            PostSetupArgs = string.Empty;
            PostSetupFunction = string.Empty;
            PostStepArgs = string.Empty;
            PostStepFunction = string.Empty;
            PrePointArgs = string.Empty;
            PrePointFunction = string.Empty;
            PreSetupArgs = string.Empty;
            PreSetupFunction = string.Empty;
            PreStepArgs = string.Empty;
            PreStepFunction = string.Empty;
            RangeCalcField = string.Empty;
            RangeFrom = string.Empty;
            RangeSteps = string.Empty;
            RangeStepSize = string.Empty;
            RangeTo = string.Empty;
            TestLimitHigh = string.Empty;
            TestLimitLow = string.Empty;
            Comment = string.Empty;
            AxisExecutionOrder = "X-Y[-Z]";
            SuspendDataLog = "TRUE";
        }

        #endregion

        public CharStep DeepClone()
        {
            using (Stream objectStream = new MemoryStream())
            {
                IFormatter formatter = new BinaryFormatter();
                formatter.Serialize(objectStream, this);
                objectStream.Seek(0, SeekOrigin.Begin);
                return formatter.Deserialize(objectStream) as CharStep;
            }
        }

        #region Property

        public string VoltageType { get; set; }
        public string SetupName { get; set; }
        public string StepName { get; set; }
        public string Mode { get; set; }

        //Parameter
        public string ParameterType { get; set; }

        public string ParameterName { get; set; }

        //Range
        public string RangeCalcField { get; set; }
        public string RangeFrom { get; set; }
        public string RangeTo { get; set; }
        public string RangeSteps { get; set; }

        public string RangeStepSize { get; set; }

        //PerformTest
        public string PerformTest { get; set; }

        //Test Limits
        public string TestLimitHigh { get; set; }

        public string TestLimitLow { get; set; }

        //Algorithm
        public string AlgorithmName { get; set; }
        public string AlgorithmArgs { get; set; }
        public string AlgorithmResultsCheck { get; set; }

        public string AlgorithmTransition { get; set; }

        //Apply To
        public string ApplyToPins { get; set; }
        public string ApplyToPinExecMode { get; set; }
        public string ApplyToTimeSets { get; set; }

        //Device Margin
        public string DeviceMarginContexts { get; set; }
        public string DeviceMarginInstances { get; set; }

        public string DeviceMarginPatterns { get; set; }

        //Adjust
        public string AdjustBackOff { get; set; }
        public string AdjustFromSetup { get; set; }

        public string AdjustSpecName { get; set; }

        //Function
        public string AxisExecutionOrder { set; get; }

        public string Function { get; set; }

        //Arguments
        public string Arguments { get; set; }

        //Interpose Functions
        public string PreSetupFunction { get; set; }
        public string PreSetupArgs { get; set; }
        public string PreStepFunction { get; set; }
        public string PreStepArgs { get; set; }
        public string PrePointFunction { get; set; }
        public string PrePointArgs { get; set; }
        public string PostPointFunction { get; set; }
        public string PostPointArgs { get; set; }
        public string PostStepFunction { get; set; }
        public string PostStepArgs { get; set; }
        public string PostSetupFunction { get; set; }

        public string PostSetupArgs { get; set; }

        //Output
        public string OutputFormat { get; set; }
        public string OutputTextFile { get; set; }
        public string OutputSheet { get; set; }
        public string SuspendDataLog { set; get; }

        //Output Destinations
        public string OutputToTextFile { get; set; }
        public string OutputToSheet { get; set; }
        public string OutputToDataLog { get; set; }
        public string OutputToImmediateWin { get; set; }

        public string OutputToOutputWin { get; set; }

        //Comment
        public string Comment { get; set; }

        #endregion
    }
}