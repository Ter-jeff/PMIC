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

        public static readonly List<string> AlgorithmName = new List<string> { "Binary", "Edge", "Jump", "Linear", "List" };
        public static readonly Dictionary<string, string> ParameterType = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { { "AC Spec", "AC Spec" }, { "DCSpec", "DC Spec" }, { "Edge", "Edge" }, { "GlobalSpec", "Global Spec" }, { "Level", "Level" }, { "Period", "Period" }, { "Protocol Aware", "Protocol Aware" }, { "SerialTiming", "Serial Timing" } };
        public static readonly Dictionary<string, string> ParameterName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) { { "RefOffset", "Ref Offset" }, { "DriveDelay", "Drive Delay" }, { "ReceiveDelay", "Receive Delay" }, { "HiZDelay", "HiZ Delay" }, { "ReferenceOffset", "Reference Offset" } };
    }

    [Serializable]
    public class CharStep : IgxlRow
    {
        #region Properity
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
        public string TestLimitLow { get; set; }
        public string TestLimitHigh { get; set; }
        //Algorithm
        public string AlgorithmName { get; set; }
        public string AlgorithmArguments { get; set; }
        public string AlgorithmResultsCheck { get; set; }
        public string AlgorithmTransition { get; set; }
        //Apply To
        public string ApplyToPinExecMode { get; set; }
        public string ApplyToPins { get; set; }
        public string ApplyToTimeSets { get; set; }
        //Device Margin
        public string DeviceMarginContexts { get; set; }
        public string DeviceMarginInstances { get; set; }
        public string DeviceMarginPatterns { get; set; }
        //Adjust
        public string AdjustBackoff { get; set; }
        public string AdjustFromSetup { get; set; }
        public string AdjustSpecName { get; set; }
        //Function
        public string AxisExecutionOrder { set; get; }
        public string Function { get; set; }
        //Arguments
        public string Arguments { get; set; }
        public string InterposeFunctions { get; set; }
        //Interpose Functions
        public string PreSetup { get; set; }
        public string PreSetupArguments { get; set; }
        public string PreStep { get; set; }
        public string PreStepArguments { get; set; }
        public string PrePoint { get; set; }
        public string PrePointArguments { get; set; }
        public string PostPoint { get; set; }
        public string PostPointArguments { get; set; }
        public string PostStep { get; set; }
        public string PostStepArguments { get; set; }
        public string PostSetup { get; set; }
        public string PostSetupArguments { get; set; }
        //Output
        public string OutputFormat { get; set; }
        public string OutputTextFile { get; set; }
        public string OutputSheet { get; set; }
        public string OutputSuspendDatalog { set; get; }
        //Output Destinations
        public string OutputDestinationsTextFile { get; set; }
        public string OutputDestinationsSheet { get; set; }
        public string OutputDestinationsDatalog { get; set; }
        public string OutputDestinationsImmediateWin { get; set; }
        public string OutputDestinationsOutputWin { get; set; }
        //Comment
        public string Comment { get; set; }
       

        #endregion

        #region Constructor
        public CharStep(string setupName, string stepName)
        {
            VoltageType = string.Empty;
            SetupName = setupName;
            StepName = stepName;
            AdjustBackoff = string.Empty;
            AdjustFromSetup = string.Empty;
            AdjustSpecName = string.Empty;
            AlgorithmArguments = string.Empty;
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
            OutputDestinationsDatalog = string.Empty;
            OutputDestinationsImmediateWin = string.Empty;
            OutputDestinationsOutputWin = string.Empty;
            OutputDestinationsSheet = string.Empty;
            OutputDestinationsTextFile = string.Empty;
            ParameterName = string.Empty;
            ParameterType = string.Empty;
            PerformTest = string.Empty;
            PostPointArguments = string.Empty;
            PostPoint = string.Empty;
            PostSetupArguments = string.Empty;
            PostSetup = string.Empty;
            PostStepArguments = string.Empty;
            PostStep = string.Empty;
            PrePointArguments = string.Empty;
            PrePoint = string.Empty;
            PreSetupArguments = string.Empty;
            PreSetup = string.Empty;
            PreStepArguments = string.Empty;
            PreStep = string.Empty;
            RangeCalcField = string.Empty;
            RangeFrom = string.Empty;
            RangeSteps = string.Empty;
            RangeStepSize = string.Empty;
            RangeTo = string.Empty;
            TestLimitHigh = string.Empty;
            TestLimitLow = string.Empty;
            Comment = string.Empty;
            AxisExecutionOrder = "X-Y[-Z]";
            OutputSuspendDatalog = "TRUE";
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
    }
}