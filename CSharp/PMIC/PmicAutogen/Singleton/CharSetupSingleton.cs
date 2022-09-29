using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.Others;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace PmicAutogen.Singleton
{
    public class CharSetupSingleton
    {
        private static CharSetupSingleton _instance;

        private CharSetupSingleton()
        {
            _specFinder = new SpecFinder(new List<GlobalSpecSheet> { TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value },
                TestProgram.IgxlWorkBk.AcSpecSheets.Select(x => x.Value).ToList(),
                TestProgram.IgxlWorkBk.DcSpecSheets.Select(x => x.Value).ToList(),
                TestProgram.IgxlWorkBk.LevelSheets.Select(x => x.Value).ToList());
            _shmooParameterTypeDictionary = GetShmooParameterTypeDictionary();
        }

        #region Singleton

        public static CharSetupSingleton Instance()
        {
            return _instance ?? (_instance = new CharSetupSingleton());
        }

        #endregion

        public Dictionary<string, string> GetShmooParameterTypeDictionary()
        {
            var newList = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (TestProgram.IgxlWorkBk == null)
                return newList;

            #region AC spec

            foreach (var item in TestProgram.IgxlWorkBk.AcSpecSheets)
            {
                var data = TestProgram.IgxlWorkBk.AcSpecSheets[item.Key];
                foreach (var row in data.AcSpecs)
                {
                    if (string.IsNullOrEmpty(row.Symbol))
                        break;
                    if (!newList.ContainsKey(row.Symbol))
                        newList.Add(row.Symbol, "AC spec");
                }
            }

            #endregion

            #region DC spec

            foreach (var item in TestProgram.IgxlWorkBk.DcSpecSheets)
            {
                var data = TestProgram.IgxlWorkBk.DcSpecSheets[item.Key];
                foreach (var row in data.GetDcSpecsData())
                    if (!newList.ContainsKey(row.Symbol))
                        newList.Add(row.Symbol, "DC spec");
            }

            #endregion

            #region Global Spec

            if (TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value != null)
            {
                var lGlobalSpecsListSource = TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.GetGlobalSpecs();
                for (var i = 0; i < lGlobalSpecsListSource.Count; i++)
                {
                    var glbSpec = lGlobalSpecsListSource[i];
                    if (!newList.ContainsKey(glbSpec.Symbol)) newList.Add(glbSpec.Symbol, "Global Spec");
                }
            }

            #endregion

            #region Level

            foreach (var item in TestProgram.IgxlWorkBk.LevelSheets)
            {
                var data = TestProgram.IgxlWorkBk.LevelSheets[item.Key];
                foreach (var row in data.LevelRows)
                    if (!newList.ContainsKey(row.Parameter))
                        newList.Add(row.Parameter, "Level");
            }

            #endregion

            #region TimeSet

            foreach (var item in TestProgram.IgxlWorkBk.TimeSetSheets)
            {
                var data = TestProgram.IgxlWorkBk.TimeSetSheets[item.Key];
                foreach (var row in data.TimeSetsData)
                    if (!newList.ContainsKey(row.Name))
                        newList.Add(row.Name, "Period");
            }

            #endregion

            var defaultList = new Dictionary<string, string>
            {
                //Edge: One of the timing edges (Close,Data,Off,On,Open,Ref Offset,Return) 
                {"Close", "Edge"},
                {"Data", "Edge"},
                {"Off", "Edge"},
                {"On", "Edge"},
                {"Open", "Edge"},
                {"RefOffset", "Edge"},
                {"Return", "Edge"},
                {"d0", "Edge"},
                {"d1", "Edge"},
                {"d2", "Edge"},
                {"d3", "Edge"},
                {"r0", "Edge"},
                {"r1", "Edge"},
                //Protocol Aware: One of these values: Clock Offset, Drive Delay, Receive Delay, HiZ Delay, Reference Offset. 
                {"ClockOffset", "Protocol Aware"},
                {"DriveDelay", "Protocol Aware"},
                {"ReceiveDelay", "Protocol Aware"},
                {"HiZDelay", "Protocol Aware"},
                {"ReferenceOffset", "Protocol Aware"}
                //Serial Timing: Timing set defined on a Serial Timing sheet (Master Period,DUT Period,Drive Delay,Receive Delay)
                //{"Master Period", "Serial Timing"},{"DUT Period", "Serial Timing"},{"Drive Delay", "Serial Timing"},{"Receive Delay", "Serial Timing"}
            };

            foreach (var item in defaultList)
                if (!newList.ContainsKey(item.Key))
                    newList.Add(item.Key, item.Value);
            return newList;
        }

        public string GetShmooParameterType(string name)
        {
            if (_shmooParameterTypeDictionary.ContainsKey(name)) return _shmooParameterTypeDictionary[name];
            return "";
        }

        public CharStep CreatePmicPwrPinCharStep(string setupName, string pin, ProdCharRow prodCharRow,
            string xYShmoo = XShmoo, string algorithm = CharStepConst.AlgorithmNameJump)
        {
            var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            var vmain = _specFinder.GetValue(prodCharRow.InstanceRow, pin);
            var parameterName = pinMap.GetChannelType(pin) == "DCVI" ? Vps : VMain;
            var applyToPins = pin;
            var parameterType = Instance().GetShmooParameterType(parameterName);
            double vmainValue;
            double.TryParse(vmain, out vmainValue);

            var rangeFrom = vmainValue * MinusRatio;
            var rangeTo = vmainValue * PlusRatio;

            var step = new CharStep(setupName, pin);
            step.Mode = xYShmoo;
            step.ParameterName = parameterName;
            step.ParameterType = parameterType;

            step.RangeCalcField = CharStepConst.RangeCalcFieldStepSize;
            step.RangeFrom = rangeFrom.ToString(CultureInfo.InvariantCulture);
            step.RangeTo = rangeTo.ToString(CultureInfo.InvariantCulture);
            step.RangeSteps = "10";

            step.AlgorithmName = algorithm;
            step.AlgorithmArgs = "6";

            step.ApplyToPins = applyToPins;
            step.ApplyToPinExecMode = Simultaneous;

            step.PostStepFunction = CharStepConst.PostStepFunctionPrintShmooInfo;
            step.PostStepArgs = pin;

            step.OutputFormat = Enhanced;
            step.OutputTextFile = Disable;
            step.OutputSheet = Disable;
            step.OutputToDataLog = Enable;
            step.OutputToImmediateWin = Disable;
            step.OutputToOutputWin = Disable;
            return step;
        }

        public CharStep CreatePwrPinCharStepPmicPeriod(string parameterName, string setupName, string xYShmoo = XShmoo,
            string algorithm = CharStepConst.AlgorithmNameJump)
        {
            var step = new CharStep(setupName, "PERIOD");
            step.Mode = xYShmoo;
            step.ParameterName = parameterName;
            step.ParameterType = "Global Spec";

            step.RangeCalcField = CharStepConst.RangeCalcFieldStepSize;
            step.RangeFrom = 5e6.ToString(CultureInfo.InvariantCulture);
            step.RangeTo = 28e6.ToString(CultureInfo.InvariantCulture);
            step.RangeSteps = "10";

            step.AlgorithmName = algorithm;
            step.AlgorithmArgs = "6";

            step.ApplyToPinExecMode = Simultaneous;

            step.PostStepFunction = CharStepConst.PostStepFunctionPrintShmooInfo;
            step.PostStepArgs = CorePower; //step.PostSetupArgs = CorePower;  2017/7/21 Anderson update

            step.OutputFormat = Enhanced;
            step.OutputTextFile = Disable;
            step.OutputSheet = Disable;
            step.OutputToDataLog = Enable;
            step.OutputToImmediateWin = Disable;
            step.OutputToOutputWin = Disable;
            return step;
        }

        public static void Initialize()
        {
            _instance = null;
        }

        public CharSetup Create1DPin(string setupName, string pinName, ProdCharRow prodCharRow)
        {
            var setup = new CharSetup();
            setup.SetupName = setupName;
            setup.TestMethod = CharSetupConst.TestMethodRetest;

            var pins = new List<string> { pinName };
            foreach (var pin in pins)
            {
                var step = CreatePmicPwrPinCharStep(setupName, pin, prodCharRow, XShmoo, Linear);
                setup.AddStep(step);
            }

            return setup;
        }

        public CharSetup Create2DPin(string setupName, List<string> pins, ProdCharRow prodCharRow, string pinY)
        {
            var setup = new CharSetup();
            CharStep step;
            setup.SetupName = setupName;
            setup.TestMethod = CharSetupConst.TestMethodRetest;

            foreach (var pin in pins)
            {
                step = CreatePmicPwrPinCharStep(setupName, pin, prodCharRow, XShmoo, Linear);
                setup.AddStep(step);
            }

            step = CreatePwrPinCharStepPmicPeriod(pinY, setupName, YShmoo, Linear);
            setup.AddStep(step);

            return setup;
        }

        #region Constant filed

        private const string Linear = "Linear";
        private const string Simultaneous = "Simultaneous";
        private const string Disable = "Disable";
        private const string Enable = "Enable";
        private const string CorePower = "CorePower";
        private const string VMain = "VMain";
        private const string Vps = "Vps";
        private const string XShmoo = "X Shmoo";
        private const string YShmoo = "Y Shmoo";
        private const string Enhanced = "Enhanced";
        private const double PlusRatio = 1.167;
        private const double MinusRatio = 0.6;

        private readonly Dictionary<string, string> _shmooParameterTypeDictionary;
        private readonly SpecFinder _specFinder;

        #endregion
    }
}