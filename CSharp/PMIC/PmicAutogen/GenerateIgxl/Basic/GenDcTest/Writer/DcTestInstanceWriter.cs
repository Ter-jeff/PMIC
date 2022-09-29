using CommonLib.Extension;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.NonIgxlSheets;
using IgxlData.VBT;
using PmicAutogen.Inputs.TestPlan.Reader.DcTest.Base;
using PmicAutogen.Local;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.Basic.GenDcTest.Writer
{
    internal class DcTestInstanceWriter : DcTestWriter
    {
        public DcTestInstanceWriter(string sheetName, List<HardIpPattern> patternList) : base(sheetName, patternList)
        {
        }

        public InstanceSheet GenInsSheet(string sheetName)
        {
            var instanceSheet = new InstanceSheet(sheetName);
            instanceSheet.AddHeaderFooter();
            foreach (var hardIpPattern in PatternList)
            {
                foreach (var voltage in LabelVoltages)
                    instanceSheet.InstanceRows.Add(GenInstanceRow(hardIpPattern, voltage));
                if (LocalSpecs.HasUltraVoltageULv)
                    instanceSheet.InstanceRows.Add(GenInstanceRow(hardIpPattern, ULv));
                if (LocalSpecs.HasUltraVoltageUHv)
                    instanceSheet.InstanceRows.Add(GenInstanceRow(hardIpPattern, UHv));
            }

            return instanceSheet;
        }

        private InstanceRow GenInstanceRow(HardIpPattern hardIpPattern, string voltage)
        {
            var insRow = new InstanceRow();
            insRow.SheetName = hardIpPattern.SheetName;
            insRow.TestName = CreateTestName(hardIpPattern, voltage);
            insRow.Type = "VBT";
            var vbtFunction = TestProgram.VbtFunctionLib.GetFunctionByName(VbtFunctionLib.FunctionalTUpdated);
            insRow.Name = vbtFunction.FunctionName;
            SetArgsListValue(hardIpPattern, vbtFunction);
            insRow.ArgList = vbtFunction.Parameters;
            insRow.Args = vbtFunction.Args;
            insRow.TimeSets = CreateDcTestTimeSets(hardIpPattern);
            insRow.DcCategory = CreateDcTestDcCategory(hardIpPattern, voltage);

            insRow.DcSelector = CreateDcSelector(hardIpPattern, voltage);
            insRow.AcCategory = CreateDcTestAcCategory(hardIpPattern);
            insRow.AcSelector = string.IsNullOrEmpty(insRow.AcCategory) ? "" : CreateAcSelector(hardIpPattern, voltage);
            insRow.PinLevels = CreateDcTestPinLevel(hardIpPattern);
            return insRow;
        }

        private void SetArgsListValue(HardIpPattern pattern, VbtFunctionBase function)
        {
            function.CheckParam = false;

            #region Set value for Functional_T_Updated

            //Patterns
            function.Args[0] = pattern.Pattern.GetInstancePatternName();

            #region Default value

            //RelayMode
            function.SetParamValue("RelayMode", "1");
            function.SetParamValue("PatternTimeout", "30");

            #endregion

            if (pattern.IsNonHardIpBlock && Regex.IsMatch(pattern.Pattern.GetLastPayload(), @"_[D]SRA*M\w*DSSC",
                    RegexOptions.IgnoreCase))
                function.SetParamValue("DigSource", "Test_AutoSwitch:JTAG_TDI");
            var interposePrePat = "";

            if (pattern.ForceConditionList.Count > 0)
            {
                var forceCondition = pattern.ForceConditionList[0];
                var termInfo = "";

                #region Check multiple force type in one force condition

                foreach (var pin in forceCondition.ForcePins)
                {
                    if (!Regex.IsMatch(pin.ForceType, "TERM", RegexOptions.IgnoreCase))
                    {
                        if (pin.Type == ForceConditionType.Normal)
                        {
                            if (pin.ForceJob == "")
                                interposePrePat += pin.PinName + ":" + pin.ForceType + ":" +
                                                   ConvertForceValueToGlbSpec(pin) + ";";
                            else
                                interposePrePat += pin.PinName + ":" + pin.ForceType + ":" +
                                                   ConvertForceValueToGlbSpec(pin) + ":" + pin.ForceJob +
                                                   ";";
                        }

                        interposePrePat = interposePrePat.Trim(',').Replace(",", ";");
                    }
                    else
                    {
                        termInfo = pin.PinName + ":" + pin.ForceType + ":" +
                                   ConvertForceValueToGlbSpec(pin) + pin.ForceJob + ";";
                    }

                    if (pin.Type == ForceConditionType.Others)
                    {
                        interposePrePat += pin.PinName + ":" + pin.ForceValue + ";";
                        interposePrePat = interposePrePat.Trim(',').Replace(",", ";");
                    }
                }

                #endregion

                interposePrePat += termInfo;
            }

            function.SetParamValue("Interpose_PrePat", interposePrePat);
            function.SetParamValue("CharInputString", interposePrePat);

            #endregion
        }

        private string ConvertForceValueToGlbSpec(ForcePin forcePin)
        {
            var result = "";
            {
                var forceValue = forcePin.ForceValue;
                if (!forceValue.ToUpper().Contains("VDD") && !forceValue.ToUpper().Contains("PINS"))
                    return UnitExtensions.ConvertNumber(forceValue);
                if (forceValue.ToUpper().Contains("VDD"))
                {
                    var reg = new Regex(@"VDD\w+", RegexOptions.IgnoreCase);
                    result = reg.Replace(forceValue, m => "_" + m.Value + "_VAR");
                }
                else if (forceValue.ToUpper().Contains("PINS"))
                {
                    var reg = new Regex(@"\w*PINS\w+", RegexOptions.IgnoreCase);
                    result = reg.Replace(forceValue, m => "_" + m.Value + "_VAR");
                }
            }
            return result;
        }

        private string CreateAcSelector(HardIpPattern pattern, string voltage)
        {
            var patternName = pattern.Pattern.GetLastPayload();

            if (InputFiles.PatternListMap == null)
                return "";

            if (!InputFiles.PatternListMap.PatternListCsvRows.Exists(x =>
                    x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase))) return "";

            var acSelector = Typ;
            if (!string.IsNullOrEmpty(pattern.AcSelectorUsed))
            {
                var newAcSelector = GetAcSelector(voltage, pattern.AcSelectorUsed);
                if (!string.IsNullOrEmpty(newAcSelector))
                    acSelector = newAcSelector;
            }

            return acSelector;
        }

        private string CreateDcSelector(HardIpPattern pattern, string voltage)
        {
            string dcSelector;
            switch (voltage)
            {
                case Hv:
                case UHv:
                    dcSelector = Max;
                    break;
                case Lv:
                case ULv:
                    dcSelector = Min;
                    break;
                case Nv:
                    dcSelector = Typ;
                    break;
                default:
                    dcSelector = Typ;
                    break;
            }

            if (!string.IsNullOrEmpty(pattern.DcSelectorUsed))
            {
                var newDcSelector = GetDcSelector(voltage, pattern.DcSelectorUsed);
                if (!string.IsNullOrEmpty(newDcSelector))
                    dcSelector = newDcSelector;
            }

            return dcSelector;
        }

        private string CreateDcTestDcCategory(HardIpPattern pattern, string voltage)
        {
            var dcCategory = DcDefault;

            var dc = GetSpecifyInfo(pattern.DcCategory, "DC");
            if (!string.IsNullOrEmpty(dc))
                dcCategory = dc;


            if (voltage.Equals(UHv, StringComparison.OrdinalIgnoreCase) ||
                voltage.Equals(ULv, StringComparison.OrdinalIgnoreCase))
                dcCategory = LocalSpecs.GetUltraCategory(dcCategory);
            return dcCategory;
        }

        private string CreateDcTestAcCategory(HardIpPattern pattern)
        {
            var ac = GetSpecifyInfo(pattern.AcCategory, "AC");
            if (!string.IsNullOrEmpty(ac))
                return ac;

            return AcDefault;
        }

        private string CreateDcTestPinLevel(HardIpPattern pattern)
        {
            if (!string.IsNullOrEmpty(pattern.LevelUsed))
                return "Levels_" + pattern.LevelUsed;

            return LevelDefault;
        }

        private string CreateDcTestTimeSets(HardIpPattern pattern)
        {
            var patternName = pattern.Pattern.GetLastPayload();
            if (InputFiles.PatternListMap.PatternListCsvRows.Exists(x =>
                    x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase)))
            {
                var row = InputFiles.PatternListMap.PatternListCsvRows.Find(x =>
                    x.PatternName.Equals(patternName, StringComparison.CurrentCultureIgnoreCase));
                return row.ActualTimeSetVersion;
            }

            return string.Empty;
        }
    }
}