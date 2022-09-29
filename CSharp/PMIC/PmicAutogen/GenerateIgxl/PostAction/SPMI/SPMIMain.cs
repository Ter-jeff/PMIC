using CommonLib.Enum;
using CommonLib.WriteMessage;
using IgxlData.IgxlBase;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Globalization;

namespace PmicAutogen.GenerateIgxl.PostAction.SPMI
{
    public class SpmiMain : MainBase
    {
        private const string SpmiSetUpName = "Shmoo_2D_SPMI";

        public void WorkFlow()
        {
            try
            {
                Response.Report("Generating SPMI Files ...", EnumMessageLevel.General, 60);

                SpmiAutoGen();

                TestProgram.IgxlWorkBk.AddIgxlSheets(IgxlSheets);

                Response.Report("SPMI Completed!", EnumMessageLevel.General, 100);
            }
            catch (Exception e)
            {
                var message = "SPMI AutoGen Failed: " + e.Message;
                Response.Report(message, EnumMessageLevel.Error, 100);
            }
        }

        private void SpmiAutoGen()
        {
            GenCharacterization();

            AddGlobalSpmiSpeed();
        }

        private void GenCharacterization()
        {
            var charSheet = TestProgram.IgxlWorkBk.GetCharSheet(PmicConst.CharSetUpPmic);
            var charSetups = GenSpmiCharSetUp();
            foreach (var charSetup in charSetups)
                if (!charSheet.CharSetups.Exists(p =>
                        p.SetupName.Equals(charSetup.SetupName, StringComparison.CurrentCultureIgnoreCase)))
                    charSheet.AddRow(charSetup);
            IgxlSheets.Add(charSheet, FolderStructure.DirDigitalScanMbist);
        }

        private List<CharSetup> GenSpmiCharSetUp()
        {
            var charSetUps = new List<CharSetup>();
            var setup = new CharSetup();
            setup.SetupName = SpmiSetUpName;
            setup.TestMethod = CharSetupConst.TestMethodRetest;
            setup.CharSteps.Add(CreateXShmooSpmiCharStep());
            setup.CharSteps.Add(CreateYShmooSpmiCharStep());
            charSetUps.Add(setup);
            return charSetUps;
        }

        private CharStep CreateXShmooSpmiCharStep()
        {
            var setup = new CharStep(SpmiSetUpName, "Axis_1");
            setup.Mode = CharStepConst.ModeXShmoo;

            setup.ParameterType = CharStepConst.ParameterTypeGlobalSpec;
            setup.ParameterName = "SPMI_Speed";

            setup.RangeCalcField = CharStepConst.RangeCalcFieldStepSize;
            setup.RangeFrom = 5e6.ToString(CultureInfo.InvariantCulture);
            setup.RangeTo = 6e7.ToString(CultureInfo.InvariantCulture);
            setup.RangeSteps = "20";

            setup.AlgorithmName = CharStepConst.AlgorithmNameLinear;

            setup.ApplyToPinExecMode = "Simultaneous";

            setup.AxisExecutionOrder = "X-Y[-Z]";

            setup.OutputFormat = "Enhanced";
            setup.SuspendDataLog = "TRUE";
            setup.OutputToTextFile = "Disable";
            setup.OutputToSheet = "Disable";
            setup.OutputToDataLog = "Enable";
            setup.OutputToImmediateWin = "Disable";
            setup.OutputToOutputWin = "Disable";

            return setup;
        }

        private CharStep CreateYShmooSpmiCharStep()
        {
            var setup = new CharStep(SpmiSetUpName, "Axis_2");
            setup.Mode = CharStepConst.ModeYShmoo;

            setup.ParameterType = CharStepConst.ParameterTypeGlobalSpec;
            setup.ParameterName = "SPMIDataStrobe";

            setup.RangeCalcField = CharStepConst.RangeCalcFieldStepSize;
            setup.RangeFrom = "0.1";
            setup.RangeTo = "0.9";
            setup.RangeSteps = "8";

            setup.AlgorithmName = CharStepConst.AlgorithmNameLinear;

            setup.ApplyToPinExecMode = "Simultaneous";
            setup.AxisExecutionOrder = "";

            return setup;
        }

        private void AddGlobalSpmiSpeed()
        {
            if (TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value == null)
                return;
            var comment = "User need to adjust the value by different project!";
            var spec = new GlobalSpec("JTAG_Speed");
            spec.Value = "=16.E+06";
            spec.Comment = comment;

            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec);
            var spec1 = new GlobalSpec("JTAG_Period");
            spec1.Value = "=1/_JTAG_Speed";
            spec1.Comment = comment;
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec1);

            var spec2 = new GlobalSpec("SPMI_Speed");
            spec2.Value = "=16.E+06";
            spec2.Comment = comment;
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec2);

            var spec3 = new GlobalSpec("SPMI_Period");
            spec3.Value = "=1/_SPMI_Speed";
            spec3.Comment = comment;
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec3);

            var spec4 = new GlobalSpec("SPMIDataStrobe");
            spec4.Value = "=750.E-03";
            spec4.Comment = comment;
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec4);

            var spec5 = new GlobalSpec("IO_Speed");
            spec5.Value = "=10.E+06";
            spec5.Comment = comment;
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec5);

            var spec6 = new GlobalSpec("IO_Period");
            spec6.Value = "=1/_IO_Speed";
            spec6.Comment = comment;
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec6);

            var spec7 = new GlobalSpec("AHB_Speed");
            spec7.Value = "=8.E+06";
            spec7.Comment = comment;
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec7);

            var spec8 = new GlobalSpec("AHB_Period");
            spec8.Value = "=1/_AHB_Speed";
            spec8.Comment = comment;
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec8);
        }
    }
}