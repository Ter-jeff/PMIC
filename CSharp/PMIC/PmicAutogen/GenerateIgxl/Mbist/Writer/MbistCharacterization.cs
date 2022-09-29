using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using PmicAutogen.Singleton;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class MbistCharacterization : ScanCharacterization
    {
        protected override void AddGlobalSpeed()
        {
            if (TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value == null)
                return;
            var spec = new GlobalSpec("Mbist_Speed");
            spec.Value = "=16.E+06";
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec);
            var spec1 = new GlobalSpec("Mbist_Period");
            spec1.Value = "=1/_Mbist_Speed";
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec1);
        }

        protected override List<CharSetup> Single1D(IEnumerable<ProdCharRow> prodCharRows)
        {
            var charSetup1DRows = new List<CharSetup>();
            foreach (var prodCharRow in prodCharRows)
            {
                var prodCharTestInstance = (ProdCharRowMbist)prodCharRow;
                if (prodCharTestInstance.PeripheralVoltage.Equals("") ||
                    prodCharTestInstance.PeripheralVoltage.Equals(Na))
                    continue;

                var pins = prodCharTestInstance.GetSinglePins(prodCharTestInstance.PeripheralVoltage);
                foreach (var pin in pins)
                {
                    var setupName = prodCharTestInstance.Get1DCharNamePeriod(pin);
                    if (!charSetup1DRows.Exists(p => p.SetupName.Equals(setupName)))
                        charSetup1DRows.Add(CharSetupSingleton.Instance()
                            .Create1DPin(setupName, pin, prodCharTestInstance));
                }
            }

            return charSetup1DRows;
        }

        protected override List<CharSetup> Single2D(IEnumerable<ProdCharRow> prodCharRows)
        {
            var charSetup2DRows = new List<CharSetup>();
            foreach (var prodCharRow in prodCharRows)
            {
                var prodCharTestInstance = (ProdCharRowMbist)prodCharRow;
                if (prodCharTestInstance.PeripheralVoltage.Equals("") ||
                    prodCharTestInstance.PeripheralVoltage.Equals(Na))
                    continue;

                var trackingGroups = prodCharTestInstance.GetTrackingGroup(prodCharTestInstance.PeripheralVoltage);
                foreach (var pins in trackingGroups)
                    foreach (var pin in pins)
                    {
                        var xShmooPins = new List<string>();
                        xShmooPins.Add(pin);
                        var setupName = prodCharTestInstance.Get2DCharNamePeriod(pin, "MBIST");
                        charSetup2DRows.Add(CharSetupSingleton.Instance()
                            .Create2DPin(setupName, xShmooPins, prodCharTestInstance, "Mbist_Speed"));
                    }
            }

            return charSetup2DRows;
        }
    }
}