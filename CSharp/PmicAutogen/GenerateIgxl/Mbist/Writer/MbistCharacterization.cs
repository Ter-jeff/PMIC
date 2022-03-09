using System.Collections.Generic;
using IgxlData.IgxlBase;
using PmicAutogen.GenerateIgxl.Scan.Writer;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using PmicAutogen.Singleton;

namespace PmicAutogen.GenerateIgxl.Mbist.Writer
{
    public class MbistCharacterization : ScanCharacterization
    {
        private const string Na = "N/A";

        public List<CharSetup> WorkFlow(List<ProdCharRowMbist> prodCharRowMbists)
        {
            var charSetups = new List<CharSetup>();

            charSetups.AddRange(MbistCharFlow1DSinglePmic(prodCharRowMbists)); //1D Single

            charSetups.AddRange(MbistCharFlow2DPowerPmic(prodCharRowMbists)); //2D Power

            AddGlobalScanSpeed();

            return charSetups;
        }

        private void AddGlobalScanSpeed()
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

        private List<CharSetup> MbistCharFlow1DSinglePmic(List<ProdCharRowMbist> instanceList)
        {
            var charSetup1DRows = new List<CharSetup>();
            foreach (var prodCharTestInstance in instanceList)
            {
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

        private List<CharSetup> MbistCharFlow2DPowerPmic(List<ProdCharRowMbist> instanceList)
        {
            var charSetup2DRows = new List<CharSetup>();
            foreach (var prodCharTestInstance in instanceList)
            {
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