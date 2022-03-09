using System.Collections.Generic;
using IgxlData.IgxlBase;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using PmicAutogen.Singleton;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanCharacterization
    {
        private const string Na = "N/A";

        public List<CharSetup> WorkFlow(List<ProdCharRowScan> prodCharRowScans)
        {
            var charSetups = new List<CharSetup>();

            charSetups.AddRange(ScanCharFlow1DSinglePmic(prodCharRowScans)); //1D Single

            charSetups.AddRange(ScanCharFlow2DPowerPmic(prodCharRowScans)); //2D Power

            AddGlobalScanSpeed();

            return charSetups;
        }

        private void AddGlobalScanSpeed()
        {
            if (TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value == null)
                return;
            var spec = new GlobalSpec("SCAN_Speed");
            spec.Value = "=16.E+06";
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec);
            var spec1 = new GlobalSpec("SCAN_Period");
            spec1.Value = "=1/_SCAN_Speed";
            TestProgram.IgxlWorkBk.GlbSpecSheetPair.Value.AddRow(spec1);
        }

        private List<CharSetup> ScanCharFlow1DSinglePmic(List<ProdCharRowScan> prodCharRowScans)
        {
            var charSetup1DRows = new List<CharSetup>();
            foreach (var prodCharTestInstance in prodCharRowScans)
            {
                if (prodCharTestInstance.SupplyVoltage.Equals("") || prodCharTestInstance.SupplyVoltage.Equals(Na))
                    continue;

                var pins = prodCharTestInstance.GetSinglePins(prodCharTestInstance.SupplyVoltage);
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

        private List<CharSetup> ScanCharFlow2DPowerPmic(List<ProdCharRowScan> prodCharRowScans)
        {
            var charSetup2DRows = new List<CharSetup>();
            foreach (var prodCharTestInstance in prodCharRowScans)
            {
                if (prodCharTestInstance.SupplyVoltage.Equals("") || prodCharTestInstance.SupplyVoltage.Equals(Na))
                    continue;

                var trackingGroups = prodCharTestInstance.GetTrackingGroup(prodCharTestInstance.SupplyVoltage);
                foreach (var pins in trackingGroups)
                foreach (var pin in pins)
                {
                    var xShmooPins = new List<string>();
                    xShmooPins.Add(pin);
                    var setupName = prodCharTestInstance.Get2DCharNamePeriod(pin, "SCAN");
                    if (!charSetup2DRows.Exists(p => p.SetupName.Equals(setupName)))
                        charSetup2DRows.Add(CharSetupSingleton.Instance()
                            .Create2DPin(setupName, xShmooPins, prodCharTestInstance, "SCAN_Speed"));
                }
            }

            return charSetup2DRows;
        }
    }
}