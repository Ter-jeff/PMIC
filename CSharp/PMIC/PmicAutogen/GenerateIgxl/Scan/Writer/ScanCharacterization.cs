using IgxlData.IgxlBase;
using PmicAutogen.Inputs.ScghFile.ProChar.Base;
using PmicAutogen.Local;
using PmicAutogen.Singleton;
using System.Collections.Generic;
using System.Linq;

namespace PmicAutogen.GenerateIgxl.Scan.Writer
{
    public class ScanCharacterization
    {
        protected const string Na = "N/A";

        public List<CharSetup> WorkFlow(IEnumerable<ProdCharRow> prodCharRows)
        {
            var charSetups = new List<CharSetup>();

            var charRows = prodCharRows.ToList();

            charSetups.AddRange(Single1D(charRows)); //1D Single

            charSetups.AddRange(Single2D(charRows)); //2D Power

            AddGlobalSpeed();

            return charSetups;
        }

        protected virtual void AddGlobalSpeed()
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

        protected virtual List<CharSetup> Single1D(IEnumerable<ProdCharRow> prodCharRows)
        {
            var charSetup1DRows = new List<CharSetup>();
            foreach (var prodCharRow in prodCharRows)
            {
                var prodCharTestInstance = (ProdCharRowScan)prodCharRow;
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

        protected virtual List<CharSetup> Single2D(IEnumerable<ProdCharRow> prodCharRows)
        {
            var charSetup2DRows = new List<CharSetup>();
            foreach (var prodCharRow in prodCharRows)
            {
                var prodCharTestInstance = (ProdCharRowScan)prodCharRow;
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