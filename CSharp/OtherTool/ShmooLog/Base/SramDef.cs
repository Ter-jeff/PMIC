using System.Collections.Concurrent;
using System.Linq;

namespace ShmooLog.Base
{
    public class SramDef
    {
        private readonly BlockingCollection<string> _powerPinSequence = new BlockingCollection<string>();

        private ConcurrentDictionary<string, string> _sramDefMap;
        public bool HasSramDef = false;

        public BlockingCollection<SelSramCondition> SramDataSet { get; } = new BlockingCollection<SelSramCondition>();

        public void InitialMap(string sramData)
        {
            if (_sramDefMap == null)
                _sramDefMap = new ConcurrentDictionary<string, string>();

            //[SELSRM_Def,VDD_DISP:VDD_SRAM_SOC;VDD_AVE:VDD_SRAM_SOC;VDD_GPU:VDD_SRAM_GPU;VDD_ECPU:VDD_SRAM_CPU;VDD_PCPU:VDD_SRAM_CPU;VDD_DCS_DDR:VDD_SRAM_SOC;VDD_SOC:VDD_SRAM_SOC]

            var sramDef = sramData.Replace("[", "").Replace("]", "").Split(',').Last().Split(';');
            foreach (var map in sramDef)
            {
                var corePowerPin = map.Split(':').First();
                var ramPowerPin = map.Split(':').Last();
                if (_sramDefMap.ContainsKey(corePowerPin))
                    continue;
                if (!_powerPinSequence.Contains(corePowerPin))
                    _powerPinSequence.Add(corePowerPin);
                _sramDefMap[corePowerPin] = ramPowerPin;
            }
        }

        public void AddData(string data)
        {
            var selSramCond = new SelSramCondition(data);

            SramDataSet.Add(selSramCond);
        }

        public void BulitCompressStr()
        {
            foreach (var sramData in SramDataSet)
            {
                var algCompressStr = "";
                foreach (var powerPin in _powerPinSequence)
                {
                    var sramPowerSetting = sramData.GetPowerPinItem(_sramDefMap[powerPin]);
                    var corePowerSetting = sramData.GetPowerPinItem(powerPin);
                    algCompressStr += corePowerSetting.Value > sramPowerSetting.Value ? "0" : "1";
                }

                sramData.CompareCompressStr = algCompressStr;
            }
        }
    }
}