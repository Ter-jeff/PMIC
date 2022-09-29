using CommonLib.Enum;
using CommonLib.ErrorReport;
using IgxlData.IgxlBase;
using IgxlData.IgxlSheets;
using IgxlData.Others;
using PmicAutogen.GenerateIgxl.Basic.GenDc.PowerOverWrite;
using PmicAutogen.Inputs.TestPlan.Reader;
using PmicAutogen.Local;
using PmicAutogen.Local.Const;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace PmicAutogen.GenerateIgxl.Basic.GenLevel
{
    public class LevelMain
    {
        private readonly IoLevelsSheet _ioLevelsSheet;
        private readonly VddLevelsSheet _vddLevelsSheet;

        public LevelMain(VddLevelsSheet vddLevelsSheet, IoLevelsSheet ioLevelsSheet)
        {
            _vddLevelsSheet = vddLevelsSheet;
            _ioLevelsSheet = ioLevelsSheet;
        }

        public void GenLevelSheets(ref List<LevelSheet> levelSheets)
        {
            //Sort PinGroup fist,then get individual pin
            var rowList = new List<IoLevelsRow>();
            var group = _ioLevelsSheet.Rows.Where(y => !string.IsNullOrEmpty(y.Domain) && y.IsGroupPin)
                .GroupBy(x => x.Domain).Select(item => item.First()).ToList();
            rowList.AddRange(group);
            rowList.AddRange(_ioLevelsSheet.Rows.Where(x => !string.IsNullOrEmpty(x.Domain) && !x.IsGroupPin));

            //Gen Levels_Analog
            var levelSheet = levelSheets.Find(x =>
                x.SheetName.Equals(PmicConst.LevelsAnalog, StringComparison.OrdinalIgnoreCase));
            if (levelSheet == null)
            {
                levelSheet = new LevelSheet(PmicConst.LevelsAnalog);
                levelSheets.Add(levelSheet);
                levelSheet.LevelRows.AddRange(CreateIoPinLevels(rowList));
            }

            //Gen Levels_Func
            levelSheet = levelSheets.Find(x =>
                x.SheetName.Equals(PmicConst.LevelsFunc, StringComparison.OrdinalIgnoreCase));
            if (levelSheet == null)
            {
                levelSheet = new LevelSheet(PmicConst.LevelsFunc);
                levelSheet.LevelRows.AddRange(CreatePowerPinLevels());
                levelSheet.LevelRows.AddRange(CreateIoPinLevels(rowList));
                levelSheets.Add(levelSheet);
            }

            //Gen Levels_BSCAN
            levelSheet = levelSheets.Find(x =>
                x.SheetName.Equals(PmicConst.LevelsBscan, StringComparison.OrdinalIgnoreCase));
            if (levelSheet == null)
            {
                levelSheet = new LevelSheet(PmicConst.LevelsBscan);
                levelSheet.LevelRows.AddRange(CreatePowerPinLevels());
                levelSheet.LevelRows.AddRange(CreateIoPinLevels(rowList, true));
                rowList.Clear();
                var bscanGroup = _ioLevelsSheet.Rows.Where(y => !string.IsNullOrEmpty(y.Domain) && !y.IsBscanApplyPins)
                    .GroupBy(x => x.Domain).Select(item => item.First()).ToList();
                rowList.AddRange(bscanGroup);
                var noBscanGroup = _ioLevelsSheet.Rows.Where(y => !string.IsNullOrEmpty(y.Domain) && y.IsBscanApplyPins)
                    .GroupBy(x => x.Domain).Select(item => item.First()).ToList();
                rowList.AddRange(noBscanGroup);
                levelSheet.LevelRows.AddRange(CreateBscanApplyPinsPinLevels(rowList));
                levelSheets.Add(levelSheet);
            }
        }

        #region Override

        public void OverrideLevels(ref LevelSheet level, PowerOverWrite powerOverWrite)
        {
            foreach (var hardIpDcRow in powerOverWrite.DataRows)
            {
                var specValues = powerOverWrite.GetSpecValueFromDef(hardIpDcRow);
                var isPowerPin = _vddLevelsSheet.Rows.Exists(x =>
                    x.WsBumpName.Equals(hardIpDcRow.PinName, StringComparison.OrdinalIgnoreCase));

                foreach (var specValue in specValues)
                {
                    if (isPowerPin)
                        if (!(specValue.Parameter.StartsWith("Ifold", StringComparison.OrdinalIgnoreCase) ||
                              specValue.Parameter.Equals("Isc", StringComparison.OrdinalIgnoreCase)))
                            continue;

                    LevelRow levelRow = null;
                    foreach (var row in level.LevelRows)
                        if (row.PinName.Equals(specValue.PinName, StringComparison.OrdinalIgnoreCase))
                        {
                            if (row.Parameter.Equals(specValue.Parameter, StringComparison.OrdinalIgnoreCase))
                            {
                                levelRow = row;
                                break;
                            }

                            if ((specValue.Parameter.StartsWith("Ifold", StringComparison.OrdinalIgnoreCase) ||
                                 specValue.Parameter.Equals("Isc", StringComparison.OrdinalIgnoreCase)) &&
                                (row.Parameter.StartsWith("Ifold", StringComparison.OrdinalIgnoreCase) ||
                                 row.Parameter.Equals("Isc", StringComparison.OrdinalIgnoreCase)))
                            {
                                levelRow = row;
                                break;
                            }
                        }

                    if (levelRow != null)
                        levelRow.Value = specValue.Value;
                    else
                        level.AddBaseLevel(specValue);
                }
            }
        }

        #endregion

        #region Bscan Apply pins

        private List<LevelRow> CreateBscanApplyPinsPinLevels(List<IoLevelsRow> ioLevelsRows)
        {
            var levelRows = new List<LevelRow>();
            var bscanApplyPinsDomains =
                ioLevelsRows.Where(o => o.IsBscanApplyPins).Select(o => o.Domain).Distinct().ToList();
            if (!bscanApplyPinsDomains.Any())
                return levelRows;

            var domainTypeValues = StaticTestPlan.IoLevelsSheet.GetBscanDomainTypeValues();
            foreach (var bscanDomain in bscanApplyPinsDomains)
            {
                if (!domainTypeValues.ContainsKey(bscanDomain)) continue;

                var pinName = bscanDomain;
                pinName = pinName + "_BSCAN_applyPins";
                var vih = "_" + pinName + "_VIH_VAR";
                var vil = "_" + pinName + "_VIL_VAR";
                var voh = "_" + pinName + "_VOH_VAR";
                var vol = "_" + pinName + "_VOL_VAR";
                var ioh = "_" + pinName + "_IOH_VAR";
                var iol = "_" + pinName + "_IOL_VAR";
                var vt = "_" + pinName + "_VT_VAR";


                if (vih == "") vih = "0";
                if (vil == "") vil = "0";
                if (voh == "") voh = "0";
                if (vol == "") vol = "0";
                if (ioh == "") ioh = "0";
                if (iol == "") iol = "0";
                if (vt == "") vt = vih + @"/2";

                const string vohAlt1 = "0";
                const string vohAlt2 = "0";
                //const string iol = "0";
                //const string ioh = "0";
                const string vcl = "-1";
                const string vch = "5.5"; //4.56->5.5
                const string vOutLoTyp = "0";
                const string vOutHiTyp = "0";
                const string driverMode = "HiZ";
                var ioLevel = new IoLevel(pinName, vil, vih, vol, voh, vohAlt1, vohAlt2, iol, ioh, vt, vcl, vch,
                    vOutLoTyp, vOutHiTyp, driverMode);
                levelRows.AddRange(AddIoPinLevel(ioLevel));
            }

            return levelRows;
        }

        #endregion

        #region power pins

        public List<LevelRow> CreatePowerPinLevels()
        {
            var levelRows = new List<LevelRow>();
            foreach (var row in _vddLevelsSheet.Rows)
            {
                var pinName = row.WsBumpName;
                var pinType = GetPinTypeFromChannelMap(pinName);
                var dcType = "";
                var vMain = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName, dcType);
                var vAlt = SpecFormat.GenDcSpecSymbolAtLevelSheet(pinName + "_" + SpecFormat.DcSpecValtSuffix, dcType);
                var iFoldLevel = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, "Ifold");
                var tDelay = SpecFormat.GenGlbSpecSymbolAtLevelSheet(pinName, "Tdelay");

                //modified by Terry to support DC30
                if (Regex.IsMatch(pinName, "DC30", RegexOptions.IgnoreCase))
                {
                    levelRows.AddRange(AddDc30PinLevel(new Dc30Level(pinName, vMain)));
                }
                else
                {
                    //need to support DCVIMerged, added by Terry
                    if (pinType.Equals("DCVI", StringComparison.CurrentCultureIgnoreCase) ||
                        pinType.Equals("DCVIMerged", StringComparison.CurrentCultureIgnoreCase))
                        levelRows.AddRange(
                            AddDcviPowerPinLevel(new DcviPowerLevel(pinName, vMain, iFoldLevel, tDelay)));
                    else
                        levelRows.AddRange(AddPowerPinLevel(new PowerLevel(pinName, vAlt, vMain, iFoldLevel, tDelay)));
                }
            }

            return levelRows;
        }

        protected string GetPinTypeFromChannelMap(string pinName)
        {
            var pinType = "";
            if (TestProgram.IgxlWorkBk.ChannelMapSheets != null)
                if (TestProgram.IgxlWorkBk.ChannelMapSheets.SelectMany(x => x.Value.ChannelMapRows).ToList().Exists(y =>
                        y.DeviceUnderTestPinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)))
                    pinType = TestProgram.IgxlWorkBk.ChannelMapSheets.SelectMany(x => x.Value.ChannelMapRows).ToList()
                        .Find(y => y.DeviceUnderTestPinName.Equals(pinName, StringComparison.OrdinalIgnoreCase)).Type;
            return pinType;
        }

        public List<LevelRow> AddDcviPowerPinLevel(DcviPowerLevel dcviPowerLevel)
        {
            var levelRows = new List<LevelRow>();

            var lLevelRow = new LevelRow(dcviPowerLevel.PinName, "Vps", dcviPowerLevel.Vps, "");

            levelRows.Add(lLevelRow);

            lLevelRow = new LevelRow(dcviPowerLevel.PinName, "Isc", dcviPowerLevel.Isc, "");

            levelRows.Add(lLevelRow);

            lLevelRow = new LevelRow(dcviPowerLevel.PinName, "tdelay", dcviPowerLevel.Tdelay, "");

            levelRows.Add(lLevelRow);

            return levelRows;
        }

        /// <summary>
        ///     Added by Terry to support DC30 type
        /// </summary>
        /// <param name="dc30Level"></param>
        /// <returns></returns>
        public List<LevelRow> AddDc30PinLevel(Dc30Level dc30Level)
        {
            var levelRows = new List<LevelRow>();

            var lLevelRow = new LevelRow(dc30Level.PinName, "Vlevel", dc30Level.Vlevel, "");

            levelRows.Add(lLevelRow);

            return levelRows;
        }

        public List<LevelRow> AddPowerPinLevel(PowerLevel powerLevel)
        {
            var levelRows = new List<LevelRow>();

            var levelRow = new LevelRow(powerLevel.PinName, "Vmain", powerLevel.Vmain, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(powerLevel.PinName, "valt", powerLevel.Valt, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(powerLevel.PinName, "iFoldLevel", powerLevel.FoldLevel, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(powerLevel.PinName, "tdelay", powerLevel.Tdelay, "");

            levelRows.Add(levelRow);

            return levelRows;
        }

        #endregion

        #region IO pins

        private List<LevelRow> CreateIoPinLevels(List<IoLevelsRow> ioLevelsRows, bool isBscanApplyPins = false)
        {
            var levelRows = new List<LevelRow>();
            foreach (var row in ioLevelsRows)
            {
                var pinName = row.IsGroupPin ? row.Domain : row.PinName;
                var iol = "0";
                var ioh = "0";
                var vih = "_" + pinName + "_VIH_VAR";
                var vil = "_" + pinName + "_VIL_VAR";
                var voh = "_" + pinName + "_VOH_VAR";
                var vol = "_" + pinName + "_VOL_VAR";
                var vt = "_" + pinName + "_VT_VAR";

                if (vih == "") vih = "0";
                if (vil == "") vil = "0";
                if (voh == "") voh = "0";
                if (vol == "") vol = "0";
                if (vt == "") vt = vih + @"/2";

                if (isBscanApplyPins)
                {
                    iol = "_" + pinName + "_IOL_VAR";
                    ioh = "_" + pinName + "_IOH_VAR";
                    if (iol == "") iol = "0";
                    if (ioh == "") ioh = "0";
                }

                const string vohAlt1 = "0";
                const string vohAlt2 = "0";
                const string vcl = "-1";
                const string vch = "5.5"; //4.56->5.5
                const string vOutLoTyp = "0";
                const string vOutHiTyp = "0";
                const string driverMode = "HiZ";
                var ioLevel = new IoLevel(pinName, vil, vih, vol, voh, vohAlt1, vohAlt2, iol, ioh, vt, vcl, vch,
                    vOutLoTyp, vOutHiTyp, driverMode);
                levelRows.AddRange(AddIoPinLevel(ioLevel));
            }

            return levelRows;
        }

        public List<LevelRow> AddIoPinLevel(IoLevel ioLevel)
        {
            var levelRows = new List<LevelRow>();


            var levelRow = new LevelRow(ioLevel.PinName, "vil", ioLevel.Vil, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "vih", ioLevel.Vih, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "vol", ioLevel.Vol, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "voh", ioLevel.Voh, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "voh_alt1", ioLevel.VohAlt1, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "voh_alt2", ioLevel.VohAlt2, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "iol", ioLevel.Iol, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "ioh", ioLevel.Ioh, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "vt", ioLevel.Vt, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "vcl", ioLevel.Vcl, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "vch", ioLevel.Vch, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "voutLoTyp", ioLevel.VoutLoTyp, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "voutHiTyp", ioLevel.VoutHiTyp, "");

            levelRows.Add(levelRow);

            levelRow = new LevelRow(ioLevel.PinName, "driverMode", ioLevel.DriverMode, "");

            levelRows.Add(levelRow);

            return levelRows;
        }

        #endregion

        #region pin group

        public void GenPinGroup(IoLevelsSheet ioLevels)
        {
            var pinMap = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMap == null) return;

            var pinGroups = CreatePinGroup(ioLevels);
            foreach (var pinGroup in pinGroups)
            {
                if (!pinGroup.PinList.Any())
                    continue;

                if (pinMap.IsGroupExist(pinGroup.PinName))
                {
                    var outString = "The pin group " + pinGroup.PinName + " is already existed !!!";
                    var rowNum = ioLevels.Rows
                        .Find(x => x.Domain.Equals(pinGroup.PinName, StringComparison.OrdinalIgnoreCase)).RowNum;
                    ErrorManager.AddError(EnumErrorType.Duplicate, EnumErrorLevel.Error, ioLevels.SheetName,
                        rowNum, ioLevels.TypeIndex, outString, pinGroup.PinName);
                }
                else
                {
                    pinMap.AddRow(pinGroup);
                }
            }
        }

        private List<PinGroup> CreatePinGroup(IoLevelsSheet ioLevels)
        {
            var groups = new List<PinGroup>();
            var group = ioLevels.Rows.Where(y => y.IsGroupPin).GroupBy(x => x.Domain).ToList();
            foreach (var item in group)
            {
                var rows = item.ToList();
                if (string.IsNullOrEmpty(rows[0].Domain)) continue;
                var pinGroup = new PinGroup(rows[0].Domain);
                foreach (var row in rows)
                {
                    var newPin = new Pin(row.PinName, PinMapConst.TypeIo, "IO sheet");
                    pinGroup.AddPin(newPin);
                }

                groups.Add(pinGroup);
            }

            var bscanApplyPinGroups = ioLevels.GenBscanApplyPinGroups();
            if (bscanApplyPinGroups.Any())
                groups.AddRange(bscanApplyPinGroups);
            return groups;
        }

        #endregion
    }
}