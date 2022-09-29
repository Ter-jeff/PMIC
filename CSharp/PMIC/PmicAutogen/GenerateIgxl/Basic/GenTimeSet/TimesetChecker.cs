using CommonLib.Enum;
using CommonLib.ErrorReport;
using IgxlData.Others.MultiTimeSet;
using PmicAutogen.Local;
using System.Collections.Generic;

namespace PmicAutogen.GenerateIgxl.Basic.GenTimeSet
{
    public class TimeSetChecker
    {
        private List<ComTimeSetBasicSheet> _comTimeSetBasicSheets = new List<ComTimeSetBasicSheet>();

        public void CheckTimeSet(List<ComTimeSetBasicSheet> comTimeSetBasicSheets)
        {
            _comTimeSetBasicSheets = comTimeSetBasicSheets;
            CheckMissingGroupName();
            CheckMultiShiftTset();
        }

        private void CheckMissingGroupName()
        {
            var pinMapSheet = TestProgram.IgxlWorkBk.PinMapPair.Value;
            if (pinMapSheet == null)
                return;
            foreach (var sheet in _comTimeSetBasicSheets)
                foreach (var timeSet in sheet.TimeSetsData)
                    foreach (var timingRow in timeSet.TimingRows)
                    {
                        var groupName = timingRow.PinGrpName;
                        if (!pinMapSheet.IsGroupExist(groupName) && !pinMapSheet.IsPinExist(groupName))
                        {
                            var outString = "Pin (group) used in Time Set file " + timeSet.Name + " missed in PinMap. -- " +
                                            groupName;
                            ErrorManager.AddError(EnumErrorType.MissingPinName, EnumErrorLevel.Error,
                                timeSet.Name, 1, outString, groupName);
                        }
                    }
        }

        private void CheckMultiShiftTset()
        {
            foreach (var igxlSheet in _comTimeSetBasicSheets)
                if (igxlSheet.IsMultiShiftInTSet)
                {
                    var alarmStr = string.Format("Multi-ShiftInFreq TimeSet Sheet problem @ {0} : {1}",
                        igxlSheet.SheetName, igxlSheet.GetMultiShiftInStr);
                    igxlSheet.InsertAlarmDataInFirstRow(alarmStr);
                    ErrorManager.AddError(EnumErrorType.FormatError, EnumErrorLevel.Error,
                        igxlSheet.SheetName, 1, alarmStr, "Multi-ShiftInFreq");
                }
        }
    }
}