using System.Collections.Generic;
using AutomationCommon.EpplusErrorReport;
using IgxlData.Others.MultiTimeSet;
using PmicAutogen.Local;

namespace PmicAutogen.GenerateIgxl.Basic.Writer.GenTimeSet
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


                    EpplusErrorManager.AddError(BasicErrorType.MissingPinName.ToString(), ErrorLevel.Error,
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
                    EpplusErrorManager.AddError(BasicErrorType.FormatError.ToString(), ErrorLevel.Error,
                        igxlSheet.SheetName, 1, alarmStr, "Multi-ShiftInFreq");
                }
        }
    }
}