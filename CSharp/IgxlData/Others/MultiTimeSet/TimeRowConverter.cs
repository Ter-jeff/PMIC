using IgxlData.IgxlBase;

namespace IgxlData.Others.MultiTimeSet
{
    public interface ITimeRowConverter
    {
        TimingRow ConvertTimeRow(string[] dataArr);
    }

    public class TimeRow1P4 : ITimeRowConverter
    {
        public TimingRow ConvertTimeRow(string[] dataArr)
        {

            var row = new TimingRow
            {
                PinGrpName = dataArr[3],
                PinGrpClockPeriod = dataArr[4],
                PinGrpSetup = dataArr[5],
                DataSrc = dataArr[6],
                DataFmt = dataArr[7],
                DriveOn = dataArr[8],
                DriveData = dataArr[9],
                DriveReturn = dataArr[10],
                DriveOff = dataArr[11],
                CompareMode = dataArr[12],
                CompareOpen = dataArr[13],
                CompareClose = dataArr[14],
                EdgeMode = dataArr[15]
            };
         
            row.Comment = dataArr.Length > 16 ? dataArr[16] : "";
            return row;
        }
    }

    public class TimeRow2P3 : ITimeRowConverter
    {
        public TimingRow ConvertTimeRow(string[] dataArr)
        {
            if (dataArr.Length < 18)
            {
            }
            var row = new TimingRow
            {
                PinGrpName = dataArr[3],
                PinGrpClockPeriod = dataArr[4],
                PinGrpSetup = dataArr[5],
                DataSrc = dataArr[6],
                DataFmt = dataArr[7],
                DriveOn = dataArr[8],
                DriveData = dataArr[9],
                DriveReturn = dataArr[10],
                DriveOff = dataArr[11],
                CompareMode = dataArr[12],
                CompareOpen = dataArr[13],
                CompareClose = dataArr[14],
                EdgeMode = dataArr[15],
                CompareRefOffset = dataArr[16]
            };
            row.Comment = dataArr.Length > 17 ? dataArr[17] : "";
            return row;
        }
    }
}