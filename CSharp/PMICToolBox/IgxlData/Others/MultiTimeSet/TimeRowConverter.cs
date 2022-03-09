using System.Linq;
using IgxlData.IgxlBase;

namespace IgxlData.Others.MultiTimeSet
{

    public class TimeRow1P4Converter
    {
        protected int MustHaveColumnCnt;

        public TimeRow1P4Converter()
        {
            MustHaveColumnCnt = 16;
        }

        public bool NeedCompensate(string[] dataArr)
        {
            if (dataArr.Length < MustHaveColumnCnt)
            {
                return true;
            }
            return false;
        }

        protected string[] DoCompensate(string[] dataArr)
        {

            if (!NeedCompensate(dataArr))
                return dataArr;


            var dataList = dataArr.ToList();
            while (dataList.Count < MustHaveColumnCnt)
            {
             dataList.Add("");   
            }
            return dataList.ToArray();
        }

        public virtual TimingRow ConvertTimeRow(string[] dataArr)
        {
            dataArr = DoCompensate(dataArr);

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
         
            if (dataArr.Length > 16) //comment might be lost
                row.Comment = dataArr[16];
            else
                row.Comment = "";
            return row;
        }

  
    }

    public class TimeRow2P3Converter : TimeRow1P4Converter
    {

        public TimeRow2P3Converter()
        {
            MustHaveColumnCnt = 17;
        }
        

        public override TimingRow ConvertTimeRow(string[] dataArr)
        {
            dataArr = DoCompensate(dataArr);

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
                CompareRefOffset = dataArr[15], // bug fixed wrong sequence by Jn 
                EdgeMode = dataArr[16]
            };
            if (dataArr.Length > 17) //comment might be lost
                row.Comment = dataArr[17];
            else
                row.Comment = "";
            return row;
        }


    }
}