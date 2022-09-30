using System.Collections.Generic;

namespace ShmooLog.Base
{
    public class ShmooId //只有純粹的數值 被ShmooSetup引用
    {
        public bool Abnormal;
        public double Abnormal_Hvcc;

        public double Abnormal_Lvcc; //應永良要求 專門給Shmoo_Hole page用的
        public string DieXY = "NONE";

        public Dictionary<int, char> FailFlagIndexDict;
        public double Hvcc;

        //新版要記得檢查抓到的內容有沒有符合格式!!!!  <--------------------------------
        public bool IsAllFailed; //外面留選項 如果All Fail看要不要畫

        public bool IsMergeByDeviceId;
        public string LotId = "NONE";

        public double Lvcc; //當Step沒辦法被整除 或者 Step Size到小數點三位的時候 []內的Low / High會對不起來


        public int MergeInstanceCnt;
        public double PassRate;

        public bool ShmooAlarm;

        public double ShmooAlarmValue = -7777;


        public List<string> ShmooContent = new List<string>(); //或者用char[][]? Dictionary?
        public List<string> ShmooContentHVCC = new List<string>(); //    [.....NH,0.565,1.200]]

        // in shmoo log  LVCC ,HVCC  20161209 by JN
        public List<string> ShmooContentLVCC = new List<string>(); //    [.....NH,0.565,1.200]]

        public string ShmooHole = "NH"; //1D 用 2D再看看
        public bool ShmooHoleInOperationRange; //1D 用 2D再看看

        public string ShmooInstanceName = "N/A";

        public string ShmooSetupUniqueName = "N/A";
        public string Site = "0"; //反正只是列印用 不用宣告int
        public string Sort = "0"; //反正只是列印用 不用宣告int
        public string SourceFileName = "N/A";

        public string GetIdUniqleName
        {
            get { return Site + "-" + LotId + "-" + DieXY; }
        }
    }
}