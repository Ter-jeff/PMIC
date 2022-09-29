namespace ShmooLog.Base
{
    public class Device //Parallel Test是以Device為單位, 且我用ConcurrentDictionary需要以物件來儲存第一手資訊 隨後再轉為DataTable
    {
        //從結果得到的資訊
        public int Bin = 0;
        //Datalog 含 Site 的統計資訊 -> Device -> Test Instance -> Test Number 依狀況判斷是以哪個為基礎去整理Data!!
        //                                                                     -> Test Instance -> Debug Mode
        //                                                                     -> Shmoo


        //身分一

        public int DeviceNo; //可當作檢索條件 不能重複
        public int ExecutedTest = -1; //<-- 目前發現有的Log沒有
        public int FailedTest = -1; //<-- 目前發現有的Log沒有
        public int Site;
        public int Sort = 0;

        public int X = -999; //可當作檢索條件
        public int Y = -999; //可當作檢索條件

        //要儲存Test Instance/Number的Sequence!!
        //public List<string> ListSeqTestInstance = new List<string>(); //照順序收集到的Test Instance, 關鍵Key

        public Device()
        {
        }

        public Device(int siteNum, int deviceNum)
        {
            Site = siteNum;
            DeviceNo = deviceNum;
        }

        public string SiteDevice
        {
            get
            {
                return string.Format("Site{0}:Device{1}", DeviceNo, Site);
                //只有Site+DeviceNumber都一樣才算重複
            }
        }

        public string DieXY
        {
            get { return string.Format("{0},{1}", X, Y); }
        }
    }
}