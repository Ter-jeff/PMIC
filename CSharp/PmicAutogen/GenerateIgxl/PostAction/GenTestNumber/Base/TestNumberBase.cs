namespace PmicAutogen.GenerateIgxl.PostAction.GenTestNumber.Base
{
    public class TestNumberBase
    {
        public TestNumberBase()
        {
            StartNum = 0;
            Interval = 100;
            MaxNum = 999999999;
        }

        public long StartNum { set; get; }
        public long Interval { set; get; }
        public long MaxNum { set; get; }
    }
}