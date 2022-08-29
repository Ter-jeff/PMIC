namespace IgxlData.Others.PatternListCsvFile
{
    public class OriPatListItem
    {
        #region Field

        private string _timeset;
        #endregion

        public int RowNum { get; set; }
        public string Idx { get; set; } //Both
        public string Pattern{ get; set; }  //  Both
        public string LatestVersion{ get; set; }    //Only TW
        public string ReleaseDate{ get; set; }      //If not exist TW will added
        public string UseNoUse{ get; set; }         //Both
        public string DRi { get; set; }                 //If not exist TW will added
        public string ReleaseNote { get; set; }         //If not exist TW will added
        public string RadarNum { get; set; }             //If not exist TW will added
        public string Org { get; set; }                  //If not exist TW will added
        public string TypeSpec { get; set; }             //If not exist TW will added
       
        public string TimesetLatest         //Ori pattern List
        {
            get { return _timeset; }
            set { _timeset = value; }
        }

        public string TimesetVersion           //TW Pattern List
        {
            get { return _timeset; }
            set { _timeset = value; }
        }  

        public string FileVersions{ get; set; }         //Both
        public string OpCode{ get; set; }           //Added in TW
        public string ScanMode { get; set; }         //Added in TW
        public string Halt { get; set; }             //Added in TW
        public string Compilation { get; set; }          //Added in TW
        public string HLv { get; set; }                  //Added in TW
        public string OriTimeMod { get; set; }               //Added in TW
        public string Checkrt { get; set; }                  //Added in TW
        public string TpCategory { get; set; }               //Added in TW

        public OriPatListItem()
        {
            RowNum = 1;
            Idx = "";
            Pattern = "";
            LatestVersion = "";
            ReleaseDate = "";
            UseNoUse = "";
            DRi = "";
            ReleaseNote = "";
            RadarNum = "";
            Org = "";
            TypeSpec = "";
            TimesetLatest = "";
            FileVersions = "";
            OpCode = "";
            ScanMode = "";
            Halt = "";
            Compilation = "";
            HLv = "";
            OriTimeMod = "";
            Checkrt = "";
            TpCategory = "";
        }
    }

}