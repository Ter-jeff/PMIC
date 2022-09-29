using CommonLib.Extension;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ShmooLog.Base
{
    public class ShmooSets //包含Shmoo Setup 及相應的資料處理
    {
        public const string TwoDimsionMergeStr = "2DMerged";

        private readonly Regex _regexFreqDector = new Regex(@"_(?<Freq>(shiftin)*\d+MHZ[a-zA-Z0-9]*)_",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public bool AlignHighLowVcc = false;

        // BinCut Excel
        public BinCutDataForShmoo BinCutRefData;


        public bool BypassAllFailed = false; //外面留選項 如果All Fail看要不要畫
        public string CatchSubString = "Full";

        public int CatchTiTill = 1; //如何切Ti決定Sheet Name(Category)
        public DataSet CurrShmooReport = new DataSet(); //有缺陷 這種寫法只會保留最後一筆

        public Dictionary<string, List<ShmooSetup>> DicCategoryShmooSetups = new Dictionary<string, List<ShmooSetup>>();
        //整合通過各種Filter集合在一起的ShmooSetup(來源不一) 以及 Shmoo Options(Spec ...)

        public DataSet DsShmooSets = new DataSet();

        public List<string> FailedShmooMsg = new List<string>();
        public bool FindShmooHole1D = true;

        public bool GRRLoopByPassErrCount = false;
        public double HgbRatio = 1.1;

        public string HgbRatioString = "+10%";

        public bool InstacneFreqMode = false;
        public bool JoinCategory = false;
        public double LgbRatio = 0.9;
        public string LgbRatioString = "-10%";

        // 根據Filter的條件 決定 Category -> 每個Table清一次 
        // 感想: 不用Dictionary沒辦法排序, 用List又需要搜尋 浪費時間
        public List<string> ListAllCategories = new List<string>(); //已經轉成全部大寫
        public bool LVCCHVCC2DRReport = false;

        public bool Merge2D2Inst = false;
        public bool Merge2DSameSetUp = false;
        public bool Merge2DSameSetupBySite = false;
        public ColorSetting MergeColorSetting = new ColorSetting();
        public bool OnlyOverlay2DShmoo = false;

        public bool OrderByTestNum = true;
        public bool OverlayBySite = false;

        public bool PlotShmoo1D = true;
        public bool PlotShmoo2D = true;

        public bool PlotShmooOverlay = true;

        public List<string> SelectedShmooItems = new List<string>(); ////Test Instance :: Setup Name 如果不是空的 就是只轉指定的

        public string ShmooLogName;
        public bool ShowPercBySite = false;


        public List<double> SpecRatio = new List<double>(); //會乘以+-

        public int SplitFilePerDevice = 0;

        // SelSram Condition
        public List<SelSramCondition> SramDataSet = new List<SelSramCondition>();
        public int UserDefLessErr5555Count = 1;
        public int UserDefLessErr9999Count = 1;

        public void ParseShmooSetupIdPerTable(int tableNo) //循序處理
        {
            //var args = new OrangeXl.ProgressStatus() { Percentage = 0 }; //To Report Progress
            //progress.Report(args);

            var rgexShmPass = new Regex(@"\+|\*|P|p", RegexOptions.Compiled);
            var rgexShmFail = new Regex(@"\-|\~|F|f", RegexOptions.Compiled);


            //分類每一個Shmoo Setup & Id
            var dt = DsShmooSets.Tables[tableNo];

            ListAllCategories.Clear();
            DicCategoryShmooSetups.Clear();
            FailedShmooMsg.Clear();

            try
            {
                if (Merge2D2Inst)
                {
                    Instance2DMergeSelectorMode();
                }
                else
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        var shmooString = (string)row["Shmoo Setup"];
                        var shmooContent = (string)row["Shmoo Content"];
                        var shmooType = (string)row["Shmoo Type"];
                        var iArray = Regex.Split(shmooString.Replace(@"[", "").Replace("]", ""), @",");

                        var ti = ((string)row["Test Instance"]).ToUpper();
                        var setupName = ((string)row["Setup Name"]).ToUpper();

                        var uniqueName = ti + "::" + setupName; //已經是大寫

                        #region 刪掉不符合條-件的Shmoo

                        if (SelectedShmooItems.Count > 0 && !SelectedShmooItems.Contains(uniqueName))
                            continue; //如果User有指定Shmoo Item
                        if (shmooType == "1D" && !PlotShmoo1D) continue;
                        if (shmooType == "2D" && !PlotShmoo2D) continue;

                        #endregion

                        var category = "";
                        if (JoinCategory)
                            category = string.Join("_", ti.Split('_').Take(CatchTiTill)); //看設定怎麼分類的 抓到哪個位置
                        else
                            category = ti.Split('_')[CatchTiTill - 1];

                        if (!CatchSubString.Equals("Full"))
                        {
                            int subIndex = Convert.ToInt16(CatchSubString);
                            category = category.Substring(0, subIndex);
                        }


                        if (!ListAllCategories.Contains(category)) //已經轉成全部大寫
                        {
                            ListAllCategories.Add(category);
                            DicCategoryShmooSetups[category] = new List<ShmooSetup>();
                        }

                        //ShmooSetup shmSetup;
                        var shmSetup =
                            DicCategoryShmooSetups[category].FirstOrDefault(q => q.UniqeName.Equals(uniqueName));
                        if (shmSetup == null)
                        {
                            shmSetup = InitialShmooSetup(row, category, uniqueName, ti, iArray);
                            DicCategoryShmooSetups[category].Add(shmSetup); //掛上去
                        }

                        //shmSetup.InstanceList.Add(ti);
                        shmSetup.ForceCondition = (string)row["ForceCondition"];
                        shmSetup.TimeSetInfo = (string)row["Activity TimeSet"];

                        var shmooContantInSetup = iArray[iArray.Length - 4];


                        var shmoostring = "";
                        // high to low need reverse
                        if (shmSetup.AcurateStep[0][0] > shmSetup.AcurateStep[0][1])
                            shmoostring = string.Join("", shmooContantInSetup.Reverse());
                        else
                            shmoostring = shmooContantInSetup;

                        var FailFlagIndexDict = new Dictionary<int, char>();
                        var index = 0;
                        foreach (var c in shmoostring)
                        {
                            var failFlag = 'F';

                            if (c.Equals('B'))
                                failFlag = 'B';
                            else if (c.Equals('S'))
                                failFlag = 'S';
                            else if (c.Equals('C')) failFlag = 'C';
                            FailFlagIndexDict.Add(index, failFlag);
                            index++;
                        }


                        //開始處理 Shmoo ID
                        var shmooId = new ShmooId
                        {
                            Site = shmooType == "1D" ? iArray[5] : iArray[1],
                            LotId = shmooType == "1D" ? iArray[1] : iArray[3],
                            DieXY = shmooType == "1D" ? iArray[2] + "," + iArray[3] : iArray[4] + "," + iArray[5],
                            Sort = (string)row["Sort"],
                            SourceFileName = (string)row["File Name"],
                            ShmooSetupUniqueName = shmSetup.UniqeName,

                            FailFlagIndexDict = FailFlagIndexDict
                        };

                        //Pass                +
                        //Fail                -
                        //Assumed Pass        *
                        //Assumed Fail        ~
                        //Programmed Pass     P
                        //Programmed Fail     F
                        //Assumed Prog Pass   p
                        //Assumed Prog Fail   f
                        //Alarm               A
                        //Error               E
                        //N/A                 .
                        //No Test             _

                        if (shmSetup.Type == "1D")
                        {
                            //如果是1D 直接Add, 2D要先切割好
                            shmooId.ShmooContent.Add(shmooContent); //那些++--

                            if (shmooContent.Length != shmSetup.StepCount[0] * shmSetup.StepCount[1])
                                FailedShmooMsg.Add(string.Format("{0} has wrong Shmoo Content at {1}",
                                    shmooId.ShmooSetupUniqueName, shmooId.DieXY));

                            //shmSetup.StepCount[0]
                            var passed = shmooContent.Count(c => Regex.IsMatch(c.ToString(), @"\+|\*|P|p"));

                            //全過或全Fail都算Abnormal!!
                            if (passed == shmooContent.Length)
                            {
                                shmooId.Abnormal = true;
                            }

                            else if (passed == 0)
                            {
                                shmooId.Abnormal = true;
                                shmooId.IsAllFailed = true;
                            }

                            if (shmooContent.Length > 0)
                                shmooId.PassRate = Math.Round(passed / shmooContent.Length * 100.0, 2,
                                    MidpointRounding.AwayFromZero);

                            //Chihome 算好的 Shmoo Hole
                            shmooId.ShmooHole = iArray[iArray.Length - 3] == string.Empty
                                ? @"N/A"
                                : iArray[iArray.Length - 3];

                            //Chihome 算好的 Hvcc Lvcc
                            if (iArray[iArray.Length - 1] == string.Empty || iArray[iArray.Length - 1] == @"N/A")
                                shmooId.Hvcc = 0;
                            else
                                shmooId.Hvcc = Convert.ToDouble(iArray[iArray.Length - 1]);
                            //Math.Round(Convert.ToDouble(iArray[iArray.Length - 1]), 4, MidpointRounding.AwayFromZero); // Mod for PC

                            if (iArray[iArray.Length - 2] == string.Empty || iArray[iArray.Length - 2] == @"N/A")
                                shmooId.Lvcc = 0;
                            else
                                shmooId.Lvcc = Convert.ToDouble(iArray[iArray.Length - 2]);
                            //Math.Round(Convert.ToDouble(iArray[iArray.Length - 2]), 4, MidpointRounding.AwayFromZero); // Mod for PC


                            //  判斷shmoo Alarm record alram index
                            //  -~~~~~-~~~~~-~~~~~-~~~~~-~~~~~-~~~~~-~~~~~-~~~~~-~~~~~-~~~~~-~~~~~-~~~~~-~~~~~A~-
                            var alarmindex = 0;
                            foreach (var c in shmooContent)
                            {
                                var isalarm = Regex.IsMatch(c.ToString(), @"A");
                                if (isalarm)
                                    shmooId.ShmooAlarmValue =
                                        Convert.ToDouble(
                                            shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()]
                                                [alarmindex]);
                                shmooId.ShmooAlarm = shmooId.ShmooAlarm || isalarm;
                                alarmindex++;
                            }


                            //加保險 Check LVCC HVCC的點 有沒有出現在座標軸上 否則畫圖表會出問題
                            //if (shmSetup.NeedAlignHighLowVcc[0] || AlignHighLowVcc)
                            if (AlignHighLowVcc)
                            {
                                var leftPoint = 0;
                                var rightPoint = 0;
                                shmooContent = rgexShmPass.Replace(shmooContent, "+");

                                leftPoint = shmooContent.IndexOf('+');
                                rightPoint = shmooContent.LastIndexOf('+');

                                if (rightPoint >= 0)
                                    shmooId.Hvcc =
                                        Convert.ToDouble(shmSetup.AcurateStep[0][2] > 0
                                            ? shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()][
                                                rightPoint]
                                            : shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()][leftPoint
                                            ]);

                                if (leftPoint >= 0)
                                    shmooId.Lvcc =
                                        Convert.ToDouble(shmSetup.AcurateStep[0][2] > 0
                                            ? shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()][leftPoint
                                            ]
                                            : shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()][
                                                rightPoint]);
                            }

                            if (shmooId.ShmooHole != "NH")
                            {
                                var leftPointFail = 0;
                                var rightPointFail = 0;
                                var leftPointPass = 0;
                                var rightPointPass = 0;

                                shmooContent = rgexShmFail.Replace(shmooContent, "-");
                                shmooContent = rgexShmPass.Replace(shmooContent, "+");

                                leftPointPass = shmooContent.IndexOf('+');
                                rightPointPass = shmooContent.LastIndexOf('+');
                                leftPointFail = shmooContent.IndexOf('-');
                                rightPointFail = shmooContent.LastIndexOf('-');

                                if (!shmooId.IsAllFailed)
                                {
                                    if (shmSetup.AcurateStep[0][2] > 0) //遞增型態
                                        shmooId.Abnormal_Hvcc =
                                            Convert.ToDouble(
                                                shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()][
                                                    rightPointPass]);
                                    else
                                        shmooId.Abnormal_Hvcc =
                                            Convert.ToDouble(
                                                shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()][
                                                    leftPointPass]);

                                    if (shmSetup.AcurateStep[0][2] > 0) //遞增型態
                                    {
                                        if (rightPointFail != -1 &&
                                            rightPointFail + 1 <
                                            shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()].Count)
                                            shmooId.Abnormal_Lvcc =
                                                Convert.ToDouble(
                                                    shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()][
                                                        rightPointFail + 1]);
                                    }
                                    else
                                    {
                                        if (leftPointFail != -1 && leftPointFail - 1 >= 0)
                                            shmooId.Abnormal_Lvcc =
                                                Convert.ToDouble(
                                                    shmSetup.DicAxisValue.XAxis.PinValueSet[shmSetup.Axis[0].Last()][
                                                        leftPointFail - 1]);
                                    }
                                }

                                if (shmSetup.SpecPoints[0].Count > 0 && SpecRatio.Count > 0)
                                {
                                    var tStr = new List<string>();
                                    var t = shmooContent.ToCharArray();
                                    for (var i = shmSetup.SpecPoints[0].Min(); i <= shmSetup.SpecPoints[0].Max(); i++)
                                        tStr.Add(t[i].ToString());
                                    shmooId.ShmooHoleInOperationRange = FindShmooHoleIn1D(tStr);
                                }
                            }
                        }
                        else
                        {
                            var shmooContainLVCC = (string)row["LVCC Value"];
                            var shmooContainHVCC = (string)row["HVCC Value"];

                            //如果是1D 直接Add, 2D要先切割好
                            shmooId.ShmooContent.AddRange(shmooContent.Split(',')); //那些++--
                            shmooId.ShmooContentHVCC.AddRange(shmooContainHVCC.Split(','));
                            shmooId.ShmooContentLVCC.AddRange(shmooContainLVCC.Split(','));


                            var pureContent = Regex.Replace(shmooContent, @",", "");

                            if (pureContent.Length != shmSetup.StepCount[0] * shmSetup.StepCount[1])
                                FailedShmooMsg.Add(string.Format("{0} has wrong Shmoo Content at {1}",
                                    shmooId.ShmooSetupUniqueName, shmooId.DieXY));

                            var passed = pureContent.Count(c => Regex.IsMatch(c.ToString(), @"\+|\*|P|p"));

                            if (passed == shmSetup.StepCount[0] * shmSetup.StepCount[1])
                            {
                                shmooId.Abnormal = true;
                            }
                            else if (passed == 0)
                            {
                                shmooId.Abnormal = true;
                                shmooId.IsAllFailed = true;
                            }

                            if (shmSetup.StepCount[0] * shmSetup.StepCount[1] > 0)
                                shmooId.PassRate = Math.Round(
                                    passed / (shmSetup.StepCount[0] * shmSetup.StepCount[1]) * 100.0, 2,
                                    MidpointRounding.AwayFromZero);
                        }

                        if (shmooId.IsAllFailed && BypassAllFailed) continue; //All Failed的濾掉

                        shmSetup.ShmooIDs.Add(shmooId); //一個Setup掛一串 ID -> Content
                    }


                    if (Merge2DSameSetUp)
                    {
                        foreach (DataRow row in dt.Rows)
                        {
                            var shmooString = (string)row["Shmoo Setup"];
                            var shmooContent = (string)row["Shmoo Content"];
                            var shmooType = (string)row["Shmoo Type"];
                            var iArray = Regex.Split(shmooString.Replace(@"[", "").Replace("]", ""), @",");

                            var ti = ((string)row["Test Instance"]).ToUpper();
                            var setupName = ((string)row["Setup Name"]).ToUpper();

                            var uniqueName = ti + "::" + setupName; //已經是大寫

                            #region 刪掉不符合條-件的Shmoo

                            if (SelectedShmooItems.Count > 0 && !SelectedShmooItems.Contains(uniqueName))
                                continue; //如果User有指定Shmoo Item
                            if (shmooType == "1D") continue;
                            if (shmooType == "2D" && !PlotShmoo2D) continue;

                            #endregion

                            var category = TwoDimsionMergeStr;

                            if (!ListAllCategories.Contains(category)) //已經轉成全部大寫
                            {
                                ListAllCategories.Add(category);
                                DicCategoryShmooSetups[category] = new List<ShmooSetup>();
                            }

                            //ShmooSetup shmSetup;
                            var shmSetup =
                                DicCategoryShmooSetups[category].FirstOrDefault(q => q.UniqeName.Equals(setupName));
                            if (shmSetup == null)
                            {
                                shmSetup = InitialShmooSetup(row, category, setupName, "MergedItem", iArray, true);
                                DicCategoryShmooSetups[category].Add(shmSetup); //掛上去
                            }

                            shmSetup.InstanceList.Add(ti);
                            shmSetup.ForceCondition = (string)row["ForceCondition"];
                            shmSetup.TimeSetInfo = (string)row["Activity TimeSet"];

                            var shmooContantInSetup = iArray[iArray.Length - 4];


                            var shmoostring = "";
                            // high to low need reverse
                            if (shmSetup.AcurateStep[0][0] > shmSetup.AcurateStep[0][1])
                                shmoostring = string.Join("", shmooContantInSetup.Reverse());
                            else
                                shmoostring = shmooContantInSetup;

                            var failFlagIndexDict = new Dictionary<int, char>();
                            var index = 0;
                            foreach (var c in shmoostring)
                            {
                                var failFlag = 'F';

                                if (c.Equals('B'))
                                    failFlag = 'B';
                                else if (c.Equals('S'))
                                    failFlag = 'S';
                                else if (c.Equals('C')) failFlag = 'C';
                                failFlagIndexDict.Add(index, failFlag);
                                index++;
                            }


                            //開始處理 Shmoo ID
                            var shmooId = new ShmooId
                            {
                                Site = shmooType == "1D" ? iArray[5] : iArray[1],
                                LotId = shmooType == "1D" ? iArray[1] : iArray[3],
                                DieXY =
                                    shmooType == "1D" ? iArray[2] + "," + iArray[3] : iArray[4] + "," + iArray[5],
                                Sort = (string)row["Sort"],
                                SourceFileName = (string)row["File Name"],
                                ShmooSetupUniqueName = uniqueName,
                                ShmooInstanceName = ti,
                                FailFlagIndexDict = failFlagIndexDict
                            };

                            //Pass                +
                            //Fail                -
                            //Assumed Pass        *
                            //Assumed Fail        ~
                            //Programmed Pass     P
                            //Programmed Fail     F
                            //Assumed Prog Pass   p
                            //Assumed Prog Fail   f
                            //Alarm               A
                            //Error               E
                            //N/A                 .
                            //No Test             _


                            var shmooContainLvcc = (string)row["LVCC Value"];
                            var shmooContainHvcc = (string)row["HVCC Value"];

                            //如果是1D 直接Add, 2D要先切割好
                            shmooId.ShmooContent.AddRange(shmooContent.Split(',')); //那些++--
                            shmooId.ShmooContentHVCC.AddRange(shmooContainHvcc.Split(','));
                            shmooId.ShmooContentLVCC.AddRange(shmooContainLvcc.Split(','));


                            var pureContent = Regex.Replace(shmooContent, @",", "");

                            if (pureContent.Length != shmSetup.StepCount[0] * shmSetup.StepCount[1])
                                FailedShmooMsg.Add(string.Format("{0} has wrong Shmoo Content at {1}",
                                    shmooId.ShmooSetupUniqueName, shmooId.DieXY));

                            var passed = pureContent.Count(c => Regex.IsMatch(c.ToString(), @"\+|\*|P|p"));

                            if (passed == shmSetup.StepCount[0] * shmSetup.StepCount[1])
                            {
                                shmooId.Abnormal = true;
                            }
                            else if (passed == 0)
                            {
                                shmooId.Abnormal = true;
                                shmooId.IsAllFailed = true;
                            }

                            if (shmSetup.StepCount[0] * shmSetup.StepCount[1] > 0)
                                shmooId.PassRate =
                                    Math.Round(passed / (shmSetup.StepCount[0] * shmSetup.StepCount[1]) * 100.0, 2,
                                        MidpointRounding.AwayFromZero);

                            if (shmooId.IsAllFailed && BypassAllFailed) continue; //All Failed的濾掉

                            shmSetup.ShmooIDs.Add(shmooId); //一個Setup掛一串 ID -> Content
                        }


                        if (Merge2DSameSetupBySite && DicCategoryShmooSetups.ContainsKey(TwoDimsionMergeStr))
                        {
                            var twoDimsionMergeShmooSetup = DicCategoryShmooSetups[TwoDimsionMergeStr];

                            var category = "2DMergeByDie";

                            ListAllCategories.Add(category);
                            DicCategoryShmooSetups.Add(category, new List<ShmooSetup>());


                            foreach (var sameSetup in twoDimsionMergeShmooSetup)
                            {
                                var sameSetupItems = new List<ShmooSetup>();

                                foreach (var categoryItem in DicCategoryShmooSetups)
                                    sameSetupItems.AddRange(categoryItem.Value.Where(shmSetup =>
                                        !shmSetup.IsSameSetupMerge && shmSetup.SetupName.Equals(sameSetup.SetupName)));

                                DicCategoryShmooSetups[category].AddRange(sameSetupItems);

                                var cloneSetUp = sameSetup.CloneJson();
                                cloneSetUp.ShmooIDs.Clear();

                                cloneSetUp.IsMergeBySiteSetup = true;

                                DicCategoryShmooSetups[category].Add(cloneSetUp);


                                var idInfoHashSet = new HashSet<string>();
                                foreach (var shmId in sameSetup.ShmooIDs) idInfoHashSet.Add(shmId.GetIdUniqleName);

                                foreach (var uniqleIdName in idInfoHashSet)
                                {
                                    var sameIdshm =
                                        sameSetup.ShmooIDs.FindAll(i => i.GetIdUniqleName.Equals(uniqleIdName))
                                            .ToList();

                                    var dicOverlay2D =
                                        new Dictionary<int,
                                            Dictionary<int,
                                                int>>(); //處理疊圖 Step X Y -> Pass Count                                  
                                    for (var x = 0; x < sameSetup.StepCount[0]; x++)
                                    {
                                        dicOverlay2D[x] = new Dictionary<int, int>();

                                        for (var y = 0; y < sameSetup.StepCount[1]; y++) dicOverlay2D[x][y] = 0; //初始化
                                    }

                                    foreach (var shmooId in sameIdshm) //收集Die Info, Content, 統計Lvcc
                                        for (var y = 0; y < shmooId.ShmooContent.Count; y++)
                                            foreach (Match m in rgexShmPass.Matches(shmooId.ShmooContent[y])) //重要技巧!
                                                dicOverlay2D[m.Index][y]++;

                                    var shmooIdByDieMerge = new ShmooId
                                    {
                                        LotId = sameIdshm.First().LotId,
                                        DieXY = sameIdshm.First().DieXY,
                                        Site = sameIdshm.First().Site,
                                        IsMergeByDeviceId = true,
                                        ShmooSetupUniqueName = sameSetup.UniqeName,
                                        MergeInstanceCnt = sameIdshm.Count
                                    };

                                    cloneSetUp.ShmooIDs.Add(shmooIdByDieMerge);

                                    var rebuildOverlayStr2D = new List<string>(); //Rebuild Overlay 2D

                                    for (var y = 0; y < sameSetup.StepCount[1]; y++)
                                    {
                                        var tmpStr = "";
                                        for (var x = 0; x < sameSetup.StepCount[0]; x++)
                                            tmpStr += string.Format("{0:F0}",
                                                dicOverlay2D[x][y] * 100 / sameIdshm.Count) + "|";
                                        rebuildOverlayStr2D.Add(tmpStr.Remove(tmpStr.Length - 1));
                                    }

                                    shmooIdByDieMerge.ShmooContent.AddRange(rebuildOverlayStr2D);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }


            //最後做總體檢 因為有可能Setup有 但是裡面的ID因為種種原因被濾掉 那就不行
            foreach (var category in ListAllCategories) //每個Setup最少要有一個ID
            {
                var setups = DicCategoryShmooSetups[category];
                var toBeRemovedSetup = setups.Where(setup => setup.ShmooIDs.Count == 0).ToList();
                foreach (var shmooSetup in toBeRemovedSetup) setups.Remove(shmooSetup); //這樣行嗎?
            }

            //每個Category最少要有一個Setup
            var toBeRemovedCat = DicCategoryShmooSetups.Keys.Where(cat => DicCategoryShmooSetups[cat].Count == 0)
                .ToList();
            foreach (var cat in toBeRemovedCat)
            {
                DicCategoryShmooSetups.Remove(cat);
                ListAllCategories.Remove(cat);
            }

            if (InstacneFreqMode)
            {
                var lOrderDicCategoryShmooSetups = new Dictionary<string, List<ShmooSetup>>();
                foreach (var cateShmItem in DicCategoryShmooSetups)
                {
                    var lSortDic = new Dictionary<string, List<ShmooSetup>>();
                    foreach (var shmItem in cateShmItem.Value)
                    {
                        var lInstance = shmItem.TestInstanceName;
                        var match = _regexFreqDector.Match(lInstance);
                        if (match.Success)
                        {
                            var lFreq = match.Groups["Freq"].ToString();
                            lInstance = lInstance.Replace("_" + lFreq, "");
                        }

                        if (!lSortDic.ContainsKey(lInstance))
                            lSortDic.Add(lInstance, new List<ShmooSetup>());
                        lSortDic[lInstance].Add(shmItem);
                    }

                    foreach (var sortItem in lSortDic)
                    {
                        if (!lOrderDicCategoryShmooSetups.ContainsKey(cateShmItem.Key))
                            lOrderDicCategoryShmooSetups[cateShmItem.Key] = new List<ShmooSetup>();
                        lOrderDicCategoryShmooSetups[cateShmItem.Key]
                            .AddRange(sortItem.Value.OrderBy(p => p.TestInstanceName).ToList());
                    }
                }

                DicCategoryShmooSetups = null;
                DicCategoryShmooSetups = lOrderDicCategoryShmooSetups;
            }

            //Report Progress Here
            //args.Percentage = 100;
            //args.Result = "Current Process: Done~";
            //progress.Report(args);
        }

        private void Instance2DMergeSelectorMode()
        {
            var mergeIndex = 0;
            foreach (DataTable dtMerge in DsShmooSets.Tables)
            {
                var instanceIndex = 0;
                foreach (DataRow row in dtMerge.Rows)
                {
                    var shmooString = (string)row["Shmoo Setup"];
                    var shmooContent = (string)row["Shmoo Content"];
                    var shmooType = (string)row["Shmoo Type"];
                    var iArray = Regex.Split(shmooString.Replace(@"[", "").Replace("]", ""), @",");

                    var ti = ((string)row["Test Instance"]).ToUpper();
                    instanceIndex++;
                    var setupName = ((string)row["Setup Name"]).ToUpper();

                    var uniqueName = ti + "::" + setupName; //已經是大寫

                    #region 刪掉不符合條-件的Shmoo

                    if (SelectedShmooItems.Count > 0 && !SelectedShmooItems.Contains(uniqueName))
                        continue; //如果User有指定Shmoo Item
                    if (shmooType == "1D") continue;
                    if (shmooType == "2D" && !PlotShmoo2D) continue;

                    #endregion

                    var category = "2DMerged" + "_" + mergeIndex;

                    if (!ListAllCategories.Contains(category)) //已經轉成全部大寫
                    {
                        ListAllCategories.Add(category);
                        DicCategoryShmooSetups[category] = new List<ShmooSetup>();
                    }

                    //ShmooSetup shmSetup;
                    var shmSetup =
                        DicCategoryShmooSetups[category].FirstOrDefault(q => q.UniqeName.Equals(setupName));
                    if (shmSetup == null)
                    {
                        shmSetup = InitialShmooSetup(row, category, setupName, "MergedItem", iArray, true);
                        DicCategoryShmooSetups[category].Add(shmSetup); //掛上去
                    }


                    var shmooPatternStr = string.Join(" ", GetShmooPatternFromShmSetup(iArray));

                    shmSetup.PatternList.Add(shmooPatternStr);

                    shmSetup.InstanceList.Add("Inst" + instanceIndex + ": " + ti);

                    shmSetup.ForceCondition = (string)row["ForceCondition"];
                    shmSetup.TimeSetInfo = (string)row["Activity TimeSet"];

                    var shmooContantInSetup = iArray[iArray.Length - 4];


                    var shmoostring = "";
                    // high to low need reverse
                    if (shmSetup.AcurateStep[0][0] > shmSetup.AcurateStep[0][1])
                        shmoostring = string.Join("", shmooContantInSetup.Reverse());
                    else
                        shmoostring = shmooContantInSetup;

                    var FailFlagIndexDict = new Dictionary<int, char>();
                    var index = 0;
                    foreach (var c in shmoostring)
                    {
                        var failFlag = 'F';

                        if (c.Equals('B'))
                            failFlag = 'B';
                        else if (c.Equals('S'))
                            failFlag = 'S';
                        else if (c.Equals('C')) failFlag = 'C';
                        FailFlagIndexDict.Add(index, failFlag);
                        index++;
                    }


                    //開始處理 Shmoo ID
                    var shmooId = new ShmooId
                    {
                        Site = shmooType == "1D" ? iArray[5] : iArray[1],
                        LotId = shmooType == "1D" ? iArray[1] : iArray[3],
                        DieXY =
                            shmooType == "1D" ? iArray[2] + "," + iArray[3] : iArray[4] + "," + iArray[5],
                        Sort = (string)row["Sort"],
                        SourceFileName = (string)row["File Name"],
                        ShmooSetupUniqueName = uniqueName,
                        ShmooInstanceName = ti,
                        FailFlagIndexDict = FailFlagIndexDict
                    };

                    //Pass                +
                    //Fail                -
                    //Assumed Pass        *
                    //Assumed Fail        ~
                    //Programmed Pass     P
                    //Programmed Fail     F
                    //Assumed Prog Pass   p
                    //Assumed Prog Fail   f
                    //Alarm               A
                    //Error               E
                    //N/A                 .
                    //No Test             _


                    var shmooContainLVCC = (string)row["LVCC Value"];
                    var shmooContainHVCC = (string)row["HVCC Value"];

                    //如果是1D 直接Add, 2D要先切割好
                    shmooId.ShmooContent.AddRange(shmooContent.Split(',')); //那些++--
                    shmooId.ShmooContentHVCC.AddRange(shmooContainHVCC.Split(','));
                    shmooId.ShmooContentLVCC.AddRange(shmooContainLVCC.Split(','));


                    var pureContent = Regex.Replace(shmooContent, @",", "");

                    if (pureContent.Length != shmSetup.StepCount[0] * shmSetup.StepCount[1])
                        FailedShmooMsg.Add(string.Format("{0} has wrong Shmoo Content at {1}",
                            shmooId.ShmooSetupUniqueName, shmooId.DieXY));

                    var passed = pureContent.Count(c => Regex.IsMatch(c.ToString(), @"\+|\*|P|p"));

                    if (passed == shmSetup.StepCount[0] * shmSetup.StepCount[1])
                    {
                        shmooId.Abnormal = true;
                    }
                    else if (passed == 0)
                    {
                        shmooId.Abnormal = true;
                        shmooId.IsAllFailed = true;
                    }

                    if (shmSetup.StepCount[0] * shmSetup.StepCount[1] > 0)
                        shmooId.PassRate = Math.Round(
                            passed / (shmSetup.StepCount[0] * shmSetup.StepCount[1]) * 100.0, 2,
                            MidpointRounding.AwayFromZero);


                    if (shmooId.IsAllFailed && BypassAllFailed) continue; //All Failed的濾掉

                    shmSetup.ShmooIDs.Add(shmooId); //一個Setup掛一串 ID -> Content
                }

                mergeIndex++;
            }
        }

        private List<string> GetShmooPatternFromShmSetup(string[] iArray)
        {
            var patternSet = new HashSet<string>();

            foreach (var s in iArray)
                if (Regex.IsMatch(s, @"\\pattern|\.pat|\.gz", RegexOptions.IgnoreCase))
                {
                    var pat = Regex.Replace(s, @".+\\|\..+|\s+", "");
                    pat = pat.ToUpper() + ".PAT";
                    patternSet.Add(pat);
                }

            return patternSet.ToList();
        }

        private ShmooSetup InitialShmooSetup(DataRow row, string category, string uniqeName, string testinstance,
            IList<string> shmooArray, bool isMerge = false)
        {
            var ti = ((string)row["Test Instance"]).ToUpper();
            var shmooContent = (string)row["Shmoo Content"];
            var setupName = ((string)row["Setup Name"]).ToUpper();


            var shmSetup = new ShmooSetup
            {
                IsSameSetupMerge = isMerge,
                Type = (string)row["Shmoo Type"],
                TestInstanceName = testinstance,
                SetupName = setupName,
                UniqeName = uniqeName, //已經是大寫,
                Category = category
            };

            if (!isMerge)
                shmSetup.TestNum = Convert.ToInt32(shmSetup.Type == "1D" ? shmooArray[9] : shmooArray[8]);
            //Int16 : -32768 ~ +32768 不夠


            shmSetup.AxisType[0] = (string)row["X Axis Unit"]; //由Shmoo Content算出來的  AxisType[0] x axis name
            //這是從DataLog抓的 如果從Setup String判斷的話 不見得有VDD或FREQ字樣 如 XIO
            shmSetup.AxisType[1] = (string)row["Y Axis Unit"]; //由Shmoo Content算出來的  AxisType[1] y axis name

            //由Content直接得到Step Count
            if (shmSetup.Type == "1D")
            {
                shmSetup.StepCount[0] = shmooContent.Length; //1D的不用處理
                if (shmooContent.Length == 1) shmSetup.Special = "RET"; //Special Case: Retention
            }
            else
            {
                var tmpContent = shmooContent.Split(','); //----+++++,---++++++,-++++++
                shmSetup.StepCount[0] = tmpContent[0].Length; //由Shmoo Content算出來的  StepCount[0] x axis count 
                shmSetup.StepCount[1] = tmpContent.Length; //由Shmoo Content算出來的  StepCount[1] y axis count 
            }


            foreach (var s in shmooArray)
            {
                if (!isMerge)
                {
                    if (Regex.IsMatch(s, @"\\pattern|\.pat|\.gz", RegexOptions.IgnoreCase))
                    {
                        var pat = Regex.Replace(s, @".+\\|\..+|\s+", "");
                        pat = pat.ToUpper() + ".PAT";

                        if (!shmSetup.PatternList.Contains(pat))
                            shmSetup.PatternList.Add(pat);
                        else
                            for (var rt = 1; rt < 99; rt++)
                            {
                                var pat_retest = pat + " (Retest-" + rt + ")";
                                if (!shmSetup.PatternList.Contains(pat_retest))
                                {
                                    shmSetup.PatternList.Add(pat_retest);
                                    rt = 99;
                                }
                            }

                        if (pat.Split('_').Count() > 3)
                            if (!Regex.IsMatch(pat.Split('_')[3], @"^IN"))
                                shmSetup.PayloadList.Add(pat);

                        continue;
                    }

                    // Rtos Cmd mode 20180227 , m9 only use cmd to access rtos
                    if (Regex.IsMatch(s, @"(bbq|pmgr mode|sc run)(\w|\s)+;", RegexOptions.IgnoreCase))
                    {
                        var cmd = s.ToUpper().Split(';').ToList();
                        shmSetup.PatternList.AddRange(cmd);
                        continue;
                    }
                }

                if (Regex.IsMatch(s, @".+=.+:.+:.+")) //Setup可能有很多組Tracking, StepSize不可信 但開頭結尾要遵守
                {
                    var axisAry = Regex.Split(s, @"\@|\=|\:", RegexOptions.IgnoreCase);

                    if (axisAry.Length == 5) //2D X軸 X@SocShift_Freq=10000000.000:100000000.000:10000000.000
                    {
                        var labelAxis = axisAry[1];
                        var startPoint = Convert.ToDouble(axisAry[2]);
                        var endPoint = Convert.ToDouble(axisAry[3]);

                        if (axisAry[0] == "X")
                        {
                            var currStepSize = (endPoint - startPoint) / (shmSetup.StepCount[0] - 1);

                            shmSetup.AcurateStep[0][0] = startPoint;
                            shmSetup.AcurateStep[0][1] = endPoint;
                            shmSetup.AcurateStep[0][2] = currStepSize; //Real Step Size!!

                            shmSetup.Axis[0].Add(labelAxis); //依序記錄遇到的X
                            shmSetup.SettingsX.Add(axisAry[1] + "=" + axisAry[2] + ":" + axisAry[3] + ":" +
                                                   axisAry[4]);

                            if (!shmSetup.DicAxisValue.XAxis.PinValueSet.ContainsKey(labelAxis))
                            {
                                shmSetup.DicAxisValue.XAxis.PinValueSet[labelAxis] = new List<double>();
                                if (Regex.IsMatch(labelAxis, @"Freq", RegexOptions.IgnoreCase))
                                {
                                    startPoint = startPoint / 1000000.0;
                                    endPoint = endPoint / 1000000.0;
                                    currStepSize = currStepSize / 1000000.0;
                                }

                                for (var i = 0; i < shmSetup.StepCount[0]; i++)
                                    shmSetup.DicAxisValue.XAxis.PinValueSet[labelAxis].Add(Math.Round(
                                        startPoint + currStepSize * i, 3,
                                        MidpointRounding.AwayFromZero));
                            }


                            //var originalStepSize = currStepSize;
                            //currStepSize = Math.Round(currStepSize, 3, MidpointRounding.AwayFromZero);

                            ////Check if Step Size是不是可以整除
                            //shmSetup.NeedAlignHighLowVcc[0] = Math.Abs(originalStepSize - currStepSize) >
                            //                                  0.0000000001;
                        }
                        else //Y軸
                        {
                            var currStepSize = (endPoint - startPoint) / (shmSetup.StepCount[1] - 1);

                            shmSetup.AcurateStep[1][0] = startPoint;
                            shmSetup.AcurateStep[1][1] = endPoint;
                            shmSetup.AcurateStep[1][2] = currStepSize; //Real Step Size!!

                            shmSetup.Axis[1].Add(labelAxis); //依序記錄遇到的Y
                            shmSetup.SettingsY.Add(axisAry[1] + "=" + axisAry[2] + ":" + axisAry[3] + ":" +
                                                   axisAry[4]);
                            //if (shmSetup.Axis[1].Count == 0) shmSetup.Special = "INVERSE";

                            if (!shmSetup.DicAxisValue.YAxis.PinValueSet.ContainsKey(labelAxis))
                                shmSetup.DicAxisValue.YAxis.PinValueSet[labelAxis] = new List<double>();

                            if (Regex.IsMatch(labelAxis, @"Freq", RegexOptions.IgnoreCase))
                            {
                                startPoint = startPoint / 1000000.0;
                                endPoint = endPoint / 1000000.0;
                                currStepSize = currStepSize / 1000000.0;
                            }

                            for (var i = 0; i < shmSetup.StepCount[1]; i++)
                                shmSetup.DicAxisValue.YAxis.PinValueSet[labelAxis].Add(Math.Round(
                                    startPoint + currStepSize * i, 3,
                                    MidpointRounding.AwayFromZero));

                            //var originalStepSize = currStepSize;
                            //currStepSize = Math.Round(currStepSize, 3, MidpointRounding.AwayFromZero);

                            ////Check if Step Size是不是可以整除
                            //shmSetup.NeedAlignHighLowVcc[1] = Math.Abs(originalStepSize - currStepSize) >
                            //                                  0.0000000001;
                        }
                    }
                    else //1D 的 X XI0_Freq=20000000.000:32000000.000:200000.000
                    {
                        var labelAxis = axisAry[0];
                        var startPoint = Convert.ToDouble(axisAry[1]);
                        var endPoint = Convert.ToDouble(axisAry[2]);
                        var currStepSize = (endPoint - startPoint) / (shmSetup.StepCount[0] - 1);

                        shmSetup.AcurateStep[0][0] = startPoint;
                        shmSetup.AcurateStep[0][1] = endPoint;
                        shmSetup.AcurateStep[0][2] = currStepSize; //Real Step Size!!

                        //有特例 同樣的Power竟然重複出現 且 Range不同(但Step Size一樣)
                        if (shmSetup.Axis[0].Contains(labelAxis))
                        {
                            var find = shmSetup.Axis[0].Count(l => Regex.IsMatch(l, labelAxis));
                            labelAxis = string.Format("{0}_{1}", labelAxis, find);
                        }

                        shmSetup.Axis[0].Add(labelAxis); //依序記錄遇到的X

                        shmSetup.SettingsX.Add(axisAry[0] + "=" + axisAry[1] + ":" + axisAry[2] + ":" +
                                               axisAry[3]);

                        if (!shmSetup.DicAxisValue.XAxis.PinValueSet.ContainsKey(labelAxis))
                            shmSetup.DicAxisValue.XAxis.PinValueSet[labelAxis] = new List<double>();

                        if (Regex.IsMatch(labelAxis, @"Freq", RegexOptions.IgnoreCase))
                        {
                            startPoint = startPoint / 1000000.0;
                            endPoint = endPoint / 1000000.0;
                            currStepSize = currStepSize / 1000000.0;
                        }

                        for (var i = 0; i < shmSetup.StepCount[0]; i++)
                            shmSetup.DicAxisValue.XAxis.PinValueSet[labelAxis]
                                .Add(startPoint +
                                     currStepSize *
                                     i); // Math.Round(startPoint + currStepSize * i, 3,MidpointRounding.AwayFromZero)

                        //var originalStepSize = currStepSize;
                        //currStepSize = Math.Round(currStepSize, 3, MidpointRounding.AwayFromZero);

                        ////Check if Step Size是不是可以整除
                        //shmSetup.NeedAlignHighLowVcc[0] = Math.Abs(originalStepSize - currStepSize) >
                        //                                  0.0000000001;
                    }

                    continue;
                }

                if (!Regex.IsMatch(s, @"=")) continue;


                var matchFreeRun = Regex.Match(s, @"(?<FreeRun>XI0|XO0)", RegexOptions.IgnoreCase);
                if (matchFreeRun.Success && !Regex.IsMatch(s, "^VDD", RegexOptions.IgnoreCase))
                {
                    var freeRun = matchFreeRun.Groups["FreeRun"].ToString(); //XI0 or XO0  20180402
                    //var mhz = Convert.ToDouble(Regex.Split(s, @"=")[1])/1000000;
                    //shmSetup.FreeRunningClk = "XI0 = " + mhz.ToString() + " Mhz";
                    //考慮改Try Parse <-- 要改
                    //XI0_DIFF=30000000;RT_CLK32768=32768
                    if (s.Split(';').Count() > 1) //XI0_0_DIFF=24000000;XI0_1_DIFF=24000000;XI0_2_DIFF=24000000
                    {
                        var s_first = s.Split(';')[0];
                        var outD = 0.0;
                        foreach (var freeRunStr in s.Split(';'))
                            if (double.TryParse(freeRunStr.Split('=')[1], out outD))
                            {
                                if (outD > 1000000.0)
                                    shmSetup.FreeRunningClk +=
                                        string.Format(freeRunStr.Split('=')[0] + " = {0:0.00} MHz ,",
                                            outD / 1000000.0); //寫死 因為Shmoo Output是錯的
                                else
                                    shmSetup.FreeRunningClk +=
                                        string.Format(freeRunStr.Split('=')[0] + " = {0} Hz ,",
                                            outD); //寫死 因為Shmoo Output是錯的
                            }
                            else
                            {
                                shmSetup.FreeRunningClk = string.Format(freeRunStr.Split('=')[0] + " = -0 MHz ,");
                            }
                    }
                    else
                    {
                        var outD = 0.0;
                        if (double.TryParse(s.Split('=')[1], out outD))
                            shmSetup.FreeRunningClk =
                                string.Format(s.Split('=')[0] + " = {0:0.00} MHz ,",
                                    outD / 1000000.0); //寫死 因為Shmoo Output是錯的
                        else shmSetup.FreeRunningClk = string.Format(s.Split('=')[0] + " = -0 MHz ,");
                    }
                }
                else
                {
                    // VDD_CPU=0.56
                    var outD = 0.0;
                    if (double.TryParse(s.Split('=')[1], out outD)) shmSetup.SettingsInit.Add(s);
                }
            } //End of iArray

            if (shmSetup.PatternList.Count == 0) shmSetup.PatternList.Add(@"N/A");

            //尋找最靠右邊的Setup 才不會選到Tracking的
            var firstX = Regex.Split(shmSetup.Axis[0][shmSetup.Axis[0].Count - 1], @":")[0];
            var origialX = firstX;
            var hadConvertX = false;
            // cyprus sepcial converter 
            if (firstX.ToUpper() == "VDD_FIXED_GROUP")
            {
                firstX = "VDD_FIXED";
                hadConvertX = true;
            }

            else if (firstX.ToUpper() == "VDD_LOW_GROUP")
            {
                firstX = "VDD_LOW";
                hadConvertX = true;
            }
            else if (firstX.ToUpper() == "VDD_CPU_GROUP")
            {
                firstX = "VDD_CPU";
                hadConvertX = true;
            }

            //"VDD_SOC:1.33,1.32,...,0.5"
            var firstY = "N/A";
            if (shmSetup.Type == "2D") firstY = Regex.Split(shmSetup.Axis[1][shmSetup.Axis[1].Count - 1], @":")[0];

            //尋找Spec 
            foreach (var init in shmSetup.SettingsInit)
            {
                // mod 2017 03 21 by JN
                // init = XO0_0=24000000 XO0_1=24000000 XO0_2=24000000 split('=') will have bug in here

                var initSets = init.Split();
                var PinValueDict = new Dictionary<string, string>();
                foreach (var set in initSets)
                {
                    var PinValue = set.Trim().Split('='); //"VDD_SOC=1.33
                    PinValueDict[PinValue[0]] = PinValue[1];
                }

                var findXKey = PinValueDict.Keys.ToList()
                    .Find(key => Regex.IsMatch(key, firstX + @"\b", RegexOptions.IgnoreCase));
                //var initSet = init.Split('='); //"VDD_SOC=1.33
                if (Regex.IsMatch(firstX, @"XI0|XO0", RegexOptions.IgnoreCase))
                    shmSetup.Spec[0] = Convert.ToDouble(shmSetup.FreeRunningClk.Split(' ')[2]) * 1000000.0;
                //else if (Regex.IsMatch(initSet[0], firstX + @"\b", RegexOptions.IgnoreCase))
                //    shmSetup.Spec[0] = Convert.ToDouble(initSet[1]);
                else if (findXKey != null) shmSetup.Spec[0] = Convert.ToDouble(PinValueDict[findXKey]);


                if (shmSetup.Type == "2D")
                {
                    var findYKey = PinValueDict.Keys.ToList()
                        .Find(key => Regex.IsMatch(key, firstY + @"\b", RegexOptions.IgnoreCase));
                    if (Regex.IsMatch(firstY, @"XI0|XO0", RegexOptions.IgnoreCase))
                    {
                        if (!string.IsNullOrEmpty(shmSetup.FreeRunningClk))
                            shmSetup.Spec[1] = Convert.ToDouble(shmSetup.FreeRunningClk.Split(' ')[2]) * 1000000.0;
                    }
                    //else if (Regex.IsMatch(initSet[0], firstY + @"\b", RegexOptions.IgnoreCase))
                    //    shmSetup.Spec[1] = Convert.ToDouble(initSet[1]);
                    else if (findYKey != null)
                    {
                        shmSetup.Spec[1] = Convert.ToDouble(PinValueDict[findYKey]);
                    }
                }
            }

            if (SpecRatio.Count > 0)
                //尋找Spec的Point 看哪一個點最靠近就可以了
                for (var i = 0; i < SpecRatio.Count; i++)
                {
                    var currSpec = SpecRatio[i] * Convert.ToDouble(shmSetup.Spec[0]);

                    var currDistance = 999999.0;
                    var minStep = 0;
                    if (shmSetup.Spec[0] > 0) // shmSetup.Spec[0] x axis spec
                    {
                        if (hadConvertX) firstX = origialX;
                        for (var step = 0; step < shmSetup.DicAxisValue.XAxis.PinValueSet[firstX].Count; step++)
                        {
                            var tmpDis =
                                Math.Abs(shmSetup.DicAxisValue.XAxis.PinValueSet[firstX][step] -
                                         currSpec); //  與Vmain值差距
                            if (tmpDis <= currDistance)
                            {
                                minStep = step;
                                currDistance = tmpDis;
                            }
                        }

                        shmSetup.SpecPoints[0].Add(minStep);
                    }

                    if (shmSetup.Type == "2D" && shmSetup.Spec[1] > 0)
                    {
                        currSpec = SpecRatio[i] * Convert.ToDouble(shmSetup.Spec[1]);
                        currDistance = 999999.0;
                        minStep = 0;
                        var findPonit = false;
                        for (var step = 0; step < shmSetup.DicAxisValue.YAxis.PinValueSet[firstY].Count; step++)
                        {
                            var tmpDis = Math.Abs(shmSetup.DicAxisValue.YAxis.PinValueSet[firstY][step] - currSpec);
                            if (tmpDis <= currDistance)
                            {
                                minStep = step;
                                currDistance = tmpDis;
                                findPonit = true;
                            }
                        }

                        if (minStep != 0 || findPonit)
                            shmSetup.SpecPoints[1].Add(minStep);
                    }
                }


            #region ShmooBinCutInformation

            if (BinCutRefData != null)
            {
                // CPUTDF_VDDECPU_Vmain_vs_Shiftin_Freq_Jump
                // VDDECPU
                var binCutDomain = setupName.Split('_')[1];
                // DFTLH_ECPU_ECPU_MC701_CPUTDF_C0F3_PL00_COM_LPC23_50MHz_CZ_NV_CPUTDF_VDDECPU
                // MC701
                var tiPMode = testinstance.Split('_')[3];

                var domain = ShmooBinCutPModeBusiness.DomainGetting(binCutDomain);
                var domainPmodeRow =
                    ShmooBinCutPModeBusiness.GetTargerBinnedDominItem(BinCutRefData, domain, ref tiPMode);

                if (domainPmodeRow != null)
                {
                    shmSetup.PlotBinCutSpec = true;
                    shmSetup.BinCutPerfromanceName = tiPMode;
                    if (hadConvertX) firstX = origialX;
                    if (!domainPmodeRow.IsOtherRail)
                    {
                        var cpMax = domainPmodeRow.CPVmax / 1000; //mV to V
                        shmSetup.BinCutSpecPoints[0] = FindBinCutSpecPoint(shmSetup, firstX, cpMax, true);
                        shmSetup.BinCutSpecValue[0] = cpMax;

                        var cpMin = domainPmodeRow.CPVmin / 1000; //mV to V
                        shmSetup.BinCutSpecPoints[1] = FindBinCutSpecPoint(shmSetup, firstX, cpMin, false);
                        shmSetup.BinCutSpecValue[1] = cpMin;
                    }
                    else // other rail only one spec
                    {
                        shmSetup.IsOtherRailSpec = true;
                        var cp = domainPmodeRow.CPVmax / 1000; //mV to V
                        shmSetup.BinCutSpecPoints[0] = FindBinCutSpecPoint(shmSetup, firstX, cp, true);
                        shmSetup.BinCutSpecValue[0] = cp;
                    }
                }
            }

            #endregion

            return shmSetup;
        }

        private static int FindBinCutSpecPoint(ShmooSetup shmSetup, string firstX, double cpValue, bool isMax)
        {
            var specPoint = isMax ? -1 : shmSetup.DicAxisValue.XAxis.PinValueSet[firstX].Count;
            for (var step = 0; step < shmSetup.DicAxisValue.XAxis.PinValueSet[firstX].Count; step++)
            {
                var tmpDis = Math.Abs(shmSetup.DicAxisValue.XAxis.PinValueSet[firstX][step]);

                if (tmpDis > cpValue)
                {
                    specPoint = step;
                    break;
                }
            }

            return specPoint;
        }

        public void CookCurrentShmooReport() //只針對當前處理的Tablet產生Report!!
        {
            CurrShmooReport.Tables.Clear(); //先清空前面的記錄

            var rgexShmPass = new Regex(@"\+|\*|P|p", RegexOptions.Compiled);
            var rgexShmReplacePass = new Regex(@"\+", RegexOptions.Compiled);
            var rgexShmReplaceFail = new Regex(@"\-", RegexOptions.Compiled);


            var lFreqMode1DShmooResult = new Dictionary<string, Dictionary<string, List<FreqModeShmooId>>>();
            var lCurrentFreqency = "";
            var lFreqencyList = new List<string>();
            var lKeyname = "";

            //var args = new OrangeXl.ProgressStatus() { Percentage = 0 }; //To Report Progress
            //progress.Report(args);

            #region 建立各種DataTable

            //Shmoo Setup Summary
            var dtSetup = new DataTable("Shmoo_Setup");
            dtSetup.Columns.Add(new DataColumn("Test Num", typeof(int)) { DefaultValue = 0 }); //
            dtSetup.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("Category", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("Free Running Clock", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("TimeSet", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("Clock", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("X Settings", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("Y Settings", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("Vdd Domain", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("ForceCondition", typeof(string)) { DefaultValue = "'-" }); //
            dtSetup.Columns.Add(new DataColumn("Pattern List", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("Payload", typeof(string)) { DefaultValue = "N/A" }); //
            dtSetup.Columns.Add(new DataColumn("Instance List", typeof(string)) { DefaultValue = "N/A" }); //
            CurrShmooReport.Tables.Add(dtSetup);

            //1D Shmoo Summary
            var dtSum1D = new DataTable("Summary_1D");
            dtSum1D.Columns.Add(new DataColumn("Test Num", typeof(int)) { DefaultValue = 0 }); //
            dtSum1D.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum1D.Columns.Add(new DataColumn("Category", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum1D.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum1D.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum1D.Columns.Add(new DataColumn("Die Count", typeof(int)) { DefaultValue = 0 }); //
            dtSum1D.Columns.Add(new DataColumn("Pass Rate", typeof(double)) { DefaultValue = 0.0 }); //
            dtSum1D.Columns.Add(new DataColumn("LVCC:Min", typeof(double)) { DefaultValue = 0.0 }); //
            dtSum1D.Columns.Add(new DataColumn("LVCC:Max", typeof(double)) { DefaultValue = 0.0 }); //
            dtSum1D.Columns.Add(new DataColumn("LVCC:Avg", typeof(double)) { DefaultValue = 0.0 }); //
            dtSum1D.Columns.Add(new DataColumn("Spec", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum1D.Columns.Add(new DataColumn("LVCC <= Spec " + LgbRatioString, typeof(double))
            { DefaultValue = 0.0 }); //
            //dtSum1D.Columns.Add(new DataColumn("LVCC > Spec - 11%", typeof(double)) { DefaultValue = 0.0 }); //
            //dtSum1D.Columns.Add(new DataColumn("LVCC > Spec - 12%", typeof(double)) { DefaultValue = 0.0 }); //
            dtSum1D.Columns.Add(new DataColumn("HVCC:Min", typeof(double)) { DefaultValue = 0.0 }); //
            dtSum1D.Columns.Add(new DataColumn("HVCC:Max", typeof(double)) { DefaultValue = 0.0 }); //
            dtSum1D.Columns.Add(new DataColumn("HVCC:Avg", typeof(double)) { DefaultValue = 0.0 }); //
            dtSum1D.Columns.Add(new DataColumn("HVCC >= Spec " + HgbRatioString, typeof(double))
            { DefaultValue = 0.0 }); //
            dtSum1D.Columns.Add(new DataColumn("Pattern List", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum1D.Columns.Add(new DataColumn("Payload", typeof(string)) { DefaultValue = "N/A" }); //
            CurrShmooReport.Tables.Add(dtSum1D);

            //2D Shmoo Summary
            var dtSum2D = new DataTable("Summary_2D");
            dtSum2D.Columns.Add(new DataColumn("Test Num", typeof(int)) { DefaultValue = 0 }); //
            dtSum2D.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum2D.Columns.Add(new DataColumn("Category", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum2D.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum2D.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum2D.Columns.Add(new DataColumn("Die Count", typeof(int)) { DefaultValue = 0 }); //
            dtSum2D.Columns.Add(new DataColumn("Pass Rate", typeof(double)) { DefaultValue = 0.0 }); //
            dtSum2D.Columns.Add(new DataColumn("Pattern List", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum2D.Columns.Add(new DataColumn("Payload", typeof(string)) { DefaultValue = "N/A" }); //
            dtSum2D.Columns.Add(new DataColumn("Instance List", typeof(string)) { DefaultValue = "N/A" }); //

            if (PlotShmooOverlay)
            {
                dtSum2D.Columns.Add(new DataColumn("Spec " + LgbRatioString, typeof(string)) { DefaultValue = "N/A" }); //
                dtSum2D.Columns.Add(new DataColumn("Spec", typeof(string)) { DefaultValue = "N/A" }); //
                dtSum2D.Columns.Add(new DataColumn("Spec " + HgbRatioString, typeof(string)) { DefaultValue = "N/A" }); //
            }

            CurrShmooReport.Tables.Add(dtSum2D);

            //如果有其他客製化的報表就插在這裡
            var dtSumAbnormal = new DataTable("Summary_Abnormal");
            dtSumAbnormal.Columns.Add(new DataColumn("Test Num", typeof(int)) { DefaultValue = 0 }); //
            dtSumAbnormal.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAbnormal.Columns.Add(new DataColumn("Category", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAbnormal.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAbnormal.Columns.Add(new DataColumn("Payload", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAbnormal.Columns.Add(new DataColumn("Die Count", typeof(int)) { DefaultValue = 0 }); //
            dtSumAbnormal.Columns.Add(new DataColumn("All Pass %", typeof(double)) { DefaultValue = 0.0 }); //
            dtSumAbnormal.Columns.Add(new DataColumn("All Fail %", typeof(double)) { DefaultValue = 0.0 }); //
            dtSumAbnormal.Columns.Add(new DataColumn("Shmoo Hole %", typeof(double)) { DefaultValue = 0.0 }); //
            dtSumAbnormal.Columns.Add(new DataColumn("Hole In +-10 Spec %", typeof(double)) { DefaultValue = 0.0 }); //
            dtSumAbnormal.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAbnormal.Columns.Add(new DataColumn("Pattern List", typeof(string)) { DefaultValue = "N/A" }); //
            CurrShmooReport.Tables.Add(dtSumAbnormal);


            var dtShmooAlram = new DataTable("ShmooAlarm");
            dtShmooAlram.Columns.Add(new DataColumn("Test Num", typeof(int)) { DefaultValue = 0 }); //
            dtShmooAlram.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmooAlram.Columns.Add(new DataColumn("Category", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmooAlram.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" });
            dtShmooAlram.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmooAlram.Columns.Add(new DataColumn("Lot ID", typeof(string)) { DefaultValue = "N/A" });
            dtShmooAlram.Columns.Add(new DataColumn("X,Y", typeof(string)) { DefaultValue = "N/A" });
            dtShmooAlram.Columns.Add(new DataColumn("Alarm Value", typeof(string)) { DefaultValue = "N/A" });
            dtShmooAlram.Columns.Add(new DataColumn("Pattern List", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmooAlram.Columns.Add(new DataColumn("Payload", typeof(string)) { DefaultValue = "N/A" });
            CurrShmooReport.Tables.Add(dtShmooAlram);

            //List 1D Shmoo Hole Device - X,Y
            var dtShmHole1D = new DataTable("ShmooHole_1D");
            dtShmHole1D.Columns.Add(new DataColumn("Test Num", typeof(int)) { DefaultValue = 0 }); //
            dtShmHole1D.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("Category", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("Lot ID", typeof(string))
            { DefaultValue = "N/A" }); //Shmoo Hole看的是全部機率 所以不用全部印出來
            dtShmHole1D.Columns.Add(new DataColumn("X,Y", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("Hole Type", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("X Settings", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("LVCC", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("HVCC", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("Shmoo Result", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("Pattern List", typeof(string)) { DefaultValue = "N/A" }); //
            dtShmHole1D.Columns.Add(new DataColumn("Payload", typeof(string)) { DefaultValue = "N/A" }); //
            CurrShmooReport.Tables.Add(dtShmHole1D);

            //Abnormal Shmoo, 100% Pass or Fail
            var dtSumAllPassOrFail = new DataTable("AllPassOrFail");
            dtSumAllPassOrFail.Columns.Add(new DataColumn("Test Num", typeof(int)) { DefaultValue = 0 }); //
            dtSumAllPassOrFail.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAllPassOrFail.Columns.Add(new DataColumn("Category", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAllPassOrFail.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAllPassOrFail.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAllPassOrFail.Columns.Add(new DataColumn("Lot ID", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAllPassOrFail.Columns.Add(new DataColumn("X,Y", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAllPassOrFail.Columns.Add(new DataColumn("Pass Rate", typeof(double)) { DefaultValue = 0.0 }); //
            dtSumAllPassOrFail.Columns.Add(new DataColumn("Pattern List", typeof(string)) { DefaultValue = "N/A" }); //
            dtSumAllPassOrFail.Columns.Add(new DataColumn("Payload", typeof(string)) { DefaultValue = "N/A" }); //
            CurrShmooReport.Tables.Add(dtSumAllPassOrFail);

            // LVCCHVCC list report
            var dtSum2DShmooLVCCHVCC = new DataTable("2D_LVCCHVCC");
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("Test Num", typeof(int)) { DefaultValue = 0 });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("Category", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("Lot ID", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("X,Y", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("Pattern List", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("Payload", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("X,Y Axis", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("LVCC List", typeof(string)) { DefaultValue = "N/A" });
            dtSum2DShmooLVCCHVCC.Columns.Add(new DataColumn("HVCC List", typeof(string)) { DefaultValue = "N/A" });
            CurrShmooReport.Tables.Add(dtSum2DShmooLVCCHVCC);

            var dtSelSramDigSrc = new DataTable("SelSramDigSrcCheck");
            dtSelSramDigSrc.Columns.Add(new DataColumn("Test Num", typeof(string)) { DefaultValue = "N/A" });
            dtSelSramDigSrc.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" });
            dtSelSramDigSrc.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" });
            dtSelSramDigSrc.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" });
            dtSelSramDigSrc.Columns.Add(new DataColumn("Site", typeof(string)) { DefaultValue = "N/A" });
            dtSelSramDigSrc.Columns.Add(new DataColumn("X,Y", typeof(string)) { DefaultValue = "N/A" });
            dtSelSramDigSrc.Columns.Add(new DataColumn("DigSrc Code  =>  GatingRule", typeof(string))
            { DefaultValue = "N/A" });
            dtSelSramDigSrc.Columns.Add(new DataColumn("PowerSetting", typeof(string)) { DefaultValue = "N/A" });
            CurrShmooReport.Tables.Add(dtSelSramDigSrc);


            //標準Shmoo圖的Sheet
            foreach (var category in ListAllCategories) //每個Category一張Sheet 以後可以平行處理去算
            {
                var dt = new DataTable(category);
                dt.Columns.Add(new DataColumn("Shmoo Type", typeof(string))
                { DefaultValue = "N/A" }); //1D 2D 混在同一個Category DataTable裡面!! 

                //一份就好 可以共用
                dt.Columns.Add(new DataColumn("X Axis", typeof(string)) { DefaultValue = "N/A" });
                dt.Columns.Add(new DataColumn("Y Axis", typeof(string)) { DefaultValue = "N/A" });
                dt.Columns.Add(new DataColumn("Titles", typeof(string)) { DefaultValue = "N/A" });
                dt.Columns.Add(new DataColumn("Patterns", typeof(string)) { DefaultValue = "N/A" });
                dt.Columns.Add(new DataColumn("TestInstances", typeof(string)) { DefaultValue = "N/A" });

                dt.Columns.Add(new DataColumn("Info Headers", typeof(string)) { DefaultValue = "N/A" }); //1D左邊那些統計的資訊
                dt.Columns.Add(new DataColumn("Shmoo Step", typeof(string)) { DefaultValue = "N/A" }); //X,Y Step Count

                //不見得會有
                dt.Columns.Add(new DataColumn("Spec X Point", typeof(string))
                { DefaultValue = "N/A" }); //標示Spec的Row Col"範圍" 給Epplus畫線
                dt.Columns.Add(new DataColumn("Spec Y Point", typeof(string))
                { DefaultValue = "N/A" }); //標示Spec的Row Col"範圍" 給Epplus畫線

                //By Device 數量隔開
                dt.Columns.Add(new DataColumn("Die Info", typeof(string)) { DefaultValue = "N/A" });
                dt.Columns.Add(new DataColumn("All Content", typeof(string)) { DefaultValue = "N/A" });

                // BinCut Spec
                dt.Columns.Add(new DataColumn("BinCutPlan", typeof(string)) { DefaultValue = "N/A" });
                dt.Columns.Add(new DataColumn("BinCutVersion", typeof(string)) { DefaultValue = "N/A" });
                dt.Columns.Add(new DataColumn("BinCutPmode", typeof(string)) { DefaultValue = "N/A" });
                dt.Columns.Add(new DataColumn("BinCutSpec", typeof(string)) { DefaultValue = "N/A" });
                dt.Columns.Add(new DataColumn("BinCutSpecName", typeof(string)) { DefaultValue = "N/A" });

                CurrShmooReport.Tables.Add(dt);
            }

            #endregion

            var currCatCount = 1; //用來回報進度用
            foreach (var category in ListAllCategories)
            {
                //args.Result = string.Format("Analyzing " + category + "!!");
                //progress.Report(args);

                var listShmooSetups = DicCategoryShmooSetups[category]; //這個Sheet內所有的Shmoo Setup(含1D2D)


                if (OrderByTestNum)
                    listShmooSetups = listShmooSetups.OrderBy(p => p.TestNum).ToList();

                foreach (var shmooSetup in listShmooSetups)
                    try
                    {
                        #region Strat loop

                        //For 第一頁Shmoo Setup
                        var tmpRowShmSetup = dtSetup.NewRow(); //一個Setup一行
                        tmpRowShmSetup["Test Num"] = shmooSetup.TestNum;
                        tmpRowShmSetup["Type"] = shmooSetup.Type;
                        tmpRowShmSetup["Category"] = shmooSetup.Category;
                        tmpRowShmSetup["Test Instance"] = shmooSetup.TestInstanceName;
                        tmpRowShmSetup["Test Setup"] = shmooSetup.SetupName;
                        tmpRowShmSetup["Free Running Clock"] = shmooSetup.FreeRunningClk;
                        tmpRowShmSetup["X Settings"] = string.Join(",", shmooSetup.SettingsX);
                        tmpRowShmSetup["Y Settings"] = string.Join(",", shmooSetup.SettingsY);
                        tmpRowShmSetup["Vdd Domain"] = string.Join(",", shmooSetup.SettingsInit);
                        tmpRowShmSetup["Pattern List"] = string.Join(",\r\n", shmooSetup.PatternList);
                        tmpRowShmSetup["Payload"] = string.Join(",\r\n", shmooSetup.PayloadList);
                        tmpRowShmSetup["Instance List"] = string.Join(",\r\n", shmooSetup.InstanceList);
                        tmpRowShmSetup["ForceCondition"] = shmooSetup.ForceCondition;
                        //For 每一獨立的Shmoo Sheet 每一個Row代表將來的一組要畫的區域

                        // 20180402  timeset & clock setting
                        if (shmooSetup.TimeSetInfo != "")
                        {
                            var timeArray = shmooSetup.TimeSetInfo.Split(',');
                            tmpRowShmSetup["TimeSet"] = timeArray[0];
                            if (timeArray.Count() > 1)
                            {
                                var clockSetting = "";

                                for (var i = 1; i < timeArray.Count(); i++)
                                {
                                    var clockName = timeArray[i].Split('=')[0];
                                    var outD = 0.0;
                                    if (double.TryParse(timeArray[i].Split('=')[1], out outD))
                                        clockSetting += string.Format(clockName + " = {0:0.00} Mhz", outD / 1000000.0);
                                    else
                                        clockSetting = timeArray[i] + ",\r\n";
                                }

                                tmpRowShmSetup["Clock"] = clockSetting;
                            }
                        }


                        var tmpRowCategory = CurrShmooReport.Tables[category].NewRow(); //一個Setup一行
                        tmpRowCategory["Shmoo Type"] = shmooSetup.Type;
                        tmpRowCategory["Patterns"] = string.Join(",", shmooSetup.PatternList);
                        tmpRowCategory["TestInstances"] = string.Join(",", shmooSetup.InstanceList);
                        tmpRowCategory["Titles"] = "Test Num:" + shmooSetup.TestNum + ":Test Instance:" +
                                                   shmooSetup.TestInstanceName + ":Setup:" + shmooSetup.SetupName;

                        var axisAllX = ""; //1D 2D通用 要考慮Tracking

                        shmooSetup.DicAxisValue.XAxis.PointUnitConvert(); // 小數單位置換 

                        foreach (var axis in shmooSetup.Axis[0])
                            axisAllX += axis + ";" + shmooSetup.AxisType[0] + ", SI Prefix: " +
                                        shmooSetup.DicAxisValue.XAxis.Unit + ";" + string.Join(";",
                                            shmooSetup.DicAxisValue.XAxis.PinValueSet[axis]) + "#";
                        tmpRowCategory["X Axis"] = axisAllX.Remove(axisAllX.Length - 1);

                        //1D 2D通用 Spec
                        if (shmooSetup.SpecPoints[0].Count > 0 && SpecRatio.Count > 0)
                            tmpRowCategory["Spec X Point"] = string.Join(",", shmooSetup.SpecPoints[0]);


                        if (shmooSetup.PlotBinCutSpec)
                        {
                            tmpRowCategory["BinCutPlan"] = BinCutRefData.BinCutFileName;
                            tmpRowCategory["BinCutVersion"] = BinCutRefData.BinCutVersion;
                            tmpRowCategory["BinCutPmode"] = shmooSetup.BinCutPerfromanceName;

                            if (shmooSetup.IsOtherRailSpec)
                            {
                                tmpRowCategory["BinCutSpec"] = shmooSetup.BinCutSpecPoints[0];
                                tmpRowCategory["BinCutSpecName"] =
                                    BinCutRefData.Job + "Max\r\n" + shmooSetup.BinCutSpecValue[0];
                            }

                            else
                            {
                                // same point merge same cell
                                if (shmooSetup.BinCutSpecPoints[0] == shmooSetup.BinCutSpecPoints[1])
                                {
                                    tmpRowCategory["BinCutSpec"] = shmooSetup.BinCutSpecPoints[0];
                                    tmpRowCategory["BinCutSpecName"] =
                                        BinCutRefData.Job + "Max:" + shmooSetup.BinCutSpecValue[0] +
                                        "\r\n" + BinCutRefData.Job + "Min:" + shmooSetup.BinCutSpecValue[1];
                                }
                                else
                                {
                                    tmpRowCategory["BinCutSpec"] = shmooSetup.BinCutSpecPoints[0] + "," +
                                                                   shmooSetup.BinCutSpecPoints[1];
                                    tmpRowCategory["BinCutSpecName"] = BinCutRefData.Job + "Max:" +
                                                                       shmooSetup.BinCutSpecValue[0] + "," +
                                                                       BinCutRefData.Job + "Min:" +
                                                                       shmooSetup.BinCutSpecValue[1];
                                }
                            }
                        }

                        if (shmooSetup.Type == "1D")
                        {
                            //For 每一獨立的Shmoo Sheet 每一個Row代表將來的一組要畫的區域
                            tmpRowCategory["Info Headers"] = "Source File,Site,Lot ID,Die XY,S.Bin,Lvcc,Hvcc,Hole";
                            tmpRowCategory["Shmoo Step"] = shmooSetup.StepCount[0].ToString();

                            var allContent = "";
                            var allDieInfo = "";
                            var allLvcc = new List<double>();
                            var allHvcc = new List<double>();

                            if (InstacneFreqMode)
                            {
                                if (!lFreqMode1DShmooResult.ContainsKey(category))
                                    lFreqMode1DShmooResult[category] = new Dictionary<string, List<FreqModeShmooId>>();


                                var lInstacneName = shmooSetup.TestInstanceName;
                                var lShmooSetupName = shmooSetup.SetupName;
                                var match = _regexFreqDector.Match(lInstacneName);
                                if (match.Success)
                                {
                                    lCurrentFreqency = match.Groups["Freq"].ToString(); //20MHZ
                                    //SOCSACHAIN_ALLFV_PP_CYPA0_S_PL00_CH_CLXB_SAA_UNC_AUT_ALLFV_DM_20MHZ_NV_VDD_FIXED_VDD_SOC_VDD_LOW 
                                    //SOCSACHAIN_ALLFV_PP_CYPA0_S_PL00_CH_CLXB_SAA_UNC_AUT_ALLFV_DM_NV_VDD_FIXED_VDD_SOC_VDD_LOW 
                                    lKeyname = lInstacneName.Replace("_" + lCurrentFreqency, "") + ":" +
                                               lShmooSetupName;
                                }
                                else
                                {
                                    lKeyname = lInstacneName + ":" + lShmooSetupName;
                                    lCurrentFreqency = "Default";
                                }

                                if (!lFreqencyList.Contains(lCurrentFreqency))
                                    lFreqencyList.Add(lCurrentFreqency);

                                if (!lFreqMode1DShmooResult[category].ContainsKey(lKeyname))
                                    lFreqMode1DShmooResult[category].Add(lKeyname, new List<FreqModeShmooId>());
                            }


                            var dicOverlay1D = new Dictionary<int, int>(); //處理疊圖 Step -> Pass Count
                            var dicOverlay1DSite =
                                new Dictionary<string, Dictionary<int, int>>(); //處理疊圖 Step -> Pass Count
                            for (var s = 0; s < shmooSetup.StepCount[0]; s++) dicOverlay1D[s] = 0; //初始化

                            var allLvccSite = new Dictionary<string, List<double>>();
                            var allHvccSite = new Dictionary<string, List<double>>();
                            var eachSiteCount = new Dictionary<string, int>();
                            var dicSiteInfo = new Dictionary<string, string>();

                            foreach (var shmooId in shmooSetup.ShmooIDs)
                            {
                                if (InstacneFreqMode)
                                {
                                    if (!lFreqMode1DShmooResult[category][lKeyname].Exists(
                                            p => p.LotId.Equals(shmooId.LotId) &&
                                                 p.Site.Equals(shmooId.Site) &&
                                                 p.DieXY.Equals(shmooId.DieXY)))
                                        lFreqMode1DShmooResult[category][lKeyname].Add(
                                            new FreqModeShmooId
                                            {
                                                LotId = shmooId.LotId,
                                                Site = shmooId.Site,
                                                DieXY = shmooId.DieXY
                                            });


                                    var idFreqShmooRestlt = lFreqMode1DShmooResult[category][lKeyname].FirstOrDefault(
                                        p => p.LotId.Equals(shmooId.LotId) &&
                                             p.Site.Equals(shmooId.Site) &&
                                             p.DieXY.Equals(shmooId.DieXY));

                                    if (idFreqShmooRestlt != null)
                                    {
                                        if (!idFreqShmooRestlt.FreqLvccDict.ContainsKey(lCurrentFreqency))
                                            idFreqShmooRestlt.FreqLvccDict.Add(lCurrentFreqency, shmooId.Lvcc);
                                        if (!idFreqShmooRestlt.FreqHvccDict.ContainsKey(lCurrentFreqency))
                                            idFreqShmooRestlt.FreqHvccDict.Add(lCurrentFreqency, shmooId.Hvcc);
                                    }


                                    //    .FreqLvccDict.Add(lCurrentFreqency, shmooId.Lvcc);

                                    //lFreqMode1DShmooResult[category][lKeyname].Find(
                                    //    p => p.LotId.Equals(shmooId.LotId) &&
                                    //         p.Site.Equals(shmooId.Site) &&
                                    //         p.DieXY.Equals(shmooId.DieXY))
                                    //    .FreqHvccDict.Add(lCurrentFreqency, shmooId.Hvcc);

                                    // SY shmoo loop use
                                    //var freqMode = new FreqModeShmooId()
                                    //{
                                    //    LotId = shmooId.LotId,
                                    //    Site = shmooId.Site,
                                    //    DieXY = shmooId.DieXY,
                                    //};

                                    //freqMode.FreqLvccDict.Add(lCurrentFreqency, shmooId.Lvcc);
                                    //freqMode.FreqHvccDict.Add(lCurrentFreqency, shmooId.Hvcc);
                                    //lFreqMode1DShmooResult[category][lKeyname].Add(freqMode);
                                }


                                if (!allLvccSite.ContainsKey(shmooId.Site))
                                {
                                    allLvccSite[shmooId.Site] = new List<double>();
                                    allHvccSite[shmooId.Site] = new List<double>();
                                    dicOverlay1DSite[shmooId.Site] = new Dictionary<int, int>();
                                    for (var s = 0; s < shmooSetup.StepCount[0]; s++)
                                        dicOverlay1DSite[shmooId.Site][s] = 0; //初始化
                                }

                                if (!eachSiteCount.ContainsKey(shmooId.Site))
                                    eachSiteCount[shmooId.Site] = 1;
                                else
                                    eachSiteCount[shmooId.Site]++;

                                if (!dicSiteInfo.ContainsKey("DieXY" + shmooId.Site))
                                    dicSiteInfo.Add("DieXY" + shmooId.Site, shmooId.DieXY);
                                if (!dicSiteInfo.ContainsKey("sBin" + shmooId.Site))
                                    dicSiteInfo.Add("sBin" + shmooId.Site, shmooId.Sort);
                                if (!dicSiteInfo.ContainsKey("LotId" + shmooId.Site))
                                    dicSiteInfo.Add("LotId" + shmooId.Site, shmooId.LotId);

                                if (SramDataSet.Any())
                                {
                                    var dieSramCondition = SramDataSet.FindAll(p => p.GetXYDie().Equals(shmooId.DieXY)
                                        & p.InstanceName.Equals(shmooSetup.TestInstanceName)
                                        & p.Site.Equals(shmooId.Site));

                                    if (dieSramCondition.Any())
                                        foreach (var sramCond in dieSramCondition)
                                        {
                                            var sramRow = dtSelSramDigSrc.NewRow(); //一個Setup一行
                                            sramRow["Test Num"] = shmooSetup.TestNum;
                                            sramRow["Type"] = shmooSetup.Type;
                                            sramRow["Test Instance"] = shmooSetup.TestInstanceName;
                                            sramRow["Test Setup"] = shmooSetup.SetupName;
                                            sramRow["Site"] = sramCond.Site;
                                            sramRow["X,Y"] = sramCond.GetXYDie();
                                            if (sramCond.Check())
                                                sramRow["DigSrc Code  =>  GatingRule"] =
                                                    sramCond.OrgCompressStr + " (P)";
                                            else
                                                sramRow["DigSrc Code  =>  GatingRule"] = sramCond.OrgCompressStr +
                                                    " (F) => " + sramCond.CompareCompressStr;
                                            sramRow["PowerSetting"] = sramCond.PowerSettingStr;
                                            dtSelSramDigSrc.Rows.Add(sramRow);
                                        }
                                }
                            }

                            //定義Abnormal的條件
                            var shmooHoleCount = 0;
                            var allPassedCount = 0;
                            var allFailedCount = 0;
                            var holeInSpecRange = 0;

                            // add for TPDD2 filter Noice Die
                            shmooSetup.ShmooIDs = ByPassGrrNoiseDieByCriteria(shmooSetup.ShmooIDs);

                            foreach (var shmooId in
                                     shmooSetup.ShmooIDs) //收集Die Info, Content, 統計Lvcc  <-- 這一段很有潛力做平行處理 應該在更外層就處理好?
                            {
                                foreach (Match m in rgexShmPass.Matches(shmooId.ShmooContent[0]))
                                {
                                    dicOverlay1D[m.Index]++; //重要技巧 尋找String內符合條件的每個Char
                                    dicOverlay1DSite[shmooId.Site][m.Index]++;
                                }

                                var processedContent = rgexShmReplacePass.Replace(shmooId.ShmooContent[0], "P");

                                processedContent = rgexShmReplaceFail.Replace(processedContent, "F");

                                allContent += string.Join(",", processedContent) + "#";
                                allDieInfo += shmooId.SourceFileName + "|" + shmooId.Site + "|" + shmooId.LotId + "|" +
                                              shmooId.DieXY + "|" + shmooId.Sort + "|"
                                              + shmooId.Lvcc + "|" + shmooId.Hvcc + "|" + shmooId.ShmooHole + "#";

                                if (shmooId.ShmooHoleInOperationRange) holeInSpecRange++;

                                if (FindShmooHole1D && shmooId.ShmooHole != "NH")
                                {
                                    //統計Shmoo Hole
                                    var tmpRowShmHole1D = dtShmHole1D.NewRow(); //一個Setup一行
                                    tmpRowShmHole1D["Test Num"] = shmooSetup.TestNum;
                                    tmpRowShmHole1D["Type"] = shmooSetup.Type;
                                    tmpRowShmHole1D["Category"] = shmooSetup.Category;
                                    tmpRowShmHole1D["Test Instance"] = shmooSetup.TestInstanceName;
                                    tmpRowShmHole1D["Test Setup"] = shmooSetup.SetupName;
                                    tmpRowShmHole1D["X Settings"] = string.Join(",", shmooSetup.SettingsX);
                                    tmpRowShmHole1D["Lot ID"] = shmooId.LotId;
                                    tmpRowShmHole1D["X,Y"] = shmooId.DieXY;
                                    tmpRowShmHole1D["Hole Type"] = shmooId.ShmooHole;
                                    tmpRowShmHole1D["Shmoo Result"] = shmooId.ShmooContent[0];
                                    tmpRowShmHole1D["LVCC"] = shmooId.Abnormal_Lvcc;
                                    tmpRowShmHole1D["HVCC"] = shmooId.Abnormal_Hvcc;
                                    tmpRowShmHole1D["Pattern List"] = string.Join(",\r\n", shmooSetup.PatternList);
                                    tmpRowShmHole1D["Payload"] = string.Join(",\r\n", shmooSetup.PayloadList);
                                    dtShmHole1D.Rows.Add(tmpRowShmHole1D);
                                    shmooHoleCount++;
                                }

                                // Fix Sweep on Vil issue: low Vil pass and high Vil fail (Lvcc = 0 and should not block Hvcc) 
                                /* original
                                 * if (shmooId.Lvcc > 0)
                                 * {
                                 *   allLvcc.Add(shmooId.Lvcc);
                                 *   allHvcc.Add(shmooId.Hvcc);
                                 * }
                                 */


                                var lowerBound = shmooSetup.AcurateStep[0][0] > shmooSetup.AcurateStep[0][1]
                                    ? shmooSetup.AcurateStep[0][1]
                                    : shmooSetup.AcurateStep[0][0];

                                var higherBound = shmooSetup.AcurateStep[0][0] > shmooSetup.AcurateStep[0][1]
                                    ? shmooSetup.AcurateStep[0][0]
                                    : shmooSetup.AcurateStep[0][1];

                                if (shmooId.Lvcc >= lowerBound)
                                {
                                    allLvccSite[shmooId.Site].Add(shmooId.Lvcc);
                                    allLvcc.Add(shmooId.Lvcc);
                                }

                                if (shmooId.Hvcc <= higherBound)
                                {
                                    allHvccSite[shmooId.Site].Add(shmooId.Hvcc);
                                    allHvcc.Add(shmooId.Hvcc);
                                }
                                //End of Fix Sepp on Vil issue

                                if (shmooId.Abnormal)
                                {
                                    var tmpAbRow = dtSumAllPassOrFail.NewRow();
                                    tmpAbRow["Test Num"] = shmooSetup.TestNum;
                                    tmpAbRow["Type"] = shmooSetup.Type;
                                    tmpAbRow["Category"] = shmooSetup.Category;
                                    tmpAbRow["Test Instance"] = shmooSetup.TestInstanceName;
                                    tmpAbRow["Test Setup"] = shmooSetup.SetupName;
                                    tmpAbRow["Lot ID"] = shmooId.LotId;
                                    tmpAbRow["X,Y"] = shmooId.DieXY;
                                    tmpAbRow["Pass Rate"] = shmooId.PassRate;
                                    tmpAbRow["Pattern List"] = string.Join(",\r\n", shmooSetup.PatternList);
                                    tmpAbRow["Payload"] = string.Join(",\r\n", shmooSetup.PayloadList);
                                    dtSumAllPassOrFail.Rows.Add(tmpAbRow);

                                    if (shmooId.IsAllFailed) allFailedCount++;
                                    else allPassedCount++;
                                }

                                if (shmooId.ShmooAlarm)
                                {
                                    var tmpAlarmRow = dtShmooAlram.NewRow();
                                    tmpAlarmRow["Test Num"] = shmooSetup.TestNum;
                                    tmpAlarmRow["Type"] = shmooSetup.Type;
                                    tmpAlarmRow["Category"] = shmooSetup.Category;
                                    tmpAlarmRow["Test Instance"] = shmooSetup.TestInstanceName;
                                    tmpAlarmRow["Test Setup"] = shmooSetup.SetupName;
                                    tmpAlarmRow["Lot ID"] = shmooId.LotId;
                                    tmpAlarmRow["X,Y"] = shmooId.DieXY;
                                    tmpAlarmRow["Alarm Value"] = shmooId.ShmooAlarmValue;
                                    tmpAlarmRow["Pattern List"] = string.Join(",\r\n", shmooSetup.PatternList);
                                    tmpAlarmRow["Payload"] = string.Join(",\r\n", shmooSetup.PayloadList);
                                    dtShmooAlram.Rows.Add(tmpAlarmRow);
                                }
                            }

                            //如果有要處理疊圖 在這裡把資訊加上去
                            if (PlotShmooOverlay && shmooSetup.ShmooIDs.Count > 1)
                            {
                                var tmpAllStr = "";
                                var rebuildOverlayStr1D = ""; //Rebuild Overlay 1D
                                for (var s = 0; s < shmooSetup.StepCount[0]; s++)
                                    rebuildOverlayStr1D += string.Format("{0:F0}",
                                        dicOverlay1D[s] * 100 / shmooSetup.ShmooIDs.Count) + ",";
                                tmpAllStr += rebuildOverlayStr1D.Remove(rebuildOverlayStr1D.Length - 1) + "#";

                                if (ShowPercBySite)
                                    foreach (var key in dicOverlay1DSite.Keys)
                                    {
                                        rebuildOverlayStr1D = "";
                                        for (var s = 0; s < shmooSetup.StepCount[0]; s++)
                                            rebuildOverlayStr1D += string.Format("{0:F0}",
                                                dicOverlay1DSite[key][s] * 100 / eachSiteCount[key]) + ",";
                                        tmpAllStr += rebuildOverlayStr1D.Remove(rebuildOverlayStr1D.Length - 1) + "#";
                                    }

                                allContent = tmpAllStr + allContent;

                                tmpAllStr = "";
                                if (allLvcc.Count > 0 && allHvcc.Count > 0)
                                    tmpAllStr = "All|All|All|All|All|" + string.Format("{0:F2}", allLvcc.Average()) +
                                                "|" + string.Format("{0:F2}", allHvcc.Average()) + "|All#";
                                else
                                    tmpAllStr = "All|All|All|All|All|" + "0.0" + "|" + "0.0" + "|All#";

                                if (ShowPercBySite)
                                    foreach (var key in allLvccSite.Keys)
                                    {
                                        double allLvccSiteVal = 0;
                                        double allHvccSiteVal = 0;
                                        if (allLvccSite[key].Count > 0) allLvccSiteVal = allLvccSite[key].Average();
                                        if (allHvccSite[key].Count > 0) allHvccSiteVal = allHvccSite[key].Average();
                                        //tmpAllStr += "All|" + key + "|" + dicSiteInfo["LotId" + key] + "|" +
                                        //dicSiteInfo["DieXY" + key] + "|" + dicSiteInfo["sBin" + key] + "|" +
                                        //String.Format("{0:F2}", allLvccSiteVal) + "|" +
                                        //String.Format("{0:F2}", allHvccSiteVal) + "|All#";

                                        tmpAllStr += "All|" + key + "|" + dicSiteInfo["LotId" + key] + "|" +
                                                     "Site:" + key + "|" + "Site:" + key + "|" +
                                                     string.Format("{0:F2}", allLvccSiteVal) + "|" +
                                                     string.Format("{0:F2}", allHvccSiteVal) + "|All#";
                                    }

                                allDieInfo = tmpAllStr + allDieInfo;
                            }

                            tmpRowCategory["Die Info"] = allDieInfo.Remove(allDieInfo.Length - 1);
                            tmpRowCategory["All Content"] = allContent.Remove(allContent.Length - 1);

                            //看看是不是需要Report Abnormal
                            if (shmooHoleCount > 0 || allPassedCount > 0 || allFailedCount > 0)
                            {
                                var tmpRowAbnormalSum = dtSumAbnormal.NewRow();
                                tmpRowAbnormalSum["Test Num"] = shmooSetup.TestNum;
                                tmpRowAbnormalSum["Type"] = shmooSetup.Type;
                                tmpRowAbnormalSum["Category"] = shmooSetup.Category;
                                tmpRowAbnormalSum["Test Instance"] = shmooSetup.TestInstanceName;
                                tmpRowAbnormalSum["Test Setup"] = shmooSetup.SetupName;
                                tmpRowAbnormalSum["Die Count"] = shmooSetup.ShmooIDs.Count;
                                tmpRowAbnormalSum["Payload"] = string.Join(",\r\n", shmooSetup.PayloadList);
                                tmpRowAbnormalSum["Pattern List"] = string.Join(",\r\n", shmooSetup.PatternList);

                                //dtSumAbnormal.Columns.Add(new DataColumn("Hole In +-10 Spec %", typeof(double)) { DefaultValue = 0.0 }); //
                                if (holeInSpecRange > 0)
                                    tmpRowAbnormalSum["Hole In +-10 Spec %"] = Math.Round(
                                        holeInSpecRange / (shmooSetup.ShmooIDs.Count * 1.0) * 100.0, 2,
                                        MidpointRounding.AwayFromZero);
                                if (shmooHoleCount > 0)
                                    tmpRowAbnormalSum["Shmoo Hole %"] = Math.Round(
                                        shmooHoleCount / (shmooSetup.ShmooIDs.Count * 1.0) * 100.0, 2,
                                        MidpointRounding.AwayFromZero);
                                if (allPassedCount > 0)
                                    tmpRowAbnormalSum["All Pass %"] = Math.Round(
                                        allPassedCount / (shmooSetup.ShmooIDs.Count * 1.0) * 100.0, 2,
                                        MidpointRounding.AwayFromZero);
                                if (allFailedCount > 0)
                                    tmpRowAbnormalSum["All Fail %"] = Math.Round(
                                        allFailedCount / (shmooSetup.ShmooIDs.Count * 1.0) * 100.0, 2,
                                        MidpointRounding.AwayFromZero);

                                dtSumAbnormal.Rows.Add(tmpRowAbnormalSum);
                            }

                            //For 1D Shmoo Summary統計用 統計每顆Die的資訊
                            var tmpRowShmSum1D = dtSum1D.NewRow(); //一個Setup一行
                            tmpRowShmSum1D["Test Num"] = shmooSetup.TestNum;
                            tmpRowShmSum1D["Type"] = shmooSetup.Type;
                            tmpRowShmSum1D["Category"] = shmooSetup.Category;
                            tmpRowShmSum1D["Test Instance"] = shmooSetup.TestInstanceName;
                            tmpRowShmSum1D["Test Setup"] = shmooSetup.SetupName;
                            tmpRowShmSum1D["Die Count"] = shmooSetup.ShmooIDs.Count;
                            tmpRowShmSum1D["Pattern List"] = string.Join(",\r\n", shmooSetup.PatternList);
                            tmpRowShmSum1D["Payload"] = string.Join(",\r\n", shmooSetup.PayloadList);

                            if (shmooSetup.ShmooIDs.Count > 0)
                                tmpRowShmSum1D["Pass Rate"] =
                                    Math.Round(allLvcc.Count / (shmooSetup.ShmooIDs.Count * 1.0) * 100.0, 2,
                                        MidpointRounding.AwayFromZero);
                            tmpRowShmSum1D["LVCC:Min"] = allLvcc.Count > 0 ? allLvcc.Min() : 0.0;
                            tmpRowShmSum1D["LVCC:Max"] = allLvcc.Count > 0 ? allLvcc.Max() : 0.0;
                            tmpRowShmSum1D["LVCC:Avg"] = allLvcc.Count > 0
                                ? Math.Round(allLvcc.Average(), 3, MidpointRounding.AwayFromZero)
                                : 0.0;


                            if (allLvcc.Count > 0 && shmooSetup.Spec[0] > 0) //表示有找到Spec
                            {
                                tmpRowShmSum1D["Spec"] =
                                    Regex.Split(shmooSetup.Axis[0][shmooSetup.Axis[0].Count - 1], @":")[0] + ":" +
                                    shmooSetup.Spec[0]; //就是Spec Name
                                var passedSpecL = (from q in allLvcc where q <= shmooSetup.Spec[0] * LgbRatio select q)
                                    .ToList();
                                tmpRowShmSum1D["LVCC <= Spec " + LgbRatioString] = Math.Round(
                                    passedSpecL.Count * 100.0 / allLvcc.Count, 3, MidpointRounding.AwayFromZero);
                                var passedSpecH = (from q in allHvcc where q >= shmooSetup.Spec[0] * HgbRatio select q)
                                    .ToList();
                                tmpRowShmSum1D["HVCC >= Spec " + HgbRatioString] = Math.Round(
                                    passedSpecH.Count * 100.0 / allHvcc.Count, 3, MidpointRounding.AwayFromZero);
                            }

                            tmpRowShmSum1D["HVCC:Min"] = allHvcc.Count > 0 ? allHvcc.Min() : 0.0;
                            tmpRowShmSum1D["HVCC:Max"] = allHvcc.Count > 0 ? allHvcc.Max() : 0.0;
                            tmpRowShmSum1D["HVCC:Avg"] = allHvcc.Count > 0
                                ? Math.Round(allHvcc.Average(), 3, MidpointRounding.AwayFromZero)
                                : 0.0;
                            dtSum1D.Rows.Add(tmpRowShmSum1D);
                        }
                        else //2D Shmoo
                        {
                            var axisAllY = "";
                            foreach (var axis in shmooSetup.Axis[1])
                                axisAllY += axis + ":" + shmooSetup.AxisType[1] + ":" +
                                            string.Join(":", shmooSetup.DicAxisValue.YAxis.PinValueSet[axis]) + "#";
                            tmpRowCategory["Y Axis"] = axisAllY.Remove(axisAllY.Length - 1);

                            if (shmooSetup.SpecPoints[1].Count > 0 && SpecRatio.Count > 0)
                                tmpRowCategory["Spec Y Point"] = string.Join(",", shmooSetup.SpecPoints[1]);

                            //For 每一獨立的Shmoo Sheet 每一個Row代表將來的一組要畫的區域
                            tmpRowCategory["Shmoo Step"] = shmooSetup.StepCount[0] + ":" + shmooSetup.StepCount[1];

                            var allContent = "";
                            var allDieInfo = "";


                            var XPointYMaxMinIndexRange = new Dictionary<int, List<int>>();
                            var passPoint = 0;

                            if (!shmooSetup.IsMergeBySiteSetup)
                            {
                                var dicOverlay2D = new Dictionary<int, Dictionary<int, int>>();
                                //處理疊圖 Step X Y -> Pass Count

                                var dicMerge2D2Inse = new Dictionary<int, Dictionary<int, string>>();
                                //處理疊圖 Step X Y -> PP, PF, FP , FF

                                for (var x = 0; x < shmooSetup.StepCount[0]; x++)
                                {
                                    dicOverlay2D[x] = new Dictionary<int, int>();
                                    dicMerge2D2Inse[x] = new Dictionary<int, string>();
                                    for (var y = 0; y < shmooSetup.StepCount[1]; y++)
                                    {
                                        dicOverlay2D[x][y] = 0; //初始化
                                        dicMerge2D2Inse[x][y] = "";
                                    }
                                }


                                foreach (var shmooId in shmooSetup.ShmooIDs) //收集Die Info, Content, 統計Lvcc
                                {
                                    for (var y = 0; y < shmooId.ShmooContent.Count; y++)
                                    {
                                        var passPointIndex = new List<int>();
                                        foreach (Match m in rgexShmPass.Matches(shmooId.ShmooContent[y])) //重要技巧!
                                        {
                                            dicOverlay2D[m.Index][y]++;
                                            passPointIndex.Add(m.Index);
                                            passPoint++;
                                        }

                                        for (var i = 0; i < shmooId.ShmooContent[y].Length; i++)
                                            if (passPointIndex.Contains(i))
                                                dicMerge2D2Inse[i][y] += "P";
                                            else
                                                dicMerge2D2Inse[i][y] += "F";
                                    }

                                    var processedContent =
                                        rgexShmReplacePass.Replace(string.Join(",", shmooId.ShmooContent), "P");
                                    //把+換成P比較習慣

                                    var index = 0;
                                    var thisChar = new List<string>();
                                    foreach (var c in processedContent)
                                    {
                                        var mc = rgexShmReplaceFail.Replace(c.ToString(), "F");
                                        //shmooId.FailFlagIndexDict[index].ToString()
                                        thisChar.Add(mc);
                                        index++;
                                    }

                                    processedContent = string.Join("", thisChar.ToArray());


                                    allContent += processedContent + "#";

                                    allDieInfo += shmooId.SourceFileName + "|" + shmooId.Site + "|" + shmooId.LotId +
                                                  "|" + shmooId.DieXY + "|" + shmooId.Sort + "|" +
                                                  shmooId.ShmooInstanceName + "#";

                                    if (shmooId.Abnormal)
                                    {
                                        var tmpAbRow = dtSumAllPassOrFail.NewRow();
                                        tmpAbRow["Test Num"] = shmooSetup.TestNum;
                                        tmpAbRow["Type"] = shmooSetup.Type;
                                        tmpAbRow["Category"] = shmooSetup.Category;
                                        tmpAbRow["Test Instance"] = shmooSetup.TestInstanceName;
                                        tmpAbRow["Test Setup"] = shmooSetup.SetupName;
                                        tmpAbRow["Lot ID"] = shmooId.LotId;
                                        tmpAbRow["X,Y"] = shmooId.DieXY;
                                        tmpAbRow["Pass Rate"] = shmooId.PassRate;
                                        tmpAbRow["Pattern List"] = string.Join(",\r\n", shmooSetup.PatternList);
                                        tmpAbRow["Payload"] = string.Join(",\r\n", shmooSetup.PayloadList);
                                        dtSumAllPassOrFail.Rows.Add(tmpAbRow);
                                    }

                                    if (LVCCHVCC2DRReport)
                                    {
                                        var tmp2DLVCCHVCRow = dtSum2DShmooLVCCHVCC.NewRow();
                                        var templvcclist = new List<string>();
                                        var temphvcclist = new List<string>();
                                        var axisY = "";
                                        var axisYType = "";
                                        var axisX = "";
                                        var axisXType = "";
                                        foreach (var axis in shmooSetup.Axis[1])
                                        {
                                            axisY = axis;
                                            axisYType = shmooSetup.AxisType[1];
                                        }

                                        foreach (var axis in shmooSetup.Axis[0])
                                        {
                                            axisX = axis;
                                            axisXType = shmooSetup.AxisType[0];
                                        }


                                        for (var i = 0; i < shmooSetup.DicAxisValue.YAxis.PinValueSet[axisY].Count; i++)
                                        {
                                            templvcclist.Add(shmooSetup.DicAxisValue.YAxis.PinValueSet[axisY][i] + ":" +
                                                             shmooId.ShmooContentLVCC[i]);
                                            temphvcclist.Add(shmooSetup.DicAxisValue.YAxis.PinValueSet[axisY][i] + ":" +
                                                             shmooId.ShmooContentHVCC[i]);
                                        }

                                        var xyaxis = "X:" + axisX + ":" + axisXType + ",\r\n" + "Y:" + axisY + ":" +
                                                     axisYType;

                                        tmp2DLVCCHVCRow["Test Num"] = shmooSetup.TestNum;
                                        tmp2DLVCCHVCRow["Type"] = shmooSetup.Type;
                                        tmp2DLVCCHVCRow["Category"] = shmooSetup.Category;
                                        tmp2DLVCCHVCRow["Test Instance"] = shmooSetup.TestInstanceName;
                                        tmp2DLVCCHVCRow["Test Setup"] = shmooSetup.SetupName;
                                        tmp2DLVCCHVCRow["Lot ID"] = shmooId.LotId;
                                        tmp2DLVCCHVCRow["X,Y"] = shmooId.DieXY;
                                        tmp2DLVCCHVCRow["Pattern List"] = string.Join(",\r\n", shmooSetup.PatternList);
                                        tmp2DLVCCHVCRow["Payload"] = string.Join(",\r\n", shmooSetup.PayloadList);
                                        tmp2DLVCCHVCRow["X,Y Axis"] = xyaxis;
                                        tmp2DLVCCHVCRow["LVCC List"] = string.Join(",\r\n", templvcclist);
                                        tmp2DLVCCHVCRow["HVCC List"] = string.Join(",\r\n", temphvcclist);
                                        dtSum2DShmooLVCCHVCC.Rows.Add(tmp2DLVCCHVCRow);
                                    }
                                }

                                //if (dicShmSettings["Plot Overlay 2D"]) //只畫2D Overlay 我覺得好像沒甚麼必要
                                //{
                                //    allContent = ""; allDieInfo = "";
                                //}


                                //如果有要處理疊圖 在這裡把資訊加上去  0|10|50|100,0|20|60|100,...#   |分隔每一點(只有疊圖的才有) ,分隔每一條Y #分隔每一個Device
                                if (PlotShmooOverlay && !Merge2D2Inst && shmooSetup.ShmooIDs.Count > 1)
                                {
                                    var tempdictXMaxMinDict = new Dictionary<int, List<string>>();

                                    var rebuildOverlayStr2D = ""; //Rebuild Overlay 1D

                                    for (var y = 0; y < shmooSetup.StepCount[1]; y++)
                                    {
                                        var tmpStr = "";
                                        for (var x = 0; x < shmooSetup.StepCount[0]; x++)
                                        {
                                            tmpStr +=
                                                string.Format("{0:F0}",
                                                    dicOverlay2D[x][y] * 100 / shmooSetup.ShmooIDs.Count) +
                                                "|";

                                            if (!tempdictXMaxMinDict.ContainsKey(x))
                                                tempdictXMaxMinDict.Add(x, new List<string>());
                                            tempdictXMaxMinDict[x].Add(string.Format("{0:F0}",
                                                dicOverlay2D[x][y] * 100 / shmooSetup.ShmooIDs.Count));
                                        }

                                        rebuildOverlayStr2D += tmpStr.Remove(tmpStr.Length - 1) + ",";
                                    }


                                    if (OnlyOverlay2DShmoo)
                                    {
                                        allContent = rebuildOverlayStr2D.Remove(rebuildOverlayStr2D.Length - 1) + "#";
                                        allDieInfo = shmooSetup.ShmooIDs.Count +
                                                     " Devices' Overlayed Shmoo #";
                                    }
                                    else
                                    {
                                        allContent = rebuildOverlayStr2D.Remove(rebuildOverlayStr2D.Length - 1) + "#" +
                                                     allContent;
                                        allDieInfo = shmooSetup.ShmooIDs.Count +
                                                     " Devices' Overlayed Shmoo #" + allDieInfo;
                                    }

                                    // to find each point max Y and Min Y value

                                    foreach (var xPointData in tempdictXMaxMinDict)
                                    {
                                        var xIndex = xPointData.Key;
                                        var dictGroup = new Dictionary<int, List<int>>();
                                        var CurrentGroup = 1;
                                        for (var i = 0; i < xPointData.Value.Count; i++)
                                        {
                                            var currPercentage = xPointData.Value[i];

                                            if ((currPercentage == "100") &
                                                !dictGroup.Any(item => item.Value.Contains(i)))
                                            {
                                                if (!dictGroup.ContainsKey(CurrentGroup))
                                                    dictGroup.Add(CurrentGroup, new List<int>());
                                                dictGroup[CurrentGroup].Add(i);
                                            }

                                            var nextPercentage = "";
                                            if (i < xPointData.Value.Count - 1)
                                                nextPercentage = xPointData.Value[i + 1];
                                            else
                                                break;
                                            if ((currPercentage == "100") & (nextPercentage != "100")) CurrentGroup++;
                                        }

                                        if (dictGroup.Count > 0)
                                            XPointYMaxMinIndexRange.Add(xIndex,
                                                dictGroup.Values.OrderBy(x => x.Count).Last());
                                        else
                                            XPointYMaxMinIndexRange.Add(xIndex, new List<int>());
                                    }
                                }

                                if (Merge2D2Inst)
                                {
                                    var rebuildMergeStr2D = ""; //Rebuild Overlay 1D

                                    for (var y = 0; y < shmooSetup.StepCount[1]; y++)
                                    {
                                        var tmpStr = "";
                                        for (var x = 0; x < shmooSetup.StepCount[0]; x++)
                                            tmpStr += dicMerge2D2Inse[x][y] + "^";
                                        rebuildMergeStr2D += tmpStr.Remove(tmpStr.Length - 1) + ",";
                                    }

                                    allContent = rebuildMergeStr2D.Remove(rebuildMergeStr2D.Length - 1) + "#" +
                                                 allContent;
                                    allDieInfo = shmooSetup.ShmooIDs.Count +
                                                 " Devices' Merge Instance Shmoo #" + allDieInfo;
                                }
                            }
                            else
                            {
                                var allContextList = new List<string>();
                                var allDeciceInfo = new List<string>();


                                var totalContextValueDict = new Dictionary<int, Dictionary<int, int>>();
                                foreach (var shmooId in shmooSetup.ShmooIDs)
                                {
                                    allContextList.Add(string.Join(",", shmooId.ShmooContent));


                                    allDeciceInfo.Add(shmooId.MergeInstanceCnt + " Instances Merge Shmoo" +
                                                      " Site:" + shmooId.Site + " LotID:" + shmooId.LotId + "Die XY" +
                                                      shmooId.DieXY);

                                    for (var y = 0; y < shmooId.ShmooContent.Count; y++)
                                    {
                                        var xlist = shmooId.ShmooContent[y].Split('|').ToList();
                                        for (var x = 0; x < xlist.Count; x++)
                                        {
                                            if (!totalContextValueDict.ContainsKey(y))
                                                totalContextValueDict[y] = new Dictionary<int, int>();
                                            if (!totalContextValueDict[y].ContainsKey(x))
                                                totalContextValueDict[y][x] = 0;
                                            totalContextValueDict[y][x] += Convert.ToInt16(xlist[x]);
                                        }
                                    }
                                }


                                var totalContext = new List<string>();
                                foreach (var totalItem in totalContextValueDict)
                                {
                                    var xValues = totalItem.Value.Values.ToList(); //12,52,46,30
                                    var xValueStrList = new List<string>();
                                    foreach (var xValue in xValues)
                                        xValueStrList.Add((xValue / shmooSetup.ShmooIDs.Count).ToString());
                                    totalContext.Add(string.Join("|", xValueStrList));
                                }

                                allContent = string.Join(",", totalContext) + "#" + string.Join("#", allContextList) +
                                             "#";
                                allDieInfo = "#" + string.Join("#", allDeciceInfo) + "#";
                            }


                            tmpRowCategory["Die Info"] = allDieInfo.Remove(allDieInfo.Length - 1);
                            tmpRowCategory["All Content"] = allContent.Remove(allContent.Length - 1);

                            //For 2D Shmoo Summary統計用 統計每顆Die的資訊
                            var tmpRowShmSum2D = dtSum2D.NewRow(); //一個Setup一行
                            tmpRowShmSum2D["Test Num"] = shmooSetup.TestNum;
                            tmpRowShmSum2D["Type"] = shmooSetup.Type;
                            tmpRowShmSum2D["Category"] = shmooSetup.Category;
                            tmpRowShmSum2D["Test Instance"] = shmooSetup.TestInstanceName;
                            tmpRowShmSum2D["Test Setup"] = shmooSetup.SetupName;
                            tmpRowShmSum2D["Die Count"] = shmooSetup.ShmooIDs.Count;
                            tmpRowShmSum2D["Pattern List"] = string.Join("\r\n", shmooSetup.PatternList);
                            tmpRowShmSum2D["Payload"] = string.Join(",\r\n", shmooSetup.PayloadList);
                            tmpRowShmSum2D["Instance List"] = string.Join(",\r\n", shmooSetup.InstanceList);


                            if (PlotShmooOverlay && !Merge2D2Inst & (XPointYMaxMinIndexRange.Count > 0) &
                                (shmooSetup.SpecPoints[0].Count > 0)) // 有去判斷overlay 才需要進來做
                            {
                                // x axis point
                                var xAxisPin = Regex.Split(shmooSetup.Axis[0][shmooSetup.Axis[0].Count - 1], @":")[0];
                                var yAxisPin = Regex.Split(shmooSetup.Axis[1][shmooSetup.Axis[1].Count - 1], @":")[0];

                                var middleIndex =
                                    (shmooSetup.SpecPoints[0].Count - 1) /
                                    2; // hopefully, your list has an odd number of elements!
                                var sortPoints = shmooSetup.SpecPoints[0];
                                sortPoints.Sort();
                                var middlePoint = sortPoints[middleIndex];
                                var middleSpec = shmooSetup.DicAxisValue.XAxis.PinValueSet[xAxisPin][middlePoint];

                                foreach (var point in shmooSetup.SpecPoints[0])
                                {
                                    var xValue = shmooSetup.DicAxisValue.XAxis.PinValueSet[xAxisPin][point];
                                    try
                                    {
                                        var maxYValue = "N/A";
                                        var minYValue = "N/A";
                                        if (XPointYMaxMinIndexRange.ContainsKey(point) &
                                            (XPointYMaxMinIndexRange[point].Count > 0))
                                        {
                                            // index 與 實際 value 是反過來的 @o@
                                            var v1 = shmooSetup.DicAxisValue.YAxis.PinValueSet[yAxisPin][
                                                shmooSetup.StepCount[1] - 1 - XPointYMaxMinIndexRange[point].Max()];
                                            var v2 = shmooSetup.DicAxisValue.YAxis.PinValueSet[yAxisPin][
                                                shmooSetup.StepCount[1] - 1 - XPointYMaxMinIndexRange[point].Min()];
                                            if (v1 > v2)
                                            {
                                                maxYValue = v1.ToString();
                                                minYValue = v2.ToString();
                                            }
                                            else
                                            {
                                                maxYValue = v2.ToString();
                                                minYValue = v1.ToString();
                                            }
                                        }

                                        var resultStr = string.Format("X@ {0}, MaxY= {1},  MinY={2}", xValue, maxYValue,
                                            minYValue);
                                        if (middleSpec > xValue)
                                            tmpRowShmSum2D["Spec " + LgbRatioString] = resultStr;
                                        else if (middleSpec < xValue)
                                            tmpRowShmSum2D["Spec " + HgbRatioString] = resultStr;
                                        else
                                            tmpRowShmSum2D["Spec"] = resultStr;
                                    }
                                    catch (Exception ex)
                                    {
                                        throw new Exception(string.Format(
                                            "Get Max/Min Y axis value fail, please check your datalog {0}", ex));
                                    }
                                }
                            }

                            if (shmooSetup.ShmooIDs.Count > 0) //計算相同ShnooIDs passpoint / total point 
                                tmpRowShmSum2D["Pass Rate"] =
                                    Math.Round(
                                        passPoint / (shmooSetup.StepCount[0] * shmooSetup.StepCount[1] *
                                                     shmooSetup.ShmooIDs.Count * 1.0) * 100.0, 2,
                                        MidpointRounding.AwayFromZero);
                            dtSum2D.Rows.Add(tmpRowShmSum2D);
                        }

                        CurrShmooReport.Tables[category].Rows.Add(tmpRowCategory);
                        dtSetup.Rows.Add(tmpRowShmSetup);

                        #endregion
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Err", ex.Message);
                        throw;
                    }

                //args.Percentage = Convert.ToInt16(currCatCount * 100 / ListAllCategories.Count);
                //args.Result = string.Format("Adding " + category + " Done!!");
                //progress.Report(args);
                currCatCount++;
            }


            if (InstacneFreqMode) Gen1DFreqencyModeTable(lFreqencyList, lFreqMode1DShmooResult);


            if (!FindShmooHole1D) CurrShmooReport.Tables.Remove("ShmooHole_1D");
            if (CurrShmooReport.Tables["Summary_Abnormal"].Rows.Count == 0)
                CurrShmooReport.Tables.Remove("Summary_Abnormal");
            if (CurrShmooReport.Tables["ShmooAlarm"].Rows.Count == 0) CurrShmooReport.Tables.Remove("ShmooAlarm");
            if (CurrShmooReport.Tables["2D_LVCCHVCC"].Rows.Count == 0) CurrShmooReport.Tables.Remove("2D_LVCCHVCC");
            if (CurrShmooReport.Tables["SelSramDigSrcCheck"].Rows.Count == 0)
                CurrShmooReport.Tables.Remove("SelSramDigSrcCheck");
        }

        private void Gen1DFreqencyModeTable(List<string> lFreqencyList,
            Dictionary<string, Dictionary<string, List<FreqModeShmooId>>> lFreqMode1DShmooResult)
        {
            var dt1DFrequency = new DataTable("1DShmooFreqSummary");
            dt1DFrequency.Columns.Add(new DataColumn("Type", typeof(string)) { DefaultValue = "N/A" });
            dt1DFrequency.Columns.Add(new DataColumn("Category", typeof(string)) { DefaultValue = "N/A" });
            dt1DFrequency.Columns.Add(new DataColumn("Test Instance", typeof(string)) { DefaultValue = "N/A" });
            dt1DFrequency.Columns.Add(new DataColumn("Test Setup", typeof(string)) { DefaultValue = "N/A" });
            dt1DFrequency.Columns.Add(new DataColumn("Lot ID", typeof(string)) { DefaultValue = "N/A" });
            dt1DFrequency.Columns.Add(new DataColumn("Site", typeof(string)) { DefaultValue = "N/A" });
            dt1DFrequency.Columns.Add(new DataColumn("X,Y", typeof(string)) { DefaultValue = "N/A" });


            var regexDig = new Regex(@"(?<Freq>\d+)MHZ", RegexOptions.IgnoreCase);
            List<string> orderlFreqList = null;
            try
            {
                orderlFreqList =
                    lFreqencyList.OrderBy(p => Convert.ToInt16(regexDig.Match(p).Groups["Freq"].ToString())).ToList();
            }
            catch (Exception)
            {
                orderlFreqList = lFreqencyList;
            }


            foreach (var freqItem in orderlFreqList)
                dt1DFrequency.Columns.Add(new DataColumn(freqItem + " LVCC", typeof(double)));

            foreach (var freqItem in orderlFreqList)
                dt1DFrequency.Columns.Add(new DataColumn(freqItem + " HVCC", typeof(double)));

            CurrShmooReport.Tables.Add(dt1DFrequency);

            foreach (var lFreqModeCate in lFreqMode1DShmooResult)
            {
                var category = lFreqModeCate.Key;
                foreach (var lInstanceItem in lFreqModeCate.Value)
                {
                    var instName = lInstanceItem.Key.Split(':')[0];
                    var setUp = lInstanceItem.Key.Split(':')[1];
                    foreach (var lFreqModeId in lInstanceItem.Value)
                    {
                        var tmp1DFreq = dt1DFrequency.NewRow();
                        tmp1DFreq["Type"] = "1D";
                        tmp1DFreq["Category"] = category;
                        tmp1DFreq["Test Instance"] = instName;
                        tmp1DFreq["Test Setup"] = setUp;
                        tmp1DFreq["Lot ID"] = lFreqModeId.LotId;
                        tmp1DFreq["Site"] = lFreqModeId.Site;
                        tmp1DFreq["X,Y"] = lFreqModeId.DieXY;
                        foreach (var freqItem in lFreqencyList)
                        {
                            if (lFreqModeId.FreqLvccDict.ContainsKey(freqItem))
                                tmp1DFreq[freqItem + " LVCC"] = lFreqModeId.FreqLvccDict[freqItem];
                            if (lFreqModeId.FreqHvccDict.ContainsKey(freqItem))
                                tmp1DFreq[freqItem + " HVCC"] = lFreqModeId.FreqHvccDict[freqItem];
                        }

                        dt1DFrequency.Rows.Add(tmp1DFreq);
                    }
                }
            }
        }

        private List<ShmooId> ByPassGrrNoiseDieByCriteria(List<ShmooId> shmooIDs)
        {
            if (!GRRLoopByPassErrCount)
                return shmooIDs;

            var ErrCode9999Count = 0;
            var ErrCode5555Count = 0;

            foreach (var shmooID in shmooIDs)
            {
                if (shmooID.Hvcc == 9999 || shmooID.Lvcc == -9999)
                    ErrCode9999Count++;
                if (shmooID.Hvcc == 5555 || shmooID.Lvcc == -5555)
                    ErrCode5555Count++;
            }

            if (ErrCode9999Count < UserDefLessErr9999Count)
                shmooIDs = shmooIDs.Where(shmooID => shmooID.Hvcc != 9999 && shmooID.Lvcc != -9999).ToList();


            if (ErrCode5555Count < UserDefLessErr5555Count)
                shmooIDs = shmooIDs.Where(shmooID => shmooID.Hvcc != 5555 && shmooID.Lvcc != -5555).ToList();


            return shmooIDs;
        }

        private static bool FindShmooHoleIn1D(List<string> shmooContent)
        {
            //var rgexShmPass = new Regex(@"\+|\*|P|p", RegexOptions.Compiled); //已經轉化完了
            //var rgexShmFail = new Regex(@"\-|\~|F|f", RegexOptions.Compiled);

            //簡單說就是有暫態的轉換PFP or FPF  ++-+-
            var transition = 0;
            var init = shmooContent[0];
            for (var i = 1; i < shmooContent.Count; i++)
            {
                if (shmooContent[i] == init) continue;
                init = shmooContent[i];
                transition++;
            }

            return transition >= 2;
        }
    }

    public class ShmooSetup //以TestInstanceName::SetupName當Key 
    {
        public double[][]
            AcurateStep = { new[] { 0.0, 0.0, 0.0 }, new[] { 0.0, 0.0, 0.0 } }; //計算X Y真實的LVCC HVCC用 因為[]內的Step不準

        public List<string>[]
            Axis = new List<string>[2] { new List<string>(), new List<string>() }; //"VDD_SOC" , "VDD_CORE" 對應SettingsX的順序

        public string[] AxisType = new string[2] { @"N/A", @"N/A" }; //V or MHz
        public string BinCutPerfromanceName = "";
        public int[] BinCutSpecPoints = new int[2]; //  BinCutSpecPoints[0] : cpmax  BinCutSpecPoints[1] : cpmin
        public double[] BinCutSpecValue = new double[2];

        public string Category = "NONE";

        //public Dictionary<string, List<double>> DicAxisValue = new Dictionary<string, List<double>>(); //VDD_SOC -> 0 = 1.33
        public AxisDictValue DicAxisValue = new AxisDictValue();
        public string ForceCondition;
        public string FreeRunningClk = "";
        public HashSet<string> InstanceList = new HashSet<string>();
        public bool IsMergeBySiteSetup;

        public bool IsOtherRailSpec;
        //要能滿足除了Shmoo內容以外的所有需要, 最少要有一個ID!!

        //一個Setup 加上 一堆Shmoo
        // 1D Shmoo
        //[Char,-0,16,7,V,0,XI0=24000000,SPI_Non_DDR_Sc19Mode1_NV,SPI_VDD_CPU_SRAM_P1,1325,
        //.\Pattern\SPI\FIJI_index19_M1_0814_modify_GLC40_mod_scenario19.PAT,
        //NV,VDD_CPU=0.950,VDD_GPU=0.950,VDD_FIXED=0.950,VDD_GPU_SRAM=0.950,VDD_CPU_SRAM=0.950,VDD_VAR_SOC=0.950,
        //VDD_CPU_SRAM=0.950:0.500:-0.010,----------------------------------------------,NH,N/A,N/A]

        // 2D Shmoo
        //[V,0,XI0=24000000,-0,0,0,
        //DD_FIJA0_L_FULP_CH_CLUA_SAA_UNC_AUT_ALLFV_1308041227_S100_1011_SM_NV,GPU_Scan_SA_VddGpu_vs_LdcSAFreq,1065,
        //.\pattern\LdcScan\DD_FIJA0_L_FULP_CH_CLUA_SAA_UNC_AUT_ALLFV_1308041227_S100_1011_SM.pat,
        //NV,VDD_GPU=0.950,
        //X@LdcSA_Freq=10000000.000:30000000.000:1000000.000,Y@VDD_GPU=0.400:1.600:0.050]


        public bool IsSameSetupMerge;
        //public bool[] NeedAlignHighLowVcc = new bool[2] { false, false };  // 20180814 by JN comment it , no usage 

        public List<string> PatternList = new List<string>();
        public List<string> PayloadList = new List<string>();

        // bin cut spec
        public bool PlotBinCutSpec;


        public List<string> SettingsInit = new List<string>(); //As Spec
        public List<string> SettingsX = new List<string>(); //列印Setup Sheet用
        public List<string> SettingsY = new List<string>(); //列印Setup Sheet用
        public string SetupName = "NONE";

        public List<ShmooId> ShmooIDs = new List<ShmooId>(); //或者用Dictionary?

        public double[] Spec = new double[2] { -1, -1 }; //標示Spec用 第一個X

        public string Special = "NONE"; //Retention用1D Shmoo的Format...., INVERSE代表XY軸互換的Shmoo

        public List<int>[]
            SpecPoints = new List<int>[2] { new List<int>(), new List<int>() }; //"VDD_SOC" , "VDD_CORE" 對應SettingsX的順序

        public int[] StepCount = new int[2] { 1, 1 }; //由Shmoo Content算出來的  X:0 Y:1
        public string TestInstanceName = "NONE";
        public int TestNum;
        public string TimeSetInfo;


        public string Type = "NONE"; //1D or 2D
        public string UniqeName = "NONE"; //同一個Device內的Unique Name不會重複!! 可以解決以Test Num為Key的問題!!
    }

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

    public class FreqModeShmooId
    {
        public string DieXY = "NONE";
        public Dictionary<string, double> FreqHvccDict = new Dictionary<string, double>();

        public Dictionary<string, double> FreqLvccDict = new Dictionary<string, double>();
        public string LotId = "NONE";
        public string Site = "0"; //反正只是列印用 不用宣告int
    }

    public class AxisDictValue
    {
        public AxisData XAxis = new AxisData("X");
        public AxisData YAxis = new AxisData("Y");
    }

    public class AxisData
    {
        public string AxisName;
        public Dictionary<string, List<double>> PinValueSet = new Dictionary<string, List<double>>();
        public string Unit = "";


        public AxisData(string axisName)
        {
            AxisName = axisName;
        }

        public void PointUnitConvert()
        {
            var lBased = 0.0;
            var lUnit = "";
            if (PinValueSet.Any())
            {
                var checkValue = PinValueSet.First().Value.FirstOrDefault();

                if (checkValue > 1)
                    return;

                if (0.001 > checkValue && checkValue > 0.000001) // u
                {
                    lUnit = "u";
                    lBased = 0.000001;
                }
                else if (0.000001 > checkValue && checkValue > 0.000000001) // n
                {
                    lUnit = "n";
                    lBased = 0.000000001;
                }
                else if (0.000000001 > checkValue && checkValue > 0.000000000001) // p 
                {
                    lUnit = "p";
                    lBased = 0.000000000001;
                }
                else if (0.000000000001 > checkValue && checkValue > 0.000000000000001) // f 
                {
                    lUnit = "f";
                    lBased = 0.000000000000001;
                }
                else
                {
                    return;
                }

                var reNew = new Dictionary<string, List<double>>();
                foreach (var pinItem in PinValueSet)
                {
                    if (!reNew.ContainsKey(pinItem.Key))
                        reNew.Add(pinItem.Key, new List<double>());
                    foreach (var value in pinItem.Value)
                        reNew[pinItem.Key].Add(Math.Round(value / lBased, 3, MidpointRounding.AwayFromZero));
                    // Math.Round(startPoint + currStepSize * i, 3,MidpointRounding.AwayFromZero)
                }

                PinValueSet = reNew;
                Unit = lUnit;
            }
        }
    }
}