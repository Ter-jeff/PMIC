using CommonLib.Enum;
using NLog;
using ShmooLog.Base;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ShmooLog
{
    public class ShmooLog
    {
        public static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly List<DeviceTestedSummary> _deviceTestedSummaries = new List<DeviceTestedSummary>();

        private readonly Regex _regexDevice = new Regex(@"^\s+Device#:", RegexOptions.Compiled);

        private readonly Regex _regexNumbers =
            new Regex(@"^ \d+\s+\d\s+|^ \d+\s+\d\d\s+|^ \d+\s+\d\d\d\s+", RegexOptions.Compiled);

        private readonly Regex _regexSepArray = new Regex(@"\s+|\t+", RegexOptions.Compiled);
        private readonly Regex _regexTestInstance = new Regex(@"^<\w+>", RegexOptions.Compiled);
        public readonly string FilePath;

        public ConcurrentDictionary<int, DataTable>
            ConcurrentDataTable = new ConcurrentDictionary<int, DataTable>(); //Key = DeviceNo, DataTable依需求訂製Column

        public ConcurrentDictionary<int, Device>
            ConcurrentDicDevice = new ConcurrentDictionary<int, Device>(); //Key = DeviceNo

        public List<int> CurrentDeviceNoList = new List<int>(); //用來當ConcurrentDictionary的Key

        public bool HasDuplicateTestInstanceSetup = false; //判定為True之後 這個檔案就廢了 因為Shmoo Excel以Test Instance + Setup為Key

        public List<string> ListOfColumnsPerMode = new List<string>
        {
            "Device#:int:0",
            "Site:int:0",
            "Test Instance:string:N/A",
            "Setup Name:string:N/A",
            "Shmoo Type:string:N/A",
            "Shmoo Setup:string:N/A",
            "Shmoo Content:string:N/A",
            "X Axis Unit:string:N/A",
            "Y Axis Unit:string:N/A",
            "LVCC Value:string:N/A",
            "HVCC Value:string:N/A",
            "ForceCondition:string:N/A",
            "Activity TimeSet:string:N/A"
        };

        public List<string> ListWarningMsg = new List<string>(); //收集所有的Warning Message

        public string Lot = "N/A"; //N9B889-15
        public SramDef SramDef = new SramDef();

        public ShmooLog(string log)
        {
            FilePath = log;
        }

        public List<DeviceTestedSummary> GetShmooDeviceTestedSummaries
        {
            get
            {
                foreach (var deviceSum in _deviceTestedSummaries)
                {
                    var dieInfo = ConcurrentDicDevice[deviceSum.DeviceNo];
                    deviceSum.DieXY = dieInfo.DieXY;
                }

                return _deviceTestedSummaries;
            }
        }

        public long FileSize { get; set; }
        public string DirectoryPath { get; set; }
        public string FileName { get; set; }
        public string UniqueName { get; set; }
        public string JobName { get; set; }
        public string ProgramName { get; set; }

        public void ParseEachDevices()
        {
            if (!File.Exists(FilePath)) return;

            var fInfo = new FileInfo(FilePath);
            FileName = fInfo.Name;
            if (fInfo.Directory != null)
                DirectoryPath = fInfo.Directory.ToString();
            FileSize = fInfo.Length;

            //最花時間的就是Compile Regex!!! 所以要寫在迴圈外
            var rgexWarningMsg = new Regex(@"alarm|warning|error|fatal", RegexOptions.Compiled);
            var rgexDevice = new Regex(@"^\s+Device#:", RegexOptions.Compiled);

            var sb = new StringBuilder();
            var lines = File.ReadLines(FilePath).ToList();

            for (var index = 0; index < lines.Count; index++)
            {
                var line = lines[index];
                if (rgexDevice.IsMatch(line))
                {
                    ParseHeaderInfo(lines.GetRange(0, index));
                    break;
                }
            }


            for (var index = 0; index < lines.Count; index++)
            {
                var line = lines[index];
                if (string.IsNullOrEmpty(line))
                    continue;
                if (rgexDevice.IsMatch(line)) //看到Device#: 先把上一顆的加入List 
                {
                    CurrentDeviceNoList.AddRange(ParseDeviceNumber(line));
                    ConsumerModeOfEachDevice(lines.GetRange(index, lines.Count - index));
                }
                else
                {
                    sb.AppendLine(line);
                }

                if (rgexWarningMsg.IsMatch(line)) ListWarningMsg.Add(line);
            }

            foreach (var msg in ListWarningMsg.Distinct()) //利用Task處理完的時間做
            {
                //Logger.Warn(msg + "\n");
            }
        }

        public int ParseTestInstanceData(List<string> tmpArray, DataTable[] currDataTables,
            Dictionary<int, int> dicActiveSiteDeviceNum, List<int> currActiveSiteNum)
        {
            #region 宣告Regex.Compile

            var rgexTestSummary =
                new Regex(@"Site    Sort|Site    X_Coord     Y_Coord|Site Failed tests/Executed tests",
                    RegexOptions.Compiled);
            var rgexShmHint1 = new Regex(@"^\[", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            var rgexShmHint2 = new Regex(@"\\pattern\\|\.pat|[\*|\+|\~|\-]{2,}|\~",
                RegexOptions.Compiled | RegexOptions.IgnoreCase);
            var rgexShmTimeSet = new Regex(@"\[Activity_Timing_Sheet", RegexOptions.Compiled);
            var rgexShmForceCond = new Regex(@"\[Force_condition_during_shmoo",
                RegexOptions.Compiled | RegexOptions.IgnoreCase); // add for get timeset name
            var rgexShmStart = new Regex(@"^\[Start_Shmoo", RegexOptions.Compiled);
            var rgexShmSite = new Regex(@"Site           Pattern\(s\)", RegexOptions.Compiled);
            var rgexShmContent = new Regex(@"[\+\-A\*\~PFpfE_]+", RegexOptions.Compiled); // . 不支援
            var rgexNotShmContent = new Regex(@"X|Y|:|\d+|M+|V+|m+", RegexOptions.Compiled);
            var rgexInverseShm2D = new Regex(@"Y\@.+X\@.+", RegexOptions.Compiled); //區別先跑Y再跑X的情況
            var rgexShmDummyContent = new Regex(@"CharIP,", RegexOptions.Compiled); // . 不支援
            var regex2DFinalContest = new Regex(@"(X\@.+Y\@.+)|(Y\@.+X\@.+)", RegexOptions.Compiled); //最後的2D 資訊line
            var regexSramDef = new Regex(@"\[SELSRM_Def,", RegexOptions.Compiled);
            var regexSramData = new Regex(@"\[SELSRM_Condition,", RegexOptions.Compiled);
            var shmContentStart = "Y Axis:";
            var shmContentEnd = "X Axis:";

            #endregion

            var binOutlineStart = 0; //偵測Bin Out資訊的位置
            var keyLineNumber = 0;
            var keyLineTitle = "N/A";

            var currShmooType = "N/A"; //1D or 2D
            var currShmooSetupName = "N/A";
            var currShmooTestInstance = "N/A";
            var currShmooContentSite = 0;
            var currShmooContent = new List<string>();
            var currShmooLvcc = new Dictionary<int, List<string>>();
            var currShmooHvcc = new Dictionary<int, List<string>>();
            var dicSiteShmooSetup = new Dictionary<int, string>();
            var currYaxisUnit = "N/A";

            var collectShmooContent = false;
            int resI;
            var forceCondition = "";
            var timeSetInfo = "";

            for (var i = 1; i < tmpArray.Count; i++) //DataLog的每一行, 從1開始
            {
                var currLine = tmpArray[i].Trim();

                #region 處理Shmoo Setup類型的資料

                if (rgexShmStart.IsMatch(currLine))
                {
                    forceCondition = "";
                    timeSetInfo = "";
                }

                if (regexSramDef.IsMatch(currLine))
                {
                    SramDef.HasSramDef = true;
                    SramDef.InitialMap(currLine);
                }

                if (regexSramData.IsMatch(currLine)) SramDef.AddData(currLine);

                // Force Condition
                if (rgexShmForceCond.IsMatch(currLine) && forceCondition.Length == 0)
                {
                    forceCondition = currLine.Replace(@"[Force_condition_during_shmoo:", "");
                    forceCondition = forceCondition.Replace(@"]", "");
                }

                // Timing Set
                if (rgexShmTimeSet.IsMatch(currLine) && timeSetInfo.Length == 0)
                {
                    timeSetInfo = currLine.Replace(@"[Activity_Timing_Sheet:", "");
                    timeSetInfo = timeSetInfo.Replace(@"]", "");
                }


                if (rgexShmHint1.IsMatch(currLine) && rgexShmHint2.IsMatch(currLine)) //特例 Shmoo Item 
                {
                    var shmooContent = currLine;
                    var iA = Regex.Split(shmooContent.Replace("[", ""), @",");
                    if (iA.Count() < 8) continue;
                    currShmooType = Regex.IsMatch(currLine, @"Y\@") ? "2D" : "1D";

                    if (currShmooType == "2D")
                        shmooContent = rgexShmDummyContent.Replace(tmpArray[i], ""); //完全不能理解當初的目的?


                    var currSite = 0;
                    if (currShmooType == "1D" || !regex2DFinalContest.IsMatch(currLine))
                    {
                        if (int.TryParse(iA[5], out resI)) currSite = resI;
                    }
                    else
                    {
                        if (int.TryParse(iA[1], out resI)) currSite = resI;
                    }

                    //因為1D 2D的格式不一樣 當初VBT沒有想清楚
                    currShmooTestInstance = currShmooType == "1D" ? iA[7] : iA[6];
                    currShmooSetupName = currShmooType == "1D" ? iA[8] : iA[7];

                    dicSiteShmooSetup[currSite] = shmooContent; //tmpArray[i];

                    if (!regex2DFinalContest
                            .IsMatch(currLine)) // beacuse 2D is shmoo set up , 2D shmoo is combined by 1D shmoo
                    {
                        if (!currShmooLvcc.ContainsKey(currSite)) currShmooLvcc.Add(currSite, new List<string>());

                        if (!currShmooHvcc.ContainsKey(currSite)) currShmooHvcc.Add(currSite, new List<string>());
                        currShmooLvcc[currSite].Add(iA[iA.Count() - 2]);
                        currShmooHvcc[currSite].Add(iA[iA.Count() - 1].Replace("]", ""));
                    }

                    currShmooContentSite = 0;
                    currShmooContent.Clear();

                    shmContentStart = "Y Axis:";
                    shmContentEnd = "X Axis:";

                    if (currShmooType == "2D" && rgexInverseShm2D.IsMatch(currLine))
                    {
                        shmContentStart = "X Axis:";
                        shmContentEnd = "Y Axis:";
                    }
                    else
                    {
                        if (iA[0] == "CharRet") //Must Be 1D
                        {
                            var dtRow = currDataTables[currSite].NewRow();
                            dtRow["Device#"] = dicActiveSiteDeviceNum[currSite];
                            dtRow["Site"] = currSite;
                            dtRow["Test Instance"] = currShmooTestInstance;
                            dtRow["Setup Name"] = currShmooSetupName;
                            dtRow["Shmoo Type"] = currShmooType;
                            dtRow["Shmoo Setup"] = dicSiteShmooSetup[currSite];
                            dtRow["Shmoo Content"] = iA[iA.Length - 4];
                            dtRow["ForceCondition"] = forceCondition;
                            dtRow["Activity TimeSet"] = timeSetInfo;
                            currDataTables[currSite].Rows.Add(dtRow);
                        }
                    }
                }

                #endregion

                if (rgexShmSite.IsMatch(currLine)) //Site           Pattern(s)          X Pin(s)       Z
                {
                    keyLineTitle = "ShmooSite";
                    keyLineNumber = i + 1;
                    continue;
                }

                if (keyLineNumber == i)
                {
                    switch (keyLineTitle) //看看前一行標註的是哪一種類型!
                    {
                        case "ShmooSite":
                            currShmooContentSite = Convert.ToInt16(currLine.Split(' ')[0]);
                            collectShmooContent = true;
                            break;
                    }

                    continue;
                }

                #region 處理Shmoo Content類型的資料

                if (collectShmooContent && rgexShmContent.IsMatch(currLine)) //無論1D 2D都一樣由下面收
                {
                    var tmpContent = currLine.Split(' ')[0];

                    if (!rgexNotShmContent.IsMatch(tmpContent))
                    {
                        if (currShmooType == "1D" && currShmooContent.Count == 0)
                            currShmooContent.Add(tmpContent);
                        else if (currShmooType == "2D")
                            if (currLine.Split(' ').Count() >= 2)
                                currShmooContent.Add(tmpContent);

                        //為了避免以下的情況 誤判 單獨存在的 "-" 會躲過   <------  留待日後抓完Content之後再檢查Length格式
                        //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                        //                                                                                -
                        //322222222222222222222111111111111111111110000000000000000000-0------------------1
                        //0998877665544332211009988776655443322110099887766554433221101-1122334455667788990
                        //050505050505050505050505050505050505050505050505050505050505750505050505050505050
                        //.................................................................................
                        //000000000000000000000000000000000000000000000000000000000000300000000000000000000
                        //mmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmmammmmmmmmmmmmmmmmmmmm
                        //VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
                        //X Axis: Vih(Level)
                    }
                }

                if (Regex.IsMatch(currLine, shmContentStart)) //收集2D Shmoo Y軸的單位
                {
                    //Y Axis: VMain(Level)
                    currYaxisUnit = currLine.Split('(')[1].Replace(")", "");
                    continue;
                }

                if (Regex.IsMatch(currLine, shmContentEnd)) //1D 2D Shmoo結尾 準備加入Die 然後宣告新的, 收集2D Shmoo Y軸的單位
                {
                    collectShmooContent = false;

                    //X Axis: XI0_Freq(AC Spec)
                    var currXaxisUnit = currLine.Split('(')[1].Replace(")", "");

                    //有抓到跑的Shmoo Content結果卻沒有對應的Setup? 主要是預防那種先有Content才印Setup的情況
                    if (!dicSiteShmooSetup.ContainsKey(currShmooContentSite))
                    {
                        Logger.Debug(currShmooTestInstance + " : " + currShmooSetupName +
                                     "has Shmoo Content without Shmoo Setup?");
                        continue;
                    }

                    if (currShmooContent.Count == 0) continue; //理論上應該要Show Warning Message!!

                    var dtRow = currDataTables[currShmooContentSite].NewRow();
                    dtRow["Device#"] = dicActiveSiteDeviceNum[currShmooContentSite];
                    dtRow["Site"] = currShmooContentSite;
                    dtRow["Test Instance"] = currShmooTestInstance;
                    dtRow["Setup Name"] = currShmooSetupName;
                    dtRow["Shmoo Type"] = currShmooType;
                    dtRow["Shmoo Setup"] = dicSiteShmooSetup[currShmooContentSite];
                    dtRow["Shmoo Content"] = string.Join(",", currShmooContent);
                    dtRow["X Axis Unit"] = currXaxisUnit;
                    dtRow["Y Axis Unit"] = currYaxisUnit;

                    dtRow["ForceCondition"] = forceCondition;
                    forceCondition = ""; // must reset it , othwise it can not be update 20190610 by JN
                    dtRow["Activity TimeSet"] = timeSetInfo;
                    dtRow["LVCC Value"] = string.Join(",", currShmooLvcc[currShmooContentSite]);
                    dtRow["HVCC Value"] = string.Join(",", currShmooHvcc[currShmooContentSite]);


                    currDataTables[currShmooContentSite].Rows.Add(dtRow);
                    dicSiteShmooSetup.Remove(currShmooContentSite);

                    currShmooContent.Clear();
                    currShmooLvcc[currShmooContentSite].Clear();
                    currShmooHvcc[currShmooContentSite].Clear();

                    // for record testinstance and setup
                    var deviceNum = dicActiveSiteDeviceNum[currShmooContentSite];
                    var deviceTestedSum = _deviceTestedSummaries.FirstOrDefault(p => p.DeviceNo.Equals(deviceNum));
                    if (deviceTestedSum == null)
                    {
                        deviceTestedSum = new DeviceTestedSummary(deviceNum);
                        _deviceTestedSummaries.Add(deviceTestedSum);
                    }

                    if (!deviceTestedSum.ShmooSetupTestedInst.ContainsKey(currShmooSetupName))
                        deviceTestedSum.ShmooSetupTestedInst.Add(currShmooSetupName, new HashSet<string>());
                    deviceTestedSum.ShmooSetupTestedInst[currShmooSetupName].Add(currShmooTestInstance);


                    continue;
                }

                #endregion

                if (rgexTestSummary.IsMatch(tmpArray[i])) //從這行開始就是Device的Bin Summary處!!
                {
                    binOutlineStart = i;
                    break; //進入Bin Summary階段
                }
            }

            return binOutlineStart;
        }

        private static IEnumerable<int> ParseDeviceNumber(string line) //Device#: 1-4
        {
            var devList = new List<int>();
            var devices = line.Replace("Device#:", "").Trim().Split(',');
            foreach (var dev in devices)
                if (Regex.IsMatch(dev, @"-"))
                {
                    var tmpAry = dev.Split('-');
                    for (var t = Convert.ToInt16(tmpAry[0]); t <= Convert.ToInt16(tmpAry[1]); t++)
                        devList.Add(t); //沒有檢查能不能轉INT有點風險
                }
                else
                {
                    devList.Add(Convert.ToInt16(dev)); //沒有檢查能不能轉INT有點風險
                }

            return devList;
        }

        private void ParseHeaderInfo(List<string> lines)
        {
            for (var i = 0; i < lines.Count; i++)
            {
                var iArray = Regex.Split(lines[i], @"\s+|\t+"); //以space or \t Split
                if (Regex.IsMatch(lines[i], "Prog Name"))
                    ProgramName = Regex.Replace(iArray[iArray.Length - 1], @"(.xlsm)|(.xls)", "");
                if (Regex.IsMatch(lines[i], "Job Name"))
                    JobName = iArray[iArray.Length - 1];
                if (Regex.IsMatch(lines[i], "Lot")) Lot = iArray[iArray.Length - 1];
            }

            if (ProgramName != "N/A" && JobName != "N/A")
                UniqueName = ProgramName + "__" + JobName;
        }

        private void ConsumerModeOfEachDevice(List<string> lines) //每次從Producer那邊取"一次Touch Down"的資料來處理
        {
            #region 抓Device Number

            var currDeviceNum = new List<int>(); //Device - Site 順序應該有對應
            var currActiveSiteNum = new List<int>(); //由第一個TestInstance判斷(通常是OP/SH)
            //*******************************************************************
            //先抓Device Number 理論上 同DataLog裡面應該也是不能重複!!
            //*******************************************************************
            if (_regexDevice.IsMatch(lines[0])) //    Device#: 1,2
            {
                currDeviceNum.AddRange(ParseDeviceNumber(lines[0]));
            }
            else
            {
                Logger.Error("[" + EnumNLogMessage.Input + "] " + "Datalog Format Error,1st Line: " + lines[0] +
                             " is not Device#:?");
                return;
            } //第一行不是Device#: 就去掉了

            #endregion

            #region 抓Active的Site

            //*******************************************************************
            //從第一個Test Instance抓到Active的Site!! 理論上應該是Open/Short
            //*******************************************************************
            var testSequence = 0; //紀錄測試順序
            for (var i = 1; i < lines.Count; i++)
            {
                var iArray = _regexSepArray.Split(lines[i].Trim()); //每一行內的分割

                if (_regexTestInstance.IsMatch(lines[i]))
                {
                    if (testSequence < 10) testSequence++; //第二個TestInstance就跳掉
                    else break;
                }

                //Collect Site Number
                if (iArray.Length >= 4 &&
                    _regexNumbers.IsMatch(lines[i])) //原本寫 >= 6是對應expand輸出 不過有些shmoo log為了節省記憶體會用simple
                {
                    var siteNum = Convert.ToInt16(iArray[1]);
                    if (!currActiveSiteNum.Contains(siteNum)) currActiveSiteNum.Add(siteNum);
                }
            }

            #endregion

            #region 比對Active Site與Device Number有沒有一致

            //*******************************************************************
            //如果SiteCount跟DeviceCount不一樣 嚴重Fail!!
            //*******************************************************************
            if (currActiveSiteNum.Count != currDeviceNum.Count)
            {
                Logger.Error(
                    "[" + EnumNLogMessage.Input + "] " + "Parse Error,{0}: Site Number {1} != Device Number {2}",
                    lines[0], currActiveSiteNum.Count, currDeviceNum.Count);
                return;
            }

            #endregion

            #region 宣告每個Device的對應關係及DataTable

            //*******************************************************************
            //建立Active Site 每一個Site都要宣告 
            //因為有可能Site跑1,3 不能只宣告兩個 否則就不能用site當array index
            //*******************************************************************
            var maxSiteNum = currActiveSiteNum.Max() + 1; //執行的site有可能是  0 3 4 所以不能用個數 要用Max來宣告
            var dicActiveSiteDeviceNum = new Dictionary<int, int>(); //建立Site -> Device的對應關係

            var currDevices = new Device[maxSiteNum]; //一切以Site為準
            var currDataTables = new DataTable[maxSiteNum]; //每個Device一個

            foreach (var activeSite in currActiveSiteNum) //只針對有Active的Site宣告就好 可以省一點空間!
            {
                currDevices[activeSite] = new Device(activeSite, currDeviceNum[0]); //因為是順序對應 最後一行有砍掉 所以Always 0

                //var deviceLabel = currDeviceNum[0].ToString();
                dicActiveSiteDeviceNum[activeSite] = currDeviceNum[0]; //建立Site -> Device的對應關係
                currDataTables[activeSite] =
                    InitializaDataTableOfEachModePerDevice(currDeviceNum[0].ToString()); //Device Number不能重複
                currDeviceNum.RemoveAt(0); //因為是順序對應 
            }

            #endregion

            #region 從頭開始掃 核心功能 根據Mode不同 決定處理的Method~

            //*******************************************************************
            //從頭開始掃
            //*******************************************************************
            var binOutlineStart =
                ParseTestInstanceData(lines, currDataTables, dicActiveSiteDeviceNum,
                    currActiveSiteNum); //處理每一行的資訊存到DataTable 然後回傳Bin Out資訊的位置

            #endregion

            #region 處理 Bin Summary

            //*******************************************************************
            //進入Bin Summary階段
            //*******************************************************************
            ParseTestResultPerDevice(lines, binOutlineStart, currDevices);

            #endregion

            #region 把結果掛上去Concurrent類型的Data Structure!!

            //*******************************************************************
            //進入收尾階段
            //*******************************************************************
            foreach (var activeSite in currActiveSiteNum) //這邊用Concurrent資料類型比較保險 因為有其他Task可能會Update
            {
                ConcurrentDicDevice[currDevices[activeSite].DeviceNo] =
                    currDevices[activeSite]; //由Device Number去操弄DataTable
                //if (LocalSpecs.UseDb)
                //{}  //LogDb.SaveNewDataTable(currDataTables[activeSite]);
                //else
                ConcurrentDataTable[currDevices[activeSite].DeviceNo] = currDataTables[activeSite];
            }

            //foreach (DataTable dt in currDataTables)
            //{
            //    if (dt != null)
            //        if (Convert.ToInt16(dt.TableName.ToString())<=6)
            //        {
            //            var dataSet = new DataSet();
            //            dataSet.Tables.Add(dt);
            //            dataSet.WriteXml(@"D:\Support\Gib_Pat\"+Convert.ToInt16(dt.TableName.ToString())+@"-old.xml");
            //        }
            //}

            #endregion
        }

        private void
            ParseTestResultPerDevice(List<string> perTouchDownLog, int binOutlineStart,
                Device[] currDevices) //最後Bin Summary的整理
        {
            var rgexTestSummary =
                new Regex(@"Site    Sort|Site    X_Coord     Y_Coord|Site Failed tests/Executed tests",
                    RegexOptions.Compiled);
            var keyLineNumber = 0;
            var keyLineTitle = "N/A";

            //最後Bin Summary階段
            for (var i = binOutlineStart; i < perTouchDownLog.Count; i++) //DataLog的每一行
            {
                if (rgexTestSummary.IsMatch(perTouchDownLog[i]))
                {
                    keyLineNumber = i + 2;
                    switch (perTouchDownLog[i])
                    {
                        case " Site    Sort     Bin":
                            keyLineTitle = "Bin";
                            break;
                        case " Site Failed tests/Executed tests":
                            keyLineTitle = "Executed Test";
                            break;
                        case " Site    X_Coord     Y_Coord":
                            keyLineTitle = "DieXY";
                            break;
                    }

                    continue;
                }

                if (keyLineNumber == i) //前面有特定指標需要到後一行處理的 很煩
                {
                    var tmp = Regex.Split(perTouchDownLog[i].Trim(), @"\s+");
                    if (tmp.Length != 3) continue;

                    switch (keyLineTitle) //看看前一行標註的是哪一種類型!
                    {
                        case @"Executed Test":
                            // Site Failed tests/Executed tests
                            currDevices[Convert.ToInt16(tmp[0])].FailedTest = Convert.ToInt32(tmp[1]);
                            currDevices[Convert.ToInt16(tmp[0])].ExecutedTest = Convert.ToInt32(tmp[2]);
                            keyLineNumber++; //因為他是一行一行印Site Sort Bin 所以
                            break;

                        case @"Bin":
                            //Site    Sort     Bin
                            if (tmp[1] == "N/A") tmp[1] = "0";
                            if (tmp[2] == "N/A") tmp[2] = "0";
                            currDevices[Convert.ToInt16(tmp[0])].Sort = Convert.ToInt32(tmp[1]);
                            currDevices[Convert.ToInt16(tmp[0])].Bin = Convert.ToInt32(tmp[2]);
                            keyLineNumber++; //因為他是一行一行印Site Sort Bin 所以
                            break;

                        case @"DieXY":
                            //Site    X     Y
                            int resA;
                            if (int.TryParse(tmp[1], out resA)) //不能保證X Y的資訊一定有 因為讀不到會表示N/A
                            {
                                currDevices[Convert.ToInt16(tmp[0])].X = Convert.ToInt16(tmp[1]);
                                currDevices[Convert.ToInt16(tmp[0])].Y = Convert.ToInt16(tmp[2]);
                            }
                            else
                            {
                                currDevices[Convert.ToInt16(tmp[0])].X = -999;
                                currDevices[Convert.ToInt16(tmp[0])].Y = -999;
                            }

                            keyLineNumber++; //因為他是一行一行印Site Sort Bin 所以
                            break;
                    }
                }
            }
        }

        private DataTable InitializaDataTableOfEachModePerDevice(string deviceNum) //根據每種Mode定義不同的DataTable~
        {
            //if(ListOfColumnsPerMode == null || ListOfColumnsPerMode.Count == 0) return null;

            var dt = new DataTable(deviceNum);

            foreach (var colSet in ListOfColumnsPerMode)
            {
                var set = colSet.Split(':'); //Name:Type:Init Value

                switch (set[1])
                {
                    case "int":
                        dt.Columns.Add(new DataColumn(set[0], typeof(int)) { DefaultValue = Convert.ToInt16(set[2]) });
                        break;
                    case "string":
                        dt.Columns.Add(new DataColumn(set[0], typeof(string)) { DefaultValue = set[2] });
                        break;
                    case "double":
                        dt.Columns.Add(new DataColumn(set[0], typeof(double))
                        { DefaultValue = Convert.ToDouble(set[2]) });
                        break;
                }
            }

            return dt;
        }
    }
}