using System.IO;
using PmicAutogen.Config.ProjectConfig;
using System.Collections.Generic;

namespace PmicAutogen.Local
{
    public static class FolderStructure
    {
        private const string StrIgLink = "IGLink";
        private const string StrTrunk = "trunk";
        private const string StrXmlFiles = "xml_Files";
        private const string StrModulesCommonSheets = "Modules_CommonSheets";
        private const string StrModulesLib = "Modules_Lib_TER";
        private const string StrModulesBlock = "Modules_Block";
        private const string StrCommon = "Common";
        private const string StrAcSpec = "AC_Spec";
        private const string StrBinTbl = "BinTable";
        private const string StrChannelMap = "ChannelMap";
        private const string StrCommonSheets = "Common_Sheets";
        private const string StrExtraSheets = "Extra";
        private const string StrDcSpec = "DC_Spec";
        private const string StrDevChar = "DevChar";
        private const string StrGlbSpec = "Global_Spec";
        private const string StrLevel = "Levels";
        private const string StrPinMap = "PinMap";
        private const string StrPortMap = "Ports";
        private const string StrTimings = "TimeSets";
        private const string StrJob = "Job";
        private const string StrPatSetsAll = "PatternSet";
        private const string StrModule = "Module";
        private const string StrScan = "Scan";
        private const string StrMbist = "Mbist";
        private const string StrHardIp = "HardIP";
        private const string StrVbt = "VBT";
        private const string StrLib = "Library";
        private const string StrDirVbtGenTool = "VBTGenTool";
        private const string StrConti = "DC_Conti";
        private const string StrMain = "Main_Flow";
        private const string StrLimitSheet = "LimitSheet";
        private const string StrReference = "Reference";
        private const string StrTestInstance = "TestIntance";
        private const string StrSpikeCheck = "SpikeCheck";

        private const string StrLibCommon = "Library_Common";
        private const string StrLibDC = "Library_DC";
        private const string StrLibDigital = "Library_Digital";
        private const string StrLibDSP = "Library_DSP";
        private const string StrLibLimitSheet = "Library_LimitSheet";
        private const string StrLibRelay = "Library_Relay";
        private const string StrLibVbtPopgen = "Library_VBT_POP_Gen";
        private const string StrLibPowerup = "Library_Powerup";

        private const string StrAcore = "ACORE";
        private const string StrBSTSQ = "BSTSQ";
        private const string StrBUCKMUTIPHASE = "BUCK_MUTIPHASE";
        private const string StrBUCK1P = "BUCK1P";
        private const string StrBUCK1PH = "BUCK1PH";
        private const string StrBUCKSW = "BUCKSW";
        private const string StrDc = "DC";
        private const string StrDCTestFunc = "DCTest_Func";
        private const string StrDigitalScanMbist = "Digital_Scan_Mbist";
        private const string StrLdo = "LDO";
        private const string StrOtp = "OTP";
        private const string StrPowerUpDown = "Power_Up_Down";

        private const string StrOtherWaitForClassify = "Other_WaitForClassify";
        #region old folder
        ////Folder1
        //private static string SubDirIgLink = Path.Combine(LocalSpecs.TarDir, StrIgLink);

        ////Folder2
        //private static string SubDirTrunk = Path.Combine(SubDirIgLink, StrTrunk);

        //private static string SubDirXmlFiles = Path.Combine(SubDirIgLink, StrXmlFiles);

        ////Folder3
        //private static string SubDirCommon = Path.Combine(SubDirTrunk, StrCommon);

        ////Folder4
        //private static string SubDirAcSpec = Path.Combine(SubDirCommon, StrAcSpec);
        //private static string SubDirBinTbl = Path.Combine(SubDirCommon, StrBinTbl);
        //private static string SubDirChannelMap = Path.Combine(SubDirCommon, StrChannelMap);
        //private static string SubDirCommonSheets = Path.Combine(SubDirCommon, StrCommonSheets);
        //private static string SubDirExtraSheets = Path.Combine(SubDirCommon, StrExtraSheets);
        //private static string SubDirDcSpec = Path.Combine(SubDirCommon, StrDcSpec);
        //private static string SubDirDevChar = Path.Combine(SubDirCommon, StrDevChar);
        //private static string SubDirGlbSpec = Path.Combine(SubDirCommon, StrGlbSpec);
        //private static string SubDirLevel = Path.Combine(SubDirCommon, StrLevel);
        //private static string SubDirPinMap = Path.Combine(SubDirCommon, StrPinMap);
        //private static string SubDirPortMap = Path.Combine(SubDirCommon, StrPortMap);
        //private static string SubDirTimings = Path.Combine(SubDirCommon, StrTimings);
        //private static string SubDirJob = Path.Combine(SubDirCommon, StrJob);

        //private static string SubDirPatSetsAll = Path.Combine(SubDirCommon, StrPatSetsAll);

        ////Folder3
        //private static string SubDirModule = Path.Combine(SubDirTrunk, StrModule);

        ////Folder4
        //private static  string SubDirOtp = Path.Combine(SubDirModule, StrOtp);
        //private static  string SubDirScan = Path.Combine(SubDirModule, StrScan);
        //private static  string SubDirMbist = Path.Combine(SubDirModule, StrMbist);
        //private static  string SubDirHardIp = Path.Combine(SubDirModule, StrHardIp);
        //private static  string SubDirVbt = Path.Combine(SubDirModule, StrVbt);
        //private static  string SubDirLib = Path.Combine(SubDirModule, StrLib);
        //private static  string SubDirVbtGenTool = Path.Combine(SubDirModule, StrDirVbtGenTool);
        //private static  string SubDirConti = Path.Combine(SubDirModule, StrConti);
        #endregion

        #region new folder
        //Folder1
        private static string SubDirIgLink = Path.Combine(LocalSpecs.TarDir, StrIgLink);

        //Folder2
        private static string SubDirTrunk = Path.Combine(SubDirIgLink, StrTrunk);

        private static string SubDirXmlFiles = Path.Combine(SubDirIgLink, StrXmlFiles);

        //Folder3
        private static string SubDirModulesCommonSheets = Path.Combine(SubDirTrunk, StrModulesCommonSheets);
        private static string SubDirModulesLib = Path.Combine(SubDirTrunk, StrModulesLib);
        private static string SubDirModulesBlock = Path.Combine(SubDirTrunk, StrModulesBlock);
        private static string SubDirOtherWaitForClassify = Path.Combine(SubDirTrunk, StrOtherWaitForClassify);

        //Folder4 ModulesCommonSheets
        private static string SubDirAcSpec = Path.Combine(SubDirModulesCommonSheets, StrAcSpec);
        private static string SubDirBinTbl = Path.Combine(SubDirModulesCommonSheets, StrBinTbl);
        private static string SubDirChannelMap = Path.Combine(SubDirModulesCommonSheets, StrChannelMap);
        private static string SubDirCommonSheets = Path.Combine(SubDirModulesCommonSheets, StrCommonSheets);
        private static string SubDirExtraSheets = Path.Combine(SubDirModulesCommonSheets, StrExtraSheets);
        private static string SubDirDcSpec = Path.Combine(SubDirModulesCommonSheets, StrDcSpec);
        private static string SubDirDevChar = Path.Combine(SubDirModulesCommonSheets, StrDevChar);
        private static string SubDirGlbSpec = Path.Combine(SubDirModulesCommonSheets, StrGlbSpec);
        private static string SubDirLevel = Path.Combine(SubDirModulesCommonSheets, StrLevel);
        private static string SubDirPinMap = Path.Combine(SubDirModulesCommonSheets, StrPinMap);
        private static string SubDirPortMap = Path.Combine(SubDirModulesCommonSheets, StrPortMap);
        private static string SubDirTimings = Path.Combine(SubDirModulesCommonSheets, StrTimings);
        private static string SubDirJob = Path.Combine(SubDirModulesCommonSheets, StrJob);
        private static string SubDirSpikeCheck = Path.Combine(SubDirModulesCommonSheets, StrSpikeCheck);
        private static string SubDirPatSetsAll = Path.Combine(SubDirModulesCommonSheets, StrPatSetsAll);
        private static string SubDirLimitSheet = Path.Combine(SubDirModulesCommonSheets, StrLimitSheet);
        private static string SubDirReference = Path.Combine(SubDirModulesCommonSheets, StrReference);
        private static string SubDirTestInstance = Path.Combine(SubDirModulesCommonSheets, StrTestInstance);

        //New Folder4 ModulesBlock
        private static string SubDirAcore = Path.Combine(SubDirModulesBlock, StrAcore);
        private static string SubDirBSTSQ = Path.Combine(SubDirModulesBlock, StrBSTSQ);
        private static string SubDirBUCKMUTIPHASE = Path.Combine(SubDirModulesBlock, StrBUCKMUTIPHASE);
        private static string SubDirBUCK1P = Path.Combine(SubDirModulesBlock, StrBUCK1P);
        private static string SubDirBUCK1PH = Path.Combine(SubDirModulesBlock, StrBUCK1PH);
        private static string SubDirBUCKSW = Path.Combine(SubDirModulesBlock, StrBUCKSW);
        private static string SubDirDc = Path.Combine(SubDirModulesBlock, StrDc);
        private static string SubDirDCTestFunc = Path.Combine(SubDirModulesBlock, StrDCTestFunc);
        private static string SubDirDigitalScanMbist = Path.Combine(SubDirModulesBlock, StrDigitalScanMbist);
        private static string SubDirLdo = Path.Combine(SubDirModulesBlock, StrLdo);
        private static string SubDirOtp = Path.Combine(SubDirModulesBlock, StrOtp);
        private static string SubDirPowerUpDown = Path.Combine(SubDirModulesBlock, StrPowerUpDown);

        //Folder4 ModulesLib
        private static string SubDirLibCommon = Path.Combine(SubDirModulesLib, StrLibCommon);
        private static string SubDirLibDC = Path.Combine(SubDirModulesLib, StrLibDC);
        private static string SubDirLibDigital = Path.Combine(SubDirModulesLib, StrLibDigital);
        private static string SubDirLibDSP = Path.Combine(SubDirModulesLib, StrLibDSP);
        private static string SubDirLibLimitSheet = Path.Combine(SubDirModulesLib, StrLibLimitSheet);
        private static string SubDirLibRelay = Path.Combine(SubDirModulesLib, StrLibRelay);
        private static string SubDirLibVbtPopgen = Path.Combine(SubDirModulesLib, StrLibVbtPopgen);
        private static string SubDirLibPowerup = Path.Combine(SubDirModulesLib, StrLibPowerup);

        private static Dictionary<string, string> _modulesLibMap;
        #endregion



        private static string SubDirMain = Path.Combine(SubDirModulesCommonSheets, StrMain);

        public static string DirIgLink => Path.Combine(SubDirIgLink);

        public static string DirTrunk => Path.Combine(SubDirTrunk);

        public static string DirXmlFiles => Path.Combine(SubDirXmlFiles);

        public static string DirModulesCommonSheets => Path.Combine(SubDirModulesCommonSheets);
        public static string DirModulesBlock => Path.Combine(SubDirModulesBlock);
        public static string DirModulesLib => Path.Combine(SubDirModulesLib);
        public static string DirOtherWaitForClassify => Path.Combine(SubDirOtherWaitForClassify);

        public static string DirAcSpec => Path.Combine(SubDirAcSpec);

        public static string DirBinTable => Path.Combine(SubDirBinTbl);

        public static string DirChannelMap => Path.Combine(SubDirChannelMap);

        public static string DirCommonSheets => Path.Combine(SubDirPowerUpDown);//CommonSheets to Power_Up_Down

        public static string DirExtraSheets => Path.Combine(SubDirExtraSheets);

        public static string DirDcSpec => Path.Combine(SubDirDcSpec);

        public static string DirDevChar => Path.Combine(SubDirDigitalScanMbist); //DevChar to Digital_Scan_Mbist

        public static string DirGlbSpec => Path.Combine(SubDirGlbSpec);

        public static string DirLevel => Path.Combine(SubDirLevel);

        public static string DirPinMap => Path.Combine(SubDirPinMap);

        public static string DirPortMap => Path.Combine(SubDirPortMap);

        public static string DirTimings => Path.Combine(SubDirTimings);

        public static string DirJob => Path.Combine(SubDirJob);

        public static string DirPatSetsAll => Path.Combine(SubDirPatSetsAll);

        public static string DirOtp => Path.Combine(SubDirOtp);

        public static string DirSpikeCheck => Path.Combine(SubDirSpikeCheck);

        public static string DirScan
        {
            get
            {
                var scan = ProjectConfigSingleton.Instance()
                    .ReplaceItemNameByConfigGroup(ProjectConfigSingleton.ConModuleName, SubDirDigitalScanMbist);
                return Path.Combine(scan);
            }
        }

        public static string DirMbist
        {
            get
            {
                var mbist = ProjectConfigSingleton.Instance()
                    .ReplaceItemNameByConfigGroup(ProjectConfigSingleton.ConModuleName, SubDirDigitalScanMbist);
                return Path.Combine(mbist);
            }
        }

        public static string DirHardIp => Path.Combine(SubDirDCTestFunc);

        public static string DirVbt => Path.Combine(SubDirAcore);

        public static string DirLib => Path.Combine(DirOtherWaitForClassify);//TODO

        public static string DirVbtGenTool => Path.Combine(SubDirOtp);

        public static string DirConti => Path.Combine(SubDirDc);

        public static string DirMain => Path.Combine(SubDirMain);

        public static string DirLimitSheet => Path.Combine(SubDirLimitSheet);
        public static string DirReference => Path.Combine(SubDirReference);
        public static string DirTestInstance => Path.Combine(SubDirTestInstance);

        //Folder4 ModulesLib
        public static string DirLibCommon => Path.Combine(SubDirLibCommon);
        public static string DirLibDC => Path.Combine(SubDirLibDC);
        public static string DirLibDigital => Path.Combine(SubDirLibDigital);
        public static string DirLibDSP => Path.Combine(SubDirLibDSP);
        public static string DirLibLimitSheet => Path.Combine(SubDirLibLimitSheet);
        public static string DirLibRelay => Path.Combine(SubDirLibRelay);
        public static string DirLibVbtPopgen => Path.Combine(SubDirLibVbtPopgen);
        public static string DirLibPowerup => Path.Combine(SubDirLibPowerup);

        //Folder4 ModulesBlock      
        public static string DirBSTSQ => Path.Combine(SubDirBSTSQ);
        public static string DirBUCKMUTIPHASE => Path.Combine(SubDirBUCKMUTIPHASE);
        public static string DirBUCK1P => Path.Combine(SubDirBUCK1P);
        public static string DirBUCK1PH => Path.Combine(SubDirBUCK1PH);
        public static string DirBUCKSW => Path.Combine(SubDirBUCKSW);
        public static string DirLdo => Path.Combine(SubDirLdo);

        public static Dictionary<string, string> ModulesLibMap
        {
            get
            {
                if (_modulesLibMap == null)
                {
                    _modulesLibMap = new Dictionary<string, string>();
                    _modulesLibMap.Add(StrLibCommon, DirLibCommon);
                    _modulesLibMap.Add(StrLibDC, DirLibDC);
                    _modulesLibMap.Add(StrLibDigital, DirLibDigital);
                    _modulesLibMap.Add(StrLibDSP, DirLibDSP);
                    _modulesLibMap.Add(StrLimitSheet, DirLibLimitSheet);
                    _modulesLibMap.Add(StrLibRelay, DirLibRelay);
                    _modulesLibMap.Add(StrLibVbtPopgen, DirLibVbtPopgen);
                }
                return _modulesLibMap;
            }
        }

        public static void ResetFolderVaribles()
        {
            //Folder1
            SubDirIgLink = Path.Combine(LocalSpecs.TarDir, StrIgLink);

            //Folder2
            SubDirTrunk = Path.Combine(SubDirIgLink, StrTrunk);

            SubDirXmlFiles = Path.Combine(SubDirIgLink, StrXmlFiles);

            //Folder3
            SubDirModulesCommonSheets = Path.Combine(SubDirTrunk, StrModulesCommonSheets);
            SubDirModulesLib = Path.Combine(SubDirTrunk, StrModulesLib);
            SubDirModulesBlock = Path.Combine(SubDirTrunk, StrModulesBlock);
            SubDirOtherWaitForClassify = Path.Combine(SubDirTrunk, StrOtherWaitForClassify);

            //Folder4 ModulesCommonSheets
            SubDirAcSpec = Path.Combine(SubDirModulesCommonSheets, StrAcSpec);
            SubDirBinTbl = Path.Combine(SubDirModulesCommonSheets, StrBinTbl);
            SubDirChannelMap = Path.Combine(SubDirModulesCommonSheets, StrChannelMap);
            SubDirCommonSheets = Path.Combine(SubDirModulesCommonSheets, StrCommonSheets);
            SubDirExtraSheets = Path.Combine(SubDirModulesCommonSheets, StrExtraSheets);
            SubDirDcSpec = Path.Combine(SubDirModulesCommonSheets, StrDcSpec);
            SubDirDevChar = Path.Combine(SubDirModulesCommonSheets, StrDevChar);
            SubDirGlbSpec = Path.Combine(SubDirModulesCommonSheets, StrGlbSpec);
            SubDirLevel = Path.Combine(SubDirModulesCommonSheets, StrLevel);
            SubDirPinMap = Path.Combine(SubDirModulesCommonSheets, StrPinMap);
            SubDirPortMap = Path.Combine(SubDirModulesCommonSheets, StrPortMap);
            SubDirTimings = Path.Combine(SubDirModulesCommonSheets, StrTimings);
            SubDirJob = Path.Combine(SubDirModulesCommonSheets, StrJob);
            SubDirSpikeCheck = Path.Combine(SubDirModulesCommonSheets, StrSpikeCheck);
            SubDirPatSetsAll = Path.Combine(SubDirModulesCommonSheets, StrPatSetsAll);
            SubDirLimitSheet = Path.Combine(SubDirModulesCommonSheets, StrLimitSheet);
            SubDirReference = Path.Combine(SubDirModulesCommonSheets, StrReference);
            SubDirTestInstance = Path.Combine(SubDirModulesCommonSheets, StrTestInstance);

            //Folder4 ModulesBlock
            SubDirAcore = Path.Combine(SubDirModulesBlock, StrAcore);
            SubDirBSTSQ = Path.Combine(SubDirModulesBlock, StrBSTSQ);
            SubDirBUCKMUTIPHASE = Path.Combine(SubDirModulesBlock, StrBUCKMUTIPHASE);
            SubDirBUCK1P = Path.Combine(SubDirModulesBlock, StrBUCK1P);
            SubDirBUCK1PH = Path.Combine(SubDirModulesBlock, StrBUCK1PH);
            SubDirBUCKSW = Path.Combine(SubDirModulesBlock, StrBUCKSW);
            SubDirDc = Path.Combine(SubDirModulesBlock, StrDc);
            SubDirDCTestFunc = Path.Combine(SubDirModulesBlock, StrDCTestFunc);
            SubDirDigitalScanMbist = Path.Combine(SubDirModulesBlock, StrDigitalScanMbist);
            SubDirLdo = Path.Combine(SubDirModulesBlock, StrLdo);
            SubDirOtp = Path.Combine(SubDirModulesBlock, StrOtp);
            SubDirPowerUpDown = Path.Combine(SubDirModulesBlock, StrPowerUpDown);

            //Folder4 ModulesLib
            SubDirLibCommon = Path.Combine(SubDirModulesLib, StrLibCommon);
            SubDirLibDC = Path.Combine(SubDirModulesLib, StrLibDC);
            SubDirLibDigital = Path.Combine(SubDirModulesLib, StrLibDigital);
            SubDirLibDSP = Path.Combine(SubDirModulesLib, StrLibDSP);
            SubDirLibLimitSheet = Path.Combine(SubDirModulesLib, StrLibLimitSheet);
            SubDirLibRelay = Path.Combine(SubDirModulesLib, StrLibRelay);
            SubDirLibVbtPopgen = Path.Combine(SubDirModulesLib, StrLibVbtPopgen);
            SubDirLibPowerup = Path.Combine(SubDirModulesLib, StrLibPowerup);
        }
    }
}