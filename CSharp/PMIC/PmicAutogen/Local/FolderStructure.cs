using PmicAutogen.Config.ProjectConfig;
using System.IO;

namespace PmicAutogen.Local
{
    public static class FolderStructure
    {
        private const string StrIgLink = "IGLink";

        private const string StrTrunk = "trunk";
        private const string StrXmlFiles = "xml_Files";

        private const string StrModulesCommonSheets = "Modules_CommonSheets";
        private const string StrModulesBlock = "Modules_Block";
        private const string StrModulesLibTer = "Modules_Lib_TER";
        private const string StrOtherWaitForClassify = "Other_WaitForClassify";

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
        private const string StrMainFlow = "Main_Flow";

        private const string StrReference = "Reference";
        private const string StrTestInstance = "TestInstance";
        private const string StrSpikeCheck = "SpikeCheck";

        public const string StrLibCommon = "Library_Common";
        public const string StrLibDc = "Library_DC";
        public const string StrLibDigital = "Library_Digital";
        public const string StrLibDsp = "Library_DSP";
        public const string StrLibLimitSheet = "Library_LimitSheet";
        public const string StrLimitSheet = "LimitSheet";
        public const string StrLibRelay = "Library_Relay";
        public const string StrLibVbtPopgen = "Library_VBT_POP_Gen";
        public const string StrLibPowerup = "Library_Powerup";

        private const string StrAcore = "ACORE";
        private const string StrBstsq = "BSTSQ";
        private const string StrBuckmutiphase = "BUCK_MUTIPHASE";
        private const string StrBuck1P = "BUCK1P";
        private const string StrBuck1Ph = "BUCK1PH";
        private const string StrBucksw = "BUCKSW";
        private const string StrDc = "DC";
        private const string StrDcTestFunc = "DCTest_Func";
        private const string StrDigitalScanMbist = "Digital_Scan_Mbist";
        private const string StrLdo = "LDO";
        private const string StrOtp = "OTP";
        private const string StrPowerUpDown = "Power_Up_Down";

        #region new folder

        //Folder1
        public static string DirIgLink = Path.Combine(LocalSpecs.TarDir, StrIgLink);

        //Folder2
        public static string DirTrunk
        {
            get { return Path.Combine(DirIgLink, StrTrunk); }
        }

        public static string DirXmlFiles
        {
            get { return Path.Combine(DirIgLink, StrXmlFiles); }
        }

        //Folder3
        public static string DirModulesCommonSheets
        {
            get { return Path.Combine(DirTrunk, StrModulesCommonSheets); }
        }

        public static string DirModulesLibTer
        {
            get { return Path.Combine(DirTrunk, StrModulesLibTer); }
        }

        public static string DirModulesBlock
        {
            get { return Path.Combine(DirTrunk, StrModulesBlock); }
        }

        public static string DirOtherWaitForClassify
        {
            get { return Path.Combine(DirTrunk, StrOtherWaitForClassify); }
        }

        //Folder4 Modules_CommonSheets
        public static string DirMainFlow = Path.Combine(DirModulesCommonSheets, StrMainFlow);
        public static string DirAcSpec
        {
            get { return Path.Combine(DirModulesCommonSheets, StrAcSpec); }
        }

        public static string DirBinTable
        {
            get { return Path.Combine(DirModulesCommonSheets, StrBinTbl); }
        }

        public static string DirChannelMap
        {
            get { return Path.Combine(DirModulesCommonSheets, StrChannelMap); }
        }

        public static string DirDcSpec
        {
            get { return Path.Combine(DirModulesCommonSheets, StrDcSpec); }
        }

        public static string DirGlbSpec
        {
            get { return Path.Combine(DirModulesCommonSheets, StrGlbSpec); }
        }

        public static string DirLevel
        {
            get { return Path.Combine(DirModulesCommonSheets, StrLevel); }
        }

        public static string DirPinMap
        {
            get { return Path.Combine(DirModulesCommonSheets, StrPinMap); }
        }

        public static string DirPortMap
        {
            get { return Path.Combine(DirModulesCommonSheets, StrPortMap); }
        }

        public static string DirTimings
        {
            get { return Path.Combine(DirModulesCommonSheets, StrTimings); }
        }

        public static string DirJob
        {
            get { return Path.Combine(DirModulesCommonSheets, StrJob); }
        }

        public static string DirSpikeCheck
        {
            get { return Path.Combine(DirModulesCommonSheets, StrSpikeCheck); }
        }

        public static string DirPatSetsAll
        {
            get { return Path.Combine(DirModulesCommonSheets, StrPatSetsAll); }
        }

        public static string DirLimitSheet
        {
            get { return Path.Combine(DirModulesCommonSheets, StrLimitSheet); }
        }

        public static string DirReference
        {
            get { return Path.Combine(DirModulesCommonSheets, StrReference); }
        }

        public static string DirTestInstance
        {
            get { return Path.Combine(DirModulesCommonSheets, StrTestInstance); }
        }

        //New Folder4 Modules_Block
        public static string DirAcore
        {
            get { return Path.Combine(DirModulesBlock, StrAcore); }
        }

        public static string DirBstsq
        {
            get { return Path.Combine(DirModulesBlock, StrBstsq); }
        }

        public static string DirBuckmutiphase
        {
            get { return Path.Combine(DirModulesBlock, StrBuckmutiphase); }
        }

        public static string DirBuck1P
        {
            get { return Path.Combine(DirModulesBlock, StrBuck1P); }
        }

        public static string DirBuck1Ph
        {
            get { return Path.Combine(DirModulesBlock, StrBuck1Ph); }
        }

        public static string DirBucksw
        {
            get { return Path.Combine(DirModulesBlock, StrBucksw); }
        }

        public static string DirDc
        {
            get { return Path.Combine(DirModulesBlock, StrDc); }
        }

        public static string DirDcTestFunc
        {
            get { return Path.Combine(DirModulesBlock, StrDcTestFunc); }
        }

        public static string DirDigitalScanMbist
        {
            get { return Path.Combine(DirModulesBlock, StrDigitalScanMbist); }
        }

        public static string DirLdo
        {
            get { return Path.Combine(DirModulesBlock, StrLdo); }
        }

        public static string DirOtp
        {
            get { return Path.Combine(DirModulesBlock, StrOtp); }
        }

        public static string DirPowerUpDown
        {
            get { return Path.Combine(DirModulesBlock, StrPowerUpDown); }
        }

        public static string DirScan
        {
            get
            {
                var scan = ProjectConfigSingleton.Instance()
                    .ReplaceItemNameByConfigGroup(ProjectConfigSingleton.ConModuleName, DirDigitalScanMbist);
                return Path.Combine(scan);
            }
        }

        public static string DirMbist
        {
            get
            {
                var mbist = ProjectConfigSingleton.Instance()
                    .ReplaceItemNameByConfigGroup(ProjectConfigSingleton.ConModuleName, DirDigitalScanMbist);
                return Path.Combine(mbist);
            }
        }

        //Folder4 Modules_Lib_TER
        public static string DirLibCommon
        {
            get { return Path.Combine(DirModulesLibTer, StrLibCommon); }
        }

        public static string DirLibDc
        {
            get { return Path.Combine(DirModulesLibTer, StrLibDc); }
        }

        public static string DirLibDigital
        {
            get { return Path.Combine(DirModulesLibTer, StrLibDigital); }
        }

        public static string DirLibDsp
        {
            get { return Path.Combine(DirModulesLibTer, StrLibDsp); }
        }

        public static string DirLibLimitSheet
        {
            get { return Path.Combine(DirModulesLibTer, StrLibLimitSheet); }
        }

        public static string DirLibRelay
        {
            get { return Path.Combine(DirModulesLibTer, StrLibRelay); }
        }

        public static string DirLibVbtPopgen
        {
            get { return Path.Combine(DirModulesLibTer, StrLibVbtPopgen); }
        }

        public static string DirLibPowerup
        {
            get { return Path.Combine(DirModulesLibTer, StrLibPowerup); }
        }

        #endregion
    }
}