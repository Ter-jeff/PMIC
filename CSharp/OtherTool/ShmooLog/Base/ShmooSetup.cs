using System;
using System.Collections.Generic;

namespace ShmooLog.Base
{
    [Serializable]
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
}