Attribute VB_Name = "LIB_Common_GlobalConstant"
Option Explicit
'variable declaration
'Public Const Version_Lib_GlobalConstant = "0.1"  'lib version
Public Const Version_Lib_GlobalConstant = "0.2"  'add, remove efuse variant for Cayman, remove unused items

' Setup for DebugPrint
Public Const AllDCVIPinlist = "All_DCVI"
Public Const AllPowerPinlist = "All_Power"
Public Const All_DigitalPinlist = "All_Digital" 'need to remove refclk pins
Public Const All_DigitalPinlist_Disc = "All_Digital_Disc"   'need to remove refclk, PA pins pins
Public Const All_Utility_list = "All_Utilities"
'CHWu modify 10/14 remove useless pin group
'Public Const PinGrouplist = "Pins_0pv,Pins_0p4v,Pins_0p8v,Pins_0p9v,Pins_1p1v,Pins_1p2v,Pins_1p8v,Pins_3p3v"    ',DDR_IO_GP,DDR_Vref,Efuse_Data_Out,PLL_Pins_1p8v,LPDP_IO_GP,LPDP_TX3_Diff,PCIE_IO_GP,MIPI_IO_GP,PCIE_REF,Pcie_txrx_io,Group_A,gpio20,SEP_SPI_MISO,GPIO_leak,ULPI_DIR"
Public Const PinGrouplist = "Pins_1p1v,Pins_1p2v,Pins_1p8v"    ',DDR_IO_GP,DDR_Vref,Efuse_Data_Out,PLL_Pins_1p8v,LPDP_IO_GP,LPDP_TX3_Diff,PCIE_IO_GP,MIPI_IO_GP,PCIE_REF,Pcie_txrx_io,Group_A,gpio20,SEP_SPI_MISO,GPIO_leak,ULPI_DIR"
Public Const XI0_GP = ""
Public Const XI0_Diff_GP = "XI0_Diff_PA"
Public Const RTCLK_GP = "RT_CLK32768_PA"
Public Const RTCLK_Diff_GP = ""
Public FreeRunFreq_debug As Double
Public clock_Vih_debug As Double
Public clock_Vil_debug As Double
Public CurrentXi0Freq As Double
Public DebugPrintFlag_Chk As Boolean

'universal
Public Const MaxNumSite = 8
Public currentJobName As String
Public gL_ProductionTemp As String
Public gS_SPI_Version As String    'SPIROM version printing

'Control flags
Public Flag_RAK_INIT As Boolean
Public Flag_RSCR_INIT As Boolean
Public Flag_Shmoo_INIT As Boolean
Public Flag_MBISTFailBlock_INIT As Boolean
'Public Flag_Uart_INIT As Boolean

'nWire Setup, dual port
Public Const XI0_PA_Pin = "XI0_PA, XO0_PA"
Public Const XI0_1_PA_Pin = "RT_CLK32768_PA"
Public Const XI0_PA_Refclk_Pin = "REFCLK1"
Public Const XI0_1_PA_Refclk_Pin = "REFCLK2"
Public Const Clock_Port = "Clock_Port"
Public Const Clock_Port1 = "RTCLK_Port"
Public Const XI0_ref_VOH = 1.8  'use 1.8v buffer
Public Const XI0_ref_VOL = 0  'use 1.8v buffer
Public Const Relay_Off_nWire = "K0"
Public Const Relay_On_nWire = ""
Public Const Relay_Off_SupportBoard = ""
Public Const Relay_On_SupportBoard = ""
Public Const Level_nWire = "Levels_nWire_XI0"
Public Const Level_nWire_Diff = "Levels_nWire_XI0_Diff"
Public Const TSB_nWire = "TSB_nWire_XI0"

'====================================================
'=   Define the variables for ECID Fuse data        =
'====================================================
Public XCoord As New SiteLong
Public YCoord As New SiteLong
Public WaferID As Long
Public LotID As String
Public HramWaferId As New SiteLong
Public HramLotId As New SiteVariant
Public HramXCoord As New SiteLong
Public HramYCoord As New SiteLong

Public TMPS_TD1_1 As New SiteLong
Public TMPS_TD1_2 As New SiteLong
Public TMPS_TD1_3 As New SiteLong
Public TMPS_TD1_4 As New SiteLong
Public TMPS_TD1_5 As New SiteLong
Public TMPS_TD1_6 As New SiteLong
Public TMPS_TD1_7 As New SiteLong

Public ADC_trim_V3 As New SiteLong
Public ADC_trim_V3_ECID As New SiteDouble
Public REFERENCE_CTRL_25C As New SiteLong
Public VOLTAGE_TRIM_BITS As New SiteLong
Public TEMP_TRIM_BITS1 As New SiteLong

'====================================================
'=  ECID fuse test flags                            =
'====================================================
Public FailFlag_untrim25c As New SiteBoolean 'TMPS TD1 25C
Public FailFlag_ADC25C As New SiteBoolean
Public FailFlag_Freq_Detect As New SiteBoolean

'====================================================
'=   Define the variables for Config Fuse data      =
'====================================================
Public Synth_Trim As New SiteLong
Public TRIMG_SOC_0 As New SiteLong
Public TRIMO_SOC_0 As New SiteLong
Public TRIMG_SOC_1 As New SiteLong
Public TRIMO_SOC_1 As New SiteLong
Public TRIMG_SOC_2 As New SiteLong
Public TRIMO_SOC_2 As New SiteLong
Public TRIMG_SOC_3 As New SiteLong
Public TRIMO_SOC_3 As New SiteLong


'[  for SPI - Define the IDS code resolution according to Table 32 and 33 in Test Plan ]
Public I_VDD_CPU_SPI As New SiteDouble
Public I_VDD_GPU_SPI As New SiteDouble
Public I_VDD_SOC_SPI As New SiteDouble
Public I_VDD_FIXED_SPI As New SiteDouble
Public I_VDD_CPU_SRAM_SPI As New SiteDouble
Public I_VDD_GPU_SRAM_SPI As New SiteDouble
Public I_VDD_LOW_SPI As New SiteDouble

'use in vdd binning VBT codes
Public I_VDD_CPU_IDS_Check As New SiteDouble
Public I_VDD_GPU_IDS_Check As New SiteDouble
Public I_VDD_SOC_IDS_Check As New SiteDouble
Public I_VDD_FIXED_IDS_Check As New SiteDouble
Public I_VDD_CPU_SRAM_IDS_Check As New SiteDouble
Public I_VDD_GPU_SRAM_IDS_Check As New SiteDouble
Public I_VDD_LOW_IDS_Check As New SiteDouble

Public IDS_CPU_Decimal As New SiteLong
Public IDS_GPU_Decimal As New SiteLong
Public IDS_SOC_Decimal As New SiteLong
Public IDS_FIXED_Decimal As New SiteLong
Public IDS_CPU_SRAM_Decimal As New SiteLong
Public IDS_GPU_SRAM_Decimal As New SiteLong
Public IDS_LOW_Decimal As New SiteLong

'20150610 update
Public IDS_CPU_Resolution As Double      ' 0.0002    ''0.2mA
Public IDS_GPU_Resolution As Double      ' 0.0002    ''0.2mA
Public IDS_SOC_Resolution As Double      ' 0.0001    ''0.1mA
Public IDS_FIXED_Resolution As Double    ' 0.0001    ''0.1mA
Public IDS_CPU_SRAM_Resolution As Double ' 0.0001    ''0.1mA
Public IDS_GPU_SRAM_Resolution As Double ' 0.0001    ''0.1mA
Public IDS_LOW_Resolution As Double      ' 0.0001    ''0.1mA

'20150121 define the MaxDecimal according to test plan
Public IDS_CPU_MaxDecimal As Double
Public IDS_GPU_MaxDecimal As Double
Public IDS_SOC_MaxDecimal As Double
Public IDS_FIXED_MaxDecimal As Double
Public IDS_CPU_SRAM_MaxDecimal As Double
Public IDS_GPU_SRAM_MaxDecimal As Double
Public IDS_LOW_MaxDecimal As Double


Public DPTX_LPDP0_PLL_FCAL As New SiteLong
Public PCIE_REFPLL_FCAL_VCO_DIGCTRL As New SiteLong
Public LPDP_C_RX As New SiteLong
Public LS3B As New SiteLong

'''20150207 add FCAL_VCO_DIGCTRL    'old Elba
Public FCAL_VCO_DIGCTRL_Decimal As New SiteLong
Public PCIE_FCAL_VCO_DIGCTRL_1st_Value As New SiteLong
Public PCIE_FCAL_VCO_DIGCTRL_2nd_Value As New SiteLong

'''20150819 add PLL_CPU_KVCO    'Cayman
Public PLL_CPU_KVCO_Decimal As New SiteLong


'''20150819 add PLL_LPDP_FCAL    'Cayman
Public PLL_LPDP_FCAL_Decimal As New SiteLong

'''20160725 add ADCLK trim    'Skye
Public PLL_GPU_FCAL_Decimal As New SiteLong
Public pblk_PLL_CFG1_kvco_trim_Decimal As New SiteLong
Public eblk_PLL_CFG1_kvco_trim_Decimal As New SiteLong

'====================================================
'=  Config fuse test flags                          =
'====================================================
Public FailFlag_Fcal_LPDP As New SiteBoolean
Public FailFlag_TrimVerify85c As New SiteBoolean    'trimG, trimO
Public FailFlag_Freq_Synth As New SiteBoolean
Public FailFlag_FCAL_VCO As New SiteBoolean

'Real VddBinning Check revision
Public Const Real_VddBinning_version = 99

'====================================================
'=    Define the variables for UDR Fuse data        =
'====================================================
Public TRIMG_CPU_0 As New SiteLong
Public TRIMO_CPU_0 As New SiteLong
Public TRIMG_CPU_1 As New SiteLong
Public TRIMO_CPU_1 As New SiteLong
Public TRIMG_CPU_2 As New SiteLong
Public TRIMO_CPU_2 As New SiteLong
Public ADC_vTrim As New SiteLong
Public ADC_tTrim As New SiteLong
Public pllTrimFusedBit As New SiteLong
Public PLL_CFG1_kvco_trim As New SiteLong
Public ADCLK_SCR2_vsns_cal_fuse As New SiteLong


'====================================================
'=  UDR fuse test flags                             =
'====================================================
Public FailFlag_ADC85C As New SiteBoolean
Public FailFlag_pllTrimFusedBit As New SiteBoolean
Public FailFlag_PLL_CFG1_kvco_trim As New SiteBoolean
Public FailFlag_ADCLK_SCR2_vsns_cal_fuse As New SiteBoolean

'20150319 add CPU PLL Fcal  'old Elba
''Public CPU_PLL_Fcal_Decimal As New SiteLong
''Public CPU_PLL_Fcal_V1_Decimal As New SiteLong

'====================================================
'=   Define the variables for Sensor Fuse data      =
'====================================================
Public AFREQ_CTRL_EN As New SiteLong
Public LATENCY As New SiteLong
Public Freq_Det_Precision As New SiteLong
Public Freq_Det_Decimal As New SiteLong
Public DFREQ_CTRL_EN As New SiteLong
Public DFREQ_CTRL_OFFSET As New SiteLong
Public SEN_SOC_TRIMG_0 As New SiteLong ''''should be equal to TRIMG_SOC_0
Public SEN_SOC_TRIMO_0 As New SiteLong ''''should be equal to TRIMO_SOC_0

'Public gS_SEN_CRC_HexStr As New SiteVariant



'====================================================
'=   Define the variables for IEDA registry     =
'====================================================
Public gS_TMPS1_Untrim As New SiteVariant
Public gS_TMPS2_Untrim As New SiteVariant
Public gS_TMPS3_Untrim As New SiteVariant
Public gS_TMPS4_Untrim As New SiteVariant
Public gS_TMPS5_Untrim As New SiteVariant
Public gS_TMPS6_Untrim As New SiteVariant
Public gS_TMPS7_Untrim As New SiteVariant
Public gS_TMPS8_Untrim As New SiteVariant
Public gS_TMPS9_Untrim As New SiteVariant
Public gS_TMPS10_Untrim As New SiteVariant
Public gS_TMPS11_Untrim As New SiteVariant
Public gS_TMPS12_Untrim As New SiteVariant
Public gS_TMPS13_Untrim As New SiteVariant
Public gS_TMPS14_Untrim As New SiteVariant

Public gS_TMPS1_Trim As New SiteVariant
Public gS_TMPS2_Trim As New SiteVariant
Public gS_TMPS3_Trim As New SiteVariant
Public gS_TMPS4_Trim As New SiteVariant
Public gS_TMPS5_Trim As New SiteVariant
Public gS_TMPS6_Trim As New SiteVariant
Public gS_TMPS7_Trim As New SiteVariant
Public gS_TMPS8_Trim As New SiteVariant
Public gS_TMPS9_Trim As New SiteVariant
Public gS_TMPS10_Trim As New SiteVariant
Public gS_TMPS11_Trim As New SiteVariant
Public gS_TMPS12_Trim As New SiteVariant
Public gS_TMPS13_Trim As New SiteVariant
Public gS_TMPS14_Trim As New SiteVariant

Public gS_TMPS1 As New SiteVariant
Public gS_TMPS2 As New SiteVariant
Public gS_TMPS3 As New SiteVariant
Public gS_TMPS4 As New SiteVariant
Public gS_TMPS5 As New SiteVariant
Public gS_TMPS6 As New SiteVariant
Public gS_TMPS7 As New SiteVariant
Public gS_TMPS8 As New SiteVariant
Public gS_TMPS9 As New SiteVariant
Public gS_TMPS10 As New SiteVariant
Public gS_TMPS11 As New SiteVariant
Public gS_TMPS12 As New SiteVariant
Public gS_TMPS13 As New SiteVariant
Public gS_TMPS14 As New SiteVariant

'====================================================
'=   Define the variables for alarm happened     =
'====================================================
Public alarmFail As New SiteBoolean

'====================================================
'=   Define the variables for mbist loop     =
'====================================================
Public mbist_sheet_init As Boolean
Public currentBlock_loopCnt As Integer
Public currentAPK_loopCnt As Integer
'====================================================
'=   Define DAC Trim Flag                           =
'====================================================
'Public DACInitialFlag As Boolean

'========================
'HardIP Test Name
'========================
Public gl_Tname_Alg_Index As Long
Public gl_Tname_Meas As String
'----------------20180523----------------
Public gl_Tname_Meas_FromFlow() As String
Public gl_Tname_Alg As String
Public gl_Sweep_Name As String
Public gl_SweepY_Name As String
'=====================================================
'20171207 - HardIP use Standard Test Name Format Flag,Roger add
Public gl_UseStandardTestName_Flag As Boolean
'=====================================================
Public gl_Disable_HIP_debug_log As Boolean

Public XVal As Double
Public YVal As Double
Public gl_flag_end_shmoo As Boolean
Public gl_flag_CZ_Nominal_Measured_1st_Point As Boolean

Public gl_FlowForLoop_DigSrc_SweepCode As String

'========================
'Powerup/down sequence
'========================
Public power_dcvs_exit As Boolean
Public power_dcvi_exit As Boolean
Public io_h_pins As String
Public io_l_pins As String
Public io_hz_pins As String

Public PowerSequencePin_GB() As String
Public nwire_seq_GB() As Long
Public nwire_port_GB() As String
Public IO_H_seq_nu_GB() As Long
Public IO_H_seq_pin_total_GB() As String
Public IO_L_seq_nu_GB() As Long
Public IO_L_seq_pin_total_GB() As String
Public IO_HZ_seq_nu_GB() As Long
Public IO_HZ_seq_pin_total_GB() As String

Public TempMaxSequence_GB As Long
Public power_up_en As Boolean
'================================================
Public site As Variant

Public gL_License_check As Long
