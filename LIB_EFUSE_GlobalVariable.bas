Attribute VB_Name = "LIB_EFUSE_GlobalVariable"

Option Explicit

''''=========================================
''''Convention Definition
''''gB:: global Boolean
''''gC:: global Const
''''gD:: global Double
''''gI:: global Integer
''''gL:: global Long
''''gS:: global String
''''gT:: global Type
''''gW:: globle DspWave
''''=========================================

Public gL_1st_FuseSheetRead As Long
Public gB_EFUSE_DVRV_ENABLE As Boolean
Public gS_JobName As String
Public gD_BaseStepVoltage As Double
Public gD_BaseVoltage As Double ''''was gL_BaseVoltage, 20160608 update
Public gD_VBaseFuse As Double   ''''was gL_VBaseFuse  , 20160608 update
Public gL_CFG_SegCNT As Long
Public gStr_PatName As String

''''20171103 add
Public gD_UDRE_BaseStepVoltage As Double
Public gD_UDRE_BaseVoltage As Double
Public gD_UDRE_VBaseFuse As Double
''''20171103 add
Public gD_UDRP_BaseStepVoltage As Double
Public gD_UDRP_BaseVoltage As Double
Public gD_UDRP_VBaseFuse As Double

Public gS_EFuse_Orientation As String ''''FM66,KP88:UP2DOWN; Malta,Elba,Rhea:RIGHT2LEFT; Cayman:SingleUp
Public gB_CFG_SVM As Boolean          ''''20160729 add, to save TTR
Public gB_eFuse_to_STDF As Boolean    ''''20160926 Add
Public gB_CFGSVM_BIT_Read_ValueisONE As New SiteBoolean
Public gB_CFGSVM_A00_CP1 As Boolean               ''''20161107 add
Public gB_eFuse_CFG_Cond_FTF_done_Flag As Boolean ''''20170923 Add

Public gS_BKM_Lot_Wafer_ID As String
Public gS_efuse_BKM_Ver As String
Public gS_BKM_Number As String
Public gS_BKM_IEDA As New SiteVariant
Public gS_BKM_Fuse_IEDA As New SiteVariant
Public Dic_BKM As New Dictionary '''2010710

'''' Table Column Sequence:: (eFuse_BitDef_Table)
'''' MSB Bit    LSB Bit Bit Width   programming stage   Low Limit   High Limit  IDS Resolution  Algorithm   Comment  Use Default or Real value  Default Value   Difference
''''
''''=========================================
''''Definition
''''=========================================
''''Single-Bit (1-Bit), Double-Bit (2-Bit) Orientation
''''ECID eFuse Bit Def
''''Config eFuse Bit Def
''''UID(AES) eFuse Bit Def
''''UDR eFuse Bit Def
''''Sensor Trim value Bit Def (instead by Monitor)
''''Monitor eFuse Bit Def
''''CMP eFuse Bit Def
''''end
''''=========================================

''''---------------------------------------------------------------------------------------------------
'''' eFuse Common Global Variables
''''---------------------------------------------------------------------------------------------------
''''must be global and used in auto_XXXSingleDoubleBit()
''''It was extracted in the module auto_XXX_Read_by_OR_2Blocks() then used in auto_XXXSingleDoubleBit()
Public gS_SingleStrArray() As String

Public gL_eFuse_catename_maxLen As Long
''''20160630 add
Public gS_ECID_SingleBit_Str As New SiteVariant ''''be used in JTAG_Read
Public gS_CFG_SingleBit_Str As New SiteVariant  ''''be used in JTAG_Read

Public gS_ECID_CRC_HexStr As New SiteVariant    ''''20161003 ADD CRC
Public gS_CFG_CRC_HexStr As New SiteVariant     ''''20161003 ADD CRC

''''---------------------------------------------------------------------------------------------------
'''' ECID Fuse
''''---------------------------------------------------------------------------------------------------
Public EcidBlock As Long                       ' Num of Block
Public EcidRowPerBlock As Long                 ' Num of word per Block in V1.3 test plan
Public EcidBitsPerRow As Long                  ' Num of bit per Row
Public EcidWriteBitExpandWidth As Long         ' Expand Write cycles per Ecid Bit
''20191230 , make bank size to be dynamic
'Public EcidReadBitWidth As Long                ' Num of bit per Read cycle
Public EcidBitPerBlockUsed As Long             ' Num of bit used per Block in V1.3 test plan
Public EcidBitPerBlock As Long                 ' Num of bit per Block
Public EcidReadCycle As Long                   ' Num of cycles to read out all Blocks
Public ECIDTotalBits As Long
Public ECIDBitPerCycle As Long
Public EcidHiLimitSingleDoubleBitCheck As Long 'Hi Limitfor Single-Double-Bit Check

''20191230 , make bank size to be dynamic
'Public EcidCharPerLotId As Long    '' = 6
'Public EcidBitPerLotIdChar As Long '' = 6

Public gS_ECID_Direct_Access_Str As New SiteVariant ''''be used in DAP
Public gI_Index_DEID As Long
Public gI_ECID_catename_maxLen As Long
Public gL_ECIDFuse_Pgm_Bit() As Long
Public gB_ReadWaferData_flag As Boolean
Public gB_eFuse_Disable_ChkLMT_Flag As Boolean
Public gB_ECID_decode_flag As New SiteBoolean

''''------ ECIDFuse Global Variable --------------------------------
''''ECID start bit and end bit
Public LOTID_FIRST_BIT As Long
Public LOTID_LAST_BIT As Long
Public LOTID_BITWIDTH As Long

Public WAFERID_FIRST_BIT As Long
Public WAFERID_LAST_BIT As Long
Public WAFERID_BITWIDTH As Long

Public XCOORD_FIRST_BIT As Long
Public XCOORD_LAST_BIT As Long
Public XCOORD_BITWIDTH As Long
Public XCOORD_LoLMT As Long
Public XCOORD_HiLMT As Long

Public YCOORD_FIRST_BIT As Long
Public YCOORD_LAST_BIT As Long
Public YCOORD_BITWIDTH As Long
Public YCOORD_LoLMT As Long
Public YCOORD_HiLMT As Long

''Public ECID_BlankCheck_Retest_Compare_Bit As Long
''Public ECID_BlankCheck_Compare_Bit As Long
''''--------------------------------------------------------------------

''''---------------------------------------------------------------------------------------------------
'''' Config Fuse
''''---------------------------------------------------------------------------------------------------
Public EConfigBlock As Long                         ' Num of Block
Public EConfigRowPerBlock As Long                   ' Num of word per Block in V1.3 test plan
Public EConfigBitPerBlockUsed As Long               ' Num of bit used per Block in V1.3 test plan
Public EConfig_Repeat_Cyc_for_Pgm As Long           ' Repeat cycle per bit program needed
''20191230 , make bank size to be dynamic
'Public EConfigReadBitWidth As Long                  ' Num of bit per Read cycle
Public EConfigBitsPerRow As Long                    ' Num of bit per Word
Public EConfigHiLimitSingleDoubleBitCheck As Long   ' Max allowed failed bit count (FBC)
Public EConfigReadCycle As Long                     ' 32(rows)x32(bits/row)=1024 bits
Public EConfigTotalBitCount As Long

Public RealVDDBin As Long
Public gS_cfgFlagname As String
Public gS_cfgFlagname_pre As String
Public gS_ConfigCondition As String
Public gS_DevRevArr() As Variant
Public gS_Major_DevRevArr() As Variant
Public gS_CFG_Direct_Access_Str As New SiteVariant
Public gS_cfgTable_First64bitsStr As String
Public gS_CFGCondTable_bitsStr As String    ''''20170630 add
Public gS_CFG_Cond_Read_bitStrM As New SiteVariant  ''''201812XX add
Public gS_CFG_Cond_Read_pkgname As New SiteVariant  ''''201812XX add
Public gL_CFG_Cond_compResult As New SiteLong       ''''201812XX add
Public gI_CFG_catename_maxLen As Long
Public gI_CFG_firstbits_index As Long
Public gS_CFG_firstbits_stage As String
Public gL_CFG_Cond_JobvsStage As Long
Public gL_CFGFuse_Pgm_Bit() As Long
Public gB_CFG_decode_flag As New SiteBoolean
''''20170630 add
Public gS_CFG_SCAN_stage As String
Public gB_CFG_blank_SCAN As New SiteBoolean
Public gB_CFG_blank_Cond As New SiteBoolean

'''''---------------------------------------------------------------------------------------------------
''''' UID (AES)  Fuse
'''''---------------------------------------------------------------------------------------------------
Public UIDBlock As Long                       ' Num of Block
Public UIDBitsPerCode As Long                 ' was AESBitPerBlock, Because there are 128 random bits are generated from C651 *.dll file
Public UIDRowPerBlock As Long                 ' Num of Row per block
Public UIDBitsPerRow As Long                  ' Num of bit per Word (was=>LeftRight_Bit_perRow)
Public UIDWriteBitExpandWidth As Long         ' Expand Write cycles per UID Bit, refer to the pattern
Public UIDReadBitWidth As Long                ' Num of bit per Read cycle
Public UIDBitsPerBlockUsed As Long            ' Num of bit used per Block
Public UIDBitsPerBlock As Long                ' Num of bit per Block
Public UIDReadCycle As Long                   ' Num of cycles to read out all Blocks
Public UIDTotalBits As Long                   ' was SEFTotalBits
Public UIDBitsPerCycle As Long

''''-------------------------------------------------------------
Public UID_Code_BitStr() As String
Public UID_ChkSum_LoLimit As Double ''= 0.2       'low limit decided by anh (2013/11/26)
Public UID_ChkSum_HiLimit As Double ''= 0.65
Public DisplayUID As Boolean
Public gL_UIDFuse_Pgm_Bit() As Long
Public gI_UID_catename_maxLen As Long
Public gL_UIDCodeBitWidth As Long
Public gL_UIDCode_Block As Long    ''''how many UIDCode(128bits) blocks
Public gD_UIDBlock1Sum As New SiteDouble ''''20160817 add
Public gD_UIDBlock2Sum As New SiteDouble ''''20160817 add

''''---------------------------------------------------------------------------------------------------
'''' UDR Fuse (was DVFM)
''''---------------------------------------------------------------------------------------------------
Public gS_USI_BitStr As New SiteVariant
Public Trim_code As New DSPWave
Public TMPS_fail_flag As SiteBoolean
Public gL_Trim_Code_Size As Long
Public gI_UDR_catename_maxLen As Long
Public gB_UDR_decode_flag As New SiteBoolean

''''---------------------------------------------------------------------------------------------------
'''' SENSOR Trim Fuse
''''---------------------------------------------------------------------------------------------------
Public SENSORBlock As Long                         ' Num of Block
Public SENSORRowPerBlock As Long                   ' Num of word per Block
Public SENSORBitPerBlockUsed As Long               ' Num of bit used per Block
Public SENSOR_Repeat_Cyc_for_Pgm As Long           ' Repeat cycle per bit program needed
Public SENSORReadBitWidth As Long                  ' Num of bit per Read cycle
Public SENSORBitsPerRow As Long                    ' Num of bit per Word
Public SENSORHiLimitSingleDoubleBitCheck As Long   ' Max allowed failed bit count (FBC)
Public SENSORReadCycle As Long                     ' 32(rows)x32(bits/row)=1024 bits
Public SENSORTotalBitCount As Long
Public gI_SEN_catename_maxLen As Long
Public gL_SENFuse_Pgm_Bit() As Long
Public gB_SEN_decode_flag As New SiteBoolean
Public gS_SEN_CRC_HexStr As New SiteVariant
Public gS_SEN_Direct_Access_Str As New SiteVariant

''''---------------------------------------------------------------------------------------------------
'''' MONITOR Trim Fuse
''''---------------------------------------------------------------------------------------------------
Public MONITORBlock As Long                         ' Num of Block
Public MONITORRowPerBlock As Long                   ' Num of word per Block
Public MONITORBitPerBlockUsed As Long               ' Num of bit used per Block
Public MONITOR_Repeat_Cyc_for_Pgm As Long           ' Repeat cycle per bit program needed
''20191230 , make bank size to be dynamic
'Public MONITORReadBitWidth As Long                  ' Num of bit per Read cycle
Public MONITORBitsPerRow As Long                    ' Num of bit per Word
Public MONITORHiLimitSingleDoubleBitCheck As Long   ' Max allowed failed bit count (FBC)
Public MONITORReadCycle As Long                     ' 32(rows)x32(bits/row)=1024 bits
Public MONITORTotalBitCount As Long
Public gI_MON_catename_maxLen As Long
Public gL_MONFuse_Pgm_Bit() As Long
Public gB_MON_decode_flag As New SiteBoolean
Public gS_MON_CRC_HexStr As New SiteVariant
Public gS_MON_Direct_Access_Str As New SiteVariant

''''---------------------------------------------------------------------------------------------------
'''' CMP Fuse (Compare with UDR)
''''---------------------------------------------------------------------------------------------------
Public gI_CMP_catename_maxLen As Long
Public gB_CMP_decode_flag As New SiteBoolean
Public gL_CMP_PGM_BITS As Long
Public gL_CMP_TOTAL_BITS As Long

'[ Define Fail Bit count (FBC) for all eFuse circuits   ]
'==========================================================
Public gL_ECID_FBC As New SiteLong
Public gL_CFG_FBC As New SiteLong
Public gL_UID_FBC As New SiteLong
Public gL_SEN_FBC As New SiteLong
Public gL_MON_FBC As New SiteLong

''''20150617 add
Public gL_CFG_eFuse_Revision As Long
Public gL_CFG_Device_Revision As Long
Public gL_UDR_eFuse_Revision As Long

'===============================================

''''---------------------------------------------------------------------------------------------------
'''' WAT Application Variables
''''---------------------------------------------------------------------------------------------------
''''20160930 move here from LIB_EFUSE_WAT
Public gS_currWATFileName As String
Public gS_xxxxLotWfID As String ''''means it's X- (previous).
Public gS_currLotWfID As String ''''means it's current.
Public gS_currLotID As String   ''''means it's current.

''''---------------------------------------------------------------------------------------------------
'''' UDRE,UDRP Fuse
''''---------------------------------------------------------------------------------------------------
Public gI_UDRE_catename_maxLen As Long
Public gI_UDRP_catename_maxLen As Long
Public gL_UDRE_eFuse_Revision As Long
Public gL_UDRP_eFuse_Revision As Long
Public gS_UDRE_USI_BitStr As New SiteVariant
Public gS_UDRP_USI_BitStr As New SiteVariant
Public gB_UDRE_decode_flag As New SiteBoolean
Public gB_UDRP_decode_flag As New SiteBoolean
''''---------------------------------------------------------------------------------------------------
'''' CMPE Fuse (Compare with UDRE)
''''---------------------------------------------------------------------------------------------------
Public gI_CMPE_catename_maxLen As Long
Public gB_CMPE_decode_flag As New SiteBoolean
Public gL_CMPE_PGM_BITS As Long
Public gL_CMPE_TOTAL_BITS As Long
''''---------------------------------------------------------------------------------------------------
'''' CMPP Fuse (Compare with UDRP)
''''---------------------------------------------------------------------------------------------------
Public gI_CMPP_catename_maxLen As Long
Public gB_CMPP_decode_flag As New SiteBoolean
Public gL_CMPP_PGM_BITS As Long
Public gL_CMPP_TOTAL_BITS As Long



''''---------------------------------------------------------------------------------------------------
''''---------------------------------------------------------------------------------------------------
''''201812XX New Method for PTE/TTR (Trial Run)''''201811XX, 201808XX
''''---------------------------------------------------------------------------------------------------
'''' Define global Fuse Bits as DspWave
''''---------------------------------------------------------------------------------------------------
Public gB_eFuse_newMethod As Boolean
Public gB_eFuse_printBitMap As Boolean   ''''Pgm/Read BitMap
Public gB_eFuse_printPgmCate As Boolean  ''''Pgm  Category
Public gB_eFuse_printReadCate As Boolean ''''Read Category
Public gB_eFuse_DSPMode As Boolean
Public gL_eFuse_Sim_Blank As Long

Public gL_ECID_msbbit_arr() As Long
Public gL_ECID_lsbbit_arr() As Long
Public gL_ECID_bitwidth_arr() As Long
Public gL_ECID_DefaultOrReal_arr() As Long
Public gL_ECID_stage_bitFlag_arr() As Long
Public gL_ECID_stage_early_bitFlag_arr() As Long
Public gL_ECID_allDefaultBits_arr() As Long
Public gL_ECID_stageLEQjob_bitFlag_arr() As Long
Public gL_CFG_stage_real_bitFlag_arr() As Long
Public gL_CFG_msbbit_arr() As Long
Public gL_CFG_lsbbit_arr() As Long
Public gL_CFG_bitwidth_arr() As Long
Public gL_CFG_DefaultOrReal_arr() As Long
Public gL_CFG_stage_bitFlag_arr() As Long
Public gL_CFG_stage_early_bitFlag_arr() As Long
'Public gL_CFG_SegCNT As Long
Public EConfigTotalBitCount_Seg As Long
Public EFUSE_DVRV_ENABLE As Boolean

Public gL_CFG_allDefaultBits_arr() As Long
Public gL_CFG_stageLEQjob_bitFlag_arr() As Long
Public gL_CFG_SegFlag_arr() As Long

''201904 Ter---------------------------
Public gL_UID_msbbit_arr() As Long
Public gL_UID_lsbbit_arr() As Long
Public gL_UID_bitwidth_arr() As Long
Public gL_UID_DefaultOrReal_arr() As Long
Public gL_UID_stage_bitFlag_arr() As Long
Public gL_UID_stage_early_bitFlag_arr() As Long
Public gL_UID_allDefaultBits_arr() As Long
Public gL_UID_stageLEQjob_bitFlag_arr() As Long

Public gL_MON_msbbit_arr() As Long
Public gL_MON_lsbbit_arr() As Long
Public gL_MON_bitwidth_arr() As Long
Public gL_MON_DefaultOrReal_arr() As Long
Public gL_MON_stage_bitFlag_arr() As Long
Public gL_MON_stage_early_bitFlag_arr() As Long
Public gL_MON_allDefaultBits_arr() As Long
Public gL_MON_stageLEQjob_bitFlag_arr() As Long

Public gL_SEN_msbbit_arr() As Long
Public gL_SEN_lsbbit_arr() As Long
Public gL_SEN_bitwidth_arr() As Long
Public gL_SEN_DefaultOrReal_arr() As Long
Public gL_SEN_stage_bitFlag_arr() As Long
Public gL_SEN_stage_early_bitFlag_arr() As Long
Public gL_SEN_allDefaultBits_arr() As Long
Public gL_SEN_stageLEQjob_bitFlag_arr() As Long

Public gL_UDR_msbbit_arr() As Long
Public gL_UDR_lsbbit_arr() As Long
Public gL_UDR_bitwidth_arr() As Long
Public gL_UDR_DefaultOrReal_arr() As Long
Public gL_UDR_stage_bitFlag_arr() As Long
Public gL_UDR_stage_early_bitFlag_arr() As Long
Public gL_UDR_allDefaultBits_arr() As Long
Public gL_UDR_stageLEQjob_bitFlag_arr() As Long

Public gL_UDRP_msbbit_arr() As Long
Public gL_UDRP_lsbbit_arr() As Long
Public gL_UDRP_bitwidth_arr() As Long
Public gL_UDRP_DefaultOrReal_arr() As Long
Public gL_UDRP_stage_bitFlag_arr() As Long
Public gL_UDRP_stage_early_bitFlag_arr() As Long
Public gL_UDRP_allDefaultBits_arr() As Long
Public gL_UDRP_stageLEQjob_bitFlag_arr() As Long

Public gL_UDRE_msbbit_arr() As Long
Public gL_UDRE_lsbbit_arr() As Long
Public gL_UDRE_bitwidth_arr() As Long
Public gL_UDRE_DefaultOrReal_arr() As Long
Public gL_UDRE_stage_bitFlag_arr() As Long
Public gL_UDRE_stage_early_bitFlag_arr() As Long
Public gL_UDRE_allDefaultBits_arr() As Long
Public gL_UDRE_stageLEQjob_bitFlag_arr() As Long

Public gL_CMP_msbbit_arr() As Long
Public gL_CMP_lsbbit_arr() As Long
Public gL_CMP_bitwidth_arr() As Long
Public gL_CMP_DefaultOrReal_arr() As Long
Public gL_CMP_stage_bitFlag_arr() As Long
Public gL_CMP_stage_early_bitFlag_arr() As Long
Public gL_CMP_allDefaultBits_arr() As Long
Public gL_CMP_stageLEQjob_bitFlag_arr() As Long

Public gL_CMPP_msbbit_arr() As Long
Public gL_CMPP_lsbbit_arr() As Long
Public gL_CMPP_bitwidth_arr() As Long
Public gL_CMPP_DefaultOrReal_arr() As Long
Public gL_CMPP_stage_bitFlag_arr() As Long
Public gL_CMPP_stage_early_bitFlag_arr() As Long
Public gL_CMPP_allDefaultBits_arr() As Long
Public gL_CMPP_stageLEQjob_bitFlag_arr() As Long

Public gL_CMPE_msbbit_arr() As Long
Public gL_CMPE_lsbbit_arr() As Long
Public gL_CMPE_bitwidth_arr() As Long
Public gL_CMPE_DefaultOrReal_arr() As Long
Public gL_CMPE_stage_bitFlag_arr() As Long
Public gL_CMPE_stage_early_bitFlag_arr() As Long
Public gL_CMPE_allDefaultBits_arr() As Long
Public gL_CMPE_stageLEQjob_bitFlag_arr() As Long

''-------------------------------------

Public Enum eFuseOrientation
    eFuse_UP2DOWN = 0
    
    eFuse_1_Bit = 1
    eFuse_SingleUp = 1
    
    eFuse_2_Bit = 2
    eFuse_RIGHT2LEFT = 2
    
    ''''undefined, should be not happened
    eFuse_SingleDown = 3
    eFuse_SingleRight = 4
    eFuse_SingleLeft = 5
    eFuse_Orient_Unknown = 999
End Enum
Public gE_eFuse_Orientation As eFuseOrientation
Public gE_eFuse_Orientation_1Bits As eFuseOrientation

Public Enum eFuseBlockType
    eFuse_ECID = 1
    eFuse_CFG = 2
    eFuse_UID = 3
    eFuse_SEN = 4
    eFuse_MON = 5
    eFuse_UDR = 6
    eFuse_UDRE = 7
    eFuse_UDRP = 8
    eFuse_CMP = 9
    eFuse_CMPE = 10
    eFuse_CMPP = 11
    eFuse_CFGTab = 20
    eFuse_CFGCond = 21
    eFuse_Block_Unknown = 999
End Enum
Public gE_eFuseBlockType As eFuseBlockType

