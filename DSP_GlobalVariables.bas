Attribute VB_Name = "DSP_GlobalVariables"
Option Explicit

''''global DSP Site Variables
Private gDL_eFuse_Orientation_ As New SiteLong
Private gDB_SerialType_ As New SiteBoolean
Private gDL_BitsPerRow_ As New SiteLong
Private gDL_ReadCycles_ As New SiteLong
Private gDL_BitsPerCycle_ As New SiteLong
Private gDL_BitsPerBlock_ As New SiteLong
Private gDL_TotalBits_ As New SiteLong
Private gDL_DigSrcRepeatN_ As New SiteLong
Private gDD_BaseVoltage_ As New SiteDouble
Private gDD_BaseStepVoltage_ As New SiteDouble
Private gDL_CRC_EndBit_ As New SiteLong

Private gDW_Pgm_RawBitWave_ As New DSPWave
Private gDW_ECID_CRC_calcBitsWave_ As New DSPWave
Private gDW_CFG_CRC_calcBitsWave_ As New DSPWave
Private gDW_UID_CRC_calcBitsWave_ As New DSPWave
Private gDW_MON_CRC_calcBitsWave_ As New DSPWave
Private gDW_SEN_CRC_calcBitsWave_ As New DSPWave
Private gDW_Pgm_BitWaveForCRCCalc_ As New DSPWave
Private gDW_Read_BitWaveForCRCCalc_ As New DSPWave

''''global DSP DSPWave
Private gDW_ECID_MSBBit_Cate_ As New DSPWave
Private gDW_ECID_LSBBit_Cate_ As New DSPWave
Private gDW_ECID_BitWidth_Cate_ As New DSPWave
Private gDW_ECID_DefaultReal_Cate_ As New DSPWave
Private gDW_ECID_Stage_BitFlag_ As New DSPWave
Private gDW_ECID_Stage_Early_BitFlag_ As New DSPWave
Private gDW_ECID_allDefaultBitWave_ As New DSPWave
Private gDW_ECID_Read_Decimal_Cate_ As New DSPWave
Private gDW_ECID_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_ECID_Pgm_SingleBitWave_ As New DSPWave
Private gDW_ECID_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_ECID_Read_SingleBitWave_ As New DSPWave
Private gDW_ECID_Read_DoubleBitWave_ As New DSPWave
Private gDW_ECID_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

Private gDW_CFG_MSBBit_Cate_ As New DSPWave
Private gDW_CFG_LSBBit_Cate_ As New DSPWave
Private gDW_CFG_BitWidth_Cate_ As New DSPWave
Private gDW_CFG_DefaultReal_Cate_ As New DSPWave
Private gDW_CFG_Stage_BitFlag_ As New DSPWave
Private gDW_CFG_Stage_Early_BitFlag_ As New DSPWave
Private gDW_CFG_Stage_Real_BitFlag_ As New DSPWave
Private gDW_CFG_allDefaultBitWave_ As New DSPWave
Private gDW_CFG_Read_Decimal_Cate_ As New DSPWave
Private gDW_CFG_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_CFG_Pgm_SingleBitWave_ As New DSPWave
Private gDW_CFG_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_CFG_Read_SingleBitWave_ As New DSPWave
Private gDW_CFG_Read_DoubleBitWave_ As New DSPWave
Private gDW_CFG_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)
Private gDW_CFG_SegFlag_ As New DSPWave

''201904 Ter
''#########################################################
'' 201901xx add
Private gDW_UID_MSBBit_Cate_ As New DSPWave
Private gDW_UID_LSBBit_Cate_ As New DSPWave
Private gDW_UID_BitWidth_Cate_ As New DSPWave
Private gDW_UID_DefaultReal_Cate_ As New DSPWave
Private gDW_UID_Stage_BitFlag_ As New DSPWave
Private gDW_UID_Stage_Early_BitFlag_ As New DSPWave
Private gDW_UID_allDefaultBitWave_ As New DSPWave
Private gDW_UID_Read_Decimal_Cate_ As New DSPWave
Private gDW_UID_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_UID_Pgm_SingleBitWave_ As New DSPWave
Private gDW_UID_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_UID_Read_SingleBitWave_ As New DSPWave
Private gDW_UID_Read_DoubleBitWave_ As New DSPWave
Private gDW_UID_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

Private gDW_UDR_MSBBit_Cate_ As New DSPWave
Private gDW_UDR_LSBBit_Cate_ As New DSPWave
Private gDW_UDR_BitWidth_Cate_ As New DSPWave
Private gDW_UDR_DefaultReal_Cate_ As New DSPWave
Private gDW_UDR_Stage_BitFlag_ As New DSPWave
Private gDW_UDR_Stage_Early_BitFlag_ As New DSPWave
Private gDW_UDR_allDefaultBitWave_ As New DSPWave
Private gDW_UDR_Read_Decimal_Cate_ As New DSPWave
Private gDW_UDR_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_UDR_Pgm_SingleBitWave_ As New DSPWave
Private gDW_UDR_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_UDR_Read_SingleBitWave_ As New DSPWave
Private gDW_UDR_Read_DoubleBitWave_ As New DSPWave
Private gDW_UDR_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

Private gDW_SEN_MSBBit_Cate_ As New DSPWave
Private gDW_SEN_LSBBit_Cate_ As New DSPWave
Private gDW_SEN_BitWidth_Cate_ As New DSPWave
Private gDW_SEN_DefaultReal_Cate_ As New DSPWave
Private gDW_SEN_Stage_BitFlag_ As New DSPWave
Private gDW_SEN_Stage_Early_BitFlag_ As New DSPWave
Private gDW_SEN_allDefaultBitWave_ As New DSPWave
Private gDW_SEN_Read_Decimal_Cate_ As New DSPWave
Private gDW_SEN_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_SEN_Pgm_SingleBitWave_ As New DSPWave
Private gDW_SEN_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_SEN_Read_SingleBitWave_ As New DSPWave
Private gDW_SEN_Read_DoubleBitWave_ As New DSPWave
Private gDW_SEN_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

Private gDW_MON_MSBBit_Cate_ As New DSPWave
Private gDW_MON_LSBBit_Cate_ As New DSPWave
Private gDW_MON_BitWidth_Cate_ As New DSPWave
Private gDW_MON_DefaultReal_Cate_ As New DSPWave
Private gDW_MON_Stage_BitFlag_ As New DSPWave
Private gDW_MON_Stage_Early_BitFlag_ As New DSPWave
Private gDW_MON_allDefaultBitWave_ As New DSPWave
Private gDW_MON_Read_Decimal_Cate_ As New DSPWave
Private gDW_MON_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_MON_Pgm_SingleBitWave_ As New DSPWave
Private gDW_MON_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_MON_Read_SingleBitWave_ As New DSPWave
Private gDW_MON_Read_DoubleBitWave_ As New DSPWave
Private gDW_MON_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

Private gDW_CMP_MSBBit_Cate_ As New DSPWave
Private gDW_CMP_LSBBit_Cate_ As New DSPWave
Private gDW_CMP_BitWidth_Cate_ As New DSPWave
Private gDW_CMP_DefaultReal_Cate_ As New DSPWave
Private gDW_CMP_Stage_BitFlag_ As New DSPWave
Private gDW_CMP_Stage_Early_BitFlag_ As New DSPWave
Private gDW_CMP_allDefaultBitWave_ As New DSPWave
Private gDW_CMP_Read_Decimal_Cate_ As New DSPWave
Private gDW_CMP_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_CMP_Pgm_SingleBitWave_ As New DSPWave
Private gDW_CMP_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_CMP_Read_SingleBitWave_ As New DSPWave
Private gDW_CMP_Read_DoubleBitWave_ As New DSPWave
Private gDW_CMP_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

Private gDW_UDRE_MSBBit_Cate_ As New DSPWave
Private gDW_UDRE_LSBBit_Cate_ As New DSPWave
Private gDW_UDRE_BitWidth_Cate_ As New DSPWave
Private gDW_UDRE_DefaultReal_Cate_ As New DSPWave
Private gDW_UDRE_Stage_BitFlag_ As New DSPWave
Private gDW_UDRE_Stage_Early_BitFlag_ As New DSPWave
Private gDW_UDRE_allDefaultBitWave_ As New DSPWave
Private gDW_UDRE_Read_Decimal_Cate_ As New DSPWave
Private gDW_UDRE_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_UDRE_Pgm_SingleBitWave_ As New DSPWave
Private gDW_UDRE_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_UDRE_Read_SingleBitWave_ As New DSPWave
Private gDW_UDRE_Read_DoubleBitWave_ As New DSPWave
Private gDW_UDRE_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

Private gDW_UDRP_MSBBit_Cate_ As New DSPWave
Private gDW_UDRP_LSBBit_Cate_ As New DSPWave
Private gDW_UDRP_BitWidth_Cate_ As New DSPWave
Private gDW_UDRP_DefaultReal_Cate_ As New DSPWave
Private gDW_UDRP_Stage_BitFlag_ As New DSPWave
Private gDW_UDRP_Stage_Early_BitFlag_ As New DSPWave
Private gDW_UDRP_allDefaultBitWave_ As New DSPWave
Private gDW_UDRP_Read_Decimal_Cate_ As New DSPWave
Private gDW_UDRP_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_UDRP_Pgm_SingleBitWave_ As New DSPWave
Private gDW_UDRP_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_UDRP_Read_SingleBitWave_ As New DSPWave
Private gDW_UDRP_Read_DoubleBitWave_ As New DSPWave
Private gDW_UDRP_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

Private gDW_CMPE_MSBBit_Cate_ As New DSPWave
Private gDW_CMPE_LSBBit_Cate_ As New DSPWave
Private gDW_CMPE_BitWidth_Cate_ As New DSPWave
Private gDW_CMPE_DefaultReal_Cate_ As New DSPWave
Private gDW_CMPE_Stage_BitFlag_ As New DSPWave
Private gDW_CMPE_Stage_Early_BitFlag_ As New DSPWave
Private gDW_CMPE_allDefaultBitWave_ As New DSPWave
Private gDW_CMPE_Read_Decimal_Cate_ As New DSPWave
Private gDW_CMPE_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_CMPE_Pgm_SingleBitWave_ As New DSPWave
Private gDW_CMPE_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_CMPE_Read_SingleBitWave_ As New DSPWave
Private gDW_CMPE_Read_DoubleBitWave_ As New DSPWave
Private gDW_CMPE_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

Private gDW_CMPP_MSBBit_Cate_ As New DSPWave
Private gDW_CMPP_LSBBit_Cate_ As New DSPWave
Private gDW_CMPP_BitWidth_Cate_ As New DSPWave
Private gDW_CMPP_DefaultReal_Cate_ As New DSPWave
Private gDW_CMPP_Stage_BitFlag_ As New DSPWave
Private gDW_CMPP_Stage_Early_BitFlag_ As New DSPWave
Private gDW_CMPP_allDefaultBitWave_ As New DSPWave
Private gDW_CMPP_Read_Decimal_Cate_ As New DSPWave
Private gDW_CMPP_Read_cmpsgWavePerCyc_ As New DSPWave
Private gDW_CMPP_Pgm_SingleBitWave_ As New DSPWave
Private gDW_CMPP_Pgm_DoubleBitWave_ As New DSPWave
Private gDW_CMPP_Read_SingleBitWave_ As New DSPWave
Private gDW_CMPP_Read_DoubleBitWave_ As New DSPWave
Private gDW_CMPP_StageLEQJob_BitFlag_ As New DSPWave ''''stage less equal Job BitFlag (Stage<= Job)

''#########################################################

'
' ****************************************************
' *** DO NOT ALTER OR ADD ANY CODE BELOW THIS LINE ***
' ***  THE CODE BELOW IS AUTOMATICALLY GENERATED   ***
' Place the declarations for DSP global variables above
' this comment using the following format:
' Private Name_ As New Type
'
Private tl_DSPGlobalVariableCheckVar_ As New SiteLong
Private Function tl_DSPGlobalVariableInitializer() As Long
    tl_DSPGlobalVariableInitializer = tl_DSPGlobalVariableCheckVar_
    tl_DSPGlobalVariableCheckVar_ = 1
End Function

' Called to also reset VBA variables when embedded variables need to be reset, e.g., after re-validation.
Public Sub tlDSPGlobalVariableReset()
    Set gDL_eFuse_Orientation_ = Nothing
    Set gDB_SerialType_ = Nothing
    Set gDL_BitsPerRow_ = Nothing
    Set gDL_ReadCycles_ = Nothing
    Set gDL_BitsPerCycle_ = Nothing
    Set gDL_BitsPerBlock_ = Nothing
    Set gDL_TotalBits_ = Nothing
    Set gDL_DigSrcRepeatN_ = Nothing
    Set gDD_BaseVoltage_ = Nothing
    Set gDD_BaseStepVoltage_ = Nothing
    Set gDL_CRC_EndBit_ = Nothing
    Set gDW_Pgm_RawBitWave_ = Nothing
    Set gDW_ECID_CRC_calcBitsWave_ = Nothing
    Set gDW_CFG_CRC_calcBitsWave_ = Nothing
    Set gDW_UID_CRC_calcBitsWave_ = Nothing
    Set gDW_MON_CRC_calcBitsWave_ = Nothing
    Set gDW_SEN_CRC_calcBitsWave_ = Nothing
    Set gDW_Pgm_BitWaveForCRCCalc_ = Nothing
    Set gDW_Read_BitWaveForCRCCalc_ = Nothing
    Set gDW_ECID_MSBBit_Cate_ = Nothing
    Set gDW_ECID_LSBBit_Cate_ = Nothing
    Set gDW_ECID_BitWidth_Cate_ = Nothing
    Set gDW_ECID_DefaultReal_Cate_ = Nothing
    Set gDW_ECID_Stage_BitFlag_ = Nothing
    Set gDW_ECID_Stage_Early_BitFlag_ = Nothing
    Set gDW_ECID_allDefaultBitWave_ = Nothing
    Set gDW_ECID_Read_Decimal_Cate_ = Nothing
    Set gDW_ECID_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_ECID_Pgm_SingleBitWave_ = Nothing
    Set gDW_ECID_Pgm_DoubleBitWave_ = Nothing
    Set gDW_ECID_Read_SingleBitWave_ = Nothing
    Set gDW_ECID_Read_DoubleBitWave_ = Nothing
    Set gDW_ECID_StageLEQJob_BitFlag_ = Nothing
    Set gDW_CFG_MSBBit_Cate_ = Nothing
    Set gDW_CFG_LSBBit_Cate_ = Nothing
    Set gDW_CFG_BitWidth_Cate_ = Nothing
    Set gDW_CFG_DefaultReal_Cate_ = Nothing
    Set gDW_CFG_Stage_BitFlag_ = Nothing
    Set gDW_CFG_Stage_Early_BitFlag_ = Nothing
    Set gDW_CFG_Stage_Real_BitFlag_ = Nothing
    Set gDW_CFG_allDefaultBitWave_ = Nothing
    Set gDW_CFG_Read_Decimal_Cate_ = Nothing
    Set gDW_CFG_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_CFG_Pgm_SingleBitWave_ = Nothing
    Set gDW_CFG_Pgm_DoubleBitWave_ = Nothing
    Set gDW_CFG_Read_SingleBitWave_ = Nothing
    Set gDW_CFG_Read_DoubleBitWave_ = Nothing
    Set gDW_CFG_StageLEQJob_BitFlag_ = Nothing
    Set gDW_CFG_SegFlag_ = Nothing
    Set gDW_UID_MSBBit_Cate_ = Nothing
    Set gDW_UID_LSBBit_Cate_ = Nothing
    Set gDW_UID_BitWidth_Cate_ = Nothing
    Set gDW_UID_DefaultReal_Cate_ = Nothing
    Set gDW_UID_Stage_BitFlag_ = Nothing
    Set gDW_UID_Stage_Early_BitFlag_ = Nothing
    Set gDW_UID_allDefaultBitWave_ = Nothing
    Set gDW_UID_Read_Decimal_Cate_ = Nothing
    Set gDW_UID_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_UID_Pgm_SingleBitWave_ = Nothing
    Set gDW_UID_Pgm_DoubleBitWave_ = Nothing
    Set gDW_UID_Read_SingleBitWave_ = Nothing
    Set gDW_UID_Read_DoubleBitWave_ = Nothing
    Set gDW_UID_StageLEQJob_BitFlag_ = Nothing
    Set gDW_UDR_MSBBit_Cate_ = Nothing
    Set gDW_UDR_LSBBit_Cate_ = Nothing
    Set gDW_UDR_BitWidth_Cate_ = Nothing
    Set gDW_UDR_DefaultReal_Cate_ = Nothing
    Set gDW_UDR_Stage_BitFlag_ = Nothing
    Set gDW_UDR_Stage_Early_BitFlag_ = Nothing
    Set gDW_UDR_allDefaultBitWave_ = Nothing
    Set gDW_UDR_Read_Decimal_Cate_ = Nothing
    Set gDW_UDR_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_UDR_Pgm_SingleBitWave_ = Nothing
    Set gDW_UDR_Pgm_DoubleBitWave_ = Nothing
    Set gDW_UDR_Read_SingleBitWave_ = Nothing
    Set gDW_UDR_Read_DoubleBitWave_ = Nothing
    Set gDW_UDR_StageLEQJob_BitFlag_ = Nothing
    Set gDW_SEN_MSBBit_Cate_ = Nothing
    Set gDW_SEN_LSBBit_Cate_ = Nothing
    Set gDW_SEN_BitWidth_Cate_ = Nothing
    Set gDW_SEN_DefaultReal_Cate_ = Nothing
    Set gDW_SEN_Stage_BitFlag_ = Nothing
    Set gDW_SEN_Stage_Early_BitFlag_ = Nothing
    Set gDW_SEN_allDefaultBitWave_ = Nothing
    Set gDW_SEN_Read_Decimal_Cate_ = Nothing
    Set gDW_SEN_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_SEN_Pgm_SingleBitWave_ = Nothing
    Set gDW_SEN_Pgm_DoubleBitWave_ = Nothing
    Set gDW_SEN_Read_SingleBitWave_ = Nothing
    Set gDW_SEN_Read_DoubleBitWave_ = Nothing
    Set gDW_SEN_StageLEQJob_BitFlag_ = Nothing
    Set gDW_MON_MSBBit_Cate_ = Nothing
    Set gDW_MON_LSBBit_Cate_ = Nothing
    Set gDW_MON_BitWidth_Cate_ = Nothing
    Set gDW_MON_DefaultReal_Cate_ = Nothing
    Set gDW_MON_Stage_BitFlag_ = Nothing
    Set gDW_MON_Stage_Early_BitFlag_ = Nothing
    Set gDW_MON_allDefaultBitWave_ = Nothing
    Set gDW_MON_Read_Decimal_Cate_ = Nothing
    Set gDW_MON_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_MON_Pgm_SingleBitWave_ = Nothing
    Set gDW_MON_Pgm_DoubleBitWave_ = Nothing
    Set gDW_MON_Read_SingleBitWave_ = Nothing
    Set gDW_MON_Read_DoubleBitWave_ = Nothing
    Set gDW_MON_StageLEQJob_BitFlag_ = Nothing
    Set gDW_CMP_MSBBit_Cate_ = Nothing
    Set gDW_CMP_LSBBit_Cate_ = Nothing
    Set gDW_CMP_BitWidth_Cate_ = Nothing
    Set gDW_CMP_DefaultReal_Cate_ = Nothing
    Set gDW_CMP_Stage_BitFlag_ = Nothing
    Set gDW_CMP_Stage_Early_BitFlag_ = Nothing
    Set gDW_CMP_allDefaultBitWave_ = Nothing
    Set gDW_CMP_Read_Decimal_Cate_ = Nothing
    Set gDW_CMP_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_CMP_Pgm_SingleBitWave_ = Nothing
    Set gDW_CMP_Pgm_DoubleBitWave_ = Nothing
    Set gDW_CMP_Read_SingleBitWave_ = Nothing
    Set gDW_CMP_Read_DoubleBitWave_ = Nothing
    Set gDW_CMP_StageLEQJob_BitFlag_ = Nothing
    Set gDW_UDRE_MSBBit_Cate_ = Nothing
    Set gDW_UDRE_LSBBit_Cate_ = Nothing
    Set gDW_UDRE_BitWidth_Cate_ = Nothing
    Set gDW_UDRE_DefaultReal_Cate_ = Nothing
    Set gDW_UDRE_Stage_BitFlag_ = Nothing
    Set gDW_UDRE_Stage_Early_BitFlag_ = Nothing
    Set gDW_UDRE_allDefaultBitWave_ = Nothing
    Set gDW_UDRE_Read_Decimal_Cate_ = Nothing
    Set gDW_UDRE_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_UDRE_Pgm_SingleBitWave_ = Nothing
    Set gDW_UDRE_Pgm_DoubleBitWave_ = Nothing
    Set gDW_UDRE_Read_SingleBitWave_ = Nothing
    Set gDW_UDRE_Read_DoubleBitWave_ = Nothing
    Set gDW_UDRE_StageLEQJob_BitFlag_ = Nothing
    Set gDW_UDRP_MSBBit_Cate_ = Nothing
    Set gDW_UDRP_LSBBit_Cate_ = Nothing
    Set gDW_UDRP_BitWidth_Cate_ = Nothing
    Set gDW_UDRP_DefaultReal_Cate_ = Nothing
    Set gDW_UDRP_Stage_BitFlag_ = Nothing
    Set gDW_UDRP_Stage_Early_BitFlag_ = Nothing
    Set gDW_UDRP_allDefaultBitWave_ = Nothing
    Set gDW_UDRP_Read_Decimal_Cate_ = Nothing
    Set gDW_UDRP_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_UDRP_Pgm_SingleBitWave_ = Nothing
    Set gDW_UDRP_Pgm_DoubleBitWave_ = Nothing
    Set gDW_UDRP_Read_SingleBitWave_ = Nothing
    Set gDW_UDRP_Read_DoubleBitWave_ = Nothing
    Set gDW_UDRP_StageLEQJob_BitFlag_ = Nothing
    Set gDW_CMPE_MSBBit_Cate_ = Nothing
    Set gDW_CMPE_LSBBit_Cate_ = Nothing
    Set gDW_CMPE_BitWidth_Cate_ = Nothing
    Set gDW_CMPE_DefaultReal_Cate_ = Nothing
    Set gDW_CMPE_Stage_BitFlag_ = Nothing
    Set gDW_CMPE_Stage_Early_BitFlag_ = Nothing
    Set gDW_CMPE_allDefaultBitWave_ = Nothing
    Set gDW_CMPE_Read_Decimal_Cate_ = Nothing
    Set gDW_CMPE_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_CMPE_Pgm_SingleBitWave_ = Nothing
    Set gDW_CMPE_Pgm_DoubleBitWave_ = Nothing
    Set gDW_CMPE_Read_SingleBitWave_ = Nothing
    Set gDW_CMPE_Read_DoubleBitWave_ = Nothing
    Set gDW_CMPE_StageLEQJob_BitFlag_ = Nothing
    Set gDW_CMPP_MSBBit_Cate_ = Nothing
    Set gDW_CMPP_LSBBit_Cate_ = Nothing
    Set gDW_CMPP_BitWidth_Cate_ = Nothing
    Set gDW_CMPP_DefaultReal_Cate_ = Nothing
    Set gDW_CMPP_Stage_BitFlag_ = Nothing
    Set gDW_CMPP_Stage_Early_BitFlag_ = Nothing
    Set gDW_CMPP_allDefaultBitWave_ = Nothing
    Set gDW_CMPP_Read_Decimal_Cate_ = Nothing
    Set gDW_CMPP_Read_cmpsgWavePerCyc_ = Nothing
    Set gDW_CMPP_Pgm_SingleBitWave_ = Nothing
    Set gDW_CMPP_Pgm_DoubleBitWave_ = Nothing
    Set gDW_CMPP_Read_SingleBitWave_ = Nothing
    Set gDW_CMPP_Read_DoubleBitWave_ = Nothing
    Set gDW_CMPP_StageLEQJob_BitFlag_ = Nothing
End Sub

Public Property Get gDL_eFuse_Orientation() As SiteLong
    TheHdw.DSP.SyncRead "gDL_eFuse_Orientation", gDL_eFuse_Orientation_
    Set gDL_eFuse_Orientation = gDL_eFuse_Orientation_
End Property

Public Property Let gDL_eFuse_Orientation(RHS As Variant)
    gDL_eFuse_Orientation_ = RHS
    TheHdw.DSP.SyncWrite "gDL_eFuse_Orientation", gDL_eFuse_Orientation_
End Property

Public Property Get gDB_SerialType() As SiteBoolean
    TheHdw.DSP.SyncRead "gDB_SerialType", gDB_SerialType_
    Set gDB_SerialType = gDB_SerialType_
End Property

Public Property Let gDB_SerialType(RHS As Variant)
    gDB_SerialType_ = RHS
    TheHdw.DSP.SyncWrite "gDB_SerialType", gDB_SerialType_
End Property

Public Property Get gDL_BitsPerRow() As SiteLong
    TheHdw.DSP.SyncRead "gDL_BitsPerRow", gDL_BitsPerRow_
    Set gDL_BitsPerRow = gDL_BitsPerRow_
End Property

Public Property Let gDL_BitsPerRow(RHS As Variant)
    gDL_BitsPerRow_ = RHS
    TheHdw.DSP.SyncWrite "gDL_BitsPerRow", gDL_BitsPerRow_
End Property

Public Property Get gDL_ReadCycles() As SiteLong
    TheHdw.DSP.SyncRead "gDL_ReadCycles", gDL_ReadCycles_
    Set gDL_ReadCycles = gDL_ReadCycles_
End Property

Public Property Let gDL_ReadCycles(RHS As Variant)
    gDL_ReadCycles_ = RHS
    TheHdw.DSP.SyncWrite "gDL_ReadCycles", gDL_ReadCycles_
End Property

Public Property Get gDL_BitsPerCycle() As SiteLong
    TheHdw.DSP.SyncRead "gDL_BitsPerCycle", gDL_BitsPerCycle_
    Set gDL_BitsPerCycle = gDL_BitsPerCycle_
End Property

Public Property Let gDL_BitsPerCycle(RHS As Variant)
    gDL_BitsPerCycle_ = RHS
    TheHdw.DSP.SyncWrite "gDL_BitsPerCycle", gDL_BitsPerCycle_
End Property

Public Property Get gDL_BitsPerBlock() As SiteLong
    TheHdw.DSP.SyncRead "gDL_BitsPerBlock", gDL_BitsPerBlock_
    Set gDL_BitsPerBlock = gDL_BitsPerBlock_
End Property

Public Property Let gDL_BitsPerBlock(RHS As Variant)
    gDL_BitsPerBlock_ = RHS
    TheHdw.DSP.SyncWrite "gDL_BitsPerBlock", gDL_BitsPerBlock_
End Property

Public Property Get gDL_TotalBits() As SiteLong
    TheHdw.DSP.SyncRead "gDL_TotalBits", gDL_TotalBits_
    Set gDL_TotalBits = gDL_TotalBits_
End Property

Public Property Let gDL_TotalBits(RHS As Variant)
    gDL_TotalBits_ = RHS
    TheHdw.DSP.SyncWrite "gDL_TotalBits", gDL_TotalBits_
End Property

Public Property Get gDL_DigSrcRepeatN() As SiteLong
    TheHdw.DSP.SyncRead "gDL_DigSrcRepeatN", gDL_DigSrcRepeatN_
    Set gDL_DigSrcRepeatN = gDL_DigSrcRepeatN_
End Property

Public Property Let gDL_DigSrcRepeatN(RHS As Variant)
    gDL_DigSrcRepeatN_ = RHS
    TheHdw.DSP.SyncWrite "gDL_DigSrcRepeatN", gDL_DigSrcRepeatN_
End Property

Public Property Get gDD_BaseVoltage() As SiteDouble
    TheHdw.DSP.SyncRead "gDD_BaseVoltage", gDD_BaseVoltage_
    Set gDD_BaseVoltage = gDD_BaseVoltage_
End Property

Public Property Let gDD_BaseVoltage(RHS As Variant)
    gDD_BaseVoltage_ = RHS
    TheHdw.DSP.SyncWrite "gDD_BaseVoltage", gDD_BaseVoltage_
End Property

Public Property Get gDD_BaseStepVoltage() As SiteDouble
    TheHdw.DSP.SyncRead "gDD_BaseStepVoltage", gDD_BaseStepVoltage_
    Set gDD_BaseStepVoltage = gDD_BaseStepVoltage_
End Property

Public Property Let gDD_BaseStepVoltage(RHS As Variant)
    gDD_BaseStepVoltage_ = RHS
    TheHdw.DSP.SyncWrite "gDD_BaseStepVoltage", gDD_BaseStepVoltage_
End Property

Public Property Get gDL_CRC_EndBit() As SiteLong
    TheHdw.DSP.SyncRead "gDL_CRC_EndBit", gDL_CRC_EndBit_
    Set gDL_CRC_EndBit = gDL_CRC_EndBit_
End Property

Public Property Let gDL_CRC_EndBit(RHS As Variant)
    gDL_CRC_EndBit_ = RHS
    TheHdw.DSP.SyncWrite "gDL_CRC_EndBit", gDL_CRC_EndBit_
End Property

Public Property Get gDW_Pgm_RawBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_Pgm_RawBitWave", gDW_Pgm_RawBitWave_
    Set gDW_Pgm_RawBitWave = gDW_Pgm_RawBitWave_
End Property

Public Property Let gDW_Pgm_RawBitWave(RHS As DSPWave)
    gDW_Pgm_RawBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_Pgm_RawBitWave", gDW_Pgm_RawBitWave_
End Property

Public Property Get gDW_ECID_CRC_calcBitsWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_CRC_calcBitsWave", gDW_ECID_CRC_calcBitsWave_
    Set gDW_ECID_CRC_calcBitsWave = gDW_ECID_CRC_calcBitsWave_
End Property

Public Property Let gDW_ECID_CRC_calcBitsWave(RHS As DSPWave)
    gDW_ECID_CRC_calcBitsWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_CRC_calcBitsWave", gDW_ECID_CRC_calcBitsWave_
End Property

Public Property Get gDW_CFG_CRC_calcBitsWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_CRC_calcBitsWave", gDW_CFG_CRC_calcBitsWave_
    Set gDW_CFG_CRC_calcBitsWave = gDW_CFG_CRC_calcBitsWave_
End Property

Public Property Let gDW_CFG_CRC_calcBitsWave(RHS As DSPWave)
    gDW_CFG_CRC_calcBitsWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_CRC_calcBitsWave", gDW_CFG_CRC_calcBitsWave_
End Property

Public Property Get gDW_UID_CRC_calcBitsWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_CRC_calcBitsWave", gDW_UID_CRC_calcBitsWave_
    Set gDW_UID_CRC_calcBitsWave = gDW_UID_CRC_calcBitsWave_
End Property

Public Property Let gDW_UID_CRC_calcBitsWave(RHS As DSPWave)
    gDW_UID_CRC_calcBitsWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_CRC_calcBitsWave", gDW_UID_CRC_calcBitsWave_
End Property

Public Property Get gDW_MON_CRC_calcBitsWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_CRC_calcBitsWave", gDW_MON_CRC_calcBitsWave_
    Set gDW_MON_CRC_calcBitsWave = gDW_MON_CRC_calcBitsWave_
End Property

Public Property Let gDW_MON_CRC_calcBitsWave(RHS As DSPWave)
    gDW_MON_CRC_calcBitsWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_CRC_calcBitsWave", gDW_MON_CRC_calcBitsWave_
End Property

Public Property Get gDW_SEN_CRC_calcBitsWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_CRC_calcBitsWave", gDW_SEN_CRC_calcBitsWave_
    Set gDW_SEN_CRC_calcBitsWave = gDW_SEN_CRC_calcBitsWave_
End Property

Public Property Let gDW_SEN_CRC_calcBitsWave(RHS As DSPWave)
    gDW_SEN_CRC_calcBitsWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_CRC_calcBitsWave", gDW_SEN_CRC_calcBitsWave_
End Property

Public Property Get gDW_Pgm_BitWaveForCRCCalc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_Pgm_BitWaveForCRCCalc", gDW_Pgm_BitWaveForCRCCalc_
    Set gDW_Pgm_BitWaveForCRCCalc = gDW_Pgm_BitWaveForCRCCalc_
End Property

Public Property Let gDW_Pgm_BitWaveForCRCCalc(RHS As DSPWave)
    gDW_Pgm_BitWaveForCRCCalc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_Pgm_BitWaveForCRCCalc", gDW_Pgm_BitWaveForCRCCalc_
End Property

Public Property Get gDW_Read_BitWaveForCRCCalc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_Read_BitWaveForCRCCalc", gDW_Read_BitWaveForCRCCalc_
    Set gDW_Read_BitWaveForCRCCalc = gDW_Read_BitWaveForCRCCalc_
End Property

Public Property Let gDW_Read_BitWaveForCRCCalc(RHS As DSPWave)
    gDW_Read_BitWaveForCRCCalc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_Read_BitWaveForCRCCalc", gDW_Read_BitWaveForCRCCalc_
End Property

Public Property Get gDW_ECID_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_MSBBit_Cate", gDW_ECID_MSBBit_Cate_
    Set gDW_ECID_MSBBit_Cate = gDW_ECID_MSBBit_Cate_
End Property

Public Property Let gDW_ECID_MSBBit_Cate(RHS As DSPWave)
    gDW_ECID_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_MSBBit_Cate", gDW_ECID_MSBBit_Cate_
End Property

Public Property Get gDW_ECID_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_LSBBit_Cate", gDW_ECID_LSBBit_Cate_
    Set gDW_ECID_LSBBit_Cate = gDW_ECID_LSBBit_Cate_
End Property

Public Property Let gDW_ECID_LSBBit_Cate(RHS As DSPWave)
    gDW_ECID_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_LSBBit_Cate", gDW_ECID_LSBBit_Cate_
End Property

Public Property Get gDW_ECID_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_BitWidth_Cate", gDW_ECID_BitWidth_Cate_
    Set gDW_ECID_BitWidth_Cate = gDW_ECID_BitWidth_Cate_
End Property

Public Property Let gDW_ECID_BitWidth_Cate(RHS As DSPWave)
    gDW_ECID_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_BitWidth_Cate", gDW_ECID_BitWidth_Cate_
End Property

Public Property Get gDW_ECID_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_DefaultReal_Cate", gDW_ECID_DefaultReal_Cate_
    Set gDW_ECID_DefaultReal_Cate = gDW_ECID_DefaultReal_Cate_
End Property

Public Property Let gDW_ECID_DefaultReal_Cate(RHS As DSPWave)
    gDW_ECID_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_DefaultReal_Cate", gDW_ECID_DefaultReal_Cate_
End Property

Public Property Get gDW_ECID_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_Stage_BitFlag", gDW_ECID_Stage_BitFlag_
    Set gDW_ECID_Stage_BitFlag = gDW_ECID_Stage_BitFlag_
End Property

Public Property Let gDW_ECID_Stage_BitFlag(RHS As DSPWave)
    gDW_ECID_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_Stage_BitFlag", gDW_ECID_Stage_BitFlag_
End Property

Public Property Get gDW_ECID_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_Stage_Early_BitFlag", gDW_ECID_Stage_Early_BitFlag_
    Set gDW_ECID_Stage_Early_BitFlag = gDW_ECID_Stage_Early_BitFlag_
End Property

Public Property Let gDW_ECID_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_ECID_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_Stage_Early_BitFlag", gDW_ECID_Stage_Early_BitFlag_
End Property

Public Property Get gDW_ECID_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_allDefaultBitWave", gDW_ECID_allDefaultBitWave_
    Set gDW_ECID_allDefaultBitWave = gDW_ECID_allDefaultBitWave_
End Property

Public Property Let gDW_ECID_allDefaultBitWave(RHS As DSPWave)
    gDW_ECID_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_allDefaultBitWave", gDW_ECID_allDefaultBitWave_
End Property

Public Property Get gDW_ECID_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_Read_Decimal_Cate", gDW_ECID_Read_Decimal_Cate_
    Set gDW_ECID_Read_Decimal_Cate = gDW_ECID_Read_Decimal_Cate_
End Property

Public Property Let gDW_ECID_Read_Decimal_Cate(RHS As DSPWave)
    gDW_ECID_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_Read_Decimal_Cate", gDW_ECID_Read_Decimal_Cate_
End Property

Public Property Get gDW_ECID_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_Read_cmpsgWavePerCyc", gDW_ECID_Read_cmpsgWavePerCyc_
    Set gDW_ECID_Read_cmpsgWavePerCyc = gDW_ECID_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_ECID_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_ECID_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_Read_cmpsgWavePerCyc", gDW_ECID_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_ECID_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_Pgm_SingleBitWave", gDW_ECID_Pgm_SingleBitWave_
    Set gDW_ECID_Pgm_SingleBitWave = gDW_ECID_Pgm_SingleBitWave_
End Property

Public Property Let gDW_ECID_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_ECID_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_Pgm_SingleBitWave", gDW_ECID_Pgm_SingleBitWave_
End Property

Public Property Get gDW_ECID_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_Pgm_DoubleBitWave", gDW_ECID_Pgm_DoubleBitWave_
    Set gDW_ECID_Pgm_DoubleBitWave = gDW_ECID_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_ECID_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_ECID_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_Pgm_DoubleBitWave", gDW_ECID_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_ECID_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_Read_SingleBitWave", gDW_ECID_Read_SingleBitWave_
    Set gDW_ECID_Read_SingleBitWave = gDW_ECID_Read_SingleBitWave_
End Property

Public Property Let gDW_ECID_Read_SingleBitWave(RHS As DSPWave)
    gDW_ECID_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_Read_SingleBitWave", gDW_ECID_Read_SingleBitWave_
End Property

Public Property Get gDW_ECID_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_Read_DoubleBitWave", gDW_ECID_Read_DoubleBitWave_
    Set gDW_ECID_Read_DoubleBitWave = gDW_ECID_Read_DoubleBitWave_
End Property

Public Property Let gDW_ECID_Read_DoubleBitWave(RHS As DSPWave)
    gDW_ECID_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_Read_DoubleBitWave", gDW_ECID_Read_DoubleBitWave_
End Property

Public Property Get gDW_ECID_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_ECID_StageLEQJob_BitFlag", gDW_ECID_StageLEQJob_BitFlag_
    Set gDW_ECID_StageLEQJob_BitFlag = gDW_ECID_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_ECID_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_ECID_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_ECID_StageLEQJob_BitFlag", gDW_ECID_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_CFG_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_MSBBit_Cate", gDW_CFG_MSBBit_Cate_
    Set gDW_CFG_MSBBit_Cate = gDW_CFG_MSBBit_Cate_
End Property

Public Property Let gDW_CFG_MSBBit_Cate(RHS As DSPWave)
    gDW_CFG_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_MSBBit_Cate", gDW_CFG_MSBBit_Cate_
End Property

Public Property Get gDW_CFG_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_LSBBit_Cate", gDW_CFG_LSBBit_Cate_
    Set gDW_CFG_LSBBit_Cate = gDW_CFG_LSBBit_Cate_
End Property

Public Property Let gDW_CFG_LSBBit_Cate(RHS As DSPWave)
    gDW_CFG_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_LSBBit_Cate", gDW_CFG_LSBBit_Cate_
End Property

Public Property Get gDW_CFG_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_BitWidth_Cate", gDW_CFG_BitWidth_Cate_
    Set gDW_CFG_BitWidth_Cate = gDW_CFG_BitWidth_Cate_
End Property

Public Property Let gDW_CFG_BitWidth_Cate(RHS As DSPWave)
    gDW_CFG_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_BitWidth_Cate", gDW_CFG_BitWidth_Cate_
End Property

Public Property Get gDW_CFG_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_DefaultReal_Cate", gDW_CFG_DefaultReal_Cate_
    Set gDW_CFG_DefaultReal_Cate = gDW_CFG_DefaultReal_Cate_
End Property

Public Property Let gDW_CFG_DefaultReal_Cate(RHS As DSPWave)
    gDW_CFG_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_DefaultReal_Cate", gDW_CFG_DefaultReal_Cate_
End Property

Public Property Get gDW_CFG_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_Stage_BitFlag", gDW_CFG_Stage_BitFlag_
    Set gDW_CFG_Stage_BitFlag = gDW_CFG_Stage_BitFlag_
End Property

Public Property Let gDW_CFG_Stage_BitFlag(RHS As DSPWave)
    gDW_CFG_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_Stage_BitFlag", gDW_CFG_Stage_BitFlag_
End Property

Public Property Get gDW_CFG_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_Stage_Early_BitFlag", gDW_CFG_Stage_Early_BitFlag_
    Set gDW_CFG_Stage_Early_BitFlag = gDW_CFG_Stage_Early_BitFlag_
End Property

Public Property Let gDW_CFG_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_CFG_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_Stage_Early_BitFlag", gDW_CFG_Stage_Early_BitFlag_
End Property

Public Property Get gDW_CFG_Stage_Real_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_Stage_Real_BitFlag", gDW_CFG_Stage_Real_BitFlag_
    Set gDW_CFG_Stage_Real_BitFlag = gDW_CFG_Stage_Real_BitFlag_
End Property

Public Property Let gDW_CFG_Stage_Real_BitFlag(RHS As DSPWave)
    gDW_CFG_Stage_Real_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_Stage_Real_BitFlag", gDW_CFG_Stage_Real_BitFlag_
End Property

Public Property Get gDW_CFG_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_allDefaultBitWave", gDW_CFG_allDefaultBitWave_
    Set gDW_CFG_allDefaultBitWave = gDW_CFG_allDefaultBitWave_
End Property

Public Property Let gDW_CFG_allDefaultBitWave(RHS As DSPWave)
    gDW_CFG_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_allDefaultBitWave", gDW_CFG_allDefaultBitWave_
End Property

Public Property Get gDW_CFG_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_Read_Decimal_Cate", gDW_CFG_Read_Decimal_Cate_
    Set gDW_CFG_Read_Decimal_Cate = gDW_CFG_Read_Decimal_Cate_
End Property

Public Property Let gDW_CFG_Read_Decimal_Cate(RHS As DSPWave)
    gDW_CFG_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_Read_Decimal_Cate", gDW_CFG_Read_Decimal_Cate_
End Property

Public Property Get gDW_CFG_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_Read_cmpsgWavePerCyc", gDW_CFG_Read_cmpsgWavePerCyc_
    Set gDW_CFG_Read_cmpsgWavePerCyc = gDW_CFG_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_CFG_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_CFG_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_Read_cmpsgWavePerCyc", gDW_CFG_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_CFG_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_Pgm_SingleBitWave", gDW_CFG_Pgm_SingleBitWave_
    Set gDW_CFG_Pgm_SingleBitWave = gDW_CFG_Pgm_SingleBitWave_
End Property

Public Property Let gDW_CFG_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_CFG_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_Pgm_SingleBitWave", gDW_CFG_Pgm_SingleBitWave_
End Property

Public Property Get gDW_CFG_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_Pgm_DoubleBitWave", gDW_CFG_Pgm_DoubleBitWave_
    Set gDW_CFG_Pgm_DoubleBitWave = gDW_CFG_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_CFG_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_CFG_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_Pgm_DoubleBitWave", gDW_CFG_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_CFG_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_Read_SingleBitWave", gDW_CFG_Read_SingleBitWave_
    Set gDW_CFG_Read_SingleBitWave = gDW_CFG_Read_SingleBitWave_
End Property

Public Property Let gDW_CFG_Read_SingleBitWave(RHS As DSPWave)
    gDW_CFG_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_Read_SingleBitWave", gDW_CFG_Read_SingleBitWave_
End Property

Public Property Get gDW_CFG_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_Read_DoubleBitWave", gDW_CFG_Read_DoubleBitWave_
    Set gDW_CFG_Read_DoubleBitWave = gDW_CFG_Read_DoubleBitWave_
End Property

Public Property Let gDW_CFG_Read_DoubleBitWave(RHS As DSPWave)
    gDW_CFG_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_Read_DoubleBitWave", gDW_CFG_Read_DoubleBitWave_
End Property

Public Property Get gDW_CFG_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_StageLEQJob_BitFlag", gDW_CFG_StageLEQJob_BitFlag_
    Set gDW_CFG_StageLEQJob_BitFlag = gDW_CFG_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_CFG_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_CFG_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_StageLEQJob_BitFlag", gDW_CFG_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_CFG_SegFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CFG_SegFlag", gDW_CFG_SegFlag_
    Set gDW_CFG_SegFlag = gDW_CFG_SegFlag_
End Property

Public Property Let gDW_CFG_SegFlag(RHS As DSPWave)
    gDW_CFG_SegFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CFG_SegFlag", gDW_CFG_SegFlag_
End Property

Public Property Get gDW_UID_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_MSBBit_Cate", gDW_UID_MSBBit_Cate_
    Set gDW_UID_MSBBit_Cate = gDW_UID_MSBBit_Cate_
End Property

Public Property Let gDW_UID_MSBBit_Cate(RHS As DSPWave)
    gDW_UID_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_MSBBit_Cate", gDW_UID_MSBBit_Cate_
End Property

Public Property Get gDW_UID_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_LSBBit_Cate", gDW_UID_LSBBit_Cate_
    Set gDW_UID_LSBBit_Cate = gDW_UID_LSBBit_Cate_
End Property

Public Property Let gDW_UID_LSBBit_Cate(RHS As DSPWave)
    gDW_UID_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_LSBBit_Cate", gDW_UID_LSBBit_Cate_
End Property

Public Property Get gDW_UID_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_BitWidth_Cate", gDW_UID_BitWidth_Cate_
    Set gDW_UID_BitWidth_Cate = gDW_UID_BitWidth_Cate_
End Property

Public Property Let gDW_UID_BitWidth_Cate(RHS As DSPWave)
    gDW_UID_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_BitWidth_Cate", gDW_UID_BitWidth_Cate_
End Property

Public Property Get gDW_UID_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_DefaultReal_Cate", gDW_UID_DefaultReal_Cate_
    Set gDW_UID_DefaultReal_Cate = gDW_UID_DefaultReal_Cate_
End Property

Public Property Let gDW_UID_DefaultReal_Cate(RHS As DSPWave)
    gDW_UID_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_DefaultReal_Cate", gDW_UID_DefaultReal_Cate_
End Property

Public Property Get gDW_UID_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_Stage_BitFlag", gDW_UID_Stage_BitFlag_
    Set gDW_UID_Stage_BitFlag = gDW_UID_Stage_BitFlag_
End Property

Public Property Let gDW_UID_Stage_BitFlag(RHS As DSPWave)
    gDW_UID_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_Stage_BitFlag", gDW_UID_Stage_BitFlag_
End Property

Public Property Get gDW_UID_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_Stage_Early_BitFlag", gDW_UID_Stage_Early_BitFlag_
    Set gDW_UID_Stage_Early_BitFlag = gDW_UID_Stage_Early_BitFlag_
End Property

Public Property Let gDW_UID_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_UID_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_Stage_Early_BitFlag", gDW_UID_Stage_Early_BitFlag_
End Property

Public Property Get gDW_UID_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_allDefaultBitWave", gDW_UID_allDefaultBitWave_
    Set gDW_UID_allDefaultBitWave = gDW_UID_allDefaultBitWave_
End Property

Public Property Let gDW_UID_allDefaultBitWave(RHS As DSPWave)
    gDW_UID_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_allDefaultBitWave", gDW_UID_allDefaultBitWave_
End Property

Public Property Get gDW_UID_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_Read_Decimal_Cate", gDW_UID_Read_Decimal_Cate_
    Set gDW_UID_Read_Decimal_Cate = gDW_UID_Read_Decimal_Cate_
End Property

Public Property Let gDW_UID_Read_Decimal_Cate(RHS As DSPWave)
    gDW_UID_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_Read_Decimal_Cate", gDW_UID_Read_Decimal_Cate_
End Property

Public Property Get gDW_UID_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_Read_cmpsgWavePerCyc", gDW_UID_Read_cmpsgWavePerCyc_
    Set gDW_UID_Read_cmpsgWavePerCyc = gDW_UID_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_UID_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_UID_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_Read_cmpsgWavePerCyc", gDW_UID_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_UID_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_Pgm_SingleBitWave", gDW_UID_Pgm_SingleBitWave_
    Set gDW_UID_Pgm_SingleBitWave = gDW_UID_Pgm_SingleBitWave_
End Property

Public Property Let gDW_UID_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_UID_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_Pgm_SingleBitWave", gDW_UID_Pgm_SingleBitWave_
End Property

Public Property Get gDW_UID_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_Pgm_DoubleBitWave", gDW_UID_Pgm_DoubleBitWave_
    Set gDW_UID_Pgm_DoubleBitWave = gDW_UID_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_UID_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_UID_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_Pgm_DoubleBitWave", gDW_UID_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_UID_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_Read_SingleBitWave", gDW_UID_Read_SingleBitWave_
    Set gDW_UID_Read_SingleBitWave = gDW_UID_Read_SingleBitWave_
End Property

Public Property Let gDW_UID_Read_SingleBitWave(RHS As DSPWave)
    gDW_UID_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_Read_SingleBitWave", gDW_UID_Read_SingleBitWave_
End Property

Public Property Get gDW_UID_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_Read_DoubleBitWave", gDW_UID_Read_DoubleBitWave_
    Set gDW_UID_Read_DoubleBitWave = gDW_UID_Read_DoubleBitWave_
End Property

Public Property Let gDW_UID_Read_DoubleBitWave(RHS As DSPWave)
    gDW_UID_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_Read_DoubleBitWave", gDW_UID_Read_DoubleBitWave_
End Property

Public Property Get gDW_UID_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UID_StageLEQJob_BitFlag", gDW_UID_StageLEQJob_BitFlag_
    Set gDW_UID_StageLEQJob_BitFlag = gDW_UID_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_UID_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_UID_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UID_StageLEQJob_BitFlag", gDW_UID_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_UDR_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_MSBBit_Cate", gDW_UDR_MSBBit_Cate_
    Set gDW_UDR_MSBBit_Cate = gDW_UDR_MSBBit_Cate_
End Property

Public Property Let gDW_UDR_MSBBit_Cate(RHS As DSPWave)
    gDW_UDR_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_MSBBit_Cate", gDW_UDR_MSBBit_Cate_
End Property

Public Property Get gDW_UDR_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_LSBBit_Cate", gDW_UDR_LSBBit_Cate_
    Set gDW_UDR_LSBBit_Cate = gDW_UDR_LSBBit_Cate_
End Property

Public Property Let gDW_UDR_LSBBit_Cate(RHS As DSPWave)
    gDW_UDR_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_LSBBit_Cate", gDW_UDR_LSBBit_Cate_
End Property

Public Property Get gDW_UDR_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_BitWidth_Cate", gDW_UDR_BitWidth_Cate_
    Set gDW_UDR_BitWidth_Cate = gDW_UDR_BitWidth_Cate_
End Property

Public Property Let gDW_UDR_BitWidth_Cate(RHS As DSPWave)
    gDW_UDR_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_BitWidth_Cate", gDW_UDR_BitWidth_Cate_
End Property

Public Property Get gDW_UDR_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_DefaultReal_Cate", gDW_UDR_DefaultReal_Cate_
    Set gDW_UDR_DefaultReal_Cate = gDW_UDR_DefaultReal_Cate_
End Property

Public Property Let gDW_UDR_DefaultReal_Cate(RHS As DSPWave)
    gDW_UDR_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_DefaultReal_Cate", gDW_UDR_DefaultReal_Cate_
End Property

Public Property Get gDW_UDR_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_Stage_BitFlag", gDW_UDR_Stage_BitFlag_
    Set gDW_UDR_Stage_BitFlag = gDW_UDR_Stage_BitFlag_
End Property

Public Property Let gDW_UDR_Stage_BitFlag(RHS As DSPWave)
    gDW_UDR_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_Stage_BitFlag", gDW_UDR_Stage_BitFlag_
End Property

Public Property Get gDW_UDR_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_Stage_Early_BitFlag", gDW_UDR_Stage_Early_BitFlag_
    Set gDW_UDR_Stage_Early_BitFlag = gDW_UDR_Stage_Early_BitFlag_
End Property

Public Property Let gDW_UDR_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_UDR_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_Stage_Early_BitFlag", gDW_UDR_Stage_Early_BitFlag_
End Property

Public Property Get gDW_UDR_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_allDefaultBitWave", gDW_UDR_allDefaultBitWave_
    Set gDW_UDR_allDefaultBitWave = gDW_UDR_allDefaultBitWave_
End Property

Public Property Let gDW_UDR_allDefaultBitWave(RHS As DSPWave)
    gDW_UDR_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_allDefaultBitWave", gDW_UDR_allDefaultBitWave_
End Property

Public Property Get gDW_UDR_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_Read_Decimal_Cate", gDW_UDR_Read_Decimal_Cate_
    Set gDW_UDR_Read_Decimal_Cate = gDW_UDR_Read_Decimal_Cate_
End Property

Public Property Let gDW_UDR_Read_Decimal_Cate(RHS As DSPWave)
    gDW_UDR_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_Read_Decimal_Cate", gDW_UDR_Read_Decimal_Cate_
End Property

Public Property Get gDW_UDR_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_Read_cmpsgWavePerCyc", gDW_UDR_Read_cmpsgWavePerCyc_
    Set gDW_UDR_Read_cmpsgWavePerCyc = gDW_UDR_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_UDR_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_UDR_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_Read_cmpsgWavePerCyc", gDW_UDR_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_UDR_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_Pgm_SingleBitWave", gDW_UDR_Pgm_SingleBitWave_
    Set gDW_UDR_Pgm_SingleBitWave = gDW_UDR_Pgm_SingleBitWave_
End Property

Public Property Let gDW_UDR_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_UDR_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_Pgm_SingleBitWave", gDW_UDR_Pgm_SingleBitWave_
End Property

Public Property Get gDW_UDR_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_Pgm_DoubleBitWave", gDW_UDR_Pgm_DoubleBitWave_
    Set gDW_UDR_Pgm_DoubleBitWave = gDW_UDR_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_UDR_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_UDR_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_Pgm_DoubleBitWave", gDW_UDR_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_UDR_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_Read_SingleBitWave", gDW_UDR_Read_SingleBitWave_
    Set gDW_UDR_Read_SingleBitWave = gDW_UDR_Read_SingleBitWave_
End Property

Public Property Let gDW_UDR_Read_SingleBitWave(RHS As DSPWave)
    gDW_UDR_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_Read_SingleBitWave", gDW_UDR_Read_SingleBitWave_
End Property

Public Property Get gDW_UDR_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_Read_DoubleBitWave", gDW_UDR_Read_DoubleBitWave_
    Set gDW_UDR_Read_DoubleBitWave = gDW_UDR_Read_DoubleBitWave_
End Property

Public Property Let gDW_UDR_Read_DoubleBitWave(RHS As DSPWave)
    gDW_UDR_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_Read_DoubleBitWave", gDW_UDR_Read_DoubleBitWave_
End Property

Public Property Get gDW_UDR_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDR_StageLEQJob_BitFlag", gDW_UDR_StageLEQJob_BitFlag_
    Set gDW_UDR_StageLEQJob_BitFlag = gDW_UDR_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_UDR_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_UDR_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDR_StageLEQJob_BitFlag", gDW_UDR_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_SEN_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_MSBBit_Cate", gDW_SEN_MSBBit_Cate_
    Set gDW_SEN_MSBBit_Cate = gDW_SEN_MSBBit_Cate_
End Property

Public Property Let gDW_SEN_MSBBit_Cate(RHS As DSPWave)
    gDW_SEN_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_MSBBit_Cate", gDW_SEN_MSBBit_Cate_
End Property

Public Property Get gDW_SEN_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_LSBBit_Cate", gDW_SEN_LSBBit_Cate_
    Set gDW_SEN_LSBBit_Cate = gDW_SEN_LSBBit_Cate_
End Property

Public Property Let gDW_SEN_LSBBit_Cate(RHS As DSPWave)
    gDW_SEN_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_LSBBit_Cate", gDW_SEN_LSBBit_Cate_
End Property

Public Property Get gDW_SEN_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_BitWidth_Cate", gDW_SEN_BitWidth_Cate_
    Set gDW_SEN_BitWidth_Cate = gDW_SEN_BitWidth_Cate_
End Property

Public Property Let gDW_SEN_BitWidth_Cate(RHS As DSPWave)
    gDW_SEN_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_BitWidth_Cate", gDW_SEN_BitWidth_Cate_
End Property

Public Property Get gDW_SEN_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_DefaultReal_Cate", gDW_SEN_DefaultReal_Cate_
    Set gDW_SEN_DefaultReal_Cate = gDW_SEN_DefaultReal_Cate_
End Property

Public Property Let gDW_SEN_DefaultReal_Cate(RHS As DSPWave)
    gDW_SEN_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_DefaultReal_Cate", gDW_SEN_DefaultReal_Cate_
End Property

Public Property Get gDW_SEN_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_Stage_BitFlag", gDW_SEN_Stage_BitFlag_
    Set gDW_SEN_Stage_BitFlag = gDW_SEN_Stage_BitFlag_
End Property

Public Property Let gDW_SEN_Stage_BitFlag(RHS As DSPWave)
    gDW_SEN_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_Stage_BitFlag", gDW_SEN_Stage_BitFlag_
End Property

Public Property Get gDW_SEN_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_Stage_Early_BitFlag", gDW_SEN_Stage_Early_BitFlag_
    Set gDW_SEN_Stage_Early_BitFlag = gDW_SEN_Stage_Early_BitFlag_
End Property

Public Property Let gDW_SEN_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_SEN_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_Stage_Early_BitFlag", gDW_SEN_Stage_Early_BitFlag_
End Property

Public Property Get gDW_SEN_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_allDefaultBitWave", gDW_SEN_allDefaultBitWave_
    Set gDW_SEN_allDefaultBitWave = gDW_SEN_allDefaultBitWave_
End Property

Public Property Let gDW_SEN_allDefaultBitWave(RHS As DSPWave)
    gDW_SEN_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_allDefaultBitWave", gDW_SEN_allDefaultBitWave_
End Property

Public Property Get gDW_SEN_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_Read_Decimal_Cate", gDW_SEN_Read_Decimal_Cate_
    Set gDW_SEN_Read_Decimal_Cate = gDW_SEN_Read_Decimal_Cate_
End Property

Public Property Let gDW_SEN_Read_Decimal_Cate(RHS As DSPWave)
    gDW_SEN_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_Read_Decimal_Cate", gDW_SEN_Read_Decimal_Cate_
End Property

Public Property Get gDW_SEN_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_Read_cmpsgWavePerCyc", gDW_SEN_Read_cmpsgWavePerCyc_
    Set gDW_SEN_Read_cmpsgWavePerCyc = gDW_SEN_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_SEN_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_SEN_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_Read_cmpsgWavePerCyc", gDW_SEN_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_SEN_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_Pgm_SingleBitWave", gDW_SEN_Pgm_SingleBitWave_
    Set gDW_SEN_Pgm_SingleBitWave = gDW_SEN_Pgm_SingleBitWave_
End Property

Public Property Let gDW_SEN_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_SEN_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_Pgm_SingleBitWave", gDW_SEN_Pgm_SingleBitWave_
End Property

Public Property Get gDW_SEN_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_Pgm_DoubleBitWave", gDW_SEN_Pgm_DoubleBitWave_
    Set gDW_SEN_Pgm_DoubleBitWave = gDW_SEN_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_SEN_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_SEN_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_Pgm_DoubleBitWave", gDW_SEN_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_SEN_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_Read_SingleBitWave", gDW_SEN_Read_SingleBitWave_
    Set gDW_SEN_Read_SingleBitWave = gDW_SEN_Read_SingleBitWave_
End Property

Public Property Let gDW_SEN_Read_SingleBitWave(RHS As DSPWave)
    gDW_SEN_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_Read_SingleBitWave", gDW_SEN_Read_SingleBitWave_
End Property

Public Property Get gDW_SEN_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_Read_DoubleBitWave", gDW_SEN_Read_DoubleBitWave_
    Set gDW_SEN_Read_DoubleBitWave = gDW_SEN_Read_DoubleBitWave_
End Property

Public Property Let gDW_SEN_Read_DoubleBitWave(RHS As DSPWave)
    gDW_SEN_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_Read_DoubleBitWave", gDW_SEN_Read_DoubleBitWave_
End Property

Public Property Get gDW_SEN_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_SEN_StageLEQJob_BitFlag", gDW_SEN_StageLEQJob_BitFlag_
    Set gDW_SEN_StageLEQJob_BitFlag = gDW_SEN_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_SEN_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_SEN_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_SEN_StageLEQJob_BitFlag", gDW_SEN_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_MON_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_MSBBit_Cate", gDW_MON_MSBBit_Cate_
    Set gDW_MON_MSBBit_Cate = gDW_MON_MSBBit_Cate_
End Property

Public Property Let gDW_MON_MSBBit_Cate(RHS As DSPWave)
    gDW_MON_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_MSBBit_Cate", gDW_MON_MSBBit_Cate_
End Property

Public Property Get gDW_MON_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_LSBBit_Cate", gDW_MON_LSBBit_Cate_
    Set gDW_MON_LSBBit_Cate = gDW_MON_LSBBit_Cate_
End Property

Public Property Let gDW_MON_LSBBit_Cate(RHS As DSPWave)
    gDW_MON_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_LSBBit_Cate", gDW_MON_LSBBit_Cate_
End Property

Public Property Get gDW_MON_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_BitWidth_Cate", gDW_MON_BitWidth_Cate_
    Set gDW_MON_BitWidth_Cate = gDW_MON_BitWidth_Cate_
End Property

Public Property Let gDW_MON_BitWidth_Cate(RHS As DSPWave)
    gDW_MON_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_BitWidth_Cate", gDW_MON_BitWidth_Cate_
End Property

Public Property Get gDW_MON_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_DefaultReal_Cate", gDW_MON_DefaultReal_Cate_
    Set gDW_MON_DefaultReal_Cate = gDW_MON_DefaultReal_Cate_
End Property

Public Property Let gDW_MON_DefaultReal_Cate(RHS As DSPWave)
    gDW_MON_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_DefaultReal_Cate", gDW_MON_DefaultReal_Cate_
End Property

Public Property Get gDW_MON_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_Stage_BitFlag", gDW_MON_Stage_BitFlag_
    Set gDW_MON_Stage_BitFlag = gDW_MON_Stage_BitFlag_
End Property

Public Property Let gDW_MON_Stage_BitFlag(RHS As DSPWave)
    gDW_MON_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_Stage_BitFlag", gDW_MON_Stage_BitFlag_
End Property

Public Property Get gDW_MON_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_Stage_Early_BitFlag", gDW_MON_Stage_Early_BitFlag_
    Set gDW_MON_Stage_Early_BitFlag = gDW_MON_Stage_Early_BitFlag_
End Property

Public Property Let gDW_MON_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_MON_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_Stage_Early_BitFlag", gDW_MON_Stage_Early_BitFlag_
End Property

Public Property Get gDW_MON_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_allDefaultBitWave", gDW_MON_allDefaultBitWave_
    Set gDW_MON_allDefaultBitWave = gDW_MON_allDefaultBitWave_
End Property

Public Property Let gDW_MON_allDefaultBitWave(RHS As DSPWave)
    gDW_MON_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_allDefaultBitWave", gDW_MON_allDefaultBitWave_
End Property

Public Property Get gDW_MON_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_Read_Decimal_Cate", gDW_MON_Read_Decimal_Cate_
    Set gDW_MON_Read_Decimal_Cate = gDW_MON_Read_Decimal_Cate_
End Property

Public Property Let gDW_MON_Read_Decimal_Cate(RHS As DSPWave)
    gDW_MON_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_Read_Decimal_Cate", gDW_MON_Read_Decimal_Cate_
End Property

Public Property Get gDW_MON_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_Read_cmpsgWavePerCyc", gDW_MON_Read_cmpsgWavePerCyc_
    Set gDW_MON_Read_cmpsgWavePerCyc = gDW_MON_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_MON_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_MON_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_Read_cmpsgWavePerCyc", gDW_MON_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_MON_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_Pgm_SingleBitWave", gDW_MON_Pgm_SingleBitWave_
    Set gDW_MON_Pgm_SingleBitWave = gDW_MON_Pgm_SingleBitWave_
End Property

Public Property Let gDW_MON_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_MON_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_Pgm_SingleBitWave", gDW_MON_Pgm_SingleBitWave_
End Property

Public Property Get gDW_MON_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_Pgm_DoubleBitWave", gDW_MON_Pgm_DoubleBitWave_
    Set gDW_MON_Pgm_DoubleBitWave = gDW_MON_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_MON_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_MON_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_Pgm_DoubleBitWave", gDW_MON_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_MON_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_Read_SingleBitWave", gDW_MON_Read_SingleBitWave_
    Set gDW_MON_Read_SingleBitWave = gDW_MON_Read_SingleBitWave_
End Property

Public Property Let gDW_MON_Read_SingleBitWave(RHS As DSPWave)
    gDW_MON_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_Read_SingleBitWave", gDW_MON_Read_SingleBitWave_
End Property

Public Property Get gDW_MON_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_Read_DoubleBitWave", gDW_MON_Read_DoubleBitWave_
    Set gDW_MON_Read_DoubleBitWave = gDW_MON_Read_DoubleBitWave_
End Property

Public Property Let gDW_MON_Read_DoubleBitWave(RHS As DSPWave)
    gDW_MON_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_Read_DoubleBitWave", gDW_MON_Read_DoubleBitWave_
End Property

Public Property Get gDW_MON_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_MON_StageLEQJob_BitFlag", gDW_MON_StageLEQJob_BitFlag_
    Set gDW_MON_StageLEQJob_BitFlag = gDW_MON_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_MON_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_MON_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_MON_StageLEQJob_BitFlag", gDW_MON_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_CMP_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_MSBBit_Cate", gDW_CMP_MSBBit_Cate_
    Set gDW_CMP_MSBBit_Cate = gDW_CMP_MSBBit_Cate_
End Property

Public Property Let gDW_CMP_MSBBit_Cate(RHS As DSPWave)
    gDW_CMP_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_MSBBit_Cate", gDW_CMP_MSBBit_Cate_
End Property

Public Property Get gDW_CMP_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_LSBBit_Cate", gDW_CMP_LSBBit_Cate_
    Set gDW_CMP_LSBBit_Cate = gDW_CMP_LSBBit_Cate_
End Property

Public Property Let gDW_CMP_LSBBit_Cate(RHS As DSPWave)
    gDW_CMP_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_LSBBit_Cate", gDW_CMP_LSBBit_Cate_
End Property

Public Property Get gDW_CMP_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_BitWidth_Cate", gDW_CMP_BitWidth_Cate_
    Set gDW_CMP_BitWidth_Cate = gDW_CMP_BitWidth_Cate_
End Property

Public Property Let gDW_CMP_BitWidth_Cate(RHS As DSPWave)
    gDW_CMP_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_BitWidth_Cate", gDW_CMP_BitWidth_Cate_
End Property

Public Property Get gDW_CMP_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_DefaultReal_Cate", gDW_CMP_DefaultReal_Cate_
    Set gDW_CMP_DefaultReal_Cate = gDW_CMP_DefaultReal_Cate_
End Property

Public Property Let gDW_CMP_DefaultReal_Cate(RHS As DSPWave)
    gDW_CMP_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_DefaultReal_Cate", gDW_CMP_DefaultReal_Cate_
End Property

Public Property Get gDW_CMP_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_Stage_BitFlag", gDW_CMP_Stage_BitFlag_
    Set gDW_CMP_Stage_BitFlag = gDW_CMP_Stage_BitFlag_
End Property

Public Property Let gDW_CMP_Stage_BitFlag(RHS As DSPWave)
    gDW_CMP_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_Stage_BitFlag", gDW_CMP_Stage_BitFlag_
End Property

Public Property Get gDW_CMP_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_Stage_Early_BitFlag", gDW_CMP_Stage_Early_BitFlag_
    Set gDW_CMP_Stage_Early_BitFlag = gDW_CMP_Stage_Early_BitFlag_
End Property

Public Property Let gDW_CMP_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_CMP_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_Stage_Early_BitFlag", gDW_CMP_Stage_Early_BitFlag_
End Property

Public Property Get gDW_CMP_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_allDefaultBitWave", gDW_CMP_allDefaultBitWave_
    Set gDW_CMP_allDefaultBitWave = gDW_CMP_allDefaultBitWave_
End Property

Public Property Let gDW_CMP_allDefaultBitWave(RHS As DSPWave)
    gDW_CMP_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_allDefaultBitWave", gDW_CMP_allDefaultBitWave_
End Property

Public Property Get gDW_CMP_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_Read_Decimal_Cate", gDW_CMP_Read_Decimal_Cate_
    Set gDW_CMP_Read_Decimal_Cate = gDW_CMP_Read_Decimal_Cate_
End Property

Public Property Let gDW_CMP_Read_Decimal_Cate(RHS As DSPWave)
    gDW_CMP_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_Read_Decimal_Cate", gDW_CMP_Read_Decimal_Cate_
End Property

Public Property Get gDW_CMP_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_Read_cmpsgWavePerCyc", gDW_CMP_Read_cmpsgWavePerCyc_
    Set gDW_CMP_Read_cmpsgWavePerCyc = gDW_CMP_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_CMP_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_CMP_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_Read_cmpsgWavePerCyc", gDW_CMP_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_CMP_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_Pgm_SingleBitWave", gDW_CMP_Pgm_SingleBitWave_
    Set gDW_CMP_Pgm_SingleBitWave = gDW_CMP_Pgm_SingleBitWave_
End Property

Public Property Let gDW_CMP_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_CMP_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_Pgm_SingleBitWave", gDW_CMP_Pgm_SingleBitWave_
End Property

Public Property Get gDW_CMP_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_Pgm_DoubleBitWave", gDW_CMP_Pgm_DoubleBitWave_
    Set gDW_CMP_Pgm_DoubleBitWave = gDW_CMP_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_CMP_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_CMP_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_Pgm_DoubleBitWave", gDW_CMP_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_CMP_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_Read_SingleBitWave", gDW_CMP_Read_SingleBitWave_
    Set gDW_CMP_Read_SingleBitWave = gDW_CMP_Read_SingleBitWave_
End Property

Public Property Let gDW_CMP_Read_SingleBitWave(RHS As DSPWave)
    gDW_CMP_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_Read_SingleBitWave", gDW_CMP_Read_SingleBitWave_
End Property

Public Property Get gDW_CMP_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_Read_DoubleBitWave", gDW_CMP_Read_DoubleBitWave_
    Set gDW_CMP_Read_DoubleBitWave = gDW_CMP_Read_DoubleBitWave_
End Property

Public Property Let gDW_CMP_Read_DoubleBitWave(RHS As DSPWave)
    gDW_CMP_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_Read_DoubleBitWave", gDW_CMP_Read_DoubleBitWave_
End Property

Public Property Get gDW_CMP_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMP_StageLEQJob_BitFlag", gDW_CMP_StageLEQJob_BitFlag_
    Set gDW_CMP_StageLEQJob_BitFlag = gDW_CMP_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_CMP_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_CMP_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMP_StageLEQJob_BitFlag", gDW_CMP_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_UDRE_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_MSBBit_Cate", gDW_UDRE_MSBBit_Cate_
    Set gDW_UDRE_MSBBit_Cate = gDW_UDRE_MSBBit_Cate_
End Property

Public Property Let gDW_UDRE_MSBBit_Cate(RHS As DSPWave)
    gDW_UDRE_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_MSBBit_Cate", gDW_UDRE_MSBBit_Cate_
End Property

Public Property Get gDW_UDRE_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_LSBBit_Cate", gDW_UDRE_LSBBit_Cate_
    Set gDW_UDRE_LSBBit_Cate = gDW_UDRE_LSBBit_Cate_
End Property

Public Property Let gDW_UDRE_LSBBit_Cate(RHS As DSPWave)
    gDW_UDRE_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_LSBBit_Cate", gDW_UDRE_LSBBit_Cate_
End Property

Public Property Get gDW_UDRE_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_BitWidth_Cate", gDW_UDRE_BitWidth_Cate_
    Set gDW_UDRE_BitWidth_Cate = gDW_UDRE_BitWidth_Cate_
End Property

Public Property Let gDW_UDRE_BitWidth_Cate(RHS As DSPWave)
    gDW_UDRE_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_BitWidth_Cate", gDW_UDRE_BitWidth_Cate_
End Property

Public Property Get gDW_UDRE_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_DefaultReal_Cate", gDW_UDRE_DefaultReal_Cate_
    Set gDW_UDRE_DefaultReal_Cate = gDW_UDRE_DefaultReal_Cate_
End Property

Public Property Let gDW_UDRE_DefaultReal_Cate(RHS As DSPWave)
    gDW_UDRE_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_DefaultReal_Cate", gDW_UDRE_DefaultReal_Cate_
End Property

Public Property Get gDW_UDRE_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_Stage_BitFlag", gDW_UDRE_Stage_BitFlag_
    Set gDW_UDRE_Stage_BitFlag = gDW_UDRE_Stage_BitFlag_
End Property

Public Property Let gDW_UDRE_Stage_BitFlag(RHS As DSPWave)
    gDW_UDRE_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_Stage_BitFlag", gDW_UDRE_Stage_BitFlag_
End Property

Public Property Get gDW_UDRE_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_Stage_Early_BitFlag", gDW_UDRE_Stage_Early_BitFlag_
    Set gDW_UDRE_Stage_Early_BitFlag = gDW_UDRE_Stage_Early_BitFlag_
End Property

Public Property Let gDW_UDRE_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_UDRE_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_Stage_Early_BitFlag", gDW_UDRE_Stage_Early_BitFlag_
End Property

Public Property Get gDW_UDRE_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_allDefaultBitWave", gDW_UDRE_allDefaultBitWave_
    Set gDW_UDRE_allDefaultBitWave = gDW_UDRE_allDefaultBitWave_
End Property

Public Property Let gDW_UDRE_allDefaultBitWave(RHS As DSPWave)
    gDW_UDRE_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_allDefaultBitWave", gDW_UDRE_allDefaultBitWave_
End Property

Public Property Get gDW_UDRE_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_Read_Decimal_Cate", gDW_UDRE_Read_Decimal_Cate_
    Set gDW_UDRE_Read_Decimal_Cate = gDW_UDRE_Read_Decimal_Cate_
End Property

Public Property Let gDW_UDRE_Read_Decimal_Cate(RHS As DSPWave)
    gDW_UDRE_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_Read_Decimal_Cate", gDW_UDRE_Read_Decimal_Cate_
End Property

Public Property Get gDW_UDRE_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_Read_cmpsgWavePerCyc", gDW_UDRE_Read_cmpsgWavePerCyc_
    Set gDW_UDRE_Read_cmpsgWavePerCyc = gDW_UDRE_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_UDRE_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_UDRE_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_Read_cmpsgWavePerCyc", gDW_UDRE_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_UDRE_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_Pgm_SingleBitWave", gDW_UDRE_Pgm_SingleBitWave_
    Set gDW_UDRE_Pgm_SingleBitWave = gDW_UDRE_Pgm_SingleBitWave_
End Property

Public Property Let gDW_UDRE_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_UDRE_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_Pgm_SingleBitWave", gDW_UDRE_Pgm_SingleBitWave_
End Property

Public Property Get gDW_UDRE_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_Pgm_DoubleBitWave", gDW_UDRE_Pgm_DoubleBitWave_
    Set gDW_UDRE_Pgm_DoubleBitWave = gDW_UDRE_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_UDRE_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_UDRE_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_Pgm_DoubleBitWave", gDW_UDRE_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_UDRE_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_Read_SingleBitWave", gDW_UDRE_Read_SingleBitWave_
    Set gDW_UDRE_Read_SingleBitWave = gDW_UDRE_Read_SingleBitWave_
End Property

Public Property Let gDW_UDRE_Read_SingleBitWave(RHS As DSPWave)
    gDW_UDRE_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_Read_SingleBitWave", gDW_UDRE_Read_SingleBitWave_
End Property

Public Property Get gDW_UDRE_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_Read_DoubleBitWave", gDW_UDRE_Read_DoubleBitWave_
    Set gDW_UDRE_Read_DoubleBitWave = gDW_UDRE_Read_DoubleBitWave_
End Property

Public Property Let gDW_UDRE_Read_DoubleBitWave(RHS As DSPWave)
    gDW_UDRE_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_Read_DoubleBitWave", gDW_UDRE_Read_DoubleBitWave_
End Property

Public Property Get gDW_UDRE_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRE_StageLEQJob_BitFlag", gDW_UDRE_StageLEQJob_BitFlag_
    Set gDW_UDRE_StageLEQJob_BitFlag = gDW_UDRE_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_UDRE_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_UDRE_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRE_StageLEQJob_BitFlag", gDW_UDRE_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_UDRP_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_MSBBit_Cate", gDW_UDRP_MSBBit_Cate_
    Set gDW_UDRP_MSBBit_Cate = gDW_UDRP_MSBBit_Cate_
End Property

Public Property Let gDW_UDRP_MSBBit_Cate(RHS As DSPWave)
    gDW_UDRP_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_MSBBit_Cate", gDW_UDRP_MSBBit_Cate_
End Property

Public Property Get gDW_UDRP_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_LSBBit_Cate", gDW_UDRP_LSBBit_Cate_
    Set gDW_UDRP_LSBBit_Cate = gDW_UDRP_LSBBit_Cate_
End Property

Public Property Let gDW_UDRP_LSBBit_Cate(RHS As DSPWave)
    gDW_UDRP_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_LSBBit_Cate", gDW_UDRP_LSBBit_Cate_
End Property

Public Property Get gDW_UDRP_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_BitWidth_Cate", gDW_UDRP_BitWidth_Cate_
    Set gDW_UDRP_BitWidth_Cate = gDW_UDRP_BitWidth_Cate_
End Property

Public Property Let gDW_UDRP_BitWidth_Cate(RHS As DSPWave)
    gDW_UDRP_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_BitWidth_Cate", gDW_UDRP_BitWidth_Cate_
End Property

Public Property Get gDW_UDRP_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_DefaultReal_Cate", gDW_UDRP_DefaultReal_Cate_
    Set gDW_UDRP_DefaultReal_Cate = gDW_UDRP_DefaultReal_Cate_
End Property

Public Property Let gDW_UDRP_DefaultReal_Cate(RHS As DSPWave)
    gDW_UDRP_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_DefaultReal_Cate", gDW_UDRP_DefaultReal_Cate_
End Property

Public Property Get gDW_UDRP_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_Stage_BitFlag", gDW_UDRP_Stage_BitFlag_
    Set gDW_UDRP_Stage_BitFlag = gDW_UDRP_Stage_BitFlag_
End Property

Public Property Let gDW_UDRP_Stage_BitFlag(RHS As DSPWave)
    gDW_UDRP_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_Stage_BitFlag", gDW_UDRP_Stage_BitFlag_
End Property

Public Property Get gDW_UDRP_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_Stage_Early_BitFlag", gDW_UDRP_Stage_Early_BitFlag_
    Set gDW_UDRP_Stage_Early_BitFlag = gDW_UDRP_Stage_Early_BitFlag_
End Property

Public Property Let gDW_UDRP_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_UDRP_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_Stage_Early_BitFlag", gDW_UDRP_Stage_Early_BitFlag_
End Property

Public Property Get gDW_UDRP_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_allDefaultBitWave", gDW_UDRP_allDefaultBitWave_
    Set gDW_UDRP_allDefaultBitWave = gDW_UDRP_allDefaultBitWave_
End Property

Public Property Let gDW_UDRP_allDefaultBitWave(RHS As DSPWave)
    gDW_UDRP_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_allDefaultBitWave", gDW_UDRP_allDefaultBitWave_
End Property

Public Property Get gDW_UDRP_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_Read_Decimal_Cate", gDW_UDRP_Read_Decimal_Cate_
    Set gDW_UDRP_Read_Decimal_Cate = gDW_UDRP_Read_Decimal_Cate_
End Property

Public Property Let gDW_UDRP_Read_Decimal_Cate(RHS As DSPWave)
    gDW_UDRP_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_Read_Decimal_Cate", gDW_UDRP_Read_Decimal_Cate_
End Property

Public Property Get gDW_UDRP_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_Read_cmpsgWavePerCyc", gDW_UDRP_Read_cmpsgWavePerCyc_
    Set gDW_UDRP_Read_cmpsgWavePerCyc = gDW_UDRP_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_UDRP_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_UDRP_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_Read_cmpsgWavePerCyc", gDW_UDRP_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_UDRP_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_Pgm_SingleBitWave", gDW_UDRP_Pgm_SingleBitWave_
    Set gDW_UDRP_Pgm_SingleBitWave = gDW_UDRP_Pgm_SingleBitWave_
End Property

Public Property Let gDW_UDRP_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_UDRP_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_Pgm_SingleBitWave", gDW_UDRP_Pgm_SingleBitWave_
End Property

Public Property Get gDW_UDRP_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_Pgm_DoubleBitWave", gDW_UDRP_Pgm_DoubleBitWave_
    Set gDW_UDRP_Pgm_DoubleBitWave = gDW_UDRP_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_UDRP_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_UDRP_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_Pgm_DoubleBitWave", gDW_UDRP_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_UDRP_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_Read_SingleBitWave", gDW_UDRP_Read_SingleBitWave_
    Set gDW_UDRP_Read_SingleBitWave = gDW_UDRP_Read_SingleBitWave_
End Property

Public Property Let gDW_UDRP_Read_SingleBitWave(RHS As DSPWave)
    gDW_UDRP_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_Read_SingleBitWave", gDW_UDRP_Read_SingleBitWave_
End Property

Public Property Get gDW_UDRP_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_Read_DoubleBitWave", gDW_UDRP_Read_DoubleBitWave_
    Set gDW_UDRP_Read_DoubleBitWave = gDW_UDRP_Read_DoubleBitWave_
End Property

Public Property Let gDW_UDRP_Read_DoubleBitWave(RHS As DSPWave)
    gDW_UDRP_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_Read_DoubleBitWave", gDW_UDRP_Read_DoubleBitWave_
End Property

Public Property Get gDW_UDRP_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_UDRP_StageLEQJob_BitFlag", gDW_UDRP_StageLEQJob_BitFlag_
    Set gDW_UDRP_StageLEQJob_BitFlag = gDW_UDRP_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_UDRP_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_UDRP_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_UDRP_StageLEQJob_BitFlag", gDW_UDRP_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_CMPE_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_MSBBit_Cate", gDW_CMPE_MSBBit_Cate_
    Set gDW_CMPE_MSBBit_Cate = gDW_CMPE_MSBBit_Cate_
End Property

Public Property Let gDW_CMPE_MSBBit_Cate(RHS As DSPWave)
    gDW_CMPE_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_MSBBit_Cate", gDW_CMPE_MSBBit_Cate_
End Property

Public Property Get gDW_CMPE_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_LSBBit_Cate", gDW_CMPE_LSBBit_Cate_
    Set gDW_CMPE_LSBBit_Cate = gDW_CMPE_LSBBit_Cate_
End Property

Public Property Let gDW_CMPE_LSBBit_Cate(RHS As DSPWave)
    gDW_CMPE_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_LSBBit_Cate", gDW_CMPE_LSBBit_Cate_
End Property

Public Property Get gDW_CMPE_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_BitWidth_Cate", gDW_CMPE_BitWidth_Cate_
    Set gDW_CMPE_BitWidth_Cate = gDW_CMPE_BitWidth_Cate_
End Property

Public Property Let gDW_CMPE_BitWidth_Cate(RHS As DSPWave)
    gDW_CMPE_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_BitWidth_Cate", gDW_CMPE_BitWidth_Cate_
End Property

Public Property Get gDW_CMPE_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_DefaultReal_Cate", gDW_CMPE_DefaultReal_Cate_
    Set gDW_CMPE_DefaultReal_Cate = gDW_CMPE_DefaultReal_Cate_
End Property

Public Property Let gDW_CMPE_DefaultReal_Cate(RHS As DSPWave)
    gDW_CMPE_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_DefaultReal_Cate", gDW_CMPE_DefaultReal_Cate_
End Property

Public Property Get gDW_CMPE_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_Stage_BitFlag", gDW_CMPE_Stage_BitFlag_
    Set gDW_CMPE_Stage_BitFlag = gDW_CMPE_Stage_BitFlag_
End Property

Public Property Let gDW_CMPE_Stage_BitFlag(RHS As DSPWave)
    gDW_CMPE_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_Stage_BitFlag", gDW_CMPE_Stage_BitFlag_
End Property

Public Property Get gDW_CMPE_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_Stage_Early_BitFlag", gDW_CMPE_Stage_Early_BitFlag_
    Set gDW_CMPE_Stage_Early_BitFlag = gDW_CMPE_Stage_Early_BitFlag_
End Property

Public Property Let gDW_CMPE_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_CMPE_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_Stage_Early_BitFlag", gDW_CMPE_Stage_Early_BitFlag_
End Property

Public Property Get gDW_CMPE_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_allDefaultBitWave", gDW_CMPE_allDefaultBitWave_
    Set gDW_CMPE_allDefaultBitWave = gDW_CMPE_allDefaultBitWave_
End Property

Public Property Let gDW_CMPE_allDefaultBitWave(RHS As DSPWave)
    gDW_CMPE_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_allDefaultBitWave", gDW_CMPE_allDefaultBitWave_
End Property

Public Property Get gDW_CMPE_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_Read_Decimal_Cate", gDW_CMPE_Read_Decimal_Cate_
    Set gDW_CMPE_Read_Decimal_Cate = gDW_CMPE_Read_Decimal_Cate_
End Property

Public Property Let gDW_CMPE_Read_Decimal_Cate(RHS As DSPWave)
    gDW_CMPE_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_Read_Decimal_Cate", gDW_CMPE_Read_Decimal_Cate_
End Property

Public Property Get gDW_CMPE_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_Read_cmpsgWavePerCyc", gDW_CMPE_Read_cmpsgWavePerCyc_
    Set gDW_CMPE_Read_cmpsgWavePerCyc = gDW_CMPE_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_CMPE_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_CMPE_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_Read_cmpsgWavePerCyc", gDW_CMPE_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_CMPE_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_Pgm_SingleBitWave", gDW_CMPE_Pgm_SingleBitWave_
    Set gDW_CMPE_Pgm_SingleBitWave = gDW_CMPE_Pgm_SingleBitWave_
End Property

Public Property Let gDW_CMPE_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_CMPE_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_Pgm_SingleBitWave", gDW_CMPE_Pgm_SingleBitWave_
End Property

Public Property Get gDW_CMPE_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_Pgm_DoubleBitWave", gDW_CMPE_Pgm_DoubleBitWave_
    Set gDW_CMPE_Pgm_DoubleBitWave = gDW_CMPE_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_CMPE_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_CMPE_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_Pgm_DoubleBitWave", gDW_CMPE_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_CMPE_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_Read_SingleBitWave", gDW_CMPE_Read_SingleBitWave_
    Set gDW_CMPE_Read_SingleBitWave = gDW_CMPE_Read_SingleBitWave_
End Property

Public Property Let gDW_CMPE_Read_SingleBitWave(RHS As DSPWave)
    gDW_CMPE_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_Read_SingleBitWave", gDW_CMPE_Read_SingleBitWave_
End Property

Public Property Get gDW_CMPE_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_Read_DoubleBitWave", gDW_CMPE_Read_DoubleBitWave_
    Set gDW_CMPE_Read_DoubleBitWave = gDW_CMPE_Read_DoubleBitWave_
End Property

Public Property Let gDW_CMPE_Read_DoubleBitWave(RHS As DSPWave)
    gDW_CMPE_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_Read_DoubleBitWave", gDW_CMPE_Read_DoubleBitWave_
End Property

Public Property Get gDW_CMPE_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPE_StageLEQJob_BitFlag", gDW_CMPE_StageLEQJob_BitFlag_
    Set gDW_CMPE_StageLEQJob_BitFlag = gDW_CMPE_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_CMPE_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_CMPE_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPE_StageLEQJob_BitFlag", gDW_CMPE_StageLEQJob_BitFlag_
End Property

Public Property Get gDW_CMPP_MSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_MSBBit_Cate", gDW_CMPP_MSBBit_Cate_
    Set gDW_CMPP_MSBBit_Cate = gDW_CMPP_MSBBit_Cate_
End Property

Public Property Let gDW_CMPP_MSBBit_Cate(RHS As DSPWave)
    gDW_CMPP_MSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_MSBBit_Cate", gDW_CMPP_MSBBit_Cate_
End Property

Public Property Get gDW_CMPP_LSBBit_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_LSBBit_Cate", gDW_CMPP_LSBBit_Cate_
    Set gDW_CMPP_LSBBit_Cate = gDW_CMPP_LSBBit_Cate_
End Property

Public Property Let gDW_CMPP_LSBBit_Cate(RHS As DSPWave)
    gDW_CMPP_LSBBit_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_LSBBit_Cate", gDW_CMPP_LSBBit_Cate_
End Property

Public Property Get gDW_CMPP_BitWidth_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_BitWidth_Cate", gDW_CMPP_BitWidth_Cate_
    Set gDW_CMPP_BitWidth_Cate = gDW_CMPP_BitWidth_Cate_
End Property

Public Property Let gDW_CMPP_BitWidth_Cate(RHS As DSPWave)
    gDW_CMPP_BitWidth_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_BitWidth_Cate", gDW_CMPP_BitWidth_Cate_
End Property

Public Property Get gDW_CMPP_DefaultReal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_DefaultReal_Cate", gDW_CMPP_DefaultReal_Cate_
    Set gDW_CMPP_DefaultReal_Cate = gDW_CMPP_DefaultReal_Cate_
End Property

Public Property Let gDW_CMPP_DefaultReal_Cate(RHS As DSPWave)
    gDW_CMPP_DefaultReal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_DefaultReal_Cate", gDW_CMPP_DefaultReal_Cate_
End Property

Public Property Get gDW_CMPP_Stage_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_Stage_BitFlag", gDW_CMPP_Stage_BitFlag_
    Set gDW_CMPP_Stage_BitFlag = gDW_CMPP_Stage_BitFlag_
End Property

Public Property Let gDW_CMPP_Stage_BitFlag(RHS As DSPWave)
    gDW_CMPP_Stage_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_Stage_BitFlag", gDW_CMPP_Stage_BitFlag_
End Property

Public Property Get gDW_CMPP_Stage_Early_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_Stage_Early_BitFlag", gDW_CMPP_Stage_Early_BitFlag_
    Set gDW_CMPP_Stage_Early_BitFlag = gDW_CMPP_Stage_Early_BitFlag_
End Property

Public Property Let gDW_CMPP_Stage_Early_BitFlag(RHS As DSPWave)
    gDW_CMPP_Stage_Early_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_Stage_Early_BitFlag", gDW_CMPP_Stage_Early_BitFlag_
End Property

Public Property Get gDW_CMPP_allDefaultBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_allDefaultBitWave", gDW_CMPP_allDefaultBitWave_
    Set gDW_CMPP_allDefaultBitWave = gDW_CMPP_allDefaultBitWave_
End Property

Public Property Let gDW_CMPP_allDefaultBitWave(RHS As DSPWave)
    gDW_CMPP_allDefaultBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_allDefaultBitWave", gDW_CMPP_allDefaultBitWave_
End Property

Public Property Get gDW_CMPP_Read_Decimal_Cate() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_Read_Decimal_Cate", gDW_CMPP_Read_Decimal_Cate_
    Set gDW_CMPP_Read_Decimal_Cate = gDW_CMPP_Read_Decimal_Cate_
End Property

Public Property Let gDW_CMPP_Read_Decimal_Cate(RHS As DSPWave)
    gDW_CMPP_Read_Decimal_Cate_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_Read_Decimal_Cate", gDW_CMPP_Read_Decimal_Cate_
End Property

Public Property Get gDW_CMPP_Read_cmpsgWavePerCyc() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_Read_cmpsgWavePerCyc", gDW_CMPP_Read_cmpsgWavePerCyc_
    Set gDW_CMPP_Read_cmpsgWavePerCyc = gDW_CMPP_Read_cmpsgWavePerCyc_
End Property

Public Property Let gDW_CMPP_Read_cmpsgWavePerCyc(RHS As DSPWave)
    gDW_CMPP_Read_cmpsgWavePerCyc_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_Read_cmpsgWavePerCyc", gDW_CMPP_Read_cmpsgWavePerCyc_
End Property

Public Property Get gDW_CMPP_Pgm_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_Pgm_SingleBitWave", gDW_CMPP_Pgm_SingleBitWave_
    Set gDW_CMPP_Pgm_SingleBitWave = gDW_CMPP_Pgm_SingleBitWave_
End Property

Public Property Let gDW_CMPP_Pgm_SingleBitWave(RHS As DSPWave)
    gDW_CMPP_Pgm_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_Pgm_SingleBitWave", gDW_CMPP_Pgm_SingleBitWave_
End Property

Public Property Get gDW_CMPP_Pgm_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_Pgm_DoubleBitWave", gDW_CMPP_Pgm_DoubleBitWave_
    Set gDW_CMPP_Pgm_DoubleBitWave = gDW_CMPP_Pgm_DoubleBitWave_
End Property

Public Property Let gDW_CMPP_Pgm_DoubleBitWave(RHS As DSPWave)
    gDW_CMPP_Pgm_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_Pgm_DoubleBitWave", gDW_CMPP_Pgm_DoubleBitWave_
End Property

Public Property Get gDW_CMPP_Read_SingleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_Read_SingleBitWave", gDW_CMPP_Read_SingleBitWave_
    Set gDW_CMPP_Read_SingleBitWave = gDW_CMPP_Read_SingleBitWave_
End Property

Public Property Let gDW_CMPP_Read_SingleBitWave(RHS As DSPWave)
    gDW_CMPP_Read_SingleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_Read_SingleBitWave", gDW_CMPP_Read_SingleBitWave_
End Property

Public Property Get gDW_CMPP_Read_DoubleBitWave() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_Read_DoubleBitWave", gDW_CMPP_Read_DoubleBitWave_
    Set gDW_CMPP_Read_DoubleBitWave = gDW_CMPP_Read_DoubleBitWave_
End Property

Public Property Let gDW_CMPP_Read_DoubleBitWave(RHS As DSPWave)
    gDW_CMPP_Read_DoubleBitWave_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_Read_DoubleBitWave", gDW_CMPP_Read_DoubleBitWave_
End Property

Public Property Get gDW_CMPP_StageLEQJob_BitFlag() As DSPWave
    TheHdw.DSP.SyncRead "gDW_CMPP_StageLEQJob_BitFlag", gDW_CMPP_StageLEQJob_BitFlag_
    Set gDW_CMPP_StageLEQJob_BitFlag = gDW_CMPP_StageLEQJob_BitFlag_
End Property

Public Property Let gDW_CMPP_StageLEQJob_BitFlag(RHS As DSPWave)
    gDW_CMPP_StageLEQJob_BitFlag_ = RHS
    TheHdw.DSP.SyncWrite "gDW_CMPP_StageLEQJob_BitFlag", gDW_CMPP_StageLEQJob_BitFlag_
End Property

