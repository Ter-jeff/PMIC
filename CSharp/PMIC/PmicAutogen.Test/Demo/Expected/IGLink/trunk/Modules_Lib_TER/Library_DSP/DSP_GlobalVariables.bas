Attribute VB_Name = "DSP_GlobalVariables"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

'___ OTP-DSP
Private gD_wPGMData_ As New DSPWave
Private gD_wReadData_ As New DSPWave
Private gDW_DefaultDSPRawData_ As New DSPWave   '20190416 Toppy
Private gD_wDEIDPGMBits_ As New DSPWave    'totoal DEID 64 bits per Site

'___ OTP Constant-DSP
Private gD_slOTP_ADDR_BW_ As New SiteLong
Private gD_slOTP_REGDATA_BW_ As New SiteLong

'___ OTP CRC DSP Procedure
Private gD_wCRCSelfLUT_ As New DSPWave   ''''CRC8 Loop Up Table
Private gD_slCRCSelfValue_ As New SiteLong

'Private gDL_AHB_BW_ As New SiteLong '20190524
' ****************************************************
' *** DO NOT ALTER OR ADD ANY CODE BELOW THIS LINE ***
' ***  THE CODE BELOW IS AUTOMATICALLY GENERATED   ***
' Place the declarations for DSP global variables above
' this comment using the following format:
' Private Name_ As New Type
'
Private tl_DSPGlobalVariableCheckVar_ As New SiteLong
Private Function tl_DSPGlobalVariableInitializer() As Long
    On Error Resume Next
    tl_DSPGlobalVariableInitializer = tl_DSPGlobalVariableCheckVar_
    tl_DSPGlobalVariableCheckVar_ = 1
    Exit Function
End Function

' Called to also reset VBA variables when embedded variables need to be reset, e.g., after re-validation.
Public Sub tlDSPGlobalVariableReset()
    Set gD_wPGMData_ = Nothing
    Set gD_wReadData_ = Nothing
    Set gDW_DefaultDSPRawData_ = Nothing
    Set gD_wDEIDPGMBits_ = Nothing
    Set gD_slOTP_ADDR_BW_ = Nothing
    Set gD_slOTP_REGDATA_BW_ = Nothing
    Set gD_wCRCSelfLUT_ = Nothing
    Set gD_slCRCSelfValue_ = Nothing
End Sub

Public Property Get gD_wPGMData() As DSPWave
    TheHdw.DSP.SyncRead "gD_wPGMData", gD_wPGMData_
    Set gD_wPGMData = gD_wPGMData_
End Property

Public Property Let gD_wPGMData(RHS As DSPWave)
    gD_wPGMData_ = RHS
    TheHdw.DSP.SyncWrite "gD_wPGMData", gD_wPGMData_
End Property

Public Property Get gD_wReadData() As DSPWave
    TheHdw.DSP.SyncRead "gD_wReadData", gD_wReadData_
    Set gD_wReadData = gD_wReadData_
End Property

Public Property Let gD_wReadData(RHS As DSPWave)
    gD_wReadData_ = RHS
    TheHdw.DSP.SyncWrite "gD_wReadData", gD_wReadData_
End Property

Public Property Get gDW_DefaultDSPRawData() As DSPWave
    TheHdw.DSP.SyncRead "gDW_DefaultDSPRawData", gDW_DefaultDSPRawData_
    Set gDW_DefaultDSPRawData = gDW_DefaultDSPRawData_
End Property

Public Property Let gDW_DefaultDSPRawData(RHS As DSPWave)
    gDW_DefaultDSPRawData_ = RHS
    TheHdw.DSP.SyncWrite "gDW_DefaultDSPRawData", gDW_DefaultDSPRawData_
End Property

Public Property Get gD_wDEIDPGMBits() As DSPWave
    TheHdw.DSP.SyncRead "gD_wDEIDPGMBits", gD_wDEIDPGMBits_
    Set gD_wDEIDPGMBits = gD_wDEIDPGMBits_
End Property

Public Property Let gD_wDEIDPGMBits(RHS As DSPWave)
    gD_wDEIDPGMBits_ = RHS
    TheHdw.DSP.SyncWrite "gD_wDEIDPGMBits", gD_wDEIDPGMBits_
End Property

Public Property Get gD_slOTP_ADDR_BW() As SiteLong
    TheHdw.DSP.SyncRead "gD_slOTP_ADDR_BW", gD_slOTP_ADDR_BW_
    Set gD_slOTP_ADDR_BW = gD_slOTP_ADDR_BW_
End Property

Public Property Let gD_slOTP_ADDR_BW(RHS As Variant)
    gD_slOTP_ADDR_BW_ = RHS
    TheHdw.DSP.SyncWrite "gD_slOTP_ADDR_BW", gD_slOTP_ADDR_BW_
End Property

Public Property Get gD_slOTP_REGDATA_BW() As SiteLong
    TheHdw.DSP.SyncRead "gD_slOTP_REGDATA_BW", gD_slOTP_REGDATA_BW_
    Set gD_slOTP_REGDATA_BW = gD_slOTP_REGDATA_BW_
End Property

Public Property Let gD_slOTP_REGDATA_BW(RHS As Variant)
    gD_slOTP_REGDATA_BW_ = RHS
    TheHdw.DSP.SyncWrite "gD_slOTP_REGDATA_BW", gD_slOTP_REGDATA_BW_
End Property

Public Property Get gD_wCRCSelfLUT() As DSPWave
    TheHdw.DSP.SyncRead "gD_wCRCSelfLUT", gD_wCRCSelfLUT_
    Set gD_wCRCSelfLUT = gD_wCRCSelfLUT_
End Property

Public Property Let gD_wCRCSelfLUT(RHS As DSPWave)
    gD_wCRCSelfLUT_ = RHS
    TheHdw.DSP.SyncWrite "gD_wCRCSelfLUT", gD_wCRCSelfLUT_
End Property

Public Property Get gD_slCRCSelfValue() As SiteLong
    TheHdw.DSP.SyncRead "gD_slCRCSelfValue", gD_slCRCSelfValue_
    Set gD_slCRCSelfValue = gD_slCRCSelfValue_
End Property

Public Property Let gD_slCRCSelfValue(RHS As Variant)
    gD_slCRCSelfValue_ = RHS
    TheHdw.DSP.SyncWrite "gD_slCRCSelfValue", gD_slCRCSelfValue_
End Property

