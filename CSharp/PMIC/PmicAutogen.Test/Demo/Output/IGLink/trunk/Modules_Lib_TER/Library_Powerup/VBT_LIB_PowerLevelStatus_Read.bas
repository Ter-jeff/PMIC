Attribute VB_Name = "VBT_LIB_PowerLevelStatus_Read"
Option Explicit
Public g_sDatalog_VDDSupply As String
Public g_sVDD_MAIN_UVI80 As String
Public g_sVDD_MAIN_LDO_UVI80 As String
Public g_sVDD_MAIN_SNS_UVI80 As String
Public g_sVDD_BUCK0_2_7_11_UVI80 As String
Public g_sVDD_BUCK1_8_9_UVI80 As String
Public g_sVDD_BUCK3_14_UVI80 As String
Public g_sVDD_MAIN_SNS_WLED_UVI80 As String
Public g_sVDD_MAIN_DRV_UVI80 As String
Public g_sVDD_MAIN_WIDAC_UVI80 As String
Public g_sVDD_MAIN1_UVI80 As String
Public g_sVDD_RTC_ALT_UVI80 As String
Public g_sVDD_MAIN_WBOOST_UVI80 As String
Public g_sVDD_LDO19_UVI80 As String
Public g_sVDD_ANA_UVI80 As String
Public g_sVDD_DIG_UVI80 As String
Public g_sVDD_BOOST_UVI80 As String
Public g_sVDD_BOOST_LDO_UVI80 As String
Public g_sVDD_BOOST_SNS_UVI80 As String
Public g_sVDD_LDO2_UVI80 As String
Public g_sVDD_LDO5_UVI80 As String
Public g_sVDDIO_1V2_UVI80 As String
Public g_sVDDIO_BUCK3_UVI80 As String
Public g_sVDD_LDO3_14_UVI80 As String
Public g_sVDD_HI_INT1_UVI80 As String
Public g_sVDD_HI_INT2_UVI80 As String
Public g_sVDD_HI_INT3_UVI80 As String
Public g_sVDD_HI_INT4_UVI80 As String
Public g_sVDD_HI_INT5_UVI80 As String
Public g_sVDD_HI_INT6_UVI80 As String

Public Enum e_VDDPin_All
    eVDD_MAIN_UVI80 = 0
    eVDD_MAIN_LDO_UVI80 = 1
    eVDD_MAIN_SNS_UVI80 = 2
    eVDD_BUCK0_2_7_11_UVI80 = 3
    eVDD_BUCK1_8_9_UVI80 = 4
    eVDD_BUCK3_14_UVI80 = 5
    eVDD_MAIN_SNS_WLED_UVI80 = 6
    eVDD_MAIN_DRV_UVI80 = 7
    eVDD_MAIN_WIDAC_UVI80 = 8
    eVDD_MAIN1_UVI80 = 9
    eVDD_RTC_ALT_UVI80 = 10
    eVDD_MAIN_WBOOST_UVI80 = 11
    eVDD_LDO19_UVI80 = 12
    eVDD_ANA_UVI80 = 13
    eVDD_DIG_UVI80 = 14
    eVDD_BOOST_UVI80 = 15
    eVDD_BOOST_LDO_UVI80 = 16
    eVDD_BOOST_SNS_UVI80 = 17
    eVDD_LDO2_UVI80 = 18
    eVDD_LDO5_UVI80 = 19
    eVDDIO_1V2_UVI80 = 20
    eVDDIO_BUCK3_UVI80 = 21
    eVDD_LDO3_14_UVI80 = 22
    eVDD_HI_INT1_UVI80 = 23
    eVDD_HI_INT2_UVI80 = 24
    eVDD_HI_INT3_UVI80 = 25
    eVDD_HI_INT4_UVI80 = 26
    eVDD_HI_INT5_UVI80 = 27
    eVDD_HI_INT6_UVI80 = 28
End Enum

Public Enum e_BootUpSeq
    eSEQ1 = 0
    eSEQ2 = 1
    eSEQ3 = 2
    eSEQ4 = 3
End Enum

Public Function VBTPOPGen_PowerPin_LevelStatus_Read()
    On Error GoTo ErrHandler
    Dim sFuncName As String:: sFuncName = "VBTPOPGen_PowerPin_LevelStatus_Read"

    Static bParsingDone As Boolean

    If bParsingDone = False Or TheExec.Datalog.Setup.LotSetup.TestMode = Engineeringmode Then
        Call VDD_Parsing_Levels_Information
        bParsingDone = True
    End If

    g_sDatalog_VDDSupply = "VDD" & GetVoltageCorner

    Call VDD_Checking_Levels_Information

    g_sVDD_MAIN_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_MAIN_UVI80").Voltage, "0.00") & "V"
    g_sVDD_MAIN_LDO_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_MAIN_LDO_UVI80").Voltage, "0.00") & "V"
    g_sVDD_MAIN_SNS_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_MAIN_SNS_UVI80").Voltage, "0.00") & "V"
    g_sVDD_BUCK0_2_7_11_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_BUCK0_2_7_11_UVI80").Voltage, "0.00") & "V"
    g_sVDD_BUCK1_8_9_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_BUCK1_8_9_UVI80").Voltage, "0.00") & "V"
    g_sVDD_BUCK3_14_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_BUCK3_14_UVI80").Voltage, "0.00") & "V"
    g_sVDD_MAIN_SNS_WLED_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_MAIN_SNS_WLED_UVI80").Voltage, "0.00") & "V"
    g_sVDD_MAIN_DRV_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_MAIN_DRV_UVI80").Voltage, "0.00") & "V"
    g_sVDD_MAIN_WIDAC_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_MAIN_WIDAC_UVI80").Voltage, "0.00") & "V"
    g_sVDD_MAIN1_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_MAIN1_UVI80").Voltage, "0.00") & "V"
    g_sVDD_RTC_ALT_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_RTC_ALT_UVI80").Voltage, "0.00") & "V"
    g_sVDD_MAIN_WBOOST_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_MAIN_WBOOST_UVI80").Voltage, "0.00") & "V"
    g_sVDD_LDO19_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_LDO19_UVI80").Voltage, "0.00") & "V"
    g_sVDD_ANA_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_ANA_UVI80").Voltage, "0.00") & "V"
    g_sVDD_DIG_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_DIG_UVI80").Voltage, "0.00") & "V"
    g_sVDD_BOOST_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_BOOST_UVI80").Voltage, "0.00") & "V"
    g_sVDD_BOOST_LDO_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_BOOST_LDO_UVI80").Voltage, "0.00") & "V"
    g_sVDD_BOOST_SNS_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_BOOST_SNS_UVI80").Voltage, "0.00") & "V"
    g_sVDD_LDO2_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_LDO2_UVI80").Voltage, "0.00") & "V"
    g_sVDD_LDO5_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_LDO5_UVI80").Voltage, "0.00") & "V"
    g_sVDDIO_1V2_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDDIO_1V2_UVI80").Voltage, "0.00") & "V"
    g_sVDDIO_BUCK3_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDDIO_BUCK3_UVI80").Voltage, "0.00") & "V"
    g_sVDD_LDO3_14_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_LDO3_14_UVI80").Voltage, "0.00") & "V"
    g_sVDD_HI_INT1_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_HI_INT1_UVI80").Voltage, "0.00") & "V"
    g_sVDD_HI_INT2_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_HI_INT2_UVI80").Voltage, "0.00") & "V"
    g_sVDD_HI_INT3_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_HI_INT3_UVI80").Voltage, "0.00") & "V"
    g_sVDD_HI_INT4_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_HI_INT4_UVI80").Voltage, "0.00") & "V"
    g_sVDD_HI_INT5_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_HI_INT5_UVI80").Voltage, "0.00") & "V"
    g_sVDD_HI_INT6_UVI80 = "VDD" & Format(TheHdw.DCVI.Pins("VDD_HI_INT6_UVI80").Voltage, "0.00") & "V"

    Exit Function
ErrHandler:
    TheExec.AddOutput "<Error>" + sFuncName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + sFuncName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
