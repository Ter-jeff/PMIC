Attribute VB_Name = "SP_IDS_Pins_Cond"
'======================================================================================================================================
Public Function IDS_OFF_TEST(AHB_WRITE_OPTION As Boolean, FLAT_PATTERN_OPTION As Boolean) As Long

    Dim Realy_On as string
    Dim Relay_Off as string
    Dim WaitTime as double
    Dim VDD_DIG_UVI80_val As New SiteDouble
    Dim VDDC_UVI80_val As New SiteDouble
    Dim VDDH_UVI80_val As New SiteDouble
    Dim VDDIO_UVI80_val As New SiteDouble
    Dim VDDKKK_UVI80_val As New SiteDouble
    Dim VDDENG_UVI80_val As New SiteDouble
    Dim TestName1 As String
    Dim TestName2 As String
    Dim TestName3 As String
    Dim TestName4 As String
    Dim TestName5 As String
    Dim TestName6 As String
    Dim meas1 As New PinListData
    Dim meas2 As New PinListData
    Dim meas3 As New PinListData
    Dim meas4 As New PinListData

    Realy_On = "K4474,K5575"
    Relay_Off = "K3333,K9999"
    WaitTime = 0.03


'===================== Relay setup =====================
    TheHdw.Utility.Pins(Relay_Off).State = tlUtilBitOff
    TheHdw.Utility.Pins(Realy_On).State = tlUtilBitOn


'===================== Special Pin status setup according to project =====================
'user need to put special setting in here 
'Like some pin need force H/L or some pin need disconnect
    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered


'===================== Instrument setup =====================
    With TheHdw.DCVI.Pins("VDD_DIG_UVI80")
        .Mode = tlDCVIModeVoltage
        .SetCurrentAndRange 0.0002, Abs(0.0002)
        .Meter.Mode = tlDCVIMeterCurrent
        .Connect
        TheHdw.Wait 1 * ms
        .Gate = True
    End With
    With TheHdw.DCVI.Pins("VDDC_UVI80,VDDIO_UVI80")
        .Mode = tlDCVIModeVoltage
        .SetCurrentAndRange 0.002, Abs(0.002)
        .Meter.Mode = tlDCVIMeterCurrent
        .Connect
        TheHdw.Wait 1 * ms
        .Gate = True
    End With
    With TheHdw.DCVI.Pins("VDDH_UVI80")
        .Mode = tlDCVIModeVoltage
        .SetCurrentAndRange 0.02, Abs(0.02)
        .Meter.Mode = tlDCVIMeterCurrent
        .Connect
        TheHdw.Wait 1 * ms
        .Gate = True
    End With
    With TheHdw.DCVI.Pins("VDDKKK_UVI80,VDDENG_UVI80")
        .Mode = tlDCVIModeVoltage
        .SetCurrentAndRange 0.2, Abs(0.2)
        .Meter.Mode = tlDCVIMeterCurrent
        .Connect
        TheHdw.Wait 1 * ms
        .Gate = True
    End With
    TheHdw.Wait WaitTime

'===================== Measure =====================
    If TheExec.TesterMode = testModeOffline Then
        VDD_DIG_UVI80_val = 1
        VDDC_UVI80_val = 1
        VDDH_UVI80_val = 1
        VDDIO_UVI80_val = 1
        VDDKKK_UVI80_val = 1
        VDDENG_UVI80_val = 1
    Else
        meas1 = TheHdw.DCVI.Pins("VDD_DIG_UVI80").Meter.Read(tlStrobe, 100, 100 * KHz, tlDCVIMeterReadingFormatAverage)
        VDD_DIG_UVI80_val = meas1.Pins("VDD_DIG_UVI80")
        meas2 = TheHdw.DCVI.Pins("VDDC_UVI80,VDDIO_UVI80").Meter.Read(tlStrobe, 100, 100 * KHz, tlDCVIMeterReadingFormatAverage)
        VDDC_UVI80_val = meas2.Pins("VDDC_UVI80")
        VDDIO_UVI80_val = meas2.Pins("VDDIO_UVI80")
        meas3 = TheHdw.DCVI.Pins("VDDH_UVI80").Meter.Read(tlStrobe, 100, 100 * KHz, tlDCVIMeterReadingFormatAverage)
        VDDH_UVI80_val = meas3.Pins("VDDH_UVI80")
        meas4 = TheHdw.DCVI.Pins("VDDKKK_UVI80,VDDENG_UVI80").Meter.Read(tlStrobe, 100, 100 * KHz, tlDCVIMeterReadingFormatAverage)
        VDDKKK_UVI80_val = meas4.Pins("VDDKKK_UVI80")
        VDDENG_UVI80_val = meas4.Pins("VDDENG_UVI80")

    End If


'===================== Datalog =====================
'===================== TestName num depend on power pin =====================
    TestName1 = "IDS_OFF_VDD-DIG-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName2 = "IDS_OFF_VDDC-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName3 = "IDS_OFF_VDDH-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName4 = "IDS_OFF_VDDIO-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName5 = "IDS_OFF_VDDKKK-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName6 = "IDS_OFF_VDDENG-UVI80_X_X_X_P_X_X_MeasI_X_X"


'===================== Hi/Low limit according to Fuji PE =====================
    Call TheExec.Flow.TestLimit(ResultVal:=VDD_DIG_UVI80_val, TName:=TestName1, hiVal:=0.0002, lowVal:=5E-06, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDC_UVI80_val, TName:=TestName2, hiVal:=0.0002, lowVal:=0.00005, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDH_UVI80_val, TName:=TestName3, hiVal:=0.007, lowVal:=0.006, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDIO_UVI80_val, TName:=TestName4, hiVal:=0.0002, lowVal:=0.00005, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDKKK_UVI80_val, TName:=TestName5, hiVal:=0.05, lowVal:=0.01, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDENG_UVI80_val, TName:=TestName6, hiVal:=0.05, lowVal:=0.01, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)


'===================== Special Pin status setup according to project =====================
    '' User need to put special pin reset setting in here


'===================== Relay reset =====================
    TheHdw.Utility.Pins(Relay_On).State = tlUtilBitOff
    TheHdw.Utility.Pins(Relay_Off).State = tlUtilBitOn

    Exit Function

ErrHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function IDS_ACTIVE_TEST(AHB_WRITE_OPTION As Boolean, FLAT_PATTERN_OPTION As Boolean) As Long

    Dim Realy_On as string
    Dim Relay_Off as string
    Dim WaitTime as double
    Dim VDD_DIG_UVI80_val As New SiteDouble
    Dim VDDC_UVI80_val As New SiteDouble
    Dim VDDH_UVI80_val As New SiteDouble
    Dim VDDIO_UVI80_val As New SiteDouble
    Dim VDDKKK_UVI80_val As New SiteDouble
    Dim VDDENG_UVI80_val As New SiteDouble
    Dim VDD_DIG_UVI80_Original_CurrentRange As Double
    Dim VDDC_UVI80_Original_CurrentRange As Double
    Dim VDDH_UVI80_Original_CurrentRange As Double
    Dim VDDIO_UVI80_Original_CurrentRange As Double
    Dim VDDKKK_UVI80_Original_CurrentRange As Double
    Dim VDDENG_UVI80_Original_CurrentRange As Double
    Dim VDD_DIG_UVI80_Original_Current As Double
    Dim VDDC_UVI80_Original_Current As Double
    Dim VDDH_UVI80_Original_Current As Double
    Dim VDDIO_UVI80_Original_Current As Double
    Dim VDDKKK_UVI80_Original_Current As Double
    Dim VDDENG_UVI80_Original_Current As Double
    Dim TestName1 As String
    Dim TestName2 As String
    Dim TestName3 As String
    Dim TestName4 As String
    Dim TestName5 As String
    Dim TestName6 As String
    Dim meas1 As New PinListData
    Dim meas2 As New PinListData
    Dim meas3 As New PinListData

    Realy_On = "K4474,K5575,K7878"
    Relay_Off = "K3333,K9999,K7777"
    WaitTime = 0.05


'===================== Relay setup =====================
    TheHdw.Utility.Pins(Relay_Off).State = tlUtilBitOff
    TheHdw.Utility.Pins(Realy_On).State = tlUtilBitOn


'===================== save original power pin current range =====================

    VDD_DIG_UVI80_Original_CurrentRange = TheHdw.DCVI.Pins("VDD_DIG_UVI80").CurrentRange
    VDDC_UVI80_Original_CurrentRange = TheHdw.DCVI.Pins("VDDC_UVI80").CurrentRange
    VDDH_UVI80_Original_CurrentRange = TheHdw.DCVI.Pins("VDDH_UVI80").CurrentRange
    VDDIO_UVI80_Original_CurrentRange = TheHdw.DCVI.Pins("VDDIO_UVI80").CurrentRange
    VDDKKK_UVI80_Original_CurrentRange = TheHdw.DCVI.Pins("VDDKKK_UVI80").CurrentRange
    VDDENG_UVI80_Original_CurrentRange = TheHdw.DCVI.Pins("VDDENG_UVI80").CurrentRange

    VDD_DIG_UVI80_Original_Current = TheHdw.DCVI.Pins("VDD_DIG_UVI80").Current
    VDDC_UVI80_Original_Current = TheHdw.DCVI.Pins("VDDC_UVI80").Current
    VDDH_UVI80_Original_Current = TheHdw.DCVI.Pins("VDDH_UVI80").Current
    VDDIO_UVI80_Original_Current = TheHdw.DCVI.Pins("VDDIO_UVI80").Current
    VDDKKK_UVI80_Original_Current = TheHdw.DCVI.Pins("VDDKKK_UVI80").Current
    VDDENG_UVI80_Original_Current = TheHdw.DCVI.Pins("VDDENG_UVI80").Current


'===================== Special pin setup according project=====================
'user need to put special setting in here 
'Like some pin need force H/L or some pin need disconnect
    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered


'=====================Run different Test Mode setup pattern=====================
'    TheExec.Datalog.WriteComment "Running Default Pattern:User define by diff. project"
'    TheHdw.Patterns("xxx").Load
'    TheHdw.Patterns("xxx").Start
'    TheHdw.Digital.Patgen.HaltWait


'===================== Instrument setup =====================
    With TheHdw.DCVI.Pins("VDD_DIG_UVI80,VDDC_UVI80,VDDH_UVI80")
        .Mode = tlDCVIModeVoltage
        .SetCurrentAndRange 0.2, Abs(0.2)
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 20 * mA
        .Connect
        TheHdw.Wait 1 * ms
        .Gate = True
    End With
    With TheHdw.DCVI.Pins("VDDIO_UVI80,VDDKKK_UVI80")
        .Mode = tlDCVIModeVoltage
        .SetCurrentAndRange 0.02, Abs(0.02)
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 20 * mA
        .Connect
        TheHdw.Wait 1 * ms
        .Gate = True
    End With
    With TheHdw.DCVI.Pins("VDDENG_UVI80")
        .Mode = tlDCVIModeVoltage
        .SetCurrentAndRange 0.002, Abs(0.002)
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 20 * mA
        .Connect
        TheHdw.Wait 1 * ms
        .Gate = True
    End With
TheHdw.Wait WaitTime

'===================== Measure =====================
    If TheExec.TesterMode = testModeOffline Then
        VDD_DIG_UVI80_val = 10
        VDDC_UVI80_val = 10
        VDDH_UVI80_val = 10
        VDDIO_UVI80_val = 10
        VDDKKK_UVI80_val = 10
        VDDENG_UVI80_val = 10
    Else
        meas1 = TheHdw.DCVI.Pins("VDD_DIG_UVI80,VDDC_UVI80,VDDH_UVI80").Meter.Read(tlStrobe, 100, 100 * KHz, tlDCVIMeterReadingFormatAverage)
        VDD_DIG_UVI80_val = meas1.Pins("VDD_DIG_UVI80")
        VDDC_UVI80_val = meas1.Pins("VDDC_UVI80")
        VDDH_UVI80_val = meas1.Pins("VDDH_UVI80")
        meas2 = TheHdw.DCVI.Pins("VDDIO_UVI80,VDDKKK_UVI80").Meter.Read(tlStrobe, 100, 100 * KHz, tlDCVIMeterReadingFormatAverage)
        VDDIO_UVI80_val = meas2.Pins("VDDIO_UVI80")
        VDDKKK_UVI80_val = meas2.Pins("VDDKKK_UVI80")
        meas3 = TheHdw.DCVI.Pins("VDDENG_UVI80").Meter.Read(tlStrobe, 100, 100 * KHz, tlDCVIMeterReadingFormatAverage)
        VDDENG_UVI80_val = meas3.Pins("VDDENG_UVI80")
    End If


'===================== Datalog =====================
'===================== TestName num depend on power pin =====================
    TestName1 = "IDS_ACTIVE_VDD-DIG-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName2 = "IDS_ACTIVE_VDDC-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName3 = "IDS_ACTIVE_VDDH-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName4 = "IDS_ACTIVE_VDDIO-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName5 = "IDS_ACTIVE_VDDKKK-UVI80_X_X_X_P_X_X_MeasI_X_X"
    TestName6 = "IDS_ACTIVE_VDDENG-UVI80_X_X_X_P_X_X_MeasI_X_X"


'===================== Hi/Low limit according to Fuji PE =====================
    Call TheExec.Flow.TestLimit(ResultVal:=VDD_DIG_UVI80_val, TName:=TestName1, hiVal:=0.02, lowVal:=0.01, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDC_UVI80_val, TName:=TestName2, hiVal:=0.02, lowVal:=0.01, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDH_UVI80_val, TName:=TestName3, hiVal:=0.02, lowVal:=0.01, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDIO_UVI80_val, TName:=TestName4, hiVal:=0.008, lowVal:=0.005, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDKKK_UVI80_val, TName:=TestName5, hiVal:=0.006, lowVal:=0.003, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDENG_UVI80_val, TName:=TestName6, hiVal:=0.0015, lowVal:=0.001, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)


'''===================== Power up status check this project Power Status 2 = Active mode =====================
'''   Dim AHBVal As New SiteLong
'''   AHB_READNWIRE POWER_CONTROL_MAINFSM_POWER_STATE_STATUS.Addr, g_RegVal
'''   AHBVal = g_RegVal
'''   For Each g_Site In TheExec.Sites
'''        If AHBVal <> 2 Then F_IDS_AWAKE = True
'''    Next g_Site
'''
'''TestName = TNameCombine("IDS", "Power", "Status", , , , TName_NonTrimItem, , NHLV, TName_None)
'''TheExec.Flow.TestLimit AHBVal, 2, 2, , , , , , TestName


'===================== power pin reset for original range =====================
    TheHdw.DCVI.Pins("VDD_DIG_UVI80").SetCurrentAndRange VDD_DIG_UVI80_Original_Current, VDD_DIG_UVI80_Original_CurrentRange
    TheHdw.DCVI.Pins("VDDC_UVI80").SetCurrentAndRange VDDC_UVI80_Original_Current, VDDC_UVI80_Original_CurrentRange
    TheHdw.DCVI.Pins("VDDH_UVI80").SetCurrentAndRange VDDH_UVI80_Original_Current, VDDH_UVI80_Original_CurrentRange
    TheHdw.DCVI.Pins("VDDIO_UVI80").SetCurrentAndRange VDDIO_UVI80_Original_Current, VDDIO_UVI80_Original_CurrentRange
    TheHdw.DCVI.Pins("VDDKKK_UVI80").SetCurrentAndRange VDDKKK_UVI80_Original_Current, VDDKKK_UVI80_Original_CurrentRange
    TheHdw.DCVI.Pins("VDDENG_UVI80").SetCurrentAndRange VDDENG_UVI80_Original_Current, VDDENG_UVI80_Original_CurrentRange


'===================== Special Pin status setup according to project =====================
'User need to put special pin reset setting in here


'===================== Relay reset =====================
    TheHdw.Utility.Pins(Relay_On).State = tlUtilBitOff
    TheHdw.Utility.Pins(Relay_Off).State = tlUtilBitOn

    Exit Function

ErrHandler:
    If AbortTest Then Exit Function Else Resume Next

End Function

