Attribute VB_Name = "VBT_LIB_DC_IDS_PMIC"
Option Explicit
'Revision History:
'V0.0 initial

'======================================================================================================================================
Public Function IDS_OFF_TEST(AHB_WRITE_OPTION As Boolean, FLAT_PATTERN_OPTION As Boolean) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IDS_OFF_TEST"

    Dim VDD_DIG_UVI80_val As New SiteDouble, VDD_DIG_UVI80_Cal As New SiteDouble
    Dim VDDC_UVI80_val As New SiteDouble, VDDC_UVI80_Cal As New SiteDouble
    Dim VDDH_UVI80_val As New SiteDouble, VDDH_UVI80_Cal As New SiteDouble
    Dim VDDIO_UVI80_val As New SiteDouble, VDDIO_UVI80_Cal As New SiteDouble
    Dim TestName1 As String
    Dim TestName2 As String
    Dim TestName3 As String
    Dim TestName4 As String

    Dim meas      As New PinListData


    '===================== Relay setup =====================
    TheHdw.Utility.Pins("K2101,K2701,K2801,K2901").State = tlUtilBitOff

    '===================== Special Pin status setup according to project =====================
    TheHdw.Digital.Pins("RESET_L").Disconnect
    With TheHdw.PPMU.Pins("RESET_L")
        .Gate = tlOff
        .ForceV 0, 0.002
        .ClampVHi = 6
        .ClampVLo = 0
        .Connect
        .Gate = tlOn
    End With
    TheHdw.Wait 0.003
    TheHdw.Digital.Pins("CRASH_L").Disconnect
    With TheHdw.PPMU.Pins("CRASH_L")
        .Gate = tlOff
        .ForceV 1.8, 0.002
        .ClampVHi = 6
        .ClampVLo = 0
        .Connect
        .Gate = tlOn
    End With
    TheHdw.Wait 0.003
    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered

    '===================== Instrument setup =====================
    With TheHdw.DCVI.Pins("VDD_DIG_UVI80, LKG_UVI80")
        .Mode = tlDCVIModeVoltage
        .SetCurrentAndRange 200 * mA, Abs(200 * mA)
        .Meter.Mode = tlDCVIMeterCurrent
        .Connect
        TheHdw.Wait 1 * ms

        .Gate = True
    End With
    TheHdw.Wait 30 * ms

    If TheExec.CurrentJob Like "*CHAR*" Then

        With TheHdw.DCVI.Pins("VDD_DIG_UVI80, LKG_UVI80")
            .SetCurrentAndRange 2 * mA, Abs(2 * mA)
            .Meter.CurrentRange = 2 * mA
        End With

        With TheHdw.DCVI.Pins("VDDC_UVI80, VDDIO_UVI80")
            .SetCurrentAndRange 200 * uA, Abs(200 * uA)
            .Meter.CurrentRange = 200 * uA
        End With
        TheHdw.Wait 200 * ms
    Else

        With TheHdw.DCVI.Pins("LKG_UVI80")
            .SetCurrentAndRange 2 * mA, Abs(2 * mA)
            .Meter.CurrentRange = 2 * mA
        End With

        With TheHdw.DCVI.Pins("VDD_DIG_UVI80, VDDC_UVI80, VDDIO_UVI80")
            .SetCurrentAndRange 200 * uA, Abs(200 * uA)
            .Meter.CurrentRange = 200 * uA
        End With
        TheHdw.Wait 200 * ms
    End If

    '===================== Measure =====================
    If TheExec.TesterMode = testModeOffline Then
        VDD_DIG_UVI80_val = 1
        VDDC_UVI80_val = 1
        VDDH_UVI80_val = 1
        VDDIO_UVI80_val = 1
    Else

        meas = TheHdw.DCVI.Pins("VDD_DIG_UVI80, VDDC_UVI80, VDDH_UVI80, VDDIO_UVI80").Meter.Read(tlStrobe, 100, 100 * KHz, tlDCVIMeterReadingFormatAverage)

        VDD_DIG_UVI80_val = meas.Pins("VDD_DIG_UVI80")
        VDDC_UVI80_val = meas.Pins("VDDC_UVI80")
        VDDH_UVI80_val = meas.Pins("VDDH_UVI80")
        VDDIO_UVI80_val = meas.Pins("VDDIO_UVI80")

    End If


    '===================== Datalog =====================

    '===================== TestName num depend on power pin =====================
    TestName1 = TNameCombine("IDS", "POR", "VDD-DIG-UVI80", , , , TName_NonTrimItem, , , TName_MeasI)
    TestName2 = TNameCombine("IDS", "POR", "VDDC-UVI80", , , , TName_NonTrimItem, , , TName_MeasI)
    TestName3 = TNameCombine("IDS", "POR", "VDDH-UVI80", , , , TName_NonTrimItem, , , TName_MeasI)
    TestName4 = TNameCombine("IDS", "POR", "VDDIO-UVI80", , , , TName_NonTrimItem, , , TName_MeasI)


    '===================== Hi/Low limit according to Fuji PE =====================
    Call TheExec.Flow.TestLimit(ResultVal:=VDD_DIG_UVI80_val, TName:=TestName1, hiVal:=0.0002, lowVal:=0.00005, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDC_UVI80_val, TName:=TestName2, hiVal:=0.00012, lowVal:=0.00003, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDH_UVI80_val, TName:=TestName3, hiVal:=0.0004, lowVal:=0.0001, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDIO_UVI80_val, TName:=TestName4, hiVal:=0.00007, lowVal:=0.00001, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)


    '===================== Special Pin status setup according to project =====================
    TheHdw.PPMU.Pins("CRASH_L").Disconnect
    TheHdw.Digital.Pins("CRASH_L").Connect


    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IDS_ACTIVE_TEST(AHB_WRITE_OPTION As Boolean, FLAT_PATTERN_OPTION As Boolean) As Long
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IDS_ACTIVE_TEST"

    Dim VDD_DIG_UVI80_val As New SiteDouble, VDD_DIG_UVI80_Cal As New SiteDouble
    Dim VDDC_UVI80_val As New SiteDouble, VDDC_UVI80_Cal As New SiteDouble
    Dim VDDH_UVI80_val As New SiteDouble, VDDH_UVI80_Cal As New SiteDouble
    Dim VDDIO_UVI80_val As New SiteDouble, VDDIO_UVI80_Cal As New SiteDouble
    Dim F_PB As String, F_Delta As String
    Dim F_IDS_ACTIVE_PB As SiteBoolean, F_IDS_ACTIVE_Delta As SiteBoolean
    Dim TestName1 As String
    Dim TestName2 As String
    Dim TestName3 As String
    Dim TestName4 As String
    Dim AHBVal    As New SiteLong
    Dim meas      As New PinListData

    '===================== Special pin setup according project=====================
    TheHdw.Digital.Pins("RESET_L").Disconnect
    With TheHdw.PPMU.Pins("RESET_L")
        .Gate = tlOff
        .ForceV 1.8, 0.002
        .ClampVHi = 6
        .ClampVLo = 0
        .Connect
        .Gate = tlOn
    End With
    TheHdw.Wait 0.003
    TheHdw.Digital.Pins("CRASH_L").Disconnect
    With TheHdw.PPMU.Pins("CRASH_L")
        .Gate = tlOff
        .ForceV 1.8, 0.002
        .ClampVHi = 6
        .ClampVLo = 0
        .Connect
        .Gate = tlOn
    End With
    TheHdw.Wait 0.003
    TheHdw.Digital.ApplyLevelsTiming True, True, False, tlPowered

    '=====================Run different Test Mode setup pattern=====================
    TheExec.Datalog.WriteComment "Running Default Pattern:PP_SPAA0_C_IN00_AN_XXXX_PFF_JTG_UNS_ALLFV_SI_V02P01DF_AVCTSU_1_A0_1902171900.PAT"
    TheHdw.Patterns(".\Patterns\HARD_IP\PP_SPAA0_C_IN00_AN_XXXX_PFF_JTG_UNS_ALLFV_SI_V02P01DF_AVCTSU_1_A0_1902171900.PAT").Load
    TheHdw.Patterns(".\Patterns\HARD_IP\PP_SPAA0_C_IN00_AN_XXXX_PFF_JTG_UNS_ALLFV_SI_V02P01DF_AVCTSU_1_A0_1902171900.PAT").Start
    TheHdw.Digital.Patgen.HaltWait


    '===================== Instrument setup =====================
    With TheHdw.DCVI.Pins("VDD_DIG_UVI80, VDDC_UVI80, VDDH_UVI80, VDDIO_UVI80")
        .Mode = tlDCVIModeVoltage
        .SetCurrentAndRange 20 * mA, Abs(20 * mA)
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange = 20 * mA
        .Connect
        TheHdw.Wait 1 * ms
        .Gate = True
    End With

    With TheHdw.DCVI.Pins("VDDIO_UVI80")
        .SetCurrentAndRange 200 * uA, Abs(200 * uA)
        .Meter.CurrentRange = 200 * uA
    End With

    With TheHdw.DCVI.Pins("VDDC_UVI80, VDDH_UVI80")
        .SetCurrentAndRange 2 * mA, Abs(2 * mA)
        .Meter.CurrentRange = 2 * mA
    End With
    TheHdw.Wait 30 * ms


    '===================== Measure =====================
    If TheExec.TesterMode = testModeOffline Then
        VDD_DIG_UVI80_val = 10
        VDDC_UVI80_val = 10
        VDDH_UVI80_val = 10
        VDDIO_UVI80_val = 10
    Else
        meas = TheHdw.DCVI.Pins("VDD_DIG_UVI80, VDDC_UVI80, VDDH_UVI80, VDDIO_UVI80").Meter.Read(tlStrobe, 100, 100 * KHz, tlDCVIMeterReadingFormatAverage)

        VDD_DIG_UVI80_val = meas.Pins("VDD_DIG_UVI80")    '1mA
        VDDC_UVI80_val = meas.Pins("VDDC_UVI80")    '200uA
        VDDH_UVI80_val = meas.Pins("VDDH_UVI80")    '200uA
        VDDIO_UVI80_val = meas.Pins("VDDIO_UVI80")    '15uA
    End If


    '===================== Datalog =====================

    '===================== TestName num depend on power pin =====================
    TestName1 = TNameCombine("IDS", "ACTIVE", "VDD-DIG-UVI80", , , , TName_NonTrimItem, , , TName_MeasI)
    TestName2 = TNameCombine("IDS", "ACTIVE", "VDDC-UVI80", , , , TName_NonTrimItem, , , TName_MeasI)
    TestName3 = TNameCombine("IDS", "ACTIVE", "VDDH-UVI80", , , , TName_NonTrimItem, , , TName_MeasI)
    TestName4 = TNameCombine("IDS", "ACTIVE", "VDDIO-UVI80", , , , TName_NonTrimItem, , , TName_MeasI)

    '===================== Hi/Low limit according to Fuji PE =====================
    Call TheExec.Flow.TestLimit(ResultVal:=VDD_DIG_UVI80_val, TName:=TestName1, hiVal:=0.0048, lowVal:=0.00005, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDC_UVI80_val, TName:=TestName2, hiVal:=0.0012, lowVal:=0.00003, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDH_UVI80_val, TName:=TestName3, hiVal:=0.0004, lowVal:=0.0001, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)
    Call TheExec.Flow.TestLimit(ResultVal:=VDDIO_UVI80_val, TName:=TestName4, hiVal:=0.00007, lowVal:=0.00005, Unit:=unitAmp, formatStr:="%6.4f", scaletype:=scaleMicro)



    '''===================== Power up status check this project Power Status 2 = Active mode =====================
    '''   AHB_READNWIRE POWER_CONTROL_MAINFSM_POWER_STATE_STATUS.Addr, g_RegVal
    '''   AHBVal = g_RegVal
    '''   For Each g_Site In TheExec.Sites
    '''        If AHBVal <> 2 Then F_IDS_AWAKE = True
    '''    Next g_Site
    '''
    '''TestName = TNameCombine("IDS", "Power", "Status", , , , TName_NonTrimItem, , NHLV, TName_None)
    '''TheExec.Flow.TestLimit AHBVal, 2, 2, , , , , , TestName

    '===================== power pin reset for fit range =====================
    TheHdw.DCVI.Pins("VDD_DIG_UVI80").SetCurrentAndRange 200# * mA, 200 * mA
    TheHdw.DCVI.Pins("VDDC_UVI80").SetCurrentAndRange 200# * mA, 200 * mA
    TheHdw.DCVI.Pins("VDDIO_UVI80").SetCurrentAndRange 200# * mA, 200 * mA
    TheHdw.DCVI.Pins("VDDH_UVI80").SetCurrentAndRange 1000# * mA, 2000 * mA

    '===================== Special Pin status setup according to project =====================
    TheHdw.PPMU.Pins("CRASH_L").Disconnect
    TheHdw.Digital.Pins("CRASH_L").Connect

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



