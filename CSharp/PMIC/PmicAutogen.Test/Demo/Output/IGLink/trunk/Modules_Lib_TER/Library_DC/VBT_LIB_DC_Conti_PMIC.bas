Attribute VB_Name = "VBT_LIB_DC_Conti_PMIC"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.57.85/ with Build Version - 2.23.57.85
'Test Plan:Z:\OTC1\Vance\Corella\Sylvester_A0_TestPlan_190326-3.xlsx, MD5=e17350c3b0c250eeaa3d8694aa12fc3e
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:D:\T-Autogen
'Server Uri:URI: https://stdp-c651-tt.tsmc:426/svn/VBT_Library/
'VBT Version in Server:Version:576
'VBT Version Used:Version:576
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\.svn -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\Characterization -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\Common -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\DC -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\Digital -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\eFuse -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\HardIP -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\Mbist -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\PFA_SONE -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\PMIC -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\SPIROM -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\TMPS_ADCTrim_FreqSync_Det -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\VDDBinning -> Version[576]
'VBT Version:O:\Teradyne\ADC\L\Log\LibBas\Wireless -> Version[576]
Option Explicit

Public g_RelayOnIndex As Integer
Public g_RelayStarNum As Integer


'Type T_ContiGroup
'    PinName As String
'    SpecificForceI As String
'    SpecificWaitTime As String
'    MustDiscnctPin As String
'    OnRelay As String
'    OffRelay As String
'    SpecialCondiPin As String
'    SpacialConPinVolt As String
'    TestItem As String
'End Type

Public ContiPinDic As New Dictionary


Public Const IO_COND_SPLIT = ","
Public Const PIN_COND_SPLIT = ";"
Public Const VOLT_CUR_SPLIT = "/"

Private Enum ContiDataLog
    eOPEN = 0
    eSHORT = 1
    eTotalDataLogItem = 1
End Enum


Public Function IO_Continuity_Parallel(digital_pins As PinList, force_i As Double, TestLimitMode As tlLimitForceResults, _
                                       Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double, _
                                       Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short", Optional connect_all_pins As PinList) As Long

    Dim IOMeasure As New PinListData
    Dim DUTPin    As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim TName     As String
    Dim TestInstanceName As String
    Dim bSerialTest As Boolean

    TestInstanceName = LCase(TheExec.DataManager.InstanceName)

    '20180618 : keep going to check each pin status if parallel test fail
    bSerialTest = False

    '////////////////////////////////////////////////
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IO_Continuity_Parallel"

    If Flag_Open Like "" Then Flag_Open = "F_open"
    If Flag_Short Like "" Then Flag_Short = "F_short"

    '20180822 evans : it's for fix spike issue
    TheHdw.DCVI.Pins("All_Power").BleederResistor = tlDCVIBleederResistorOn
    TheHdw.Wait 30# * ms
    TheHdw.DCVI.Pins("All_Power").BleederResistor = tlDCVIBleederResistorOff

    TheHdw.Digital.ApplyLevelsTiming True, True, True
    TheHdw.Wait 5# * ms

    Call IO_Continuity_Parallel_PreSetup
    '------------------------------------------------------------------------------------------------

    '*************   Test digital channel continuity  ******************************
    ''''Disconnect os_pins Pin Electronics from pins in order to connect PPMU's''''
    TheHdw.Digital.Pins(digital_pins).Disconnect
    If connect_all_pins <> "" Then
        TheHdw.Digital.Pins(connect_all_pins).Disconnect
        With TheHdw.PPMU.Pins(connect_all_pins)
            .ForceV 0#
            .Connect
            .Gate = tlOn
        End With
    End If
    '''''' Connect all os_pins to ppmu and ppmu force 0v for each one
    With TheHdw.PPMU.Pins(digital_pins)
        .ForceV 0#
        .Connect
        .Gate = tlOn
    End With

    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt

    IOMeasure.AddPin (digital_pins)

    With TheHdw.PPMU.Pins(digital_pins)
        .ClampVHi = 2
        .ForceI (force_i)
    End With

    TheHdw.Wait 0.002

    DebugPrintFunc_PPMU CStr(digital_pins)
    IOMeasure = TheHdw.PPMU.Pins(digital_pins).Read(tlPPMUReadMeasurements, 20)    'normal measure

    'recover measure dut pin to 0V before next pin
    TheHdw.PPMU.Pins(digital_pins).ForceV (0)    'correct it to force v, not force i.

    bSerialTest = IOContiDatalog(IOMeasure, force_i, TestLimitMode, False, HiLimit_Short, LowLimit_Short, HiLimit_Open, LowLimit_Open, Flag_Open, Flag_Short)

    '20180618 : keep going to check each pin status if parallel test fail
    If bSerialTest = True Then
        Call IO_Continuity_Serial(digital_pins, force_i, TestLimitMode, HiLimit_Short, LowLimit_Short, HiLimit_Open, LowLimit_Open, Flag_Open, Flag_Short, connect_all_pins)
    End If

    Call IO_Continuity_Parallel_PostSetup

    'Disconnect PPMU from digital channels
    With TheHdw.PPMU.Pins(digital_pins)
        .Gate = tlOff
        .Disconnect
    End With

    If connect_all_pins <> "" Then
        With TheHdw.PPMU.Pins(connect_all_pins)
            '.ForceV 0#
            .Gate = tlOff
            .Disconnect
        End With
    End If

    Exit Function
ErrHandler:
    TheExec.AddOutput "Error in Continuity"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Split_Concat(IP As String, idx As Long) As String
    Dim OP()      As String
    Dim i         As Double

    For i = 0 To UBound(Split(IP, "&"))
        ReDim Preserve OP(i)
        OP(i) = Split(IP, "&")(i)
    Next i
    Split_Concat = OP(idx)
    Exit Function
End Function

Public Sub IO_Continuity_Parallel_PreSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IO_Continuity_Parallel_PreSetup"


    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub IO_Continuity_Parallel_PostSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IO_Continuity_Parallel_PostSetup"



    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function IO_Continuity_Serial(digital_pins As PinList, force_i As Double, TestLimitMode As tlLimitForceResults, _
                                     Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double, _
                                     Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short", Optional connect_all_pins As PinList) As Long

    Dim IOMeasure As New PinListData
    Dim DUTPin    As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim TName     As String
    Dim TestInstanceName As String
    Dim bSerialTest As Boolean

    TestInstanceName = LCase(TheExec.DataManager.InstanceName)

    '////////////////////////////////////////////////
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IO_Continuity_Serial"

    If Flag_Open Like "" Then Flag_Open = "F_open"
    If Flag_Short Like "" Then Flag_Short = "F_short"

    TheHdw.Digital.ApplyLevelsTiming True, True, True
    TheHdw.Wait 5# * ms

    Call IO_Continuity_Serial_PreSetup

    '*************   Test digital channel continuity  ******************************
    ''''Disconnect os_pins Pin Electronics from pins in order to connect PPMU's''''
    TheHdw.Digital.Pins(digital_pins).Disconnect
    If connect_all_pins <> "" Then
        TheHdw.Digital.Pins(connect_all_pins).Disconnect
        With TheHdw.PPMU.Pins(connect_all_pins)
            .ForceV 0#
            .Connect
            .Gate = tlOn
        End With
    End If
    '''''' Connect all os_pins to ppmu and ppmu force 0v for each one
    With TheHdw.PPMU.Pins(digital_pins)
        .ForceV 0#
        .Connect
        .Gate = tlOn
    End With

    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt

    For Each DUTPin In Pins
        IOMeasure.AddPin (DUTPin)
        With TheHdw.PPMU.Pins(DUTPin)
            .ClampVHi = 2
            .ForceI (force_i)
        End With

        'perpin Setting before measure ( Relay / wait time ... etc )

        'Call Conti_PerPinSetting_beforeMeasure(DUTPin)

        TheHdw.Wait 0.005

        '''        DebugPrintFunc_PPMU CStr(DUTPin)
        IOMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)    'normal measure

        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each Site In TheExec.Sites
                TheExec.Datalog.WriteComment ("** Offline Mode **")
                If LCase(TheExec.DataManager.InstanceName) Like "*neg*" Then IOMeasure.Pins(DUTPin).Value(Site) = -0.5    ' Use fixed value for easily checking offline mode.
                If LCase(TheExec.DataManager.InstanceName) Like "*pos*" Then IOMeasure.Pins(DUTPin).Value(Site) = 0.5    ' Use fixed value for easily checking offline mode.
            Next Site
        End If

        'recover measure dut pin to 0V before next pin
        TheHdw.PPMU.Pins(DUTPin).ForceV (0)

        'Call Conti_PerPinSetting_afterMeasure(DUTPin)

    Next DUTPin

    Call IOContiDatalog(IOMeasure, force_i, TestLimitMode, True, HiLimit_Short, LowLimit_Short, HiLimit_Open, LowLimit_Open, Flag_Open, Flag_Short)

    Call IO_Continuity_Serial_PostSetup

    'Disconnect PPMU from digital channels
    With TheHdw.PPMU.Pins(digital_pins)
        .Gate = tlOff
        .Disconnect
    End With

    '    thehdw.Digital.Pins(digital_pins).Connect

    If connect_all_pins <> "" Then
        ' TheHdw.Digital.Pins (connect_all_pins)
        With TheHdw.PPMU.Pins(connect_all_pins)
            '.ForceV 0#
            .Gate = tlOff
            .Disconnect
        End With
    End If

    Exit Function
ErrHandler:
    TheExec.AddOutput "Error in Continuity"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Sub IO_Continuity_Serial_PreSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IO_Continuity_Serial_PreSetup"


    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub IO_Continuity_Serial_PostSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IO_Continuity_Serial_PostSetup"


    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function IOContiDatalog(MeasReault As PinListData, force_i As Double, TestLimitMode As tlLimitForceResults, bSerial As Boolean, _
                               Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double, _
                               Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short") As Boolean

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "IOContiDatalog"


    Dim PinSequence As Integer
    Dim ContiLog  As Integer
    Dim i         As Integer
    Dim Site      As Variant
    Dim LowLimit, HiLimit As Double
    Dim FlowLimitObj As IFlowLimitsInfo
    Dim Lolimit_new() As Double
    Dim HiLimit_new() As Double
    Dim Lolimit_str() As String
    Dim Hilimit_str() As String
    Dim TName     As String
    Dim TestName  As String

    If TestLimitMode = tlForceFlow Then
        Call TheExec.Flow.GetTestLimits(FlowLimitObj)
        Call FlowLimitObj.GetLowLimits(Lolimit_str)
        Call FlowLimitObj.GetHighLimits(Hilimit_str)

        ReDim Lolimit_new(UBound(Lolimit_str))
        ReDim HiLimit_new(UBound(Hilimit_str))

        For i = 0 To UBound(Lolimit_str)
            If Lolimit_str(i) <> "" Then Lolimit_new(i) = CDbl(Lolimit_str(i))
        Next i
        For i = 0 To UBound(Hilimit_str)
            If Hilimit_str(i) <> "" Then HiLimit_new(i) = CDbl(Hilimit_str(i))
        Next i
    End If

    PinSequence = 0
    IOContiDatalog = False

    For ContiLog = 0 To eTotalDataLogItem
        For i = 0 To MeasReault.Pins.Count - 1
            If ContiLog = ContiDataLog.eOPEN Then
                If bSerial = True Then
                    If TheExec.DataManager.InstanceName Like "*PostBurn*" Then
                        DC_CreateTestName "PostBurn-OPEN-SERIAL"
                    Else
                        DC_CreateTestName "OPEN-SERIAL"
                    End If
                Else
                    If TheExec.DataManager.InstanceName Like "*PostBurn*" Then
                        DC_CreateTestName "PostBurn-OPEN"
                    Else
                        DC_CreateTestName "OPEN"
                    End If
                End If
            ElseIf ContiLog = ContiDataLog.eSHORT Then
                If bSerial = True Then
                    If TheExec.DataManager.InstanceName Like "*PostBurn*" Then
                        DC_CreateTestName "PostBurn-SHORT-SERIAL"
                    Else
                        DC_CreateTestName "SHORT-SERIAL"
                    End If
                Else
                    If TheExec.DataManager.InstanceName Like "*PostBurn*" Then
                        DC_CreateTestName "PostBurn-SHORT"
                    Else
                        DC_CreateTestName "SHORT"
                    End If
                End If
            Else
                TheExec.Datalog.WriteComment "ContiDatalogForParallel : Please Check the Conti Datalog Condition!!!"
                '                Stop
                Exit Function    '//2019_1213
            End If

            DC_UpdatePinName UCase(CStr(MeasReault.Pins(i))), CStr(force_i * 10 ^ 6) & "uA"

            '''            HiLimit = TheExec.Flow.TestLimit.GetHiLimit(Replace(DC_GetTestName, "-SERIAL", ""))
            '''            LowLimit = TheExec.Flow.TestLimit.GetLowLimit(Replace(DC_GetTestName, "-SERIAL", ""))

            If TheExec.DataManager.ChannelType(MeasReault.Pins(i)) <> "N/C" Then  'if N/C jump next pin

                TName = DC_GetTestName
                If TestLimitMode = tlForceFlow Then
                    LL.FlowTestLimit ResultVal:=MeasReault.Pins(MeasReault.Pins(i)), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=TName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    '                    TheExec.Flow.TestLimit ResultVal:=MeasReault.Pins(MeasReault.Pins(i)), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=TName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    'Judge failed open or failed short for tlForceFlow
                    For Each Site In TheExec.Sites
                        If MeasReault.Pins(MeasReault.Pins(i)).Value(Site) > HiLimit_new(PinSequence) Then
                            If HiLimit_new(PinSequence) > 0 Then
                                TheExec.Sites.Item(Site).FlagState(Flag_Open) = logicTrue
                                IOContiDatalog = True
                            Else
                                TheExec.Sites.Item(Site).FlagState(Flag_Short) = logicTrue
                                IOContiDatalog = True
                            End If

                        ElseIf MeasReault.Pins(MeasReault.Pins(i)).Value(Site) < Lolimit_new(PinSequence) Then
                            If Lolimit_new(PinSequence) > 0 Then
                                TheExec.Sites.Item(Site).FlagState(Flag_Short) = logicTrue
                                IOContiDatalog = True
                            Else
                                TheExec.Sites.Item(Site).FlagState(Flag_Open) = logicTrue
                                IOContiDatalog = True
                            End If

                        End If
                    Next Site
                    '///////////////////////////////////////////////////////////////////////////////////////
                ElseIf TestLimitMode = tlForceNone Then
                    'offline mode simulation
                    If TheExec.TesterMode = testModeOffline Then
                        For Each Site In TheExec.Sites
                            If InStr(UCase(TheExec.DataManager.InstanceName), "NEG") > 0 Then
                                MeasReault.Pins(MeasReault.Pins(i)).Value(Site) = -0.5 + Rnd() * 0.1
                            ElseIf InStr(UCase(TheExec.DataManager.InstanceName), "POS") > 0 Then
                                MeasReault.Pins(MeasReault.Pins(i)).Value(Site) = 0.5 + Rnd() * 0.1
                            Else
                                MeasReault.Pins(MeasReault.Pins(i)).Value(Site) = 0.5 + Rnd() * 0.1
                            End If
                        Next Site
                    End If

                    If InStr(UCase(TheExec.DataManager.InstanceName), "NEG") > 0 Then
                        If ContiLog = ContiDataLog.eOPEN Then
                            HiLimit = LL.GetHiLimit(DC_GetTestName)
                            LowLimit = LL.GetLowLimit(DC_GetTestName)
                            '                            LowLimit = LowLimit_Open '-1.2 * v
                            '                            HiLimit = HiLimit_Open '200 * mV
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(MeasReault.Pins(i)).Value(Site) < LowLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_open") = logicTrue
                                    IOContiDatalog = True
                                End If
                            Next Site
                        ElseIf ContiLog = ContiDataLog.eSHORT Then
                            HiLimit = LL.GetHiLimit(DC_GetTestName)
                            LowLimit = LL.GetLowLimit(DC_GetTestName)
                            '                            LowLimit = LowLimit_Short '-3 * v
                            '                            HiLimit = HiLimit_Short '-200 * mV
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(MeasReault.Pins(i)).Value(Site) > HiLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_short") = logicTrue
                                    IOContiDatalog = True
                                End If
                            Next Site
                        End If
                        LL.FlowTestLimit TName:=DC_GetTestName, ResultVal:=MeasReault.Pins(MeasReault.Pins(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        '                        TheExec.Flow.TestLimit TName:=DC_GetTestName, ResultVal:=MeasReault.Pins(MeasReault.Pins(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                    Else
                        If ContiLog = ContiDataLog.eOPEN Then
                            HiLimit = LL.GetHiLimit(TestName)
                            LowLimit = LL.GetLowLimit(TestName)
                            '                            LowLimit = LowLimit_Open '-200 * mV
                            '                            HiLimit = HiLimit_Open '1.2 * v
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(MeasReault.Pins(i)).Value(Site) > HiLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_open") = logicTrue
                                    IOContiDatalog = True
                                End If
                            Next Site
                        ElseIf ContiLog = ContiDataLog.eSHORT Then
                            HiLimit = LL.GetHiLimit(TestName)
                            LowLimit = LL.GetLowLimit(TestName)
                            '                            LowLimit = LowLimit_Short '200 * mV
                            '                            HiLimit = HiLimit_Short '1.2 * v
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(MeasReault.Pins(i)).Value(Site) < LowLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_short") = logicTrue
                                    IOContiDatalog = True
                                End If
                            Next Site
                        End If
                        LL.FlowTestLimit TName:=DC_GetTestName, ResultVal:=MeasReault.Pins(MeasReault.Pins(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        '                        TheExec.Flow.TestLimit TName:=DC_GetTestName, ResultVal:=MeasReault.Pins(MeasReault.Pins(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone

                    End If

                End If
            End If

            PinSequence = PinSequence + 1
        Next i

    Next ContiLog

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Power_Continuity_Parallel(analog_pins As PinList, force_i As Double, TestLimitMode As tlLimitForceResults, _
                                          Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double _
                                                                                                                                                ) As Long

    Dim PowerMeasure As New PinListData
    Dim PinGroup  As IPinListData
    Dim DUTPin    As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim Power_conti_volt As New PinListData
    Dim PPMUMeas_HexVs As New PinListData
    Dim i         As Long
    Dim TName     As String
    Dim Site      As Variant
    Dim TestInstanceName As String
    Dim MeasPins  As String
    Dim bSerialTest As Boolean

    TestInstanceName = LCase(TheExec.DataManager.InstanceName)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Power_Continuity_Parallel"

    'Call Power_Continuity_Parallel_PreSetup

    '20180618 : keep going to check each pin status if parallel test fail
    bSerialTest = False

    PowerMeasure.AddPin (analog_pins)

    With TheHdw.DCVI.Pins(analog_pins)
        .Gate = False
        .Mode = tlDCVIModeCurrent
        .Current = force_i
        .SetVoltageAndRange -1, 7
        .Connect tlDCVIConnectDefault
        TheHdw.Wait 1 * ms
        .Gate = True
    End With

    'measure
    TheHdw.DCVI.Pins(analog_pins).Meter.Mode = tlDCVIMeterVoltage
    TheHdw.Wait 30 * ms  '3 * ms
    PowerMeasure = TheHdw.DCVI.Pins(analog_pins).Meter.Read(tlStrobe, 1, 10 * KHz)

    'reset
    With TheHdw.DCVI.Pins(analog_pins)
        .Current = 0#
        .Gate = False
        .Disconnect tlDCVIConnectDefault
        .Mode = tlDCVIModeVoltage
    End With

    bSerialTest = ContiDatalogForParallel(PowerMeasure, force_i, TestLimitMode, HiLimit_Short, LowLimit_Short, HiLimit_Open, LowLimit_Open)

    'Call Power_Continuity_Parallel_PostSetup

    '20180618 : keep going to check each pin status if parallel test fail
    If bSerialTest = True Then
        Call Power_Continuity_Serial(analog_pins, force_i, TestLimitMode, HiLimit_Short, LowLimit_Short, HiLimit_Open, LowLimit_Open)
    End If

    Exit Function
ErrHandler:
    TheExec.AddOutput "Error in Continuity"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'Public Sub Power_Continuity_Parallel_PreSetup()
'
'End Sub
'
'
'Public Sub Power_Continuity_Parallel_PostSetup()
'
'        ElseIf (LCase(A_Slot_Type(i)) = "dc-07") Then
'            p_dc07 = p_dc07 & "," & Pin_Ary(i)
'        End If
'    Next i
'
'    If (p_hexvs <> "" Or p_uvs <> "") Then
'        If InitRange200mAPins <> "" Then
'            TheHdw.DCVS.Pins(InitRange200mAPins).SetCurrentRanges 0.2, 0.2
'            TheHdw.DCVS.Pins(InitRange200mAPins).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
'        End If
'        If InitRange20mAPins <> "" Then
'            TheHdw.DCVS.Pins(InitRange20mAPins).SetCurrentRanges 0.02, 0.02
'            TheHdw.DCVS.Pins(InitRange20mAPins).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
'        End If
'        If InitRange2mAPins <> "" Then
'            TheHdw.DCVS.Pins(InitRange2mAPins).SetCurrentRanges 0.002, 0.002
'            TheHdw.DCVS.Pins(InitRange2mAPins).CurrentLimit.Source.FoldLimit.Level.Value = 0.002
'        End If
'
'        '=======================================================================================================
'        If p_hexvs <> "" Then
'            p_hexvs = Right(p_hexvs, Len(p_hexvs) - 1)
'            If auto_range_flag = True Then
'                TheHdw.DCVS.Pins(p_hexvs).Voltage.Main.Value = ForceV
'                TheHdw.Wait 100 * ms
'            End If
'            HexVS_Power_data = TheHdw.DCVS.Pins(p_hexvs).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
'        End If
'        If p_uvs <> "" Then
'            p_uvs = Right(p_uvs, Len(p_uvs) - 1)
'            If auto_range_flag = True Then
'                TheHdw.DCVS.Pins(p_uvs).Voltage.Main.Value = ForceV
'                TheHdw.Wait 5 * ms
'            End If
'            UVS_Power_data = TheHdw.DCVS.Pins(p_uvs).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
'        End If
'
'        '=======================================================================================================
'        TheHdw.DCVS.Pins(PowerPins).Voltage.Main.Value = 0#
'        TheHdw.Wait 3 * ms
'        'Start search I range
'        For i = 0 To UBound(Pin_Ary)
'            TheHdw.DCVS.Pins(Pin_Ary(i)).Voltage.Main.Value = ForceV
'            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            If (LCase(A_Slot_Type(i)) = "hexvs") Then
'                If TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.CurrentRange = 0.01 Then
'                    TheHdw.Wait 100 * ms
'                Else
'                    TheHdw.Wait 30 * ms
'                End If
'            Else
'                TheHdw.Wait 5 * ms  'align Cyprus TTR
'            End If
'
'            If (LCase(A_Slot_Type(i)) = "hexvs") Then
'                '===============================================================================auto range
'                If auto_range_flag = True Then
'                    For Each Site In TheExec.sites
'                        StepNo = j + step_ary(i)
'                        If StepNo = 6 Then j = Stop_Step
'                        Val = Abs(HexVS_Power_data.Pins(Pin_Ary(i)).Value(Site))
'                        Select Case StepNo
'                            Case 1: '15A => 1A
'                                    If ((Val + (0.05 + 0.15) * 2) < 1) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 1, 1: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 1
''                                    SattleTime = 1 * ms
''                                    If SattleTime > WaitTime Then WaitTime = SattleTime
'                            Case 2: '1A => 100mA
'                                    If ((Val + (0.01 + 0.005) * 2) < 0.1) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.1, 0.1: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.1
''                                    SattleTime = 10 * ms
''                                    If SattleTime > WaitTime Then WaitTime = SattleTime
'    '                        Case 3: '100mA => 10mA
'    '                                If ((Val + 0.001*2) < 0.01) Then TheHdw.DCVS.Pins(Power_data.Pins(i)).SetCurrentRanges 0.01, 0.01: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
'    '                                SattleTime = 100 * ms
'    '                                If SattleTime > WaitTime Then WaitTime = SattleTime
'                        End Select
'                    Next Site
'                    Wait 0.01
'                End If
'                '===============================================================================
'                HexVS_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
'            ElseIf (LCase(A_Slot_Type(i)) = "vhdvs") Then
'                '===============================================================================auto range
'                If auto_range_flag = True Then
'                    For Each Site In TheExec.sites
'                        StepNo = j + step_ary(i)
'                        Val = Abs(UVS_Power_data.Pins(Pin_Ary(i)).Value(Site))
'
'                        Select Case StepNo
'                            Case 1: '200mA => 40mA
'                                    If ((Val + (0.0012 + 0.001) * 2) < 0.04) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.04, 0.04: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.04: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.04
''                                    SattleTime = 260 * us
''                                    If SattleTime > WaitTime Then WaitTime = SattleTime
'                            Case 2: '40mA =>20mA
'                                    If ((Val + (0.00024 + 0.0003) * 2) < 0.02) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.02, 0.02: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.02: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.02
''                                    SattleTime = 540 * us
''                                    If SattleTime > WaitTime Then WaitTime = SattleTime
'                            Case 3: '20mA =>2mA
'                                    If ((Val + (0.00012 + 0.0001) * 2) < 0.002) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.002, 0.002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.002
''                                    SattleTime = 3.5 * ms
''                                    If SattleTime > WaitTime Then WaitTime = SattleTime
'                            Case 4: '2mA =>200uA
'                                    If ((Val + (0.000012 + 0.00001) * 2) < 0.0002) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.0002, 0.0002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.0002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.0002
''                                    SattleTime = 4 * ms
''                                    If SattleTime > WaitTime Then WaitTime = SattleTime
'    '                        Case 5: '200uA =>20uA
'    '                                If ((Val + (0.0000012 + 0.000001) * 2) < 0.00002) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.00002, 0.00002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.00002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.00002
'    '                                SattleTime = 4 * ms
'    '                                If SattleTime > WaitTime Then WaitTime = SattleTime
'    '                        Case 6: '20uA =>4uA
'    '                                If ((Val + (0.00000012 + 0.0000001) * 2) < 0.000004) Then thehdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.000004, 0.000004: thehdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.000004: thehdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.000004
'    '                                SattleTime = 18 * ms
'    '                                If SattleTime > WaitTime Then WaitTime = SattleTime
'                        End Select
'                    Next Site
'                    Wait 0.004
'                End If
'                '===============================================================================
'                UVS_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 1)
'            End If
'
'            TheHdw.DCVS.Pins(Pin_Ary(i)).Voltage.Main.Value = 0#
'
'            Tname = "pwr_FVMI_" & Pin_Ary(i)
'            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'            If (LCase(A_Slot_Type(i)) = "hexvs") Then
'                'offline mode simulation
'                If TheExec.TesterMode = testModeOffline Then
'                    For Each Site In TheExec.sites
'                        HexVS_Power_data.Pins(Pin_Ary(i)).Value(Site) = 0.01 + Rnd() * 0.0001
'                    Next Site
'                End If
'
'                If TestLimitMode = tlForceFlow Then
'                    TheExec.Flow.TestLimit ResultVal:=HexVS_Power_data.Pins(Pin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, forceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
'                ElseIf TestLimitMode = tlForceNone Then
'                    TheExec.Flow.TestLimit ResultVal:=HexVS_Power_data.Pins(Pin_Ary(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, forceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
'                End If
'            ElseIf (LCase(A_Slot_Type(i)) = "vhdvs") Then
'                'offline mode simulation
'                If TheExec.TesterMode = testModeOffline Then
'                    For Each Site In TheExec.sites
'                        UVS_Power_data.Pins(Pin_Ary(i)).Value(Site) = 0.01 + Rnd() * 0.0001
'                    Next Site
'                End If
'
'                If TestLimitMode = tlForceFlow Then
'                    TheExec.Flow.TestLimit ResultVal:=UVS_Power_data.Pins(Pin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, forceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
'                ElseIf TestLimitMode = tlForceNone Then
'                    TheExec.Flow.TestLimit ResultVal:=UVS_Power_data.Pins(Pin_Ary(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, forceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
'                End If
'            End If
'            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'        Next i
'
'
'End Sub

Public Function Power_Continuity_Serial(analog_pins As PinList, force_i As Double, TestLimitMode As tlLimitForceResults, _
                                        Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double _
                                                                                                                                              ) As Long

    Dim PowerMeasure As New PinListData
    Dim PinGroup  As IPinListData
    Dim DUTPin    As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim Power_conti_volt As New PinListData
    Dim PPMUMeas_HexVs As New PinListData
    Dim i         As Long
    Dim TName     As String
    Dim Site      As Variant
    Dim TestInstanceName As String

    TestInstanceName = LCase(TheExec.DataManager.InstanceName)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Power_Continuity_Serial"

    Call Power_Continuity_Serial_PreSetup

    TheExec.DataManager.DecomposePinList analog_pins, Pins(), Pin_Cnt

    For Each DUTPin In Pins
        PowerMeasure.AddPin (DUTPin)

        If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then

            With TheHdw.DCVI.Pins(DUTPin)
                .Gate = False
                .Mode = tlDCVIModeCurrent
                .Current = force_i
                .SetVoltageAndRange -1, 7
                .Connect tlDCVIConnectDefault
                TheHdw.Wait (10 * ms)
                .Gate = True
            End With

            'measure
            TheHdw.DCVI.Pins(DUTPin).Meter.Mode = tlDCVIMeterVoltage
            TheHdw.Wait (0.03)

            'Call Conti_PerPinSetting_beforeMeasure(DUTPin)
            PowerMeasure.Pins(DUTPin) = TheHdw.DCVI.Pins(DUTPin).Meter.Read(tlStrobe, 1, 10 * KHz)

            'reset
            With TheHdw.DCVI.Pins(DUTPin)
                .Current = 0#
                .Gate = False
                .Disconnect tlDCVIConnectDefault
                .Mode = tlDCVIModeVoltage
            End With

        End If

        'Call Conti_PerPinSetting_afterMeasure(DUTPin)

    Next DUTPin

    Call ContiDatalogForSerial(PowerMeasure, DUTPin, Pins, force_i, TestLimitMode, HiLimit_Short, LowLimit_Short, HiLimit_Open, LowLimit_Open)

    Call Power_Continuity_Serial_PostSetup

    Exit Function
ErrHandler:
    TheExec.AddOutput "Error in Continuity"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Sub Power_Continuity_Serial_PreSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Power_Continuity_Serial_PreSetup"

    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Power_Continuity_Serial_PostSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Power_Continuity_Serial_PostSetup"

    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function Analog_Continuity_Parallel(analog_pins As PinList, force_i As Double, TestLimitMode As tlLimitForceResults, _
                                           Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double _
                                                                                                                                                 ) As Long

    Dim AnalogMeasure As New PinListData
    Dim DUTPin    As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim TName     As String
    Dim TestInstanceName As String
    Dim bSerialTest As Boolean

    TestInstanceName = LCase(TheExec.DataManager.InstanceName)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Analog_Continuity_Parallel"

    Call Analog_Continuity_Parallel_PreSetup

    '20180618 : keep going to check each pin status if parallel test fail
    bSerialTest = False

    AnalogMeasure.AddPin (analog_pins)

    With TheHdw.DCVI.Pins(analog_pins)
        .Gate = False
        .Mode = tlDCVIModeCurrent
        .Current = force_i
        .SetVoltageAndRange -1, 7
        .Connect tlDCVIConnectDefault
        TheHdw.Wait 1 * ms
        .Gate = True
    End With

    'measure
    TheHdw.DCVI.Pins(analog_pins).Meter.Mode = tlDCVIMeterVoltage
    TheHdw.Wait 3 * ms
    AnalogMeasure = TheHdw.DCVI.Pins(analog_pins).Meter.Read(tlStrobe, 1, 10 * KHz)

    'reset
    With TheHdw.DCVI.Pins(analog_pins)
        .Current = 0#
        .Gate = False
        .Disconnect tlDCVIConnectDefault
        .Mode = tlDCVIModeVoltage
    End With

    bSerialTest = ContiDatalogForParallel(AnalogMeasure, force_i, TestLimitMode, HiLimit_Short, LowLimit_Short, HiLimit_Open, LowLimit_Open)

    Call Analog_Continuity_Parallel_PostSetup

    '20180618 : keep going to check each pin status if parallel test fail
    If bSerialTest = True Then
        Call Analog_Continuity_Serial(analog_pins, force_i, TestLimitMode, HiLimit_Short, LowLimit_Short, HiLimit_Open, LowLimit_Open)
    End If

    Exit Function
ErrHandler:
    TheExec.AddOutput "Error in Continuity"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Sub Analog_Continuity_Parallel_PreSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Analog_Continuity_Parallel_PreSetup"

    If InStr(UCase(TheExec.DataManager.InstanceName), "ANALOG") > 0 Then
        '20180517 : connect Buck6~12 to UVI80 S
        'TheHdw.Utility("k3902,k3903,k3904,k3905,k3906,k3907").State = tlUtilBitOn
        '20180517 : connect Buck6~12 to UVI80 F
        'TheHdw.Utility("k1601,k1701,k1801,k1901,k2001,k2201").State = tlUtilBitOn
    End If

    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Analog_Continuity_Parallel_PostSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Analog_Continuity_Parallel_PostSetup"

    If InStr(UCase(TheExec.DataManager.InstanceName), "ANALOG") > 0 Then
        '20180517 : disconnect Buck6~12 to UVI80 S
        'TheHdw.Utility("k3902,k3903,k3904,k3905,k3906,k3907").State = tlUtilBitOff
        '20180517 : disconnect Buck6~12 to UVI80 F
        'TheHdw.Utility("k1601,k1701,k1801,k1901,k2001,k2201").State = tlUtilBitOff
    End If

    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function Analog_Continuity_Serial(analog_pins As PinList, force_i As Double, TestLimitMode As tlLimitForceResults, _
                                         Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double _
                                                                                                                                               ) As Long

    Dim AnalogMeasure As New PinListData
    Dim DUTPin    As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim TName     As String
    Dim TestInstanceName As String
    'Dim m_Pins() As T_ContiGroup

    TestInstanceName = LCase(TheExec.DataManager.InstanceName)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Analog_Continuity_Serial"

    Call Analog_Continuity_Serial_PreSetup

    TheExec.DataManager.DecomposePinList analog_pins, Pins(), Pin_Cnt


    For Each DUTPin In Pins
        AnalogMeasure.AddPin (DUTPin)
        If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then

            Call Analog_Continuity_Serial_PrePerPinSetup(DUTPin)

            With TheHdw.DCVI.Pins(DUTPin)
                .Gate = False
                If TheHdw.DCVI.Pins(DUTPin).DCVIType = "DC07" Then
                    .BleederResistor = tlDCVIBleederResistorOff
                End If
                .Disconnect tlDCVIConnectDefault
                .Mode = tlDCVIModeCurrent
                '                .Current = force_i
                ' .CurrentRange = force_i
                '                .CurrentRange = 0.002

                'If TheHdw.DCVI.Pins(DUTPin).Current > force_i Then
                If Abs(TheHdw.DCVI.Pins(DUTPin).Current) > Abs(force_i) Then  'add by kevin
                    .Current = force_i
                    .CurrentRange = force_i

                Else
                    .CurrentRange = force_i
                    .Current = force_i

                End If

                ''                .SetCurrentAndRange force_i, force_i
                If TheHdw.DCVI.Pins(DUTPin).DCVIType = "DC07" Then
                    If (force_i > 0) Then
                        .SetVoltageAndRange 1, 7
                    Else
                        .SetVoltageAndRange -1, 7
                    End If
                Else
                    If (force_i > 0) Then
                        .SetVoltageAndRange 5, 7

                    Else
                        .SetVoltageAndRange -5, 7
                    End If
                End If
                '                .SetCurrentAndRange force_i, force_i

                .Connect tlDCVIConnectDefault
                TheHdw.Wait 1 * ms
                .Gate = True
            End With

            'measure
            TheHdw.DCVI.Pins(DUTPin).Meter.Mode = tlDCVIMeterVoltage
            TheHdw.DCVI.Pins(DUTPin).Meter.VoltageRange = 2

            If (DUTPin = "VDDNEG_SUB_DC30") Then
                TheHdw.Wait 30 * ms
            Else
                TheHdw.Wait 20 * ms
            End If
            'perpin Setting before measure ( Relay / wait time ... etc )

            'Call Conti_PerPinSetting_beforeMeasure(DUTPin)
            '                If (DUTPin = "VDDNEG_AZ_DC30") Then Stop
            AnalogMeasure.Pins(DUTPin) = TheHdw.DCVI.Pins(DUTPin).Meter.Read(tlStrobe, 1, 10 * KHz)

            'reset
            With TheHdw.DCVI.Pins(DUTPin)
                .Current = 0#
                .Voltage = 0
                .Gate = False
                .Disconnect tlDCVIConnectDefault
                .Mode = tlDCVIModeVoltage
            End With

            Call Analog_Continuity_Serial_PostPerPinSetup(DUTPin)

        End If

        'Call Conti_PerPinSetting_afterMeasure(DUTPin)

    Next DUTPin

    Call ContiDatalogForSerial(AnalogMeasure, DUTPin, Pins, force_i, TestLimitMode, HiLimit_Short, LowLimit_Short, HiLimit_Open, LowLimit_Open)

    Call Analog_Continuity_Serial_PostSetup

    Exit Function
ErrHandler:
    TheExec.AddOutput "Error in Continuity"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function





Public Sub Analog_Continuity_Serial_PrePerPinSetup(DUTPin As Variant)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Analog_Continuity_Serial_PrePerPinSetup"



    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Analog_Continuity_Serial_PostPerPinSetup(DUTPin As Variant)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Analog_Continuity_Serial_PostPerPinSetup"



    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Analog_Continuity_Serial_PreSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Analog_Continuity_Serial_PreSetup"

    If InStr(UCase(TheExec.DataManager.InstanceName), "ANALOG") > 0 Then
        '20180517 : disconnect Buck6~12 to UVI80 S
        'TheHdw.Utility("k3902,k3903,k3904,k3905,k3906,k3907").State = tlUtilBitOn
        '20180517 : disconnect Buck6~12 to UVI80 F
        'TheHdw.Utility("k1601,k1701,k1801,k1901,k2001,k2201").State = tlUtilBitOn
    End If


    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Analog_Continuity_Serial_PostSetup()
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Analog_Continuity_Serial_PostSetup"

    '    If InStr(UCase(TheExec.DataManager.InstanceName), "ANALOG") > 0 Then
    '        '20180517 : disconnect Buck6~12 to UVI80 S
    '        'TheHdw.Utility("k3902,k3903,k3904,k3905,k3906,k3907").State = tlUtilBitOff
    '        '20180517 : disconnect Buck6~12 to UVI80 F
    '        'TheHdw.Utility("k1601,k1701,k1801,k1901,k2001,k2201").State = tlUtilBitOff
    '    End If


    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub



Public Function ContiDatalogForParallel(MeasReault As PinListData, force_i As Double, TestLimitMode As tlLimitForceResults, _
                                        Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double _
                                                                                                                                              ) As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "ContiDatalogForParallel"

    Dim ContiLog  As Integer
    Dim i         As Integer
    Dim Site      As Variant
    Dim LowLimit, HiLimit As Double

    ContiDatalogForParallel = False

    For ContiLog = 0 To eTotalDataLogItem

        For i = 0 To MeasReault.Pins.Count - 1

            If TheExec.DataManager.ChannelType(MeasReault.Pins(i)) <> "N/C" Then

                If ContiLog = ContiDataLog.eOPEN Then
                    DC_CreateTestName "OPEN"
                ElseIf ContiLog = ContiDataLog.eSHORT Then
                    DC_CreateTestName "SHORT"
                Else
                    TheExec.Datalog.WriteComment "ContiDatalogForParallel : Please Check the Conti Datalog Condition!!!"
                    '                    Stop
                    Exit Function    '//2019_1213
                End If

                DC_UpdatePinName UCase(CStr(MeasReault.Pins(i))), CStr(force_i * 10 ^ 6) & "uA"

                '''                HiLimit = TheExec.Flow.TestLimit.GetHiLimit(DC_GetTestName)
                '''                LowLimit = TheExec.Flow.TestLimit.GetLowLimit(DC_GetTestName)

                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit ResultVal:=MeasReault.Pins(MeasReault.Pins(i)), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Else:
                    TestLimitMode = tlForceNone
                    If TheExec.TesterMode = testModeOffline Then
                        For Each Site In TheExec.Sites
                            If InStr(UCase(TheExec.DataManager.InstanceName), "NEG") > 0 Then
                                MeasReault.Pins(MeasReault.Pins(i)).Value(Site) = -0.5 + Rnd() * 0.1
                            ElseIf InStr(UCase(TheExec.DataManager.InstanceName), "POS") > 0 Then
                                MeasReault.Pins(MeasReault.Pins(i)).Value(Site) = 0.5 + Rnd() * 0.1
                            Else
                                MeasReault.Pins(MeasReault.Pins(i)).Value(Site) = 0.5 + Rnd() * 0.1
                            End If
                        Next Site
                    End If

                    If InStr(UCase(TheExec.DataManager.InstanceName), "NEG") > 0 Then
                        If ContiLog = ContiDataLog.eOPEN Then
                            LowLimit = LowLimit_Open    '-1.2 * v
                            HiLimit = HiLimit_Open    '200 * mV
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(MeasReault.Pins(i)).Value(Site) < LowLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_open") = logicTrue
                                    ContiDatalogForParallel = True
                                End If
                            Next Site
                        ElseIf ContiLog = ContiDataLog.eSHORT Then
                            LowLimit = LowLimit_Short    '-3 * v
                            HiLimit = HiLimit_Short    '-200 * mV
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(MeasReault.Pins(i)).Value(Site) > HiLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_short") = logicTrue
                                    ContiDatalogForParallel = True
                                End If
                            Next Site
                        End If
                        TheExec.Flow.TestLimit ResultVal:=MeasReault.Pins(MeasReault.Pins(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone

                    Else
                        If ContiLog = ContiDataLog.eOPEN Then
                            LowLimit = LowLimit_Open    '-200 * mV
                            HiLimit = HiLimit_Open    '1.2 * v
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(MeasReault.Pins(i)).Value(Site) > HiLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_open") = logicTrue
                                    ContiDatalogForParallel = True
                                End If
                            Next Site
                        ElseIf ContiLog = ContiDataLog.eSHORT Then
                            LowLimit = LowLimit_Short    '200 * mV
                            HiLimit = HiLimit_Short    '1.2 * v
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(MeasReault.Pins(i)).Value(Site) < LowLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_short") = logicTrue
                                    ContiDatalogForParallel = True
                                End If
                            Next Site
                        End If
                        TheExec.Flow.TestLimit ResultVal:=MeasReault.Pins(MeasReault.Pins(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone

                    End If
                End If

            End If

        Next i

    Next ContiLog

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub ContiDatalogForSerial(MeasReault As PinListData, DUTPin As Variant, Pins() As String, force_i As Double, TestLimitMode As tlLimitForceResults, _
                                 Optional HiLimit_Short As Double, Optional LowLimit_Short As Double, Optional HiLimit_Open As Double, Optional LowLimit_Open As Double)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "ContiDatalogForSerial"

    Dim ContiLog  As Integer
    Dim Site      As Variant
    Dim LowLimit, HiLimit As Double
    Dim TestName  As String

    For ContiLog = 0 To eTotalDataLogItem

        For Each DUTPin In Pins

            If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then
                If ContiLog = ContiDataLog.eOPEN Then
                    If TheExec.DataManager.InstanceName Like "*PostBurn*" Then
                        DC_CreateTestName "PostBurn-OPEN-SERIAL"
                    Else
                        DC_CreateTestName "OPEN-SERIAL"
                    End If
                ElseIf ContiLog = ContiDataLog.eSHORT Then
                    If TheExec.DataManager.InstanceName Like "*PostBurn*" Then
                        DC_CreateTestName "PostBurn-SHORT-SERIAL"
                    Else
                        DC_CreateTestName "SHORT-SERIAL"
                    End If
                Else
                    TheExec.Datalog.WriteComment "ContiDatalogForParallel : Please Check the Conti Datalog Condition!!!"
                    '                    Stop
                    Exit Sub    ' //2019_1213
                End If

                DC_UpdatePinName CStr(DUTPin), CStr(force_i * 10 ^ 6) & "uA"

                '''                HiLimit = TheExec.Flow.TestLimit.GetHiLimit(Replace(DC_GetTestName, "-SERIAL", ""))
                '''                LowLimit = TheExec.Flow.TestLimit.GetLowLimit(Replace(DC_GetTestName, "-SERIAL", ""))

                If TestLimitMode = tlForceFlow Then
                    LL.FlowTestLimit ResultVal:=MeasReault.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    '                    TheExec.Flow.TestLimit ResultVal:=MeasReault.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Else: TestLimitMode = tlForceNone

                    If TheExec.TesterMode = testModeOffline Then
                        For Each Site In TheExec.Sites
                            If InStr(UCase(TheExec.DataManager.InstanceName), "NEG") > 0 Then
                                MeasReault.Pins(DUTPin).Value(Site) = -0.5 + Rnd() * 0.1
                            ElseIf InStr(UCase(TheExec.DataManager.InstanceName), "POS") > 0 Then
                                MeasReault.Pins(DUTPin).Value(Site) = 0.5 + Rnd() * 0.1
                            Else
                                MeasReault.Pins(DUTPin).Value(Site) = 0.5 + Rnd() * 0.1
                            End If
                        Next Site
                    End If

                    If InStr(UCase(TheExec.DataManager.InstanceName), "NEG") > 0 Then
                        If ContiLog = ContiDataLog.eOPEN Then
                            HiLimit = LL.GetHiLimit(DC_GetTestName)
                            LowLimit = LL.GetLowLimit(DC_GetTestName)
                            '                            LowLimit = LowLimit_Open '-1.2 * v
                            '                            HiLimit = HiLimit_Open '200 * mV
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(DUTPin).Value(Site) < LowLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_open") = logicTrue
                                End If
                            Next Site
                        ElseIf ContiLog = ContiDataLog.eSHORT Then
                            HiLimit = LL.GetHiLimit(DC_GetTestName)
                            LowLimit = LL.GetLowLimit(DC_GetTestName)
                            '                            LowLimit = LowLimit_Short '-3 * v
                            '                            HiLimit = HiLimit_Short '-200 * mV
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(DUTPin).Value(Site) > HiLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_short") = logicTrue
                                End If
                            Next Site
                        End If
                        LL.FlowTestLimit ResultVal:=MeasReault.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        '                        TheExec.Flow.TestLimit ResultVal:=MeasReault.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone

                        'Judge failed open or failed short for tlForceNone
                        For Each Site In TheExec.Sites
                            If MeasReault.Pins(DUTPin).Value(Site) < LowLimit Then
                                TheExec.Sites.Item(Site).FlagState("F_open") = logicTrue
                            End If
                        Next Site
                    Else
                        If ContiLog = ContiDataLog.eOPEN Then
                            HiLimit = LL.GetHiLimit(DC_GetTestName)
                            LowLimit = LL.GetLowLimit(DC_GetTestName)
                            '                            LowLimit = LowLimit_Open '-200 * mV
                            '                            HiLimit = HiLimit_Open '1.2 * v
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(DUTPin).Value(Site) > HiLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_open") = logicTrue
                                End If
                            Next Site
                        ElseIf ContiLog = ContiDataLog.eSHORT Then
                            HiLimit = LL.GetHiLimit(DC_GetTestName)
                            LowLimit = LL.GetLowLimit(DC_GetTestName)
                            '                            LowLimit = LowLimit_Short '200 * mV
                            '                            HiLimit = HiLimit_Short ' 1.2 * v
                            'Judge failed open or failed short for tlForceNone
                            For Each Site In TheExec.Sites
                                If MeasReault.Pins(DUTPin).Value(Site) < LowLimit Then
                                    TheExec.Sites.Item(Site).FlagState("F_short") = logicTrue
                                End If
                            Next Site
                        End If
                        LL.FlowTestLimit ResultVal:=MeasReault.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        '                        TheExec.Flow.TestLimit ResultVal:=MeasReault.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=DC_GetTestName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone

                    End If
                End If

            End If

        Next DUTPin

    Next ContiLog

    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Public Function Conti_WalkingZ(patset As Pattern, digital_pins As PinList)
    On Error GoTo ErrHandler
    TheHdw.PPMU.Pins(digital_pins).Disconnect

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.005

    '' float pin before fuction run
    If (digital_pins <> "") Then
        TheHdw.Patterns(patset).Load
        TheHdw.Patterns(patset).test pfAlways, 0
    End If

    DebugPrintFunc patset.Value

    Exit Function

ErrHandler:
    TheExec.AddOutput "Error in DC_Conti_pattern"
End Function

'Public Function IO_Continuity_IV_Curve(digital_pins As PinList, force_i_S As Double, force_i_E As Double, force_i_Step As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, Optional Separate_limit As Boolean = False, Optional independt_meas As Boolean) As Long
'
'    Dim DUTPin    As Variant
'    Dim Pins() As String, Pin_Cnt As Long
'    Dim i         As Long
'    Dim force_i   As Double
'    Dim PPMUMeasure As New PinListData
'    Dim Site      As SiteVariant
'
'    On Error GoTo ErrHandler
'    '    thehdw.DCVS.pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff 'chihome
'
'    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
'
'    For Each DUTPin In Pins
'        PPMUMeasure.AddPin (DUTPin)
'    Next DUTPin
'
'    For force_i = force_i_S To force_i_E Step (force_i_E - force_i_S) / force_i_Step
'
'
'        Dim PinGroup As IPinListData
'        Dim Power_conti_volt As New PinListData
'        Dim PPMUMeas_HexVs As New PinListData
'        Dim TName As String
'
'
'        TheHdw.Digital.ApplyLevelsTiming True, True, True
'        'DisconnectVDDCA 'SEC DRAM
'        TheHdw.Wait 0.001
'
'        '*************   Test digital channel continuity  ******************************
'        ''''Disconnect os_pins Pin Electronics from pins in order to connect PPMU's''''
'        'TheHdw.digital.Pins("Non_Conti_IO").Disconnect
'        TheHdw.Digital.Pins(digital_pins).Disconnect
'
'        '''''' Connect all os_pins to ppmu and ppmu force 0v for each one
'        If independt_meas = False Then
'            With TheHdw.PPMU.Pins(digital_pins)
'                .ForceV 0#
'                .Connect
'                .Gate = tlOn
'            End With
'        End If
'
'
'
'        '    TheExec.DataManager.DecomposePinList digital_pins, Pins(), pin_cnt
'
'        For Each DUTPin In Pins
'            '        PPMUMeasure.AddPin (DUTPin)
'
'            If independt_meas = False Then
'                With TheHdw.PPMU.Pins(DUTPin)
'                    ''            .ClampVHi = 1.2
'                    ''            .ClampVLo = -1
'                    .ForceI (force_i)
'                End With
'            Else
'                With TheHdw.PPMU.Pins(DUTPin)
'                    .Connect
'                    .ForceI (force_i)
'                    .Gate = tlOn
'                End With
'            End If
'
'
'            TheHdw.Wait 0.005
'
'            DebugPrintFunc_PPMU CStr(DUTPin)
'            PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)
'
'
'            'offline mode simulation
'            If TheExec.TesterMode = testModeOffline Then
'                For Each Site In TheExec.Sites
'                    If LCase(TheExec.DataManager.InstanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = -0.5 + Rnd() * 0.1
'                    If LCase(TheExec.DataManager.InstanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = 0.5 + Rnd() * 0.1
'                Next Site
'            End If
'
'
'            'recover measure dut pin to 0V before next pin
'            If independt_meas = False Then
'                TheHdw.PPMU.Pins(DUTPin).ForceV 0
'            Else
'                TheHdw.PPMU.Pins(DUTPin).ForceI 0
'                TheHdw.PPMU.Pins(DUTPin).Gate = tlOff
'                TheHdw.PPMU.Pins(DUTPin).Disconnect
'            End If
'            'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
'
'        Next DUTPin
'
'
'        For Each DUTPin In Pins
'            If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop1
'            TName = "Conti1_" & CStr(DUTPin)
'            If TestLimitMode = tlForceFlow Then
'                TheExec.Flow.TestLimit ResultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=TName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
'            Else: TestLimitMode = tlForceNone
'                TheExec.Flow.TestLimit ResultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=TName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
'            End If
'
'
'loop1:
'
'        Next DUTPin
'
'        If Separate_limit = True Then
'            For Each DUTPin In Pins
'                If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop2
'                TName = "Conti2_" & CStr(DUTPin)
'                If TestLimitMode = tlForceFlow Then
'                    TheExec.Flow.TestLimit ResultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=TName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
'                Else: TestLimitMode = tlForceNone
'                    TheExec.Flow.TestLimit ResultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", TName:=TName, forceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
'                End If
'loop2:
'
'            Next DUTPin
'        End If
'
'        'Disconnect PPMU from digital channels
'        With TheHdw.PPMU.Pins(digital_pins)
'            .Gate = tlOff
'            .Disconnect
'        End With
'
'        TheHdw.Digital.Pins(digital_pins).Connect
'
'    Next force_i
'
'
'    Exit Function
'ErrHandler:
'    TheExec.AddOutput "Error in Continuity"
'    If AbortTest Then Exit Function Else Resume Next
'End Function



