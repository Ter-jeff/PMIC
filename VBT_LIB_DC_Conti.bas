Attribute VB_Name = "VBT_LIB_DC_Conti"
Option Explicit

Public Function PPMU_Continuity(digital_pins As PinList, force_i As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, Optional PN_Disconnect As Boolean = False, Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short", Optional connect_all_pins As PinList) As Long

    Dim PPMUMeasure As New PinListData
    Dim PinGroup As IPinListData
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim Power_conti_volt As New PinListData
    Dim PPMUMeas_HexVs As New PinListData
    Dim i As Long
    Dim Tname As String
    Dim site As Variant
    Dim PinStr As String
    '////////////////////////////////////////////////
    Dim FlowLimitObj As IFlowLimitsInfo
    Dim Lolimit_new() As Double
    Dim HiLimit_new() As Double
    Dim Lolimit_str() As String
    Dim Hilimit_str() As String
    
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
    '////////////////////////////////////////////////
    On Error GoTo errHandler
    
    If Flag_Open Like "" Then Flag_Open = "F_open"
    If Flag_Short Like "" Then Flag_Short = "F_short"
    
    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    'If LCase(GetInstrument(Pins(0), 0)) = "hsd-u" Then '''remove the judgement, there will be error if the 1st pin is n/c.
        TheHdw.Digital.Pins(digital_pins).Disconnect
    'End If
    
    If connect_all_pins <> "" Then
        TheHdw.Digital.Pins(connect_all_pins).Disconnect
        With TheHdw.PPMU.Pins(connect_all_pins)
            .ForceV 0#
            .Connect
            .Gate = tlOn
        End With
    End If
    '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
    With TheHdw.PPMU.Pins(digital_pins)
        .ForceV 0#
        .Connect
        .Gate = tlOn
    End With

    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)

        With TheHdw.PPMU.Pins(DUTPin)
            .ForceI (force_i)
        End With
        
        TheHdw.Wait 0.005
        
        If PN_Disconnect = False Then
            DebugPrintFunc_PPMU CStr(DUTPin)
            PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)    'normal measure
        Else
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////
            If DicDiffPairs.Exists(LCase(CStr(DUTPin))) Then    '<--------------------Add Line #1
                 PinStr = DicDiffPairs(LCase(CStr(DUTPin))) '<--------------------Add Line #2
                 TheHdw.PPMU.Pins(PinStr).Gate = tlOff
                 DebugPrintFunc_PPMU CStr(DUTPin)
                 PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 32)
                 TheHdw.PPMU.Pins(PinStr).Gate = tlOn   'recover
            Else
                  DebugPrintFunc_PPMU CStr(DUTPin)
                  PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 32)
            End If '<--------------------Add Line #3

        End If
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                If LCase(TheExec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(site) = -0.5 ' Use fixed value for easily checking offline mode.
                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(site) = 0.5 ' Use fixed value for easily checking offline mode.
            Next site
        End If
        'recover measure dut Pin to 0V before next Pin
        TheHdw.PPMU.Pins(DUTPin).ForceV (0) 'correct it to force v, not force i.
    Next DUTPin
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        Dim PinSequence As Integer
        PinSequence = 0
        For Each DUTPin In Pins
        
            If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then 'if N/C jump next Pin
            
                Tname = "Conti1_" & CStr(DUTPin)
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    '///////////////////////////////////////////////////////////////////////////////////////
                    'Judge failed open or failed short for tlForceFlow
                    For Each site In TheExec.sites
                        If PPMUMeasure.Pins(DUTPin).Value(site) > HiLimit_new(PinSequence) Then
                            If HiLimit_new(PinSequence) > 0 Then
                                TheExec.sites.Item(site).FlagState(Flag_Open) = logicTrue
                            Else
                                TheExec.sites.Item(site).FlagState(Flag_Short) = logicTrue
                            End If
                            
                        ElseIf PPMUMeasure.Pins(DUTPin).Value(site) < Lolimit_new(PinSequence) Then
                            If Lolimit_new(PinSequence) > 0 Then
                                TheExec.sites.Item(site).FlagState(Flag_Short) = logicTrue
                            Else
                                TheExec.sites.Item(site).FlagState(Flag_Open) = logicTrue
                            End If
                                
                        End If
                    Next site
                ElseIf TestLimitMode = tlForceNone Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                    '///////////////////////////////////////////////////////////////////////////////////////
                    'Judge failed open or failed short for tlForceNone
                    For Each site In TheExec.sites
                        If PPMUMeasure.Pins(DUTPin).Value(site) > HiLimit Then
                            If HiLimit > 0 Then
                                TheExec.sites.Item(site).FlagState(Flag_Open) = logicTrue
                            Else
                                TheExec.sites.Item(site).FlagState(Flag_Short) = logicTrue
                            End If
                            
                        ElseIf PPMUMeasure.Pins(DUTPin).Value(site) < LowLimit Then
                            If LowLimit > 0 Then
                                TheExec.sites.Item(site).FlagState(Flag_Short) = logicTrue
                            Else
                                TheExec.sites.Item(site).FlagState(Flag_Open) = logicTrue
                            End If
                                
                        End If
                    Next site
                    '///////////////////////////////////////////////////////////////////////////////////////
                End If
            End If

            PinSequence = PinSequence + 1
        Next DUTPin
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
     'Disconnect PPMU from digital channels
    With TheHdw.PPMU.Pins(digital_pins)
        .Gate = tlOff
        .Disconnect
    End With
    If LCase(GetInstrument(Pins(0), 0)) = "hsd-u" Then
        TheHdw.Digital.Pins(digital_pins).Connect
    End If
    If connect_all_pins <> "" Then
        With TheHdw.PPMU.Pins(connect_all_pins)
            '.ForceV 0#
            .Gate = tlOff
            .Disconnect
        End With
    End If
    
    DebugPrintFunc ""
    
    Exit Function
errHandler:
        TheExec.AddOutput "Error in Continuity"
        If AbortTest Then Exit Function Else Resume Next
End Function
Public Function UVI80_Continuity(digital_pins As PinList, force_i As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, Optional PN_Disconnect As Boolean = False, _
                                        Optional Separate_limit As Boolean = False, Optional LowLimit2 As Double, Optional HiLimit2 As Double, Optional connect_all_pins As PinList) As Long

    Dim PPMUMeasure As New PinListData
    Dim PinGroup As IPinListData
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim Power_conti_volt As New PinListData
    Dim PPMUMeas_HexVs As New PinListData
    Dim i As Long
    Dim Tname As String
    Dim site As Variant

    On Error GoTo errHandler
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.001

    '*************   Test digital channel continuity  ******************************
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    'TheHdw.digital.Pins("Non_Conti_IO").Disconnect
''    TheHdw.Digital.Pins(digital_pins).Disconnect
''
''    '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
''    With TheHdw.PPMU.Pins(digital_pins)
''        .ForceV 0#
''        .Gate = tlOn
''        .Connect
''    End With
    If connect_all_pins <> "" Then
        TheHdw.Digital.Pins(connect_all_pins).Disconnect
        With TheHdw.PPMU.Pins(connect_all_pins)
            .ForceV 0#
            .Connect
            .Gate = tlOn
        End With
    End If

    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt

    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)

            'If LCase(DUTPin) Like "*uvi80*" And TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then
            If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then
                'bypass alarm
               
                TheHdw.DCVI.Pins(DUTPin).Alarm(tlDCVIAlarmOverRange) = tlAlarmOff
                TheHdw.DCVI.Pins(DUTPin).Alarm(tlDCVIAlarmMode) = tlAlarmOff
                TheHdw.DCVI.Pins(DUTPin).Alarm(tlDCVIAlarmCapture) = tlAlarmOff
                
                With TheHdw.DCVI.Pins(DUTPin)
'''                'The following VBT syntax programs a DCVI in high impedance mode:
'''                ' Only required if force was previously connected
'''                TheHdw.DCVI.Pins("MyPin").Disconnect tlDCVIConnectDefault
'''                ' Program the DCVI mapped to MyPin to high impedance mode
'''                TheHdw.DCVI.Pins("MyPin").mode = tlDCVIModeHighImpedance
'''                ' Connect only the sense to use with high impedance mode
'''                TheHdw.DCVI.Pins("MyPin").Connect tlDCVIConnectHighSense
 
                    .Disconnect tlDCVIConnectDefault
                    .mode = tlDCVIModeCurrent
                   ' .CurrentRange.Autorange = True
                    .Voltage = 1    'original setting is clamp voltage=0 which is not correct 2018/01/02
                    .VoltageRange.Value = -2
                    .current = force_i
                    '.CurrentRange.Value = force_i * 2
                    .CurrentRange.Autorange = True
                    .Connect tlDCVIConnectDefault
                   ' thehdw.wait (10 * ms)
                    .Gate = True
                End With
                
                'measure
                TheHdw.DCVI.Pins(DUTPin).Meter.mode = tlDCVIMeterVoltage
                TheHdw.Wait (0.005) ' TTR from 0.5 to 0.03 'to 5ms 180430
                PPMUMeasure.Pins(DUTPin) = TheHdw.DCVI.Pins(DUTPin).Meter.Read(tlStrobe, 1, 1000000)
                
                'reset
                With TheHdw.DCVI.Pins(DUTPin)
                     .current = 0#
                    .Gate = False
                    .Disconnect tlDCVIConnectDefault
                End With
                

            End If

        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                If LCase(TheExec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(site) = -0.5 + Rnd() * 0.1
                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(site) = 0.5 + Rnd() * 0.1
            Next site
        End If


        'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome

    Next DUTPin
    
    If connect_all_pins <> "" Then
        With TheHdw.PPMU.Pins(connect_all_pins)
            '.ForceV 0#
            .Gate = tlOff
            .Disconnect
        End With
    End If

        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then
            Tname = "Conti1_" & CStr(DUTPin)
            If TestLimitMode = tlForceFlow Then
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            Else: TestLimitMode = tlForceNone
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
            
                'Judge failed open or failed short for tlForceNone
                For Each site In TheExec.sites
                    If PPMUMeasure.Pins(DUTPin).Value(site) > HiLimit Then
                        If HiLimit > 0 Then
                            TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                        Else
                            TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                        End If
                    ElseIf PPMUMeasure.Pins(DUTPin).Value(site) < LowLimit Then
                        If LowLimit > 0 Then
                            TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                        Else
                            TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                        End If
                    End If
                Next site
            
            End If
            


             'reset alarm
            TheHdw.DCVI.Pins(DUTPin).Alarm(tlDCVIAlarmOverRange) = tlAlarmForceFail
            TheHdw.DCVI.Pins(DUTPin).Alarm(tlDCVIAlarmMode) = tlAlarmForceFail
            TheHdw.DCVI.Pins(DUTPin).Alarm(tlDCVIAlarmCapture) = tlAlarmForceFail
           End If
           
        Next DUTPin

        If Separate_limit = True Then
            For Each DUTPin In Pins
                If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop2
                Tname = "Conti2_" & CStr(DUTPin)
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Else: TestLimitMode = tlForceNone
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit2, hiVal:=HiLimit2, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                
                    'Judge failed open or failed short for tlForceNone
                    For Each site In TheExec.sites
                        If PPMUMeasure.Pins(DUTPin).Value(site) > HiLimit Then
                            If HiLimit > 0 Then
                                TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                            Else
                                TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                            End If
                        ElseIf PPMUMeasure.Pins(DUTPin).Value(site) < LowLimit Then
                            If LowLimit > 0 Then
                                    TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                                Else
                                    TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                                End If
                        End If
                    Next site
                
                End If
loop2:

            Next DUTPin
        End If

    DebugPrintFunc ""

    Exit Function
errHandler:
        TheExec.AddOutput "Error in Continuity"
        If AbortTest Then Exit Function Else Resume Next
End Function







Public Function p2p_short_Power(AllHexVsDig As PinList, AllHexVs As PinList, UVSPinsHigh As PinList, UVSPinsLow As PinList, _
                                HexvsForceI As Double, UvsHigh_ForceI As Double, UvsLow_ForceI As Double, HexLoLimit As Double, HexHiLimit As Double, LoLimit As Double, HiLimit As Double _
                                , Optional TestLimitMode As tlLimitForceResults = tlForceNone, Optional VClampHi As Double = 6.5, Optional VClampLo As Double = -1.6) As Long
'Hexvs use ppmu to measure, current <0, Uvs use DCVS sink current to measure, current >0

    Dim HexVSMeasure As New PinListData
    Dim UvsOutMeasure_Low As New PinListData
    Dim UvsOutMeasure_Mid As New PinListData
    Dim UvsOutMeasure_High As New PinListData
    Dim HexvsOutMeasure_Low As New PinListData
    Dim HexvsOutMeasure_Mid As New PinListData
    Dim HexvsOutMeasure_High As New PinListData
    Dim MDA_UVS As New PinListData
    Dim MDA_HEXVS As New PinListData
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim p As Variant
''    Dim TempString As String
''    Dim PowerSequence As Double

    On Error GoTo errHandler
    'inital all power Pins alarm bin out at powershort beginning
    With TheHdw.DCVS.Pins(AllHexVs & "," & UVSPinsHigh & "," & UVSPinsLow)
         .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin
         .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail
    End With
    

    With TheHdw.DCVS.Pins(AllHexVs)
        .Voltage.Main = 0#
        .Gate = False
        .Disconnect
    End With

    TheHdw.Digital.Pins(AllHexVsDig).Disconnect
    With TheHdw.PPMU.Pins(AllHexVsDig) ' need sink current with ppmu for hexVS Pins
        .ClampVHi = VClampHi
        .ClampVLo = VClampLo
        .Connect
        .ForceV (0)
        .Gate = tlOn
    End With
    
    With TheHdw.DCVS.Pins(UVSPinsHigh & "," & UVSPinsLow)
        .Alarm(tlDCVSAlarmSinkFoldCurrentLimitTimeout) = tlAlarmOff      ''Turn off  alarm for FIMV only
        .Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout) = tlAlarmOff    ''Turn off  alarm for FIMV only
        .Alarm(tlDCVSAlarmSourceFoldCurrentLimitLevel) = tlAlarmOff      ''Turn off  alarm for FIMV only
'========================Modified by Carter for Central Compile Review, 20181126, Variable not defined
'        .Alarm(tlDCVSAlarmSinkFoldCurrentLimitLevel) = tlAlarmOff        ''Turn off  alarm for FIMV only
'====================================================================================================
        .mode = tlDCVSModeVoltage
        .Voltage.Main = 0
        .Meter.mode = tlDCVSMeterVoltage
        .Connect
        .Gate = True
    End With
    
    Call Pre_PowerUp
    
    'PPMU Measure
    If HexvsForceI > 0.1 Or HexvsForceI < -0.1 Then 'check forceI
        TheExec.Datalog.WriteComment "print: higher HexvsForceI" & HexvsForceI & " , exit"
        Exit Function
    End If

    PPMU_FIMV_P2P_short AllHexVsDig, HexvsForceI, 0.1, 10, HexLoLimit, HexHiLimit, TestLimitMode   'fiji -0.05
    If TheExec.sites.Active.Count = 0 Then Exit Function

    'UVS256 Measure
    If UvsHigh_ForceI > 0.1 Or UvsHigh_ForceI < -0.1 Then   'check forceI
        TheExec.Datalog.WriteComment "print: higher UvsHigh_ForceI" & UvsHigh_ForceI & " , exit"
        Exit Function
    End If

    If UvsLow_ForceI > 0.1 Or UvsLow_ForceI < -0.1 Then
        TheExec.Datalog.WriteComment "print: higher UvsLow_ForceI" & UvsLow_ForceI & " , exit"
        Exit Function
    End If

    TheHdw.Digital.Pins("Pins_Monitor").Disconnect 'disconnect power monitor Pins
    
    DCVS_P2P_short_FIMV DCVS_UVS256, UVSPinsHigh, 1, UvsHigh_ForceI * 2, UvsHigh_ForceI * 2, UvsHigh_ForceI, 10, LoLimit, HiLimit, TestLimitMode, 0.1      'fiji 0.05
    
    If TheExec.sites.Active.Count = 0 Then Exit Function

    DCVS_P2P_short_FIMV DCVS_UVS256, UVSPinsLow, 1, UvsLow_ForceI * 2, UvsLow_ForceI * 2, UvsLow_ForceI, 10, LoLimit, HiLimit, tlForceNone, 0.1 'fiji 0.01
    If TheExec.sites.Active.Count = 0 Then Exit Function

''    With TheHdw.PPMU.Pins(AllHexVsDig) ' need sink current with ppmu for hexVS Pins
''        .ForceV (0)
''        .Gate = tlOff
''        .Disconnect
''    End With

    With TheHdw.DCVS.Pins(UVSPinsHigh & "," & UVSPinsLow)  ''Turn on alarm for FIMV
        .Alarm(tlDCVSAlarmAll) = tlAlarmForceBin           ''Turn on alarm for FIMV
        .Alarm(tlDCVSAlarmAll) = tlAlarmForceFail          ''Turn on alarm for FIMV
   End With



Exit Function

errHandler:

    TheExec.Datalog.WriteComment " P-to-P power short happens Error."
    If AbortTest Then Exit Function Else Resume Next

End Function

'
 Public Function Pre_PowerUp(Optional Power As String = "All_Power", Optional DebugFlag As Boolean = False)
    Dim CurrentChans As String
    Dim site As Variant
    Dim WaitConnectTime As Double
    Dim Pins() As String, Pin_Cnt As Long
    Dim powerPin As Variant
    Dim PowerName As String
    Dim TempString As String
    Dim Vmain As Double
    Dim Irange As Double
    Dim step As Integer
    Dim RiseTime As Double
    Dim PowerSequence As Double
    Dim i As Long
    
    
    On Error GoTo errHandler
    
    'DebugFlag = True

    CurrentChans = TheExec.CurrentChanMap 'obtain FT or CP channel map information

''UVS265 HighCurrent Mode Per Channel Flod Limits 0.8A else 0.2
''Typical function will import Level Sheet Setting
''Goal Level set to Spec

    
    TheExec.DataManager.DecomposePinList Power, Pins(), Pin_Cnt
    
    For i = 0 To Pin_Cnt - 1
        TempString = ""
        PowerName = CStr(Pins(i))

        TempString = PowerName & "_PowerSequence_GLB"
        PowerSequence = TheExec.specs.Globals(TempString).ContextValue
        If PowerSequence = 99 Then
            TheHdw.DCVS.Pins(Pins(i)).Disconnect
            If DebugFlag = True Then    'debugprint
                TheExec.Datalog.WriteComment "print: Pin " & PowerName & " is 'NA' Pin, disconnect, PowerSequence " & PowerSequence
            End If
        End If
    Next i
    

    
    Exit Function
    
errHandler:
        TheExec.Datalog.WriteComment "power up function is error "
        If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function Conti_WalkingZ(patset As Pattern, digital_pins As PinList, PN_Disconnect As String)
    On Error GoTo errHandler
    Dim FailPinsGroup() As String
    Dim Pin As Variant
    Dim site As Variant
    TheHdw.PPMU.Pins(digital_pins).Disconnect
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.001
    TheHdw.Digital.Pins(digital_pins).InitState = chInitLo
    If PN_Disconnect <> "" Then TheHdw.Digital.Pins(PN_Disconnect).Disconnect
    TheHdw.Wait 0.002
'' float Pin before fuction run
    If (digital_pins <> "") Then
'        Thehdw.Patterns(PatSet).Load
        TheHdw.Patterns(patset).Test pfAlways, 0
    End If
   
    Dim WalkingZFailPins As String
        
    For Each site In TheExec.sites
        
        FailPinsGroup = TheHdw.Digital.FailedPins(site)
        For Each Pin In FailPinsGroup
            TheExec.Flow.TestLimit 1, 0, 0, PinName:=Pin
        Next Pin
'        TheExec.Flow.TestLimit WalkingZ_Faillist, 1, 1
        
'        WalkingZFailPins = Join(FailPinsGroup, ",")
'        Theexec.Datalog.WriteComment "Site" & site & ", Walking Fail Pins : " & UCase(WalkingZFailPins)
    Next site
    
    DebugPrintFunc patset.Value

Exit Function

errHandler:
        TheExec.AddOutput "Error in DC_Conti_pattern"
                If AbortTest Then Exit Function Else Resume Next
End Function


Public Function PowerSensePins_continuity(PowerPins As PinList, LowLimit As Double, HiLimit As Double, force_v As Double, SensePin_additionName As String) As Long
    On Error GoTo errHandler
       
    Dim ResultPower As New PinListData
    
    Dim power_Pins_array() As String
    Dim digital_pins_array() As String
    Dim Ts As Variant
    Dim PowerSeqNum As Integer
    Dim DigitalSeqNum As Integer
    Dim seqnum As Integer
    Dim seqnum_check As Integer
    Dim Pin_Cnt As Long
    Dim power_sense As String
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheExec.DataManager.DecomposePinList PowerPins, power_Pins_array(), Pin_Cnt ''160725 Turn string to Pinlist
                
    For Each Ts In power_Pins_array
        TheHdw.DCVS.Pins(Ts).Voltage.Main = force_v
        
        If SensePin_additionName = "" Then
            SensePin_additionName = "Sense"
        End If
        
        power_sense = ""
        power_sense = Ts + "_" + SensePin_additionName
        TheHdw.Digital.Pins(power_sense).Disconnect ''160725 Digital.disconnect first
        With TheHdw.PPMU.Pins(power_sense)
            .ForceI 0
            .Gate = tlOn
            .Connect
        End With
        TheHdw.Wait 0.05
        ResultPower.AddPin (power_sense)
        ResultPower.Pins(power_sense) = TheHdw.PPMU.Pins(power_sense).Read
        TheHdw.DCVS.Pins(Ts).Voltage.Main = 0
        TheHdw.PPMU.Pins(power_sense).Disconnect
        TheHdw.PPMU.Pins(power_sense).Gate = tlOff
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                ResultPower.Pins(power_sense).Value(site) = 0.2 - Rnd() * 0.01
            Next site
        End If
    Next Ts
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
    TheExec.Flow.TestLimit resultVal:=ResultPower, lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitVolt, ForceVal:=force_v, ForceUnit:=unitVolt

    DebugPrintFunc ""                                        ' add for Miner 20151103


    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function








Public Function PPMU_Continuity_PN_Disconnect_IV_Curve(digital_pins As PinList, force_i_S As Double, force_i_E As Double, force_i_Step As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, Optional Separate_limit As Boolean = False) As Long
 
    Dim PPMUMeasure As New PinListData
'    Dim PinGroup As IPinListData
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
'    Dim Power_conti_volt As New PinListData
'    Dim PPMUMeas_HexVs As New PinListData
    Dim i As Long
    Dim Tname As String
    Dim PinStr As String
    Dim force_i As Double
    Dim site As Variant
    
    On Error GoTo errHandler
'    thehdw.DCVS.Pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff 'chihome
    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
    
    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)
    Next DUTPin
    
For force_i = force_i_S To force_i_E Step (force_i_E - force_i_S) / force_i_Step

   
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    'DisconnectVDDCA 'SEC DRAM
    TheHdw.Wait 0.001
    
    '*************   Test digital channel continuity  ******************************
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    'TheHdw.digital.Pins("Non_Conti_IO").Disconnect
    TheHdw.Digital.Pins(digital_pins).Disconnect
        
    '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
    With TheHdw.PPMU.Pins(digital_pins)
        .ForceV 0#
        .Connect
        .Gate = tlOn
    End With
    
'    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt

    For Each DUTPin In Pins
'        PPMUMeasure.AddPin (DUTPin)

        With TheHdw.PPMU.Pins(DUTPin)
''            .ClampVHi = 1.2
''            .ClampVLo = -1
            .ForceI (force_i)
        End With

        TheHdw.Wait 0.005
        
        'disconnect differential pair
        If LCase(CStr(DUTPin)) Like "*_n" Then PinStr = Replace(LCase(CStr(DUTPin)), "_n", "_p")
        If LCase(CStr(DUTPin)) Like "*_p" Then PinStr = Replace(LCase(CStr(DUTPin)), "_p", "_n")
        
        'for _pll -> _nll, _ncie...
        PinStr = Replace(LCase(PinStr), "_nll", "_pll")
        PinStr = Replace(LCase(PinStr), "_ncie", "_pcie")

        TheHdw.PPMU.Pins(PinStr).Disconnect
        DebugPrintFunc_PPMU CStr(DUTPin)
        PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)
           
        TheHdw.PPMU.Pins(PinStr).Connect
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                If LCase(TheExec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(site) = -0.5 + Rnd() * 0.1
                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(site) = 0.5 + Rnd() * 0.1
            Next site
        End If
        
        
        'recover measure dut Pin to 0V before next Pin
        TheHdw.PPMU.Pins(DUTPin).ForceV 0

        'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome

    Next DUTPin
    
    
        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop1
            Tname = "Conti1_" & CStr(DUTPin)
            If TestLimitMode = tlForceFlow Then
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            Else: TestLimitMode = tlForceNone
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
            End If
            
        
loop1:
            
        Next DUTPin
        
        If Separate_limit = True Then
            For Each DUTPin In Pins
                If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop2
                Tname = "Conti2_" & CStr(DUTPin)
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Else: TestLimitMode = tlForceNone
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                End If
loop2:
                
            Next DUTPin
        End If
    
''      'initialize ppmu to suitable clamp
''    With TheHdw.PPMU.Pins("Pins_1p0v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p1v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p8v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''    End With
    
     'Disconnect PPMU from digital channels
    With TheHdw.PPMU.Pins(digital_pins)
        .Gate = tlOff
        .Disconnect
    End With
                      
    TheHdw.Digital.Pins(digital_pins).Connect
    
Next force_i

    Exit Function
errHandler:
        TheExec.AddOutput "Error in Continuity"
        If AbortTest Then Exit Function Else Resume Next
End Function


Public Function PPMU_Continuity_PN_Disconnect(digital_pins As PinList, force_i As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, Optional Separate_limit As Boolean = False) As Long
 
    Dim PPMUMeasure As New PinListData
    Dim PinGroup As IPinListData
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim Power_conti_volt As New PinListData
    Dim PPMUMeas_HexVs As New PinListData
    Dim i As Long
    Dim Tname As String
    Dim PinStr As String
    Dim site As Variant
    
    On Error GoTo errHandler
'    thehdw.DCVS.Pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff 'chihome
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    'DisconnectVDDCA 'SEC DRAM
    TheHdw.Wait 0.001
    
    '*************   Test digital channel continuity  ******************************
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    'TheHdw.digital.Pins("Non_Conti_IO").Disconnect
    TheHdw.Digital.Pins(digital_pins).Disconnect
        
    '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
    With TheHdw.PPMU.Pins(digital_pins)
        .ForceV 0#
        .Connect
        .Gate = tlOn
    End With
    
    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt

    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)

        With TheHdw.PPMU.Pins(DUTPin)
''            .ClampVHi = 1.2
''            .ClampVLo = -1
            .ForceI (force_i)
        End With

        TheHdw.Wait 0.005
        
        'disconnect differential pair
'        If LCase(CStr(DUTPin)) Like "*_n" Then PinStr = Replace(LCase(CStr(DUTPin)), "_n", "_p")
        If LCase(CStr(DUTPin)) Like "*_n" Then PinStr = Left(CStr(DUTPin), Len(CStr(DUTPin)) - 2) & "_p"
'        If LCase(CStr(DUTPin)) Like "*_p" Then PinStr = Replace(LCase(CStr(DUTPin)), "_p", "_n")
        If LCase(CStr(DUTPin)) Like "*_p" Then PinStr = Left(CStr(DUTPin), Len(CStr(DUTPin)) - 2) & "_n"
        If LCase(CStr(DUTPin)) Like "*_dn*" Then PinStr = Replace(LCase(CStr(DUTPin)), "_dn", "_dp")
        If LCase(CStr(DUTPin)) Like "*_dp*" Then PinStr = Replace(LCase(CStr(DUTPin)), "_dp", "_dn")
        
        
        'for _pll -> _nll, _ncie...
        PinStr = Replace(LCase(PinStr), "_nll", "_pll")
        PinStr = Replace(LCase(PinStr), "_ncie", "_pcie")

        TheHdw.PPMU.Pins(PinStr).Disconnect
        
         'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                If LCase(TheExec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(site) = -0.5 + Rnd() * 0.1
                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(site) = 0.5 + Rnd() * 0.1
            Next site
            
        Else
            DebugPrintFunc_PPMU CStr(DUTPin)
            PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)
            
        End If
        
           
        TheHdw.PPMU.Pins(PinStr).Connect
        

        
        
        'recover measure dut Pin to 0V before next Pin
        TheHdw.PPMU.Pins(DUTPin).ForceV 0

        'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome

    Next DUTPin
    
    
        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop1
            Tname = "Conti1_" & CStr(DUTPin)
            If TestLimitMode = tlForceFlow Then
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            Else: TestLimitMode = tlForceNone
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
            
            
            'Judge failed open or failed short for tlForceNone
            For Each site In TheExec.sites
                If PPMUMeasure.Pins(DUTPin).Value(site) > HiLimit Then
                    If HiLimit > 0 Then
                        TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                    Else
                        TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                    End If
                    
                ElseIf PPMUMeasure.Pins(DUTPin).Value(site) < LowLimit Then
                    If LowLimit > 0 Then
                        TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                    Else
                        TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                    End If
                        
                End If
            Next site
            
            
            End If
            
        
loop1:
            
        Next DUTPin
        
        If Separate_limit = True Then
            For Each DUTPin In Pins
                If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop2
                Tname = "Conti2_" & CStr(DUTPin)
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Else: TestLimitMode = tlForceNone
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                End If
loop2:
                
            Next DUTPin
        End If
    
''      'initialize ppmu to suitable clamp
''    With TheHdw.PPMU.Pins("Pins_1p0v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p1v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p8v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''    End With
    
     'Disconnect PPMU from digital channels
    With TheHdw.PPMU.Pins(digital_pins)
        .Gate = tlOff
        .Disconnect
    End With
                      
    TheHdw.Digital.Pins(digital_pins).Connect
    
    Exit Function
errHandler:
        TheExec.AddOutput "Error in Continuity"
        If AbortTest Then Exit Function Else Resume Next
End Function




Public Function p2p_short_Power_FVMI(allpowerpins As PinList, _
                     ForceV As Double, _
                     LowLimit As Double, _
                     HiLimit As Double, _
                     TestLimitMode As tlLimitForceResults, _
                     FlowLimitForInitIRange As Boolean, _
                     digital_pins As PinList, _
                     Optional InitRange200mAPins As String, _
                     Optional InitRange20mAPins As String, _
                     Optional InitRange2mAPins As String, _
                     Optional auto_range_flag As Boolean = False) As Long

   ''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Testing Method:  Force 0.1V , measure smaller than 199ma,set clamp to 200ma, if higher than 199 ma then fail
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim PowerPins As String
    Dim p As Variant, Pin_Ary() As String, p_cnt As Long, Pin_dcvi_Ary() As String
    Dim Tname As String
    Dim TempString As String
    Dim PowerSequence As Double
    Dim site As Variant
    Dim Pin As New PinList
    Dim FoldLimit As Double
    
    Dim MaxCurr As Double
    Dim MeasRangeVal As Double
    Dim i As Integer, StepNo As Integer, j As Integer, Stop_Step As Integer
    
        
    On Error GoTo errHandler
    
    Dim FlowLimitsInfo As IFlowLimitsInfo

    Dim Val As Double
    Dim Val_Hi() As String
    Dim Val_Lo() As String

    If (FlowLimitForInitIRange = True Or TestLimitMode = tlForceFlow) Then
        Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
        FlowLimitsInfo.GetHighLimits Val_Hi
        FlowLimitsInfo.GetLowLimits Val_Lo
    End If

    
    Dim Merge_Type, Slot_Type As String
    Dim A_Slot_Type() As String
    Dim Split_Ary() As String
    Dim SattleTime As Double
    Dim WaitTime As Double
    
    Dim p_hexvs As String
    Dim p_uvs As String
    Dim p_dc07 As String
    
    Dim A_HexVS() As String
    Dim A_UVS() As String
    Dim A_DC07() As String
    
    Dim HexVS_Power_data As New PinListData
    Dim UVS_Power_data As New PinListData
    Dim DC07_Power_data As New PinListData

    Dim MaxValue As Double
    Dim getMaxValue As Double
    Dim MeasRangeList() As Double
    Dim UVI80MeasRangeList() As Double
    
    TheExec.DataManager.DecomposePinList allpowerpins, Pin_Ary, p_cnt
    
    For Each p In Pin_Ary
        If TheExec.DataManager.ChannelType(p) <> "N/C" Then PowerPins = PowerPins & "," & p
    Next p
    
    If PowerPins <> "" Then PowerPins = Right(PowerPins, Len(PowerPins) - 1)
    
    Pin_Ary = Split(PowerPins, ",")
    
    ReDim A_Slot_Type(UBound(Pin_Ary)) As String
    ReDim step_ary(UBound(Pin_Ary)) As Long
    
    If (LCase(GetInstrument(Pin_Ary(0), 0)) = "hexvs" Or LCase(GetInstrument(Pin_Ary(0), 0)) = "vhdvs") Then
        TheHdw.DCVS.Pins(PowerPins).Gate = False
        TheHdw.DCVS.Pins(allpowerpins).CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
    ElseIf (LCase(GetInstrument(Pin_Ary(0), 0)) = "dc-07") Then
        TheHdw.DCVI.Pins(PowerPins).Gate = False
        TheHdw.DCVI.Pins(allpowerpins).FoldCurrentLimit.Behavior = tlDCVIFoldCurrentLimitBehaviorGateOff
    End If
    

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered     'SEC DRAM
    
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    If digital_pins.Value <> "" Then TheHdw.Digital.Pins(digital_pins).Disconnect
    
    TheHdw.Wait 3 * ms
    WaitTime = 260 * us
    
    '==================== Auto IRange =========================
    ' Set init IRange
    For i = 0 To UBound(Pin_Ary)
        A_Slot_Type(i) = GetInstrument(Pin_Ary(i), 0)

        If (LCase(A_Slot_Type(i)) = "hexvs") Then
            p_hexvs = p_hexvs & "," & Pin_Ary(i)
            If (FlowLimitForInitIRange = True Or TestLimitMode = tlForceFlow) Then
                Val = Abs(Val_Hi(i))
            Else
                Val = HiLimit
            End If
                If Val < 0.01 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.01, 0.01  ' HexVS
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.05
                    SattleTime = 100 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 3
                    
                ElseIf Val < 0.1 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.1, 0.1    ' HexVS
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.1
                    SattleTime = 10 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 2
                    
                ElseIf Val < 1 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 1, 1    ' HexVS
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 1
                    SattleTime = 1 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 1
                    
                ElseIf Val < 15 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 15, 15   ' HexVS
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 15
                    SattleTime = 100 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 0
                    
                Else
                    'Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Max
                    Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges Val, Val
                    SattleTime = 100 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 0
                    
                End If
            
        ElseIf (LCase(A_Slot_Type(i)) = "vhdvs") Then
            p_uvs = p_uvs & "," & Pin_Ary(i)
            If (FlowLimitForInitIRange = True Or TestLimitMode = tlForceFlow) Then
                Val = Abs(Val_Hi(i))
            Else
                Val = HiLimit
            End If
                If Val < 0.000004 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.000004, 0.000004
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.000004
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.000004
                    SattleTime = 18 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 6
                    
                ElseIf Val < 0.00002 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.00002, 0.00002
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.00002
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.00002
                    SattleTime = 4 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 5
                    
                ElseIf Val < 0.0002 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.0002, 0.0002
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.0002
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.0002
                    SattleTime = 4 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 4
                    
                ElseIf Val < 0.002 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.002, 0.002
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.002
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.002
                    SattleTime = 3.5 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 3
                    
                ElseIf Val < 0.02 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.02, 0.02
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.02
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
                    SattleTime = 540 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 2
                    
                ElseIf Val < 0.04 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.04, 0.04
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.04
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.04
                    SattleTime = 260 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 1
                    
                ElseIf Val < 0.2 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.2, 0.2
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.2
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
                    SattleTime = 210 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 0
                
                Else
                    'Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Max
                    Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges Val, Val
                    SattleTime = 210 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    step_ary(i) = 0
                    
                End If

        ElseIf (LCase(A_Slot_Type(i)) = "dc-07") Then
            p_dc07 = p_dc07 & "," & Pin_Ary(i)
        End If
    Next i
    
    If (p_hexvs <> "" Or p_uvs <> "") Then
        If InitRange200mAPins <> "" Then
            TheHdw.DCVS.Pins(InitRange200mAPins).SetCurrentRanges 0.2, 0.2
            TheHdw.DCVS.Pins(InitRange200mAPins).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
        End If
        If InitRange20mAPins <> "" Then
            TheHdw.DCVS.Pins(InitRange20mAPins).SetCurrentRanges 0.02, 0.02
            TheHdw.DCVS.Pins(InitRange20mAPins).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
        End If
        If InitRange2mAPins <> "" Then
            TheHdw.DCVS.Pins(InitRange2mAPins).SetCurrentRanges 0.002, 0.002
            TheHdw.DCVS.Pins(InitRange2mAPins).CurrentLimit.Source.FoldLimit.Level.Value = 0.002
        End If
        
        '=======================================================================================================
        If p_hexvs <> "" Then
            p_hexvs = Right(p_hexvs, Len(p_hexvs) - 1)
            If auto_range_flag = True Then
                TheHdw.DCVS.Pins(p_hexvs).Voltage.Main.Value = ForceV
            End If
            TheHdw.Wait 100 * ms    'prevent for error, need to be turned
            HexVS_Power_data = TheHdw.DCVS.Pins(p_hexvs).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
        End If
        If p_uvs <> "" Then
            p_uvs = Right(p_uvs, Len(p_uvs) - 1)
            If auto_range_flag = True Then
                TheHdw.DCVS.Pins(p_uvs).Voltage.Main.Value = ForceV
                TheHdw.Wait 5 * ms
            End If
            UVS_Power_data = TheHdw.DCVS.Pins(p_uvs).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
        End If
        
        '=======================================================================================================
        TheHdw.DCVS.Pins(PowerPins).Voltage.Main.Value = 0#
        TheHdw.Wait 3 * ms
        'Start search I range
        For i = 0 To UBound(Pin_Ary)
            TheHdw.DCVS.Pins(Pin_Ary(i)).Voltage.Main.Value = ForceV
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If (LCase(A_Slot_Type(i)) = "hexvs") Then
                If TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.CurrentRange = 0.01 Then
                    TheHdw.Wait 100 * ms
                Else
                    TheHdw.Wait 30 * ms
                End If
            Else
                TheHdw.Wait 5 * ms  'align Cyprus TTR
            End If
                
            If (LCase(A_Slot_Type(i)) = "hexvs") Then
                '===============================================================================auto range
                If auto_range_flag = True Then
                    For Each site In TheExec.sites
                        StepNo = j + step_ary(i)
                        If StepNo = 6 Then j = Stop_Step
                        Val = Abs(HexVS_Power_data.Pins(Pin_Ary(i)).Value(site))
                        Select Case StepNo
                            Case 1: '15A => 1A
                                    If ((Val + (0.05 + 0.15) * 2) < 1) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 1, 1: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 1
'                                    SattleTime = 1 * ms
'                                    If SattleTime > WaitTime Then WaitTime = SattleTime
                            Case 2: '1A => 100mA
                                    If ((Val + (0.01 + 0.005) * 2) < 0.1) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.1, 0.1: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.1
'                                    SattleTime = 10 * ms
'                                    If SattleTime > WaitTime Then WaitTime = SattleTime
    '                        Case 3: '100mA => 10mA
    '                                If ((Val + 0.001*2) < 0.01) Then TheHdw.DCVS.Pins(Power_data.Pins(i)).SetCurrentRanges 0.01, 0.01: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
    '                                SattleTime = 100 * ms
    '                                If SattleTime > WaitTime Then WaitTime = SattleTime
                        End Select
                    Next site
                    Wait 0.01
                End If
                '===============================================================================
                HexVS_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
            ElseIf (LCase(A_Slot_Type(i)) = "vhdvs") Then
                '===============================================================================auto range
                If auto_range_flag = True Then
                    For Each site In TheExec.sites
                        StepNo = j + step_ary(i)
                        Val = Abs(UVS_Power_data.Pins(Pin_Ary(i)).Value(site))
    
                        Select Case StepNo
                            Case 1: '200mA => 40mA
                                    If ((Val + (0.0012 + 0.001) * 2) < 0.04) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.04, 0.04: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.04: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.04
'                                    SattleTime = 260 * us
'                                    If SattleTime > WaitTime Then WaitTime = SattleTime
                            Case 2: '40mA =>20mA
                                    If ((Val + (0.00024 + 0.0003) * 2) < 0.02) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.02, 0.02: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.02: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.02
'                                    SattleTime = 540 * us
'                                    If SattleTime > WaitTime Then WaitTime = SattleTime
                            Case 3: '20mA =>2mA
                                    If ((Val + (0.00012 + 0.0001) * 2) < 0.002) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.002, 0.002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.002
'                                    SattleTime = 3.5 * ms
'                                    If SattleTime > WaitTime Then WaitTime = SattleTime
                            Case 4: '2mA =>200uA
                                    If ((Val + (0.000012 + 0.00001) * 2) < 0.0002) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.0002, 0.0002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.0002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.0002
'                                    SattleTime = 4 * ms
'                                    If SattleTime > WaitTime Then WaitTime = SattleTime
    '                        Case 5: '200uA =>20uA
    '                                If ((Val + (0.0000012 + 0.000001) * 2) < 0.00002) Then TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.00002, 0.00002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.00002: TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.00002
    '                                SattleTime = 4 * ms
    '                                If SattleTime > WaitTime Then WaitTime = SattleTime
    '                        Case 6: '20uA =>4uA
    '                                If ((Val + (0.00000012 + 0.0000001) * 2) < 0.000004) Then thehdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.000004, 0.000004: thehdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.000004: thehdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level = 0.000004
    '                                SattleTime = 18 * ms
    '                                If SattleTime > WaitTime Then WaitTime = SattleTime
                        End Select
                    Next site
                    Wait 0.004
                End If
                '===============================================================================
                UVS_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 1)
            End If
            
            TheHdw.DCVS.Pins(Pin_Ary(i)).Voltage.Main.Value = 0#
            
            Tname = "pwr_FVMI_" & Pin_Ary(i)
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If (LCase(A_Slot_Type(i)) = "hexvs") Then
                'offline mode simulation
                If TheExec.TesterMode = testModeOffline Then
                    For Each site In TheExec.sites
                        HexVS_Power_data.Pins(Pin_Ary(i)).Value(site) = 0.01 + Rnd() * 0.0001
                    Next site
                End If
            
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=HexVS_Power_data.Pins(Pin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                ElseIf TestLimitMode = tlForceNone Then
                    TheExec.Flow.TestLimit resultVal:=HexVS_Power_data.Pins(Pin_Ary(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                End If
            ElseIf (LCase(A_Slot_Type(i)) = "vhdvs") Then
                'offline mode simulation
                If TheExec.TesterMode = testModeOffline Then
                    For Each site In TheExec.sites
                        UVS_Power_data.Pins(Pin_Ary(i)).Value(site) = 0.01 + Rnd() * 0.0001
                    Next site
                End If
                
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=UVS_Power_data.Pins(Pin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                ElseIf TestLimitMode = tlForceNone Then
                    TheExec.Flow.TestLimit resultVal:=UVS_Power_data.Pins(Pin_Ary(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                End If
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Next i

    ElseIf (p_dc07 <> "") Then
        p_dc07 = Right(p_dc07, Len(p_dc07) - 1)
'''        TheHdw.DCVI.Pins(p_dc07).Alarm(tlDCVIAlarmOverRange) = tlAlarmOff
'''        TheHdw.DCVI.Pins(p_dc07).Alarm(tlDCVIAlarmMode) = tlAlarmOff
'''        TheHdw.DCVI.Pins(p_dc07).Alarm(tlDCVIAlarmCapture) = tlAlarmOff
        Pin_dcvi_Ary = Split(p_dc07, ",")
        TheHdw.DCVI.Pins(p_dc07).SetCurrentAndRange 0.02, 0.02
        FoldLimit = TheHdw.DCVI.Pins(p_dc07).current
        
        With TheHdw.DCVI.Pins(p_dc07)
            .Disconnect tlDCVIConnectDefault
            .Gate = False
            .mode = tlDCVIModeVoltage
            .Voltage = 0
            .CurrentRange.Autorange = True
            '.CurrentRange.Value = FoldLimit
            .VoltageRange.Autorange = True
            .Connect tlDCVIConnectDefault
            .Gate = True
            .Meter.mode = tlDCVIMeterCurrent
            .Meter.CurrentRange.Value = FoldLimit
        End With
        
        TheHdw.Wait 3 * ms
        DC07_Power_data = TheHdw.DCVI.Pins(p_dc07).Meter.Read(tlStrobe, 1, , tlDCVIMeterReadingFormatAverage)
        '=========================================================================
        For i = 0 To DC07_Power_data.Pins.Count - 1
''            UVI80MeasRangeList = TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).Meter.CurrentRange.List
            TheHdw.DCVI.Pins(Pin_dcvi_Ary(i)).Voltage = ForceV
''            TheHdw.Wait 3 * ms
''            DC07_Power_data.Pins(i).Value = TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).Meter.Read(tlStrobe, 10, , tlDCVIMeterReadingFormatAverage)
            
''            For Each Site In TheExec.sites.Active
''                MeasRangeVal = GetMeasRange(Math.Abs(DC07_Power_data.Pins(i).Value), UVI80MeasRangeList)
''                If (MeasRangeVal < FoldLimit) Then
''                     TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).CurrentRange = MeasRangeVal
''                     TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).Meter.CurrentRange.Value = MeasRangeVal
''                Else
''                     TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).Current = FoldLimit
''                     TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).CurrentRange = FoldLimit
''                     TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).Meter.CurrentRange.Value = FoldLimit
''                End If
''            Next Site
            TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).CurrentRange.Autorange = True
            TheHdw.Wait 3 * ms
            DC07_Power_data.Pins(i).Value = TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).Meter.Read(tlStrobe, 10, , tlDCVIMeterReadingFormatAverage)
            TheHdw.DCVI.Pins(Pin_dcvi_Ary(i)).Voltage = 0
            'offline mode simulation
            If TheExec.TesterMode = testModeOffline Then
                For Each site In TheExec.sites
                    DC07_Power_data.Pins(Pin_dcvi_Ary(i)).Value(site) = 0.01 + Rnd() * 0.0001
                Next site
            End If
        Next i
        '=========================================================================
        For i = 0 To UBound(Pin_Ary)
            If TestLimitMode = tlForceFlow Then
                TheExec.Flow.TestLimit resultVal:=DC07_Power_data.Pins(Pin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            ElseIf TestLimitMode = tlForceNone Then
                TheExec.Flow.TestLimit resultVal:=DC07_Power_data.Pins(Pin_Ary(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
            End If
        Next i
        '=========================================================================
'''        TheHdw.DCVI.Pins(p_dc07).Alarm(tlDCVIAlarmOverRange) = tlAlarmDefault
'''        TheHdw.DCVI.Pins(p_dc07).Alarm(tlDCVIAlarmMode) = tlAlarmDefault
'''        TheHdw.DCVI.Pins(p_dc07).Alarm(tlDCVIAlarmCapture) = tlAlarmDefault
        
        TheHdw.DCVI.Pins(p_dc07).Voltage = 0
    End If
    
    TheHdw.Digital.ApplyLevelsTiming False, True, False, tlPowered       'SEC DRAM
    DebugPrintFunc ""
    
    Exit Function
 
errHandler:
 
    TheExec.Datalog.WriteComment " P-to-P power short happens Error."
    If AbortTest Then Exit Function Else Resume Next
    

End Function

Public Function PPMU_Continuity_IV_Curve(digital_pins As PinList, force_i_S As Double, force_i_E As Double, force_i_Step As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, Optional Separate_limit As Boolean = False, Optional independt_meas As Boolean) As Long
    
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim i As Long
    Dim force_i As Double
    Dim PPMUMeasure As New PinListData
    Dim site As Variant

    On Error GoTo errHandler
'    thehdw.DCVS.Pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff 'chihome

    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
    
    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)
    Next DUTPin
    
        For force_i = force_i_S To force_i_E Step (force_i_E - force_i_S) / force_i_Step
    

        Dim PinGroup As IPinListData
        Dim Power_conti_volt As New PinListData
        Dim PPMUMeas_HexVs As New PinListData
        Dim Tname As String

    
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    'DisconnectVDDCA 'SEC DRAM
    TheHdw.Wait 0.001
    
    '*************   Test digital channel continuity  ******************************
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    'TheHdw.digital.Pins("Non_Conti_IO").Disconnect
    TheHdw.Digital.Pins(digital_pins).Disconnect
        
    '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
    If independt_meas = False Then
        With TheHdw.PPMU.Pins(digital_pins)
            .ForceV 0#
            .Connect
            .Gate = tlOn
        End With
    End If

    
    
'    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt

    For Each DUTPin In Pins
'        PPMUMeasure.AddPin (DUTPin)

        If independt_meas = False Then
            With TheHdw.PPMU.Pins(DUTPin)
    ''            .ClampVHi = 1.2
    ''            .ClampVLo = -1
                .ForceI (force_i)
            End With
        Else
            With TheHdw.PPMU.Pins(DUTPin)
                .Connect
                .ForceI (force_i)
                .Gate = tlOn
            End With
        End If


        TheHdw.Wait 0.005
        
        DebugPrintFunc_PPMU CStr(DUTPin)
        PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)
           
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
'                If LCase(theexec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = -0.5 + Rnd() * 0.1
'                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = 0.5 + Rnd() * 0.1
                PPMUMeasure.Pins(DUTPin).Value(site) = 0.5 - Rnd() * 0.01
            Next site
        End If
        
        
        'recover measure dut Pin to 0V before next Pin
        If independt_meas = False Then
            TheHdw.PPMU.Pins(DUTPin).ForceV 0
        Else
            TheHdw.PPMU.Pins(DUTPin).ForceI 0
            TheHdw.PPMU.Pins(DUTPin).Gate = tlOff
            TheHdw.PPMU.Pins(DUTPin).Disconnect
        End If
        'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome

    Next DUTPin
    
    
        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop1
            Tname = "Conti1_" & CStr(DUTPin)
            If TheExec.TesterMode = testModeOffline Then
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=0.5, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
            Else
                If TestLimitMode = tlForceFlow Then
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Else: TestLimitMode = tlForceNone
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                End If
            End If
        
loop1:
            
        Next DUTPin
        
        If Separate_limit = True Then
            For Each DUTPin In Pins
                If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop2
                Tname = "Conti2_" & CStr(DUTPin)
                If TheExec.TesterMode = testModeOffline Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=0.5, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                Else
                        If TestLimitMode = tlForceFlow Then
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                        Else: TestLimitMode = tlForceNone
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        End If
                End If
loop2:
                
            Next DUTPin
        End If
    
''      'initialize ppmu to suitable clamp
''    With TheHdw.PPMU.Pins("Pins_1p0v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p1v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p8v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''    End With
    
     'Disconnect PPMU from digital channels
    With TheHdw.PPMU.Pins(digital_pins)
        .Gate = tlOff
        .Disconnect
    End With
             
    TheHdw.Digital.Pins(digital_pins).Connect
    
Next force_i
    
    
    Exit Function
errHandler:
        TheExec.AddOutput "Error in Continuity"
        If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Power_open_measurement(allpowerpins As PinList, _
                     ForceV As Double, _
                     LowLimit As Double, _
                     HiLimit As Double, _
                     TestLimitMode As tlLimitForceResults, _
                     FlowLimitForInitIRange As Boolean, _
                     digital_pins As PinList, _
                     Optional InitRange200mAPins As String, _
                     Optional InitRange20mAPins As String, _
                     Optional InitRange2mAPins As String) As Long
'used for metal wafer measurement.
   ''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Testing Method:  Force 0.1V , measure smaller than 199ma,set clamp to 200ma, if higher than 199 ma then fail
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim PowerPins As String
    Dim p As Variant, Pin_Ary() As String, p_cnt As Long
    Dim Tname As String
    Dim TempString As String
    Dim PowerSequence As Double
    Dim site As Variant
    Dim Pin As New PinList
    
    On Error GoTo errHandler
    
    Dim FlowLimitsInfo As IFlowLimitsInfo

    Dim Val As Double
    Dim Val_Hi() As String
    Dim Val_Lo() As String

'    If (FlowLimitForInitIRange = True) Then
'        Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
'        FlowLimitsInfo.GetHighLimits Val_Hi
'        FlowLimitsInfo.GetLowLimits Val_Lo
'    End If

    
    Dim Merge_Type, Slot_Type As String
    Dim A_Slot_Type() As String
    Dim Split_Ary() As String
    Dim SattleTime As Double
    Dim WaitTime As Double
    Dim p_hexvs As String
    Dim p_uvs As String
    Dim A_HexVS() As String
    Dim A_UVS() As String
    Dim HexVS_Power_data As New PinListData
    Dim UVS_Power_data As New PinListData
    
    TheExec.DataManager.DecomposePinList allpowerpins, Pin_Ary, p_cnt
    
    For Each p In Pin_Ary
        If TheExec.DataManager.ChannelType(p) <> "N/C" Then PowerPins = PowerPins & "," & p
    Next p
    
    If PowerPins <> "" Then PowerPins = Right(PowerPins, Len(PowerPins) - 1)
    
    Pin_Ary = Split(PowerPins, ",")
    
    ReDim A_Slot_Type(UBound(Pin_Ary)) As String
    ReDim step_ary(UBound(Pin_Ary)) As Long
    
    TheHdw.DCVS.Pins(PowerPins).Gate = False
    TheHdw.DCVS.Pins(allpowerpins).CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
    
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    If digital_pins.Value <> "" Then TheHdw.Digital.Pins(digital_pins).Disconnect

    WaitTime = 260 * us
    
    Dim i As Integer
      
    '==================== Auto IRange =========================
    ' Set init IRange
    For i = 0 To UBound(Pin_Ary)
        A_Slot_Type(i) = GetInstrument(Pin_Ary(i), 0)

        If (LCase(A_Slot_Type(i)) = "hexvs") Then
            p_hexvs = p_hexvs & "," & Pin_Ary(i)
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.05, 0.05    ' HexVS
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
                    SattleTime = 1 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
'                    step_ary(i) = 1
        ElseIf (LCase(A_Slot_Type(i)) = "vhdvs") Then
            p_uvs = p_uvs & "," & Pin_Ary(i)
                'Val = Abs(Val_Hi(i))
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.05, 0.05
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.05
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
                    SattleTime = 210 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
'                    step_ary(i) = 0
        End If
    Next i
    
    If InitRange200mAPins <> "" Then
        TheHdw.DCVS.Pins(InitRange200mAPins).SetCurrentRanges 0.2, 0.2
        TheHdw.DCVS.Pins(InitRange200mAPins).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
    End If
    If InitRange20mAPins <> "" Then
        TheHdw.DCVS.Pins(InitRange20mAPins).SetCurrentRanges 0.02, 0.02
        TheHdw.DCVS.Pins(InitRange20mAPins).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
    End If
    If InitRange2mAPins <> "" Then
        TheHdw.DCVS.Pins(InitRange2mAPins).SetCurrentRanges 0.002, 0.002
        TheHdw.DCVS.Pins(InitRange2mAPins).CurrentLimit.Source.FoldLimit.Level.Value = 0.002
    End If
    
    If p_hexvs <> "" Then p_hexvs = Right(p_hexvs, Len(p_hexvs) - 1)
    If p_uvs <> "" Then p_uvs = Right(p_uvs, Len(p_uvs) - 1)
'
    TheHdw.Wait WaitTime
    
'    If p_hexvs <> "" Then HexVS_Power_data = TheHdw.DCVS.Pins(p_hexvs).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
'    If p_uvs <> "" Then UVS_Power_data = TheHdw.DCVS.Pins(p_uvs).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
'    TheHdw.Wait 3 * ms
'
'
'    TheHdw.DCVS.Pins(PowerPins).Voltage.Main.Value = 0#
'    TheHdw.Wait 3 * ms
    
    'Start search I range
    For i = 0 To UBound(Pin_Ary)
    
        TheHdw.DCVS.Pins(Pin_Ary(i)).Gate = True
        TheHdw.DCVS.Pins(Pin_Ary(i)).Connect tlDCVSConnectDefault
        TheHdw.DCVS.Pins(Pin_Ary(i)).Voltage.Main.Value = ForceV
       
        WaitTime = 260 * us
        
        If (LCase(A_Slot_Type(i)) = "hexvs") Then
            If TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.CurrentRange = 0.05 Then
                TheHdw.Wait 100 * ms
            Else
                TheHdw.Wait 30 * ms
            End If
        Else
            TheHdw.Wait 5 * ms
        End If
            
            If (LCase(A_Slot_Type(i)) = "hexvs") Then
                HexVS_Power_data.AddPin (Pin_Ary(i))
                HexVS_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                
            ElseIf (LCase(A_Slot_Type(i)) = "vhdvs") Then
                UVS_Power_data.AddPin (Pin_Ary(i))
                UVS_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 1)

            End If

        TheHdw.DCVS.Pins(Pin_Ary(i)).Voltage.Main.Value = 0#

    Next i
    '===========================================================

    For i = 0 To UBound(Pin_Ary)
    
        Tname = "pwr_FVMI_" & Pin_Ary(i)
        
        If (LCase(A_Slot_Type(i)) = "hexvs") Then
        
            'offline mode simulation
            If TheExec.TesterMode = testModeOffline Then
                For Each site In TheExec.sites
                    HexVS_Power_data.Pins(Pin_Ary(i)).Value(site) = 0.01 + Rnd() * 0.0001
                Next site
            End If
        
            If TestLimitMode = tlForceFlow Then
                TheExec.Flow.TestLimit resultVal:=HexVS_Power_data.Pins(Pin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            ElseIf TestLimitMode = tlForceNone Then
                TheExec.Flow.TestLimit resultVal:=HexVS_Power_data.Pins(Pin_Ary(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
            End If
        ElseIf (LCase(A_Slot_Type(i)) = "vhdvs") Then
        
            'offline mode simulation
            If TheExec.TesterMode = testModeOffline Then
                For Each site In TheExec.sites
                    UVS_Power_data.Pins(Pin_Ary(i)).Value(site) = 0.01 + Rnd() * 0.0001
                Next site
            End If
            
            If TestLimitMode = tlForceFlow Then
                TheExec.Flow.TestLimit resultVal:=UVS_Power_data.Pins(Pin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            ElseIf TestLimitMode = tlForceNone Then
                TheExec.Flow.TestLimit resultVal:=UVS_Power_data.Pins(Pin_Ary(i)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
            End If
        End If
        
    Next i
    
    'TheHdw.Digital.ApplyLevelsTiming False, True, False, tlPowered       'SEC DRAM
    
    DebugPrintFunc ""
    
    Exit Function
 
errHandler:
 
    TheExec.Datalog.WriteComment " P-to-P power short happens Error."
    If AbortTest Then Exit Function Else Resume Next
    

End Function


Public Function PPMU_IO_measure_R(allpowerpins As PinList, digital_pins As PinList, force_i As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, Optional PN_Disconnect As Boolean = False, Optional Separate_limit As Boolean = False, Optional LowLimit2 As Double, Optional HiLimit2 As Double, Optional VClampHi As Double = 6.5, Optional VClampLo As Double = -1.6) As Long
'used for metal wafer measurement.
    Dim PPMUMeasure As New PinListData
    Dim PinGroup As IPinListData
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim Power_conti_volt As New PinListData
    Dim PPMUMeas_HexVs As New PinListData
    Dim i As Long
    Dim Tname As String
    Dim site As Variant
    Dim IOPins As New PinListData
    Dim PinStr As String
    On Error GoTo errHandler


    TheHdw.DCVS.Pins(allpowerpins).Gate = False

    TheHdw.Wait 0.001

    '*************   Test digital channel continuity  ******************************
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    TheHdw.Digital.Pins(digital_pins).Disconnect

    '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
    With TheHdw.PPMU.Pins(digital_pins)
        .ForceV 0#
        .Gate = tlOn
        .Connect
    End With

    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt

    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)

        With TheHdw.PPMU.Pins(DUTPin)
            .ClampVHi = VClampHi
            .ClampVLo = VClampLo
            .ForceI (force_i)
        End With
        
        If False Then   ' go to no addtional wait time, 5ms
            'additional wait time handling
            If LCase(DUTPin) Like "*ddr*reset_n*" Or LCase(DUTPin) Like "*ddr*cke*" Then
                TheHdw.Wait 0.1    'relax wait time for these two Pins sensitively
            Else
                TheHdw.Wait 0.005
            End If
        Else
            TheHdw.Wait 0.005
        End If
        
        If PN_Disconnect = False Then
            PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)   'normal measure
            'IOPins = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)     'normal measure
            PPMUMeasure.Pins(DUTPin) = PPMUMeasure.Pins(DUTPin).Divide(force_i)
        Else
            'disconnect differential pair, might be no longer use, use offset to evaluate.
            If LCase(CStr(DUTPin)) Like "*pcie*ref*" Then
                If LCase(CStr(DUTPin)) Like "*_n" Then PinStr = Replace(LCase(CStr(DUTPin)), "_n", "_p")
                If LCase(CStr(DUTPin)) Like "*_p" Then PinStr = Replace(LCase(CStr(DUTPin)), "_p", "_n")
            End If
            
            If LCase(CStr(DUTPin)) Like "*mipi*" Then
                If LCase(CStr(DUTPin)) Like "*_dn*" Then PinStr = Replace(LCase(CStr(DUTPin)), "_dn", "_dp")
                If LCase(CStr(DUTPin)) Like "*_dp*" Then PinStr = Replace(LCase(CStr(DUTPin)), "_dp", "_dn")
            End If
''            If LCase(CStr(DUTPin)) Like "*ddr*reset_n*" Or LCase(DUTPin) Like "*ddr*cke*" Then
''                If LCase(CStr(DUTPin)) Like "*reset_n*" Then PinStr = Replace(LCase(CStr(DUTPin)), "_reset_n", "_cke")
''                If LCase(CStr(DUTPin)) Like "*cke*" Then PinStr = Replace(LCase(CStr(DUTPin)), "_cke", "_reset_n")
''            End If
            
            'for _pll -> _nll, _ncie...
            PinStr = Replace(LCase(PinStr), "_nll", "_pll")
            PinStr = Replace(LCase(PinStr), "_ncie", "_pcie")
    
            TheHdw.PPMU.Pins(PinStr).Disconnect
            PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)
            TheHdw.PPMU.Pins(PinStr).Connect    'recover
            
        End If
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
'                If LCase(theexec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = -0.5 + Rnd() * 0.1
'                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = 0.5 + Rnd() * 0.1
                PPMUMeasure.Pins(DUTPin).Value(site) = 0.05 + Rnd() * 0.01
            Next site
        End If


        'recover measure dut Pin to 0V before next Pin
        TheHdw.PPMU.Pins(DUTPin).ForceI (0)

        'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome

    Next DUTPin


        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop1  'jump next Pin
            
            Tname = "IO_R_" & CStr(DUTPin)
            If TestLimitMode = tlForceFlow Then
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            Else: TestLimitMode = tlForceNone
                'TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceNone
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitCustom, scaletype:=scaleNone, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, customUnit:="ohm", ForceResults:=tlForceNone
            'Judge failed open or failed short for tlForceNone
            For Each site In TheExec.sites
                If PPMUMeasure.Pins(DUTPin).Value(site) > HiLimit Then
                    If HiLimit > 0 Then
                        TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                    Else
                        TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                    End If
                    
                ElseIf PPMUMeasure.Pins(DUTPin).Value(site) < LowLimit Then
                    If LowLimit > 0 Then
                        TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                    Else
                        TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                    End If
                        
                End If
            Next site
            
            
            End If
loop1:

        Next DUTPin

        If Separate_limit = True Then
            For Each DUTPin In Pins
                If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop2  'jump next Pin
                Tname = "Conti2_" & CStr(DUTPin)
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Else: TestLimitMode = tlForceNone
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit2, hiVal:=HiLimit2, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                
            'Judge failed open or failed short for tlForceNone
            For Each site In TheExec.sites
                If PPMUMeasure.Pins(DUTPin).Value(site) > HiLimit Then
                    If HiLimit > 0 Then
                        TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                    Else
                        TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                    End If
                    
                ElseIf PPMUMeasure.Pins(DUTPin).Value(site) < LowLimit Then
                    If LowLimit > 0 Then
                        TheExec.sites.Item(site).FlagState("F_short") = logicTrue
                    Else
                        TheExec.sites.Item(site).FlagState("F_open") = logicTrue
                    End If
                        
                End If
            Next site
                
                
                
                End If
loop2:

            Next DUTPin
        End If

     'Disconnect PPMU from digital channels
    With TheHdw.PPMU.Pins(digital_pins)
        .Gate = tlOff
        .Disconnect
    End With

    TheHdw.Digital.Pins(digital_pins).Connect
    
    DebugPrintFunc ""
    
    Exit Function
errHandler:
        TheExec.AddOutput "Error in Continuity"
        If AbortTest Then Exit Function Else Resume Next
End Function
Public Function GndSensePins_continuity(PowerPins As String, digital_pins As String, LowLimit As Double, HiLimit As Double, Power_Force_V As Double, ch_force_i As Double) As Long
    On Error GoTo errHandler
        
    Dim ResultPower As New PinListData
    
    Dim power_Pins_array() As String
    Dim digital_pins_array() As String
    Dim Ts As Variant
    Dim PowerSeqNum As Long
    Dim DigitalSeqNum As Long
    Dim seqnum As Integer
    Dim seqnum_check As Integer
    Dim i As Long
        
    Dim power_sense As String

    digital_pins_array = Split(digital_pins, ",")
    DigitalSeqNum = UBound(digital_pins_array) + 1
'    DigitalSeqNum = CLng(DigitalSeqNum)
   Call Trim_NC_Pin(digital_pins_array, DigitalSeqNum)
    
    If DigitalSeqNum > 0 Then
        For i = 0 To UBound(digital_pins_array)
            If i = 0 Then
                digital_pins = digital_pins_array(i)
            Else
                digital_pins = digital_pins & "," & digital_pins_array(i)
            End If
        Next i
    Else
        digital_pins = ""
    End If

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Pins(digital_pins).Disconnect  '//digital_pins
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    power_Pins_array = Split(PowerPins, ",")
    PowerSeqNum = UBound(power_Pins_array)

    digital_pins_array = Split(digital_pins, ",")
    DigitalSeqNum = UBound(digital_pins_array)

    If (PowerSeqNum = DigitalSeqNum) Then
        seqnum_check = 0
    Else
        seqnum_check = 1
    End If
    
    If digital_pins <> "" Then
        TheHdw.DCVS.Pins(PowerPins).Voltage.Main = Power_Force_V
        TheHdw.Wait 0.005
        TheHdw.PPMU.Pins(digital_pins).Connect
        TheHdw.PPMU.Pins(digital_pins).ForceI ch_force_i
        TheHdw.Wait 0.005
        ResultPower = TheHdw.PPMU(digital_pins).Read(tlPPMUReadMeasurements)
        TheHdw.PPMU.Pins(digital_pins).Gate = tlOff
        TheHdw.PPMU.Pins(digital_pins).Disconnect
        
        TheExec.Flow.TestLimit resultVal:=ResultPower, lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitVolt, ForceVal:=ch_force_i
        DebugPrintFunc ""                                        ' add for Miner 20151103
    End If

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function p2p_short_Power_FVMI_VI_Curve(allpowerpins As PinList, _
                     ForceV_S As Double, _
                     ForceV_E As Double, _
                     ForceV_Step_Count As Double, _
                     LowLimit As Double, _
                     HiLimit As Double, _
                     digital_pins As PinList, _
                     Optional InitRange200mAPins As String, _
                     Optional InitRange20mAPins As String, _
                     Optional InitRange2mAPins As String) As Long

   ''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Testing Method:  ForceV from ForceV_S to ForceV_E, stepping ForceV_Step_Count. Measure current by using  p2p_short_Power_FVMI()
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    Dim ForceV As Double
    
    'Force FlowLimitForInitIRange as False.  Not to use flow limit due to unknow stepping
    
    For ForceV = ForceV_S To ForceV_E Step (ForceV_E - ForceV_S) / ForceV_Step_Count
    
        Call p2p_short_Power_FVMI(allpowerpins, ForceV, LowLimit, HiLimit, 0, False, digital_pins, InitRange200mAPins, InitRange20mAPins, InitRange2mAPins)
    
    Next ForceV
   
    
    On Error GoTo errHandler
    
    Exit Function
 
errHandler:
 
    TheExec.Datalog.WriteComment " P-to-P power short VI Curve happens Error."
    If AbortTest Then Exit Function Else Resume Next
    

End Function
Public Function RetrieveDictionaryOfDiffPairs()    'Wherever the DcConti module, put it there
    On Error GoTo errHandler
    Dim Pins() As String, Pin_Cnt As Long, iPin As Long
    Dim DiffGroup  As String: DiffGroup = "All_DiffPairs"                            'T-Autogen will create it."
    TheExec.DataManager.DecomposePinList DiffGroup, Pins(), Pin_Cnt
    DicDiffPairs.RemoveAll
    If Pin_Cnt Mod 2 <> 0 Or Pin_Cnt < 1 Then GoTo errHandler
        For iPin = 0 To Pin_Cnt - 1 Step 2
            DicDiffPairs.Add LCase(CStr(Pins(iPin))), LCase(CStr(Pins(iPin + 1)))
            DicDiffPairs.Add LCase(CStr(Pins(iPin + 1))), LCase(CStr(Pins(iPin)))
        Next iPin
    Exit Function
errHandler:
    HandleExecIPError "RetrieveDictionaryOfDiffPairs"
End Function

Public Function PPMU_Measure_Contact_Resistance(digital_pins As PinList, force_i_S As Double, force_i_E As Double, force_i_Step As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, Optional Separate_limit As Boolean = False, Optional independt_meas As Boolean, Optional Alldigital_pins As String, Optional HiLimit_Resistance As Double) As Long
    
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim i As Long
    Dim force_i As Double
    Dim PPMUMeasure As New PinListData
    Dim site As Variant
    Dim MeasV_Force50mA As New SiteDouble: Dim MeasV_Force0mA As New SiteDouble: Dim Calculate_Contact_R As New SiteDouble
    
    Dim TestNum As Long
    
    On Error GoTo errHandler
'    thehdw.DCVS.Pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff 'chihome

    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
    
    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)
    Next DUTPin
    
        For force_i = force_i_S To force_i_E Step (force_i_E - force_i_S) / force_i_Step
    

        Dim PinGroup As IPinListData
        Dim Power_conti_volt As New PinListData
        Dim PPMUMeas_HexVs As New PinListData
        Dim Tname As String

    
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    'DisconnectVDDCA 'SEC DRAM
    TheHdw.Wait 0.001
    
    '*************   Test digital channel continuity  ******************************
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    'TheHdw.digital.Pins("Non_Conti_IO").Disconnect
    TheHdw.Digital.Pins(digital_pins).Disconnect
        
    '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
    If independt_meas = False Then
        With TheHdw.PPMU.Pins(digital_pins)
            .ForceV 0#
            .Connect
            .Gate = tlOn
        End With
    End If

    
    
'    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt

    For Each DUTPin In Pins
'        PPMUMeasure.AddPin (DUTPin)

        If independt_meas = False Then
            With TheHdw.PPMU.Pins(DUTPin)
    ''            .ClampVHi = 1.2
    ''            .ClampVLo = -1
                .ForceI (force_i)
            End With
        Else
            With TheHdw.PPMU.Pins(DUTPin)
                .Connect
                .ForceI (force_i)
                .Gate = tlOn
            End With
        End If


        TheHdw.Wait 0.005
        
        DebugPrintFunc_PPMU CStr(DUTPin)
        PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)

        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
'                If LCase(TheExec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = -0.5 + Rnd() * 0.1
'                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = 0.5 + Rnd() * 0.1
                If force_i = 0.05 Then
                    PPMUMeasure.Pins(DUTPin).Value(site) = 0.5 + Rnd() * 0.1
                Else
                    PPMUMeasure.Pins(DUTPin).Value(site) = Rnd() * 0.1
                End If
            Next site
        End If

'///////////////////////ZB add code///////////////////////////////////////
        For Each site In TheExec.sites
            If force_i = 0.05 Then
                MeasV_Force50mA = PPMUMeasure.Pins(DUTPin).Value
            Else
                MeasV_Force0mA = PPMUMeasure.Pins(DUTPin).Value
            End If
        Next site
'/////////////////////////////////////////////////////////////////////////
        
        'recover measure dut Pin to 0V before next Pin
        If independt_meas = False Then
            TheHdw.PPMU.Pins(DUTPin).ForceV 0
        Else
            TheHdw.PPMU.Pins(DUTPin).ForceI 0
            TheHdw.PPMU.Pins(DUTPin).Gate = tlOff
            TheHdw.PPMU.Pins(DUTPin).Disconnect
        End If
        'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome

    Next DUTPin
    
    
        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop1
            Tname = "Conti1_" & CStr(DUTPin)
            If TestLimitMode = tlForceFlow Then
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            Else: TestLimitMode = tlForceNone
                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone 'original
                'TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceNone 'ZB
            End If

loop1:
            
        Next DUTPin
        
        If Separate_limit = True Then
            For Each DUTPin In Pins
                If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop2
                Tname = "Conti2_" & CStr(DUTPin)
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Else: TestLimitMode = tlForceNone
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                End If
loop2:
                
            Next DUTPin
        End If
    
''      'initialize ppmu to suitable clamp
''    With TheHdw.PPMU.Pins("Pins_1p0v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p1v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p8v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''    End With
    
     'Disconnect PPMU from digital channels
    With TheHdw.PPMU.Pins(digital_pins)
        .Gate = tlOff
        .Disconnect
    End With
             
    TheHdw.Digital.Pins(digital_pins).Connect
    
Next force_i
'///////////////////////////////////ZB add for measure contact resistance//////////////////////////////////////////////////////

    For Each site In TheExec.sites
        Calculate_Contact_R = (MeasV_Force50mA - MeasV_Force0mA) / 0.05
    Next site
    
    Dim RakV() As Double
        For Each site In TheExec.sites
            For Each DUTPin In PPMUMeasure.Pins
                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(DUTPin, Site)
                        Calculate_Contact_R = Calculate_Contact_R - (CurrentJob_Card_RAK.Pins(DUTPin).Value(site))
            Next DUTPin
        Next site
        
'    If LCase(CurrentJobName) Like LCase("*cp2*") Or LCase(CurrentJobName) Like LCase("*85*") Or LCase(CurrentJobName) Like LCase("*25*") Then
'        HiLimit_Resistance = 10
'    ElseIf LCase(CurrentJobName) Like LCase("*wlft*") Then
'        HiLimit_Resistance = 5
'    End If

If TheExec.Flow.EnableWord("DebugContact") = True Then
        TestNum = TheExec.Datalog.LastTestNumLogged
        
'========================Modified by Carter for Central Compile Review, 20181126, Function "VBT_print" not defined
'    Call VBT_print.needlepolish_r(Testnum, Calculate_Contact_R)
'====================================================================================================

End If

TheExec.Flow.TestLimit resultVal:=Calculate_Contact_R, hiVal:=HiLimit_Resistance, scaletype:=scaleNone, Unit:=unitOhm, PinName:=digital_pins    'ZB add for measure contact resistance

'TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceNone 'ZB modify
'///////////////////////////////////End measure contact resistance//////////////////////////////////////////////////////
    'if Alldigital_pins ="" then
        Alldigital_pins = "All_Digital"
        TheHdw.Digital.Pins(Alldigital_pins).Disconnect
        'end if
    Exit Function
errHandler:
        TheExec.AddOutput "Error in PPMU_Measure_Contact_Resistance"
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PPMU_Measure_Contact_Resistance_Corner_Vss(digital_pins As PinList, force_i_S As Double, force_i_E As Double, force_i_Step As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, PairNum As Double, Optional Separate_limit As Boolean = False, Optional independt_meas As Boolean) As Long
    
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim i As Long
    Dim force_i As Double
    Dim PPMUMeasure As New PinListData
    Dim site As Variant
    Dim MeasV_Force50mA As New PinListData: Dim MeasV_Force0mA As New PinListData: Dim Calculate_Contact_R As New PinListData

    On Error GoTo errHandler
'    thehdw.DCVS.pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff 'chihome

    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
    
    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)
        Calculate_Contact_R.AddPin (DUTPin)
        MeasV_Force50mA.AddPin (DUTPin)
        MeasV_Force0mA.AddPin (DUTPin)
    Next DUTPin
    
        For force_i = force_i_S To force_i_E Step (force_i_E - force_i_S) / force_i_Step
    

        Dim PinGroup As IPinListData
        Dim Power_conti_volt As New PinListData
        Dim PPMUMeas_HexVs As New PinListData
        Dim Tname As String

    
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    'DisconnectVDDCA 'SEC DRAM
    TheHdw.Wait 0.001
    
    '*************   Test digital channel continuity  ******************************
    ''''Disconnect os_pins Pin Electronics from pins in order to connect PPMU's''''
    'TheHdw.digital.Pins("Non_Conti_IO").Disconnect
    TheHdw.Digital.Pins(digital_pins).Disconnect
        
    '''''' Connect all os_pins to ppmu and ppmu force 0v for each one
    If independt_meas = False Then
        With TheHdw.PPMU.Pins(digital_pins)
            .ForceV 0#
            .Connect
            .Gate = tlOn
        End With
    End If

    
    
'    TheExec.DataManager.DecomposePinList digital_pins, Pins(), pin_cnt

    For Each DUTPin In Pins
'        PPMUMeasure.AddPin (DUTPin)

        If independt_meas = False Then
            With TheHdw.PPMU.Pins(DUTPin)
    ''            .ClampVHi = 1.2
    ''            .ClampVLo = -1
                .ForceI (force_i)
            End With
        Else
            With TheHdw.PPMU.Pins(DUTPin)
                .Connect
                .ForceI (force_i)
                .Gate = tlOn
            End With
        End If


        TheHdw.Wait 0.005
        
        DebugPrintFunc_PPMU CStr(DUTPin)
        PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)

'///////////////////////ZB add code///////////////////////////////////////
        For Each site In TheExec.sites
            If force_i = 0.05 Then
                MeasV_Force50mA = PPMUMeasure.Pins(DUTPin).Value
            Else
                MeasV_Force0mA = PPMUMeasure.Pins(DUTPin).Value
            End If
        Next site
'/////////////////////////////////////////////////////////////////////////
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
'                If LCase(TheExec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = -0.5 + Rnd() * 0.1
'                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = 0.5 + Rnd() * 0.1
                MeasV_Force50mA.Pins(DUTPin).Value(site) = 0.5 + Rnd() * 0.1
                MeasV_Force0mA.Pins(DUTPin).Value(site) = Rnd() * 0.1
            Next site
        End If
        
        
        'recover measure dut pin to 0V before next pin
        If independt_meas = False Then
            TheHdw.PPMU.Pins(DUTPin).ForceV 0
        Else
            TheHdw.PPMU.Pins(DUTPin).ForceI 0
            TheHdw.PPMU.Pins(DUTPin).Gate = tlOff
            TheHdw.PPMU.Pins(DUTPin).Disconnect
        End If
        'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome

    Next DUTPin
    
    
        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop1
            Tname = "Conti1_" & CStr(DUTPin)
            If TestLimitMode = tlForceFlow Then
'                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceFlow
            Else: TestLimitMode = tlForceNone
'                TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowval:=LowLimit, hival:=HiLimit, ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceNone 'original
                'TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceNone 'ZB
            End If

loop1:
            
        Next DUTPin
        
        If Separate_limit = True Then
            For Each DUTPin In Pins
                If TheExec.DataManager.ChannelType(DUTPin) = "N/C" Then GoTo loop2
                Tname = "Conti2_" & CStr(DUTPin)
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Else: TestLimitMode = tlForceNone
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                End If
loop2:
                
            Next DUTPin
        End If
    
''      'initialize ppmu to suitable clamp
''    With TheHdw.PPMU.Pins("Pins_1p0v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p1v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p8v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''    End With
    
     'Disconnect PPMU from digital channels
    With TheHdw.PPMU.Pins(digital_pins)
        .Gate = tlOff
        .Disconnect
    End With
             
    TheHdw.Digital.Pins(digital_pins).Connect
    
Next force_i
'///////////////////////////////////ZB add for measure contact resistance//////////////////////////////////////////////////////
    For Each DUTPin In Pins
        For Each site In TheExec.sites
            Calculate_Contact_R.Pins(DUTPin).Value(site) = (MeasV_Force50mA.Pins(DUTPin).Value(site) - MeasV_Force0mA.Pins(DUTPin).Value(site)) / 0.05
        Next site
    Next DUTPin
    
    Dim RakV() As Double
        For Each site In TheExec.sites
            For Each DUTPin In PPMUMeasure.Pins
                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(DUTPin, Site)
                            Calculate_Contact_R.Pins(DUTPin).Value(site) = Calculate_Contact_R.Pins(DUTPin).Value(site) - (CurrentJob_Card_RAK.Pins(DUTPin).Value(site))
            Next DUTPin
        Next site

TheExec.Flow.TestLimit resultVal:=Calculate_Contact_R, lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitOhm, PinName:=digital_pins      'ZB add for measure contact resistance
'TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceNone 'ZB modify
'///////////////////////////////////End measure contact resistance//////////////////////////////////////////////////////
    
    Exit Function
errHandler:
        TheExec.AddOutput "Error in Resistance Corner Vss"
        If AbortTest Then Exit Function Else Resume Next
End Function
Private Function GetMeasRange(InitResult As Double, MeasRangeList() As Double)
     Dim i As Integer
     Dim j As Integer
     Dim MinCurr As Double
     MinCurr = 0.0002
     GetMeasRange = MeasRangeList(UBound(MeasRangeList))
     For i = 0 To UBound(MeasRangeList)
        If (MeasRangeList(i) > 2 * Math.Abs(InitResult)) Then
            If (MeasRangeList(i) < 0.0002) Then
               For j = 0 To UBound(MeasRangeList)
                If (MeasRangeList(j) >= MinCurr) Then
                   GetMeasRange = MeasRangeList(j)
                   Exit For
                End If
               Next
            Else
              GetMeasRange = MeasRangeList(i)
            End If
          Exit For
        End If
     Next
End Function

Public Function p2p_short_Power_FVMI_Parallel(hexpowerpins As PinList, _
                     uvspowerpins As PinList, _
                     ForceV As Double, _
                     LowLimit As Double, _
                     HiLimit As Double, _
                     MaxForceV As Double, _
                     PinGroup_Cnt As Double, _
                     Step_Voltage_Level As Double, _
                     TestLimitMode As tlLimitForceResults, _
                     FlowLimitForInitIRange As Boolean, _
                     digital_pins As PinList, _
                     Optional InitRange200mAPins As String, _
                     Optional InitRange20mAPins As String, _
                     Optional InitRange2mAPins As String) As Long


   ''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Testing Method:  Force 0.1V , measure smaller than 199ma,set clamp to 200ma, if higher than 199 ma then fail
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim PowerPins As String
    Dim p As Variant, Pin_dcvi_Ary() As String
    Dim hexPin_Ary() As String, hexp_cnt As Long, uvsPin_Ary() As String, uvsp_cnt As Long
    Dim Tname As String
    Dim TempString As String
    Dim PowerSequence As Double
    Dim site As Variant
    Dim Pin As New PinList
    Dim FoldLimit As Double
    
    Dim MaxCurr As Double
    Dim MeasRangeVal As Double
    
    On Error GoTo errHandler
    
    Dim FlowLimitsInfo As IFlowLimitsInfo

    Dim Val As Double
    Dim Val_Hi() As String
    Dim Val_Lo() As String
    Dim HexVal_Hi() As String
    Dim HexVal_Lo() As String
    Dim UvsVal_Hi() As String
    Dim UvsVal_Lo() As String


    If hexpowerpins = "" And uvspowerpins = "" Then
        TheExec.Datalog.WriteComment "No Input pins."
        GoTo errHandler
    End If

    If (FlowLimitForInitIRange = True) Then
        Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
        FlowLimitsInfo.GetHighLimits Val_Hi
        FlowLimitsInfo.GetLowLimits Val_Lo
    End If

    
    Dim Merge_Type, Slot_Type As String
    Dim A_Slot_Type() As String
    Dim Split_Ary() As String
    Dim SattleTime As Double
    Dim WaitTime As Double
    
    Dim p_hexvs As String
    Dim p_uvs As String
    '''Dim p_dc07 As String
    
    Dim A_HexVS() As String
    Dim A_UVS() As String
    '''Dim A_DC07() As String
    
    Dim HexVS_Power_data As New PinListData
    Dim UVS_Power_data As New PinListData
    Dim DC07_Power_data As New PinListData

    Dim MaxValue As Double
    Dim getMaxValue As Double
    Dim MeasRangeList() As Double
    Dim UVI80MeasRangeList() As Double
    
 
    TheExec.DataManager.DecomposePinList hexpowerpins, hexPin_Ary, hexp_cnt
    TheExec.DataManager.DecomposePinList uvspowerpins, uvsPin_Ary, uvsp_cnt
    If hexpowerpins <> "" Then
        For Each p In hexPin_Ary
            If TheExec.DataManager.ChannelType(p) <> "N/C" Then p_hexvs = p_hexvs & "," & p
        Next p
    End If
    If uvspowerpins <> "" Then
        For Each p In uvsPin_Ary
            If TheExec.DataManager.ChannelType(p) <> "N/C" Then p_uvs = p_uvs & "," & p
        Next p
    End If
    If p_hexvs <> "" Then p_hexvs = Right(p_hexvs, Len(p_hexvs) - 1)
    If p_uvs <> "" Then p_uvs = Right(p_uvs, Len(p_uvs) - 1)
 
    'Pin_Ary = Split(PowerPins, ",")
    A_HexVS = Split(p_hexvs, ",")
    A_UVS = Split(p_uvs, ",")
    

    '''====================  Re-write
    Dim Num_PinGroup As Integer
    Dim AllPin_Cnt As Integer
    'Dim PinGroup_Cnt As Integer     '''input parameters, how many pins for one group
    'Dim Step_Voltage_Level As Double      '''input parameters, how much level pre step
    Dim PowerPins_Group() As String
    Dim Current_PinsGroup_ary() As String
    Dim Current_Val_Hi() As String
    Dim Current_Val_Lo() As String
    Dim CurrentGroup_HexVS() As String
    Dim CurrentGroup_UVS() As String
    
    Dim iCount As Integer: iCount = 0
    Dim GroupIndex As Integer: GroupIndex = 0
    Dim PinsIndex As Integer: PinsIndex = 0
    Dim EFUSELimit As Double: EFUSELimit = 0.1
    Dim TempForceV As Double: TempForceV = 0
'    Step_Voltage_Level = 0
    ''' setup by init default
    If MaxForceV = 0 Then MaxForceV = 3
    If PinGroup_Cnt = 0 Then PinGroup_Cnt = 10
    If Step_Voltage_Level = 0 Then Step_Voltage_Level = 0 '0.01

    AllPin_Cnt = UBound(A_HexVS) + UBound(A_UVS) + (2 * 1)
    If PinGroup_Cnt > AllPin_Cnt Then PinGroup_Cnt = AllPin_Cnt
    
    If AllPin_Cnt <> 0 And PinGroup_Cnt <> 0 Then
        ''' check PinGroup_Cnt is over max voltage
        If (ForceV + ((PinGroup_Cnt - 1) * Step_Voltage_Level)) > MaxForceV Then
            PinGroup_Cnt = CInt((MaxForceV - ForceV) / Step_Voltage_Level)
            TheExec.Datalog.WriteComment " Over max voltage, change PinGroup_Cnt to" & PinGroup_Cnt
        End If
        If AllPin_Cnt Mod PinGroup_Cnt <> 0 Then
            Num_PinGroup = Floor(AllPin_Cnt / PinGroup_Cnt) + 1
        Else
            Num_PinGroup = (AllPin_Cnt / PinGroup_Cnt)
        End If
    End If

    ReDim PowerPins_Group(Num_PinGroup - 1) As String
    ReDim p_hexVs_Group(Num_PinGroup - 1) As String
    ReDim p_uvs_Group(Num_PinGroup - 1) As String
    ReDim HexVS_Val_Hi(Num_PinGroup - 1) As String
    ReDim HexVS_Val_Lo(Num_PinGroup - 1) As String
    ReDim UVS_Val_Hi(Num_PinGroup - 1) As String
    ReDim UVS_Val_Lo(Num_PinGroup - 1) As String

    For GroupIndex = 0 To (Num_PinGroup - 1)
        For PinsIndex = 0 To (PinGroup_Cnt - 1)
            If iCount < UBound(A_HexVS) + 1 Then
                PowerPins_Group(GroupIndex) = PowerPins_Group(GroupIndex) & "," & A_HexVS(iCount)
                p_hexVs_Group(GroupIndex) = p_hexVs_Group(GroupIndex) & "," & A_HexVS(iCount)
                p_uvs_Group(GroupIndex) = ""
                HexVS_Val_Hi(GroupIndex) = HexVS_Val_Hi(GroupIndex) & "," & Val_Hi(iCount)
                HexVS_Val_Lo(GroupIndex) = HexVS_Val_Lo(GroupIndex) & "," & Val_Lo(iCount)
            Else
                PowerPins_Group(GroupIndex) = PowerPins_Group(GroupIndex) & "," & A_UVS(iCount - (UBound(A_HexVS) + 1))
                p_uvs_Group(GroupIndex) = p_uvs_Group(GroupIndex) & "," & A_UVS(iCount - (UBound(A_HexVS) + 1))
                UVS_Val_Hi(GroupIndex) = UVS_Val_Hi(GroupIndex) & "," & Val_Hi(iCount)
                UVS_Val_Lo(GroupIndex) = UVS_Val_Lo(GroupIndex) & "," & Val_Lo(iCount)
            End If
            iCount = iCount + 1

            If (iCount >= AllPin_Cnt) Then Exit For
        Next PinsIndex
        
        PowerPins_Group(GroupIndex) = Right(PowerPins_Group(GroupIndex), Len(PowerPins_Group(GroupIndex)) - 1)
        If p_hexVs_Group(GroupIndex) <> "" Then p_hexVs_Group(GroupIndex) = Right(p_hexVs_Group(GroupIndex), Len(p_hexVs_Group(GroupIndex)) - 1)
        If p_uvs_Group(GroupIndex) <> "" Then p_uvs_Group(GroupIndex) = Right(p_uvs_Group(GroupIndex), Len(p_uvs_Group(GroupIndex)) - 1)
        If HexVS_Val_Hi(GroupIndex) <> "" Then HexVS_Val_Hi(GroupIndex) = Right(HexVS_Val_Hi(GroupIndex), Len(HexVS_Val_Hi(GroupIndex)) - 1)
        If HexVS_Val_Lo(GroupIndex) <> "" Then HexVS_Val_Lo(GroupIndex) = Right(HexVS_Val_Lo(GroupIndex), Len(HexVS_Val_Lo(GroupIndex)) - 1)
        If UVS_Val_Hi(GroupIndex) <> "" Then UVS_Val_Hi(GroupIndex) = Right(UVS_Val_Hi(GroupIndex), Len(UVS_Val_Hi(GroupIndex)) - 1)
        If UVS_Val_Lo(GroupIndex) <> "" Then UVS_Val_Lo(GroupIndex) = Right(UVS_Val_Lo(GroupIndex), Len(UVS_Val_Lo(GroupIndex)) - 1)
    Next GroupIndex

    '''====================  Re-write end
'    ReDim A_Slot_Type(UBound(Pin_Ary)) As String
'    ReDim step_ary(UBound(Pin_Ary)) As Long

    TheHdw.DCVS.Pins(p_hexvs).Gate = False
    TheHdw.DCVS.Pins(p_hexvs).CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
    TheHdw.DCVS.Pins(p_uvs).Gate = False
    TheHdw.DCVS.Pins(p_uvs).CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff

    ' Slot_Type = GetInstrument(Pin_Ary(0), 0)
    ' If (LCase(Slot_Type) = "hexvs" Or LCase(Slot_Type) = "vhdvs") Then
    '     thehdw.DCVS.Pins(PowerPins).Gate = False
    '     thehdw.DCVS.Pins(allpowerpins).CurrentLimit.Source.FoldLimit.Behavior = tlDCVSCurrentLimitBehaviorGateOff
    ' ElseIf (LCase(Slot_Type) = "dc-07") Then
    '     thehdw.DCVI.Pins(PowerPins).Gate = False
    '     thehdw.DCVI.Pins(allpowerpins).FoldCurrentLimit.Behavior = tlDCVIFoldCurrentLimitBehaviorGateOff
    ' End If
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered     'SEC DRAM
    
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    If digital_pins.Value <> "" Then TheHdw.Digital.Pins(digital_pins).Disconnect
    
    TheHdw.Wait 3 * ms
    WaitTime = 260 * us
    
    '''====================  Re-write
    'Current_PinsGroup_ary = Split(PowerPins_Group(0), ",")  '''Get the frist pin group and first pins instrument only
    If hexpowerpins <> "" Then
        TheHdw.DCVS.Pins(hexpowerpins).Voltage.Main.Value = 0# '''Set all power pins to 0 V first
    End If

    If uvspowerpins <> "" Then
        TheHdw.DCVS.Pins(uvspowerpins).Voltage.Main.Value = 0# '''Set all power pins to 0 V first
    End If

    TheHdw.Wait 3 * ms

    For GroupIndex = 0 To (Num_PinGroup - 1)

        '''' init PinListData
        Set HexVS_Power_data = Nothing
        Set UVS_Power_data = Nothing
        iCount = 0

        Current_PinsGroup_ary = Split(PowerPins_Group(GroupIndex), ",")

        ''' set Step Voltage for each pins in Pin group
        For PinsIndex = 0 To UBound(Current_PinsGroup_ary)
            If UCase(Current_PinsGroup_ary(PinsIndex)) Like "*EFUSE*" And ((ForceV + (PinsIndex * Step_Voltage_Level)) >= EFUSELimit) Then
                TheExec.Datalog.WriteComment "Pin: " & Current_PinsGroup_ary(PinsIndex) & " Voltage Level over " & EFUSELimit & " v"
            End If
            TheHdw.DCVS.Pins(Current_PinsGroup_ary(PinsIndex)).Voltage.Main.Value = ForceV + (PinsIndex * Step_Voltage_Level) ''force V by pin
            TheExec.Datalog.WriteComment " Pin:" & Current_PinsGroup_ary(PinsIndex) & " Current Range: " & TheHdw.DCVS.Pins(Current_PinsGroup_ary(PinsIndex)).Meter.CurrentRange
        Next PinsIndex
        TheHdw.Wait 10 * ms

        '''Call set current range sub function
        ' Call SetCurrentRange(PowerPins_Group(GroupIndex), Current_Val_Hi(GroupIndex), Current_Val_Low(GroupIndex), FlowLimitForInitIRange, waittime, HiLimit, LowLimit, _
        '                 InitRange200mAPins, InitRange20mAPins, InitRange2mAPins)
        Call SetCurrentRange(p_hexVs_Group(GroupIndex), p_uvs_Group(GroupIndex), HexVS_Val_Hi(GroupIndex), UVS_Val_Hi(GroupIndex), HexVS_Val_Lo(GroupIndex), UVS_Val_Lo(GroupIndex), _
                                                            FlowLimitForInitIRange, WaitTime, HiLimit, LowLimit, InitRange200mAPins, InitRange20mAPins, InitRange2mAPins)

'        Debug.Print "wait time: " & waittime
        
        ' If (LCase(Slot_Type) = "hexvs") Then
        '     HexVS_Power_data = thehdw.DCVS.Pins(PowerPins_Group(GroupIndex)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
        ' Else If (LCase(Slot_Type) = "vhdvs") Then
        '     UVS_Power_data = thehdw.DCVS.Pins(PowerPins_Group(GroupIndex)).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
        ' Else If (LCase(Slot_Type) = "dc-07") Then

        ' End If
        
        'SattleTime = 0 * ms   ''' Do init

        ''' execute wiat time
        TheHdw.Wait WaitTime   ''calc in sub function
'        thehdw.Wait 100 * ms
        
        ''' execute meter Read
        If p_hexVs_Group(GroupIndex) <> "" Then HexVS_Power_data = TheHdw.DCVS.Pins(p_hexVs_Group(GroupIndex)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
        If p_uvs_Group(GroupIndex) <> "" Then UVS_Power_data = TheHdw.DCVS.Pins(p_uvs_Group(GroupIndex)).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)

        CurrentGroup_HexVS = Split(p_hexVs_Group(GroupIndex), ",")
        CurrentGroup_UVS = Split(p_uvs_Group(GroupIndex), ",")
        
        If p_hexVs_Group(GroupIndex) <> "" Then
            For PinsIndex = 0 To UBound(CurrentGroup_HexVS)
                Tname = "pwr_FVMI_" & CurrentGroup_HexVS(PinsIndex)
                'offline mode simulation
                If TheExec.TesterMode = testModeOffline Then
                    For Each site In TheExec.sites
                        HexVS_Power_data.Pins(CurrentGroup_HexVS(PinsIndex)).Value(site) = 0.01 + Rnd() * 0.0001
                    Next site
                End If
                TempForceV = (ForceV + (iCount * Step_Voltage_Level)) ''TheHdw.DCVS.Pins(CurrentGroup_HexVS(PinsIndex)).Voltage.Main.Value
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=HexVS_Power_data.Pins(CurrentGroup_HexVS(PinsIndex)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=TempForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                ElseIf TestLimitMode = tlForceNone Then
                    TheExec.Flow.TestLimit resultVal:=HexVS_Power_data.Pins(CurrentGroup_HexVS(PinsIndex)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=TempForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                End If
                iCount = iCount + 1
            Next PinsIndex
        End If
        
        If p_uvs_Group(GroupIndex) <> "" Then
            For PinsIndex = 0 To UBound(CurrentGroup_UVS)
                Tname = "pwr_FVMI_" & CurrentGroup_UVS(PinsIndex)
                'offline mode simulation
                If TheExec.TesterMode = testModeOffline Then
                    For Each site In TheExec.sites
                        UVS_Power_data.Pins(CurrentGroup_UVS(PinsIndex)).Value(site) = 0.01 + Rnd() * 0.0001
                    Next site
                End If
                TempForceV = (ForceV + (iCount * Step_Voltage_Level)) ''TheHdw.DCVS.Pins(CurrentGroup_UVS(PinsIndex)).Voltage.Main.Value
                If TestLimitMode = tlForceFlow Then
                    TheExec.Flow.TestLimit resultVal:=UVS_Power_data.Pins(CurrentGroup_UVS(PinsIndex)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=TempForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                ElseIf TestLimitMode = tlForceNone Then
                    TheExec.Flow.TestLimit resultVal:=UVS_Power_data.Pins(CurrentGroup_UVS(PinsIndex)), lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=(ForceV + (PinsIndex * Step_Voltage_Level)), ForceUnit:=unitVolt, ForceResults:=tlForceNone
                End If
                iCount = iCount + 1
            Next PinsIndex
        End If

        ''' reset this pins group voltage
        If p_hexVs_Group(GroupIndex) <> "" Then
            TheHdw.DCVS.Pins(p_hexVs_Group(GroupIndex)).Voltage.Main.Value = 0#
        End If
        If p_uvs_Group(GroupIndex) <> "" Then
            TheHdw.DCVS.Pins(p_uvs_Group(GroupIndex)).Voltage.Main.Value = 0#
        End If
        TheHdw.Wait 0.01
        

    Next GroupIndex

    TheHdw.Digital.ApplyLevelsTiming False, True, False, tlPowered       'SEC DRAM
    DebugPrintFunc ""
    '''====================  Re-write end

    Exit Function
 
errHandler:
 
    TheExec.Datalog.WriteComment " p2p_short_Power_FVMI_Parallel happens Error."
    If AbortTest Then Exit Function Else Resume Next
   
End Function

Public Function SetCurrentRange(HexMeasurePins As String, UVSMeasPins As String, HexLimitInfoHi As String, UVSLimitInfoHi As String, HexLimitInfoLow As String, UVSLimitInfoLow As String, _
                                                            FlowLimitInit As Boolean, Wait As Double, ArgHiLimit As Double, ArgLowLimit As Double, _
                                                            Optional PinsInitRange200mA As String, Optional PinsInitRange20mA As String, Optional PinsInitRange2mA As String)
    Dim HexMeasPinsAry() As String
    Dim UVSMeasPinsAry() As String
    Dim HexLimitHiAry() As String
    Dim UVSLimitHiAry() As String
    Dim HexLimitLowAry() As String
    Dim UVSLimitLowAry() As String
    Dim PinsInstrument As String
    Dim Irange As Double
    Dim PinIdx As Long
    ' Dim UVSPinIdx As Long
    Dim SetupTime As Double
    Dim CheckIdx As Long
    HexMeasPinsAry = Split(HexMeasurePins, ",")
    UVSMeasPinsAry = Split(UVSMeasPins, ",")
    HexLimitHiAry = Split(HexLimitInfoHi, ",")
    UVSLimitHiAry = Split(UVSLimitInfoHi, ",")
    HexLimitLowAry = Split(HexLimitInfoLow, ",")
    UVSLimitLowAry = Split(UVSLimitInfoLow, ",")
    If UBound(HexMeasPinsAry) <> UBound(HexLimitHiAry) Then
        GoTo errHandler
    ElseIf UBound(UVSMeasPinsAry) <> UBound(UVSLimitHiAry) Then
        GoTo errHandler
    End If
    
    If HexMeasurePins <> "" Then
        For PinIdx = 0 To UBound(HexMeasPinsAry)
        'If PinIdx = 0 Then PinsInstrument = LCase(GetInstrument(MeasPinsAry(0), 0))
        ' Select Case PinsInstrument
        '     Case "hexvs"
        '         'For i = 0 To UBound(Pin_Ary)
            If (FlowLimitInit = True) Then
                Irange = Abs(HexLimitHiAry(PinIdx))
            Else
                Irange = ArgHiLimit
            End If

            If Irange < 0.01 Then
                TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).SetCurrentRanges 0.01, 0.01  ' HexVS
                TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.05
                SetupTime = 100 * ms
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 3
            ElseIf Irange < 0.1 Then
                TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).SetCurrentRanges 0.1, 0.1    ' HexVS
                TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.1
                SetupTime = 10 * ms
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 2
            ElseIf Irange < 1 Then
                TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).SetCurrentRanges 1, 1    ' HexVS
                TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 1
                SetupTime = 1 * ms
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 1
            ElseIf Irange < 15 Then
                TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).SetCurrentRanges 15, 15   ' HexVS
                TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 15
                SetupTime = 100 * us
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 0
            Else
                'Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Max
                Irange = TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value
                TheHdw.DCVS.Pins(HexMeasPinsAry(PinIdx)).SetCurrentRanges Irange, Irange
                SetupTime = 100 * us
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 0
            End If

        Next PinIdx
    End If
    If UVSMeasPins <> "" Then
        For PinIdx = 0 To UBound(UVSMeasPinsAry)
                ' Case "vhdvs"
            If (FlowLimitInit = True) Then
                Irange = Abs(UVSLimitHiAry(PinIdx))
            Else
                Irange = ArgHiLimit
            End If
            If Irange < 0.000004 Then
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).SetCurrentRanges 0.000004, 0.000004
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentRange.Value = 0.000004
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.000004
                SetupTime = 18 * ms
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 6
            ElseIf Irange < 0.00002 Then
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).SetCurrentRanges 0.00002, 0.00002
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentRange.Value = 0.00002
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.00002
                SetupTime = 4 * ms
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 5
            ElseIf Irange < 0.0002 Then
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).SetCurrentRanges 0.0002, 0.0002
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentRange.Value = 0.0002
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.0002
                SetupTime = 4 * ms
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 4
            ElseIf Irange < 0.002 Then
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).SetCurrentRanges 0.002, 0.002
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentRange.Value = 0.002
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.002
                SetupTime = 3.5 * ms
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 3
            
            ElseIf Irange < 0.02 Then
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).SetCurrentRanges 0.02, 0.02
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentRange.Value = 0.02
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
                SetupTime = 540 * us
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 2
            ElseIf Irange < 0.04 Then
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).SetCurrentRanges 0.04, 0.04
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentRange.Value = 0.04
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.04
                SetupTime = 260 * us
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 1
            
            ElseIf Irange < 0.2 Then
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).SetCurrentRanges 0.2, 0.2
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentRange.Value = 0.2
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
                SetupTime = 210 * us
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 0
            ElseIf Irange < 0.7 Then
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).SetCurrentRanges 0.7, 0.7
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentRange.Value = 0.7
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value = 0.7
                SetupTime = 210 * us
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 0
            Else
                'Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Max
                Irange = TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).CurrentLimit.Source.FoldLimit.Level.Value
                TheHdw.DCVS.Pins(UVSMeasPinsAry(PinIdx)).SetCurrentRanges Irange, Irange
                SetupTime = 210 * us
                If SetupTime > Wait Then Wait = SetupTime
                'step_ary(i) = 0
            End If
        Next PinIdx
'    Else
                
        ' p_dc07 = p_dc07 & "," & MeasPinsAry(i)
        '     ' End Select
        ' ' Next PinIdx
    End If
        If (HexMeasurePins <> "" Or UVSMeasPins = "") Then
            If PinsInitRange200mA <> "" Then
                TheHdw.DCVS.Pins(PinsInitRange200mA).SetCurrentRanges 0.2, 0.2
                TheHdw.DCVS.Pins(PinsInitRange200mA).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
            End If
            If PinsInitRange20mA <> "" Then
                TheHdw.DCVS.Pins(PinsInitRange20mA).SetCurrentRanges 0.02, 0.02
                TheHdw.DCVS.Pins(PinsInitRange20mA).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
            End If
            If PinsInitRange2mA <> "" Then
                TheHdw.DCVS.Pins(PinsInitRange2mA).SetCurrentRanges 0.002, 0.002
                TheHdw.DCVS.Pins(PinsInitRange2mA).CurrentLimit.Source.FoldLimit.Level.Value = 0.002
            End If
        End If
    
        
        ''Debug.Print "Sub Wait:" & Wait
    Exit Function
errHandler:
        TheExec.Datalog.WriteComment "Pin Count nonmatch with Limit Count"
        If AbortTest Then Exit Function Else Resume Next
End Function
