Attribute VB_Name = "LIB_Common_DCVS_PPMU"
Option Explicit
'Revision History:
'V0.0 initial bring up
Public Const Version_Lib_Common_DC = "0.1" 'library version

'enums
Enum dcvs_type
    DCVS_HexVs = 1
    DCVS_HDVS = 2
    DCVS_UVS256 = 3
End Enum

Enum PinType
     SingleEnd = 0
     Differential = 1
End Enum

'variable declaration

'*****************************************
'******              DCVS Operations******
'*****************************************
Public Function DCVS_Trim_NC_Pin(ByRef original_ary() As String, ByRef original_pin_cnt As Long)
'for power pin special handling
Dim i As Long, j As Long
Dim p As Variant
Dim TempArray() As String
Dim TempPinCnt As Long
Dim NullArray() As String
Dim TempString As String
Dim PowerSequence As Double

If original_pin_cnt <> 0 Then
    i = 0   'init
    For Each p In original_ary
        TempString = p & "_PowerSequence_GLB"
        TempString = Replace(TempString, "MONITOR_", "")
        PowerSequence = TheExec.specs.Globals(TempString).ContextValue
        If PowerSequence <> 99 And TheExec.DataManager.ChannelType(p) <> "N/C" Then i = i + 1
        'If TheExec.DataManager.ChannelType(p) <> "N/C" Then i = i + 1
    Next p
    
    'redim
    ReDim TempArray(i - 1)
    
    j = 0   'init
    Dim k As Long
    k = 0
    For Each p In original_ary
        If TheExec.DataManager.ChannelType(p) <> "N/C" Then
            TempArray(j) = original_ary(k)
            j = j + 1
        Else
            j = j
        End If
        k = k + 1
    Next p
    
    original_ary = TempArray
    original_pin_cnt = j
End If
    
End Function

Public Function DCVS_MeterRead(dcvs_type As dcvs_type, powerPin As String, sample_size As Long, ByRef MeasCurr As PinListData)
'strobe measure values with DCVS instrument
    On Error GoTo errHandler
    Dim site As Variant
    Dim p As Variant
    Dim ResetI As Boolean
    Dim Default_CurrentRange As Double
    ResetI = False
    Select Case dcvs_type
    
        Case DCVS_HexVs, DCVS_HDVS:
            MeasCurr = TheHdw.DCVS.Pins(powerPin).Meter.Read(tlStrobe, sample_size, 10000, tlDCVSMeterReadingFormatAverage)
             
        Case DCVS_UVS256:
            TheHdw.DCVS.Pins(powerPin).Meter.Filter.Value = TheHdw.DCVS.Pins(powerPin).Meter.Filter.max / sample_size ' UVS average 10 samples
            MeasCurr = TheHdw.DCVS.Pins(powerPin).Meter.Read(tlStrobe, 1)  ' UVS only allow one sample
            
        Case Else
            GoTo errHandler 'catch out of bag
    End Select
    
    Exit Function
errHandler:
    ErrorDescription ("DCVS_MeterRead")
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function DCVS_Set_Meter_Range(MeasureI_pin As PinList, Irange As String)
'setup current range for HardIP I measurement
    Dim Pins() As String
    Dim Pin_Cnt As Long
    Dim UVS_Ring() As String
    Dim UVs_rng() As String
    Dim defautIrange As Double
    Dim i As Integer
    Dim WaitTime As Double
    defautIrange = 0.05
    WaitTime = 0.001
   ' If Irange = "" Then TheExec.Datalog.WriteComment (" Please set the Meter range"): Exit Function
    If Irange = "" Then Irange = defautIrange
    TheExec.DataManager.DecomposePinList MeasureI_pin, Pins(), Pin_Cnt
    
    UVs_rng = Split(Irange, ",")
    While UBound(UVs_rng) < UBound(Pins)
    
    Irange = Irange & "," & defautIrange
    UVs_rng = Split(Irange, ",")
    Wend
    
    
    For i = 0 To Pin_Cnt - 1 'UBound(UVs_rng) '- 1
         If UVs_rng(i) = "" Then MsgBox "Please set the Meter range"
         
          If CDbl(UVs_rng(i)) > 0.2 Then
              UVs_rng(i) = 0.4
              WaitTime = 0.001
         ElseIf CDbl(UVs_rng(i)) > 0.02 Then
              UVs_rng(i) = 0.2
              WaitTime = 0.001
         ElseIf CDbl(UVs_rng(i)) > 0.002 Then
               UVs_rng(i) = 0.02
               WaitTime = 0.001
         ElseIf CDbl(UVs_rng(i)) > 0.0002 Then
              UVs_rng(i) = 0.002
              WaitTime = 0.03
         ElseIf CDbl(UVs_rng(i)) > 0.00002 Then
              UVs_rng(i) = 0.0002
              WaitTime = 0.05
         ElseIf CDbl(UVs_rng(i)) > 0.000002 Then
              UVs_rng(i) = 0.00002
              WaitTime = 0.07
         Else
              UVs_rng(i) = 0.000004
              WaitTime = 0.1
         End If
    If TheExec.DataManager.ChannelType(Pins(i)) <> "N/C" Then
         With TheHdw.DCVS.Pins(Pins(i))
             .Meter.mode = tlDCVSMeterCurrent
             
             '.Meter.Filter.Bypass = False
             '.Meter.Filter.Value = 98
             .SetCurrentRanges CDbl(UVs_rng(i)), CDbl(UVs_rng(i))
             
             .Gate = True
         End With
         
         
         
   ' TheExec.Datalog.WriteComment ("                                                     =====> Curr_meas Meter I range setting, " & Pins(i) & " =" & UVs_rng(i))
    TheExec.Datalog.WriteComment ("                       =====> Curr_meas Meter I range setting, " & Pins(i) & " =" & TheHdw.DCVS.Pins(Pins(i)).Meter.CurrentRange.Value)
    End If
    Next i
    
    
    
    TheHdw.Wait (WaitTime)
End Function

Public Function DCVS_PowerOn_I_Meter(Pin As String, v As Double, i_rng As Double, wait_before_gate As Double, wait_after_gate As Double, Steps As Integer, RiseTime As Double, Optional DebugPrintEnable As Boolean = False)
'set Force voltage and Current/Meter Range
''===============================================
''Description: __                __
''                __|
''            __|
''        __|
''    __| >|   |<--stepT  __        v
''__|                   __ stepV   __
''|<-- steps -->
''|<-FallTime ->
''===============================================
    On Error GoTo errHandler
    
    Dim i_meter_rng As Double
    Dim setV As Double
    Dim StepV As Double
    Dim stepT As Double
    
    Dim i As Integer
    
    i_meter_rng = i_rng
    StepV = v / Steps
    stepT = RiseTime / Steps

    With TheHdw.DCVS.Pins(Pin)
        .Connect
        .mode = tlDCVSModeVoltage
        .Voltage.Main = 0
        .Meter.mode = tlDCVSMeterCurrent
        
        If i_rng <> -99 Then    'bypass range setup if does not need to
            .SetCurrentRanges i_rng, i_meter_rng
            .CurrentRange.Value = i_rng
            .CurrentLimit.Source.FoldLimit.Level.Value = i_rng
            .Meter.CurrentRange = i_rng
        End If
        
        TheHdw.Wait wait_before_gate   'wait for relay connect
        
        .Gate = True
    End With
    
''Pwr On Ramp up slew-rate control============================
    For i = 1 To Steps
        setV = i * StepV
        TheHdw.DCVS.Pins(Pin).Voltage.Main = setV
        
        If DebugPrintEnable = True Then
            TheExec.Datalog.WriteComment "  Curr_" & Pin & " Pwr Up Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        
        TheHdw.Wait stepT
    Next i
''============================================================

    TheHdw.Wait wait_after_gate
    
    Exit Function
    
errHandler:
    ErrorDescription ("DCVS_PowerOn_I_Meter")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DCVS_PowerOff_I_Meter_Parallel(Pin As String, wait_before_gate As Double, wait_after_gate As Double, PowerSequence As Long, Optional DebugPrintEnable As Boolean = False)
'set Force voltage and Current/Meter Range
''===============================================
''Description
''__                              __
''  |__
''     |_>|  |<--stepT             v
''        |__          __
''           |__       __ stepV   __
''|<-- steps -->
''|<-FallTime ->
''===============================================
    On Error GoTo errHandler
    
    Dim i_meter_rng As Double   'meter range
    Dim setV As Double          'current voltage
    Dim StepV As Double         'step voltage
    Dim stepT As Double         'step time
    Dim PowerPins() As String
    Dim PinCnt As Long
    Dim powerPin As Variant
    Dim TempString As String
    Dim Vmain(50) As Double
    Dim Irange As Double
    Dim step As Integer
    Dim PreStep As Integer:: PreStep = 0
    Dim RiseTime As Double
    
    Dim i As Integer:: i = 1
    Dim j As Integer:: j = 1

    TheExec.DataManager.DecomposePinList Pin, PowerPins(), PinCnt
    
    For Each powerPin In PowerPins
        'If TheExec.DataManager.ChannelType(PowerPin) <> "N/C" Then 'check CP for FT form NC pins
            TempString = powerPin & "_GLB"
            'Vmain(i) = TheExec.specs.Globals(TempString).ContextValue
            Vmain(i) = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
            'get Ifold limit spec value
            TempString = powerPin & "_Ifold_GLB"
            Irange = TheExec.specs.Globals(TempString).ContextValue
            
            'auto calculate steps
            step = Vmain(i) / 0.1 '0.1v per step
            If step = 0 Then step = 10 'default value
            If step > PreStep Then PreStep = step
            
            RiseTime = step * ms
            i_meter_rng = Irange
        
            With TheHdw.DCVS.Pins(powerPin)
                .mode = tlDCVSModeVoltage
                .SetCurrentRanges Irange, i_meter_rng
                '.Meter.mode = tlDCVSMeterCurrent
                .CurrentRange.Value = Irange
                .CurrentLimit.Source.FoldLimit.Level.Value = Irange
                .Meter.CurrentRange = Irange

            End With
            
            If DebugPrintEnable = True Then    'debugprint
                TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin, 30, False) & ", Vmain " & Format(Vmain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", FallTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
            End If
            
            i = i + 1
'        Else
'            If DebugPrintEnable = True Then    'debugprint
'                TheExec.Datalog.WriteComment "print: Pin " & PowerPin & " not turn on by 'NC pin', PowerSequence " & PowerSequence & " ,Warning!!!"
'            End If
'        End If
    Next powerPin
    
    step = PreStep
    RiseTime = step * ms
    stepT = RiseTime / step
    
    With TheHdw.DCVS.Pins(Pin)
        .Connect
        TheHdw.Wait wait_before_gate
        .Gate = True
    End With
    
    ''Pwr On Ramp Down slew-rate control============================
    For j = 1 To step
        i = 1
        For Each powerPin In PowerPins
''            If TheExec.DataManager.ChannelType(PowerPin) <> "N/C" Then 'check CP for FT form NC pins  'no need to double check NC pin
                setV = Vmain(i) - (j * Vmain(i) / step)
                TheHdw.DCVS.Pins(powerPin).Voltage.Main = setV
                
''                If DebugPrintEnable = True Then
''                    TheExec.Datalog.WriteComment "  Curr_" & PowerPin & " Pwr Down Voltage (" & CStr(i) & ") : " & Format(setV, "0.000") & " V"
''                End If
            i = i + 1
''            End If
        Next powerPin
        TheHdw.Wait stepT   'wait step time
    Next j
    
    setV = 0    'final step, return to 0V anyway
    TheHdw.DCVS.Pins(Pin).Voltage.Main = setV
    
''    If DebugPrintEnable = True Then
''        TheExec.Datalog.WriteComment "  Curr_" & PowerPin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
''    End If
    ''==============================================================
    
    TheHdw.Wait wait_after_gate
    
    Exit Function
    
errHandler:
    ErrorDescription ("DCVS_PowerOff_I_Meter")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DCVS_PowerOn_I_Meter_Parallel(Pin As String, wait_before_gate As Double, wait_after_gate As Double, PowerSequence As Long, Optional DebugPrintEnable As Boolean = False)
'set Force voltage and Current/Meter Range
''===============================================
''Description: __                __
''                __|
''            __|
''        __|
''    __| >|   |<--stepT  __        v
''__|                   __ stepV   __
''|<-- steps -->
''|<-FallTime ->
''===============================================
    On Error GoTo errHandler
    
    Dim i_meter_rng As Double
    Dim setV As Double
    Dim StepV As Double
    Dim stepT As Double
    Dim PowerPins() As String
    Dim PinCnt As Long
    Dim powerPin As Variant
    Dim TempString As String
    Dim Vmain(50) As Double
    Dim Irange As Double
    Dim step As Integer
    Dim PreStep As Integer:: PreStep = 0
    Dim RiseTime As Double
    
    Dim i As Integer:: i = 1
    Dim j As Integer:: j = 1
    Dim temp_str As String
    
    TheExec.DataManager.DecomposePinList Pin, PowerPins(), PinCnt

    If TheExec.specs.DC.Contains(PowerPins(0) & "_VAR_H") Then 'Carter, 20190502
        temp_str = "_VAR_H"
    Else
        temp_str = "_VAR"
    End If

    For Each powerPin In PowerPins
        TempString = powerPin & temp_str
        Vmain(i) = TheExec.specs.DC.Item(TempString).ContextValue
        
        'get Ifold limit spec value
        TempString = powerPin & "_Ifold_GLB"
        Irange = TheExec.specs.Globals(TempString).ContextValue
        
        'auto calculate steps
        step = Vmain(i) / 0.1 '0.1v per step
        If step = 0 Then step = 10 'default value
        If step > PreStep Then PreStep = step   'calculate largest ramp up steps from all powers in the same sequence
        
        RiseTime = step * ms
        i_meter_rng = Irange
    
        With TheHdw.DCVS.Pins(powerPin)
            .mode = tlDCVSModeVoltage
            .SetCurrentRanges Irange, i_meter_rng
            '.Meter.mode = tlDCVSMeterCurrent
            .CurrentRange.Value = Irange
            .CurrentLimit.Source.FoldLimit.Level.Value = Irange
            .Meter.CurrentRange = Irange
        End With
        
        If DebugPrintEnable = True Then    'debugprint
            TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin, 30, False) & ", Vmain " & Format(Vmain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", RiseTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
        End If
        
        i = i + 1
    Next powerPin
    
    step = PreStep
    RiseTime = step * ms
    stepT = RiseTime / step
    
    With TheHdw.DCVS.Pins(Pin)
        .Connect
        .Voltage.Main = 0
        TheHdw.Wait wait_before_gate
        .Gate = True
    End With
    

''Pwr On Ramp up slew-rate control============================
    For j = 1 To step
        i = 1
        For Each powerPin In PowerPins
            setV = j * Vmain(i) / step
            TheHdw.DCVS.Pins(powerPin).Voltage.Main = setV
''            If DebugPrintEnable = True Then
''                TheExec.Datalog.WriteComment "  Curr_" & PowerPin & " Pwr Up Voltage (" & CStr(i) & ") : " & Format(setV, "0.000") & " V"
''            End If
            i = i + 1
        Next powerPin
        
        TheHdw.Wait stepT
        
    Next j
''============================================================

    TheHdw.Wait wait_after_gate
    
    Exit Function
    
errHandler:
    ErrorDescription ("DCVS_PowerOn_I_Meter_Parallel")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DCVS_PowerOff_I_Meter(Pin As String, v As Double, i_rng As Double, wait_before_gate As Double, wait_after_gate As Double, Steps As Integer, FallTime As Double, Optional DebugPrintEnable As Boolean = False)
'set Force voltage and Current/Meter Range
''===============================================
''Description
''__                              __
''  |__
''     |_>|  |<--stepT             v
''        |__          __
''           |__       __ stepV   __
''|<-- steps -->
''|<-FallTime ->
''===============================================
    On Error GoTo errHandler
    
    Dim i_meter_rng As Double   'meter range
    Dim setV As Double          'current voltage
    Dim StepV As Double         'step voltage
    Dim stepT As Double         'step time
    
    Dim i As Integer
    Dim stepsm As Integer
    
    i_meter_rng = i_rng
    StepV = v / Steps
    stepT = FallTime / Steps
    
    With TheHdw.DCVS.Pins(Pin)
        .Connect
        .mode = tlDCVSModeVoltage
        .Voltage.Main = v
        .SetCurrentRanges i_rng, i_meter_rng
        .Meter.mode = tlDCVSMeterCurrent
        .CurrentRange.Value = i_rng
        .CurrentLimit.Source.FoldLimit.Level.Value = i_rng
        .Meter.CurrentRange = i_rng
        
        TheHdw.Wait wait_before_gate   'wait for relay connect
        
        .Gate = True
    End With
        
    ''Pwr On Ramp Down slew-rate control============================
    stepsm = Steps - 1
    For i = 0 To stepsm
        setV = v - (i * StepV)
        TheHdw.DCVS.Pins(Pin).Voltage.Main = setV
        
        If DebugPrintEnable = True Then
            TheExec.Datalog.WriteComment "  Curr_" & Pin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        
        TheHdw.Wait stepT   'wait step time
    Next i
    
    setV = 0    'final step, return to 0V anyway
    TheHdw.DCVS.Pins(Pin).Voltage.Main = setV
    
    If DebugPrintEnable = True Then
        TheExec.Datalog.WriteComment "  Curr_" & Pin & " Pwr Down Voltage (" & CStr(i) & ") : " & CStr(setV) & " V"
    End If
    ''==============================================================
    
    TheHdw.Wait wait_after_gate
    
    Exit Function
    
errHandler:
    ErrorDescription ("DCVS_PowerOff_I_Meter")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function DCVS_P2P_short_FIMV(dcvs_type As dcvs_type, Pins As PinList, ForceV As Double, Output_Current_range As Double, Meter_Current_Range As Double, SinkFoldLimit As Double, SampleSize As Long, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, wait_time As Double)
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim p As Variant
    Dim measure_result As New PinListData
    Dim Tname As String
    Dim TempString As String
    Dim PowerSequence As Double
    Dim site As Variant
    
    If Pins = "" Then Exit Function ' check if p is null
    
    TheExec.DataManager.DecomposePinList Pins, Pin_Ary, Pin_Cnt
    For Each p In Pin_Ary
        TempString = ""
        TempString = p & "_PowerSequence_GLB"
        PowerSequence = TheExec.specs.Globals(TempString).ContextValue
        If PowerSequence = 99 Then GoTo loop1
        
        If TheExec.DataManager.ChannelType(p) = "N/C" Then GoTo loop1   'NC pin, no measure
        
        Tname = "uvs_FIMV_" & p
        With TheHdw.DCVS.Pins(p)
            .Voltage.Main = ForceV
            .SetCurrentRanges Output_Current_range, Meter_Current_Range
            .CurrentLimit.Sink.FoldLimit.Level.Value = SinkFoldLimit
        End With
        
        Wait wait_time
        
        DCVS_MeterRead DCVS_UVS256, CStr(p), 10, measure_result
        TheHdw.DCVS.Pins(p).Voltage.Main = 0# ' recover to 0 after measurement
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                    measure_result.Pins(p).Value(site) = -0.7 + 0 * 0.1
                    'measure_result.Pins(p).Value(site) = -0.7 + Rnd() * 0.1
            Next site
        End If
        
        
        If TestLimitMode = tlForceFlow Then
            Call TheExec.Flow.TestLimit(resultVal:=measure_result, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=SinkFoldLimit, ForceUnit:=unitAmp, ForceResults:=tlForceFlow)
        Else
            Call TheExec.Flow.TestLimit(resultVal:=measure_result, lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=SinkFoldLimit, ForceUnit:=unitAmp, ForceResults:=tlForceNone)

        End If
        
        
        DebugPrintFunc ""                                        ' add for Miner 20151103
        If TheExec.sites.Active.Count = 0 Then Exit Function
loop1:
        
        
    Next p
    
    Exit Function
    
End Function

'*****************************************
'******              PPMU Operations******
'*****************************************

Public Function PPMU_FIMV_P2P_short(Pins As PinList, force_i As Double, wait_time As Double, sample_size As Long, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults)
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim p As Variant
    Dim MeasureResult As New PinListData
    Dim Tname As String
    Dim tmp_p As String
    Dim TempString As String
    Dim PowerSequence As Double
    Dim site As Variant

    'If p = "" Then Exit Function ' check if p is null

    TheExec.DataManager.DecomposePinList Pins, Pin_Ary, Pin_Cnt
    For Each p In Pin_Ary
        TempString = ""
        TempString = p & "_PowerSequence_GLB"
        'TempString = Replace(TempString, "Monitor_", "")
        TempString = Replace(TempString, "_SENSE", "")
        PowerSequence = TheExec.specs.Globals(TempString).ContextValue
        
        If PowerSequence = 99 Then GoTo loop1
        If TheExec.DataManager.ChannelType(p) = "N/C" Then GoTo loop1
        
        Tname = "hexvs_FIMV_" & p
        'tmp_p = Replace(p, "_Monitor", "")
        tmp_p = Replace(p, "_SENSE", "")
        TheHdw.DCVS.Pins(tmp_p).BleederResistor = tlDCVSOff
        With TheHdw.PPMU.Pins(p)
        .ForceI force_i
        .Connect
        .Gate = tlOn
'        .ClampVHi = 1.8 * 1.1
'        .ClampVLo = -1
        End With

        TheHdw.Wait wait_time
        DebugPrintFunc_PPMU CStr(p)
        MeasureResult = TheHdw.PPMU.Pins(p).Read(tlPPMUReadMeasurements, sample_size)
        
''        DebugPrintFunc ""                                        ' add for Miner 20151103
''        TheHdw.PPMU.Pins(p).ForceV 0
''        TheHdw.PPMU.Pins(p).Gate = tlOff
''        TheHdw.PPMU.Pins(p).Disconnect

        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                    MeasureResult.Pins(p).Value(site) = -0.7 + 0 * 0.1
                    'MeasureResult.Pins(p).Value(site) = -0.7 + Rnd() * 0.1
            Next site
        End If

        If TestLimitMode = tlForceFlow Then
            Call TheExec.Flow.TestLimit(resultVal:=MeasureResult, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow)
        Else
            Call TheExec.Flow.TestLimit(resultVal:=MeasureResult, lowVal:=LowLimit, hiVal:=HiLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone)
        End If

        DebugPrintFunc ""                                        ' add for Miner 20151103
        TheHdw.PPMU.Pins(p).ForceV 0
        TheHdw.PPMU.Pins(p).Gate = tlOff
        TheHdw.PPMU.Pins(p).Disconnect
        
        If TheExec.sites.Active.Count = 0 Then Exit Function 'chihome


loop1:


    Next p

    Exit Function

End Function

Public Function PowerOff_I_Meter_Parallel(Pin As String, wait_before_gate As Double, wait_after_gate As Double, PowerSequence As Long, Optional DebugPrintEnable As Boolean = False)

    On Error GoTo errHandler
    
    Dim i_meter_rng As Double   'meter range
    Dim setV As Double          'current voltage
    Dim StepV As Double         'step voltage
    Dim stepT As Double         'step time
    Dim PowerPins() As String
    Dim PinCnt As Long
    Dim powerPin As Variant
    Dim TempString As String
    Dim Vmain(50) As Double
    Dim Irange As Double
    Dim step As Integer
    Dim PreStep As Integer:: PreStep = 0
    Dim RiseTime As Double
    Dim pin_name As String
    Dim Pin_Type() As String
    Dim SlotType As String
    Dim pins_dcvs As String, pins_dcvi As String

    Dim i As Integer:: i = 1
    Dim j As Integer:: j = 1
    Dim k As Integer
    
    TheExec.DataManager.DecomposePinList Pin, PowerPins(), PinCnt
    ReDim Pin_Type(PinCnt)
    
    For Each powerPin In PowerPins
            TempString = powerPin & "_GLB"
            pin_name = powerPin
            SlotType = LCase(GetInstrument(pin_name, 0))
            Select Case SlotType
                Case "dc-07":
                        Vmain(i) = TheHdw.DCVI.Pins(powerPin).Voltage
                Case Else
                        Vmain(i) = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
            End Select

            'get Ifold limit spec value
            TempString = powerPin & "_Ifold_GLB"
            Irange = TheExec.specs.Globals(TempString).ContextValue
'''            Vmain(i) = 0.1
'''            Irange = 0.1

            'auto calculate steps
            step = Vmain(i) / 0.1 '0.1v per step
            If step = 0 Then step = 10 'default value
            If step > PreStep Then PreStep = step
            
            RiseTime = step * ms
            i_meter_rng = Irange
            '---------------------------------------------------------------------
            Select Case SlotType
                Case "dc-07":
                            With TheHdw.DCVI.Pins(powerPin)
                                 .mode = tlDCVIModeVoltage
                                 .SetCurrentAndRange Irange, i_meter_rng
                                 .CurrentRange.Value = Irange
                                 .current = Irange
                                 .Meter.CurrentRange = Irange
                            End With
                            Pin_Type(k) = "dcvi"
                            pins_dcvi = pins_dcvi + "," + powerPin
                Case Else
                            Vmain(i) = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
                            With TheHdw.DCVS.Pins(powerPin)
                                .mode = tlDCVSModeVoltage
                                .SetCurrentRanges Irange, i_meter_rng
                                '.Meter.mode = tlDCVSMeterCurrent
                                .CurrentRange.Value = Irange
                                .CurrentLimit.Source.FoldLimit.Level.Value = Irange
                                .Meter.CurrentRange = Irange
                            End With
                            Pin_Type(k) = "dcvs"
                            pins_dcvs = pins_dcvs + "," + powerPin
            End Select
            '---------------------------------------------------------------------
            If DebugPrintEnable = True Then    'debugprint
                TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin, 30, False) & ", Vmain " & Format(Vmain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", FallTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
            End If
            
            i = i + 1:: k = k + 1
    Next powerPin
    
    step = PreStep
    RiseTime = step * ms
    stepT = RiseTime / step

    pins_dcvi = Mid(pins_dcvi, 2, Len(pins_dcvi))
    pins_dcvs = Mid(pins_dcvs, 2, Len(pins_dcvs))
    
    If pins_dcvi <> "" Then
        With TheHdw.DCVI.Pins(pins_dcvi)
            .Connect
            'TheHdw.Wait wait_before_gate
            .Gate = True
        End With
    End If
    
    If pins_dcvs <> "" Then
        With TheHdw.DCVS.Pins(pins_dcvs)
            .Connect
            'TheHdw.Wait wait_before_gate
            .Gate = True
        End With
    End If
    TheHdw.Wait wait_before_gate

    '=============================================================='Pwr On Ramp Down slew-rate control
    For j = 1 To step
        i = 1
        For Each powerPin In PowerPins
                setV = Vmain(i) - (j * Vmain(i) / step)
                If Pin_Type(i - 1) = "dcvs" Then
                    TheHdw.DCVS.Pins(powerPin).Voltage.Main = setV
                ElseIf Pin_Type(i - 1) = "dcvi" Then
                    TheHdw.DCVI.Pins(powerPin).Voltage = setV
                End If
            i = i + 1
        Next powerPin
        TheHdw.Wait stepT   'wait step time
    Next j
    
    setV = 0    'final step, return to 0V anyway
    
    If pins_dcvi <> "" Then
        TheHdw.DCVI.Pins(pins_dcvi).Voltage = setV
    End If
    
    If pins_dcvs <> "" Then
        TheHdw.DCVS.Pins(pins_dcvs).Voltage.Main = setV
    End If
    ''==============================================================
    
    TheHdw.Wait wait_after_gate
    Exit Function
    
errHandler:
    ErrorDescription ("PowerOff_I_Meter_Parallel")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PowerOn_I_Meter_Parallel(Pin As String, wait_before_gate As Double, wait_after_gate As Double, PowerSequence As Long, Optional DebugPrintEnable As Boolean = False)

    On Error GoTo errHandler
    
    Dim i_meter_rng As Double
    Dim setV As Double
    Dim StepV As Double
    Dim stepT As Double
    Dim PowerPins() As String
    Dim PinCnt As Long
    Dim powerPin As Variant
    Dim TempString As String
    Dim Vmain(50) As Double
    Dim Irange As Double
    Dim step As Integer
    Dim PreStep As Integer:: PreStep = 0
    Dim RiseTime As Double
    Dim pin_name As String
    Dim Pin_Type() As String
    Dim SlotType As String
    Dim pins_dcvs As String, pins_dcvi As String

    Dim i As Integer:: i = 1
    Dim j As Integer:: j = 1
    Dim k As Integer
    
    TheExec.DataManager.DecomposePinList Pin, PowerPins(), PinCnt
    
    ReDim Pin_Type(PinCnt)
    
    For Each powerPin In PowerPins
'        TempString = powerPin & "_GLB"
'        Vmain(i) = TheExec.Specs.Globals(TempString).ContextValue
        TempString = powerPin & "_VAR"                                                          'new test setting for AP
        Vmain(i) = TheExec.specs.DC(TempString).ContextValue

        'get Ifold limit spec value
        TempString = powerPin & "_Ifold_GLB"
        Irange = TheExec.specs.Globals(TempString).ContextValue
'''        Vmain(i) = 0.1
'''        Irange = 0.1
        'auto calculate steps
        step = Vmain(i) / 0.1 '0.1v per step
        If step = 0 Then step = 10 'default value
        If step > PreStep Then PreStep = step   'calculate largest ramp up steps from all powers in the same sequence
        
        RiseTime = step * ms
        i_meter_rng = Irange
        '---------------------------------------------------------------------
        pin_name = powerPin
        SlotType = LCase(GetInstrument(pin_name, 0))
        Select Case SlotType
            Case "dc-07":
                    With TheHdw.DCVI.Pins(powerPin)
                         .mode = tlDCVIModeVoltage
                         .SetCurrentAndRange Irange, i_meter_rng
                         .CurrentRange.Value = Irange
                         .current = Irange
                         .Meter.CurrentRange = Irange
                    End With
                    Pin_Type(k) = "dcvi"
                    pins_dcvi = pins_dcvi + "," + powerPin
            Case Else
                    With TheHdw.DCVS.Pins(powerPin)
                        .mode = tlDCVSModeVoltage
                        .SetCurrentRanges Irange, i_meter_rng
                        '.Meter.mode = tlDCVSMeterCurrent
                        .CurrentRange.Value = Irange
                        .CurrentLimit.Source.FoldLimit.Level.Value = Irange
                        .Meter.CurrentRange = Irange
                    End With
                    Pin_Type(k) = "dcvs"
                    pins_dcvs = pins_dcvs + "," + powerPin
        End Select
        '---------------------------------------------------------------------
        
        If DebugPrintEnable = True Then    'debugprint
            TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(powerPin, 30, False) & ", Vmain " & Format(Vmain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", RiseTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
        End If
        
        i = i + 1:: k = k + 1
    Next powerPin
    
    step = PreStep
    RiseTime = step * ms
    stepT = RiseTime / step
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    pins_dcvi = Mid(pins_dcvi, 2, Len(pins_dcvi))
    pins_dcvs = Mid(pins_dcvs, 2, Len(pins_dcvs))
    If pins_dcvi <> "" Then
        With TheHdw.DCVI.Pins(pins_dcvi)
            .Connect
            .Voltage = 0
            'TheHdw.Wait wait_before_gate
            .Gate = True
        End With
    End If
    
    If pins_dcvs <> "" Then
        With TheHdw.DCVS.Pins(pins_dcvs)
            .Connect
            .Voltage.Main = 0
            'TheHdw.Wait wait_before_gate
            .Gate = True
        End With
    End If
    
    TheHdw.Wait wait_before_gate
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '========================================================'Pwr On Ramp up slew-rate control
    For j = 1 To step
        i = 1
        For Each powerPin In PowerPins
            setV = j * Vmain(i) / step
            If Pin_Type(i - 1) = "dcvs" Then
               TheHdw.DCVS.Pins(powerPin).Voltage.Main = setV
            ElseIf Pin_Type(i - 1) = "dcvi" Then
               TheHdw.DCVI.Pins(powerPin).Voltage = setV
            End If
            i = i + 1
        Next powerPin
        
        TheHdw.Wait stepT
    Next j
    ''============================================================

    TheHdw.Wait wait_after_gate
    
    Exit Function
    
errHandler:
    ErrorDescription ("PowerOn_I_Meter_Parallel")
    If AbortTest Then Exit Function Else Resume Next
End Function
