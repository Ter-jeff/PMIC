Attribute VB_Name = "VBT_LIB_DC_Func"
Option Explicit

Public Type GetShmooVihVil
    Vih As New SiteDouble
    Vil As New SiteDouble
End Type

Public Type Pin_parameter
            Voltage As Double
            Start_voltage As Double
            End_voltage As Double
End Type
Public Type Pins_detail
    PinName As String
    Step_value As Double
    Step_value_up As Double
    Step_value_down As Double
    Start_voltage As Double
    Start_voltage_up As Double
    Start_voltage_down As Double
    Pin_rise As Boolean
    LatchUp_Final_Value As Double
    Gate_check As New SiteBoolean
End Type

Public GPIO_Vih_Vil(100) As GetShmooVihVil

Public g_CFG_GPIO_PF_Val_LV As New PinListData '' 0 is pass, 1 is fail
Public g_CFG_GPIO_PF_Val_NV As New PinListData '' 0 is pass, 1 is fail
Public g_CFG_GPIO_PF_Val_HV As New PinListData '' 0 is pass, 1 is fail
Public GPIO_driver_result As New SiteDouble


'Revision History:
'V0.0 initial bring up

' This module should be used for VBT Tests.  All functions in this module
' will be available to be used from the Test Instance sheet.
' Additional modules may be added as needed (all starting with "VBT_").
'
' The required signature for a VBT Test is:
'
' Public Function FuncName(<arglist>) As Long
'   where <arglist> is any list of arguments supported by VBT Tests.
'
' See online help for supported argument types in VBT Tests.
'
'
' It is highly suggested to use error handlers in VBT Tests.  A sample
' VBT Test with a suggeseted error handler is shown below:
'
' Function FuncName() As Long
'     On Error GoTo errHandler
'
'     Exit Function
' errHandler:
'     If AbortTest Then Exit Function Else Resume Next
' End Function
Public Function DC_Func_WriteFuncResult(Optional PerPinFailLog As Boolean = False) As Long
    Dim site As Variant
    Dim TestNumber As Long
    Dim FailCount As New PinListData
    Dim allpins As PinList
    Dim Pin As Variant
    Dim Pins() As String
    Dim Pin_Cnt As Long
    
     TheExec.DataManager.DecomposePinList "All_Digital", Pins(), Pin_Cnt
    'AllPins = "AllPins"
   ' allpins.Value=
  
    For Each site In TheExec.sites
        TestNumber = TheExec.sites.Item(site).TestNumber
        If TheHdw.Digital.Patgen.PatternBurstPassed(site) Then  'collect pattern burst result
            Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
            Else
            Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
            TheExec.sites.Item(site).testResult = siteFail

            
            If PerPinFailLog Then   'per pin fail log collection
                For Each Pin In Pins
                'thehdw.Digital.Pins(pin)
                    If TheExec.DataManager.ChannelType(Pin) <> "N/C" Then
                        FailCount = TheHdw.Digital.Pins(Pin).FailCount
                        If FailCount <> 0 Then TheExec.Datalog.WriteComment "===> Pin " & Pin & " Fail count =" & FailCount
                    End If
                Next Pin
            End If
        End If
        TheExec.sites.Item(site).TestNumber = TestNumber + 1
    Next site
    

End Function


Public Function Meas_LeakCurr_Univeral_func(patset As Pattern, DisableComparePins As PinList, DisableConnectPins As PinList, TestSequence As String, CPUA_Flag_In_Pat As Boolean, PpmuMeasureI_Pin As PinList, _
                                                 MeasOnHalt As Boolean, Flow_Limit_Result As tlLimitForceResults, _
                                                 I_LowLimit As Double, I_HighLimit As Double, IMeasRange As String, ForceV As Double, ForceVMidPoint As Double, Optional EnableTwoPoint As Boolean = False) As Long
'Parameter information     **2013/04/12 by JT
'Patset : Test Pattern
'DisableComparePins     :   Disable Pin Compare H/L.
'TestSequence           :   Decide to test which function and sequence, "v,i,vi,f" means test sequence will be
'                           1. Meas voltagee
'                           2. Meas voltage and current at the same CPU flag loop
'                           3. Meas frequence
'CPUA_Flag_In_Pat       :   If CPUA flag in Pattern?
'PpmuMeasureV_Pin       :   Meas voltage pin
'FreqCtrMeasurePins     :   Meas frequence pin
'MeasureI_pin           :   Meas current pin
'MeasFreqPinType        :   Meas frequence pin is different or single end.
'MeasOnHalt             :   If need to measure any thing after Pattern halt


    Dim i As Integer ' for the use of loading pattern
    
    Dim pat_count As Long
    
    Dim MeasVoltage As New PinListData
    Dim MeasCurr As New PinListData
    Dim MeasCurr_MidPoint As New PinListData
    Dim MeasCurr_Delta As New PinListData
    
    Dim idx As Long
    Dim k As Long
    
    Dim Status As Boolean
    Dim p As Variant, Pin_Ary() As String, p_cnt As Long
    Dim CPUA_Flag_Cnt As Integer
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim p_hexvs As Variant, p_hexvs_ary() As String, PinCnt_hexvs As Long
    Dim p_uvs As Variant, p_uvs_ary() As String, PinCnt_uvs As Long
    Dim TestNum As Long, Cnt1 As Long
    Dim p_uvs_idx As Integer
    Dim p_uvs_str As String
    Dim patt_ary() As String
    Dim Pat As Variant
    Dim MaxSpec As Double
    Dim ppmuRange As Double
    Dim site As Variant

    
    On Error GoTo errHandler
    
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
    'actual leakage test start
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    'disconnect pins
    If DisableConnectPins <> "" Then
        TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    End If
    
    If (DisableComparePins <> "") Then
        TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    End If
    
    TheHdw.Patterns(patset).Load
    Status = GetPatListFromPatternSet(patset.Value, patt_ary, pat_count)
    pat_count = 1
 
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    
    TheExec.DataManager.DecomposePinList PpmuMeasureI_Pin, Pin_Ary, p_cnt
     
    For Each Pat In patt_ary
        Call TheHdw.Patterns(Pat).start
        TheHdw.Wait 0.001
        

            
        TestSeqNum = 0
        
        For Each Ts In TestSequenceArray
            
            If MeasOnHalt = True And TestSeqNum = UBound(TestSequenceArray) Then
                Call TheHdw.Digital.Patgen.HaltWait 'Meas after patgen halt
            ElseIf (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If

            TestOptLen = Len(Ts)

            For k = 1 To TestOptLen
                TestOption = Mid(Ts, k, 1)
                Select Case TestOption
                
                    Case "i", "I"
    
                        '#####################################
                        ' Current Measurement
                        '#####################################
                        If (PpmuMeasureI_Pin <> "") Then
                            TheHdw.Digital.Pins(PpmuMeasureI_Pin).Disconnect
                            With TheHdw.PPMU.Pins(PpmuMeasureI_Pin)   '' make sure which pins
                               .ForceV ForceV, IMeasRange  'avoid hard code here
                               .Connect
                               .Gate = True
                            End With
                        End If
                
                        'check ppmu range vs spec
                        ppmuRange = TheHdw.PPMU.Pins(PpmuMeasureI_Pin).MeasureCurrentRange
                        MaxSpec = max(Abs(I_LowLimit), Abs(I_HighLimit))
                        
                        TheExec.Flow.TestLimit ppmuRange, MaxSpec, , tlSignGreaterEqual, tlSignLessEqual, Tname:="Abs(Spec) vs Range" 'PPMU Range check PASS
                        
                        
                        TestNum = TheExec.sites.Item(0).TestNumber
                        DebugPrintFunc_PPMU PpmuMeasureI_Pin.Value
                        MeasCurr = TheHdw.PPMU.Pins(PpmuMeasureI_Pin).Read(tlPPMUReadMeasurements, 10)
                        
                        '20160301 - Format alignment from C651 Mai (Discard following line) -- TYCHENGG
                        'TheExec.Flow.TestLimit MeasCurr, I_LowLimit, I_HighLimit, , , , unitAmp, , Tname:="Current_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceResults:=Flow_Limit_Result, forceVal:=ForceV
                        TheExec.Flow.TestLimit MeasCurr, I_LowLimit, I_HighLimit, , , , unitAmp, , Tname:="Current_meas" + "_" + CStr(TestSeqNum), ForceResults:=Flow_Limit_Result, ForceVal:=ForceV
                        If EnableTwoPoint Then   '2-points
                        
                         If (PpmuMeasureI_Pin <> "") Then
                            TheHdw.Digital.Pins(PpmuMeasureI_Pin).Disconnect
                            With TheHdw.PPMU.Pins(PpmuMeasureI_Pin)   '' make sure which pins
                               .ForceV ForceVMidPoint, IMeasRange  'avoid hard code here
                               .Connect
                               .Gate = True
                            End With
                        End If
                        DebugPrintFunc_PPMU PpmuMeasureI_Pin.Value
                        MeasCurr_MidPoint = TheHdw.PPMU.Pins(PpmuMeasureI_Pin).Read(tlPPMUReadMeasurements, 10)
                        
                        '20160301 - Format alignment from C651 Mai (Discard following line) -- TYCHENGG
                        'TheExec.Flow.TestLimit MeasCurr_MidPoint, I_LowLimit, I_HighLimit, , , , unitAmp, , Tname:="Current_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceResults:=Flow_Limit_Result, forceVal:=ForceVMidPoint
                        TheExec.Flow.TestLimit MeasCurr_MidPoint, I_LowLimit, I_HighLimit, , , , unitAmp, , Tname:="Current_meas" + "_" + CStr(TestSeqNum), ForceResults:=Flow_Limit_Result, ForceVal:=ForceVMidPoint
                        
                        MeasCurr_Delta = MeasCurr.Copy
                        'Delta result
                        For Each site In TheExec.sites
                            For idx = 0 To p_cnt - 1
                                MeasCurr_Delta.Pins.Item(idx).Value = MeasCurr.Pins.Item(idx).Value - MeasCurr_MidPoint.Pins.Item(idx).Value
                            Next idx
                        Next site
                        
                        '20160301 - Format alignment from C651 Mai (Discard following line) -- TYCHENGG
                        'TheExec.Flow.TestLimit MeasCurr_Delta, I_LowLimit, I_HighLimit, , , , unitAmp, , Tname:="Current_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceResults:=Flow_Limit_Result, forceVal:=(ForceV - ForceVMidPoint)
                        TheExec.Flow.TestLimit MeasCurr_Delta, I_LowLimit, I_HighLimit, , , , unitAmp, , Tname:="Current_meas" + "_" + CStr(TestSeqNum), ForceResults:=Flow_Limit_Result, ForceVal:=(ForceV - ForceVMidPoint)
                        
                        End If
                        
                        'If CurrentJobName Like "*char*" Then Call Char_show(MeasCurr, I_HighLimit, I_LowLimit, TestNum, "A", "_" + CStr(TestSeqNum))

                    Case Else
                         TheExec.Datalog.WriteComment "Error Test Option, please select I"
                End Select
            Next k
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
            End If
  
        Next Ts
        
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        
        Call DC_Func_WriteFuncResult
                
        pat_count = pat_count + 1
    Next Pat
    
    'reset digital pin conditions
    If PpmuMeasureI_Pin <> "" Then
        With TheHdw.PPMU.Pins(PpmuMeasureI_Pin)
            .ForceV (0)
            .Gate = tlOff
            .Disconnect
        End With
        TheHdw.Digital.Pins(PpmuMeasureI_Pin).Connect
    End If
    
    If DisableConnectPins <> "" Then TheHdw.Digital.Pins(DisableConnectPins).Connect
    
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    
    DebugPrintFunc patset.Value  ' print all debug information
    
Exit Function


errHandler:
    TheExec.Datalog.WriteComment "Meas_LeakCurr_Univeral_func is error "
    If AbortTest Then Exit Function Else Resume Next
End Function





Public Function Meas_VOHL_Univeral_func(patset As Pattern, DisableComparePins As PinList, DisableConnectPins As PinList, TestSequence As String, CPUA_Flag_In_Pat As Boolean, PpmuMeasureV_Pin As PinList, _
                                                 MeasOnHalt As Boolean, Flow_Limit_Result As tlLimitForceResults, _
                                                 V_LowLimit As Double, V_HighLimit As Double, Irange As String, ForceI As Double, StartLabel As String, StopLabel As String) As Long
'Parameter information     **2013/04/12 by JT
'Patset : Test Pattern
'DisableComparePins     :   Disable Pin Compare H/L.
'TestSequence           :   Decide to test which function and sequence, "v,i,vi,f" means test sequence will be
'                           1. Meas voltagee
'                           2. Meas voltage and current at the same CPU flag loop
'                           3. Meas frequence
'CPUA_Flag_In_Pat       :   If CPUA flag in Pattern?
'PpmuMeasureV_Pin       :   Meas voltage pin
'FreqCtrMeasurePins     :   Meas frequence pin
'MeasureI_pin           :   Meas current pin
'MeasFreqPinType        :   Meas frequence pin is different or single end.
'MeasOnHalt             :   If need to measure any thing after Pattern halt


    Dim i As Integer ' for the use of loading pattern
    
    Dim pat_count As Long, Pin As Variant
    Dim MeasFreqSingle As New PinListData
    Dim MeasFreqDifferential As New PinListData
    
    Dim MeasVoltage As New PinListData
    Dim MeasCurr As New PinListData
    
    Dim j As Long
    Dim k As Long
    Dim freq As Double
    Dim freq_limit_upper As Double
    Dim freq_limit_lower As Double
    Dim Status As Boolean
    Dim p As Long
    Dim CPUA_Flag_Cnt As Integer
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim BurstResult As New SiteLong
    Dim AllSitePass As Boolean
    Dim patt_ary() As String
    Dim Pat As Variant
    Dim site As Variant
    
    On Error GoTo errHandler
    
    TestSequenceArray = Split(TestSequence, ",")
    
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered 'SEC DRAM
    
    If (DisableConnectPins <> "") Then
        TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    End If

  '  Dim patout As String
  '  Call ShowPrintFunc(patset.Value, patout)
    
    
    If (DisableComparePins <> "") Then
        TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    End If
    
    TheHdw.Patterns(patset).Load
    Status = GetPatListFromPatternSet(patset.Value, patt_ary, pat_count)
    pat_count = 1

  
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Wait 0.001
     
    For Each Pat In patt_ary
        'TheExec.Datalog.WriteComment " "

        TheExec.Datalog.WriteComment "Pat =  " & Pat
       ' Call thehdw.Patterns(pat).Reload
        'Call thehdw.Patterns(pat).test(pfAlways, 0)

        For i = 0 To 1
            Call TheHdw.Patterns(Pat).start(StartLabel, StopLabel)
            TheHdw.Digital.Patgen.HaltWait
            AllSitePass = True
            For Each site In TheExec.sites
                BurstResult(site) = 1
                
                If (TheHdw.Digital.Patgen.PatternBurstPassed(site) = False) Then
                    TheExec.Datalog.WriteComment vbCrLf & Pat & "_" & StopLabel & vbTab & "Run " & i & " : Fail."
                    BurstResult(site) = 0
                    AllSitePass = False
                End If
            Next site
            If AllSitePass = True Then Exit For
        Next i
        TheExec.Flow.TestLimit BurstResult, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:=StartLabel & "_Result" 'BurstResult=1:Pass


        If (AllSitePass = False) Then GoTo End_Meas

        With TheHdw.PPMU.Pins(PpmuMeasureV_Pin)   '' make sure which pins
           .ForceI 0
           .Gate = False
           .Disconnect
''           .ClampVHi = 1.8 * 1.1
''           .ClampVLo = -1#
        End With

        TheHdw.Wait 0.001


        TheExec.Datalog.WriteComment "Force Current = " & ForceI
        TestSeqNum = 0
        For Each Ts In TestSequenceArray
            
            If MeasOnHalt = True And TestSeqNum = UBound(TestSequenceArray) Then
                Call TheHdw.Digital.Patgen.HaltWait 'Meas after patgen halt
            ElseIf (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            
            TestOptLen = Len(Ts)
            
            Dim p_hexvs As Variant, p_hexvs_ary() As String, PinCnt_hexvs As Long
            Dim p_uvs As Variant, p_uvs_ary() As String, PinCnt_uvs As Long
            Dim TestNum As Long, Cnt1 As Long
            Dim p_uvs_idx As Integer
            Dim p_uvs_str As String
            
            For k = 1 To TestOptLen
            
         
                TestOption = Mid(Ts, k, 1)
                Select Case TestOption
                    Case "V", "v"
    
                        '#####################################
                        ' Voltage Measurement
                        '#####################################
                        Dim pin_array() As String
                        Dim PinCnt As Long
                        
                        TheExec.DataManager.DecomposePinList PpmuMeasureV_Pin, pin_array, PinCnt
                        
                        For Each Pin In pin_array
                                TestNum = TheExec.sites.Item(0).TestNumber
                                If (PpmuMeasureV_Pin <> "") Then
                                    TheHdw.Digital.Pins(Pin).Disconnect
                                    With TheHdw.PPMU.Pins(Pin)   '' make sure which pins
                                        .ForceI ForceI
                                        .Connect
''                                        .ClampVHi = 1.8 * 1.1
''                                        .ClampVLo = -1#
                                        .Gate = True
                                    End With
                                End If
                                TheHdw.Wait 0.001
                                DebugPrintFunc_PPMU CStr(Pin)
                                MeasVoltage = TheHdw.PPMU.Pins(Pin).Read(tlPPMUReadMeasurements, 20)

                                TheHdw.Wait 0.001
                                'Call Char_show(MeasVoltage, V_HighLimit, V_LowLimit, TestNum, "V", "_" + CStr(TestSeqNum))
                                                
                                If (PpmuMeasureV_Pin <> "") Then
                                    'thehdw.digital.pins(pin).Disconnect
                                    With TheHdw.PPMU.Pins(Pin)   '' make sure which pins
                                       .ForceI 0
                                       .Gate = False
                                       .Disconnect
''                                       .ClampVHi = 2#
                                    End With
                                    TheHdw.Digital.Pins(Pin).Connect

                                End If
'
                                'offline mode simulation
                                If TheExec.TesterMode = testModeOffline Then
                                    For Each site In TheExec.sites
                                        If LCase(TheExec.DataManager.instanceName) Like "*voh*" Then MeasVoltage.Pins(Pin).Value = 1.5 + Rnd() * 0.1
                                        If LCase(TheExec.DataManager.instanceName) Like "*vol*" Then MeasVoltage.Pins(Pin).Value = 0.2 + Rnd() * 0.1
                                    Next site
                                End If
                                
                                'TheExec.Flow.TestLimit MeasVoltage.Pins(pin), V_LowLimit, V_HighLimit, , , scaleNone, unitVolt, , tname:="Volt_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), forceResults:=Flow_Limit_Result
                                If TheExec.DataManager.ChannelType(Pin) = "N/C" Then GoTo loop1
                                    
                                TheExec.Flow.TestLimit resultVal:=MeasVoltage.Pins(Pin), lowVal:=V_LowLimit, hiVal:=V_HighLimit, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", _
                                Tname:="Volt_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceVal:=ForceI, ForceUnit:=unitAmp, ForceResults:=tlForceNone

loop1:
                      Next Pin
                     

                    Case Else
                         TheExec.Datalog.WriteComment "Error Test Option, please select V,I or F"
                
                End Select
            
            Next k
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
            End If
            


        Next Ts
        
        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & pat_count & "): " & Pat & ""
        
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end

End_Meas:
    
        pat_count = pat_count + 1
        
    Next Pat
    

    
    If DisableComparePins <> "" Then
        TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    End If
    
    DebugPrintFunc patset.Value  ' print all debug information
Exit Function


errHandler:
    TheExec.Datalog.WriteComment "Meas_VOHL_Univeral_func is error "
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Meas_VOHL_Univeral_func_Parallel(patset As Pattern, DisableComparePins As PinList, DisableConnectPins As PinList, TestSequence As String, CPUA_Flag_In_Pat As Boolean, PpmuMeasureV_Pin As PinList, _
                                                 MeasOnHalt As Boolean, Flow_Limit_Result As tlLimitForceResults, _
                                                 V_LowLimit As Double, V_HighLimit As Double, Irange As String, ForceI As Double, StartLabel As String, StopLabel As String, Optional mod_count As Integer = 1) As Long
'Parameter information     **2013/04/12 by JT
'Patset : Test Pattern
'DisableComparePins     :   Disable Pin Compare H/L.
'TestSequence           :   Decide to test which function and sequence, "v,i,vi,f" means test sequence will be
'                           1. Meas voltagee
'                           2. Meas voltage and current at the same CPU flag loop
'                           3. Meas frequence
'CPUA_Flag_In_Pat       :   If CPUA flag in Pattern?
'PpmuMeasureV_Pin       :   Meas voltage pin
'FreqCtrMeasurePins     :   Meas frequence pin
'MeasureI_pin           :   Meas current pin
'MeasFreqPinType        :   Meas frequence pin is different or single end.
'MeasOnHalt             :   If need to measure any thing after Pattern halt


    Dim i As Integer ' for the use of loading pattern
    
    Dim pat_count As Long, Pin As Variant
    Dim MeasFreqSingle As New PinListData
    Dim MeasFreqDifferential As New PinListData
    
    Dim MeasVoltage As New PinListData
    Dim MeasCurr As New PinListData
    
    Dim j As Long
    Dim k As Long
    Dim freq As Double
    Dim freq_limit_upper As Double
    Dim freq_limit_lower As Double
    Dim Status As Boolean
    'Dim p As Long
    Dim CPUA_Flag_Cnt As Integer
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim BurstResult As New SiteLong
    Dim AllSitePass As Boolean
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim group_count As Long
    'Dim mod_count As Long
    Dim PinString As Variant    'For MaxQQ
    Dim patt_ary() As String
    Dim Pat As Variant
    Dim site As Variant
    Dim p As Variant
    Dim NewPins() As Variant    'For MaxQQ
    
    On Error GoTo errHandler
    
    TestSequenceArray = Split(TestSequence, ",")
    TheExec.DataManager.DecomposePinList PpmuMeasureV_Pin, Pins(), Pin_Cnt
    
    'mod_count = 19
    
    NewPins = GroupPinsByMod(Pins, mod_count)    'For MaxQQ
       
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
    group_count = 0
    
    For Each PinString In NewPins
    
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
            TheHdw.Digital.Pins(PinString).Levels.Value(chIol) = ForceI
            TheHdw.Digital.Pins(PinString).Levels.Value(chIoh) = ForceI * (-1)
            TheHdw.Digital.Pins(PinString).Levels.Value(chVol) = V_LowLimit
            TheHdw.Digital.Pins(PinString).Levels.Value(chVoh) = V_HighLimit
        
            If (DisableComparePins <> "") Then
                TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
            End If
            
            If (DisableConnectPins <> "") Then
                TheHdw.Digital.Pins(DisableComparePins).Disconnect
            End If
            
            TheHdw.Patterns(patset).Load
            Status = GetPatListFromPatternSet(patset.Value, patt_ary, pat_count)
            pat_count = 1
        
            Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
            TheHdw.Wait 0.001
             
            For Each Pat In patt_ary
        
                    Call TheHdw.Patterns(Pat).start
                    TheHdw.Digital.Patgen.HaltWait
                    'all site pattern burst pass checking
''                    AllSitePass = True
''                    For Each Site In TheExec.Sites
''                        BurstResult(Site) = 1
''
''                        If (thehdw.Digital.Patgen.PatternBurstPassed(Site) = False) Then
''                            'TheExec.Datalog.WriteComment vbCrLf & Pat & "_" & StopLabel & vbTab & "Run " & i & " : Fail."
''                            BurstResult(Site) = 0
''                            AllSitePass = False
''                        End If
''                    Next Site
        
               'assign test name as pin name
                If group_count = 0 Then TheExec.Datalog.WriteComment vbCrLf
                TheExec.Datalog.WriteComment "print: Group " & (group_count + 1) & ", Pins: " & FormatNumericDatalog(PinString, 10 * mod_count, True) & ", " & "Voh spec: " & Format(V_HighLimit, "0.000") & " V, Vol spec: " & Format(V_LowLimit, "0.000") & " V, ForceI: " & Format(ForceI, "0.000") & " A"
                
                Call DC_Func_WriteFuncResult(False)
''                TheExec.DataManager.DecomposePinList PinString, Pins(), pin_cnt
''                For Each p In Pins
''                    TheExec.Flow.TestLimit BurstResult, 1, 1, tlSignGreaterEqual, tlSignLessEqual, ScaleType:=scaleNone, Tname:=p & " in Group " & Format((group_count + 1), "0"), forceVal:=ForceI, forceunit:=unitAmp  'BurstResult=1:Pass
''                Next p
           
            pat_count = pat_count + 1
            Next Pat
                              
            If DisableComparePins <> "" Then
                TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
            End If
            
            If (DisableConnectPins <> "") Then
                TheHdw.Digital.Pins(DisableComparePins).Connect
            End If
            
            group_count = group_count + 1
            
    Next PinString
    
    
''    For Each Pat In patt_ary
''        TheExec.Datalog.WriteComment "Pat =  " & Pat
''    Next Pat
    
    DebugPrintFunc patset.Value  ' print all debug information

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Meas_VOHL_Univeral_func_Parallel is error "
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Meas_VOH_MeasVI_Univeral_func_DC(patset As Pattern, DisableComparePins As PinList, DisableConnectPins As PinList, TestSequence As String, CPUA_Flag_In_Pat As Boolean, PpmuMeasure_Pin As PinList, FreqMeasSingleEndPins As PinList, FreqMeasDiffEndPins As PinList, MeasureI_pin As PinList, _
                                                 MeasOnHalt As Boolean, Flow_Limit_Result As tlLimitForceResults, PPMU_V_LowLimit As String, PPMU_V_HighLimit As String, PPMU_I_LowLimit As String, PPMU_I_HighLimit As String, PPMU_R_LowLimit As String, PPMU_R_HighLimit As String, Reset_V As String, Reset_I As String _
                                                , I_LowLimit As Double, I_HighLimit As Double, Freq_LowLimit As Double, Freq_HighLimit As Double, ForceV As String, ForceI As String, PortName As String, Optional Flag_Meas_VOH_Freq As Boolean = False, Optional DisableClock As Boolean = False, Optional Flag_bypass_meas_I As Boolean = False, Optional Irange As Double) As Long
' EDITFORMAT1 1,,Pattern,Group1,,patset|2,,PinList,,,DisableComparePins|3,,PinList,,,DisableConnectPins|4,,String,,,TestSequence|5,,Boolean,,,CPUA_Flag_In_Pat|6,,PinList,,,PpmuMeasure_Pin|7,,PinList,,,FreqMeasSingleEndPins|8,,PinList,,,FreqMeasDiffEndPins|9,,PinList,,,MeasureI_pin|10,,Boolean,,,MeasOnHalt|11,,tlLimitForceResults,,,Flow_Limit_Result|12,,Double,,,PPMU_V_LowLimit|13,,Double,,,PPMU_V_HighLimit|14,,Double,,,PPMU_I_LowLimit|15,,Double,,,PPMU_I_HighLimit|16,,Double,,,PPMU_R_LowLimit|17,,Double,,,PPMU_R_HighLimit|18,,Double,,,I_LowLimit|19,,Double,,,I_HighLimit|20,,Double,,,Freq_LowLimit|21,,Double,,,Freq_HighLimit|22,,String,,,ForceV|23,,String,,,ForceI|24,,String,,,PortName|25,,Boolean,,,Flag_Meas_VOH_Freq|26,,Boolean,,,DisableClock|27,,Boolean,,,Flag_bypass_meas_I
'Parameter information     **2013/04/12 by JT
'Patset : Test Pattern
'DisableComparePins     :   Disable Pin Compare H/L.
'TestSequence           :   Decide to test which function and sequence, "v,i,vi,f" means test sequence will be
'                           1. Meas voltagee
'                           2. Meas voltage and current at the same CPU flag loop
'                           3. Meas frequence
'CPUA_Flag_In_Pat       :   If CPUA flag in Pattern?
'PpmuMeasure_Pin       :   Meas voltage pin
'FreqCtrMeasurePins     :   Meas frequence pin
'MeasureI_pin           :   Meas current pin
'MeasFreqPinType        :   Meas frequence pin is different or single end.
'MeasOnHalt             :   If need to measure any thing after Pattern halt


    Dim i As Integer ' for the use of loading pattern
    
    Dim pat_count As Long
    Dim MeasFreqSingle As New PinListData
    Dim MeasFreqDifferential As New PinListData
    
    Dim MeasVoltage As New PinListData
    Dim MeasCurrent As New PinListData
    Dim MeasCurr As New PinListData
    Dim z As Integer
    Dim j As Long
    Dim k As Long
    Dim freq As Double
    Dim freq_limit_upper As Double
    Dim freq_limit_lower As Double
    Dim Status As Boolean
    Dim p As Long
    Dim CPUA_Flag_Cnt As Integer
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, ForceISequenceArray() As String, ForceVSequenceArray() As String
    Dim PPMU_V_LowLimitArray() As String, PPMU_I_LowLimitArray() As String, PPMU_R_LowLimitArray() As String
    Dim PPMU_V_HighLimitArray() As String, PPMU_I_HighLimitArray() As String, PPMU_R_HighLimitArray() As String
    Dim Reset_V_Array() As String, Reset_I_Array() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim Pins() As String, Pin_Cnt As Long
    Dim DUTPin As Variant
    Dim patout As String
    Dim patt_ary() As String
    Dim Pat As Variant
    Dim site As Variant
    
    TestSequenceArray = Split(TestSequence, ",")
    ForceISequenceArray = Split(ForceI, ",")
    ForceVSequenceArray = Split(ForceV, ",")
    PPMU_V_LowLimitArray = Split(PPMU_V_LowLimit, ",")
    PPMU_V_HighLimitArray = Split(PPMU_V_HighLimit, ",")
    PPMU_I_LowLimitArray = Split(PPMU_I_LowLimit, ",")
    PPMU_I_HighLimitArray = Split(PPMU_I_HighLimit, ",")
    PPMU_R_LowLimitArray = Split(PPMU_R_LowLimit, ",")
    PPMU_R_HighLimitArray = Split(PPMU_R_HighLimit, ",")
    Reset_V_Array = Split(Reset_V, ",")
    Reset_I_Array = Split(Reset_I, ",")
    
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    On Error GoTo errHandler
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If (DisableConnectPins <> "") Then
        TheHdw.Digital.Pins(DisableComparePins).Disconnect
    End If
    
    If (DisableComparePins <> "") Then
        TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    End If
    TheHdw.Patterns(patset).Load
    Call GetPatListFromPatternSet(patset.Value, patt_ary, pat_count)

    If FreqMeasSingleEndPins <> "" Then
        Call Freq_MeasFreqSetup(FreqMeasSingleEndPins, 0.001)
    End If
    If FreqMeasDiffEndPins <> "" Then
        Call Freq_MeasFreqSetup(FreqMeasDiffEndPins, 0.001)
    End If
     Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
     
  'per pin measurement
  TheExec.DataManager.DecomposePinList PpmuMeasure_Pin, Pins(), Pin_Cnt
     
    For Each Pat In patt_ary

        TheExec.Datalog.WriteComment "Pat =  " & Pat
        Call TheHdw.Patterns(Pat).start
        TestSeqNum = 0

        For Each Ts In TestSequenceArray
            
            If MeasOnHalt = True And TestSeqNum = UBound(TestSequenceArray) Then
                Call TheHdw.Digital.Patgen.HaltWait 'Meas after patgen halt
            ElseIf (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If

            TestOptLen = Len(Ts)
            
        Dim p_hexvs As Variant, p_hexvs_ary() As String, PinCnt_hexvs As Long
        Dim p_uvs As Variant, p_uvs_ary() As String, PinCnt_uvs As Long
        Dim TestNum As Long, Cnt1 As Long
        Dim p_uvs_idx As Integer
        Dim p_uvs_str As String
            
            For k = 1 To TestOptLen
                TestOption = Mid(Ts, k, 1)
                

                
                Select Case TestOption
                    Case "V", "v"
                     For Each DUTPin In Pins
                        If (PpmuMeasure_Pin <> "") Then
                            TheHdw.Digital.Pins(DUTPin).Disconnect
                            With TheHdw.PPMU.Pins(DUTPin)   '' make sure which pins
                                If ForceISequenceArray(TestSeqNum) <> "" Then
                                      
                                    .ForceI ForceISequenceArray(TestSeqNum)
                              
                                Else
                                    .ForceI 0
                                End If
                               .Connect
                               .Gate = True
                            End With
                        End If
                        
                        If ForceISequenceArray(TestSeqNum) = "" Then ForceISequenceArray(TestSeqNum) = 0
    
                        '#####################################
                        ' Voltage Measurement
                        '#####################################
                        TestNum = TheExec.sites.Item(0).TestNumber
'                        MeasVoltage.GlobalSort = False
                            DebugPrintFunc_PPMU CStr(DUTPin)
                            MeasVoltage = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 10)
                            TheHdw.PPMU.Pins(DUTPin).ForceI Reset_I_Array(TestSeqNum) 'reset to 0A
                        TheExec.Flow.TestLimit MeasVoltage, CDbl(PPMU_V_LowLimitArray(TestSeqNum)), CDbl(PPMU_V_HighLimitArray(TestSeqNum)), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:="Volt_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceVal:=ForceISequenceArray(TestSeqNum), ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        Next DUTPin
                        
                        
                        'If CurrentJobName Like "*char*" Then Call Char_show(MeasVoltage, CDbl(PPMU_V_HighLimitArray(TestSeqNum)), CDbl(PPMU_V_LowLimitArray(TestSeqNum)), TestNum, "V", "_" + CStr(TestSeqNum))
                        

                    Case "I", "i"
                        '#####################################
                        ' Current Measurement
                        '#####################################
                        If Flag_bypass_meas_I = True Then GoTo Skip_i_meas    'skip I meas in Voh,Vol tests
                        
                        If (PpmuMeasure_Pin <> "") Then
                            TheHdw.Digital.Pins(PpmuMeasure_Pin).Disconnect
                            With TheHdw.PPMU.Pins(PpmuMeasure_Pin)   '' make sure which pins
                                If ForceVSequenceArray(TestSeqNum) <> "" Then
                                
                                If Irange <> 0 Then
                                .ForceV ForceVSequenceArray(TestSeqNum), Irange
                                   Else
                                      
                                    .ForceV ForceVSequenceArray(TestSeqNum)
                                    End If
                                Else
                                    .ForceV 0
                                End If
                               .Connect
                               .Gate = True
                            End With
                        End If
                        
                        If ForceVSequenceArray(TestSeqNum) = "" Then ForceVSequenceArray(TestSeqNum) = 0
    
                        '#####################################
                        ' Voltage Measurement
                        '#####################################
                        TestNum = TheExec.sites.Item(0).TestNumber
'                        'MeasCurrent.GlobalSort = False
                        For Each DUTPin In Pins
                            DebugPrintFunc_PPMU CStr(DUTPin)
                            MeasCurrent = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 10)
                            TheHdw.PPMU.Pins(DUTPin).ForceV Reset_V_Array(TestSeqNum) 'reset to 0V
                            If LCase(TestOption) Like "i" Then TheExec.Flow.TestLimit MeasCurrent, CDbl(PPMU_I_LowLimitArray(TestSeqNum)), CDbl(PPMU_I_HighLimitArray(TestSeqNum)), scaletype:=scaleNone, Unit:=unitAmp, Tname:="Curr_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceVal:=ForceVSequenceArray(TestSeqNum), ForceUnit:=unitVolt, ForceResults:=tlForceNone
                        'If LCase(TestOption) Like "r" Then TheExec.Flow.TestLimit MeasCurrent.Math.Divide(ForceVSequenceArray(TestSeqNum)).Invert, PPMU_R_LowLimit, PPMU_R_HighLimit, ScaleType:=scaleNone, unit:=unitCustom, tname:="Imp_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), forceVal:=ForceVSequenceArray(TestSeqNum), forceunit:=unitVolt, customUnit:="ohm", forceResults:=tlForceNone
                        Next DUTPin
                        
                        'If CurrentJobName Like "*char*" Then Call Char_show(MeasVoltage, CDbl(PPMU_I_HighLimitArray(TestSeqNum)), CDbl(PPMU_I_LowLimitArray(TestSeqNum)), TestNum, "I", "_" + CStr(TestSeqNum))
                    Case "R", "r"
                        '#####################################
                        ' Current Measurement
                        '#####################################
                        If (PpmuMeasure_Pin <> "") Then
                            TheHdw.Digital.Pins(PpmuMeasure_Pin).Disconnect
                            With TheHdw.PPMU.Pins(PpmuMeasure_Pin)   '' make sure which pins
                                If ForceVSequenceArray(TestSeqNum) <> "" Then
                                    .ForceV ForceVSequenceArray(TestSeqNum), 0.02
                                Else
                                    .ForceV 0
                                End If
                               .Connect
                               .Gate = True
                            End With
                        End If
                        
                        If ForceVSequenceArray(TestSeqNum) = "" Then ForceVSequenceArray(TestSeqNum) = 0
    
                        '#####################################
                        ' Voltage Measurement
                        '#####################################
                        TestNum = TheExec.sites.Item(0).TestNumber
'                        'MeasCurrent.GlobalSort = False
                        For Each DUTPin In Pins
                            DebugPrintFunc_PPMU CStr(DUTPin)
                            MeasCurrent = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 10)
                            TheHdw.PPMU.Pins(DUTPin).ForceV Reset_V_Array(TestSeqNum) 'reset to 0V
                            'If LCase(TestOption) Like "i" Then TheExec.Flow.TestLimit MeasCurrent, PPMU_I_LowLimit, PPMU_I_HighLimit, ScaleType:=scaleNone, unit:=unitAmp, tname:="Curr_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), forceVal:=ForceVSequenceArray(TestSeqNum), forceunit:=unitVolt, forceResults:=tlForceNone
                            If LCase(TestOption) Like "r" Then TheExec.Flow.TestLimit MeasCurrent.Math.Divide(ForceVSequenceArray(TestSeqNum)).Invert, CDbl(PPMU_R_LowLimitArray(TestSeqNum)), CDbl(PPMU_R_HighLimitArray(TestSeqNum)), scaletype:=scaleNone, Unit:=unitCustom, Tname:="Imp_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceVal:=ForceVSequenceArray(TestSeqNum), ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                        Next DUTPin
                        'If CurrentJobName Like "*char*" Then Call Char_show(MeasVoltage, CDbl(PPMU_I_HighLimitArray(TestSeqNum)), CDbl(PPMU_I_LowLimitArray(TestSeqNum)), TestNum, "R", "_" + CStr(TestSeqNum))
                                  

                    Case Else
                         TheExec.Datalog.WriteComment "Error Test Option, please select V,I,R"
                End Select
Skip_i_meas:
                
            Next k
            TestSeqNum = TestSeqNum + 1
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
            End If

        Next Ts
        
        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & pat_count & "): " & Pat & ""
        
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        pat_count = pat_count + 1
        
        AllSitePass = True
        For Each site In TheExec.sites
            BurstResult(site) = 1
            
            If (TheHdw.Digital.Patgen.PatternBurstPassed(site) = False) Then
                'TheExec.Datalog.WriteComment vbCrLf & Pat & "_" & StopLabel & vbTab & "Run " & i & " : Fail."
                BurstResult(site) = 0
                AllSitePass = False
            End If
        Next site
        'If AllSitePass = True Then Exit For

        TheExec.Flow.TestLimit BurstResult, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="Burst Result"  'BurstResult=1:Pass

'debug!!!
        'If (AllSitePass = False) Then GoTo End_Meas
        
    Next Pat
    

    
    If PpmuMeasure_Pin <> "" Then
        TheHdw.PPMU.Pins(PpmuMeasure_Pin).Disconnect
        'TheHdw.Digital.Pins(PpmuMeasure_Pin).Connect
    End If
    
    If (DisableConnectPins <> "") Then
        TheHdw.Digital.Pins(DisableComparePins).Connect
    End If
    
    If DisableComparePins <> "" Then
        TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    End If
    
End_Meas:
    
    DebugPrintFunc patset.Value  ' print all debug information

Exit Function


errHandler:
    TheExec.Datalog.WriteComment "error in Meas_FreqVoltCurr_Univeral_func"
    If AbortTest Then Exit Function Else Resume Next
  
End Function




Public Function Meas_VOH_MeasVIR_delta_Univeral_func_DC(patset As Pattern, DisableComparePins As PinList, DisableConnectPins As PinList, TestSequence As String, CPUA_Flag_In_Pat As Boolean, PpmuMeasure_Pin As PinList, FreqMeasSingleEndPins As PinList, FreqMeasDiffEndPins As PinList, MeasureI_pin As PinList, _
                                                 MeasOnHalt As Boolean, Flow_Limit_Result As tlLimitForceResults, PPMU_V_LowLimit As String, PPMU_V_HighLimit As String, PPMU_I_LowLimit As String, PPMU_I_HighLimit As String, PPMU_R_LowLimit As String, PPMU_R_HighLimit As String, Reset_V As String, Reset_I As String _
                                                , I_LowLimit As Double, I_HighLimit As Double, Freq_LowLimit As Double, Freq_HighLimit As Double, ForceV As String, ForceI As String, PortName As String, Optional Flag_Meas_VOH_Freq As Boolean = False, Optional DisableClock As Boolean = False, Optional Flag_bypass_meas_I As Boolean = False, Optional Irange As Double) As Long
' EDITFORMAT1 1,,Pattern,Group1,,patset|2,,PinList,,,DisableComparePins|3,,PinList,,,DisableConnectPins|4,,String,,,TestSequence|5,,Boolean,,,CPUA_Flag_In_Pat|6,,PinList,,,PpmuMeasure_Pin|7,,PinList,,,FreqMeasSingleEndPins|8,,PinList,,,FreqMeasDiffEndPins|9,,PinList,,,MeasureI_pin|10,,Boolean,,,MeasOnHalt|11,,tlLimitForceResults,,,Flow_Limit_Result|12,,Double,,,PPMU_V_LowLimit|13,,Double,,,PPMU_V_HighLimit|14,,Double,,,PPMU_I_LowLimit|15,,Double,,,PPMU_I_HighLimit|16,,Double,,,PPMU_R_LowLimit|17,,Double,,,PPMU_R_HighLimit|18,,Double,,,I_LowLimit|19,,Double,,,I_HighLimit|20,,Double,,,Freq_LowLimit|21,,Double,,,Freq_HighLimit|22,,String,,,ForceV|23,,String,,,ForceI|24,,String,,,PortName|25,,Boolean,,,Flag_Meas_VOH_Freq|26,,Boolean,,,DisableClock|27,,Boolean,,,Flag_bypass_meas_I
'Parameter information     **2013/04/12 by JT
'Patset : Test Pattern
'DisableComparePins     :   Disable Pin Compare H/L.
'TestSequence           :   Decide to test which function and sequence, "v,i,vi,f" means test sequence will be
'                           1. Meas voltagee
'                           2. Meas voltage and current at the same CPU flag loop
'                           3. Meas frequence
'CPUA_Flag_In_Pat       :   If CPUA flag in Pattern?
'PpmuMeasure_Pin       :   Meas voltage pin
'FreqCtrMeasurePins     :   Meas frequence pin
'MeasureI_pin           :   Meas current pin
'MeasFreqPinType        :   Meas frequence pin is different or single end.
'MeasOnHalt             :   If need to measure any thing after Pattern halt


    Dim i As Integer ' for the use of loading pattern
    
    Dim pat_count As Long
    Dim MeasFreqSingle As New PinListData
    Dim MeasFreqDifferential As New PinListData
    
    Dim MeasVoltage As New PinListData
    Dim MeasCurrent As New PinListData
    Dim MeasCurrent1 As New PinListData
    Dim MeasCurr As New PinListData
    Dim z As Integer
    Dim j As Long
    Dim k As Long
    Dim freq As Double
    Dim freq_limit_upper As Double
    Dim freq_limit_lower As Double
    Dim Status As Boolean
    Dim p As Long
    Dim CPUA_Flag_Cnt As Integer
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, ForceISequenceArray() As String, ForceVSequenceArray() As String
    Dim PPMU_V_LowLimitArray() As String, PPMU_I_LowLimitArray() As String, PPMU_R_LowLimitArray() As String
    Dim PPMU_V_HighLimitArray() As String, PPMU_I_HighLimitArray() As String, PPMU_R_HighLimitArray() As String
    Dim Reset_V_Array() As String, Reset_I_Array() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim Pins() As String, Pin_Cnt As Long
    Dim DUTPin As Variant
    Dim patout As String
    Dim patt_ary() As String
    Dim Pat As Variant
    Dim site As Variant
    
    TestSequenceArray = Split(TestSequence, ",")
    ForceISequenceArray = Split(ForceI, ",")
    ForceVSequenceArray = Split(ForceV, ",")
    PPMU_V_LowLimitArray = Split(PPMU_V_LowLimit, ",")
    PPMU_V_HighLimitArray = Split(PPMU_V_HighLimit, ",")
    PPMU_I_LowLimitArray = Split(PPMU_I_LowLimit, ",")
    PPMU_I_HighLimitArray = Split(PPMU_I_HighLimit, ",")
    PPMU_R_LowLimitArray = Split(PPMU_R_LowLimit, ",")
    PPMU_R_HighLimitArray = Split(PPMU_R_HighLimit, ",")
    Reset_V_Array = Split(Reset_V, ",")
    Reset_I_Array = Split(Reset_I, ",")
    
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    On Error GoTo errHandler
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    
    If (DisableConnectPins <> "") Then
        TheHdw.Digital.Pins(DisableComparePins).Disconnect
    End If
    
    If (DisableComparePins <> "") Then
        TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    End If
    TheHdw.Patterns(patset).Load
    Call GetPatListFromPatternSet(patset.Value, patt_ary, pat_count)

    If FreqMeasSingleEndPins <> "" Then
        Call Freq_MeasFreqSetup(FreqMeasSingleEndPins, 0.001)
    End If
    If FreqMeasDiffEndPins <> "" Then
        Call Freq_MeasFreqSetup(FreqMeasDiffEndPins, 0.001)
    End If
     Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
     
  'per pin measurement
  TheExec.DataManager.DecomposePinList PpmuMeasure_Pin, Pins(), Pin_Cnt
     
    For Each Pat In patt_ary

        TheExec.Datalog.WriteComment "Pat =  " & Pat
        Call TheHdw.Patterns(Pat).start
        TestSeqNum = 0

        For Each Ts In TestSequenceArray
            
            If MeasOnHalt = True And TestSeqNum = UBound(TestSequenceArray) Then
                Call TheHdw.Digital.Patgen.HaltWait 'Meas after patgen halt
            ElseIf (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If

            TestOptLen = Len(Ts)
            
        Dim p_hexvs As Variant, p_hexvs_ary() As String, PinCnt_hexvs As Long
        Dim p_uvs As Variant, p_uvs_ary() As String, PinCnt_uvs As Long
        Dim TestNum As Long, Cnt1 As Long
        Dim p_uvs_idx As Integer
        Dim p_uvs_str As String
            
            For k = 1 To TestOptLen
                TestOption = Mid(Ts, k, 1)
                

                
                Select Case TestOption
                    Case "V", "v"
    
                        If (PpmuMeasure_Pin <> "") Then
                            TheHdw.Digital.Pins(PpmuMeasure_Pin).Disconnect
                            With TheHdw.PPMU.Pins(PpmuMeasure_Pin)   '' make sure which pins
                                If ForceISequenceArray(TestSeqNum) <> "" Then
                                    .ForceI ForceISequenceArray(TestSeqNum)
                                Else
                                    .ForceI 0
                                End If
                               .Connect
                               .Gate = True
                            End With
                        End If
                        
                        If ForceISequenceArray(TestSeqNum) = "" Then ForceISequenceArray(TestSeqNum) = 0
    
                        '#####################################
                        ' Voltage Measurement
                        '#####################################
                        TestNum = TheExec.sites.Item(0).TestNumber
'                        MeasVoltage.GlobalSort = False
                        For Each DUTPin In Pins
                            DebugPrintFunc_PPMU CStr(DUTPin)
                            MeasVoltage = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 10)
                            TheHdw.PPMU.Pins(DUTPin).ForceI Reset_I_Array(TestSeqNum) 'reset to 0A
                        TheExec.Flow.TestLimit MeasVoltage, CDbl(PPMU_V_LowLimitArray(TestSeqNum)), CDbl(PPMU_V_HighLimitArray(TestSeqNum)), scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:="Volt_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceVal:=ForceISequenceArray(TestSeqNum), ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        Next DUTPin
                        
                        
                        'If CurrentJobName Like "*char*" Then Call Char_show(MeasVoltage, CDbl(PPMU_V_HighLimitArray(TestSeqNum)), CDbl(PPMU_V_LowLimitArray(TestSeqNum)), TestNum, "V", "_" + CStr(TestSeqNum))
                        

                    Case "I", "i"
                        '#####################################
                        ' Current Measurement
                        '#####################################
                        If Flag_bypass_meas_I = True Then GoTo Skip_i_meas    'skip I meas in Voh,Vol tests
                        
                        If (PpmuMeasure_Pin <> "") Then
                            TheHdw.Digital.Pins(PpmuMeasure_Pin).Disconnect
                            With TheHdw.PPMU.Pins(PpmuMeasure_Pin)   '' make sure which pins
                                If ForceVSequenceArray(TestSeqNum) <> "" Then
                                    .ForceV ForceVSequenceArray(TestSeqNum), 0.000002
                                    
                                Else
                                    .ForceV 0
                                End If
                               .Connect
                               .Gate = True
                            End With
                        End If
                        
                        If ForceVSequenceArray(TestSeqNum) = "" Then ForceVSequenceArray(TestSeqNum) = 0
    
                        '#####################################
                        ' Voltage Measurement
                        '#####################################
                        TestNum = TheExec.sites.Item(0).TestNumber
'                        'MeasCurrent.GlobalSort = False
                        For Each DUTPin In Pins
                            DebugPrintFunc_PPMU CStr(DUTPin)
                            MeasCurrent = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 10)
                            TheHdw.PPMU.Pins(DUTPin).ForceV Reset_V_Array(TestSeqNum) 'reset to 0V
                            If LCase(TestOption) Like "i" Then TheExec.Flow.TestLimit MeasCurrent, CDbl(PPMU_I_LowLimitArray(TestSeqNum)), CDbl(PPMU_I_HighLimitArray(TestSeqNum)), scaletype:=scaleNone, Unit:=unitAmp, Tname:="Curr_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceVal:=ForceVSequenceArray(TestSeqNum), ForceUnit:=unitVolt, ForceResults:=tlForceNone
                        'If LCase(TestOption) Like "r" Then TheExec.Flow.TestLimit MeasCurrent.Math.Divide(ForceVSequenceArray(TestSeqNum)).Invert, PPMU_R_LowLimit, PPMU_R_HighLimit, ScaleType:=scaleNone, unit:=unitCustom, tname:="Imp_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), forceVal:=ForceVSequenceArray(TestSeqNum), forceunit:=unitVolt, customUnit:="ohm", forceResults:=tlForceNone
                        Next DUTPin
                        
                        'If CurrentJobName Like "*char*" Then Call Char_show(MeasVoltage, CDbl(PPMU_I_HighLimitArray(TestSeqNum)), CDbl(PPMU_I_LowLimitArray(TestSeqNum)), TestNum, "I", "_" + CStr(TestSeqNum))
                    Case "R", "r"
                        '#####################################
                        ' Current Measurement
                        '#####################################
                         For Each DUTPin In Pins
                        If (PpmuMeasure_Pin <> "") Then
                            TheHdw.Digital.Pins(PpmuMeasure_Pin).Disconnect
                            With TheHdw.PPMU.Pins(PpmuMeasure_Pin)   '' make sure which pins
                                If ForceVSequenceArray(TestSeqNum) <> "" Then
                                    .ForceV 0.7, Irange
                                Else
                                    .ForceV 0
                                End If
                               .Connect
                               .Gate = True
                            End With
                        End If
                        
                        If ForceVSequenceArray(TestSeqNum) = "" Then ForceVSequenceArray(TestSeqNum) = 0
    
                        '#####################################
                        ' Voltage Measurement
                        '#####################################
                        TestNum = TheExec.sites.Item(0).TestNumber
'                        'MeasCurrent.GlobalSort = False
                        
                       
                            DebugPrintFunc_PPMU CStr(DUTPin)
                            MeasCurrent = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 10)
                            
                            TheHdw.PPMU.Pins(DUTPin).ForceV Reset_V_Array(TestSeqNum) 'reset to 0V
                            'If LCase(TestOption) Like "i" Then TheExec.Flow.TestLimit MeasCurrent, PPMU_I_LowLimit, PPMU_I_HighLimit, ScaleType:=scaleNone, unit:=unitAmp, tname:="Curr_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), forceVal:=ForceVSequenceArray(TestSeqNum), forceunit:=unitVolt, forceResults:=tlForceNone
                           ' If LCase(TestOption) Like "r" Then TheExec.Flow.TestLimit MeasCurrent.Math.Divide(ForceVSequenceArray(TestSeqNum)).Invert, CDbl(PPMU_R_LowLimitArray(TestSeqNum)), CDbl(PPMU_R_HighLimitArray(TestSeqNum)), ScaleType:=scaleNone, unit:=unitCustom, tname:="Imp_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), forceVal:=ForceVSequenceArray(TestSeqNum), forceunit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                      
                        
                       
                            TheHdw.Digital.Pins(PpmuMeasure_Pin).Disconnect
                            With TheHdw.PPMU.Pins(PpmuMeasure_Pin)   '' make sure which pins
                                If ForceVSequenceArray(TestSeqNum) <> "" Then
                                    .ForceV 1.1, Irange
                                Else
                                    .ForceV 0
                                End If
                               .Connect
                               .Gate = True
                            End With
                       
                        
                        If ForceVSequenceArray(TestSeqNum) = "" Then ForceVSequenceArray(TestSeqNum) = 0
    
                        '#####################################
                        ' Voltage Measurement
                        '#####################################
                        TestNum = TheExec.sites.Item(0).TestNumber
'                        'MeasCurrent.GlobalSort = False
                        
                       
                            DebugPrintFunc_PPMU CStr(DUTPin)
                            MeasCurrent1 = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 10)
                            
                            TheHdw.PPMU.Pins(DUTPin).ForceV Reset_V_Array(TestSeqNum) 'reset to 0V
                            'If LCase(TestOption) Like "i" Then TheExec.Flow.TestLimit MeasCurrent, PPMU_I_LowLimit, PPMU_I_HighLimit, ScaleType:=scaleNone, unit:=unitAmp, tname:="Curr_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), forceVal:=ForceVSequenceArray(TestSeqNum), forceunit:=unitVolt, forceResults:=tlForceNone
                           ' If LCase(TestOption) Like "r" Then TheExec.Flow.TestLimit MeasCurrent.Math.Divide(ForceVSequenceArray(TestSeqNum)).Invert, CDbl(PPMU_R_LowLimitArray(TestSeqNum)), CDbl(PPMU_R_HighLimitArray(TestSeqNum)), ScaleType:=scaleNone, unit:=unitCustom, tname:="Imp_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), forceVal:=ForceVSequenceArray(TestSeqNum), forceunit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                            
                            
                      
                                                
                       ' (MeasCurrent1.Pins(dutpin).Value -MeasCurrent1.Pins(dutpin).Value)/0.2
                        
                            'If LCase(TestOption) Like "i" Then TheExec.Flow.TestLimit MeasCurrent, PPMU_I_LowLimit, PPMU_I_HighLimit, ScaleType:=scaleNone, unit:=unitAmp, tname:="Curr_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), forceVal:=ForceVSequenceArray(TestSeqNum), forceunit:=unitVolt, forceResults:=tlForceNone
                       
                            If LCase(TestOption) Like "r" Then TheExec.Flow.TestLimit MeasCurrent1.Pins(DUTPin).Subtract(MeasCurrent.Pins(DUTPin)).Divide(0.4).Invert, CDbl(PPMU_R_LowLimitArray(TestSeqNum)), CDbl(PPMU_R_HighLimitArray(TestSeqNum)), scaletype:=scaleNone, Unit:=unitCustom, Tname:="Imp_meas" + "_" + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PatExculdePath(Pat), ForceVal:=ForceVSequenceArray(TestSeqNum), ForceUnit:=unitVolt, customUnit:="ohm", ForceResults:=tlForceNone
                        
                         Next DUTPin
                        'If CurrentJobName Like "*char*" Then Call Char_show(MeasVoltage, CDbl(PPMU_I_HighLimitArray(TestSeqNum)), CDbl(PPMU_I_LowLimitArray(TestSeqNum)), TestNum, "R", "_" + CStr(TestSeqNum))
                        
                        'reset ppmu and return pins to digital channel
                        With TheHdw.PPMU.Pins(PpmuMeasure_Pin)   '' make sure which pins
                            .ForceV 0
                            .Connect
                            .Gate = True
                        End With
                        TheHdw.PPMU.Pins(PpmuMeasure_Pin).Disconnect
                        TheHdw.Digital.Pins(PpmuMeasure_Pin).Connect
                    Case Else
                         TheExec.Datalog.WriteComment "Error Test Option, please select V,I,R"
                End Select
Skip_i_meas:
                
            Next k
            TestSeqNum = TestSeqNum + 1
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
            End If

        Next Ts
        
        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & pat_count & "): " & Pat & ""
        
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        pat_count = pat_count + 1
        
        AllSitePass = True
        For Each site In TheExec.sites
            BurstResult(site) = 1
            
            If (TheHdw.Digital.Patgen.PatternBurstPassed(site) = False) Then
                'TheExec.Datalog.WriteComment vbCrLf & Pat & "_" & StopLabel & vbTab & "Run " & i & " : Fail."
                BurstResult(site) = 0
                AllSitePass = False
            End If
        Next site
        'If AllSitePass = True Then Exit For

        TheExec.Flow.TestLimit BurstResult, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="Burst Result"  'BurstResult=1:Pass

'debug!!!
        'If (AllSitePass = False) Then GoTo End_Meas
        
    Next Pat
    

    
    If PpmuMeasure_Pin <> "" Then
        TheHdw.PPMU.Pins(PpmuMeasure_Pin).Disconnect
        'TheHdw.Digital.Pins(PpmuMeasure_Pin).Connect
    End If
    
    If (DisableConnectPins <> "") Then
        TheHdw.Digital.Pins(DisableComparePins).Connect
    End If
    
    If DisableComparePins <> "" Then
        TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    End If
    
End_Meas:
    
    DebugPrintFunc patset.Value  ' print all debug information

Exit Function


errHandler:
    TheExec.Datalog.WriteComment "error in Meas_FreqVoltCurr_Univeral_func"
    If AbortTest Then Exit Function Else Resume Next
  
End Function


Public Function Interpose_TurnOff_CPUFLagA(argc As Long, argv() As String)
    Call TheHdw.Digital.Patgen.Continue(0, cpuA)
End Function

Public Function Meas_VIHL_VOHL_Universal_Functional(patset As Pattern, DisableComparePins As PinList, DisableConnectPins As PinList, CPUA_Flag_In_Pat As Boolean, digital_pins As PinList, _
            V_VOL As String, V_VOH As String, Force_Iol_Ioh As String, _
            V_DriveLow As String, V_DriveHigh As String, _
            Optional CPUA_Flag_In_Pat_Times As String) As Long
            
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim pat_count As Long, Pin As Variant

    Dim freq_limit_upper As Double
    Dim freq_limit_lower As Double
    Dim Status As Boolean

    Dim CPUA_Flag_Cnt As Integer
    Dim TestOptLen As Integer

    Dim Pins() As String
    Dim Pin_Cnt As Long

    Dim patt_ary() As String
    Dim Pat As Variant
    Dim site As Variant
    
    On Error GoTo errHandler
    
''    TheExec.DataManager.DecomposePinList PpmuMeasureV_Pin, Pins(), pin_cnt
           
     ''20151105 Check CPUA subroutine times
     Dim l_CPUA_Flag_In_Pat_Times   As Long
     If CPUA_Flag_In_Pat_Times <> "" Then
         l_CPUA_Flag_In_Pat_Times = CDbl(CPUA_Flag_In_Pat_Times)
     End If
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Force_Iol_Ioh <> "" Then
        TheHdw.Digital.Pins(digital_pins).Levels.Value(chIol) = CDbl(Force_Iol_Ioh)
        TheHdw.Digital.Pins(digital_pins).Levels.Value(chIoh) = CDbl(Force_Iol_Ioh) * (-1)
    End If
        
    If V_VOL <> "" Then
        TheHdw.Digital.Pins(digital_pins).Levels.Value(chVol) = CDbl(V_VOL)
    End If
        
    If V_VOH <> "" Then
        TheHdw.Digital.Pins(digital_pins).Levels.Value(chVoh) = CDbl(V_VOH)
    End If

    If V_DriveLow <> "" Then
        TheHdw.Digital.Pins(digital_pins).Levels.Value(chVil) = CDbl(V_DriveLow)
    End If
    
    If V_DriveHigh <> "" Then
        TheHdw.Digital.Pins(digital_pins).Levels.Value(chVih) = CDbl(V_DriveHigh)
    End If
        
    If DisableComparePins <> "" Then
        TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    End If
    
    If DisableConnectPins <> "" Then
        TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    End If
        
    TheHdw.Patterns(patset).Load
    Status = GetPatListFromPatternSet(patset.Value, patt_ary, pat_count)
    
    pat_count = 1
    
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    
    TheHdw.Wait 0.001
         
    For Each Pat In patt_ary
    
        Call TheHdw.Patterns(Pat).start
            
        If (CPUA_Flag_In_Pat) Then
            For i = 0 To l_CPUA_Flag_In_Pat_Times - 1
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)  'Meas during CPUA loop
                
                Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
            Next i
            
        Else
            Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
        End If

        pat_count = pat_count + 1
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
    Next Pat
    
    Call DC_Func_WriteFuncResult(False)
                          
    If DisableComparePins <> "" Then
        TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    End If
        
    If (DisableConnectPins <> "") Then
        TheHdw.Digital.Pins(DisableConnectPins).Connect
    End If
    
    DebugPrintFunc patset.Value  ' print all debug information

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Meas_VOHL_Univeral_func_Parallel is error "
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Cal_Hysteresis(powerPin As String)

On Error GoTo errHandler:
Dim Subroutine_Name As String
Dim site As Variant
Dim PowerVolt As Double, i As Long
Dim V_hysteresis As Double
Dim vpwr As Double
Dim PinName As String
Subroutine_Name = "Cal_Hysteresis"

vpwr = TheHdw.DCVS.Pins(powerPin).Voltage.Main

For i = 0 To UBound(GPIO_Vih_Vil)
    For Each site In TheExec.sites.Active
        If GPIO_Vih_Vil(i).Vih(site) <> -999 And GPIO_Vih_Vil(i).Vil(site) <> -999 Then
            V_hysteresis = GPIO_Vih_Vil(i).Vih(site) - GPIO_Vih_Vil(i).Vil(site)
            PinName = "GPIO" & CStr(i)
            TheExec.Flow.TestLimit resultVal:=V_hysteresis, lowVal:=vpwr * 0.1, Tname:="Hystere", PinName:=PinName, Unit:=unitVolt
        End If
    Next site
Next i

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in " & Subroutine_Name
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Judge_GPIO_Vil(powerPin As String)

On Error GoTo errHandler:
Dim Subroutine_Name As String
Dim site As Variant
Dim PowerVolt As Double, i As Long
Dim V_GPIO_Vil As Double
Dim vpwr As Double
Dim PinName As String
Subroutine_Name = "Judge_GPIO_Vil"

vpwr = TheHdw.DCVS.Pins(powerPin).Voltage.Main

For i = 0 To UBound(GPIO_Vih_Vil)
    For Each site In TheExec.sites.Active
        If GPIO_Vih_Vil(i).Vil(site) <> -999 Then
            V_GPIO_Vil = GPIO_Vih_Vil(i).Vil(site)
            PinName = "GPIO" & CStr(i)
            TheExec.Flow.TestLimit resultVal:=V_GPIO_Vil, lowVal:=vpwr * 0.3, Tname:="GPIO_Vil", PinName:=PinName, Unit:=unitVolt
        End If
    Next site
Next i

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in " & Subroutine_Name
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Judge_GPIO_Vih(powerPin As String)

On Error GoTo errHandler:
Dim Subroutine_Name As String
Dim site As Variant
Dim PowerVolt As Double, i As Long
Dim V_GPIO_Vih As Double
Dim vpwr As Double
Dim PinName As String

Subroutine_Name = "Judge_GPIO_Vih"

vpwr = TheHdw.DCVS.Pins(powerPin).Voltage.Main

For i = 0 To UBound(GPIO_Vih_Vil)
    For Each site In TheExec.sites.Active
        If GPIO_Vih_Vil(i).Vih(site) <> -999 Then
            V_GPIO_Vih = GPIO_Vih_Vil(i).Vih(site)
            PinName = "GPIO" & CStr(i)
            TheExec.Flow.TestLimit resultVal:=V_GPIO_Vih, hiVal:=vpwr * 0.7, Tname:="GPIO_Vih", PinName:=PinName, Unit:=unitVolt
        End If
    Next site
Next i

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in " & Subroutine_Name
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function HiZ_Leakage_Power(patset As Pattern, _
                                  ForceV_IiH As Double, _
                                  ForceV_IiL As Double, _
                                  I_Meas_Range As Double, _
                                  leakage_pins As PinList, _
                                  hi_limit As Double, _
                                  lo_limit As Double, _
                                  MeasureI_pin As PinList, _
                                  I_LowLimit As Double, _
                                  I_HighLimit As Double, _
                                  Irange As String, _
                                  Port_name As String, _
                                  CPUA_Flag_In_Pat As Boolean, _
                                  Optional DisableClock As Boolean = False, _
                                  Optional Flag_Low As Boolean = False, _
                                  Optional Init_H_pin As PinList, _
                                  Optional Init_L_pin As PinList)

    Dim site As Variant
    'Dim SeqLeakPins As String
    Dim PinArr() As String, PinCount As Long, i As Long
    Dim p As Variant
    Dim MeasVal As New PinListData
    Dim TestNum As Long
    Dim Tname As String
    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    
    Dim MeasCurr As New PinListData
    Dim TestSeqNum As Integer
    
    On Error GoTo errHandler
    
    Call TheHdw.Digital.ApplyLevelsTiming(True, True, False, tlPowered, leakage_pins, , leakage_pins)
    
    '============= Init State ==============
    If Init_H_pin <> "" Then TheHdw.Digital.Pins(Init_H_pin).InitState = chInitHi
    If Init_L_pin <> "" Then TheHdw.Digital.Pins(Init_L_pin).InitState = chInitLo
    '============= Init State ==============
    
    TheHdw.Patterns(patset).Load
    ''Call TheHdw.Patterns(patset).Start
    ''TheHdw.digital.Patgen.HaltWait
    
    'Check if pattern passed
    For i = 0 To 1
        Call TheHdw.Patterns(patset).start
        TheHdw.Digital.Patgen.HaltWait
        AllSitePass = True
        For Each site In TheExec.sites
            BurstResult(site) = 1
        Next site
        
        For Each site In TheExec.sites
            If (TheHdw.Digital.Patgen.PatternBurstPassed(site) = False) Then
                TheExec.Datalog.WriteComment vbCrLf & patset & "_" & vbTab & "Run " & i & " : Fail."
                BurstResult(site) = 0
                AllSitePass = False
            End If
        Next site
        If AllSitePass = True Then Exit For
    Next i
    
    TheExec.Flow.TestLimit BurstResult, 1, 1, tlSignGreaterEqual, tlSignLessEqual, Tname:="Burst_Result" 'BurstResult=1:Pass
    TheExec.Datalog.WriteComment ""
    
    TheHdw.Digital.Pins(leakage_pins).InitState = chInitoff
    
    TheHdw.PPMU.Pins(leakage_pins).ForceV (0)
    TheHdw.PPMU.Pins(leakage_pins).Gate = tlOff
    TheHdw.PPMU.Pins(leakage_pins).Disconnect

    ''''''use the "theexec.DataManager.DecomposePinList" to serialize the pins to be tested sequentially'''''
    TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount
'    DCVS_Trim_NC_Pin PinArr(), PinCount
    
    ''High
    
    'If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
    TheHdw.Digital.Pins(leakage_pins).Disconnect
    
    With TheHdw.PPMU(leakage_pins)
        .Connect
        .Gate = tlOn
        .ForceV ForceV_IiH, I_Meas_Range
         TheHdw.Wait 0.001
         DebugPrintFunc_PPMU leakage_pins.Value
         MeasVal = .Read(tlPPMUReadMeasurements)
         
        .ForceV (0)
        .Gate = tlOff
        .Disconnect
        
    End With
        
    'TheHdw.digital.Pins(leakage_pins).Connect 'connect the tested pin back to the PE
    
    'offline mode simulation
    If TheExec.TesterMode = testModeOffline Then

        For Each site In TheExec.sites
            For Each p In PinArr()
                If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.Pins(p).Value = 5 * uA + Rnd() * 0.1 * uA
            Next p
        Next site
    End If
  
  
    TheExec.Flow.TestLimit resultVal:=MeasVal, lowVal:=lo_limit, hiVal:=hi_limit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:="Leak_Hi", ForceVal:=ForceV_IiH, ForceUnit:=unitVolt, ForceResults:=tlForceNone
    'TheExec.Flow.TestLimit resultval:=PPMUMeasure.Pins(DUTPin), lowval:=LowLimit, hival:=HiLimit, ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", tname:=tname, forceVal:=force_i, forceunit:=unitAmp, forceResults:=tlForceNone

    'DCVS meas
    TestNum = TheExec.sites.Item(0).TestNumber
    Call DCVS_Set_Meter_Range(MeasureI_pin, Irange)
    TheHdw.Digital.Pins(leakage_pins).Connect
    TheHdw.Wait 0.01     'new add
    
    'DCVS_MeterRead DCVS_UVS256, CStr(MeasureI_pin), 10, MeasCurr
    
    'TheExec.Flow.TestLimit resultVal:=MeasCurr, LowVal:=I_LowLimit, HiVal:=I_HighLimit, ScaleType:=scaleNone, unit:=unitAmp, formatstr:="%.3f", Tname:="Power_Curr_meas_Hi", forceunit:=unitVolt, ForceResults:=tlForceFlow

    'TheExec.Flow.TestLimit resultVal:=MeasCurr, ScaleType:=scaleNone, unit:=unitAmp, formatstr:="%.3f", Tname:="Power_Curr_meas_Hi", forceunit:=unitVolt, ForceResults:=tlForceFlow
           

    '' Low
    
    If Flag_Low = True Then
        ''If TheExec.DataManager.ChannelType(PinArr(i)) <> "N/C" Then
        TheHdw.Digital.Pins(leakage_pins).Disconnect
    
        With TheHdw.PPMU(leakage_pins)
            .Connect
            .Gate = tlOn
            .ForceV ForceV_IiL, I_Meas_Range
             TheHdw.Wait 0.001
             DebugPrintFunc_PPMU leakage_pins.Value
             MeasVal = .Read(tlPPMUReadMeasurements)
             
            .ForceV (0)
            .Gate = tlOff
            .Disconnect
        End With
        
    
    
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
    ''        TheExec.DataManager.DecomposePinList leakage_pins, PinArr(), PinCount
    ''        DCVS_Trim_NC_Pin PinArr(), PinCount
            For Each site In TheExec.sites
                For Each p In PinArr()
                    If TheExec.DataManager.ChannelType(p) <> "N/C" Then MeasVal.Pins(p).Value = -(5 * uA + Rnd() * 0.1 * uA)
                Next p
            Next site
        End If
    
        TheExec.Flow.TestLimit resultVal:=MeasVal, lowVal:=lo_limit, hiVal:=hi_limit, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:="Leak_Lo", ForceVal:=ForceV_IiL, ForceUnit:=unitVolt, ForceResults:=tlForceNone
    
        TheHdw.Digital.Pins(leakage_pins).Connect 'connect the tested pin back to the PE
            
    End If
    
            '#####################################
            ' Current Measurement_Power Pin
            '#####################################
    


        
        'If DisableClock = True Then FreeRunClk_Disable (PortName)
        'DCVS meas
        TestNum = TheExec.sites.Item(0).TestNumber
        Call DCVS_Set_Meter_Range(MeasureI_pin, Irange)
        TheHdw.Digital.Pins(leakage_pins).Connect
        'TheHdw.Wait 0.01     'new add
        
        'DCVS_MeterRead DCVS_UVS256, CStr(MeasureI_pin), 10, MeasCurr
       
        'TheExec.Flow.TestLimit resultVal:=MeasCurr, LowVal:=I_LowLimit, HiVal:=I_HighLimit, ScaleType:=scaleNone, unit:=unitAmp, formatstr:="%.3f", Tname:="Power_Curr_meas_Lo", forceunit:=unitVolt, ForceResults:=tlForceFlow
        'TheExec.Flow.TestLimit resultVal:=MeasCurr, ScaleType:=scaleNone, unit:=unitAmp, formatstr:="%.3f", Tname:="Power_Curr_meas_Lo", forceunit:=unitVolt, ForceResults:=tlForceFlow
                         
 
    
    DebugPrintFunc patset.Value


Exit Function
errHandler:
    TheExec.AddOutput "Error in the Seq Leakage Test"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Meas_VIR_IO_Universal_func_GPIO(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
Optional Measure_Pin_PPMU As String, Optional ForceV As String, Optional ForceI As String, Optional MeasureI_Range As String = "0.05", _
Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
Optional InstSpecialSetting As InstrumentSpecialSetup = 0, Optional SpecialCalcValSetting As CalculateMethodSetup = 0, Optional RAK_Flag As Enum_RAK = 0, _
Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String, _
Optional Flag_SingleLimit As Boolean = False, Optional TestLimitPerPin_VIR As String = "FFF", _
Optional InterFunc_PrePat As InterposeName, Optional InterFuncArgs_PrePat As String, _
Optional InterFunc_PostPat As InterposeName, Optional InterFuncArgs_PostPat As String, _
Optional CharInputString As String, Optional ForceFunctional_Flag As Boolean = False, Optional MeasIGrpPinCnt As Integer = 0, Optional KeepEmptyLimit As Boolean = False, _
Optional CFG_GPIO_Pins As String) As Long

''20150924 - Remove argument as below
''Remove ,Optional MeasOnHalt As Boolean
''Remove ,Optional InterFunc_SequenceN As InterposeName
''Remove ,Optional InterFuncArgs_SequenceN As String

''==================================================================================
'' 20150621 - Check with CCWu: FRCPortName As String, Optional DisableFRC As Boolean = False not use in this function
'' 20150717 - Impedance measurement by using 2 point measure method, Define "Z" for TestSequence - On going
''                - EX: Pin1, Pin2 + Pin3, Pin4     V1, V2 + V3, V4
''                - V1 and V2 use for Pin1 of impedence measurement
''                - V1 and V2 use for Pin2 of impedence measurement
'' 20150717 - Get I from previous item and apply the current value to next item, use enum for the feature
''                - EX: TestSequence: "V,V,V"
''                  If second V want to apply calcuated I value that Force I value argument should be "0,keyword,0"
'' 20150727 - MeasureI_Range is use for test sequence "I", "R" and "Z"
''==================================================================================
    
    Dim PatCount As Long
    Dim k As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, ForceISequenceArray() As String, ForceVSequenceArray() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim PatMeas As String
    Dim TestPinArrayIV() As String
    Dim TestIrange() As String
    Dim TestSeqNumIdx As Long
    Dim InDSPwave As New DSPWave
    Dim OutDspWave() As New DSPWave
    Dim ShowDec As String
    Dim ShowOut As String
    
    ''20141223
    Dim site As Variant
    Dim PattArray() As String
    Dim Pat As String
    Dim patt As Variant
    
    Dim HighLimitVal() As Double
    Dim LowLimitVal() As Double
    
    ''20150728
    Dim ReturnMeasVolt As New PinListData
    Dim FlowForLoopName() As String   ' Sequences : Code , Voltage , Loop Index
    
    Dim i, j As Integer

    ''20160821-Add judgement by 7.75mA for Fuse
    Dim CFGTestPins(0) As String
    Dim CFGTest_FirstSequence As Boolean
    CFGTest_FirstSequence = True
    
    On Error GoTo errHandler
    
    Shmoo_Pattern = patset.Value

    Call tl_PinListDataSort(True)
    ''========================================================================================
    '' 20150121 - Range Check
    If Range_Check_Enable_Word = True Then         'Change to "Range_Check_Enable_Word" by Martin 20151225  for TTR
        If TheExec.DataManager.MemberIndex = 0 Then
            gl_UseLimitCheck_Counter = 0
        End If
    End If
    ''========================================================================================
   
    '' 20160111 - Check input condition for measure I, R and Z.
''    Call CheckCondition_Measure_I_R_Z(TestSequence, Measure_Pin_PPMU, ForceV, MeasureI_Range)
   
    TestSequenceArray = Split(TestSequence, ",")
    
    If ForceI = "" Then ForceI = 0
    If ForceV = "" Then ForceV = 0
    
    
    
    
    ForceISequenceArray = Split(ForceI, "|")
    ForceVSequenceArray = Split(ForceV, "|")
        '''''''''''Apply DC spec'''''''''''
    For i = 0 To UBound(ForceVSequenceArray)
        ForceVSequenceArray(i) = EvaluateEachBlock(ForceVSequenceArray(i)) ''zhhuangf
    Next i
    '''''''''''Apply DC spec'''''''''''
    '' 20150812-Decompose DigCap_Pin by ","
    Dim DigCap_Pin_Ary() As String
    DigCap_Pin_Ary = Split(DigCap_Pin, ",")
    FlowForLoopName = Split(DigSrc_FlowForLoopIntegerName, ",")
    
    Char_Test_Name_Curr_Loc = 0 'init char datalog test name index
    
    TestPinArrayIV = Split((Measure_Pin_PPMU), "+")
    
    If MeasureI_Range = "" Then MeasureI_Range = "50e-3"
    
    TestIrange = Split(MeasureI_Range, "+")
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
  
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    
    If InterFunc_PrePat <> "" Then Call Interpose(InterFunc_PrePat, InterFuncArgs_PrePat)
    
    '20151028  CUS_MeasV_And_CalR -- TYCHENGG
    ''ex.  CUS_Str_MainProgram = "CalR;1.88;RVOH,RVOL,RVOH"
    ''========================================================================================
    Dim CUS_CalR_VDD As Double
    Dim CUS_CalR_Seq_Ary() As String
    Dim CUS_CalR_Arg_Ary() As String
    If (CUS_Str_MainProgram <> "") Then
        If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
            CUS_CalR_Arg_Ary = Split(CUS_Str_MainProgram, ";")
            CUS_CalR_VDD = CDbl(CUS_CalR_Arg_Ary(1))
            CUS_CalR_Seq_Ary = Split(CUS_CalR_Arg_Ary(2), ",")
        End If
    End If
    ''========================================================================================
    
    '' 20150625 - Apply Char setup
    If CharInputString <> "" Then
'        Call SetForceCondition(CharInputString)
    End If
    
    TheHdw.Patterns(patset).Load
    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    If KeepEmptyLimit = True Then
        Call GetFlowSingleUseLimit_KeepEmpty(HighLimitVal, LowLimitVal)  ''20141223
    Else
        Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)  ''20141223
    End If
    For Each patt In PattArray
        Pat = CStr(patt)
        PatMeas = Pat
         
        If DigSrc_Sample_Size <> 0 Then
             
             '' 20150810 - Source dssc index by For opcode from flow table
            If DigSrc_FlowForLoopIntegerName <> "" Then        '20151201
                If FlowForLoopName(0) <> "" Then        '20151201
                    Call DSSCSrcBitFromFlowForLoop(FlowForLoopName(0), DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment)
                End If
            End If
            
            For Each site In TheExec.sites.Active
                ''20150708- Add customize string for digsrc data process
                Call Create_DigSrc_Data(DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPwave, site, CUS_Str_DigSrcData)
            Next site
            
            Call SetupDigSrcDspWave(PatMeas, DigSrc_pin, "Meas_src", DigSrc_Sample_Size, InDSPwave)
        End If

        ''20150812- Add program to setup multiple DigCap_Pin.
        If DigCap_Sample_Size <> 0 Then
            Dim DigCap_Pin_Num As Integer
            DigCap_Pin_Num = UBound(DigCap_Pin_Ary)
            ReDim OutDspWave(DigCap_Pin_Num) As New DSPWave
            For i = 0 To DigCap_Pin_Num
                TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test for " & DigCap_Pin_Ary(i) & " ========")
                OutDspWave(i).CreateConstant 0, DigCap_Sample_Size
                DigCap_Pin.Value = DigCap_Pin_Ary(i)
                Call DigCapSetup(Pat, DigCap_Pin, "Meas_cap", DigCap_Sample_Size, OutDspWave(i))
           Next i
        End If
          
        Call TheHdw.Patterns(Pat).start
        TestSeqNum = 0
        
        Dim FlowTestName() As String
        
        For Each Ts In TestSequenceArray
            ''20150907 - Only need CPUA_Flag_In_Pat to do control
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            TestOptLen = Len(Ts)
            
            TestSeqNumIdx = TestSeqNum
            For k = 1 To TestOptLen
                TestOption = Mid(Ts, k, 1)
                
                '' 20160106 - If "ForceFunctional_Flag" = True to let TestOption = "N" to make the test instance only run functional test
                If ForceFunctional_Flag = True Then
                    TestOption = "N"
                End If
                
                If (Measure_Pin_PPMU <> "") Then
                    Call Meas_VIR_IO_PreSetupBeforeMeasurement(TestPinArrayIV, TestSeqNumIdx)
                    
                    Select Case UCase(TestOption)
                    
                        Case "V"
                        
                            Call IO_HardIP_PPMU_Measure_V(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceISequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(TestSeqNum), LowLimitVal(TestSeqNum), TestLimitPerPin_VIR, ReturnMeasVolt, _
                                    FlowTestName, SpecialCalcValSetting, InstSpecialSetting, RAK_Flag, CUS_Str_MainProgram)
 
                             ''20151028  CUS_MeasV_And_CalR -- TYCHENGG
                            ''========================================================================================
                            If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
                                Call CUS_VIR_MainProgram_MeasV_CalR(TestPinArrayIV, TestSeqNum, CUS_CalR_Seq_Ary, ForceISequenceArray, ReturnMeasVolt, CUS_CalR_VDD)
                            End If
                            ''========================================================================================
                            
                        Case "I"
                            
                            If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
                            
'                            Call IO_HardIP_PPMU_Measure_I_TTR(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
'                                    k, Pat, Flag_SingleLimit, HighLimitVal(TestSeqNum), LowLimitVal(TestSeqNum), TestLimitPerPin_VIR, TestIrange, CUS_Str_MainProgram, SpecialCalcValSetting, MeasIGrpPinCnt, CFG_GPIO_Pins, CFGTest_FirstSequence)
                            If CFGTest_FirstSequence = True Then CFGTest_FirstSequence = False
                            
                        Case "R"
                            
                            Call IO_HardIP_PPMU_Measure_R(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, TestIrange, FlowTestName, RAK_Flag)
                        
                        Case "Z"
                            
                            Call IO_HardIP_PPMU_Measure_Z(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, TestIrange, FlowTestName, RAK_Flag)
                                    
                        Case "E"
                            Dim ForceI1 As String, ForceI2 As String, ForceI3 As String
                            Dim ForceISequenceArray1() As String, ForceISequenceArray2() As String, ForceISequenceArray3() As String
                            'ForceI = "0,-0.01155+0,-0.01155_0,-0.00655+0,-0.00655_0,-0.01655+0,-0.01655_0,-0.01155+0,-0.01155"
                            Dim ForceI_split() As String
                            ForceI_split = Split(ForceI, "_")
                            If ForceI = "" Then ForceI = 0
                            If ForceI1 = "" Then ForceI1 = 0
                            If ForceI2 = "" Then ForceI2 = 0
                            If ForceI3 = "" Then ForceI3 = 0
                            If ForceV = "" Then ForceV = 0
                    
                            ForceISequenceArray = Split(ForceI_split(0), "+") 'Split(ForceI, "+")
                            ForceISequenceArray1 = Split(ForceI_split(1), "+") 'Split(ForceI1, "+")
                            ForceISequenceArray2 = Split(ForceI_split(2), "+") 'Split(ForceI2, "+")
                            ForceISequenceArray3 = Split(ForceI_split(3), "+") 'Split(ForceI3, "+")
                            'ForceVSequenceArray = Split(ForceV, "+")
                    
                            Dim ForceByPin() As String
                            Dim result_v1(4) As New PinListData
                            Dim delta_V As New SiteDouble
                            Dim delta_I As New SiteDouble
                            Dim delta_R As New SiteDouble
                            Dim delta_R0 As New SiteDouble
                            Dim delta_R1 As New SiteDouble
                            Dim cal_V As New SiteDouble
                            Dim cal_V0 As New SiteDouble
                            Dim cal_V1 As New SiteDouble
                            Dim R_trace As New SiteDouble
                            Dim delta_Rtx As New SiteDouble
                            Dim ForceISequenceArrayI0() As String, ForceISequenceArrayI1() As String, ForceISequenceArrayI2() As String
                            Dim X As Integer
                            Dim Cal_method As Integer
                            Cal_method = 2
                            Dim R_trace1 As Double, R_trace2 As Double
                       
                            R_trace1 = 4.5
                            R_trace2 = 4.5
                            ' ForceISequenceArray
                            'Call IO_HardIP_PPMU_Measure_VIR_USB_HSTX(TestOption, PpmuMeasure_Pin, TestPinArrayIV, _
                                    TestSeqNum, TestSeqNumIdx, ForceISequenceArray, _
                                    k, Pat, CurrentJobName, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, _
                                    TestIrange)
                            'result_v1(0) = result_v
                            
                            'Call IO_HardIP_PPMU_Measure_VIR_USB_HSTX(TestOption, PpmuMeasure_Pin, TestPinArrayIV, _
                                    TestSeqNum, TestSeqNumIdx, ForceISequenceArray1, _
                                    k, Pat, CurrentJobName, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, _
                                    TestIrange)
                        
                            Call IO_HardIP_PPMU_Measure_V(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceISequenceArray1, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, ReturnMeasVolt, _
                                    FlowTestName, SpecialCalcValSetting)
                               
                            result_v1(1) = ReturnMeasVolt
                            
                            'Call IO_HardIP_PPMU_Measure_VIR_USB_HSTX(TestOption, PpmuMeasure_Pin, TestPinArrayIV, _
                                    TestSeqNum, TestSeqNumIdx, ForceISequenceArray2, _
                                    k, Pat, CurrentJobName, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, _
                                    TestIrange)
                               
                            Call IO_HardIP_PPMU_Measure_V(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceISequenceArray2, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, ReturnMeasVolt, _
                                    FlowTestName, SpecialCalcValSetting)
       
                            result_v1(2) = ReturnMeasVolt
                        
                            'Call IO_HardIP_PPMU_Measure_VIR_USB_HSTX(TestOption, PpmuMeasure_Pin, TestPinArrayIV, _
                                    TestSeqNum, TestSeqNumIdx, ForceISequenceArray3, _
                                    k, Pat, CurrentJobName, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, _
                                    TestIrange)
                            'result_v1(3) = result_v

                            For X = 0 To ReturnMeasVolt.Pins.Count - 1
                    
                                ForceISequenceArrayI0 = Split(ForceISequenceArray(1), ",")
                                ForceISequenceArrayI1 = Split(ForceISequenceArray1(1), ",")
                                ForceISequenceArrayI2 = Split(ForceISequenceArray2(1), ",")
                    
                                Select Case Cal_method
                                     
                                    Case "1"
                                        
                                        For Each site In TheExec.sites.Active
                                            delta_I = Abs(ForceISequenceArrayI2(1) - ForceISequenceArrayI1(1))
                                            delta_V = Abs(result_v1(1).Pins(X).Value(site) - result_v1(2).Pins(X).Value(site))
                                            delta_R = delta_V / delta_I
                                            cal_V = Abs(result_v1(0).Pins(X).Value(site)) * (45 / (delta_R + 45))
                                            TheExec.Datalog.WriteComment "          " & site & "   Cal_Reslut " & result_v1(1).Pins(X) & " => Delta_V =" & delta_V & " => Delta_R =" & delta_R & " => CAL_Volt =" & cal_V & " mV"
                                        Next site

                                        TheExec.Flow.TestLimit cal_V, , , Unit:=unitCustom, customUnit:="V", formatStr:="%.3f", Tname:="Volt_meas" + "_" + CStr(TestSeqNum), PinName:=UCase(result_v1(1).Pins(X).Name), ForceResults:=tlForceFlow
                        
                                    Case "2"
                                        
                                        For Each site In TheExec.sites.Active
                                            delta_R1 = Abs(result_v1(2).Pins(X).Value(site) - result_v1(1).Pins(X).Value(site)) / Abs(ForceISequenceArrayI2(1) - ForceISequenceArrayI1(1))
                                  
                                            If UCase(result_v1(1).Pins(X)) = "USB_DM" Then
                                                R_trace = R_trace1
                                            Else
                                                R_trace = R_trace2
                                            End If
                                  
                                            delta_Rtx = Abs(delta_R1 - R_trace)

                                            cal_V1 = 50 * (Abs(result_v1(1).Pins(X).Value(site)) + Abs(ForceISequenceArrayI1(1)) * delta_R1) / (delta_R1 - R_trace + 50)
                                            TheExec.Datalog.WriteComment "          " & site & "   Cal_Reslut " & result_v1(1).Pins(X) & " => R0 =" & delta_R0 & " => R1 =" & delta_R1 & " => Trace_R =" & R_trace & " => CAL_Volt1 =" & cal_V1 * 1000 & " mV"
                                 
                                        Next site

                                        If LCase(currentJobName) Like "*char*" Then
                                            TheExec.Flow.TestLimit cal_V1, , , Unit:=unitCustom, customUnit:="V", formatStr:="%.3f", PinName:=UCase(result_v1(1).Pins(X).Name), ForceResults:=tlForceFlow
                                        Else
                                            TheExec.Flow.TestLimit cal_V1, , , Unit:=unitCustom, customUnit:="V", formatStr:="%.3f", Tname:="Volt_meas" + "_" + CStr(TestSeqNum), ForceResults:=tlForceFlow
                                        End If
                        
                                    Case Else
                                        TheExec.Datalog.WriteComment "Error Cal_methodology, please select correct one"
                                End Select
                            Next X
                       
                        Case "N"
                        
                        Case Else
                            TheExec.Datalog.WriteComment "Error Test Option, please select V, I or R"
                    
                    End Select
                    
                    Call Meas_VIR_IO_PostSetupAfterMeasurement(TestPinArrayIV, TestSeqNumIdx)
                End If
            Next k
            
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
            End If
        Next Ts
        
        If DebugPrintEnable = True Then
            TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & ""
        End If
                
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        PatCount = PatCount + 1
    
        If DigCap_Sample_Size <> 0 Then
'            Call HardIP_Digcap_Print(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, ShowDec, ShowOut, , DigCap_DSPWaveSetting) '''change
        End If
    Next patt
    
    If InterFunc_PostPat <> "" Then Call Interpose(InterFunc_PostPat, InterFuncArgs_PostPat)
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Connect
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    
    If DisableFRC = True Then
        Call FreeRunclk_Enable(FRCPortName)
    End If

    Call HardIP_WriteFuncResult
    
    DebugPrintFunc patset.Value  ' print all debug information
    
    '' 20150728 - Print Char setup for power pins.
    If CurrentJobName_U Like "*CHAR*" Then
        If CharInputString <> "" Then
          '  Call PrintCharPowerSet(CharInputString)
        End If
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Meas_VIR_IO_Universal_func"
    If AbortTest Then Exit Function Else Resume Next
  
End Function

Public Function IO_HardIP_PPMU_Measure_I_TTR(TestPinArrayIV() As String, _
                                             TestSeqNum As Integer, _
                                             TestSeqNumIdx As Long, _
                                             ForceSequenceArray() As String, _
                                             k As Long, Pat As Variant, _
                                             Flag_SingleLimit As Boolean, _
                                             HighLimitVal As Double, _
                                             LowLimitVal As Double, _
                                             TestLimitPerPin_VIR As String, _
                                             TestIrange() As String, _
                                             VDD_IO_1p2 As String, _
                                             VDD_IO_1p8 As String, _
                                    Optional CUS_Str_MainProgram As String, _
                                    Optional PPMU_TestLimit_TTR As Boolean, _
                                    Optional MeasIGrpPinCnt As Integer = 0, _
                                    Optional CFG_GPIO_Pins As String, _
                                    Optional CFGTest_FirstSequence As Boolean) As Long
'VDD_IO_1p2 = "VDDIO12_GRP5"
'VDD_IO_1p8 = "VDDIO18_GRP"
On Error GoTo errHandler
    Dim funcName As String:: funcName = "IO_HardIP_PPMU_Measure_I_TTR"
    
    Dim MeasureValue As New PinListData
    Dim TestNum As Long
    Dim Pin As Variant
    Dim p As Long
    Dim ForceV  As Double
    Dim MeasureValueGrp As New PinListData
    Dim PinName As String
    Dim Vdiff As Double


 ''========================================================================================
    '' 20150202 - Range Check
    Dim RangeCheck_HighLimitVal() As Double
    Dim RangeCheck_LowLimitVal() As Double
    If Range_Check_Enable_Word = True Then
        Call GetFlowSingleUseLimit(RangeCheck_HighLimitVal, RangeCheck_LowLimitVal)
    End If
    ''========================================================================================
    
    ''=========================================================================================================
    '' 20160108 - Add rule to cover force value is different
    Dim ForceByPin() As String
    Dim ForceValByPin() As String
''    Dim ForceValIdx As Integer
    Dim Measure_I_Range() As String
    Dim MeasurePin As String
    Dim MI_Range_Index As Long
    Dim i, j As Long
    Dim OutputTname As String
    
    
'    If LCase(TheExec.DataManager.InstanceName) Like "*lv*" Then
'        Vratio = 0.9
'    ElseIf LCase(TheExec.DataManager.InstanceName) Like "*hv*" Then
'        Vratio = 1.1
'    Else
'        Vratio = 1
'    End If
    
    MI_Range_Index = 0
    
    '' Force Pin
    If UBound(TestPinArrayIV) = 0 Then
        ForceByPin = Split(TestPinArrayIV(0), ",")
        MeasurePin = TestPinArrayIV(0)
        TheHdw.Digital.Pins(TestPinArrayIV(0)).Disconnect
    Else
        ForceByPin = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        MeasurePin = TestPinArrayIV(TestSeqNumIdx)
        TheHdw.Digital.Pins(TestPinArrayIV(TestSeqNumIdx)).Disconnect
    End If
    
    '' Force Volt value
    If UBound(ForceSequenceArray) = 0 Then
        ForceValByPin = Split(ForceSequenceArray(0), ",")
    Else
        ForceValByPin = Split(ForceSequenceArray(TestSeqNumIdx), ",")
    End If
    
    '' Measure Current range
    If UBound(TestIrange) = 0 Then
        Measure_I_Range = Split(TestIrange(0), ",")
        Dim MeasurePinArry() As String
        MeasurePinArry = Split(MeasurePin, ",")    'expend I range for all point
        If UBound(Measure_I_Range) = 0 Then
            ReDim Preserve Measure_I_Range(UBound(MeasurePinArry))
            'ReDim Measure_I_Range(UBound(MeasurePinArry))
            For i = 0 To UBound(MeasurePinArry)
                Measure_I_Range(i) = Measure_I_Range(0)
            Next i
        End If
    Else
        Measure_I_Range = Split(TestIrange(TestSeqNumIdx), ",")
    End If
    
    ''=========================================================================================================
    '' 20150108 - Check number whether differrent between measure current range and force pin, add defalut value to let input number are the same.
    Call VIR_CheckTestCondition_Measure_I_R_Z("I", ForceByPin, Measure_I_Range)
     '' 20150111 - Check force value is the same or different
   ' Dim i As Long
    Dim b_ForceDiffVolt As Boolean
    Dim PastVal As Double
    b_ForceDiffVolt = False
    For i = 0 To UBound(ForceValByPin)
        If i <> 0 Then
            If ForceValByPin(i) <> PastVal Then
                                b_ForceDiffVolt = True
                                Exit For
            End If
        End If
        
        PastVal = (ForceValByPin(i)) '' use global para need to use the Evalute
    Next i
    ''=========================================================================================================
        If MeasIGrpPinCnt = 0 Then
                MeasIGrpPinCnt = 1
        End If
        
        Dim MeasIGrpCnt As Integer
    Dim TestPinArray() As String
    Dim PinCount As Long
    Dim TestPinStr As String
    
    For Each Pin In ForceByPin
        TheExec.DataManager.DecomposePinList Pin, TestPinArray, PinCount
    
        If PinCount Mod MeasIGrpPinCnt <> 0 Then
            MeasIGrpCnt = PinCount \ MeasIGrpPinCnt + 1
        Else
            MeasIGrpCnt = PinCount / MeasIGrpPinCnt
        End If
        
        If b_ForceDiffVolt = False Then
            If ForceValByPin(0) > 0.8 Then
                If LCase(Pin) Like "*1p2*" Then
                    Vdiff = TheHdw.DCVS.Pins(VDD_IO_1p2).Voltage.Value - ForceValByPin(0) ''vddio12_grp!!!
                Else
                    Vdiff = TheHdw.DCVS.Pins(VDD_IO_1p8).Voltage.Value - ForceValByPin(0) ''vddio18_grp!!!
                End If
            Else
                Vdiff = ForceValByPin(0)
            End If
        Else
            If ForceValByPin(MI_Range_Index) > 0.8 Then
                If LCase(Pin) Like "*1p2*" Then
                    Vdiff = TheHdw.DCVS.Pins(VDD_IO_1p2).Voltage.Value - ForceValByPin(MI_Range_Index)
                Else
                    Vdiff = TheHdw.DCVS.Pins(VDD_IO_1p8).Voltage.Value - ForceValByPin(MI_Range_Index)
                End If
            Else
                Vdiff = ForceValByPin(MI_Range_Index)
            End If
        End If

        For i = 0 To MeasIGrpCnt - 1
            TestPinStr = ""
            
            For j = 0 To MeasIGrpPinCnt - 1
                If (i * MeasIGrpPinCnt + j) < PinCount Then
                    TestPinStr = TestPinStr & "," & TestPinArray(i * MeasIGrpPinCnt + j)
                    MeasureValue.AddPin (TestPinArray(i * MeasIGrpPinCnt + j))
                End If
            Next j
            TestPinStr = Right(TestPinStr, Len(TestPinStr) - 1)
    
            With TheHdw.PPMU.Pins(TestPinStr)
            '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
                .ForceI 0, 0.05
                .Connect
                .Gate = tlOn
                If b_ForceDiffVolt = False Then
                    .ForceV ForceValByPin(0), Measure_I_Range(MI_Range_Index)
'                    If ForceValByPin(0) > 0.7 Then ''161219
'                        Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP").Voltage.Value - ForceValByPin(0) ''vddio18_grp!!!
'                    Else
'                        Vdiff = ForceValByPin(0)
'                    End If
                Else
                    .ForceV ForceValByPin(MI_Range_Index), Measure_I_Range(MI_Range_Index)
'                    If ForceValByPin(0) > 0.7 Then
'                        Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP").Voltage.Value - ForceValByPin(MI_Range_Index)
'                    Else
'                        Vdiff = ForceValByPin(MI_Range_Index)
'                    End If
                End If
                '' 20160108 - Only keep 1 force value but current range can be different for force pin
            End With
            
            If PPMU_TestLimit_TTR = False Then TheExec.Datalog.WriteComment "Pin = " & (TestPinStr & " Measure Current Range = " & TheHdw.PPMU.Pins(Pin).MeasureCurrentRange)
            
            TheHdw.Wait (100 * us)
            
            MeasureValueGrp = TheHdw.PPMU.Pins(TestPinStr).Read(tlPPMUReadMeasurements, 10)
            
            For j = 0 To MeasIGrpPinCnt - 1
                If (i * MeasIGrpPinCnt + j) < PinCount Then
                    PinName = TestPinArray(i * MeasIGrpPinCnt + j)
                    
                    ' 20161207 Add RAK
                                        MeasureValue.Pins(PinName) = MeasureValueGrp.Pins(PinName).Multiply(CurrentJob_Card_RAK.Pins(PinName)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(MeasureValueGrp.Pins(PinName))
                    If TheExec.TesterMode = testModeOffline Then MeasureValue.Pins(PinName) = 0.0005 + Rnd() * 0.0001
                End If
            Next j
            
            TheHdw.PPMU.Pins(TestPinStr).Disconnect
        Next i
        MI_Range_Index = MI_Range_Index + 1
    Next Pin
    
    
    If UBound(ForceSequenceArray) <> 0 Then
        If ForceSequenceArray(TestSeqNum) = "" Then
            ForceSequenceArray(TestSeqNum) = 0
        End If
    End If
    
    
'''   TheHdw.Wait (100 * us)
    
    '' 20160112 - Comment this
''    MeasureValue = TheHdw.PPMU.Pins(TestPinArrayIV(TestSeqNumIdx)).Read(tlPPMUReadMeasurements, 10)
    DebugPrintFunc_PPMU CStr(MeasurePin)
    
    ''20150904 - Move to CUS_VIR_MainProgram_MeasureI
''    If CUS_Str_MainProgram <> "" Then
''        Call CUS_VIR_MainProgram_MeasureI(CUS_Str_MainProgram, VIR_MI_AFTER_MEASUREMENT, MeasureValue)
''    End If

    
    Dim TestNameInput As String
    TestNameInput = "Curr_meas_" + CStr(TestSeqNum)
        
    ''20151103 print force condition
    Call Print_Force_Condition("i", MeasureValue)
    
''    ''20160112 - Force value index for test limit if force voltage value is different
''    Dim ForceVal_Index As Long
''    ForceVal_Index = 0
                
                Flag_SingleLimit = True '''Driver strength test using single limit.

    If PPMU_TestLimit_TTR = True Then
        Dim Lowlimitval_temp As Double
        Dim Hilimitval_temp As Double
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        
        Lowlimitval_temp = GetLowLimitFromFlow
        Hilimitval_temp = GetHiLimitFromFlow
        For Each site In TheExec.sites.Active
            For p = 0 To MeasureValue.Pins.Count - 1
                If MeasureValue.Pins(p).Value > Hilimitval_temp Or MeasureValue.Pins(p).Value < Lowlimitval_temp Then
                    If gl_UseStandardTestName_Flag = True Then                       'Roger add
                                                TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(p), OutputTname, TestSeqNum, p)
                    End If
                    TheExec.Datalog.WriteComment "Pin = " & (MeasureValue.Pins(p) & " Measure Current Range = " & TheHdw.PPMU.Pins(MeasureValue.Pins(p)).MeasureCurrentRange)
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitAmp, , Tname:=TestNameInput, ForceResults:=tlForceNone
                End If
            Next p
        Next site

        'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
    ElseIf Flag_SingleLimit = True Then
        If b_ForceDiffVolt = False Then

                        For p = 0 To MeasureValue.Pins.Count - 1
                If LowLimitVal = -123456.123456 Then
                  If gl_UseStandardTestName_Flag = True Then                       'Roger add
                    TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(p), OutputTname, TestSeqNum, p)
                    End If
                    
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), , HighLimitVal, scaletype:=scaleNone, Unit:=unitAmp, Tname:=TestNameInput, ForceVal:=ForceValByPin(0), ForceUnit:=unitVolt, ForceResults:=tlForceNone
                ElseIf HighLimitVal = -123456.123456 Then
                    If gl_UseStandardTestName_Flag = True Then                       'Roger add
                    TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(p), OutputTname, TestSeqNum, p)
                    End If
                    
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), LowLimitVal, , scaletype:=scaleNone, Unit:=unitAmp, Tname:=TestNameInput, ForceVal:=ForceValByPin(0), ForceUnit:=unitVolt, ForceResults:=tlForceNone
                Else
                    If gl_UseStandardTestName_Flag = True Then                       'Roger add
                    TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(p), OutputTname, TestSeqNum, p)
                    End If
                    If p = 0 Then
                        TheExec.Flow.TestLimit MeasureValue.Pins(p), LowLimitVal, HighLimitVal, scaletype:=scaleNone, Unit:=unitAmp, Tname:=TestNameInput, ForceVal:=ForceValByPin(0), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                    Else
                        TheExec.Flow.TestLimit MeasureValue.Pins(p), GetLowLimitFromFlow, GetHiLimitFromFlow, scaletype:=scaleNone, Unit:=unitAmp, Tname:=TestNameInput, ForceVal:=ForceValByPin(0), ForceUnit:=unitVolt, ForceResults:=tlForceNone
                        
                    End If
                
                End If
                        Next p

        Else
            For p = 0 To MeasureValue.Pins.Count - 1
                If gl_UseStandardTestName_Flag = True Then                       'Roger add
                    TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(p), OutputTname, TestSeqNum, p)
                End If
                If p = 0 Then
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), LowLimitVal, HighLimitVal, scaletype:=scaleNone, Unit:=unitAmp, Tname:=TestNameInput, ForceVal:=FormatNumber(TheHdw.PPMU(MeasureValue.Pins(p)).Voltage.Value, 3), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                Else
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), GetLowLimitFromFlow, GetHiLimitFromFlow, scaletype:=scaleNone, Unit:=unitAmp, Tname:=TestNameInput, ForceVal:=FormatNumber(TheHdw.PPMU(MeasureValue.Pins(p)).Voltage.Value, 3), ForceUnit:=unitVolt, ForceResults:=tlForceNone
                End If
            
            Next p
        End If

        
    Else

    End If


Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function IO_HardIP_PPMU_Measure_V_TTR(TestPinArrayIV() As String, _
                                             TestSeqNum As Integer, _
                                             TestSeqNumIdx As Long, _
                                             ForceSequenceArray() As String, _
                                             k As Long, Pat As Variant, _
                                             Flag_SingleLimit As Boolean, _
                                             HighLimitVal As Double, _
                                             LowLimitVal As Double, _
                                             TestLimitPerPin_VIR As String, _
                                       ByRef ReturnMeasVolt As PinListData, _
                                    Optional InstSpecialSetting As InstrumentSpecialSetup = 0, _
                                    Optional RAK_Flag As Enum_RAK = 0, _
                                    Optional CUS_Str_MainProgram As String = "", _
                                    Optional MeasVGrpPinCnt As Integer) As Long

    Dim MeasureValue As New PinListData
    Dim Force_idx As Integer
    Dim site As Variant
    Dim TestNum As Long
    Dim Pin  As Variant
    
    Dim p As Long
    Dim ForceV  As Double
    Dim ForceByPin() As String
    Dim ForceValByPin() As String
    Dim ForceValIdx As Integer
    Dim IdxV As Integer
    Dim MeasurePinStr As String
    Dim i, j As Long
    Dim MeasureValueGrp As New PinListData
    Dim PinName As String
    Dim PinArr() As String
    Dim PinCount As Long
    Dim MeasVGrpCnt As Integer
    Dim TestPinStr As String
    Dim TestPinArrayIVsize As Long

    ''========================================================================================
    '' 20150202 - Range Check
    Dim RangeCheck_HighLimitVal() As Double
    Dim RangeCheck_LowLimitVal() As Double
    Dim TempMeasVal_PerPin(100) As New PinListData
    If Range_Check_Enable_Word = True Then
        Call GetFlowSingleUseLimit(RangeCheck_HighLimitVal, RangeCheck_LowLimitVal)
    End If
    
    If UBound(ForceSequenceArray) = 0 Then
        ForceValByPin = Split(ForceSequenceArray(0), ",")
    Else
        ForceValByPin = Split(ForceSequenceArray(TestSeqNumIdx), ",")
    End If
    
    ForceValIdx = 0
    
    If UBound(TestPinArrayIV) = 0 Then
        ForceByPin = Split(TestPinArrayIV(0), ",")
        MeasurePinStr = TestPinArrayIV(0)       '20160224 add to allow every seq with the same pins
        'TheHdw.Digital.Pins(TestPinArrayIV(0)).Disconnect
    Else
        ForceByPin = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        MeasurePinStr = TestPinArrayIV(TestSeqNumIdx)       '20160224 add to allow every seq with the same pins
        'TheHdw.Digital.Pins(TestPinArrayIV(TestSeqNumIdx)).Disconnect
    End If
    '' 20150721 - Apply force I value from Stored_MeasI_PPMU,
    ''                - Can not coexist between stored value and force value at the same sequence
    ''
    Dim b_IsNumeral As Boolean
    Dim b_UseStoredForceVal As Boolean
    
    b_IsNumeral = ContentIsNumeral(ForceValByPin(0))
    If b_IsNumeral Then
        b_UseStoredForceVal = False
    Else
        b_UseStoredForceVal = True
    End If
    Dim ForceValI As Double
    If b_UseStoredForceVal = False Then '' 20150721 - Normal usage
        If MeasVGrpPinCnt > 0 And InstSpecialSetting <> InstrumentSpecialSetup.PPMU_SerialMeasurement Then '' Split pin group
            For Each Pin In ForceByPin
                TheExec.DataManager.DecomposePinList Pin, PinArr, PinCount
        
                If PinCount Mod MeasVGrpPinCnt <> 0 Then
                    MeasVGrpCnt = PinCount \ MeasVGrpPinCnt + 1
                Else
                    MeasVGrpCnt = PinCount / MeasVGrpPinCnt
                End If
                
                For i = 0 To MeasVGrpCnt - 1
                    TestPinStr = ""
                    
                    For j = 0 To MeasVGrpPinCnt - 1
                        If (i * MeasVGrpPinCnt + j) < PinCount Then
                            TestPinStr = TestPinStr & "," & PinArr(i * MeasVGrpPinCnt + j)
                            MeasureValue.AddPin (PinArr(i * MeasVGrpPinCnt + j))
                        End If
                    Next j
                    TestPinStr = Right(TestPinStr, Len(TestPinStr) - 1)
                
                    With TheHdw.PPMU.Pins(TestPinStr)
                        '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
                        .ForceI 0, 0
                        .Connect
                        .Gate = tlOn
                       
                        If UBound(ForceValByPin) = 0 Then
                            .ForceI ForceValByPin(0), ForceValByPin(0)
                            ForceValI = ForceValByPin(0)
                        ElseIf ForceValByPin(ForceValIdx) <> "" Then
                            .ForceI ForceValByPin(ForceValIdx), ForceValByPin(ForceValIdx)
                            ForceValI = ForceValByPin(ForceValIdx)
                        Else:
                            .ForceI 0
                            ForceValI = 0
                        End If
                    End With
                    
                    TheHdw.Wait (1 * ms)
                    DebugPrintFunc_PPMU CStr(TestPinStr)
                    MeasureValueGrp = TheHdw.PPMU.Pins(TestPinStr).Read(tlPPMUReadMeasurements, 10)
                    
                    For j = 0 To MeasVGrpPinCnt - 1
                        If (i * MeasVGrpPinCnt + j) < PinCount Then
                            PinName = PinArr(i * MeasVGrpPinCnt + j)
                            MeasureValue.Pins(PinName) = MeasureValueGrp.Pins(PinName)
                            If TheExec.TesterMode = testModeOffline Then MeasureValue.Pins(PinName) = 0.005 + Rnd() * 0.001
                        End If
                    Next j
                    
                    TheHdw.PPMU.Pins(TestPinStr).Disconnect
                Next i
                ForceValIdx = ForceValIdx + 1
            Next Pin
        Else '' No split pin group
            '' 20150721 - Normal usage
            If InstSpecialSetting = InstrumentSpecialSetup.PPMU_SerialMeasurement Then
                TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
                TheHdw.PPMU.Pins(MeasurePinStr).ForceI 0, 0
        
                For Each Pin In PinArr
                    MeasureValue.AddPin (Pin)
                    TheHdw.PPMU.Pins(Pin).ForceI ForceValByPin(0), Abs(ForceValByPin(0))
                    ForceValI = ForceValByPin(0)
                    TheHdw.Wait 0.001
                    DebugPrintFunc_PPMU CStr(Pin)
                    MeasureValue.Pins(Pin) = TheHdw.PPMU.Pins(Pin).Read(tlPPMUReadMeasurements, 10)
                    TheHdw.PPMU.Pins(Pin).ForceI 0, 0
                    TheHdw.PPMU.Pins(Pin).Disconnect
                Next Pin
            Else
                For Each Pin In ForceByPin
                    With TheHdw.PPMU.Pins(Pin)
                        '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
                        .ForceI 0, 0
                        .Connect
                        .Gate = tlOn
                       
                        If UBound(ForceValByPin) = 0 Then
                            .ForceI ForceValByPin(0), ForceValByPin(0)
                            ForceValI = ForceValByPin(0)
                        ElseIf ForceValByPin(ForceValIdx) <> "" Then
                            .ForceI ForceValByPin(ForceValIdx), ForceValByPin(ForceValIdx)
                             ForceValI = ForceValByPin(ForceValIdx)
                        Else:
                            .ForceI 0
                            ForceValI = 0
                        End If
                    End With
                    ForceValIdx = ForceValIdx + 1
                Next Pin
            
                TheHdw.Wait (1 * ms)
                DebugPrintFunc_PPMU CStr(MeasurePinStr)
                MeasureValue = TheHdw.PPMU.Pins(MeasurePinStr).Read(tlPPMUReadMeasurements, 10)
                If TheExec.TesterMode = testModeOffline Then MeasureValue = 0.005 + Rnd() * 0.001
            End If
        End If
    
    Else
        '' 20150721 - Apply stored value
        Dim AfterformulaVal_PPMU As New PinListData
''        Call CUS_FormulaCalc(Stored_MeasI_PPMU, AfterformulaVal_PPMU)
        
        '' 20150721 - Store ForceValue for each site for test limit use.
        Dim TestPinMaxNum As Integer
        TestPinMaxNum = UBound(ForceByPin)
        ReDim StoreForceI(TestPinMaxNum) As New SiteDouble
        
        If MeasVGrpPinCnt > 0 And InstSpecialSetting <> InstrumentSpecialSetup.PPMU_SerialMeasurement Then '' Split pin group
            For Each Pin In ForceByPin
                TheExec.DataManager.DecomposePinList Pin, PinArr, PinCount
        
                If PinCount Mod MeasVGrpPinCnt <> 0 Then
                    MeasVGrpCnt = PinCount \ MeasVGrpPinCnt + 1
                Else
                    MeasVGrpCnt = PinCount / MeasVGrpPinCnt
                End If
                
                For i = 0 To MeasVGrpCnt - 1
                    TestPinStr = ""
                    
                    For j = 0 To MeasVGrpPinCnt - 1
                        If (i * MeasVGrpPinCnt + j) < PinCount Then
                            TestPinStr = TestPinStr & "," & PinArr(i * MeasVGrpPinCnt + j)
                            MeasureValue.AddPin (PinArr(i * MeasVGrpPinCnt + j))
                        End If
                    Next j
                    TestPinStr = Right(TestPinStr, Len(TestPinStr) - 1)
                    
                    For Each site In TheExec.sites.Active
                        With TheHdw.PPMU.Pins(TestPinStr)
                            If UBound(ForceValByPin) = 0 Then
                                .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(0)).Value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(0)).Value(site)
                                StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site)
                            ElseIf ForceValByPin(ForceValIdx) <> "" Then
                                .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site)
                                StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site)
                            Else
                                .ForceI 0
                            End If
                        End With
                    Next site
                    
                    TheHdw.Wait (1 * ms)
                    DebugPrintFunc_PPMU CStr(TestPinStr)
                    MeasureValueGrp = TheHdw.PPMU.Pins(TestPinStr).Read(tlPPMUReadMeasurements, 10)
                    
                    For j = 0 To MeasVGrpPinCnt - 1
                        If (i * MeasVGrpPinCnt + j) < PinCount Then
                            PinName = PinArr(i * MeasVGrpPinCnt + j)
                            MeasureValue.Pins(PinName) = MeasureValueGrp.Pins(PinName)
                            If TheExec.TesterMode = testModeOffline Then MeasureValue.Pins(PinName) = 0.005 + Rnd() * 0.001
                        End If
                    Next j
                    
                    TheHdw.PPMU.Pins(TestPinStr).Disconnect
                Next i
                ForceValIdx = ForceValIdx + 1
            Next Pin
        Else '' No split pin group
            If InstSpecialSetting = InstrumentSpecialSetup.PPMU_SerialMeasurement Then
                TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
                TheHdw.PPMU.Pins(MeasurePinStr).ForceI 0, 0
        
                For Each Pin In PinArr
                    MeasureValue.AddPin (Pin)
                    TheHdw.PPMU.Pins(Pin).ForceI ForceValByPin(0), Abs(ForceValByPin(0))
                    ForceValI = ForceValByPin(0)
                    TheHdw.Wait 0.001
                    DebugPrintFunc_PPMU CStr(Pin)
                    MeasureValue.Pins(Pin) = TheHdw.PPMU.Pins(Pin).Read(tlPPMUReadMeasurements, 10)
                    TheHdw.PPMU.Pins(Pin).ForceI 0, 0
                    TheHdw.PPMU.Pins(Pin).Disconnect
                Next Pin
            Else
                For Each Pin In ForceByPin
                    For Each site In TheExec.sites.Active
                        With TheHdw.PPMU.Pins(Pin)
                            If UBound(ForceValByPin) = 0 Then
                                .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(0)).Value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(0)).Value(site)
                                
                                StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site)
                                
                            ElseIf ForceValByPin(ForceValIdx) <> "" Then
                                .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site)
                                
                                StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site)
                            Else
                                .ForceI 0
                            End If
                        End With
                    Next site
                    ForceValIdx = ForceValIdx + 1
                Next Pin
            
                TheHdw.Wait (1 * ms)
                DebugPrintFunc_PPMU CStr(MeasurePinStr)
                MeasureValue = TheHdw.PPMU.Pins(MeasurePinStr).Read(tlPPMUReadMeasurements, 10)
                If TheExec.TesterMode = testModeOffline Then MeasureValue = 0.005 + Rnd() * 0.001
            End If
        End If
    End If

    If UBound(ForceSequenceArray) <> 0 Then
        If ForceSequenceArray(TestSeqNum) = "" Then
            ForceSequenceArray(TestSeqNum) = 0
        End If
    End If

    For Each site In TheExec.sites.Active
        TestNum = TheExec.sites.Item(site).TestNumber
    Next site

'    TheHdw.Wait (1 * ms)
'
'    If InstSpecialSetting = DigitalConnectPPMU2 Then
'        TheHdw.PPMU.AllowPPMUFuncRelayConnection (True)
'        TheHdw.PPMU.Pins(MeasurePinStr).ForceI 0, 0.0002 '20160224 add to allow every seq with the same pins
'        TheHdw.Digital.Pins(MeasurePinStr).Connect
'    End If
'
'    If InstSpecialSetting = PPMU_SerialMeasurement Then
'        Dim PinArr() As String
'        Dim PinCount As Long
'
'        TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
'        TheHdw.PPMU.Pins(MeasurePinStr).ForceI 0, 0
'
'        For Each Pin In PinArr
'            MeasureValue.AddPin (Pin)
'            TheHdw.PPMU.Pins(Pin).ForceI ForceValByPin(0), Abs(ForceValByPin(0))
'            TheHdw.Wait 0.001
'            DebugPrintFunc_PPMU CStr(Pin)
'            MeasureValue.Pins(Pin) = TheHdw.PPMU.Pins(Pin).Read(tlPPMUReadMeasurements, 10)
'            TheHdw.PPMU.Pins(Pin).ForceI 0, 0
'            TheHdw.PPMU.Pins(Pin).Disconnect
'        Next Pin
'    Else
'        DebugPrintFunc_PPMU CStr(MeasurePinStr)
'        MeasureValue = TheHdw.PPMU.Pins(MeasurePinStr).Read(tlPPMUReadMeasurements, 10)
'    End If
    
    '' Calculate RAK
    Dim RakV() As Double
    If RAK_Flag = True Then
        For Each site In TheExec.sites
            For Each Pin In MeasureValue.Pins
    
                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(Pin, Site)
                
                If InStr(TheExec.CurrentChanMap, "CP") <> 0 Then
                    MeasureValue.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site) - ForceValI * (CP_Card_RAK.Pins(Pin).Value(site))
                Else
                    MeasureValue.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site) - ForceValI * (FT_Card_RAK.Pins(Pin).Value(site))
                End If

            Next Pin
        Next site
    End If
      
    '' 20150728 - Add return measure volt to main function.

    ReturnMeasVolt = MeasureValue
    
    Force_idx = TestSeqNum
    If UBound(ForceSequenceArray) = 0 Then
        Force_idx = 0
    End If

    Dim TestNameInput As String
    TestNameInput = "Volt_meas_" + CStr(TestSeqNum)
    
    '''20151103 print force condition
    Call Print_Force_Condition("v", MeasureValue)
    
    '' 20150721 - Test limit for force stored value
    Dim ForceIndex As Integer
    ForceIndex = 0
    If b_UseStoredForceVal = True Then
        For Each Pin In ForceByPin
            'For Each Site In TheExec.Sites
            If CurrentJobName_L Like "*char*" Then
                Disable_Inst_pinname_in_PTR
                    TheExec.Flow.TestLimit MeasureValue.Pins(Pin), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", ForceVal:=StoreForceI(ForceIndex).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Enable_Inst_pinname_in_PTR
            Else
                TheExec.Flow.TestLimit MeasureValue.Pins(Pin), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=StoreForceI(ForceIndex).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            
            End If
            'Next Site
            ForceIndex = ForceIndex + 1
        Next Pin
      '  Exit Function
    'End If
    ElseIf InStr(CUS_Str_MainProgram, "DDR_VOHL") <> 0 Then
         
            Dim HiLimitVal As Integer
            Dim LoLimitVal As Integer
            HiLimitVal = 0: LoLimitVal = 0
            If CUS_Str_MainProgram = "DDR_VOHL_1" Then
                HiLimitVal = 132: LoLimitVal = 108
            ElseIf CUS_Str_MainProgram = "DDR_VOHL_2" Then
                If TestSeqNumIdx = 0 Then HiLimitVal = 42: LoLimitVal = 38
                If TestSeqNumIdx = 1 Then HiLimitVal = 176: LoLimitVal = 144
            ElseIf CUS_Str_MainProgram = "DDR_VOHL_3" Then
                HiLimitVal = 264: LoLimitVal = 216
            Else
              
              If TestSeqNumIdx = 0 Then HiLimitVal = 132: LoLimitVal = 108
            End If
         
            If TestSeqNumIdx = 0 Then
                TheExec.Flow.TestLimit MeasureValue, , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
                TheExec.Flow.TestLimit MeasureValue.Math.Divide(ForceValI), LoLimitVal, HiLimitVal, scaletype:=scaleNone, Unit:=unitCustom, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI, customUnit:="ohm" 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            Else
                TheExec.Flow.TestLimit MeasureValue, , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
                TheExec.Flow.TestLimit MeasureValue.Math.Subtract(1.1).Divide(ForceValI).Abs, LoLimitVal, HiLimitVal, scaletype:=scaleNone, Unit:=unitCustom, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI, customUnit:="ohm" 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            End If
        
        ElseIf Flag_SingleLimit = True Then
            If LCase(currentJobName) Like "*char*" Then
                TheExec.Flow.TestLimit MeasureValue, LowLimitVal, HighLimitVal, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", ForceUnit:=unitAmp, ForceResults:=tlForceFlow, ForceVal:=ForceValI  'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            Else
                TheExec.Flow.TestLimit MeasureValue, LowLimitVal, HighLimitVal, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            End If
        '' 20150202 - Range Check
        If Range_Check_Enable_Word = True Then
            Call CheckRangesAndClamps(MeasureValue, "V", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
            gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
        End If
    
    Else
        If Mid(TestLimitPerPin_VIR, 1, 1) = "F" And UBound(ForceValByPin) = 0 Then
             If LCase(currentJobName) Like "*char*" Then
                        TheExec.Flow.TestLimit MeasureValue, , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", ForceVal:=ForceSequenceArray(Force_idx), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            Else
                        '' TheExec.Flow.TestLimit MeasureValue, , , ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=TestNameInput + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PATT_ExculdePath(Pat), forceVal:=ForceSequenceArray(Force_Idx), forceunit:=unitAmp, ForceResults:=tlForceFlow
                        TheExec.Flow.TestLimit MeasureValue, , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=ForceSequenceArray(Force_idx), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            End If
            '' 20150202 - Range Check
            If Range_Check_Enable_Word = True Then
                Call CheckRangesAndClamps(MeasureValue, "V", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
                gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
            End If
        Else
            IdxV = 0
            For p = 0 To MeasureValue.Pins.Count - 1
                If CurrentJobName_L Like "*char*" Then
                    Disable_Inst_pinname_in_PTR
                        TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", ForceVal:=ForceValByPin(IdxV), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    Enable_Inst_pinname_in_PTR
                Else
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=ForceValByPin(IdxV), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                End If
                
                If UBound(ForceValByPin) = 0 Then
                    IdxV = 0
                Else
                    IdxV = IdxV + 1
                End If
                
                '' 20150202 - Range Check
                If Range_Check_Enable_Word = True Then
                    TempMeasVal_PerPin(p).AddPin (MeasureValue.Pins(p))
                    TempMeasVal_PerPin(p).Pins(MeasureValue.Pins(p)) = MeasureValue.Pins(p)
                    Call CheckRangesAndClamps(TempMeasVal_PerPin(p), "V", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
                    gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
                End If
            Next p
        End If
    End If
       
    ' 20160105: Steph added for Refbuf test (Autogen) --- start
'    Call CUS_VFI_MeasureVolt(CUS_Str_MainProgram, MeasureValue, TestSeqNum, Pat)
    ' 20160105: Steph added for Refbuf test (Autogen) --- end
    
End Function


Public Function Meas_VIR_IO_Universal_func_GPIO_TTR(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
Optional Measure_Pin_PPMU As String, Optional ForceV As String, Optional ForceI As String, Optional MeasureI_Range As String = "0.05", _
Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
Optional InstSpecialSetting As InstrumentSpecialSetup = 0, Optional PPMU_TestLimit_TTR As Boolean = False, Optional RAK_Flag As Enum_RAK = 0, _
Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String, _
Optional Flag_SingleLimit As Boolean = False, Optional TestLimitPerPin_VIR As String = "FFF", _
Optional InterFunc_PrePat As InterposeName, Optional InterFuncArgs_PrePat As String, _
Optional InterFunc_PostPat As InterposeName, Optional InterFuncArgs_PostPat As String, _
Optional CharInputString As String, Optional ForceFunctional_Flag As Boolean = False, Optional MeasIGrpPinCnt As Integer = 0, Optional KeepEmptyLimit As Boolean = False, _
Optional CFG_GPIO_Pins As String, Optional MeasVGrpPinCnt As Integer = 0, Optional VDD_IO_1p2 As String, Optional VDD_IO_1p8 As String, Optional Validating_ As Boolean) As Long

''20191028, Add VDD_IO_1p2 and VDD_IO_1p8 for calculating differential voltage, Carter
''20150924 - Remove argument as below
''Remove ,Optional MeasOnHalt As Boolean
''Remove ,Optional InterFunc_SequenceN As InterposeName
''Remove ,Optional InterFuncArgs_SequenceN As String

''==================================================================================
'' 20150621 - Check with CCWu: FRCPortName As String, Optional DisableFRC As Boolean = False not use in this function
'' 20150717 - Impedance measurement by using 2 point measure method, Define "Z" for TestSequence - On going
''                - EX: Pin1, Pin2 + Pin3, Pin4     V1, V2 + V3, V4
''                - V1 and V2 use for Pin1 of impedence measurement
''                - V1 and V2 use for Pin2 of impedence measurement
'' 20150717 - Get I from previous item and apply the current value to next item, use enum for the feature
''                - EX: TestSequence: "V,V,V"
''                  If second V want to apply calcuated I value that Force I value argument should be "0,keyword,0"
'' 20150727 - MeasureI_Range is use for test sequence "I", "R" and "Z"
''==================================================================================
    
    Dim PatCount As Long
    Dim k As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, ForceISequenceArray() As String, ForceVSequenceArray() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim PatMeas As String
    Dim TestPinArrayIV() As String
    Dim TestIrange() As String
    Dim TestSeqNumIdx As Long
    Dim InDSPwave As New DSPWave
    Dim OutDspWave() As New DSPWave
    Dim ShowDec As String
    Dim ShowOut As String
    
    ''20141223
    Dim site As Variant
    Dim PattArray() As String
    Dim Pat As String
    Dim patt As Variant
    
    Dim HighLimitVal() As Double
    Dim LowLimitVal() As Double
    
    ''20150728
    Dim ReturnMeasVolt As New PinListData
    Dim FlowForLoopName() As String   ' Sequences : Code , Voltage , Loop Index
    
    Dim i, j As Integer

    ''20160821-Add judgement by 7.75mA for Fuse
    Dim CFGTestPins(0) As String
    Dim CFGTest_FirstSequence As Boolean
    CFGTest_FirstSequence = True
    
    On Error GoTo errHandler
    
    If Validating_ Then 'Carter, 20190315
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If

    Shmoo_Pattern = patset.Value

    Call tl_PinListDataSort(True)
    ''========================================================================================
    '' 20150121 - Range Check
    If Range_Check_Enable_Word = True Then         'Change to "Range_Check_Enable_Word" by Martin 20151225  for TTR
        If TheExec.DataManager.MemberIndex = 0 Then
            gl_UseLimitCheck_Counter = 0
        End If
    End If
    ''========================================================================================
   
    '' 20160111 - Check input condition for measure I, R and Z.
''    Call CheckCondition_Measure_I_R_Z(TestSequence, Measure_Pin_PPMU, ForceV, MeasureI_Range)
   
    TestSequenceArray = Split(TestSequence, ",")
    
    If ForceI = "" Then ForceI = 0
    If ForceV = "" Then ForceV = 0
    
    Call GetFlowTName
    
    
    
    ForceISequenceArray = Split(ForceI, "+")
    ForceVSequenceArray = Split(ForceV, "+")
        '''''''''''Apply DC spec'''''''''''
    For i = 0 To UBound(ForceVSequenceArray)
        ForceVSequenceArray(i) = EvaluateEachBlock(ForceVSequenceArray(i)) ''zhhuangf
    Next i
    '''''''''''Apply DC spec'''''''''''
    '' 20150812-Decompose DigCap_Pin by ","
    Dim DigCap_Pin_Ary() As String
    DigCap_Pin_Ary = Split(DigCap_Pin, ",")
    FlowForLoopName = Split(DigSrc_FlowForLoopIntegerName, ",")
    
    Char_Test_Name_Curr_Loc = 0 'init char datalog test name index
    
    TestPinArrayIV = Split((Measure_Pin_PPMU), "+")
    
    If MeasureI_Range = "" Then MeasureI_Range = "50e-3"
    
    TestIrange = Split(MeasureI_Range, "+")
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
  
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    
    If InterFunc_PrePat <> "" Then Call Interpose(InterFunc_PrePat, InterFuncArgs_PrePat)
    
    '20151028  CUS_MeasV_And_CalR -- TYCHENGG
    ''ex.  CUS_Str_MainProgram = "CalR;1.88;RVOH,RVOL,RVOH"
    ''========================================================================================
    Dim CUS_CalR_VDD As Double
    Dim CUS_CalR_Seq_Ary() As String
    Dim CUS_CalR_Arg_Ary() As String
    If (CUS_Str_MainProgram <> "") Then
        If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
            CUS_CalR_Arg_Ary = Split(CUS_Str_MainProgram, ";")
            CUS_CalR_VDD = CDbl(CUS_CalR_Arg_Ary(1))
            CUS_CalR_Seq_Ary = Split(CUS_CalR_Arg_Ary(2), ",")
        End If
    End If
    ''========================================================================================
    
    '' 20150625 - Apply Char setup
    If CharInputString <> "" Then
        Call SetForceCondition(CharInputString)
    End If
    
    TheHdw.Patterns(patset).Load
    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    If KeepEmptyLimit = True Then
        Call GetFlowSingleUseLimit_KeepEmpty(HighLimitVal, LowLimitVal)  ''20141223
    Else
        Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)  ''20141223
    End If
    For Each patt In PattArray
        Pat = CStr(patt)
        PatMeas = Pat
         
        If DigSrc_Sample_Size <> 0 Then
             
             '' 20150810 - Source dssc index by For opcode from flow table
            If DigSrc_FlowForLoopIntegerName <> "" Then        '20151201
                If FlowForLoopName(0) <> "" Then        '20151201
                    Call DSSCSrcBitFromFlowForLoop(FlowForLoopName(0), DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment)
                End If
            End If
            
            For Each site In TheExec.sites.Active
                ''20150708- Add customize string for digsrc data process
                Call Create_DigSrc_Data(DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPwave, site, CUS_Str_DigSrcData)
            Next site
            
            Call SetupDigSrcDspWave(PatMeas, DigSrc_pin, "Meas_src", DigSrc_Sample_Size, InDSPwave)
        End If

        ''20150812- Add program to setup multiple DigCap_Pin.
        If DigCap_Sample_Size <> 0 Then
            Dim DigCap_Pin_Num As Integer
            DigCap_Pin_Num = UBound(DigCap_Pin_Ary)
            ReDim OutDspWave(DigCap_Pin_Num) As New DSPWave
            For i = 0 To DigCap_Pin_Num
                TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test for " & DigCap_Pin_Ary(i) & " ========")
                OutDspWave(i).CreateConstant 0, DigCap_Sample_Size
                DigCap_Pin.Value = DigCap_Pin_Ary(i)
                Call DigCapSetup(Pat, DigCap_Pin, "Meas_cap", DigCap_Sample_Size, OutDspWave(i))
           Next i
        End If
          
        Call TheHdw.Patterns(Pat).start
        TestSeqNum = 0
        
        Dim FlowTestName() As String
        
        For Each Ts In TestSequenceArray
            ''20150907 - Only need CPUA_Flag_In_Pat to do control
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            TestOptLen = Len(Ts)
            
            TestSeqNumIdx = TestSeqNum
            For k = 1 To TestOptLen
                TestOption = Mid(Ts, k, 1)
                
                '' 20160106 - If "ForceFunctional_Flag" = True to let TestOption = "N" to make the test instance only run functional test
                If ForceFunctional_Flag = True Then
                    TestOption = "N"
                End If
                
                If (Measure_Pin_PPMU <> "") Then
                    Call Meas_VIR_IO_PreSetupBeforeMeasurement(TestPinArrayIV, TestSeqNumIdx)
                    
                    Select Case UCase(TestOption)
                    
                        Case "V"
                        
                            Call IO_HardIP_PPMU_Measure_V_TTR(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceISequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(TestSeqNum), LowLimitVal(TestSeqNum), TestLimitPerPin_VIR, ReturnMeasVolt, _
                                    InstSpecialSetting, RAK_Flag, CUS_Str_MainProgram, MeasVGrpPinCnt)
 
                             ''20151028  CUS_MeasV_And_CalR -- TYCHENGG
                            ''========================================================================================
                            If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
                                Call CUS_VIR_MainProgram_MeasV_CalR(TestPinArrayIV, TestSeqNum, CUS_CalR_Seq_Ary, ForceISequenceArray, ReturnMeasVolt, CUS_CalR_VDD)
                            End If
                            ''========================================================================================
                            
                        Case "I"
                            
                            If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
                            
                            Call IO_HardIP_PPMU_Measure_I_TTR(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(TestSeqNum), LowLimitVal(TestSeqNum), TestLimitPerPin_VIR, TestIrange, VDD_IO_1p2, VDD_IO_1p8, _
                                    CUS_Str_MainProgram, PPMU_TestLimit_TTR, MeasIGrpPinCnt, CFG_GPIO_Pins, CFGTest_FirstSequence)

                            If CFGTest_FirstSequence = True Then CFGTest_FirstSequence = False
                                                   
                        Case Else
                            TheExec.Datalog.WriteComment "Error Test Option, please select V, I"
                    
                    End Select
                    
                    Call Meas_VIR_IO_PostSetupAfterMeasurement(TestPinArrayIV, TestSeqNumIdx)
                End If
            Next k
            
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
            End If
        Next Ts
        
        If DebugPrintEnable = True Then
            TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & ""
        End If
                
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        PatCount = PatCount + 1
        
        'This function is no longer exist
'        If DigCap_Sample_Size <> 0 Then
'            Call HardIP_Digcap_Print(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, ShowDec, ShowOut, , DigCap_DSPWaveSetting) '''change
'        End If
    Next patt
    
    If InterFunc_PostPat <> "" Then Call Interpose(InterFunc_PostPat, InterFuncArgs_PostPat)
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Connect
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    
    If DisableFRC = True Then
        Call FreeRunclk_Enable(FRCPortName)
    End If

    Call HardIP_WriteFuncResult
    
    DebugPrintFunc patset.Value  ' print all debug information
    
    '' 20150728 - Print Char setup for power pins.
    If CurrentJobName_U Like "*CHAR*" Then
        If CharInputString <> "" Then
          '  Call PrintCharPowerSet(CharInputString)
        End If
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Meas_VIR_IO_Universal_func"
    If AbortTest Then Exit Function Else Resume Next
  
End Function


Public Function EvaluateEachBlock(ForceVSeq As String) As String ''zhhuangf
Dim i As Integer, j As Integer
    ''_VDDIO18_GRP_VAR_H
Dim ForceVarray() As String
If InStr(ForceVSeq, ",") Then
    Dim tempV1() As String
    Dim Variable_Spec As String
    ForceVarray = Split(ForceVSeq, ",")
    For i = 0 To UBound(ForceVarray)
        If InStr(ForceVarray(i), "*") Then
            tempV1 = Split(ForceVarray(i), "*")
            For j = 0 To UBound(tempV1)
                If InStr(tempV1(j), "_") Then
                    Variable_Spec = Right(tempV1(j), Len(tempV1(j)) - 1)
                    tempV1(j) = CStr(TheExec.specs.DC.Item(Variable_Spec).ContextValue)
                End If
                If j = 0 Then
                    ForceVarray(i) = tempV1(0)
                Else
                    ForceVarray(i) = ForceVarray(i) & "*" & tempV1(j)
                End If
            Next j
        ElseIf InStr(ForceVarray(i), "/") Then
            tempV1 = Split(ForceVarray(i), "/")
            For j = 0 To UBound(tempV1)
                If InStr(tempV1(j), "_") Then
                    Variable_Spec = Right(tempV1(j), Len(tempV1(j)) - 1)
                    tempV1(j) = CStr(TheExec.specs.DC.Item(Variable_Spec).ContextValue)
                End If
                If j = 0 Then
                    ForceVarray(i) = tempV1(0)
                Else
                    ForceVarray(i) = ForceVarray(i) & "/" & tempV1(j)
                End If
            Next j
        ElseIf InStr(ForceVSeq, "_") Then
                        Variable_Spec = Right(ForceVSeq, ForceVSeq - 1)
                        ForceVSeq = CStr(TheExec.specs.DC.Item(Variable_Spec).ContextValue)
        End If
        If i = 0 Then
            EvaluateEachBlock = Evaluate(ForceVarray(0))
        Else
            EvaluateEachBlock = EvaluateEachBlock & "," & Evaluate(ForceVarray(i))
        End If
    Next i
Else
    If InStr(ForceVSeq, "*") Then
        tempV1 = Split(ForceVSeq, "*")
        For j = 0 To UBound(tempV1)
            If InStr(tempV1(j), "_") Then
                Variable_Spec = Right(tempV1(j), Len(tempV1(j)) - 1)
                tempV1(j) = CStr(TheExec.specs.DC.Item(Variable_Spec).ContextValue)
            End If
            If j = 0 Then
                ForceVSeq = tempV1(0)
            Else
                ForceVSeq = ForceVSeq & "*" & tempV1(j)
            End If
        Next j
    ElseIf InStr(ForceVSeq, "/") Then
        tempV1 = Split(ForceVSeq, "/")
        For j = 0 To UBound(tempV1)
            If InStr(tempV1(j), "_") Then
                Variable_Spec = Right(tempV1(j), Len(tempV1(j)) - 1)
                tempV1(j) = CStr(TheExec.specs.DC.Item(Variable_Spec).ContextValue)
            End If
            If j = 0 Then
                ForceVSeq = ForceVSeq & tempV1(0)
            Else
                ForceVSeq = ForceVSeq & "/" & tempV1(j)
            End If
        Next j
    ElseIf InStr(ForceVSeq, "_") Then
        Variable_Spec = Right(ForceVSeq, ForceVSeq - 1)
        ForceVSeq = CStr(TheExec.specs.DC.Item(Variable_Spec).ContextValue)
    End If
    EvaluateEachBlock = Evaluate(ForceVSeq)
End If
End Function


Public Function GetFlowSingleUseLimit_KeepEmpty(ByRef d_HighLimitVal() As Double, ByRef d_LowLimitVal() As Double) As Double
    ' Get the limits info
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Dim HighLimitValArray() As String
    Dim LowLimitValArray() As String
    Dim HighLimitArraySize As Long
    Dim LowLimitArraySize As Long
    
    Dim i As Integer
    
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    
    
   ' TheExec.Flow.GetTestLimits
    If FlowLimitsInfo Is Nothing Then
   ' i = i
  '  End If
     ReDim d_HighLimitVal(0) As Double
     ReDim d_LowLimitVal(0) As Double
     ReDim HighLimitValArray(0) As String
     ReDim LowLimitValArray(0) As String
     d_HighLimitVal(0) = 0
     d_LowLimitVal(0) = 0
     HighLimitValArray(0) = 0
     LowLimitValArray(0) = 0
  Else
    Call FlowLimitsInfo.GetHighLimits(HighLimitValArray())
    Call FlowLimitsInfo.GetLowLimits(LowLimitValArray())
    
    HighLimitArraySize = UBound(HighLimitValArray)
    ReDim d_HighLimitVal(HighLimitArraySize) As Double
    LowLimitArraySize = UBound(LowLimitValArray)
    ReDim d_LowLimitVal(LowLimitArraySize) As Double
    
  End If
    For i = 0 To HighLimitArraySize
        If (HighLimitValArray(i)) = "" Then HighLimitValArray(i) = -123456.123456 '''ZHHUANGF 20160728
        d_HighLimitVal(i) = CDbl(HighLimitValArray(i))
    Next i
        
    For i = 0 To LowLimitArraySize
        If LowLimitValArray(i) = "" Then LowLimitValArray(i) = -123456.123456
        
        d_LowLimitVal(i) = CDbl(LowLimitValArray(i))
    Next i
       
End Function

Public Function EVS_Static_Power_Ramp(dc_spec As String, S_WaitTime As Double, Power_pin As String, Optional Step_number As Integer = 10, Optional Rising_Delay_time As Double = 0.02, Optional Looping_Contorl As Boolean = False, Optional Looping_Range As Double = 0.4, Optional Looping_Index_Name As String = "", Optional Looping_Max_Steps_Name As String = "", Optional Open_LatchUp_measure As Boolean = False, Optional Multi_Function As Boolean = False, Optional Mulit_EVS_Index As Integer = 1, Optional Test_time_breakdown As Boolean = False) As Long

On Error GoTo errHandler
    If Looping_Contorl And (Looping_Max_Steps_Name = "" Or Looping_Index_Name = "") Then
        TheExec.Datalog.WriteComment "Error!! Please make sure you fill in argument both Looping_Max_Steps_Name and Looping_Index_Name "
        GoTo errHandler
    End If
    Dim funcName As String:: funcName = "EVS_Static_Power_Ramp"
    Dim inst_name As String
    'Dim Pin As Variant
    Dim i As Long
    Dim spec As Variant
    Dim Spec_Var As String
    Dim Pin As Variant
    Dim power_pin_ary() As String
    Dim Dc_spec_type As String
    Dim EVS_Detail_value() As Pins_detail ''Define each pin information
    Dim Mulit_Loop_EVS As Integer
    Dim TExec_Before_Pat As Double
    Dim Total_time As Double
    Dim EVS_Sequence As String
    Dim EVS_Type As String
    'If Test_time_breakdown = True Then
        'Total_time = TheExec.Timer(0)
    'End If
    EVS_Sequence = "Total_time"
    Call Test_time_breakdown_Start(Total_time, Test_time_breakdown, EVS_Sequence)
    power_pin_ary = Split(Power_pin, ",")
    ReDim EVS_Detail_value(UBound(power_pin_ary))
    '// ----Multi EVS ----
    If Multi_Function = False Then
        Mulit_EVS_Index = 1
        EVS_Type = "Regular-EVS"
    Else
        Mulit_EVS_Index = Mulit_EVS_Index
        EVS_Type = "Multi-EVS"
    End If
    S_WaitTime = S_WaitTime / Mulit_EVS_Index
    
    'Dim preEVSseup As Long
    'preEVSseup = ProfileMarkEnter(2, "PreEVS Setup")
   'evs pre
    
    'TExec_Before_Pat = theexec.Timer(0)
    EVS_Sequence = "EVS_Pre_Setting"
    Call Test_time_breakdown_Start(TExec_Before_Pat, Test_time_breakdown, EVS_Sequence)
    Call EVS_Pre_Setting(dc_spec, power_pin_ary, EVS_Detail_value, Step_number, Looping_Contorl, Looping_Range, Looping_Index_Name, Looping_Max_Steps_Name, Open_LatchUp_measure)
    Call Test_time_breakdown_End(TExec_Before_Pat, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)
    'TExec_Before_Pat = theexec.Timer(TExec_Before_Pat)
    'theexec.DataLog.WriteComment "EVS_Pre_Setting : " + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
    'evs pre
    'ProfileMarkLeave (preEVSseup)
    If Open_LatchUp_measure = False Then 'Or Looping_Contorl = False Then
        TheHdw.Alarms.StartMonitoringAlarms
    End If
    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = 170
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width = 100
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 75
    TheExec.Datalog.ApplySetup

    For Mulit_Loop_EVS = 0 To Mulit_EVS_Index - 1
        'TExec_Before_Pat = TheExec.Timer(0)
        EVS_Sequence = "Evs_Ramp_UP"
        'TExec_Before_Pat = 0
        Dim TExec_up As Double
        Call Test_time_breakdown_Start(TExec_up, Test_time_breakdown, EVS_Sequence)
        Call Evs_Ramp_UPorDown(EVS_Detail_value, "UP", S_WaitTime, power_pin_ary, Step_number, Rising_Delay_time, Open_LatchUp_measure, Looping_Contorl, Looping_Index_Name, Looping_Max_Steps_Name)
        Call Test_time_breakdown_End(TExec_up, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)
        'TExec_Before_Pat = TheExec.Timer(TExec_Before_Pat) - S_WaitTime
        'TheExec.DataLog.WriteComment "EVS_Ramp_up" & Mulit_Loop_EVS + 1 & " : " + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
        'TheExec.DataLog.WriteComment "EVS Stress time: " & S_WaitTime
        DebugPrintFunc ""
        'TExec_Before_Pat = TheExec.Timer(0)
        EVS_Sequence = "Evs_Ramp_DOWN"
        Dim TExec_down As Double
        Call Test_time_breakdown_Start(TExec_down, Test_time_breakdown, EVS_Sequence)
        Call Evs_Ramp_UPorDown(EVS_Detail_value, "DOWN", S_WaitTime, power_pin_ary, Step_number, Rising_Delay_time, Open_LatchUp_measure, Looping_Contorl, Looping_Index_Name, Looping_Max_Steps_Name)
        Call Test_time_breakdown_End(TExec_down, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)
        'TExec_Before_Pat = TheExec.Timer(TExec_Before_Pat)
        'TheExec.DataLog.WriteComment "EVS_Ramp_down" & Mulit_Loop_EVS + 1 & " : " + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
    Next Mulit_Loop_EVS
    
        Dim pin_count As Integer
        Dim Print_Looping_Inf As String
        Dim Pins As String
        Dim EVS_Level As Double
        Dim EVS_Instance_name As String
        Dim Gatecheck As Boolean
        Dim EVS_Gate_check As String
        Dim Die_X_location As New SiteLong
        Dim Die_Y_location As New SiteLong
        Dim m_InstanceName As String
        m_InstanceName = TheExec.DataManager.instanceName

        Print_Looping_Inf = ""
        EVS_Instance_name = m_InstanceName
     If Looping_Contorl = True Then
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        TheHdw.Wait 0.002
        TheExec.Datalog.WriteComment "-------------------------EVS Collection mode Start-------------------------"
        TheExec.Datalog.WriteComment "Instance name:" & EVS_Instance_name
        For Each site In TheExec.sites
            pin_count = 0
            Print_Looping_Inf = ""
            Die_X_location = XCoord(site)
            Die_Y_location = YCoord(site)
            For Each Pin In power_pin_ary
                'If TheHdw.DCVS.Pins(Pin).Gate = False Then
                If EVS_Detail_value(pin_count).Gate_check(site) = False Then
                    EVS_Gate_check = "Alarm:True"
                Else
                    EVS_Gate_check = "Alarm:False"
                End If
                Pins = EVS_Detail_value(pin_count).PinName
                EVS_Level = EVS_Detail_value(pin_count).LatchUp_Final_Value
                Print_Looping_Inf = Print_Looping_Inf + Pins & ":" & EVS_Level & ";" & EVS_Gate_check & ";" ', Stress time:" & S_WaitTime & "(msec);"
                pin_count = pin_count + 1
            Next Pin
            TheExec.Datalog.WriteComment "EVS results: SITE" & site & ";" & "X,Y: " & Die_X_location & "," & Die_Y_location & ";" & Print_Looping_Inf & EVS_Type    '& Chr(10)
        Next site
        'TheExec.DataLog.WriteComment "EVS pin's Level: " & Print_Looping_Inf
        TheExec.Datalog.WriteComment "-------------------------EVS Collection mode End-------------------------"
    End If
    
    EVS_Sequence = "Total_time"
    Call Test_time_breakdown_End(Total_time, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)
    
    If Open_LatchUp_measure = False Then 'Or Looping_Contorl = False Then
        TheHdw.Alarms.StopMonitoringAlarms
        TheHdw.Alarms.CloseAlarmWindow
    End If
    
    'Total_time = TheExec.Timer(Total_time)
    'TheExec.DataLog.WriteComment "Total test time : " + Format(Total_time * 1000#, "##0.000") + " msec"

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

Public Function EVS_Pre_Setting(dc_spec As String, power_pin_ary() As String, ByRef EVS_Detail_value() As Pins_detail, Optional Step_number As Integer = 10, Optional Looping_Control As Boolean, Optional Looping_Range As Double = 0.4, Optional Looping_Index_Name As String, Optional Looping_Max_Steps_Name As String, Optional Open_LatchUp_measure As Boolean)
    Dim spec As Variant, Spec_Var As String, Pin As Variant, Dc_spec_type As String
    Dim Step_value As Double
    Dim pin_count As Integer
    Dim End_Value As Double
    Dim Start_Value As Double
    Dim diff_value As Double
    Dim Looping_Index As Double, Looping_Max_Step As Double, Looping_Step_range As Double
    Dim Gap_Persent As Double
    Dim EVS_Index As Integer, Evs_End_Index As Integer
    pin_count = 0  '  initial value
    Evs_End_Index = Step_number
 
    '\\\\\\\\\\\\\\\\\\\Calculate EVS Ramp up/down information\\\\\\\\\\\\\\\\\\\
    For Each Pin In power_pin_ary
'        If Open_LatchUp_measure Then
'            TheHdw.DCVS.Pins(Pin).Meter.mode = tlDCVSMeterCurrent
'            TheHdw.DCVS.Pins(Pin).Meter.CurrentRange = 1
'        End If
'        Dim DC_Spec_Var As String
'        If Pin Like "*SRAM*" Then
'            DC_Spec_Var = "_VAR_"
'        Else
'            DC_Spec_Var = "_VOP_VAR_"
'        End If
        Spec_Var = Pin & "_VAR" '& Dc_spec_type
        Start_Value = TheHdw.DCVS.Pins(Pin).Voltage.Value ''Start Voltage is read by hardware
        End_Value = TheExec.specs.DC.Item(Spec_Var).Categories.Item(dc_spec).max.Value

        ''//////////////////////////Looping Function for EVS experiment//////////////////////////
        '' If we want to apply this method , we need to additionly change the EVS flow and assign varient at main flow
        ''
        '' Example as below:
        ''//////// Main flow ////////
        '' |     Opcode     |      Parameter     |
        '' | assign-integer |     Max_step 5     | <= Assign integer at IGXL Main flow
        '' | create-integer | EVS_Looping_INDEX  | <= Create integer at IGXL Main flow
        ''//////// Main flow ////////
        
        ''//////// EVS flow ////////
        '' |     Opcode    |                             Parameter                                 |
        '' |      For      | EVS_Looping_INDEX=0; EVS_Looping_INDEX< Max_Step; EVS_Looping_INDEX++ | <= For index for looping ,Start at 0
        '' |      test     |                         CpuSaEVS_Static_*pp**                         | <= Run Pattern
        '' |      test     |                       EVS_Static_Power_Up_CpuSa                       | <= Power up
        '' |      test     |                      EVS_Static_Power_Down_CpuSa                      | <= Power down
        '' |      Next     |                          EVS_Looping_INDEX                            |
        ''//////// EVS flow ////////
        
        '' Instance setting for "EVS_Static_Power_Up(Down)_CpuSa
        '' Addtional setting argument Looping_Control = true, Looping_Index_Name = "EVS_Looping_INDEX" and Looping_Max_Steps_Name = "Max_Step"
'        If Looping_Control And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then  ''Only need to change final value of Power Up
'            Looping_Index = TheExec.Flow.var(Looping_Index_Name).Value '' get index from flow
'            Looping_Max_Step = TheExec.Flow.var(Looping_Max_Steps_Name).Value '' get max step from flow
'            End_Value = Start_Value + (Looping_Index + 1) * ((End_Value - Start_Value) / Looping_Max_Step) ''' update end value for each looping
'        End If
        If Looping_Control And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then  ''Only need to change final value of Power Up
            Looping_Index = TheExec.Flow.var(Looping_Index_Name).Value '' get index from flow
            Looping_Max_Step = TheExec.Flow.var(Looping_Max_Steps_Name).Value '' get max step from flow
            Looping_Step_range = Looping_Range / (Looping_Max_Step - 1) '' get max step from flow
            End_Value = TheExec.specs.DC.Item(Spec_Var).Categories.Item(dc_spec).max.Value ''' update end value for each looping
            'Start_Value = End_Value - Looping_Range
            Dim Looping_Start_Voltage As Double
            Looping_Start_Voltage = End_Value - Looping_Range
            End_Value = Looping_Start_Voltage + Looping_Step_range * Looping_Index
        End If
        If Not (Looping_Control) And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then
        '' You can keep the opcode "For" on IGXL flow and trun looping_Control into false.
        '' If you still have Looping_Max_Steps_Name , it will set it to 1 which means it won't loop on the IGXL flow anymore.
            TheExec.Flow.var(Looping_Max_Steps_Name).Value = 1
'        ElseIf Not (Looping_Control) And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then
'            TheExec.Flow.var(Looping_Max_Steps_Name).Value = 1
        End If
        ''//////////////////////////Looping Function for EVS experiment//////////////////////////
        
        EVS_Detail_value(pin_count).LatchUp_Final_Value = Format(End_Value, "0.00")
        diff_value = End_Value - Start_Value
        Gap_Persent = Abs(diff_value / Start_Value) '''+-0.5%
        If Gap_Persent < 0.005 Then 'small than 0.5 % pin don't need to rise at all
            EVS_Detail_value(pin_count).Pin_rise = False
        Else
        
            Step_value = diff_value / Evs_End_Index ' calculate step value by end index

            EVS_Detail_value(pin_count).Start_voltage = Start_Value
            EVS_Detail_value(pin_count).Start_voltage_up = Start_Value
            EVS_Detail_value(pin_count).Start_voltage_down = End_Value
            
            EVS_Detail_value(pin_count).Step_value = Step_value
            EVS_Detail_value(pin_count).Step_value_up = Step_value
            EVS_Detail_value(pin_count).Step_value_down = Step_value * (-1)
            
            EVS_Detail_value(pin_count).Pin_rise = True
            EVS_Detail_value(pin_count).PinName = CStr(Pin)
            'Pin_detail_dict.Add CStr(Pin), Temp_value
        End If
        pin_count = pin_count + 1
    Next Pin
    '\\\\\\\\\\\\\\\\\\\ Calculate EVS Ramp up/down information\\\\\\\\\\\\\\\\\\\

End Function

Public Function Evs_Ramp_UPorDown(EVS_Detail_value() As Pins_detail, Direction As String, S_WaitTime As Double, power_pin_ary() As String, Optional Step_number As Integer = 10, Optional Rising_Delay_time As Double = 0.02, Optional Open_LatchUp_measure As Boolean = False, Optional Looping_Control As Boolean = False, Optional Looping_Index_Name As String = "", Optional Looping_Max_Steps_Name As String = "")
    Dim pin_count As Integer
    Dim PowerPin_Final_Value As New PinListData
    Dim ForceV As Double
    Dim LatchUp_measure_Value As New SiteDouble
    Dim EVS_Index As Integer, Evs_End_Index As Integer
    Dim Gatecheck As Boolean
    Dim power_pin_value As Double
    Dim Latch_Up_name As String
    Dim Pin As Variant
    Dim EVS_Index_Value As Integer
    Dim EVS_End_Index_Value As Integer:: Evs_End_Index = Step_number
    Dim ifold As Double
                Dim m_InstanceName As String
    m_InstanceName = TheExec.DataManager.instanceName

    For EVS_Index = 1 To Evs_End_Index
        If UCase(Direction) = "UP" Then
            pin_count = 0 '' reset pin index
        ElseIf UCase(Direction) = "DOWN" Then
            pin_count = UBound(power_pin_ary())
        End If
        
        For Each Pin In power_pin_ary
            Dim PinName As String
            'TheHdw.DCVS.Pins(Pin).CurrentRange = ifold
            'ifold = TheHdw.DCVS.Pins(Pin).CurrentLimit.Source.FoldLimit.Level
            If UCase(Direction) = "UP" Then
                EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_up
                EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_up
            ElseIf UCase(Direction) = "DOWN" Then
                EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_down
                EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_down
            End If
            
            PinName = EVS_Detail_value(pin_count).PinName
            If EVS_Detail_value(pin_count).Pin_rise Then
                ForceV = EVS_Detail_value(pin_count).Start_voltage + EVS_Index * EVS_Detail_value(pin_count).Step_value
                TheHdw.DCVS.Pins(PinName).Voltage.Value = ForceV
        
                If EVS_Index = Evs_End_Index Then
                    PowerPin_Final_Value.AddPin (PinName)
                    PowerPin_Final_Value.Pins(PinName) = TheHdw.DCVS.Pins(PinName).Voltage.Value
                End If
            End If
            
            If Open_LatchUp_measure Then
                TheHdw.DCVS.Pins(PinName).Meter.mode = tlDCVSMeterCurrent
                If Pin = "VDD_SOC" Then
                    TheHdw.DCVS.Pins(PinName).Meter.CurrentRange = 15
                Else
                    TheHdw.DCVS.Pins(PinName).Meter.CurrentRange = 1
                End If
                           
          '' Print out measure current value if want to collect the Latch up data
                LatchUp_measure_Value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                Latch_Up_name = m_InstanceName & "_" & "Latch_up_data_" & Replace(CStr(EVS_Detail_value(pin_count).LatchUp_Final_Value), ".", "p") & "V"
                TheExec.Flow.TestLimit resultVal:=LatchUp_measure_Value, PinName:=PinName, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Latch_Up_name, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
            End If
         '' Print out measure current value if want to collect the Latch up data

            If UCase(Direction) = "UP" Then
                pin_count = pin_count + 1
            ElseIf UCase(Direction) = "DOWN" Then
               pin_count = pin_count - 1
            End If
        Next Pin
        
        TheHdw.Wait Rising_Delay_time 'delay time of each ramp up
        
    Next EVS_Index
    '\\\\\\\\\\\\\\\\\\\ End Power up/Down \\\\\\\\\\\\\\\\\
    pin_count = 0 'reset pin count
    '///////////stress time after power up///////////
    If UCase(Direction) = "UP" Then
        TheHdw.Wait S_WaitTime
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "-------------------------EVS Power ramp start-------------------------"
        TheExec.Flow.TestLimit S_WaitTime, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, Tname:="EVS stress time", customUnit:="sec"  'BurstResult=1:Pass
    '///////////Print out final value of each pin (Power up and down)///////////
        For Each Pin In power_pin_ary
            If EVS_Detail_value(pin_count).Pin_rise Then
                TheExec.Flow.TestLimit PowerPin_Final_Value.Pins(Pin), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, customUnit:="V"   'BurstResult=1:Pass
            End If
            pin_count = pin_count + 1
        Next Pin
    '///////////Print out final value of each pin (Power up and down)///////////
    '///////////Alarm check after stress during power up///////////
        For Each site In TheExec.sites
            pin_count = 0
            For Each Pin In power_pin_ary
            Gatecheck = TheHdw.DCVS.Pins(Pin).Gate
            EVS_Detail_value(pin_count).Gate_check(site) = Gatecheck
                If TheHdw.DCVS.Pins(Pin).Gate = False Then
'                    Dim power_pin_value As Double
                    power_pin_value = TheHdw.DCVS.Pins(Pin).Voltage.Value
                    TheExec.Flow.TestLimit power_pin_value, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, customUnit:="V", PinName:=Pin, Tname:="Vlotage_PatternAlarm_After_Stress"
                    TheExec.Flow.TestLimit Gatecheck, 1, 1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_Stress", PinName:=Pin
                End If
                pin_count = pin_count + 1
            Next Pin
        Next site

'        ///////////Alarm check after stress during power up///////////
    ElseIf UCase(Direction) = "DOWN" Then
    '///////////Alarm check after stress during power down///////////
        Dim All_Core_Power() As String
        Dim Core_power As String
        Dim All_power As Variant
        Dim CorePower_Cnt As Long
        'Core_power = "VDD_AVE,VDD_DCS_DDR,VDD_DISP,VDD_ECPU,VDD_GPU,VDD_PCPU,VDD_SOC,VDD_SRAM_ANE,VDD_SRAM_CPU,VDD_SRAM_GPU,VDD_SRAM_SOC,VDD_LOW,VDD_FIXED"
        'All_Core_Power() = Split(Core_power, ",")
        TheExec.DataManager.DecomposePinList "CorePower", All_Core_Power, CorePower_Cnt

        For Each site In TheExec.sites
            'For Each Pin In power_pin_ary
            For Each All_power In All_Core_Power
            Gatecheck = TheHdw.DCVS.Pins(All_power).Gate
                If TheHdw.DCVS.Pins(All_power).Gate = False Then
                    TheExec.Flow.TestLimit Gatecheck, -1, -1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_EVS_ramp_down", PinName:=All_power
                End If
            Next All_power
        Next site
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "--------------------------EVS Power ramp end--------------------------"
    
    End If
    '///////////stress time after power up///////////
End Function

Public Function Test_time_breakdown_End(ByRef Timer As Double, Test_time_breakdown As Boolean, EVS_Sequence As String, Mulit_Loop_EVS As Integer, S_WaitTime As Double)
    If Test_time_breakdown = True Then
        Timer = TheExec.Timer(Timer)
        If EVS_Sequence = "EVS_Pre_Setting" Then
            TheExec.Datalog.WriteComment "EVS_Pre_Setting : " + Format(Timer * 1000#, "##0.000") + " msec"
        End If
        If EVS_Sequence = "Evs_Ramp_UP" Then
            TheExec.Datalog.WriteComment "EVS_Ramp_up" & Mulit_Loop_EVS + 1 & " : " + Format((Timer - S_WaitTime) * 1000#, "##0.000") + " msec"
        End If
        If EVS_Sequence = "Evs_Ramp_DOWN" Then
            TheExec.Datalog.WriteComment "EVS_Ramp_down" & Mulit_Loop_EVS + 1 & " : " + Format(Timer * 1000#, "##0.000") + " msec"
        End If
        If EVS_Sequence = "Total_time" Then
            TheExec.Datalog.WriteComment "Total test time : " + Format(Timer * 1000#, "##0.000") + " msec"
        End If
    End If
End Function

Public Function Test_time_breakdown_Start(ByRef Timer As Double, Test_time_breakdown As Boolean, EVS_Sequence As String)
    If Test_time_breakdown = True Then
        Timer = TheExec.Timer(0)
    End If
End Function

Public Function auto_Checkboard_EVS_Probe_Location()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Check_Probe_Location"
    Dim site As Variant
    
    For Each site In TheExec.sites
        
        Dim Sum_XY As New SiteLong
        Dim Flag_Checkboard As String
        Dim Flag_Reverse_checkboard As String
        Flag_Checkboard = "Checkboard"
        Flag_Reverse_checkboard = "Reverse_Checkboard"
        Sum_XY(site) = XCoord(site) + YCoord(site)
        If Sum_XY Mod 2 = 0 Then
            TheExec.sites(site).FlagState(Flag_Checkboard) = logicTrue
            TheExec.sites(site).FlagState(Flag_Reverse_checkboard) = logicFalse
            'theexec.Flow.SiteFlag(site, Flag_Checkboard) = 1
            'theexec.Flow.SiteFlag(site, Flag_Reverse_checkboard) = 0
        Else
            TheExec.sites(site).FlagState(Flag_Checkboard) = logicFalse
            TheExec.sites(site).FlagState(Flag_Reverse_checkboard) = logicTrue
            'theexec.Flow.SiteFlag(site, Flag_Checkboard) = 0
            'theexec.Flow.SiteFlag(site, Flag_Reverse_checkboard) = 1
        End If
        
    Next site

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function
