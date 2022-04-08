Attribute VB_Name = "VBT_LIB_HardIP_JitterEye"
Option Explicit

Public Function Time_Measure_kit_UP1600(pat_name As Pattern, pin_name As PinList, jitter_meas As Boolean, eye_meas As Boolean, _
    Optional CPUA_Flag_In_Pat As Boolean, Optional TestSequence As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
    Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String = "", _
    Optional MeasPin_Differential As PinList, _
    Optional MeasF_WalkingStrobe_Flag As Boolean, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, Optional MeasF_WalkingStrobe_StepVoltage As Double, _
    Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double) As Long

''duty_freq_meas As Boolean, log_on As Boolean,

    Dim DSPCapture As New PinListData
    Dim RJ As New PinListData, DDJ As New PinListData, Tj As New PinListData
    Dim measUI As New PinListData
    Dim Tr As New PinListData, Tf As New PinListData
    Dim Eye20 As New PinListData, Eye50 As New PinListData, Eye80 As New PinListData
    Dim dspStatus As New PinListData
    
    Dim RJ_J As New PinListData, DDJ_J As New PinListData
    Dim MeasUI_J As New PinListData
    Dim dspStatus_J As New PinListData
    
    Dim dutycycle As New PinListData
    Dim freq As New PinListData
    
    Dim site As Variant
    Dim Pin As Variant
    Dim InDSPwave As New DSPWave, OutDspWave As New DSPWave
    
    Dim Pat As String
''    Dim ShowDec As String, ShowOut As String
    Dim PattArray() As String
    Dim patt As Variant
    Dim PatCount As Long
    
    Dim index As Long
    
    On Error GoTo errHandler
    
    Dim TestSequenceArray() As String
    TestSequenceArray = Split(TestSequence, ",")
    Dim Ts As Variant, TestOption As Variant
    Dim TestOptLen As Integer
    Dim i As Long, j As Long, k As Long
    Dim TestLimitPinName As String

   'setup and run pattern
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    Call HardIP_InitialSetupForPatgen
    TheHdw.Patterns(pat_name).Load
    Call PATT_GetPatListFromPatternSet(pat_name.Value, PattArray, PatCount)
    
    ''20161107-Return sweep test name
    Dim Rtn_SweepTestName As String
    Rtn_SweepTestName = ""
    
    For Each patt In PattArray
    
        Pat = CStr(patt)
    
        TheHdw.Patterns(Pat).Load
        
        Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, _
                                               DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPwave, Rtn_SweepTestName)
        
        Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
        
        Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
        
        '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Patterns(Pat).start
        Else
            Call TheHdw.Patterns(Pat).Test(pfAlways, 0)
        End If

        TheHdw.Wait 0.5
       
        For Each Ts In TestSequenceArray
        
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            ''20170105-Add freq walking strobe function to fine turn Voh/Vol(single end) Vt(differential) before jitter measurement
            If pin_name <> "" Then
                If MeasF_WalkingStrobe_Flag = True Then
                    Call Freq_WalkingStrobe_Meas_VOHVOL(pin_name, MeasF_WalkingStrobe_StartV, MeasF_WalkingStrobe_EndV, MeasF_WalkingStrobe_StepVoltage, MeasF_WalkingStrobe_BothVohVolDiffV, MeasF_WalkingStrobe_interval, MeasF_WalkingStrobe_miniFreq)
                End If
            ElseIf MeasPin_Differential <> "" Then
                If MeasF_WalkingStrobe_Flag = True Then
                    Call Freq_WalkingStrobe_Meas_VOD_Diff(MeasPin_Differential, MeasF_WalkingStrobe_StartV, MeasF_WalkingStrobe_EndV, MeasF_WalkingStrobe_StepVoltage, MeasF_WalkingStrobe_BothVohVolDiffV, MeasF_WalkingStrobe_interval, MeasF_WalkingStrobe_miniFreq)
                End If
            End If
            
            DSPCapture = TheHdw.Digital.Jitter.SingleDSPWaves
            TheHdw.Digital.Jitter.TimeoutEnable = True          'enable time out
            TheHdw.Digital.Jitter.TimeOut = 10                      'setup time out value
       
            TheHdw.Wait 0.1
    
            'start time measure
            TheHdw.Digital.Jitter.start     'start jitter
            TheHdw.Digital.Jitter.Wait
        
            Call TheHdw.Digital.Jitter.UpdateSingleDSPWaves(DSPCapture)     'update the dspwave
        
            'measure jitter block
            If jitter_meas = True Then
                Call rundsp.LoopJitterMeas(DSPCapture, RJ_J, DDJ_J, MeasUI_J, dspStatus_J, dutycycle, freq)     'DSP
            End If
        
            'measure eye block
            If eye_meas = True Then
                Call rundsp.LoopEyeMeas(DSPCapture, RJ, DDJ, Tj, measUI, Tr, Tf, Eye20, Eye50, Eye80, dspStatus)        'DSP
            End If
        
''            'measure Duty Freq block
''            If duty_freq_meas = True Then
''                'dutycycle = TheHdw.PPMU.Pins(pin_name).Read(tlPPMUReadMeasurements, 1)   ' for assign pin information to dutycycle
''                'freq = TheHdw.PPMU.Pins(pin_name).Read(tlPPMUReadMeasurements, 1)        ' for assign pin information to dutycycle
''                'Call rundsp.duty_freq_meas(DSPCapture, dutycycle, freq)    'DSP
''            End If
            
''            ''generate raw data
''            '' 20161116 pinlist data seems not workable
''            If log_on = True Then
''                For Each Pin In DSPCapture.Pins
''                    For Each Site In TheExec.Sites
''                        TheHdw.Digital.Jitter.FileExport DSPCapture.Pins(Pin).Value(Site), TheExec.DataManager.InstanceName & "_raw_data_" & DSPCapture.Pins(Pin) & "_site" & Site & ".txt"
''                    Next Site
''                Next Pin
''            End If
            
            'judgment
            
            If pin_name <> "" Then
                TestLimitPinName = pin_name
            ElseIf MeasPin_Differential <> "" Then
                TestLimitPinName = MeasPin_Differential
            End If
            
            Dim testName As String
            testName = ""
            
            If jitter_meas = True Then
                For index = 0 To (RJ_J.Pins.Count - 1)
                    Call TheExec.Flow.TestLimit(resultVal:=RJ_J.Pins(index), scaletype:=scalePico, Unit:=unitTime, Tname:=testName & "Jitter_RJ", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                    Call TheExec.Flow.TestLimit(resultVal:=DDJ_J.Pins(index), scaletype:=scalePico, Unit:=unitTime, Tname:=testName & "Jitter_DDJ", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                    Call TheExec.Flow.TestLimit(resultVal:=MeasUI_J.Pins(index), scaletype:=scalePico, Unit:=unitTime, Tname:=testName & "Jitter_UI", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                    Call TheExec.Flow.TestLimit(resultVal:=dutycycle.Pins(index), Unit:=unitCustom, Tname:=testName & "Duty_cycle", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                    Call TheExec.Flow.TestLimit(resultVal:=freq.Pins(index), scaletype:=scaleMega, Unit:=unitHz, Tname:=testName & "Freq", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                Next index
            End If
            
            
            If eye_meas = True Then
                Call TheExec.Flow.TestLimit(resultVal:=RJ, scaletype:=scalePico, Unit:=unitTime, Tname:="Eye_RJ", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                Call TheExec.Flow.TestLimit(resultVal:=DDJ, scaletype:=scalePico, Unit:=unitTime, Tname:="Eye_DDJ", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                Call TheExec.Flow.TestLimit(resultVal:=measUI, scaletype:=scalePico, Unit:=unitTime, Tname:="Eye_Width", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                Call TheExec.Flow.TestLimit(resultVal:=Eye20, scaletype:=scalePico, Unit:=unitTime, Tname:="Eye_Widthhigh", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                Call TheExec.Flow.TestLimit(resultVal:=Eye50, scaletype:=scalePico, Unit:=unitTime, Tname:="Eye_Widthmid", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                Call TheExec.Flow.TestLimit(resultVal:=Eye80, scaletype:=scalePico, Unit:=unitTime, Tname:="Eye_Widthlow", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
            End If
        
''            If duty_freq_meas = True Then
''                Call TheExec.Flow.TestLimit(resultVal:=dutycycle, unit:=unitCustom, Tname:="Duty_cycle", PinName:=pin_name, ForceResults:=tlForceNone)
''                Call TheExec.Flow.TestLimit(resultVal:=freq, ScaleType:=scaleMega, unit:=unitHz, Tname:="Freq", PinName:=pin_name, ForceResults:=tlForceNone)
''            End If
           
''            TestSeqNum = TestSeqNum + 1
                
            If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Next Ts
        
        TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end

        If DigCap_Sample_Size <> 0 Then
            Dim DigCapPinAry() As String, NumberPins As Long
            Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
            
            If NumberPins > 1 Then
                Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
                Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins)
            ElseIf NumberPins = 1 Then
                Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
                Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave)
                Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
            End If
        End If
    
    Next patt
    
    '' 20160713 - Call write functional result if cpu flag in pattern
    If (CPUA_Flag_In_Pat) Then
        Call HardIP_WriteFuncResult
    End If
  
    Exit Function
  
errHandler:
    TheExec.Datalog.WriteComment "error in Time_Measure_kit_UP1600"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DDR_RO_Time_Measure_KIT_UP1600(pat_name As Pattern, MeasPin_SingleEnd As PinList, jitter_meas As Boolean, eye_meas As Boolean, _
    Optional CPUA_Flag_In_Pat As Boolean, Optional TestSequence As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
    Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String = "", _
    Optional MeasPin_Differential As String, _
    Optional MeasF_WalkingStrobe_Flag As Boolean, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, Optional MeasF_WalkingStrobe_StepVoltage As Double, _
    Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double, _
    Optional FreqSeq_Start As Long, Optional FreqSeq_Stop As Long, Optional Interpose_PrePat As String, _
    Optional TestNameInput As String) As Long

    Dim DSPCapture As New PinListData
    Dim RJ As New PinListData, DDJ As New PinListData, Tj As New PinListData
    Dim measUI As New PinListData
    Dim Tr As New PinListData, Tf As New PinListData
    Dim Eye20 As New PinListData, Eye50 As New PinListData, Eye80 As New PinListData
    Dim dspStatus As New PinListData
    
    Dim RJ_J As New PinListData, DDJ_J As New PinListData, MeasUI_J As New PinListData
    Dim dspStatus_J As New PinListData
    Dim DutyCycle_J As New PinListData, Freq_J As New PinListData
    Dim PWHigh_J As New PinListData, PWLow_J As New PinListData, MeasuredPeriod_J As New PinListData
    
    Dim site As Variant
    Dim Pin As Variant
    Dim InDSPwave As New DSPWave
    Dim OutDspWave As New DSPWave
     
    Dim Pat As String
    Dim PattArray() As String
    Dim patt As Variant
    Dim PatCount As Long
    Dim Org_Test_Number As Long
    Dim Set_Align_Test_Number As Long
    
    On Error GoTo errHandler
    
    Dim TestSequenceArray() As String
    Dim Ts As Variant, TestOption As Variant

    Dim i As Long, j As Long, k As Long
    Dim TestLimitPinName As String
    
''    Dim TestNameSplit() As String
''    Dim TestNameStart As String, TestNameEnd As String
''    TestNameSplit = Split(TestNameInput, ",")
''    TestNameStart = TestNameSplit(0)
''    TestNameEnd = TestNameSplit(1)
    
   'setup and run pattern
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    Call HardIP_InitialSetupForPatgen
    TheHdw.Patterns(pat_name).Load
    Call PATT_GetPatListFromPatternSet(pat_name.Value, PattArray, PatCount)
    
    ''20161107-Return sweep test name
    Dim Rtn_SweepTestName As String
    Rtn_SweepTestName = ""
    
''    Dim TestName As String
''    TestName = "Jitter_"
    Dim index As Long
    Dim TestOptLen As Integer
    Dim MeasPinAry_F_SingleEnd() As String, MeasureF_Pin_SingleEnd As New PinList
    Dim MeasPinAry_F_Differential() As String, MeasureF_Pin_Differential As New PinList
    MeasPinAry_F_SingleEnd = Split(MeasPin_SingleEnd, "+")
    MeasPinAry_F_Differential = Split(MeasPin_Differential, "+")
    
    Dim TestSeqNum As Integer
    Dim Rtn_MeasFreq As New PinListData
    Dim d_Freq As Double
    
    Dim b_FirstTime As Boolean
    b_FirstTime = True
    
    Dim SeqMaxNum As Long
    Dim SeqIndex As Long
    SeqMaxNum = FreqSeq_Stop - FreqSeq_Start
    
    Dim ExecMaxNum As Long
    Dim ExecIndex As Long
    ExecMaxNum = 20
    
    Dim Dict_F_KeyName As String
    Dict_F_KeyName = ""
    
    ''20170322-Store MeasF mid value for VT
    Dim SplitFreqVtValue() As String
    Dim DictKey_StoreVT As String
    Dim Dict_VT_Value As New SiteDouble
    If CUS_Str_MainProgram <> "" Then
        SplitFreqVtValue = Split(CUS_Str_MainProgram, ":")
        If UCase(SplitFreqVtValue(0)) = "SETUP_STORE_VT" Then
            DictKey_StoreVT = SplitFreqVtValue(1)
            Dict_VT_Value = GetStoredMeasurement(DictKey_StoreVT)
        End If
    End If
    
    For SeqIndex = 0 To SeqMaxNum
        TestSequenceArray = Split(TestSequence, ",")
        TestSequenceArray(FreqSeq_Start + SeqIndex) = "F"
        Dict_F_KeyName = "F" & CStr(FreqSeq_Start + SeqIndex)
    '===========================================================================================================================================
    '(1)For TestNumber align , the code is assembled with (2)
    '===========================================================================================================================================
        For Each site In TheExec.sites.Active
            Org_Test_Number = TheExec.sites(site).TestNumber
            Exit For
        Next site
        For Each site In TheExec.sites
            TheExec.sites(site).TestNumber = Org_Test_Number
            For ExecIndex = 1 To ExecMaxNum
                
                Rtn_MeasFreq = GetStoredMeasurement(Dict_F_KeyName)
                
                TheExec.Flow.TestLimit resultVal:=Rtn_MeasFreq.Pins(1).Value(site), Unit:=unitHz, Tname:="FreqFromDict_" & SeqIndex & "_" & TestNameInput, ForceResults:=tlForceNone
                
                If Rtn_MeasFreq.Pins(1).Value(site) > 100000000# Then
                    d_Freq = Rtn_MeasFreq.Pins(1).Value(site)
                    
                Else
                    d_Freq = 100000000#
                    TheExec.Datalog.WriteComment ("Site " & site & " cant not get reasonable freq value")
                End If
                
                TheHdw.Digital.Jitter.BitPeriod = (1 / ((d_Freq) * 2))
                TheHdw.Digital.Jitter.ApplyJitterSet
                TheExec.Datalog.WriteComment ("ExecIndex=" & ExecIndex & " Bit period = " & CStr(TheHdw.Digital.Jitter.BitPeriod) & " Freq = " & CStr(1 / (2 * (TheHdw.Digital.Jitter.BitPeriod))))
                
                For Each patt In PattArray
                
                    Pat = CStr(patt)
                    TheHdw.Patterns(Pat).Load
                    
                    If SeqIndex = 0 And ExecIndex = 0 Then
                        Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, _
                                                               DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPwave, Rtn_SweepTestName)
                        
                        Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                        
                        Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
                    End If
                    
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Patterns(Pat).start
                    Else
                        Call TheHdw.Patterns(Pat).Test(pfAlways, 0)
                    End If
    
                    TestSeqNum = 0
    
                    For Each Ts In TestSequenceArray
    
                        If (CPUA_Flag_In_Pat) Then
                            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
                            If Ts = "F" Then
                                TheHdw.Wait 0.01
                            End If
                        Else
                            Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
                        End If
    
                        TestOptLen = Len(Ts)
                        TestOption = Ts
                        If TestOption = "F" And MeasPin_SingleEnd <> "" Then Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_F_SingleEnd, MeasureF_Pin_SingleEnd)
                        If TestOption = "F" And MeasPin_Differential <> "" Then Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_F_Differential, MeasureF_Pin_Differential)
 
                        Select Case UCase(Ts)
                            Case "F"
                                If ExecIndex = 0 Then
    
                                ElseIf ExecIndex > 0 Then
    
                                    If MeasPin_SingleEnd <> "" Then
                                        Call HardIP_Duty_Frequency(MeasureF_Pin_SingleEnd, False, TestSeqNum, 0.01, Rtn_MeasFreq, False, False)
                                        
                                    ElseIf MeasPin_Differential <> "" Then
                                        If CUS_Str_MainProgram <> "" Then
                                            SplitFreqVtValue = Split(CUS_Str_MainProgram, ":")
                                            If UCase(SplitFreqVtValue(0)) = "SETUP_STORE_VT" Then
                                                    TheHdw.Digital.Pins(MeasureF_Pin_Differential).DifferentialLevels.Value(chDiff_Vt) = Dict_VT_Value(site)
                                            End If
                                        End If
                                        Call HardIP_Duty_Frequency(MeasureF_Pin_Differential, True, TestSeqNum, 0.01, Rtn_MeasFreq, False, False)
                                        
                                    End If
     
                                    Dim d_Freq_duty_Cycle As Double
                                    d_Freq_duty_Cycle = Rtn_MeasFreq.Pins(1).Value(site)
    
                                    If Abs(d_Freq - d_Freq_duty_Cycle) < 200000 Or ExecIndex = 20 Then
                                        '============================================================================================================================
                                        '(2)For TestNumber align , calculate the EQs in Bin Cut Tables and get a guard band to avoid the TestNumber is different from
                                        'another touch down
                                        '============================================================================================================================
                                        
                                        Set_Align_Test_Number = Org_Test_Number + (ExecMaxNum * 3) + 5
                                        TheExec.sites(site).TestNumber = Set_Align_Test_Number
                                     
                                        ExecIndex = 20
                                        Dim DSPCapture1 As New PinListData
                                        DSPCapture = TheHdw.Digital.Jitter.SingleDSPWaves
                                        TheHdw.Digital.Jitter.TimeoutEnable = True          'enable time out
                                        TheHdw.Digital.Jitter.TimeOut = 10                      'setup time out value
    
                                        TheHdw.Wait 0.1
    
                                        'start time measure
                                        TheHdw.Digital.Jitter.start     'start jitter
                                        TheHdw.Digital.Jitter.Wait
    
                                        Call TheHdw.Digital.Jitter.UpdateSingleDSPWaves(DSPCapture)     'update the dspwave
    
                                        'measure jitter block
                                        If jitter_meas = True Then
                                            Call rundsp.DDR_LoopJitterMeas(DSPCapture, RJ_J, DDJ_J, MeasUI_J, dspStatus_J, DutyCycle_J, Freq_J, PWHigh_J, PWLow_J, MeasuredPeriod_J)      'DSP
                                        End If
    
''                                        'measure eye block
''                                        If eye_meas = True Then
''                                            Call rundsp.LoopEyeMeas(DSPCapture, RJ, DDJ, Tj, measUI, Tr, Tf, Eye20, Eye50, Eye80, dspStatus)        'DSP
''                                        End If
    
                                        If MeasPin_SingleEnd <> "" Then
                                            TestLimitPinName = MeasureF_Pin_SingleEnd
                                        ElseIf MeasPin_Differential <> "" Then
                                            TestLimitPinName = MeasureF_Pin_Differential
                                        End If
    
                                        If jitter_meas = True Then
                                            For index = 0 To (RJ_J.Pins.Count - 1)
''                                                Call TheExec.Flow.TestLimit(resultVal:=RJ_J.Pins(index), ScaleType:=scalePico, unit:=unitTime, Tname:=TestName & "Jitter_RJ", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
''                                                Call TheExec.Flow.TestLimit(resultVal:=DDJ_J.Pins(index), ScaleType:=scalePico, unit:=unitTime, Tname:=TestName & "Jitter_DDJ", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
''                                                Call TheExec.Flow.TestLimit(resultVal:=measUI_j.Pins(index), ScaleType:=scalePico, unit:=unitTime, Tname:=TestName & "Jitter_UI", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                                                Call TheExec.Flow.TestLimit(resultVal:=PWHigh_J.Pins(index), Unit:=unitTime, Tname:="PWH_" & SeqIndex & "_" & TestNameInput, PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                                                Call TheExec.Flow.TestLimit(resultVal:=MeasuredPeriod_J.Pins(index), Unit:=unitTime, Tname:="Period_" & SeqIndex & "_" & TestNameInput, PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                                                Call TheExec.Flow.TestLimit(resultVal:=DutyCycle_J.Pins(index), Unit:=unitCustom, Tname:="DutyCycle_" & SeqIndex & "_" & TestNameInput, PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                                                Call TheExec.Flow.TestLimit(resultVal:=Freq_J.Pins(index), scaletype:=scaleMega, Unit:=unitHz, Tname:="Freq_" & SeqIndex & "_" & TestNameInput, PinName:=TestLimitPinName, ForceResults:=tlForceNone)
                                            Next index
                                        End If
    
''                                        If eye_meas = True Then
''                                            Call TheExec.Flow.TestLimit(resultVal:=RJ, ScaleType:=scalePico, unit:=unitTime, Tname:="Eye_RJ", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
''                                            Call TheExec.Flow.TestLimit(resultVal:=DDJ, ScaleType:=scalePico, unit:=unitTime, Tname:="Eye_DDJ", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
''                                            Call TheExec.Flow.TestLimit(resultVal:=measUI, ScaleType:=scalePico, unit:=unitTime, Tname:="Eye_Width", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
''                                            Call TheExec.Flow.TestLimit(resultVal:=Eye20, ScaleType:=scalePico, unit:=unitTime, Tname:="Eye_Widthhigh", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
''                                            Call TheExec.Flow.TestLimit(resultVal:=Eye50, ScaleType:=scalePico, unit:=unitTime, Tname:="Eye_Widthmid", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
''                                            Call TheExec.Flow.TestLimit(resultVal:=Eye80, ScaleType:=scalePico, unit:=unitTime, Tname:="Eye_Widthlow", PinName:=TestLimitPinName, ForceResults:=tlForceNone)
''                                        End If
                                     End If
                                End If
    
                            Case "N"
    
                            Case Else
                        End Select
                        TestSeqNum = TestSeqNum + 1
    
                        If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                    Next Ts
    
                    TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end
    
                    If DigCap_Sample_Size <> 0 Then
                        Dim DigCapPinAry() As String, NumberPins As Long
                        Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
    
                        If NumberPins > 1 Then
                            Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
                            Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins)
                        ElseIf NumberPins = 1 Then
                            Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
                            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave)
                            Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
                        End If
                    End If
    
                Next patt
            Next ExecIndex
        Next site
        TestSequenceArray(FreqSeq_Start + SeqIndex) = "N"
    Next SeqIndex
    
    '' 20160713 - Call write functional result if cpu flag in pattern
    If (CPUA_Flag_In_Pat) Then
        Call HardIP_WriteFuncResult
    End If
  
    Exit Function
  
errHandler:
    TheExec.Datalog.WriteComment "error in DDR_RO_Time_Measure_KIT_UP1600"
    If AbortTest Then Exit Function Else Resume Next
End Function








