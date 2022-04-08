Attribute VB_Name = "VBT_LIB_HardIP_PCIE_PI"
Option Explicit

Public PhAdj As New SiteLong  'As Integer
Public lane As Integer
Public init_phase As Integer
Public Qd As New SiteDouble
Public Ph0Rel As New SiteDouble
'Public site As Variant
Public pat4_check As New SiteLong  'As Integer
Public Ph0 As New SiteDouble
Public Adj_steps As New SiteDouble
Public offset_angle As New SiteDouble


Public Function PCIE_PI_TEST(patset1 As Pattern, patset2 As Pattern, patset3 As Pattern, patset4 As Pattern, PhAdjMax As Integer, Total_Lane_Num As Long, Tested_Lane_Num As Long) As Long
    
    Dim pat1_count As Long
    Dim pat2_count As Long
    Dim pat3_count As Long
    Dim pat4_count As Long
    Dim PatTrimArray_pat1() As String
    Dim PatTrimArray_pat2() As String
    Dim PatTrimArray_pat3() As String
    Dim PatTrimArray_pat4() As String
    Dim site As Variant
    Dim Current_test_number As Long
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    Call GetPatFromPatternSet(patset1.Value, PatTrimArray_pat1, pat1_count)
    Call GetPatFromPatternSet(patset2.Value, PatTrimArray_pat2, pat2_count)
    Call GetPatFromPatternSet(patset3.Value, PatTrimArray_pat3, pat3_count)
    Call GetPatFromPatternSet(patset4.Value, PatTrimArray_pat4, pat4_count)
    
    offset_angle = 5.625
    
    For Each site In TheExec.sites
        
       Current_test_number = TheExec.sites(site).TestNumber
        
    Next site
    
    Call PCIE_PI_Pat1(PatTrimArray_pat1(0), Total_Lane_Num)
    
    For Each site In TheExec.sites
        
        pat4_check(site) = 0
        
    Next site
        
        For lane = 0 To Tested_Lane_Num - 1
                 
            Call PCIE_PI_Pat2(PatTrimArray_pat2(lane), PhAdjMax)
            Call PCIE_PI_Pat3(PatTrimArray_pat3(lane), "JTAG_TDI", PatTrimArray_pat2(lane), PhAdjMax)

        Next lane
        
     
    For Each site In TheExec.sites
     
     TheExec.sites(site).TestNumber = Current_test_number + 50000
     
    Next site
     
     Call PCIE_PI_Pat4(PatTrimArray_pat4(0), Total_Lane_Num, Tested_Lane_Num)
     
     DebugPrintFunc patset1 & "," & patset2 & "," & patset3 & "," & patset4 ' print all debug information
     
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in PCIE_PI_TEST"
    If AbortTest Then Exit Function Else Resume Next
  
End Function

Public Function PCIE_PI_Pat1(pat1 As String, Total_Lane_Num As Long)

    Dim pat1_dspwave As New DSPWave
    Dim abc As Integer
    Dim output_pin As New PinList
    Dim merge_bit As String
    output_pin = "JTAG_TDO"
    Dim site As Variant
    Dim total_bit As Long
    
    total_bit = Total_Lane_Num * 7
    
    Call DigCapSetup(pat1, output_pin, "test", total_bit, pat1_dspwave)

     
    TheHdw.Patterns(pat1).Load
    Call TheHdw.Patterns(pat1).Test(pfAlways, 0, tlResultModeDomain)
    
    Call TheHdw.Digital.Patgen.HaltWait
    
    For Each site In TheExec.sites
    
    pat1_dspwave = TheHdw.DSSC.Pins("JTAG_TDO").Pattern(pat1).Capture.Signals("test").DSPWave
    
        For abc = UBound(pat1_dspwave(site).Data) To 0 Step -1
        merge_bit = merge_bit & pat1_dspwave(site).Element(abc)
        Next abc
        
        TheExec.Datalog.WriteComment "site : " & site & " " & merge_bit & " <=LSB"
        merge_bit = ""
        
    Next site
   
End Function

Public Function PCIE_PI_Pat2(pat2 As String, PhAdjMax As Integer)

    Dim pat2_dspwave As New DSPWave
    Dim abc As Integer
    Dim output_pin As New PinList
    output_pin = "JTAG_TDO"
    Dim merge_bit As String
    Dim site As Variant
    Call DigCapSetup(pat2, output_pin, "test1", 8, pat2_dspwave)
    
    TheHdw.Patterns(pat2).Load
    Call TheHdw.Patterns(pat2).Test(pfAlways, 0, tlResultModeDomain)
    Call TheHdw.Digital.Patgen.HaltWait
    
    

    
        
    pat2_dspwave = TheHdw.DSSC.Pins("JTAG_TDO").Pattern(pat2).Capture.Signals("test1").DSPWave
        
    For Each site In TheExec.sites
    
        For abc = UBound(pat2_dspwave(site).Data) To 0 Step -1
            merge_bit = merge_bit & pat2_dspwave(site).Element(abc)
        Next abc
            
        'TheExec.Datalog.WriteComment "site : " & Site & " " & merge_bit & " <=LSB"
    
        If merge_bit Like "00000000" Or merge_bit Like "00100000" Then
            PhAdj(site) = PhAdjMax + 1
            pat4_check(site) = 1
        End If
        
        
        Qd(site) = (pat2_dspwave(site).Element(6) + pat2_dspwave(site).Element(7) * 2)
        Ph0Rel(site) = pat2_dspwave(site).Element(0) + pat2_dspwave(site).Element(1) * 2 + pat2_dspwave(site).Element(2) * 4 + pat2_dspwave(site).Element(3) * 8 + pat2_dspwave(site).Element(4) * 16


        Select Case Qd(site)
            Case 0:
               Ph0(site) = 0 + (Ph0Rel(site) * offset_angle)
               Adj_steps(site) = Ph0(site) / offset_angle
            Case 1:
               Ph0(site) = 360 - (Ph0Rel(site) * offset_angle)
               Adj_steps(site) = (360 - Ph0(site)) / offset_angle
            Case 2:
               Ph0(site) = 180 - (Ph0Rel(site) * offset_angle)
               Adj_steps(site) = Ph0(site) / offset_angle
            Case 3:
               Ph0(site) = 180 + (Ph0Rel(site) * offset_angle)
               Adj_steps(site) = (360 - Ph0(site)) / offset_angle
        End Select
      
        TheExec.Datalog.WriteComment "Lane" & lane & "  site " & site & ":" & merge_bit & "<=LSB" & "  Angle: " & Format(Ph0(site), "##0.000") & "  Steps: " & Adj_steps(site)

        merge_bit = ""
    Next site

End Function

Public Function PCIE_PI_Pat3(Adjust_pat As String, DigSrcPin As String, PCIE_Pat2 As String, PCIE_PhAdjMax As Integer) As Long

  Dim Data_Out As New DSPWave
  Dim pat_count As Long
  Dim Adj_patArray() As String
  Dim Adj_common_steps As Integer
  Dim i As Integer
  Dim i_loop As Integer
  Dim j_loop As Integer
  Dim DigSrcPin1 As New PinList
  Dim site As Variant
  Dim ChecPhaseResult As Boolean
  Dim InWave As New DSPWave
  Dim Current_test_number1 As Long
  InWave.CreateConstant 0, 3
  TheHdw.Patterns(Adjust_pat).Load
  
  For Each site In TheExec.sites
       Current_test_number1 = TheExec.sites(site).TestNumber
  Next site
  
  
  For i_loop = 0 To PCIE_PhAdjMax
  
      For Each site In TheExec.sites
    
'      Select Case Qd(Site)
'      Case 0:
'         Ph0(Site) = 0 + (Ph0Rel(Site) * offset_angle)
'         Adj_steps(Site) = Ph0(Site) / offset_angle
'      Case 1:
'         Ph0(Site) = 360 - (Ph0Rel(Site) * offset_angle)
'         Adj_steps(Site) = (360 - Ph0(Site)) / offset_angle
'      Case 2:
'         Ph0(Site) = 180 - (Ph0Rel(Site) * offset_angle)
'         Adj_steps(Site) = Ph0(Site) / offset_angle
'      Case 3:
'         Ph0(Site) = 180 + (Ph0Rel(Site) * offset_angle)
'         Adj_steps(Site) = (360 - Ph0(Site)) / offset_angle
'      End Select
         
      If Qd(site) = 0 Or Qd(site) = 2 Then
         InWave.Element(0) = 1
         InWave.Element(1) = 0
         InWave.Element(2) = 1
      Else
         InWave.Element(0) = 1
         InWave.Element(1) = 1
         InWave.Element(2) = 0
      End If
      DigSrcPin1 = DigSrcPin
    
    Next site
    
    Call SetupDigSrcDspWave(Adjust_pat, DigSrcPin1, "phase_adj", 3, InWave)
     
  
    Adj_common_steps = 999
    
    For Each site In TheExec.sites
        
        If Adj_common_steps > Adj_steps(site) Then
            Adj_common_steps = Adj_steps(site)
        End If
        
    Next site
  
    For j_loop = 0 To Adj_common_steps - 1
'        Call TheHdw.Patterns(Adjust_pat).start
        Call TheHdw.Patterns(Adjust_pat).Test(pfAlways, 0, tlResultModeDomain)
        TheHdw.Digital.Patgen.HaltWait
    Next j_loop
  
  
     For Each site In TheExec.sites
            For i = 0 To Adj_steps(site) - 1 - Adj_common_steps
'                   If i < Adj_steps(Site) - 1 - Adj_common_steps Then
'                    Call TheHdw.Patterns(Adjust_pat).start
'                   Else
                 Call TheHdw.Patterns(Adjust_pat).Test(pfAlways, 0, tlResultModeDomain)
'                   End If
                 TheHdw.Digital.Patgen.HaltWait
            Next i
     Next site


    For Each site In TheExec.sites
        
        TheExec.sites(site).TestNumber = Current_test_number1 + 2000 * (i_loop + 1)
        
    Next site
     
    Call PCIE_PI_Pat2(PCIE_Pat2, PCIE_PhAdjMax)
     
    ChecPhaseResult = True
    For Each site In TheExec.sites
    
        If pat4_check(site) = 0 Then
           ChecPhaseResult = False
           If i_loop = PCIE_PhAdjMax Then TheExec.sites.Item(site).testResult = siteFail
        End If
    
    Next site
    
    If ChecPhaseResult = True Then i_loop = PCIE_PhAdjMax + 1
                                 
  Next i_loop
  
  
End Function

Public Function PCIE_PI_Pat4(pat4 As String, Total_Lane_Num As Long, Tested_Lane_Num As Long)

    Dim pat4_dspwave As New DSPWave
    Dim i As Integer
    Dim output_pin As New PinList
    Dim site As Variant
    Dim show_out As String
    Dim lane_status_str() As String
    Dim lane_phase_result() As String
    Dim phase_result As String
    Dim phase_num As Integer
    Dim Lane_Num As Integer
    Dim total_bit As Long
    Dim X As Integer
    Dim phase_Array() As String
    Dim eye_width_site() As New SiteDouble
    Dim eye_count_site() As New SiteDouble
    Dim con_fail_count_site() As New SiteDouble
    Dim pat4_output_bit_num As Long
    
    ReDim lane_phase_result(0 To Total_Lane_Num - 1)
    ReDim eye_width_site(0 To Total_Lane_Num - 1)
    ReDim eye_count_site(0 To Total_Lane_Num - 1)
    ReDim con_fail_count_site(0 To Total_Lane_Num - 1)
    
    output_pin = "JTAG_TDO"
    total_bit = Total_Lane_Num * 16 * 17

    TheHdw.Patterns(pat4).Load
    Call DigCapSetup(pat4, output_pin, "test1", total_bit, pat4_dspwave)

    Call TheHdw.Patterns(pat4).Test(pfAlways, 0, tlResultModeDomain)
 
    Call TheHdw.Digital.Patgen.HaltWait
    
    pat4_dspwave = TheHdw.DSSC.Pins("JTAG_TDO").Pattern(pat4).Capture.Signals("test1").DSPWave
    

    For Each site In TheExec.sites
    
        show_out = ""
        
        For X = 0 To total_bit - 1
            show_out = show_out & pat4_dspwave.Element(X)
        Next X
     
        For X = 0 To Total_Lane_Num - 1
            lane_phase_result(X) = ""
        Next X

        For phase_num = 0 To 15
            For X = 0 To Total_Lane_Num - 1
                phase_result = Mid(show_out, (X + 1) + (phase_num * (17 * Total_Lane_Num)), 1) & Mid(show_out, (Total_Lane_Num + 1 + (X * 16)) + (phase_num * (17 * Total_Lane_Num)), 16)
                If phase_result = "10000000000000000" Then
                    lane_phase_result(X) = lane_phase_result(X) & "1,"
                Else
                    lane_phase_result(X) = lane_phase_result(X) & "0,"
                End If
            Next X
        Next phase_num
        
 
       For i = 0 To Tested_Lane_Num - 1
            
                phase_Array = Split(lane_phase_result(i), ",")
                eye_count_site(i) = EyeCount(phase_Array)
                
                phase_Array = Split(lane_phase_result(i), ",")
                eye_width_site(i) = EyeWidth(phase_Array)
                
                phase_Array = Split(lane_phase_result(i), ",")
                con_fail_count_site(i) = PI_continuous_fai_count(phase_Array)

                TheExec.Datalog.WriteComment "Site: " & site & ", Lane" & i & " phase result " & " = " & lane_phase_result(i) & ""
        Next i
        
        TheExec.Datalog.WriteComment "Site: " & site & ", Capture bits " & total_bit & " = " & show_out & " "
        
    Next site
    
   For i = 0 To Tested_Lane_Num - 1
                TheExec.Flow.TestLimit eye_count_site(i), 1, 2, Tname:=TheExec.DataManager.instanceName & "con_fail_count", PinName:="Lane_" & i & "_eye_count", ForceResults:=tlForceFlow
                TheExec.Flow.TestLimit eye_width_site(i), 6, 16, Tname:=TheExec.DataManager.instanceName & "con_fail_count", PinName:="Lane_" & i & "_eye_width", ForceResults:=tlForceFlow
                TheExec.Flow.TestLimit con_fail_count_site(i), 0, 2, Tname:=TheExec.DataManager.instanceName & "con_fail_count", PinName:="Lane_" & i & "_max_con_fail_count", ForceResults:=tlForceFlow
  Next i
End Function

Public Function PI_continuous_fai_count(phase_Array() As String) As Integer
'================continuous fail count============================================
        Dim temp_count As Integer
        Dim con_fail_count As Integer
        Dim loop_2nd As Integer
        Dim i As Integer
        temp_count = 0
        con_fail_count = 0
        loop_2nd = 0
        For i = 0 To 15
            If temp_count = 16 Then
                con_fail_count = 16
                i = 16
            End If
            If phase_Array(i) = "0" Then
                temp_count = temp_count + 1
                If i = 15 Then
                    i = 0
                    loop_2nd = 1
                End If
                    
            ElseIf phase_Array(i) = "1" Then
                If con_fail_count <= temp_count Then con_fail_count = temp_count
                temp_count = 0
                If loop_2nd = 1 Then i = i + 15
            Else
            End If
        Next i
        
    PI_continuous_fai_count = con_fail_count

End Function

'Public Check_Eye(17) As String
Public Function EyeCount(EyeStrArray() As String) As Integer
'Public Function EyeCount() As Integer
'Dim EyeStrArray(17) As String
Dim EyeArrayCnt As Integer
Dim loop_i As Integer
Dim Record_bit As String
Dim Current_bit As String
Dim transitionCnt As Integer
Dim CycleCnt As Integer
Dim First_bit_Test As Boolean
Dim check_result As Integer

'EyeStrArray(0) = "1"
'EyeStrArray(1) = "1"
'EyeStrArray(2) = "1"
'EyeStrArray(3) = "1"
'EyeStrArray(4) = "1"
'EyeStrArray(5) = "1"
'EyeStrArray(6) = "1"
'EyeStrArray(7) = "1"
'EyeStrArray(8) = "1"
'EyeStrArray(9) = "1"
'EyeStrArray(10) = "1"
'EyeStrArray(11) = "1"
'EyeStrArray(12) = "1"
'EyeStrArray(13) = "1"
'EyeStrArray(14) = "1"
'EyeStrArray(15) = "1"
'EyeStrArray(16) = ""

    First_bit_Test = True
    CycleCnt = 0

    EyeArrayCnt = UBound(EyeStrArray)
    
    transitionCnt = 0
    
    For loop_i = 0 To EyeArrayCnt
    
        Current_bit = EyeStrArray(loop_i) ' current bit

        If First_bit_Test = False Then
        
            If Record_bit <> Current_bit Then transitionCnt = transitionCnt + 1 'cal the transition
            
            If CycleCnt = 1 Then Exit For
            
        End If

        Record_bit = EyeStrArray(loop_i) ' recored bit
    
        If loop_i = (EyeArrayCnt - 1) And CycleCnt = 0 Then
        
            CycleCnt = 1 'set start another cycle
            loop_i = -1
        End If
    
    First_bit_Test = False
    
    Next

    check_result = transitionCnt Mod 2
    transitionCnt = transitionCnt - check_result
    If transitionCnt = 0 Then
        If Current_bit Like "1" Then
            EyeCount = 1
        Else
            EyeCount = 0
        End If
    Else
        EyeCount = transitionCnt / 2 'Mod 2
    End If
    
End Function

Public Function EyeWidth(EyeWidthStrArray() As String) As Integer
'
'End Function
'Public Function EyeWidth() As Integer
'Dim EyeWidthStrArray(17) As String
Dim EyeArrayCnt As Integer
Dim loop_i As Integer
Dim Record_bit As String
Dim Current_bit As String
Dim transitionCnt As Integer
Dim CycleCnt As Integer
Dim First_bit_Test As Boolean
Dim check_result As Integer
Dim MaxEyeWidth As Integer
Dim CurrentEyeWidth As Integer
Dim Start_record As Boolean
'EyeWidthStrArray(0) = "0"
'EyeWidthStrArray(1) = "0"
'EyeWidthStrArray(2) = "0"
'EyeWidthStrArray(3) = "0"
'EyeWidthStrArray(4) = "0"
'EyeWidthStrArray(5) = "0"
'EyeWidthStrArray(6) = "0"
'EyeWidthStrArray(7) = "0"
'EyeWidthStrArray(8) = "0"
'EyeWidthStrArray(9) = "0"
'EyeWidthStrArray(10) = "0"
'EyeWidthStrArray(11) = "0"
'EyeWidthStrArray(12) = "0"
'EyeWidthStrArray(13) = "0"
'EyeWidthStrArray(14) = "0"
'EyeWidthStrArray(15) = "0"
'EyeWidthStrArray(16) = ""

    MaxEyeWidth = 16
    CurrentEyeWidth = 16
    Start_record = False
    First_bit_Test = True
    CycleCnt = 0

    EyeArrayCnt = UBound(EyeWidthStrArray) '- 1
    transitionCnt = 0
    
    For loop_i = 0 To EyeArrayCnt - 1
    
        Current_bit = EyeWidthStrArray(loop_i) ' current bit

        If First_bit_Test = False Then
        
            If Record_bit <> Current_bit Then
            
                If Current_bit Like "1" Then
                    Start_record = True
                End If
                
                If Current_bit Like "0" Then
                    If CurrentEyeWidth < MaxEyeWidth Then MaxEyeWidth = CurrentEyeWidth
                    CurrentEyeWidth = 0
                End If
                
                
                transitionCnt = transitionCnt + 1 'cal the transition
                
            End If
            
            If Current_bit Like "1" And Start_record = True Then
                CurrentEyeWidth = CurrentEyeWidth + 1
            End If
            
                
            
        End If

        Record_bit = EyeWidthStrArray(loop_i) ' recored bit
    
        If loop_i = (EyeArrayCnt - 1) And CycleCnt = 0 Then
        
            CycleCnt = 1 'set start another cycle
            loop_i = -1
        End If
    
    First_bit_Test = False
    
    Next

'    check_result = transitionCnt Mod 2
'    transitionCnt = transitionCnt - check_result
'    If transitionCnt = 0 Then
'        EyeWidth = 1
'    Else
'        EyeWidth = transitionCnt / 2 'Mod 2
'    End If
    
    If transitionCnt = 0 And Record_bit Like "0" Then
    
        EyeWidth = 0
    
    Else
    
        EyeWidth = MaxEyeWidth
        
    End If
    
    
    
End Function
