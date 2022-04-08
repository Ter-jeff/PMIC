Attribute VB_Name = "LIB_Digital_Debug"
Option Explicit

Public Const LVCC_boundary_Switch = 4 '1~10 means only get fail log at LVCC boundary with how many times
                                      '0 means get fail log with full search range

Public Const Shmoo_faillog_test_number = 100


Enum Shmoo_direction_enum
    High_to_Low = 1
    Low_to_High = 2
End Enum

Public Function FailingBoundaryDatalog_Func_Multi_Power(Power_Search_String As String, _
                    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                     Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant) As Long
'Shmoo with faillog capture version 1.0 careated by JT 2014/02/20.

Dim PinCnt As Long
Dim PinAry() As String
Dim i As Integer
Dim k As Integer
Dim j As Integer
Dim Fail_log_cnt As Integer
Dim patternArray() As String
Dim PowerV As Double
Dim p As Integer
Dim Org_Test_Number As Long
Dim current_site As Integer
Dim Timelist As String
Dim TimeGroup() As String
Dim CurrTiming As Variant
Dim TimeDomainlist As String
Dim TimeDomaingroup() As String
Dim CurrTimeDomain As Variant
Dim TimeDomainIn As String
On Error GoTo errHandler_faillog
'    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_pattern As String, _
'                                     Shmoo_status As String)

'                    Shmoo_faillog_header LotID, CStr(WaferId), CStr(XCoord(site)), CStr(YCoord(site)), Shmoo_pattern, "Shmoo hole"
'                    Call FailingBoundaryDatalog_Func_TER(shmoo_pin_string, RangeFrom, RangeTo, CDbl(RangeStepSize))
'                    Call Shmoo_faillog_ending


'Shmoo_faillog_test_number
'TheExec.Sites(site).TestNumber = 100
'current_site = TheExec.Sites.SiteNumber

'Org_test_number = TheExec.Sites(CurrSite).TestNumber

'TheExec.Sites(current_site).TestNumber = Shmoo_faillog_test_number

TheExec.Datalog.WriteComment "***************** Shmoo fail log capture start *****************"
TheExec.Datalog.WriteComment ""
TheExec.Datalog.WriteComment "Lot : " & Shmoo_LotID
TheExec.Datalog.WriteComment "Wafer :" & Shmoo_wafer
TheExec.Datalog.WriteComment "Die X :" & Shmoo_X
TheExec.Datalog.WriteComment "Die Y :" & Shmoo_Y
TheExec.Datalog.WriteComment "Pattern :" & Shmoo_Pattern
TheExec.Datalog.WriteComment "Shmoo status :" & Shmoo_status
TheExec.Datalog.WriteComment ""

'list time ing and frerunning clock
                'DomainList = thehdw.Patterns(InitPats(TestPatIdx)).TimeDomains
                TimeDomainlist = TheHdw.Digital.Timing.TimeDomainlist
                'TheExec.Datalog.WriteComment "Time Doamin : " & TimeDomainlist
                
                TimeDomaingroup = Split(TimeDomainlist, ",")
                
                For Each CurrTimeDomain In TimeDomaingroup
                    
                    'TheExec.Datalog.WriteComment "Time Doamin : " & CurrTimeDomain
                    If CStr(CurrTimeDomain) = "All" Then
                    TimeDomainIn = ""
                    Else
                    TimeDomainIn = CStr(CurrTimeDomain)
                    End If
                    
                    Timelist = TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
                    'TimeGroup
                    TimeGroup = Split(Timelist, ",")
                    For Each CurrTiming In TimeGroup
                        If CurrTiming = "" Then Exit For
                        TheExec.Datalog.WriteComment "Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & (1 / TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.Period(CStr(CurrTiming))) / 1000000 & " Mhz"
    
                    Next CurrTiming
                Next CurrTimeDomain

                '' add for XI0 free running clk
               'TheExec.Datalog.WriteComment "  FreeRunFreq : " & thehdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz , clock_Vih: " & thehdw.DIB.SupportBoardClock.Vih & " v , clock_Vil: " & thehdw.DIB.SupportBoardClock.vil & " v"
               'TheExec.Datalog.WriteComment "FreeRunFreq : " & TheHdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz" ', clock_Vih: " & clock_Vih_debug & " v , clock_Vil: " & clock_Vil_debug & " v"
               TheExec.Datalog.WriteComment ""


        Dim power_list() As String
        Dim Power_number As Integer
    
        Dim power_pins(20) As String
        Dim Power_RangeA(20) As Double
        Dim Power_RangeB(20) As Double
        Dim Power_StepSize(20) As Double
        Dim power_Temp() As String
        Dim Power_range_temp() As String
        Dim Shmoo_steps As Double
        Dim axis_type As tlDevCharShmooAxis
        Dim SetupName As String
        Dim VmainOrValt As String
        
    
        power_list = Split(Power_Search_String, ",")
        Power_number = UBound(power_list)
        
        For i = 0 To Power_number
            power_Temp() = Split(power_list(i), "=")
            Power_range_temp() = Split(power_Temp(1), ":")
            power_pins(i) = power_Temp(0)
            Power_RangeA(i) = CDbl(Power_range_temp(0))
            Power_RangeB(i) = CDbl(Power_range_temp(1))
            Power_StepSize(i) = CDbl(Power_range_temp(2))
            Shmoo_steps = Abs(Power_RangeA(i) - Power_RangeB(i)) / Abs(Power_StepSize(i))
            'Power_setting()
        Next i

k = Shmoo_steps
Fail_log_cnt = 1

        SetupName = TheExec.DevChar.Setups.ActiveSetupName
        VmainOrValt = LCase(TheExec.DevChar.Setups.Item(SetupName).Shmoo.Axes.Item(axis_type).Parameter.Name)

        For j = 0 To k
        
            'loop power by step
            If Direction = Low_to_High Then
            
                For i = 0 To Power_number
                    If Power_RangeA(i) < Power_RangeB(i) Then
                        If VmainOrValt = "vmain" Then
                            TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value = Power_RangeA(i) + Abs(Power_StepSize(i)) * j
                        ElseIf VmainOrValt = "valt" Then
                            TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value = Power_RangeA(i) + Abs(Power_StepSize(i)) * j
                        End If
                    Else
                        If VmainOrValt = "vmain" Then
                            TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value = Power_RangeB(i) + Abs(Power_StepSize(i)) * j
                        ElseIf VmainOrValt = "valt" Then
                            TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value = Power_RangeB(i) + Abs(Power_StepSize(i)) * j
                        End If
                    End If
                Next i
                
            Else
            
                For i = 0 To Power_number
                    If Power_RangeA(i) > Power_RangeB(i) Then
                        If VmainOrValt = "vmain" Then
                            TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value = Power_RangeA(i) - Abs(Power_StepSize(i)) * j
                        ElseIf VmainOrValt = "valt" Then
                            TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value = Power_RangeA(i) - Abs(Power_StepSize(i)) * j
                        End If
                    Else
                        If VmainOrValt = "vmain" Then
                            TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value = Power_RangeB(i) - Abs(Power_StepSize(i)) * j
                        ElseIf VmainOrValt = "valt" Then
                            TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value = Power_RangeB(i) - Abs(Power_StepSize(i)) * j
                        End If
                    End If
                Next i
            End If
            
            For i = 0 To Power_number
                If VmainOrValt = "vmain" Then
                    TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.000"))
                    TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.000")) + ("V")
                ElseIf VmainOrValt = "valt" Then
                    TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value, "0.000"))
                    TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value, "0.000")) + ("V")
                End If
            Next i
            
            TheHdw.Wait 0.002
            TheHdw.Patterns(Shmoo_Pattern).Test pfAlways, 0
            
            
            For i = 0 To Power_number
                If VmainOrValt = "vmain" Then
                    TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.000"))
                    TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.000")) + ("V")
                ElseIf VmainOrValt = "valt" Then
                    TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value, "0.000"))
                    TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value, "0.000")) + ("V")
                End If
            Next i
                    
                    
            If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False And LVCC_boundary_Switch = Fail_log_cnt Then Exit For
            If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then Fail_log_cnt = Fail_log_cnt + 1

        Next j


TheExec.Datalog.WriteComment "***************** Shmoo fail log capture end *****************"

'TheExec.Sites(current_site).TestNumber = Org_test_number

errHandler_faillog:
TheExec.Datalog.WriteComment "*****************Shmoo with fail log capture error *****************"

End Function

Public Function FailingDatalog_Lvcc_Boundary(Power_Search_String As String, _
                    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                     Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeSteps As Long, Optional RangeStepSize As Double) As Long
'Shmoo with faillog capture version 1.0 careated by JT 2014/02/20.

Dim PinCnt As Long
Dim PinAry() As String
Dim i As Integer
Dim k As Integer
Dim j As Integer
Dim Fail_log_cnt As Integer
Dim patternArray() As String
Dim PowerV As Double
Dim p As Integer
Dim Org_Test_Number As Long
Dim current_site As Integer
Dim Timelist As String
Dim TimeGroup() As String
Dim CurrTiming As Variant
Dim TimeDomainlist As String
Dim TimeDomaingroup() As String
Dim CurrTimeDomain As Variant
Dim TimeDomainIn As String
Dim ShmooPatternSplit() As String
Dim TestNumber As Long
Dim inst_name As String
Dim shmoopowerpin As String
Dim Failed_Pins() As String
Dim AllFailPins As String
Dim OutputString As String

On Error GoTo errHandler_faillog
'    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_pattern As String, _
'                                     Shmoo_status As String)

'                    Shmoo_faillog_header LotID, CStr(WaferId), CStr(XCoord(site)), CStr(YCoord(site)), Shmoo_pattern, "Shmoo hole"
'                    Call FailingBoundaryDatalog_Func_TER(shmoo_pin_string, RangeFrom, RangeTo, CDbl(RangeStepSize))
'                    Call Shmoo_faillog_ending


'Shmoo_faillog_test_number
'TheExec.Sites(site).TestNumber = 100
'current_site = TheExec.Sites.SiteNumber

'Org_test_number = TheExec.Sites(CurrSite).TestNumber

'TheExec.Sites(current_site).TestNumber = Shmoo_faillog_test_number

ShmooPatternSplit() = Split(Shmoo_Pattern, ",")
inst_name = UCase(TheExec.DataManager.instanceName)

Dim Context As String: Context = ""
Dim TimeSet_Str As String: TimeSet_Str = ""

Context = TheExec.Contexts.ActiveSelection
TimeSet_Str = TheExec.Contexts(Context).Sheets.Timesets

If TheExec.Flow.EnableWord("FailPinsOnly") = False Then
TheExec.Datalog.WriteComment "***************** Shmoo fail log capture start *****************"
TheExec.Datalog.WriteComment ""
TheExec.Datalog.WriteComment "Lot : " & Shmoo_LotID
TheExec.Datalog.WriteComment "Wafer :" & Shmoo_wafer
TheExec.Datalog.WriteComment "Die X :" & Shmoo_X
TheExec.Datalog.WriteComment "Die Y :" & Shmoo_Y
TheExec.Datalog.WriteComment "Pattern :" & Shmoo_Pattern
TheExec.Datalog.WriteComment "Shmoo status :" & Shmoo_status
TheExec.Datalog.WriteComment ""
TheExec.Datalog.WriteComment "Activity Timeset Sheet :" & TimeSet_Str

'list time ing and frerunning clock
                'DomainList = thehdw.Patterns(InitPats(TestPatIdx)).TimeDomains
                TimeDomainlist = TheHdw.Digital.Timing.TimeDomainlist
                'TheExec.Datalog.WriteComment "Time Domain : " & TimeDomainlist
                
                TimeDomaingroup = Split(TimeDomainlist, ",")
                
                For Each CurrTimeDomain In TimeDomaingroup
                    
                    'TheExec.Datalog.WriteComment "Time Domain : " & CurrTimeDomain
                    If CStr(CurrTimeDomain) = "All" Then
                    TimeDomainIn = ""
                    Else
                    TimeDomainIn = CStr(CurrTimeDomain)
                    End If
                    
                    Timelist = TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
                    'TimeGroup
                    TimeGroup = Split(Timelist, ",")
                    For Each CurrTiming In TimeGroup
                        If CurrTiming = "" Then Exit For
                        TheExec.Datalog.WriteComment "Time Domain : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & (1 / TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.Period(CStr(CurrTiming))) / 1000000 & " Mhz"
    
                    Next CurrTiming
                Next CurrTimeDomain

                '' add for XI0 free running clk
               'TheExec.Datalog.WriteComment "  FreeRunFreq : " & thehdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz , clock_Vih: " & thehdw.DIB.SupportBoardClock.Vih & " v , clock_Vil: " & thehdw.DIB.SupportBoardClock.vil & " v"
               'TheExec.Datalog.WriteComment "FreeRunFreq : " & TheHdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz" ', clock_Vih: " & clock_Vih_debug & " v , clock_Vil: " & clock_Vil_debug & " v"
               TheExec.Datalog.WriteComment ""
End If

        Dim power_list() As String
        Dim Power_number As Integer
    
        Dim power_pins() As String
        Dim Power_RangeA(20) As Double
        Dim Power_RangeB(20) As Double
        Dim Power_StepSize(20) As Double
        Dim power_Temp() As String
        Dim Power_range_temp() As String
        Dim Shmoo_steps As Double
        Dim StepValue As Double
    
        power_list = Split(Power_Search_String, ",")
        Power_number = UBound(power_list)
        ReDim power_pins(Power_number)
        
        For i = 0 To Power_number
             power_Temp() = Split(power_list(i), "=")
             Power_range_temp() = Split(power_Temp(1), ":")
             power_pins(i) = power_Temp(0)
        Next i
        
        
        If RangeFrom > RangeTo Then
           Shmoo_steps = ((Shmoo_Vcc_Min(CurrSite) - RangeTo) / RangeStepSize)
        ElseIf RangeTo > RangeFrom Then
           Shmoo_steps = (Shmoo_Vcc_Min(CurrSite) - RangeFrom) / RangeStepSize
        Else
           Shmoo_steps = 3
        End If
       
If TheExec.Flow.EnableWord("FailPinsOnly") = False Then
   k = Shmoo_steps
Else
   k = 0
End If
   Fail_log_cnt = 1

    For j = 0 To k
        
            'loop power by step
            
                For i = 0 To Power_number
                  If TheExec.Flow.EnableWord("FailPinsOnly") = True Then
                      TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value = Shmoo_Vcc_Min(CurrSite) - 0.005
                  Else
                      TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value = Shmoo_Vcc_Min(CurrSite) - RangeStepSize * (j + 1)
                  End If
                Next i
                
          If TheExec.Flow.EnableWord("FailPinsOnly") = False Then

          StepValue = Fail_log_cnt * 3.125
          
          TheExec.Datalog.WriteComment "Power setup (Vmin- " & StepValue & "mV) "
            For i = 0 To Power_number
                TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.00000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.00000")) + ("V")
            Next i
          End If
            
            TheHdw.Wait 0.002
            
         If TheExec.Flow.EnableWord("FailPinsOnly") = True Then
            For i = 0 To UBound(ShmooPatternSplit)
             Call TheHdw.Patterns(ShmooPatternSplit(i)).Load
             Call TheHdw.Patterns(ShmooPatternSplit(i)).start
             TheHdw.Digital.Patgen.HaltWait
             If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then
              shmoopowerpin = Join(power_pins, ",")
              Failed_Pins() = TheHdw.Digital.FailedPins(CurrSite)
              AllFailPins = Join(Failed_Pins, ",")
              OutputString = "[" & "FailPins" & "," & Shmoo_LotID & "-" & Shmoo_wafer & "," & Shmoo_X & "," & Shmoo_Y & "," & "Site" & CStr(CurrSite)
              OutputString = OutputString & "," & inst_name & "," & "Pattern./" & ShmooPatternSplit(i) & "," & "ShmooPowerPin:" & shmoopowerpin & "," & "ApplyVoltage(Vmin-Guardband 5mV)" & "=" & CStr(Format((Shmoo_Vcc_Min(CurrSite) - 0.005), "0.00000"))
              OutputString = OutputString & "," & "FailPins = " & UCase(AllFailPins)
              TheExec.Datalog.WriteComment OutputString & "]"
             End If
            Next i
         Else
         
            Dim Temp_patary() As String
            Dim tempPat As Variant
            Temp_patary() = Split(Shmoo_Pattern, ",")
            For Each tempPat In Temp_patary
                TestNumber = TheExec.sites.Item(CurrSite).TestNumber
                TheHdw.Patterns(tempPat).Test pfNever, 0
                If (TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = True) Then
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestPass)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestFail)
                End If
                TestNumber = TestNumber + 1
                TheExec.sites.Item(CurrSite).TestNumber = TestNumber
            Next tempPat
            
            TheExec.Datalog.WriteComment "                                                "
            
''             For i = 0 To Power_number
''                TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.00000"))
''                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.00000")) + ("V")
''             Next i
  
            If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False And LVCC_boundary_Switch = Fail_log_cnt Then GoTo Endfor
            If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then Fail_log_cnt = Fail_log_cnt + 1
         End If
         
                    
    Next j
Endfor:

If TheExec.Flow.EnableWord("FailPinsOnly") = False Then
TheExec.Datalog.WriteComment "***************** Shmoo fail log/Pins capture end *****************"
End If

'TheExec.Sites(current_site).TestNumber = Org_test_number

errHandler_faillog:
'TheExec.Datalog.WriteComment "*****************Shmoo with fail log capture error *****************"

End Function

Public Function FailingDatalog_Hvcc_Boundary(Power_Search_String As String, _
                    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                     Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeSteps As Long, Optional RangeStepSize As Double) As Long
'Shmoo with faillog capture version 1.0 careated by JT 2014/02/20.

Dim PinCnt As Long
Dim PinAry() As String
Dim i As Integer
Dim k As Integer
Dim j As Integer
Dim Fail_log_cnt As Integer
Dim patternArray() As String
Dim PowerV As Double
Dim p As Integer
Dim Org_Test_Number As Long
Dim current_site As Integer
Dim Timelist As String
Dim TimeGroup() As String
Dim CurrTiming As Variant
Dim TimeDomainlist As String
Dim TimeDomaingroup() As String
Dim CurrTimeDomain As Variant
Dim TimeDomainIn As String
Dim ShmooPatternSplit() As String
Dim TestNumber As Long
Dim inst_name As String
Dim shmoopowerpin As String
Dim Failed_Pins() As String
Dim AllFailPins As String
Dim OutputString As String

On Error GoTo errHandler_faillog
'    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_pattern As String, _
'                                     Shmoo_status As String)

'                    Shmoo_faillog_header LotID, CStr(WaferId), CStr(XCoord(site)), CStr(YCoord(site)), Shmoo_pattern, "Shmoo hole"
'                    Call FailingBoundaryDatalog_Func_TER(shmoo_pin_string, RangeFrom, RangeTo, CDbl(RangeStepSize))
'                    Call Shmoo_faillog_ending


'Shmoo_faillog_test_number
'TheExec.Sites(site).TestNumber = 100
'current_site = TheExec.Sites.SiteNumber

'Org_test_number = TheExec.Sites(CurrSite).TestNumber

'TheExec.Sites(current_site).TestNumber = Shmoo_faillog_test_number

ShmooPatternSplit() = Split(Shmoo_Pattern, ",")
inst_name = UCase(TheExec.DataManager.instanceName)

Dim Context As String: Context = ""
Dim TimeSet_Str As String: TimeSet_Str = ""
Context = TheExec.Contexts.ActiveSelection
TimeSet_Str = TheExec.Contexts(Context).Sheets.Timesets

If TheExec.Flow.EnableWord("FailPinsOnly") = False Then
TheExec.Datalog.WriteComment "***************** Shmoo fail log capture start *****************"
TheExec.Datalog.WriteComment ""
TheExec.Datalog.WriteComment "Lot : " & Shmoo_LotID
TheExec.Datalog.WriteComment "Wafer :" & Shmoo_wafer
TheExec.Datalog.WriteComment "Die X :" & Shmoo_X
TheExec.Datalog.WriteComment "Die Y :" & Shmoo_Y
TheExec.Datalog.WriteComment "Pattern :" & Shmoo_Pattern
TheExec.Datalog.WriteComment "Shmoo status :" & Shmoo_status
TheExec.Datalog.WriteComment ""
TheExec.Datalog.WriteComment "Activity Timeset Sheet :" & TimeSet_Str
'list time ing and frerunning clock
                'DomainList = thehdw.Patterns(InitPats(TestPatIdx)).TimeDomains
                TimeDomainlist = TheHdw.Digital.Timing.TimeDomainlist
                'TheExec.Datalog.WriteComment "Time Domain : " & TimeDomainlist
                
                TimeDomaingroup = Split(TimeDomainlist, ",")
                
                For Each CurrTimeDomain In TimeDomaingroup
                    
                    'TheExec.Datalog.WriteComment "Time Domain : " & CurrTimeDomain
                    If CStr(CurrTimeDomain) = "All" Then
                    TimeDomainIn = ""
                    Else
                    TimeDomainIn = CStr(CurrTimeDomain)
                    End If
                    
                    Timelist = TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
                    'TimeGroup
                    TimeGroup = Split(Timelist, ",")
                    For Each CurrTiming In TimeGroup
                        If CurrTiming = "" Then Exit For
                        TheExec.Datalog.WriteComment "Time Domain : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & (1 / TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.Period(CStr(CurrTiming))) / 1000000 & " Mhz"
    
                    Next CurrTiming
                Next CurrTimeDomain

                '' add for XI0 free running clk
               'TheExec.Datalog.WriteComment "  FreeRunFreq : " & thehdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz , clock_Vih: " & thehdw.DIB.SupportBoardClock.Vih & " v , clock_Vil: " & thehdw.DIB.SupportBoardClock.vil & " v"
               'TheExec.Datalog.WriteComment "FreeRunFreq : " & TheHdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz" ', clock_Vih: " & clock_Vih_debug & " v , clock_Vil: " & clock_Vil_debug & " v"
               TheExec.Datalog.WriteComment ""
End If

        Dim power_list() As String
        Dim Power_number As Integer
    
        Dim power_pins() As String
        Dim Power_RangeA(20) As Double
        Dim Power_RangeB(20) As Double
        Dim Power_StepSize(20) As Double
        Dim power_Temp() As String
        Dim Power_range_temp() As String
        Dim Shmoo_steps As Double
        Dim StepValue As Double
    
        power_list = Split(Power_Search_String, ",")
        Power_number = UBound(power_list)
        ReDim power_pins(Power_number)
        
        For i = 0 To Power_number
             power_Temp() = Split(power_list(i), "=")
             Power_range_temp() = Split(power_Temp(1), ":")
             power_pins(i) = power_Temp(0)
        Next i
        
        
        If RangeFrom > RangeTo Then
           Shmoo_steps = ((RangeFrom - Shmoo_Vcc_Max(CurrSite)) / RangeStepSize)
        ElseIf RangeTo > RangeFrom Then
           Shmoo_steps = (RangeTo - Shmoo_Vcc_Max(CurrSite)) / RangeStepSize
        Else
           Shmoo_steps = 3
        End If
       
If TheExec.Flow.EnableWord("FailPinsOnly") = False Then
   k = Shmoo_steps
Else
   k = 0
End If
   Fail_log_cnt = 1

    For j = 0 To k
        
            'loop power by step
            
                For i = 0 To Power_number
                  If TheExec.Flow.EnableWord("FailPinsOnly") = True Then
                      TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value = Shmoo_Vcc_Max(CurrSite) + 0.005
                  Else
                      TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value = Shmoo_Vcc_Max(CurrSite) + RangeStepSize * (j + 1)
                  End If
                Next i
                
          If TheExec.Flow.EnableWord("FailPinsOnly") = False Then

          StepValue = Fail_log_cnt * 3.125
          
          TheExec.Datalog.WriteComment "Power setup (Vmax+ " & StepValue & "mV) "
            For i = 0 To Power_number
                TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.00000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.00000")) + ("V")
            Next i
          End If
            
            TheHdw.Wait 0.002
            
         If TheExec.Flow.EnableWord("FailPinsOnly") = True Then
            For i = 0 To UBound(ShmooPatternSplit)
             Call TheHdw.Patterns(ShmooPatternSplit(i)).Load
             Call TheHdw.Patterns(ShmooPatternSplit(i)).start
             TheHdw.Digital.Patgen.HaltWait
             If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then
              shmoopowerpin = Join(power_pins, ",")
              Failed_Pins() = TheHdw.Digital.FailedPins(CurrSite)
              AllFailPins = Join(Failed_Pins, ",")
              OutputString = "[" & "FailPins" & "," & Shmoo_LotID & "-" & Shmoo_wafer & "," & Shmoo_X & "," & Shmoo_Y & "," & "Site" & CStr(CurrSite)
              OutputString = OutputString & "," & inst_name & "," & "Pattern./" & ShmooPatternSplit(i) & "," & "ShmooPowerPin:" & shmoopowerpin & "," & "ApplyVoltage(Vmax+Guardband 5mV)" & "=" & CStr(Format((Shmoo_Vcc_Max(CurrSite) + 0.005), "0.00000"))
              OutputString = OutputString & "," & "FailPins = " & UCase(AllFailPins)
              TheExec.Datalog.WriteComment OutputString & "]"
             End If
            Next i
         Else
            Dim Temp_patary() As String
            Dim tempPat As Variant
            Temp_patary() = Split(Shmoo_Pattern, ",")
            For Each tempPat In Temp_patary
                TestNumber = TheExec.sites.Item(CurrSite).TestNumber
                TheHdw.Patterns(tempPat).Test pfNever, 0
                If (TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = True) Then
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestPass)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestFail)
                End If
                TestNumber = TestNumber + 1
                TheExec.sites.Item(CurrSite).TestNumber = TestNumber
            Next tempPat
            
            TheExec.Datalog.WriteComment "                                                "
            
''             For i = 0 To Power_number
''                TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.00000"))
''                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Main.Value, "0.00000")) + ("V")
''             Next i
  
            If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False And LVCC_boundary_Switch = Fail_log_cnt Then GoTo Endfor
            If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then Fail_log_cnt = Fail_log_cnt + 1
         End If
         
                    
    Next j
Endfor:

If TheExec.Flow.EnableWord("FailPinsOnly") = False Then
TheExec.Datalog.WriteComment "***************** Shmoo fail log/Pins capture end *****************"
End If

'TheExec.Sites(current_site).TestNumber = Org_test_number

errHandler_faillog:
'TheExec.Datalog.WriteComment "*****************Shmoo with fail log capture error *****************"

End Function
Public Function FailingDatalog_HLvcc_Boundary_SELSRM(Power_Search_String As String, _
                    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                     Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeSteps As Long, Optional RangeStepSize As Double, Optional dssc_pat As String) As Long
'Shmoo SELSRM HLVCC boundary failure log by Cebu 201807

Dim PinCnt As Long
Dim PinAry() As String
Dim i As Integer
Dim k As Integer
Dim j As Integer
Dim M As Integer
Dim Fail_log_cnt As Integer
Dim patternArray() As String
Dim PowerV As Double
Dim p As Integer
Dim Org_Test_Number As Long
Dim current_site As Integer
Dim Timelist As String
Dim TimeGroup() As String
Dim CurrTiming As Variant
Dim TimeDomainlist As String
Dim TimeDomaingroup() As String
Dim CurrTimeDomain As Variant
Dim TimeDomainIn As String
Dim ShmooPatternSplit() As String
Dim TestNumber As Long
Dim inst_name As String
Dim shmoopowerpin As String
Dim Failed_Pins() As String
Dim AllFailPins As String
Dim OutputString As String

Dim nWire_port_ary() As String
nWire_port_ary = Split(nWire_Ports_GLB, ",")
Dim pat_count As Long, pat_array_count As Long
Dim pat_array() As String, pat_array_nu() As String
Dim power_list_dic() As String
Dim power_count_dic As Long

Dim power_list() As String
Dim Power_number As Integer
Dim power_pins() As String
Dim Power_RangeA(20) As Double
Dim Power_RangeB(20) As Double
Dim Power_StepSize(20) As Double
Dim power_Temp() As String
Dim Power_range_temp() As String
Dim Shmoo_steps As Double
Dim StepValue As Double

Dim pat_list_array() As String
Dim pat_list_count As Long
Dim confirm_pat_load As Boolean
''''''''''''''''''''''''''''''''''''''
Dim DC_Spec_Level As New PinListData   '''from Vmain or Valt
Dim SELSRM_Rails As String
Dim Shmoo_value As New PinListData
Dim DigSrc_wav As New DSPWave
Dim powerPin As String
Dim pin_name() As String
Dim n As Integer
Dim DSSC_string As String
Dim pat_str_ary() As String
''''''''''''''''''''''''''''''''''''''
'Dim Pattern_Decompose() As String
Dim PatCnt As Long
Dim LVCC As Boolean:: LVCC = True
Dim dssc_pat_exist As Boolean:: dssc_pat = False
If Shmoo_status <> "" Then
   If UCase(Shmoo_status) Like "SHMOO HVCC" Then LVCC = False
End If
   
'On Error GoTo errHandler_faillog
    Set DigSrc_wav = Nothing
    If InStr(Shmoo_Pattern, ",") > 0 Then
       ShmooPatternSplit() = Split(Shmoo_Pattern, ",")
    Else
         ShmooPatternSplit() = TheExec.DataManager.Raw.GetPatternsInSet(Shmoo_Pattern, PatCnt)
    End If
    inst_name = UCase(TheExec.DataManager.instanceName)
    '----------------------------------------------------------------------
    'Decide_DC_Level DC_Spec_Level, g_ApplyLevelTimingValt, g_ApplyLevelTimingVmain, BlockType_GB
    
    If digSrc_EQ_GB <> "" Then
       DigSrc_wav.CreateConstant 0, Len(digSrc_EQ_GB), DspLong
       Decide_DC_Level DC_Spec_Level, g_ApplyLevelTimingValt, g_ApplyLevelTimingVmain, BlockType_GB
    End If
    
    pin_name() = Split(g_VDDForce, ",")

    '----------------------------------------------------------------------
         TheExec.Datalog.WriteComment "***************** Shmoo fail log capture start *****************"
         TheExec.Datalog.WriteComment ""
         TheExec.Datalog.WriteComment "Lot : " & Shmoo_LotID
         TheExec.Datalog.WriteComment "Wafer :" & Shmoo_wafer
         TheExec.Datalog.WriteComment "Die X :" & Shmoo_X
         TheExec.Datalog.WriteComment "Die Y :" & Shmoo_Y
         TheExec.Datalog.WriteComment "Pattern :" & Shmoo_Pattern
         TheExec.Datalog.WriteComment "Shmoo status :" & Shmoo_status
         TheExec.Datalog.WriteComment ""

         TimeDomainlist = TheHdw.Digital.Timing.TimeDomainlist
         TimeDomaingroup = Split(TimeDomainlist, ",")

         For Each CurrTimeDomain In TimeDomaingroup
             If CStr(CurrTimeDomain) = "All" Then
             TimeDomainIn = ""
             Else
             TimeDomainIn = CStr(CurrTimeDomain)
             End If

             Timelist = TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
             'TimeGroup
             TimeGroup = Split(Timelist, ",")
             For Each CurrTiming In TimeGroup
                If CurrTiming = "" Then Exit For
                If TimeDomainIn = "" Then
                      TheExec.Datalog.WriteComment "Time Doamin : TimeDomainIn = N/A"
                Else
                      TheExec.Datalog.WriteComment "Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & (1 / TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.Period(CStr(CurrTiming))) / 1000000 & " Mhz"
                End If

             Next CurrTiming
         Next CurrTimeDomain
        TheExec.Datalog.WriteComment ""
    'End If

    power_list = Split(Power_Search_String, ",")
    Power_number = UBound(power_list)
    ReDim power_pins(Power_number)

    For i = 0 To Power_number
         power_Temp() = Split(power_list(i), "=")
         Power_range_temp() = Split(power_Temp(1), ":")
         power_pins(i) = power_Temp(0)
         powerPin = powerPin + "," + power_pins(i)
         Shmoo_value.AddPin power_pins(i)
    Next i
    powerPin = Mid(powerPin, 2, Len(powerPin))

'''    Shmoo_Vcc_Max(CurrSite) = 0.5
'''    Shmoo_Vcc_Min(CurrSite) = 0.5

  If LVCC = True Then
     If RangeFrom > RangeTo Then
        Shmoo_steps = Fix((Shmoo_Vcc_Min(CurrSite) - RangeTo) / RangeStepSize)
     ElseIf RangeTo > RangeFrom Then
        Shmoo_steps = Fix((Shmoo_Vcc_Min(CurrSite) - RangeFrom) / RangeStepSize)
     Else
        Shmoo_steps = 3
     End If
  Else
     If RangeFrom > RangeTo Then
        Shmoo_steps = Fix((RangeFrom - Shmoo_Vcc_Max(CurrSite)) / RangeStepSize)
     ElseIf RangeTo > RangeFrom Then
        Shmoo_steps = Fix((RangeTo - Shmoo_Vcc_Max(CurrSite)) / RangeStepSize)
     Else
        Shmoo_steps = 3
     End If
  End If

    If Not Shmoo_steps < 1 Then
       k = Shmoo_steps
    Else
       k = 0
    End If

    Fail_log_cnt = 1


    
   If Not k = 0 Then
    '=============================================================================================================================
        For j = 0 To k                      '''''step loop
            For M = 0 To UBound(ShmooPatternSplit)      '''''pat loop
                '----------------------------------------------------------------------confirm init or pl
                GetPatFromPatternSet CStr(ShmooPatternSplit(M)), pat_list_array, pat_list_count
                pat_str_ary = Split(pat_list_array(0), ":")
                If pat_list_array(0) Like "*:*" Then
                    pat_str_ary = Split(pat_str_ary(1), "_")
                Else
                    pat_str_ary = Split(pat_str_ary(0), "\")
                    pat_str_ary = Split(pat_str_ary(UBound(pat_str_ary)), "_")
                End If

                If LCase(pat_str_ary(3)) Like LCase("pl*") Or LCase(pat_str_ary(3)) Like LCase("fu*") Then
                    confirm_pat_load = True
                    If LCase(pat_str_ary(3)) Like LCase("pllp") Or LCase(pat_str_ary(3)) Like LCase("fulp") Then
                       confirm_pat_load = True
                    Else
                        For k = 3 To Len(pat_str_ary(3))
                            If Mid(pat_str_ary(3), k, 1) >= "0" And Mid(pat_str_ary(3), k, 1) <= "9" Then
                                If confirm_pat_load <> False Then confirm_pat_load = True
                            Else
                                confirm_pat_load = False
                            End If
                        Next k
                    End If
                Else
                    confirm_pat_load = False
                End If
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''init
                If confirm_pat_load <> True Then
                    '-----------------------------------------------------------------assign power voltage
                    TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
                    TheHdw.Wait 0.0001

                    For i = 0 To g_ApplyLevelTimingVmain.Pins.Count - 1
                        TheHdw.DCVS.Pins(g_ApplyLevelTimingVmain.Pins(i).Name).Voltage.Value = g_ApplyLevelTimingVmain.Pins(i).Value
                    Next i

                    '''''''''''''''''''''''''''''''''confirm whether DSSC source pat
                    If InStr(dssc_pat_init_GB, ShmooPatternSplit(M)) > 0 Then
                      dssc_pat_exist = True
                        For i = 0 To Power_number
                           If LVCC = True Then
                              Shmoo_value.Pins(power_pins(i)).Value = Shmoo_Vcc_Min(CurrSite) - RangeStepSize * (j + 1)
                           Else
                              Shmoo_value.Pins(power_pins(i)).Value = Shmoo_Vcc_Max(CurrSite) + RangeStepSize * (j + 1)
                           End If
                        Next i
                        
                        DSSC_string = ""
                        DSSC_string = Decide_Switching_Bit_Debug_LVCC(digSrc_EQ_GB, DigSrc_wav, DC_Spec_Level, BlockType_GB, SELSRM_Rails, powerPin, Shmoo_value, g_VDDForce, g_CharInputString_Voltage_Dict)
                        Call SetupDigSrcDspWave(dssc_pat_init_GB, DigSrc_pin_GB, "FUNC_SRC", CLng(DigSrcSize_GB), DigSrc_wav)
                    End If
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    If M = 0 Then
                       StepValue = (j + 1) * RangeStepSize
                       If LVCC = True Then
                          TheExec.Datalog.WriteComment "Power setup (Vmin - " & StepValue * 1000 & "mV) "
                       Else
                          TheExec.Datalog.WriteComment "Power setup (Vmax + " & StepValue * 1000 & "mV) "
                       End If
                    End If
                    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                    thehdw.Wait 0.002
                      'thehdw.Patterns(ShmooPatternSplit(M)).test pfAlways, 0    ''''''''pat set for scan test
                      TheHdw.Patterns(ShmooPatternSplit(M)).start
                      TheHdw.Digital.Patgen.HaltWait
                      HardIP_WriteFuncResult
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''pl
                ElseIf confirm_pat_load = True Then
                    '---------------------------------------------for non shmoo pins
                    If g_CharInputString_Voltage_Dict.Count <> 0 Then
                        For i = 0 To UBound(pin_name)       '--------apply force_condition(Dictionary)
                           If Not g_CharInputString_Voltage_Dict.Exists(pin_name(i)) Then
                              TheHdw.DCVS.Pins(pin_name(i)).Voltage.Alt.Value = g_CharInputString_Voltage_Dict(pin_name(i))
                           End If
                        Next i
                    End If
                    '---------------------------------------------for shmoo pins
                    For i = 0 To Power_number
                       If LVCC = True Then
                          TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value = Shmoo_Vcc_Min(CurrSite) - RangeStepSize * (j + 1)
                       Else
                          TheHdw.DCVS.Pins(power_pins(i)).Voltage.Alt.Value = Shmoo_Vcc_Max(CurrSite) + RangeStepSize * (j + 1)
                       End If
                    Next i
                    '---------------------------------------------
                    '---------------------------------------------
                    TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt
                    TheHdw.Wait 0.0001
                    '---------------------------------------------
                    'thehdw.Patterns(ShmooPatternSplit(M)).test pfAlways, 0
                    TheHdw.Patterns(ShmooPatternSplit(M)).start
                    TheHdw.Digital.Patgen.HaltWait
                    HardIP_WriteFuncResult
                    TheExec.Datalog.WriteComment "                                                "
                    If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False And LVCC_boundary_Switch = Fail_log_cnt Then
                       For i = 0 To Power_number
                           TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Value, "0.000"))
                           TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Value, "0.000")) + ("V")
                       Next i
                       GoTo Endfor
                    End If
                    If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then Fail_log_cnt = Fail_log_cnt + 1
                    '=============================================================================================================print log
                       For i = 0 To Power_number
                            TheExec.Datalog.WriteComment power_pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Value, "0.000"))
                            TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(power_pins(i)).Voltage.Value, "0.000")) + ("V")
                       Next i
                       If dssc_pat_exist = True Then TheExec.Datalog.WriteComment "DSSC Sorce = " & DSSC_string & ";  SELSRM_Rails=" & SELSRM_Rails
'                     End If
                    '=============================================================================================================
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''error
                End If
            Next M
        Next j
    '=============================================================================================================================
    Else
       TheExec.Datalog.WriteComment "Shmoo HVCC all pass"
    End If
Endfor:

TheExec.Datalog.WriteComment "***************** Shmoo fail log/Pins capture end *****************"
'End If

Exit Function

errHandler_faillog:
'TheExec.Datalog.WriteComment "*****************Shmoo with fail log capture error *****************"

End Function
