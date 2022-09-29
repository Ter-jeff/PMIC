Attribute VB_Name = "VBT_LIB_DebugPrint"

Option Explicit
Public gb_DebugPrintFlag_Chk As Boolean
Public Const AllDiffmeterlist As String = "All_DMPINS"
Public Function DebugPrintFunc_mod(Test_Pattern As String, Optional testname_enable As Boolean = False) As Long
'for debug printing generation
    Dim funcName As String:: funcName = "DebugPrintFunc_mod"
    Dim PinCnt As Long, pinary() As String
    Dim i As Long
    Dim PowerVolt As Double
    Dim Powerfoldlimit As Double
    Dim AlramCheck As String
    Dim PowerAlramTime As Double
    Dim All_power_list As PinList
    Dim CurrentChans As String
    Dim PatSetArray() As String
    Dim PrintPatSet As Variant
    Dim patt As Variant 'patt1
    Dim patt1 As Variant
    Dim patt_ary_debug() As String
    Dim pat_count_debug As Long
    Dim patt_ary_debug1() As String
    Dim pat_count_debug1 As Long
    Dim PinGroup() As String
    Dim EachPinGroup As Variant
    Dim Timelist As String
    Dim TimeGroup() As String
    Dim CurrTiming As Variant
    Dim TimeDomainlist As String
    Dim TimeDomaingroup() As String
    Dim CurrTimeDomain As Variant
    Dim TimeDomainIn As String
    Dim TempString As String
    Dim TempStringOffline As String
    Dim AlarmBehavior As tlAlarmBehavior
    Dim DebugPrint_version As Double
    Dim Vmain As Double
    Dim IRange As Double
    Dim Gate_State As Boolean
    Dim Gate_State_str As String
    Dim PinData As New PinListData
    Dim out_line As String
    Dim CurrSite As Variant
    Dim XI0_Vicm  As Double
    Dim XI0_Vid As Double
    Dim XI0_Vihd As Double
    Dim XI0_Vild As Double
    Dim SlotType As String
    Dim AlarmBehavior_DCVI As String
    
    On Error GoTo ErrHandler
    
    'version history
    'DebugPrint_version = 1.3   'copy from Fiji
    'DebugPrint_version = 1.4   'implement offline simulation for Rhea bring up
    'DebugPrint_version = 1.5   'Update for Multi-Port nWire setting
     DebugPrint_version = 1.6   'Add differential nWire frequency capture, DCVS tl* modes put in strings, support no pattern items
     DebugPrint_version = 1.7   'Add DC/AC cetegory setup, remove off-limt timing simulation, offline could get real timing.
     DebugPrint_version = 1.71   'Add PPMU debug print function.
     DebugPrint_version = 1.72   'Add DCVI debug print support.
     Shmoo_Pattern = Test_Pattern

    
'    for debug
 '   gb_DebugPrintFlag_Chk = True
    'setups
    If gb_DebugPrintFlag_Chk = True Then
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "================debug print start=================="
        'list all power pin's level
        TheExec.Datalog.WriteComment "  DebugPrint version = " & DebugPrint_version
        If testname_enable Then
            TheExec.Datalog.WriteComment "  TestInstanceName = " & G_TestName
            testname_enable = False
        Else
            TheExec.Datalog.WriteComment "  TestInstanceName = " & TheExec.DataManager.InstanceName
        End If
        
        TheExec.Datalog.WriteComment "***** List all Category info Start ******"
        ''''Get the current TestInstance Context
        Dim m_DCCategory As String
        Dim m_DCSelector As String
        Dim m_ACCategory As String
        Dim m_ACSelector As String
        Dim m_TimeSetSheet As String
        Dim m_EdgeSetSheet As String
        Dim m_LevelsSheet As String
        Dim m_tmpPMname As String
    
        ''''20151109
        ''''Use the local module private global variable to be flexible if it could be used anywhere in this Module. (Just in case)
        Call TheExec.DataManager.GetInstanceContext(m_DCCategory, m_DCSelector, _
                                                    m_ACCategory, m_ACSelector, _
                                                    m_TimeSetSheet, m_EdgeSetSheet, _
                                                    m_LevelsSheet, "")
    
        TempString = "DC Category ="
        TempString = TempString + " " + m_DCCategory
        TheExec.Datalog.WriteComment TempString
    
        TempString = "AC Category ="
        TempString = TempString + " " + m_ACCategory
        TheExec.Datalog.WriteComment TempString
    
        TempString = "Level ="
        TempString = TempString + " " + m_LevelsSheet
        TheExec.Datalog.WriteComment TempString

        TheExec.Datalog.WriteComment "***** List all Category info end ******"
        TheExec.Datalog.WriteComment "***** List all power Start ******"

        TheExec.DataManager.DecomposePinList AllPowerPinlist, pinary(), PinCnt
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        TempStringOffline = PinAry(i) & "_GLB"
'                        If LCase(TheExec.DataManager.InstanceName) Like "*_hv*" Then
'                            Vmain = TheExec.specs.Globals(TempStringOffline).ContextValue * TheExec.specs.Globals("Ratio_Plus").ContextValue
'                        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*_nv*" Then
'                            Vmain = TheExec.specs.Globals(TempStringOffline).ContextValue '* TheExec.Specs.Globals("Ratio_Plus").ContextValue
'                        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*_lv*" Then
'                            Vmain = TheExec.specs.Globals(TempStringOffline).ContextValue * TheExec.specs.Globals("Ratio_Minus").ContextValue
'                        End If
'                        'PowerVolt = Vmain
'                        PowerVolt = TheHdw.DCVS.Pins(PinAry(i)).Voltage.Main.Value
'                    Else    'online
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                                
                        Select Case SlotType
                            Case "hexvs": PowerVolt = TheHdw.DCVS.Pins(pinary(i)).Voltage.Main.Value
                            Case "vhdvs": PowerVolt = TheHdw.DCVS.Pins(pinary(i)).Voltage.Main.Value
                            Case "dc-07": PowerVolt = TheHdw.DCVI.Pins(pinary(i)).Voltage
                            Case "dc-30": PowerVolt = TheHdw.DCVI.Pins(pinary(i)).Voltage
                        End Select
'                    End If
                        
                    TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Format(PowerVolt, "0.000") & " v"
                End If
            Next i

            TheExec.Datalog.WriteComment "***** List all power end ******"
            TheExec.Datalog.WriteComment "***** List all power FoldLimit TimeOut Start ******"

            TempString = "FoldLimit TimeOut :"
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'PowerAlramTime = 0.001 * i
'                        PowerAlramTime = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.TimeOut
'                    Else    'online
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerAlramTime = TheHdw.DCVS.Pins(pinary(i)).CurrentLimit.Source.FoldLimit.TimeOut
                            Case "vhdvs": PowerAlramTime = TheHdw.DCVS.Pins(pinary(i)).CurrentLimit.Source.FoldLimit.TimeOut
                            Case "dc-07": PowerAlramTime = TheHdw.DCVI.Pins(pinary(i)).FoldCurrentLimit.TimeOut
                            Case Else: PowerAlramTime = -999
                            'Case "dc-30": PowerAlramTime = TheHdw.DCVI.Pins(pinary(i)).FoldCurrentLimit.TimeOut
                        End Select
'                    End If

                If PowerAlramTime = -999 Then
                       
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & "N / A"
                    Else
                        
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Format(1000 * PowerAlramTime, "0.000") & " ms"
                    End If
                End If
            Next i

          '  TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power FoldLimit TimeOut End ******"
            TheExec.Datalog.WriteComment "***** List all power FoldLimit Current Start ******"

            TempString = "FoldLimit Current :"
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        TempStringOffline = PinAry(i) & "_Ifold_GLB"
'                        Irange = TheExec.specs.Globals(TempStringOffline).ContextValue
'                        'Powerfoldlimit = Irange
'                        Powerfoldlimit = TheHdw.DCVS.Pins(PinAry(i)).CurrentLimit.Source.FoldLimit.Level.Value
'                    Else    'online
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": Powerfoldlimit = TheHdw.DCVS.Pins(pinary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                            Case "vhdvs": Powerfoldlimit = TheHdw.DCVS.Pins(pinary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                            Case "dc-07": Powerfoldlimit = TheHdw.DCVI.Pins(pinary(i)).Current
                            Case "dc-30": Powerfoldlimit = TheHdw.DCVI.Pins(pinary(i)).Current
                        End Select
'                    End If

                    If i <> (PinCnt - 1) Then
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Format(Powerfoldlimit, "0.000000") & " A" '+ ","
                    Else
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Format(Powerfoldlimit, "0.000000") & " A"
                    End If
                End If
            Next i

            'TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power FoldLimit Current End ******"
            TheExec.Datalog.WriteComment "***** List all power Alram Check Start ******"

            TempString = "Alram Check :"
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'AlarmBehavior = tlAlarmDefault
'                        AlarmBehavior = TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
'                    Else
                        
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": AlarmBehavior = TheHdw.DCVS.Pins(pinary(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
                            Case "vhdvs": AlarmBehavior = TheHdw.DCVS.Pins(pinary(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout)
                            Case "dc-07": AlarmBehavior_DCVI = TheHdw.DCVI.Pins(pinary(i)).FoldCurrentLimit.Behavior
                            'Case "dc-30": a = TheHdw.DCVI.Pins(pinary(i)).FoldCurrentLimit.Behavior
                            Case Else: AlarmBehavior_DCVI = -999
                        End Select
'                    End If
                    
                    If AlarmBehavior_DCVI = "0" Then
                        AlramCheck = "tlDCVICurrentLimitBehaviorDoNotGateOff"
                    ElseIf AlarmBehavior_DCVI = "1" Then
                        AlramCheck = "tlDCVICurrentLimitBehaviorGateOff"
                    ElseIf AlarmBehavior_DCVI = "-999" Then
                        AlramCheck = "N/A"
                    End If
                    
'                    If AlarmBehavior = tlAlarmOff Then
'                        AlramCheck = "tlAlarmOff"
'                    ElseIf AlarmBehavior = tlAlarmContinue Then
'                        AlramCheck = "tlAlarmContinue"
'                    ElseIf AlarmBehavior = tlAlarmDefault Then
'                        AlramCheck = "tlAlarmDefault"
'                    ElseIf AlarmBehavior = tlAlarmForceBin Then
'                        AlramCheck = "tlAlarmForceBin"
'                    ElseIf AlarmBehavior = tlAlarmForceFail Then
'                        AlramCheck = "tlAlarmForceFail"
'                    End If
                    If i <> (PinCnt - 1) Then
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & AlramCheck
                    Else
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & AlramCheck
                    End If
                End If
            Next i

            'TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power Alram Check End ******"
            TheExec.Datalog.WriteComment "***** List all power Connection Check Start ******"

            TempString = "Power Relay Connection:"
            Dim PowerConnect_State As tlDCVSConnectWhat
            Dim PowerConnect_State_str As String
            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'PowerConnect_State = tlDCVSConnectForce
'                        PowerConnect_State = TheHdw.DCVS.Pins(PinAry(i)).Connected
'                    Else
                        SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": PowerConnect_State = TheHdw.DCVS.Pins(pinary(i)).Connected
                            Case "vhdvs": PowerConnect_State = TheHdw.DCVS.Pins(pinary(i)).Connected
                            Case "dc-07": PowerConnect_State = TheHdw.DCVI.Pins(pinary(i)).Connected
                            Case "dc-30": PowerConnect_State = TheHdw.DCVI.Pins(pinary(i)).Connected
                        End Select
'                    End If

                If SlotType = "dc-07" Then
                        Select Case PowerConnect_State
                                 Case "0": PowerConnect_State_str = " Not connected "
                                 Case "3": PowerConnect_State_str = "ConnecteDefault"
                                 Case "1": PowerConnect_State_str = "High force connected"
                                 Case "2": PowerConnect_State_str = "High sense connected"
'                                 Case "4": PowerConnect_State_str = " High guard connected "
'                                 Case "8": PowerConnect_State_str = " Low force connected "
'                                 Case "16": PowerConnect_State_str = " Low sense connected "
                        End Select
                
                ElseIf SlotType = "dc-30" Then
                        Select Case PowerConnect_State
                                 Case "0": PowerConnect_State_str = " Not connected "
                                 Case "1": PowerConnect_State_str = "High force connected"
                                 Case "2": PowerConnect_State_str = "High sense connected"
                                 Case "3": PowerConnect_State_str = "High force and High sense connected"
                                 Case "4": PowerConnect_State_str = "High guard connected"
                                 Case "5": PowerConnect_State_str = "High force and High guard connected"
                                 Case "6": PowerConnect_State_str = "High sense and High guard connected"
                                 Case "7": PowerConnect_State_str = "ConnecteDefault"""
                        End Select
                
                End If
'                    Select Case PowerConnect_State
'                         Case tlDCVSConnectDefault: PowerConnect_State_str = "tlDCVSConnectDefault"
'                         Case tlDCVSConnectNone: PowerConnect_State_str = "tlDCVSConnectNone"
'                         Case tlDCVSConnectForce: PowerConnect_State_str = "tlDCVSConnectForce"
'                         Case tlDCVSConnectSense: PowerConnect_State_str = "tlDCVSConnectSense"
'                    End Select
                    If i <> (PinCnt - 1) Then
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & PowerConnect_State_str
                    Else
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & PowerConnect_State_str
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power Connection Check End ******"
            TheExec.Datalog.WriteComment "***** List all power Gate Start ******"

            TempString = "Power Gate Status:"

            For i = 0 To PinCnt - 1
                If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
'                    If TheExec.TesterMode = testModeOffline Then
'                        'Gate_State = True
'                        Gate_State = TheHdw.DCVS.Pins(PinAry(i)).Gate
'                    Else
                       SlotType = LCase(GetInstrument(pinary(i), 0))
                        Select Case SlotType
                            Case "hexvs": Gate_State = TheHdw.DCVS.Pins(pinary(i)).Gate
                            Case "vhdvs": Gate_State = TheHdw.DCVS.Pins(pinary(i)).Gate
                            Case "dc-07": Gate_State = TheHdw.DCVI.Pins(pinary(i)).Gate
                            Case "dc-30": Gate_State = TheHdw.DCVI.Pins(pinary(i)).Gate
                        End Select
'                    End If
                    
                    Select Case Gate_State
                         Case True: Gate_State_str = "on"
                         Case False: Gate_State_str = "off"
                    End Select
                    If i <> (PinCnt - 1) Then
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Gate_State_str + ","
                    Else
                        TheExec.Datalog.WriteComment "  " & pinary(i) & " = " & Gate_State_str
                    End If
                End If
            Next i

            TheExec.Datalog.WriteComment TempString
            TheExec.Datalog.WriteComment "***** List all power Gate Check End ******"
            TheExec.Datalog.WriteComment "***** List Pattern Start ******"

            'Print test pattern
            If Test_Pattern <> "" Then
                PatSetArray = Split(Test_Pattern, ",")

                For Each PrintPatSet In PatSetArray
                    If LCase(PrintPatSet) Like "*.pat*" Then
                        TheExec.Datalog.WriteComment "  Pattern : " & PrintPatSet
                    Else
                        GetPatListFromPatternSet CStr(PrintPatSet), patt_ary_debug, pat_count_debug
                        For Each patt In patt_ary_debug
                            If patt <> "" Then TheExec.Datalog.WriteComment "  Pattern : " & patt
                        Next patt
                    End If
                Next PrintPatSet
            Else
                'do nothing, no printing
            End If

            TheExec.Datalog.WriteComment "***** List Pattern end ******"
            TheExec.Datalog.WriteComment "***** List Level Start ******"
'
            PinGroup = Split(PinGrouplist, ",")
            For Each EachPinGroup In PinGroup   'EachPinGroup
                TheExec.Datalog.WriteComment "  Pins : " & CStr(EachPinGroup) _
                & " , Vih = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVih), "0.000") & " v" _
                & " , Vil = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVil), "0.000") & " v" _
                & " , Voh = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVoh), "0.000") & " v" _
                & " , Vol = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVol), "0.000") & " v" _
                & " , Iol = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chIoh), "0.000") & " v" _
                & " , Ioh = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chIol), "0.000") & " v" _
                & " , Vt  = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVt), "0.000") & " v" _
                & " , Vch = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVch), "0.000") & " v" _
                & " , Vcl = " & Format(TheHdw.Digital.Pins(CStr(EachPinGroup)).Levels.Value(chVcl), "0.000") & " v" _
                & " , PPMU_VclampHi = " & Format(TheHdw.PPMU.Pins(CStr(EachPinGroup)).ClampVHi, "0.000") & " v" _
                & " , PPMU_VclampLow = " & Format(TheHdw.PPMU.Pins(CStr(EachPinGroup)).ClampVLo, "0.000") & " v"
            Next EachPinGroup

            TheExec.Datalog.WriteComment "***** List Level end ******"
            TheExec.Datalog.WriteComment "***** List Timing Start ******"

'            If Test_Pattern <> "" Then
                TimeDomainlist = TheHdw.Digital.Timing.TimeDomainlist
                TimeDomaingroup = Split(TimeDomainlist, ",")
                For Each CurrTimeDomain In TimeDomaingroup
                    If CStr(CurrTimeDomain) = "All" Then
                        TimeDomainIn = ""
                    Else
                        TimeDomainIn = CStr(CurrTimeDomain)
                    End If

                    Timelist = TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.TimeSetNameList
'                    TimeGroup
                    TimeGroup = Split(Timelist, ",")
                    For Each CurrTiming In TimeGroup
                        If CurrTiming <> "" Then
                            If TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.Period(CStr(CurrTiming)) > 0 Then
                                TheExec.Datalog.WriteComment "  Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & Format((1 / TheHdw.Digital.TimeDomains(TimeDomainIn).Timing.Period(CStr(CurrTiming))) / 1000000, "0.000") & " Mhz"
                            Else
                                TheExec.Datalog.WriteComment "  Time Doamin : " & CurrTimeDomain & ", TimeSet : " & CStr(CurrTiming) & " = " & Format(0, "0.000") & " Mhz"
                            End If
                        End If
                    Next CurrTiming
                Next CurrTimeDomain
'            Else
'                TheExec.Datalog.WriteComment "  Time Doamin : " & "N/A" & ", TimeSet : " & "N/A" & " = " & Format(0 / 1000000, "0.000") & " Mhz"
'            End If
'
            '' add for XI0 free running clk
'               TheExec.Datalog.WriteComment "  FreeRunFreq : " & TheHdw.DIB.SupportBoardClock.Frequency / 1000000 & " Mhz , clock_Vih: " & TheHdw.DIB.SupportBoardClock.Vih & " v , clock_Vil: " & TheHdw.DIB.SupportBoardClock.Vil & " v"
            Dim XI0_Freq_pl As New PinListData, RTCLK_Freq_pl As New PinListData, Pin_XI0 As New PinList, Pin_RTCLK As New PinList
            Dim Site As Variant

            If XI0_GP <> "" Then 'differential(false) or single end(true)
                Pin_XI0.Value = XI0_GP
                TheHdw.Digital.Pins(Pin_XI0).Levels.Value(chVoh) = TheHdw.Digital.Pins(Pin_XI0).Levels.Value(chVih) / 4
            ElseIf XI0_Diff_GP <> "" Then
                'Vod=0, do nothing
                Pin_XI0.Value = XI0_Diff_GP
                TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVod) = TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVid) / 4
            End If

            If XI0_Diff_GP <> "" Or XI0_GP <> "" Then
                Freq_MeasFreqSetup Pin_XI0, 0.001
                Freq_MeasFreqStart Pin_XI0, 0.001, XI0_Freq_pl
            End If

            If TheExec.TesterMode = testModeOffline Then
                For Each Site In TheExec.Sites
                    XI0_Freq_pl.Pins(0).Value = 24000000
                Next Site
            End If

            For Each Site In TheExec.Sites
                    If XI0_GP <> "" Then 'differential(false) or single end(true)
                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0) : " & Format(XI0_Freq_pl.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(Pin_XI0).Levels.Value(chVih), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(Pin_XI0).Levels.Value(chVil), "0.000") & " v"
                        'CHWu modify 10/14 to add Xio_PA_1 and remove RTCLK
'                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0_1) : " & Format(XI0_Freq_pl_1.pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.pins(Pin_XI0_1).Levels.Value(chVih), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.pins(Pin_XI0_1).Levels.Value(chVil), "0.000") & " v"
                    ElseIf XI0_Diff_GP <> "" Then
                      'CHWu modify 11/17 modify for Xio_PA printout
                       XI0_Vicm = TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVicm)
                       XI0_Vid = TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVid)
                       XI0_Vihd = XI0_Vicm + XI0_Vid / 2
                       XI0_Vild = XI0_Vicm - XI0_Vid / 2
                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0) : " & Format(XI0_Freq_pl.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(XI0_Vihd, "0.000") & " v , clock_Vil: " & Format(XI0_Vild, "0.000") & " v"
'                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0_1) : " & Format(XI0_Freq_pl_1.pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(XI0_Vihd, "0.000") & " v , clock_Vil: " & Format(XI0_Vild, "0.000") & " v"
'                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0) : " & Format(XI0_Freq_pl.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVid), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(Pin_XI0).DifferentialLevels.Value(chVod), "0.000") & " v"
'                        TheExec.Datalog.WriteComment "  FreeRunFreq (XI0_1) : " & Format(XI0_Freq_pl_1.Pins(0).Value / 1000000, "0.000") & " Mhz , clock_Vih: " & Format(TheHdw.Digital.Pins(Pin_XI0_1).DifferentialLevels.Value(chVid), "0.000") & " v , clock_Vil: " & Format(TheHdw.Digital.Pins(Pin_XI0_1).DifferentialLevels.Value(chVod), "0.000") & " v"
                    End If
            Next Site

            TheExec.Datalog.WriteComment "***** List Timing end ******"
'            TheExec.Datalog.WriteComment "***** List Disable Compare check Start ******"
'
'            'EachPinGroup
'            PinGroup = Split(PinGrouplist, ",")
'            For Each EachPinGroup In PinGroup
'                TheExec.Datalog.WriteComment "  Pins : " & CStr(EachPinGroup) _
'                & " , Disable Compare = " & TheHdw.Digital.Pins(EachPinGroup).DisableCompare
'            Next EachPinGroup
'
'            TheExec.Datalog.WriteComment "***** List List Disable Compare check End ******"
            TheExec.Datalog.WriteComment "***** List all utility bit status Start ******"
'            TheExec.DataManager.DecomposePinList All_Utility_list, pinary(), PinCnt
'
            'Utility bits
            out_line = "Utility_list : "
            TheExec.DataManager.DecomposePinList All_Utility_list, pinary(), PinCnt
            For Each CurrSite In TheExec.Sites.Active
                For i = 0 To PinCnt - 1
                    If TheExec.DataManager.ChannelType(pinary(i)) <> "N/C" Then
                        PinData = TheHdw.Utility.Pins(pinary(i)).States(tlUBStateProgrammed)    'TheHdw.Utility.pins((pinary(i)) '.States(tlUBStateCompared)
                        If i = 0 Then
                              TheExec.Datalog.WriteComment out_line + pinary(i) & " = " & PinData.Pins(0).Value(CurrSite)
                              'out_line = out_line + pinary(i) & " = " & PinData.Pins(0).Value(CurrSite) '''& ","
                        Else
                              TheExec.Datalog.WriteComment out_line + pinary(i) & " = " & PinData.Pins(0).Value(CurrSite)
                              'out_line = out_line & "," & pinary(i) & " = " & PinData.Pins(0).Value(CurrSite)
                        End If
                    End If
                Next i
                'TheExec.Datalog.WriteComment out_line
                out_line = "Utility_list : "
            Next CurrSite
            TheExec.Datalog.WriteComment "***** List all utility bit status end ******"

  TheExec.Datalog.WriteComment "***** ADG1414 bit status Start ******"
        '************add ADG1414
        Dim m_ADG1414ArgList() As String
        Dim PinName        As String
        Dim ChannelType    As String
        Dim GetStatus      As Long
        Dim mS_Temp    As String
        
        m_ADG1414ArgList = Split(g_ADG1414ArgList, ",")
        If UBound(m_ADG1414ArgList) = -1 Then
            TheExec.AddOutput "Please SVN update VBT_DibChecker-ADG1414_CONTROL."
        Else
            For i = 0 To UBound(m_ADG1414ArgList)
                PinName = m_ADG1414ArgList(i)
                ChannelType = "ADG1414"
                GetStatus = g_ADG1414Data(i)
              '  mS_Temp = FormatNumeric(PinName, -30) & "," & FormatNumeric(ChannelType, -15) & "," & FormatNumeric(GetStatus, -15) & "," & FormatNumeric("", -15) & "," & FormatNumeric("", -25)
                mS_Temp = PinName & " = " & GetStatus
                
                TheExec.Datalog.WriteComment mS_Temp
            Next i
        End If

 TheExec.Datalog.WriteComment "***** ADG1414 bit status End ******"


 TheExec.Datalog.WriteComment "***** Digital pins connected status start ******"

    TheExec.DataManager.DecomposePinList All_DigitalPinlist, pinary(), PinCnt
        For i = 0 To PinCnt - 1
        
            If TheHdw.Digital.Raw.Chans(pinary(i)).IsConnected = True Then
            TheExec.Datalog.WriteComment pinary(i) & " : is connected"
            Else
            TheExec.Datalog.WriteComment pinary(i) & " : disconnected"
            End If
        
        Next

 TheExec.Datalog.WriteComment "***** Digital pins connected status  End  ******"



 TheExec.Datalog.WriteComment "***** List DCVI pins NominalBandwidth start ******"
        TheExec.DataManager.DecomposePinList AllDCVIPinlist, pinary(), PinCnt
            For i = 0 To PinCnt - 1
            TheExec.Datalog.WriteComment pinary(i) & " : " & TheHdw.DCVI.Pins(pinary(i)).NominalBandwidth.Value & "Hz"
            Next
    
    TheExec.Datalog.WriteComment "***** List DCVI pins NominalBandwidth  End  ******"



 TheExec.Datalog.WriteComment "***** List DCVI pins Filter start ******"
   TheExec.DataManager.DecomposePinList AllDCVIPinlist, pinary(), PinCnt
   
            For i = 0 To PinCnt - 1
            TheExec.Datalog.WriteComment pinary(i) & "_Bypaas = " & TheHdw.DCVI.Pins(pinary(i)).Meter.Filter.Bypass & "| Filter = " & TheHdw.DCVI.Pins(pinary(i)).Meter.Filter.Value & "Hz"
            Next

 TheExec.Datalog.WriteComment "***** List DCVI pins Filter  End  ******"

 TheExec.Datalog.WriteComment "***** List Differmeter pins Filter start ******"
    TheExec.DataManager.DecomposePinList AllDCVIPinlist, pinary(), PinCnt
   
            For i = 0 To PinCnt - 1
            TheExec.Datalog.WriteComment pinary(i) & "_Bypaas = " & TheHdw.DCVI.Pins(pinary(i)).Meter.Filter.Bypass & "| Filter = " & TheHdw.DCVI.Pins(pinary(i)).Meter.Filter.Value & "Hz"
            Next

 TheExec.Datalog.WriteComment "***** List DCVI pins Filter  End  ******"
 
 
 TheExec.Datalog.WriteComment "***** List DiffMeter Mode start ******"
   TheExec.DataManager.DecomposePinList AllDiffmeterlist, pinary(), PinCnt
   
   
            For i = 0 To PinCnt - 1
            If UCase(pinary(i)) Like "*UVI80*" Then
                If TheHdw.DCDiffMeter.Pins(pinary(i)).MeterMode = tlDCDiffMeterModeHighAccuracy Then
                TheExec.Datalog.WriteComment pinary(i) & " = HighAccuracy "
                ElseIf TheHdw.DCDiffMeter.Pins(pinary(i)).MeterMode = tlDCDiffMeterModeHighSpeed Then
                TheExec.Datalog.WriteComment pinary(i) & " = HighSpeed "
                End If
            End If
            Next

TheExec.Datalog.WriteComment "***** List DiffMeter Mode End  ******"
        
            TheExec.Datalog.WriteComment "================debug print end  =================="
            TheExec.Datalog.WriteComment ""
        End If






    Exit Function
    
ErrHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function LIB_GEN_3Step_Trim_MultiTarget_DebugPrint(sPreTrimCodeVal As String, _
sMeasPinGrp As String, sCodeSeq As String, sLSB As String, _
 sTestBlock As String) As Long
    Dim funcName As String:: funcName = "LIB_GEN_3Step_Trim_MultiTarget_DebugPrint"
    On Error GoTo ErrorHandler

    Dim lPinCnt As Long
    Dim lTestIndex As Long
    Dim asTmpCodeArr() As String
    Dim asTmpTargetArr() As String
    Dim asTmpLSBArr() As String
    Dim asTmpCodeSeqArr() As String
    Dim asTmpMeasPin() As String
    Dim asTmpTestBlockArr() As String
    Dim sInstName As String

    TheExec.DataManager.DecomposePinList sMeasPinGrp, asTmpMeasPin(), lPinCnt
    asTmpCodeArr = Split(sPreTrimCodeVal, ",")
   
    asTmpLSBArr = Split(sLSB, ",")
    asTmpCodeSeqArr = Split(sCodeSeq, "+")
    asTmpTestBlockArr = Split(sTestBlock, ",")
    sInstName = UCase(TheExec.DataManager.InstanceName)

        
        
        TheExec.Datalog.WriteComment "***************** DebugPrint 3Step MultiTarget Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "

        
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  Block = " & asTmpTestBlockArr(lTestIndex)
        TheExec.Datalog.WriteComment "  CodeDistribution = " & asTmpCodeSeqArr(lTestIndex)
      '  TheExec.Datalog.WriteComment "  Target = " & asTmpTargetArr(lTestIndex)
        TheExec.Datalog.WriteComment "  LSB = " & asTmpLSBArr(lTestIndex)
        TheExec.Datalog.WriteComment "  PreTrimCode = " & asTmpCodeArr(lTestIndex)
        TheExec.Datalog.WriteComment "  MeasPin = " & asTmpMeasPin(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint 3Step MultiTarget Parameter End  ********************"
    



    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function LIB_GEN_3Step_Trim_DebugPrint(sPreTrimCodeVal As String, _
sMeasPinGrp As String, sCodeSeq As String, sLSB As String, _
sTarget As String, sTestBlock As String) As Long
    Dim funcName As String:: funcName = "LIB_GEN_3Step_Trim_DebugPrint"
    On Error GoTo ErrorHandler

    Dim lPinCnt As Long
    Dim lTestIndex As Long
    Dim asTmpCodeArr() As String
    Dim asTmpTargetArr() As String
    Dim asTmpLSBArr() As String
    Dim asTmpCodeSeqArr() As String
    Dim asTmpMeasPin() As String
    Dim asTmpTestBlockArr() As String
    Dim sInstName As String

    TheExec.DataManager.DecomposePinList sMeasPinGrp, asTmpMeasPin(), lPinCnt
    asTmpCodeArr = Split(sPreTrimCodeVal, ",")
    asTmpTargetArr = Split(sTarget, ",")
    asTmpLSBArr = Split(sLSB, ",")
    asTmpCodeSeqArr = Split(sCodeSeq, "+")
    asTmpTestBlockArr = Split(sTestBlock, ",")
    sInstName = UCase(TheExec.DataManager.InstanceName)
    


        
        
        TheExec.Datalog.WriteComment "***************** DebugPrint 3Step Trim Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  Block = " & asTmpTestBlockArr(lTestIndex)
        TheExec.Datalog.WriteComment "  CodeDistribution = " & asTmpCodeSeqArr(lTestIndex)
        TheExec.Datalog.WriteComment "  Target = " & asTmpTargetArr(lTestIndex)
        TheExec.Datalog.WriteComment "  LSB = " & asTmpLSBArr(lTestIndex)
        TheExec.Datalog.WriteComment "  PreTrimCode = " & asTmpCodeArr(lTestIndex)
        TheExec.Datalog.WriteComment "  MeasPin = " & asTmpMeasPin(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint 3Step Trim Parameter End  ********************"
    



    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function AnalogSweep_Voltage_MonitorDTB_DebugPrint(sSweep_Pin As String, dSweepStartVal As Double, dSweepStopVal As Double, dSweepStepVal As Double, dSweepInterval As Double, _
                                sToggle_Pin As String, dToggle_Threshold As Double, lToggle_Sample As Long) As Long
    Dim funcName As String:: funcName = "AnalogSweep_Voltage_MonitorDTB_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sSweepStartVal As String
    Dim sSweepStopVal As String
    Dim sSweepStepVal As String
    Dim sSweepInterval As String
    Dim sToggle_Threshold As String
    Dim sToggle_Sample As String
    Dim sInstName As String

  
    sSweepStartVal = CStr(dSweepStartVal)
    sSweepStopVal = CStr(dSweepStopVal)
    sSweepStepVal = CStr(dSweepStepVal)
    sSweepInterval = CStr(dSweepInterval)
    sToggle_Threshold = CStr(dToggle_Threshold)
    sToggle_Sample = CStr(lToggle_Sample)
    sInstName = UCase(TheExec.DataManager.InstanceName)
    
        
        TheExec.Datalog.WriteComment "***************** DebugPrint AnalogSweep Voltage MonitorDTB Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  SweepPin = " & sSweep_Pin
        TheExec.Datalog.WriteComment "  SweepStartValue = " & sSweepStartVal & " V "
        TheExec.Datalog.WriteComment "  SweepStopValue = " & sSweepStopVal & " V "
        TheExec.Datalog.WriteComment "  SweepStepValue = " & sSweepStepVal & " V "
'        TheExec.Datalog.WriteComment "  WaitTime = " & sSweepInterval & " s "
        TheExec.Datalog.WriteComment "  TogglePin = " & sToggle_Pin
        TheExec.Datalog.WriteComment "  ToggleThreshold = " & sToggle_Threshold & " V "
'        TheExec.Datalog.WriteComment "  SampleSize = " & sToggle_Sample
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "***************** DebugPrint AnalogSweep Voltage MonitorDTB Parameter End  ********************"


    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function AnalogSweep_Current_MonitorDTB_DebugPrint(dSweepStartVal As Double, dSweepStopVal As Double, dSweepStepVal As Double, dSweepInterval As Double, sSweep_Pin As String, _
                                sToggle_Pin As String, dToggle_Threshold As Double, lToggle_Sample As Long) As Long

    Dim funcName As String:: funcName = "AnalogSweep_Current_MonitorDTB_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sSweepStartVal As String
    Dim sSweepStopVal As String
    Dim sSweepStepVal As String
    Dim sSweepInterval As String
    Dim sToggle_Threshold As String
    Dim sToggle_Sample As String
    Dim sInstName As String

  
    sSweepStartVal = CStr(dSweepStartVal)
    sSweepStopVal = CStr(dSweepStopVal)
    sSweepStepVal = CStr(dSweepStepVal)
    sSweepInterval = CStr(dSweepInterval)
    sToggle_Threshold = CStr(dToggle_Threshold)
    sToggle_Sample = CStr(lToggle_Sample)
    sInstName = UCase(TheExec.DataManager.InstanceName)
    
        
        TheExec.Datalog.WriteComment "***************** DebugPrint AnalogSweep Current MonitorDTB Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  SweepPin = " & sSweep_Pin
        TheExec.Datalog.WriteComment "  SweepStartValue = " & sSweepStartVal & " A "
        TheExec.Datalog.WriteComment "  SweepStopValue = " & sSweepStopVal & " A "
        TheExec.Datalog.WriteComment "  SweepStepValue = " & sSweepStepVal & " A "
'        TheExec.Datalog.WriteComment "  WaitTime = " & sSweepInterval & " s "
        TheExec.Datalog.WriteComment "  TogglePin = " & sToggle_Pin
        TheExec.Datalog.WriteComment "  ToggleThreshold = " & sToggle_Threshold & " V "
'        TheExec.Datalog.WriteComment "  SampleSize = " & sToggle_Sample
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "***************** DebugPrint AnalogSweep Current MonitorDTB Parameter End  ********************"


    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function AnalogSweep_Dutycycle_MonitorDTB_DebugPrint(sSweep_Pattern As String, sSweep_Pin As String, dSweepStartVal As Double, dSweepStopVal As Double, dSweepStepVal As Double, dSweepInterval As Double, _
                                sToggle_Pin As String, dToggle_Threshold As Double, lToggle_Sample As Long) As Long
    Dim funcName As String:: funcName = "AnalogSweep_Dutycycle_MonitorDTB_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sSweepStartVal As String
    Dim sSweepStopVal As String
    Dim sSweepStepVal As String
    Dim sSweepInterval As String
    Dim sToggle_Threshold As String
    Dim sToggle_Sample As String
    Dim sInstName As String

  
    sSweepStartVal = CStr(dSweepStartVal)
    sSweepStopVal = CStr(dSweepStopVal)
    sSweepStepVal = CStr(dSweepStepVal)
    sSweepInterval = CStr(dSweepInterval)
    sToggle_Threshold = CStr(dToggle_Threshold)
    sToggle_Sample = CStr(lToggle_Sample)
    sInstName = UCase(TheExec.DataManager.InstanceName)
    
        
        TheExec.Datalog.WriteComment "***************** DebugPrint AnalogSweep DutyCycle Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  SweepPattern = " & sSweep_Pattern
        TheExec.Datalog.WriteComment "  SweepPin = " & sSweep_Pin
        TheExec.Datalog.WriteComment "  SweepStartValue = " & sSweepStartVal & " s "
        TheExec.Datalog.WriteComment "  SweepStopValue = " & sSweepStopVal & " s "
        TheExec.Datalog.WriteComment "  SweepStepValue = " & sSweepStepVal & " s "
'        TheExec.Datalog.WriteComment "  WaitTime = " & sSweepInterval & " s "
        TheExec.Datalog.WriteComment "  TogglePin = " & sToggle_Pin
        TheExec.Datalog.WriteComment "  ToggleThreshold = " & sToggle_Threshold & " V "
'        TheExec.Datalog.WriteComment "  SampleSize = " & sToggle_Sample
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "***************** DebugPrint AnalogSweep DutyCycle Parameter End  ********************"


    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function AnalogSweep_Voltage_MonitorREG_DebugPrint(sSweep_Pin As String, dSweepStartVal As Double, dSweepStopVal As Double, dSweepStepVal As Double, dSweepInterval As Double, _
                                lMeasureRegAddr As Long, lMeasureRegFieldOrBit As Long) As Long
    Dim funcName As String:: funcName = "AnalogSweep_Voltage_MonitorREG_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sSweepStartVal As String
    Dim sSweepStopVal As String
    Dim sSweepStepVal As String
    Dim sSweepInterval As String
    Dim sMeasureRegAddr As String
    Dim sMeasureRegFieldOrBit As String
    Dim sInstName As String

    sSweepStartVal = CStr(dSweepStartVal)
    sSweepStopVal = CStr(dSweepStopVal)
    sSweepStepVal = CStr(dSweepStepVal)
    sSweepInterval = CStr(dSweepInterval)
    sMeasureRegAddr = CStr(lMeasureRegAddr)
    sMeasureRegFieldOrBit = CStr(lMeasureRegFieldOrBit)
    sInstName = UCase(TheExec.DataManager.InstanceName)
    
        
        TheExec.Datalog.WriteComment "***************** DebugPrint AnalogSweep Voltage MonitorREG Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  SweepPin = " & sSweep_Pin
        TheExec.Datalog.WriteComment "  SweepStartValue = " & sSweepStartVal & " V "
        TheExec.Datalog.WriteComment "  SweepStopValue = " & sSweepStopVal & " V "
        TheExec.Datalog.WriteComment "  SweepStepValue = " & sSweepStepVal & " V "
'        TheExec.Datalog.WriteComment "  WaitTime = " & sSweepInterval & " s "
'        TheExec.Datalog.WriteComment "  TrimRegAddr = " & sMeasureRegAddr
'        TheExec.Datalog.WriteComment "  TrimRegAddrBitField = " & sMeasureRegFieldOrBit
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "***************** DebugPrint AnalogSweep Voltage MonitorREG Parameter End  ********************"


    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function Sweep_Voltage_BestPointSearch_DebugPrint(sSweep_Pin As String, sSweepStartVal As String, sSweepStepVal As String, lHowManyStep As Long, _
                                dWaitTime As Double, sMeasPinGrp As String, lSampleSize As Long, lSampleRate As Long) As Long
    Dim funcName As String:: funcName = "Sweep_Voltage_BestPointSearch_DebugPrint"
    On Error GoTo ErrorHandler
    
    Dim lPinCnt As Long
    Dim lTestIndex As Long
    Dim asTmpMeasPin() As String
    Dim asTmpSweepPin() As String
    Dim asSweepStartVal() As String
    Dim asSweepStepVal() As String
    Dim sSampleSize As String
    Dim sSamplerate As String
    Dim sHowmanystep As String
    Dim sWaittime As String
    Dim sInstName As String
    Dim iSite As Variant
    
    TheExec.DataManager.DecomposePinList sMeasPinGrp, asTmpMeasPin(), lPinCnt
    
    asTmpSweepPin = Split(sSweep_Pin, ",")
    asSweepStartVal = Split(sSweepStartVal, ",")
    asSweepStepVal = Split(sSweepStepVal, ",")
    sHowmanystep = CStr(lHowManyStep)
    sWaittime = CStr(dWaitTime)
    sSampleSize = CStr(lSampleSize)
    sSamplerate = CStr(lSampleRate)
    sInstName = UCase(TheExec.DataManager.InstanceName)
    
        
        TheExec.Datalog.WriteComment "***************** DebugPrint SweepVoltage BestPointSearch Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  SweepPin = " & asTmpSweepPin(lTestIndex)
        TheExec.Datalog.WriteComment "  SweepStartValue = " & asSweepStartVal(lTestIndex) & " V "
        TheExec.Datalog.WriteComment "  SweepStepValue = " & asSweepStepVal(lTestIndex) & " V "
        TheExec.Datalog.WriteComment "  WaitTime = " & sWaittime & " s "
        TheExec.Datalog.WriteComment "  SweepHowManyStep = " & sHowmanystep
        TheExec.Datalog.WriteComment "  MeasPin = " & asTmpMeasPin(lTestIndex)
        TheExec.Datalog.WriteComment "  SampleSize = " & sSampleSize
        TheExec.Datalog.WriteComment "  SampleRate = " & sSamplerate
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint SweepVoltage BestPointSearch Parameter End  ********************"


    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function Trimlink_Comparator_GenTool_DebugPrint(cFirst As Long, cMid1 As Long, cMid2 As Long, cLast As Long, _
                                    lRegName As Long, lBitField_Name As Long, _
                                    sSweep_Pin As String, dSweepStartVal As Double, dSweepStopVal As Double, dSweepStepVal As Double, dSweepInterval As Double, dAnalogTarget As Double, _
                                    sToggle_Pin As String, sToggle_Direction As String, dToggle_Threshold As Double, lToggle_Sample As Long, sSweepVoltageOrCurrent As String) As Long
    Dim funcName As String:: funcName = "Trimlink_Comparator_GenTool_DebugPrint"
    On Error GoTo ErrorHandler
    
    Dim sFirst As String
    Dim sMid1 As String
    Dim sMid2 As String
    Dim sLast As String
    Dim sCodeSeq As String
    Dim sSweepStartVal As String
    Dim sSweepStopVal As String
    Dim sSweepStepVal As String
    Dim sSweepInterval As String
    Dim sRegName As String
    Dim sBitField_Name As String
    Dim sAnalogTarget As String
    Dim sToggleThreshold As String
    Dim sSampleSize As String
    Dim sInstName As String
    Dim sUnit As String
    Dim sSweepStatus As String
    
    If sSweepVoltageOrCurrent = "I" Then
        sUnit = "A"
        sSweepStatus = "Current"
    Else
        sUnit = "V"
        sSweepStatus = "Voltage"
    End If
    
    sCodeSeq = CStr(cFirst) + "," + CStr(cMid1) + "," + CStr(cMid2) + "," + CStr(cLast)
    sSweepStartVal = CStr(dSweepStartVal)
    sSweepStopVal = CStr(dSweepStopVal)
    sSweepStepVal = CStr(dSweepStepVal)
    sSweepInterval = CStr(dSweepInterval)
    sAnalogTarget = CStr(dAnalogTarget)
    sRegName = CStr(lRegName)
    sBitField_Name = CStr(lBitField_Name)
    sToggleThreshold = CStr(dToggle_Threshold)
    sSampleSize = CStr(lToggle_Sample)
    sInstName = UCase(TheExec.DataManager.InstanceName)
    
        
        TheExec.Datalog.WriteComment "***************** DebugPrint Trimlink Comparator Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  CodeDistribution = " & sCodeSeq
        TheExec.Datalog.WriteComment "  SweepPin = " & sSweep_Pin
        TheExec.Datalog.WriteComment "  SweepStartValue = " & sSweepStartVal & " " & sUnit
        TheExec.Datalog.WriteComment "  SweepStopValue = " & sSweepStopVal & " " & sUnit
        TheExec.Datalog.WriteComment "  SweepStepValue = " & sSweepStepVal & " " & sUnit
        TheExec.Datalog.WriteComment "  WaitTime = " & sSweepInterval & " s "
        TheExec.Datalog.WriteComment "  TrimRegAddr = " & sRegName
        TheExec.Datalog.WriteComment "  TrimRegAddrBitField = " & sBitField_Name
        TheExec.Datalog.WriteComment "  TogglePin = " & sToggle_Pin
        TheExec.Datalog.WriteComment "  ToggleDirection = " & sToggle_Direction
        TheExec.Datalog.WriteComment "  ToggleThreshold = " & sToggleThreshold
        TheExec.Datalog.WriteComment "  ToggleSampleSize = " & sSampleSize
        TheExec.Datalog.WriteComment "  SweepVoltageOrCurrent = " & sSweepStatus
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "***************** DebugPrint Trimlink Comparator Parameter End  ********************"


    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function



Public Function LIB_GEN_BestCodeSearch_DebugPrint(lFirst As Long, lMid1 As Long, lMid2 As Long, lLast As Long, _
                    sMeasPinGrp As String, sTarget As String, sTestBlock As String) As Long
    Dim funcName As String:: funcName = "LIB_GEN_BestCodeSearch_DebugPrint"
    On Error GoTo ErrorHandler

    Dim lPinCnt As Long
    Dim lTestIndex As Long
    Dim asTmpCodeArr() As String
    Dim asTmpTargetArr() As String
    Dim asTmpLSBArr() As String
    Dim asTmpCodeSeqArr() As String
    Dim asTmpMeasPin() As String
    Dim asTmpTestBlockArr() As String
    Dim sInstName As String

    TheExec.DataManager.DecomposePinList sMeasPinGrp, asTmpMeasPin(), lPinCnt
    asTmpTargetArr = Split(sTarget, ",")
    asTmpTestBlockArr = Split(sTestBlock, ",")
    sInstName = UCase(TheExec.DataManager.InstanceName)
    


        
        
        TheExec.Datalog.WriteComment "***************** DebugPrint BestCodeSearch Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  CodeDistribution = " & lFirst & ", " & lMid1 & ", " & lMid2 & ", " & lLast
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  Block = " & asTmpTestBlockArr(lTestIndex)

        TheExec.Datalog.WriteComment "  Target = " & asTmpTargetArr(lTestIndex)
        TheExec.Datalog.WriteComment "  MeasPin = " & asTmpMeasPin(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint BestCodeSearch Parameter End  ********************"
    



    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function LIB_GEN_CalcCodeSearch_DebugPrint(lFirst As Long, lMid1 As Long, lMid2 As Long, lLast As Long, _
                    sMeasPinGrp As String, sTestBlock As String, sTarget As SiteDouble) As Long
    Dim funcName As String:: funcName = "LIB_GEN_CalcCodeSearch_DebugPrint"
    On Error GoTo ErrorHandler

    Dim lPinCnt As Long
    Dim lTestIndex As Long
    Dim asTmpCodeArr() As String
    Dim asTmpTargetArr() As String
    Dim asTmpLSBArr() As String
    Dim asTmpCodeSeqArr() As String
    Dim asTmpMeasPin() As String
    Dim asTmpTestBlockArr() As String
    Dim sInstName As String

    TheExec.DataManager.DecomposePinList sMeasPinGrp, asTmpMeasPin(), lPinCnt
    asTmpTargetArr = Split(sTarget, ",")
    asTmpTestBlockArr = Split(sTestBlock, ",")
    sInstName = UCase(TheExec.DataManager.InstanceName)
    


        
        
        TheExec.Datalog.WriteComment "***************** DebugPrint CalcCodeSearch Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  CodeDistribution = " & lFirst & ", " & lMid1 & ", " & lMid2 & ", " & lLast
''        For Each Site In TheExec.Sites
''        TheExec.Datalog.WriteComment "  Target(Site" & Site & ") = " & sTarget(Site)
''        Next Site
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  Block = " & asTmpTestBlockArr(lTestIndex)
        TheExec.Datalog.WriteComment "  MeasPin = " & asTmpMeasPin(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint CalcCodeSearch Parameter End  ********************"
    



    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function LIB_GEN_FourthStep_LinearSearch_DebugPrint(lFirst As Long, lMid1 As Long, lMid2 As Long, lLast As Long, _
                    sMeasPinGrp As String, sTarget As Double, sLSB As Double) As Long
    Dim funcName As String:: funcName = "LIB_GEN_FourthStep_LinearSearch_DebugPrint"
    On Error GoTo ErrorHandler

    Dim lPinCnt As Long
    Dim lTestIndex As Long
    Dim asTmpCodeArr() As String
    Dim asTmpTargetArr() As String
    Dim asTmpLSBArr() As String
    Dim asTmpCodeSeqArr() As String
    Dim asTmpMeasPin() As String
    Dim asTmpTestBlockArr() As String
    Dim sInstName As String

    TheExec.DataManager.DecomposePinList sMeasPinGrp, asTmpMeasPin(), lPinCnt
    asTmpTargetArr = Split(sTarget, ",")
  
    sInstName = UCase(TheExec.DataManager.InstanceName)
    asTmpLSBArr = Split(sLSB, ",")


        
        
        TheExec.Datalog.WriteComment "***************** DebugPrint FourthStep LinearSearch Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  CodeDistribution = " & lFirst & ", " & lMid1 & ", " & lMid2 & ", " & lLast
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  LSB = " & asTmpLSBArr(lTestIndex)
        TheExec.Datalog.WriteComment "  Target = " & asTmpTargetArr(lTestIndex)
        TheExec.Datalog.WriteComment "  MeasPin = " & asTmpMeasPin(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint FourthStep LinearSearch Parameter End  ********************"
    



    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function LIB_GEN_FourthStep_LsbBasedSearch_DebugPrint(lFirst As Long, lMid1 As Long, lMid2 As Long, lLast As Long, _
                    sMeasPinGrp As String, sTarget As Double, sLSB As Double) As Long
    Dim funcName As String:: funcName = "LIB_GEN_FourthStep_LsbBasedSearch_DebugPrint"
    On Error GoTo ErrorHandler

    Dim lPinCnt As Long
    Dim lTestIndex As Long
    Dim asTmpCodeArr() As String
    Dim asTmpTargetArr() As String
    Dim asTmpLSBArr() As String
    Dim asTmpCodeSeqArr() As String
    Dim asTmpMeasPin() As String
    Dim asTmpTestBlockArr() As String
    Dim sInstName As String

    TheExec.DataManager.DecomposePinList sMeasPinGrp, asTmpMeasPin(), lPinCnt
    asTmpTargetArr = Split(sTarget, ",")
  
    sInstName = UCase(TheExec.DataManager.InstanceName)
    asTmpLSBArr = Split(sLSB, ",")


        
        
        TheExec.Datalog.WriteComment "***************** DebugPrint FourthStep LsbBasedSearch Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  CodeDistribution = " & lFirst & ", " & lMid1 & ", " & lMid2 & ", " & lLast
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  LSB = " & asTmpLSBArr(lTestIndex)
        TheExec.Datalog.WriteComment "  Target = " & asTmpTargetArr(lTestIndex)
        TheExec.Datalog.WriteComment "  MeasPin = " & asTmpMeasPin(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint FourthStep LsbBasedSearch Parameter End  ********************"
    



    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function




Public Function Trimlink_GenTool_DebugPrint(sNumbits As String, _
sMeasPinGrp As String, sTarget As String, sTestBlock As String) As Long
    Dim funcName As String:: funcName = "Trimlink_GenTool_DebugPrint"
    On Error GoTo ErrorHandler
    Dim NumBits As String
    Dim lPinCnt As Long
    Dim lTestIndex As Long
    Dim asTmpCodeArr() As String
    Dim asTmpTargetArr() As String
    Dim asTmpLSBArr() As String
    Dim asTmpCodeSeqArr() As String
    Dim asTmpMeasPin() As String
    Dim asTmpTestBlockArr() As String
    Dim sInstName As String
    NumBits = sNumbits
    TheExec.DataManager.DecomposePinList sMeasPinGrp, asTmpMeasPin(), lPinCnt
    asTmpTargetArr = Split(sTarget, ",")
    asTmpTestBlockArr = Split(sTestBlock, ",")
    sInstName = UCase(TheExec.DataManager.InstanceName)
    


        
        
        TheExec.Datalog.WriteComment "***************** DebugPrint Trimlink GenTool Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  Block = " & asTmpTestBlockArr(lTestIndex)
        TheExec.Datalog.WriteComment "  Numbits = " & NumBits
        TheExec.Datalog.WriteComment "  Target = " & asTmpTargetArr(lTestIndex)
        TheExec.Datalog.WriteComment "  MeasPin = " & asTmpMeasPin(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint Trimlink GenTool Parameter End  ********************"
    



    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function Trimlink_GenTool_MultiTarget_DebugPrint(sNumbits As String, _
sMeasPinGrp As String, sTestBlock As String) As Long
    Dim funcName As String:: funcName = "Trimlink_GenTool_MultiTarget_DebugPrint"
    On Error GoTo ErrorHandler
    Dim NumBits As String
    Dim lPinCnt As Long
    Dim lTestIndex As Long
    Dim asTmpCodeArr() As String
    Dim asTmpTargetArr() As String
    Dim asTmpLSBArr() As String
    Dim asTmpCodeSeqArr() As String
    Dim asTmpMeasPin() As String
    Dim asTmpTestBlockArr() As String
    Dim sInstName As String
    

    
    NumBits = sNumbits
    TheExec.DataManager.DecomposePinList sMeasPinGrp, asTmpMeasPin(), lPinCnt
    asTmpTestBlockArr = Split(sTestBlock, ",")
    sInstName = UCase(TheExec.DataManager.InstanceName)
    


        
        
        TheExec.Datalog.WriteComment "***************** DebugPrint Trimlink GenTool MultiTarget Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        
        
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  Block = " & asTmpTestBlockArr(lTestIndex)
        TheExec.Datalog.WriteComment "  Numbits = " & NumBits
   '     TheExec.Datalog.WriteComment "  Target = " & asTmpTargetArr(lTestIndex)
        TheExec.Datalog.WriteComment "  MeasPin = " & asTmpMeasPin(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint Trimlink GenTool MultiTarget Parameter End  ********************"
    



    Exit Function
    
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function Code_Sweep_Parallel_FreeRunningDTB_DebugPrint(sCodeSeq As String, sToggle_Pin As String, _
                                                            sToggle_Direction As String, sTestBlock As String) As Variant
    Dim funcName As String:: funcName = "Code_Sweep_Parallel_FreeRunningDTB_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sInstName       As String
    Dim lTestIndex      As Long
    Dim lPinCnt         As Long

    Dim asTmpCodeSeqArr()       As String
    Dim asTmpTogglePin()        As String
    Dim asTmpToggle_Direction() As String
    Dim asTmpTestBlockArr()     As String

    TheExec.DataManager.DecomposePinList sToggle_Pin, asTmpTogglePin(), lPinCnt

    asTmpCodeSeqArr = Split(sCodeSeq, "+")
    asTmpTogglePin = Split(sToggle_Pin, ",")
    asTmpToggle_Direction = Split(sToggle_Direction, ",")
    asTmpTestBlockArr = Split(sTestBlock, ",")
    sInstName = UCase(TheExec.DataManager.InstanceName)
    
        TheExec.Datalog.WriteComment "***************** DebugPrint CodeSweep Parallel FreeRunningDTB Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  Block = " & asTmpTestBlockArr(lTestIndex)
        TheExec.Datalog.WriteComment "  CodeDistribution = " & asTmpCodeSeqArr(lTestIndex)
        TheExec.Datalog.WriteComment "  TogglePin = " & asTmpTogglePin(lTestIndex)
        TheExec.Datalog.WriteComment "  ToggleDirection = " & asTmpToggle_Direction(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint CodeSweep Parallel FreeRunningDTB Parameter End  ********************"

    Exit Function
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Code_Sweep_Parallel_MonitorDTB_DebugPrint(sCodeSeq As String, sToggle_Pin As String, sToggle_Direction As String, _
                                                        sToggle_Threshold As String, sTestBlock As String) As Variant
    Dim funcName As String:: funcName = "Code_Sweep_Parallel_MonitorDTB_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sInstName       As String
    Dim lTestIndex      As Long
    Dim lPinCnt         As Long

    Dim asTmpCodeSeqArr()       As String
    Dim asTmpTogglePin()        As String
    Dim asTmpToggle_Direction() As String
    Dim asTmpToggle_Threshold() As String
    Dim asTmpTestBlockArr()     As String

    TheExec.DataManager.DecomposePinList sToggle_Pin, asTmpTogglePin(), lPinCnt

    asTmpCodeSeqArr = Split(sCodeSeq, "+")
    asTmpToggle_Direction = Split(sToggle_Direction, ",")
    asTmpToggle_Threshold = Split(sToggle_Threshold, ",")
    asTmpTestBlockArr = Split(sTestBlock, ",")
    sInstName = UCase(TheExec.DataManager.InstanceName)

        TheExec.Datalog.WriteComment "***************** DebugPrint CodeSweep Parallel MonitorDTB Parameter Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
    For lTestIndex = 0 To lPinCnt - 1
        TheExec.Datalog.WriteComment "  Block = " & asTmpTestBlockArr(lTestIndex)
        TheExec.Datalog.WriteComment "  CodeDistribution = " & asTmpCodeSeqArr(lTestIndex)
        TheExec.Datalog.WriteComment "  TogglePin = " & asTmpTogglePin(lTestIndex)
        TheExec.Datalog.WriteComment "  ToggleDirection = " & asTmpToggle_Direction(lTestIndex)
        TheExec.Datalog.WriteComment "  ToggleThreshold = " & asTmpToggle_Threshold(lTestIndex)
        TheExec.Datalog.WriteComment "                                                                "
    Next lTestIndex
        TheExec.Datalog.WriteComment "***************** DebugPrint CodeSweep Parallel MonitorDTB Parameter End  ********************"

    Exit Function
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Code_Sweep_Serial_FreeRunningDTB_DebugPrint(lFirst As Long, lMid1 As Long, lMid2 As Long, lLast As Long, _
                                                            sToggle_Pin As String, sToggle_Direction As String, _
                                                            lRegName As Long, lBitField_Name As Long) As Variant
    Dim funcName As String:: funcName = "Code_Sweep_Serial_FreeRunningDTB_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sInstName       As String

    sInstName = UCase(TheExec.DataManager.InstanceName)

        TheExec.Datalog.WriteComment "***************** DebugPrint Parameter CodeSweep Serial FreeRunningDTB Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  CodeDistribution = " & lFirst & ", " & lMid1 & ", " & lMid2 & ", " & lLast
        TheExec.Datalog.WriteComment "  TogglePin = " & sToggle_Pin
        TheExec.Datalog.WriteComment "  ToggleDirection = " & sToggle_Direction
'        TheExec.Datalog.WriteComment "  TrimRegName = " & lRegName
'        TheExec.Datalog.WriteComment "  TrimBitFieldName = " & lBitField_Name
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "***************** DebugPrint Parameter CodeSweep Serial FreeRunningDTB End  ********************"

    Exit Function
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Code_Sweep_Serial_MonitorDTB_DebugPrint(lFirst As Long, lMid1 As Long, lMid2 As Long, lLast As Long, _
                                                            sToggle_Pin As String, sToggle_Direction As String, dToggle_Threshold As Double, _
                                                            lRegName As Long, lBitField_Name As Long) As Variant
    Dim funcName As String:: funcName = "Code_Sweep_Serial_MonitorDTB_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sInstName       As String
    
    sInstName = UCase(TheExec.DataManager.InstanceName)

        TheExec.Datalog.WriteComment "***************** DebugPrint Parameter CodeSweep Serial MonitorDTB Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  CodeDistribution = " & lFirst & ", " & lMid1 & ", " & lMid2 & ", " & lLast
        TheExec.Datalog.WriteComment "  TogglePin = " & sToggle_Pin
        TheExec.Datalog.WriteComment "  ToggleDirection = " & sToggle_Direction
        TheExec.Datalog.WriteComment "  ToggleThreshold = " & dToggle_Threshold
'        TheExec.Datalog.WriteComment "  TrimRegName = " & lRegName
'        TheExec.Datalog.WriteComment "  TrimBitFieldName = " & lBitField_Name
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "***************** DebugPrint Parameter CodeSweep Serial MonitorDTB End  ********************"

    Exit Function
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Code_Sweep_Tweak_Serial_MonitorDTB_DebugPrint(lFirst As Long, lMid1 As Long, lMid2 As Long, lLast As Long, _
                                                            sToggle_Pin As String, dToggle_Threshold As Double, _
                                                            lRegName As Long, lBitField_Name As Long, _
                                                            sStartDirection As String, lStartCode As Long, lStopCode As Long) As Variant
    Dim funcName As String:: funcName = "Code_Sweep_Tweak_Serial_MonitorDTB_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sInstName       As String
    
    sInstName = UCase(TheExec.DataManager.InstanceName)

        TheExec.Datalog.WriteComment "***************** DebugPrint Parameter CodeSweep Tweak Serial MonitorDTB Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  CodeDistribution = " & lFirst & ", " & lMid1 & ", " & lMid2 & ", " & lLast
        TheExec.Datalog.WriteComment "  TogglePin = " & sToggle_Pin
        TheExec.Datalog.WriteComment "  ToggleThreshold = " & dToggle_Threshold
'        TheExec.Datalog.WriteComment "  TrimRegName = " & lRegName
'        TheExec.Datalog.WriteComment "  TrimBitFieldName = " & lBitField_Name
        TheExec.Datalog.WriteComment "  StartDirection = " & sStartDirection
        TheExec.Datalog.WriteComment "  StartCode = " & lStartCode
        TheExec.Datalog.WriteComment "  StopCode = " & lStopCode
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "***************** DebugPrint Parameter CodeSweep Tweak Serial MonitorDTB End  ********************"

    Exit Function
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Code_Sweep_Serial_MonitorREG_DebugPrint(lFirst As Long, lMid1 As Long, lMid2 As Long, lLast As Long, _
                                                            lMeasureRegAddr As Long, sToggle_Direction As String, lMeasureRegFieldOrBit As Long, _
                                                            lRegName As Long, lBitField_Name As Long) As Variant
    Dim funcName As String:: funcName = "Code_Sweep_Serial_MonitorREG_DebugPrint"
    On Error GoTo ErrorHandler

    Dim sInstName       As String
    
    sInstName = UCase(TheExec.DataManager.InstanceName)

        TheExec.Datalog.WriteComment "***************** DebugPrint Parameter CodeSweep Serial MonitorREG Start *******************"
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  Instance Name : " & sInstName
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "  CodeDistribution = " & lFirst & ", " & lMid1 & ", " & lMid2 & ", " & lLast
        TheExec.Datalog.WriteComment "  ToggleReg = " & lMeasureRegAddr
        TheExec.Datalog.WriteComment "  ToggleRegBitFieldName = " & lMeasureRegFieldOrBit
        TheExec.Datalog.WriteComment "  Toggle_Direction = " & sToggle_Direction
'        TheExec.Datalog.WriteComment "  TrimRegName = " & lRegName
'        TheExec.Datalog.WriteComment "  TrimBitFieldName = " & lBitField_Name
        TheExec.Datalog.WriteComment "                                                                "
        TheExec.Datalog.WriteComment "***************** DebugPrint Parameter CodeSweep Serial MonitorREG End  ********************"

    Exit Function
ErrorHandler:
    TheExec.AddOutput "<Error>" + funcName + ":: Please check it out."
    TheExec.Datalog.WriteComment "<Error>" + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
