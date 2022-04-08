Attribute VB_Name = "VBT_LIB_Digital_Shmoo"
Option Explicit
'Revision History:
'V0.0 initial bring up

Public Interpose_PrePat_GLB As String
Public ReadHWPowerValue_GLB As String
Public PL_DC_conditions_GLB As String
Public Vbump_for_Interpose As Boolean



' ============
' Private Data
' ============

' Context values on the Test Instances sheet
Private m_TimeSetSheet As String, m_LevelsSheet As String

' States of driver features which are saved and restored
Private m_OldPatThreading As Boolean
Private m_OldFlagMatchEnable As Boolean
Private m_OldWaitFlagsHigh As Long
Private m_OldWaitFlagsLow As Long
Private m_OldMatchAllSites As Boolean

' Cached parameters for PostTest POSTPATBPF interpose function. This
' is needed for the pattern set breakpoint feature.
Private m_DrivePins As String
Private m_FloatPins As String
Private m_EndOfBodyF As String
Private m_EndOfBodyFArgs As String

Private m_InterposeFunctionsSet As Boolean

Private Const TL_E_AT_PATSET_BREAKPT = &HC0000014


'Public Type Tracking_Axis_Para
'
'End Type
Public Type Axis_Para
    X_axis As New SiteDouble ' only store X_axis info
    Y_axis As New SiteDouble ' only store Y_axis info
    Z_axis As New SiteDouble ' only store Z_axis info
    
'    X_axis_Tracking() As New SiteDouble ' only store X_axis Tracking info
'    Y_axis_Tracking() As New SiteDouble ' only store Y_axis Tracking info
'    Z_axis_Tracking() As New SiteDouble ' only store Z_axis Tracking info
    X_axis_Tracking() As New SiteDouble  ' only store X_axis Tracking info
    Y_axis_Tracking() As New SiteDouble ' only store Y_axis Tracking info
    Z_axis_Tracking() As New SiteDouble ' only store Z_axis Tracking info
    
    CurrResult As New SiteVariant
End Type

Public Type g_3DShmooCurrPointResult

    Axis_CurrPoint() As Axis_Para
'    Axis_CurrResult() As Axis_Para
    
End Type

Public g_ShmooResult As g_3DShmooCurrPointResult



Public g_TestNum As Long
Public Xaxis_index As Long
Public Yaxis_index As Long
Public Zaxis_index As Long
Public Count_Point As Long
Public MaxArrIndex As Long

Public X_Tracking_Point As Long
Public Y_Tracking_Point As Long
Public Z_Tracking_Point As Long
Public X_dimemsion As Boolean
Public Y_dimemsion As Boolean
Public Z_dimemsion As Boolean

Public X_Point As Long
Public Y_Point As Long
Public Z_Point As Long
Public LVCC_flag As Boolean
Public HVCC_flag As Boolean
Public StartPoint As Double
Public StopPoint As Double
Public StepSize As Double
Public RangeSeq(2)  As Boolean '0:X-axis, 1:Y-axis, 2:Z-axis

Public axis_val() As New SiteVariant 'XYZ
Public axis_pin() As String 'XYZ

' ===============
' Private Helpers
' ===============

' This template needs to know timing and levels sheet names.
' Fetch them from the Context Manager
Private Sub FetchContext()
    Dim A(0 To 4) As String

    ' For compatibility with 7.01.01 and earlier:
    ' In earlier versions, a contextmgr bug made using a MemberIndex > 0 act like the CurrentlyAppliedContext parameter was False.
    ' This caused "" to be returned for the output parameters...so that ApplyLevelsTiming was NOT called for 2nd & later members of a test group
    
    Dim MemberIndex As Long
    MemberIndex = TheExec.DataManager.MemberIndex
    
    Dim UseCurrentContext As Boolean
    UseCurrentContext = (MemberIndex = 0)
    
    Call m_STDSvcClient.dmgr.ContextMgr.GetInstanceContextInformation(TheExec.DataManager.instanceName, MemberIndex, _
                A(0), A(1), m_TimeSetSheet, A(2), A(3), A(4), m_LevelsSheet, True, UseCurrentContext)

End Sub

' Restore tester state to the default
Private Sub PostBody(DrivePins As String, FloatPins As String, WaitTimeDomain As String, WaitFlagA As tlWaitVal, _
                    WaitFlagB As tlWaitVal, WaitFlagC As tlWaitVal, WaitFlagD As tlWaitVal)

    If TheExec.Flow.IsRunning = False Then Exit Sub
    
    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF, TL_C_POSTPATBPF)
    m_InterposeFunctionsSet = False

    ' Return channels to the default start-state condition, as needed
    If NonBlank(DrivePins) Then Call tl_SetStartState(DrivePins, chstartNone)

    ' Return specified DUT pins, if any, to connection with tester pin-electronics & power
    If NonBlank(FloatPins) Then Call tl_ConnectTester(FloatPins)
    
    ' Restore flag match feature
    ' for compatibility, the flag set/restore should be conditional if asynchronous pattern start not disabled and not suspended
    If ((TheHdw.Patterns.EnableAsyncPatternStart <> tlAsyncPatternModeDisabled) And (TheHdw.Patterns.SuspendAsyncPatternStart = False)) Then
       ' If the flag match settings are defaults then should not call SetFlagMatch
       If ((WaitFlagA <> waitoff) And (WaitFlagB <> waitoff) And (WaitFlagC <> waitoff) And (WaitFlagD <> waitoff)) Then
          Call TheHdw.Digital.TimeDomains(WaitTimeDomain).Patgen.SetFlagMatch( _
                   m_OldFlagMatchEnable, m_OldWaitFlagsHigh, m_OldWaitFlagsLow, _
                   m_OldMatchAllSites)
       End If
    Else
        Call TheHdw.Digital.TimeDomains(WaitTimeDomain).Patgen.SetFlagMatch( _
         m_OldFlagMatchEnable, m_OldWaitFlagsHigh, m_OldWaitFlagsLow, _
                 m_OldMatchAllSites)
    End If
    ' Restore pattern threading
    TheHdw.Patterns().Threading.Enable = m_OldPatThreading
End Sub


' Run the pattern and see if it passed or failed
Private Sub Body(FloatPins As PinList, PatternTimeout As Double, Patterns As Pattern, _
                 ReportResult As PFType, ResultMode As tlResultMode)
    ' Remove specified DUT pins, if any, from connection to tester pin-electronics and other resources
    If NonBlank(FloatPins) Then Call tl_SetFloatState(FloatPins)
    m_FloatPins = FloatPins.Value
    
    ' Enable the pattern timeout counter
    TheHdw.Digital.Patgen.TimeoutEnable = True
    TheHdw.Digital.Patgen.TimeOut = PatternTimeout
End Sub


' =====================
' Private work routines
' =====================

' Do test setup.  This involves setting the timing and levels, connecting the pins, and
' various other functions in preparation for running the pattern.
Private Sub PreBody(DriveHiPins As PinList, DriveLoPins As PinList, DriveZPins As PinList, DisablePins As PinList, _
                    Util1Pins As PinList, Util0Pins As PinList, WaitFlagA As tlWaitVal, _
                    WaitFlagB As tlWaitVal, WaitFlagC As tlWaitVal, WaitFlagD As tlWaitVal, MatchAllSites As Boolean, _
                    PatThreading As Boolean, RelayMode As tlRelayMode, _
                    WaitTimeDomain As String, Interpose_PrePat As String)

    Dim ConnectAllPins As Boolean, LoadLevels As Boolean, LoadTiming As Boolean

    ' Save previous state of pattern threading and set according to parameter.
    m_OldPatThreading = TheHdw.Patterns().Threading.Enable
    TheHdw.Patterns().Threading.Enable = PatThreading

    ' Set drive state on specified utility pins
    If NonBlank(Util0Pins) Then Call tl_SetUtilState(Util0Pins, 0)
    If NonBlank(Util1Pins) Then Call tl_SetUtilState(Util1Pins, 1)
    
    
    ' Instruct functional voltages/currents hardware drivers to acquire
    '   drive/receive values from the DataManager and apply them.
    If NonBlank(m_LevelsSheet) Then LoadLevels = True
    
    ' Instruct functional timing hardware drivers to acquire timing values
    '   from the DataManager and apply them.
    If NonBlank(m_TimeSetSheet) Then LoadTiming = True
    
    ' Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '   of pins noted on the active levels sheet, if needed
    ConnectAllPins = True
    If (RelayMode <> TL_C_RELAYPOWERED) Then
        LoadLevels = True   'If levels are powered down, they must be powered up again
    End If
        
    ' ApplyLevelTiming will
    '   Optionally power down instruments and power supplies
    '   Optionally Close Pin-Electronics, High-Voltage, & Power Supply Relays,
    '       of pins noted on the active levels sheet
    '   Optionally load Timing and Levels information
    '   Set init-state driver conditions on specified pins
    '       Setting init state causes the pin to drive the specified value.  Init
    '       state is set once, during the prebody, before the first pattern burst.
    '       Default is to leave the pin driving whatever value it last drove during
    '       the previous pattern burst.

    '     thehdw.DCVS.pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff
    Call TheHdw.Digital.ApplyLevelsTiming(ConnectAllPins, LoadLevels, LoadTiming, RelayMode)
    
    
      '' 20150625 - Apply Char setup
'    If UCase(TheExec.CurrentJob) Like "*CHAR*" Then
'        If Interpose_PrePat <> "" Then
'            Call SetForceCondition(Interpose_PrePat)
'        End If
'    End If

    
''    Call StartSBClock(24000000)
''    Call ReStartFRC
    'add wait time here
    'Call thehdw.Wait(5 * 0.001)
    'theexec.Datalog.WriteComment ("add 5ms wait time for level switch")
    'end add wait time
    
    'thehdw.DCVS.pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff
    If NonBlank(DriveLoPins) Then Call tl_SetInitState(DriveLoPins, chInitLo)
    If NonBlank(DriveHiPins) Then Call tl_SetInitState(DriveHiPins, chInitHi)
    If NonBlank(DriveZPins) Then Call tl_SetInitState(DriveZPins, chInitoff)
    
    If NonBlank(DisablePins) Then Call tl_SetDisableState(DisablePins)
    
    ' Set start-state driver conditions on specified pins.
    ' Start state determines the driver value the pin is set to as each pattern burst starts.
    ' Default is to have start state automatically selected appropriately
    '   depending on the Format of the first vector of each pattern burst.
    If NonBlank(DriveLoPins) Then Call tl_SetStartState(DriveLoPins, chStartLo)
    If NonBlank(DriveHiPins) Then Call tl_SetStartState(DriveHiPins, chStartHi)
    If NonBlank(DriveZPins) Then Call tl_SetStartState(DriveZPins, chStartOff)
    m_DrivePins = tl_tm_CombineCslStrings(DriveHiPins, DriveLoPins)
    m_DrivePins = tl_tm_CombineCslStrings(DriveZPins, m_DrivePins)
    
    ' Read back state of flag feature for later restoration
    ' for compatibility, the flag set/restore should be conditional if asynchronous pattern start not disabled and not suspended
    If ((TheHdw.Patterns.EnableAsyncPatternStart <> tlAsyncPatternModeDisabled) And (TheHdw.Patterns.SuspendAsyncPatternStart = False)) Then
       ' If the flag match settings are defaults then should not call GetFlagMatch
       If ((WaitFlagA <> waitoff) And (WaitFlagB <> waitoff) And (WaitFlagC <> waitoff) And (WaitFlagD <> waitoff)) Then
          Call TheHdw.Digital.TimeDomains(WaitTimeDomain).Patgen.GetFlagMatch( _
                   m_OldFlagMatchEnable, m_OldWaitFlagsHigh, m_OldWaitFlagsLow, _
                   m_OldMatchAllSites)
       End If
    Else
        Call TheHdw.Digital.TimeDomains(WaitTimeDomain).Patgen.GetFlagMatch( _
                 m_OldFlagMatchEnable, m_OldWaitFlagsHigh, m_OldWaitFlagsLow, _
         m_OldMatchAllSites)
    End If

    ' Set desired state according to arguments.
    Call SetFlagMatch(WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD, _
        MatchAllSites, WaitTimeDomain)
End Sub


Public Function TPmode_Char_on()
TPModeAsCharz_GLB = True

''======char shmoo error code initial=====''

F_shmoo_abnormal_counter = True ''default turn on shmoo_abnormal_counter
shmoohole_count = 0
shmooallfail_count = 0
shmooalarm_count = 0
total_shmoo_count = 0
included_shmoo_count = 0
excluded_shmoo_count = 0
        
        
 Parse_EMA_DigSrcInfo
        
  ''======char shmoo error code initial=====''
        
    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = 75
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width = 60
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 70

    TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
        
End Function

Public Function TPmode_Char_off()
TPModeAsCharz_GLB = False
End Function

Private Function Validate(Patterns As Pattern, PatThreading As Boolean, _
                          DriveLoPins As PinList, DriveHiPins As PinList, _
                          DriveZPins As PinList, DisablePins As PinList, FloatPins As PinList, _
                          Util1Pins As PinList, Util0Pins As PinList, _
                          PatternTimeout As String, WaitTimeDomain As String) As Boolean
    
    Validate = True ' Assume the best and override if trouble is found
    
    If Not ValidatePatternThreading(Patterns, PatThreading, 1, True, 26) Then Validate = False
    
    ' Validate the pin state parameters.
    If Not ValidatePinStates(DriveLoPins, DriveHiPins, DriveZPins, DisablePins, _
                             FloatPins, Util1Pins, Util0Pins) Then Validate = False
        
    If ValidateNumeric(PatternTimeout, "PatternTimeout", 33) Then
        ' Validate  0.0 <= PatternTimeout
        If Not ValidateInRange(StrToDbl(PatternTimeout), "PatternTimeout", 0#, , , , 33) Then Validate = False
    Else
        Validate = False
    End If
    
    'validate timedomain
    If Not ValidateTimeDomain(WaitTimeDomain, "WaitTimeDomain", 34) Then Validate = False
    
End Function



Private Function Validate_Char(PatternString As String, PatThreading As Boolean, _
                          DriveLoPins As PinList, DriveHiPins As PinList, _
                          DriveZPins As PinList, DisablePins As PinList, FloatPins As PinList, _
                          Util1Pins As PinList, Util0Pins As PinList, _
                          PatternTimeout As String, WaitTimeDomain As String) As Boolean
    Dim Patterns As New Pattern
    Dim PatternStringArr() As String
    Dim Pat As Variant
    Dim PatArr() As String
    
    PatternStringArr = Split(PatternString, ",")
    
    Validate_Char = True ' Assume the best and override if trouble is found
    
    For Each Pat In PatternStringArr
        If Pat <> "" Then
           If InStr(CStr(Pat), ":") > 0 Then
              PatArr = Split(Pat, ":")
              Patterns.Value = PatArr(0)
           Else
              Patterns.Value = Pat
           End If
           If Not ValidatePatternThreading(Patterns, PatThreading, 1, True, 26) Then Validate_Char = False
              If Validate_Char Then Call PrLoadPattern(Patterns.Value)
        Else
        End If
    Next Pat
    
    ' Validate the pin state parameters.
    If Not ValidatePinStates(DriveLoPins, DriveHiPins, DriveZPins, DisablePins, _
                             FloatPins, Util1Pins, Util0Pins) Then Validate_Char = False
        
    If ValidateNumeric(PatternTimeout, "PatternTimeout", 33) Then
        ' Validate  0.0 <= PatternTimeout
        If Not ValidateInRange(StrToDbl(PatternTimeout), "PatternTimeout", 0#, , , , 33) Then Validate_Char = False
    Else
        Validate_Char = False
    End If
    
    'validate timedomain
    If Not ValidateTimeDomain(WaitTimeDomain, "WaitTimeDomain", 34) Then Validate_Char = False
    
End Function
Public Function freerunclk_set_XY(argc As Integer, argv() As String) As Long
    On Error GoTo errHandler
    ''argv(0) is shmoo axis
    ''argv(1) is FRC nWire port name
    ''argv(2) is FRC AC spec symbol
    
    '' 20151029 - Add tracking info for second nWire-FRC
    '' argv(3) is tracking FRC nWire port name >> Clock_Port1
    '' argv(4) is tracking FRC AC spec symbol >> XI0_Freq_1_S
    Dim site As Variant
    Dim suspended As Boolean
    Dim specval As Long
    Dim pointval As Double
    Dim pointss As Long
    Dim axis_type As tlDevCharShmooAxis
    
    
    '' 20151029 - If Shmoo XI0_Freq and tracking condition as XI0_Freq_1 that should be vary the XI0_Freq_1 by using nWire-FRC
    Dim TrackingStepName() As String
    Dim Trackingcount As Integer
    Dim TeackingPointVal() As Double
    Dim b_IsTracking As Boolean
    
    If argc > 3 Then
        b_IsTracking = True
    Else
        b_IsTracking = False
    End If
    
    Select Case UCase(argv(0))
        Case "X":
            axis_type = tlDevCharShmooAxis_X
            
        Case "Y":
            axis_type = tlDevCharShmooAxis_Y
            
        Case "Z":
            axis_type = tlDevCharShmooAxis_Z
            
        Case Else:
            axis_type = tlDevCharShmooAxis_Invalid
    End Select
    
    suspended = TheExec.Datalog.DatalogSuspended
    
    TheExec.Datalog.DatalogSuspended = False
    
    For Each site In TheExec.sites
        pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(axis_type).Value
        
        '' 20151029 - Get Tracking point value
        If b_IsTracking = True Then
            TrackingStepName = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(axis_type).TrackingParameters.List
            Trackingcount = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(axis_type).TrackingParameters.Count
            Dim i As Double
            ReDim TeackingPointVal(Trackingcount)
            For i = 1 To Trackingcount
                TeackingPointVal(i) = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(axis_type).TrackingParameters(TrackingStepName(i - 1)).Value
            Next i
        End If
    Next site
    
    Dim FRC_PortName As String
    Dim FRC_ACSpec As String
    FRC_PortName = argv(1)
    FRC_ACSpec = argv(2)
    
    Call VaryFreq(FRC_PortName, pointval, FRC_ACSpec)
    
    TheHdw.Wait 0.005
    
    Dim Track_FRC_PortName As String
    Dim Track_FRC_ACSpec As String
        If b_IsTracking = True Then
            Dim j As Double
            Dim k As Double
                k = 1
                For j = 3 To argc - 1 Step 2
                    Track_FRC_PortName = argv(j)
                    Track_FRC_ACSpec = argv(j + 1)
                        Call VaryFreq(Track_FRC_PortName, TeackingPointVal(k), Track_FRC_ACSpec)
                    TheHdw.Wait 0.005
                    k = k + 1
                Next j
            
''        If argc > 5 Then
''    Dim FRC_PortName_2 As String
''    Dim FRC_ACSpec_2 As String
''            FRC_PortName_2 = argv(5)
''            FRC_ACSpec_2 = argv(6)
''            Call VaryFreq(FRC_PortName_2, TeackingPointVal_2, FRC_ACSpec_2)
''
''            thehdw.Wait 0.005
''        End If
        End If
    
  
     
    TheExec.Datalog.DatalogSuspended = False


Exit Function
errHandler:
        TheExec.AddOutput "Error in the VBT Icc"
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function freerunclk_stop(argc As Integer, argv() As String) As Long
    ''argv(0) is FRC nWire port name
    ''argv(1) is tracking FRC nWire port name >> Clock_Port1
    
    Dim FRC_PortName As String
    Dim site As Variant
    
    '' 20151029 - Stop tracking nWireFRC
    Dim b_IsTracking As Boolean
    Dim Track_FRC_PortName() As String
    Dim FRC_PortName_1 As String
    Dim FRC_PortName_2 As String
    If argc > 1 Then
        b_IsTracking = True
    Else
        b_IsTracking = False
    End If
    
    FRC_PortName = argv(0)
'    FRC_PortName = Replace(FRC_PortName, "+", ",")
    
    '' 20151029 - Stop tracking nWireFRC
    If b_IsTracking = True Then
    Dim i As Double
    ReDim Track_FRC_PortName(argc - 1)
    For i = 1 To argc - 1
        Track_FRC_PortName(i) = argv(i)
    Next i
        
''        If argc > 2 Then
''        FRC_PortName_2 = argv(2)
''        End If
    End If
    
    For Each site In TheExec.sites.Active
        If TheHdw.Protocol.ports(FRC_PortName).Enabled = True Then
            TheHdw.Protocol.ports(FRC_PortName).Halt
           ' TheHdw.Protocol.Ports(FRC_PortName).Enabled = False   ' marked for shmoo XI0 at PA mode
        End If
        
        If b_IsTracking = True Then
        For i = 1 To argc - 1
            If TheHdw.Protocol.ports(Track_FRC_PortName(i)).Enabled = True Then
                TheHdw.Protocol.ports(Track_FRC_PortName(i)).Halt
            End If
        Next i
''             If argc > 2 Then
''             If thehdw.Protocol.ports(FRC_PortName_2).Enabled = True Then
''                thehdw.Protocol.ports(FRC_PortName_2).Halt
''             End If
''             End If
''
        End If
    Next site

End Function
Public Function CharStoreResultsUntilNextRun()
    On Error GoTo err1
    TheExec.DevChar.Configuration.Features.Item(tlDevCharFeature_StoreResultsUntilNextRun).Enabled = False
    m_STDSvcClient.SelfTest.MemoryCollectRunInterval = 1
    Exit Function
err1:
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function setup_patgen_counter(argc As Integer, argv() As String) As Long
    Dim x_pointval As Double, x_count As Double
    Dim y_pointval As Double, y_count As Double
    Dim site As Variant
    On Error GoTo err1
    
    For Each site In TheExec.sites
        x_pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
        y_pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
        x_count = x_pointval / ((65536 + 2) / 24000000)
        y_count = y_pointval / ((65536 + 2) / 24000000)
    Next site
    TheHdw.Digital.Patgen.counter(tlPgCounter10) = x_count  'boot ok
    TheHdw.Digital.Patgen.counter(tlPgCounter11) = y_count  'bist done
    Exit Function
err1:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function run_shmoo(shmoo_setup As String)
    On Error GoTo err1
    If TheExec.DevChar.Setups.IsRunning = True Then Exit Function
        With TheExec.DevChar.Setups(shmoo_setup)
            .SaveState ("current")
            .Execute False
            .RestoreState ("current")
        End With
    Exit Function
err1:
    If AbortTest Then Exit Function Else Resume Next
End Function
' ==============
' Public Methods
' ==============
' Perform a digital functional test.
' Return TL_SUCCESS if the test executes without problems, else TL_ERROR.
'201612 Add DigSrc Arguments
Public Function Functional_T_char(StartOfBodyF As InterposeName, _
                             PrePatF As InterposeName, PreTestF As InterposeName, _
                             PostTestF As InterposeName, PostPatF As InterposeName, EndOfBodyF As InterposeName, _
                             ReportResult As PFType, ResultMode As tlResultMode, DriveLoPins As PinList, DriveHiPins As PinList, _
                             DriveZPins As PinList, DisablePins As PinList, FloatPins As PinList, StartOfBodyFArgs As String, _
                             PrePatFArgs As String, PreTestFArgs As String, PostTestFArgs As String, _
                             PostPatFArgs As String, EndOfBodyFArgs As String, Util1Pins As PinList, _
                             Util0Pins As PinList, PatFlagF As InterposeName, _
                             PatFlagFArgs As String, Validating_ As Boolean, _
                             Optional PatternTimeout As String = "30", Optional Step_ As SubType, _
                             Optional WaitTimeDomain As String, _
                             Optional ConcurrentMode As tlPatConcurrentMode = tlPatConcurrentModeCached, _
                             Optional Interpose_PrePat As String, Optional Init_Patt1 As Pattern, Optional Init_Patt2 As Pattern, Optional Init_Patt3 As Pattern, _
                             Optional Init_Patt4 As Pattern, Optional Init_Patt5 As Pattern, _
                             Optional Init_Patt6 As Pattern, Optional Init_Patt7 As Pattern, Optional Init_Patt8 As Pattern, _
                             Optional Init_Patt9 As Pattern, Optional Init_Patt10 As Pattern, _
                             Optional PayLoad_Patt1 As Pattern, _
                             Optional PayLoad_Patt2 As Pattern, _
                             Optional PayLoad_Patt3 As Pattern, _
                             Optional PayLoad_Patt4 As Pattern, _
                             Optional PayLoad_Patt5 As Pattern, _
                             Optional Power_Run_Scenario As String, Optional Wait As String, _
                             Optional digsrc_BitSize As String, Optional digsrc_Seg As String, Optional digsrc_DigSrcPin As String, Optional digSrc_EQ As String, _
                             Optional BlockType As String, Optional SELSRAM_DSSC As String, Optional Pmode As String, _
                             Optional Vbump As Boolean = False) As Long
' EDITFORMAT1 1,,Pattern,,,Patterns|7,,InterposeName,Interpose Functions,,StartOfBodyF|9,,InterposeName,,,PrePatF|11,,InterposeName,,,PreTestF|13,,InterposeName,,,PostTestF|15,,InterposeName,,,PostPatF|17,,InterposeName,,,EndOfBodyF|2,,PFType,,,ReportResult|6,,tlResultMode,,,ResultMode|19,,pinlist,Pin States,,DriveLoPins|20,,pinlist,,,DriveHiPins|21,,pinlist,,,DriveZPins|22,,pinlist,,,DisablePins|23,,pinlist,,,FloatPins|8,,String,,,StartOfBodyFArgs|10,,String,,,PrePatFArgs|12,,String,,,PreTestFArgs|14,,String,,,PostTestFArgs|16,,String,,,PostPatFArgs|18,,String,,,EndOfBodyFArgs|24,,pinlist,,,Util1Pins|25,,pinlist,,,Util0Pins|31,,InterposeName,,,PatFlagF|32,,String,,,PatFlagFArgs|5,,tlRelayMode,,,RelayMode|3,,Boolean,,,PatThreading|30,,Boolean,,,MatchAllSites|26,,tlWaitVal,Flag Match,,WaitFlagA|27,,tlWaitVal,,,WaitFlagB|28,,tlWaitVal,,,WaitFlagC|29,,tlWaitVal,,,WaitFlagD|0,,Boolean,,,Validating_|4,,String,,0 <= PatternTimeout,PatternTimeout|6,,tlPatStartConcurrentMode,,,ConcurrentMode

''==============================================================================================
''---------- 20171020 for releasing more argument ----------
Dim RelayMode As tlRelayMode
Dim PatThreading As Boolean
Dim MatchAllSites As Boolean
Dim WaitFlagA As tlWaitVal
Dim WaitFlagB As tlWaitVal
Dim WaitFlagC As tlWaitVal
Dim WaitFlagD As tlWaitVal

RelayMode = tlPowered
WaitFlagA = waitoff
WaitFlagB = waitoff
WaitFlagC = waitoff
WaitFlagD = waitoff
''==============================================================================================

    Interpose_PrePat_GLB = Interpose_PrePat
    Dim Test_Pattern As String
    Functional_T_char = TL_SUCCESS   ' be optimistic
    If Not TheExec.Flow.IsRunning Then Exit Function
    
'    On Error GoTo errHandler
    
    ' Cache parameters for PostTest
    m_EndOfBodyF = EndOfBodyF
    m_EndOfBodyFArgs = EndOfBodyFArgs
    
    ' Apply default values to parameters whose values were not specified.
    ApplyDefaults PatternTimeout
    
   
    If Validating_ Then
    
        Dim PatString As String: PatString = ""
        If Init_Patt1 <> "" Then PatString = Init_Patt1.Value
        If Init_Patt2 <> "" Then PatString = PatString & "," & Init_Patt2.Value
        If Init_Patt3 <> "" Then PatString = PatString & "," & Init_Patt3.Value
        If Init_Patt4 <> "" Then PatString = PatString & "," & Init_Patt4.Value
        If Init_Patt5 <> "" Then PatString = PatString & "," & Init_Patt5.Value
        If Init_Patt6 <> "" Then PatString = PatString & "," & Init_Patt6.Value
        If Init_Patt7 <> "" Then PatString = PatString & "," & Init_Patt7.Value
        If Init_Patt8 <> "" Then PatString = PatString & "," & Init_Patt8.Value
        If Init_Patt9 <> "" Then PatString = PatString & "," & Init_Patt9.Value
        If Init_Patt10 <> "" Then PatString = PatString & "," & Init_Patt10.Value
        
        If PayLoad_Patt1 <> "" Then PatString = PatString & "," & PayLoad_Patt1.Value
        If PayLoad_Patt2 <> "" Then PatString = PatString & "," & PayLoad_Patt2.Value
        If PayLoad_Patt3 <> "" Then PatString = PatString & "," & PayLoad_Patt3.Value
        If PayLoad_Patt4 <> "" Then PatString = PatString & "," & PayLoad_Patt4.Value
        If PayLoad_Patt5 <> "" Then PatString = PatString & "," & PayLoad_Patt5.Value
               
        If Not Validate_Char(PatString, PatThreading, DriveLoPins, DriveHiPins, DriveZPins, DisablePins, _
            FloatPins, Util1Pins, Util0Pins, PatternTimeout, WaitTimeDomain) Then Functional_T_char = TL_ERROR
        Exit Function
    End If
    
    If Step_ = subAllBody Or Step_ = subPrebody Or _
       m_InterposeFunctionsSet = False Then

        ' Register certain interpose function names with flow controller
        Call tl_SetInterpose(TL_C_PREPATF, PrePatF.Value, PrePatFArgs, _
                             TL_C_POSTPATF, PostPatF.Value, PostPatFArgs, _
                             TL_C_PRETESTF, PreTestF.Value, PreTestFArgs, _
                             TL_C_POSTTESTF, PostTestF.Value, PostTestFArgs, _
                             TL_C_FLAGMATCHF, PatFlagF.Value, PatFlagFArgs, _
                             TL_C_POSTPATBPF, "PostTest", "")

        m_InterposeFunctionsSet = True

    End If

    ' PreBody
    If Step_ = subAllBody Or Step_ = subPrebody Then
        FetchContext
            
        ' Set up the test
        
    
    Call PreBody(DriveHiPins, DriveLoPins, DriveZPins, DisablePins, Util1Pins, Util0Pins, _
                 WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD, MatchAllSites, _
                 PatThreading, RelayMode, WaitTimeDomain, Interpose_PrePat)
    
    
    g_Vbump_function = False
    g_VminBoundary_selsrm = False
    
    If Vbump = True Then
      '===========================DC_LEVEL Powers Stored===============================
        Shmoo_Save_core_power_per_site_for_Vbump
        g_VminBoundary_selsrm = True
        g_FirstSetp = True
        g_Print_SELSRM_Def = True
        g_InitSeq = ""
        g_Vbump_function = True
        g_shmoo_ret = False
       '===========================DC_LEVEL Powers Stored===============================
    End If
    
    End If ' PreBody
    
    Vbump_for_Interpose = False
    If Vbump = True Then '' for printshmooinfo outputstring printing
        Vbump_for_Interpose = True
    Else
        Vbump_for_Interpose = False
    End If
    
    Dim CurConcurrentContext As Long
    CurConcurrentContext = m_STDSvcClient.FlowDomainService.ConcurrentContext
    
    ' Body
    If Step_ = subAllBody Or Step_ = subBody Then
        
        ' cache member variables
        ' there are statements below which can cause us to jump to the next subflow if we're running with concurrent test.
        ' if the next test in the next subflow runs this function then it will overwrite the below member variables, such
        ' that when we get back to this call they will have different values.  so we cache the values here and then
        ' restore them right after the code that can cause us to jump to the next subflow.  then later on in
        ' postbody and posttest when they're used they'll have the proper values.
        
        Dim tempendofbody As String
        Dim tempendofbodyfargs As String
        Dim tempdrivepins As String
        Dim tempfloatpins As String
              
        If CurConcurrentContext Then
            tempendofbody = m_EndOfBodyF
            tempendofbodyfargs = m_EndOfBodyFArgs
            tempdrivepins = m_DrivePins
            tempfloatpins = m_FloatPins
        End If
                
'        Body_patt.Value = Shmoo_Pattern
        ' Perform the test
        Call Interpose(StartOfBodyF, StartOfBodyFArgs)
        
        'job char flag
        If Vbump = True Then
            '=======================================Dyanmic DSSC Source bits =================================
'            Dim dyanmicDSSCbits As String
            If g_FirstSetp = True Then
               g_dyanmicDSSCbits = ""
               If SELSRAM_DSSC <> "" And BlockType <> "" Then
                  If InStr(SELSRAM_DSSC, "'") > 0 Then SELSRAM_DSSC = Replace(SELSRAM_DSSC, "'", "")
                  If InStr(UCase(SELSRAM_DSSC), "SELSRM") > 0 Then SELSRAM_DSSC = Replace(UCase(SELSRAM_DSSC), "SELSRM", "")
                  g_dyanmicDSSCbits = dynamic_SELSRM_source_bits(SELSRAM_DSSC, BlockType)
               Else
                  g_dyanmicDSSCbits = ""
               End If
            End If
            '=======================================Dyanmic DSSC Source bits =================================
            '=====================Pattern Decompose for bring up shmoo=======================
            If TheExec.EnableWord("BringUp_Shmoo") = True Then
               DecomposePattSet Init_Patt1, Init_Patt2, Init_Patt3, Init_Patt4, Init_Patt5, Init_Patt6, Init_Patt7, Init_Patt8, Init_Patt9, Init_Patt10, PayLoad_Patt1, PayLoad_Patt2, PayLoad_Patt3, PayLoad_Patt4, PayLoad_Patt5
            End If
            '=====================Pattern Decompose for bring up shmoo=======================
            '===================================Pmode transfer to Force condition ============================
            If g_FirstSetp = True Then
            Dim Pmode_Voltage As String::   Pmode_Voltage = ""
               If Pmode <> "" Then
                  g_CharInputString_Voltage_Dict.RemoveAll
                  Decide_Pmode_ForceVoltage Pmode, "CorePower", Pmode_Voltage
                  PL_DC_conditions_GLB = Pmode_Voltage
                  Call SetForceCondition(Pmode_Voltage & ";STOREPREPAT")
                  
                  
               End If
               If Interpose_PrePat <> "" Then
                  Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
               End If
               If Pmode_Voltage <> "" Then
                  Interpose_PrePat = Pmode_Voltage
               End If
            End If
            '===================================Pmode transfer to Force condition ============================
             
        Else ' without Vbump selsram function
            
                        CHAR_USL_HVCC = 9999
                        CHAR_USL_LVCC = 9999
                        CHAR_LSL_HVCC = 9999
                        CHAR_LSL_LVCC = 9999
                
            If Interpose_PrePat <> "" Then
               Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
            End If
        End If
        
        
        Call Body(FloatPins, StrToDbl(PatternTimeout), PayLoad_Patt1, ReportResult, ResultMode)
                
       
        
        ' Run the pattern.  Perform functional test.
        Dim site As Variant
        Dim i As Long
        
       
        If TheExec.sites.ActiveCount > 0 Then

        Dim pin_count As Long
        Dim Pin_Ary() As String
        Dim VDD_Force As String
        Dim AddPinToDict As Boolean
        
            On Error Resume Next
            If Vbump = True Then
                If g_FirstSetp = True Then
                   Getforcecondition_VDD g_ForceCond_VDD, Interpose_PrePat
                End If
            Else ' without selsram function
                Decide_shmoo_patt Init_Patt1, Init_Patt2, Init_Patt3, Init_Patt4, Init_Patt5, Init_Patt6, Init_Patt7, Init_Patt8, Init_Patt9, Init_Patt10, PayLoad_Patt1, PayLoad_Patt2, PayLoad_Patt3, PayLoad_Patt4, PayLoad_Patt5
                Getforcecondition_VDD g_ForceCond_VDD, Interpose_PrePat
            End If
                        
            
            
            Get_Shmoo_Set_Pin Shmoo_Apply_Pin, g_ForceCond_VDD, pin_count
            
           If Vbump = False Then
              For Each site In TheExec.sites
                  For i = 0 To pin_count - 1
                      ShmooSweepPower(i) = 0
                  Next i
                  Shmoo_Save_core_power_per_site Shmoo_Apply_Pin, ShmooSweepPower               'read h/w power setup to array
              Next site
           End If
            
            Power_Level_Last = "" 'Right(theexec.DataManager.InstanceName, 2)
            
            Power_Level_Vmode_Last = "" 'SelSram parameters

            
            Dim wait_time_ary() As String

            If InStr(Wait, ",") > 0 Then
               Wait = Replace(Wait, "'", "")
               wait_time_ary = Split(Wait, ",")
            Else
               ReDim wait_time_ary(14) As String
            End If
                
        '///////////////////Multi SrcCode Initialize///////////////////
                Dim digsrc_BitSize_arr() As String
                Dim digsrc_Seg_arr() As String
                Dim digsrc_DigSrcPin_arr() As String
                Dim digSrc_EQ_arr() As String
                
                'digsrc_BitSize
                If InStr(digsrc_BitSize, ",") > 0 Then
                    digsrc_BitSize = Replace(digsrc_BitSize, "'", "") 'SelSram parameters
                    digsrc_BitSize_arr = Split(digsrc_BitSize, ",")
                Else
                    ReDim digsrc_BitSize_arr(14) As String
                End If
                'digsrc_Seg
                If InStr(digsrc_Seg, ",") > 0 Then
                    digsrc_Seg = Replace(digsrc_Seg, "'", "") 'SelSram parameters
                    digsrc_Seg_arr = Split(digsrc_Seg, ",")
                Else
                    ReDim digsrc_Seg_arr(14) As String
                End If
                'digsrc_DigSrcPin
                If InStr(digsrc_DigSrcPin, ",") > 0 Then
                    digsrc_DigSrcPin = Replace(digsrc_DigSrcPin, "'", "") 'SelSram parameters
                    digsrc_DigSrcPin_arr = Split(digsrc_DigSrcPin, ",")
                Else
                    ReDim digsrc_DigSrcPin_arr(14) As String
                End If
                'digSrc_EQ
                If InStr(digSrc_EQ, ",") > 0 Then
                    digSrc_EQ = Replace(digSrc_EQ, "'", "") 'SelSram parameters
                    digSrc_EQ_arr = Split(digSrc_EQ, ",")
                Else
                    ReDim digSrc_EQ_arr(14) As String
                End If
                
'==============================================================================
'
Dim RTOSRelaySwith As Boolean
Dim all_powerpins As PinList
Dim DecideSPIMatchLoopFlag As Boolean
Dim SPIMatchLoopCountValue As Long
If LCase(TheExec.DataManager.instanceName) Like "*rtos*" Then
    RTOSRelaySwith = True ' True--> Skye RTOS method(change on the fly)
Else
    RTOSRelaySwith = False ' False---> normal RTOS method
End If
'==============================================================================
                
        '///////////////////////////////////////////////////
            TheExec.Datalog.WriteComment Power_Run_Scenario
            
            If Vbump = True Then

                 If TheExec.EnableWord("Shmoo_TTR") = True Then
                    If Not g_FirstSetp = True And g_InitSeq = "1" Then GoTo Init1
                    If Not g_FirstSetp = True And g_InitSeq = "2" Then GoTo Init2
                    If Not g_FirstSetp = True And g_InitSeq = "3" Then GoTo Init3
                    If Not g_FirstSetp = True And g_InitSeq = "4" Then GoTo Init4
                    If Not g_FirstSetp = True And g_InitSeq = "5" Then GoTo Init5
                    If Not g_FirstSetp = True And g_InitSeq = "6" Then GoTo Init6
                    If Not g_FirstSetp = True And g_InitSeq = "7" Then GoTo Init7
                    If Not g_FirstSetp = True And g_InitSeq = "8" Then GoTo Init8
                    If Not g_FirstSetp = True And g_InitSeq = "9" Then GoTo Init9
                    If Not g_FirstSetp = True And g_InitSeq = "10" Then GoTo Init10
                    If Not g_FirstSetp = True And g_InitSeq = "Payload1" Then GoTo Payload1
                End If
Init1:
                Shmoo_Test_Pattern Init_Patt1, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 1, wait_time_ary(0), digsrc_BitSize_arr(0), digsrc_Seg_arr(0), digsrc_DigSrcPin_arr(0), digSrc_EQ_arr(0), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
Init2:
                Shmoo_Test_Pattern Init_Patt2, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 2, wait_time_ary(1), digsrc_BitSize_arr(1), digsrc_Seg_arr(1), digsrc_DigSrcPin_arr(1), digSrc_EQ_arr(1), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
Init3:
                Shmoo_Test_Pattern Init_Patt3, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 3, wait_time_ary(2), digsrc_BitSize_arr(2), digsrc_Seg_arr(2), digsrc_DigSrcPin_arr(2), digSrc_EQ_arr(2), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
Init4:
                Shmoo_Test_Pattern Init_Patt4, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 4, wait_time_ary(3), digsrc_BitSize_arr(3), digsrc_Seg_arr(3), digsrc_DigSrcPin_arr(3), digSrc_EQ_arr(3), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
Init5:
                Shmoo_Test_Pattern Init_Patt5, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 5, wait_time_ary(4), digsrc_BitSize_arr(4), digsrc_Seg_arr(4), digsrc_DigSrcPin_arr(4), digSrc_EQ_arr(4), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
Init6:
                Shmoo_Test_Pattern Init_Patt6, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 6, wait_time_ary(5), digsrc_BitSize_arr(5), digsrc_Seg_arr(5), digsrc_DigSrcPin_arr(5), digSrc_EQ_arr(5), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
Init7:
                Shmoo_Test_Pattern Init_Patt7, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 7, wait_time_ary(6), digsrc_BitSize_arr(6), digsrc_Seg_arr(6), digsrc_DigSrcPin_arr(6), digSrc_EQ_arr(6), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
Init8:
                Shmoo_Test_Pattern Init_Patt8, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 8, wait_time_ary(7), digsrc_BitSize_arr(7), digsrc_Seg_arr(7), digsrc_DigSrcPin_arr(7), digSrc_EQ_arr(7), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
Init9:
                Shmoo_Test_Pattern Init_Patt9, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 9, wait_time_ary(8), digsrc_BitSize_arr(8), digsrc_Seg_arr(8), digsrc_DigSrcPin_arr(8), digSrc_EQ_arr(8), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
Init10:
                Shmoo_Test_Pattern Init_Patt10, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 10, wait_time_ary(9), digsrc_BitSize_arr(9), digsrc_Seg_arr(9), digsrc_DigSrcPin_arr(9), digSrc_EQ_arr(9), , , , , , , BlockType, g_dyanmicDSSCbits, Vbump
                
                If g_ForceCond_VDD <> "" Then Power_Level_Last = ""
                   TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
                   TheHdw.Wait 0.0001
Payload1:
                Shmoo_Test_Pattern PayLoad_Patt1, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 1, wait_time_ary(10), digsrc_BitSize_arr(10), digsrc_Seg_arr(10), digsrc_DigSrcPin_arr(10), digSrc_EQ_arr(10), RTOSRelaySwith, , , , , 3, BlockType, g_dyanmicDSSCbits, Vbump
                
                Shmoo_Test_Pattern PayLoad_Patt2, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 2, wait_time_ary(11), digsrc_BitSize_arr(11), digsrc_Seg_arr(11), digsrc_DigSrcPin_arr(11), digSrc_EQ_arr(11), RTOSRelaySwith, , , , , 3, BlockType, g_dyanmicDSSCbits, Vbump
                
                Shmoo_Test_Pattern PayLoad_Patt3, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 3, wait_time_ary(12), digsrc_BitSize_arr(12), digsrc_Seg_arr(12), digsrc_DigSrcPin_arr(12), digSrc_EQ_arr(12), RTOSRelaySwith, , , , , 3, BlockType, g_dyanmicDSSCbits, Vbump
                
                Shmoo_Test_Pattern PayLoad_Patt4, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 4, wait_time_ary(13), digsrc_BitSize_arr(13), digsrc_Seg_arr(13), digsrc_DigSrcPin_arr(13), digSrc_EQ_arr(13), RTOSRelaySwith, , , , , 3, BlockType, g_dyanmicDSSCbits, Vbump
                
                Shmoo_Test_Pattern PayLoad_Patt5, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 5, wait_time_ary(14), digsrc_BitSize_arr(14), digsrc_Seg_arr(14), digsrc_DigSrcPin_arr(14), digSrc_EQ_arr(14), RTOSRelaySwith, , , , , 3, BlockType, g_dyanmicDSSCbits, Vbump
                 
            Else ' without Vump/SELSRM function
                Shmoo_Test_Pattern Init_Patt1, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 1, wait_time_ary(0), digsrc_BitSize_arr(0), digsrc_Seg_arr(0), digsrc_DigSrcPin_arr(0), digSrc_EQ_arr(0)
                Shmoo_Test_Pattern Init_Patt2, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 2, wait_time_ary(1), digsrc_BitSize_arr(1), digsrc_Seg_arr(1), digsrc_DigSrcPin_arr(1), digSrc_EQ_arr(1)
                Shmoo_Test_Pattern Init_Patt3, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 3, wait_time_ary(2), digsrc_BitSize_arr(2), digsrc_Seg_arr(2), digsrc_DigSrcPin_arr(2), digSrc_EQ_arr(2)
                Shmoo_Test_Pattern Init_Patt4, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 4, wait_time_ary(3), digsrc_BitSize_arr(3), digsrc_Seg_arr(3), digsrc_DigSrcPin_arr(3), digSrc_EQ_arr(3)
                Shmoo_Test_Pattern Init_Patt5, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 5, wait_time_ary(4), digsrc_BitSize_arr(4), digsrc_Seg_arr(4), digsrc_DigSrcPin_arr(4), digSrc_EQ_arr(4)
                Shmoo_Test_Pattern Init_Patt6, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 6, wait_time_ary(5), digsrc_BitSize_arr(5), digsrc_Seg_arr(5), digsrc_DigSrcPin_arr(5), digSrc_EQ_arr(5)
                Shmoo_Test_Pattern Init_Patt7, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 7, wait_time_ary(6), digsrc_BitSize_arr(6), digsrc_Seg_arr(6), digsrc_DigSrcPin_arr(6), digSrc_EQ_arr(6)
                Shmoo_Test_Pattern Init_Patt8, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 8, wait_time_ary(7), digsrc_BitSize_arr(7), digsrc_Seg_arr(7), digsrc_DigSrcPin_arr(7), digSrc_EQ_arr(7)
                Shmoo_Test_Pattern Init_Patt9, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 9, wait_time_ary(8), digsrc_BitSize_arr(8), digsrc_Seg_arr(8), digsrc_DigSrcPin_arr(8), digSrc_EQ_arr(8)
                Shmoo_Test_Pattern Init_Patt10, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, True, 10, wait_time_ary(9), digsrc_BitSize_arr(9), digsrc_Seg_arr(9), digsrc_DigSrcPin_arr(9), digSrc_EQ_arr(9)
                

                Shmoo_Test_Pattern PayLoad_Patt1, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 1, wait_time_ary(10), digsrc_BitSize_arr(10), digsrc_Seg_arr(10), digsrc_DigSrcPin_arr(10), digSrc_EQ_arr(10), RTOSRelaySwith, , , , , 3
                Shmoo_Test_Pattern PayLoad_Patt2, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 2, wait_time_ary(11), digsrc_BitSize_arr(11), digsrc_Seg_arr(11), digsrc_DigSrcPin_arr(11), digSrc_EQ_arr(11), RTOSRelaySwith, , , , , 3
                Shmoo_Test_Pattern PayLoad_Patt3, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 3, wait_time_ary(12), digsrc_BitSize_arr(12), digsrc_Seg_arr(12), digsrc_DigSrcPin_arr(12), digSrc_EQ_arr(12), RTOSRelaySwith, , , , , 3
                Shmoo_Test_Pattern PayLoad_Patt4, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 4, wait_time_ary(13), digsrc_BitSize_arr(13), digsrc_Seg_arr(13), digsrc_DigSrcPin_arr(13), digSrc_EQ_arr(13), RTOSRelaySwith, , , , , 3
                Shmoo_Test_Pattern PayLoad_Patt5, ReportResult, CLng(TL_C_YES), ConcurrentMode, Power_Run_Scenario, Shmoo_Apply_Pin, False, 5, wait_time_ary(14), digsrc_BitSize_arr(14), digsrc_Seg_arr(14), digsrc_DigSrcPin_arr(14), digSrc_EQ_arr(14), RTOSRelaySwith, , , , , 3
            End If
            
            If TheExec.DevChar.Setups.IsRunning = False And CharSetName_GLB <> "" Then
                Dim p As Variant, p_ary() As String, p_cnt As Long, ApplyPins As String, Setup_mode As String
                If TheExec.DevChar.Setups(CharSetName_GLB).TestMethod.Value = tlDevCharTestMethod_Reburst Then TheExec.Datalog.WriteComment "[PrintCharCondition:" & PrintCharSetup(Interpose_PrePat_GLB) & ",Test]"
                Setup_mode = TheExec.DevChar.Setups(CharSetName_GLB).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Name
                If (LCase(Setup_mode) <> "vid" And LCase(Setup_mode) <> "vicm") Then
                    ApplyPins = TheExec.DevChar.Setups(CharSetName_GLB).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins
                    TheExec.DataManager.DecomposePinList ApplyPins, p_ary, p_cnt
                    For Each p In p_ary
                        TheExec.DevChar.Setups(CharSetName_GLB).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = p
                        run_shmoo CharSetName_GLB
                    Next p
                    TheExec.DevChar.Setups(CharSetName_GLB).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = ApplyPins
                Else
                    run_shmoo CharSetName_GLB
                End If
            End If


            If err.number <> 0 Then
                If err.number = TL_E_AT_PATSET_BREAKPT Then
                    Exit Function
                Else
                    GoTo errHandler
                End If
            End If
            On Error GoTo errHandler
        End If
                
        ' restore the member variables for posttest
        If CurConcurrentContext Then
            m_EndOfBodyF = tempendofbody
            m_EndOfBodyFArgs = tempendofbodyfargs
        End If
    
        ' Calls End of Body Interpose Function, anything from here to the end of the Body
        ' should be added to PostTest()
        Dim argv() As String
        PostTest 0, argv

        '20170213 prevent over write shmoo pattern
        If Vbump = True Then
            Decide_shmoo_patt Init_Patt1, Init_Patt2, Init_Patt3, Init_Patt4, Init_Patt5, Init_Patt6, Init_Patt7, Init_Patt8, Init_Patt9, Init_Patt10, PayLoad_Patt1, PayLoad_Patt2, PayLoad_Patt3, PayLoad_Patt4, PayLoad_Patt5
        End If
        DebugPrintFunc Shmoo_Pattern
        
        If Vbump = False Then
           For Each site In TheExec.sites
               Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "Instance Level"
           Next site
        End If
        ' restore the member variables for postbody (do this here instead of a couple of lines above since posttest could
        ' possibly jump to the next subflow in a concurrent test and cause the below memeber variables to change again.
        If CurConcurrentContext Then
            m_DrivePins = tempdrivepins
            m_FloatPins = tempfloatpins
        End If
        
    End If ' Body
    
    ' PostBody
    If Step_ = subAllBody Or Step_ = subPostbody Then
        ReadHWPowerValue_GLB = PrintCharSetup(Interpose_PrePat)
        
        If Vbump = True Then 'Switch voltage to VRS(Safe voltage) and change to Vmain
           g_Vbump_function = False
           Shmoo_Restore_Power_per_site_Vbump_NV True
           TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
           TheHdw.Wait 0.001
        Else ' without Vbump selsram function
            If Interpose_PrePat <> "" Then
                Call SetForceCondition("RESTOREPREPAT")
            End If
        End If
        
        Call PostBody(m_DrivePins, m_FloatPins, WaitTimeDomain, WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD)
    End If ' PostBody
    
    ' There shouldn't be any code below this line. Any other necessary
    ' code should be added to the PostTest method to support pattern set
    ' breakpoints.
    
    '' 20190327 only 3D Shmoo result record\
    If TheExec.DevChar.Setups.IsRunning = True Then
        If TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).startTime Like "1/1/0001*" Or TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).startTime Like "0001/1/1*" Then Exit Function  ' initial run of shmoo, not the first point
        If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = False And (TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Count = 3 Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Count = 2) Then
            StoreEachPoint
        End If
    End If
    
    Exit Function
    
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instanceName)
    Call TheExec.ErrorReport
    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF)
    m_InterposeFunctionsSet = False

    Functional_T_char = TL_ERROR
                If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub ApplyDefaults(ByRef PatternTimeout As String)
    ' If the worksheet doesn't have a value then apply 30 as the default.
    If Not NonBlank(PatternTimeout) Then
        PatternTimeout = "30"
    End If
End Sub

'removed by SY, no one use this function and this will causing compiler fail for non-AP project since different efuse structure

''Public Function Read_Waferdata_char()
''
''    Dim site As Variant
''    If UCase(TheExec.CurrentJob) Like "*FT*" Then
''        For Each site In TheExec.sites
''            HramLotId(site) = ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(site) + "-" + ECIDFuse.Category(ECIDIndex("Wafer_ID")).Read.ValStr(site)
''            'HramWaferId(Site) = TheExec.Datalog.Setup.WaferSetup.ID
''            XCoord(site) = ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.ValStr(site)
''            YCoord(site) = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.ValStr(site)
''        Next site
''    ElseIf UCase(TheExec.CurrentJob) Like "*CP*" Then
''        For Each site In TheExec.sites
''            HramLotId(site) = TheExec.Datalog.setup.LotSetup.LotID
''            HramWaferId(site) = TheExec.Datalog.setup.WaferSetup.ID
''            'HramLotId(Site) = HramLotId(Site) & "-" & HramWaferId(Site)
''            XCoord(site) = TheExec.Datalog.setup.WaferSetup.GetXCoord(site)
''            YCoord(site) = TheExec.Datalog.setup.WaferSetup.GetYCoord(site)
''
''            TheExec.Datalog.WriteComment "[XY_Coordinate_Read,Site:" & site & ",X:" & XCoord(site) & ",Y:" & YCoord(site) & ",LotId:" & HramLotId(site) & "]"
''            'TheExec.AddOutput "[XY_Coordination_Read,Site:" & site & ",X:" & XCoord(site) & ",Y:" & YCoord(site) & ",LotId:" & HramLotId(site) & "]"
''        Next site
''    End If
''
''
''
''End Function

Public Function PrintShmooInfo(argc As Long, argv() As String)
    Dim SetupName As String
    Dim method As String

    '20180118 Refresh shmoo overlay count
    If TheExec.Overlays.Count > 10000 Then
        TheExec.Overlays.RemoveAll
    End If

    SetupName = TheExec.DevChar.Setups.ActiveSetupName
    With TheExec.DevChar.Setups(SetupName)
        If .Shmoo.Axes.Count > 1 Then
            Call ShmooPostStep2Dto1D(argc, argv)
            Call ShmooPostStep2D(argc, argv)
        Else
            TheExec.Datalog.WriteComment "[Start_Shmoo]"
            Call ShmooPostStep1D(argc, argv)
            '20170120 Ignore HW after Shmoo
'           If (.TestMethod.Value = tlDevCharTestMethod_Retest) Then
'                TheExec.Datalog.WriteComment "[PrintCharCondition:" & PrintCharSetup(Interpose_PrePat_GLB) & "]"
'            End If
            TheExec.Datalog.WriteComment "[End_Shmoo]"
        End If
        If TheExec.DevChar.Setups.Item(SetupName).output.SuspendDatalog = False And .Shmoo.Axes.Count = 2 Then
            StoreEachPointResult_2D
            TheExec.Datalog.WriteComment "<<<<< 2D Shmoo Print Start >>>>>"
            Call Print2DShmooInfo(argc, argv)
            TheExec.Datalog.WriteComment "<<<<< 2D Shmoo Print Stop >>>>>"
            Count_Point = 0
        End If
        If TheExec.DevChar.Setups.Item(SetupName).output.SuspendDatalog = False And .Shmoo.Axes.Count = 3 Then
            StoreEachPointResult
            TheExec.Datalog.WriteComment "<<<<< 3D Shmoo Print Start >>>>>"
            Call Print3DShmooInfo(argc, argv)
            TheExec.Datalog.WriteComment "<<<<< 3D Shmoo Print Stop >>>>>"
            Count_Point = 0
        End If
        ''''' 20180710 Initialize GLlobal power condition
        ReadHWPowerValue_GLB = ""
        Charz_Force_Power_condition = ""
        
        ''''' 20180710 Add initialize value ''''''''''''
        CHAR_USL_HVCC = 9999
        CHAR_USL_LVCC = 9999
        CHAR_LSL_HVCC = 9999
        CHAR_LSL_LVCC = 9999
        
        
    End With

    Dim AcCat As String
    Dim site As Variant
    Dim SetSite As Integer
    AcCat = TheExec.Contexts.ActiveSelection.ACCategory

'''''''''''''''''Obsolete due to Support multiple nWire port 20170503'''''''''''''
'    If XI0_GP <> "" Then
'        Call VaryFreq("XI0_Port", TheExec.Specs.AC("XI0_Freq_VAR").ContextValue, "XI0_Freq_VAR")
'    ElseIf XI0_Diff_GP <> "" Then
'        Call VaryFreq("XI0_Diff_Port", TheExec.Specs.AC("XI0_Diff_Freq_VAR").ContextValue, "XI0_Diff_Freq_VAR")
'    End If
'    If RTCLK_GP <> "" Then
'        Call VaryFreq("RT_CLK32768_Port", TheExec.Specs.AC("RT_CLK32768_Freq_VAR").ContextValue, "RT_CLK32768_Freq_VAR")
'    ElseIf RTCLK_Diff_GP <> "" Then
'        Call VaryFreq("RT_CLK32768_Diff_Port", TheExec.Specs.AC("RT_CLK32768_Diff_Freq_VAR").ContextValue, "RT_CLK32768_Diff_Freq_VAR")
'    End If
'

    'For turn off TName Sweep point
    gl_flag_end_shmoo = True
    gl_flag_CZ_Nominal_Measured_1st_Point = False
    
    If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = False Then    '20180718 add
        Call TheExec.sites(site).IncrementTestNumber
    End If
    
  '''''''''''''''''Support multiple nWire port 20170503'''''''''''''
    Dim nWire_port_ary() As String
    Dim nwp As Variant ', all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String

    nWire_port_ary = Split(nWire_Ports_GLB, ",")
'    nWire_port_ary = Split("XO0_Port,RT_CLK32768_Port", ",")
    ' Convert nWire_ports to all_ports and all_pins

    With TheExec.DevChar.Setups(SetupName)
        If .Shmoo.Axes.Count > 1 Then
            If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value = "AC Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Type.Value = "AC Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Type.Value = "Global Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value = "Global Spec" Then
        '    If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value = "AC Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value = "Global Spec" Then
                For Each nwp In nWire_port_ary
                    Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
                    Call VaryFreq(port_pa, TheExec.specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
                Next nwp
            End If
        Else
            For Each nwp In nWire_port_ary
                Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
                Call VaryFreq(port_pa, TheExec.specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
            Next nwp
        End If
    End With
    Shmoo_End = True
End Function

Public Function Flow_Shmoo_Setup()
    Dim DevChar_Setup As String
    Dim shmoo_axis As Variant, Shmoo_Tracking_Item As Variant
    Dim axis_name As Variant, axis_type As String, Tracking_num As Long
    Dim i As Long, Shmoo_Spec As String, Shmoo_StepSize As Double, shmoo_step As Long
    Dim StepSize As Double
    Dim arg_ary() As String
    Dim site As Variant
    Shmoo_setup_str = ""
    Flow_Shmoo_Axis_Count = 0
    
    Flow_Shmoo_X_Current_Step = -1
    Flow_Shmoo_Y_Current_Step = -1
    Flow_Shmoo_X_Last_Value = -99
    Flow_Shmoo_Y_Last_Value = -99
    Flow_Shmoo_X_Fast = False
    
    For Each site In TheExec.sites
        DevChar_Setup = TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup")
        With TheExec.DevChar.Setups(DevChar_Setup)
            For Each shmoo_axis In .Shmoo.Axes.List
                Select Case shmoo_axis
                    Case tlDevCharShmooAxis_X:
                        axis_type = "X"
                    Case tlDevCharShmooAxis_Y:
                        axis_type = "Y"
                End Select
                With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis)
                    TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_" & axis_type & "_Start") = .Parameter.range.from
                    TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_" & axis_type & "_Stop") = .Parameter.range.To
                    TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_" & axis_type & "_StepSize") = .Parameter.range.StepSize
                    shmoo_step = Abs(Floor((.Parameter.range.To - .Parameter.range.from) / .Parameter.range.StepSize))
                    If axis_type = "X" Then
                        Flow_Shmoo_X_Step = shmoo_step
                        TheExec.sites(site).SiteVariableValue("Flow_Shmoo_X_Step") = shmoo_step
                    Else
                        Flow_Shmoo_Y_Step = shmoo_step
                        TheExec.sites(site).SiteVariableValue("Flow_Shmoo_Y_Step") = shmoo_step
                    End If
                End With
            Next shmoo_axis
        End With
    Next site
     With TheExec.DevChar.Setups(DevChar_Setup)
        For Each shmoo_axis In .Shmoo.Axes.List
            Select Case shmoo_axis
                Case tlDevCharShmooAxis_X:
                    axis_type = "X"
                Case tlDevCharShmooAxis_Y:
                    axis_type = "Y"
            End Select
            With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis)
                Select Case .Parameter.Type.Value
                    Case "Level": Shmoo_Spec = .ApplyTo.Pins & "(" & .Parameter.Name & ")"
                    Case "AC Spec": Shmoo_Spec = .Parameter.Name
                    Case "Global Spec":
                        arg_ary = Split(.InterposeFunctions.PrePoint.Arguments, ",")
                        If LCase(.InterposeFunctions.PrePoint.Name) Like "freerunclk_set_xy" Then
                            Shmoo_Spec = arg_ary(2)
                        Else
                            Shmoo_Spec = .Parameter.Name
                        End If
                    
                End Select

                Flow_Shmoo_Axis(Flow_Shmoo_Axis_Count) = axis_type
                Flow_Shmoo_Axis_Count = Flow_Shmoo_Axis_Count + 1
                If .Parameter.range.from < .Parameter.range.To Then
                    StepSize = .Parameter.range.StepSize
                Else
                    StepSize = -(.Parameter.range.StepSize)
                End If
                If Shmoo_setup_str = "" Then
                    Shmoo_setup_str = "Shmoo_Setup(" & DevChar_Setup & ")" & axis_type & ":" & Shmoo_Spec & "=" & .Parameter.range.from & "," & .Parameter.range.To & "," & StepSize & "; "
                Else
                    Shmoo_setup_str = Shmoo_setup_str & axis_type & ":" & Shmoo_Spec & "=" & .Parameter.range.from & "," & .Parameter.range.To & "," & StepSize & "; "
                End If
            End With
            With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).TrackingParameters
                For Each Shmoo_Tracking_Item In .List
                    Shmoo_Spec = .Item(Shmoo_Tracking_Item).ApplyTo.Pins
                    Shmoo_StepSize = (.Item(Shmoo_Tracking_Item).range.To - .Item(Shmoo_Tracking_Item).range.from) / shmoo_step
                    Shmoo_setup_str = Shmoo_setup_str & axis_type & ":" & Shmoo_Spec & "=" & .Item(Shmoo_Tracking_Item).range.from & "," & .Item(Shmoo_Tracking_Item).range.To & "," & Shmoo_StepSize & "; "
                    Flow_Shmoo_Axis(Flow_Shmoo_Axis_Count) = axis_type & Tracking_num
                    Flow_Shmoo_Axis_Count = Flow_Shmoo_Axis_Count + 1
                Next Shmoo_Tracking_Item
            End With
        Next shmoo_axis
    End With
    TheExec.Datalog.WriteComment "******************************    " & Shmoo_setup_str & "   ******************************"

End Function

'///20170125 Add Check FRC status
Public Function NWireFRCIsEnable(ClockPort As String) As Boolean
Dim site As Variant
    For Each site In TheExec.sites
    
        If TheHdw.Protocol.ports(ClockPort).Enabled = False Then
            NWireFRCIsEnable = False
        Else
            NWireFRCIsEnable = True
        End If
        Exit For
    Next site
End Function


' judge if the shmoo abnormal ratio is large then spec

Public Function CheckCharErrorCount(shmoo_abnormal_type As String, shmoo_abnormal_ratio_hi_lim As Double) As Long

    'Dim site As Long
    Dim shmoo_abnormal_ratio As New SiteDouble
    Dim test_name As String
    
    '' select one of the abnormal type [alarm, shmoo_hole, all_fail]
    For Each site In TheExec.sites
        If included_shmoo_count(site) <> 0 Then
            Select Case LCase(Trim(shmoo_abnormal_type))
                Case "alarm":
                    'For Each site In TheExec.sites
                        shmoo_abnormal_ratio(site) = Format(shmooalarm_count(site) / included_shmoo_count(site), "0.000")
                    'Next site
                    test_name = "CheckChar_Alarm"
                    shmooalarm_count = 0
                Case "shmoo_hole":
                    'For Each site In TheExec.sites
                        shmoo_abnormal_ratio(site) = Format(shmoohole_count(site) / included_shmoo_count(site), "0.000")
                   ' Next site
                    test_name = "CheckChar_Hole"
                     shmoohole_count = 0
                
                Case "all_fail":
                    'For Each site In TheExec.sites
                        shmoo_abnormal_ratio(site) = Format(shmooallfail_count(site) / included_shmoo_count(site), "0.000")
                    'Next site
                    test_name = "CheckChar_AllFail"
                    shmooallfail_count = 0
                Case Else:
                    TheExec.Datalog.WriteComment " [Warning] Unrecognized shmoo_abnormal_type = " & shmoo_abnormal_type + vbCrLf
                    Exit Function
                
            End Select
            
        
            '' Apply Limit
            TheExec.Datalog.WriteComment ""
            'For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site " & site & ", Total_shmoo_count = " & total_shmoo_count(site)
            TheExec.Datalog.WriteComment "site " & site & ", Included_shmoo_count = " & included_shmoo_count(site)
            TheExec.Datalog.WriteComment "site " & site & ", Excluded_shmoo_count = " & excluded_shmoo_count(site)
            'Next site
            TheExec.Datalog.WriteComment ""
            TheExec.Flow.TestLimit resultVal:=shmoo_abnormal_ratio, hiVal:=shmoo_abnormal_ratio_hi_lim, Tname:=test_name, scaletype:=scalePercent
        Else
            'For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ""
            TheExec.Datalog.WriteComment " site " & site & " " & shmoo_abnormal_type & ", [Warning] Included shmoo count = 0 "
            'Next site
        End If
    Next site
End Function

Public Function EnableShmooAbnormalCounter()

    F_shmoo_abnormal_counter = True
    
End Function

Public Function DisableShmooAbnormalCounter()

    F_shmoo_abnormal_counter = False
    
End Function

Public Function Init_Datalog_Setup_Char()    'in on program validated

    If LCase(TheExec.CurrentJob) Like "*char*" Then
        TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = 75
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width = 60
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 70
    End If
    TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
End Function
Public Function Re_PowerOn_WhileSweep(argc As Long, argv() As String)

'argv(0)=PowerPin
'argv(1)=wait time
'argv(2)=cool down voltage

Dim Pin As Variant
Dim Pins() As String
Dim Pin_Cnt As Long
Dim InstName As String
Dim site As Variant
Dim gate_off As Boolean
Dim PowerVolage As New PinListData



'Initialize
Dim powerPin, wait_time, cool_down_voltage As String


'exit function

powerPin = "CorePower"
wait_time = "0.1"
'cool_down_voltage = "0.1"
gate_off = False


If argc = 2 Then
    powerPin = CStr(argv(0))
    wait_time = CDbl(argv(1))
    'cool_down_voltage = CDbl(argv(2))
Else
    TheExec.Datalog.WriteComment "Using default string, Mointor Pin: CorePower, Wait_Time 0.1s"
    TheExec.Datalog.WriteComment "Please fill argument in this format to change default value: Power_pins, wait_time"
    'Exit Function
End If


TheExec.DataManager.DecomposePinList powerPin, Pins(), Pin_Cnt


For Each Pin In Pins
    PowerVolage.AddPin CStr(Pin)
Next Pin

For Each site In TheExec.sites.Active
    For Each Pin In Pins
        If TheExec.DataManager.ChannelType(Pin) <> "N/C" Then
            PowerVolage.Pins(CStr(Pin)).Value = TheHdw.DCVS.Pins(CStr(Pin)).Voltage.Value
        End If
    Next Pin
Next site


For Each site In TheExec.sites.Active
    For Each Pin In Pins
        If TheExec.DataManager.ChannelType(Pin) <> "N/C" Then
            InstName = GetInstrument(CStr(Pin), 0)
            
            Select Case InstName
                   Case "VHDVS"
                    If TheHdw.DCVS.Pins(Pin).Gate = False Then
                        gate_off = True
                        Exit For
                    End If
                   Case "HexVS"
                     If TheHdw.DCVS.Pins(Pin).Gate = False Then
                        gate_off = True
                        Exit For
                     End If
                   Case Else
                        TheExec.Datalog.WriteComment "Error in Re_PowerOn_WhileSweep()"
             End Select
        End If
    Next Pin
    
    
    If gate_off Then
    
        For Each Pin In Pins
            If TheExec.DataManager.ChannelType(Pin) <> "N/C" Then
                InstName = GetInstrument(CStr(Pin), 0)
                Select Case InstName
                       Case "VHDVS"
                            PowerVolage.Pins(CStr(Pin)).Value = TheHdw.DCVS.Pins(CStr(Pin)).Voltage.Value
                            'thehdw.DCVS.Pins(Pin).Voltage.Main.Value = thehdw.DCVS.Pins(Pin).Voltage.Main.Value - cool_down_voltage
                       Case "HexVS"
                            PowerVolage.Pins(CStr(Pin)).Value = TheHdw.DCVS.Pins(CStr(Pin)).Voltage.Value
                            'thehdw.DCVS.Pins(Pin).Voltage.Main.Value = thehdw.DCVS.Pins(Pin).Voltage.Main.Value - cool_down_voltage
                       Case Else
                                TheExec.Datalog.WriteComment "Error in Re_PowerOn_WhileSweep()"
                                TheExec.Datalog.WriteComment "Pin: " & Pin
                 End Select
            'If thehdw.DCVS.Pins(Pin).Gate = False Then thehdw.DCVS.Pins(Pin).Gate = True
            End If
        Next Pin
    End If
Next site

If gate_off = False Then
    Exit Function
End If

DCVS_PowerDown_Parallel_Interpose AllPowerPinlist, 0.001, False


TheHdw.Wait wait_time

DCVS_PowerUp_Parallel_Interpose AllPowerPinlist, "", 0.001, False


For Each site In TheExec.sites.Active
    For Each Pin In Pins
        If TheExec.DataManager.ChannelType(Pin) <> "N/C" Then
            InstName = GetInstrument(CStr(Pin), 0)
            Select Case InstName
                   Case "VHDVS"
                        TheHdw.DCVS.Pins(Pin).Voltage.Main.Value = PowerVolage.Pins(Pin).Value 'thehdw.DCVS.Pins(Pin).Voltage.Main.Value + cool_down_voltage
                   Case "HexVS"
                        TheHdw.DCVS.Pins(Pin).Voltage.Main.Value = PowerVolage.Pins(Pin).Value 'thehdw.DCVS.Pins(Pin).Voltage.Main.Value + cool_down_voltage
                   Case Else
                            TheExec.Datalog.WriteComment "Error in Re_PowerOn_WhileSweep()"
             End Select
        End If
     Next Pin
Next site

End Function


Public Function DCVS_PowerUp_Parallel_Interpose(PowerPinList As String, DisconnectPinList As String, Optional WaitConnectTime As Double = 0.001, Optional DebugFlag As Boolean = False)
'power up sequence at flow start
    Dim CurrentChans As String
    Dim site As Variant
    Dim Pins() As String, PinCnt As Long
    Dim powerPin As Variant
    Dim PowerName As String
    Dim TempString As String
    Dim Vmain As Double
    Dim Irange As Double
    Dim step As Integer
    Dim RiseTime As Double
    Dim PowerSequence As Double
    Dim nwire_port1 As Double
    Dim nwire_port2 As Double
    Dim i As Long
    Dim PowerSequencePin() As String
    Dim TempMaxSequence As Long:: TempMaxSequence = 0
    
    Dim XO0_Port As New PinList
    Dim CLK32768_Port As New PinList
    Dim nwire01_name As String
    Dim nwire02_name As String
    
        Dim XI0_Pin As String
    Dim XI0_SeqName As String
    Dim XI0_Seq As Long
    Dim RTCLK_Pin As String
    Dim RTCLK_SeqName As String
    Dim RTCLK_Seq As Long

    On Error GoTo errHandler
    
     nwire01_name = ""
    nwire02_name = ""
    
    '/////1226///
    ''  ------------------- 20180305 nWire pin form nWire_Ports_GLB ---------------------------
    Dim nWire_port_ary() As String
'    Dim i As Integer
    nWire_port_ary = Split(nWire_Ports_GLB, ",")
    For i = 0 To UBound(nWire_port_ary)
        If LCase(nWire_port_ary(i)) Like "*diff*" Then ' Diff nWire pin
            If LCase(nWire_port_ary(i)) Like "rt*" Then
                RTCLK_Pin = nWire_port_ary(i)                                   '//nWire port name
                nwire01_name = "RT_CLK32768_Diff_Port_PowerSequence_GLB"
            Else
                XI0_Pin = nWire_port_ary(i)
                nwire02_name = "XO0_Diff_Port_PowerSequence_GLB"      '//GB sequence number
            End If
        Else 'SE nWire pin
            If LCase(nWire_port_ary(i)) Like "rt*" Then
                RTCLK_Pin = nWire_port_ary(i)                                   '//nWire port name
                nwire01_name = "RT_CLK32768_Port_PowerSequence_GLB"   '//GB sequence number
            Else
                XI0_Pin = nWire_port_ary(i)
                nwire02_name = "XO0_Port_PowerSequence_GLB"           '//GB sequence number
            End If
        End If
    Next i
    
   
''
''    If RTCLK_GP <> "" Then
''         CLK32768_Port.Value = "RT_CLK32768_Port"                   '//nWire port name
''         nwire01_name = "RT_CLK32768_Port_PowerSequence_GLB"        '//GB sequence number
''    ElseIf RTCLK_Diff_GP <> "" Then
''         CLK32768_Port.Value = "RT_CLK32768_Diff_Port"
''         nwire01_name = "RT_CLK32768_Diff_Port_PowerSequence_GLB"
''    End If
''
''    If XI0_GP <> "" Then
''        XO0_Port.Value = "XO0_Port"                                 '//nWire port name
''        nwire02_name = "XO0_Port_PowerSequence_GLB"                 '//GB sequence number
''    ElseIf XI0_Diff_GP <> "" Then
''        XO0_Port.Value = "XO0_Diff_Port"
''        nwire02_name = "XO0_Diff_Port_PowerSequence_GLB"
''    End If
''    '/////1226///
    
                    

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
     '///1226
     TheHdw.Utility.Pins("k0,k1").State = tlUtilBitOn
     '///1226
    
    Call Print_Header("Power up sequence_Interpose")

    'theexec.Datalog.WriteComment "print: Power up start, Power pins: " & PowerPinList
    'theexec.Datalog.WriteComment RepeatChr("*", 120)
    TheHdw.DCVS.Pins(PowerPinList).Voltage.Main = 0  'reset to 0V
    
    TheExec.DataManager.DecomposePinList PowerPinList, Pins(), PinCnt

    ReDim PowerSequencePin(PinCnt)  'get pin count numbers to arrange array's memory

    For Each powerPin In Pins
        TempString = ""
        PowerName = CStr(powerPin)

        'get power sequence global spec
        TempString = PowerName & "_PowerSequence_GLB"
        PowerSequence = TheExec.specs.Globals(TempString).ContextValue

        If TheExec.DataManager.ChannelType(powerPin) <> "N/C" Then
            If PowerSequence <> 99 Then
                'string power sequence pin
                If PowerSequencePin(PowerSequence) = "" Then
                    PowerSequencePin(PowerSequence) = PowerName
                Else
                    PowerSequencePin(PowerSequence) = PowerSequencePin(PowerSequence) & "," & PowerName
                End If
                If PowerSequence >= TempMaxSequence Then TempMaxSequence = PowerSequence
            'sequence 99, means disconnect pins
            Else
                'TheHdw.DCVS.Pins(PowerPin).Disconnect ' it cause voltage spike, removed it
                Vmain = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
                Irange = TheHdw.DCVS.Pins(powerPin).CurrentRange.Value
                If DebugFlag = True Then    'debugprint
                    'TheExec.Datalog.WriteComment "print: Pin " & PowerName & " is 'NA' pin, disconnect, PowerSequence " & PowerSequence
                    'theexec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(PowerPin & "(N/A)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", RiseTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
                End If
            End If
        Else
            Vmain = 0   'Can not read from DCVS
            Irange = 0
            If DebugFlag = True Then    'debugprint
                'TheExec.Datalog.WriteComment "print: Pin " & PowerName & " not turn on by 'NC pin', PowerSequence " & PowerSequence & " ,Warning!!!"
                'theexec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(PowerPin & "(N/C)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", RiseTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
            End If
        End If
    Next powerPin
    
    '////1226
    nwire_port1 = TheExec.specs.Globals(nwire01_name).ContextValue
    nwire_port2 = TheExec.specs.Globals(nwire02_name).ContextValue
    '///1226
    Dim clk_value As Double
    Dim sites As Variant
    For Each sites In TheExec.sites
    If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Contains(tlDevCharShmooAxis_Y) Then
        If UCase(TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Item(tlDevCharShmooAxis_Y).Parameter.Name) Like "X?#*" Then
            clk_value = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
        End If
    End If
    
    If UCase(TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Item(tlDevCharShmooAxis_X).Parameter.Name) Like "X?#*" Then
        clk_value = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
    End If
    Next sites
    
    For i = 0 To PinCnt
        If PowerSequencePin(i) <> "" Then
        ''power up
        'theexec.Datalog.WriteComment vbCrLf & "print: power up action(" & i & ")" & vbCrLf & RepeatChr("*", 120)
        DCVS_PowerOn_I_Meter_Parallel PowerSequencePin(i), WaitConnectTime, WaitConnectTime, i, DebugFlag
        
        '///1226
            If nwire_port1 = i Then
                TheExec.Datalog.WriteComment vbCrLf & "print: power up for nwire(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                PowerUp_Interpose CLK32768_Port, DebugFlag
            End If
            If nwire_port2 = i Then
                TheExec.Datalog.WriteComment vbCrLf & "print: power up for nwire(" & i & ")" & vbCrLf & RepeatChr("*", 120)
                PowerUp_Interpose XO0_Port, DebugFlag
                '//1226
                Call VaryFreq("XO0_Port", clk_value, "XO0_Freq_Var")
                '//1226
            End If
        '////1226
        
        ''power up
        End If
    Next i
    
    
    Call Print_Footer("Power up sequence_Interpose")
    
    Exit Function
    
errHandler:
        ErrorDescription ("DCVS_PowerUp")
        If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function DCVS_PowerDown_Parallel_Interpose(PowerPinList As String, Optional WaitConnectTime As Double = 0.001, Optional DebugFlag As Boolean = True) ', _
                Optional DriveLowPinList As PinList, Optional ClockPat As Pattern, Optional RTCLK_Relay As PinList, Optional XI0_Relay As PinList)
    Dim CurrentChans As String
    Dim Pins() As String, PinCnt As Long
    Dim powerPin As Variant
    Dim PowerName As String
    Dim TempString As String
    Dim Vmain As Double
    Dim Irange As Double
    Dim step As Integer
    Dim FallTime As Double
    Dim PowerSequence As Double
    Dim i As Long
    Dim site As Variant
    
    Dim RTCLK_Relay As New PinList
    Dim XI0_Relay As New PinList
    
    Dim PowerSequencePin() As String
    Dim seqnum As Long
    Dim TempMaxSequence As Long:: TempMaxSequence = 0
    
    Dim XI0_Pin As String
    Dim XI0_SeqName As String
    Dim XI0_Seq As Long
    Dim RTCLK_Pin As String
    Dim RTCLK_SeqName As String
    Dim RTCLK_Seq As Long
    
    On Error GoTo errHandler
    
    
    Call Print_Header("Power down sequence_Interpose")
    'theexec.Datalog.WriteComment vbCrLf & "print: Power down start, Power pins: " & PowerPinList
    'theexec.Datalog.WriteComment RepeatChr("*", 120)
    
    TheExec.DataManager.DecomposePinList PowerPinList, Pins(), PinCnt
    ReDim PowerSequencePin(PinCnt)

    For Each powerPin In Pins
        TempString = ""
        PowerName = CStr(powerPin)
        
        'get power sequence global spec
        TempString = PowerName & "_PowerSequence_GLB"
        PowerSequence = TheExec.specs.Globals(TempString).ContextValue

        If TheExec.DataManager.ChannelType(powerPin) <> "N/C" Then 'check CP or FT NC pins
            If PowerSequence <> 99 Then
                If PowerSequencePin(PowerSequence) = "" Then
                    PowerSequencePin(PowerSequence) = PowerName
                Else
                    PowerSequencePin(PowerSequence) = PowerSequencePin(PowerSequence) & "," & PowerName
                End If
                If PowerSequence >= TempMaxSequence Then TempMaxSequence = PowerSequence
            'sequence 99, means disconnect pins
            Else
                TheHdw.DCVS.Pins(powerPin).Disconnect
                Vmain = TheHdw.DCVS.Pins(powerPin).Voltage.Main.Value
                Irange = TheHdw.DCVS.Pins(powerPin).CurrentRange.Value
                
                If DebugFlag = True Then    'debugprint
                    'TheExec.Datalog.WriteComment "print: Pin " & PowerName & " is 'NA' pin, disconnect, PowerSequence " & PowerSequence
                    'theexec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(PowerPin & "(N/A)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", FallTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
                End If
            End If
        'NC pins, does not need to power off
        Else
            Vmain = 0   'Can not read from DCVS
            Irange = 0
            If DebugFlag = True Then    'debugprint
                'TheExec.Datalog.WriteComment "print: Pin " & PowerName & " not turn off by 'NC pin', PowerSequence " & PowerSequence & " ,Warning!!!"
                'theexec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(PowerPin & "(N/C)", 30, False) & ", Vmain " & Format(Vmain, "0.000") & " V, Irange " & FormatNumericDatalog(Format(Irange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(0, 2, True) & ", FallTime " & FormatNumericDatalog(0 * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
            End If
        End If
    Next powerPin
    
'/////////////////////////////////////////////////////////////////////////////////////////
    

    For i = PinCnt To 0 Step -1
        If PowerSequencePin(i) <> "" Then
            ''power up
            'theexec.Datalog.WriteComment vbCrLf & "print: power down action(" & i & ")" & vbCrLf & RepeatChr("*", 120)
            DCVS_PowerOff_I_Meter_Parallel PowerSequencePin(i), WaitConnectTime, WaitConnectTime, i, DebugFlag 'WaitConnectTime, WaitConnectTime, i, DebugFlag
            ''power up
        End If
    Next i
'/////////////////////////////////////////////////////////////////////////////////////////
    
    Call Print_Footer("Power down sequence")

    Exit Function
    
errHandler:
        ErrorDescription ("DCVS_PowerDown")
        If AbortTest Then Exit Function Else Resume Next
    
End Function
'*****************************************
'******     current, voltage profile******
'*****************************************




Public Function PostPointInterpose_nWire_ReStore(argc As Long, argv() As String)

    Dim site As Variant
    Dim axis_type As tlDevCharShmooAxis
    
    On Error GoTo errHandler

    Select Case UCase(argv(0))
        Case "X":
            axis_type = tlDevCharShmooAxis_X
        Case "Y":
            axis_type = tlDevCharShmooAxis_Y
        Case Else:
            axis_type = tlDevCharShmooAxis_Invalid
    End Select
    
    Dim nWire_port_ary() As String
    Dim nwp As Variant ', all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
    Dim X_RangeTo As Double
    Dim Y_RangeTo As Double
    Dim SetupName As String
    
    SetupName = TheExec.DevChar.Setups.ActiveSetupName
'    nWire_port_ary = Split(nWire_Ports_GLB, ",")
    'Convert nWire_ports to all_ports and all_pins
    nWire_port_ary = Split("XO0_Port,RT_CLK32768_Port", ",")
    
    For Each site In TheExec.sites
        With TheExec.DevChar
            X_RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.To
            XVal = .Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
            If .Setups(SetupName).Shmoo.Axes.Count > 1 Then
                Y_RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.To
                YVal = .Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
            End If
        End With
    Next site
    
    
    With TheExec.DevChar.Setups(SetupName)
        If .Shmoo.Axes.Count > 1 Then
            If YVal = Y_RangeTo And XVal = X_RangeTo Then
                For Each nwp In nWire_port_ary
                    Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
                    Call VaryFreq(port_pa, TheExec.specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
                Next nwp
            End If
        Else
            If XVal = X_RangeTo Then
                For Each nwp In nWire_port_ary
                    Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
                    Call VaryFreq(port_pa, TheExec.specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
                Next nwp
            End If
        End If
    End With
    
    
    
    
'    With TheExec.DevChar
''        StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
''        RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.From
'        X_RangeTo = .Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.To
'        Y_RangeTo = .Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Range.To
'    End With
''
''    If TheExec.DevChar.Setups(SetupName).Shmoo.Axes.Count > 1 Then
''        XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(axis_type).Value
''        YVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(axis_type).Value
''    End If
''
''
''    For Each site In TheExec.sites
''        XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(axis_type).Value
''        If TheExec.DevChar.Setups(SetupName).Shmoo.Axes.Count > 1 Then
''            YVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(axis_type).Value
''        End If
'''        If YVal = 60000000 And XVal = 4 Then
''    Next site
''
''        If XVal = X_RangeTo Then
''            With TheExec.DevChar.Setups(SetupName)
''                If .Shmoo.Axes.Count > 1 Then
'''                    If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value = "AC Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Type.Value = "AC Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Type.Value = "Global Spec" Or TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value = "Global Spec" Then
''                        For Each nwp In nWire_port_ary
''                            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
''                            Call VaryFreq(port_pa, TheExec.specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
''                        Next nwp
'''                    End If
''                Else
''                End If
''            End With
''        End If
    
    Exit Function
errHandler:

    Stop

End Function






Public Function StoreMaxNum(argc As Long, argv() As String)

    Dim DevSetupName As String
    Dim RangeFrom(2) As Double, RangeTo(2) As Double, RangeStepSize(2) As Double, RangeSteps(2) As Long
    Dim RangeLow(2) As Double
    Dim RangeCalcType(2) As tlDevCharRangeField
    Dim curr_axis As Variant
    Dim Index_val As Long
'    Dim ShmooResult As g_3DShmooResult
    Dim sheetName As String
    Dim TNum_column As Long
    Dim i, j, k As Long
    Dim Tracking_Item As Variant
    Dim tmpStr As String
    On Error GoTo err
    
    Count_Point = 0
    X_Point = 0
    Y_Point = 0
    Z_Point = 0
    Xaxis_index = 0
    Yaxis_index = 0
    Zaxis_index = 0
    X_Tracking_Point = 0
    Y_Tracking_Point = 0
    Z_Tracking_Point = 0
    
    LVCC_flag = False
    HVCC_flag = False
'    Exit Function
    X_dimemsion = False
    Y_dimemsion = False
    Z_dimemsion = False
    
    tmpStr = UCase(Join(argv))
    If InStr(tmpStr, "X") <> 0 Then X_dimemsion = True
    If InStr(tmpStr, "Y") <> 0 Then Y_dimemsion = True
    If InStr(tmpStr, "Z") <> 0 Then Z_dimemsion = True
    
    sheetName = TheExec.Flow.Raw.SheetInRun
    TNum_column = TheExec.Flow.Raw.GetCurrentLineNumber + 5
    g_TestNum = Sheets(sheetName).Cells(TNum_column, 10)
    
    DevSetupName = TheExec.DevChar.Setups.ActiveSetupName
'    ReDim RangeSeq(TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.count - 1)
    
    For Each curr_axis In TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.List
        With TheExec.DevChar
            RangeFrom(curr_axis) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.from
            RangeTo(curr_axis) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.To
            RangeSteps(curr_axis) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.Steps + 1
            RangeStepSize(curr_axis) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.StepSize
            RangeCalcType(curr_axis) = .Setups(DevSetupName).Shmoo.Axes(curr_axis).Parameter.range.CalculatedField

            If RangeCalcType(curr_axis) = 1 Then 'RangeCalcType="STEPS"
                If RangeFrom(curr_axis) < RangeTo(curr_axis) Then
                    RangeSeq(curr_axis) = True ' small---->lager
                Else
                    RangeSeq(curr_axis) = False
                End If
                Index_val = Abs((RangeTo(curr_axis) - RangeFrom(curr_axis))) / RangeStepSize(curr_axis)
                If InStr(Index_val, ".") <> 0 Then Index_val = Index_val + 1
            ElseIf RangeCalcType(curr_axis) = 0 Then '"STEP SIZES"
                Index_val = RangeSteps(curr_axis)
            End If
'            ReDim Preserve ShmooResult.Three_axis(curr_axis).Axis_CurrPoint(Index_val)
'            ReDim Preserve ShmooResult.Three_axis(curr_axis).Axis_CurrResult(Index_val)
            
            Select Case curr_axis
                Case 0
                    Xaxis_index = Index_val + 1
                    X_Tracking_Point = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count
'                    ReDim g_ShmooResult.Axis_CurrPoint.X_axis_Tracking(X_Tracking_Point)
               Case 1
                    Yaxis_index = Index_val + 1
                    Y_Tracking_Point = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count
                Case 2
                    Zaxis_index = Index_val + 1
                    Z_Tracking_Point = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count
            End Select
        End With
''        With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
''            For Each Tracking_Item In .List
'''                 TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(Tracking_Item).Value
''            Next Tracking_Item
''        End With
'        If RangeFrom(curr_axis) < RangeTo(curr_axis) Then
'            LVCC_flag = True
'        Else
'            HVCC_flag = True
'        End If
'        StartPoint = RangeFrom(curr_axis)
'        StopPoint = RangeTo(curr_axis)
'        StepSize = RangeStepSize(curr_axis)
    Next curr_axis
    If TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 3 Then '  3D shmoo only
        MaxArrIndex = Xaxis_index * Yaxis_index * Zaxis_index
        ReDim g_ShmooResult.Axis_CurrPoint(MaxArrIndex - 1)
        If X_Tracking_Point = 0 And Y_Tracking_Point = 0 And Z_Tracking_Point = 0 Then ' means no any tracking parameters.
            'Do nothing
        Else
            For k = 0 To MaxArrIndex - 1
                If X_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).X_axis_Tracking(X_Tracking_Point)
                Else
                    ReDim g_ShmooResult.Axis_CurrPoint(k).X_axis_Tracking(X_Tracking_Point - 1)
                End If
                
                If Y_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).Y_axis_Tracking(Y_Tracking_Point)
                Else
                    ReDim g_ShmooResult.Axis_CurrPoint(k).Y_axis_Tracking(Y_Tracking_Point - 1)
                End If
                
                If Z_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).Z_axis_Tracking(Z_Tracking_Point)
                Else
                    ReDim g_ShmooResult.Axis_CurrPoint(k).Z_axis_Tracking(Z_Tracking_Point - 1)
                End If
            Next k
        End If
    ElseIf TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 2 Then ' 2D case
        MaxArrIndex = Xaxis_index * Yaxis_index ' only X-Y case
        ReDim g_ShmooResult.Axis_CurrPoint(MaxArrIndex - 1)
        If X_Tracking_Point = 0 And Y_Tracking_Point = 0 Then  ' means no any tracking parameters.
            'Do nothing
        Else
            For k = 0 To MaxArrIndex - 1
                If X_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).X_axis_Tracking(X_Tracking_Point)
                Else
                    ReDim g_ShmooResult.Axis_CurrPoint(k).X_axis_Tracking(X_Tracking_Point - 1)
                End If
                
                If Y_Tracking_Point = 0 Then
                    ReDim g_ShmooResult.Axis_CurrPoint(k).Y_axis_Tracking(Y_Tracking_Point)
                Else
                    ReDim g_ShmooResult.Axis_CurrPoint(k).Y_axis_Tracking(Y_Tracking_Point - 1)
                End If
            Next k
        End If
    End If
'    ReDim g_ShmooResult.Axis_CurrResult(MaxArrIndex)
'    For i = 0 To 2
'        For j = 0 To 2
'            For k = 0 To 2
'                ReDim g_ShmooResult.X_axis(i).Y_axis(j).Z_axis(k).Axis_CurrPoint(MaxArrIndex)  ' in order to fill in Max Array number to record all shmoo result
'            Next k
'        Next j
'    Next i
    Exit Function
err:
    If AbortTest Then Exit Function Else Resume Next
End Function

