Attribute VB_Name = "VBT_LIB_Digital_Functional_T"

Option Explicit
'Revision History:
'V0.0 initial bring up
'V0.1 Add Mbist finger print function

Public Const LVCC_boundary_Switch = 1 '1~10 means only get fail log at LVCC boundary with how many times
                                      '0 means get fail log with full search range

' Digital Functional Test

' (c) Teradyne, Inc, 1997-2008
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
' Revision History:
' Date                  Description
' 12-jun-12 s.bullock           tersw00187233 - locally save/restored some member variables in functioal_t to handle concurrent test use cases.
' 27-Mar-12 Venkata Kotireddy tersw00184533 - Fixed the issue to not save/restore flag match if the user did not specify flags to match when APS is enabled, and not suspended
' 13-Mar-12 R.Stimson   tersw00184562 - Restore error handler after Patterns.Test.
' 03-Jul-11 Obula Reddy   tersw00172420 - added getdefaults() to set the defaule value for template arguments
' 08-Jul-10 Pavan         tersw00166339  Added validation support for WaitTimeDomain.
' 09/10/09  David Sanders tersw00146124 Template code for patgen flag matching
'                         does not work with multiple time domains
' 01/06/09  Tim Orr     tersw001334426, slowdown in template test groups: see FetchContext
' 09/1/2005             Ported from Flex


' ============
' Private Data
' ============

' Context values on the Test Instances sheet
Private m_TimeSetSheet As String, m_LevelsSheet As String
Private m_InstanceName As String

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

' ============
' Public Enums
' ============

' CPU flag wait conditions
'Public Enum tlWaitVal
'    waitoff = -2    ' default value is first
'    waitLo = 0
'    waitHi = -1
'End Enum

Public Enum CusWaitVal
    waitoff = 0    ' default value is first
    waitLo = -1
    waitHi = -2
End Enum

Private Const TL_E_AT_PATSET_BREAKPT = &HC0000014

''''20151106 Set Variable for Functional_T_char_Mbist()
Private gm_Patterns As String
Private gm_bistType As String
Private gm_Power_Run_Scenario As String
Private gm_CharInputString As String
Private gm_AI_fail_point As String
Private gm_testName As String
Private gm_patcnt As Long
Private gm_rtnINITpattArr() As String
Private gm_rtnPLLPpattArr() As String
Private gm_wait_time_ary() As String
Public gB_shmooAccumResult As New SiteLong ''''20151110 New for the Accumlated shmoo result for the multi-patterns
Private gm_freqPattSet As New Pattern ''''20151111 New
Public Block As String
Public mbist_flag_set_placement As Long




' Perform a digital functional test.
' Return TL_SUCCESS if the test executes without problems, else TL_ERROR.
Public Function Functional_T_updated(Patterns As Pattern, StartOfBodyF As InterposeName, _
                             PrePatF As InterposeName, PreTestF As InterposeName, _
                             PostTestF As InterposeName, PostPatF As InterposeName, EndOfBodyF As InterposeName, _
                             ReportResult As PFType, ResultMode As tlResultMode, DriveLoPins As PinList, DriveHiPins As PinList, _
                             DriveZPins As PinList, DisablePins As PinList, FloatPins As PinList, StartOfBodyFArgs As String, _
                             PrePatFArgs As String, PreTestFArgs As String, PostTestFArgs As String, _
                             PostPatFArgs As String, EndOfBodyFArgs As String, Util1Pins As PinList, _
                             Util0Pins As PinList, PatFlagF As InterposeName, _
                             PatFlagFArgs As String, RelayMode As tlRelayMode, PatThreading As Boolean, _
                             MatchAllSites As Boolean, WaitFlagA As CusWaitVal, WaitFlagB As CusWaitVal, _
                             WaitFlagC As CusWaitVal, WaitFlagD As CusWaitVal, Validating_ As Boolean, _
                             Optional PatternTimeout As String = "30", Optional Step_ As SubType, _
                             Optional WaitTimeDomain As String, Optional ConcurrentMode As tlPatConcurrentMode = tlPatConcurrentModeCached, _
                             Optional Interpose_PrePat As String, _
                             Optional RunFailCycle As Boolean = False, Optional EnableBinOut As Boolean = False, Optional inst_width As Integer, _
                             Optional DigSource As String = "", Optional DigCapture As String = "", Optional ForceResults As Boolean = False, Optional MbistIndicator As Boolean = False, _
                             Optional MbistMatchLoopCountValue As Long = 0) As Long
' EDITFORMAT1 1,,Pattern,,,Patterns|7,,InterposeName,Interpose Functions,,StartOfBodyF|9,,InterposeName,,,PrePatF|11,,InterposeName,,,PreTestF|13,,InterposeName,,,PostTestF|15,,InterposeName,,,PostPatF|17,,InterposeName,,,EndOfBodyF|2,,PFType,,,ReportResult|6,,tlResultMode,,,ResultMode|19,,pinlist,Pin States,,DriveLoPins|20,,pinlist,,,DriveHiPins|21,,pinlist,,,DriveZPins|22,,pinlist,,,DisablePins|23,,pinlist,,,FloatPins|8,,String,,,StartOfBodyFArgs|10,,String,,,PrePatFArgs|12,,String,,,PreTestFArgs|14,,String,,,PostTestFArgs|16,,String,,,PostPatFArgs|18,,String,,,EndOfBodyFArgs|24,,pinlist,,,Util1Pins|25,,pinlist,,,Util0Pins|31,,InterposeName,,,PatFlagF|32,,String,,,PatFlagFArgs|5,,tlRelayMode,,,RelayMode|3,,Boolean,,,PatThreading|30,,Boolean,,,MatchAllSites|26,,CusWaitVal,Flag Match,,WaitFlagA|27,,CusWaitVal,,,WaitFlagB|28,,CusWaitVal,,,WaitFlagC|29,,CusWaitVal,,,WaitFlagD|0,,Boolean,,,Validating_|4,,String,,0 <= PatternTimeout,PatternTimeout|6,,tlPatStartConcurrentMode,,,ConcurrentMode


    Functional_T_updated = TL_SUCCESS   ' be optimistic
    If Not TheExec.Flow.IsRunning Then Exit Function
    
    On Error GoTo errHandler

    ' Cache parameters for PostTest
    m_EndOfBodyF = EndOfBodyF
    m_EndOfBodyFArgs = EndOfBodyFArgs
    
    ' Apply default values to parameters whose values were not specified.
    ApplyDefaults PatternTimeout
    
    If Validating_ Then
        ' Perform additional parameter validation
        If Not Validate(Patterns, PatThreading, DriveLoPins, DriveHiPins, DriveZPins, DisablePins, _
            FloatPins, Util1Pins, Util0Pins, PatternTimeout, WaitTimeDomain) Then Functional_T_updated = TL_ERROR
        If Patterns.Value <> "" Then Call PrLoadPattern(Patterns.Value)
        If PrePatF.Value <> "" Then Call PrLoadPattern(PrePatF.Value)
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
        '==============================20180226 Vramp to prevent voltage spike==============================
        m_InstanceName = LCase(TheExec.DataManager.instanceName)
        
        If UCase(m_InstanceName) Like UCase("SocMbist*") Or UCase(m_InstanceName) Like UCase("CpuMbist*") Or UCase(m_InstanceName) Like UCase("GfxMbist*") Then
            Call MbistRampApplyLevel_AutoReadingContext(, , , m_InstanceName)
        End If
        '===================================================================================================
        
        ' Set up the test
        Call PreBody(DriveHiPins, DriveLoPins, DriveZPins, DisablePins, Util1Pins, Util0Pins, _
                 WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD, MatchAllSites, _
                 PatThreading, RelayMode, WaitTimeDomain, "")
    End If ' PreBody
        
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
                
        ' Perform the test
        Call Interpose(StartOfBodyF, StartOfBodyFArgs)
        
        '''20180621 for shmoo PTR high/low limit
        'Add for force condition.
        '2017/11/02 Add STORE Pre Pat String
        If Interpose_PrePat <> "" Then
            Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
        End If
        
        If (RelayMode = tlUnpowered) Then MsgBox "Please change relay mode to powered", vbOKOnly, "IG-XL alarm"
     
        Call Body(FloatPins, StrToDbl(PatternTimeout), Patterns, ReportResult, ResultMode)
        
                
        ' Run the pattern.  Perform functional test.
        If TheExec.sites.ActiveCount > 0 Then
            On Error Resume Next
            
            Shmoo_Pattern = Patterns.Value
            
            'Add for Mbist finger print
   ' =========================================================================DSSC============================================================================'
        'PC modified for DSSC Mapping
        
         If DigSource <> "" Then
            Shmoo_Save_core_power_per_site_for_Vbump
            Dim Pattern_Decompose() As String
            Dim PatCnt As Long
            Dim DSSC_Pattern As String
            Dim DSSC_Pattern_Count As Long
            Dim i As Long, j As Long
            Dim TestCase As String
            Dim digSrc_EQ As String
            Dim DigSrc_Size As Double
            Dim DigSrc_pin As New PinList
            Dim DigSource_Arr() As String
            Dim DigSrc_wav As New DSPWave
            Dim PattArray() As String
            Dim site As Variant
            Dim Store_PinList As New PinListData
            Dim BlockType As String, BlockHeader As String
            BlockType = Split(m_InstanceName, "_")(0)
            BlockHeader = Left(BlockType, 3)
            If UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
                BlockType = "SCAN"
            ElseIf UCase(BlockType) Like "*MBIST*" Then
                BlockType = "MBIST"
            End If
            BlockType = UCase(BlockHeader) & BlockType
            DigSource_Arr() = Split(DigSource, ":")
            TestCase = DigSource_Arr(0)
            DigSrc_pin.Value = DigSource_Arr(1)

            Pattern_Decompose = TheExec.DataManager.Raw.GetPatternsInSet(Shmoo_Pattern, PatCnt)
            DSSC_Pattern_Count = 0
            Dim DecodeBit_Str As String
            Dim DecodeBit_Ary() As String
            For i = 0 To UBound(Pattern_Decompose)
            
                For j = 0 To UBound(SrcStock)
                
                    If UCase(Pattern_Decompose(i)) Like UCase("*" & SrcStock(j).PatternName & "*") Then
                        DSSC_Pattern = SrcStock(j).PatternName
                        DSSC_Pattern_Count = DSSC_Pattern_Count + 1 'Prevent DSSC patterns more than one
                        Call GetSrcString_fromEMAArray(DSSC_Pattern, TestCase, digSrc_EQ, DigSrc_Size, DecodeBit_Ary)
                        Set DigSrc_wav = Nothing
                        DigSrc_wav.CreateConstant 0, CLng(DigSrc_Size)
                        If UCase(digSrc_EQ) Like "*S*" Then
                            Dim TempStr As String
                            TempStr = Decide_Switching_Bit(digSrc_EQ, DigSrc_wav, g_ApplyLevelTimingValt, BlockType, DecodeBit_Str)
                        Else
                            Dim k As Integer
                            For k = 0 To Len(digSrc_EQ) - 1
                                For Each site In TheExec.sites.Active
                                      DigSrc_wav.Element(k) = CDbl(Mid(digSrc_EQ, k + 1, 1))
                                Next site
                            Next k
                            TempStr = digSrc_EQ
                        End If
                        Call PATT_GetPatListFromPatternSet(DSSC_Pattern, PattArray, PatCnt)
                        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "FUNC_SRC", CLng(DigSrc_Size), DigSrc_wav)
                        For Each site In TheExec.sites
                            TheExec.Datalog.WriteComment "Site" & site & " " & "DigCap pattern = " & "DSSC_Pattern" & ": " & DSSC_Pattern & "," & "Src Bits = " & Len(TempStr) & "," & "SELSRAM_SEND [ First(L) ==> Last(R) ]" & TempStr & ", IAGEPDS :" & DecodeBit_Str
                        Next site
                        
                    End If
                Next j
              j = 0
              Next i
           
            'If DSSC_Pattern_Count > 1 Then TheExec.ErrorLogMessage "Number of DSSC Patterns more than one   "
            If DSSC_Pattern = "" Then TheExec.ErrorLogMessage " Can not find corresponding DSSC pattern from DSSC Mapping table"
            
            'Call GetSrcString_fromEMAArray(DSSC_Pattern, TestCase, digSrc_EQ, DigSrc_Size)
               digSrc_EQ_GB = digSrc_EQ:: BlockType_GB = BlockType:: DigSrcSize_GB = DigSrc_Size:: dssc_pat_init_GB = PattArray(0):: DigSrc_pin_GB = DigSrc_pin
 
 
            'CpuSaChain_B00_CL10_SAA_PP_CONA0_C_PLLP_CH_CL10_SAA_UNC_AUT_ALLFV_SI
            'Call SetupDigSrcDspWave(PattArray(0), DigSrc_Pin, "FUNC_SRC", CLng(DigSrc_Size), DigSrc_wav)
            
'            For Each Site In TheExec.sites
'              TheExec.Datalog.WriteComment "Site" & Site & " " & "DigCap pattern = " & "DSSC_Pattern" & ": " & DSSC_Pattern & "," & "Src Bits = " & Len(digSrc_EQ) & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & digSrc_EQ
'            Next Site
         End If
''======================================================================DSSC Capture Set up================================================================
'Dim OutDspWave(0) As New DSPWave
'Dim OutDspWave() As New DSPWave
'''ReDim OutDspWave(2)
''
'''PC modified for DSSC Capture
'         If DigCapture <> "" Then
'            Dim DigCap_Pin As New PinList
'            Dim DigCap_Sample_Size As Long
'
'            Dim DigCap_DataWidth As Long
'            Dim DSSC_Capture_Out As String
'            Dim DigCap_Arr() As String
'            Set OutDspWave = Nothing
'
'            DigCap_Arr() = Split(DigCapture, ":")
'            DigCap_Sample_Size = DigCap_Arr(0)
'            DigCap_Pin = DigCap_Arr(1)
'            Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
'            TheExec.Datalog.WriteComment ("Cap Bits = " & DigCap_Sample_Size)
'            TheExec.Datalog.WriteComment ("Cap Pin = " & DigCap_Pin)
'            TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test End   ========")
'          End If
' =========================================================================DSSC============================================================================'
           
            
        If MbistMatchLoopCountValue > 0 Then
           
            Dim MatchLoopNum As Long

            MatchLoopNum = MbistMatchLoopCountValue
            TheHdw.Digital.Patgen.counter(tlPgCounter10) = MatchLoopNum
         
            '''====================================
            If TheExec.TesterMode = testModeOffline Then
                TheHdw.Digital.Patgen.counter(tlPgCounter10) = 1
            End If
            '''====================================
        End If
    
        Dim original_fun_width As Integer
        
        If (inst_width > 0) Then
                original_fun_width = TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width
          TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
          'TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.TestName.Width = 75
          TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width = inst_width
          'TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 70
          TheExec.Datalog.ApplySetup
        End If

        Call pattern_module_test(Patterns.Value, RunFailCycle, EnableBinOut, ReportResult, TL_C_YES, ResultMode, ConcurrentMode)     ' test chip block loop function

        If (inst_width > 0) Then
                TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
          'TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.TestName.Width = 75
          TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width = original_fun_width
          'TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 70
          TheExec.Datalog.ApplySetup
        End If

'            If RunFailCycle Then
'              If theexec.EnableWord("Mbist_FingerPrint") = True Then
'                Call auto_FuncTest_Mbist_ExecuteForShowFailBlock(Patterns.Value, EnableBinout) 'Mbist finger print VBT
'              Else
'                Call thehdw.Patterns(Patterns).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode) 'Function T org execute function
'              End If
'            Else
'                Call thehdw.Patterns(Patterns).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode) 'Function T org execute function
'            End If
        
            If MbistMatchLoopCountValue > 0 Then
            
            Dim RealLoopNum As Long
            
            RealLoopNum = (MatchLoopNum - TheHdw.Digital.Patgen.counter(tlPgCounter10))
            TheExec.Datalog.WriteComment "Set C10 of " & TheExec.DataManager.instanceName & " : " & MatchLoopNum & " run down to " & " : " & TheHdw.Digital.Patgen.counter(tlPgCounter10) & " Total Loop Counts " & " : " & RealLoopNum
            
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
        DebugPrintFunc Patterns.Value
        
        Dim PatSetArray() As String
        Dim PrintPatSet As Variant
        Dim patt_ary_debug() As String
        Dim pat_count_debug As Long
        Dim patt As Variant
        If False Then
        PatSetArray = Split(Patterns.Value, ",")
        
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
    
    '2017/11/02 Add RESTORE Pre Pat String in post body
        If Interpose_PrePat <> "" Then
            Call SetForceCondition("RESTOREPREPAT")
        End If
        
        Call PostBody(m_DrivePins, m_FloatPins, WaitTimeDomain, WaitFlagA, WaitFlagB, WaitFlagC, WaitFlagD)
    End If ' PostBody
    
    ' There shouldn't be any code below this line. Any other necessary
    ' code should be added to the PostTest method to support pattern set
    ' breakpoints.
        

   
    Exit Function
    
errHandler:
    Call TheExec.ErrorLogMessage("Test " & TL_C_ERRORSTR & ", Instance: " & TheExec.DataManager.instanceName)
    Call TheExec.ErrorReport
    ' Clear previously registered interpose function names
    Call tl_ClearInterpose(TL_C_PREPATF, TL_C_POSTPATF, TL_C_PRETESTF, TL_C_POSTTESTF)
    m_InterposeFunctionsSet = False

    Functional_T_updated = TL_ERROR
                If AbortTest Then Exit Function Else Resume Next
End Function

' Perform a digital functional test.
' Return TL_SUCCESS if the test executes without problems, else TL_ERROR.




Public Function DatalogType() As Integer
    DatalogType = logFunctional
End Function

' =====================
' Private work routines
' =====================

' Do test setup.  This involves setting the timing and levels, connecting the pins, and
' various other functions in preparation for running the pattern.
Private Sub PreBody(DriveHiPins As PinList, DriveLoPins As PinList, DriveZPins As PinList, DisablePins As PinList, _
                    Util1Pins As PinList, Util0Pins As PinList, WaitFlagA As CusWaitVal, _
                    WaitFlagB As CusWaitVal, WaitFlagC As CusWaitVal, WaitFlagD As CusWaitVal, MatchAllSites As Boolean, _
                    PatThreading As Boolean, RelayMode As tlRelayMode, _
                    WaitTimeDomain As String, CharInputString As String)

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
'        If CharInputString <> "" Then
'            Call SetCharPower(CharInputString)
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

' Restore tester state to the default
Private Sub PostBody(DrivePins As String, FloatPins As String, WaitTimeDomain As String, WaitFlagA As CusWaitVal, _
                    WaitFlagB As CusWaitVal, WaitFlagC As CusWaitVal, WaitFlagD As CusWaitVal)

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

' PostPat Breakpoint interpose function. This is need to support the pattern set
' breakpoint feature.
Public Sub PostTest(argc As Long, argv() As String)

    Call Interpose(m_EndOfBodyF, m_EndOfBodyFArgs)
    
End Sub

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

Private Sub ApplyDefaults(ByRef PatternTimeout As String)
    ' If the worksheet doesn't have a value then apply 30 as the default.
    If Not NonBlank(PatternTimeout) Then
        PatternTimeout = "30"
    End If
End Sub

'this function is used by test instance sheet,to write the default value for the argument when new test instance is created
'the array can be redimed to hold more values.
Public Function getdefaults() As Variant
        Dim argdefaults(4) As Variant
        argdefaults(0) = "waitflaga,0" 'argument name and value .
        argdefaults(1) = "waitflagb,0"
        argdefaults(2) = "waitflagc,0"
        argdefaults(3) = "waitflagd,0"
        argdefaults(4) = "patterntimeout,30"
        getdefaults = argdefaults
End Function
' Return TL_SUCCESS if the test executes without problems, else TL_ERROR.
'=============================20160413==================================

Public Function pattern_module_test(pattern_load As String, RunFailCycle As Boolean, EnableBinOut As Boolean, ReportResult As PFType, TL_C_YES As Long, ResultMode As tlResultMode, ConcurrentMode As tlPatConcurrentMode)

    Dim ins_name As String
    Dim i As Long:: i = 0
    Dim site As Variant
    Dim site_BK_loop_count As Long
    Dim pattern_name(8) As String
    Dim Flag_Name As String
    Dim c As Boolean
    Dim confirm_inst As Boolean
    Dim ws_def As Worksheet
    Dim wb As Workbook

    Dim maxDepth As Integer
    Dim patt As Variant
    Dim rtnPatternNames() As String, rtnPatternCount As Long
    Dim astrPattPathSplit() As String
    Dim astrPattPathSplit_01() As String
    Dim blPatPass As New SiteBoolean
    Dim numcap As Long
    Dim PinData_d As New PinListData
    Dim Mbist_repair_cycle As Long
    Dim Pins As New PinData
    Dim Cdata As Variant
    Dim TestNumber As New SiteLong
    Dim ins_new_name As String
    Dim tested As New SiteBoolean
    Dim strPattName As String
    Dim inst_match As Boolean
    Dim Temp As Long
    Dim allpins As String
    Dim PinData As New PinListData
    
    Dim blMbistFP_Binout As Boolean
    Dim MBISTFailBlockFlag As Boolean
    Dim PassOrFail As New SiteLong
    Dim lGetFlagIdx As Long
    Dim blJump As Boolean
    Dim m_testName As String
    Dim k As Long, p As Long, g As Long, j As Long:: k = 0:: p = 0:: g = 0:: j = 0
    
    On Error GoTo errHandler
    'LogLimited = 255
    allpins = "JTAG_TDO"
    m_testName = TheExec.DataManager.instanceName
    'Dim pattern_load As String
    'pattern_load = ".\Patterns\vreg_test_pop_student.pat"
    '-----------------------------------------------------------------------------------------
    For Each site In TheExec.sites
        If (TheExec.sites(site).SiteVariableValue("LP_BM") = TheExec.sites(site).SiteVariableValue("Lcount_BM")) Then
            confirm_in_loop = False
            Exit For
        End If
    Next site
    '-----------------------------------------------------------------------------------------
    If (confirm_in_loop = True) Then
        confirm_inst = False
        '============================================================================init setting
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        
        ins_name = TheExec.DataManager.instanceName
        
        For Each site In TheExec.sites
            site_BK_loop_count = TheExec.sites(site).SiteVariableValue("LP_BM")
            'pattern_name(Site) = mbist_dynamic.Block_dynamic(0).pat_name_dynamic(site_BK_loop_count)
            
            If UCase(ins_name) Like "*IVDM*" Or UCase(ins_name) Like "*SNSUHS*" Then
            Else
                TheExec.sites.Item(site).TestNumber = TheExec.sites.Item(site).TestNumber + site_BK_loop_count * 100001
            End If
            
            'Exit For
        Next site
        '============================================================================excute test
        For i = 0 To UBound(mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(0).instance_dynamic)
            If (UCase(ins_name) Like UCase("*" + mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(0).instance_dynamic(i) + "*") And (mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(0).instance_dynamic(i) <> "")) Then
                confirm_inst = True
                Exit For
            End If
        Next i
 
         If (confirm_inst = True) Then
            Call auto_Mbist_Block_loop_inst_match(ins_name, pattern_load, site_BK_loop_count, Flag_Name, EnableBinOut, RunFailCycle)
         Else
            Call auto_Mbist_Block_loop_inst_non_match(ins_name, pattern_load, site_BK_loop_count, Flag_Name, EnableBinOut, RunFailCycle)
         End If
        '============================================================================
    Else
        If LCase(TheExec.DataManager.instanceName) Like "*bist*" And TheExec.EnableWord("Mbist_FingerPrint") = True Then
                Call Finger_print(pattern_load, RunFailCycle, Flag_Name, True)
        Else
        
            If TheExec.TesterMode = testModeOffline Then
                Call ATPG_offline(pattern_load, ResultMode)
            Else
                Call TheHdw.Patterns(pattern_load).Test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode)
            End If
        End If
    End If
    '--------------------------------------------------------------------------------print out flag sheet
'''    create_flag_sheet = True
'''    If (create_flag_sheet And confirm_in_loop = True) Then
'''        For Each Site In TheExec.sites
''''            If ws_def Is Nothing Then
'''            If Not WorksheetFunction.IsErr(Evaluate("'" & "Mbist_Block_loop_flag_list" & "'!A1")) = False Then
'''                Sheets.Add after:=Sheets(Sheets.Count)
'''                Sheets(Sheets.Count).Name = "Mbist_Block_loop_flag_list"
'''            Else
''''                Application.DisplayAlerts = False
''''                Sheets("BinOutCalcScanMbistTable").delete
'''            End If
'''            Set wb = Application.ActiveWorkbook
'''            Set ws_def = wb.Sheets("Mbist_Block_loop_flag_list")
'''
'''            If (TheExec.sites(Site).SiteVariableValue("LP_BM") = 0 And index_flag_y = 1) Then
'''                ws_def.Cells(index_flag_y, 2 * index_flag_x - 1).Value = "Flag-" + bist_type + "_ " + mbist_dynamic.Block_dynamic(0).block_name_dynamic
'''                ws_def.Cells(index_flag_y, 2 * index_flag_x).Value = "Binout-" + bist_type + "_ " + mbist_dynamic.Block_dynamic(0).block_name_dynamic
'''                index_flag_y = index_flag_y + 1
'''            End If
'''            ws_def.Cells(index_flag_y, 2 * index_flag_x - 1).Value = flag_name
'''
'''            If (TheExec.sites.Item(Site).FlagState(flag_name) = logicTrue) Then
'''                ws_def.Cells(index_flag_y, 2 * index_flag_x).Value = "Fail"
'''            ElseIf (TheExec.sites.Item(Site).FlagState(flag_name) = logicFalse) Then
'''                ws_def.Cells(index_flag_y, 2 * index_flag_x).Value = "Pass"
'''            Else
'''                ws_def.Cells(index_flag_y, 2 * index_flag_x).Value = "Clean"
'''            End If
'''            index_flag_y = index_flag_y + 1
'''            Exit For
'''        Next Site
'''    End If
    '--------------------------------------------------------------------------------
    
    '--------------------------------------------------------------------------------print out flag sheet for txt/csv file
    If (create_flag_sheet And confirm_in_loop = True) Then
        Dim FileExists As Boolean
        Dim string_store As String, string_store01 As String
        For Each site In TheExec.sites
                FileExists = (Dir(File_path) <> "")
                If FileExists = False Then
                    Open File_path For Output As #1
                End If
                If (TheExec.sites(site).SiteVariableValue("LP_BM") = 0 And index_flag_y = 1) Then
                    string_store = ""
                    string_store01 = ""
                    string_store = "Flag-" + bist_type + "_" + mbist_dynamic.Block_dynamic(0).block_name_dynamic
                    string_store01 = "Binout-" + bist_type + "_" + mbist_dynamic.Block_dynamic(0).block_name_dynamic
                    Print #1, "===============================================================,======================="
                    Print #1, string_store + "," + string_store01
                    'Write #1, "Flag-SOC_ SOC,Binout-SOC_ SOC"
                    index_flag_y = index_flag_y + 1
                End If
                Print #1, Flag_Name + ",";
                If (TheExec.sites.Item(site).FlagState(Flag_Name) = logicTrue) Then
                    Write #1, "Fail"
                ElseIf (TheExec.sites.Item(site).FlagState(Flag_Name) = logicFalse) Then
                    Write #1, "Pass"
                Else
                    Write #1, "Clean"
                End If
                'Close #1
                Exit For
        Next site
    End If
    '--------------------------------------------------------------------------------
    
    Exit Function
    
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_Mbist_Block_loop_inst_match(instance_name As String, m_pattname As String, bk_loop_count As Long, ByRef Flag_Name As String, EnableBinOut As Boolean, RunFailCycle As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Mbist_Block_loop_inst_match"
    Dim site As Variant
    'Dim flag_name As String
    Dim ins_array_name_pp() As String
    Dim ins_array_name_type() As String
    Dim ins_array_name_others() As String
    
    Dim flag_array_string_match() As String
    Dim flag_array_string_inst() As String
    Dim match_begin As Long, match_end As Long:: match_begin = match_end = 0
    Dim flag_spilt As String
    Dim ins_array_name_long() As String

    Dim ins_array_name_perf_v() As String
    Dim i As Long, k As Long, p As Long, g As Long, j As Long:: i = 0:: k = 0:: p = 0:: g = 0:: j = 0
    Dim confirm As Boolean
    
    Dim LNH_V As String
    Dim perofmrance As String
    Dim maxDepth As Integer
    Dim patt As Variant
    Dim rtnPatternNames() As String, rtnPatternCount As Long
    Dim astrPattPathSplit() As String
    Dim astrPattPathSplit_01() As String
    Dim blPatPass As New SiteBoolean
    Dim numcap As Long
    Dim PinData_d As New PinListData
    Dim Mbist_repair_cycle As Long
    Dim Pins As New PinData
    Dim Cdata As Variant
    Dim TestNumber As New SiteLong
    Dim ins_new_name As String
    Dim tested As New SiteBoolean
    Dim strPattName As String
    Dim flag_match As Boolean
    Dim Temp As Long
    Dim allpins As String
    Dim PinData As New PinListData
    
    Dim match_string_1st As String
    
    Dim blMbistFP_Binout As Boolean
    Dim MBISTFailBlockFlag As Boolean
    Dim PassOrFail As New SiteLong
    Dim lGetFlagIdx As Long
    Dim blJump As Boolean
    
    Dim m_testName As String
    Dim for_confirm_ins_name As String
    for_confirm_ins_name = ""
    Dim for_confirm_ins_name_array() As String
    
    
    allpins = "JTAG_TDO"
    ins_array_name_perf_v = Split(instance_name, "_")
    m_testName = ins_array_name_perf_v(0)
    '=================================================================================================test flag
    For i = 0 To UBound(ins_array_name_perf_v)
        If (UCase(ins_array_name_perf_v(i)) Like "NV" Or UCase(ins_array_name_perf_v(i)) Like "LV" Or UCase(ins_array_name_perf_v(i)) Like "HV" Or UCase(ins_array_name_perf_v(i)) Like "MNV" Or UCase(ins_array_name_perf_v(i)) Like "MLV" Or UCase(ins_array_name_perf_v(i)) Like "MHV") Then
            LNH_V = "_" + ins_array_name_perf_v(i)            ''''''''''N/L/HV
        Else
             LNH_V = ""
        End If

        If (UCase(ins_array_name_perf_v(i)) Like "MC*" Or UCase(ins_array_name_perf_v(i)) Like "MS*" Or UCase(ins_array_name_perf_v(i)) Like "MG*" Or UCase(ins_array_name_perf_v(i)) Like "MA*") Then
            If (IsNumeric(Mid(ins_array_name_perf_v(i), 3, 1))) Then
                perofmrance = "_" + ins_array_name_perf_v(i) '''''''''''performance name
            Else
                perofmrance = ""
            End If
        End If
    Next i
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''match instance name, prepare conbine flag name
    For i = 0 To mbist_match(type_nu).inst_count - 1
        match_begin = match_end = 0
        confirm = False
        p = 0:: g = 0
        If (UCase(instance_name) Like UCase("*" + mbist_match(type_nu).inst_nu(i).binflag_match_name + "*")) Then
                 flag_match = True
                 Flag_Name = ""
                 flag_array_string_match = Split(mbist_match(type_nu).inst_nu(i).binflag_match_name, "_")
                 flag_array_string_inst = Split(instance_name, "_")
                 match_string_1st = ""
                 If (flag_array_string_match(0) = "" Or flag_array_string_match(0) = " ") Then
                    match_string_1st = flag_array_string_match(1)
                    g = 1
                 Else
                    match_string_1st = flag_array_string_match(0)
                    g = 0
                 End If
                 For k = 0 To UBound(flag_array_string_inst)
                    If (UCase(flag_array_string_inst(k)) Like UCase(match_string_1st)) Then
                        match_begin = k
                        confirm = True
                        g = g + 1
                    ElseIf (confirm = True And k <= UBound(flag_array_string_inst) And g <= UBound(flag_array_string_match) And flag_array_string_match(g) <> "") Then
                        If (UCase(flag_array_string_inst(k)) Like UCase(flag_array_string_match(g))) Then
                            confirm = True
                        Else
                            confirm = False
                        End If

                        If (flag_array_string_match(UBound(flag_array_string_match)) = "") Then
                            If (UCase(flag_array_string_inst(k)) Like UCase(flag_array_string_match(UBound(flag_array_string_match) - 1)) And confirm = True) Then
                                match_end = k
                                Exit For
                            End If
                        Else
                            If (UCase(flag_array_string_inst(k)) Like UCase(flag_array_string_match(UBound(flag_array_string_match))) And confirm = True) Then
                                match_end = k
                                Exit For
                            End If
                        End If

                        g = g + 1
                    End If
                 Next k

                 If (confirm = True And match_end <> 0) Then
                    For k = 0 To UBound(flag_array_string_inst)
                        If (k >= match_begin And k <= match_end) Then
                            If (p = 0 And Flag_Name <> "") Then
                                Flag_Name = Flag_Name + "_" + mbist_match(type_nu).inst_nu(i).binflag_mid_name  '//check
                            Else
                                Flag_Name = mbist_match(type_nu).inst_nu(i).binflag_mid_name
                            End If
                            p = p + 1
                        Else
                            If (Flag_Name <> "") Then
                                Flag_Name = Flag_Name + "_" + flag_array_string_inst(k)
                            Else
                                Flag_Name = flag_array_string_inst(k)
                            End If

                        End If
                    Next k

                 End If
                 Exit For
        End If
    Next i
    
    If (instance_name Like "*Mbist_*") Then
        ins_array_name_type = Split(instance_name, "_")
    End If
    
    If (instance_name Like "CpuMbist_*") Then
        ins_array_name_others = Split(instance_name, "CpuMbist_")
    ElseIf (instance_name Like "SocMbist_*") Then
        ins_array_name_others = Split(instance_name, "SocMbist_")
    End If
    
    '=================================================================================================instance name
    ins_new_name = ins_array_name_type(0) + "_" + mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(bk_loop_count) + "_" + ins_array_name_others(1)
    Block = mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(bk_loop_count)
    '===================================================Print debug information===============================================================

    'TheExec.Datalog.WriteComment "~~~~~~~ Instance match ~~~~~~~"
    If (Flag_Name <> "" And flag_match = True) Then
        'TheExec.Datalog.WriteComment "~~~~~~~ Flag match ~~~~~~~"
        If (instance_name Like "*_PP_*") Then
            ins_array_name_pp = Split(Flag_Name, "_PP_")
            for_confirm_ins_name = "PP_" + ins_array_name_pp(1)
        ElseIf (instance_name Like "*_DD_*") Then
            ins_array_name_pp = Split(Flag_Name, "_DD_")
            for_confirm_ins_name = "DD_" + ins_array_name_pp(1)
        ElseIf (instance_name Like "*_CZ_*") Then
            ins_array_name_pp = Split(Flag_Name, "_CZ_")
            for_confirm_ins_name = "CZ_" + ins_array_name_pp(1)
        Else
            ins_array_name_pp = Split(Flag_Name, "_")
            ins_array_name_pp(0) = ""
            for_confirm_ins_name = "" + ins_array_name_pp(1)
        End If
        Flag_Name = ins_array_name_pp(0) + LNH_V + "_" + mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(bk_loop_count)
    ElseIf (flag_match = False) Then
        'TheExec.Datalog.WriteComment "~~~~~~~ Flag no match ~~~~~~~"
        If (ins_new_name Like "*_PP_*") Then
            ins_array_name_pp = Split(ins_new_name, "_PP_")
            for_confirm_ins_name = "PP_" + ins_array_name_pp(1)
        ElseIf (ins_new_name Like "*_DD_*") Then
            ins_array_name_pp = Split(ins_new_name, "_DD_")
            for_confirm_ins_name = "DD_" + ins_array_name_pp(1)
        ElseIf (ins_new_name Like "*_CZ_*") Then
            ins_array_name_pp = Split(ins_new_name, "_CZ_")
            for_confirm_ins_name = "CZ_" + ins_array_name_pp(1)
        Else
            ins_array_name_pp = Split(ins_new_name, "_")
            ins_array_name_pp(0) = ""
            for_confirm_ins_name = "" + ins_array_name_pp(1)
        End If

        ins_array_name_pp = Split(ins_new_name, "_")
        flag_spilt = ins_array_name_pp(mbist_flag_set_placement + 1)
        ins_array_name_long = Split(ins_new_name, "_" + flag_spilt)
        Flag_Name = ins_array_name_long(0) + LNH_V
    Else
        'TheExec.Datalog.WriteComment "~~~~~~~ Flag conbine Erro ~~~~~~~"
        'TheExec.Flow.TestLimit -1, 0, 1, , , , unitNone, , "Test_Falg"
    End If
    
    '=========================================================================================================================================
    
    Flag_Name = "F_" + Flag_Name
    
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_NV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_LV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_HV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_MNV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_MLV", "")
    for_confirm_ins_name = Replace(for_confirm_ins_name, "_MHV", "")
    
'''    For Each Site In theExec.sites
'''        theExec.sites.Item(Site).FlagState(flag_name) = logicFalse ''''mean Pass
'''    Next Site

    Call TheExec.Datalog.SetDynamicTestName(ins_new_name, False)
    '=================================================================================================pattern
    For k = 0 To UBound(mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(bk_loop_count).pat_dynamic)
        If (Trim(for_confirm_ins_name) = Trim(mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(bk_loop_count).instance_dynamic(k))) Then
            m_pattname = mbist_dynamic.Block_dynamic(0).block_type_pat_dynamic(bk_loop_count).pat_dynamic(k)
        End If
    Next k
    '=================================================================================================patt test
    blMbistFP_Binout = EnableBinOut And gl_MbistFP_Binout
    TheHdw.Patterns(m_pattname).Load
    
    If TheExec.EnableWord("Mbist_FingerPrint") = True Then
        Call Finger_print(m_pattname, RunFailCycle, Flag_Name, True)
    Else
        Call PATT_GetPatListFromPatternSet(m_pattname, rtnPatternNames, rtnPatternCount)
        For Each patt In rtnPatternNames
            TheExec.Datalog.WriteComment "<" & ins_new_name & ">" & " dummy "
        For Each site In TheExec.sites
            tested(site) = False 'swinza move to out of pattern-loop
        Next site
            Call TheHdw.Patterns(patt).Test(pfAlways, 0, tlResultModeDomain)
            '===================================================================
            For Each site In TheExec.sites
                'testnumber(Site) = TheExec.sites.Item(Site).testnumber
                'tested(Site) = False
                blPatPass(site) = TheHdw.Digital.Patgen.PatternBurstPassed
                '-------------------------------------------------------------------------------------------------
                If blPatPass(site) = False Or alarmFail(site) = True Then   'pattern test fail or alarm
                    TheExec.sites.Item(site).FlagState(Flag_Name) = logicTrue 'pattern test fail
                    TheExec.sites.Item(site).testResult = siteFail
                    tested(site) = True
                    'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
               '-------------------------------------------------------------------------------------------------
                Else    'blPatPass(Site) = True ; pattern test pass
                    If (tested(site) = False) Then
                        If (TheExec.sites.Item(site).FlagState(Flag_Name) <> logicTrue) Then 'confirm flag is true(pattern fail)
                            TheExec.sites.Item(site).FlagState(Flag_Name) = logicFalse       'pattern test pass
                        End If
                        TheExec.sites.Item(site).testResult = sitePass
                    End If
                        'Call TheExec.Datalog.WriteFunctionalResult(Site, testnumber(Site), logTestPass, , ins_new_name)
                        'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
                End If  '' If blPatPass(Site) End
                '-------------------------------------------------------------------------------------------------
                'TheExec.Datalog.WriteComment "Instance                = " & ins_new_name
                'TheExec.Datalog.WriteComment "Pat Name                = " & m_pattname
                'TheExec.Datalog.WriteComment "Test Falg               =>" & flag_name & "(" & Site & ") = " & TheExec.sites.Item(Site).FlagState(flag_name) & ",     if pattern pass=> flag is logicFalse => 0" & ",     if pattern fail=> flag is logicTrue => 1"
                blPatPass(site) = False
                alarmFail(site) = False
            Next site
            '===================================================================
        Next patt
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
        If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_Mbist_Block_loop_inst_non_match(instance_name As String, m_pattname As String, bk_loop_count As Long, ByRef Flag_Name As String, EnableBinOut As Boolean, RunFailCycle As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Mbist_Block_loop_inst_non_match"
    Dim site As Variant
    'Dim flag_name As String
    Dim ins_array_name_pp() As String
    Dim ins_array_name_long() As String
    Dim ins_array_name_type() As String
    Dim ins_array_name_others() As String
    Dim flag_spilt As String
    
    Dim flag_array_string_match() As String
    Dim flag_array_string_inst() As String
    Dim match_begin As Long, match_end As Long:: match_begin = match_end = 0
    
    Dim ins_array_name_perf_v() As String
    Dim i As Long, k As Long, p As Long, g As Long, j As Long:: i = 0:: k = 0:: p = 0:: g = 0:: j = 0
    Dim confirm As Boolean
    
    Dim LNH_V As String
    Dim perofmrance As String
    Dim maxDepth As Integer
    Dim patt As Variant
    Dim rtnPatternNames() As String, rtnPatternCount As Long
    Dim astrPattPathSplit() As String
    Dim astrPattPathSplit_01() As String
    Dim blPatPass As New SiteBoolean
    Dim numcap As Long
    Dim PinData_d As New PinListData
    Dim Mbist_repair_cycle As Long
    Dim Pins As New PinData
    Dim Cdata As Variant
    Dim TestNumber As New SiteLong
    Dim ins_new_name As String
    Dim tested As New SiteBoolean
    Dim strPattName As String
    Dim inst_match As Boolean
    Dim Temp As Long
    Dim allpins As String
    Dim PinData As New PinListData
    
    Dim blMbistFP_Binout As Boolean
    Dim MBISTFailBlockFlag As Boolean
    Dim PassOrFail As New SiteLong
    Dim lGetFlagIdx As Long
    Dim blJump As Boolean
    Dim m_testName As String
    
    
    
    allpins = "JTAG_TDO"
    ins_array_name_perf_v = Split(instance_name, "_")
    m_testName = ins_array_name_perf_v(0)
    '=================================================================================================test flag
    For i = 0 To UBound(ins_array_name_perf_v)
        If (UCase(ins_array_name_perf_v(i)) Like "NV" Or UCase(ins_array_name_perf_v(i)) Like "LV" Or UCase(ins_array_name_perf_v(i)) Like "HV" Or UCase(ins_array_name_perf_v(i)) Like "MNV" Or UCase(ins_array_name_perf_v(i)) Like "MLV" Or UCase(ins_array_name_perf_v(i)) Like "MHV") Then
            LNH_V = "_" + ins_array_name_perf_v(i)            ''''''''''N/L/HV
        Else
             LNH_V = ""
        End If

        If (UCase(ins_array_name_perf_v(i)) Like "MC*" Or UCase(ins_array_name_perf_v(i)) Like "MS*" Or UCase(ins_array_name_perf_v(i)) Like "MG*" Or UCase(ins_array_name_perf_v(i)) Like "MA*") Then
            If (IsNumeric(Mid(ins_array_name_perf_v(i), 3, 1))) Then
                perofmrance = "_" + ins_array_name_perf_v(i) '''''''''''performance name
            Else
                perofmrance = ""
            End If
        End If
    Next i
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'TheExec.Datalog.WriteComment "~~~~~~~ Instance no match ~~~~~~~"
    
    If (instance_name Like "*Mbist_*") Then
        ins_array_name_type = Split(instance_name, "_")
    End If
    
    If (instance_name Like "CpuMbist_*") Then
        ins_array_name_others = Split(instance_name, "CpuMbist_")
    ElseIf (instance_name Like "SocMbist_*") Then
        ins_array_name_others = Split(instance_name, "SocMbist_")
    End If
    
    '=================================================================================================instance name
    ins_new_name = ins_array_name_type(0) + "_" + mbist_dynamic.Block_dynamic(0).block_count_name_dynamic(bk_loop_count) + "_" + ins_array_name_others(1)
    ins_array_name_pp = Split(ins_new_name, "_")
    flag_spilt = ins_array_name_pp(mbist_flag_set_placement + 1)
    ins_array_name_long = Split(ins_new_name, "_" + flag_spilt)
    
    Flag_Name = ins_array_name_long(0) + LNH_V
    Flag_Name = "F_" + Flag_Name

    Call TheExec.Datalog.SetDynamicTestName(ins_new_name, False)
    '=================================================================================================patt test
    blMbistFP_Binout = EnableBinOut And gl_MbistFP_Binout
    TheHdw.Patterns(m_pattname).Load
    
    If TheExec.EnableWord("Mbist_FingerPrint") = True Then
        Call Finger_print(m_pattname, RunFailCycle, Flag_Name, True)
    Else
        Call PATT_GetPatListFromPatternSet(m_pattname, rtnPatternNames, rtnPatternCount)
        For Each patt In rtnPatternNames
            TheExec.Datalog.WriteComment "<" & ins_new_name & ">" & " dummy "
        For Each site In TheExec.sites
            tested(site) = False 'swinza move to out of pattern-loop
        Next site
            Call TheHdw.Patterns(patt).Test(pfAlways, 0, tlResultModeDomain)
            '===================================================================
            For Each site In TheExec.sites
                'testnumber(Site) = TheExec.sites.Item(Site).testnumber
                'tested(Site) = False
                blPatPass(site) = TheHdw.Digital.Patgen.PatternBurstPassed
                '-------------------------------------------------------------------------------------------------
                If blPatPass(site) = False Or alarmFail(site) = True Then   'pattern test fail or alarm
                    TheExec.sites.Item(site).FlagState(Flag_Name) = logicTrue 'pattern test fail
                    TheExec.sites.Item(site).testResult = siteFail
                    tested(site) = True
                    'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
               '-------------------------------------------------------------------------------------------------
                Else    'blPatPass(Site) = True ; pattern test pass
                    If (tested(site) = False) Then
                        If (TheExec.sites.Item(site).FlagState(Flag_Name) <> logicTrue) Then 'confirm flag is true(pattern fail)
                            TheExec.sites.Item(site).FlagState(Flag_Name) = logicFalse       'pattern test pass
                        End If
                        TheExec.sites.Item(site).testResult = sitePass
                    End If
                        'Call TheExec.Datalog.WriteFunctionalResult(Site, testnumber(Site), logTestPass, , ins_new_name)
                        'TheExec.sites.Item(Site).testnumber = TheExec.sites.Item(Site).testnumber + 1
                End If  '' If blPatPass(Site) End
                '-------------------------------------------------------------------------------------------------
                'TheExec.Datalog.WriteComment "Instance                = " & ins_new_name
                'TheExec.Datalog.WriteComment "Pat Name                = " & m_pattname
                'TheExec.Datalog.WriteComment "Test Falg               =>" & flag_name & "(" & Site & ") = " & TheExec.sites.Item(Site).FlagState(flag_name) & ",     if pattern pass=> flag is logicFalse => 0" & ",     if pattern fail=> flag is logicTrue => 1"
                blPatPass(site) = False
                alarmFail(site) = False
            Next site
            '===================================================================
        Next patt
        
    End If
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
