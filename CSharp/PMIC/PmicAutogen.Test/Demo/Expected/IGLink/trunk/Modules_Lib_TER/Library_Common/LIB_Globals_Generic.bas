Attribute VB_Name = "LIB_Globals_Generic"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit
'************************************************
'         For ReadLDODynLimitRefTable
'************************************************
'Global bDynLimitRefTable As Boolean
'Global dictLDODynLimitRefTable As New Dictionary
'Global strLDODynLimitRefTableVSel() As String
'Global strLDODynLimitRefTableVout() As String
'************************************************
'     Initialize The TrimTable and Test Name
'************************************************

' put InitTestNameArray() into OnProgramStart()
' put ReadAllTrimTableByName() into OnProgramStart()

'************************************************
'                  For Test Name
'************************************************
Public Enum TEST_NAME_INDEX
    GRP1_TESTBLOCK = 0
    GRP2_PHASENUMBER = 1
    GRP3_TESTMODE = 2
    GRP4_SUBTESTMODE = 3
    GRP5_SUBTESTCONDITION = 4
    GRP6_SUBTESTCONDITION1 = 5
    GRP7_TRIMCONDITION = 6
    GRP8_LINKNUMBER = 7
    GRP9_VDDSUPPLY = 8
    GRP10_MEASURETYPE = 9
    GRP11_GNGMIN = 10
    GRP12_GNGMAX = 11
End Enum

Public Enum INSTANTANCE_NAME_INDEX
    INSTTBLOCK = 0
    INSTTPHASE = 1
    INSTTMODE = 2
End Enum

Public Enum TOGGLE_CHECK
    DTBLOW = 1
    DTBHIGH = 2
    DTBTOGGLE = 3
End Enum

Public Enum LIB_DCVI_TYPE
    LIB_DCVI_DC30 = 1
    LIB_DCVI_UVI80 = 2
End Enum

Public Const gTestNameTemplate = "X_X_X_X_X_X_X_X_X_X_X_X"
Public gArrTestName() As String
Public Const gNonToggleCode = 999

Public Const TrimCodeLength = 3

Public Const InterposeConcat = ";"

Public Const dbgUseTrimLink = True

'************************************************
'            For Check H/W Setup
'************************************************
Public Enum CHECK_DATA
    CHECK_DATA_BEFORE
    CHECK_DATA_After
End Enum

Public g_PreCheckData() As String
Public g_CurrCheckData() As String
Public g_PreFlowSheetName As String
Public g_CurrFlowSheetName As String
Public g_ResultCheckData() As Boolean

Public g_ADG1414ArgList As String
Public g_ADG1414Data(10) As Long
'************************************************
'            For Trim Table
'************************************************

Public bTrimTableReady As Boolean
Public gTrimActiveSite As New SiteBoolean
Public g_AllTrimTable() As TrimTable

Public Const c_CheckEachTable = "NOTREADY"
Public Const c_TrimDefault = "DEFAULT"
Public Const c_EndOfTrimTable = "END"
Public Const c_TrimTableStartRow = 2

Public Enum TRIM_TABLE_INDEX
    TRIM_TABLE_INDEX_CHECK = 1
    TRIM_TABLE_INDEX_TABLENAME = 2
    TRIM_TABLE_INDEX_TRIMCODE = 4
    TRIM_TABLE_INDEX_TRIMVAL = 5
    TRIM_TABLE_INDEX_PERCENT = 6
    TRIM_TABLE_INDEX_DEFAULT = 10
End Enum

Public Type TrimTable
    Trim_Val      As New DSPWave
    Trim_Per      As New DSPWave
    TableName     As String
    TargetCode    As New SiteLong
    ArrTrim_Val() As Double
    ArrTrim_Per() As Double
    PerMean       As Double
End Type

'************************************************
'         For Ton Test
'************************************************
Public Const CaptFailSize = 128

'************************************************
'         For Trim Link and Tweak Trim
'************************************************
Public LIB_TRIM_MEASVAL As New DSPWave
Public LIB_TWEAK_TRIM_VAL As New SiteDouble
Public LIB_TWEAK_TRIM_DELTA As New SiteDouble

'************************************************
'   For PostBurn to run TTR or Analogsweep
'   PostBurnTTR: True -> GNG (CP1/FT)
'                False -> AnalogSweep (CHAR/QA)
'   TrimTTR: True -> Firmware pattern only
'            False -> CodeSweep/HybridSweep
'************************************************
Public Const PostBurnTTR As Boolean = True
Public Const TrimTTR As Boolean = True

'************************************************
'         For ReadLDODynLimitRefTable
'************************************************
Global bDynLimitRefTable As Boolean
Global dictLDODynLimitRefTable As New Dictionary
Global strLDODynLimitRefTableVSel() As String
Global strLDODynLimitRefTableVout() As String

'************************************************
' flag for temp solution to control 40 relays all sites becasue some risk support board
'************************************************
Global gbRlyDefault As Boolean


'20190412 add by Endy****************************************************************************************************
'in LIB_ACORE
'Public g_sTPPath  As String ''''move to LIB_Common_GlobalConstant
'Public g_Site As Variant    ''''move to LIB_Common_GlobalConstant

'in LIB_COMMON_NWIRE
Public g_Nwire_On As Boolean
Public g_Nwire_JTAG As Boolean
Public g_Nwire_SPMI As Boolean

'From unused/LIB_Digital_Debug'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Enum Shmoo_direction_enum                                   'Used for Module(LIB_Digital_Shmoo)
    High_to_Low = 1
    Low_to_High = 2
End Enum


'''From Lib_Pool/LIB_OTP_Type_Syntax'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Enum YESNO_type                                      'Used for Module(LIB_Globals_Generic)
'    Yes = 0
'    No = 1
'End Enum


'From unused/LIB_HardIP_Dictionary''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Private gDictDSPWaves As New Dictionary                     'Used for Module(LIB_Globals_Generic)
Private gDictCurrMeasurements As New Dictionary             'Used for Module(LIB_Globals_Generic)
Private gDictSiteLong As New Dictionary                     'Used for Module(LIB_Globals_Generic)


'From unused/VBT_LIB_HardIP'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''20160729 - Use global value to denfine default setting
Public Const pc_Def_VFI_UVI80_VoltageRange = 7              'Used for Module(LIB_Digital_Shmoo_Setup,)

Public Const pc_Def_DSSC_Amplitude = 1                      'Used for Module(LIB_Globals_Generic)

Public CP_Card_RAK As New PinListData                       'Used for Module(LIB_Digital_Shmoo_Setup)
Public FT_Card_RAK As New PinListData                       'Used for Module(LIB_Digital_Shmoo_Setup)
Public CurrentJob_Card_RAK As New PinListData               'Used for Module(LIB_Digital_Shmoo_Setup)


'From unused/LIB_HardIP'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public glb_Disable_CurrRangeSetting_Print As Boolean        'Used for Module(LIB_Digital_Shmoo_Setup)


'From unused/LIB_MBIST''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Type EvaPattMbistCycleBlock                         'Used for Module(LIB_Globals_Generic)
    strBlaclName  As String
    lVector       As Long
    lCycle        As Long
    strCompare    As String
    strFlagName   As String   '' Add for MbistFP Binning 201606014 Webster
End Type

Private Type MbistCycleBlock                                'Used for Module(LIB_Globals_Generic)
    strPattName   As String
    tpMbistCycleBlock() As EvaPattMbistCycleBlock
End Type

Public tpEvaPattCycleBlockInfor() As EvaPattMbistCycleBlock    'Used for Module(LIB_Globals_Generic)
Public gl_MbistFP_Binout As Boolean                         'Used for Module(LIB_Globals_Generic)
Public tpCycleBlockInfor() As MbistCycleBlock               'Used for Module(LIB_Globals_Generic)

Private Type FlagInfo                                       'Used for Module(LIB_Globals_Generic)
    FlagName      As String
    CheckInfo     As Boolean
End Type

Public tyFlagInfoArr() As FlagInfo                          'Used for Module(LIB_Globals_Generic)
Public gS_currPayload_pattSetName As String                 'Used for Module(LIB_Globals_Generic)
Public MatchFlag  As Boolean                                 'Used for Module(LIB_Globals_Generic)


'From unused/LIB_Digital_Debug'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used for Module(LIB_Digital_Shmoo)
Public Function FailingBoundaryDatalog_Func_Multi_Power(Power_Search_String As String, _
                                                        Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                                        Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant) As Long
    'Shmoo with faillog capture version 1.0 careated by JT 2014/02/20.

    Dim PinCnt    As Long
    Dim pinary()  As String
    Dim i         As Integer
    Dim k         As Integer
    Dim j         As Integer
    Dim Fail_log_cnt As Integer
    Dim patternArray() As String
    Dim PowerV    As Double
    Dim p         As Integer
    Dim Org_Test_Number As Long
    Dim current_site As Integer
    Dim Timelist  As String
    Dim TimeGroup() As String
    Dim CurrTiming As Variant
    Dim TimeDomainlist As String
    Dim TimeDomaingroup() As String
    Dim CurrTimeDomain As Variant
    Dim TimeDomainIn As String
    On Error GoTo errHandler_faillog
    Dim funcName  As String:: funcName = "FailingBoundaryDatalog_Func_Multi_Power"
    
    '    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_pattern As String, _
         '                                     Shmoo_status As String)

    '                    Shmoo_faillog_header LotID, CStr(WaferId), CStr(g_slXCoord(site)), CStr(g_slYCoord(site)), Shmoo_pattern, "Shmoo hole"
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

    Dim Power_Pins(20) As String
    Dim Power_RangeA(20) As Double
    Dim Power_RangeB(20) As Double
    Dim Power_StepSize(20) As Double
    Dim Power_temp() As String
    Dim Power_range_temp() As String
    Dim Shmoo_steps As Double
    Dim axis_type As tlDevCharShmooAxis
    Dim SetupName As String
    Dim VmainOrValt As String


    power_list = Split(Power_Search_String, ",")
    Power_number = UBound(power_list)

    For i = 0 To Power_number
        Power_temp() = Split(power_list(i), "=")
        Power_range_temp() = Split(Power_temp(1), ":")
        Power_Pins(i) = Power_temp(0)
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
                        TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value = Power_RangeA(i) + Abs(Power_StepSize(i)) * j
                    ElseIf VmainOrValt = "valt" Then
                        TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Alt.Value = Power_RangeA(i) + Abs(Power_StepSize(i)) * j
                    End If
                Else
                    If VmainOrValt = "vmain" Then
                        TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value = Power_RangeB(i) + Abs(Power_StepSize(i)) * j
                    ElseIf VmainOrValt = "valt" Then
                        TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Alt.Value = Power_RangeB(i) + Abs(Power_StepSize(i)) * j
                    End If
                End If
            Next i

        Else

            For i = 0 To Power_number
                If Power_RangeA(i) > Power_RangeB(i) Then
                    If VmainOrValt = "vmain" Then
                        TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value = Power_RangeA(i) - Abs(Power_StepSize(i)) * j
                    ElseIf VmainOrValt = "valt" Then
                        TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Alt.Value = Power_RangeA(i) - Abs(Power_StepSize(i)) * j
                    End If
                Else
                    If VmainOrValt = "vmain" Then
                        TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value = Power_RangeB(i) - Abs(Power_StepSize(i)) * j
                    ElseIf VmainOrValt = "valt" Then
                        TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Alt.Value = Power_RangeB(i) - Abs(Power_StepSize(i)) * j
                    End If
                End If
            Next i
        End If

        For i = 0 To Power_number
            If VmainOrValt = "vmain" Then
                TheExec.Datalog.WriteComment Power_Pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value, "0.000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value, "0.000")) + ("V")
            ElseIf VmainOrValt = "valt" Then
                TheExec.Datalog.WriteComment Power_Pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Alt.Value, "0.000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Alt.Value, "0.000")) + ("V")
            End If
        Next i

        TheHdw.Wait 0.002
        TheHdw.Patterns(Shmoo_Pattern).test pfAlways, 0


        For i = 0 To Power_number
            If VmainOrValt = "vmain" Then
                TheExec.Datalog.WriteComment Power_Pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value, "0.000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value, "0.000")) + ("V")
            ElseIf VmainOrValt = "valt" Then
                TheExec.Datalog.WriteComment Power_Pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Alt.Value, "0.000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Alt.Value, "0.000")) + ("V")
            End If
        Next i


        If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False And LVCC_boundary_Switch = Fail_log_cnt Then Exit For
        If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then Fail_log_cnt = Fail_log_cnt + 1

    Next j


    TheExec.Datalog.WriteComment "***************** Shmoo fail log capture end *****************"

    'TheExec.Sites(current_site).TestNumber = Org_test_number

errHandler_faillog:
    TheExec.Datalog.WriteComment "*****************Shmoo with fail log capture error *****************"
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'Used for Module(LIB_Digital_Shmoo)
Public Function FailingDatalog_Lvcc_Boundary(Power_Search_String As String, _
                                             Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                             Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeSteps As Long, Optional RangeStepSize As Double) As Long
    'Shmoo with faillog capture version 1.0 careated by JT 2014/02/20.

    Dim PinCnt    As Long
    Dim pinary()  As String
    Dim i         As Integer
    Dim k         As Integer
    Dim j         As Integer
    Dim Fail_log_cnt As Integer
    Dim patternArray() As String
    Dim PowerV    As Double
    Dim p         As Integer
    Dim Org_Test_Number As Long
    Dim current_site As Integer
    Dim Timelist  As String
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

    Dim funcName  As String:: funcName = "FailingDatalog_Lvcc_Boundary"
    On Error GoTo errHandler_faillog
    '    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_pattern As String, _
         '                                     Shmoo_status As String)

    '                    Shmoo_faillog_header LotID, CStr(WaferId), CStr(g_slXCoord(site)), CStr(g_slYCoord(site)), Shmoo_pattern, "Shmoo hole"
    '                    Call FailingBoundaryDatalog_Func_TER(shmoo_pin_string, RangeFrom, RangeTo, CDbl(RangeStepSize))
    '                    Call Shmoo_faillog_ending


    'Shmoo_faillog_test_number
    'TheExec.Sites(site).TestNumber = 100
    'current_site = TheExec.Sites.SiteNumber

    'Org_test_number = TheExec.Sites(CurrSite).TestNumber

    'TheExec.Sites(current_site).TestNumber = Shmoo_faillog_test_number

    ShmooPatternSplit() = Split(Shmoo_Pattern, ",")
    inst_name = UCase(TheExec.DataManager.InstanceName)

    Dim Context   As String: Context = ""
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

    Dim Power_Pins() As String
    Dim Power_RangeA(20) As Double
    Dim Power_RangeB(20) As Double
    Dim Power_StepSize(20) As Double
    Dim Power_temp() As String
    Dim Power_range_temp() As String
    Dim Shmoo_steps As Double
    Dim StepValue As Double

    power_list = Split(Power_Search_String, ",")
    Power_number = UBound(power_list)
    ReDim Power_Pins(Power_number)

    For i = 0 To Power_number
        Power_temp() = Split(power_list(i), "=")
        Power_range_temp() = Split(Power_temp(1), ":")
        Power_Pins(i) = Power_temp(0)
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
                TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value = Shmoo_Vcc_Min(CurrSite) - 0.005
            Else
                TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value = Shmoo_Vcc_Min(CurrSite) - RangeStepSize * (j + 1)
            End If
        Next i

        If TheExec.Flow.EnableWord("FailPinsOnly") = False Then

            StepValue = Fail_log_cnt * 3.125

            TheExec.Datalog.WriteComment "Power setup (Vmin- " & StepValue & "mV) "
            For i = 0 To Power_number
                TheExec.Datalog.WriteComment Power_Pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value, "0.00000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value, "0.00000")) + ("V")
            Next i
        End If

        TheHdw.Wait 0.002

        If TheExec.Flow.EnableWord("FailPinsOnly") = True Then
            For i = 0 To UBound(ShmooPatternSplit)
                Call TheHdw.Patterns(ShmooPatternSplit(i)).Load
                Call TheHdw.Patterns(ShmooPatternSplit(i)).Start
                TheHdw.Digital.Patgen.HaltWait
                If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then
                    shmoopowerpin = Join(Power_Pins, ",")
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
                TestNumber = TheExec.Sites.Item(CurrSite).TestNumber
                TheHdw.Patterns(tempPat).test pfNever, 0
                If (TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = True) Then
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestPass)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestFail)
                End If
                TestNumber = TestNumber + 1
                TheExec.Sites.Item(CurrSite).TestNumber = TestNumber
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
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'Used for Module(LIB_Digital_Shmoo)
Public Function FailingDatalog_Hvcc_Boundary(Power_Search_String As String, _
                                             Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_Pattern As String, _
                                             Shmoo_status As String, Direction As Shmoo_direction_enum, CurrSite As Variant, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeSteps As Long, Optional RangeStepSize As Double) As Long
    'Shmoo with faillog capture version 1.0 careated by JT 2014/02/20.

    Dim PinCnt    As Long
    Dim pinary()  As String
    Dim i         As Integer
    Dim k         As Integer
    Dim j         As Integer
    Dim Fail_log_cnt As Integer
    Dim patternArray() As String
    Dim PowerV    As Double
    Dim p         As Integer
    Dim Org_Test_Number As Long
    Dim current_site As Integer
    Dim Timelist  As String
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
    Dim funcName  As String:: funcName = "FailingDatalog_Hvcc_Boundary"
    
    
    '    Shmoo_LotID As String, Shmoo_wafer As String, Shmoo_X As String, Shmoo_Y As String, Shmoo_pattern As String, _
         '                                     Shmoo_status As String)

    '                    Shmoo_faillog_header LotID, CStr(WaferId), CStr(g_slXCoord(site)), CStr(g_slYCoord(site)), Shmoo_pattern, "Shmoo hole"
    '                    Call FailingBoundaryDatalog_Func_TER(shmoo_pin_string, RangeFrom, RangeTo, CDbl(RangeStepSize))
    '                    Call Shmoo_faillog_ending


    'Shmoo_faillog_test_number
    'TheExec.Sites(site).TestNumber = 100
    'current_site = TheExec.Sites.SiteNumber

    'Org_test_number = TheExec.Sites(CurrSite).TestNumber

    'TheExec.Sites(current_site).TestNumber = Shmoo_faillog_test_number

    ShmooPatternSplit() = Split(Shmoo_Pattern, ",")
    inst_name = UCase(TheExec.DataManager.InstanceName)

    Dim Context   As String: Context = ""
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

    Dim Power_Pins() As String
    Dim Power_RangeA(20) As Double
    Dim Power_RangeB(20) As Double
    Dim Power_StepSize(20) As Double
    Dim Power_temp() As String
    Dim Power_range_temp() As String
    Dim Shmoo_steps As Double
    Dim StepValue As Double

    power_list = Split(Power_Search_String, ",")
    Power_number = UBound(power_list)
    ReDim Power_Pins(Power_number)

    For i = 0 To Power_number
        Power_temp() = Split(power_list(i), "=")
        Power_range_temp() = Split(Power_temp(1), ":")
        Power_Pins(i) = Power_temp(0)
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
                TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value = Shmoo_Vcc_Max(CurrSite) + 0.005
            Else
                TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value = Shmoo_Vcc_Max(CurrSite) + RangeStepSize * (j + 1)
            End If
        Next i

        If TheExec.Flow.EnableWord("FailPinsOnly") = False Then

            StepValue = Fail_log_cnt * 3.125

            TheExec.Datalog.WriteComment "Power setup (Vmax+ " & StepValue & "mV) "
            For i = 0 To Power_number
                TheExec.Datalog.WriteComment Power_Pins(i) + "= " & CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value, "0.00000"))
                TheExec.Datalog.WriteComment ("VOLTAGE=") + CStr(Format(TheHdw.DCVS.Pins(Power_Pins(i)).Voltage.Main.Value, "0.00000")) + ("V")
            Next i
        End If

        TheHdw.Wait 0.002

        If TheExec.Flow.EnableWord("FailPinsOnly") = True Then
            For i = 0 To UBound(ShmooPatternSplit)
                Call TheHdw.Patterns(ShmooPatternSplit(i)).Load
                Call TheHdw.Patterns(ShmooPatternSplit(i)).Start
                TheHdw.Digital.Patgen.HaltWait
                If TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = False Then
                    shmoopowerpin = Join(Power_Pins, ",")
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
                TestNumber = TheExec.Sites.Item(CurrSite).TestNumber
                TheHdw.Patterns(tempPat).test pfNever, 0
                If (TheHdw.Digital.Patgen.PatternBurstPassed(CurrSite) = True) Then
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestPass)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(CurrSite, TestNumber, logTestFail)
                End If
                TestNumber = TestNumber + 1
                TheExec.Sites.Item(CurrSite).TestNumber = TestNumber
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
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'From unused/LIB_HardIP_Dictionary'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used for Module(Exec_IP_Module)
'Public Function RemoveAllStored()
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "RemoveAllStored"
'
'    gDictCurrMeasurements.RemoveAll
'    gDictDSPWaves.RemoveAll
'    gDictSiteLong.RemoveAll
'
'    Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'From unused/LIB_HardIP'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used for Module(Exec_IP_Module)
Public Function SetupDatalogFormat(TestNameW As Integer, PatternW As Integer)
    'Init_Datalog_Setup
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SetupDatalogFormat_Test"

    If TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.TestName.Width < TestNameW Then
        TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.TestName.Width = TestNameW    '70
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.TestName.Width = TestNameW    '70
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = PatternW      '102
        '        TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
    End If

    TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
    TheExec.Datalog.Setup.DatalogSetup.DisableChannelNumberInPTR = True
    TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = True
    TheExec.Datalog.ApplySetup

    Exit Function

ErrHandler:
    TheExec.AddOutput "<Error> " + funcName + ":: please check it out."
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'Used for Module(VBT_LIB_DC_IDS, LIB_Digital_Shmoo)
Public Function HardIP_WriteFuncResult(Optional SpecialReserve As String = "", Optional CodeSearchPatternResult As SiteBoolean, Optional m_TestName As String = "") As Long
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "HardIP_WriteFuncResult"


    Dim Site      As Variant
    Dim TestNumber As Long
    Dim FailCount As New PinListData
    Dim Allpins   As PinList
    Dim Pin       As Variant
    Dim Pins()    As String
    Dim Pin_Cnt   As Long

    '' 20150604: Need to modify "All_Digital" to the parameter.
    TheExec.DataManager.DecomposePinList "All_Digital", Pins(), Pin_Cnt

    If SpecialReserve <> "" Then
        If SpecialReserve = "DSSC_CODESEARCH" Then
            For Each Site In TheExec.Sites
                TestNumber = TheExec.Sites.Item(Site).TestNumber
                If CodeSearchPatternResult(Site) Then
                    If TheExec.DevChar.Setups.IsRunning = True Then TheExec.Sites.Item(Site).TestResult = sitePass

                    ''''20151106 update
                    If (m_TestName <> "") Then
                        Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestPass, , m_TestName)
                    Else
                        Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestPass)
                    End If
                Else
                    ''''20151106 update
                    If (m_TestName <> "") Then
                        Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestFail, , m_TestName)
                    Else
                        Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestFail)
                    End If
                    '' 20160218 - Modify sequence to let TestResult after WriteFunctionalResult to cover test number increment 2 issue if souce sink time out alarm happen.
                    TheExec.Sites.Item(Site).TestResult = siteFail
                End If
                TheExec.Sites.Item(Site).TestNumber = TestNumber + 1
            Next Site
        End If
    Else

        Dim patPassed As New SiteBoolean
        patPassed = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        For Each Site In TheExec.Sites
            TestNumber = TheExec.Sites.Item(Site).TestNumber
            Exit For
        Next Site

        For Each Site In TheExec.Sites
            If patPassed Then
                TheExec.Sites.Item(Site).TestResult = sitePass
                If TheExec.DevChar.Setups.IsRunning = True Then TheExec.Sites.Item(Site).TestResult = sitePass
                ''''20151106 update
                If (m_TestName <> "") Then
                    Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestPass, , m_TestName)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestPass)
                End If
            Else
                ''''20151106 update
                If (m_TestName <> "") Then
                    Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestFail, , m_TestName)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(Site, TestNumber, logTestFail)
                End If
                '' 20160218 - Modify sequence to let TestResult after WriteFunctionalResult to cover test number increment 2 issue if souce sink time out alarm happen.
                TheExec.Sites.Item(Site).TestResult = siteFail

            End If

            If TheExec.DevChar.Setups.IsRunning = False Then TheExec.Sites.Item(Site).TestNumber = TestNumber + 1

            '20180607 TER **************************************************************************************
            If TheExec.DevChar.Setups.IsRunning = True Then
                Dim SetupName As String

                SetupName = TheExec.DevChar.Setups.ActiveSetupName
                If Not ((TheExec.DevChar.Results(SetupName).StartTime Like "1/1/0001*" Or TheExec.DevChar.Results(SetupName).StartTime Like "0001/1/1*")) Then
                    With TheExec.DevChar.Setups(SetupName)
                        If .Shmoo.Axes.Count > 1 Then
                            XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                            YVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
                            If gl_flag_end_shmoo = True Then
                                TheExec.Sites.Item(Site).TestNumber = TestNumber + 1
                            End If
                        Else
                            XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                            If gl_flag_end_shmoo = True Then
                                TheExec.Sites.Item(Site).TestNumber = TestNumber + 1
                            End If
                        End If
                    End With
                End If
            End If
            '*****************************************************************************************************
        Next Site
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'Used for Module(VBT_LIB_Digital_Functional_T, LIB_Digital_Shmoo)
Public Function GeneralDigCapSetting(Pat As String, DigCap_Pin As PinList, DigCap_Sample_Size As Long, ByRef OutDspWave As DSPWave) As Long
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "GeneralDigCapSetting"


    Dim i         As Long

    If DigCap_Sample_Size <> 0 Then

        Dim Str_FinalPatName As String
        Str_FinalPatName = ""
        Call AnalyzePatName(Pat, Str_FinalPatName)

        ''        Dim DigCap_Pin_Num As Integer
        ''        DigCap_Pin_Num = UBound(DigCap_Pin_Ary)
        ''        ReDim OutDspWave(DigCap_Pin_Num) As New DSPWave

        ''        For i = 0 To DigCap_Pin_Num
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test Start ========")
        OutDspWave.CreateConstant 0, DigCap_Sample_Size
        ''            DigCap_Pin.Value = DigCap_Pin_Ary(i)
        Call DigCapSetup(Pat, DigCap_Pin, Str_FinalPatName, DigCap_Sample_Size, OutDspWave)
        ''          Call DigCapSetup(Pat, DigCap_Pin, "S" & CStr(PatCount), DigCap_Sample_Size, OutDspWave)
        ''       Next i
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'Used for Module(LIB_Globals_Generic)
Public Function AnalyzePatName(Pat As String, ByRef Str_FinalPatName As String) As Long
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "AnalyzePatName"

    Dim Str_Before_UnderLine As String, Str_After_UnderLine As String
    Dim pat_name() As String
    Dim pat_name_module() As String
    Dim Pat_name1() As String

    pat_name_module = Split(Pat, ":")
    pat_name = Split(pat_name_module(0), "\")

    pat_name(0) = pat_name(UBound(pat_name))
    pat_name(0) = Replace(pat_name(0), ".", "_")
    Pat_name1 = Split(TheExec.DataManager.InstanceName, "_")

    Str_Before_UnderLine = pat_name(0)
    Str_After_UnderLine = Pat_name1(UBound(Pat_name1))

    Str_FinalPatName = Str_Before_UnderLine & "_" & Str_After_UnderLine

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'Used for Module(LIB_Digital_Shmoo)
Public Function CreateSimulateDataDSPWave(OutDspWave As DSPWave, DigCap_Sample_Size As Long, DigCap_DataWidth As Long)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "CreateSimulateDataDSPWave"

    Dim Site      As Variant
    Dim i         As Integer
    Dim TempStr_DSP As New SiteVariant

    If TheExec.TesterMode = testModeOffline Then
        If DigCap_DataWidth <> 0 Then
            For Each Site In TheExec.Sites
                For i = 0 To DigCap_Sample_Size - 1
                    If Site = 0 Then
                        If i Mod 3 = 0 Then
                            OutDspWave(Site).Element(i) = 1
                        Else
                            OutDspWave(Site).Element(i) = 0
                        End If
                    Else
                        If i Mod 2 = 0 Then
                            OutDspWave(Site).Element(i) = 0
                        Else
                            OutDspWave(Site).Element(i) = 1
                        End If

                    End If
                Next i
            Next Site
        Else
            For Each Site In TheExec.Sites
                If Site = 0 Then
                    OutDspWave(Site).Element(0) = 1
                    OutDspWave(Site).Element(1) = 1
                    OutDspWave(Site).Element(2) = 1
                    OutDspWave(Site).Element(3) = 1
                    OutDspWave(Site).Element(4) = 1
                    OutDspWave(Site).Element(5) = 0

                    'OutDspWave(site).Element(13) = 1
                    'OutDspWave(site).Element(14) = 1
                    'OutDspWave(site).Element(15) = 0

                    ''                    OutDspWave(Site).Element(16) = 1
                    ''                    OutDspWave(Site).Element(17) = 1
                    ''                    OutDspWave(Site).Element(18) = 1
                    ''                    OutDspWave(Site).Element(19) = 1
                    ''                    OutDspWave(Site).Element(20) = 1
                    ''                    OutDspWave(Site).Element(21) = 0
                    ''
                    ''                    OutDspWave(Site).Element(32) = 1
                    ''                    OutDspWave(Site).Element(33) = 1
                    ''                    OutDspWave(Site).Element(34) = 1
                    ''                    OutDspWave(Site).Element(35) = 1
                    ''                    OutDspWave(Site).Element(36) = 1
                    ''                    OutDspWave(Site).Element(37) = 0

                Else
                    OutDspWave(Site).Element(0) = 1
                    OutDspWave(Site).Element(1) = 0
                    OutDspWave(Site).Element(2) = 0
                    OutDspWave(Site).Element(3) = 0
                    OutDspWave(Site).Element(4) = 0
                    OutDspWave(Site).Element(5) = 1

                    'OutDspWave(site).Element(13) = 0
                    'OutDspWave(site).Element(14) = 0
                    'OutDspWave(site).Element(15) = 1
                    ''                    OutDspWave(Site).Element(16) = 1
                    ''                    OutDspWave(Site).Element(17) = 0
                    ''                    OutDspWave(Site).Element(18) = 0
                    ''                    OutDspWave(Site).Element(19) = 0
                    ''                    OutDspWave(Site).Element(20) = 0
                    ''                    OutDspWave(Site).Element(21) = 1
                    ''
                    ''                    OutDspWave(Site).Element(32) = 1
                    ''                    OutDspWave(Site).Element(33) = 0
                    ''                    OutDspWave(Site).Element(34) = 0
                    ''                    OutDspWave(Site).Element(35) = 0
                    ''                    OutDspWave(Site).Element(36) = 0
                    ''                    OutDspWave(Site).Element(37) = 1
                End If
            Next Site
        End If

        For Each Site In TheExec.Sites
            For i = 0 To OutDspWave(Site).SampleSize - 1
                TempStr_DSP(Site) = TempStr_DSP(Site) & CStr(OutDspWave(Site).Element(i))
            Next i

        Next Site
        If gl_Disable_HIP_debug_log = False Then
            For Each Site In TheExec.Sites
                TheExec.Datalog.WriteComment ("Site_" & Site & " simulate data = " & TempStr_DSP(Site))
            Next Site
        End If

    End If
    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Used for Module(LIB_Digital_Shmoo)
Public Function HardIP_Freq_MeasFreqStart(Pin As PinList, Interval As Double, ByRef freq As PinListData, Optional CustomizeWaitTime As String)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "HardIP_Freq_MeasFreqStart"

    '' 20150918 - Check whether have duplicated pins, change measure mrthod from parallel to serial if pin name duplicate.
    Dim FlagDuplicatePins As Boolean
    FlagDuplicatePins = CheckDuplicateInputPins(CStr(Pin))


    Dim CounterValue As New PinListData
    Dim Site      As Variant

    ''20150623 - Add CustomizeWaitTime
    If CustomizeWaitTime <> "" Then
        TheHdw.Wait (CDbl(CustomizeWaitTime))
    End If

    TheHdw.Digital.Pins(Pin).FreqCtr.Clear
    TheHdw.Digital.Pins(Pin).FreqCtr.Start

    '' 20150918 - Check whether have duplicated pins, change measure mrthod from parallel to serial if pin name duplicate.
    Dim i         As Long
    Dim InputPins() As String
    InputPins = Split(Pin, ",")
    If FlagDuplicatePins = True Then
        For i = 0 To UBound(InputPins)
            CounterValue.AddPin(InputPins(i)).Value = TheHdw.Digital.Pins(InputPins(i)).FreqCtr.Read

        Next i
    Else
        CounterValue = TheHdw.Digital.Pins(Pin).FreqCtr.Read
        ''        freq = CounterValue.Math.Divide(interval)
    End If
    freq = CounterValue.Math.Divide(Interval)

    ''    ''20150623 - Remove site loop
    ''''    For Each Site In TheExec.Sites
    ''        CounterValue = TheHdw.Digital.Pins(pin).FreqCtr.Read
    ''        freq = CounterValue.Math.Divide(interval)
    ''''    Next Site
    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



'Used for Module(LIB_Globals_Generic)
Public Function CheckDuplicateInputPins(CheckPins As String) As Boolean
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "CheckDuplicateInputPins"

    '' 20150918 - Check whether have duplicated pins, change measure mrthod from parallel to serial if pin name duplicate.
    Dim ActualPins() As String
    Dim ActtualNumberPins As Long
    Call TheExec.DataManager.DecomposePinList(CheckPins, ActualPins(), ActtualNumberPins)

    Dim i         As Long
    Dim InputPins() As String
    Dim Pins()    As String
    Dim InputPinsNum As Long
    Dim TotalInputPinsNum As Long
    Dim SinglePinFlag As Boolean
    SinglePinFlag = True
    InputPins = Split(CheckPins, ",")

    For i = 0 To UBound(InputPins)
        Call TheExec.DataManager.DecomposePinList(InputPins(i), Pins(), InputPinsNum)
        If InputPinsNum <> 1 Then
            SinglePinFlag = False
        End If
        TotalInputPinsNum = TotalInputPinsNum + InputPinsNum
    Next i

    If ActtualNumberPins <> TotalInputPinsNum Then
        If SinglePinFlag = True Then
            CheckDuplicateInputPins = True
        Else
            CheckDuplicateInputPins = False
            TheExec.AddOutput ("Check input pins whether duplicated")
            TheExec.Datalog.WriteComment ("Check input pins whether duplicated")
        End If
    Else
        CheckDuplicateInputPins = False
    End If
    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'Used for Module(LIB_Digital_Shmoo_Setup)
Public Function Merge_TName(TName_Ary() As String) As String
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Merge_TName"

    Dim i         As Integer
    Dim TName     As String: TName = ""

    For i = 0 To UBound(TName_Ary)
        If TName_Ary(i) = "" Then TName_Ary(i) = "X"
        TName_Ary(i) = Replace(TName_Ary(i), "_", "")

        If TName <> "" Then
            TName = TName & TName_Ary(i) & "_"
        Else
            TName = TName_Ary(i) & "_"
        End If
    Next i


    'Merge_TName = Tname
    Merge_TName = TName    '& "(" & Replace(gl_TName_Pat, "_", "-") & ")"

    'If AbortTest Then Exit Function Else Resume Next
    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'From unused/LIB_MBIST''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used for Module(VBT_LIB_Digital_Functional_T)
'=======================20160301=======================================
Public Function auto_FuncTest_Mbist_ExecuteForShowFailBlock(m_pattname As String, EnableBinout As Boolean) As Long

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "auto_FuncTest_Mbist_ExecuteForShowFailBlock"

    ''''-------------------------------------------------------------------------------------------------
    ''''20151020 Update (Check for Mbist Function)
    ''''-------------------------------------------------------------------------------------------------
    Dim m_tn      As Long
    Dim Site      As Variant
    ''''-------------------------------------------------------------------------------------------------
    ''''20160301 Update (Check for MBISTFailBlock)
    ''''-------------------------------------------------------------------------------------------------
    Dim numcap    As Long
    Dim numPrecap As Long
    Dim Mbist_repair_cycle As Long
    Dim k As Long, Count As Long, j As Long

    Dim TestPatName As String, rtnPatternNames() As String, rtnPatternCount As Long
    Dim PatternNamesArray() As String

    Dim mem_location As String, i As Long
    Dim InstanceName As String
    Dim maxDepth  As Integer
    '    Dim Shift_Pat As Pattern
    Dim patt      As Variant
    Dim PatternName As String
    Dim PassOrFail As New SiteLong
    Dim MBISTFailBlockFlag As Boolean
    Dim PMAndBlock As String
    Dim Allpins   As String
    Dim blJump    As Boolean
    Dim Pins      As New PinListData
    Dim PinData   As New PinListData
    Dim Cdata     As Variant
    Dim Temp      As Long
    Dim m_TestName As String
    '    Dim IsLargeThanMaxDepth As Boolean
    Dim FailCount As New PinListData
    Dim blPatPass As New SiteBoolean
    Dim lFlagIdx  As Long
    Dim astrPattPathSplit() As String
    Dim strPattName As String
    Dim blMbistFP_Binout As Boolean
    Dim lGetFlagIdx As Long

    blMbistFP_Binout = EnableBinout And gl_MbistFP_Binout       '' 20160629  webster


    ''''-------------------------------------------------------------------------------------------------
    ''''20151102, Reset F_Payload Every time before runing payload
    Dim m_flagname As String
    Count = 0
    m_flagname = "F_Payload"
    Allpins = "JTAG_TDO"
    '    Shift_Pat = m_pattname
    For Each Site In TheExec.Sites.Existing
        TheExec.Sites.Item(Site).FlagState(m_flagname) = logicFalse    ''''mean Pass
    Next Site
    gS_currPayload_pattSetName = m_pattname    ''''for SONE datalog
    ''''-------------------------------------------------------------------------------------------------
    m_TestName = TheExec.DataManager.InstanceName
    InstanceName = LCase(m_TestName)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    maxDepth = TheHdw.Digital.HRAM.maxDepth
    TheHdw.Digital.HRAM.Size = maxDepth
    TheHdw.Digital.HRAM.CaptureType = captFail
    ''PMAndBlock = Mid(m_testName, InStr(m_testName, "_MC") + 1, 9)
    ''    GetPatListFromPatternSet m_pattname, rtnPatternNames, rtnPatternCount
    Call PATT_GetPatListFromPatternSet(m_pattname, rtnPatternNames, rtnPatternCount)

    blPatPass = True    '' 20160629
    For Each patt In rtnPatternNames
        TheHdw.Digital.HRAM.SetTrigger trigFail, False, 0, True
        TheHdw.Digital.Patgen.ClearFail

        astrPattPathSplit = Split(CStr(patt), "\")
        strPattName = UCase(astrPattPathSplit(UBound(astrPattPathSplit)))
        If strPattName Like "*.GZ" Then strPattName = Replace(strPattName, ".GZ", "")

        MatchFlag = False
        For Temp = 0 To UBound(tpCycleBlockInfor)
            If UCase(tpCycleBlockInfor(Temp).strPattName) = strPattName Then
                tpEvaPattCycleBlockInfor = tpCycleBlockInfor(Temp).tpMbistCycleBlock
                MatchFlag = True
                Exit For
            End If
        Next Temp

        TheHdw.Patterns(patt).Start
        TheHdw.Digital.Patgen.HaltWait


        ''blPatPass = TheHdw.Digital.Patgen.PatternBurstPassed
        numcap = TheHdw.Digital.HRAM.CapturedCycles

        '    FailCount = TheHdw.Digital.Pins(AllPins).FailCount         '' webster add 20160428

        For Each Site In TheExec.Sites

            m_tn = TheExec.Sites.Item(Site).TestNumber

            If TheHdw.Digital.Patgen.PatternBurstPassed(Site) = True Then
                Call TheExec.Datalog.WriteFunctionalResult(Site, m_tn, logTestPass, , m_TestName)

                If blPatPass(Site) <> False Then  '''20160624 for pattern group, PrePattern not fail
                    TheExec.Sites.Item(Site).TestResult = sitePass  ''''20160506
                End If

                blPatPass(Site) = True

                '            If (UCase(m_BinFlagName) <> UCase("Default")) Then
                '                If (TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicTrue) Then
                '                    ''''<Important>
                '                    ''''Because it was Failed on previous test, so it will NOT do any change here.
                '                Else
                '                    TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicFalse ''''mean Pass
                '                End If
                '            End If
            Else
                ''''Fail/Alarm Case
                Call TheExec.Datalog.WriteFunctionalResult(Site, m_tn, logTestFail, , m_TestName)

                TheExec.Sites.Item(Site).TestResult = siteFail    ''''20151112 update
                blPatPass(Site) = False
                '            If (UCase(m_BinFlagName) <> UCase("Default")) Then
                '                TheExec.Sites.Item(Site).FlagState(m_BinFlagName) = logicTrue ''''mean Fail
                '            End If
                '            ''''20151102, for SONE
                '            TheExec.Sites.Item(Site).FlagState(m_flagname) = logicTrue ''''mean Fail
            End If

            TheExec.Sites.Item(Site).TestNumber = m_tn + 1

        Next Site

        If MatchFlag = False Then
            TheExec.Datalog.WriteComment ("Warning!! Pattern Name not match ")
            Exit Function
        End If

        If MatchFlag And blMbistFP_Binout Then
            For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                    lGetFlagIdx = GetFlagInfoArrIndex(tpEvaPattCycleBlockInfor(k).strFlagName)
                    If lGetFlagIdx >= 0 Then
                        tyFlagInfoArr(lGetFlagIdx).CheckInfo = True
                    End If
                End If
            Next k
        End If

        For Each Site In TheExec.Sites
            If blPatPass(Site) = False Then     ''  patt fail
                For i = 0 To numcap - 1
                    ''  For i = 0 To FailCount(Site) - 1 ''  webster add 20160428
                    Set PinData = TheHdw.Digital.Pins(Allpins).HRAM.PinData(i)
                    Mbist_repair_cycle = TheHdw.Digital.HRAM.PatGenInfo(i, pgCycle)
                    'Mbist_repair_cycle = Mbist_repair_cycle + 1 'no shift
                    '                mem_location = "Not Match"
                    'Array selection
                    For Each Pins In PinData.Pins
                        Cdata = Pins.Value(Site)
                        If InstanceName Like "*bist*" Then
                            For j = 0 To UBound(tpEvaPattCycleBlockInfor)
                                If Mbist_repair_cycle = tpEvaPattCycleBlockInfor(j).lCycle Then
                                    MBISTFailBlockFlag = True
                                    Exit For
                                End If
                            Next j
                            If MBISTFailBlockFlag Then
                                MBISTFailBlockFlag = False
                                For k = Count To UBound(tpEvaPattCycleBlockInfor)

                                    If Mbist_repair_cycle = tpEvaPattCycleBlockInfor(k).lCycle Then
                                        If tpEvaPattCycleBlockInfor(k).strCompare <> Cdata Then
                                            PassOrFail(Site) = 0
                                            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                                            End If
                                        Else
                                            PassOrFail(Site) = 1
                                            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                                If TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                    TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                                End If
                                            End If
                                        End If
                                        blJump = True
                                        Count = k + 1
                                    Else
                                        PassOrFail(Site) = 1
                                        If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                            If TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                                TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                            End If
                                        End If
                                    End If

                                    TheExec.Flow.TestLimit PassOrFail(Site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
                                                           tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                                    If blJump = True Then
                                        blJump = False
                                        Exit For
                                    End If
                                Next k
                            End If
                        End If
                    Next Pins
                Next i
                '''            ''' ===========================
                '''            If numcap = maxDepth Then  '' HRAM is full
                '''                TheHdw.Digital.Patgen.HaltMode = tlHaltOnHRAMFull
                '''                TheHdw.Digital.HRAM.Size = maxDepth
                '''                TheHdw.Digital.HRAM.CaptureType = captFail
                '''                TheHdw.Digital.HRAM.SetTrigger trigFail, False, 0, True
                '''
                '''' '''               TheHdw.Digital.Patgen.Events.SetCycleCount True, Mbist_repair_cycle + 1 ' from last fail cycle+1
                '''                TheHdw.Digital.Patgen.MaskTilCycle = True
                '''
                '''                For Each patt In rtnPatternNames
                '''                    TheHdw.Patterns(m_pattname).start
                '''                    TheHdw.Digital.Patgen.HaltWait
                '''                Next patt
                '''                Dim numcap_1 As Long
                '''                numcap_1 = TheHdw.Digital.HRAM.CapturedCycles
                '''                If numcap_1 Then
                '''                    IsLargeThanMaxDepth = True
                '''                Else
                '''                    IsLargeThanMaxDepth = False
                '''                End If
                '''                TheHdw.Digital.Patgen.MaskTilCycle = False
                '''            End If
                '''
                '''            If k < UBound(tpEvaPattCycleBlockInfor) Then
                '''                If IsLargeThanMaxDepth = False Then
                '''                    PassOrFail(Site) = 1
                '''                    For k = Count To UBound(tpEvaPattCycleBlockInfor)
                '''                        Theexec.Flow.TestLimit PassOrFail(Site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
                 '''                                tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                '''                    Next k
                '''                Else
                '''                    Theexec.Datalog.WriteComment ("Warning!! The pattern fail cycle exceed HRAM maxDepth: " & maxDepth)
                '''                End If       ' If IsLargeThanMaxDepth
                '''            End If
                '''
                '''            If IsLargeThanMaxDepth = False Then
                '''                Theexec.Flow.TestLimit 1, 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
                 '''                                           "Pattern_fail_cycle_size_check", , , , , tlForceNone
                '''            Else
                '''                Theexec.Flow.TestLimit 0, 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
                 '''                                           "Pattern_fail_cycle_size_check", , , , , tlForceNone
                '''            End If

                If k < UBound(tpEvaPattCycleBlockInfor) Then        '' in unread all info of  tpEvaPattCycleBlockInfor case
                    If numcap < maxDepth Then
                        PassOrFail(Site) = 1
                        For k = Count To UBound(tpEvaPattCycleBlockInfor)
                            If blMbistFP_Binout And tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                                If TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                    TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                                End If
                            End If
                            TheExec.Flow.TestLimit PassOrFail(Site), 0.5, 1.5, , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
                                                   tpEvaPattCycleBlockInfor(k).strBlaclName & " " & tpEvaPattCycleBlockInfor(k).lCycle, , , , , tlForceNone
                        Next k
                    Else
                        '' add for HRAM is full and still have some cycles need to judge, to set all flag status = true
                        If gl_MbistFP_Binout Then
                            For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                                If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicTrue
                            Next k
                        End If
                    End If
                End If

                If numcap >= maxDepth Then   '' HRAM is full
                    TheExec.Flow.TestLimit 0, , , , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
                                           "Fail_cycle_size_check", , , , , tlForceNone
                    TheExec.Datalog.WriteComment ("The number of pattern fail cycles full or exceed HRAM maxDepth: " & maxDepth)
                Else
                    TheExec.Flow.TestLimit 1, , , , , scaleNoScaling, , , "MemFP_" & m_TestName, , _
                                           "Fail_cycle_size_check", , , , , tlForceNone
                End If


                Count = 0
                k = 0
                '            IsLargeThanMaxDepth = False
                '        End If

            Else    ''blPatPass(Site) = True
                If blMbistFP_Binout Then
                    For k = 0 To UBound(tpEvaPattCycleBlockInfor)
                        If tpEvaPattCycleBlockInfor(k).strFlagName <> "" Then
                            If TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) <> logicTrue Then
                                TheExec.Sites.Item(Site).FlagState(tpEvaPattCycleBlockInfor(k).strFlagName) = logicFalse
                            End If
                        End If
                    Next k
                End If
            End If    '' If blPatPass(Site)
        Next Site
    Next patt

    '' TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode 'recover halt on opcode
    '    TheExec.Flow.IncrementTestNumber
    ''''-------------------------------------------------------------------------------------------------

    '''    Dim m_instname As String
    '''    Dim FlowTestName() As String
    '''    m_instname = TheExec.DataManager.InstanceName
    '''    If UCase(m_instname) Like UCase("*RING*") Then
    '''        Dim MeasF_Pin As New PinList
    '''        MeasF_Pin.Value = "RINGS_RO_CLK_OUT"
    '''        'Call HardIP_FrequencyMeasure(MeasureF_Pin_SingleEnd, False, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, Pat, Flag_SingleLimit, d_MeasF_Interval, MeasF_WaitTime, MeasF_EventSource)
    '''         Call HardIP_FrequencyMeasure(MeasF_Pin, False, "FFF", 0, 0, 0, m_pattname, True, 0.01, FlowTestName)
    '''    End If
    auto_FuncTest_Mbist_ExecuteForShowFailBlock = 1

    Exit Function

ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function



'Used for Module(LIB_Globals_Generic)
Public Function GetFlagInfoArrIndex(FlagName As String) As Long
    On Error GoTo ErrHandler
    Dim funcName  As String: funcName = "GetFlagInfoArrIndex"
    Dim lIdxTemp  As Long

    GetFlagInfoArrIndex = -1

    For lIdxTemp = 0 To UBound(tyFlagInfoArr)
        If tyFlagInfoArr(lIdxTemp).FlagName = FlagName Then
            GetFlagInfoArrIndex = lIdxTemp
            Exit For
        End If
    Next lIdxTemp

    If GetFlagInfoArrIndex = -1 Then
        TheExec.Datalog.WriteComment "<Warnning> the flag(" & FlagName & ") can not be found in MBISTFailBlock excel sheet"
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'From unused / VBT_LIB_DC_Conti''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used for Module(Exec_IP_Module)
'Wherever the DcConti module, put it there
'Public Function RetrieveDictionaryOfDiffPairs()
'    On Error GoTo ErrHandler
'    Dim Pins() As String, Pin_Cnt As Long, iPin As Long
'    Dim DiffGroup As String: DiffGroup = "All_DiffPairs"                            'T-Autogen will create it."
'    TheExec.DataManager.DecomposePinList DiffGroup, Pins(), Pin_Cnt
'    DicDiffPairs.RemoveAll
'    If Pin_Cnt Mod 2 <> 0 Or Pin_Cnt < 1 Then GoTo ErrHandler
'    For iPin = 0 To Pin_Cnt - 1 Step 2
'        DicDiffPairs.Add LCase(CStr(Pins(iPin))), LCase(CStr(Pins(iPin + 1)))
'        DicDiffPairs.Add LCase(CStr(Pins(iPin + 1))), LCase(CStr(Pins(iPin)))
'    Next iPin
'    Exit Function
'ErrHandler:
'    HandleExecIPError "RetrieveDictionaryOfDiffPairs"
'End Function


'From unused/LIB_Common_Custom''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used for Module(VBT_LIB_Common & VBT_LIB_Digital_Shmoo)
Function RepeatChr(Str As String, repeat As Long) As String
    Dim i         As Long
    RepeatChr = ""
    For i = 0 To repeat - 1
        RepeatChr = RepeatChr & Str
    Next i
End Function



'Used for Module(VBT_LIB_Common & VBT_LIB_Digital_Shmoo)
Public Function FormatNumericDatalog(num As Variant, length As Long, LeftZero As Boolean) As String

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "FormatNumericDatalog"

    ''''Example
    ''''----------------------------------------
    '''' length > 0  is to right shift
    '''' length < 0  is to left  shift
    ''''----------------------------------------
    ''''FormatLog(123456, 8) + "...end"
    ''''  123456...end
    ''''
    ''''FormatLog(123456,-8) + "...end"
    ''''123456  ...end
    ''''
    ''''----------------------------------------

    Dim numStr    As String
    Dim tmpLen    As Long
    Dim spcLen    As Long

    numStr = CStr(num)
    tmpLen = Len(numStr)

    If (tmpLen > Abs(length)) Then
        spcLen = 0
    Else
        spcLen = Abs(length) - tmpLen
    End If

    If (length < 0) Then   ''''number shift to the very left
        FormatNumericDatalog = CStr(num) + Space(spcLen)
    ElseIf LeftZero Then    ''''default: shift to the very right
        FormatNumericDatalog = Space(spcLen) + CStr(num)
    Else
        FormatNumericDatalog = CStr(num) + Space(spcLen)
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'From unused / LIB_Common_DCVS_PPMU''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used for Module(VBT_LIB_Common & VBT_LIB_Digital_Shmoo)
Public Function PowerOff_I_Meter_Parallel(Pin As String, wait_before_gate As Double, wait_after_gate As Double, PowerSequence As Long, Optional DebugPrintEnable As Boolean = False)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "PowerOff_I_Meter_Parallel"

    Dim i_meter_rng As Double   'meter range
    Dim setV      As Double          'current voltage
    Dim stepV     As Double         'step voltage
    Dim stepT     As Double         'step time
    Dim PowerPins() As String
    Dim PinCnt    As Long
    Dim PowerPin  As Variant
    Dim TempString As String
    Dim Vmain(50) As Double
    Dim IRange    As Double
    Dim step      As Integer
    Dim PreStep   As Integer:: PreStep = 0
    Dim RiseTime  As Double
    Dim pin_name  As String
    Dim Pin_Type() As String
    Dim SlotType  As String
    Dim pins_dcvs As String, pins_dcvi As String

    Dim i         As Integer:: i = 1
    Dim j         As Integer:: j = 1
    Dim k         As Integer

    TheExec.DataManager.DecomposePinList Pin, PowerPins(), PinCnt
    ReDim Pin_Type(PinCnt)

    For Each PowerPin In PowerPins
        TempString = PowerPin & "_GLB"
        pin_name = PowerPin
        SlotType = LCase(GetInstrument(pin_name, 0))
        Select Case SlotType
            Case "dc-07":
                Vmain(i) = TheHdw.DCVI.Pins(PowerPin).Voltage
            Case Else
                Vmain(i) = TheHdw.DCVS.Pins(PowerPin).Voltage.Main.Value
        End Select

        'get Ifold limit spec value
        TempString = PowerPin & "_Ifold_GLB"
        IRange = TheExec.Specs.Globals(TempString).ContextValue
        '''            Vmain(i) = 0.1
        '''            Irange = 0.1

        'auto calculate steps
        step = Vmain(i) / 0.1    '0.1v per step
        If step = 0 Then step = 10    'default value
        If step > PreStep Then PreStep = step

        RiseTime = step * ms
        i_meter_rng = IRange
        '---------------------------------------------------------------------
        Select Case SlotType
            Case "dc-07":
                With TheHdw.DCVI.Pins(PowerPin)
                    .Mode = tlDCVIModeVoltage
                    .SetCurrentAndRange IRange, i_meter_rng
                    .CurrentRange.Value = IRange
                    .Current = IRange
                    .Meter.CurrentRange = IRange
                End With
                Pin_Type(k) = "dcvi"
                pins_dcvi = pins_dcvi + "," + PowerPin
            Case Else
                Vmain(i) = TheHdw.DCVS.Pins(PowerPin).Voltage.Main.Value
                With TheHdw.DCVS.Pins(PowerPin)
                    .Mode = tlDCVSModeVoltage
                    .SetCurrentRanges IRange, i_meter_rng
                    '.Meter.mode = tlDCVSMeterCurrent
                    .CurrentRange.Value = IRange
                    .CurrentLimit.Source.FoldLimit.Level.Value = IRange
                    .Meter.CurrentRange = IRange
                End With
                Pin_Type(k) = "dcvs"
                pins_dcvs = pins_dcvs + "," + PowerPin
        End Select
        '---------------------------------------------------------------------
        If DebugPrintEnable = True Then    'debugprint
            TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(PowerPin, 30, False) & ", Vmain " & Format(Vmain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(IRange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", FallTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
        End If

        i = i + 1:: k = k + 1
    Next PowerPin

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
        For Each PowerPin In PowerPins
            setV = Vmain(i) - (j * Vmain(i) / step)
            If Pin_Type(i - 1) = "dcvs" Then
                TheHdw.DCVS.Pins(PowerPin).Voltage.Main = setV
            ElseIf Pin_Type(i - 1) = "dcvi" Then
                TheHdw.DCVI.Pins(PowerPin).Voltage = setV
            End If
            i = i + 1
        Next PowerPin
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

ErrHandler:
    ErrorDescription ("PowerOff_I_Meter_Parallel")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


'Used for Module(VBT_LIB_Common, VBT_LIB_Digital_Shmoo)
Public Function PowerOn_I_Meter_Parallel(Pin As String, wait_before_gate As Double, wait_after_gate As Double, PowerSequence As Long, Optional DebugPrintEnable As Boolean = False)

    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "PowerOn_I_Meter_Parallel"

    Dim i_meter_rng As Double
    Dim setV      As Double
    Dim stepV     As Double
    Dim stepT     As Double
    Dim PowerPins() As String
    Dim PinCnt    As Long
    Dim PowerPin  As Variant
    Dim TempString As String
    Dim Vmain(50) As Double
    Dim IRange    As Double
    Dim step      As Integer
    Dim PreStep   As Integer:: PreStep = 0
    Dim RiseTime  As Double
    Dim pin_name  As String
    Dim Pin_Type() As String
    Dim SlotType  As String
    Dim pins_dcvs As String, pins_dcvi As String

    Dim i         As Integer:: i = 1
    Dim j         As Integer:: j = 1
    Dim k         As Integer

    TheExec.DataManager.DecomposePinList Pin, PowerPins(), PinCnt

    ReDim Pin_Type(PinCnt)

    For Each PowerPin In PowerPins

        TempString = PowerPin & "_GLB"
        Vmain(i) = TheExec.Specs.Globals(TempString).ContextValue

        'get Ifold limit spec value
        TempString = PowerPin & "_Ifold_GLB"
        IRange = TheExec.Specs.Globals(TempString).ContextValue
        '''        Vmain(i) = 0.1
        '''        Irange = 0.1
        'auto calculate steps
        step = Vmain(i) / 0.1    '0.1v per step
        If step = 0 Then step = 10    'default value
        If step > PreStep Then PreStep = step   'calculate largest ramp up steps from all powers in the same sequence

        RiseTime = step * ms
        i_meter_rng = IRange
        '---------------------------------------------------------------------
        pin_name = PowerPin
        SlotType = LCase(GetInstrument(pin_name, 0))
        Select Case SlotType
            Case "dc-07":
                With TheHdw.DCVI.Pins(PowerPin)
                    .Mode = tlDCVIModeVoltage
                    .SetCurrentAndRange IRange, i_meter_rng
                    .CurrentRange.Value = IRange
                    .Current = IRange
                    .Meter.CurrentRange = IRange
                End With
                Pin_Type(k) = "dcvi"
                pins_dcvi = pins_dcvi + "," + PowerPin
            Case Else
                With TheHdw.DCVS.Pins(PowerPin)
                    .Mode = tlDCVSModeVoltage
                    .SetCurrentRanges IRange, i_meter_rng
                    '.Meter.mode = tlDCVSMeterCurrent
                    .CurrentRange.Value = IRange
                    .CurrentLimit.Source.FoldLimit.Level.Value = IRange
                    .Meter.CurrentRange = IRange
                End With
                Pin_Type(k) = "dcvs"
                pins_dcvs = pins_dcvs + "," + PowerPin
        End Select
        '---------------------------------------------------------------------

        If DebugPrintEnable = True Then    'debugprint
            TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(PowerPin, 30, False) & ", Vmain " & Format(Vmain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(IRange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", RiseTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
        End If

        i = i + 1:: k = k + 1
    Next PowerPin

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
        For Each PowerPin In PowerPins
            setV = j * Vmain(i) / step
            If Pin_Type(i - 1) = "dcvs" Then
                TheHdw.DCVS.Pins(PowerPin).Voltage.Main = setV
            ElseIf Pin_Type(i - 1) = "dcvi" Then
                TheHdw.DCVI.Pins(PowerPin).Voltage = setV
            End If
            i = i + 1
        Next PowerPin

        TheHdw.Wait stepT
    Next j
    ''============================================================

    TheHdw.Wait wait_after_gate

    Exit Function

ErrHandler:
    ErrorDescription ("PowerOn_I_Meter_Parallel")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


'Used for Module(VBT_LIB_Digital_Shmoo)
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
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "DCVS_PowerOff_I_Meter_Parallel"

    Dim i_meter_rng As Double   'meter range
    Dim setV      As Double          'current voltage
    Dim stepV     As Double         'step voltage
    Dim stepT     As Double         'step time
    Dim PowerPins() As String
    Dim PinCnt    As Long
    Dim PowerPin  As Variant
    Dim TempString As String
    Dim Vmain(50) As Double
    Dim IRange    As Double
    Dim step      As Integer
    Dim PreStep   As Integer:: PreStep = 0
    Dim RiseTime  As Double

    Dim i         As Integer:: i = 1
    Dim j         As Integer:: j = 1

    TheExec.DataManager.DecomposePinList Pin, PowerPins(), PinCnt

    For Each PowerPin In PowerPins
        'If TheExec.DataManager.ChannelType(PowerPin) <> "N/C" Then 'check CP for FT form NC pins
        TempString = PowerPin & "_GLB"
        'Vmain(i) = TheExec.specs.Globals(TempString).ContextValue
        Vmain(i) = TheHdw.DCVS.Pins(PowerPin).Voltage.Main.Value
        'get Ifold limit spec value
        TempString = PowerPin & "_Ifold_GLB"
        IRange = TheExec.Specs.Globals(TempString).ContextValue

        'auto calculate steps
        step = Vmain(i) / 0.1    '0.1v per step
        If step = 0 Then step = 10    'default value
        If step > PreStep Then PreStep = step

        RiseTime = step * ms
        i_meter_rng = IRange

        With TheHdw.DCVS.Pins(PowerPin)
            .Mode = tlDCVSModeVoltage
            .SetCurrentRanges IRange, i_meter_rng
            '.Meter.mode = tlDCVSMeterCurrent
            .CurrentRange.Value = IRange
            .CurrentLimit.Source.FoldLimit.Level.Value = IRange
            .Meter.CurrentRange = IRange

        End With

        If DebugPrintEnable = True Then    'debugprint
            TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(PowerPin, 30, False) & ", Vmain " & Format(Vmain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(IRange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", FallTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
        End If

        i = i + 1
        '        Else
        '            If DebugPrintEnable = True Then    'debugprint
        '                TheExec.Datalog.WriteComment "print: Pin " & PowerPin & " not turn on by 'NC pin', PowerSequence " & PowerSequence & " ,Warning!!!"
        '            End If
        '        End If
    Next PowerPin

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
        For Each PowerPin In PowerPins
            ''            If TheExec.DataManager.ChannelType(PowerPin) <> "N/C" Then 'check CP for FT form NC pins  'no need to double check NC pin
            setV = Vmain(i) - (j * Vmain(i) / step)
            TheHdw.DCVS.Pins(PowerPin).Voltage.Main = setV

            ''                If DebugPrintEnable = True Then
            ''                    TheExec.Datalog.WriteComment "  Curr_" & PowerPin & " Pwr Down Voltage (" & CStr(i) & ") : " & Format(setV, "0.000") & " V"
            ''                End If
            i = i + 1
            ''            End If
        Next PowerPin
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

ErrHandler:
    ErrorDescription ("DCVS_PowerOff_I_Meter")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


'Used for Module(VBT_LIB_Digital_Shmoo)
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
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "DCVS_PowerOn_I_Meter_Parallel"

    Dim i_meter_rng As Double
    Dim setV      As Double
    Dim stepV     As Double
    Dim stepT     As Double
    Dim PowerPins() As String
    Dim PinCnt    As Long
    Dim PowerPin  As Variant
    Dim TempString As String
    Dim Vmain(50) As Double
    Dim IRange    As Double
    Dim step      As Integer
    Dim PreStep   As Integer:: PreStep = 0
    Dim RiseTime  As Double

    Dim i         As Integer:: i = 1
    Dim j         As Integer:: j = 1


    TheExec.DataManager.DecomposePinList Pin, PowerPins(), PinCnt

    For Each PowerPin In PowerPins
        TempString = PowerPin & "_GLB"
        Vmain(i) = TheExec.Specs.Globals(TempString).ContextValue

        'get Ifold limit spec value
        TempString = PowerPin & "_Ifold_GLB"
        IRange = TheExec.Specs.Globals(TempString).ContextValue

        'auto calculate steps
        step = Vmain(i) / 0.1    '0.1v per step
        If step = 0 Then step = 10    'default value
        If step > PreStep Then PreStep = step   'calculate largest ramp up steps from all powers in the same sequence

        RiseTime = step * ms
        i_meter_rng = IRange

        With TheHdw.DCVS.Pins(PowerPin)
            .Mode = tlDCVSModeVoltage
            .SetCurrentRanges IRange, i_meter_rng
            '.Meter.mode = tlDCVSMeterCurrent
            .CurrentRange.Value = IRange
            .CurrentLimit.Source.FoldLimit.Level.Value = IRange
            .Meter.CurrentRange = IRange
        End With

        If DebugPrintEnable = True Then    'debugprint
            TheExec.Datalog.WriteComment "print: Pin " & FormatNumericDatalog(PowerPin, 30, False) & ", Vmain " & Format(Vmain(i), "0.000") & " V, Irange " & FormatNumericDatalog(Format(IRange, "0.000"), 7, True) & " A, Step " & FormatNumericDatalog(step, 2, True) & ", RiseTime " & FormatNumericDatalog(RiseTime * 1000, 2, True) & " ms" & ", PowerSequence " & FormatNumericDatalog(PowerSequence, 3, True)
        End If

        i = i + 1
    Next PowerPin

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
        For Each PowerPin In PowerPins
            setV = j * Vmain(i) / step
            TheHdw.DCVS.Pins(PowerPin).Voltage.Main = setV
            ''            If DebugPrintEnable = True Then
            ''                TheExec.Datalog.WriteComment "  Curr_" & PowerPin & " Pwr Up Voltage (" & CStr(i) & ") : " & Format(setV, "0.000") & " V"
            ''            End If
            i = i + 1
        Next PowerPin

        TheHdw.Wait stepT

    Next j
    ''============================================================

    TheHdw.Wait wait_after_gate

    Exit Function

ErrHandler:
    ErrorDescription ("DCVS_PowerOn_I_Meter_Parallel")
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'
''From Lib_Pool/LIB_OTP_Function'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'                                                            'Used for Module(VBT_Spotcal, LIB_Common)
'Public Function File_CheckAndCreateFolder(Optional Path_FolderName As String = g_sOTPDATA_FILEDIR, Optional CreateFolder As YESNO_type = Yes)
'Dim funcName As String:: funcName = "CheckAndCreateFolder"
'On Error GoTo ErrHandler
'
'Dim mS_TempString As String
'Dim mB_DebugPrtDlog As Boolean
'
'mB_DebugPrtDlog = True
'mS_TempString = ""
'
'     Dim ffs As New FileSystemObject
'
'     If (Right(Trim(Path_FolderName), 1) <> "\") Then Path_FolderName = Path_FolderName + "\"
'
'     If Not (ffs.FolderExists(Path_FolderName)) Then
'        TheExec.Datalog.WriteComment "<Notice!!> The Folder::" + Path_FolderName + " Is Not Exist."
'        Select Case CreateFolder
'            Case Yes
'                ffs.CreateFolder Path_FolderName
'                 mS_TempString = "The New Folder::" + Path_FolderName
'            Case No
'                 mS_TempString = "Skip To Create The New Folder::" + Path_FolderName
'        End Select
'     Else
'                 'mS_TempString = "The Datalog Folder::" + Path_FolderName
'     End If
'
'      If (mB_DebugPrtDlog) Then TheExec.Datalog.WriteComment mS_TempString
'
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function
'
'
'                                                            'Used for Module(VBT_Spotcal, LIB_Common)
'Public Function File_CreateAFile(FileName As String, Text As String)
'
'Dim fso As FileSystemObject
'Dim fid As TextStream
'Dim i As Long
'Dim funcName As String:: funcName = "File_CreateAFile"
'
'    On Error GoTo ErrHandler
'
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    Set fid = fso.CreateTextFile(FileName, True)
'
'    fid.WriteLine (Text)
'    fid.Close
'
'Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function


'From unused/LIB_Digital_Shmoo_Sub''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used for Module(LIB_Globals_Generic)
Public Function Search_Low2High_First_Pass(Shmoo_result_PF As String) As Integer
    Dim char_pt   As String
    Dim max       As Integer
    Dim point     As Integer
    Dim i         As Long
    On Error GoTo err1
    Dim funcName  As String:: funcName = "Search_Low2High_First_Pass"
    max = Len(Shmoo_result_PF)

    For i = 1 To max
        char_pt = Mid(Shmoo_result_PF, i, 1)
        If (char_pt = "P") Then point = i: i = max

    Next i
    Search_Low2High_First_Pass = point
    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


'Used for Module(LIB_Globals_Generic)
Public Function Search_High2Low_First_Pass(Shmoo_result_PF As String) As Integer
    Dim char_pt   As String
    Dim max       As Integer
    Dim point     As Integer
    Dim i         As Long

    On Error GoTo ErrHandler:
    Dim funcName  As String:: funcName = "Search_High2Low_First_Pass"
    max = Len(Shmoo_result_PF)

    For i = max To 1 Step -1
        char_pt = Mid(Shmoo_result_PF, i, 1)
        If (char_pt = "P") Then point = i: i = 1

    Next i

    Search_High2Low_First_Pass = point
    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


'Used for Module(LIB_Globals_Generic)
Public Function Search_HVCC(Shmoo_result_PF As String) As Integer
    Dim str_len, start_point As Integer
    Dim search_dif As Boolean
    Dim char_pt   As String
    Dim report_point As Integer
    Dim i         As Long
    On Error GoTo err1
    Dim funcName  As String:: funcName = "Search_HVCC"

    str_len = Len(Shmoo_result_PF)
    start_point = Search_Low2High_First_Pass(Shmoo_result_PF)
    search_dif = False
    report_point = 0
    For i = start_point To str_len
        char_pt = Mid(Shmoo_result_PF, i, 1)

        If Not (search_dif) Then
            If (char_pt = "F") Then
                report_point = i - 1
                search_dif = True
            End If
        Else
            If (char_pt = "P") Then
                report_point = -1
                i = str_len
            End If
        End If
    Next i
    If (report_point = 0) Then report_point = str_len: ReportHVCC = False
    Search_HVCC = report_point
    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next


End Function


'Used for Module(LIB_Globals_Generic)
Public Function Search_LVCC(Shmoo_result_PF As String) As Integer
    Dim str_len, start_point As Integer
    Dim search_dif As Boolean
    Dim char_pt   As String
    Dim report_point As Integer
    Dim i         As Long

    On Error GoTo err1
    Dim funcName  As String:: funcName = "Search_LVCC"

    str_len = Len(Shmoo_result_PF)
    start_point = Search_High2Low_First_Pass(Shmoo_result_PF)
    search_dif = False
    report_point = 0

    For i = start_point To 1 Step -1
        char_pt = Mid(Shmoo_result_PF, i, 1)

        If Not (search_dif) Then
            If (char_pt = "F") Then
                report_point = i + 1
                search_dif = True
            End If
        Else
            If (char_pt = "P") Then
                report_point = -1
                i = 1
            End If
        End If
    Next i
    If (report_point = 0) Then report_point = 1: ReportLVCC = False
    Search_LVCC = report_point
    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


'Used for Module(LIB_Globals_Generic, LIB_Digital_Shmoo)
Public Function Search_VIH_LVCC(Shmoo_result_PF As String) As Integer
    'Report -1 is shmoo hole,report -2 is first point fail
    Dim str_len, start_point As Integer
    Dim search_dif As Boolean
    Dim char_pt   As String
    Dim report_point As Integer
    Dim i         As Long

    On Error GoTo err1
    Dim funcName  As String:: funcName = "Search_VIH_LVCC"
    
    str_len = Len(Shmoo_result_PF)
    start_point = str_len
    If (Mid(Shmoo_result_PF, start_point, 1) = "F") Then Search_VIH_LVCC = -2: Exit Function
    search_dif = False
    report_point = 0

    For i = start_point To 1 Step -1
        char_pt = Mid(Shmoo_result_PF, i, 1)

        If Not (search_dif) Then
            If (char_pt = "F") Then
                report_point = i + 1
                search_dif = True
            End If
        Else
            If (char_pt = "P") Then
                report_point = -1
                i = 1
            End If
        End If
    Next i
    If (report_point = 0) Then report_point = 1
    Search_VIH_LVCC = report_point
    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


'Used for Module(LIB_Globals_Generic, LIB_Digital_Shmoo)
Public Function Search_VIL_HVCC(Shmoo_result_PF As String) As Integer
    'Report -1 is shmoo hole,report -2 is first point fail
    Dim str_len, start_point As Integer
    Dim search_dif As Boolean
    Dim char_pt   As String
    Dim report_point As Integer
    Dim i         As Long

    On Error GoTo err1
    Dim funcName  As String:: funcName = "Search_VIL_HVCC"
    
    str_len = Len(Shmoo_result_PF)
    start_point = 1
    If (Mid(Shmoo_result_PF, 1, 1) = "F") Then Search_VIL_HVCC = -2: Exit Function
    search_dif = False
    report_point = 0
    For i = start_point To str_len
        char_pt = Mid(Shmoo_result_PF, i, 1)

        If Not (search_dif) Then
            If (char_pt = "F") Then
                report_point = i - 1
                search_dif = True
            End If
        Else
            If (char_pt = "P") Then
                report_point = -1
                i = str_len
            End If
        End If
    Next i
    If (report_point = 0) Then report_point = str_len
    Search_VIL_HVCC = report_point
    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


' US: Removed is now part of "LIB_VBT_Rly_Ctrl"
'From Lib_Pool/LIB_VBT_Rly_Ctrl''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Used for Module(VBT_LIB_DC_Conti_PMIC, VBT_LIB_DC_Leak_PMIC)
''Public Function ADG1414_CONTROL(VOUT_0A4A_UVI80_S As Long, VOUT_0B4B_UVI80_S As Long, _
 ''                                VOUT_8A12A_UVI80_S As Long, VOUT_8B12B_UVI80_S As Long, _
 ''                                VOUT_16A20A_UVI80_S As Long, VOUT_16B20B_UVI80_S As Long, _
 ''                                VOUT_24A24B_UVI80_S As Long, _
 ''                                Optional ADGOverWrite As Boolean = False, _
 ''                                Optional ADGResetN As Boolean = False) As Long
''On Error GoTo ErrHandler
''Dim funcName As String:: funcName = "ADG1414_CONTROL"
''
''
''    Dim SPIReadBack As New SiteLong
''    Dim Site As Variant
''    Dim TestStatus As SiteLong
''    Dim Result As New SiteLong
''    Dim PortsResult As New PinListData
''    Dim LastTestPassed As New SiteBoolean
''
''    TheHdw.DIB.PowerOn = True
''     '----------------------------------
''     'Trace the ADG1414
''     g_ADG1414ArgList = "VOUT_24A24B_UVI80_S,VOUT_16B20B_UVI80_S,VOUT_16A20A_UVI80_S,VOUT_8B12B_UVI80_S,VOUT_8A12A_UVI80_S,VOUT_0B4B_UVI80_S,VOUT_0A4A_UVI80_S"
''     g_ADG1414Data(0) = VOUT_24A24B_UVI80_S
''     g_ADG1414Data(1) = VOUT_16B20B_UVI80_S
''     g_ADG1414Data(2) = VOUT_16A20A_UVI80_S
''     g_ADG1414Data(3) = VOUT_8B12B_UVI80_S
''     g_ADG1414Data(4) = VOUT_8A12A_UVI80_S
''     g_ADG1414Data(5) = VOUT_0B4B_UVI80_S
''     g_ADG1414Data(6) = VOUT_0A4A_UVI80_S
'''     g_ADG1414Data(6) = B1LX012
'''     g_ADG1414Data(7) = B0LX013
'''     g_ADG1414Data(8) = B0LX024
''    '----------------------------------
''
''
''    '___Change FRC Path
''    Call SetFRCPath("ADG1414_SCLK", True) 'IMOLA20171013 To Prevent the ADG1414 Reset while running any pattern on CL
''    Call SetFRCPath("ADG1414_RESET", True) 'IMOLA20171013 To Prevent the ADG1414 Reset while running any pattern on CL
''
''    With TheHdw.Digital.Pins("ADG1414_PINS")
''        .StartState = chStartOff
''        .InitState = chInitoff
''    End With
'''      'Apply levels and timing for the protocol pins
'''    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
''
''    '___Enable the ports, HRAM setup
''    TheHdw.Protocol.ports("ADG1414_PINS").Enabled = True
''    TheHdw.Protocol.ports("ADG1414_PINS").NWire.HRAM.Setup.TriggerType = tlNWireHRAMTriggerType_First
''    TheHdw.Protocol.ports("ADG1414_PINS").NWire.HRAM.Setup.WaitForEvent = False
''    TheHdw.Protocol.ModuleRecordingEnabled = True
''
''    Call SPI_BYTE_WRITE1(VOUT_24A24B_UVI80_S, "ADG1414_PINS") '9
''    Call SPI_BYTE_WRITE2(VOUT_16B20B_UVI80_S, "ADG1414_PINS") '9
''    Call SPI_BYTE_WRITE2(VOUT_16A20A_UVI80_S, "ADG1414_PINS") '9
''    Call SPI_BYTE_WRITE2(VOUT_8B12B_UVI80_S, "ADG1414_PINS") '9
''    Call SPI_BYTE_WRITE2(VOUT_8A12A_UVI80_S, "ADG1414_PINS") '9
''    Call SPI_BYTE_WRITE2(VOUT_0B4B_UVI80_S, "ADG1414_PINS") '9
''    Call SPI_BYTE_WRITE2(VOUT_0A4A_UVI80_S, "ADG1414_PINS") '9
''
''
'''    Call SPI_BYTE_READ_WRITE2(B11LX01, "ADG1414_PINS") '8
'''    Call SPI_BYTE_READ_WRITE2(B4LX01, "ADG1414_PINS") '7
'''    Call SPI_BYTE_READ_WRITE2(B2LX01, "ADG1414_PINS") '6
'''    Call SPI_BYTE_READ_WRITE2(B3LXABC, "ADG1414_PINS") '5
'''    Call SPI_BYTE_READ_WRITE2(B1LX0B5LX01, "ADG1414_PINS") '4
'''    Call SPI_BYTE_READ_WRITE2(B1LX012, "ADG1414_PINS") '3
'''    Call SPI_BYTE_READ_WRITE2(B0LX013, "ADG1414_PINS") '2
'''    Call SPI_BYTE_READ_WRITE2(B0LX024, "ADG1414_PINS") '1
''
''''    Call SPI_BYTE_READ_WRITE2(B1LX031_B8LX0, "ADG1414_PINS") '8
''''    Call SPI_BYTE_READ_WRITE2(B1LX420, "ADG1414_PINS") '7
''''    Call SPI_BYTE_READ_WRITE2(B10LX_B6LXAB, "ADG1414_PINS") '6
''''    Call SPI_BYTE_READ_WRITE2(B3LXABC, "ADG1414_PINS") '5
''''    Call SPI_BYTE_READ_WRITE2(B4LX01_B3LXC_B10LX, "ADG1414_PINS") '4
''''    Call SPI_BYTE_READ_WRITE2(B2LX01, "ADG1414_PINS") '3
''''    Call SPI_BYTE_READ_WRITE2(AMUXB, "ADG1414_PINS") '2
''''    Call SPI_BYTE_READ_WRITE2(AMUXA, "ADG1414_PINS") '1      'Last byte to write and read dummy data.
''    ' ****************************************************************************************************
''
''
''    TheHdw.Protocol.ports("ADG1414_PINS").Halt
''    TheHdw.Protocol.ports("ADG1414_PINS").Enabled = False
''
''Exit Function
''ErrHandler:
''    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
''    If AbortTest Then Exit Function Else Resume Next
''End Function


'Used for Module(VBT_LIB_DC_Conti_PMIC, VBT_LIB_DC_Leak_PMIC)
Public Function SetFRCPath(PinName As String, Bypass As Boolean)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SetFRCPath"

    ' Setup IKS mask for PIK & PGIK
    Call TheHdw.Raw.TSIO.Wr("A_GL_IK_MSK", 9)

    ' Set the steering to select the pin
    Dim PinKey    As Long
    Call m_stdsvcclient.IkSvc.GetPinKey(PinName, PinKey)
    Call m_stdsvcclient.IkSvc.SelectPinKey(PinKey)

    If Bypass Then
        ' Bypass the DDR flop
        Call TheHdw.Raw.TSIO.Wr("A_P_TIM_PA_DRV_SEL", 6)
    Else
        ' Use the clock path
        Call TheHdw.Raw.TSIO.Wr("A_P_TIM_PA_DRV_SEL", 5)
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'Used for Module(LIB_Globals_Generic)
Public Function SPI_BYTE_WRITE1(ByteVal As Variant, PortName As String) As Long

    If TheExec.Sites.ActiveCount = 0 Then Exit Function  ' bin out if stop on fail
    On Error GoTo ErrHandler

    With TheHdw.Protocol.ports(PortName).NWire.Frames("SPI_BYTE_WRITE1")
        .Fields("DIN").Value = ByteVal
        .Execute    ' tlNWireExecutionType_CaptureInCMEM
    End With
    TheHdw.Protocol.ports(PortName).IdleWait
    Exit Function

ErrHandler:
    Debug.Print "Error in SPI_BYTE_WRITE: " & err.Description
    Resume Next
End Function


'Used for Module(LIB_Globals_Generic)
Public Function SPI_BYTE_WRITE2(ByteVal As Variant, PortName As String) As Long

    If TheExec.Sites.ActiveCount = 0 Then Exit Function  ' bail out if stop on fail
    On Error GoTo ErrHandler

    With TheHdw.Protocol.ports(PortName).NWire.Frames("SPI_BYTE_WRITE2")
        .Fields("DIN").Value = ByteVal
        .Execute    ' tlNWireExecutionType_CaptureInCMEM
    End With
    TheHdw.Protocol.ports(PortName).IdleWait
    Exit Function

ErrHandler:
    Debug.Print "Error in SPI_BYTE_WRITE: " & err.Description
    Resume Next
End Function

