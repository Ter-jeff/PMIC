Attribute VB_Name = "LIB_Digital_Shmoo"
Option Explicit

Public pseudo_result_index As Long
'Revision History
' 1.3 add support for edge shmoo with skip tests
' 1.4 add support for "retest" in characterization
' 1.5 add shmoo pattern gloabal variable for Char setup retest function
' 1.6 add Char key word Shmoo_header for hard IP used
' 1.7 add IG-XL 8.10.12 coding for pattern list
' 1.8 change lotId and wafer Id to site variable
' 1.9 add Char_map.txt for in job characterization
' 1.10 add function Shmoo_Test_Pattern for in round-1 characterization
' 1.11 make VBT_shmoo.bas more independent

Public Set_Pin_NV As Double
Public Const VBT_Shmoo_Version = "1.11"
Public Const MAX_CHAR_ENABLE_ROW = 30000
Public Const MAX_CHAR_SETUP_ROW = 100
Public Const Char_Flow_Enable_Sheet = ".\Setup\Char_Flow_Enable.txt"
Public Const Shmoo_Setup_Sheet_Enable = ".\Setup\Char_Enable.txt"
Public Const Shmoo_Setup_Sheet_Setup = ".\Setup\Char_Setup.txt"
Public Const char_map_Sheet = ".\Common\Char_Map.txt"
Public wb As Workbook
Public Shmoo_Pattern As String
Public Shmoo_Pattern_Payload As String
Public Shmoo_header As String
Public Shmoo_Vcc_Min As New SiteDouble
Public Shmoo_Vcc_Max As New SiteDouble
Public ShmooPowerName As String
Public Shmoo_Instance_Name() As String
Public Shmoo_setup_name() As String
Public Shmoo_Setup_Name_New() As String
Public Shmoo_Setup_idx As Integer
''For AI use 20150715
Public Voltage_fail_point As Long
Public Voltage_fail_point_request As Long
Public Voltage_fail_collect(10) As String
Public ReportHVCC As Boolean
Public ReportLVCC As Boolean
Public ShmResult As New SiteVariant
Public RTOSPatResult As New SiteBoolean

Type Char_Enable
    Enable As String
    TestInstance As String
    charSetup As String
    Pattern As String
    Count As Long
End Type
Type Char_setup
    Setup_Name As String
    Test_Method As String
    Step_Name As String
    mode As String
    Parameter_Type As String
    Parameter_Name As String
    Range_Calc_Field As String
    Range_From As String
    Range_To As String
    Range_Steps As String
    Range_Step_Size As String
    Perform_Test  As String
    Test_Limits_Low As String
    Test_Limits_High As String
    Algorithm_Name As String
    Algorithm_Arguments As String
    Algorith_Results_Check As String
    Algorithm_Transition As String
    Apply_To_Pins As String
    Apply_To_Pin_Exec_Mode As String
    Apply_To_Time_Sets As String
    Adjust_Backoff As String
    Adjust_Spec_Name As String
    Adjust_From_Setup As String
    Function As String
    Function_Arguments As String
    Interpose_Functions_Pre_Setup As String
    Interpose_Functions_Pre_Setup_Arguments As String
    Interpose_Functions_Pre_Step As String
    Interpose_Functions_Pre_Step_Arguments As String
    Interpose_Functions_Pre_Point As String
    Interpose_Functions_Pre_Point_Arguments As String
    Interpose_Functions_Post_Point As String
    Interpose_Functions_Post_Point_Arguments As String
    Interpose_Functions_Post_Step As String
    Interpose_Functions_Post_Step_Arguments As String
    Interpose_Functions_Post_Setup As String
    Interpose_Functions_Post_Setup_Arguments As String
    Output_Format As String
    Output_Text_File As String
    Output_Sheet As String
    Output_Destinations_Text_File As String
    Output_Destinations_Sheet As String
    Output_Destinations_Datalog As String
    Output_Destinations_Immediate_Win As String
    Output_Destinations_Output_Win As String
    comment As String
    Count As Long
End Type
Public Const MaxCharSetup = 15
Public Const MaxCharCorePower = 7
Public Const MaxCharInitPatt = 4
Public Const MaxFuncBlock = 100
'Enum Char_Enable_Enum
'    Fail_Enable = 1
'    Disable = 2
'    Enable = 3
'End Enum
Type Char_map
     TestNum(MaxCharSetup) As String
     Func_Block As String
     PowerCondition(MaxCharSetup) As String
     Enable(MaxCharSetup) As String
     Char_setup(MaxCharSetup) As String
     NV_Power(MaxCharSetup) As String
     Core_power(MaxCharSetup, MaxCharCorePower) As Double
     Init_Patt(MaxCharSetup, MaxCharInitPatt) As String
     Count As Long
End Type
Type Current_Shmoo_Setup
    TestNum As Long
    Enable As String
    Func_Block As String
    Func_block_index As Long
    PowerCondition As String
    Char_Setup_Index As Long 'index of  char setup within a function block
    Char_Setup_Name As String
    Pins_Apply As String
End Type
Public Curr_Shmoo_Condition As Current_Shmoo_Setup
Public char_map_entry(MaxFuncBlock) As Char_map
Public Char_Setup_Collection_Index As New Collection
Public count_func_block As Long
Public ShmooSweepPower(100) As New SiteDouble
Public Power_Level_Last As New SiteVariant
Public Shmoo_Apply_Pin As String

Dim char_flow_enable_entry(MAX_CHAR_ENABLE_ROW) As Char_Enable
Dim char_enable_entry(MAX_CHAR_ENABLE_ROW) As Char_Enable
Dim char_setup_entry(MAX_CHAR_SETUP_ROW) As Char_setup
Dim char_flow_enable_key As New Collection
Dim char_enable_key As New Collection
Dim char_setup_key As New Collection
Dim char_setup_count As Long
Dim char_enable_count As Long
Dim shmoo_mode As tlDevCharShmooAxis
Dim shmoo_algorithm As tlDevCharShmooPGA
Dim shmoo_Calc_Field As tlDevCharRangeField
Dim shmoo_Apply_To_Pin_Exec_Mode As tlDevCharPinExecMode
Dim shmoo_Destination_DataLog As tlDevCharOutputDestinationState
Dim shmoo_Destination_OutputWindow As tlDevCharOutputDestinationState
Dim shmoo_Destination_Sheet As tlDevCharOutputDestinationState
Dim shmoo_Destination_TextFile As tlDevCharOutputDestinationState
Dim shmoo_Destination_ImmediateWindow As tlDevCharOutputDestinationState

Public Flow_Shmoo_Axis(20) As String
Public Flow_Shmoo_Axis_Count As Long
Public Flow_Shmoo_X_Step As Long
Public Flow_Shmoo_Y_Step As Long
Public Flow_Shmoo_X_Current_Step As Long
Public Flow_Shmoo_Y_Current_Step As Long
Public Flow_Shmoo_X_Last_Value As Long
Public Flow_Shmoo_Y_Last_Value As Long
Public Flow_Shmoo_X_Start As Long
Public Flow_Shmoo_Y_Start As Long
Public Flow_Shmoo_X_Fast As Boolean
Public Flow_Shmoo_Force_Condition As String
Public Shmoo_setup_str As String
Public Shmoo_End As Boolean
Public Flow_Shmoo_Port_Name As String
Public FlowShmooString_GLB As String
Public shmoohole_count As New SiteLong
Public shmooallfail_count As New SiteLong
Public shmooalarm_count As New SiteLong
Public included_shmoo_count As New SiteLong
Public excluded_shmoo_count As New SiteLong
Public total_shmoo_count As New SiteLong
Public F_shmoo_abnormal_counter As Boolean
Public Type testCondition
    DigSrc_BinStr As String
    ConditionName As String
    DigSrc_BitCount As Double
End Type

Public Type DynamicSrc
    PatternName As String
    TestCase() As testCondition
End Type
Public SrcStock() As DynamicSrc
Public DSSCMappingTableIsRead As Boolean
Public g_Retention_Start As Boolean
Public g_Retention_Shmoo As Boolean
Public g_ForceCond_VDD As String
Public g_Retention_FC As String ' Retention pin/Voltage parsed from force condition "RETV", Eg. VDD1:RETV:1.0;VDD3,VDD4:RETV:1.1 => VDD1=1.0;VDD3,VDD4=1.1
Public g_Retention_VDD As String 'Retention pin parsed from force condition "RETV"
Public g_Retention_ForceV As String 'Retention Voltage parsed from force condition "RETV"

'=================================================================
' 201810 add these parameters for Select Sram START
Public Type Sub_Info
    BITS As Integer
    logicPin As String
    SramPin As String
    SelSram1 As Integer
    SelSram0 As Integer
End Type
Public Type Domain
    DomainName As String
    Pattern() As String
    DomainBits() As Sub_Info
End Type
Public Type mapping_table
    Block() As Domain
End Type
Public GetSelSram As mapping_table
Public PrintDSSCSwitchVoltage As New PinListData
Public PrintSwitchDspWave As New DSPWave
Public g_BlockType As String
Public digSrc_EQ_GB As String
Public BlockType_GB As String
Public DigSrc_pin_GB As New PinList
Public DigSrcSize_GB As String
Public dssc_pat_init_GB As String
Public g_shmoo_ret As Boolean
Public g_InitSeq As String
Public g_dyanmicDSSCbits As String
Public RTOS_Shmoo_Start As Boolean
Public g_VminBoundary_selsrm As Boolean

Public g_FirstSetp As Boolean
Public g_Vbump_function As Boolean
Public g_Print_SELSRM_Def As Boolean
Public ShmooSweepPowerDict As New PinListData
Public Power_Level_Vmode_Last As String
Public g_ApplyLevelTimingVmain As New PinListData
Public g_ApplyLevelTimingValt As New PinListData
Public g_CharInputString_Voltage_Dict As New Dictionary
Public g_Globalpointval As New PinListData
Public g_VDDForce As String
Public g_PLSWEEP As Boolean
' 201810 add these parameters for Select Sram END
'=================================================================

Public g_ShmooPin As New PinListData

Public Function Print_power_condition() As String



Dim VDD_CPU_POWER As String
Dim VDD_CPU_SRAM_POWER As String
Dim VDD_GPU_POWER As String
Dim VDD_GPU_SRAM_POWER As String
Dim VDD_SOC_POWER As String
Dim VDD_LOW_POWER As String
Dim VDD_FIXED_POWER As String
Dim Power_Condition_string As String

VDD_CPU_POWER = CStr(TheHdw.DCVS.Pins("VDD_CPU").Voltage.Main.Value)
VDD_CPU_SRAM_POWER = CStr(TheHdw.DCVS.Pins("VDD_CPU_SRAM").Voltage.Main.Value)
VDD_GPU_POWER = CStr(TheHdw.DCVS.Pins("VDD_GPU").Voltage.Main.Value)
VDD_GPU_SRAM_POWER = CStr(TheHdw.DCVS.Pins("VDD_GPU_SRAM").Voltage.Main.Value)
VDD_SOC_POWER = CStr(TheHdw.DCVS.Pins("VDD_SOC").Voltage.Main.Value)
VDD_LOW_POWER = CStr(TheHdw.DCVS.Pins("VDD_LOW").Voltage.Main.Value)
VDD_FIXED_POWER = CStr(TheHdw.DCVS.Pins("VDD_FIXED").Voltage.Main.Value)

Power_Condition_string = "[Power_Condition :" & VDD_CPU_POWER & "," & VDD_CPU_SRAM_POWER & "," & _
        VDD_GPU_POWER & "," & VDD_GPU_SRAM_POWER & "," & VDD_SOC_POWER & "," & VDD_LOW_POWER _
        & "," & VDD_FIXED_POWER
'Debug.Print Power_Condition_string

End Function

''''Move_to_Old_Shmoo
''Public Function Read_Char_Map_Excel()
''
''
''    Dim Row As Long, count_per_func_block As Long
''    Dim name_last As String
''    Dim FileNum As Integer
''    Dim DataLine() As String
''    Dim DataLine1() As String
''    Dim entry() As String
''    Dim flg_last As Boolean
''    Dim j As Long
''    Dim multiple_col_start As Long
'' '=========================================
''    Dim RowMaxNum As Long
''    Dim ColMaxNum As Long
''    Dim TempColContent As String
''    Dim ColNum As Integer
''    Dim RowNum As Integer
'' '=========================================
''
''    count_func_block = -1
''    name_last = ""
''    flg_last = False
''
''    Set Char_Setup_Collection_Index = Nothing
''
''    FileNum = FreeFile()
''''    Open char_map_Sheet For Input As #FileNum
''''    Row = 1
''
''    RowMaxNum = Worksheets("Char_map").Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Row
''    ColMaxNum = Worksheets("Char_map").Cells(1, ActiveSheet.Columns.Count).End(xlToLeft).Column
''
''    ReDim DataLine(RowMaxNum) As String
''    ReDim DataLine1(RowMaxNum) As String
''
''
''    For RowNum = 1 To RowMaxNum
''        For ColNum = 1 To ColMaxNum
''            TempColContent = Worksheets("Char_map").Cells(RowNum, ColNum)
''            If ColNum = 1 Then 'delete first space
''                DataLine1(RowNum) = LTrim(DataLine(RowNum) + "  " + TempColContent)
''                DataLine(RowNum) = DataLine1(RowNum)
''            Else
''                DataLine(RowNum) = DataLine(RowNum) + vbTab + TempColContent
''            End If
''        Next ColNum
''        entry = Split(DataLine(RowNum), vbTab)
''        If RowNum > 1 And DataLine(RowNum) <> "" Then  ' data valid from row 10
''            If entry(1) <> name_last Then
''                count_func_block = count_func_block + 1                                     ' num of function block
''                count_per_func_block = -1
''                Char_Setup_Collection_Index.Add CStr(count_func_block), entry(1)                       ' add index key for char_map_entry
''            End If
''            count_per_func_block = count_per_func_block + 1
''            char_map_entry(count_func_block).TestNum(count_per_func_block) = entry(0)
''            char_map_entry(count_func_block).Func_Block = entry(1)
''            char_map_entry(count_func_block).PowerCondition(count_per_func_block) = entry(2)
''            char_map_entry(count_func_block).Enable(count_per_func_block) = entry(3)
''            char_map_entry(count_func_block).Char_setup(count_per_func_block) = entry(4)
''            char_map_entry(count_func_block).Count = count_per_func_block
''            multiple_col_start = 5
''            For j = 0 To MaxCharCorePower - 1
''                If entry(multiple_col_start + j) <> "" Then
''                    char_map_entry(count_func_block).Core_Power(count_per_func_block, j) = CDbl(entry(multiple_col_start + j))
''                Else
''                    char_map_entry(count_func_block).Core_Power(count_per_func_block, j) = -99
''                End If
''            Next j
''            multiple_col_start = multiple_col_start + MaxCharCorePower
''            For j = 0 To MaxCharInitPatt - 1
''                char_map_entry(count_func_block).Init_Patt(count_per_func_block, j) = entry(multiple_col_start + j)
''            Next j
''            name_last = char_map_entry(count_func_block).Func_Block
''        End If
''    Next RowNum
''
''
''End Function
'''Move_to_Old_Shmoo'
''Public Function Read_Char_Map()
''
''    Dim Row As Long, count_per_func_block As Long
''    Dim name_last As String
''    Dim FileNum As Integer
''    Dim DataLine As String
''    Dim entry() As String
''    Dim flg_last As Boolean
''    Dim j As Long
''    Dim multiple_col_start As Long
''
''    count_func_block = -1
''    name_last = ""
''    flg_last = False
''
''    Set Char_Setup_Collection_Index = Nothing
''
''    FileNum = FreeFile()
''    Open char_map_Sheet For Input As #FileNum
''    Row = 1
''
''    While Not EOF(FileNum)
''        Line Input #FileNum, DataLine ' read in data 1 line at a time
''        entry = Split(DataLine, vbTab)
''        If Row > 1 And DataLine <> "" Then  ' data valid from row 10
''            If entry(1) <> name_last Then
''                count_func_block = count_func_block + 1                                     ' num of function block
''                count_per_func_block = -1
''                Char_Setup_Collection_Index.Add CStr(count_func_block), entry(1)                       ' add index key for char_map_entry
''            End If
''            count_per_func_block = count_per_func_block + 1
''            char_map_entry(count_func_block).Enable(count_per_func_block) = entry(0)
''            char_map_entry(count_func_block).Func_Block = entry(1)
''            char_map_entry(count_func_block).Char_setup(count_per_func_block) = entry(2)
''            char_map_entry(count_func_block).NV_Power(count_per_func_block) = entry(3)
''            char_map_entry(count_func_block).Count = count_per_func_block
''            multiple_col_start = 4
''            For j = 0 To MaxCharCorePower - 1
''                If entry(multiple_col_start + j) <> "" Then
''                    char_map_entry(count_func_block).Core_Power(count_per_func_block, j) = CDbl(entry(multiple_col_start + j))
''                Else
''                    char_map_entry(count_func_block).Core_Power(count_per_func_block, j) = -99
''                End If
''            Next j
''            multiple_col_start = multiple_col_start + MaxCharCorePower
''            For j = 0 To MaxCharInitPatt - 1
''                char_map_entry(count_func_block).Init_Patt(count_per_func_block, j) = entry(multiple_col_start + j)
''            Next j
''            name_last = char_map_entry(count_func_block).Func_Block
''        End If
''        Row = Row + 1
''    Wend
''
''    Close #FileNum
''End Function


Public Function Shmoo_To_Do() As Boolean
    If (TheExec.DataManager.instanceName Like "*_LV" And Curr_Shmoo_Condition.PowerCondition Like "*l*") _
        Or (TheExec.DataManager.instanceName Like "*_HV" And Curr_Shmoo_Condition.PowerCondition Like "*h*") _
        Or (TheExec.DataManager.instanceName Like "*_NV" And Curr_Shmoo_Condition.PowerCondition Like "*n*") Then
        Shmoo_To_Do = True
    Else
        Shmoo_To_Do = False
    End If
End Function


Public Function Shmoo_Test_Pattern_old(patt As Pattern, ReportResult As PFType, ResultMode As tlResultMode, ConcurrentMode As tlPatConcurrentMode, Power_Run_Scenario As String, powerPin As String, set_init As Boolean, seq As Long, wait_time As String)
    Dim External_Retention As Boolean
    Dim test_name_ary() As String
    Dim SRV_type As String
    Dim block_name As String
    Dim site As Variant
    Dim lPatternCount As Long
    Dim astrPattTemp() As String

    '' Add for Pattern loop ,20160607, KS
    If InStr(patt, ":") Then
        astrPattTemp = Split(patt, ":")
        lPatternCount = CLng(astrPattTemp(1)) - 1
        patt = astrPattTemp(0)
        TheExec.Datalog.WriteComment "Loop Pattern :" & patt & "_" & "Repeat count :" & lPatternCount + 1
    Else
        lPatternCount = 0
    End If
    
    If patt.Value = "" Then Exit Function
    
'    Call TheHdw.Patterns(patt).Load
'    Call TheHdw.Patterns(patt).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode)
'    Exit Function
    
    
    test_name_ary = Split(TheExec.DataManager.instanceName, "_")

    If UBound(test_name_ary) > 0 Then
        block_name = LCase(test_name_ary(1)) 'CPU,GPU,SOC
        If LCase(test_name_ary(4)) Like "*ext*" Then External_Retention = True
    End If
    
''    If (InStr(test_name_ary(4), "SRVA") > 0) Then
''        SRV_type = "SRVA"
''    ElseIf (InStr(test_name_ary(4), "SRVB") > 0) Then
''        SRV_type = "SRVB"
''    End If

    For Each site In TheExec.sites
        Set_Run_Level Power_Run_Scenario, powerPin, set_init, seq
    Next site
    Call TheHdw.Patterns(patt).Load
'    Call TheHdw.Patterns(patt).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode)
    Dim InDSPwave As New DSPWave
    Dim Count As Long
    
   '' Add for Pattern loop ,20160607, KS
    '---------------
    For Count = 0 To lPatternCount
        Call TheHdw.Patterns(patt).start                ' make sure to jump out  the cpu loop before halt
        While TheHdw.Digital.Patgen.IsRunning = True
            TheHdw.Digital.Patgen.Continue 0, cpuA
        Wend
        TheHdw.Digital.Patgen.HaltWait
    Next Count

    
    '------------------
'''''
'''''    Call TheHdw.Patterns(patt).start                ' make sure to jump out  the cpu loop before halt
'''''    While TheHdw.Digital.Patgen.IsRunning = True
'''''        TheHdw.Digital.Patgen.Continue 0, cpuA
'''''    Wend
'''''    TheHdw.Digital.Patgen.HaltWait
    
    '' KS update for remove fail pins count function when do shmoo
    If UCase(TheExec.CurrentJob) Like "*CHAR*" Then
        Dim TestNumber As Long
        For Each site In TheExec.sites
                TestNumber = TheExec.sites.Item(site).TestNumber
                If TheHdw.Digital.Patgen.PatternBurstPassed(site) Then
                    If TheExec.DevChar.Setups.IsRunning = True Then TheExec.sites.Item(site).testResult = sitePass
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                Else
                    TheExec.sites.Item(site).testResult = siteFail
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                End If
                If TheExec.DevChar.Setups.IsRunning = False Then TheExec.sites.Item(site).TestNumber = TestNumber + 1
        Next site
    Else
        HardIP_WriteFuncResult
    End If
    
    
    For Each site In TheExec.sites
        DebugPrintFunc patt.Value
    Next site
    'add for retention
    If wait_time <> "" Then         ' add for wait time between patterns
        For Each site In TheExec.sites
            Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "Retention"
            ' Scenario 1 init_NV_pl1_NV_pl2_NV => set Power_Level_Last to "Sweep" to force pl2_NV change level to NV
            ' Scenario 2 init_NV_pl1_NV_pl2_Sweep => set Power_Level_Last to "Sweep" to force pl2_NV stay at "Sweep" voltage
            Power_Level_Last = "Sweep"
            print_core_power "Retention Power", Shmoo_Apply_Pin
            DebugPrintFunc patt.Value
        Next site
        If set_init = True Then
            TheExec.Datalog.WriteComment "wait " & wait_time & "after init pattern " & seq
        Else
            TheExec.Datalog.WriteComment "wait " & wait_time & "after payload pattern " & seq
        End If
        TheHdw.Wait CDbl(wait_time)
    End If
End Function
'201612 Add DigSrc Arguments/201811 Add Vbump function and SELSRM
Public Function Shmoo_Test_Pattern(ByVal patt As Pattern, ReportResult As PFType, ResultMode As tlResultMode, ConcurrentMode As tlPatConcurrentMode, Power_Run_Scenario As String, powerPin As String, set_init As Boolean, seq As Long, wait_time As String, _
                                    Optional digsrc_BitSize As String, Optional digsrc_Seg As String, Optional digsrc_DigSrcPin As String, Optional digSrc_EQ As String, _
                                    Optional RTOSRelaySwith As Boolean, _
                                    Optional allpowerpins As PinList, _
                                    Optional DecideSPIMatchLoopFlag As Boolean, _
                                    Optional SPIMatchLoopCountValue As Long, _
                                    Optional CharInputString As String, _
                                    Optional RTOSPatIndex As Integer, _
                                    Optional BlockType As String, Optional DynamicSelSrmBits As String, Optional Vbump As Boolean = False)

    Dim External_Retention As Boolean
    Dim test_name_ary() As String
    Dim SRV_type As String
    Dim block_name As String
    Dim site As Variant
    Dim lPatternCount As Long
    Dim astrPattTemp() As String
    Dim bstrPattTemp() As String
    Dim TestCase As String
    Dim DigSrc_Size As Double
    Dim DigSrc_flag As Boolean
    Dim digcap_flag As Boolean 'add for DigCap function
    Dim DigSrc_wav As New DSPWave
    Dim DigSrc_pin As New PinList
    Dim PattArray() As String
    Dim PatCount As Long
    Dim Seg_Arr() As String
    Dim Pin_Ary() As String
    Dim pin_count As Long

    Dim i As Integer
    '========================== 'add for Multi Pat function ==========================
    Dim MultiPatAry() As String
    Dim MultiPat As Boolean
    Dim MultiPatCount As Long
    Dim CountMultiPat As Long
    '========================== 'add for Multi Pat function ==========================
    
    '========================== 'add for DigCap function ============================
    Dim DigCapName() As String
    Dim DigSrcPin As String, DigCapPin As String, DigSrcSize As String, DigCapSize As String
    Dim DigCap_Info_Dict As New Dictionary
    Dim DigCap_Pin As New PinList
    Dim OutDspWave As New DSPWave
    Dim DSSC_Capture_Out As String
    '========================== 'add for DigCap function ============================
    
    '========================== 'add for SELSRM function ============================
    Dim SELSRM_Fun As Boolean
    '========================== 'add for SELSRM function ============================
    
    On Error GoTo err
    '' Add for Pattern loop ,20160607, KS
    MultiPat = False 'add for Multi Pat function
    digcap_flag = False 'add for DigCap function
    g_Retention_Shmoo = False 'add for SelSram function
    DigSrc_flag = False 'init flag
    SELSRM_Fun = False 'init SELSRM flag
    lPatternCount = 0 'initial PatternCount for pat loops
    MultiPatCount = 0 'initial Multi patterns count
    
    If patt.Value = "" Then Exit Function
    '' auto convert T_update to T_char, need to modified VBT for each project
    If Vbump = True Then ' SELSM function for debug use
        If TheExec.EnableWord("BringUp_Shmoo") = True Then
                SELSRM_Fun = True ''' to avoid VBT error while BringUp_Shmoo enable word is opening
           If InStr(UCase(patt), "DSSC") > 0 Then
              DigSrcPin = "JTAG_TDI"
              DigSrc_flag = True
               If UCase(BlockType) Like "*SOC*" Then
                  digSrc_EQ = "SSSSSSSSSSSSSSSSSSSSS"
                  DigSrcSize = "21"
               ElseIf UCase(BlockType) Like "*CPU*" Then
                  digSrc_EQ = "SSSSSSSSSSSSSSSSS"
                  DigSrcSize = "17"
               ElseIf UCase(BlockType) Like "*GPU*" Or UCase(BlockType) Like "*GFX*" Then
                  digSrc_EQ = "SSSSSSS"
                  DigSrcSize = "7"
               End If
               GoTo BringUp_Shmoo
            End If
        End If
    End If
    
    If wait_time <> "" Then g_Retention_Shmoo = True ' use for non SELSRM Function
    
    If InStr(patt, ":") > 0 Then
        astrPattTemp = Split(patt, ":")
        bstrPattTemp = Split(astrPattTemp(1), "_")
        
        '========================================================================Process SELSRM format=====================================================================
        If InStr(LCase(bstrPattTemp(0)), "selsrm") > 0 Then
           If Vbump = True Then
              SELSRM_Fun = True
              If digsrc_BitSize <> "" And digsrc_DigSrcPin <> "" And digSrc_EQ <> "" Then
                 Call Char_Process_DigString(digsrc_BitSize, digsrc_Seg, digsrc_DigSrcPin, DigCapName, DigSrcPin, DigCapPin, DigSrcSize, DigCapSize, DigCap_Info_Dict)
                 If DynamicSelSrmBits <> "" Then
                    If Not UCase(digSrc_EQ) = UCase(DynamicSelSrmBits) Then
                       TheExec.ErrorLogMessage "DynamicSelSrmBits"
                       GoTo err
                    End If
                  End If
              ElseIf DynamicSelSrmBits <> "" And digsrc_BitSize = "" And digsrc_DigSrcPin = "" Then
                  DigSrcSize = Len(DynamicSelSrmBits)
                  DigSrcPin = "JTAG_TDI"
                  digSrc_EQ = DynamicSelSrmBits
              ElseIf DynamicSelSrmBits <> "" And digsrc_BitSize <> "" And digsrc_DigSrcPin <> "" And digSrc_EQ = "" Then
                  Call Char_Process_DigString(digsrc_BitSize, digsrc_Seg, digsrc_DigSrcPin, DigCapName, DigSrcPin, DigCapPin, DigSrcSize, DigCapSize, DigCap_Info_Dict)
                  digSrc_EQ = DynamicSelSrmBits
              Else
                  TheExec.ErrorLogMessage "No Digital source for SELSRM Char"
              End If
           Else
              TheExec.ErrorLogMessage "Please enable Vbump function"
              GoTo err
           End If
        End If
        '=======================================================================Process SELSRM format=======================================================================
        
        '============================================================Process DSSC string, merge DigSrc/DigCap===============================================================
        If InStr(LCase(bstrPattTemp(0)), "digsrc") > 0 Then
           Call Char_Process_DigString(digsrc_BitSize, digsrc_Seg, digsrc_DigSrcPin, DigCapName, DigSrcPin, DigCapPin, DigSrcSize, DigCapSize, DigCap_Info_Dict)
        End If
        '============================================================Process DSSC string, merge DigSrc/DigCap===============================================================
        
        '========================================================================Mapping Table Method=======================================================================
        If (UBound(bstrPattTemp()) = 1) Then
            TestCase = bstrPattTemp(1)
            Call GetSrcString_fromEMAArray(astrPattTemp(0), TestCase, digSrc_EQ, DigSrc_Size)
            digsrc_BitSize = CStr(DigSrc_Size)
        End If
        '========================================================================Mapping Table Method=======================================================================
        
        
        If (LCase(bstrPattTemp(0)) <> "digsrc") And (LCase(bstrPattTemp(0)) <> "selsrm") Then 'Pattern loops
            lPatternCount = CLng(astrPattTemp(1)) - 1
            patt = astrPattTemp(0)
            TheExec.Datalog.WriteComment "Loop Pattern :" & patt & "_" & "Repeat count :" & lPatternCount + 1
        Else ' Create DSPWave signal
            If DigSrcPin <> "" Then DigSrc_flag = True
            If DigCapPin <> "" Then digcap_flag = True
            patt = astrPattTemp(0)
            
BringUp_Shmoo:
            Call PATT_GetPatListFromPatternSet(patt.Value, PattArray, PatCount)
            If DigSrc_flag = True Then
               Set DigSrc_wav = Nothing
               DigSrc_wav.CreateConstant 0, DigSrcSize
               
               '===================================================DSSC Switching for SELSRM Function====================================================='
               If SELSRM_Fun = True Then
                  Dim DC_Spec_Level As New PinListData
                  Dim DecodeingString As String
                  If TheExec.EnableWord("Shmoo_TTR") = True Then
                    If InStr(UCase(digSrc_EQ), "S") > 0 Then
                       If set_init = True And g_InitSeq = "" Then g_InitSeq = CStr(seq)
                    Else
                       If g_InitSeq = "" Then g_InitSeq = "Payload1"
                    End If
                  End If
                  Decide_DC_Level DC_Spec_Level, g_ApplyLevelTimingValt, g_ApplyLevelTimingVmain, BlockType
                  digSrc_EQ = Decide_Switching_Bit(digSrc_EQ, DigSrc_wav, DC_Spec_Level, BlockType, DecodeingString, powerPin, g_Globalpointval, g_ForceCond_VDD, g_CharInputString_Voltage_Dict)
               '===================================================DSSC Switching for SELSRM Function====================================================='
               Else ' Without DSSC Switching
                  For i = 0 To Len(digSrc_EQ) - 1
                     For Each site In TheExec.sites.Active
                         DigSrc_wav.Element(i) = CDbl(Mid(digSrc_EQ, i + 1, 1))
                     Next site
                  Next i
               End If
               
               DigSrc_pin.Value = DigSrcPin
               Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "FUNC_SRC", CLng(DigSrcSize), DigSrc_wav)
               
               '===========================================================Debug LVCC/HVCC/Diagnostic Char=========================================================='
               If (SELSRM_Fun = True And TheExec.Flow.EnableWord("Debug_LVCC_VminBoundary") = True) Or (SELSRM_Fun = True And TheExec.Flow.EnableWord("Debug_HVCC_VminBoundary") = True) Then
                 digSrc_EQ_GB = digSrc_EQ:: BlockType_GB = BlockType:: DigSrcSize_GB = DigSrcSize:: dssc_pat_init_GB = PattArray(0):: DigSrc_pin_GB = DigSrc_pin
'                 g_VminBoundary_selsrm = True
               End If
               '===========================================================Debug LVCC/HVCC/Diagnostic Char=========================================================='
               If SELSRM_Fun = True Then
                  If set_init Then
                     TheExec.Datalog.WriteComment "DigSrc pattern = " & "Init" & seq & ": " & patt & "," & "Src Bits = " & Len(digSrc_EQ) & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & digSrc_EQ & "," & DecodeingString
                  Else
                     TheExec.Datalog.WriteComment "DigSrc pattern = " & "Payload" & seq & ": " & patt & "," & "Src Bits = " & Len(digSrc_EQ) & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & digSrc_EQ & "," & DecodeingString
                  End If
               Else
                  If set_init Then
                     TheExec.Datalog.WriteComment "DigSrc pattern = " & "Init" & seq & ": " & patt & "," & "Src Bits = " & Len(digSrc_EQ) & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & digSrc_EQ
                  Else
                     TheExec.Datalog.WriteComment "DigSrc pattern = " & "Payload" & seq & ": " & patt & "," & "Src Bits = " & Len(digSrc_EQ) & "," & "Output String [ LSB(L) ==> MSB(R) ]:" & digSrc_EQ
                  End If
               End If
            End If
            ' ==============================================================Creat DSP wave for DigCap=============================================================
            If digcap_flag = True Then
               Set OutDspWave = Nothing
               DigCap_Pin.Value = DigCapPin
               Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, CLng(DigCapSize), OutDspWave)
               TheExec.Datalog.WriteComment ("Cap Bits = " & CLng(DigCapSize))
               TheExec.Datalog.WriteComment ("Cap Pin = " & DigCap_Pin)
               TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test End   ========")
            End If
            ' ==============================================================Creat DSP wave for DigCap=============================================================
        End If
 
    ElseIf InStr(patt, ",") > 0 Then 'Multi Pattern function
        MultiPatAry = Split(patt, ",")
        MultiPatCount = UBound(MultiPatAry)
        MultiPat = True
    End If
    

''===========================================================SET RUN LEVEL=========================================================
    If Vbump = True Then  'add for SelSram function
       Set_Run_Level_Vbump Power_Run_Scenario, powerPin, set_init, seq 'add for Vbump function
    Else
       If Not UCase(Power_Run_Scenario) Like "INIT_SWEEP_PL_SWEEP" Then
       ''no need to change voltage conditions if init_sweep_pl_sweep (it apply to correct sweep condition by IG-XL)
          Set_Run_Level Power_Run_Scenario, powerPin, set_init, seq
       End If
    End If
''===========================================================SET RUN LEVEL=========================================================

    Dim InDSPwave As New DSPWave
    Dim Count As Long
    Dim TestNumber As Long
            

    For CountMultiPat = 0 To MultiPatCount  'Multi pat function
    
        If MultiPat = True Then
           Call TheHdw.Patterns(MultiPatAry(CountMultiPat)).Load
        Else
           Call TheHdw.Patterns(patt).Load
        End If
                    
     ''-------------------------------------------
     '' HRAM setup capture on first fail 20170425
        TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
        TheHdw.Digital.HRAM.Size = 512
        TheHdw.Digital.HRAM.CaptureType = captFail
        TheHdw.Digital.HRAM.SetTrigger trigFirst, True, 0, True
     ''-------------------------------------------
     
        For Count = 0 To lPatternCount
          If MultiPat = True Then
            Call TheHdw.Patterns(MultiPatAry(CountMultiPat)).start
          Else
            Call TheHdw.Patterns(patt).start ' make sure to jump out  the cpu loop before halt
          End If
            While TheHdw.Digital.Patgen.IsRunning = True
                TheHdw.Digital.Patgen.Continue 0, cpuA
            Wend
            TheHdw.Digital.Patgen.HaltWait
        Next Count
        '------------------
        '------------------
'===============================================================================
'20190319 update
'if SuspendDatalog=false the Tname need to include all information from X,Y,Z current point
    Dim Suspend_Flag As Boolean
    Dim DevSetupName As String
    Dim TnameCombShmooInfo As String
    If TheExec.DevChar.Setups.IsRunning = True Then
'    Exit Function
        Call StoreEachPoint_PreSetup
        DevSetupName = TheExec.DevChar.Setups.ActiveSetupName
        Suspend_Flag = TheExec.DevChar.Setups.Item(DevSetupName).output.SuspendDatalog
    End If
'===============================================================================
        
        '' KS update for remove fail pins count function when do shmoo
        If TPModeAsCharz_GLB = True Then
       
            For Each site In TheExec.sites
                TestNumber = TheExec.sites.Item(site).TestNumber
                If TheHdw.Digital.Patgen.PatternBurstPassed(site) Then
'                    If TheExec.DevChar.Setups.IsRunning = True Then TheExec.sites.item(site).testResult = sitePass
                    TheExec.sites.Item(site).testResult = sitePass
                    If TheExec.DevChar.Setups.IsRunning = True Then
                        If Suspend_Flag = False Then
                            If TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 3 Then '3D
                                TnameCombShmooInfo = TheExec.DataManager.instanceName & "_" & axis_pin(0) & CStr(axis_val(0)) & "_" & axis_pin(1) & CStr(axis_val(1)) & "_" & axis_pin(2) & CStr(axis_val(2))    ''20190319 update
                            ElseIf TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 2 Then '2D
                                TnameCombShmooInfo = TheExec.DataManager.instanceName & "_" & axis_pin(0) & CStr(axis_val(0)) & "_" & axis_pin(1) & CStr(axis_val(1))
                            ElseIf TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 1 Then '1D
                                TnameCombShmooInfo = TheExec.DataManager.instanceName & "_" & axis_pin(0) & CStr(axis_val(0))
                            End If
                            Call TheExec.Datalog.WriteFunctionalResult(site, g_TestNum, logTestPass, , TnameCombShmooInfo) ''20190319 update
                        Else
                            Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                        End If
                    Else
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                    End If
                Else
                    TheExec.sites.Item(site).testResult = siteFail
                    If TheExec.DevChar.Setups.IsRunning = True Then
                        If Suspend_Flag = False Then
                            If TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 3 Then '3D
                                TnameCombShmooInfo = TheExec.DataManager.instanceName & "_" & axis_pin(0) & CStr(axis_val(0)) & "_" & axis_pin(1) & CStr(axis_val(1)) & "_" & axis_pin(2) & CStr(axis_val(2))    ''20190319 update
                            ElseIf TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 2 Then '2D
                                TnameCombShmooInfo = TheExec.DataManager.instanceName & "_" & axis_pin(0) & CStr(axis_val(0)) & "_" & axis_pin(1) & CStr(axis_val(1))
                            ElseIf TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count = 1 Then '1D
                                TnameCombShmooInfo = TheExec.DataManager.instanceName & "_" & axis_pin(0) & CStr(axis_val(0))
                            End If
                            Call TheExec.Datalog.WriteFunctionalResult(site, g_TestNum, logTestFail, , TnameCombShmooInfo) ''20190319 update
                        Else
                            Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                        End If
                    Else
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                    End If
                   ''-----------------------------------------------------------------------------------------------
    '                    If LCase(patt) Like "*rtos*" Then Call RTOS_BCS(patt, RTOSPatIndex)
                    If LCase(TheExec.DataManager.instanceName) Like "*rtos*" Then Call RTOS_BCS(patt, site, RTOSPatIndex)
                   ''------------------------------------------------------------------------------------------------
                End If
                If TheExec.DevChar.Setups.IsRunning = False Then TheExec.sites.Item(site).TestNumber = TestNumber + 1
            Next site
        Else
            HardIP_WriteFuncResult
        End If
        g_TestNum = g_TestNum + 1 ''20190319 update
    Next CountMultiPat
    
    '=============================================================Process DSP Capture out =================================================================
    If digcap_flag = True Then
       Call CreateSimulateDataDSPWave(OutDspWave, CLng(DigCapSize), CLng(DigCapSize))
       Call Char_Process_DSP_Capture(DigCapName, OutDspWave, DigCap_Info_Dict, CStr(DigCap_Pin))
    End If
     '======================================================================================================================================================
            
    For Each site In TheExec.sites
        '20170213 prevent over write shmoo pattern
        DebugPrintFunc patt.Value
    Next site
    'add for retention
    
    If Vbump = True Then 'Vbump function
        If wait_time <> "" And g_PLSWEEP = False Then
           g_shmoo_ret = True
           If TheExec.DevChar.Setups.IsRunning = True Then
              Shmoo_Restore_Power_per_site_Vbump_Retention powerPin, True
           Else
              Shmoo_Restore_Power_per_site_Vbump_Retention powerPin, False
           End If
           Power_Level_Last = ""
           If set_init = True Then
              TheExec.Datalog.WriteComment "wait " & wait_time & " after init pattern " & seq
           Else
              TheExec.Datalog.WriteComment "wait " & wait_time & " after payload pattern " & seq
           End If
           TheHdw.Wait CDbl(wait_time)
           If TheExec.Flow.EnableWord("Enable_RET_RampDownUp") = True And TheExec.DevChar.Setups.IsRunning = True Then
              Retention_RampdownUp Shmoo_Apply_Pin, "UP"
           End If
           
        ElseIf wait_time <> "" And g_PLSWEEP = True Then ' Disrtub retention function
           If set_init = True Then
              TheExec.Datalog.WriteComment "wait " & wait_time & " after init pattern " & seq
           Else
              TheExec.Datalog.WriteComment "wait " & wait_time & " after payload pattern " & seq
           End If
           TheHdw.Wait CDbl(wait_time)
        End If
        
        
    Else ' without Vbump function
        If TheExec.Flow.EnableWord("Enable_RET_RampDownUp") = False Then
    
            If wait_time <> "" Then         ' add for wait time between patterns
                    Power_Level_Last = "Sweep" '20181101 move out varint from site-loop
                    For Each site In TheExec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "Restore to Sweep V"
    
                    If TheExec.DevChar.Setups.IsRunning = False Then
                        Shmoo_Set_Retention_Power False ' for functional test
                    Else
                        Shmoo_Set_Retention_Power True  ' Skip set retention power for shmoo pin
                    End If
                    ' Scenario 1 init_NV_pl1_NV_pl2_NV => set Power_Level_Last to "Sweep" to force pl2_NV change level to NV
                    ' Scenario 2 init_NV_pl1_NV_pl2_Sweep => set Power_Level_Last to "Sweep" to force pl2_NV stay at "Sweep" voltage
                   ' Power_Level_Last = "Sweep" '20181101 move out varint from site-loop
                    print_core_power "Retention Power", Shmoo_Apply_Pin
                
                    '20170213 prevent over write shmoo pattern
                    DebugPrintFunc patt.Value
                    Next site
                    If set_init = True Then
                    TheExec.Datalog.WriteComment "wait " & wait_time & " after init pattern " & seq
                    Else
                    TheExec.Datalog.WriteComment "wait " & wait_time & " after payload pattern " & seq
                    End If
                    TheHdw.Wait CDbl(wait_time)
            End If
        
        Else
            If wait_time <> "" Then         ' add for wait time between patterns
                Dim RetPowers As Double
                Dim RetPins As New PinList
                Dim Retention_V(100) As New SiteDouble
    ''                For Each Site In theexec.sites
                  
    ''                    ' Scenario 1 init_NV_pl1_NV_pl2_NV => set Power_Level_Last to "Sweep" to force pl2_NV change level to NV
    ''                    ' Scenario 2 init_NV_pl1_NV_pl2_Sweep => set Power_Level_Last to "Sweep" to force pl2_NV stay at "Sweep" voltage
    ''                    Power_Level_Last = "Sweep"
    ''                    'print_core_power "Retention Power", Shmoo_Apply_Pin
    ''
    ''                    '20170213 prevent over write shmoo pattern
    ''                    'DebugPrintFunc patt.Value
    ''                    RetPowers = ShmooSweepPower(Site)
    ''                Next Site
    ''                RetPins = Shmoo_Apply_Pin
    
    '                Call MbistRetentionLevelWait_ForChar(CDbl(wait_time) * 1000, ShmooSweepPower(), RetPins, 10, 0)
    
                For Each site In TheExec.sites: For i = 0 To 99: Retention_V(i)(site) = 0: Next i: Next site ' initialize Retention_V array
                Decide_retetntion_power Retention_V(), RetPins
                If RetPins <> "" Then
                    Call MbistRetentionLevelWait_ForChar(CDbl(wait_time) * 1000, Retention_V(), RetPins, 10, 0)
                End If
                If set_init = True Then
                    TheExec.Datalog.WriteComment "wait " & wait_time & "after init pattern " & seq
                Else
                    TheExec.Datalog.WriteComment "wait " & wait_time & "after payload pattern " & seq
                End If
                'thehdw.Wait CDbl(wait_time)
            End If
        End If
    End If
    'On Error GoTo 0
    Exit Function
err:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function RTOS_BCS(patt As Pattern, site As Variant, Optional RTOSPatIndex As Integer)

    Dim w_CurrFailingPat As String
    Dim w_CurrFailingVector As Integer
    '20170428 add case "C", judge TestDone srm
    Dim r_TestDoneIdx As Long
    Dim PattArray() As String
    Dim PatCount As Long
    Dim i As Integer
    Dim w_BootIndex As Integer, w_BistDownIndex As Integer, w_HaltIndex As Integer, w_cmdIndex As Integer
    Dim VectorStr As String
'                    Dim VectorIndex As Integer
'                    Dim TheLastVector As Integer
    Dim w_CmdStrFlag As Boolean
    Dim PatTmp() As String
    
    
''-----------------------------------------------------------------------------------------------
'                    ''20170425
      w_CurrFailingPat = TheHdw.Digital.HRAM.PatGenInfo(0, pgPattern)
      w_CurrFailingVector = TheHdw.Digital.HRAM.PatGenInfo(0, pgVector)
      w_CmdStrFlag = True

      Call PATT_GetPatListFromPatternSet(patt.Value, PattArray, PatCount)
      PatTmp = Split(PattArray(0), ":")
      PattArray(0) = PatTmp(0)
'      If LCase(PattArray(0)) Like "*rtos*" Then ' only RTOS pattern entry
          For i = 0 To 1000
              VectorStr = TheHdw.Digital.Patterns(PattArray(0)).GetCommandString("", i)
'              If LCase(VectorStr) Like "*ready_wait_loop*" Then 'Keyword from boot done
              If LCase(VectorStr) Like "*rdywait*" Then 'Keyword from boot done
                 w_BootIndex = i + RTOSPatIndex
'                  w_BootIndex = i + 0
'              ElseIf LCase(VectorStr) Like "*cmd_done*" Then 'Keyword from command done
              ElseIf LCase(VectorStr) Like "*cmddone*" Then 'Keyword from command done
                  If w_CmdStrFlag = True Then
                      w_cmdIndex = i - 35  ' Cyprus 20170902 pat
                      w_CmdStrFlag = False
                  End If
'              ElseIf LCase(VectorStr) Like "*test_done*" Then 'Keyword from Scenrio done
              ElseIf LCase(VectorStr) Like "*tstdone*" Then 'Keyword from Scenrio done
                  w_BistDownIndex = i - 1
              ElseIf LCase(VectorStr) Like "*halt*" Then
                  w_HaltIndex = i
                  Exit For
              End If
          Next i
          
          If w_BootIndex - RTOSPatIndex = 0 And w_cmdIndex = 0 And w_BistDownIndex = 0 Then
            ShmResult(site) = ShmResult(site) & "-"
            GoTo bypass
          End If
    ''
           'Judge and record test result character to shmoo result string
    '     SRM
'          If w_BootIndex - RTOSPatIndex = 0 And w_cmdIndex = 0 And w_BistDownIndex = 0 Then
'              ShmResult(Site) = ShmResult(Site) & "-"
'              theexec.Datalog.WriteComment "RTOS_BCS bypassed due to pattern keyword issue."
          If LCase(w_CurrFailingPat) Like "*rdywait*" Then
              ShmResult(site) = ShmResult(site) & "B"
          ElseIf LCase(w_CurrFailingPat) Like "*cmddone*" Then
             ShmResult(site) = ShmResult(site) & "C"
          ElseIf LCase(w_CurrFailingPat) Like "*tstdone*" Then
             ShmResult(site) = ShmResult(site) & "S"
          Else 'VM
              If w_CurrFailingVector <= w_BootIndex Then
                 ShmResult(site) = ShmResult(site) & "B"
              ElseIf w_CurrFailingVector > w_BootIndex And w_CurrFailingVector < w_BistDownIndex Then
                 ShmResult(site) = ShmResult(site) & "C"
              ElseIf w_CurrFailingVector > w_BistDownIndex Then
                 ShmResult(site) = ShmResult(site) & "S"
              End If
          End If
'      End If
'----------------------------------------------------------------------------------------------
bypass:
End Function


'20170104 Roy modified
'Replace module from "Select case" to "If-else"

Public Function Set_Run_Level(Power_Run_Scenario As String, powerPin As String, set_init As Boolean, seq As Long)
    Dim VoltageLevel As String, Scenario As String
    Dim i As Long
    Dim init_level As String
    Dim pl_level As String
    Dim Power_Run_Scenario_ary() As String
    Dim inst_name As String
    Dim inst_level As String
    Dim site As Variant
    
    Power_Run_Scenario_ary = Split(Power_Run_Scenario, "_")
    inst_name = TheExec.DataManager.instanceName
    inst_level = Right(TheExec.DataManager.instanceName, 2)
    init_level = "-99"
    pl_level = "-99"
    ' Init_NV_pl_Sweep, NV test
    '       last="", init=NV
    '       last=NV, init=NV
    '       last=NV, init=NV
    '       last=NV, pl=Sweep
    '
    '       last=Sweep, init=NV
    '       last=NV, init=NV
    '       last=NV, init=NV
    '       last=NV, pl=Sweep
    If set_init = True Then
            
        If LCase(Power_Run_Scenario) Like LCase("*Init_Sweep*") Then
            init_level = "Sweep"
            If Not (Power_Level_Last Like init_level) Then
                For Each site In TheExec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** Char_Init" & seq & "_" & inst_level & "_Sweep ***"
                Next site
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*Init_[NHL]V*") Then
            init_level = Mid(Power_Run_Scenario, InStr(LCase(Power_Run_Scenario), "init_") + 5, 2)
            If Not (Power_Level_Last Like init_level) Then Shmoo_Set_Power Shmoo_Apply_Pin, init_level, "*** Char_Init" & seq & "_" & init_level & " ***", True
        ElseIf LCase(Power_Run_Scenario) Like LCase("*init" & seq & "_Sweep*") Then
            init_level = "Sweep"
            If Not (Power_Level_Last Like init_level) Then
                For Each site In TheExec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** Char_Init" & seq & "_" & inst_level & "_Sweep ***"
                Next site
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*init" & seq & "_[NHL]V*") Then
            init_level = Mid(Power_Run_Scenario, InStr(LCase(Power_Run_Scenario), "init" & seq & "_") + 6, 2)
            If Not (Power_Level_Last Like init_level) Then Shmoo_Set_Power Shmoo_Apply_Pin, init_level, "*** Char_Init" & seq & "_" & init_level & " ***", True
        End If
        Power_Level_Last = init_level
        If init_level Like "-99" Then TheExec.ErrorLogMessage "Power Run Scenario " & Power_Run_Scenario & " is not supported"
    Else
            
        If LCase(Power_Run_Scenario) Like LCase("*pl_Sweep*") Then
            pl_level = "Sweep"
            If Not (Power_Level_Last Like pl_level) Then
                For Each site In TheExec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** PL" & seq & "_Sweep ***"
                Next site
            Else
                For Each site In TheExec.sites
                    print_core_power "*** PL" & seq & "_Sweep ***", Shmoo_Apply_Pin
                Next site
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl_[NHL]V*") Then
            pl_level = Mid(Power_Run_Scenario, InStr(LCase(Power_Run_Scenario), "pl_") + 3, 2)
            If g_Retention_Shmoo = True Then
               'For retention payload, use force condition instead of N/L/HV for force pin
                'Modify for force condition "VRET" 20171213
                    If g_ForceCond_VDD <> "" Or g_Retention_VDD <> "" Then
                        Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** PL" & seq & "_" & pl_level & " ***" & pl_level & " Force***", g_ForceCond_VDD
                    End If
                Shmoo_Set_Power Shmoo_Apply_Pin, pl_level, "*** PL" & seq & "_" & pl_level & " ***", True, g_ForceCond_VDD
            Else
                If Not (Power_Level_Last Like pl_level) Then
                    Shmoo_Set_Power Shmoo_Apply_Pin, pl_level, "*** PL" & seq & "_" & pl_level & " ***", True
'                Else
'                    print_core_power "*** PL" & seq & "_" & pl_level & " ***", Shmoo_Apply_Pin
                End If
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_Sweep*") Then
            pl_level = "Sweep"
            If Not (Power_Level_Last Like pl_level) Then
                For Each site In TheExec.sites
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** PL" & seq & "_Sweep ***"
                Next site
            Else
                For Each site In TheExec.sites
                    print_core_power "*** PL" & seq & "_Sweep ***", Shmoo_Apply_Pin
                Next site
            End If
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_[NHL]V*") Then
            pl_level = Mid(Power_Run_Scenario, InStr(LCase(Power_Run_Scenario), "pl" & seq & "_") + 4, 2)
            If g_Retention_Shmoo = True Then
               'For retention payload, use force condition instead of N/L/HV for force pin
                'Modify for force condition "VRET" 20171213
                If g_ForceCond_VDD <> "" Or g_Retention_VDD <> "" Then
                    Shmoo_Restore_Power_per_site Shmoo_Apply_Pin, ShmooSweepPower, "*** PL" & seq & "_" & pl_level & " Force***", g_ForceCond_VDD
                End If
                Shmoo_Set_Power Shmoo_Apply_Pin, pl_level, "*** PL" & seq & "_" & pl_level & " ***", True, g_ForceCond_VDD
             Else
                If Not (Power_Level_Last Like pl_level) Then
                    Shmoo_Set_Power Shmoo_Apply_Pin, pl_level, "*** PL" & seq & "_" & pl_level & " ***", True
'                Else
'                    print_core_power "*** PL" & seq & "_" & pl_level & " ***", Shmoo_Apply_Pin
                End If
            End If
        End If
           
        Power_Level_Last = pl_level
        If pl_level Like "-99" Then TheExec.ErrorLogMessage "Power Run Scenario " & Power_Run_Scenario & " is not supported"
    End If
    
End Function

Public Function gen_search_string(SetupName As String, ByRef Search_String As String, axis_type As tlDevCharShmooAxis, Optional RangeFrom As Double, Optional RangeTo As Double, Optional RangeStepSize As Double, Optional RangeSteps As Long)
On Error GoTo errHandler
''
    Search_String = ""
    ShmooPowerName = ""
    Dim axis_header As String, p As Variant
    Dim RangeFromTracking As Variant, RangeToTracking As Variant, RangeStepSizeTracking As Variant, RangeStepsTracking As Variant
    Dim StepName As Variant, Pin_Ary() As String, shmoo_pin_string As String, PinName As Variant
    Dim StepNameTrack As Variant
    Dim Search_String_Main As String
    Dim Search_String_Tracking As String
    
    Select Case axis_type
            Case tlDevCharShmooAxis_X: axis_header = "X@"
            Case tlDevCharShmooAxis_Y: axis_header = "Y@"
    End Select
    
    If TheExec.DevChar.Setups(SetupName).output.Format Like "SwapXY" Then
        Select Case axis_type
                Case tlDevCharShmooAxis_X: axis_header = "Y@"
                Case tlDevCharShmooAxis_Y: axis_header = "X@"
        End Select
    Else
        Select Case axis_type
                Case tlDevCharShmooAxis_X: axis_header = "X@"
                Case tlDevCharShmooAxis_Y: axis_header = "Y@"
        End Select
    End If
    
    With TheExec.DevChar
''        StepName = .Setups(SetupName).Shmoo.Axes(axis_type).StepName
''        RangeFrom = .Setups(SetupName).Shmoo.Axes(axis_type).Parameter.Range.From
''        RangeTo = .Setups(SetupName).Shmoo.Axes(axis_type).Parameter.Range.To
''        RangeSteps = .Setups(SetupName).Shmoo.Axes(axis_type).Parameter.Range.steps
''        If RangeSteps = 0 Then RangeSteps = 1
''        If RangeSteps > 0 Then
''           RangeStepSize = (RangeTo - RangeFrom) / RangeSteps
''        Else
''            If RangeStepSize <> 0 Then
''                RangeSteps = (RangeTo - RangeFrom) / RangeStepSize
''            Else
''                RangeSteps = 1
''            End If
''        End If
        RangeStepsTracking = RangeSteps 'tracking steps is the same as main
        If .Setups(SetupName).Shmoo.Axes(axis_type).ApplyTo.Pins <> "" Then
            Pin_Ary = Split(.Setups(SetupName).Shmoo.Axes(axis_type).ApplyTo.Pins, ",")
            shmoo_pin_string = .Setups(SetupName).Shmoo.Axes(axis_type).ApplyTo.Pins
            For Each PinName In Pin_Ary
            ShmooPowerName = ShmooPowerName & "_" & PinName
                 Search_String_Main = Search_String_Main & axis_header & PinName & "="                                      ' need to modify 0.0000
                 Search_String_Main = Search_String_Main & Format(RangeFrom, "0.0000########") & ":"                                ' need to modify 0.0000
                 Search_String_Main = Search_String_Main & Format(RangeTo, "0.0000########") & ":"                                  ' need to modify 0.0000
                 Search_String_Main = Search_String_Main & Format(RangeStepSize, "0.0000########") & ","                            ' need to modify 0.0000
            Next PinName
        ElseIf LCase(.Setups.Item(SetupName).Shmoo.Axes.Item(axis_type).Parameter.Type) Like "*spec" Then
        ShmooPowerName = ShmooPowerName & "_" & PinName
            PinName = .Setups.Item(SetupName).Shmoo.Axes.Item(axis_type).Parameter.Name
            Search_String_Main = Search_String_Main & axis_header & PinName & "="
            Search_String_Main = Search_String_Main & Format(RangeFrom, "0.0000") & ":"                                     ' need to modify 0.0000
            Search_String_Main = Search_String_Main & Format(RangeTo, "0.0000") & ":"                                       ' need to modify 0.0000
            Search_String_Main = Search_String_Main & Format(RangeStepSize, "0.0000") & ","                                 ' need to modify 0.0000
        End If
        With .Setups.Item(SetupName).Shmoo.Axes.Item(axis_type).TrackingParameters
            For Each StepNameTrack In .List
                RangeFromTracking = .Item(StepNameTrack).range.from
                RangeToTracking = .Item(StepNameTrack).range.To
                RangeStepSizeTracking = (RangeToTracking - RangeFromTracking) / RangeStepsTracking
                If .Item(StepNameTrack).ApplyTo.Pins <> "" Then
                       Pin_Ary = Split(.Item(StepNameTrack).ApplyTo.Pins, ",")
                       shmoo_pin_string = shmoo_pin_string & "," & .Item(StepNameTrack).ApplyTo.Pins
                       For Each p In Pin_Ary
                       ShmooPowerName = ShmooPowerName & "_" & p
                          Search_String_Tracking = Search_String_Tracking & axis_header & p & "="
                          Search_String_Tracking = Search_String_Tracking & Format(RangeFromTracking, "0.0000########") & ":"       ' need to modify 0.0000
                          Search_String_Tracking = Search_String_Tracking & Format(RangeToTracking, "0.0000########") & ":"         ' need to modify 0.0000
                          Search_String_Tracking = Search_String_Tracking & Format(RangeStepSizeTracking, "0.0000########") & ","   ' need to modify 0.0000
                       Next p
                ElseIf .Item(StepNameTrack).Type Like "*Spec" Then
                    PinName = .Item(StepNameTrack).Name
                    ShmooPowerName = ShmooPowerName & "_" & PinName
                    Search_String_Tracking = Search_String_Tracking & axis_header & PinName & "="
                    Search_String_Tracking = Search_String_Tracking & Format(RangeFromTracking, "0.0000") & ":"             ' need to modify 0.0000
                    Search_String_Tracking = Search_String_Tracking & Format(RangeToTracking, "0.0000") & ":"               ' need to modify 0.0000
                    Search_String_Tracking = Search_String_Tracking & Format(RangeStepSizeTracking, "0.0000") & ","         ' need to modify 0.0000
                End If
            Next StepNameTrack
        End With
        Search_String = Search_String_Tracking & Search_String_Main
     End With
Exit Function
errHandler:
                If AbortTest Then Exit Function Else Resume Next
End Function




''Public Function char_flow()
''
''    Dim SetupName As String
''    Dim i As Long
''    Dim char_enable_idx As Long
''    On Error GoTo errHandler
''
''    If TheExec.DevChar.Setups.IsRunning = True Then Exit Function
''    If TheExec.Flow.EnableWord("Char_Flow") = True Then
''        char_flow_enable_idx = char_flow_enable_key(TheExec.DataManager.InstanceName)
''        For i = 1 To char_flow_enable_entry(char_flow_enable_idx).Count
''            Select Case char_flow_enable_entry(char_flow_enable_idx + i - 1).Enable
''                Case "Enable":
''                    Call run_shmoo(char_flow_enable_entry(char_flow_enable_idx + i - 1).CharSetup)
''                Case "Pass_Enable":
''                    If (thehdw.digital.Patgen.PatternBurstPassed = True) Then
''                        Call run_shmoo(char_flow_enable_entry(char_flow_enable_idx + i - 1).CharSetup)
''                    End If
''                Case "Fail_Enable":
''                    If (thehdw.digital.Patgen.PatternBurstPassed = False) Then
''                        Call run_shmoo(char_flow_enable_entry(char_flow_enable_idx + i - 1).CharSetup)
''                    End If
''                Case "Disable":
''                Case Default:
''                    TheExec.AddOutput ("Error!!" & char_flow_enable_entry(char_flow_enable_idx + i - 1).Enable & "is not supported in sheet Char_Flow_Enable ")
''            End Select
''        Next i
''    End If
''    If TheExec.Flow.EnableWord("Debug_Shmoo") = True Then
''        char_enable_idx = char_enable_key(TheExec.DataManager.InstanceName)
''        For i = 1 To char_enable_entry(char_enable_idx).Count
''            Select Case char_enable_entry(char_enable_idx + i - 1).Enable
''                Case "Enable":
''                    Call run_shmoo(char_enable_entry(char_enable_idx + i - 1).CharSetup)
''                Case "Pass_Enable":
''                    If (thehdw.digital.Patgen.PatternBurstPassed = True) Then
''                        Call run_shmoo(char_enable_entry(char_enable_idx + i - 1).CharSetup)
''                    End If
''                Case "Fail_Enable":
''                    If (thehdw.digital.Patgen.PatternBurstPassed = False) Then
''                        Call run_shmoo(char_enable_entry(char_enable_idx + i - 1).CharSetup)
''                    End If
''                Case "Disable":
''                Case Default:
''                    TheExec.AddOutput ("Error!!" & char_enable_entry(char_enable_idx + i - 1).Enable & "is not supported in sheet Char_Enable ")
''            End Select
''        Next i
''    End If
''    Exit Function
''errHandler:
''    Exit Function
''End Function

Public Function ShmooPostStep2Dto1D(argc As Long, argv() As String)

    Dim SetupName As String
    Dim i As Long
    Dim OutputString As String
    Dim instanceName As String
    Dim TestNum As Long
    Dim lvccf As Integer
    Dim LVCC As Double
    Dim site As Variant
'    Dim v_Xi0 As Double
'    Dim xio_spec As String
    Dim TestVoltage As String
    Dim StartVoltage As Double, EndVoltage As Double, StepSize As Double
    Dim Patt_String As String, Shmoo_Result As String
    Dim Pat As Variant
    Dim PinName As Variant
    Dim StepName As Variant
    Dim RangeFrom As Double, RangeTo As Double, RangeStepSize As Double, RangeSteps As Long
    Dim RangeLow As Double
    Dim RangeCalcType As tlDevCharRangeField
    Dim allpowerpins As String
    Dim PowerPinCnt As Long, PowerPinAry() As String
    Dim FlagFirstPass As Boolean
    Dim last_point_result As tlDevCharResult, current_point_result As tlDevCharResult
    Dim min_point As Long, max_point As Long, current_point As Long
    Dim Vcc_min As String, Vcc_max As String
    Dim patt_ary() As String, pat_count As Long, p As Variant
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim shmoo_pin_string As String
    Dim tmp As String
    Dim Search_String As String
    Dim FlagHole As Boolean
    Dim Shmoo_hole As String
    Dim FlagPF(1000) As Boolean
    Dim FlagFP(1000) As Boolean
    Dim FlagPF_Count As Long
    Dim FlagFP_Count As Long
    Dim ch As String
    Dim Group As Boolean
    Dim Label As String
    Dim step_Start As Long
    Dim step_Stop As Long
    Dim Step_x As Long
    Dim Range_temp As Double
    Dim range_plus As Long
    Dim Shmoo_Pattset As New Pattern
    Dim CharShmooAxis_Inter As tlDevCharShmooAxis
    Dim CharShmooAxis_Outer As tlDevCharShmooAxis
    
    Dim CharShmooType_X As String
    Dim CharShmooType_Y As String
    Dim patset As Variant, patset1 As Variant, j As Long
    Dim Outer_StepName As Variant
    Dim Outer_RangeFrom As Double, Outer_RangeTo As Double, Outer_RangeStepSize As Double, Outer_RangeSteps As Long
    Dim Outer_ParameterName As String   '20180716 Auto parsing FRC info
    Dim Outer_Step_start As Long
    Dim Outer_Step_stop As Long
    Dim Outer_Step_Index As Long
    Dim j_Outer  As Long
    
    Dim HIO_PinName_Updated As Boolean      '20180515 TER
    
    Dim index As Long
    Dim vbump_value As String
        
    On Error GoTo errHandler_shmoo
    
    Shmoo_hole = "NH"
    
    instanceName = TheExec.DataManager.instanceName     '20180616 TER
    Call Get_Tname_FromFlowSheet(instanceName, HIO_PinName_Updated)      '20180515 TER
    
    For Each site In TheExec.sites
        OutputString = ""
        lvccf = 0

        
        If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = False Then    '20180718 add
            Call TheExec.sites(site).IncrementTestNumber
        End If
        
        TestNum = TheExec.sites(site).TestNumber
        
        If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = True Then    '20180718 add
            Call TheExec.sites(site).IncrementTestNumber
        End If
        
        'xio_spec = "XI0_Freq_VAR"

        
'        v_Xi0 = TheExec.specs.AC(xio_spec).CurrentValue
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        SetupName = TheExec.DevChar.Setups.ActiveSetupName
        With TheExec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.To
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.CalculatedField
            

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
               range_plus = -1
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step
                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
                End If
                RangeLow = RangeTo
                range_plus = 1
            End If
        End With
        
        Patt_String = ""
        
        With TheExec.DevChar.Results(SetupName).Shmoo
            FlagPF_Count = 1
            FlagFP_Count = 1
            For i = 0 To 9
                FlagPF(i) = False
                FlagFP(i) = False
            Next i
        End With
        j = 0
            
        Shmoo_Pattset.Value = Shmoo_Pattern
        Patt_String = PatSetToPat(Shmoo_Pattset)
    
        CharShmooType_X = TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value
        CharShmooType_Y = TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Type.Value
        
        If CharShmooType_X = "Level" Or CharShmooType_X = "DC Spec" Then
            CharShmooAxis_Inter = tlDevCharShmooAxis_X
            CharShmooAxis_Outer = tlDevCharShmooAxis_Y
        ElseIf CharShmooType_Y = "Level" Or CharShmooType_Y = "DC Spec" Then
            CharShmooAxis_Inter = tlDevCharShmooAxis_Y
            CharShmooAxis_Outer = tlDevCharShmooAxis_X
        Else '' 20180710 to avoid wrong result in output string when 2D shmoo is not include power pin
            CharShmooAxis_Inter = tlDevCharShmooAxis_X
            CharShmooAxis_Outer = tlDevCharShmooAxis_Y
        End If
        
        With TheExec.DevChar
            Outer_StepName = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).StepName
            Outer_RangeFrom = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.from
            Outer_RangeTo = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.To
            Outer_RangeSteps = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.Steps + 1
            Outer_ParameterName = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.Name.Value '20180716 Auto parsing FRC info
        End With
        
        If Outer_RangeFrom < Outer_RangeTo Then
            Outer_Step_start = 0
            Outer_Step_stop = Outer_RangeSteps - 1
            Outer_Step_Index = 1
        Else
            Outer_Step_start = Outer_RangeSteps - 1
            Outer_Step_stop = 0
            Outer_Step_Index = -1
        End If
         
         
        For j_Outer = Outer_Step_start To Outer_Step_stop Step Outer_Step_Index
        
            Search_String = ""
                                            ''tlDevCharShmooAxis_X
            With TheExec.DevChar
                StepName = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).StepName                                ''tlDevCharShmooAxis_X
                RangeFrom = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).Parameter.range.from            ''tlDevCharShmooAxis_X
                RangeTo = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).Parameter.range.To                    ''tlDevCharShmooAxis_X
                RangeSteps = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).Parameter.range.Steps + 1     ''tlDevCharShmooAxis_X
                               
                If RangeFrom < RangeTo Then
                    step_Start = 0
                    step_Stop = RangeSteps - 1
                    Step_x = 1
                    range_plus = -1
                Else
                     step_Start = RangeSteps - 1
                    step_Stop = 0
                    Step_x = -1
                    range_plus = 1
                End If
                
                If RangeSteps = 0 Then RangeSteps = 1
                If RangeSteps > 0 Then
                   RangeStepSize = (RangeTo - RangeFrom) / (RangeSteps - 1)
                Else
                    If RangeStepSize <> 0 Then
                        RangeSteps = (RangeTo - RangeFrom) / RangeStepSize + 1
                    Else
                        RangeSteps = 1
                    End If
                End If
                gen_search_string SetupName, Search_String, CharShmooAxis_Inter, RangeFrom, RangeTo, RangeStepSize, RangeSteps
                shmoo_pin_string = .Setups(SetupName).Shmoo.Axes(CharShmooAxis_Inter).ApplyTo.Pins    ''tlDevCharShmooAxis_X
                Shmoo_Result = ""
        ''              Dim debug_str As String
        ''              debug_str = ""
        ''
        ''              For i = Step_start To Step_stop Step Step_x
        ''              current_point_result = .Points(i).ExecutionResult
        ''
        ''
        ''
        ''               If i = 4 Then
        ''               i = i
        ''               End If
        ''
        ''                  Select Case current_point_result
        ''                    Case tlDevCharResult_Pass:  Shmoo_Result = Shmoo_Result & "+"
        ''                    Case tlDevCharResult_Fail:  Shmoo_Result = Shmoo_Result & "-"
        ''                    Case tlDevCharResult_NoTest:
        ''                                   Shmoo_Result = Shmoo_Result & "_"
        ''                                   current_point_result = last_point_result
        ''                    Case tlDevCharResult_AssumedPass:
        ''                                   Shmoo_Result = Shmoo_Result & "*"
        ''                                   current_point_result = last_point_result
        ''                    Case tlDevCharResult_AssumedFail:
        ''                                   Shmoo_Result = Shmoo_Result & "~"
        ''                                   current_point_result = last_point_result
        ''                    Case Default:  Shmoo_Result = Shmoo_Result & "?"
        ''
        ''                  End Select
        ''
        ''
        ''              Next i
        ''
        ''               theexec.Datalog.WriteComment Shmoo_Result
     ''' Debug
    '''                  ch = Mid("---------------------------------------------------------------------------------------------------------------------------------------------------------", i + 1, 1)
    '''                  ch = Mid("--------------------------------------------------------------------------------------------------------------------------------------------++++---------", i + 1, 1)
    '''                  ch = Mid("--------------------------------------------------------------------------------------------------------------------------------------------++++----+++++", i + 1, 1)
    '''                  ch = Mid("--------------------------------------------------------------------------------------------------------------------------------------------++++-+++++---", i + 1, 1)
    '''                    LotId = "N99G19"
    '''                    WaferId = 1
    '''                    XCoord(site) = 16
    '''                    YCoord(site) = 7
    '''                    v_XI0 = 24000000#
    '''                    ch = Mid("-------------------------------------------------------------------------+********+-", i + 1, 1)
    '''                    ch = Mid("-------------------------------------------------------------------------+*********+", i + 1, 1)
    '''                    ch = Mid("------------------------------------------------------------------------------------", i + 1, 1)
    '''                    ch = Mid("-------------------------------------------------------------------------+----------", i + 1, 1)
    '''                    ch = Mid("-------------------------------------------------------------------------++---------", i + 1, 1)
    '''                    ch = Mid("-------------------------------------------------------------------------++-++------", i + 1, 1)
    '''                    ch = Mid("-------------------------------------------------------------------------++-++--++++", i + 1, 1)
    '''                    ch = Mid("+++-~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~-", i + 1, 1)
    '''                     ch = Mid("+**********************************************************************************+", i + 1, 1)
    '''                    Select Case ch
    '''                       Case "+":
    '''                          Shmoo_Result = Shmoo_Result & "+"
    '''                          current_point_result = tlDevCharResult_Pass
    '''                       Case "-":
    '''                          Shmoo_Result = Shmoo_Result & "-"
    '''                          current_point_result = tlDevCharResult_Fail
    '''                       Case "*": 'assume pass
    '''                          Shmoo_Result = Shmoo_Result & "*"
    '''                          current_point_result = last_point_result
    '''                       Case "~": 'assume fail
    '''                          Shmoo_Result = Shmoo_Result & "~"
    '''                          current_point_result = last_point_result
    '''                       Case Default:  Shmoo_Result = Shmoo_Result & "?"
    '''                    End Select
    '                                    Shmoo_Result = Shmoo_Result & .CurrentPoint.Data.StandardRecords(()
    '                   --++++++++++++++++++--------++++++++---------+++++++++++++------
    '
    '                     1stFP            1stPF     2ndFP 2ndPF     3rdFP       3rdPF      <----- BH
    '                   --++++++++++++++++++--------++++++++----------------------------
    '                     1stFP            1stPF     2ndFP 2ndPF                            <----- LH/HH
    '
    '                   -----------------------++++++++++++++++++-----------------------
    '                                          1stFP            1stPF                       <----- NH
    '
               With TheExec.DevChar.Results(SetupName).Shmoo
                    min_point = 999
                    max_point = 999
                    current_point_result = tlDevCharResult_Fail
                    last_point_result = tlDevCharResult_Fail
                    FlagFirstPass = False
                       
                      ' For i = 0 To RangeSteps - 1
                      
                        For i = step_Start To step_Stop Step Step_x
                            If CharShmooType_X = "Level" Or CharShmooType_X = "DC Spec" Then
                                current_point_result = .Points(i, j_Outer).ExecutionResult
                            ElseIf CharShmooType_Y = "Level" Or CharShmooType_Y = "DC Spec" Then
                                current_point_result = .Points(j_Outer, i).ExecutionResult
                            Else '' 20180710 to avoid wrong result in output string when 2D shmoo is not include power pin
                                current_point_result = .Points(i, j_Outer).ExecutionResult
                            End If
                            
                            Select Case current_point_result
                            Case tlDevCharResult_Pass:
                                        Shmoo_Result = Shmoo_Result & "+"
                            
                            Case tlDevCharResult_Fail:
                                        Shmoo_Result = Shmoo_Result & "-"
                            
                            Case tlDevCharResult_NoTest:
                                        Shmoo_Result = Shmoo_Result & "_"
                                        current_point_result = last_point_result
                            
                            Case tlDevCharResult_AssumedPass:
                                        Shmoo_Result = Shmoo_Result & "*"
                                        current_point_result = last_point_result
                            
                            Case tlDevCharResult_AssumedFail:
                                        Shmoo_Result = Shmoo_Result & "~"
                                        current_point_result = last_point_result
                            Case Else:
                                        Shmoo_Result = Shmoo_Result & "?"
        
                            End Select
        
                        If last_point_result = tlDevCharResult_Fail And current_point_result = tlDevCharResult_Pass Then
                            FlagFP(FlagFP_Count) = True
                            FlagFP_Count = FlagFP_Count + 1
                        End If
                        
                        If last_point_result = tlDevCharResult_Pass And current_point_result = tlDevCharResult_Fail Then
                            FlagPF(FlagPF_Count) = True
                            FlagPF_Count = FlagPF_Count + 1
                        End If
                      
                        If current_point_result = tlDevCharResult_Pass And FlagFirstPass = False Then  'find first pass point
                        'If current_point_result = tlDevCharResult_Pass And last_point_result = tlDevCharResult_Fail Then  'find last F-> P
                            min_point = i
                            FlagFirstPass = True 'always take the first pass point
                        End If
                      
                        'If current_point_result = tlDevCharResult_Pass And FlagFirstPass = False Then  'find first pass point
                        If current_point_result = tlDevCharResult_Pass And last_point_result = tlDevCharResult_Fail Then  'find last F-> P
                            min_point = i
    '                       FlagFirstPass = True 'always take the first pass point
                        End If
                      
                        If current_point_result = tlDevCharResult_Fail And last_point_result = tlDevCharResult_Pass Then       'find last pass point
                            max_point = i + range_plus 'always take the last pass point
                        End If
                       
                        last_point_result = current_point_result
                    Next i
                End With
                
                If FlagFP(1) = True And FlagFP(2) = False Then
                    Shmoo_hole = "NH"
                End If
                
                If FlagFP(1) = True And FlagFP(2) = True Then
                    Shmoo_hole = "LH"
                End If
                
                If FlagFP(1) = True And FlagFP(2) = True And FlagFP(3) = True Then
                    Shmoo_hole = "BH"
                End If
                
                If min_point <> 999 Then
                    Vcc_min = CStr(RangeFrom + min_point * RangeStepSize)
                Else
                    Vcc_min = "N/A"
                End If
                
                If max_point <> 999 Then
                    Vcc_max = CStr(RangeFrom + max_point * RangeStepSize)
                Else
                    If Vcc_min <> "N/A" Then
                        If range_plus = 1 Then
                            Vcc_max = Format(RangeFrom, "0.000")
                        Else
                            Vcc_max = Format(RangeTo, "0.000")
                        End If
                   Else
                       Vcc_max = "N/A"
                   End If
                End If
                
                If last_point_result = tlDevCharResult_Pass Then
                    If range_plus = 1 Then
                        Vcc_max = Format(RangeFrom, "0.000")
                    Else
                        Vcc_max = Format(RangeTo, "0.000")
                    End If
                    '  Vcc_max = CStr(RangeTo)
                End If
                
                '  If RangeFrom > RangeTo Then
                '     tmp = Vcc_max
                '     Vcc_max = Vcc_min
                '     Vcc_min = tmp
                '  End If
            End With
            If InStr(TheExec.DataManager.instanceName, "_NV") Then TestVoltage = "NV"
            If InStr(TheExec.DataManager.instanceName, "_HV") Then TestVoltage = "HV"
            If InStr(TheExec.DataManager.instanceName, "_LV") Then TestVoltage = "LV"

    '[Char,N99G19-1,16,7,V,0,XI0=24000000,CpuBira_P0001_IN02_BIR_SI_PL00_CL16_BIR_59N_SI_PP_NV,CPU_BIST_CPU_Domain_CPU_SRAM_Domain_P1_Full_Range,1069,
    '.\pattern\CpuMbist\PP_FIJA0_C_IN00_XX_CLXX_XXX_XXX_XXX_P0001_1308131609_SI_mod.pat,.\pattern\CpuMbist\PP_FIJA0_C_IN02_BI_CLXX_BIR_JTG_XXX_ALLFV_1306250000_SI.pat,.\pattern\CpuMbist\PP_FIJA0_C_PL00_BI_CL16_BIR_JTG_59N_ALLFV_1306250000_SI.pat,
    'NV,VDD_FIXED=0.528:1.404:0.005,VDD_VAR_SOC_VAR=0.500:1.330:0.005,
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++,NH,0.5,1.260]
            
            If Shmoo_header = "" Then Shmoo_header = "Char"
            OutputString = OutputString & "[" & Shmoo_header & "," & HramLotId(site) & "-" & CStr(HramWaferId(site)) & "," & CStr(XCoord(site)) & "," & CStr(YCoord(site))
    
            Dim SetupName_New As String, k As Integer
            Dim InstanceName_New As String
            
            SetupName_New = SetupName
    
            'Shmoo_header
            Dim VIL_Flag As Boolean
            
            VIL_Flag = False
            ShmooPowerName = ShmooPowerName
            'v_Xi0 = thehdw.
            
                    


        '20180716 Auto parsing FRC info
        Dim nWire_port_ary() As String
        Dim nwp As Variant
        Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
        Dim FRC_Name As String, FRC_Value As Double, All_FRC_Status As String
        All_FRC_Status = ""
        nWire_port_ary = Split(nWire_Ports_GLB, ",")
        For Each nwp In nWire_port_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            If TheHdw.Protocol.ports(port_pa).Enabled = True Then
                FRC_Name = Replace(UCase(ac_spec_pa), "_FREQ_VAR", "")
                If UCase(Outer_ParameterName) = UCase(ac_spec_pa) Then
                    FRC_Value = Outer_RangeFrom + j_Outer * Outer_Step_Index * TheExec.DevChar.Setups(SetupName).Shmoo.Axes(CharShmooAxis_Outer).Parameter.range.StepSize
                Else
                    FRC_Value = TheExec.specs.AC(ac_spec_pa).CurrentValue
                End If
                If All_FRC_Status = "" Then
                    All_FRC_Status = FRC_Name & "=" & FRC_Value
                Else
                    All_FRC_Status = All_FRC_Status & ";" & FRC_Name & "=" & FRC_Value
                End If
            End If
        Next nwp
        If FRC_Name = "" Then ' Default use XI0, if no input of FRC info
            FRC_Name = "XI0"
            FRC_Value = TheExec.specs.AC("XI0_Freq_VAR").CurrentValue
            All_FRC_Status = FRC_Name & "=" & FRC_Value
        End If
        
        

           ' v_Xi0 = Outer_RangeFrom + j_Outer * Outer_Step_Index * TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Range.StepSize      '20180627 TER
           
            OutputString = OutputString & ",V," & site & "," & All_FRC_Status & ","    '20180716 Auto parsing FRC info

            OutputString = OutputString & TheExec.DataManager.instanceName & ShmooPowerName & "," & SetupName_New & "," & CStr(TestNum) & ","       '20180616 TER

            OutputString = OutputString & Patt_String & ","
            OutputString = OutputString & TestVoltage & ","
            
            If argv(0) <> Empty Then
                TheExec.DataManager.DecomposePinList argv(0), Pin_Ary, Pin_Cnt
                PinName = argv(0) 'setup voltage
            End If
            
            If Vbump_for_Interpose = True Then
                Dim PL_DC_conditions_str As String
                PL_DC_conditions_str = Replace(PL_DC_conditions_GLB, ":V:", "=")
                PL_DC_conditions_str = Replace(PL_DC_conditions_str, ";", ",")
                OutputString = OutputString & PL_DC_conditions_str
            
            Else
                For j = 0 To Pin_Cnt - 1
                    PinName = Pin_Ary(j)
                    If TheExec.DataManager.ChannelType(PinName) <> "N/C" Then
                        If j = 0 Then
                            OutputString = OutputString & PinName & "=" & Format(TheHdw.DCVS.Pins(PinName).Voltage.Main.Value, "0.000")
                        Else
                             OutputString = OutputString & "," & PinName & "=" & Format(TheHdw.DCVS.Pins(PinName).Voltage.Main.Value, "0.000")
                        End If
                    End If
                Next j
            End If
            
            For i = 1 To argc - 1
              If UCase(argv(i)) = "VIL" Or UCase(argv(i)) = "VOL" Then
                VIL_Flag = True
              Else
                TheExec.DataManager.DecomposePinList argv(i), Pin_Ary, Pin_Cnt
                For j = 0 To Pin_Cnt - 1
                    PinName = Pin_Ary(j)
                    If TheExec.DataManager.ChannelType(PinName) <> "N/C" Then
                        If Vbump_for_Interpose = True Then
                            index = InStr(LCase(PL_DC_conditions_str), PinName & "=")
                            vbump_value = Mid(LCase(PL_DC_conditions_str), index + Len(PinName) + 1, 5)
                            OutputString = OutputString & "," & PinName & "=" & vbump_value
                            
                        Else
                            OutputString = OutputString & "," & PinName & "=" & Format(TheHdw.DCVS.Pins(PinName).Voltage.Main.Value, "0.000")
                        End If
                    End If
                Next j
              End If
            Next
            PL_DC_conditions_str = ""
            OutputString = OutputString & ","
            Search_String = Mid(Search_String, 1, Len(Search_String) - 1) 'take out last ","
            Search_String = Replace(Search_String, "X@", "")
            OutputString = OutputString & Search_String
            OutputString = OutputString & ","
            OutputString = OutputString & Shmoo_Result & ","
            
            
            ''''''''****20180709  adding for printing Vcc_min/Vcman for specail case ****'''''''''''''''''''''''''''''''''''''''''''
            ''''''''Vcc_min/Vcman = -9999/9999(all fail), -5555/5555(shmoo hole), -7777/7777(alarm/error/unknown)'''''''''''''''''''''
            If Vcc_min = "N/A" And Vcc_max = "N/A" Then  ' shmoo points all fail
                Vcc_min = "-9999"
                Vcc_max = "9999"
            End If
            
            If FlagFP(2) = True Or FlagPF(2) = True Then  ' shmoo holes
                Vcc_min = "-5555"
                Vcc_max = "5555"
            End If
            
            If InStr(Shmoo_Result, "?") Then ' any unknown situations, like "alarm" or "error"
                Vcc_min = "-7777"
                Vcc_max = "7777"
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            
            If VIL_Flag = True Then
                OutputString = OutputString & Shmoo_hole & "," & Vcc_max & "," & Vcc_min & "]"
            Else
                OutputString = OutputString & Shmoo_hole & "," & Vcc_min & "," & Vcc_max & "]"
            End If
            
            TheExec.Datalog.WriteComment OutputString
            
            '' Reset to default
            OutputString = ""
            FlagPF_Count = 1
            FlagFP_Count = 1
            For i = 0 To 9
                FlagPF(i) = False
                FlagFP(i) = False
            Next i
            
        Next j_Outer
        
        '20180716 add for 2D shmoo to print force cnodition
        TheExec.Datalog.WriteComment "[Force_condition_during_shmoo:" & Charz_Force_Power_condition & "]"
        
        If Vcc_min = "N/A" Then
            Shmoo_Vcc_Min(site) = -0.1
        Else
            Shmoo_Vcc_Min(site) = Vcc_min
        End If
        
        If Vcc_max = "N/A" Then
            If RangeFrom > RangeTo Then
                Shmoo_Vcc_Max(site) = RangeFrom + 0.1
            Else
                Shmoo_Vcc_Max(site) = RangeTo + 0.1
            End If
            
        Else
            Shmoo_Vcc_Max(site) = Vcc_max
        End If
        
    Next site
    
    If Vcc_min = "N/A" Then Vcc_min = 9999
    
    '20170126 Add Limit judgement
    Dim print_all As Boolean
    Dim print_lvcc As Boolean
    Dim print_hvcc As Boolean

    Dim DFTH_Testname As String
    Dim DFTL_Testname As String
    print_all = False
    print_lvcc = False
    print_hvcc = False
    
    If InStr(instanceName, "DFTLH_") <> 0 Or InStr(instanceName, "DFTHL_") <> 0 Then print_all = True
    If InStr(instanceName, "HFLH_") <> 0 Or InStr(instanceName, "HFHL_") <> 0 Then print_all = True
    If InStr(instanceName, "MCLH_") <> 0 Or InStr(instanceName, "MCHL_") <> 0 Then print_all = True
    
    If InStr(instanceName, "DFTL_") <> 0 Then print_lvcc = True
    If InStr(instanceName, "HFL_") <> 0 Then print_lvcc = True
    If InStr(instanceName, "MCL_") <> 0 Then print_lvcc = True
    
    If InStr(instanceName, "DFTH_") <> 0 Then print_hvcc = True
    If InStr(instanceName, "HFH_") <> 0 Then print_hvcc = True
    If InStr(instanceName, "MCH_") <> 0 Then print_hvcc = True
    
    If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = True Then
    If RangeFrom < RangeTo Then
        If print_all Or print_lvcc Then
            TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instanceName & "_" & SetupName & "_Vmin"
        End If
        If print_all Or print_hvcc Then
            TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instanceName & "_" & SetupName & "_Vmax"
        End If
    Else
        If print_all Or print_lvcc Then
            DFTL_Testname = Replace(instanceName, "_CZ_NV", "_")
            DFTL_Testname = Replace(DFTL_Testname, "DFTLH", "DFTL")
            TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=DFTL_Testname & "_" & SetupName & "_Vmin"
        End If
        If print_all Or print_hvcc Then
            DFTH_Testname = Replace(instanceName, "_CZ_NV", "_")
            DFTH_Testname = Replace(DFTH_Testname, "DFTLH", "DFTH")
            TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=DFTH_Testname & "_" & SetupName & "_Vmax"
        End If
    End If
    Else '20190321 update: Suspend Datalog =False, add g_TestNum
        If RangeFrom < RangeTo Then
            If print_all Or print_lvcc Then
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instanceName & "_" & SetupName & "_Vmin", TNum:=g_TestNum
                 g_TestNum = g_TestNum + 1
           End If
            If print_all Or print_hvcc Then
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instanceName & "_" & SetupName & "_Vmax", TNum:=g_TestNum
                g_TestNum = g_TestNum + 1
            End If
        Else
            If print_all Or print_lvcc Then
                DFTL_Testname = Replace(instanceName, "_CZ_NV", "_")
                DFTL_Testname = Replace(DFTL_Testname, "DFTLH", "DFTL")
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=DFTL_Testname & "_" & SetupName & "_Vmin", TNum:=g_TestNum
                g_TestNum = g_TestNum + 1
            End If
            If print_all Or print_hvcc Then
                DFTH_Testname = Replace(instanceName, "_CZ_NV", "_")
                DFTH_Testname = Replace(DFTH_Testname, "DFTLH", "DFTH")
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=DFTH_Testname & "_" & SetupName & "_Vmax", TNum:=g_TestNum
                g_TestNum = g_TestNum + 1
            End If
        End If
    End If

 
    Exit Function
    
errHandler_shmoo:
    TheExec.ErrorLogMessage "Error in ShmooPostStep2Dto1D for " & TheExec.DataManager.instanceName
    
                If AbortTest Then Exit Function Else Resume Next
End Function
Public Function ShmooPostStep2D(argc As Long, argv() As String)
    Dim SetupName As String
    Dim i As Long
    Dim OutputString As String
    Dim instanceName As String
    Dim TestNum As Long
    Dim lvccf As Integer
    Dim LVCC As Double
    Dim site As Variant
    Dim v_Xi0 As Double
    Dim TestVoltage As String
    Dim StartVoltage As Double, EndVoltage As Double, StepSize As Double
    Dim Patt_String As String, Shmoo_Result As String
    Dim Pat As Variant
    Dim PinName As Variant
    Dim StepName As Variant
    Dim RangeFrom As Double, RangeTo As Double, RangeStepSize As Double, RangeSteps As Long
    Dim allpowerpins As String
    Dim PowerPinCnt As Long, PowerPinAry() As String
    Dim FlagFirstPass As Boolean
    Dim last_point_result As tlDevCharResult, current_point_result As tlDevCharResult
    Dim min_point As Long, max_point As Long, current_point As Long
    Dim Vcc_min As String, Vcc_max As String
    Dim patt_ary() As String, pat_count As Long, p As Variant
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim shmoo_pin_string As String
    Dim tmp As String
    Dim Search_String As String, Search_String_X As String, Search_String_Y As String
    Dim Group As Boolean
    Dim Label As String
    Dim Shmoo_Pattset As New Pattern
    Dim VIL_Flag As Boolean
    Dim step_Start As Long
    Dim step_Stop As Long
    Dim Step_x As Long
    Dim RangeLow As Double, RangeStart As Double
    Dim Shmoo_hole As String
    Dim RangeCalcType As tlDevCharRangeField
    Dim xio_spec As String
    Dim Range_temp As Double
    Dim range_plus As Long
    
    Dim HIO_PinName_Updated As Boolean      '20180515 TER
    
    Dim index As Long
    Dim vbump_value As String
    
    On Error GoTo errHandler_shmoo
    
    instanceName = TheExec.DataManager.instanceName     '20180616 TER add
    Call Get_Tname_FromFlowSheet(instanceName, HIO_PinName_Updated)      '20180515 TER add
    
    For Each site In TheExec.sites
        OutputString = ""
        lvccf = 0

        
        If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = False Then    '20180718 add
            Call TheExec.sites(site).IncrementTestNumber
        End If
        
        TestNum = TheExec.sites(site).TestNumber
        
        If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = True Then    '20180718 add
            Call TheExec.sites(site).IncrementTestNumber
        End If
        
        
        
        'v_Xi0 = TheHdw.DIB.SupportBoardClock.Frequency
        
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////  Read XI0 Nwire Setup value
'        If LCase(TheExec.DataManager.InstanceName) Like "*func*" Then
'            xio_spec = "XI0_Freq_H"
'        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*cpu*" Then
'            xio_spec = "XI0_Freq_C"
'        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*gpu*" Then
'            xio_spec = "XI0_Freq_G"
'        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*soc*" Then
'            xio_spec = "XI0_Freq_S"
'        Else
'            xio_spec = "XI0_Freq_H"
'        End If

        'xio_spec = "XI0_Freq_VAR"

'        v_Xi0 = TheExec.specs.AC(xio_spec).CurrentValue
        
        '20180716 Auto parsing FRC info
        Dim nWire_port_ary() As String
        Dim nwp As Variant
        Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
        Dim FRC_Name As String, FRC_Value As Double, All_FRC_Status As String
        All_FRC_Status = ""
        nWire_port_ary = Split(nWire_Ports_GLB, ",")
        For Each nwp In nWire_port_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            If TheHdw.Protocol.ports(port_pa).Enabled = True Then
                FRC_Name = Replace(UCase(ac_spec_pa), "_FREQ_VAR", "")
                FRC_Value = TheExec.specs.AC(ac_spec_pa).CurrentValue
                If All_FRC_Status = "" Then
                    All_FRC_Status = FRC_Name & "=" & FRC_Value
                Else
                    All_FRC_Status = All_FRC_Status & ";" & FRC_Name & "=" & FRC_Value
                End If
            End If
        Next nwp
        If FRC_Name = "" Then ' Default use XI0, if no input of FRC info
            FRC_Name = "XI0"
            FRC_Value = TheExec.specs.AC("XI0_Freq_VAR").CurrentValue
            All_FRC_Status = FRC_Name & "=" & FRC_Value
        End If
    
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
        
        ''Read X axis setup information
        SetupName = TheExec.DevChar.Setups.ActiveSetupName
        
        With TheExec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.To
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.CalculatedField

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
                range_plus = -1
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step
                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
                End If
                RangeLow = RangeTo
                range_plus = 1
            End If
        End With
        
        Patt_String = ""
        Dim patset As Variant, j As Long
        Shmoo_Pattset.Value = Shmoo_Pattern
        Patt_String = PatSetToPat(Shmoo_Pattset)
        gen_search_string SetupName, Search_String_X, tlDevCharShmooAxis_X, RangeFrom, RangeTo, RangeStepSize, RangeSteps
        
        ''Read Y axis setup information
        With TheExec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.To
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.CalculatedField

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
                range_plus = -1
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step
                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
                End If
                RangeLow = RangeTo
                range_plus = 1
            End If
        End With
        
        gen_search_string SetupName, Search_String_Y, tlDevCharShmooAxis_Y, RangeFrom, RangeTo, RangeStepSize, RangeSteps
        Search_String = Search_String_X & Search_String_Y
        
        If InStr(TheExec.DataManager.instanceName, "_NV") Then TestVoltage = "NV"
        If InStr(TheExec.DataManager.instanceName, "_HV") Then TestVoltage = "HV"
        If InStr(TheExec.DataManager.instanceName, "_LV") Then TestVoltage = "LV"

        
        

   
        OutputString = OutputString & "[V," & site & "," & All_FRC_Status & "," & HramLotId(site) & "-" & CStr(HramWaferId(site)) & "," & CStr(XCoord(site)) & "," & CStr(YCoord(site)) & ","  '20180716 Auto parsing FRC info

        OutputString = OutputString & TheExec.DataManager.instanceName & "," & SetupName & "," & CStr(TestNum) & ","

        OutputString = OutputString & Patt_String & ","
        OutputString = OutputString & TestVoltage & ","
         PinName = argv(0) 'setup voltage
        If argv(0) <> Empty Then
            TheExec.DataManager.DecomposePinList argv(0), Pin_Ary, Pin_Cnt
            PinName = argv(0) 'setup voltage
        End If
        
         
        If Vbump_for_Interpose = True Then
                Dim PL_DC_conditions_str As String
                PL_DC_conditions_str = Replace(PL_DC_conditions_GLB, ":V:", "=")
                PL_DC_conditions_str = Replace(PL_DC_conditions_str, ";", ",")
                OutputString = OutputString & PL_DC_conditions_str
            
        Else
            For j = 0 To Pin_Cnt - 1
                PinName = Pin_Ary(j)
                If TheExec.DataManager.ChannelType(PinName) <> "N/C" Then
                    If j = 0 Then
                        OutputString = OutputString & PinName & "=" & Format(TheHdw.DCVS.Pins(PinName).Voltage.Main.Value, "0.000")
                    Else
                         OutputString = OutputString & "," & PinName & "=" & Format(TheHdw.DCVS.Pins(PinName).Voltage.Main.Value, "0.000")
                    End If
                End If
            Next j
        End If
        
        For i = 1 To argc - 1
          If UCase(argv(i)) = "VIL" Or UCase(argv(i)) = "VOL" Then
            VIL_Flag = True
          Else
            TheExec.DataManager.DecomposePinList argv(i), Pin_Ary, Pin_Cnt
            For j = 0 To Pin_Cnt - 1
                PinName = Pin_Ary(j)
                If TheExec.DataManager.ChannelType(PinName) <> "N/C" Then
                    If Vbump_for_Interpose = True Then
                        index = InStr(LCase(PL_DC_conditions_str), PinName & "=")
                        vbump_value = Mid(LCase(PL_DC_conditions_str), index + Len(PinName) + 1, 5)
                        OutputString = OutputString & "," & PinName & "=" & vbump_value
                        
                    Else
                        OutputString = OutputString & "," & PinName & "=" & Format(TheHdw.DCVS.Pins(PinName).Voltage.Main.Value, "0.000")
                    End If
                End If
            Next j
          End If
        Next
        PL_DC_conditions_str = ""
        OutputString = OutputString & ","
       Search_String = Mid(Search_String, 1, Len(Search_String) - 1) 'take out last ","
        OutputString = OutputString & Search_String
        OutputString = OutputString & "]"
        TheExec.Datalog.WriteComment OutputString
    Next site

          ''clear forcecondition before exit function
          Charz_Force_Power_condition = ""

    Exit Function
    
errHandler_shmoo:
    TheExec.ErrorLogMessage "Error in ShmooPostStep2D for " & TheExec.DataManager.instanceName
    
                If AbortTest Then Exit Function Else Resume Next
End Function
''debug printing



'
'


Public Function ShmooPostStep1D(argc As Long, argv() As String)

    '
    ' Assume IO are all with NV value in the level sheet
    '
    Dim SetupName As String
    Dim i As Long
    Dim OutputString As String
    Dim instanceName As String
    Dim TestNum As Long
    Dim lvccf As Integer
    Dim LVCC As Double
    Dim site As Variant
'    Dim v_Xi0 As Double
    Dim v_Shiftin As Double
    
    Dim TestVoltage As String
    Dim StartVoltage As Double, EndVoltage As Double, StepSize As Double
    Dim Patt_String As String, Shmoo_Result As String, Shmoo_result_PF As String
    Dim Pat As Variant
    Dim PinName As Variant
    Dim StepName As Variant
    Dim RangeFrom As Double, RangeTo As Double, RangeStepSize As Double, RangeSteps As Long
    Dim allpowerpins As String
    Dim PowerPinCnt As Long, PowerPinAry() As String
    Dim Vcc_min As String, Vcc_max As String
    Dim patt_ary() As String, pat_count As Long, p As Variant
    Dim Pin_Ary() As String, p_cnt As Long, Pin_Cnt As Long
    Dim shmoo_pin_string As String
    Dim tmp As String
    Dim Search_String As String
    Dim ch As String
    Dim Group As Boolean
    Dim Label As String
    Dim step_Start As Long
    Dim step_Stop As Long
    Dim Step_x As Long
    Dim Step_NV As Long
    Dim Range_temp As Double
    Dim range_plus As Long
    Dim Shmoo_Pattset As New Pattern
'    Dim xio_spec As String
    Dim Shiftin_spec As String
    
    Dim SetupName_New As String, k As Integer
    Dim InstanceName_New As String
    Dim patset As Variant, patset1 As Variant, j As Long
    Dim RangeLow As Double, RangeStart As Double
    Dim Shmoo_hole As String
    Dim RangeCalcType As tlDevCharRangeField

    Dim RangeHigh As Double     '20180515 TER add
    Dim HIO_PinName_Updated As Boolean      '20180515 TER
    
    Dim index As Long
    Dim vbump_value As String
        

    
    On Error GoTo errHandler_shmoo
    

    ReportHVCC = True
    ReportLVCC = True
    Shmoo_hole = "NH"
    Patt_String = ""
    
    instanceName = TheExec.DataManager.instanceName     '20180616 TER add
    Call Get_Tname_FromFlowSheet(instanceName, HIO_PinName_Updated)      '20180515 TER add
    
    For Each site In TheExec.sites
        
        
        OutputString = ""
        lvccf = 0

        '20170125 Modify TestName width show in datalog
        TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
        If Len(instanceName) < 235 Then
            TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = Len(instanceName) + 20
        Else
            TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.MaximumWidth
        End If
        TheExec.Datalog.ApplySetup
        
        '' CHWu 20151026 - Add test name rule for "CPUBIST" block
''        If UCase(InstanceName) Like "*CPUBIST_*" Or UCase(InstanceName) Like "*CPUBIRA_*" Then
''            InstanceName = G_TestName
''        End If
        

        
        If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = False Then    '20180718 add
            Call TheExec.sites(site).IncrementTestNumber
        End If
        
        TestNum = TheExec.sites(site).TestNumber
        
        If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = True Then    '20180718 add
            Call TheExec.sites(site).IncrementTestNumber
        End If
        
        '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////  Read XI0 Nwire Setup value
        'xio_spec = "XI0_Freq_VAR"
         Shiftin_spec = "ShiftIn_Freq_VAR"
        
''        If LCase(TheExec.DataManager.InstanceName) Like "*tmps*" Then
''            xio_spec = "XI0_Freq_H"
''        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*adc*" Then
''            xio_spec = "XI0_Freq_H"
''        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*func*" Then
''            xio_spec = "XI0_Freq_H"
''        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*cpu*" Then
''            xio_spec = "XI0_Freq_C"
''        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*gpu*" Then
''            xio_spec = "XI0_Freq_G"
''        ElseIf LCase(TheExec.DataManager.InstanceName) Like "*soc*" Then
''            xio_spec = "XI0_Freq_S"
''        Else
''            xio_spec = "XI0_Freq_H"
''        End If
        
'        v_Xi0 = TheExec.specs.AC(xio_spec).CurrentValue
        v_Shiftin = TheExec.specs.AC(Shiftin_spec).CurrentValue
        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

         '20180716 Auto parsing FRC info
        Dim nWire_port_ary() As String
        Dim nwp As Variant
        Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
        Dim FRC_Name As String, FRC_Value As Double, All_FRC_Status As String
        All_FRC_Status = ""
        nWire_port_ary = Split(nWire_Ports_GLB, ",")
        For Each nwp In nWire_port_ary
            Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
            'If UCase(ac_spec_pa) Like "*XI0*" Or UCase(ac_spec_pa) Like "*XO0*" Then
            If TheHdw.Protocol.ports(port_pa).Enabled = True Then
                FRC_Name = Replace(UCase(ac_spec_pa), "_FREQ_VAR", "")
                FRC_Value = TheExec.specs.AC(ac_spec_pa).CurrentValue
                If All_FRC_Status = "" Then
                    All_FRC_Status = FRC_Name & "=" & FRC_Value
                Else
                    All_FRC_Status = All_FRC_Status & ";" & FRC_Name & "=" & FRC_Value
                End If
            End If
            'End If
        Next nwp
        If FRC_Name = "" Then ' Default use XI0, if no input of FRC info
            FRC_Name = "XI0"
            FRC_Value = TheExec.specs.AC("XI0_Freq_VAR").CurrentValue
            All_FRC_Status = FRC_Name & "=" & FRC_Value
        End If
        
        SetupName = TheExec.DevChar.Setups.ActiveSetupName
        
        Shmoo_Pattset.Value = Shmoo_Pattern
        Patt_String = PatSetToPat(Shmoo_Pattset)
        
        Search_String = ""
        With TheExec.DevChar
            StepName = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
            RangeFrom = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
            RangeTo = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.To
            RangeSteps = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.Steps + 1
            RangeStepSize = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.StepSize
            RangeCalcType = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.CalculatedField
            
            
            '20170210 Added to check Endpoint
            Dim RangeTo_New As Double
            Dim RangeFrom_New As Double
            

            If RangeFrom < RangeTo Then ' always start from lower Value
                step_Start = 0
                step_Stop = RangeSteps - 1
                Step_x = 1
                RangeLow = RangeFrom
'                range_plus = -1
            '20170210 Added to check Endpoint
                RangeFrom_New = RangeFrom
                RangeTo_New = RangeFrom + (RangeSteps - 1) * RangeStepSize
            Else
                step_Start = RangeSteps - 1
                step_Stop = 0
                Step_x = -1
'                If RangeCalcType = tlDevCharRangeField_Steps Then 'calculate step
'                    RangeTo = RangeFrom - (RangeSteps - 1) * RangeStepSize
'                End If

'                range_plus = 1
                '20170210 Added to check Endpoint
                RangeLow = Format((RangeFrom - (RangeSteps - 1) * RangeStepSize), "0.000#########")
                
                RangeFrom_New = RangeFrom
                RangeTo_New = RangeFrom - (RangeSteps - 1) * RangeStepSize
            End If
            If Abs(RangeTo) < 0.000000000001 Then RangeTo = 0
            If Abs(RangeFrom) < 0.000000000001 Then RangeFrom = 0
            '20170210 Added to check Endpoint
            gen_search_string SetupName, Search_String, tlDevCharShmooAxis_X, RangeFrom_New, RangeTo_New, RangeStepSize, RangeSteps
            
            shmoo_pin_string = .Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins
            If TheExec.EnableWord("ShmooMakePseudoData") = True Then Call ShmooMakePseudoData(SetupName, step_Start, step_Stop, Step_x)
           Call CreateShmooResultString(Shmoo_Result, Shmoo_result_PF, SetupName, step_Start, step_Stop, Step_x, site)
            
'20161229 Roy Modified,Prevent Step_NV out of range
            Step_NV = -1
'            If Not (InStr(LCase(TheExec.DevChar.ActiveDataObject.TestName), "vih") > 0 Or InStr(LCase(TheExec.DevChar.ActiveDataObject.TestName), "vil") > 0) Then
'                Call Decide_NV(Step_NV, RangeLow, RangeStepSize, Step_start, Step_x, SetupName)
'                If Step_NV > Len(Shmoo_result_PF) Or Step_NV < 0 Then Step_NV = -1
'            End If
            'Exit Function
            Call Decide_LVCC_HVCC(Vcc_min, Vcc_max, Shmoo_hole, Step_NV, RangeLow, RangeStepSize, Shmoo_result_PF, SetupName, step_Start, step_Stop, Step_x)
            
            
        End With
        If InStr(TheExec.DataManager.instanceName, "_NV") Then TestVoltage = "NV"
        If InStr(TheExec.DataManager.instanceName, "_HV") Then TestVoltage = "HV"
        If InStr(TheExec.DataManager.instanceName, "_LV") Then TestVoltage = "LV"
        
'    [Char,N99G19-1,16,7,V,0,XI0=24000000,CpuBira_P0001_IN02_BIR_SI_PL00_CL16_BIR_59N_SI_PP_NV,CPU_BIST_CPU_Domain_CPU_SRAM_Domain_P1_Full_Range,1069,
'.\pattern\CpuMbist\PP_FIJA0_C_IN00_XX_CLXX_XXX_XXX_XXX_P0001_1308131609_SI_mod.pat,.\pattern\CpuMbist\PP_FIJA0_C_IN02_BI_CLXX_BIR_JTG_XXX_ALLFV_1306250000_SI.pat,.\pattern\CpuMbist\PP_FIJA0_C_PL00_BI_CL16_BIR_JTG_59N_ALLFV_1306250000_SI.pat,
'NV,VDD_FIXED=0.528:1.404:0.005,VDD_VAR_SOC_VAR=0.500:1.330:0.005,
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++,NH,0.5,1.260]
        If Shmoo_header = "" Then Shmoo_header = "Char"
        OutputString = OutputString & "[" & Shmoo_header & "," & HramLotId(site) & "-" & CStr(HramWaferId(site)) & "," & CStr(XCoord(site)) & "," & CStr(YCoord(site))
        'OutputString = OutputString & "[" & Shmoo_header & "," & HramLotId(Site) & "," & CStr(Xcoord(Site)) & "," & CStr(Ycoord(Site))
        SetupName_New = SetupName
        
        'Shmoo_header
        Dim VIL_Flag As Boolean
        VIL_Flag = False
        ShmooPowerName = ShmooPowerName

        
       
        OutputString = OutputString & ",V," & site & "," & All_FRC_Status & ","    '20180716 Auto parsing FRC info
        OutputString = OutputString & TheExec.DataManager.instanceName & ShmooPowerName & "," & SetupName_New & "," & CStr(TestNum) & ","


        OutputString = OutputString & Patt_String & ","
        OutputString = OutputString & TestVoltage & ","

        If argv(0) <> Empty Then
            TheExec.DataManager.DecomposePinList argv(0), Pin_Ary, Pin_Cnt
            PinName = argv(0) 'setup voltage
        End If
        
        
        If Vbump_for_Interpose = True Then
            Dim PL_DC_conditions_str As String
            PL_DC_conditions_str = Replace(PL_DC_conditions_GLB, ":V:", "=")
            PL_DC_conditions_str = Replace(PL_DC_conditions_str, ";", ",")
            OutputString = OutputString & PL_DC_conditions_str
        
        Else
            For j = 0 To Pin_Cnt - 1
                PinName = Pin_Ary(j)
                If TheExec.DataManager.ChannelType(PinName) <> "N/C" Then
                    If j = 0 Then
                        OutputString = OutputString & PinName & "=" & Format(TheHdw.DCVS.Pins(PinName).Voltage.Main.Value, "0.000")
                    Else
                        OutputString = OutputString & "," & PinName & "=" & Format(TheHdw.DCVS.Pins(PinName).Voltage.Main.Value, "0.000")
                    End If
                End If
            Next j
        End If
        
        For i = 1 To argc - 1
          If UCase(argv(i)) = "VIL" Or UCase(argv(i)) = "VOL" Then
            VIL_Flag = True
          Else
            TheExec.DataManager.DecomposePinList argv(i), Pin_Ary, Pin_Cnt
            
            
            For j = 0 To Pin_Cnt - 1
                PinName = LCase(Pin_Ary(j))
                If TheExec.DataManager.ChannelType(PinName) <> "N/C" Then
                    If Vbump_for_Interpose = True Then
                        index = InStr(LCase(PL_DC_conditions_str), PinName & "=")
                        vbump_value = Mid(LCase(PL_DC_conditions_str), index + Len(PinName) + 1, 5)
                        OutputString = OutputString & "," & PinName & "=" & vbump_value
                    Else
                        OutputString = OutputString & "," & PinName & "=" & Format(TheHdw.DCVS.Pins(PinName).Voltage.Main.Value, "0.000")
                    End If
                End If
            Next j
          End If
        Next
        PL_DC_conditions_str = ""
        
        OutputString = OutputString & ","
        Search_String = Mid(Search_String, 1, Len(Search_String) - 1) 'take out last ","
        Search_String = Replace(Search_String, "X@", "")
        OutputString = OutputString & Search_String
        OutputString = OutputString & ","
        
'///////////////////////////////////////////////////////// check hole
        If Vcc_max = "5555" And Vcc_min <> "-5555" Then Shmoo_hole = "HH"
        If Vcc_max <> "5555" And Vcc_min = "-5555" Then Shmoo_hole = "LH"
        If Vcc_max = "5555" And Vcc_min = "-5555" Then Shmoo_hole = "BH"
        If Vcc_max <> "5555" And Vcc_min <> "-5555" Then Shmoo_hole = "NH"
'/////////////////////////////////////////////////////////
        OutputString = OutputString & Shmoo_Result & ","
        
        If VIL_Flag = True Then
            OutputString = OutputString & Shmoo_hole & "," & Vcc_max & "," & Vcc_min & "]"
        Else
            OutputString = OutputString & Shmoo_hole & "," & Vcc_min & "," & Vcc_max & "]"
        End If
        
        ''get current Timing set sheet''
        Dim Context As String: Context = ""
        Dim TimeSet_Str As String: TimeSet_Str = ""
        Context = TheExec.Contexts.ActiveSelection
        TimeSet_Str = TheExec.Contexts(Context).Sheets.Timesets
        
'        Debug.Print outputString
        TheExec.Datalog.WriteComment OutputString
        TheExec.Datalog.WriteComment "[Force_condition_during_shmoo:" & Charz_Force_Power_condition & "]"
        TheExec.Datalog.WriteComment "[Activity_Timing_Sheet:" & UCase(TimeSet_Str) & "," & "Shiftin_Freq=" & CStr(v_Shiftin) & "]"
        
        If Vcc_min = "N/A" Then
            Shmoo_Vcc_Min(site) = -0.1
        Else
            If Vcc_min = "" Then Vcc_min = 0
            Shmoo_Vcc_Min(site) = Vcc_min
        End If
        
        If Vcc_max = "N/A" Then
            If RangeFrom > RangeTo Then
                Shmoo_Vcc_Max(site) = RangeFrom + 0.1
            Else
                Shmoo_Vcc_Max(site) = RangeTo + 0.1
            End If
        Else
            If Vcc_max = "" Then Vcc_max = 0
            Shmoo_Vcc_Max(site) = Vcc_max
        End If
        
        '**************************************************AI**********************************************
        
        If TheExec.EnableWord("AI_Fail_Log") = True And Voltage_fail_point <> 0 Then
        Dim Setpower As String
        Dim X As Integer
        Dim y As Integer
        'Voltage_fail_point
        'Voltage_fail_collect
        Setpower = ""
        For X = 0 To Voltage_fail_point - 1
                Setpower = ""
                Setpower = Replace(shmoo_pin_string, ",", "+") & ",VDD," & Voltage_fail_collect(X)
                TheExec.Datalog.WriteComment Setpower
                Call SetForceCondition(Setpower)
                TheHdw.Patterns(Patt_String).Test pfAlways, 0
                TheHdw.Digital.Patgen.HaltWait
                    If TheHdw.Digital.Patgen.PatternBurstPassed(site) = False Then y = y + 1
                    If y = Voltage_fail_point_request Then GoTo contine1
        Next X
contine1:
    End If
        '*************************************************AI**********************************************
    
    Next site
    
        
        If Vcc_min = "N/A" Then Vcc_min = 9999
        'job char flag
    If UCase(currentJobName) Like "*CP*" Or UCase(currentJobName) Like "*FT*" Then

        Dim TestNameLVCC As String, TestNameHVCC As String
        Dim testName As String
        Dim GPIO_Char_Shmoo_Pin As String
        Dim Shmoo_setup_name As String
        Shmoo_setup_name = TheExec.DevChar.Setups.ActiveSetupName
        GPIO_Char_Shmoo_Pin = TheExec.DevChar.Setups(Shmoo_setup_name).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins

''        If UCase(InstanceName) Like "*_VDIFF*" Or UCase(InstanceName) Like "*_VCM*" Then
''
''            If UCase(InstanceName) Like "*_NV*" Then
''                testName = Replace(InstanceName, "_CZ_NV", "")
''            ElseIf UCase(InstanceName) Like "*_LV*" Then
''                testName = Replace(InstanceName, "_CZ_LV", "")
''            ElseIf UCase(InstanceName) Like "*_HV*" Then
''                testName = Replace(InstanceName, "_CZ_HV", "")
''            End If
''
''        Else
        
            If UCase(instanceName) Like "*_CHAR_CP*" Then
                testName = Replace(instanceName, "_Char_CP", "")
            ElseIf UCase(instanceName) Like "*__H*" Then
                testName = Replace(instanceName, "__H", "")
            ElseIf UCase(instanceName) Like "*__L*" Then
                testName = Replace(instanceName, "__L", "")
            Else
                testName = instanceName
            End If
            
            If UCase(testName) Like "*_NV*" Then
                testName = Replace(testName, "_CZ_NV", "_")
            ElseIf UCase(testName) Like "*_LV*" Then
                testName = Replace(testName, "_CZ_LV", "_")
            ElseIf UCase(testName) Like "*_HV*" Then
                testName = Replace(testName, "_CZ_HV", "_")
            End If
            
            Dim HVCC_DFTLH As String
            Dim LVCC_DFTLH As String
         '20160925 Multi_USL/LSL
            Dim HVCC_MCLH As String
            Dim LVCC_MCLH As String
         If UCase(testName) Like "*DFTLH*" Then
                HVCC_DFTLH = Replace(testName, "DFTLH", "DFTH")
                LVCC_DFTLH = Replace(testName, "DFTLH", "DFTL")
         End If
         '20160925 Multi_USL/LSL
         If UCase(testName) Like "*MCLH*" Then
                HVCC_MCLH = Replace(testName, "MCLH", "MCH")
                LVCC_MCLH = Replace(testName, "MCLH", "MCL")
         End If
         
            
'        End If


        TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
        TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = True
        TheExec.Datalog.ApplySetup
'--------------------------------------------------------------------------------------------
'        Public CHAR_USL_HVCC As Double
'        Public CHAR_USL_LVCC As Double
'        Public CHAR_LSL_HVCC As Double
'        Public CHAR_LSL_LVCC As Double
        
        
        
        Dim HF_HVCC_TESTNAME, HF_LVCC_TESTNAME As String
        Dim hi_limit, Low_limit As Double
        
        If RangeFrom < RangeTo Then
            hi_limit = RangeTo: Low_limit = RangeFrom
        Else
            hi_limit = RangeFrom: Low_limit = RangeTo
        End If
        
        If (CHAR_USL_HVCC = 9999) Then CHAR_USL_HVCC = hi_limit
        If (CHAR_LSL_HVCC = 9999) Then CHAR_LSL_HVCC = Low_limit
        If (CHAR_USL_LVCC = 9999) Then CHAR_USL_LVCC = hi_limit
        If (CHAR_LSL_LVCC = 9999) Then CHAR_LSL_LVCC = Low_limit
        
        'Debug.Print TheExec.DataManager.InstanceName & "," & Interpose_PrePat_GLB
        

        If (CHAR_USL_HVCC < CHAR_LSL_HVCC) Then TheExec.AddOutput TheExec.DataManager.instanceName & " : Limit Error ! " & "HVCC_USL=" & CStr(CHAR_USL_HVCC) & ",HVCC_LSL=" & CStr(CHAR_LSL_HVCC): CHAR_USL_HVCC = hi_limit: CHAR_LSL_HVCC = Low_limit
        If (CHAR_USL_LVCC < CHAR_LSL_LVCC) Then TheExec.AddOutput TheExec.DataManager.instanceName & " : Limit Error ! " & "LVCC_USL=" & CStr(CHAR_USL_LVCC) & ",LVCC_LSL=" & CStr(CHAR_LSL_LVCC): CHAR_USL_LVCC = hi_limit: CHAR_LSL_LVCC = Low_limit

        If UCase(instanceName) Like "HFHL*" Or UCase(instanceName) Like "HFLH*" Then
        
            HF_HVCC_TESTNAME = Replace(testName, "HFHL", "HFH")
            HF_HVCC_TESTNAME = Replace(testName, "HFLH", "HFH")
            
            HF_LVCC_TESTNAME = Replace(testName, "HFHL", "HFL")
            HF_LVCC_TESTNAME = Replace(testName, "HFLH", "HFL")
            
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=HF_HVCC_TESTNAME & " " & shmoo_pin_string & " <> " & HF_HVCC_TESTNAME, lowVal:=Low_limit, hiVal:=hi_limit
            '20170120 chnage print format
            TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=HF_LVCC_TESTNAME & " " & shmoo_pin_string & " <> " & HF_HVCC_TESTNAME, lowVal:=Low_limit, hiVal:=hi_limit
            TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
            
            TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            TheExec.Datalog.ApplySetup
            Exit Function
        End If
        
        If UCase(instanceName) Like "HFH*" Then
        
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC
            TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
            
            TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            TheExec.Datalog.ApplySetup
            Exit Function
        End If
        
        If UCase(instanceName) Like "HFL*" Then
        
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC
            TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
            
            TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            TheExec.Datalog.ApplySetup
            Exit Function
        End If
        
        If UCase(instanceName) Like "HIO*" And UCase(instanceName) Like "*VCM*" And UCase(instanceName) Like "*USBPICO*" Then
        
            HF_HVCC_TESTNAME = testName
            
            
'            HF_LVCC_TESTNAME = testName
            
            
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=HF_HVCC_TESTNAME & "   " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC
        
'            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=HF_LVCC_TESTNAME & "   " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC

            TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            TheExec.Datalog.ApplySetup
            Exit Function
        End If
        
        
        
        If UCase(instanceName) Like "*DIFF*" And UCase(instanceName) Like "HIO*" Then
            If (ReportHVCC And ReportLVCC) Then
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=testName & "_H_" & " " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC
                'TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=testName & "_L_" & " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC"
            ElseIf (ReportLVCC) Then
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC
            ElseIf (ReportHVCC) Then
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC
            End If

            TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            TheExec.Datalog.ApplySetup
            Exit Function
        End If
        
        If UCase(instanceName) Like "*VID*" Or UCase(instanceName) Like "*VICM*" Then
            'Cyprus USB2 20170823
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string & " <> " & testName, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC
            TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
            TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
            TheExec.Datalog.ApplySetup
            Exit Function
        End If
        
        If instanceName Like "*HAC*" And SetupName Like "*VIL*" Then
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & SetupName & "_Vmax")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string '& " <> " & TestName
                TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        ElseIf instanceName Like "*HAC*" And SetupName Like "*VIH*" Then
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & SetupName & "_Vmin")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string '& " <> " & TestName
                TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        
        ElseIf instanceName Like "DFTLH*" Then
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=HVCC_DFTLH & " " & shmoo_pin_string ' & " <> " & HVCC_DFTLH
                TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=LVCC_DFTLH & " " & shmoo_pin_string ' & " <> " & LVCC_DFTLH
                TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        
        ElseIf instanceName Like "DFTL_*" Or instanceName Like "MCL_*" Then
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string ' & " <> " & TestName
                TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        ElseIf instanceName Like "DFTH_*" Or instanceName Like "MCH_*" Then
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string ' & " <> " & TestName
                TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        
        
        
        '20160925 Multi_USL/LSL
        ElseIf instanceName Like "MCLH*" Then
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=HVCC_MCLH & " " & shmoo_pin_string '& " <> " & HVCC_MCLH
                TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
                TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
                TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=LVCC_MCLH & " " & shmoo_pin_string '& " <> " & LVCC_MCLH
                TheExec.Datalog.WriteComment "[Force_condition_during_shmoo_HW:" & ReadHWPowerValue_GLB & "]"
        End If
        
        If UCase(TheExec.DataManager.instanceName) Like "*ALLPINSGPIO1_X_VI*" Then
            
            If instanceName Like "*HIO*" And instanceName Like "*_VIL_*" Then
                    TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
                    TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=testName & " " & GPIO_Char_Shmoo_Pin ' & " <> " & TestName
            ElseIf instanceName Like "*HIO*" And instanceName Like "*_VIH_*" Then
                    TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
                    TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, PinName:=GPIO_Char_Shmoo_Pin, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=testName & " " & GPIO_Char_Shmoo_Pin '& " <> " & TestName
            End If
            
        ElseIf UCase(TheExec.DataManager.instanceName) Like "*_AMP_*" Then
            
            If instanceName Like "*HIO_VIL*" Then
                    TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
                    TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string '& " <> " & TestName
            ElseIf instanceName Like "*HIO_VIH*" Then
                    TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
                    TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string '& " <> " & TestName
            End If
                        
        Else
            If instanceName Like "*HIO*" And instanceName Like "*VIL_*" Then
                    TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
                    TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=CHAR_LSL_HVCC, hiVal:=CHAR_USL_HVCC, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string '& " <> " & TestName
            ElseIf instanceName Like "*HIO*" And instanceName Like "*VIH_*" Then
                    TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
                    TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=CHAR_LSL_LVCC, hiVal:=CHAR_USL_LVCC, ForceResults:=tlForceNone, Tname:=testName & " " & shmoo_pin_string ' & " <> " & TestName
            End If
        End If
        
'--------------------------------------------------------------------------------------------
    
        
    Else
        If RangeFrom < RangeTo Then
            TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instanceName & "_" & SetupName & "_Vmin"
            TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeFrom, hiVal:=RangeTo, ForceResults:=tlForceNone, Tname:=instanceName & "_" & SetupName & "_Vmax"
        Else
            TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmin")
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Min, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=instanceName & "_" & SetupName & "_Vmin"
            TheExec.Datalog.WriteComment ("Test name :" & instanceName & "_" & SetupName & "_Vmax")
            TheExec.Flow.TestLimit resultVal:=Shmoo_Vcc_Max, lowVal:=RangeTo, hiVal:=RangeFrom, ForceResults:=tlForceNone, Tname:=instanceName & "_" & SetupName & "_Vmax"
        End If
    End If
    
 '''''-------------  CHWUD 11/2 for print LVCC get 3 fail log -----------------------------------------------
    
For Each site In TheExec.sites


         If TheExec.Flow.EnableWord("CaptureFaillog") = True Then
         
                If Shmoo_hole = "BH" Or Shmoo_hole = "LH" Then
                
                    If (LCase(TheExec.CurrentJob) Like "*cp*") Then
                        FailingBoundaryDatalog_Func_Multi_Power Search_String, LotID, CStr(WaferID), CStr(XCoord(site)), _
                        CStr(YCoord(site)), Shmoo_Pattern, "Shmoo hole", High_to_Low, site
                         
                    Else
                        ''''0605 update to use HRAM data
                        FailingBoundaryDatalog_Func_Multi_Power Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                        CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo hole", High_to_Low, site
                    
                    End If

                End If
        End If
        If TheExec.Flow.EnableWord("Debug_LVCC") = True Then
            If (LCase(TheExec.CurrentJob) Like "*cp*") Then
                FailingBoundaryDatalog_Func_Multi_Power Search_String, LotID, CStr(WaferID), CStr(XCoord(site)), _
                CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site
                
            Else
                ''''0605 update to use HRAM data
                FailingBoundaryDatalog_Func_Multi_Power Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site
            End If
        End If
        
        If TheExec.Flow.EnableWord("Debug_HVCC") = True Then
            If (LCase(TheExec.CurrentJob) Like "*cp*") Then
                FailingBoundaryDatalog_Func_Multi_Power Search_String, LotID, CStr(WaferID), CStr(XCoord(site)), _
                CStr(YCoord(site)), Shmoo_Pattern, "Shmoo HVCC", Low_to_High, site
            Else
                ''''0605 update to use HRAM data
                FailingBoundaryDatalog_Func_Multi_Power Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo HVCC", Low_to_High, site
            End If
        End If

         If TheExec.Flow.EnableWord("Debug_LVCC_VminBoundary") = True Then
          If Shmoo_Vcc_Min(site) > 0 Then
             If (LCase(TheExec.CurrentJob) Like "*cp*") Then
                If g_VminBoundary_selsrm = True Then
                    FailingDatalog_HLvcc_Boundary_SELSRM Search_String, LotID, CStr(WaferID), CStr(XCoord(site)), _
                    CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                Else
                    FailingDatalog_Lvcc_Boundary Search_String, LotID, CStr(WaferID), CStr(XCoord(site)), _
                    CStr(YCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                End If
                
             Else
                If g_VminBoundary_selsrm = True Then
                    FailingDatalog_HLvcc_Boundary_SELSRM Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                    CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                Else
                    FailingDatalog_Lvcc_Boundary Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                    CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo LVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
                End If
             End If
          Else
          End If
        End If
        
        If TheExec.Flow.EnableWord("Debug_HVCC_VminBoundary") = True Then
        Dim Vcc_max_Limit As Double
        If RangeFrom > RangeTo Then
            Vcc_max_Limit = RangeFrom
        Else
            Vcc_max_Limit = RangeTo
        End If
          If Shmoo_Vcc_Max(site) <= Vcc_max_Limit Then
             If (LCase(TheExec.CurrentJob) Like "*cp*") Then
               If g_VminBoundary_selsrm = True Then
                   FailingDatalog_HLvcc_Boundary_SELSRM Search_String, LotID, CStr(WaferID), CStr(XCoord(site)), _
                   CStr(YCoord(site)), Shmoo_Pattern, "Shmoo HVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
               Else
                   FailingDatalog_Hvcc_Boundary Search_String, LotID, CStr(WaferID), CStr(XCoord(site)), _
                   CStr(YCoord(site)), Shmoo_Pattern, "Shmoo HVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
               End If
             Else
               If g_VminBoundary_selsrm = True Then
                   FailingDatalog_HLvcc_Boundary_SELSRM Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                   CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo HVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
               Else
                   FailingDatalog_Hvcc_Boundary Search_String, HramLotId(site), CStr(HramWaferId(site)), CStr(HramXCoord(site)), _
                   CStr(HramYCoord(site)), Shmoo_Pattern, "Shmoo HVCC", High_to_Low, site, RangeFrom, RangeTo, RangeSteps, RangeStepSize
               End If
             End If
          End If
        End If
        
Next site
        
'-----------------------------------------------------------------------

    
    TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
    TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
    '20170125 Modify TestName width show in datalog
    TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width = 60
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 70
    TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = 75
    TheExec.Datalog.ApplySetup
    '20170126 Initialize GLlobal power condition
    ReadHWPowerValue_GLB = ""
    Charz_Force_Power_condition = ""
    Exit Function
    
errHandler_shmoo:

    TheExec.ErrorLogMessage "Error in ShmooPostStep1D for " & TheExec.DataManager.instanceName
    TheExec.Datalog.WriteComment "<Error> " + TheExec.DataManager.instanceName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function



 

Public Function ShmooResultPF_1(ShmRes As SiteVariant, LVCC As SiteDouble, HVCC As SiteDouble, ShmooEachPoint() As SiteDouble)
'Public Function ShmooResultPF(ShmRes As String, LVCC As Double, HVCC As Double)

Dim strlen As Long
Dim i, j, k As Long
Dim Count As Long
Dim LocResult As String
Dim PreviousResult As String
Dim PreviousResult_tmp As String
Dim FPPoint(1000) As Integer
Dim FPCount As Integer
Dim PFPoint(1000) As Integer
Dim PFCount As Integer
Dim FirstPassLoc As Integer
Dim FirstFailLoc As Integer
Dim FPTransit As Integer
Dim PFTransit As Integer
'Dim LVCC As Double
Dim FPTransit_1 As Integer
Dim PFTransit_1 As Integer
Dim Entry_Flag As Boolean
Dim FailToPassCounr As Integer


    PFCount = 0
    FPCount = 0
    FirstPassLoc = -1
    FirstFailLoc = -1
    FPTransit = -1
    PFTransit = -1
    FPTransit_1 = -1
    PFTransit_1 = -1
    Count = 0
    Entry_Flag = True
    FailToPassCounr = 0
    k = InStr(ShmRes, "P") ' for FFPPFF case


''Dim ShmooEachPoint(5) As Double
''ShmooEachPoint(0) = 1
''ShmooEachPoint(1) = 1.2
''ShmooEachPoint(2) = 1.4
''ShmooEachPoint(3) = 1.6
''ShmooEachPoint(4) = 1.8
''ShmooEachPoint(5) = 2

''ShmooEachPoint(0) = 2
''ShmooEachPoint(1) = 1.8
''ShmooEachPoint(2) = 1.6
''ShmooEachPoint(3) = 1.4
''ShmooEachPoint(4) = 1.2
''ShmooEachPoint(5) = 1

If InStr(ShmRes, "A") Then 'Alarm
    ShmRes = "7777"
ElseIf InStr(ShmRes, "E") Then 'Error
    ShmRes = "7777"
Else
''    If LVCC_flag = True Then
        strlen = Len(ShmRes)
        For i = 1 To strlen
            LocResult = Mid(ShmRes, i, 1)
            If (i = 1) Then
                PreviousResult = LocResult
                PreviousResult_tmp = LocResult
            Else
                If (PreviousResult <> LocResult) Then
                    If (PreviousResult = "P") Then
                        PFPoint(PFCount) = i
                        PFCount = PFCount + 1
                    Else
                        FPPoint(FPCount) = i
                        FPCount = FPCount + 1
                    End If
                    PreviousResult = LocResult
                End If
            End If
            
            If (LocResult = "P") Then
                If (FirstPassLoc = -1) Then
                    FirstPassLoc = i - 1
                End If
                If (FirstFailLoc <> -1 And PFTransit = -1) Then
                    PFTransit = i - 1
                End If
                ''===============================================
                '' only looking for this scenario "FFPPFF"
'                PreviousResult_tmp = LocResult
'                If FirstFailLoc <> -1 And Entry_Flag = True Then
''                If FirstFailLoc <> -1 And (PreviousResult_tmp <> LocResult) Then
'''                    FPTransit_1 = FPTransit_1 + 1
'''                    Entry_Flag = False
''                    FailToPassCounr = FailToPassCounr + 1 ' if FailToPassCounr > 1 means the result has shmoo hole
''                End If
'                PreviousResult_tmp = LocResult
                ''===============================================
            End If
            
            If (LocResult = "F") Then
                If (FirstFailLoc = -1) Then
                    FirstFailLoc = i - 1
                End If
                If (FirstPassLoc <> -1 And FPTransit = -1) Then
                    FPTransit = i - 1
                End If
                ''===============================================
                '' only looking for this scenario "FFPPFF"
'                If FirstPassLoc <> -1 And (PreviousResult_tmp <> LocResult) Then
'                    FailToPassCounr = FailToPassCounr + 1 ' if FailToPassCounr > 1 means the result has shmoo hole
'                End If
'                If FirstPassLoc <> -1 And FailToPassCounr = 0 Then
'                    PFTransit_1 = PFTransit_1 + 1
'                    Entry_Flag = True
'                End If
'                PreviousResult_tmp = LocResult

                ''===============================================
                
           End If
           
            If Left(ShmRes, 1) = "F" Then ' only allow first point is "F"
                If Mid(ShmRes, k + 1, 1) = "P" Then ' skip  FFFFF 'P' FFFFF case
                    If FirstPassLoc <> -1 And (PreviousResult_tmp <> LocResult) Then
                        PreviousResult_tmp = LocResult
                        FailToPassCounr = FailToPassCounr + 1 '
                    End If
                End If
            End If
        Next i
    
        If (PFTransit = -1 And FPTransit = -1 And FirstPassLoc <> -1 And FirstFailLoc = -1) Then ' All PASS
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                LVCC = ShmooEachPoint(LBound(ShmooEachPoint))
                HVCC = ShmooEachPoint(UBound(ShmooEachPoint))
            Else
                LVCC = ShmooEachPoint(UBound(ShmooEachPoint))
                HVCC = ShmooEachPoint(LBound(ShmooEachPoint))
            End If
        ElseIf (PFTransit = -1 And FPTransit = -1 And FirstPassLoc = -1 And FirstFailLoc <> -1) Then ' All FAIL
'            LVCC = -9999
'            HVCC = 9999
'            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                ShmRes = "9999"
'            Else
'                ShmRes = "-9999"
'            End If
        ElseIf (PFTransit <> -1 And FPTransit = -1 And FirstPassLoc <> -1 And FirstFailLoc <> -1) Then ' Fail-Pass transition point
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                LVCC = ShmooEachPoint(FPPoint(0) - 1)
                HVCC = ShmooEachPoint(UBound(ShmooEachPoint))
            Else
                LVCC = ShmooEachPoint(UBound(ShmooEachPoint))
                HVCC = ShmooEachPoint(FPPoint(0) - 1)
            End If
        ElseIf (PFTransit = -1 And FPTransit <> -1 And FirstPassLoc <> -1 And FirstFailLoc <> -1) Then ' Pass-Fail transition point
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                LVCC = ShmooEachPoint(0)
                HVCC = ShmooEachPoint(PFPoint(0) - 2)
            Else
                LVCC = ShmooEachPoint(PFPoint(0) - 2)
                HVCC = ShmooEachPoint(0)
            End If
        ElseIf (PFTransit <> -1 And FPTransit <> -1 And FirstPassLoc <> -1 And FirstFailLoc <> -1 And FailToPassCounr <> 2) Then ' Shmoo hole, ex: PPPFFFPPP
'            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                ShmRes = "5555"
'            Else
'                ShmRes = "-5555"
'            End If
        ElseIf (FailToPassCounr = 2) Then '  ex: FFPPPFFF
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                LVCC = ShmooEachPoint(FPPoint(0) - 1)
                HVCC = ShmooEachPoint(PFPoint(0) - 2)
            Else
                LVCC = ShmooEachPoint(PFPoint(0) - 2)
                HVCC = ShmooEachPoint(FPPoint(0) - 1)
            End If
        End If
    
End If


End Function



'Public Function StoreEachPoint_PreSetup(axis_val() As SiteVariant, axis_pin() As String)
Public Function StoreEachPoint_PreSetup()
'        Dim axis_val() As SiteVariant  'XYZ
        Dim Suspend_Flag As Boolean
        Dim DevSetupName As String
'        Dim axis_pin() As String 'XYZ
        Dim TnameCombShmooInfo As String
        Dim curr_axis As Variant
        Dim Tracking_Item As Variant
        Dim Tracking_Xaxis_Pin() As String
        Dim Tracking_Xaxis_val() As New SiteDouble
        Dim Tracking_Yaxis_Pin() As String
        Dim Tracking_Yaxis_val() As New SiteDouble
        Dim Tracking_Zaxis_Pin() As String
        Dim Tracking_Zaxis_val() As New SiteDouble
        Dim ii As Integer
        Dim jj As Integer
        Dim kk As Integer
'  Exit Function
        DevSetupName = TheExec.DevChar.Setups.ActiveSetupName
        Suspend_Flag = TheExec.DevChar.Setups.Item(DevSetupName).output.SuspendDatalog
        ii = 0
        jj = 0
        kk = 0
        ReDim axis_val(TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count - 1)
        ReDim axis_pin(TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.Count - 1)
        If X_Tracking_Point <> 0 Then
            ReDim Tracking_Xaxis_Pin(X_Tracking_Point - 1)
            ReDim Tracking_Xaxis_val(X_Tracking_Point - 1)
        End If
        If Y_Tracking_Point <> 0 Then
            ReDim Tracking_Yaxis_Pin(Y_Tracking_Point - 1)
            ReDim Tracking_Yaxis_val(Y_Tracking_Point - 1)
        End If
        If Z_Tracking_Point <> 0 Then
            ReDim Tracking_Zaxis_Pin(Z_Tracking_Point - 1)
            ReDim Tracking_Zaxis_val(Z_Tracking_Point - 1)
        End If
        
'        If Suspend_Flag = False And theexec.DevChar.Setups(DevSetupName).Shmoo.Axes.count = 3 Then
        If Suspend_Flag = False Then
            For Each site In TheExec.sites
                For Each curr_axis In TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.List
                    Select Case curr_axis
                        Case 0 'X Axis
                            axis_val(curr_axis)(site) = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).Value ''20190319 update
                            Call ValueResolution(axis_val, curr_axis)
'                            Exit Function
                            If X_Tracking_Point <> 0 Then ' Tracking case
                                With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                                    For Each Tracking_Item In .List
                                        Tracking_Xaxis_val(ii) = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(Tracking_Item).Value
                                        Tracking_Xaxis_Pin(ii) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).ApplyTo.Pins), "_", "")
                                        If Tracking_Xaxis_Pin(ii) = "" Then Tracking_Xaxis_Pin(ii) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).Name), "_", "")
                                        ii = ii + 1 ' Never larger than X_Tracking_Point
                                    Next Tracking_Item
                                End With
                            End If
                            axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).ApplyTo.Pins), "_", "")
                            If axis_pin(curr_axis) = "" Then
                                axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).Parameter.Name), "_", "")
                            End If
                        Case 1 'Y Axis
                            axis_val(curr_axis) = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).Value ''20190319 update
                            Call ValueResolution(axis_val, curr_axis)
'                            Exit Function
                            If Y_Tracking_Point <> 0 Then ' Tracking case
                                With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                                    For Each Tracking_Item In .List
                                        Tracking_Yaxis_val(jj) = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(Tracking_Item).Value
                                        Tracking_Yaxis_Pin(jj) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).ApplyTo.Pins), "_", "")
                                        If Tracking_Yaxis_Pin(jj) = "" Then Tracking_Yaxis_Pin(jj) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).Name), "_", "")
                                        jj = jj + 1 ' Never larger than Y_Tracking_Point
                                    Next Tracking_Item
                                End With
                            End If
                            axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).ApplyTo.Pins), "_", "")
                            If axis_pin(curr_axis) = "" Then
                                axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).Parameter.Name), "_", "")
                            End If
                        Case 2 'Z Axis
                            axis_val(curr_axis) = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).Value ''20190319 update
                            Call ValueResolution(axis_val, curr_axis)
'                            Exit Function
                            If Z_Tracking_Point <> 0 Then ' Tracking case
                                With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                                    For Each Tracking_Item In .List
                                        Tracking_Zaxis_val(kk) = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(Tracking_Item).Value
                                        Tracking_Zaxis_Pin(kk) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).ApplyTo.Pins), "_", "")
                                        If Tracking_Zaxis_Pin(kk) = "" Then Tracking_Zaxis_Pin(kk) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).Name), "_", "")
                                        kk = kk + 1 ' Never larger than Z_Tracking_Point
                                    Next Tracking_Item
                                End With
                            End If
                            axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).ApplyTo.Pins), "_", "")
                            If axis_pin(curr_axis) = "" Then
                                axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).Parameter.Name), "_", "")
                            End If
                    End Select
                Next curr_axis
                ii = 0
                jj = 0
                kk = 0
            Next site
        End If
        
End Function

Public Function ValueResolution(Val() As SiteVariant, asix_index As Variant) As Long
'Public Function ValueResolution(Val As Variant, asix_index As Long) As Long


    If Val(asix_index) >= 1000000 Then
        Val(asix_index) = (Val(asix_index) / 1000000) & "M"
    ElseIf Val(asix_index) >= 1000 And Val(asix_index) < 1000000 Then
        Val(asix_index) = (Val(asix_index) / 1000) & "K"
    ElseIf Val(asix_index) >= 1 And Val(asix_index) < 1000 Then
        ' do nothing
    ElseIf Val(asix_index) >= 0.001 And Val(asix_index) < 1 Then
        Val(asix_index) = (Val(asix_index) * 1000) & "m"
    ElseIf Val(asix_index) >= 0.000001 And Val(asix_index) < 0.001 Then
        Val(asix_index) = (Val(asix_index) * 1000000) & "u"
    End If


End Function

Public Function ValueResolution_1(x_value As Variant, y_value As Variant, Optional z_value As Variant) As Long
'Public Function ValueResolution(Val As Variant, asix_index As Long) As Long


    If x_value >= 1000000 Then
        x_value = (x_value / 1000000) & "M"
    ElseIf x_value >= 1000 And x_value < 1000000 Then
        x_value = (x_value / 1000) & "K"
    ElseIf x_value >= 1 And x_value < 1000 Then
        ' do nothing
    ElseIf x_value >= 0.001 And x_value < 1 Then
        x_value = CStr((x_value * 1000)) & "m"
    ElseIf x_value >= 0.000001 And x_value < 0.001 Then
        x_value = (x_value * 1000000) & "u"
    End If
    
    If y_value >= 1000000 Then
        y_value = (y_value / 1000000) & "M"
    ElseIf y_value >= 1000 And y_value < 1000000 Then
        y_value = (y_value / 1000) & "K"
    ElseIf y_value >= 1 And y_value < 1000 Then
        ' do nothing
    ElseIf y_value >= 0.001 And y_value < 1 Then
        y_value = (y_value * 1000) & "m"
    ElseIf y_value >= 0.000001 And y_value < 0.001 Then
        y_value = (y_value * 1000000) & "u"
    End If
    
    If z_value >= 1000000 Then
        z_value = (z_value / 1000000) & "M"
    ElseIf z_value >= 1000 And z_value < 1000000 Then
        z_value = (z_value / 1000) & "K"
    ElseIf z_value >= 1 And z_value < 1000 Then
        ' do nothing
    ElseIf z_value >= 0.001 And z_value < 1 Then
        z_value = (z_value * 1000) & "m"
    ElseIf z_value >= 0.000001 And z_value < 0.001 Then
        z_value = (z_value * 1000000) & "u"
    End If
    
End Function

Public Sub VaryFreq(ClockPort As String, ClkFreq As Double, ACSpec As String)

Dim site As Variant

    For Each site In TheExec.sites
        TheHdw.Protocol.ports(ClockPort).Halt
        TheHdw.Protocol.ports(ClockPort).Enabled = False
    Next site

    Call TheExec.Overlays.ApplyUniformSpecToHW(ACSpec, ClkFreq)


    TheHdw.Wait 0.003
    TheHdw.Protocol.ports(ClockPort).Enabled = True
    TheHdw.Protocol.ports(ClockPort).NWire.ResetPLL

    TheHdw.Wait 0.001

    Call TheHdw.Protocol.ports(ClockPort).NWire.Frames("RunFreeClock").Execute
    TheHdw.Protocol.ports(ClockPort).IdleWait
End Sub





Public Sub MeasureFreq(MeasPin As String, ByRef result As PinListData)
    On Error GoTo errHandler
    
    With TheHdw.Digital.Pins(MeasPin).FreqCtr
        .Clear
        .EventSlope = Positive
        .EventSource = VOH
        .Interval = 0.01
        .Enable = IntervalEnable
        .start
    End With
    
    TheHdw.Wait 10 * ms
    
    result = TheHdw.Digital.Pins(MeasPin).FreqCtr.Read
    result = result.Math.Divide(TheHdw.Digital.Pins(MeasPin).FreqCtr.Interval)

    Exit Sub
    
errHandler:
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'Public Function save_core_power(power_pins As String, CorePowerStored() As Double)
'    Dim p_ary() As String, p_cnt As Long, i As Long
'    TheExec.DataManager.DecomposePinList power_pins, p_ary, p_cnt
'    ReDim CorePowerStored(p_cnt)
'    For i = 0 To p_cnt - 1
'        If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then CorePowerStored(i) = TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value
'    Next i
'End Function
Public Function restore_core_power(power_pins As String, CorePowerStored() As Double, log_header As String)
    
    Dim p_ary() As String, p_cnt As Long, i As Long
    TheExec.DataManager.DecomposePinList power_pins, p_ary, p_cnt
    For i = 0 To p_cnt - 1
        If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = CorePowerStored(i)
    Next i
    print_core_power log_header, power_pins
End Function
Public Function Decide_shmoo_patt(Init_Patt1 As Pattern, Init_Patt2 As Pattern, Init_Patt3 As Pattern, Init_Patt4 As Pattern, Init_Patt5 As Pattern, _
            Init_Patt6 As Pattern, Init_Patt7 As Pattern, Init_Patt8 As Pattern, Init_Patt9 As Pattern, Init_Patt10 As Pattern, _
            PayLoad_Patt1 As Pattern, PayLoad_Patt2 As Pattern, PayLoad_Patt3 As Pattern, PayLoad_Patt4 As Pattern, PayLoad_Patt5 As Pattern)
    
    Dim tempAry() As String
    Dim i As Integer
    Dim TempAry2() As String
    Dim TempStr As String
    
    Shmoo_Pattern = ""
    
    If Init_Patt1 <> "" Then Shmoo_Pattern = Init_Patt1
    If Shmoo_Pattern <> "" Then
        If Init_Patt2 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt2
    Else
        Shmoo_Pattern = Init_Patt2
    End If
    
    If Shmoo_Pattern <> "" Then
        If Init_Patt3 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt3
    Else
        Shmoo_Pattern = Init_Patt3
    End If
    
    If Shmoo_Pattern <> "" Then
        If Init_Patt4 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt4
    Else
        Shmoo_Pattern = Init_Patt4
    End If
    
    If Shmoo_Pattern <> "" Then
        If Init_Patt5 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt5
    Else
        Shmoo_Pattern = Init_Patt5
    End If
    
    If Shmoo_Pattern <> "" Then
        If Init_Patt6 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt6
    Else
        Shmoo_Pattern = Init_Patt6
    End If
    If Shmoo_Pattern <> "" Then
        If Init_Patt7 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt7
    Else
        Shmoo_Pattern = Init_Patt7
    End If
    If Shmoo_Pattern <> "" Then
        If Init_Patt8 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt8
    Else
        Shmoo_Pattern = Init_Patt8
    End If
    If Shmoo_Pattern <> "" Then
        If Init_Patt9 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt9
    Else
        Shmoo_Pattern = Init_Patt9
    End If
    If Shmoo_Pattern <> "" Then
        If Init_Patt10 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & Init_Patt10
    Else
        Shmoo_Pattern = Init_Patt10
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt1 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt1
    Else
        Shmoo_Pattern = PayLoad_Patt1
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt2 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt2
    Else
        Shmoo_Pattern = PayLoad_Patt2
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt3 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt3
    Else
        Shmoo_Pattern = PayLoad_Patt3
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt4 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt4
    Else
        Shmoo_Pattern = PayLoad_Patt4
    End If
    If Shmoo_Pattern <> "" Then
        If PayLoad_Patt5 <> "" Then Shmoo_Pattern = Shmoo_Pattern & "," & PayLoad_Patt5
    Else
        Shmoo_Pattern = PayLoad_Patt5
    End If
    
    
    tempAry() = Split(Shmoo_Pattern, ",")
    For i = 0 To UBound(tempAry())
        TempAry2() = Split(tempAry(i), ":")
        If i = 0 Then
            TempStr = TempAry2(0)
        Else
            TempStr = TempStr & "," & TempAry2(0)
        End If
    Next i
    Shmoo_Pattern = TempStr
    
End Function




Public Function Run_init_pattern(Shmoo_Pattern_Init As String)
    Dim i As Long
    Dim patt As String
    For i = 0 To MaxCharInitPatt - 1
        patt = char_map_entry(Curr_Shmoo_Condition.Func_block_index).Init_Patt(Curr_Shmoo_Condition.Char_Setup_Index, i)
        If patt <> "" Then
            Call TheHdw.Patterns(patt).Test(pfAlways, 0, tlResultModeDomain)
            If Shmoo_Pattern_Init = "" Then
                Shmoo_Pattern_Init = patt
            Else
                Shmoo_Pattern_Init = Shmoo_Pattern_Init & "," & patt
            End If
        End If
    Next i
End Function
Public Function print_core_power(log_str As String, power_pins As String)
    Dim p_ary() As String, p_cnt As Long, i As Long, j As Long
    Dim out_str As String, InstName As String, ShmooPower As Double
    Dim instanceName As String
    If power_pins = "" Then Exit Function
    
        instanceName = TheExec.DataManager.instanceName

    TheExec.DataManager.DecomposePinList power_pins, p_ary, p_cnt
    For i = 0 To p_cnt - 1
        If Not (TheExec.DataManager.ChannelType(p_ary(i)) Like "N/C") Then
            InstName = GetInstrument(p_ary(i), 0)
            Select Case InstName
               Case "DC-07"
                  ShmooPower = TheHdw.DCVI.Pins(p_ary(i)).Voltage
               Case "VHDVS"
                  ShmooPower = TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value
               Case "HexVS"
                   ShmooPower = TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value
               Case "HSD-U"
               Case Else
                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in print_core_power"
            End Select
            If i = 0 Then
      
                    out_str = instanceName & "(Site" & TheExec.sites.SiteNumber & ")," & Curr_Shmoo_Condition.Char_Setup_Name & "," & Left(log_str & Space(100), 20) & "," & p_ary(i) & "=" & Format(ShmooPower, "0.000")
            
            Else
                out_str = out_str & "," & p_ary(i) & "=" & Format(TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main, "0.000")
            End If
        End If
    Next i
    If TheExec.Flow.EnableWord("Datalog_Verbose") = True Then
        TheExec.Datalog.WriteComment out_str
        TheExec.AddOutput out_str
'        Debug.Print out_str
    End If
End Function
'Public Function vminSearchDCVS(pat As Pattern, TestVoltage As String, StartVoltage As Double, EndVoltage As Double, stepSize As Double, VminSerchPin As Pinlist)
''DTR:F,N99F43-10,9,8,IN03_CZ_MD00_RE0W11_PL01_32_8_LV,17065,.\Pattern\rmhMBIST\PP_ORIA0_S_IN03_BI_XXXX_XXX_JTG_XXX_ALLFV_121211111111.pat,.\Pattern\112712_mbist_char_patDir\CZ_ORIA0_S_PL01_BI_MD00_BST_JTG_XXX_ALLFV_121127111111_RE0W11.pat,LV,32000000,8000000,5,00000,VDD_SOC=0.949974060318,VDD_CPU=0.731,VDD_SRAM_CPU0=0.867,VDD_SRAM_CPU1=0.867,VDD_SRAM_SOC=0.95,XI0,----------------------------------------+++++++++,12000000
''DTR:V,XI0=24000000,N99F43-10,9,8,IN05_CZ_MD00_RE0W11_PL01_04_08_NV,17066,.\Pattern\rmhMBIST\PP_ORIA0_S_IN05_BI_XXXX_XXX_JTG_XXX_ALLFV_121211111111.pat,.\Pattern\112712_mbist_char_patDir\CZ_ORIA0_S_PL01_BI_MD00_BST_JTG_XXX_ALLFV_121127111111_RE0W11.pat,NV,0.4,,0.8,0.01,VDD_SOC=0.949974060318,VDD_CPU=0.907,VDD_SRAM_CPU0=0.867,VDD_SRAM_CPU1=0.867,VDD_SRAM_SOC=0.95,VDD_SRAM_CPU0,VDD_CPU,----------------------------+++++++++++++,0.68
'    'On Error GoTo errHandler
'
'    Dim PatternNames() As String                '<- Array of pattern names
'    Dim PatternName As String                   '<- Individual pattern name
'    Dim PatternCount As Long                    '<- Number of patterns
'    Dim PatIdx As Long                          '<- Pattern loop index
'    Dim Status As Boolean
''    Dim TestPat As Pattern
'   ''===============================================================================
'    Dim v_VDD_SOC As Double
'    Dim v_VDD_CPU As Double
'    Dim v_VDD_SRAM_CPU0 As Double
'    Dim v_VDD_SRAM_CPU1 As Double
'    Dim v_VDD_SRAM_SOC As Double
'    Dim v_Xi0 As Double
'
'    Call thehdw.digital.ApplyLevelsTiming(True, True, True, tlPowered)
'
'    v_VDD_SOC = thehdw.DCVS.pins("VDD_SOC").Voltage.Value
'    v_VDD_CPU = thehdw.DCVS.pins("VDD_CPU").Voltage.Value
'    v_VDD_SRAM_CPU0 = thehdw.DCVS.pins("VDD_SRAM_CPU0").Voltage.Value
'    v_VDD_SRAM_CPU1 = thehdw.DCVS.pins("VDD_SRAM_CPU1").Voltage.Value
'    v_VDD_SRAM_SOC = thehdw.DCVS.pins("VDD_SRAM_SOC").Voltage.Value
'
'   ''===============================================================================
'    thehdw.Patterns(pat).Load
'    Status = PATT_GetPatListFromPatternSet(pat.Value, PatternNames, PatternCount)
'    v_Xi0 = freq_free_run_clk
'
'    'Dim cnt_value As New SiteDouble
'    Dim steps As Integer
'    Dim i As Long
'    Dim outputString As String
'    Dim InstanceName As String
'    Dim TestNum As Long
'    Dim lvccf As Integer
'    Dim lvcc As Double
'
'    lvccf = 0
'    InstanceName = InstanceName
'    TestNum = TheExec.Sites.Item(0).TestNumber
'    Call TheExec.Sites.Item(0).IncrementTestNumber
'
'' Public Hram_LotID_g As String, Hram_WaferID_g As String
'' Public Hram_X_Coor_g As String, Hram_Y_Coor_g As String
''
'    outputString = outputString & "V,XI0=" & CStr(v_Xi0) & "," & Hram_LotID_g & "-" & Hram_WaferID_g & "," & Hram_X_Coor_g & "," & Hram_Y_Coor_g & "," & InstanceName & ","
'    outputString = outputString & TestNum & "," & PatternNames(0) & "," & PatternNames(1) & "," & TestVoltage & ","
'    outputString = outputString & CStr(StartVoltage) & "," & CStr(EndVoltage) & "," & CStr(stepSize) & ","
'    outputString = outputString & "VDD_SOC=" & v_VDD_SOC & "," & "VDD_CPU=" & v_VDD_CPU & "," & "VDD_SRAM_CPU0=" & v_VDD_SRAM_CPU0 & ","
'    outputString = outputString & "VDD_SRAM_CPU1=" & v_VDD_SRAM_CPU1 & "," & "VDD_SRAM_SOC=" & v_VDD_SRAM_SOC & "," & VminSerchPin & ","
'
'    steps = Abs((EndVoltage - StartVoltage) / stepSize)
'    For i = 0 To steps
'
'        thehdw.DCVS.pins(VminSerchPin).Voltage.Value = StartVoltage + i * stepSize
'        thehdw.wait 0.005
'        Call thehdw.Patterns(pat).Start("")
'        Call thehdw.digital.Patgen.HaltWait
'
'        If thehdw.digital.Patgen.PatternBurstPassed Then
'        outputString = outputString & "+"
'
'            lvccf = lvccf + 1
'                If lvccf = 1 Then
'                   lvcc = StartVoltage + i * stepSize
'                End If
'
'        Else
'        outputString = outputString & "-"
'        End If
'
'    Next i
'
'    If lvccf > 0 Then
'        outputString = outputString & "," & CStr(lvcc)
'    Else
'        outputString = outputString & "," & "NA"
'    End If
'
'    TheExec.Datalog.WriteComment outputString
'
'    'turn off instruments
''''    thehdw.Digital.Pins("all_dig").Disconnect
'
''errHandler:
''    TheExec.Datalog.WriteComment "Error encountered in PCM testing"
'End Function
'
'
'    'On Error GoTo errHandler
'
'''    Dim PatternNames() As String                '<- Array of pattern names
'''    Dim PatternName As String                   '<- Individual pattern name
'''    Dim PatternCount As Long                    '<- Number of patterns
'''    Dim PatIdx As Long                          '<- Pattern loop index
'''    Dim Status As Boolean
'''   ''===============================================================================
'''    Dim v_VDD_SOC As Double
'''    Dim v_VDD_CPU As Double
'''    Dim v_VDD_SRAM_CPU0 As Double
'''    Dim v_VDD_SRAM_CPU1 As Double
'''    Dim v_VDD_SRAM_SOC As Double
'''    Dim v_XI0 As Double
'''
'''    Call thehdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
'''
'''    v_VDD_SOC = thehdw.DCVS.pins("VDD_SOC").Voltage.Value
'''    v_VDD_CPU = thehdw.DCVS.pins("VDD_CPU").Voltage.Value
'''    v_VDD_SRAM_CPU0 = thehdw.DCVS.pins("VDD_SRAM_CPU0").Voltage.Value
'''    v_VDD_SRAM_CPU1 = thehdw.DCVS.pins("VDD_SRAM_CPU1").Voltage.Value
'''    v_VDD_SRAM_SOC = thehdw.DCVS.pins("VDD_SRAM_SOC").Voltage.Value
'''
'''   ''===============================================================================
'''    thehdw.Patterns(pat).Load
'''    Status = PATT_GetPatListFromPatternSet(pat, "DLLINTD", PatternNames, PatternCount)
'''
'''    'Dim cnt_value As New SiteDouble
'''    Dim steps As Integer
'''    Dim i As Long
'''    Dim outputString As String
'''    Dim InstanceName As String
'''    Dim testNum As Long
'''    Dim fmaxf As Integer
'''    Dim fmax As Double
'''
'''    InstanceName = InstanceName
'''    testNum = TheExec.Sites.Item(0).TestNumber
'''    Call TheExec.Sites.Item(0).IncrementTestNumber
'''
'''    outputString = outputString & "F," & Hram_LotID_g & "-" & Hram_WaferID_g & "," & Hram_X_Coor_g & "," & Hram_Y_Coor_g & "," & InstanceName & ","
'''    outputString = outputString & testNum & "," & PatternNames(0) & "," & PatternNames(1) & "," & TestVoltage & ","
'''    outputString = outputString & CStr(StartFreq) & "," & CStr(EndFreq) & "," & CStr(stepSize) & ","
'''    outputString = outputString & "VDD_SOC=" & v_VDD_SOC & "," & "VDD_CPU=" & v_VDD_CPU & "," & "VDD_SRAM_CPU0=" & v_VDD_SRAM_CPU0 & ","
'''    outputString = outputString & "VDD_SRAM_CPU1=" & v_VDD_SRAM_CPU1 & "," & "VDD_SRAM_SOC=" & v_VDD_SRAM_SOC & ","
'''    outputString = outputString & FreqSerchPin & ","
'''
'''    steps = Abs((EndFreq - StartFreq) / stepSize)
'''
'''    For i = 0 To steps
'''
'''        thehdw.DIB.SupportBoardClock.Frequency = StartFreq - i * stepSize
'''        thehdw.Wait 0.01
'''        Call thehdw.Patterns(pat).Start("")
'''        Call thehdw.Digital.Patgen.HaltWait
'''
'''        If thehdw.Digital.Patgen.PatternBurstPassed Then
'''        outputString = outputString & "+"
'''
'''            fmaxf = fmaxf + 1
'''                If fmaxf = 1 Then
'''                    fmax = StartFreq - i * stepSize
'''                End If
'''
'''        Else
'''        outputString = outputString & "-"
'''        End If
'''    Next i
'''
'''    If fmaxf > 0 Then
'''       outputString = outputString & "," & CStr(fmax)
'''    Else
'''       outputString = outputString & "," & "NA"
'''    End If
'''
'''    TheExec.Datalog.WriteComment outputString
'''    'turn off instruments
''''''    thehdw.Digital.Pins("all_dig").Disconnect
'''
''''errHandler:
''''    TheExec.Datalog.WriteComment "Error encountered in PCM testing"
'''End Function
'

''            For j = 0 To MaxCharCorePower - 1
''                If char_map_entry(Curr_Shmoo_Condition.Func_block_index).Core_Power(i, j) <> "" Then
''                    Select Case j
''                        Case 0: power_pins = power_pins & ",VDD_CPU"
''                        Case 1: power_pins = power_pins & ",VDD_SOC"
''                        Case 2: power_pins = power_pins & ",VDD_FIXED"
''                        Case 3: power_pins = power_pins & ",VDD_SRAM"
''                    End Select
''                End If
''            Next j
            
'Public Function start_nwire_XI0(freq As Double)
'    If LCase(TheExec.CurrentChanMap) Like "*_site" Then
'        Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Freq", freq, , False)
'       Call SetFRCPath("OutputClk_XI0_Diff")
'        TheHdw.Wait (10 * ms)
'        TheHdw.Protocol.Ports("Clock_port").Enabled = True
'        TheHdw.Protocol.Ports("Clock_port").NWire.ResetPLL
'        ' Start the nWire engine.
'        Call TheHdw.Protocol.Ports("Clock_port").NWire.Frames("RunFreeClock").Execute
'        TheHdw.Protocol.Ports("Clock_port").IdleWait
'    Else
'        TheHdw.DIB.SupportBoardClock.Frequency = freq
'    End If
'    TheHdw.Wait 0.003
'End Function

Public Function Shmoo_Save_core_power_per_site(power_pins As String, ShmooPower() As SiteDouble)
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    On Error GoTo errHandler
    TheExec.DataManager.DecomposePinList power_pins, p_ary, p_cnt
    For i = 0 To p_cnt - 1
        If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
            InstName = GetInstrument(p_ary(i), 0)
            Select Case InstName
               Case "DC-07"
                  ShmooPower(i) = TheHdw.DCVI.Pins(p_ary(i)).Voltage
               Case "VHDVS"
                  ShmooPower(i) = TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value
               Case "HexVS"
                   ShmooPower(i) = TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value
               Case "HSD-U"
               Case Else
                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Save_core_power_per_site"
            End Select
        End If
    Next i
    
   Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> Shmoo_Save_core_power_per_site:: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function Get_Current_Apply_Pin(power_pins As String)
' Only get Power pins needed  for shmoo
' Ignore any IO pins and FreeRun Freq pins
    Dim active_setup As String, curr_axis As Variant, curr_track As Variant, apply_Pin As String, apply_Pin_arry() As String, pin_count As Long, i As Long
    Dim p_ary() As String, p_cnt As Long
    power_pins = ""
    Set g_Globalpointval = Nothing
    active_setup = TheExec.DevChar.Setups.ActiveSetupName
    For Each curr_axis In TheExec.DevChar.Setups(active_setup).Shmoo.Axes.List
        ''exit for if any axis is not power pin -by SY
        If TheExec.DevChar.Setups(active_setup).Shmoo.Axes(curr_axis).ApplyTo.Pins = "" Then Exit For
        apply_Pin = TheExec.DevChar.Setups(active_setup).Shmoo.Axes(curr_axis).ApplyTo.Pins
        
'        Add for store shmoo global spec to avoid direct to apply Vmain used for Vbump function
        If g_Vbump_function = True Then
           Call TheExec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count)
           For i = 0 To pin_count - 1
               g_Globalpointval.AddPin (apply_Pin_arry(i))
               For Each site In TheExec.sites
                   g_Globalpointval.Pins(apply_Pin_arry(i)).Value = TheExec.DevChar.Results(active_setup).Shmoo.CurrentPoint.Axes(curr_axis).Value
               Next site
           Next i
        End If
        If apply_Pin <> "" Then
            If power_pins <> "" Then
                power_pins = power_pins & "," & apply_Pin
            Else
                power_pins = apply_Pin
            End If
        End If
        For Each curr_track In TheExec.DevChar.Setups(active_setup).Shmoo.Axes(curr_axis).TrackingParameters.List
            apply_Pin = TheExec.DevChar.Setups(active_setup).Shmoo.Axes(curr_axis).TrackingParameters.Item(curr_track).ApplyTo.Pins
            If g_Vbump_function = True Then
               Call TheExec.DataManager.DecomposePinList(apply_Pin, apply_Pin_arry, pin_count)
               For i = 0 To pin_count - 1
                   g_Globalpointval.AddPin (apply_Pin_arry(i))
                   For Each site In TheExec.sites
                       g_Globalpointval.Pins(apply_Pin_arry(i)).Value = TheExec.DevChar.Results(active_setup).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(curr_track).Value
                   Next site
               Next i
            End If
            power_pins = power_pins & "," & apply_Pin
        Next curr_track
   Next curr_axis
   TheExec.DataManager.DecomposePinList power_pins, p_ary, p_cnt
   power_pins = Join(p_ary, ",")
End Function
Public Function Shmoo_Restore_Power_per_site(ShmooPowerStored_Pins As String, ShmooPowerStored() As SiteDouble, log_header As String, Optional Restore_Pins As String = "")
    'if Restore_Pins="" then restore all ShmooPowerStored_Pins
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim rp_ary() As String, rp_cnt As Long
    Dim InstName As String
    Dim tmp_ShmooPowerStored_Pins() As String
    Dim p As Variant, pn As String
    Dim Need_ReStore_Pin As Boolean
    Dim Restore_Pins_Dict As New Dictionary
    Dim Restore_Pin_str As String
    Dim ShmooPowerStored_Pins_str  As String
    
    If ShmooPowerStored_Pins = "" Then Exit Function
    
    If Restore_Pins = "" Then
        Restore_Pin_str = ShmooPowerStored_Pins
    Else
        Restore_Pin_str = Restore_Pins
    End If
    TheExec.DataManager.DecomposePinList ShmooPowerStored_Pins, p_ary, p_cnt
    
    TheExec.DataManager.DecomposePinList Restore_Pin_str, rp_ary, rp_cnt
    Restore_Pin_str = Join(rp_ary, ",")
   Create_Pin_Dic Restore_Pin_str, Restore_Pins_Dict

    For i = 0 To p_cnt - 1
        If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" And Restore_Pins_Dict.Exists(LCase(p_ary(i))) = True Then
            InstName = GetInstrument(p_ary(i), 0)
            Select Case InstName
               Case "DC-07"
                   TheHdw.DCVI.Pins(p_ary(i)).Voltage = ShmooPowerStored(i)
               Case "VHDVS"
                   TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = ShmooPowerStored(i)
               Case "HexVS"
                   TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = ShmooPowerStored(i)
               Case "HSD-U"
               Case Else
                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site"
            End Select
        End If
    Next i
    'print_core_power log_header, ShmooPowerStored_Pins
End Function
Public Function Create_Pin_Dic(Pins As String, Pin_Dict As Dictionary)
    Dim p_ary() As String, p_cnt As Long, pn As String, p As Variant
    TheExec.DataManager.DecomposePinList Pins, p_ary, p_cnt
    Pin_Dict.RemoveAll
    For Each p In p_ary
        pn = LCase(CStr(p))
        Pin_Dict.Add pn, True
    Next p
End Function
Public Function Shmoo_Set_Power(power_pins As String, Level As String, log_header As String, Optional Use_Performance_Mode As Boolean = False, Optional skip_pin As String = "")
    Dim p_ary() As String, p_cnt As Long, i As Long, j As Long
    Dim Core_p_ary() As String, Core_p_cnt As Long
    Dim main_power As String, main_spec_name As String
    Dim Ratio As Double
    Dim Flag_core_power_found As Boolean
    Dim p_mode As String, p_mode_code As Long, block_name As String
    Dim p_mode_code_str As String
    Dim tmp_ary() As String
    Dim shmoo_pin As String
    Dim Active_Test_inst_name As String
    Dim Dc_cat As String, Dc_spec_type As String
    Dim sp As Variant, t As String
    Dim InstName As String
    Dim Skip_Pin_Dic As New Dictionary
    Dim Need_Skip_Pin As Boolean
    ' Assumption:
    ' 1. Only use Selector :Typ,Max,Min
    ' 2. DC spec name is  VDD_CPU_VAR_C/S/G/H
    ' 3. DC spec will not be changed
    If power_pins = "" Then Exit Function
    
    If skip_pin <> "" Then Create_Pin_Dic skip_pin, Skip_Pin_Dic

    TheExec.DataManager.DecomposePinList power_pins, p_ary, p_cnt
    TheExec.DataManager.GetInstanceContext Dc_cat, t, t, t, t, t, t, t
    For Each sp In TheExec.specs.DC.Categories(Dc_cat).SpecList
        sp = LCase(sp)
        If sp Like "*_var_c" Then
            Dc_spec_type = "C"
        ElseIf sp Like "*_var_g" Then
            Dc_spec_type = "G"
        ElseIf sp Like "*_var_s" Then
            Dc_spec_type = "S"
        ElseIf sp Like "*_var_h" Then
            Dc_spec_type = "H"
        Else
            TheExec.ErrorLogMessage "DC spec " & sp & " is not ended with _VAR_C/S/G/H in " & TheExec.DataManager.instanceName
        End If
        Exit For
    Next sp
    For i = 0 To p_cnt - 1
        p_ary(i) = LCase(p_ary(i))
        Need_Skip_Pin = False
        If skip_pin <> "" Then
            If Skip_Pin_Dic.Exists(p_ary(i)) = True Then Need_Skip_Pin = True
        End If
        If Not (TheExec.DataManager.ChannelType(p_ary(i)) Like "N/C") And Need_Skip_Pin = False Then
            InstName = GetInstrument(p_ary(i), 0)
            Select Case InstName
               Case "DC-07":
                        Select Case Level
                            Case "NV":  TheHdw.DCVI.Pins(p_ary(i)).Voltage = TheExec.specs.DC.Item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(Dc_cat).Typ.Value
                            Case "LV":  TheHdw.DCVI.Pins(p_ary(i)).Voltage = TheExec.specs.DC.Item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(Dc_cat).Min.Value
                            Case "HV":  TheHdw.DCVI.Pins(p_ary(i)).Voltage = TheExec.specs.DC.Item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(Dc_cat).max.Value
                            Case Else
                                TheExec.ErrorLogMessage Level & " is not supported in " & TheExec.DataManager.instanceName
                        End Select
               Case "VHDVS":
                        Select Case Level
                            Case "NV":  TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = TheExec.specs.DC.Item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(Dc_cat).Typ.Value
                            Case "LV":  TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = TheExec.specs.DC.Item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(Dc_cat).Min.Value
                            Case "HV":  TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = TheExec.specs.DC.Item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(Dc_cat).max.Value
                            Case Else
                                TheExec.ErrorLogMessage Level & " is not supported in " & TheExec.DataManager.instanceName
                        End Select
               Case "HexVS":
                        Select Case Level
                            Case "NV":  TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = TheExec.specs.DC.Item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(Dc_cat).Typ.Value
                            Case "LV":  TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = TheExec.specs.DC.Item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(Dc_cat).Min.Value
                            Case "HV":  TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = TheExec.specs.DC.Item(p_ary(i) & "_" & "VAR" & "_" & Dc_spec_type).Categories(Dc_cat).max.Value
                            Case Else
                                TheExec.ErrorLogMessage Level & " is not supported in " & TheExec.DataManager.instanceName
                        End Select
               Case "HSD-U"
               Case Else
                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Set_Power"
            End Select

        End If
'         If TheExec.Flow.EnableWord("Datalog_Verbose") = True Then TheExec.Datalog.WriteComment log_header & " Shmoo Pin (" & p_ary(i) & ")= " & thehdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value
    Next i
    'Override with Char_Map CorePower Value
    'print_core_power log_header, power_pins
    TheHdw.Wait 0.001
End Function

'20161229 Roy Modified
'Get NV value from Hardware instead of PinName
Public Function Decide_NV(Step_NV As Long, RangeStart As Double, RangeStepSize As Double, step_Start As Long, Step_x As Long, SetupName As String)
    Dim shmoo_pin As String
    Dim shmoo_Parameter_Type As String
    Dim shmoo_Parameter_Name As String
    Dim testValue As Double
    Dim p As Variant
    Dim Pins As String
    Dim p_ary() As String
    Dim P_ary_cnt As Long
    Dim pm_check As Boolean
    Dim Shmoo_TestName, tmp_ary, block_name, Level As String
    Dim p_mode_code As Long
    Dim shmooo_pin As String
    Dim Ratio As Long
    Dim real_shmoo_pin As String
    Dim InstName As String
    '-----------------------------------------------
    Dim Shmoo_setup_name, Shmoo_TestInst_Name As String
    
    Shmoo_setup_name = TheExec.DevChar.Setups.ActiveSetupName
    Shmoo_TestInst_Name = TheExec.DevChar.ActiveDataObject.testName
    shmoo_pin = TheExec.DevChar.Setups(Shmoo_setup_name).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins
    Set_Pin_NV = 0
    Call TheExec.DataManager.DecomposePinList(shmoo_pin, p_ary, P_ary_cnt)

    real_shmoo_pin = p_ary(0)
    
    InstName = GetInstrument(real_shmoo_pin, 0)
    Select Case InstName
        Case "DC-07"
            Set_Pin_NV = Format(TheHdw.DCVI.Pins(real_shmoo_pin).Voltage, "0.0000")
        Case "VHDVS"
            Set_Pin_NV = Format(TheHdw.DCVS.Pins(real_shmoo_pin).Voltage.Main, "0.0000")
        Case "HexVS"
            Set_Pin_NV = Format(TheHdw.DCVS.Pins(real_shmoo_pin).Voltage.Main, "0.0000")
        Case Else
    End Select

    
'////Original  20161229 Roy Modified///
'''    tmp_ary = Split(Shmoo_TestInst_Name, "_")
'''    If UBound(tmp_ary) > 0 Then
'''        block_name = LCase(tmp_ary(1)) ' SOC, CPU,GPU/GFX
'''
'''    End If
'''
'''    If (InStr(tmp_ary(3), "MC") > 0 Or InStr(tmp_ary(3), "MG") > 0 Or InStr(tmp_ary(3), "MS") > 0) Then
'''        p_mode_code = CLng(Mid(tmp_ary(3), InStr(tmp_ary(3), "M") + 2)) 'MC1, MG001, MS01
'''    ElseIf (InStr(LCase(tmp_ary(4)), "marginhd") > 0) Then 'MCL_CPU_CPUSRAM_X_CPUMarginHD_CL51_X_PLLP_59N_X_CZ_NV
'''        p_mode_code = 0
'''    ElseIf (InStr(LCase(tmp_ary(4)), "marginbc") > 0) Then
'''        p_mode_code = 0
'''    End If
'''
'''
'''    tmp_ary = Split(TheExec.DataManager.InstanceName, "_")
'''    Level = tmp_ary(UBound(tmp_ary))
'''    Select Case Level
'''        Case "NV": ratio = 1
'''        Case "LV": ratio = TheExec.Specs.Globals("Ratio_Minus").ContextValue
'''        Case "HV": ratio = TheExec.Specs.Globals("Ratio_Plus").ContextValue
'''        Case Else
'''            ratio = 1
'''            'TheExec.ErrorLogMessage Level & " is not supported in " & TheExec.DataManager.InstanceName
'''    End Select
'''
'''    If p_mode_code = 999 Then
'''        p_mode_code = 0
'''    End If
'''
'''    '/////////////////////////////cpu////////////////////////////////////////////
'''    If LCase(block_name) Like "cpu*" Then
'''        If shmoo_pin Like "VDD_CPU" Then
'''            If p_mode_code = 0 Then
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_CPU_GLB").ContextValue
'''            Else
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_CPU_P" & p_mode_code & "_GLB").ContextValue
'''            End If
'''        ElseIf shmoo_pin Like "VDD_CPU_SRAM" Then
'''            If p_mode_code = 0 Then
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_CPU_SRAM_GLB").ContextValue
'''            Else
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_CPU_SRAM_P" & p_mode_code & "_GLB").ContextValue
'''            End If
'''        Else
'''            If shmoo_pin Like "VDD*" Then
'''                Set_Pin_NV = TheExec.Specs.Globals(shmoo_pin & "_GLB").ContextValue
'''            End If
'''        End If
'''    End If
'''    '/////////////////////////////////////////   GPU   //////////////////////////////////////////////////////////
'''    If LCase(block_name) Like "gpu*" Then
'''        If shmoo_pin Like "VDD_GPU" Then
'''            If p_mode_code = 0 Then
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_GPU_GLB").ContextValue
'''            Else
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_GPU_P" & p_mode_code & "_GLB").ContextValue
'''            End If
'''        ElseIf shmoo_pin Like "VDD_GPU_SRAM" Then
'''            If p_mode_code = 0 Then
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_GPU_SRAM_GLB").ContextValue
'''            Else
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_GPU_SRAM_P" & p_mode_code & "_GLB").ContextValue
'''            End If
'''
'''        ElseIf shmoo_pin Like "VDD_SOC" Or shmoo_pin Like "VDD_LOW" Then
'''            If Level = "LV" Then
'''                    Set_Pin_NV = TheExec.Specs.Globals("VDD_SOC_P1_GLB").ContextValue
'''            ElseIf Level = "HV" Then
'''                    Set_Pin_NV = TheExec.Specs.Globals("VDD_SOC_P2_GLB").ContextValue
'''            Else
'''                    Set_Pin_NV = TheExec.Specs.Globals("VDD_SOC_GLB").ContextValue
'''            End If
'''        Else
'''            If shmoo_pin Like "VDD*" Then
'''                Set_Pin_NV = TheExec.Specs.Globals(shmoo_pin & "_GLB").ContextValue
'''            End If
'''        End If
'''    End If
'''    '/////////////////////////////////////////   SOC   //////////////////////////////////////////////////////////
'''    If LCase(block_name) Like "soc*" Then
'''        If shmoo_pin Like "VDD_SOC" Then
'''            If p_mode_code = 0 Then
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_SOC_GLB").ContextValue
'''            Else
'''                Set_Pin_NV = TheExec.Specs.Globals("VDD_SOC_P" & p_mode_code & "_GLB").ContextValue
'''            End If
'''        Else
'''            If shmoo_pin Like "VDD*" Then
'''                Set_Pin_NV = TheExec.Specs.Globals(shmoo_pin & "_GLB").ContextValue
'''            End If
'''        End If
'''    End If
'''    '/////////////////////////////////////////////////////////////////////////////////////////////////////////////
'''    Dim cType As String
'''    Dim cTyype_temp() As String
'''
'''    If InStr(real_shmoo_pin, ",") > 0 Then
'''        cTyype_temp = Split(real_shmoo_pin, ",")
'''        cType = TheExec.DataManager.ChannelType(cTyype_temp(0))
'''    Else
'''        cType = TheExec.DataManager.ChannelType(real_shmoo_pin)
'''    End If
'''
''''    If (InStr(tmp_ary(0), "H") > 0 And cType <> "I/O") Then
''''        Set_Pin_NV = TheExec.specs.Globals(shmoo_pin & "_GLB").ContextValue
''''    End If
''''
''''    TheExec.Datalog.WriteComment "Shmoo Pin = " & shmoo_pin
''''    TheExec.Datalog.WriteComment "NV = " & Set_Pin_NV
'''
'''
'''    '-----------------------------------------------
'////Original ///
    
    If TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins <> "" Then
        p_ary = Split(TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins, ",")
        shmoo_pin = p_ary(0) ' Pick the first pin to calculate the Step NV
    End If

    shmoo_Parameter_Name = TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Name
    shmoo_Parameter_Type = TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type
    Shmoo_TestName = TheExec.DevChar.ActiveDataObject.testName


    With TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X)
        If LCase(TheExec.DataManager.PinType(shmoo_pin)) = "power" Then
            If Step_x < 0 Then Step_NV = Abs((Set_Pin_NV - RangeStart) / (Step_x * RangeStepSize))
            If Step_x > 0 Then Step_NV = (Set_Pin_NV - RangeStart) / (Step_x * RangeStepSize) + step_Start
            
        ElseIf shmoo_Parameter_Type Like "AC Spec" Then
            testValue = TheExec.specs.AC.Item(shmoo_Parameter_Name).ContextValue
            Step_NV = (testValue - RangeStart) / (Step_x * RangeStepSize) + step_Start

        End If
    End With
    
End Function

Public Function Decide_LVCC_HVCC(Vcc_min As String, Vcc_max As String, Shmoo_hole As String, Step_NV As Long, RangeLow As Double, RangeStepSize As Double, Shmoo_result_PF As String, SetupName As String, step_Start As Long, step_Stop As Long, Step_x As Long)
    Dim FlagFirstPass As Boolean, FlagFirstFail As Boolean
    Dim last_point_result As String, current_point_result As String, char_pt As String
    Dim AllFail As Boolean
    Dim min_point As Long, max_point As Long, current_point As Long
    Dim FlagHole As Boolean
    Dim FlagPF(1000) As Boolean
    Dim FlagFP(1000) As Boolean
    Dim FlagPF_Count As Long
    Dim FlagFP_Count As Long
    Dim i As Long, j As Long
    Dim test_name As String
    Dim step_p As Long
    Dim x_pra As String
    Dim show_vcc As String
    Dim shmoo_form, shmoo_stop, shmoo_step As String
    Dim AllPass As Boolean
    Dim lvcc_point As Integer
    Dim hvcc_point As Integer
    Dim str_Temp() As String
    Dim mode_type As String
    Dim Point_Volt() As Double
    
    Vcc_min = ""
    Vcc_max = ""
    
    show_vcc = "[Shmoo,"
    
    Dim Shmoo_setup_name, Shmoo_TestInst_Name As String
    
    step_p = Len(Shmoo_result_PF)
    
    Shmoo_setup_name = TheExec.DevChar.Setups.ActiveSetupName
    Shmoo_TestInst_Name = TheExec.DevChar.ActiveDataObject.testName
    shmoo_form = CStr(RangeLow)
    shmoo_stop = CStr(RangeLow + RangeStepSize * step_p)
    shmoo_step = CStr(step_p + 1)
    x_pra = TheExec.DevChar.ActiveDataObject.XParameter
    str_Temp = Split(Shmoo_TestInst_Name, "_")
    mode_type = str_Temp(0)
'    If (TheExec.EnableWord("One_transition") = True) Then
        If (LCase(Shmoo_TestInst_Name) Like "dfth*" Or LCase(Shmoo_TestInst_Name) Like "hfh*" Or LCase(Shmoo_TestInst_Name) Like "mch*") Then
            If step_Start > step_Stop Then
                Step_NV = step_Stop
            Else
                Step_NV = step_Start
            End If
        ElseIf (LCase(Shmoo_TestInst_Name) Like "dftl*" Or LCase(Shmoo_TestInst_Name) Like "hfl*" Or LCase(Shmoo_TestInst_Name) Like "mcl*") Then
            If step_Start > step_Stop Then
                Step_NV = step_Start
            Else
                Step_NV = step_Stop
            End If
        End If
'    End If
    
    If (Step_NV > step_p Or Step_NV < 0) Then
        If (LCase(Shmoo_TestInst_Name) Like "dfth*" Or LCase(Shmoo_TestInst_Name) Like "hfh*" Or LCase(Shmoo_TestInst_Name) Like "mch*") Then
            Step_NV = step_Start
        
        ElseIf (LCase(Shmoo_TestInst_Name) Like "dftl*" Or LCase(Shmoo_TestInst_Name) Like "hfl*" Or LCase(Shmoo_TestInst_Name) Like "mcl*") Then
            Step_NV = step_Stop
        End If
    End If

    test_name = TheExec.DevChar.ActiveDataObject.testName
    
    
    Shmoo_hole = "NH"

    'Early exit if
    '   NV fails
    '   All Fail
    '   Not pass or fail
'-----------------------------------------------------------------------------------------------
    AllPass = True
    ' if fails at NV
    ' step_start is the lowest char value
    ' Check if all points fail
    
    For i = step_Start To step_Stop Step Step_x
        char_pt = Mid(Shmoo_result_PF, i + 1, 1)
        If char_pt = "F" Then
            AllPass = False
        End If
    Next i
    
    If AllPass = True Then
        Vcc_min = CStr(RangeLow): Vcc_max = CStr(RangeLow + RangeStepSize * Abs(step_Stop - step_Start))
        GoTo end_lvcc_hvcc
    End If
'-----------------------------------------------------------------------------------------------
    AllFail = True
    ' if fails at NV
    ' step_start is the lowest char value
    ' Check if all points fail
    
    For i = step_Start To step_Stop Step Step_x
        char_pt = Mid(Shmoo_result_PF, i + 1, 1)
        If char_pt = "P" Then
            AllFail = False
        End If
        If Not (char_pt = "P" Or char_pt = "F") Then
           Vcc_min = "-7777": Vcc_max = "7777"
           GoTo end_lvcc_hvcc
        End If
    Next i
    
    If AllFail = True Then
        Vcc_max = "9999": Vcc_min = "-9999"
        GoTo end_lvcc_hvcc
    End If
    
    
'%%%%%%%%%%%%%%%%%%%%%%%% NV Fail ,Report 5555  (Open while Check HH,LH)%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'20170213 add boolean to control NV hole
Dim NVffff As Boolean
NVffff = False
'20170105 Roy added
If NVffff = True Then
    If Step_NV > 0 Then
        If Mid(Shmoo_result_PF, Step_NV, 1) = "F" Then
            Vcc_max = "5555":  Vcc_min = "-5555"
            GoTo end_lvcc_hvcc
        End If
    End If
Else

'    If Step_NV > 0 Then
'        If Mid(Shmoo_result_PF, Step_NV, 1) = "F" Then
'            Vcc_max = "5555":  Vcc_min = "-5555"
'            GoTo end_lvcc_hvcc
'        End If
'    End If

End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%  HF VID,VICM%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       'hvcc_point or lvcc_point is -1,shmoo hole
        If (UCase(mode_type) Like "HF*") Or (UCase(x_pra) = "VID") Or (UCase(x_pra) = "VICM") Then
            ReDim Point_Volt(step_p) As Double
            For i = 1 To step_p
                Point_Volt(i) = RangeLow + RangeStepSize * (i - 1)
            Next i
            
            hvcc_point = Search_HVCC(Shmoo_result_PF)
            lvcc_point = Search_LVCC(Shmoo_result_PF)
            If hvcc_point > step_p Then hvcc_point = step_p
            If (hvcc_point = -1) Then
                Vcc_max = "5555"
            Else
                Vcc_max = CStr(Point_Volt(hvcc_point))
            End If
            
            If (lvcc_point = -1) Then
                Vcc_min = "-5555"
            Else
                Vcc_min = CStr(Point_Volt(lvcc_point))
            End If
            
            GoTo end_lvcc_hvcc
        End If

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%  VIH,VIL %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       'hvcc_point or lvcc_point is -1,shmoo hole
       'hvcc_point or lvcc_point is -2,first point fail
If UCase(TheExec.DataManager.instanceName) Like "*USBPICO*" Or UCase(TheExec.DataManager.instanceName) Like "*LPDPRX*" Then
Else
        If UCase(mode_type) Like "HIO*" Then
            If (InStr(LCase(x_pra), "_vih_") > 0) Or (InStr(LCase(x_pra), "_vil_") > 0) Then
                ReDim Point_Volt(step_p) As Double
                For i = 1 To step_p
                    Point_Volt(i) = RangeLow + RangeStepSize * (i - 1)
                Next i
                If (LCase(x_pra) = "vih") Then
                    lvcc_point = Search_VIH_LVCC(Shmoo_result_PF)
                    If (lvcc_point = -1) Then
                        Vcc_min = "-5555"
                    ElseIf (lvcc_point = -2) Then
                        Vcc_min = "-8888"
                    Else
                        Vcc_min = CStr(Point_Volt(lvcc_point))
                    End If
                End If
                If (LCase(x_pra) = "vil") Then
                    hvcc_point = Search_VIL_HVCC(Shmoo_result_PF)
                    If (hvcc_point = -1) Then
                        Vcc_max = "5555"
                    ElseIf (lvcc_point = -2) Then
                        Vcc_min = "8888"
                    Else
                        Vcc_max = CStr(Point_Volt(hvcc_point))
                    End If
                End If
                GoTo end_lvcc_hvcc
            End If
        End If
End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    
    '================================================
    
'    If (InStr(UCase(x_pra), "*VID*") > 0 Or InStr(UCase(x_pra), "*VICM*") > 0) Then
'        ReDim Point_Volt(step_p) As Double
'            For i = 1 To step_p
'                Point_Volt(i) = RangeLow + RangeStepSize * (i - 1)
'            Next i
'
'            hvcc_point = Search_HVCC(Shmoo_result_PF)
'            lvcc_point = Search_LVCC(Shmoo_result_PF)
'            If (hvcc_point = -1) Then
'                Vcc_max = "5555"
'            Else
'                Vcc_max = CStr(Point_Volt(hvcc_point))
'            End If
'
'            If (lvcc_point = -1) Then
'                Vcc_min = "-5555"
'            Else
'                Vcc_min = CStr(Point_Volt(lvcc_point))
'            End If
'
'            GoTo end_lvcc_hvcc
'    End If
    
    
    '================================================
    ReDim Point_Volt(step_p) As Double
    For i = 1 To step_p
        Point_Volt(i) = RangeLow + RangeStepSize * (i - 1)
    Next i
    
    hvcc_point = Search_HVCC(Shmoo_result_PF)
    lvcc_point = Search_LVCC(Shmoo_result_PF)
    If hvcc_point > step_p Then hvcc_point = step_p
    If (hvcc_point = -1) Then
        Vcc_max = "5555"
    Else
        Vcc_max = CStr(Point_Volt(hvcc_point))
    End If
    
    If (lvcc_point = -1) Then
        Vcc_min = "-5555"
    Else
        Vcc_min = CStr(Point_Volt(lvcc_point))
    End If
    
    GoTo end_lvcc_hvcc
    
    If (Step_NV = -1) Then
        
        If step_Start > step_Stop Then
            Step_NV = step_Stop
        Else
            Step_NV = step_Start
        End If
        
    End If
   
    
    
'----------------------------------------------------------------------------------------------
    If Not (InStr(LCase(test_name), "vih") > 0 Or InStr(LCase(test_name), "vil") > 0 Or InStr(LCase(x_pra), "vih") > 0 Or InStr(LCase(x_pra), "vil") > 0 Or InStr(LCase(x_pra), "vid") > 0) Then
        If Mid(Shmoo_result_PF, Step_NV + 1, 1) = "F" Then
        
            For i = Step_NV To (step_Stop - step_Start) / Step_x    'search low to high voltage
                char_pt = Mid(Shmoo_result_PF, i + 1, 1)
                If char_pt = "P" Then
                    Vcc_max = 8888
                    i = (step_Stop - step_Start) / Step_x
                End If
            Next i
            If Vcc_max = "" Then Vcc_max = 9999
            
            
            For i = Step_NV To 0 Step -1    'search low to high voltage
                char_pt = Mid(Shmoo_result_PF, i + 1, 1)
                If char_pt = "P" Then
                    Vcc_min = -8888
                    i = 0
                End If
            Next i
            If Vcc_min = "" Then Vcc_min = -9999
            

            
            Exit Function
        End If
    End If
'-------------------------------------------------------------------------------------------------
    ' Find LVCC: NV value to Low value
    ' 01         Step
    ' FFPPPpPPFF
    ' PPPPPpPPFF
    
    lvcc_point = 0
    
    If Not (InStr(LCase(test_name), "vih") > 0 Or InStr(LCase(test_name), "vil") > 0 Or InStr(LCase(x_pra), "vih") > 0 Or InStr(LCase(x_pra), "vil") > 0 Or InStr(LCase(x_pra), "vid") > 0) Then

        For i = Step_NV To 0 Step -1
            If Mid(Shmoo_result_PF, i + 1, 1) = "F" Then
                
                
                Vcc_min = CStr(RangeLow + RangeStepSize * (i + 1))
'                TheExec.Datalog.WriteComment "[LVCC=" & Vcc_min & "]"
                lvcc_point = i
                i = 0
            End If
        Next i
        
        
        
        
        If Vcc_min = "" Then Vcc_min = CStr(RangeLow)
    Else
        If InStr(LCase(test_name), "vih") > 0 Or InStr(LCase(x_pra), "vih") > 0 Or InStr(LCase(x_pra), "vid") > 0 Then
            
            
            For i = (step_Stop - step_Start) / Step_x To 0 Step -1 'search high to low voltage
                char_pt = Mid(Shmoo_result_PF, i + 1, 1)
                If char_pt = "F" Then
                    Vcc_min = CStr(RangeLow + RangeStepSize * (i + 1))
                    lvcc_point = i
                    i = 0
                End If
            Next i
            If Vcc_min = "" Then Vcc_min = CStr(RangeLow)
        End If
        
    End If
    
    
        
    If lvcc_point <> 0 Then
        For i = lvcc_point - 1 To 0 Step -1
            If Mid(Shmoo_result_PF, i + 1, 1) = "P" Then
                Vcc_min = "-5555"
            End If
            
        Next i
    End If
    
    Dim Fail_index As Integer
    ''*******************************************AI***********************************************
    If TheExec.EnableWord("AI_Fail_Log") = True And LCase(TheExec.DataManager.instanceName) Like "*lvcc*" Then
        If Vcc_min <> "-5555" Then
            Fail_index = 0
        
            For i = Step_NV To 0 Step -1
                    If Mid(Shmoo_result_PF, i + 1, 1) = "F" Then
                        Voltage_fail_collect(Fail_index) = CStr(RangeLow + RangeStepSize * (i))
                        Fail_index = Fail_index + 1
                        If Fail_index = 5 Then i = 0
                    End If
            Next i
        Else
        ''For shmoo hole collect 10 point fail cycle
            Fail_index = 0
        
            For i = Step_NV To 0 Step -1
                    If Mid(Shmoo_result_PF, i + 1, 1) = "F" Then
                        Voltage_fail_collect(Fail_index) = CStr(RangeLow + RangeStepSize * (i))
                        Fail_index = Fail_index + 1
                        If Fail_index = 10 Then i = 0
                    End If
            Next i
            
        End If
        Voltage_fail_point = Fail_index
    End If
      ''*******************************************AI***********************************************
'--------------------------------------------------------------------------------------------------------------
    ' Find HVCC: NV value to Hi value
    ' FFPPPpPPFF
    ' PPPPPpPPPP
    
    hvcc_point = 0
    
    If Not (InStr(LCase(test_name), "vih") > 0 Or InStr(LCase(test_name), "vil") > 0 Or InStr(LCase(x_pra), "vih") > 0 Or InStr(LCase(x_pra), "vil") > 0) Then
        
        For i = Step_NV To step_p Step 1
            If Mid(Shmoo_result_PF, i + 1, 1) = "F" Then
                Vcc_max = CStr(RangeLow + RangeStepSize * (i - 1))
                hvcc_point = i
                i = step_p
            End If
        Next i
        
        If Vcc_max = "" Then Vcc_max = CStr(RangeLow + RangeStepSize * Abs(step_Stop - step_Start))
    Else
        If InStr(LCase(test_name), "vil") > 0 Or InStr(LCase(x_pra), "vil") > 0 Then
            
            For i = 0 To (step_Stop - step_Start) / Step_x    'search low to high voltage
                char_pt = Mid(Shmoo_result_PF, i + 1, 1)
                If char_pt = "F" Then
                    Vcc_max = CStr(RangeLow + RangeStepSize * (i - 1))
                    hvcc_point = i
                    
                    i = (step_Stop - step_Start) / Step_x
                End If
            Next i
            If Vcc_max = "" Then Vcc_max = CStr(RangeLow + RangeStepSize * Abs(step_Stop - step_Start))
        End If
    End If
    
    'show_vcc = show_vcc & Vcc_max
        
    If hvcc_point <> 0 Then
        For i = hvcc_point + 1 To (step_Stop - step_Start) / Step_x Step 1
            If Mid(Shmoo_result_PF, i + 1, 1) = "P" Then
                Vcc_max = "5555"
            End If
            
        Next i
    End If
    
    ''*******************************************AI***********************************************
    If TheExec.EnableWord("AI_Fail_Log") = True And LCase(TheExec.DataManager.instanceName) Like "*hvcc*" Then
        If Vcc_min <> "5555" Then
            Fail_index = 0

            For i = Step_NV To step_p Step 1
                    If Mid(Shmoo_result_PF, i + 1, 1) = "F" Then
                        Voltage_fail_collect(Fail_index) = CStr(RangeLow + RangeStepSize * (i))
                        Fail_index = Fail_index + 1
                        If Fail_index = 5 Then i = step_p
                    End If
            Next i
        Else
        ''For shmoo hole collect 10 point fail cycle
            Fail_index = 0
        
            For i = Step_NV To step_p Step 1
                    If Mid(Shmoo_result_PF, i + 1, 1) = "F" Then
                        Voltage_fail_collect(Fail_index) = CStr(RangeLow + RangeStepSize * (i))
                        Fail_index = Fail_index + 1
                        If Fail_index = 10 Then i = step_p
                    End If
            Next i
            
        End If
        Voltage_fail_point = Fail_index
    End If
    ''*******************************************AI***********************************************
end_lvcc_hvcc:


    If Abs(Vcc_min) < 0.000000000001 Then Vcc_min = 0
    If Abs(Vcc_max) < 0.000000000001 Then Vcc_max = 0
    
    ''======170425 Char shmoo error code count start=====''
    
     For Each site In TheExec.sites
            total_shmoo_count = total_shmoo_count + 1
     Next site
    
     If F_shmoo_abnormal_counter = True Then

        For Each site In TheExec.sites
            If Trim(Vcc_max) = "5555" Or Trim(Vcc_min) = "-5555" Then
                shmoohole_count = shmoohole_count + 1
            End If

            If Trim(Vcc_max) = "9999" Or Trim(Vcc_min) = "-9999" Then
                shmooallfail_count = shmooallfail_count + 1
            End If

            If Trim(Vcc_max) = "7777" Or Trim(Vcc_min) = "-7777" Then
                shmooalarm_count = shmooalarm_count + 1
            End If

            included_shmoo_count = included_shmoo_count + 1

        Next site

      Else

        For Each site In TheExec.sites
            excluded_shmoo_count = excluded_shmoo_count + 1
        Next site

      End If
    ''======170425 Char shmoo error code count end=====''
    
    show_vcc = show_vcc & Vcc_min
    show_vcc = show_vcc & "," & Vcc_max
    show_vcc = show_vcc & "," & shmoo_form & "," & shmoo_stop & "," & shmoo_step & "," & CStr(RangeStepSize) & "]"
 
    TheExec.Datalog.WriteComment show_vcc
    'Call Print_power_condition
    'Debug.Print show_vcc
End Function

Public Function ShmooMakePseudoData(SetupName As String, step_Start As Long, step_Stop As Long, Step_x As Long)
    Dim pseudo_result_str(200) As String, i As Long, ch As String
    Dim cnt As Long
    cnt = 0
    
    
'    pseudo_result_str(cnt) = "+++++++++++++++": cnt = cnt + 1
'    pseudo_result_str(cnt) = "---------------": cnt = cnt + 1
'    pseudo_result_str(cnt) = "--------+++++++": cnt = cnt + 1
'    pseudo_result_str(cnt) = "+++++++--------": cnt = cnt + 1
'    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1
'    pseudo_result_str(cnt) = "---+-++++-++---": cnt = cnt + 1
    
    
    
    'Alg L2H
'    pseudo_result_str(cnt) = "+++++++++++++++": cnt = cnt + 1
'    pseudo_result_str(cnt) = "---------------": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++++++++---": cnt = cnt + 1
    pseudo_result_str(cnt) = "---++++++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++++++++--+": cnt = cnt + 1
    pseudo_result_str(cnt) = "+--++++++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "+--+++++++++--+": cnt = cnt + 1
    pseudo_result_str(cnt) = "-------++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++---------": cnt = cnt + 1
    pseudo_result_str(cnt) = "--------+++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "-----+++++-++++": cnt = cnt + 1 'HH
    pseudo_result_str(cnt) = "---+++++-++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---+++-++++++--": cnt = cnt + 1 'LH
    pseudo_result_str(cnt) = "-+++-+++++-----": cnt = cnt + 1 'LH
    pseudo_result_str(cnt) = "-+++-+++++-++--": cnt = cnt + 1 'BH
    pseudo_result_str(cnt) = "---++++-+++-+--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---++++--++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---+++--+++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "-++-+++-+++++--": cnt = cnt + 1
    'Alg H2L
    pseudo_result_str(cnt) = "+++++++++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "---------------": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++++++++---": cnt = cnt + 1
    pseudo_result_str(cnt) = "---++++++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "+++++++++------": cnt = cnt + 1
    pseudo_result_str(cnt) = "-------++++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "++++++---------": cnt = cnt + 1
    pseudo_result_str(cnt) = "--------+++++++": cnt = cnt + 1
    pseudo_result_str(cnt) = "-----+++++-++++": cnt = cnt + 1 'HH
    pseudo_result_str(cnt) = "---+++++-++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---+++-++++++--": cnt = cnt + 1 'LH
    pseudo_result_str(cnt) = "-+++-+++++-----": cnt = cnt + 1 'LH
    pseudo_result_str(cnt) = "-+++-+++++-++--": cnt = cnt + 1 'BH
    pseudo_result_str(cnt) = "---++++-+++-+--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---++++--++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "---+++--+++++--": cnt = cnt + 1
    pseudo_result_str(cnt) = "-++-+++-+++++--": cnt = cnt + 1
    
    'Misc
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_LVCConly
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_LVCConly_Over_NV
    pseudo_result_str(cnt) = "+++++++--------": cnt = cnt + 1     'CpuTd_VDD_CPU_HVCConly
    pseudo_result_str(cnt) = "+++++++--------": cnt = cnt + 1     'CpuTd_VDD_CPU_HVCConly_Under_NV
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_High_to_Low
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_Low_to_High
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_T_VDD_GPU_High_to_Low
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_T_VDD_GPU_Low_to_High
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_High_to_Low_CalcStepSize
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_Low_to_High_CalcStepSize
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_High_to_Low__StepSizeNotExact
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VDD_CPU_Low_to_High__StepSizeNotExact
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VIH_Pins_1p8v_IO
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VT_Pins_1p8v_IO
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VIH_SWD_TMS2
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_VT_SWD_TMS2
    pseudo_result_str(cnt) = "---+++++++++---": cnt = cnt + 1     'CpuTd_XI0_Freq_C


    With TheExec.DevChar.Results(SetupName).Shmoo

        For i = step_Start To step_Stop Step Step_x
            ch = Mid(pseudo_result_str(pseudo_result_index), i + 1, 1)
            Select Case ch
                Case "+":
                    .Points(i).ExecutionResult = tlDevCharResult_Pass
                Case "-":
                    .Points(i).ExecutionResult = tlDevCharResult_Fail
                Case "*": 'assume pass
                    .Points(i).ExecutionResult = tlDevCharResult_AssumedPass
                Case "~": 'assume fail
                    .Points(i).ExecutionResult = tlDevCharResult_AssumedFail
                Case Default:
                    .Points(i).ExecutionResult = tlDevCharResult_Error
            End Select
        Next i
    End With
    pseudo_result_index = pseudo_result_index + 1
End Function
Public Function CreateShmooResultString(Shmoo_Result, Shmoo_result_PF As String, SetupName As String, step_Start As Long, step_Stop As Long, Step_x As Long, Optional site As Variant)
    Dim i As Long
    Dim current_point_result As String
    Shmoo_Result = "": Shmoo_result_PF = ""
    Dim j As Long
    If Step_x > 0 Then
        j = 1
            Else
                j = Len(ShmResult(site))
    End If
    'Always from low value to hi value
    For i = step_Start To step_Stop Step Step_x
        current_point_result = TheExec.DevChar.Results(SetupName).Shmoo.Points(i).ExecutionResult
        Select Case current_point_result
            Case tlDevCharResult_Pass:
                    Shmoo_Result = Shmoo_Result & "+": Shmoo_result_PF = Shmoo_result_PF & "P"
            Case tlDevCharResult_Fail:
'                    If UCase(theexec.DataManager.InstanceName) Like "*CPUFUNC*" Or UCase(theexec.DataManager.InstanceName) Like "*SOCFUNC*" Then
                    If UCase(TheExec.DataManager.instanceName) Like "*CPUFUNC*" Or UCase(TheExec.DataManager.instanceName) Like "*SOCFUNC*" Or UCase(TheExec.DataManager.instanceName) Like "*RTOS*" Then
                        'If ShmResult(site) = "S" Then
                        If Mid(ShmResult(site), j, 1) = "S" Then
                            Shmoo_Result = Shmoo_Result & "S": Shmoo_result_PF = Shmoo_result_PF & "F"
                        ElseIf Mid(ShmResult(site), j, 1) = "B" Then
                            Shmoo_Result = Shmoo_Result & "B": Shmoo_result_PF = Shmoo_result_PF & "F"
                        ElseIf Mid(ShmResult(site), j, 1) = "C" Then
                            Shmoo_Result = Shmoo_Result & "C": Shmoo_result_PF = Shmoo_result_PF & "F"
                        ElseIf Mid(ShmResult(site), j, 1) = "-" Then
                            Shmoo_Result = Shmoo_Result & "-": Shmoo_result_PF = Shmoo_result_PF & "F"
                            TheExec.Datalog.WriteComment "RTOS_BCS bypassed due to pattern keyword issue."
                        End If
                        
                   If Step_x > 0 Then
                      j = j + 1
                        Else
                            j = j - 1
                    End If
                    
                    Else
                        Shmoo_Result = Shmoo_Result & "-": Shmoo_result_PF = Shmoo_result_PF & "F"
                    End If
            Case tlDevCharResult_NoTest:
                    Shmoo_Result = Shmoo_Result & "_": Shmoo_result_PF = Shmoo_result_PF & "_"
            Case tlDevCharResult_AssumedPass:
                    Shmoo_Result = Shmoo_Result & "*": Shmoo_result_PF = Shmoo_result_PF & "P"
            Case tlDevCharResult_AssumedFail:
                    Shmoo_Result = Shmoo_Result & "~":: Shmoo_result_PF = Shmoo_result_PF & "F"
            Case Else:
                    Shmoo_Result = Shmoo_Result & "?":: Shmoo_result_PF = Shmoo_result_PF & "?"
        End Select
    Next i
    ShmResult(site) = ""
End Function
Public Function Disable_Inst_pinname_in_PTR()

    TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
    TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = True
    TheExec.Datalog.ApplySetup

End Function
Public Function Enable_Inst_pinname_in_PTR()

    TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
    TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
    TheExec.Datalog.ApplySetup

End Function
Public Function Set_Level_Timing_Spec(Shmoo_Param_Type As String, Shmoo_Param_Name As String, shmoo_pin As String, Shmoo_TimeSets As String, Shmoo_value As Double, Port_name As String)
'Set instrument hardware
    Dim InstName As String
    Dim FRC_pin_name As String, Shmoo_Spec As String
'    If Shmoo_TimeSets <> "" Then
'        TheExec.ErrorLogMessage "Set up Timing set is not supported"
'        Exit Function
'    End If
    Select Case Shmoo_Param_Type
        Case "AC Spec", "DC Spec":
            TheExec.Overlays.ApplyUniformSpecToHW Shmoo_Param_Name, Shmoo_value
            Shmoo_Spec = Shmoo_Param_Name
        Case "Level":
        '20160925 Force to Ucase
            Select Case UCase(Shmoo_Param_Name)
                Case "VMAIN":
                    InstName = GetInstrument(shmoo_pin, 0)
                    Select Case InstName
                       Case "DC-07"
                            TheHdw.DCVI.Pins(shmoo_pin).Voltage = Shmoo_value
                       Case "VHDVS"
                            TheHdw.DCVS.Pins(shmoo_pin).Voltage.Main.Value = Shmoo_value
                       Case "HexVS"
                            TheHdw.DCVS.Pins(shmoo_pin).Voltage.Main.Value = Shmoo_value
                       Case Else
                    End Select
                Case "VT":
                   TheHdw.Digital.Pins(shmoo_pin).Levels.DriverMode = tlDriverModeVt
                   TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVt) = Shmoo_value
                Case "VIH": TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVih) = Shmoo_value
                Case "VIL": TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVil) = Shmoo_value
                Case "VOH": TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVoh) = Shmoo_value
                Case "VOL": TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVol) = Shmoo_value
                Case "VID": TheHdw.Digital.Pins(shmoo_pin).DifferentialLevels.Value(chVid) = Shmoo_value
                Case "VOD": TheHdw.Digital.Pins(shmoo_pin).DifferentialLevels.Value(chVod) = Shmoo_value
                Case "VICM":  TheHdw.Digital.Pins(shmoo_pin).DifferentialLevels.Value(chVicm) = Shmoo_value
                Case Else:
                    TheExec.ErrorLogMessage "Not supported Shmoo Parameter Name: " & Shmoo_Param_Name
            End Select
            Shmoo_Spec = shmoo_pin & "(" & Shmoo_Param_Name & ")"
        Case "Global Spec":
            If Port_name <> "" Then ' Shmoo pin with value from characterization loop and non-shmoo clock with AC context value
                Dim nWires_ary() As String
                Dim nwp As Variant, all_ports As String, all_pins As String
                Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
                nWires_ary = Split(nWire_Ports_GLB, ",")
                For Each nwp In nWires_ary
                    ' Convert nWires to all_ports and all_pins
                    Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
                    If Port_name Like nwp Then
                        Call VaryFreq(port_pa, Shmoo_value, ac_spec_pa)
                    Else
                        Call VaryFreq(port_pa, TheExec.specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
                    End If
'                    FreqMeasDebug pin_pa, 0.5, 0.01, 0.1             'Debug to print out freq in datalog
                Next nwp
            Else
                TheExec.Overlays.ApplyUniformSpecToHW Shmoo_Param_Name, Shmoo_value
                Shmoo_Spec = Shmoo_Param_Name
            End If
       '20180702 TER add for changeing "Edge"
        Case "Edge":
            Select Case UCase(Shmoo_Param_Name)
                Case "ON": TheHdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeD0) = Shmoo_value
                Case "DATA": TheHdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeD1) = Shmoo_value
                Case "RETURN": TheHdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeD2) = Shmoo_value
                Case "OFF": TheHdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeD3) = Shmoo_value
                Case "OPEN": TheHdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeR0) = Shmoo_value
                Case "CLOSE": TheHdw.Digital.Pins(shmoo_pin).Timing.EdgeTime(Shmoo_TimeSets, chEdgeR1) = Shmoo_value
                Case Else:
                    TheExec.ErrorLogMessage "Not supported To Set up Timing set Shmoo Parameter Name: " & Shmoo_Param_Name
                    Exit Function
            End Select
        Case Else:
            TheExec.ErrorLogMessage "Not supported Shmoo Parameter Name: " & Shmoo_Param_Type
    End Select
End Function
Public Function Shmoo_Set_Current_Point()
' Set up shmoo condition for current shmoo point (including tracking)
' Use Set_Level_Timing_Specto set hardware
    Dim Shmoo_Pin_Str As String
    Dim Shmoo_Tracking_Item As Variant, shmoo_axis As Variant
    Dim DevChar_Setup As String
    Dim Shmoo_Param_Type As String, Shmoo_Param_Name As String, shmoo_pin As String, Shmoo_value As Double, Port_name As String
    Dim Shmoo_Step_Name As String, Shmoo_TimeSets As String
    Dim arg_ary() As String
    Dim site As Variant
    If TheExec.DevChar.Setups.IsRunning = False Then
        Shmoo_End = False
    Else
        DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
       ' If Shmoo_End = True Then Exit Function  ' Prevent from setting  to last shmoo point; set Shmoo_End at the end of   PrintShmooInfo
        If TheExec.DevChar.Results(DevChar_Setup).startTime Like "1/1/0001*" Or TheExec.DevChar.Results(DevChar_Setup).startTime Like "0001/1/1*" Then Exit Function  ' initial run of shmoo, not the first point
        With TheExec.DevChar.Setups(DevChar_Setup).Shmoo
            For Each shmoo_axis In .Axes.List
                If LCase(.Axes(shmoo_axis).InterposeFunctions.PrePoint.Name) Like "freerunclk_set_xy" Then
                    arg_ary = Split(.Axes(shmoo_axis).InterposeFunctions.PrePoint.Arguments, ",")
                    Port_name = arg_ary(1)
                End If
                Shmoo_Param_Type = .Axes.Item(shmoo_axis).Parameter.Type
                Shmoo_Param_Name = .Axes.Item(shmoo_axis).Parameter.Name
                shmoo_pin = .Axes.Item(shmoo_axis).ApplyTo.Pins
                Shmoo_TimeSets = .Axes.Item(shmoo_axis).ApplyTo.Timesets
                For Each site In TheExec.sites
                    Shmoo_value = TheExec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).Value
                    'Debug.Print Shmoo_value
                    Set_Level_Timing_Spec Shmoo_Param_Type, Shmoo_Param_Name, shmoo_pin, Shmoo_TimeSets, Shmoo_value, Port_name
                Next site
                With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).TrackingParameters
                    For Each Shmoo_Tracking_Item In .List
                            Shmoo_Param_Type = .Item(Shmoo_Tracking_Item).Type
                            Shmoo_Param_Name = .Item(Shmoo_Tracking_Item).Name
                            shmoo_pin = .Item(Shmoo_Tracking_Item).ApplyTo.Pins
                            Shmoo_TimeSets = .Item(Shmoo_Tracking_Item).ApplyTo.Timesets
                            For Each site In TheExec.sites
                                Shmoo_value = TheExec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).TrackingParameters(Shmoo_Tracking_Item).Value
                                Set_Level_Timing_Spec Shmoo_Param_Type, Shmoo_Param_Name, shmoo_pin, Shmoo_TimeSets, Shmoo_value, Port_name
                            Next site
                    Next Shmoo_Tracking_Item
                End With
            Next shmoo_axis
        End With
    End If
End Function
Public Function Force_Flow_Shmoo_Condition()
    Dim X As Double, y As Double
    Dim X_axis As String, Y_axis As String
    Dim force_ary As AddIns
    
    Dim DevChar_Setup As String
    Dim shmoo_axis As Variant, Shmoo_Tracking_Item As Variant
    Dim axis_name As Variant, shmoo_val As Double, Shmoo_type As Double, Shmoo_Name As Double, Shmoo_Spec As String
    Dim i As Long, Shmoo_start As Double, shmoo_stop As Double, Shmoo_StepSize As Double, Shmoo_Current_Step As Long, shmoo_step As Long
    Dim X_pt As Double, Y_pt As Double
    Dim Port_name As String
    Dim shmoo_pin As String, Shmoo_TimeSet As String
    Dim arg_ary() As String, axis_type As String
    Dim site As Variant
    Flow_Shmoo_Axis_Count = 0
    Flow_Shmoo_Force_Condition = ""
    Shmoo_setup_str = ""
    Flow_Shmoo_Port_Name = ""
    For Each site In TheExec.sites
        DevChar_Setup = TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup")
        X_pt = TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_X")
        Y_pt = TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_Y")
        Exit For
    Next site
    If DevChar_Setup <> "" Then
        With TheExec.DevChar.Setups(DevChar_Setup).Shmoo
            Shmoo_Tracking_Item = -99
            For Each shmoo_axis In .Axes.List
                If (Flow_Shmoo_X_Last_Value <> X_pt _
                    Or Flow_Shmoo_X_Last_Value <> -99) _
                    And (Flow_Shmoo_Y_Last_Value <> Y_pt _
                    Or Flow_Shmoo_Y_Last_Value <> -99) Then
                    Flow_Shmoo_X_Fast = True
                End If
                If LCase(.Axes(shmoo_axis).InterposeFunctions.PrePoint.Name) Like "freerunclk_set_xy" Then
                    arg_ary = Split(.Axes(shmoo_axis).InterposeFunctions.PrePoint.Arguments, ",")
                    Port_name = arg_ary(1)
                    Flow_Shmoo_Port_Name = Port_name
                End If
                Select Case shmoo_axis
                    Case tlDevCharShmooAxis_X:
                        axis_type = "X"
                        If Flow_Shmoo_X_Current_Step <= Flow_Shmoo_X_Step _
                            And (Flow_Shmoo_X_Last_Value <> X_pt _
                            Or Flow_Shmoo_X_Last_Value = -99) Then
                            If Flow_Shmoo_X_Last_Value <> X_pt Then
                                If Flow_Shmoo_X_Current_Step = Flow_Shmoo_X_Step Then
                                    Flow_Shmoo_X_Current_Step = 0
                                Else
                                    Flow_Shmoo_X_Current_Step = Flow_Shmoo_X_Current_Step + 1
                                End If
                            End If
                        End If
                        Shmoo_Current_Step = Flow_Shmoo_X_Current_Step
                        shmoo_step = Flow_Shmoo_X_Step
                    Case tlDevCharShmooAxis_Y:
                        axis_type = "Y"
                        If Flow_Shmoo_Y_Current_Step < Flow_Shmoo_Y_Step _
                            And (Flow_Shmoo_Y_Last_Value <> Y_pt _
                            Or Flow_Shmoo_Y_Last_Value = -99) Then
                            If Flow_Shmoo_X_Fast = True Then
                                If Flow_Shmoo_Y_Last_Value <> Y_pt _
                                And Flow_Shmoo_X_Current_Step = 0 Then
                                    Flow_Shmoo_Y_Current_Step = Flow_Shmoo_Y_Current_Step + 1
                                End If
                            Else
                                Flow_Shmoo_Y_Current_Step = Flow_Shmoo_Y_Current_Step + 1
                            End If
                        End If
                        Shmoo_Current_Step = Flow_Shmoo_Y_Current_Step
                        shmoo_step = Flow_Shmoo_Y_Step
                End Select
            
                If .Axes(shmoo_axis).Parameter.range.StepSize <> Empty Then
                    Shmoo_StepSize = .Axes(shmoo_axis).Parameter.range.StepSize
                Else
                    Shmoo_StepSize = (.Axes(shmoo_axis).Parameter.range.To - .Axes(shmoo_axis).Parameter.range.from) / .Axes(shmoo_axis).Parameter.range.Steps
                End If
                shmoo_val = .Axes(shmoo_axis).Parameter.range.from + Shmoo_Current_Step * .Axes(shmoo_axis).Parameter.range.StepSize
'                Flow_Shmoo_Setup_Instrument DevChar_Setup, CLng(shmoo_axis), "", Shmoo_Current_Step, Shmoo_StepSize
                Set_Level_Timing_Spec .Axes(shmoo_axis).Parameter.Type.Value, .Axes(shmoo_axis).Parameter.Name.Value, .Axes(shmoo_axis).ApplyTo.Pins, .Axes(shmoo_axis).ApplyTo.Timesets, shmoo_val, Port_name
                If .Axes(shmoo_axis).ApplyTo.Pins <> "" Then
                    Shmoo_Spec = .Axes(shmoo_axis).Parameter.Name.Value & "(" & .Axes(shmoo_axis).ApplyTo.Pins & ")"
                Else
                    Shmoo_Spec = .Axes(shmoo_axis).Parameter.Name.Value
                End If
                If .Axes(shmoo_axis).ApplyTo.Timesets <> "" Then
                    Shmoo_Spec = .Axes(shmoo_axis).Parameter.Name.Value & "(" & .Axes(shmoo_axis).ApplyTo.Timesets & ")"
                Else
                    Shmoo_Spec = .Axes(shmoo_axis).Parameter.Name.Value
                End If
                If Shmoo_setup_str = "" Then
                    Shmoo_setup_str = axis_type & ":" & Shmoo_Spec & "=" & shmoo_val & "; "
                Else
                    Shmoo_setup_str = Shmoo_setup_str & axis_type & ":" & Shmoo_Spec & "=" & shmoo_val & "; "
                End If
                For Each Shmoo_Tracking_Item In .Axes(shmoo_axis).TrackingParameters.List
                    shmoo_pin = .Axes(shmoo_axis).ApplyTo.Pins
                    Shmoo_TimeSet = .Axes(shmoo_axis).ApplyTo.Timesets
                    With .Axes.Item(Shmoo_Tracking_Item).Parameter
                        Shmoo_StepSize = (.range.To - .range.from) / shmoo_step
                        shmoo_val = .range.from + Shmoo_Current_Step * Shmoo_StepSize
                        Set_Level_Timing_Spec .Type.Value, .Name.Value, shmoo_pin, Shmoo_TimeSet, shmoo_val, Port_name
                    End With
'                    Flow_Shmoo_Setup_Instrument DevChar_Setup, CLng(shmoo_axis), CStr(Shmoo_Tracking_Item), Shmoo_Current_Step, Shmoo_StepSize
                    If shmoo_pin <> "" Then
                        Shmoo_Spec = .Axes(shmoo_axis).Parameter.Name.Value & "(" & shmoo_pin & ")"
                    Else
                        Shmoo_Spec = .Axes(shmoo_axis).Parameter.Name.Value
                    End If
                    If Shmoo_TimeSet <> "" Then
                        Shmoo_Spec = .Axes(shmoo_axis).Parameter.Name.Value & "(" & Shmoo_TimeSet & ")"
                    End If
                    Shmoo_setup_str = axis_type & ":" & Shmoo_Spec & "=" & shmoo_val & "; "
                Next Shmoo_Tracking_Item
            Next shmoo_axis
        End With
        Flow_Shmoo_X_Last_Value = X_pt
        Flow_Shmoo_Y_Last_Value = Y_pt
        FlowShmooString_GLB = CStr(shmoo_val * 1000)
        TheExec.Datalog.WriteComment "*********** Shmoo Point   " & Shmoo_setup_str & "    ***********"
    End If
    
End Function
Public Function Get_Axis_Type(shmoo_axis As Long) As String
    Select Case shmoo_axis
        Case tlDevCharShmooAxis_X:
            Get_Axis_Type = "X"
        Case tlDevCharShmooAxis_Y:
            Get_Axis_Type = "Y"
    End Select
End Function

Public Function Restart_All_Freerun_Clk()
    Dim site As Variant
    For Each site In TheExec.sites
        Exit For
    Next
    
'''''''''''''''''Obsolete due to Support multiple nWire port 20170503'''''''''''''
'    'Stop Clock
'    If XI0_GP <> "" Then
'        Call FreeRunClk_Disable("XI0_Port")
'    ElseIf XI0_Diff_GP <> "" Then
'        Call FreeRunClk_Disable("XI0_Diff_Port")
'    End If
'    If RTCLK_GP <> "" Then
'        Call FreeRunClk_Disable("RT_CLK32768_Port")
'    ElseIf RTCLK_Diff_GP <> "" Then
'        Call FreeRunClk_Disable("RT_CLK32768_Diff_Port")
'    End If
'    'Start Clock
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
    
    
    '''''''''''''''''Support multiple nWire port 20170503'''''''''''''
    Dim nWire_port_ary() As String
    Dim nwp As Variant ', all_ports As String, all_pins As String
    Dim port_pa As String, ac_spec_pa As String, pin_pa As String, global_spec_pa As String
'    If nWire_ports = "" Then nWire_ports = nWire_Ports_GLB
    nWire_port_ary = Split(nWire_Ports_GLB, ",")
    ' Convert nWire_ports to all_ports and all_pins
    For Each nwp In nWire_port_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        'Stop Clock
        Call FreeRunClk_Disable(port_pa)
    Next nwp
    
    For Each nwp In nWire_port_ary
        Get_nWire_Name nwp, port_pa, ac_spec_pa, pin_pa, global_spec_pa
        'Start Clock
        Call VaryFreq(port_pa, TheExec.specs.AC(ac_spec_pa).ContextValue, ac_spec_pa)
    Next nwp
      
End Function





Public Function ReStart_FRC(ports As String)
    Call Enable_FRC(ports)
''    Dim port_ary() As String, Freq_ary() As String, i As Long
''    port_ary = Split(ports, ",")
''    For i = 0 To UBound(port_ary)
''        If LCase(port_ary(i)) Like "xi*" Then
''            If XI0_GP <> "" Then
''                Call VaryFreq("XI0_Port", TheExec.Specs.AC("XI0_Freq_VAR").ContextValue, "XI0_Freq_VAR")
''            ElseIf XI0_Diff_GP <> "" Then
''                Call VaryFreq("XI0_Diff_Port", TheExec.Specs.AC("XI0_Diff_Freq_VAR").ContextValue, "XI0_Diff_Freq_VAR")
''            End If
''        End If
''        If LCase(port_ary(i)) Like "rt*" Then
''            If RTCLK_GP <> "" Then
''                Call VaryFreq("RT_CLK32768_Port", TheExec.Specs.AC("RT_CLK32768_Freq_VAR").ContextValue, "RT_CLK32768_Freq_VAR")
''            ElseIf RTCLK_Diff_GP <> "" Then
''                Call VaryFreq("RT_CLK32768_Diff_Port", TheExec.Specs.AC("RT_CLK32768_Diff_Freq_VAR").ContextValue, "RT_CLK32768_Diff_Freq_VAR")
''            End If
''        End If
''    Next i
End Function

Public Function FreqMeasDebug(FreqMeasPins As String, V_threshold As Double, t_interval As Double, t_wait As Double)
    Dim MeasFreq As New PinListData, i As Long
    Dim PinMeas As New PinList
    PinMeas = FreqMeasPins
    If TheExec.DataManager.PinType(FreqMeasPins) Like "Differential" Then
        TheHdw.Digital.Pins(FreqMeasPins).DifferentialLevels.Value(chVod) = V_threshold
    Else
        TheHdw.Digital.Pins(FreqMeasPins).Levels.Value(chVoh) = V_threshold
    End If
    Call Freq_MeasFreqSetup(PinMeas, t_interval, VOH)
    Call HardIP_Freq_MeasFreqStart(PinMeas, t_interval, MeasFreq, CStr(t_wait))
    If TheExec.DataManager.PinType(FreqMeasPins) Like "Differential" Then
        For i = 0 To MeasFreq.Pins.Count - 1 Step 2
            TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(i), Tname:="Debug"
        Next i
    Else
        TheExec.Flow.TestLimit resultVal:=MeasFreq, Tname:="Debug"
    End If
End Function
Public Function PostPoint_test_faillog_2D(argc As Long, argv() As String) 'pclinzg plot faillog for 2D shmoo
    Dim SetupName As String
    Dim StepNamex As String
    Dim StepNamey As String
    Dim Value As Double
    Dim freq As Double
    Dim site As Variant
    
    
    If TheExec.EnableWord("AI_Fail") = True Then
      For Each site In TheExec.sites
        SetupName = TheExec.DevChar.Setups.ActiveSetupName
        StepNamex = TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_X).StepName
        Value = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
        TheExec.Datalog.WriteComment StepNamex & ":" & Format(Value, "0.000")
        
        StepNamey = TheExec.DevChar.Setups(SetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).StepName
        If StepNamey <> "" Then
           freq = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
           TheExec.Datalog.WriteComment StepNamey & ":" & Format(freq / 1000000, "0.000") & "Mhz"
       End If
        

        
        HardIP_WriteFuncResult
      Next site
    End If
    
End Function

Public Function Getforcecondition_VDD(VDD_Force As String, Interpose_PrePat As String)
Dim force_condition_arr() As String
Dim pin_array() As String
Dim fc As Variant
VDD_Force = ""
force_condition_arr = Split(Interpose_PrePat, ";")

For Each fc In force_condition_arr

    pin_array = Split(fc, ":")
    
    If UBound(pin_array) = 2 Then
        If UCase(pin_array(1)) = "V" Then
            If VDD_Force = "" Then
                VDD_Force = pin_array(0)
            Else
                VDD_Force = VDD_Force & "," & pin_array(0)
            End If
        End If
    End If
Next fc

End Function
Public Function Get_Shmoo_Set_Pin(Shmoo_Apply_Pin As String, VDD_Force As String, pin_count As Long)
          
            Dim tmp_Shmoo_Apply_Pin() As String
            Dim pin_list_arry() As String
            Dim Flag_IO As Boolean, Flag_VDD As Boolean
            Dim i As Long
         
           
            If TheExec.DevChar.Setups.IsRunning = True Then
    
                Get_Current_Apply_Pin Shmoo_Apply_Pin
                Call TheExec.DataManager.DecomposePinList(Shmoo_Apply_Pin, pin_list_arry, pin_count)
                
                Flag_IO = False
                Flag_VDD = False
                
                For i = 0 To pin_count - 1
                    If UCase(TheExec.DataManager.PinType(pin_list_arry(i))) = "I/O" Then Flag_IO = True
                    If UCase(TheExec.DataManager.PinType(pin_list_arry(i))) = "POWER" Then Flag_VDD = True
                Next i
                If Flag_IO = True And Flag_VDD = True Then TheExec.ErrorLogMessage "Can not  contain both I/O and Power Pin  for  Shmoo apply pin " & Shmoo_Apply_Pin
                
                If Flag_IO = True Then
                   If g_Vbump_function = True Then
                      Shmoo_Apply_Pin = ""
                   Else
                      If VDD_Force = "" Then
                         Shmoo_Apply_Pin = "CorePower"
                      Else
                         Shmoo_Apply_Pin = "CorePower, " & VDD_Force
                      End If
                   End If
                ElseIf Flag_VDD = True Then
                    If g_Vbump_function = True Then
                       Shmoo_Apply_Pin = Shmoo_Apply_Pin
                    Else
                       If VDD_Force = "" Then
                          Shmoo_Apply_Pin = Shmoo_Apply_Pin & ",CorePower"
                       Else
                          Shmoo_Apply_Pin = Shmoo_Apply_Pin & ",CorePower, " & VDD_Force
                       End If
                    End If
                End If
                
               ''SY mask
''                tmp_Shmoo_Apply_Pin = Split(Shmoo_Apply_Pin, ",")
''                If UBound(tmp_Shmoo_Apply_Pin) < 0 Then
''                    Power_Run_Scenario = "init_NV_pl_NV"   'if pin  are IO pins,  Power_Run_Scenario is not used
''                End If
                
            Else
                If g_Vbump_function = True Then
                   Shmoo_Apply_Pin = ""
                Else
                   If VDD_Force = "" Then
                      Shmoo_Apply_Pin = "CorePower"
                   Else
                      Shmoo_Apply_Pin = "CorePower, " & VDD_Force
                   End If
                End If
            End If
            
            Call TheExec.DataManager.DecomposePinList(Shmoo_Apply_Pin, pin_list_arry, pin_count)
End Function
Public Function GetSrcString_fromEMAArray(Pat As String, TestCase As String, ByRef SrcBinStr As String, ByRef SrcBits As Double, Optional A As Variant)
Dim i As Integer, j As Integer
Dim FindTestCase As Boolean
Dim FindPattern As Boolean

    For i = 0 To UBound(SrcStock)
        'Compare input pattern and test case
        FindPattern = False
        If UCase(Pat) Like UCase(SrcStock(i).PatternName & "*") Then 'Pat is match up with control table
            For j = 0 To UBound(SrcStock(i).TestCase)
                FindTestCase = False
                'compare test case
                If UCase(TestCase) = UCase(SrcStock(i).TestCase(j).ConditionName) Then
                    SrcBinStr = SrcStock(i).TestCase(j).DigSrc_BinStr
                    SrcBits = SrcStock(i).TestCase(j).DigSrc_BitCount
                    FindTestCase = True
                    FindPattern = True
                    Exit For
                End If
            Next j
            FindPattern = True
            Exit For
        End If
    Next i
    'error message
    If FindPattern = False Then
        TheExec.Datalog.WriteComment "Can NOT find Pattern in Control Table" & vbCrLf
    Else
        If FindTestCase = False Then
            TheExec.Datalog.WriteComment "Can NOT find TestCase in control Table" & vbCrLf
        End If
    End If
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + "GetSrcString_fromEMAArray" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Parse_EMA_DigSrcInfo()
    Dim ws As Worksheet
    Dim maxcolumn As Double
    Dim MaxRow As Double
    Dim CurColumn As Double
    Dim CurRow As Double
    Dim TempStr As String
    Dim TempCount As Double
    Dim ExistPatCount As Integer
    Dim ExistTestCount As Integer
    Dim CurPatNum As Integer
Dim CurTestNum As Integer

On Error GoTo errHandler

'Worksheets("ACCP Char").Select
    If DSSCMappingTableIsRead = False Then
        Set ws = Sheets("CZ_DSSCMappingTable_CP1")
        'MaxRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        maxcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' define array depth of patterns
        ExistPatCount = 0
        For CurColumn = 1 To maxcolumn
            If ws.Cells(1, CurColumn) Like "Pattern Name" Then
                ExistPatCount = ExistPatCount + 1
                ReDim SrcStock(ExistPatCount - 1) As DynamicSrc
            End If
        Next CurColumn
        
        ExistPatCount = 0
        For CurColumn = 1 To maxcolumn
            If UCase(ws.Cells(1, CurColumn)) Like UCase("Pattern Name") Then
                If ExistPatCount > 0 Then
                    ReDim Preserve SrcStock(ExistPatCount - 1).TestCase(CurTestNum - 1) As testCondition
                End If
                ExistPatCount = ExistPatCount + 1
                CurTestNum = 0
        
            ElseIf UCase(ws.Cells(1, CurColumn)) Like UCase("Test*") Then
                CurTestNum = CurTestNum + 1
                If CurColumn = maxcolumn Then
                    ReDim Preserve SrcStock(ExistPatCount - 1).TestCase(CurTestNum - 1) As testCondition
                End If
            End If
        Next CurColumn
        
        CurPatNum = 0
        CurTestNum = 0
        For CurColumn = 1 To maxcolumn
            If UCase(ws.Cells(1, CurColumn)) Like UCase("Pattern Name") Then
                SrcStock(CurPatNum).PatternName = ws.Cells(2, CurColumn)
                CurPatNum = CurPatNum + 1
                CurTestNum = 0
            ElseIf UCase(ws.Cells(1, CurColumn)) Like UCase("Test*") Then
                SrcStock(CurPatNum - 1).TestCase(CurTestNum).ConditionName = ws.Cells(1, CurColumn)
                MaxRow = ws.Cells(Rows.Count, CurColumn).End(xlUp).row
                TempStr = ""
                TempCount = 0
                For CurRow = 2 To MaxRow
                    TempStr = TempStr & ws.Cells(CurRow, CurColumn)
                    TempCount = TempCount + 1
                Next CurRow
                SrcStock(CurPatNum - 1).TestCase(CurTestNum).DigSrc_BinStr = TempStr
                SrcStock(CurPatNum - 1).TestCase(CurTestNum).DigSrc_BitCount = TempCount
                If TempCount <> Len(TempStr) Then
                    TheExec.Datalog.WriteComment "Source Bit is NOT single bit" & vbCrLf
                    TheExec.Datalog.WriteComment "PatternName :" & SrcStock(CurPatNum - 1).PatternName & vbCrLf
                    TheExec.Datalog.WriteComment "TestCase :" & SrcStock(CurPatNum - 1).TestCase(CurTestNum).ConditionName & vbCrLf
                End If
                'SrcStock.TestCase.disrc_bitcount
                CurTestNum = CurTestNum + 1
            End If
        Next CurColumn
        Dim i As Integer
        Dim j As Integer
        
        Dim tempStr_1 As String
        Dim tempStr_2 As String
        For i = 0 To UBound(SrcStock())
            tempStr_1 = SrcStock(i).PatternName
            For j = i + 1 To UBound(SrcStock())
                tempStr_2 = SrcStock(j).PatternName
                If UCase(tempStr_1) = UCase(tempStr_2) Then
                    TheExec.Datalog.WriteComment "There are two same patterns in control table" & vbCrLf
                    TheExec.Datalog.WriteComment "Pattern Name :" & tempStr_2
                    TheExec.Datalog.WriteComment "Pattern# :" & i + 1 & "," & j + 1
                    GoTo errHandler
                End If
            Next j
        Next i
    
        'debug
    '    Dim Out_SrcBinStr As String
    '    Dim Out_SrcBits As Double
        'Call GetSrcString_fromEMAArray("DD_CYPA0_AAAAAA", "Test3", Out_SrcBinStr, Out_SrcBits)
        DSSCMappingTableIsRead = True
    End If
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + "Parse_EMA_DigSrcInfo" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function Shmoo_Set_Retention_Power(Optional Skip_Sweep_Pin As Boolean = False)
    'Modify for force condition "VRET" 20171213
    Dim i As Long
    Dim rn_ary() As String, rn_ary_fv() As String
    Dim Pin_Ary() As String, p_cnt As Long
    Dim skip_pin As String
    Dim Skip_Pin_Dic  As New Dictionary
                    
    If Skip_Sweep_Pin = True Then
        If TheExec.DevChar.Setups.IsRunning = True Then Get_Current_Apply_Pin skip_pin
        If skip_pin <> "" Then Create_Pin_Dic skip_pin, Skip_Pin_Dic
    End If
    
    If g_Retention_VDD <> "" Then
        rn_ary = Split(LCase(g_Retention_VDD), ",")
        rn_ary_fv = Split(g_Retention_ForceV, ",")
        For i = 0 To UBound(rn_ary)
            If Skip_Sweep_Pin = True Then ' Do not set retention power for shmoo pin
                If Skip_Pin_Dic.Exists(rn_ary(i)) = False Then TheHdw.DCVS.Pins(rn_ary(i)).Voltage.Main.Value = CDbl(rn_ary_fv(i))
            Else
                TheHdw.DCVS.Pins(rn_ary(i)).Voltage.Main.Value = CDbl(rn_ary_fv(i))
            End If
        Next i
'    Else
'        rn_ary = Split(LCase(skip_pin), ",")
'        If skip_pin <> "" Then Shmoo_Restore_Power_per_site skip_pin, ShmooSweepPower, "*** PL" & seq & "_" & pl_level & " Force***", ""
    End If
End Function
Public Function Decide_retetntion_power(Retention_V() As SiteDouble, RetPins As PinList)
    'Modify for force condition "VRET" 20171213
    ' Assume that  shmoo pin must be included in g_Retention_VDD
    Dim i As Long
    Dim rn_ary() As String, rn_ary_fv() As String, rn_cnt As Long
    Dim Pin_Ary() As String, p_cnt As Long
    Dim shmoo_pin As String
    Dim Shmoo_pin_Dic  As New Dictionary
    Dim ShmooSweepPower_Dic  As New Dictionary
                    
    If TheExec.DevChar.Setups.IsRunning = True Then
        Get_Current_Apply_Pin shmoo_pin
        If g_Retention_VDD <> "" Then
            RetPins.Value = g_Retention_VDD
        Else
            RetPins.Value = shmoo_pin
        End If
    Else
        If g_Retention_VDD <> "" Then
            RetPins.Value = g_Retention_VDD
        Else
            RetPins.Value = ""
        End If
    End If
    
    If g_Retention_VDD <> "" Then
        rn_ary = Split(LCase(g_Retention_VDD), ",")
        rn_ary_fv = Split(g_Retention_ForceV, ",")
        If TheExec.DevChar.Setups.IsRunning = True Then
            Create_Pin_Dic shmoo_pin, Shmoo_pin_Dic
        End If
        For Each site In TheExec.sites
             ShmooSweepPower_Dic.RemoveAll
             TheExec.DataManager.DecomposePinList Shmoo_Apply_Pin, Pin_Ary, p_cnt
             For i = 0 To UBound(Pin_Ary)
                 ShmooSweepPower_Dic.Add LCase(Pin_Ary(i)), ShmooSweepPower(i)
             Next i
                          
             For i = 0 To UBound(rn_ary)
                 If ShmooSweepPower_Dic.Exists(rn_ary(i)) = True And Shmoo_pin_Dic.Exists(rn_ary(i)) = True Then
                     Retention_V(i) = ShmooSweepPower_Dic.Item(rn_ary(i))
                 Else
                     Retention_V(i)(site) = CDbl(rn_ary_fv(i))
                 End If
             Next i
        Next site
    Else
         If TheExec.DevChar.Setups.IsRunning = True Then
            For Each site In TheExec.sites
                ShmooSweepPower_Dic.RemoveAll
                TheExec.DataManager.DecomposePinList Shmoo_Apply_Pin, Pin_Ary, p_cnt
                For i = 0 To UBound(Pin_Ary)
                    ShmooSweepPower_Dic.Add LCase(Pin_Ary(i)), ShmooSweepPower(i)
                Next i
                
                TheExec.DataManager.DecomposePinList shmoo_pin, rn_ary, rn_cnt
                For i = 0 To UBound(rn_ary)
                        Retention_V(i) = ShmooSweepPower_Dic.Item(LCase(rn_ary(i)))
                Next i
            Next site
        End If
    End If
End Function



Public Function Parse_SELSRM_Mapping_Table()
Dim i As Long, j As Long
Dim maxcolumn As Long
Dim MaxRow As Long
Dim CurColumn As Long
Dim ws As Worksheet
Dim blockName As String
Dim Block_Index As Integer

Dim Block_Arr() As String
Dim Block_Arr_starting_inx() As Integer
Dim Block_Patt_Index() As Integer
Dim Block_Patt_End_Index() As Integer
Dim Block_Flag() As Boolean
Dim Block_Bits_Index() As Integer
Dim Block_Bits_End_Index() As Integer

Dim SOC_Patt_End_index As Integer
Dim SOC_Patt_End_flag As Boolean
Dim SOC_Bits_End_index As Integer
Dim SOC_Block_index As Integer
Dim SOC_Block_flag As Boolean
Dim CPU_Patt_End_index As Integer
Dim CPU_Patt_End_flag As Boolean
Dim CPU_Bits_End_index As Integer
Dim CPU_Block_index As Integer
Dim CPU_Block_flag As Boolean
Dim GPU_Patt_End_index As Integer
Dim GPU_Patt_End_flag As Boolean
Dim GPU_Bits_End_index As Integer
Dim GPU_Block_index As Integer
Dim GPU_Block_flag As Boolean
Dim RTOS_Patt_End_index As Integer
Dim RTOS_Patt_End_flag As Boolean
Dim RTOS_Bits_End_index As Integer
Dim RTOS_Block_index As Integer
Dim RTOS_Block_flag As Boolean
Dim Table_2D_arr() As String
Dim ii As Integer

Dim Int_Val As Integer
Dim tmp_Val As Integer
Int_Val = 0
tmp_Val = 0

On Error GoTo errHandler
Set ws = Sheets("SELSRM_Mapping_Table")

'    Dim GetSelSram As mapping_table
Worksheets("SELSRM_Mapping_Table").Select
'    ReDim GetSelSram.Block(3) ' temperary only 4 block, soc,cpu,gpu and rtos

maxcolumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
MaxRow = ws.Cells(Rows.Count, 1).End(xlUp).row
ReDim Table_2D_arr(MaxRow, maxcolumn)
ReDim Block_Arr(MaxRow - 2)

For i = 1 To maxcolumn ' get block name, Pattern and Bit index
    For j = 2 To MaxRow
        Table_2D_arr(j, i) = ws.Cells(j, i).Value
    Next j
Next i

For i = 1 To 3
    For j = 2 To MaxRow
        If i = 1 Then ' get block index
            If LCase(Table_2D_arr(j, i)) = "end" Then
                Block_Arr(j - 2) = LCase(Table_2D_arr(j, i)) 'Block_Arr(j - 2) is for initializing starting point "(0)"
                ReDim Block_Patt_Index(Block_Index - 1)
                ReDim Block_Patt_End_Index(Block_Index - 1)
                ReDim Block_Bits_Index(Block_Index - 1)
                ReDim Block_Bits_End_Index(Block_Index - 1)
                Exit For
            Else
                Block_Arr(j - 2) = LCase(Table_2D_arr(j, i)) 'Block_Arr(j - 2) is for initializing starting point "(0)"
                If LCase(Table_2D_arr(j, i)) <> LCase(Table_2D_arr(j + 1, i)) Then
                    Block_Index = Block_Index + 1
                End If
            End If
        ElseIf i = 2 Then ' get patt index
        
''                For ii = 0 To Block_Index - 1
''                    If LCase(Table_2D_arr(j, i)) <> "" Then
''                        If Block_Arr(j - 2) = Block_Arr(j - 2 + 1) Then
''                            Block_Patt_Index(ii) = Block_Patt_Index(ii) + 1
''                            Exit For
''                        Else
''                            Block_Patt_Flag(ii) = True
''                            Exit For
''                        End If
''                    End If
''                Next ii
        End If
    Next j
Next i




''===================================================================================
''In order to get each block / starting index
ReDim Block_Arr_starting_inx(Block_Index - 1)

ii = 0
For Block_Index = 0 To Block_Index - 1
    If Block_Index = 0 Then
        Block_Arr_starting_inx(Block_Index) = 2 ' "+2" is for align with table
    Else
        For ii = ii To UBound(Block_Arr) - 1
            If Block_Arr(ii) <> Block_Arr(ii + 1) Then
                Block_Arr_starting_inx(Block_Index) = ii + 3
                Exit For
            End If
        Next ii
    End If
    ii = ii + 1
Next Block_Index


ii = 0
For Block_Index = 0 To Block_Index - 1
    For ii = ii To UBound(Block_Arr) - 1
        If Block_Arr(ii) <> Block_Arr(ii + 1) Then
            Block_Patt_Index(Block_Index) = ii + 2 ' "+2" is for align with table
            Exit For
        End If
    Next ii
    ii = ii + 1
Next Block_Index
''===================================================================================
''In order to get pattern/bits end index
ii = 0
For Block_Index = 0 To Block_Index - 1
    For ii = ii To MaxRow
        If LCase(Table_2D_arr(ii + 2, 2)) <> "" Then ' patt
            Block_Patt_End_Index(Block_Index) = Block_Patt_End_Index(Block_Index) + 1
        End If
        If LCase(Table_2D_arr(ii + 2, 3)) <> "" Then ' Bits
            Block_Bits_End_Index(Block_Index) = Block_Bits_End_Index(Block_Index) + 1
        End If
        If ii + 2 = Block_Patt_Index(Block_Index) Then Exit For
    Next ii
    ii = ii + 1
Next Block_Index
''===================================================================================

ii = 0
ReDim GetSelSram.Block(Block_Index - 1)
ReDim Block_Flag(Block_Index - 1)
For ii = 0 To Block_Index - 1
'        For i = 0 To UBound(Block_Patt_Index)
        For j = 2 To MaxRow
            If Block_Flag(ii) = False Then
                If j = Block_Arr_starting_inx(ii) Then
                    GetSelSram.Block(ii).DomainName = Table_2D_arr(j, 1)
                    If Block_Patt_End_Index(ii) = 0 Then ' no pattern case
                        ReDim GetSelSram.Block(ii).Pattern(Block_Patt_End_Index(ii))
                    Else
                        ReDim GetSelSram.Block(ii).Pattern(Block_Patt_End_Index(ii) - 1)
                        For Block_Patt_End_Index(ii) = 0 To Block_Patt_End_Index(ii) - 1
                            GetSelSram.Block(ii).Pattern(Block_Patt_End_Index(ii)) = Table_2D_arr(j + Block_Patt_End_Index(ii), 1 + 1)
                        Next Block_Patt_End_Index(ii)
                    End If
                    
                    ReDim GetSelSram.Block(ii).DomainBits(Block_Bits_End_Index(ii) - 1) ' bits number must no empty
                    For Block_Bits_End_Index(ii) = 0 To Block_Bits_End_Index(ii) - 1
                        GetSelSram.Block(ii).DomainBits(Block_Bits_End_Index(ii)).BITS = Table_2D_arr(j + Block_Bits_End_Index(ii), 1 + 2)
                        GetSelSram.Block(ii).DomainBits(Block_Bits_End_Index(ii)).logicPin = Table_2D_arr(j + Block_Bits_End_Index(ii), 1 + 3)
                        GetSelSram.Block(ii).DomainBits(Block_Bits_End_Index(ii)).SramPin = Table_2D_arr(j + Block_Bits_End_Index(ii), 1 + 4)
                        GetSelSram.Block(ii).DomainBits(Block_Bits_End_Index(ii)).SelSram1 = Table_2D_arr(j + Block_Bits_End_Index(ii), 1 + 5)
                        GetSelSram.Block(ii).DomainBits(Block_Bits_End_Index(ii)).SelSram0 = Table_2D_arr(j + Block_Bits_End_Index(ii), 1 + 6)
                    Next Block_Bits_End_Index(ii)
                    Block_Flag(ii) = True
                End If
            End If
        Next j
'        Next i
Next ii

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + "Parse_SELSRM_Mapping_Table" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Shmoo_Save_core_power_per_site_for_Vbump()
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    On Error GoTo errHandler
    
    Set g_ApplyLevelTimingVmain = Nothing
    Set g_ApplyLevelTimingValt = Nothing
    
    TheExec.DataManager.DecomposePinList "CorePower", p_ary, p_cnt
    For i = 0 To p_cnt - 1
        If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
           g_ApplyLevelTimingVmain.AddPin UCase((p_ary(i)))
           g_ApplyLevelTimingValt.AddPin UCase((p_ary(i)))
           InstName = GetInstrument(p_ary(i), 0)
           Select Case InstName
                  Case "DC-07"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).Value = CDbl(Format(TheHdw.DCVI.Pins(p_ary(i)).Voltage, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).Value = CDbl(Format(TheHdw.DCVI.Pins(p_ary(i)).Voltage, "0.000"))
                  Case "VHDVS"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).Value = CDbl(Format(TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).Value = CDbl(Format(TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt.Value, "0.000"))
                  Case "HexVS"
                        g_ApplyLevelTimingVmain.Pins(UCase(p_ary(i))).Value = CDbl(Format(TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value, "0.000"))
                        g_ApplyLevelTimingValt.Pins(UCase(p_ary(i))).Value = CDbl(Format(TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt.Value, "0.000"))
                  Case "HSD-U"
                  Case Else
                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Save_core_power_per_site"
            End Select
        End If
    Next i
   Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> Shmoo_Save_core_power_per_site:: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function dynamic_SELSRM_source_bits(SELSRAM_DSSC As String, BlockType As String) As String

''DSSC pin seq need modified for each project\\Hard coding\\
 Dim BitsDef As String: BitsDef = "VDD_DISP,VDD_AVE,VDD_GPU,VDD_ECPU,VDD_PCPU,VDD_DCS_DDR,VDD_SOC"
 Dim BitsDefArr() As String
 Dim SELSRAMArr() As String
 Dim i As Long
 Dim BitsOrderInfo As New Dictionary
 Dim BitsNum As Long
 Dim BlockTypeNum As Long
 Dim logicPin As String
 Dim SELSRM As String
 Dim DSSCSelSrmOpposite As Long
 Dim BitValue() As String
 ReDim SELSRAMArr(Len(SELSRAM_DSSC) - 1)
 BitsOrderInfo.RemoveAll
 BitsDefArr = Split(BitsDef, ",")
 For i = 0 To Len(SELSRAM_DSSC) - 1
    SELSRAMArr(i) = CStr(Mid(SELSRAM_DSSC, i + 1, 1))
 Next i
 ReDim Preserve BitsDefArr(UBound(BitsDefArr))
 BitsNum = UBound(BitsDefArr)
 
 If UBound(BitsDefArr) <> UBound(SELSRAMArr) Then
    TheExec.ErrorLogMessage "Number of bits not match with SELSRAM Char Info "
 Else
    For i = 0 To BitsNum
       If Not BitsOrderInfo.Exists(BitsDefArr(i)) Then
          BitsOrderInfo.Add (BitsDefArr(i)), SELSRAMArr(i)
       Else
          TheExec.ErrorLogMessage "Duplicate Rails, Please check"
       End If
    Next i
 End If
 ''\\Hard coding\\
 If BlockType <> "" Then
    If UCase(BlockType) Like "*CPU*" Then
        If UCase(BlockType) Like "*SCAN*" Or UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
           BlockType = "CPUSCAN"
        ElseIf UCase(BlockType) Like "*BST*" Or UCase(BlockType) Like "*BIR*" Or UCase(BlockType) Like "*BIST*" Or UCase(BlockType) Like "*RET*" Or UCase(BlockType) Like "*MBIST*" Or UCase(BlockType) Like "*BISR*" Then
           BlockType = "CPUMBIST"
        End If
    ElseIf UCase(BlockType) Like "*SOC*" Then
        If UCase(BlockType) Like "*SCAN*" Or UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
           BlockType = "SOCSCAN"
        ElseIf UCase(BlockType) Like "*BST*" Or UCase(BlockType) Like "*BIR*" Or UCase(BlockType) Like "*BIST*" Or UCase(BlockType) Like "*RET*" Or UCase(BlockType) Like "*MBIST*" Or UCase(BlockType) Like "*BISR*" Then
           BlockType = "SOCMBIST"
        End If
    ElseIf UCase(BlockType) Like "*GPU*" Or UCase(BlockType) Like "*GFX*" Then
        If UCase(BlockType) Like "*SCAN*" Or UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
           BlockType = "GFXSCAN"
        ElseIf UCase(BlockType) Like "*BST*" Or UCase(BlockType) Like "*BIR*" Or UCase(BlockType) Like "*BIST*" Or UCase(BlockType) Like "*RET*" Or UCase(BlockType) Like "*MBIST*" Or UCase(BlockType) Like "*BISR*" Then
           BlockType = "GFXMBIST"
        End If
    End If
 End If
 
 For i = 0 To UBound(GetSelSram.Block)
    If UCase(GetSelSram.Block(i).DomainName) <> "" Then
      If UCase(BlockType) Like "*" & UCase(GetSelSram.Block(i).DomainName) & "*" Then
         BlockTypeNum = i
         Exit For
      End If
    End If
 Next i
 
 If BlockTypeNum <> -1 Then
   ReDim BitValue(UBound(GetSelSram.Block(BlockTypeNum).DomainBits))
   For i = 0 To UBound(GetSelSram.Block(BlockTypeNum).DomainBits)
      logicPin = GetSelSram.Block(BlockTypeNum).DomainBits(i).logicPin
      DSSCSelSrmOpposite = GetSelSram.Block(BlockTypeNum).DomainBits(i).SelSram1
       
      If BitsOrderInfo.Exists(logicPin) = True Then
         SELSRM = BitsOrderInfo(logicPin)
      Else
         TheExec.ErrorLogMessage "Wrong Logic Pin Name in SELSRM_Mapping_Table"
      End If
       
      If UCase(SELSRM) = "1" Then
         If DSSCSelSrmOpposite = 1 Then
            BitValue(i) = 1
         Else
            BitValue(i) = 0
         End If
      ElseIf UCase(SELSRM) = "0" Then
         If DSSCSelSrmOpposite = 1 Then
            BitValue(i) = 0
         Else
            BitValue(i) = 1
         End If
      ElseIf UCase(SELSRM) = "S" Then
            BitValue(i) = "S"
      End If
   Next i
 End If
 dynamic_SELSRM_source_bits = Join(BitValue, "")
       
End Function
Public Function DecodingRealSourceBit(Source_Bits As String, BlockType As String) As String
''\\Hard coding\\
 Dim BitsDef As String: BitsDef = "VDD_DISP,VDD_AVE,VDD_GPU,VDD_ECPU,VDD_PCPU,VDD_DCS_DDR,VDD_SOC"
 Dim BitsDefArr() As String
 Dim RailsDecodingInfo As New Dictionary
 Dim BlockTypeNum As Long
 Dim DSSCSelSrmOpposite As Long
 Dim BitsValue As String
 Dim DcodingRailInfo() As String
 Dim logicPin As String
 Dim i As Long
 
 BitsDefArr = Split(BitsDef, ",")
 ReDim Preserve BitsDefArr(UBound(BitsDefArr))
 ReDim DcodingRailInfo(UBound(BitsDefArr))
 
 For i = 0 To UBound(GetSelSram.Block)
    If UCase(GetSelSram.Block(i).DomainName) <> "" Then
      If UCase(BlockType) Like "*" & UCase(GetSelSram.Block(i).DomainName) & "*" Then
         BlockTypeNum = i
         Exit For
      End If
    End If
 Next i
 
 For i = 0 To Len(Source_Bits) - 1
  logicPin = GetSelSram.Block(BlockTypeNum).DomainBits(i).logicPin
  DSSCSelSrmOpposite = GetSelSram.Block(BlockTypeNum).DomainBits(i).SelSram1
  BitsValue = CStr(Mid(Source_Bits, i + 1, 1))
   If Not RailsDecodingInfo.Exists(logicPin) = True Then
     If DSSCSelSrmOpposite = 1 Then
        RailsDecodingInfo.Add (logicPin), BitsValue
     Else
        If BitsValue = "1" Then
          RailsDecodingInfo.Add (logicPin), 0
        ElseIf BitsValue = "0" Then
          RailsDecodingInfo.Add (logicPin), 1
        ElseIf UCase(BitsValue) = "S" Then
          RailsDecodingInfo.Add (logicPin), "S"
        End If
     End If
   End If
 Next i

 For i = 0 To UBound(BitsDefArr)
  DcodingRailInfo(i) = RailsDecodingInfo(BitsDefArr(i))
 Next i
 
 DecodingRealSourceBit = Join(DcodingRailInfo, "")
End Function
    Public Function DecomposePattSet(Init1 As Pattern, Init2 As Pattern, Init3 As Pattern, Init4 As Pattern, Init5 As Pattern, Init6 As Pattern, Init7 As Pattern, Init8 As Pattern, Init9 As Pattern, _
                                      Init10 As Pattern, Payload1 As Pattern, Payload2 As Pattern, Payload3 As Pattern, Payload4 As Pattern, Payload5 As Pattern)
    Dim Pat_init1() As String
    Dim Pats_Num As Long
    Dim patIdx As Integer
    Dim tmpIN As String: tmpIN = ""
    Dim tmpPL As String: tmpPL = ""
    Dim INIArr() As String
    Dim PLLArr() As String

    Dim PL_Start_Idx As Integer
    PL_Start_Idx = 0
    On Error GoTo errHandler
    
    If Init1 <> "" Then
        TheHdw.Patterns(Init1).ValidatePatlist
        Pat_init1 = TheExec.DataManager.Raw.GetPatternsInSet(CStr(Init1), Pats_Num)
        If UBound(Pat_init1) > 0 Then
           For patIdx = 0 To Pats_Num - 1
            
            
               If UCase(Pat_init1(patIdx)) Like "*_IN*" And PL_Start_Idx = 0 Then
                  tmpIN = tmpIN & Pat_init1(patIdx) & ","
               Else 'If UCase(Pat_init1(PatIdx)) Like "*_PL*" Or UCase(Pat_init1(PatIdx)) Like "*_FULP*" Then
                  PL_Start_Idx = PL_Start_Idx + 1
                  tmpPL = tmpPL & Pat_init1(patIdx) & ","
                  
               'Else
               '   TheExec.DataLog.WriteComment "Pattern doesn't exist in PatternSet"
               End If
           Next patIdx
           tmpIN = Mid(tmpIN, 1, Len(tmpIN) - 1)
           tmpPL = Mid(tmpPL, 1, Len(tmpPL) - 1)
           INIArr() = Split(tmpIN, ",")
           PLLArr() = Split(tmpPL, ",")
           ReDim Preserve INIArr(9)
           ReDim Preserve PLLArr(4)
           Init1.Value = INIArr(0)
           Init2.Value = INIArr(1)
           Init3.Value = INIArr(2)
           Init4.Value = INIArr(3)
           Init5.Value = INIArr(4)
           Init6.Value = INIArr(5)
           Init7.Value = INIArr(6)
           Init8.Value = INIArr(7)
           Init9.Value = INIArr(8)
           Init10.Value = INIArr(9)
           Payload1.Value = PLLArr(0)
           Payload2.Value = PLLArr(1)
           Payload3.Value = PLLArr(2)
           Payload4.Value = PLLArr(3)
           Payload5.Value = PLLArr(4)
        End If
    End If
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> DecomposePattSet:: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    End Function


Public Function Decide_Pmode_ForceVoltage(PerformanceMode As String, power_pins As String, Pmode_Voltage As String)
    Dim p_ary() As String, p_cnt As Long, i As Long, j As Long
    Dim Dc_cat As String, Dc_spec_type As String, dc_sel As String
    Dim sp As Variant, t As String
    Dim PinValue As String
    Dim PerformanceModeArr() As String

    If power_pins = "" Then Exit Function
    TheExec.DataManager.GetInstanceContext Dc_cat, dc_sel, t, t, t, t, t, t
    For Each sp In TheExec.specs.DC.Categories(UCase(Dc_cat)).SpecList
        sp = LCase(sp)
        If sp Like "*_var_c" Then
            Dc_spec_type = "C"
        ElseIf sp Like "*_var_g" Then
            Dc_spec_type = "G"
        ElseIf sp Like "*_var_s" Then
            Dc_spec_type = "S"
        ElseIf sp Like "*_var_h" Then
            Dc_spec_type = "H"
        ElseIf sp Like "*_var_r" Then
            Dc_spec_type = "R"
        ElseIf sp Like "*_var" Then ''added case for new DC spec sheets method
            Dc_spec_type = ""
        Else
            TheExec.ErrorLogMessage "DC spec " & sp & " is not ended with _VAR_C/S/G/H in " & TheExec.DataManager.instanceName
        End If
        Exit For
    Next sp
    

    If UCase(Dc_spec_type) = "C" Or UCase(Dc_spec_type) = "G" Or UCase(Dc_spec_type) = "S" Or UCase(Dc_spec_type) = "H" Or UCase(Dc_spec_type) = "R" Then ''added case for new DC spec sheets method
       Dc_spec_type = "_" & Dc_spec_type
    End If
    
    
    PerformanceModeArr = Split(PerformanceMode, ":")
    
    If UBound(PerformanceModeArr) > 0 Then
        If UCase(PerformanceModeArr(1)) Like "LV" Then dc_sel = "MIN"
        If UCase(PerformanceModeArr(1)) Like "NV" Then dc_sel = "TYP"
        If UCase(PerformanceModeArr(1)) Like "HV" Then dc_sel = "MAX"
    End If
    
    TheExec.DataManager.DecomposePinList power_pins, p_ary, p_cnt
    For i = 0 To p_cnt - 1
        p_ary(i) = LCase(p_ary(i))
''\\Hard coding "_VOP"\\
        If UCase(dc_sel) Like "TYP" Then
            PinValue = Format(CStr(TheExec.specs.DC.Item(p_ary(i) & "_VOP" & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).Typ.Value), "0.000")
        ElseIf UCase(dc_sel) Like "MAX" Then
            PinValue = Format(CStr(TheExec.specs.DC.Item(p_ary(i) & "_VOP" & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).max.Value), "0.000")
        ElseIf UCase(dc_sel) Like "MIN" Then
            PinValue = Format(CStr(TheExec.specs.DC.Item(p_ary(i) & "_VOP" & "_" & "VAR" & Dc_spec_type).Categories(PerformanceModeArr(0)).Min.Value), "0.000")
        End If
        Pmode_Voltage = Pmode_Voltage & ";" & UCase(p_ary(i)) & ":" & "V" & ":" & PinValue
    Next i
    Pmode_Voltage = Mid(Pmode_Voltage, 2, Len(Pmode_Voltage) - 1)
End Function


Public Function Char_Process_DigString(DigDSSC_BitSize As String, DigDSSC_Seg As String, DigDSSC_DigPin As String, _
                                       ByRef DigCapName() As String, _
                                       ByRef DigSrcPin As String, _
                                       ByRef DigCapPin As String, ByRef DigSrcSize As String, _
                                       ByRef DigCapSize As String, _
                                       ByRef DigCap_Info_Dict As Dictionary) As Long
                                   
                                       
        Dim DigDSSC_Seg_Arr_Split() As String
        Dim DigCapEachSgmt_Info() As String
        Dim DigDSSC_BitSize_Arr() As String
        Dim DigDSSC_DigPin_Arr() As String
        Dim DigDSSC_Seg_Arr() As String
        Dim i As Long
    
        DigDSSC_BitSize_Arr = Split(DigDSSC_BitSize, "|")
        If UBound(DigDSSC_BitSize_Arr) = 1 Then
           DigSrcSize = DigDSSC_BitSize_Arr(0)
           DigCapSize = DigDSSC_BitSize_Arr(1)
        ElseIf UBound(DigDSSC_BitSize_Arr) = 0 Then
           DigSrcSize = DigDSSC_BitSize_Arr(0)
           DigCapSize = ""
        End If
        DigDSSC_DigPin_Arr = Split(DigDSSC_DigPin, "|")
        If UBound(DigDSSC_DigPin_Arr) = 1 Then
           DigSrcPin = DigDSSC_DigPin_Arr(0)
           DigCapPin = DigDSSC_DigPin_Arr(1)
        ElseIf UBound(DigDSSC_DigPin_Arr) = 0 Then
           DigSrcPin = DigDSSC_DigPin_Arr(0)
           DigCapPin = ""
        End If
    
        DigDSSC_Seg_Arr = Split(DigDSSC_Seg, "|")
        If UBound(DigDSSC_Seg_Arr) = 1 Then
            DigDSSC_Seg_Arr_Split = Split(DigDSSC_Seg_Arr(1), "+")
            DigCap_Info_Dict.RemoveAll
            ReDim DigCapName(UBound(DigDSSC_Seg_Arr_Split))
            
            For i = 0 To UBound(DigDSSC_Seg_Arr_Split)
                DigCapEachSgmt_Info = Split(DigDSSC_Seg_Arr_Split(i), ":")
                DigCap_Info_Dict.Add DigCapEachSgmt_Info(1), CLng(DigCapEachSgmt_Info(0))
                DigCapName(i) = DigCapEachSgmt_Info(1)
            Next i
         Else
         End If
End Function


Public Function Char_Process_DSP_Capture(DigCapName() As String, OutDspWave As DSPWave, DigCap_Info_Dict As Dictionary, DigCap_Pin As String) As Long
       Dim Bin2Dec As Double
       Dim OutPutLen As Long
       Dim DSSC_Capture_Out As String
       Dim CaptureName As Variant
       Dim i As Long, StartNum As Long
       Dim CaptureBits As Long
       Dim DSSC_Capture_Out_Dict As New Dictionary
       Dim DSSC_Sgmt_Name_String As String
       Dim OutPut_Sgmt_Name As String
       Dim site As Variant

              DSSC_Capture_Out_Dict.RemoveAll
              For Each site In TheExec.sites
                  StartNum = 0
                  For i = 0 To OutDspWave(site).SampleSize - 1
                      DSSC_Capture_Out = DSSC_Capture_Out & CStr(OutDspWave(site).Element(i))
                  Next i
                  DSSC_Capture_Out_Dict.Add site, DSSC_Capture_Out
                  
                  For Each CaptureName In DigCapName
                   If DigCap_Info_Dict.Exists(CaptureName) Then
                      DSSC_Sgmt_Name_String = ""
                      Bin2Dec = 0
                      CaptureBits = CLng(DigCap_Info_Dict.Item(CaptureName))
                      For i = StartNum To (StartNum + CaptureBits - 1)
                          DSSC_Sgmt_Name_String = DSSC_Sgmt_Name_String & CStr(OutDspWave(site).Element(i))
                      Next i
                      OutPutLen = Len(DSSC_Sgmt_Name_String) - 1
                      For i = 0 To OutPutLen
                          Bin2Dec = CDbl(Bin2Dec + Mid(DSSC_Sgmt_Name_String, (Len(DSSC_Sgmt_Name_String) - i), 1) * 2 ^ i)
                      Next i
                      OutPut_Sgmt_Name = CaptureName & "(OutPutString:" & DSSC_Sgmt_Name_String & ")"
                      If TheExec.DevChar.Setups.IsRunning = True Then
                            If TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).output.SuspendDatalog = False Then
                              TheExec.Flow.TestLimit resultVal:=Bin2Dec, ForceResults:=tlForceNone, Tname:="DigCap" & ":" & OutPut_Sgmt_Name, PinName:=DigCap_Pin, lowVal:=0, hiVal:=2 ^ (OutPutLen + 1) - 1, TNum:=g_TestNum
                            Else
                              TheExec.Flow.TestLimit resultVal:=Bin2Dec, ForceResults:=tlForceNone, Tname:="DigCap" & ":" & OutPut_Sgmt_Name, PinName:=DigCap_Pin, lowVal:=0, hiVal:=2 ^ (OutPutLen + 1) - 1
                            End If
                      Else
                            TheExec.Flow.TestLimit resultVal:=Bin2Dec, ForceResults:=tlForceNone, Tname:="DigCap" & ":" & OutPut_Sgmt_Name, PinName:=DigCap_Pin, lowVal:=0, hiVal:=2 ^ (OutPutLen + 1) - 1
                      End If
                      StartNum = StartNum + CaptureBits
                    End If
                  Next CaptureName
              Next site
              g_TestNum = g_TestNum + 1 ' 20193021 update
 
End Function




Public Function Decide_DC_Level(DC_Level As PinListData, DC_Level_Alt As PinListData, DC_Level_Vmain As PinListData, BlockType As String)

''\\Hard coding\\
If BlockType <> "" Then
  If UCase(BlockType) Like "*CPU*" Then
    If UCase(BlockType) Like "*SCAN*" Or UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
       DC_Level = g_ApplyLevelTimingValt.Copy
       BlockType = "CPUSCAN"
    ElseIf UCase(BlockType) Like "*BST*" Or UCase(BlockType) Like "*BIR*" Or UCase(BlockType) Like "*BIST*" Or UCase(BlockType) Like "*RET*" Or UCase(BlockType) Like "*MBIST*" Or UCase(BlockType) Like "*BISR*" Then
       DC_Level = g_ApplyLevelTimingVmain.Copy
       BlockType = "CPUMBIST"
    Else
       DC_Level = g_ApplyLevelTimingValt.Copy
    End If
  ElseIf UCase(BlockType) Like "*SOC*" Then
    If UCase(BlockType) Like "*SCAN*" Or UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
       DC_Level = g_ApplyLevelTimingValt.Copy
       BlockType = "SOCSCAN"
    ElseIf UCase(BlockType) Like "*BST*" Or UCase(BlockType) Like "*BIR*" Or UCase(BlockType) Like "*BIST*" Or UCase(BlockType) Like "*RET*" Or UCase(BlockType) Like "*MBIST*" Or UCase(BlockType) Like "*BISR*" Then
       DC_Level = g_ApplyLevelTimingVmain.Copy
       BlockType = "SOCMBIST"
    Else
       DC_Level = g_ApplyLevelTimingValt.Copy
    End If
  ElseIf UCase(BlockType) Like "*GPU*" Or UCase(BlockType) Like "*GFX*" Then
    If UCase(BlockType) Like "*SCAN*" Or UCase(BlockType) Like "*SA*" Or UCase(BlockType) Like "*TD*" Then
       DC_Level = g_ApplyLevelTimingValt.Copy
       BlockType = "GFXSCAN"
    ElseIf UCase(BlockType) Like "*BST*" Or UCase(BlockType) Like "*BIR*" Or UCase(BlockType) Like "*BIST*" Or UCase(BlockType) Like "*RET*" Or UCase(BlockType) Like "*MBIST*" Or UCase(BlockType) Like "*BISR*" Then
       DC_Level = g_ApplyLevelTimingVmain.Copy
       BlockType = "GFXMBIST"
    Else
       DC_Level = g_ApplyLevelTimingValt.Copy
    End If
  End If
Else
     DC_Level = g_ApplyLevelTimingValt.Copy
End If

End Function




Public Function Decide_Switching_Bit(digSrc_EQ As String, DSPWaveSwitch As DSPWave, DC_Level As PinListData, BlockType As String, SELSRM_Rails As String, Optional shmoo_pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary) As String

  Dim logicPin As String
  Dim SramPin As String
  Dim DSSC_Switching_Voltage As New PinListData
  Dim Sdomain As Long
  Dim DSSCSelSrmOpposite As Long
  Dim BlockTypeNum As Long
  Dim i As Integer
  Dim ReturnString() As String
  Dim LogicValue As Double
  Dim SramValue As Double
  On Error GoTo errHandler
  BlockTypeNum = -1
  
  ReDim ReturnString(Len(digSrc_EQ) - 1)
  Decide_DSSC_Switching_Voltage DSSC_Switching_Voltage, DC_Level, shmoo_pin, ShmooPinsVoltage, ForcePin, SetForceVoltage
  For i = 0 To UBound(GetSelSram.Block)
     If UCase(GetSelSram.Block(i).DomainName) <> "" Then
       If UCase(BlockType) Like "*" & UCase(GetSelSram.Block(i).DomainName) & "*" Then
          BlockTypeNum = i
          Exit For
       End If
     End If
  Next i
  
  If BlockTypeNum <> -1 Then
    For i = 0 To Len(digSrc_EQ) - 1
     If UCase(CStr(Mid(digSrc_EQ, i + 1, 1))) Like "S" Then
         logicPin = GetSelSram.Block(BlockTypeNum).DomainBits(i).logicPin
         SramPin = GetSelSram.Block(BlockTypeNum).DomainBits(i).SramPin
         DSSCSelSrmOpposite = GetSelSram.Block(BlockTypeNum).DomainBits(i).SelSram1
         For Each site In TheExec.sites.Active
            LogicValue = CDbl(DSSC_Switching_Voltage.Pins(logicPin).Value)
            SramValue = CDbl(DSSC_Switching_Voltage.Pins(SramPin).Value)
            If DSSCSelSrmOpposite = 0 Then
                Sdomain = IIf((LogicValue > SramValue), 1, 0)
                DSPWaveSwitch.Element(i) = Sdomain
                ReturnString(i) = Sdomain
            ElseIf DSSCSelSrmOpposite = 1 Then
                Sdomain = IIf((LogicValue > SramValue), 0, 1)
                DSPWaveSwitch.Element(i) = Sdomain
                ReturnString(i) = Sdomain
            End If
         Next site
      Else
         For Each site In TheExec.sites.Active
             DSPWaveSwitch.Element(i) = CDbl(Mid(digSrc_EQ, i + 1, 1))
             ReturnString(i) = CDbl(Mid(digSrc_EQ, i + 1, 1))
         Next site
      End If
    Next i
    
    Set PrintSwitchDspWave = Nothing
    PrintSwitchDspWave = DSPWaveSwitch
    g_BlockType = BlockType

    If TheExec.DevChar.Setups.IsRunning = False Then
       Decide_Switching_Bit = Join(ReturnString, "")
       SELSRM_Rails = DecodingRealSourceBit(Decide_Switching_Bit, BlockType)
    Else
       Decide_Switching_Bit = digSrc_EQ
       SELSRM_Rails = DecodingRealSourceBit(Decide_Switching_Bit, BlockType)
    End If
  End If
  Exit Function
  
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + "Decide_Switching_Bit" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Set_Run_Level_Vbump(Power_Run_Scenario As String, powerPin As String, set_init As Boolean, seq As Long)
    Dim VoltageLevel As String, Scenario As String
    Dim i As Long
    Dim init_level As String
    Dim pl_level As String
    Dim Power_Run_Scenario_ary() As String
    Dim inst_name As String
    Dim inst_level As String
    
    inst_level = Right(TheExec.DataManager.instanceName, 2)
    init_level = "-99"
    pl_level = "-99"
    

    If set_init = True Then
        init_level = "NV"
        If Not g_FirstSetp = True Then
           If Not (Power_Level_Last Like init_level) Then
               If g_shmoo_ret = True Then
                  Shmoo_Restore_Power_per_site_Vbump_NV True
               End If
               TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
               TheHdw.Wait 0.001
               Shmoo_Restore_Power_per_site_Vbump_NV True, True
           End If
        End If
        g_FirstSetp = False
        Power_Level_Last = init_level
        If init_level Like "-99" Then TheExec.ErrorLogMessage "Power Run Scenario " & Power_Run_Scenario & " is not supported"
    Else
        g_PLSWEEP = False
        If LCase(Power_Run_Scenario) Like LCase("*pl_Sweep*") Then
            pl_level = "Sweep"
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump powerPin
            g_PLSWEEP = True
        
        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl_NV*") Then
            pl_level = "NV"
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump_NV

        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_Sweep*") Then
            pl_level = "Sweep"
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump powerPin
            g_PLSWEEP = True

        ElseIf LCase(Power_Run_Scenario) Like LCase("*pl" & seq & "_NV*") Then
            pl_level = "NV"
            If Not (Power_Level_Last Like pl_level) Then Shmoo_Restore_Power_per_site_Vbump_NV
        End If
           
        Power_Level_Last = pl_level
        If pl_level Like "-99" Then TheExec.ErrorLogMessage "Power Run Scenario " & Power_Run_Scenario & " is not supported"
    End If
    
End Function
Public Function Shmoo_Restore_Power_per_site_Vbump_Retention(Shmoo_Apply_Pin As String, RetentionShmoo As Boolean)
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim InstName As String
    Dim site As Variant
    Dim Shmoo_Apply_Pin_Arry() As String
    Dim SRAMRampUpFirst As New SiteBoolean
    Dim LogicRampdownFirst As New SiteBoolean
    Dim SramShmooPower As String: SramShmooPower = ""
    Dim RetentionShmoo_Pins_Dict As New Dictionary
    Dim Retention_ForceV_Arr() As String
    On Error GoTo errHandler
    Retention_ForceV_Arr = Split(g_Retention_ForceV, ",")
    
    If RetentionShmoo = True Then
       If TheExec.Flow.EnableWord("Enable_RET_RampDownUp") = True Then
          Retention_RampdownUp Shmoo_Apply_Pin, "DOWN"
       Else
           Create_Pin_Dic Shmoo_Apply_Pin, RetentionShmoo_Pins_Dict
           Shmoo_Apply_Pin_Arry = Split(Shmoo_Apply_Pin, ",")
           For i = 0 To UBound(Shmoo_Apply_Pin_Arry)
              If Not UCase(Shmoo_Apply_Pin_Arry(i)) Like UCase(SramShmooPower) Or SramShmooPower = "" Then
                If TheExec.DataManager.ChannelType(Shmoo_Apply_Pin_Arry(i)) <> "N/C" Then
                  InstName = GetInstrument(Shmoo_Apply_Pin_Arry(i), 0)
                  Select Case InstName
                         Case "DC-07"
                              TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                         Case "VHDVS"
                              For Each site In TheExec.sites
                                  TheHdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Main.Value = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).Value
                              Next site
                    Case "HexVS"
                              For Each site In TheExec.sites
                                  TheHdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Main.Value = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).Value
                              Next site
                    Case "HSD-U"
                    Case Else
                              TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                  End Select
                End If
              End If
           Next i
       End If
    End If
     
    If g_Retention_VDD <> "" And TheExec.Flow.EnableWord("Enable_RET_Ramping") = False Then
       TheExec.DataManager.DecomposePinList g_Retention_VDD, p_ary, p_cnt
       For i = 0 To p_cnt - 1
           If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
              If Not RetentionShmoo_Pins_Dict(p_ary(i)) = True Then
              InstName = GetInstrument(p_ary(i), 0)
                   Select Case InstName
                       Case "DC-07"
                             TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention"
                       Case "VHDVS"
'                             For Each site In TheExec.sites
                                 TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = Retention_ForceV_Arr(i)
'                             Next site
                       Case "HexVS"
'                             For Each site In TheExec.sites
                                 TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = Retention_ForceV_Arr(i)
'                             Next site
                       Case "HSD-U"
                       Case Else
                                TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump_Retention"
                   End Select
              End If
           End If
       Next i
    End If
    
    TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
    TheHdw.Wait 0.001
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> Shmoo_Restore_Power_per_site_Vbump_Retention:: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Decide_DSSC_Switching_Voltage(DSSC_Switching_Voltage As PinListData, DC_Level As PinListData, Optional Shmoo_Apply_Pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary)

Dim p_ary() As String, p_cnt As Long, i As Long
Dim InstName As String
Dim site As Variant
Dim CorePower_Pins_Dict As New Dictionary
On Error GoTo errHandler

Set PrintDSSCSwitchVoltage = Nothing
DSSC_Switching_Voltage = DC_Level.Copy
Create_Pin_Dic "CorePower", CorePower_Pins_Dict


If ForcePin <> "" Then
  TheExec.DataManager.DecomposePinList ForcePin, p_ary, p_cnt
     For i = 0 To p_cnt - 1
         If Not CorePower_Pins_Dict.Exists(LCase(p_ary(i))) = True Then
            DSSC_Switching_Voltage.AddPin (UCase(p_ary(i)))
            For Each site In TheExec.sites
'              DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).Value = SetForceVoltage(UCase(p_ary(i)))
              DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).Value = SetForceVoltage(UCase(p_ary(i)))
            Next site
          Else
            For Each site In TheExec.sites
'              DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).Value = SetForceVoltage(UCase(p_ary(i)))
              DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).Value = SetForceVoltage(UCase(p_ary(i)))
            Next site
         End If
     Next i
 End If

 If Shmoo_Apply_Pin <> "" Then
  TheExec.DataManager.DecomposePinList Shmoo_Apply_Pin, p_ary, p_cnt
     For i = 0 To p_cnt - 1
         If Not CorePower_Pins_Dict.Exists(LCase(p_ary(i))) = True Then
            DSSC_Switching_Voltage.AddPin (UCase(p_ary(i)))
            For Each site In TheExec.sites
'              DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).Value = ShmooPinsVoltage.Pins(UCase(p_ary(i))).Value
              DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).Value = ShmooPinsVoltage.Pins(UCase(p_ary(i))).Value
            Next site
          Else
            For Each site In TheExec.sites
'              DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).Value = ShmooPinsVoltage.Pins(UCase(p_ary(i))).Value
              DSSC_Switching_Voltage.Pins(UCase(p_ary(i))).Value = ShmooPinsVoltage.Pins(UCase(p_ary(i))).Value
            Next site
         End If
     Next i
 End If
 
 PrintDSSCSwitchVoltage = DSSC_Switching_Voltage.Copy
 
 Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + "Decide_DSSC_Switching_Voltage" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
 End Function
Public Function Shmoo_Restore_Power_per_site_Vbump_NV(Optional Init As Boolean = False, Optional InitAltRecover As Boolean = False)
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim InstName As String
    Dim site As Variant
    Dim CorePower_Pins_Dict As New Dictionary
    Dim OtherPower As String: OtherPower = "CorePower"
    Dim Payload_Voltage_Vmain As New PinListData
    Dim Payload_Voltage_Valt As New PinListData
    On Error GoTo errHandler
    
     Payload_Voltage_Vmain = g_ApplyLevelTimingVmain.Copy
     Payload_Voltage_Valt = g_ApplyLevelTimingValt.Copy
     Create_Pin_Dic "CorePower", CorePower_Pins_Dict
  
     If g_ForceCond_VDD <> "" And Init = False Then
        TheExec.DataManager.DecomposePinList g_ForceCond_VDD, p_ary, p_cnt
           For i = 0 To p_cnt - 1
            If g_CharInputString_Voltage_Dict.Exists((UCase(p_ary(i)))) = True Then
               If Not CorePower_Pins_Dict.Exists(LCase(p_ary(i))) = True Then
                  Payload_Voltage_Vmain.AddPin (UCase(p_ary(i)))
                  Payload_Voltage_Valt.AddPin (UCase(p_ary(i)))
'                  For Each site In TheExec.sites
                    Payload_Voltage_Vmain.Pins(UCase(p_ary(i))).Value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
                    Payload_Voltage_Valt.Pins(UCase(p_ary(i))).Value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
'                  Next site
                  OtherPower = OtherPower & "," & p_ary(i)
                Else
'                  For Each site In TheExec.sites
                    Payload_Voltage_Vmain.Pins(UCase(p_ary(i))).Value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
                    Payload_Voltage_Valt.Pins(UCase(p_ary(i))).Value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
'                  Next site
               End If
             End If
           Next i
       End If
     '=========================================== Applied voltage to Valt==================================================
      If Init = False Or InitAltRecover = True Then
        TheExec.DataManager.DecomposePinList OtherPower, p_ary, p_cnt
            For i = 0 To p_cnt - 1
               If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                   InstName = GetInstrument(p_ary(i), 0)
                   Select Case InstName
                      Case "DC-07"
                            TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                      Case "VHDVS"
'                           For Each site In TheExec.sites
                                   TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt.Value = Payload_Voltage_Valt.Pins(p_ary(i)).Value
'                           Next site
                      Case "HexVS"
'                           For Each site In TheExec.sites
                                   TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt.Value = Payload_Voltage_Valt.Pins(p_ary(i)).Value
'                           Next site
                      Case "HSD-U"
                      Case Else
                               TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                   End Select
               End If
            Next i
          If InitAltRecover = False Then
             TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt
             TheHdw.Wait 0.001
          End If
       End If
      '=========================================== Store to Vmain voltage which voltage same as Valt ==================================================
          If (Init = True And InitAltRecover = False) Or Init = False Then
            TheExec.DataManager.DecomposePinList OtherPower, p_ary, p_cnt
              For i = 0 To p_cnt - 1
                 If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                     InstName = GetInstrument(p_ary(i), 0)
                     Select Case InstName
                        Case "DC-07"
                              TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                        Case "VHDVS"
'                             For Each site In TheExec.sites
                                     TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = Payload_Voltage_Vmain.Pins(p_ary(i)).Value
'                             Next site
                        Case "HexVS"
'                             For Each site In TheExec.sites
                                     TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = Payload_Voltage_Vmain.Pins(p_ary(i)).Value
'                             Next site
                        Case "HSD-U"
                        Case Else
                                 TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                     End Select
                 End If
              Next i
           End If
     Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> Shmoo_Restore_Power_per_site_Vbump_NV:: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Shmoo_Restore_Power_per_site_Vbump(Shmoo_Apply_Pin As String)
    
    Dim p_ary() As String, p_cnt As Long, i As Long
    Dim InstName As String
    Dim site As Variant
    Dim Shmoo_Apply_Pin_Arry() As String
    Dim SRAMRampUpFirst As New SiteBoolean
    Dim LogicRampdownFirst As New SiteBoolean
    Dim SramShmooPower As String: SramShmooPower = ""
    On Error GoTo errHandler
    If g_ForceCond_VDD <> "" Then
        TheExec.DataManager.DecomposePinList g_ForceCond_VDD, p_ary, p_cnt
        For i = 0 To p_cnt - 1
           If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
               InstName = GetInstrument(p_ary(i), 0)
               Select Case InstName
                  Case "DC-07"
                        TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                  Case "VHDVS"
'                        For Each site In TheExec.sites
                               TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt.Value = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
'                        Next site
                  Case "HexVS"
'                        For Each site In TheExec.sites
                               TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt = g_CharInputString_Voltage_Dict(UCase(p_ary(i)))
'                        Next site
                  Case "HSD-U"
                  Case Else
                           TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
               End Select
           End If
        Next i
     End If

    If Shmoo_Apply_Pin <> "" Then
         Shmoo_Apply_Pin_Arry = Split(Shmoo_Apply_Pin, ",")
         For i = 0 To UBound(Shmoo_Apply_Pin_Arry)
            If Not UCase(Shmoo_Apply_Pin_Arry(i)) Like UCase(SramShmooPower) Or SramShmooPower = "" Then
              If TheExec.DataManager.ChannelType(Shmoo_Apply_Pin_Arry(i)) <> "N/C" Then
                InstName = GetInstrument(Shmoo_Apply_Pin_Arry(i), 0)
                   Select Case InstName
                      Case "DC-07"
                            TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & Shmoo_Apply_Pin_Arry(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                      Case "VHDVS"
                            For Each site In TheExec.sites
                                TheHdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Alt.Value = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).Value
                            Next site
                      Case "HexVS"
                            For Each site In TheExec.sites
                                TheHdw.DCVS.Pins(Shmoo_Apply_Pin_Arry(i)).Voltage.Alt.Value = g_Globalpointval.Pins(Shmoo_Apply_Pin_Arry(i)).Value
                            Next site
                      Case "HSD-U"
                      Case Else
                             TheExec.ErrorLogMessage "Instrument " & InstName & " for pin " & p_ary(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                   End Select
              End If
            End If
         Next i
     End If
     
     TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt
''wait time for Vmain switch to Valt
     TheHdw.Wait 0.001
     Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> Shmoo_Restore_Power_per_site_Vbump:: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Retention_RampdownUp(Shmoo_Apply_Pin As String, RampDirection As String)
    Dim StartVoltage As New PinListData
    Dim EndVoltage As New PinListData
    Dim p_ary() As String, p_cnt As Long, i As Long, InstName As String
    Dim step As Integer
    Dim StepNum As Integer

    On Error GoTo errHandler
    StepNum = 9 ' //hard coding//must be odd number and please modify this number by different project
    
    If UCase(RampDirection) = "DOWN" Then
       StartVoltage = g_ApplyLevelTimingValt.Copy
       EndVoltage = g_Globalpointval.Copy
    ElseIf UCase(RampDirection) = "UP" Then
       StartVoltage = g_Globalpointval.Copy
       EndVoltage = g_ApplyLevelTimingValt.Copy
    End If

    TheExec.DataManager.DecomposePinList Shmoo_Apply_Pin, p_ary, p_cnt
    
    If UCase(RampDirection) = "DOWN" Then
        For step = 1 To StepNum
            For i = 0 To p_cnt - 1
                If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                   If step Mod 2 = 1 Then
                      For Each site In TheExec.sites
                          TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = StartVoltage.Pins(p_ary(i)).Value - ((StartVoltage.Pins(p_ary(i)).Value - EndVoltage.Pins(p_ary(i)).Value) / StepNum) * step
                      Next site
                   ElseIf step Mod 2 = 0 Then
                      For Each site In TheExec.sites
                          TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt.Value = StartVoltage.Pins(p_ary(i)).Value - ((StartVoltage.Pins(p_ary(i)).Value - EndVoltage.Pins(p_ary(i)).Value) / StepNum) * step
                      Next site
                   End If
                End If
            Next i
            If step Mod 2 = 1 Then
               TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
               TheHdw.Wait 20 * 0.000001
            ElseIf step Mod 2 = 0 Then
               TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt
               TheHdw.Wait 20 * 0.000001
            End If
        Next step
    ElseIf UCase(RampDirection) = "UP" Then
        For step = 1 To StepNum
            For i = 0 To p_cnt - 1
                If TheExec.DataManager.ChannelType(p_ary(i)) <> "N/C" Then
                   If step Mod 2 = 1 Then
                      For Each site In TheExec.sites
                          TheHdw.DCVS.Pins(p_ary(i)).Voltage.Alt.Value = StartVoltage.Pins(p_ary(i)).Value - ((StartVoltage.Pins(p_ary(i)).Value - EndVoltage.Pins(p_ary(i)).Value) / StepNum) * step
                      Next site
                   ElseIf step Mod 2 = 0 Then
                      For Each site In TheExec.sites
                          TheHdw.DCVS.Pins(p_ary(i)).Voltage.Main.Value = StartVoltage.Pins(p_ary(i)).Value - ((StartVoltage.Pins(p_ary(i)).Value - EndVoltage.Pins(p_ary(i)).Value) / StepNum) * step
                      Next site
                   End If
                End If
            Next i
            If step Mod 2 = 1 Then
               TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageAlt
               TheHdw.Wait 20 * 0.000001
            ElseIf step Mod 2 = 0 Then
               TheHdw.DCVS.Pins("CorePower").Voltage.output = tlDCVSVoltageMain
               TheHdw.Wait 20 * 0.000001
            End If
        Next step
    End If
   
   Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> Shmoo_Save_core_power_per_site:: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'Public Function StoreEachPointResult(argc As Long, argv() As String)
Public Function StoreEachPointResult()
    Dim Suspend_Flag As Boolean
    Dim DevSetupName As String
'    Dim Execution_result As String
    Dim curr_axis As Variant
    Dim AxisOrder As Variant
    Dim Execution_result As New SiteVariant
    Dim site As Variant
    
    On Error GoTo err
'    Exit Function
    DevSetupName = TheExec.DevChar.Setups.ActiveSetupName
    Suspend_Flag = TheExec.DevChar.Setups.Item(DevSetupName).output.SuspendDatalog
    AxisOrder = TheExec.DevChar.Setups(DevSetupName).Shmoo.AxisOrder
    
    For Each site In TheExec.sites
        For MaxArrIndex = 0 To MaxArrIndex - 1
            Select Case AxisOrder
                Case tlDevCharAxisOrder_ZXY
                    If Y_Point <= Yaxis_index Then
                        If X_Point <= Xaxis_index Then
                            If Z_Point <= Zaxis_index Then
    '                        Debug.Print CStr(X_Point) & "," & CStr(Y_Point) & "," & CStr(Z_Point)
                               Execution_result = TheExec.DevChar.Results(DevSetupName).Shmoo.Points(X_Point, Y_Point, Z_Point).ExecutionResult
    '                            g_ShmooResult.Axis_CurrPoint(Count_Point).CurrResult = Execution_result
                                Select Case Execution_result
                                    Case tlDevCharResult_Pass
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "P"
                                    Case tlDevCharResult_Fail
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "F"
                                    Case tlDevCharResult_Alarm
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "A"
                                    Case tlDevCharResult_Error
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "E"
                                    Case Else
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "?"
                                End Select
                                Z_Point = Z_Point + 1
                                If Z_Point = Zaxis_index Then
                                    Z_Point = 0
                                    X_Point = X_Point + 1
                                    If X_Point = Xaxis_index Then
                                        X_Point = 0
                                        Y_Point = Y_Point + 1
                                    End If
                                End If
                            End If
                        End If
                    End If
            End Select
        Next MaxArrIndex
        Y_Point = 0
    Next site
    Exit Function
err:
    If AbortTest Then Exit Function Else Resume Next
End Function


'Public Function StoreEachPointResult(argc As Long, argv() As String)
Public Function StoreEachPointResult_2D()
    Dim Suspend_Flag As Boolean
    Dim DevSetupName As String
'    Dim Execution_result As String
    Dim curr_axis As Variant
    Dim AxisOrder As Variant
    Dim Execution_result As New SiteVariant
    Dim site As Variant
    
    On Error GoTo err
'    Exit Function
    DevSetupName = TheExec.DevChar.Setups.ActiveSetupName
    Suspend_Flag = TheExec.DevChar.Setups.Item(DevSetupName).output.SuspendDatalog
    AxisOrder = TheExec.DevChar.Setups(DevSetupName).Shmoo.AxisOrder
    
    For Each site In TheExec.sites
        For MaxArrIndex = 0 To MaxArrIndex - 1
'            Select Case AxisOrder
'                Case tlDevCharAxisOrder_ZXY
                    If Y_Point <= Yaxis_index Then
                        If X_Point <= Xaxis_index Then
'                            If Z_Point <= Zaxis_index Then
    '                        Debug.Print CStr(X_Point) & "," & CStr(Y_Point) & "," & CStr(Z_Point)
                               Execution_result = TheExec.DevChar.Results(DevSetupName).Shmoo.Points(X_Point, Y_Point).ExecutionResult
    '                            g_ShmooResult.Axis_CurrPoint(Count_Point).CurrResult = Execution_result
                                Select Case Execution_result
                                    Case tlDevCharResult_Pass
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "P"
                                    Case tlDevCharResult_Fail
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "F"
                                    Case tlDevCharResult_Alarm
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "A"
                                    Case tlDevCharResult_Error
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "E"
                                    Case Else
                                        g_ShmooResult.Axis_CurrPoint(MaxArrIndex).CurrResult = "?"
                                End Select
                                X_Point = X_Point + 1
                                If X_Point = Xaxis_index Then
                                    X_Point = 0
                                    Y_Point = Y_Point + 1
'                                    If X_Point = Xaxis_index Then
'                                        X_Point = 0
'                                        Y_Point = Y_Point + 1
'                                    End If
                                End If
'                            End If
                        End If
                    End If
'            End Select
        Next MaxArrIndex
        Y_Point = 0
    Next site
    Exit Function
err:
    If AbortTest Then Exit Function Else Resume Next
End Function




'Public Function StoreEachPointResult(argc As Long, argv() As String)
Public Function StoreEachPoint()
    Dim Suspend_Flag As Boolean
    Dim DevSetupName As String
    Dim Execution_result As String
    Dim curr_axis As Variant
    Dim AxisOrder As Variant
    Dim Tracking_Item As Variant
    Dim i, j, k As Integer
    
    On Error GoTo err
    i = 0
    j = 0
    k = 0
    DevSetupName = TheExec.DevChar.Setups.ActiveSetupName
    Suspend_Flag = TheExec.DevChar.Setups.Item(DevSetupName).output.SuspendDatalog
    AxisOrder = TheExec.DevChar.Setups(DevSetupName).Shmoo.AxisOrder
'    Exit Function
    If Count_Point > MaxArrIndex - 1 Then Exit Function ''  do not need to record
    For Each site In TheExec.sites
        For Each curr_axis In TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.List
'            Exit Function
            Select Case curr_axis
                Case 0
                    g_ShmooResult.Axis_CurrPoint(Count_Point).X_axis = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).Value
                    If TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count <> 0 Then ' Tracking case
                        With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                            For Each Tracking_Item In .List
                                g_ShmooResult.Axis_CurrPoint(Count_Point).X_axis_Tracking(i) = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(Tracking_Item).Value
                                i = i + 1 ' Never larger than X_Tracking_Point
                            Next Tracking_Item
                        End With
                    End If
                Case 1
                    g_ShmooResult.Axis_CurrPoint(Count_Point).Y_axis = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).Value
                    If TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count <> 0 Then ' Tracking case
                        With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                            For Each Tracking_Item In .List
                                g_ShmooResult.Axis_CurrPoint(Count_Point).Y_axis_Tracking(j) = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(Tracking_Item).Value
                                j = j + 1 ' Never larger than Y_Tracking_Point
                            Next Tracking_Item
                        End With
                    End If
                Case 2
                    g_ShmooResult.Axis_CurrPoint(Count_Point).Z_axis = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).Value
                    If TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Count <> 0 Then ' Tracking case
                        With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                            For Each Tracking_Item In .List
                                g_ShmooResult.Axis_CurrPoint(Count_Point).Z_axis_Tracking(k) = TheExec.DevChar.Results(DevSetupName).Shmoo.CurrentPoint.Axes(curr_axis).TrackingParameters(Tracking_Item).Value
                                k = k + 1 ' Never larger than Z_Tracking_Point
                            Next Tracking_Item
                        End With
                    End If
            End Select
        Next curr_axis
        i = 0
        j = 0
        k = 0
    Next site
    Count_Point = Count_Point + 1
''
    Exit Function
err:
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function Print3DShmooInfo(argc As Long, argv() As String)
Dim site As Variant
Dim i, j, k As Long
Dim ShmooResult As String
Dim x_val As Variant
Dim y_val As Variant
Dim z_val As Variant
Dim LVCC As Double
Dim HVCC As Double
Dim InstName As String
Dim curr_axis As Variant
Dim axis_pin(2) As String
Dim X_Axis_Val() As Double
Dim Y_Axis_Val() As Double
Dim Z_Axis_Val() As Double
Dim DevSetupName As String

Dim tmpStr() As String
Dim tmpStr1() As String
Dim tmpStr2() As String
Dim InstName_H As String
Dim InstName_L As String
Dim TmpVal As Integer
Dim X_Axis_TrackingPara() As String
Dim Y_Axis_TrackingPara() As String
Dim Z_Axis_TrackingPara() As String
Dim X_Axis_TrackingParaFrom() As Variant
Dim Y_Axis_TrackingParaFrom() As Variant
Dim Z_Axis_TrackingParaFrom() As Variant
Dim X_Axis_TrackingParaTo() As Variant
Dim Y_Axis_TrackingParaTo() As Variant
Dim Z_Axis_TrackingParaTo() As Variant
Dim Tracking_Item As Variant
Dim ii, jj, kk As Integer

Dim x_TrackingInfo As String
Dim x_stepsize As Variant
Dim x_num As Integer
Dim y_TrackingInfo As String
Dim y_stepsize As Variant
Dim y_num As Integer
Dim z_TrackingInfo As String
Dim z_stepsize As Variant
Dim z_num As Integer
Dim tmp_Tnum As Long

ii = 0
jj = 0
kk = 0
''Exit Function
DevSetupName = TheExec.DevChar.Setups.ActiveSetupName
For Each curr_axis In TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.List
        Select Case curr_axis
            Case 0
                ReDim X_Axis_Val(Xaxis_index - 1)
                If X_Tracking_Point <> 0 Then 'extract X Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim X_Axis_TrackingPara(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaFrom(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaTo(X_Tracking_Point - 1)
                    With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .List
                            X_Axis_TrackingPara(ii) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).ApplyTo.Pins), "_", "")
                            X_Axis_TrackingParaFrom(ii) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.from
                            X_Axis_TrackingParaTo(ii) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.To
                            If X_Axis_TrackingPara(ii) = "" Then X_Axis_TrackingPara(ii) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).Name), "_", "")
                            ii = ii + 1
                        Next Tracking_Item
                    End With
                    For x_num = ii - 1 To 0 Step -1
                        x_stepsize = Abs(X_Axis_TrackingParaFrom(x_num) - X_Axis_TrackingParaTo(x_num)) / (Xaxis_index - 1)
                        Call ValueResolution_1(X_Axis_TrackingParaFrom(x_num), X_Axis_TrackingParaTo(x_num), x_stepsize)
                        x_TrackingInfo = X_Axis_TrackingPara(x_num) & "," & X_Axis_TrackingParaFrom(x_num) & "," & X_Axis_TrackingParaTo(x_num) & "," & x_stepsize & ";" & x_TrackingInfo
                    Next x_num
                    TheExec.Datalog.WriteComment " X-Asix TraickingPin Info => " & x_TrackingInfo
                End If
            Case 1
                ReDim Y_Axis_Val(Yaxis_index - 1)
                If Y_Tracking_Point <> 0 Then 'extract Y Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Y_Axis_TrackingPara(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaFrom(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaTo(Y_Tracking_Point - 1)
                    With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .List
                            Y_Axis_TrackingPara(jj) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).ApplyTo.Pins), "_", "")
                            Y_Axis_TrackingParaFrom(jj) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.from
                            Y_Axis_TrackingParaTo(jj) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.To
                            If Y_Axis_TrackingPara(jj) = "" Then Y_Axis_TrackingPara(jj) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).Name), "_", "")
                            jj = jj + 1
                        Next Tracking_Item
                    End With
                    For y_num = jj - 1 To 0 Step -1
                        y_stepsize = Abs(Y_Axis_TrackingParaFrom(y_num) - Y_Axis_TrackingParaTo(y_num)) / (Yaxis_index - 1)
                        Call ValueResolution_1(Y_Axis_TrackingParaFrom(y_num), Y_Axis_TrackingParaTo(y_num), y_stepsize)
                        y_TrackingInfo = Y_Axis_TrackingPara(y_num) & "," & Y_Axis_TrackingParaFrom(y_num) & "," & Y_Axis_TrackingParaTo(y_num) & "," & y_stepsize & ";" & y_TrackingInfo
                    Next y_num
                    TheExec.Datalog.WriteComment " Y-Asix TraickingPin Info => " & y_TrackingInfo
                End If
            Case 2
                ReDim Z_Axis_Val(Zaxis_index - 1)
                If Z_Tracking_Point <> 0 Then 'extract Z Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Z_Axis_TrackingPara(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaFrom(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaTo(Z_Tracking_Point - 1)
                    With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .List
                            Z_Axis_TrackingPara(kk) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).ApplyTo.Pins), "_", "")
                            Z_Axis_TrackingParaFrom(kk) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.from
                            Z_Axis_TrackingParaTo(kk) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.To
                            If Z_Axis_TrackingPara(kk) = "" Then Z_Axis_TrackingPara(kk) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).Name), "_", "")
                            kk = kk + 1
                        Next Tracking_Item
                    End With
                    For z_num = kk - 1 To 0 Step -1
                        z_stepsize = Abs(Z_Axis_TrackingParaFrom(z_num) - Z_Axis_TrackingParaTo(z_num)) / (Zaxis_index - 1)
                        Call ValueResolution_1(Z_Axis_TrackingParaFrom(z_num), Z_Axis_TrackingParaTo(z_num), z_stepsize)
                        z_TrackingInfo = Z_Axis_TrackingPara(z_num) & "," & Z_Axis_TrackingParaFrom(z_num) & "," & Z_Axis_TrackingParaTo(z_num) & "," & z_stepsize & ";" & z_TrackingInfo
                    Next z_num
                    TheExec.Datalog.WriteComment " Z-Asix TraickingPin Info => " & z_TrackingInfo
                End If
        End Select
        axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).ApplyTo.Pins), "_", "")
        If axis_pin(curr_axis) = "" Then
            axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).Parameter.Name), "_", "")
        End If
        
Next curr_axis

InstName = TheExec.DataManager.instanceName
tmpStr = Split(InstName, "_")

tmpStr1 = Split(InstName, "_")
TmpVal = InStr(tmpStr1(0), "L")
tmpStr1(0) = Mid(tmpStr1(0), 1, TmpVal) '  "DFTL" H or "MCL" H
InstName_L = Join(tmpStr1, "_")

tmpStr2 = Split(InstName, "_")
tmpStr2(0) = Mid(tmpStr2(0), 1, TmpVal - 1) & Right(tmpStr2(0), 1) '  H
InstName_H = Join(tmpStr2, "_")

k = 0
j = 0
ShmooResult = ""

'''''------------------------------------------------------------------
''''Dim Z_Axis_Val_w() As New SiteDouble
''''Dim ShmooResult_w As New SiteVariant
''''Dim x_val_w As New SiteVariant
''''Dim y_val_w As New SiteVariant
''''Dim z_val_w As New SiteVariant
''''Dim LVCC_w As New SiteDouble
''''Dim HVCC_w As New SiteDouble
''''
''''ReDim Z_Axis_Val_w(Zaxis_index - 1)
''''
''''For i = 0 To MaxArrIndex - 1
''''    For Each site In theexec.sites
''''        Z_Axis_Val_w(j) = g_ShmooResult.Axis_CurrPoint(i).Z_axis(site)
''''        j = j + 1
''''        ShmooResult_w = ShmooResult_w & g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)
''''        x_val_w = g_ShmooResult.Axis_CurrPoint(i).X_axis(site)
''''        y_val_w = g_ShmooResult.Axis_CurrPoint(i).Y_axis(site)
''''        z_val_w = g_ShmooResult.Axis_CurrPoint(i).Z_axis(site)
''''        Call ValueResolution_2(x_val_w, y_val_w, z_val_w)
'''''        Call ShmooResultPF_1(ShmooResult_w, LVCC_w, HVCC_w, Z_Axis_Val_w)
''''    Next site
''''        If i = (Zaxis_index - 1) + k Then
''''            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val_w & "_" & CStr(axis_pin(1)) & y_val_w & "_" & CStr(axis_pin(2)) & z_val_w & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
''''            ShmooResult_w = ""
''''            k = k + Zaxis_index
''''            j = 0
''''        Else
''''            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val_w & "_" & CStr(axis_pin(1)) & y_val_w & "_" & CStr(axis_pin(2)) & z_val_w & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
''''        End If
'''''    Next site
''''Next i



''  =============================================================================================================================================
''  Z-axis
tmp_Tnum = g_TestNum
For Each site In TheExec.sites
    If (X_dimemsion = False And Y_dimemsion = False And Z_dimemsion = False) Or Z_dimemsion = True Then
        TheExec.Datalog.WriteComment "  [ Z-Axis ] "
        For i = 0 To MaxArrIndex - 1
            Z_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(i).Z_axis(site)
            j = j + 1
            ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)
            x_val = g_ShmooResult.Axis_CurrPoint(i).X_axis(site)
            y_val = g_ShmooResult.Axis_CurrPoint(i).Y_axis(site)
            z_val = g_ShmooResult.Axis_CurrPoint(i).Z_axis(site)
            Call ValueResolution_1(x_val, y_val, z_val)
    '        If i = 0 Or i <> Zaxis_index Then
            If i = (Zaxis_index - 1) + k Then
    '            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                Call ShmooResultPF(ShmooResult, LVCC, HVCC, Z_Axis_Val)
                Select Case ShmooResult
                    Case "9999"
                        ' still print out the latest point
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        If UCase(Right(tmpStr(0), 2)) = "LH" Then
                        
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                           Else
                                TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        End If
                    Case "7777"
                        ' still print out the latest point
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        If UCase(Right(tmpStr(0), 2)) = "LH" Then
                            
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        End If
                    Case "5555"
                        ' still print out the latest point
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        If UCase(Right(tmpStr(0), 2)) = "LH" Then
                            
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                           Else
                                TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            End If
                        End If
                    Case Else ' search LVCC/HVCC point
                        ' still print out the latest point
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                        If UCase(Right(tmpStr(0), 2)) = "LH" Then
                        
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                                TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                               TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                           End If
                        ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                 g_TestNum = g_TestNum + 1
                           End If
                        ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                            If RangeSeq(2) = True Then 'Z-axis small--->large
                                TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Z_Axis_Val(UBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(LBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                            Else
                                TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Z_Axis_Val(LBound(Z_Axis_Val)), lowVal:=Z_Axis_Val(UBound(Z_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & "result", TNum:=g_TestNum
                                g_TestNum = g_TestNum + 1
                           End If
                        End If
                End Select
    '            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & " LVCC or HVCC or shmoo hole "
                ShmooResult = ""
                k = k + Zaxis_index
                j = 0
            Else
                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
            End If
            
        Next i
        k = 0 ' initial it fot multi site
    End If
''  =============================================================================================================================================
''  X-axis
    If X_dimemsion = True Then
        TheExec.Datalog.WriteComment "  [ X-Axis ] "
        Dim xx As Integer
        Dim tmp_count1 As Integer
        Dim tmp_x As Integer
        'xx = 2
        tmp_count1 = 0
        For xx = 1 To Zaxis_index
            For i = tmp_count1 To MaxArrIndex * Zaxis_index - 1 Step Zaxis_index
                If i > xx * MaxArrIndex - 1 Then
                    tmp_count1 = i
                    Exit For
                Else
                    tmp_x = i - (xx - 1) * (MaxArrIndex - 1)
                    ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site)
                    X_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(tmp_x).X_axis(site)
                    j = j + 1
                    x_val = g_ShmooResult.Axis_CurrPoint(tmp_x).X_axis(site)
                    y_val = g_ShmooResult.Axis_CurrPoint(tmp_x).Y_axis(site)
                    z_val = g_ShmooResult.Axis_CurrPoint(tmp_x).Z_axis(site)
                    Call ValueResolution_1(x_val, y_val, z_val)
                End If
        'Debug.Print tmp_x
                If j = Xaxis_index Then
                    
                    Call ShmooResultPF(ShmooResult, LVCC, HVCC, X_Axis_Val)
                    Select Case ShmooResult
                        Case "9999"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
        
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "7777"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "5555"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case Else ' search LVCC/HVCC point
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            End If
                    End Select
        '''''            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & " LVCC or HVCC or shmoo hole "
                    ShmooResult = ""
                    j = 0
                Else
                    If i > MaxArrIndex - 1 Then
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
                    Else
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                    End If
                End If
        
            Next i
        '    i = tmp_count1
        Next xx
        k = 0 ' initial it fot multi site
    End If
'''''  =============================================================================================================================================
'''''  Y-axis
    If Y_dimemsion = True Then
        TheExec.Datalog.WriteComment "  [ Y-Axis ] "
        Dim yy As Integer
        Dim tmp_count As Integer
        Dim tmp_y As Integer
        'yy = 2
        tmp_count = 0
        For yy = 1 To Zaxis_index * Xaxis_index
            For i = tmp_count To (MaxArrIndex * Zaxis_index * Xaxis_index) - 1 Step (Zaxis_index * Xaxis_index)
        
                If i > yy * MaxArrIndex - 1 Then
                    tmp_count = i
                    Exit For
                Else
                    tmp_y = i - (yy - 1) * (MaxArrIndex - 1)
                    ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)
                    Y_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(tmp_y).Y_axis(site)
                    j = j + 1
                    x_val = g_ShmooResult.Axis_CurrPoint(tmp_y).X_axis(site)
                    y_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Y_axis(site)
                    z_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Z_axis(site)
                    Call ValueResolution_1(x_val, y_val, z_val)
                End If
        'Debug.Print tmp_y
                If j = Yaxis_index Then
        '            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                    Call ShmooResultPF(ShmooResult, LVCC, HVCC, Y_Axis_Val)
                    Select Case ShmooResult
                        Case "9999"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
        
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "7777"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "5555"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case Else ' search LVCC/HVCC point
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                    End Select
        ''''            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & " LVCC or HVCC or shmoo hole "
                    ShmooResult = ""
                    j = 0
                Else
                    If i > MaxArrIndex - 1 Then
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                    Else
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                    End If
        '            TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                End If
        
            Next i
        Next yy
        k = 0 ' initial it fot multi site
        '''
        '''
    End If
    g_TestNum = tmp_Tnum
Next site

End Function


Public Function Print2DShmooInfo(argc As Long, argv() As String)
Dim site As Variant
Dim i, j, k As Long
Dim ShmooResult As String
Dim x_val As Variant
Dim y_val As Variant
Dim z_val As Variant
Dim LVCC As Double
Dim HVCC As Double
Dim InstName As String
Dim curr_axis As Variant
Dim axis_pin(2) As String
Dim X_Axis_Val() As Double
Dim Y_Axis_Val() As Double
Dim Z_Axis_Val() As Double
Dim DevSetupName As String

Dim tmpStr() As String
Dim tmpStr1() As String
Dim tmpStr2() As String
Dim InstName_H As String
Dim InstName_L As String
Dim TmpVal As Integer
Dim X_Axis_TrackingPara() As String
Dim Y_Axis_TrackingPara() As String
Dim Z_Axis_TrackingPara() As String
Dim X_Axis_TrackingParaFrom() As Variant
Dim Y_Axis_TrackingParaFrom() As Variant
Dim Z_Axis_TrackingParaFrom() As Variant
Dim X_Axis_TrackingParaTo() As Variant
Dim Y_Axis_TrackingParaTo() As Variant
Dim Z_Axis_TrackingParaTo() As Variant
Dim Tracking_Item As Variant
Dim ii, jj, kk As Integer

Dim x_TrackingInfo As String
Dim x_stepsize As Variant
Dim x_num As Integer
Dim y_TrackingInfo As String
Dim y_stepsize As Variant
Dim y_num As Integer
Dim z_TrackingInfo As String
Dim z_stepsize As Variant
Dim z_num As Integer
Dim tmp_Tnum As Long

ii = 0
jj = 0
kk = 0
''Exit Function
DevSetupName = TheExec.DevChar.Setups.ActiveSetupName
For Each curr_axis In TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes.List
        Select Case curr_axis
            Case 0
                ReDim X_Axis_Val(Xaxis_index - 1)
                If X_Tracking_Point <> 0 Then 'extract X Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim X_Axis_TrackingPara(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaFrom(X_Tracking_Point - 1)
                    ReDim X_Axis_TrackingParaTo(X_Tracking_Point - 1)
                    With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .List
                            X_Axis_TrackingPara(ii) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).ApplyTo.Pins), "_", "")
                            X_Axis_TrackingParaFrom(ii) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.from
                            X_Axis_TrackingParaTo(ii) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.To
                            If X_Axis_TrackingPara(ii) = "" Then X_Axis_TrackingPara(ii) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).Name), "_", "")
                            ii = ii + 1
                        Next Tracking_Item
                    End With
                    For x_num = ii - 1 To 0 Step -1
                        x_stepsize = Abs(X_Axis_TrackingParaFrom(x_num) - X_Axis_TrackingParaTo(x_num)) / (Xaxis_index - 1)
                        Call ValueResolution_1(X_Axis_TrackingParaFrom(x_num), X_Axis_TrackingParaTo(x_num), x_stepsize)
                        x_TrackingInfo = X_Axis_TrackingPara(x_num) & "," & X_Axis_TrackingParaFrom(x_num) & "," & X_Axis_TrackingParaTo(x_num) & "," & x_stepsize & ";" & x_TrackingInfo
                    Next x_num
                    TheExec.Datalog.WriteComment " X-Asix TraickingPin Info => " & x_TrackingInfo
                End If
            Case 1
                ReDim Y_Axis_Val(Yaxis_index - 1)
                If Y_Tracking_Point <> 0 Then 'extract Y Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Y_Axis_TrackingPara(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaFrom(Y_Tracking_Point - 1)
                    ReDim Y_Axis_TrackingParaTo(Y_Tracking_Point - 1)
                    With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .List
                            Y_Axis_TrackingPara(jj) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).ApplyTo.Pins), "_", "")
                            Y_Axis_TrackingParaFrom(jj) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.from
                            Y_Axis_TrackingParaTo(jj) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.To
                            If Y_Axis_TrackingPara(jj) = "" Then Y_Axis_TrackingPara(jj) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).Name), "_", "")
                            jj = jj + 1
                        Next Tracking_Item
                    End With
                    For y_num = jj - 1 To 0 Step -1
                        y_stepsize = Abs(Y_Axis_TrackingParaFrom(y_num) - Y_Axis_TrackingParaTo(y_num)) / (Yaxis_index - 1)
                        Call ValueResolution_1(Y_Axis_TrackingParaFrom(y_num), Y_Axis_TrackingParaTo(y_num), y_stepsize)
                        y_TrackingInfo = Y_Axis_TrackingPara(y_num) & "," & Y_Axis_TrackingParaFrom(y_num) & "," & Y_Axis_TrackingParaTo(y_num) & "," & y_stepsize & ";" & y_TrackingInfo
                    Next y_num
                    TheExec.Datalog.WriteComment " Y-Asix TraickingPin Info => " & y_TrackingInfo
                End If
            Case 2
                ReDim Z_Axis_Val(Zaxis_index - 1)
                If Z_Tracking_Point <> 0 Then 'extract Z Axis tracking pin, either ApplyToPins or ParameterName
                    ReDim Z_Axis_TrackingPara(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaFrom(Z_Tracking_Point - 1)
                    ReDim Z_Axis_TrackingParaTo(Z_Tracking_Point - 1)
                    With TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters
                        For Each Tracking_Item In .List
                            Z_Axis_TrackingPara(kk) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).ApplyTo.Pins), "_", "")
                            Z_Axis_TrackingParaFrom(kk) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.from
                            Z_Axis_TrackingParaTo(kk) = TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).range.To
                            If Z_Axis_TrackingPara(kk) = "" Then Z_Axis_TrackingPara(kk) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes(curr_axis).TrackingParameters.Item(Tracking_Item).Name), "_", "")
                            kk = kk + 1
                        Next Tracking_Item
                    End With
                    For z_num = kk - 1 To 0 Step -1
                        z_stepsize = Abs(Z_Axis_TrackingParaFrom(z_num) - Z_Axis_TrackingParaTo(z_num)) / (Zaxis_index - 1)
                        Call ValueResolution_1(Z_Axis_TrackingParaFrom(z_num), Z_Axis_TrackingParaTo(z_num), z_stepsize)
                        z_TrackingInfo = Z_Axis_TrackingPara(z_num) & "," & Z_Axis_TrackingParaFrom(z_num) & "," & Z_Axis_TrackingParaTo(z_num) & "," & z_stepsize & ";" & z_TrackingInfo
                    Next z_num
                    TheExec.Datalog.WriteComment " Z-Asix TraickingPin Info => " & z_TrackingInfo
                End If
        End Select
        axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).ApplyTo.Pins), "_", "")
        If axis_pin(curr_axis) = "" Then
            axis_pin(curr_axis) = Replace((TheExec.DevChar.Setups(DevSetupName).Shmoo.Axes((curr_axis)).Parameter.Name), "_", "")
        End If
        
Next curr_axis

InstName = TheExec.DataManager.instanceName
tmpStr = Split(InstName, "_")

tmpStr1 = Split(InstName, "_")
TmpVal = InStr(tmpStr1(0), "L")
tmpStr1(0) = Mid(tmpStr1(0), 1, TmpVal) '  "DFTL" H or "MCL" H
InstName_L = Join(tmpStr1, "_")

tmpStr2 = Split(InstName, "_")
tmpStr2(0) = Mid(tmpStr2(0), 1, TmpVal - 1) & Right(tmpStr2(0), 1) '  H
InstName_H = Join(tmpStr2, "_")

k = 0
j = 0
ShmooResult = ""


''  =============================================================================================================================================
''  X-axis
tmp_Tnum = g_TestNum
For Each site In TheExec.sites
    If X_dimemsion = True Then
        TheExec.Datalog.WriteComment "  [ X-Axis ] "
        Dim xx As Integer
        Dim tmp_count1 As Integer
        Dim tmp_x As Integer
        'xx = 2
        tmp_count1 = 0
'        For xx = 1 To Zaxis_index
            For i = tmp_count1 To MaxArrIndex - 1
'                If i > xx * MaxArrIndex - 1 Then
'                    tmp_count1 = i
'                    Exit For
'                Else
'                    tmp_x = i - (xx - 1) * (MaxArrIndex - 1)
                    ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(i).CurrResult(site)
                    X_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(i).X_axis(site)
                    j = j + 1
                    x_val = g_ShmooResult.Axis_CurrPoint(i).X_axis(site)
                    y_val = g_ShmooResult.Axis_CurrPoint(i).Y_axis(site)
                    Call ValueResolution_1(x_val, y_val, z_val)
'                End If
        'Debug.Print tmp_x
                If j = Xaxis_index Then
                    
                    Call ShmooResultPF(ShmooResult, LVCC, HVCC, X_Axis_Val)
                    Select Case ShmooResult
                        Case "9999"
                            ' still print out the latest point
'                            If i > MaxArrIndex - 1 Then
'                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
'                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
'                            End If
        
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Z_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "7777"
                            ' still print out the latest point
'                            If i > MaxArrIndex - 1 Then
'                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
'                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
'                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "5555"
                            ' still print out the latest point
'                            If i > MaxArrIndex - 1 Then
'                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
'                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
'                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case Else ' search LVCC/HVCC point
                            ' still print out the latest point
'                            If i > MaxArrIndex - 1 Then
'                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
'                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
'                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=X_Axis_Val(UBound(X_Axis_Val)), lowVal:=X_Axis_Val(LBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=X_Axis_Val(LBound(X_Axis_Val)), lowVal:=X_Axis_Val(UBound(X_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val, TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            End If
                    End Select
        '''''            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & " LVCC or HVCC or shmoo hole "
                    ShmooResult = ""
                    j = 0
                Else
'                    If i > MaxArrIndex - 1 Then
'                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_x).CurrResult(site))
'                    Else
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
'                    End If
                End If
        
            Next i
        '    i = tmp_count1
'        Next xx
        k = 0 ' initial it fot multi site
    End If
'''''  =============================================================================================================================================
'''''  Y-axis
    If Y_dimemsion = True Then
        TheExec.Datalog.WriteComment "  [ Y-Axis ] "
        Dim yy As Integer
        Dim tmp_count As Integer
        Dim tmp_y As Integer
        'yy = 2
        tmp_count = 0
        For yy = 1 To Xaxis_index
            For i = tmp_count To (MaxArrIndex * Xaxis_index) - 1 Step Xaxis_index
        
                If i > yy * MaxArrIndex - 1 Then
                    tmp_count = i
                    Exit For
                Else
                    tmp_y = i - (yy - 1) * (MaxArrIndex - 1)
                    ShmooResult = ShmooResult & g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site)
                    Y_Axis_Val(j) = g_ShmooResult.Axis_CurrPoint(tmp_y).Y_axis(site)
                    j = j + 1
                    x_val = g_ShmooResult.Axis_CurrPoint(tmp_y).X_axis(site)
                    y_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Y_axis(site)
                    z_val = g_ShmooResult.Axis_CurrPoint(tmp_y).Z_axis(site)
                    Call ValueResolution_1(x_val, y_val, z_val)
                End If
        'Debug.Print tmp_y
                If j = Yaxis_index Then
        '            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                    Call ShmooResultPF(ShmooResult, LVCC, HVCC, Y_Axis_Val)
                    Select Case ShmooResult
                        Case "9999"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
        
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                               End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-9999, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "7777"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-7777, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case "5555"
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=-5555, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                        Case Else ' search LVCC/HVCC point
                            ' still print out the latest point
                            If i > MaxArrIndex - 1 Then
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                            Else
                                TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                            End If
                            If UCase(Right(tmpStr(0), 2)) = "LH" Then
        
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_L & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName_H & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "H" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=HVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result" & "_", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            ElseIf UCase(Right(tmpStr(0), 1)) = "L" Then
                                If RangeSeq(0) = True Then 'X-axis small--->large
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Y_Axis_Val(UBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(LBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                Else
                                    TheExec.Flow.TestLimit resultVal:=LVCC, hiVal:=Y_Axis_Val(LBound(Y_Axis_Val)), lowVal:=Y_Axis_Val(UBound(Y_Axis_Val)), Tname:=InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & "result", TNum:=g_TestNum
                                    g_TestNum = g_TestNum + 1
                                End If
                            End If
                    End Select
        ''''            theexec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & "result" & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & " LVCC or HVCC or shmoo hole "
                    ShmooResult = ""
                    j = 0
                Else
                    If i > MaxArrIndex - 1 Then
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(tmp_y).CurrResult(site))
                    Else
                        TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                    End If
        '            TheExec.Datalog.WriteComment "Site(" & CStr(site) & ") : " & InstName & "_" & CStr(axis_pin(0)) & x_val & "_" & CStr(axis_pin(1)) & y_val & "_" & CStr(axis_pin(2)) & z_val & " " & CStr(g_ShmooResult.Axis_CurrPoint(i).CurrResult(site))
                End If
        
            Next i
        Next yy
        k = 0 ' initial it fot multi site
        '''
        '''
    End If
    g_TestNum = tmp_Tnum
Next site

End Function



Public Function ShmooResultPF(ShmRes As String, LVCC As Double, HVCC As Double, ShmooEachPoint() As Double)
'Public Function ShmooResultPF(ShmRes As String, LVCC As Double, HVCC As Double)

Dim strlen As Long
Dim i, j, k As Long
Dim Count As Long
Dim LocResult As String
Dim PreviousResult As String
Dim PreviousResult_tmp As String
Dim FPPoint(1000) As Integer
Dim FPCount As Integer
Dim PFPoint(1000) As Integer
Dim PFCount As Integer
Dim FirstPassLoc As Integer
Dim FirstFailLoc As Integer
Dim FPTransit As Integer
Dim PFTransit As Integer
'Dim LVCC As Double
Dim FPTransit_1 As Integer
Dim PFTransit_1 As Integer
Dim Entry_Flag As Boolean
Dim FailToPassCounr As Integer


    PFCount = 0
    FPCount = 0
    FirstPassLoc = -1
    FirstFailLoc = -1
    FPTransit = -1
    PFTransit = -1
    FPTransit_1 = -1
    PFTransit_1 = -1
    Count = 0
    Entry_Flag = True
    FailToPassCounr = 0
    k = InStr(ShmRes, "P") ' for FFPPFF case


''Dim ShmooEachPoint(5) As Double
''ShmooEachPoint(0) = 1
''ShmooEachPoint(1) = 1.2
''ShmooEachPoint(2) = 1.4
''ShmooEachPoint(3) = 1.6
''ShmooEachPoint(4) = 1.8
''ShmooEachPoint(5) = 2

''ShmooEachPoint(0) = 2
''ShmooEachPoint(1) = 1.8
''ShmooEachPoint(2) = 1.6
''ShmooEachPoint(3) = 1.4
''ShmooEachPoint(4) = 1.2
''ShmooEachPoint(5) = 1

If InStr(ShmRes, "A") Then 'Alarm
    ShmRes = "7777"
ElseIf InStr(ShmRes, "E") Then 'Error
    ShmRes = "7777"
Else
''    If LVCC_flag = True Then
        strlen = Len(ShmRes)
        For i = 1 To strlen
            LocResult = Mid(ShmRes, i, 1)
            If (i = 1) Then
                PreviousResult = LocResult
                PreviousResult_tmp = LocResult
            Else
                If (PreviousResult <> LocResult) Then
                    If (PreviousResult = "P") Then
                        PFPoint(PFCount) = i
                        PFCount = PFCount + 1
                    Else
                        FPPoint(FPCount) = i
                        FPCount = FPCount + 1
                    End If
                    PreviousResult = LocResult
                End If
            End If
            
            If (LocResult = "P") Then
                If (FirstPassLoc = -1) Then
                    FirstPassLoc = i - 1
                End If
                If (FirstFailLoc <> -1 And PFTransit = -1) Then
                    PFTransit = i - 1
                End If
                ''===============================================
                '' only looking for this scenario "FFPPFF"
'                PreviousResult_tmp = LocResult
'                If FirstFailLoc <> -1 And Entry_Flag = True Then
''                If FirstFailLoc <> -1 And (PreviousResult_tmp <> LocResult) Then
'''                    FPTransit_1 = FPTransit_1 + 1
'''                    Entry_Flag = False
''                    FailToPassCounr = FailToPassCounr + 1 ' if FailToPassCounr > 1 means the result has shmoo hole
''                End If
'                PreviousResult_tmp = LocResult
                ''===============================================
            End If
            
            If (LocResult = "F") Then
                If (FirstFailLoc = -1) Then
                    FirstFailLoc = i - 1
                End If
                If (FirstPassLoc <> -1 And FPTransit = -1) Then
                    FPTransit = i - 1
                End If
                ''===============================================
                '' only looking for this scenario "FFPPFF"
'                If FirstPassLoc <> -1 And (PreviousResult_tmp <> LocResult) Then
'                    FailToPassCounr = FailToPassCounr + 1 ' if FailToPassCounr > 1 means the result has shmoo hole
'                End If
'                If FirstPassLoc <> -1 And FailToPassCounr = 0 Then
'                    PFTransit_1 = PFTransit_1 + 1
'                    Entry_Flag = True
'                End If
'                PreviousResult_tmp = LocResult

                ''===============================================
                
           End If
           
            If Left(ShmRes, 1) = "F" Then ' only allow first point is "F"
                If Mid(ShmRes, k + 1, 1) = "P" Then ' skip  FFFFF 'P' FFFFF case
                    If FirstPassLoc <> -1 And (PreviousResult_tmp <> LocResult) Then
                        PreviousResult_tmp = LocResult
                        FailToPassCounr = FailToPassCounr + 1 '
                    End If
                End If
            End If
        Next i
    
        If (PFTransit = -1 And FPTransit = -1 And FirstPassLoc <> -1 And FirstFailLoc = -1) Then ' All PASS
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                LVCC = ShmooEachPoint(LBound(ShmooEachPoint))
                HVCC = ShmooEachPoint(UBound(ShmooEachPoint))
            Else
                LVCC = ShmooEachPoint(UBound(ShmooEachPoint))
                HVCC = ShmooEachPoint(LBound(ShmooEachPoint))
            End If
        ElseIf (PFTransit = -1 And FPTransit = -1 And FirstPassLoc = -1 And FirstFailLoc <> -1) Then ' All FAIL
'            LVCC = -9999
'            HVCC = 9999
'            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                ShmRes = "9999"
'            Else
'                ShmRes = "-9999"
'            End If
        ElseIf (PFTransit <> -1 And FPTransit = -1 And FirstPassLoc <> -1 And FirstFailLoc <> -1) Then ' Fail-Pass transition point
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                LVCC = ShmooEachPoint(FPPoint(0) - 1)
                HVCC = ShmooEachPoint(UBound(ShmooEachPoint))
            Else
                LVCC = ShmooEachPoint(UBound(ShmooEachPoint))
                HVCC = ShmooEachPoint(FPPoint(0) - 1)
            End If
        ElseIf (PFTransit = -1 And FPTransit <> -1 And FirstPassLoc <> -1 And FirstFailLoc <> -1) Then ' Pass-Fail transition point
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                LVCC = ShmooEachPoint(0)
                HVCC = ShmooEachPoint(PFPoint(0) - 2)
            Else
                LVCC = ShmooEachPoint(PFPoint(0) - 2)
                HVCC = ShmooEachPoint(0)
            End If
        ElseIf (PFTransit <> -1 And FPTransit <> -1 And FirstPassLoc <> -1 And FirstFailLoc <> -1 And FailToPassCounr <> 2) Then ' Shmoo hole, ex: PPPFFFPPP
'            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                ShmRes = "5555"
'            Else
'                ShmRes = "-5555"
'            End If
        ElseIf (FailToPassCounr = 2) Then '  ex: FFPPPFFF
            If ShmooEachPoint(UBound(ShmooEachPoint)) > ShmooEachPoint(LBound(ShmooEachPoint)) Then
                LVCC = ShmooEachPoint(FPPoint(0) - 1)
                HVCC = ShmooEachPoint(PFPoint(0) - 2)
            Else
                LVCC = ShmooEachPoint(PFPoint(0) - 2)
                HVCC = ShmooEachPoint(FPPoint(0) - 1)
            End If
        End If
    
End If


End Function

Public Function Decide_Switching_Bit_Debug_LVCC(digSrc_EQ As String, DSPWaveSwitch As DSPWave, DC_Level As PinListData, BlockType As String, SELSRM_Rails As String, Optional shmoo_pin As String, Optional ShmooPinsVoltage As PinListData, Optional ForcePin As String, Optional SetForceVoltage As Dictionary) As String

  Dim logicPin As String
  Dim SramPin As String
  Dim DSSC_Switching_Voltage As New PinListData
  Dim Sdomain As Long
  Dim DSSCSelSrmOpposite As Long
  Dim BlockTypeNum As Long
  Dim i As Integer
  Dim ReturnString() As String
  On Error GoTo errHandler
  BlockTypeNum = -1
  
  ReDim ReturnString(Len(digSrc_EQ) - 1)
  Decide_DSSC_Switching_Voltage DSSC_Switching_Voltage, DC_Level, shmoo_pin, ShmooPinsVoltage, ForcePin, SetForceVoltage
  For i = 0 To UBound(GetSelSram.Block)
     If UCase(GetSelSram.Block(i).DomainName) <> "" Then
       If UCase(BlockType) Like "*" & UCase(GetSelSram.Block(i).DomainName) & "*" Then
          BlockTypeNum = i
          Exit For
       End If
     End If
  Next i
  
  If BlockTypeNum <> -1 Then
    For i = 0 To Len(digSrc_EQ) - 1
     If UCase(CStr(Mid(digSrc_EQ, i + 1, 1))) Like "S" Then
         logicPin = GetSelSram.Block(BlockTypeNum).DomainBits(i).logicPin
         SramPin = GetSelSram.Block(BlockTypeNum).DomainBits(i).SramPin
         DSSCSelSrmOpposite = GetSelSram.Block(BlockTypeNum).DomainBits(i).SelSram1
            If DSSCSelSrmOpposite = 0 Then
               Sdomain = IIf((DSSC_Switching_Voltage.Pins(logicPin).Value > DSSC_Switching_Voltage.Pins(SramPin).Value), 1, 0)
               DSPWaveSwitch.Element(i) = Sdomain
               ReturnString(i) = Sdomain
            ElseIf DSSCSelSrmOpposite = 1 Then
               Sdomain = IIf((DSSC_Switching_Voltage.Pins(logicPin).Value > DSSC_Switching_Voltage.Pins(SramPin).Value), 0, 1)
               DSPWaveSwitch.Element(i) = Sdomain
               ReturnString(i) = Sdomain
            End If
      Else
          DSPWaveSwitch.Element(i) = CDbl(Mid(digSrc_EQ, i + 1, 1))
          ReturnString(i) = CDbl(Mid(digSrc_EQ, i + 1, 1))
      End If
    Next i

    Decide_Switching_Bit_Debug_LVCC = Join(ReturnString, "")
    SELSRM_Rails = DecodingRealSourceBit(Decide_Switching_Bit_Debug_LVCC, BlockType)

  End If
  Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + "Decide_Switching_Bit" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

