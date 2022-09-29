Attribute VB_Name = "VBT_LIB_TestLimit"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

Public Const gC_Sheet_Job = "JobList"
Public Const gC_ProjectLimitSheet = "QQ_LimitSheet"
Public Const gC_Sheet_CurrentLimit = "CurrentLimit" 'Dump All Limit for current test.
Public Const gC_Sheet_UpdateLimit = "UpdateLimit"   'Dump the Limit which need to be updated.
Public Const gC_Final_PM_Want_LimitSheet = "MainLimitSheet"
Public Const gb_TNameTNum_Dupl_CHK = False  '20190222

''--------------------------------------------------------
Public Const sLimitSheet_MP_DataLog = "_FinalTrim_,_FinalTrimCode_,_PostBurn_,_PostBurnCode_" '2019/12/06 Control MP keyWord
Public g_sLimitSheet_MP_DataLog_array()         As String  '2019/12/06 Control MP keyWord


'====================================================================================================================================================
'History:
'2018/04/03: LimitTable UDT for TTR (1st Run :150ms/2nd Run:16ms) & UserTName
'2018/04/09: Add LL.ByTNum_GetLimitDetailsFromLimitTable for Buck3
'====================================================================================================================================================
'A.Usage:[LL.FlowTestLimit]
'====================================================================================================================================================
'LibLimitSet:  bEnableCurrentSheet = True:bdbgIsExistTestName = True: bEnableLogAllJob =True   [Take a long time]
'LibLimitSet:  bEnableCurrentSheet = True:bdbgIsExistTestName = False: bEnableLogAllJob =False
'Private Const bEnableCurrentSheet = True  'bEnableCurrentSheet=Flase: If you don't want to log into CurrentSheet
'Private Const bdbgIsExistTestName = False 'bdbgIsExistTestName=True: Check all item include Limit="N/A" & TName="*CODE""
'Private Const bEnableLogAllJob = False    'bEnableLogAllJob=True   : Update High/Low limit for all Job into CurrentSheet.
'----------------------------------------------------------
''Step0:Import  VB module & File.
'       Import 3 sheets  :  MP3TLimitSheet/ CurrentLimit/ UpdateLimit
'       Import 2 modules :LIB_TestLimit.bas/ LibLimitSet.cls
'       Put code here OnValidationStart:   Call LoadCurrentLimitSet
'       Put code here OnProgramStarted: Call SetCurrentLimitSet
''Step1:Select the module.
''Step2:Replace "TheExec.Flow.testlimit" with "LL.FlowTestLimit"
''Step3:Run.
'       Showing the TestName which are not existing in MP3TLimitSheet.
'       EX:TName =OPEN_AMUX-B0_X_X_X_X_-100UA_X_X_X_X_X  Is Not Existing In MP3TLimitSheet.
''Step4:Copy from "UpdateLimit" ,and paste into "MP3TLimits.
''Notice-1:  Please remember  Validate T/P again if  any change on "MP3TLimits��.
''Notice-2:  Should not have "N/A" in "MP3TLimits"

'====================================================================================================================================================
'B.Usage:[LL.ByTNum_GetLimitDetailsFromLimitTable]
'====================================================================================================================================================
'USING GLOBAL PARAMETERS:
'g_lTestNumber = 10140: g_lNumberOfLimitSets = 3 [MUST TO ASSIGN g_lTestNumber & g_lNumberOfLimitSets]
'LL.ByTNum_GetLimitDetailsFromLimitTable g_lTestNumber, g_strTestNames, g_dLowLimits, g_dHighLimits, g_stScale, g_utMeasUnit, g_lNumberOfLimitSets
'Call TheExec.Flow.TestLimit(ResultVal:=sdMeasValue(0), TName:=g_strTestNames(0), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Unit:=g_utMeasUnit(0), ForceResults:=tlForceNone, PinName:=plMeasPin, lowVal:=g_dLowLimits(0), hiVal:=g_dHighLimits(0))
'Call TheExec.Flow.TestLimit(ResultVal:=sdMeasValue(1), TName:=g_strTestNames(1), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Unit:=g_utMeasUnit(1), ForceResults:=tlForceNone, PinName:=plMeasPin, lowVal:=g_dLowLimits(1), hiVal:=g_dHighLimits(1))
'Call TheExec.Flow.TestLimit(ResultVal:=sdMeasValue(2), TName:=g_strTestNames(2), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Unit:=g_utMeasUnit(2), ForceResults:=tlForceNone, PinName:=plMeasPin, lowVal:=g_dLowLimits(2), hiVal:=g_dHighLimits(2))
'====================================================================================================================================================

'20180327 gtlForceNA mens limit is N/A
Public Enum gtlLimitForceResults
    gtlForceNone = 0
    gtlForcePass = 1
    gtlForceFail = 2
    gtlForceFlow = 3
    gtlForceNA = 4
End Enum

Public Enum gtlTestMode
    gtl_AEL = 0
    gtl_Checkermode = 1
    gtl_Development = 2
    gtl_Engineeringmode = 3
    gtl_Maintenancemode = 4
    gtl_Productiontestmode = 5
    gtl_QualityControl = 6
End Enum

Public Enum gColorIndex
    BLACKCOLOR = 1
    WHITECOLOR = 2
    REDCOLOR = 3
    LIGHTGREENCOLOR = 4
    BLUECOLOR = 5
    YELLOWCOLOR = 6
    PINKCOLOR = 7
    TurquoiseColor = 8
    DardRedColor = 9
    GREENCOLOR = 10
    DarkBuleColor = 11
    DarkYellowColor = 12
    VioletColor = 13
    TealColor = 14
    Gray25Color = 15
    Gray50Color = 16
    RoseColor = 38
    LightOrangeColor = 45
    ORANGECOLOR = 46
    BlueGrayColor = 47
    Gray40Color = 48
    DarkTealColor = 49
    SeaGreenColor = 50
    DarkGreenColor = 51
    OliveGreenColor = 52
    BrownColor = 53
    PlumColor = 54
    IndigoColor = 55
    Gray80Color = 56
End Enum

Public Enum gutTestlimit
    gut_ScaleType = 0
    gut_UnitType = 1
    gut_UnitCustomStr = 2
End Enum

Public Type LimitTableParamSyntax
    Row            As String
    FlowTable      As String
    TestName       As String
    TestNumber     As Long
    UserTName      As Variant
    HiLimit        As Variant
    LoLimit        As Variant
    Units          As Variant
    UScale         As Variant
End Type

Public Type LimitTableSyntax
    Index                 As Long
    TName                 As String
    TNum                  As Long
    CurrLimitSet          As String
    CurrLoLimit_Col       As Long
    CurrHiLimit_Col       As Long
    ParamSyntax           As LimitTableParamSyntax
End Type


Public LimitTable()             As LimitTableSyntax
Public gL_1st_LimitSheetRead    As Long
Public LL                       As New LibLimitSet

'
Public G_ForceResults           As gtlLimitForceResults
Public g_CurrentLimitSet        As String
Public g_CurrentLimitSheet      As String
Public gB_1stRun                As Boolean
Public gL_UpdateTest            As Long
Public gL_CurrTest              As Long
Public gL_JobNum                As Long
Public gS_JobList               As String
Public g_TestMode               As gtlTestMode
Public gL_UpdateTestMulti()            As Long
'Define for LL.ByTNum_GetLimitDetailsFromLimitTable
Public g_strTestNames()     As String
Public g_dLowLimits()       As Double
Public g_dHighLimits()      As Double
Public g_stScale()          As tlScaleType
Public g_utMeasUnit()       As UnitType
Public g_lTestNumber        As Long
Public g_lNumberOfLimitSets As Long
Public g_mTXT_input_JobList()   As String
Public g_mTXT_input_Header()    As String



' Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\WINDOWS\system32\scrrun.dll"
    
Public Sub SetCurrentLimitSet(Optional bDatalog As Boolean = True)
    On Error GoTo ErrHandler
    Dim funcName As String:: funcName = "SetCurrentLimitSet"
    Dim mS_dlogstr     As String
    Dim ReferenceTime  As Double
    Dim ElapsedTime    As Double
    Dim index_i As Long
    ReferenceTime = TheExec.Timer
    'Put this code on OnProgramStarted
    
    ReSetLimitSet
    gL_UpdateTest = 0: gL_CurrTest = 0

    ReDim gL_UpdateTestMulti(UBound(Split(gC_Sheet_UpdateLimit, ",")))
    For index_i = 0 To UBound(Split(gC_Sheet_UpdateLimit, ",")) Step 1
        gL_UpdateTestMulti(index_i) = 0
    Next index_i
    gS_JobList = LL.GetJobList(gC_Sheet_Job)
    gL_JobNum = UBound(Split(gS_JobList, ",")) + 1
    g_CurrentLimitSet = UCase(TheExec.CurrentJob)
    g_TestMode = TheExec.Datalog.Setup.LotSetup.TESTMODE
    
    'TheExec.Flow.Limits.Key = tlFlowLimitsKeyTestName

     'gL_1st_LimitSheetRead = 0
     LL.UpdateFlowLimit (gL_1st_LimitSheetRead)
     If (0) Then LL.CreatMainSheet_form_Multisubsheet (0)
    
     If bDatalog Then
        TheExec.Datalog.WriteComment vbCrLf & "Job = " & UCase(TheExec.CurrentJob) & ", CurrentLimitSheet =" & gC_ProjectLimitSheet
        TheExec.Datalog.WriteComment "Job = " & UCase(TheExec.CurrentJob) & ", CurrentLimitSet =" & g_CurrentLimitSet
        
        mS_dlogstr = "funcName = " & funcName & "::ElapsedTime =" & TheExec.Timer(ReferenceTime)
        TheExec.Datalog.WriteComment mS_dlogstr
        Debug.Print mS_dlogstr
     End If
     
     If gL_1st_LimitSheetRead = 0 Then gL_1st_LimitSheetRead = gL_1st_LimitSheetRead + 1
     If TheExec.Datalog.Setup.LotSetup.TESTMODE = 3 Then G_ForceResults = gtlForceFlow


''' 20191206 Add new MP_DataLog_ON flag judgement
    'Dim sLimitSheet_MP_DataLog_array()          As String
    Dim i As Long
    If TheExec.Flow.EnableWord("A_Enable_LimitSheet_Log_Precise") = True Then
        g_sLimitSheet_MP_DataLog_array = Split(sLimitSheet_MP_DataLog, ",")
        LL.MP_Datalog_FlagOff   ''' If method 1-> will control in main flow. If method 2-> this line need to set to FlagOn
    Else
        LL.MP_Datalog_FlagOff
    End If




    Exit Sub
ErrHandler:
     RunTimeError funcName
     If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub LoadCurrentLimitSet()
    Dim ws As New Worksheet
    Dim lIdx As Long
    Dim index_i As Long
    On Error GoTo ErrHandler
    Dim funcName As String:: funcName = "LoadCurrentLimitSet"
    Dim All_Sheet_arr() As String
    Dim Update_Sheet_arr() As String
    Dim Current_Sheet_arr() As String
    Static bFormatLimitSheet As Boolean
    
    Dim PorjectSheetName As String
'    ReDim All_Sheet_arr(UBound(Split(gC_ProjectLimitSheet, ",")))
    All_Sheet_arr = Split(gC_ProjectLimitSheet, ",")
    Update_Sheet_arr = Split(gC_Sheet_UpdateLimit, ",")
    Current_Sheet_arr = Split(gC_Sheet_CurrentLimit, ",")


    'Put this code on OnValidationStart
     'AutoLimits
    If UCase(TheExec.CurrentJob) Like "*LIMITS*" Then
        TheExec.Error.Behavior("FLO:202") = tlErrorIgnore
        TheExec.Flow.Limits.Load "AutoLimits.txt"
        Debug.Print "TheExec.Flow.Limits.Load AutoLimits.txt"
    End If
    gB_1stRun = True
    gL_1st_LimitSheetRead = 0
    
    gS_JobList = LL.GetJobList(gC_Sheet_Job)
    gL_JobNum = UBound(Split(gS_JobList, ",")) + 1
    
    If bFormatLimitSheet = True Then Exit Sub
    

'    LL.LimitSheetHeader gS_JobList, gC_Sheet_CurrentLimit
    For index_i = 0 To UBound(Split(gC_Sheet_CurrentLimit, ",")) Step 1
        PorjectSheetName = Current_Sheet_arr(index_i)
        If CheckSheet(PorjectSheetName) = False Then Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = PorjectSheetName
        LL.LimitSheetHeader gS_JobList, PorjectSheetName
    Next index_i

'    LL.LimitSheetHeader gS_JobList, gC_Sheet_UpdateLimit
    For index_i = 0 To UBound(Split(gC_Sheet_UpdateLimit, ",")) Step 1
        PorjectSheetName = Update_Sheet_arr(index_i)
        If CheckSheet(PorjectSheetName) = False Then Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = PorjectSheetName
        LL.LimitSheetHeader gS_JobList, PorjectSheetName
    Next index_i
    
'    LL.LimitSheetHeader gS_JobList, gC_ProjectLimitSheet
    For index_i = 0 To UBound(Split(gC_ProjectLimitSheet, ",")) Step 1
        PorjectSheetName = All_Sheet_arr(index_i)
        If CheckSheet(PorjectSheetName) = False Then Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = PorjectSheetName
        LL.LimitSheetHeader gS_JobList, PorjectSheetName
    Next index_i
    
    bFormatLimitSheet = True
     
    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub




 Public Sub ReSetLimitSet()
    On Error GoTo ErrHandler
    Dim funcName As String:: funcName = "ReSetLimitSet"

    Dim Update_Sheet_arr() As String
    Dim Current_Sheet_arr() As String
    Dim PorjectSheetName As String
    Dim index_i As Long
    
    Update_Sheet_arr = Split(gC_Sheet_UpdateLimit, ",")
    Current_Sheet_arr = Split(gC_Sheet_CurrentLimit, ",")
    'Put this code on OnProgramStart
    
'     LL.ResetSheet gC_Sheet_CurrentLimit
'     LL.ResetSheet gC_Sheet_UpdateLimit

    For index_i = 0 To UBound(Split(gC_Sheet_CurrentLimit, ",")) Step 1
        PorjectSheetName = Current_Sheet_arr(index_i)
        LL.ResetSheet PorjectSheetName
    Next index_i

    For index_i = 0 To UBound(Split(gC_Sheet_UpdateLimit, ",")) Step 1
        PorjectSheetName = Update_Sheet_arr(index_i)
        LL.ResetSheet PorjectSheetName
    Next index_i

    Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'Public Sub RunTimeError(funcName As String)
'    ' Sanity clause
'    If TheExec Is Nothing Then
'        MsgBox "IG-XL in not running!  Error encountered in Exec Interpose Function " + funcName + vbCrLf + _
'            "VBT Error # " + Trim$(CStr(Err.Number)) + ": " + Err.Description
'        Exit Sub
'    End If
'    TheExec.Datalog.WriteComment "Error encountered in Function::" + funcName
'End Sub


Public Function CheckSheet(pName As String) As Boolean

On Error GoTo ErrHandler
Dim funcName As String:: funcName = "CheckSheet"

'--------------------------------------------------------------------------------
'*************
'*Check Sheet*
'*************
'--------------------------------------------------------------------------------
Dim IsExist As Boolean
Dim i As Long
'IsExist = False
For i = 1 To Application.ActiveWorkbook.Sheets.Count
    If Application.ActiveWorkbook.Sheets(i).Name = pName Then
        IsExist = True
        Exit For
    End If
Next
CheckSheet = IsExist

Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



Sub SaveStringtoFile(ByVal WriteFileName As String, ByVal Save_str As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "SaveStringtoFile"

    Dim fs As Object
    Dim file_temp As Object
    Dim FileExists As Boolean:: FileExists = False
    

    Set fs = CreateObject("Scripting.filesystemobject")
    WriteFileName = TheExec.TestProgram.Path & "\" & WriteFileName
    
    FileExists = fs.FileExists(WriteFileName)
    
    If FileExists = False Then
        Set file_temp = fs.CreateTextFile(WriteFileName, False, False)
    Else
        Set file_temp = fs.OpenTextFile(WriteFileName, ForAppending, True)
    End If
    
    
    file_temp.WriteLine (Save_str)
    file_temp.Close
    
Exit Sub
ErrHandler:
     RunTimeError funcName
     If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function Limitsheet_MP_Datalog_FlagOn()
''-----20191206 Add for "LimitSheet_MP_DataLog_ON"
On Error GoTo ErrHandler

    LL.MP_Datalog_FlagOn

Exit Function
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Limitsheet_MP_Datalog_FlagOff()
''-----20191206 Add for "LimitSheet_MP_DataLog_ON"
On Error GoTo ErrHandler

    LL.MP_Datalog_FlagOff

Exit Function
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


