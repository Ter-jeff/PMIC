VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibLimitSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!

Option Explicit
'History:
'2018/04/03 V1:
'2018/04/13 V1:
'2018/04/24 V1:
Private Const bEnableCurrentSheet = False 'bEnableCurrentSheet=False: If you don't want to log into CurrentSheet
Private Const bEnableCurrentTextFile = False 'bEnableCurrentTextFile=False: If you don't want to log into CurrentSheet
Private Const bdbgIsExistTestName = False 'bdbgIsExistTestName=True: Check all item include Limit="N/A" & TName="*CODE""
Private Const bEnableLogAllJob = False    'bEnableLogAllJob=True   : Update High/Low limit for all Job into CurrentSheet.

Private m_DicTName        As New Dictionary  '2018/04/03 TTR:Only use this one to get index.
Private m_DicRun          As New Dictionary  '2018/04/03 TTR:Remove
Private m_DicHiLimit      As New Dictionary  '2018/04/03 TTR:Remove
Private m_DicLowLimit     As New Dictionary  '2018/04/03 TTR:Remove
Private m_DicTNum         As New Dictionary  '2018/04/03 TTR:Remove
Private m_DicUserTName    As New Dictionary  '2018/04/03 TTR:Remove
Private Const SupportType = "ISiteDouble,Double,ISiteLong,Integer,Long,IPinData,ISiteVariant,IPinListData,IPinListDataMath,ISiteComplexDouble"
Private Const RemoveLimitTName = "Delta,Code,Link,OTP"                             ' 20180424 remove Target
Private Const ListLimitTName = "Swp5,PreTrim,PreTrimCode,TweakTrim,FinalTrim,FinalTrimCode,PostBurn,PostBurnCode,P,IDS,ENG,X,FirstTrim,FirstTrimCode,TrimTarget,PostTrim,PostTrimCode,TrimLink,POSTBURNTARGET,Addr,POSTBURNGNG" ' 20180502 Update limit worksheet only need below keyword in Group7:
Private Const TNameSplit = "_"
Private Const SupportUnitType = "V,A,Ohm,,DB,Hz,S,V,uA/us"

Private g_LimitSheet_MP_DataLog_ON As Boolean '2019/12/06 Control MP keyWord








Public Function LimitSheetHeader(mS_JobList As String, mS_Sheet As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "UpdateTestNameSheet"

 
Dim mCol_Run       As Long
Dim mCol_flowtable As Long
Dim mCol_instance As Long
Dim mCol_testname As Long

Dim mCol_testnumber As Long
Dim mCol_UserTName As Long
Dim mCol_highcompsign As Long
Dim mCol_lolim() As Long
Dim mCol_hilim() As Long
Dim mCol_scale As Long
Dim mCol_units As Long
Dim mCol_format As Long
Dim mCol_passsort As Long
Dim mCol_failsort As Long
Dim mCol_passbin As Long
Dim mCol_failbin As Long
Dim mCol_result As Long


Dim mCol_passaction As Long
Dim mCol_failaction As Long
Dim mCol_assume As Long
Dim mCol_sites As Long
Dim mCol_comment As Long


Dim mVal_Run        As Variant
Dim mVal_flowtable  As Variant
Dim mVal_instance   As Variant
Dim mVal_testname   As Variant
Dim mVal_testnumber As Variant
Dim mVal_UserTName    As Variant
Dim mVal_highcompsign   As Variant
Dim mVal_lolim()    As Variant
Dim mVal_hilim()    As Variant
Dim mVal_scale  As Variant
Dim mVal_units  As Variant
Dim mVal_format As Variant
Dim mVal_passsort   As Variant
Dim mVal_failsort   As Variant
Dim mVal_passbin    As Variant
Dim mVal_failbin    As Variant
Dim mVal_result As Variant
Dim mS_Job()            As String
Dim mL_JobNum           As Long
Dim mRow_JobList        As Long
Dim mCol_JobListStart   As Long
Dim mRow_Header   As Long

Dim mSheet As Worksheet
Dim FlowLimitsInfo As IFlowLimitsInfo
Dim i As Long
Dim Index As Long
Dim mCell As Object
Dim mTestMode As gtlTestMode
 
mTestMode = TheExec.Datalog.Setup.LotSetup.TESTMODE
If mTestMode <> gtl_Engineeringmode Then Exit Function

Set mSheet = Application.ActiveWorkbook.Sheets(mS_Sheet)
ActiveWindow.Zoom = 80

mS_Job = Split(mS_JobList, ",")
mL_JobNum = UBound(mS_Job) + 1
ReDim mCol_lolim(mL_JobNum - 1)
ReDim mCol_hilim(mL_JobNum - 1)
ReDim mVal_lolim(mL_JobNum - 1)
ReDim mVal_hilim(mL_JobNum - 1)

'Define the Row and Col.
mRow_JobList = 2
mRow_Header = 3

mCol_Run = 2
mCol_flowtable = 3
mCol_instance = 4
mCol_testname = 5
mCol_UserTName = 6
mCol_testnumber = 7
mCol_scale = 8
mCol_units = 9
mCol_lolim(i) = 10 + 2 * i
mCol_hilim(i) = 11 + 2 * i
For i = 0 To mL_JobNum - 1
    mCol_lolim(i) = 10 + 2 * i
    mCol_hilim(i) = 11 + 2 * i
Next i
mCol_comment = mCol_hilim(mL_JobNum - 1) + 1


For i = 0 To mL_JobNum - 1
  mSheet.Cells(mRow_JobList, mCol_lolim(i)) = mS_Job(i)
  mSheet.Cells(mRow_JobList, mCol_hilim(i)) = mS_Job(i)
Next i

mSheet.Cells(mRow_Header, mCol_flowtable) = "flowtable"
mSheet.Cells(mRow_Header, mCol_instance) = "instance"
mSheet.Cells(mRow_Header, mCol_testname) = "testname"
mSheet.Cells(mRow_Header, mCol_UserTName) = "PETname"
mSheet.Cells(mRow_Header, mCol_testnumber) = "testnumber"
mSheet.Cells(mRow_Header, mCol_scale) = "scale"
mSheet.Cells(mRow_Header, mCol_units) = "units"
For i = 0 To mL_JobNum - 1
  mSheet.Cells(mRow_Header, mCol_lolim(i)) = "lolim"
  mSheet.Cells(mRow_Header, mCol_hilim(i)) = "hilim"
Next i

mSheet.Cells(mRow_Header, mCol_comment) = "comment"


         
Dim mIntColor As gColorIndex
Dim mFontColor As gColorIndex

Worksheets(mS_Sheet).Select
'======== Set Font and size for all cell ==========
Range("A1:AO3").Select
With Selection
    .Font.Name = "Calibri"
    .Font.Size = 11
    .Font.Bold = False
    .RowHeight = 18
    .VerticalAlignment = xlCenter
    .HorizontalAlignment = xlCenter
    .WrapText = False
End With


mFontColor = BLACKCOLOR
mIntColor = Gray25Color
Range("a2:AO2").Select
With Selection
    .Interior.ColorIndex = mIntColor
    .Font.ColorIndex = mFontColor
End With


mFontColor = WHITECOLOR
mIntColor = DarkBuleColor
Range("a3:AO3").Select
With Selection
    .Interior.ColorIndex = mIntColor
    .Font.ColorIndex = mFontColor
    .ColumnWidth = 15
End With


Dim mRange As String
mSheet.Select
Cells(mRow_Header, mCol_flowtable).Select: Selection.ColumnWidth = 20
Cells(mRow_Header, mCol_instance).Select: Selection.ColumnWidth = 30
Cells(mRow_Header, mCol_testname).Select: Selection.ColumnWidth = 50
Cells(mRow_Header, mCol_UserTName).Select: Selection.ColumnWidth = 50

For i = 0 To (mL_JobNum - 1)
        Cells(mRow_Header, mCol_lolim(i)).Select: Selection.ColumnWidth = 20
        Cells(mRow_Header, mCol_hilim(i)).Select: Selection.ColumnWidth = 20
    If UCase(mSheet.Cells(mRow_JobList, mCol_hilim(i))) = UCase(TheExec.CurrentJob) Then
        mFontColor = BLACKCOLOR
        mIntColor = YELLOWCOLOR
        Cells(mRow_JobList, mCol_lolim(i)).Select: Selection.Interior.ColorIndex = mIntColor: Selection.Font.ColorIndex = mFontColor
        Cells(mRow_JobList, mCol_hilim(i)).Select: Selection.Interior.ColorIndex = mIntColor: Selection.Font.ColorIndex = mFontColor
        'Stop
    End If
Next i
'===================================================
'Freeze panes to always see the head line and test name
  ActiveWindow.FreezePanes = False
  mSheet.Select
  mSheet.Range("A1:AT60000").ClearOutline
  mSheet.Range("F4").Select
  ActiveWindow.FreezePanes = True
           
  LL.SheetDataList "AD4:AD60000", SupportUnitType 'Limitsheet:units==>"V,A,Ohm,,DB,Hz,S,V"
   
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function LimitSheetHeader_bak(mS_JobList As String, mS_Sheet As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "UpdateTestNameSheet"

 
Dim mCol_Run       As Long
Dim mCol_flowtable As Long
Dim mCol_instance As Long
Dim mCol_testname As Long
Dim mCol_testnumber As Long
Dim mCol_UserTName As Long
Dim mCol_highcompsign As Long
Dim mCol_lolim() As Long
Dim mCol_hilim() As Long
Dim mCol_scale As Long
Dim mCol_units As Long
Dim mCol_format As Long
Dim mCol_passsort As Long
Dim mCol_failsort As Long
Dim mCol_passbin As Long
Dim mCol_failbin As Long
Dim mCol_result As Long


Dim mCol_passaction As Long
Dim mCol_failaction As Long
Dim mCol_assume As Long
Dim mCol_sites As Long
Dim mCol_comment As Long


Dim mVal_Run        As Variant
Dim mVal_flowtable  As Variant
Dim mVal_instance   As Variant
Dim mVal_testname   As Variant
Dim mVal_testnumber As Variant
Dim mVal_UserTName    As Variant
Dim mVal_highcompsign   As Variant
Dim mVal_lolim()    As Variant
Dim mVal_hilim()    As Variant
Dim mVal_scale  As Variant
Dim mVal_units  As Variant
Dim mVal_format As Variant
Dim mVal_passsort   As Variant
Dim mVal_failsort   As Variant
Dim mVal_passbin    As Variant
Dim mVal_failbin    As Variant
Dim mVal_result As Variant
Dim mS_Job()            As String
Dim mL_JobNum           As Long
Dim mRow_JobList        As Long
Dim mCol_JobListStart   As Long
Dim mRow_Header   As Long

Dim mSheet As Worksheet
Dim FlowLimitsInfo As IFlowLimitsInfo
Dim i As Long
Dim Index As Long
Dim mCell As Object
Dim mTestMode As gtlTestMode
 
mTestMode = TheExec.Datalog.Setup.LotSetup.TESTMODE
If mTestMode <> gtl_Engineeringmode Then Exit Function

Set mSheet = Application.ActiveWorkbook.Sheets(mS_Sheet)
ActiveWindow.Zoom = 80

mS_Job = Split(mS_JobList, ",")
mL_JobNum = UBound(mS_Job) + 1
ReDim mCol_lolim(mL_JobNum - 1)
ReDim mCol_hilim(mL_JobNum - 1)
ReDim mVal_lolim(mL_JobNum - 1)
ReDim mVal_hilim(mL_JobNum - 1)

'Define the Row and Col.
mRow_JobList = 2
mRow_Header = 3

mCol_Run = 2
mCol_flowtable = 3
mCol_instance = 4
mCol_testname = 5
mCol_testnumber = 6
mCol_UserTName = 7
mCol_highcompsign = 8
mCol_lolim(i) = 9 + 2 * i
mCol_hilim(i) = 10 + 2 * i
    
For i = 0 To mL_JobNum - 1
    mCol_lolim(i) = 9 + 2 * i
    mCol_hilim(i) = 10 + 2 * i
Next i
mCol_scale = mCol_hilim(mL_JobNum - 1) + 1
mCol_units = mCol_hilim(mL_JobNum - 1) + 2
mCol_format = mCol_hilim(mL_JobNum - 1) + 3
mCol_passsort = mCol_hilim(mL_JobNum - 1) + 4
mCol_failsort = mCol_hilim(mL_JobNum - 1) + 5
mCol_passbin = mCol_hilim(mL_JobNum - 1) + 6
mCol_failbin = mCol_hilim(mL_JobNum - 1) + 7
mCol_result = mCol_hilim(mL_JobNum - 1) + 8
mCol_passaction = mCol_hilim(mL_JobNum - 1) + 9
mCol_failaction = mCol_hilim(mL_JobNum - 1) + 10
mCol_assume = mCol_hilim(mL_JobNum - 1) + 11
mCol_sites = mCol_hilim(mL_JobNum - 1) + 12
mCol_comment = mCol_hilim(mL_JobNum - 1) + 13


For i = 0 To mL_JobNum - 1
  mSheet.Cells(mRow_JobList, mCol_lolim(i)) = mS_Job(i)
  mSheet.Cells(mRow_JobList, mCol_hilim(i)) = mS_Job(i)
Next i

mSheet.Cells(mRow_Header, mCol_flowtable) = "flowtable"
mSheet.Cells(mRow_Header, mCol_instance) = "instance"
mSheet.Cells(mRow_Header, mCol_testname) = "testname"
mSheet.Cells(mRow_Header, mCol_testnumber) = "testnumber"
mSheet.Cells(mRow_Header, mCol_UserTName) = "UserTName"
mSheet.Cells(mRow_Header, mCol_highcompsign) = "highcompsign"
For i = 0 To mL_JobNum - 1
  mSheet.Cells(mRow_Header, mCol_lolim(i)) = "lolim"
  mSheet.Cells(mRow_Header, mCol_hilim(i)) = "hilim"
Next i

mSheet.Cells(mRow_Header, mCol_scale) = "scale"
mSheet.Cells(mRow_Header, mCol_units) = "units"

mSheet.Cells(mRow_Header, mCol_format) = "format"
mSheet.Cells(mRow_Header, mCol_passsort) = "passsort"
mSheet.Cells(mRow_Header, mCol_failsort) = "failsort"
mSheet.Cells(mRow_Header, mCol_passbin) = "passbin"
mSheet.Cells(mRow_Header, mCol_failbin) = "failbin"
mSheet.Cells(mRow_Header, mCol_failsort) = "failsort"
mSheet.Cells(mRow_Header, mCol_result) = "result"


mSheet.Cells(mRow_Header, mCol_passaction) = "passaction"
mSheet.Cells(mRow_Header, mCol_failaction) = "failaction"
mSheet.Cells(mRow_Header, mCol_assume) = "assume"
mSheet.Cells(mRow_Header, mCol_sites) = "sites"
mSheet.Cells(mRow_Header, mCol_comment) = "comment"


         
Dim mIntColor As gColorIndex
Dim mFontColor As gColorIndex

Worksheets(mS_Sheet).Select
'======== Set Font and size for all cell ==========
Range("A1:AO3").Select
With Selection
    .Font.Name = "Calibri"
    .Font.Size = 11
    .Font.Bold = False
    .RowHeight = 18
    .VerticalAlignment = xlCenter
    .HorizontalAlignment = xlCenter
    .WrapText = False
End With


mFontColor = BLACKCOLOR
mIntColor = Gray25Color
Range("a2:AO2").Select
With Selection
    .Interior.ColorIndex = mIntColor
    .Font.ColorIndex = mFontColor
End With


mFontColor = WHITECOLOR
mIntColor = DarkBuleColor
Range("a3:AO3").Select
With Selection
    .Interior.ColorIndex = mIntColor
    .Font.ColorIndex = mFontColor
    .ColumnWidth = 15
End With


Dim mRange As String
mSheet.Select
Cells(mRow_Header, mCol_flowtable).Select: Selection.ColumnWidth = 20
Cells(mRow_Header, mCol_instance).Select: Selection.ColumnWidth = 30
Cells(mRow_Header, mCol_testname).Select: Selection.ColumnWidth = 50
Cells(mRow_Header, mCol_UserTName).Select: Selection.ColumnWidth = 50

For i = 0 To (mL_JobNum - 1)
        Cells(mRow_Header, mCol_lolim(i)).Select: Selection.ColumnWidth = 20
        Cells(mRow_Header, mCol_hilim(i)).Select: Selection.ColumnWidth = 20
    If UCase(mSheet.Cells(mRow_JobList, mCol_hilim(i))) = UCase(TheExec.CurrentJob) Then
        mFontColor = BLACKCOLOR
        mIntColor = YELLOWCOLOR
        Cells(mRow_JobList, mCol_lolim(i)).Select: Selection.Interior.ColorIndex = mIntColor: Selection.Font.ColorIndex = mFontColor
        Cells(mRow_JobList, mCol_hilim(i)).Select: Selection.Interior.ColorIndex = mIntColor: Selection.Font.ColorIndex = mFontColor
        'Stop
    End If
Next i
'===================================================
'Freeze panes to always see the head line and test name
  ActiveWindow.FreezePanes = False
  mSheet.Select
  mSheet.Range("A1:AT60000").ClearOutline
  mSheet.Range("F4").Select
  ActiveWindow.FreezePanes = True
           
  LL.SheetDataList "AD4:AD60000", SupportUnitType 'Limitsheet:units==>"V,A,Ohm,,DB,Hz,S,V"
   
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function LimitTextHeader(ByVal mS_JobList As String, ByVal mF_File As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "UpdateTestNameSheet"

 
Dim mCol_Run       As Long
Dim mCol_flowtable As Long
Dim mCol_instance As Long
Dim mCol_testname As Long
Dim mCol_testnumber As Long
Dim mCol_UserTName As Long
Dim mCol_highcompsign As Long
Dim mCol_lolim() As Long
Dim mCol_hilim() As Long
Dim mCol_scale As Long
Dim mCol_units As Long
Dim mCol_format As Long
Dim mCol_passsort As Long
Dim mCol_failsort As Long
Dim mCol_passbin As Long
Dim mCol_failbin As Long
Dim mCol_result As Long


Dim mCol_passaction As Long
Dim mCol_failaction As Long
Dim mCol_assume As Long
Dim mCol_sites As Long
Dim mCol_comment As Long


Dim mVal_Run        As Variant
Dim mVal_flowtable  As Variant
Dim mVal_instance   As Variant
Dim mVal_testname   As Variant
Dim mVal_testnumber As Variant
Dim mVal_UserTName    As Variant
Dim mVal_highcompsign   As Variant
Dim mVal_lolim()    As Variant
Dim mVal_hilim()    As Variant
Dim mVal_scale  As Variant
Dim mVal_units  As Variant
Dim mVal_format As Variant
Dim mVal_passsort   As Variant
Dim mVal_failsort   As Variant
Dim mVal_passbin    As Variant
Dim mVal_failbin    As Variant
Dim mVal_result As Variant
Dim mS_Job()            As String
Dim mL_JobNum           As Long
Dim mRow_JobList        As Long
Dim mCol_JobListStart   As Long
Dim mRow_Header   As Long

Dim mSheet As Worksheet
Dim FlowLimitsInfo As IFlowLimitsInfo
Dim i As Long
Dim Index As Long
Dim mCell As Object
Dim mTestMode As gtlTestMode
 
mTestMode = TheExec.Datalog.Setup.LotSetup.TESTMODE
If mTestMode <> gtl_Engineeringmode Then Exit Function


mS_Job = Split(mS_JobList, ",")
mL_JobNum = UBound(mS_Job) + 1
ReDim mCol_lolim(mL_JobNum - 1)
ReDim mCol_hilim(mL_JobNum - 1)
ReDim mVal_lolim(mL_JobNum - 1)
ReDim mVal_hilim(mL_JobNum - 1)

'Define the Row and Col.
mRow_JobList = 2
mRow_Header = 3
'Dim g_mTXT_input_JobList()    As String
'Dim g_mTXT_input_Header()    As String

mCol_Run = 2
mCol_flowtable = 3
mCol_instance = 4
mCol_testname = 5
mCol_testnumber = 6
mCol_UserTName = 7
mCol_highcompsign = 8
mCol_lolim(i) = 9 + 2 * i
mCol_hilim(i) = 10 + 2 * i
    
For i = 0 To mL_JobNum - 1
    mCol_lolim(i) = 9 + 2 * i
    mCol_hilim(i) = 10 + 2 * i
Next i
mCol_scale = mCol_hilim(mL_JobNum - 1) + 1
mCol_units = mCol_hilim(mL_JobNum - 1) + 2
mCol_format = mCol_hilim(mL_JobNum - 1) + 3
mCol_passsort = mCol_hilim(mL_JobNum - 1) + 4
mCol_failsort = mCol_hilim(mL_JobNum - 1) + 5
mCol_passbin = mCol_hilim(mL_JobNum - 1) + 6
mCol_failbin = mCol_hilim(mL_JobNum - 1) + 7
mCol_result = mCol_hilim(mL_JobNum - 1) + 8
mCol_passaction = mCol_hilim(mL_JobNum - 1) + 9
mCol_failaction = mCol_hilim(mL_JobNum - 1) + 10
mCol_assume = mCol_hilim(mL_JobNum - 1) + 11
mCol_sites = mCol_hilim(mL_JobNum - 1) + 12
mCol_comment = mCol_hilim(mL_JobNum - 1) + 13

ReDim g_mTXT_input_JobList(mCol_comment - 1) As String
ReDim g_mTXT_input_Header(mCol_comment - 1) As String
Dim mCurrentTXTinput_JobList As String
Dim mCurrentTXTinput_Header As String



For i = 0 To mL_JobNum - 1

    g_mTXT_input_JobList(mCol_lolim(i) - 1) = mS_Job(i)
    g_mTXT_input_JobList(mCol_hilim(i) - 1) = mS_Job(i)
Next i

g_mTXT_input_Header(mCol_flowtable - 1) = "flowtable"
g_mTXT_input_Header(mCol_instance - 1) = "instance"
g_mTXT_input_Header(mCol_testname - 1) = "testname"
g_mTXT_input_Header(mCol_testnumber - 1) = "testnumber"
g_mTXT_input_Header(mCol_UserTName - 1) = "UserTName"
g_mTXT_input_Header(mCol_highcompsign - 1) = "highcompsign"


For i = 0 To mL_JobNum - 1
    g_mTXT_input_Header(mCol_lolim(i) - 1) = "lolim"
    g_mTXT_input_Header(mCol_hilim(i) - 1) = "hilim"
Next i


g_mTXT_input_Header(mCol_scale - 1) = "scale"
g_mTXT_input_Header(mCol_units - 1) = "units"



g_mTXT_input_Header(mCol_format - 1) = "format"
g_mTXT_input_Header(mCol_passsort - 1) = "passsort"
g_mTXT_input_Header(mCol_failsort - 1) = "failsort"
g_mTXT_input_Header(mCol_passbin - 1) = "passbin"
g_mTXT_input_Header(mCol_failbin - 1) = "failbin"
g_mTXT_input_Header(mCol_failsort - 1) = "failsort"
g_mTXT_input_Header(mCol_result - 1) = "result"



g_mTXT_input_Header(mCol_passaction - 1) = "passaction"
g_mTXT_input_Header(mCol_failaction - 1) = "failaction"
g_mTXT_input_Header(mCol_assume - 1) = "assume"
g_mTXT_input_Header(mCol_sites - 1) = "sites"
g_mTXT_input_Header(mCol_comment - 1) = "comment"


mCurrentTXTinput_JobList = Join(g_mTXT_input_JobList, vbTab)
SaveStringtoFile mF_File, mCurrentTXTinput_JobList

mCurrentTXTinput_Header = Join(g_mTXT_input_Header, vbTab)
SaveStringtoFile mF_File, mCurrentTXTinput_Header


   
   
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub Class_Initialize()
If 0 Then
    m_DicHiLimit.RemoveAll
    m_DicLowLimit.RemoveAll
    m_DicTNum.RemoveAll
    m_DicTName.RemoveAll
    m_DicUserTName.RemoveAll
  End If
End Sub



'Public Function TestLimit(resultVal, [lowVal As Double = -1.#INF], [hiVal As Double = 1.#INF], [lowCompareSign As tlCompareSign = tlSignGreaterEqual], [highCompareSign As tlCompareSign = tlSignLessEqual], [scaletype As tlScaleType = scaleNone], [unit As UnitType = unitNone], [formatStr As String], [TName As String], [compareMode As tlLimitCompareType = CompareAverage], [pinName As String], [forceVal As Double], [forceunit As UnitType = unitNone], [customUnit As String], [customForceunit As String], [ForceResults As tlLimitForceResults = tlForceNone], [TNum As Long = -1]) As Long
Public Function FlowTestLimit(ResultVal As Variant, _
Optional ByVal lowVal As Variant, _
Optional ByVal hiVal As Variant, _
Optional lowCompareSign As tlCompareSign = tlSignGreaterEqual, _
Optional highCompareSign As tlCompareSign = tlSignLessEqual, _
Optional scaletype As tlScaleType = scaleNone, _
Optional Unit As UnitType = unitNone, _
Optional formatStr As String, _
Optional ByVal TName As String, _
Optional compareMode As tlLimitCompareType = CompareAverage, _
Optional PinName As Variant, _
Optional forceVal As Variant, _
Optional ForceUnit As UnitType = unitNone, _
Optional customUnit As String, _
Optional customForceunit As String, _
Optional ForceResults As tlLimitForceResults = tlForceNone, _
Optional ByVal TNum As Long = 0)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "FlowTestLimit"
    
Dim mTName           As String
Dim mUserTName       As String
Dim mS_dlogstr       As String
Dim bTimer           As Boolean
Dim ElapsedTime      As Double
Dim ReferenceTime    As Double

Dim index_i As Long
Dim Update_Sheet_arr() As String
    Update_Sheet_arr = Split(gC_Sheet_UpdateLimit, ",")
Dim PorjectSheetName As String

Dim mCheckTNameFlag  As Boolean 'TTR(2018/05/25)
Dim mCheckTNameIN_Dictionary  As Boolean
Dim Index_for_TName As Long
bTimer = False
TName = Trim(TName) '20180416 Claire add
mTName = TName
'mUserTName = TName

TheExec.Flow.ApplyScalingToLimits = True  '' --20190408 change scale rule in Limit sheet



''-----20191206 MP_DataLog_ON flag judgement
Dim SkipJudgeFlag As Boolean
If g_LimitSheet_MP_DataLog_ON = True Then '' user select EnableWord "LimitSheet_MP_DataLog_ON"
    SkipJudgeFlag = Check_LimitSheet_MP_DataLog_ON_To_Skip(TName) '' Check particular datalog name if naming not match will skip lot out
    If SkipJudgeFlag = True Then Exit Function
End If

'==============
Dim Ori_TName As String
Dim AR() As String
Dim TA(11) As String
Dim i As Integer
Ori_TName = TName
 AR = Split(TName, "_")

    If UBound(AR) = 12 Then
        For i = 0 To 11
            TA(i) = AR(i)
        Next i
            TName = UCase(Trim(Join(TA, "_")))
    End If
'==============

mCheckTNameFlag = CheckTNameFlag(TName) ''ListLimitTName = "PreTrim,PreTrimCode,FinalTrim,FinalTrimCode,PostBurn,PostBurnCode,P,IDS" ' keyword in Group7:
If mCheckTNameFlag = True Then
    mCheckTNameIN_Dictionary = CheckTName(TName) ' check This TName in dictionary or not
    mUserTName = GetUserTName(Ori_TName) 'TName
End If
''If mCheckTNameIN_Dictionary = False Then Stop
If mCheckTNameIN_Dictionary Then Index_for_TName = m_DicTName.Item(TName)


'ReferenceTime = TheExec.Timer
    
    'A.GET LIMIT:
    '20180327 only support SiteDouble, Double, SiteLong, Long, PinListData, SiteVariant, Pinlist
    'If (InStr(UCase(SupportType), UCase(TypeName(ResultVal))) = False) Then
    '    Debug.Print TypeName(ResultVal) & ", only support SiteDouble, Double, SiteLong, Long, PinListData, SiteVariant, Pinlist"
    '    Stop
    'End If
        
    If Not IsMissing(forceVal) Then
        forceVal = CDbl(forceVal)
    Else
        forceVal = CDbl(0)
    End If
    If IsMissing(PinName) Then PinName = ""

'''*********************************************************************************************
    If mCheckTNameIN_Dictionary Then
        If mCheckTNameFlag Or ForceResults = tlForceFlow Then
            If GetRun(TName, bdbgIsExistTestName) = False Then GoTo SkipLimitSheet

            If (0) Then
                hiVal = LimitTable(Index_for_TName).ParamSyntax.HiLimit
                lowVal = LimitTable(Index_for_TName).ParamSyntax.LoLimit
                TNum = LimitTable(Index_for_TName).ParamSyntax.TestNumber
            Else
                hiVal = GetHiLimit(TName)
                lowVal = GetLowLimit(TName)
                TNum = GetTNum(TName)
            End If

            Unit = GetFlowParameters(TName, gut_UnitType, Unit, scaletype, customUnit)
            scaletype = GetFlowParameters(TName, gut_ScaleType, Unit, scaletype, customUnit)
            customUnit = GetFlowParameters(TName, gut_UnitCustomStr, Unit, scaletype, customUnit)
            'If Unit = 0 Then scaletype = scaleNoScaling

        ElseIf IsMissing(lowVal) Or IsMissing(hiVal) Then
        'Special Case:
        '2-1:lowVal = "N/A" OR hiVal = "N/A"
            If bdbgIsExistTestName Then Call CheckTName(TName, bdbgIsExistTestName)
            lowVal = "N/A"
            hiVal = "N/A"
            ForceResults = gtlForceNA
            TNum = GetTNum(TName)
        Else
            If ForceResults = tlForceFlow Then GetHiLoFromFlow lowVal, hiVal, 0
        End If ' end of judge high/low spec missing or type is forceFlow
    Else
        If ForceResults = tlForceFlow Then
            GetHiLoFromFlow lowVal, hiVal, 0
        ElseIf IsMissing(lowVal) Or IsMissing(hiVal) Then
            If IsMissing(lowVal) Then lowVal = "N/A"
            If IsMissing(hiVal) Then hiVal = "N/A"
            ForceResults = gtlForceNA
        Else
        ' limit direct( no direction define , not forceflow, have limit high and low.
        End If
    End If ' dicationary judge end
'''*********************************************************************************************
''''*********************************************************************************************

SkipLimitSheet:
    If mUserTName <> TName And mUserTName <> "" Then
        TName = mUserTName 'FINAL TESTNAME
    Else
        TName = Ori_TName
    End If

'20180626 evans : fix object didn't be assigned issue
    If IsMissing(lowVal) Then
        lowVal = "N/A"
        ForceResults = gtlForceNA
    ElseIf lowVal = "N/A" Then
        ForceResults = gtlForceNA
    End If

    If IsMissing(hiVal) Then
        hiVal = "N/A"
        ForceResults = gtlForceNA
    ElseIf hiVal = "N/A" Then
        ForceResults = gtlForceNA
    End If

    
    'mS_dlogstr = "funcName = " & funcName & "::ElapsedTime -1 =" & TheExec.Timer(ReferenceTime)
    'If bTimer Then TheExec.Datalog.WriteComment mS_dlogstr
    'ReferenceTime = TheExec.Timer


   'B.FLOW TEST LIMIT:
    Select Case ForceResults
        Case gtlForceFlow
                            ForceResults = gtlForceNone
        Case gtlForceNone
                            ForceResults = gtlForceNone
        Case tlForcePass
                            ForceResults = gtlForcePass
        Case tlForceFail
                            ForceResults = tlForceFail
        Case gtlForceNA
                            ForceResults = gtlForceNone
        Case Else
                    'Put the code.
    End Select
    
    'TName = Ori_TName
    
        If lowVal = "N/A" And hiVal = "N/A" Then
            TheExec.Flow.TestLimit ResultVal:=ResultVal, scaletype:=scaletype, Unit:=Unit, _
            formatStr:=formatStr, TName:=TName, PinName:=PinName, forceVal:=forceVal, ForceUnit:=ForceUnit, customUnit:=customUnit, customForceunit:=customForceunit, ForceResults:=ForceResults ', TNum:=TNum
        ElseIf lowVal = "N/A" Then
            If TheExec.TesterMode = testModeOffline Then ResultVal = hiVal
            TheExec.Flow.TestLimit ResultVal:=ResultVal, hiVal:=hiVal, scaletype:=scaletype, Unit:=Unit, _
            formatStr:=formatStr, TName:=TName, PinName:=PinName, forceVal:=forceVal, ForceUnit:=ForceUnit, customUnit:=customUnit, customForceunit:=customForceunit, ForceResults:=ForceResults ', TNum:=TNum
        ElseIf hiVal = "N/A" Then
            If TheExec.TesterMode = testModeOffline Then ResultVal = lowVal
            TheExec.Flow.TestLimit ResultVal:=ResultVal, lowVal:=lowVal, scaletype:=scaletype, Unit:=Unit, _
            formatStr:=formatStr, TName:=TName, PinName:=PinName, forceVal:=forceVal, ForceUnit:=ForceUnit, customUnit:=customUnit, customForceunit:=customForceunit, ForceResults:=ForceResults ', TNum:=TNum
        Else
            If TheExec.TesterMode = testModeOffline Then ResultVal = (lowVal + hiVal) / 2
            TheExec.Flow.TestLimit ResultVal:=ResultVal, lowVal:=lowVal, hiVal:=hiVal, lowCompareSign:=lowCompareSign, highCompareSign:=highCompareSign, scaletype:=scaletype, Unit:=Unit, _
           formatStr:=formatStr, TName:=TName, compareMode:=compareMode, PinName:=PinName, forceVal:=forceVal, ForceUnit:=ForceUnit, customUnit:=customUnit, customForceunit:=customForceunit, ForceResults:=ForceResults ', TNum:=TNum
        End If
'    End If
      
    'mS_dlogstr = "funcName = " & funcName & "::ElapsedTime -2 =" & TheExec.Timer(ReferenceTime)
    'If bTimer Then TheExec.Datalog.WriteComment mS_dlogstr
    'ReferenceTime = TheExec.Timer
    
    'C.EXPORT TO TABLE:
'''''''    If g_TestMode <> gtl_Engineeringmode Then Exit Function
    
    If TNum = 0 Then TNum = TheExec.Datalog.LastTestNumLogged
    
    If mCheckTNameFlag Or ForceResults = tlForceFlow Then
        If Not mCheckTNameIN_Dictionary And TName <> "" Then
            If UBound(Split(gC_Sheet_UpdateLimit, ",")) > 0 Then
                If UCase(TName) Like "*BUCK*" Then
                    For index_i = 0 To UBound(Split(gC_Sheet_UpdateLimit, ",")) Step 1
                        If Update_Sheet_arr(index_i) Like "*BUCK*" Then Exit For
                    Next index_i
                    If index_i > UBound(Split(gC_Sheet_UpdateLimit, ",")) Then index_i = 0
                ElseIf UCase(TName) Like "*LDO*" Then
                    For index_i = 0 To UBound(Split(gC_Sheet_UpdateLimit, ",")) Step 1
                        If Update_Sheet_arr(index_i) Like "*LDO*" Then Exit For
                    Next index_i
                    If index_i > UBound(Split(gC_Sheet_UpdateLimit, ",")) Then index_i = 0
                ElseIf UCase(TName) Like "*ACORE*" Then
                    For index_i = 0 To UBound(Split(gC_Sheet_UpdateLimit, ",")) Step 1
                        If Update_Sheet_arr(index_i) Like "*ACORE*" Then Exit For
                    Next index_i
                    If index_i > UBound(Split(gC_Sheet_UpdateLimit, ",")) Then index_i = 0
                ElseIf UCase(TName) Like "*BOOST*" Then
                    For index_i = 0 To UBound(Split(gC_Sheet_UpdateLimit, ",")) Step 1
                        If Update_Sheet_arr(index_i) Like "*BOOST*" Then Exit For
                    Next index_i
                    If index_i > UBound(Split(gC_Sheet_UpdateLimit, ",")) Then index_i = 0
                Else
                    index_i = 0
                End If
                
                PorjectSheetName = Update_Sheet_arr(index_i)
                
                UpdateTestNameSheet PorjectSheetName, TName, lowVal, hiVal, TNum, , Unit, scaletype, customUnit, index_i
            Else
                UpdateTestNameSheet gC_Sheet_UpdateLimit, TName, lowVal, hiVal, TNum, , Unit, scaletype, customUnit
            End If ' end of mutlisheet judge
        End If ' end of dictionary check
    End If ' end of Group7Name check
    
    If bEnableCurrentSheet = True Then
        'If mCheckTNameFlag Then
        UpdateTestNameSheet gC_Sheet_CurrentLimit, TName, lowVal, hiVal, TNum, mUserTName, Unit, scaletype, customUnit
        'End If
    Else
        If bEnableCurrentTextFile = True Then
            UpdateTestNameTextFile gC_Sheet_CurrentLimit, TName, lowVal, hiVal, TNum, mUserTName, Unit, scaletype, customUnit
        End If
    End If
    
    'mS_dlogstr = "funcName = " & funcName & "::ElapsedTime -3 =" & TheExec.Timer(ReferenceTime)
    'If bTimer Then TheExec.Datalog.WriteComment mS_dlogstr
    'ReferenceTime = TheExec.Timer


    TheExec.Flow.ApplyScalingToLimits = False  '' --20190408 Reset scale rule in Limit sheet
    
    Exit Function
ErrHandler:
    RunTimeError funcName
    Stop
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetHiLimit(TName As String) As Variant
Dim funcName As String:: funcName = "GetHiLimit"

Dim i              As Long
Dim mKey           As String
Dim mVal           As Variant
Dim Index          As Long
Dim mS_Temp        As String

mKey = UCase(TName)

    If m_DicTName.Exists(mKey) Then
       Index = m_DicTName.Item(mKey)
        If CStr(LimitTable(Index).ParamSyntax.HiLimit) Like "*N*A*" Then
            GetHiLimit = "N/A"
        Else
            GetHiLimit = LimitTable(Index).ParamSyntax.HiLimit
        End If
       
'       If CStr(GetHiLimit) Like "*N*A*" Then GoTo Error
'        If CStr(GetHiLimit) Like "*N*A*" Then GetHiLimit = "N/A"
    Else
       Debug.Print "mKey=" & mKey & " Not Exists."
       GetHiLimit = 9999
    End If


Exit Function
Error:
mS_Temp = funcName & ":: Get test limit = " & CStr(GetHiLimit) & " failed: TName= '" & TName & ". Please edit it in limit table."
TheExec.AddOutput (mS_Temp)
'TheExec.Datalog.WriteComment (mS_Temp)
'Stop
GetHiLimit = 9999
End Function


Public Function GetFlowParameters(TName As String, GetFlowPara As gutTestlimit, _
                                    Optional mUnitType As UnitType = unitCustom, _
                                    Optional mScaleType As tlScaleType = scaleNone, _
                                    Optional mUnitCustomStr As String = "") As Variant
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "GetFlowParameters"
Dim Dlog              As String
Dim TableColumn       As String
Dim dScaleFactor      As Double
Dim m_UnitType        As UnitType
Dim m_ScaleType       As tlScaleType
Dim m_UnitCustomStr   As String

Dim i              As Long
Dim mKey           As String
Dim mVal           As Variant
Dim Index          As Long
Dim mS_Temp        As String

mKey = UCase(TName)

    If m_DicTName.Exists(mKey) Then
       Index = m_DicTName.Item(mKey)
                If GetFlowPara = gut_UnitType Then
                    'A.UnitType
                     mVal = LimitTable(Index).ParamSyntax.Units
                    Select Case UCase(mVal)
                        Case UCase("A"):       'Const unitAmp = 1
                             Dlog = "unitAmp"
                             TableColumn = "A"
                             m_UnitType = unitAmp
                        Case UCase("Ohm"):    'Const unitCustom = 8
                             Dlog = "unitCustom"
                             TableColumn = "unitCustom"
                             TableColumn = "Ohm"
                             m_UnitType = unitCustom
                        Case UCase("uA/us")
                             Dlog = "unitCustom"
                             TableColumn = "unitCustom"
                             TableColumn = "uA/us"
                             m_UnitType = unitCustom
                        Case UCase("DB"):        'Const unitDb = 5
                             Dlog = "unitDb"
                             TableColumn = "DB"
                             m_UnitType = unitDb
                        Case UCase("Hz"):        'Const unitHz = 3
                             Dlog = "unitHz"
                             TableColumn = "Hz"
                             m_UnitType = unitHz
                        Case UCase(""):      'Const unitNone = 0
                             Dlog = "unitNone"
                             TableColumn = ""
                             m_UnitType = unitNone
                        Case UCase("S"):      'Const unitTime = 4
                             Dlog = "unitTime"
                             TableColumn = "S"
                             m_UnitType = unitTime
                        Case UCase("V"):      'Const unitVolt = 2
                             Dlog = "unitVolt"
                             TableColumn = "V"
                             m_UnitType = unitVolt
                        Case UCase("F"):
                             Dlog = "unitCustom"
                             TableColumn = "F"
                             m_UnitType = unitCustom
                        Case UCase("%"):
                             Dlog = "unitCustom"
                             TableColumn = "%"
                             m_UnitType = unitCustom '20180413 JC add
                        Case UCase("mS"):
                             Dlog = "unitCustom"
                             TableColumn = "mS"
                             m_UnitType = unitCustom
                        Case UCase("nS"):
                             Dlog = "unitCustom"
                             TableColumn = "nS"
                             m_UnitType = unitCustom
                        Case UCase("C"):
                             Dlog = "unitCustom"
                             TableColumn = "C"
                             m_UnitType = unitCustom
                        Case UCase("V/Code"):
                             Dlog = "unitCustom"
                             TableColumn = "V/Code"
                             m_UnitType = unitCustom
                        Case UCase("V/uS"):
                             Dlog = "unitCustom"
                             TableColumn = "V/uS"
                             m_UnitType = unitCustom
                        Case UCase("pF"):
                             Dlog = "unitCustom"
                             TableColumn = "pF"
                             m_UnitType = unitCustom
                        Case UCase("code"):
                             Dlog = "unitCustom"
                             TableColumn = "code"
                             m_UnitType = unitCustom
                        Case UCase("V/pS"):
                             Dlog = "unitCustom"
                             TableColumn = "V/pS"
                             m_UnitType = unitCustom
                        Case UCase("V/nS"):
                             Dlog = "unitCustom"
                             TableColumn = "V/nS"
                             m_UnitType = unitCustom
                        Case UCase("A/A"):
                             Dlog = "unitCustom"
                             TableColumn = "A/A"
                             m_UnitType = unitCustom
                        Case Else
                            m_UnitType = unitCustom
                            m_UnitType = mUnitType
                            Stop
                    End Select
                    
                    'Final:GetFlowParameters
                    GetFlowParameters = m_UnitType
                    
                ElseIf GetFlowPara = gut_ScaleType Then
                    'B.ScaleType
                    mVal = LimitTable(Index).ParamSyntax.UScale
                    Select Case mVal
                        Case "f":        'Const scaleFemto = 15
                             Dlog = "scaleFemto"
                             TableColumn = "f"
                             dScaleFactor = 0.000000000000001
                             m_ScaleType = scaleFemto
                        Case "K", "k":       '20180827 add
                             Dlog = "scaleKilo"
                             TableColumn = "K"
                             dScaleFactor = 1000#
                             m_ScaleType = scaleKilo
                        Case scaleKilo:         'Const scaleKilo = -3 (&HFFFFFFFD)
                             Dlog = "scaleKilo"
                             TableColumn = "K"
                             dScaleFactor = 1000#
                             m_ScaleType = scaleKilo
                        Case "G":         'Const scaleGiga = -9 (&HFFFFFFF7)
                             Dlog = "scaleGiga"
                             TableColumn = "G"
                             dScaleFactor = 1000000000#
                             m_ScaleType = scaleGiga
                        Case scaleGiga:         'Const scaleGiga = -9 (&HFFFFFFF7)
                             Dlog = "scaleGiga"
                             TableColumn = "G"
                             dScaleFactor = 1000000000#
                             m_ScaleType = scaleGiga
                        Case "M":         'Const scaleMega = -6 (&HFFFFFFFA)
                             Dlog = "scaleMega"
                             TableColumn = "M"
                             dScaleFactor = 1000000#
                             m_ScaleType = scaleMega
                        Case scaleMega:         'Const scaleMega = -6 (&HFFFFFFFA)
                             Dlog = "scaleMega"
                             TableColumn = "M"
                             dScaleFactor = 1000000#
                             m_ScaleType = scaleMega
                        Case "u":        'Const scaleMicro = 6
                             Dlog = "scaleMicro"
                             TableColumn = "u"
                             dScaleFactor = 0.000001
                             m_ScaleType = scaleMicro
                        Case scaleMicro:        'Const scaleMicro = 6
                             Dlog = "scaleMicro"
                             TableColumn = "u"
                             dScaleFactor = 0.000001
                             m_ScaleType = scaleMicro
                        Case "m":        'Const scaleMilli = 3
                             Dlog = "scaleMilli"
                             TableColumn = "m"
                             dScaleFactor = 0.001
                             m_ScaleType = scaleMilli
                        Case scaleMilli:        'Const scaleMilli = 3
                             Dlog = "scaleMilli"
                             TableColumn = "m"
                             dScaleFactor = 0.001
                             m_ScaleType = scaleMilli
                        Case "n":         'Const scaleNano = 9
                             Dlog = "scaleNano"
                             TableColumn = "n"
                             dScaleFactor = 0.000000001
                             m_ScaleType = scaleNano
                        Case scaleNano:         'Const scaleNano = 9
                             Dlog = "scaleNano"
                             TableColumn = "n"
                             dScaleFactor = 0.000000001
                             m_ScaleType = scaleNano
                        Case scaleNone:         'Const scaleNone = 0
                             Dlog = "scaleNone"
                             TableColumn = ""
                             m_ScaleType = scaleNone
                        Case scaleNoScaling:    'Const scaleNoScaling = 9999 (&H270F)
                             Dlog = "scaleNoScaling"
                             TableColumn = ""
                             dScaleFactor = 1#
                             m_ScaleType = scaleNoScaling
                        Case "%":      'Const scalePercent = 2
                             Dlog = "scalePercent"
                             TableColumn = "%"
                             m_ScaleType = scalePercent
                        Case scalePercent:      'Const scalePercent = 2
                             Dlog = "scalePercent"
                             TableColumn = "%"
                             m_ScaleType = scalePercent
                        Case "p":         'Const scalePico = 12
                             Dlog = "scalePico"
                             TableColumn = "p"
                             dScaleFactor = 0.00000000001
                             m_ScaleType = scalePico
                        Case scalePico:         'Const scalePico = 12
                             Dlog = "scalePico"
                             TableColumn = "p"
                             dScaleFactor = 0.00000000001
                             m_ScaleType = scalePico
                        Case "T", "t":        'Const scaleTera = -12 (&HFFFFFFF4)
                             Dlog = "scaleTera"
                             TableColumn = "T"
                             dScaleFactor = 1000000000000#
                             m_ScaleType = scaleTera
                        Case scaleTera:         'Const scaleTera = -12 (&HFFFFFFF4)
                             Dlog = "scaleTera"
                             TableColumn = "T"
                             dScaleFactor = 1000000000000#
                             m_ScaleType = scaleTera
                        Case Else
                              m_ScaleType = mScaleType
                             Stop    '' If you stop in here, please check your Scale "Capital / low case"
                    End Select
                    'Final:GetFlowParameters
                    GetFlowParameters = m_ScaleType
                    
                ElseIf GetFlowPara = gut_UnitCustomStr Then
                
                    'C.UnitCustomSt
                     mVal = LimitTable(Index).ParamSyntax.Units
                     m_UnitCustomStr = ""
                     Select Case UCase(mVal)
                         Case UCase("A"):
                         Case UCase("Ohm"):
                              m_UnitCustomStr = "Ohm"
                         Case UCase("DB"):
                         Case UCase("Hz"):
                         Case UCase(""):
                         Case UCase("S"):
                         Case UCase("V"):
                         Case Else
                             m_UnitCustomStr = mVal
                             If mUnitCustomStr <> "" Then m_UnitCustomStr = m_UnitCustomStr
                     End Select
                     'Final:GetFlowParameters
                     GetFlowParameters = m_UnitCustomStr
                Else
                 
                 'Put the code for other case..

                End If
    Else
          Debug.Print "mKey=" & mKey & " Not Exists."
          Stop
    End If

          
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ByTNum_GetFlowParameters(TNum As Long, GetFlowPara As gutTestlimit, _
                                    Optional mUnitType As UnitType = unitCustom, _
                                    Optional mScaleType As tlScaleType = scaleNone, _
                                    Optional mUnitCustomStr As String = "") As Variant
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "ByTNum_GetFlowParameters"
Dim Dlog              As String
Dim TableColumn       As String
Dim dScaleFactor      As Double
Dim m_UnitType        As UnitType
Dim m_ScaleType       As tlScaleType
Dim m_UnitCustomStr   As String

Dim i              As Long
Dim mKey           As String
Dim mVal           As Variant
Dim Index          As Long
Dim mS_Temp        As String

mKey = UCase(TNum)

    If m_DicTNum.Exists(mKey) Then
       Index = m_DicTNum.Item(mKey)
                If GetFlowPara = gut_UnitType Then
                    'A.UnitType
                     mVal = LimitTable(Index).ParamSyntax.Units
                    Select Case UCase(mVal)
                        Case UCase("A"):       'Const unitAmp = 1
                             Dlog = "unitAmp"
                             TableColumn = "A"
                             m_UnitType = unitAmp
                        Case UCase("ohm"):    'Const unitCustom = 8
                             Dlog = "unitCustom"
                             TableColumn = "unitCustom"
                             TableColumn = "ohm"
                             m_UnitType = unitCustom
                             'If mUnitCustomStr <> "" Then TableColumn = mUnitCustomStr
                        Case UCase("DB"):        'Const unitDb = 5
                             Dlog = "unitDb"
                             TableColumn = "DB"
                             m_UnitType = unitDb
                        Case UCase("Hz"):        'Const unitHz = 3
                             Dlog = "unitHz"
                             TableColumn = "Hz"
                             m_UnitType = unitHz
                        Case UCase(""):      'Const unitNone = 0
                             Dlog = "unitNone"
                             TableColumn = ""
                             m_UnitType = unitNone
                        Case UCase("S"):      'Const unitTime = 4
                             Dlog = "unitTime"
                             TableColumn = "S"
                             m_UnitType = unitTime
                        Case UCase("V"):      'Const unitVolt = 2
                             Dlog = "unitVolt"
                             TableColumn = "V"
                             m_UnitType = unitVolt
                        Case Else
                                 m_UnitType = unitCustom
                                 m_UnitType = mUnitType
                                 Stop
                    End Select
                    
                    'Final:ByTNum_GetFlowParameters
                    ByTNum_GetFlowParameters = m_UnitType
                    
                ElseIf GetFlowPara = gut_ScaleType Then
                    'B.ScaleType
                    mVal = LimitTable(Index).ParamSyntax.UScale
                    Select Case mVal
                        Case "f":        'Const scaleFemto = 15
                             Dlog = "scaleFemto"
                             TableColumn = "f"
                             dScaleFactor = 0.000000000000001
                             m_ScaleType = scaleFemto
                        Case scaleGiga:         'Const scaleGiga = -9 (&HFFFFFFF7)
                             Dlog = "scaleGiga"
                             TableColumn = "G"
                             dScaleFactor = 1000000000#
                             m_ScaleType = scaleGiga
                        Case scaleKilo:         'Const scaleKilo = -3 (&HFFFFFFFD)
                             Dlog = "scaleKilo"
                             TableColumn = "K"
                             dScaleFactor = 1000#
                             m_ScaleType = scaleKilo
                        Case scaleMega:         'Const scaleMega = -6 (&HFFFFFFFA)
                             Dlog = "scaleMega"
                             TableColumn = "M"
                             dScaleFactor = 1000000#
                             m_ScaleType = scaleMega
                        Case scaleMicro:        'Const scaleMicro = 6
                             Dlog = "scaleMicro"
                             TableColumn = "u"
                             dScaleFactor = 0.000001
                             m_ScaleType = scaleMicro
                        Case scaleMilli:        'Const scaleMilli = 3
                             Dlog = "scaleMilli"
                             TableColumn = "m"
                             dScaleFactor = 0.001
                             m_ScaleType = scaleMilli
                        Case scaleNano:         'Const scaleNano = 9
                             Dlog = "scaleNano"
                             TableColumn = "n"
                             dScaleFactor = 0.000000001
                             m_ScaleType = scaleNano
                        Case scaleNone:         'Const scaleNone = 0
                             Dlog = "scaleNone"
                             TableColumn = ""
                             m_ScaleType = scaleNone
                        Case scaleNoScaling:    'Const scaleNoScaling = 9999 (&H270F)
                             Dlog = "scaleNoScaling"
                             TableColumn = ""
                             dScaleFactor = 1#
                             m_ScaleType = scaleNoScaling
                        Case scalePercent:      'Const scalePercent = 2
                             Dlog = "scalePercent"
                             TableColumn = "%"
                             m_ScaleType = scalePercent
                        Case scalePico:         'Const scalePico = 12
                             Dlog = "scalePico"
                             TableColumn = "p"
                             dScaleFactor = 0.00000000001
                             m_ScaleType = scalePico
                        Case scaleTera:         'Const scaleTera = -12 (&HFFFFFFF4)
                             Dlog = "scaleTera"
                             TableColumn = "T"
                             dScaleFactor = 1000000000000#
                             m_ScaleType = scaleTera
                        Case Else
                              m_ScaleType = mScaleType
                             Stop
                    End Select
                    'Final:ByTNum_GetFlowParameters
                    ByTNum_GetFlowParameters = m_ScaleType
                    
                ElseIf GetFlowPara = gut_UnitCustomStr Then
                
                    'C.UnitCustomSt
                     mVal = LimitTable(Index).ParamSyntax.Units
                     m_UnitCustomStr = ""
                     Select Case UCase(mVal)
                         Case UCase("A"):
                         Case UCase("Ohm"):
                              m_UnitCustomStr = "Ohm"
                         Case UCase("DB"):
                         Case UCase("Hz"):
                         Case UCase(""):
                         Case UCase("S"):
                         Case UCase("V"):
                         Case Else
                             m_UnitCustomStr = mVal
                             If mUnitCustomStr <> "" Then m_UnitCustomStr = m_UnitCustomStr
                     End Select
                     'Final:GetFlowParameters
                     ByTNum_GetFlowParameters = m_UnitCustomStr
                Else
                 
                 'Put the code for other case..

                End If
    Else
          Debug.Print "mKey=" & mKey & " Not Exists."
          Stop
    End If

          
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetRun(TName As String, Optional dbgPrint As Boolean = True) As Boolean
On Error GoTo Error
Dim funcName As String:: funcName = "GetRun"

Dim i              As Long
Dim mKey           As String
Dim mVal           As Variant
Dim Index          As Long
Dim mS_Temp        As String

    mKey = UCase(TName)
    If m_DicTName.Exists(mKey) Then
       Index = m_DicTName.Item(mKey)
        If UCase(LimitTable(Index).ParamSyntax.Row) = "RUN" Then
           GetRun = True
        Else
           GetRun = False
        End If
    Else
        mS_Temp = "TName =" & mKey & "  Is Not Existing In " & gC_ProjectLimitSheet & "."
       If dbgPrint Then TheExec.Datalog.WriteComment mS_Temp
       GetRun = False
    End If
    
'    If m_DicRun.Exists(mKey) Then
'       If UCase(m_DicRun.Item(mKey)) = "RUN" Then
'        GetRun = True
'       Else
'        GetRun = False
'       End If
'    Else
'        mS_Temp = "TName =" & mKey & "  Is Not Existing In " & gC_ProjectLimitSheet & "."
'       If dbgPrint Then TheExec.Datalog.WriteComment mS_Temp
'       GetRun = False
'    End If

Exit Function
Error:
Stop
GetRun = False
End Function



Public Function SheetDataList(Optional sRange As String = "A1:A5", _
                               Optional DataList As String = "", _
                               Optional mS_Sheet As String = "")
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "SheetDataList"
Dim mSheet As Worksheet
Dim mRange As Range

If mS_Sheet = "" Then
    Set mRange = Application.ActiveSheet.Range(sRange)
Else
    Set mSheet = Application.ActiveWorkbook.Sheets(mS_Sheet)
    Set mRange = mSheet.Range(sRange)
End If

With mRange.Validation
    .delete 'delete previous validation
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
                Formula1:=DataList
End With


Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetHiLoFromFlow(GetLowL As Variant, GetHighL As Variant, Optional Index As Long = 0)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "GetHiLoFromFlow"
Dim FlowLimitsInfo As IFlowLimitsInfo
Dim GetpVal() As String
Dim dLowL As Double
Dim dHighL As Double

Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
Call FlowLimitsInfo.GetLimits(Index, dLowL, dHighL)
GetLowL = dLowL
GetHighL = dHighL

Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function CheckTName(TName As String, Optional dbgPrint As Boolean = True) As Boolean
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "CheckTName"


Dim mKey As String
Dim mS_Temp As String
Dim mDictionary    As New Dictionary

mKey = UCase(Trim(TName))
    If m_DicTName.Exists(mKey) Then
       CheckTName = True
    Else
        mS_Temp = "TName =" & mKey & "  Is Not Existing In " & gC_ProjectLimitSheet & "."
       If dbgPrint Then TheExec.Datalog.WriteComment mS_Temp
       CheckTName = False
    End If
    
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetUserTName(TName As String, Optional dbgPrint As Boolean = False) As String
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "GetUserTName"

Dim i              As Long
Dim mKey           As String
Dim mVal           As Variant
Dim Index          As Long
Dim mS_Temp        As String

GetUserTName = TName
mKey = UCase(TName)

    If m_DicTName.Exists(mKey) Then
       Index = m_DicTName.Item(mKey)
       mVal = LimitTable(Index).ParamSyntax.UserTName
        If mVal <> "" Then
            GetUserTName = mVal
         End If
    Else
        mS_Temp = "TName =" & mKey & "  Is Not Existing In " & gC_ProjectLimitSheet & "."
       If dbgPrint Then TheExec.Datalog.WriteComment mS_Temp
    End If


'If m_DicUserTName.Exists(mKey) Then
'   mVal = m_DicUserTName.Item(mKey)
'   If mVal <> "" Then
'     GetUserTName = mVal
'   End If
'Else
'    mS_Temp = "TName =" & mKey & "  Is Not Existing In " & gC_ProjectLimitSheet & "."
'   If dbgPrint Then TheExec.Datalog.WriteComment mS_Temp
'End If

    
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetJobList(ByVal mWorkSheet As String) As String
Dim funcName As String:: funcName = "GetJobList"

Dim m_ws        As Worksheet
Dim mCol_Job    As Long
Dim mL_Row      As Long
Dim mL_Col      As Long
Dim mL_StartRow As Long
Dim mL_StartCol As Long
Dim mS_Temp     As String
Dim i           As Long
Dim Index       As Long
Dim mKey        As String
Dim mVal        As Variant
On Error GoTo TT


GetJobList = ""
mS_Temp = ""
Set m_ws = ActiveWorkbook.Worksheets(mWorkSheet)

 mCol_Job = 2: mL_StartRow = 5: Index = 0
 mKey = UCase(m_ws.Cells(mL_StartRow + Index, mCol_Job))
 
 While mKey <> ""
 mS_Temp = mS_Temp & "," & mKey
 Index = Index + 1
 mKey = UCase(m_ws.Cells(mL_StartRow + Index, mCol_Job))
 Wend
 
  If Len(mS_Temp) > 1 Then
    GetJobList = Mid(mS_Temp, 2, Len(mS_Temp) - 1)
  End If
  
Exit Function
TT:
 Debug.Print "WorkSheetIsNotExist:" & mWorkSheet

End Function
Public Function GetLowLimit(TName As String) As Variant
Dim funcName As String:: funcName = "GetLowLimit"

Dim i              As Long
Dim mKey           As String
Dim mVal           As Variant
Dim Index          As Long
Dim mS_Temp        As String

mKey = UCase(TName)

    If m_DicTName.Exists(mKey) Then
       Index = m_DicTName.Item(mKey)
        If CStr(LimitTable(Index).ParamSyntax.LoLimit) Like "*N*A*" Then
            GetLowLimit = "N/A"
        Else
            GetLowLimit = LimitTable(Index).ParamSyntax.LoLimit
        End If
'       If CStr(GetLowLimit) Like "*N*A*" Then GoTo Error
                
    Else
       Debug.Print "mKey=" & mKey & " Not Exists."
        GetLowLimit = -9999
    End If
    

Exit Function
Error:
mS_Temp = funcName & ":: Get test limit = " & CStr(GetLowLimit) & " failed: TName= '" & TName & ". Please edit it in limit table."
TheExec.AddOutput (mS_Temp)
'TheExec.Datalog.WriteComment (mS_Temp)
'Stop
GetLowLimit = -9999
End Function

Public Function GetTNum(TName As String) As Variant
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "GetTNum"

Dim i              As Long
Dim mKey           As String
Dim mVal           As Variant
Dim Index          As Long
Dim mS_Temp        As String

    mKey = UCase(TName)
    If m_DicTName.Exists(mKey) Then
       Index = m_DicTName.Item(mKey)
       GetTNum = LimitTable(Index).ParamSyntax.TestNumber
    Else
       Debug.Print "mKey=" & mKey & " Not Exists."
       GetTNum = 9999
    End If
    
    
'    If m_DicTNum.Exists(mKey) Then
'       GetTNum = m_DicTNum.Item(mKey)
'    Else
'       Debug.Print "mKey=" & mKey & " Not Exists."
'       GetTNum = 9999
'    End If
       
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function CheckTNameFlag(ByVal TName As String) As Boolean
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "CheckTNameFlag"
Dim i                         As Long
Dim CheckCondition()          As String
Dim CheckRemoveFlag           As Boolean
Dim CheckTNameListFlag        As Boolean
Dim TNameGrp()                As String
Dim TNameGrp7                 As String

CheckTNameFlag = False
CheckTNameListFlag = False
CheckCondition = Split(ListLimitTName, ",")
TNameGrp = Split(TName, TNameSplit)

If UBound(TNameGrp) < 11 Then
  Exit Function
Else
  TNameGrp7 = TNameGrp(6)
End If


For i = 0 To UBound(CheckCondition)
    'CheckRemoveFlag = InStr(UCase(TNameGrp7), CheckCondition(i))
    If UCase(TNameGrp7) = UCase(CheckCondition(i)) Then
        CheckTNameListFlag = True
        Exit For
    End If
        
Next i

CheckTNameFlag = CheckTNameListFlag



Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Check_LimitSheet_MP_DataLog_ON_To_Skip(ByVal TName As String) As Boolean
''-----20191206 Add for "LimitSheet_MP_DataLog_ON"
On Error GoTo ErrHandler


Dim funcName As String:: funcName = "Check_LimitSheet_MP_DataLog_ON_To_Skip"
Dim i                       As Long
Dim Check_Boolean           As Boolean
Dim TNameTemp               As String

Check_LimitSheet_MP_DataLog_ON_To_Skip = False
Check_Boolean = True

For i = 0 To UBound(g_sLimitSheet_MP_DataLog_array)
    If InStr(UCase(TName), UCase(g_sLimitSheet_MP_DataLog_array(i))) > 0 Then
        Check_Boolean = False
        Exit For
    End If
Next i

Check_LimitSheet_MP_DataLog_ON_To_Skip = Check_Boolean


Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MP_Datalog_FlagOn() As Boolean
''-----20191206 Add for "LimitSheet_MP_DataLog_ON"
On Error GoTo ErrHandler

g_LimitSheet_MP_DataLog_ON = True

Exit Function
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MP_Datalog_FlagOff() As Boolean
''-----20191206 Add for "LimitSheet_MP_DataLog_ON"
On Error GoTo ErrHandler

g_LimitSheet_MP_DataLog_ON = False

Exit Function
ErrHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function UpdateFlowLimit(Optional RunTimes As Long = 0) As Boolean
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "UFlowLimit"
 
Dim mSheet          As Worksheet
Dim mCol_TName      As Long
Dim mCol_Tnum       As Long
Dim mCol_HiLimit   As Long
Dim mCol_LowLimit   As Long

Dim mCol_Run        As Long
Dim mCol_flowtable  As Long
Dim mCol_instance   As Long
Dim mCol_testname   As Long
Dim mCol_testnumber As Long
Dim mCol_UserTName  As Long
Dim mCol_highcompsign As Long
Dim mCol_lolim()    As Long
Dim mCol_hilim()    As Long
Dim mCol_scale      As Long
Dim mCol_units      As Long
Dim mCol_format     As Long
Dim mCol_passsort   As Long
Dim mCol_failsort   As Long
Dim mCol_passbin    As Long
Dim mCol_failbin    As Long
Dim mCol_result     As Long
Dim mCol_passaction As Long
Dim mCol_failaction As Long
Dim mCol_assume     As Long
Dim mCol_sites      As Long
Dim mCol_comment    As Long


Dim mS_Job()            As String
Dim mL_JobNum           As Long
Dim mRow_JobList        As Long
Dim mCol_JobListStart   As Long
Dim mRow_Header         As Long


Dim mCellName           As String
Dim mCheckString        As String
Dim mCheckFalg          As Boolean

Dim i           As Long
Dim Index       As Long
Dim Index_bak       As Long

Dim mKey        As String
Dim mKey_TNum   As Double
Dim mVal        As Variant
Dim mL_Row      As Long
Dim mL_Col      As Long
Dim mL_Range    As Long


Dim mCell As Object

Dim sheet_index As Double
Dim All_Sheet_arr() As String
Dim All_Sheet_number As String
Dim PorjectSheetName As String
All_Sheet_arr = Split(gC_ProjectLimitSheet, ",")
All_Sheet_number = UBound(All_Sheet_arr) + 1
Dim dd As Variant
Dim TNum_arr() As Long
'---------------------------

    If True Then
        m_DicTNum.RemoveAll
        m_DicTName.RemoveAll
    End If
    
    
    Index = 0
    Index_bak = 0
    For sheet_index = 0 To All_Sheet_number - 1 Step 1
        PorjectSheetName = All_Sheet_arr(sheet_index)
        Set mSheet = Application.ActiveWorkbook.Sheets(PorjectSheetName)
        
        If RunTimes <> 0 Then GoTo SkipReadTable
        
        mCheckFalg = False
        For i = 10 To 40   ' to find Current run job name in where and information location
            mCellName = UCase(mSheet.Cells(2, i))
            If mCellName = UCase(g_CurrentLimitSet) Then
                mCol_Run = 2 'row
                mCol_TName = 5: mCol_UserTName = 6: mCol_Tnum = 7: mCol_scale = 8: mCol_units = 9
                mCol_LowLimit = i: mCol_HiLimit = i + 1
                mCheckFalg = True
                If mCheckFalg Then
                    mS_Job = Split(gS_JobList, ",")
                    mL_JobNum = UBound(mS_Job) + 1
                    'Define the Row and Col.
                    i = 10 + 2 * (mL_JobNum - 1)
                End If
                Exit For
            End If
        Next i
        
        If mCheckFalg = False Then Debug.Print "Can not find the g_CurrentLimitSet=" & g_CurrentLimitSet
        
        i = 4
        
        mKey = Trim(UCase(mSheet.Cells(i, mCol_TName)))
        mKey_TNum = mSheet.Cells(i, mCol_Tnum)
        
        Do While Len(mKey) > 0
            mKey = UCase(mSheet.Cells(i, mCol_TName))
            mKey = Trim(mKey)
            
            If Len(mKey) = 0 Then
                Exit Do
            End If
            mKey_TNum = mSheet.Cells(i, mCol_Tnum)
            mKey_TNum = Trim(mKey_TNum)
            
            
            If m_DicTName.Exists(mKey) = False Then
                 
                 ReDim Preserve LimitTable(Index)
                 LimitTable(Index).Index = Index
                 LimitTable(Index).TName = mKey
                 LimitTable(Index).TNum = mKey_TNum
                 LimitTable(Index).CurrLimitSet = UCase(g_CurrentLimitSet)
                 LimitTable(Index).CurrHiLimit_Col = mCol_HiLimit
                 LimitTable(Index).CurrLoLimit_Col = mCol_LowLimit
                 LimitTable(Index).TNum = mSheet.Cells(i, mCol_Tnum)
        
                mVal = mSheet.Cells(i, mCol_HiLimit): LimitTable(Index).ParamSyntax.HiLimit = mVal
                mVal = mSheet.Cells(i, mCol_LowLimit): LimitTable(Index).ParamSyntax.LoLimit = mVal
                mVal = mSheet.Cells(i, mCol_Tnum): LimitTable(Index).ParamSyntax.TestNumber = mVal
                mVal = mSheet.Cells(i, mCol_Run): LimitTable(Index).ParamSyntax.Row = mVal
                mVal = mSheet.Cells(i, mCol_TName): LimitTable(Index).ParamSyntax.TestName = mVal
                mVal = mSheet.Cells(i, mCol_UserTName): LimitTable(Index).ParamSyntax.UserTName = mVal
                mVal = mSheet.Cells(i, mCol_scale): LimitTable(Index).ParamSyntax.UScale = mVal
                mVal = mSheet.Cells(i, mCol_units): LimitTable(Index).ParamSyntax.Units = mVal
                
                m_DicTName.Add mKey, Index
                
                
                '20190222 Check Test Number Duplication
                If gb_TNameTNum_Dupl_CHK = True Then
                    If m_DicTNum.Exists(mKey_TNum) = False Then
                        m_DicTNum.Add mKey_TNum, Index
                        Index = Index + 1
                    Else
                        Call Dupl_TNum_Check(mKey_TNum)
                    End If
                Else
                    Index = Index + 1
                End If
                
            Else
                Debug.Print "mKey=" & mKey & " Exists."
                TheExec.AddOutput "Warning: Please check sheet" & PorjectSheetName & " The TestName: " & mKey & " is exist!!!"
                
                '20190222 Check Test Name Duplication
                If gb_TNameTNum_Dupl_CHK = True Then
                    Call Dupl_TName_Check(mKey)
                End If

            End If
            i = i + 1
            'mKey = UCase(mSheet.Cells(i, mCol_TName))
            If Len(mKey) = 0 Then
                Exit Do
            End If
        Loop   '-- Wend
    Next sheet_index
    
'    If m_DicTName.Count > 0 Then
'      Debug.Print "m_DicTName.Count=" & m_DicTName.Count
'      Debug.Print "m_DicTNum.Count=" & m_DicTNum.Count
'      Debug.Print "LimitTable.Count=" & UBound(LimitTable) + 1
'    End If
       
    UpdateFlowLimit = True
    Exit Function
    
    '-------------------------------------------------------------------------
SkipReadTable:
    ''TName:
    For i = 0 To UBound(LimitTable)
        Index = i
        mKey = UCase(LimitTable(Index).TName)
        mVal = LimitTable(Index).Index
        If m_DicTName.Exists(mKey) = False And mKey <> "" Then
           If mVal = Index Then
             m_DicTName.Add mKey, mVal
            Else
              Stop 'ERROR!
            End If
        Else
            Debug.Print "mKey=" & mKey & " Exists."
        End If
    Next i

       If m_DicTName.Count > 0 Then
         Debug.Print "m_DicTName.Count=" & m_DicTName.Count
         Debug.Print "m_DicTNum.Count=" & m_DicTNum.Count
         Debug.Print "LimitTable.Count=" & UBound(LimitTable) + 1
       End If
    UpdateFlowLimit = True
'-------------------------------------------------------------------------


Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Sub DemoLimitSet()
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "DemoLimitSet"

'===================================
Dim lTestNumber        As Long
Dim lNumberOfLimitSets As Long
Dim strTestNames()     As String
Dim dLowLimits()       As Double
Dim dHighLimits()      As Double
Dim stScale()          As tlScaleType
Dim utMeasUnit()       As UnitType
'=====================================
Dim strTestName As String
Dim strDatalogType As String
Dim plMeasPin As New PinList
Dim sdMeasValue(2) As New SiteDouble
Dim i As Long
plMeasPin.Value = "AMUX-AX"

sdMeasValue(0) = 0.001: sdMeasValue(1) = 0.002: sdMeasValue(2) = 0.003

lTestNumber = 10110: lNumberOfLimitSets = 3
LL.ByTNum_GetLimitDetailsFromLimitTable lTestNumber, strTestNames, dLowLimits, dHighLimits, stScale, utMeasUnit, lNumberOfLimitSets
    
Call TheExec.Flow.TestLimit(ResultVal:=sdMeasValue(0), TName:=strTestNames(0), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Unit:=utMeasUnit(0), ForceResults:=tlForceNone, PinName:=plMeasPin, lowVal:=dLowLimits(0), hiVal:=dHighLimits(0))
Call TheExec.Flow.TestLimit(ResultVal:=sdMeasValue(1), TName:=strTestNames(1), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Unit:=utMeasUnit(1), ForceResults:=tlForceNone, PinName:=plMeasPin, lowVal:=dLowLimits(1), hiVal:=dHighLimits(1))
Call TheExec.Flow.TestLimit(ResultVal:=sdMeasValue(2), TName:=strTestNames(2), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Unit:=utMeasUnit(2), ForceResults:=tlForceNone, PinName:=plMeasPin, lowVal:=dLowLimits(2), hiVal:=dHighLimits(2))
               

'USING GLOBAL PARAMETERS:
g_lTestNumber = 10140: g_lNumberOfLimitSets = 3
LL.ByTNum_GetLimitDetailsFromLimitTable g_lTestNumber, g_strTestNames, g_dLowLimits, g_dHighLimits, g_stScale, g_utMeasUnit, g_lNumberOfLimitSets
Call TheExec.Flow.TestLimit(ResultVal:=sdMeasValue(0), TName:=g_strTestNames(0), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Unit:=g_utMeasUnit(0), ForceResults:=tlForceNone, PinName:=plMeasPin, lowVal:=g_dLowLimits(0), hiVal:=g_dHighLimits(0))
Call TheExec.Flow.TestLimit(ResultVal:=sdMeasValue(1), TName:=g_strTestNames(1), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Unit:=g_utMeasUnit(1), ForceResults:=tlForceNone, PinName:=plMeasPin, lowVal:=g_dLowLimits(1), hiVal:=g_dHighLimits(1))
Call TheExec.Flow.TestLimit(ResultVal:=sdMeasValue(2), TName:=g_strTestNames(2), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Unit:=g_utMeasUnit(2), ForceResults:=tlForceNone, PinName:=plMeasPin, lowVal:=g_dLowLimits(2), hiVal:=g_dHighLimits(2))
         


Exit Sub
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Public Function ToGetFlowParameters(GetFlowPara As gutTestlimit, mUnitType As UnitType, mScaleType As tlScaleType, Optional mUnitCustomStr As String) As Variant
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "ToGetFlowParameters"
Dim Dlog         As String
Dim TableColumn  As String
Dim dScaleFactor As Double

If GetFlowPara = gut_UnitType Then

        Select Case mUnitType
            Case unitAmp:       'Const unitAmp = 1
                 Dlog = "unitAmp"
                 TableColumn = "A"
            Case unitCustom:    'Const unitCustom = 8
                 Dlog = "unitCustom"
                 TableColumn = "unitCustom"
                 TableColumn = "Ohm"
                 If mUnitCustomStr <> "" Then TableColumn = mUnitCustomStr
            Case unitDb:        'Const unitDb = 5
                 Dlog = "unitDb"
                 TableColumn = "DB"
            Case unitHz:        'Const unitHz = 3
                 Dlog = "unitHz"
                 TableColumn = "Hz"
            Case unitNone:      'Const unitNone = 0
                 Dlog = "unitNone"
                 TableColumn = ""
            Case unitTime:      'Const unitTime = 4
                 Dlog = "unitTime"
                 TableColumn = "S"
            Case unitVolt:      'Const unitVolt = 2
                 Dlog = "unitVolt"
                 TableColumn = "V"
            Case Else
                 Stop
        End Select
        
ElseIf GetFlowPara = gut_ScaleType Then

        Select Case mScaleType
            Case scaleFemto:        'Const scaleFemto = 15
                 Dlog = "scaleFemto"
                 TableColumn = "f"
                 dScaleFactor = 0.000000000000001
            Case scaleGiga:         'Const scaleGiga = -9 (&HFFFFFFF7)
                 Dlog = "scaleGiga"
                 TableColumn = "G"
                 dScaleFactor = 1000000000#
            Case scaleKilo:         'Const scaleKilo = -3 (&HFFFFFFFD)
                 Dlog = "scaleKilo"
                 TableColumn = "K"
                 dScaleFactor = 1000#
            Case scaleMega:         'Const scaleMega = -6 (&HFFFFFFFA)
                 Dlog = "scaleMega"
                 TableColumn = "M"
                 dScaleFactor = 1000000#
            Case scaleMicro:        'Const scaleMicro = 6
                 Dlog = "scaleMicro"
                 TableColumn = "u"
                 dScaleFactor = 0.000001
            Case scaleMilli:        'Const scaleMilli = 3
                 Dlog = "scaleMilli"
                 TableColumn = "m"
                 dScaleFactor = 0.001
            Case scaleNano:         'Const scaleNano = 9
                 Dlog = "scaleNano"
                 TableColumn = "n"
                 dScaleFactor = 0.000000001
            Case scaleNone:         'Const scaleNone = 0
                 Dlog = "scaleNone"
                 TableColumn = ""
            Case scaleNoScaling:    'Const scaleNoScaling = 9999 (&H270F)
                 Dlog = "scaleNoScaling"
                 TableColumn = ""
                 dScaleFactor = 1#
            Case scalePercent:      'Const scalePercent = 2
                 Dlog = "scalePercent"
                 TableColumn = "%"
            Case scalePico:         'Const scalePico = 12
                 Dlog = "scalePico"
                 TableColumn = "p"
                 dScaleFactor = 0.00000000001
            Case scaleTera:         'Const scaleTera = -12 (&HFFFFFFF4)
                 Dlog = "scaleTera"
                 TableColumn = "T"
                 dScaleFactor = 1000000000000#
                 
            Case Else
                 Stop
        End Select
        
Else


End If
    

ToGetFlowParameters = TableColumn

           
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ByTNum_GetLimitDetailsFromLimitTable(ByRef lTestNumber As Long, _
                                                            strTestNames() As String, _
                                                            dLowLimits() As Double, _
                                                            dHighLimits() As Double, _
                                                            stScaleTypes() As tlScaleType, _
                                                            utUnitTypes() As UnitType, _
                                                             Optional lNumberOfLimitSets As Long = 1)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "ByTNum_GetLimitDetailsFromLimitTable"

Dim lIdx As Long
ReDim dLowLimits(lNumberOfLimitSets - 1)
ReDim dHighLimits(lNumberOfLimitSets - 1)
ReDim strTestNames(lNumberOfLimitSets - 1)
ReDim stScaleTypes(lNumberOfLimitSets - 1)
ReDim utUnitTypes(lNumberOfLimitSets - 1)

Dim i              As Long
Dim mTNumKey       As Long
Dim mVal           As Variant
Dim Index          As Long
Dim mS_Temp        As String

For lIdx = 0 To lNumberOfLimitSets - 1
    mTNumKey = lTestNumber + lIdx
    If m_DicTNum.Exists(UCase(mTNumKey)) Then
        Index = m_DicTNum.Item(UCase(mTNumKey))
        dLowLimits(lIdx) = LimitTable(Index).ParamSyntax.LoLimit
        dHighLimits(lIdx) = LimitTable(Index).ParamSyntax.HiLimit
        strTestNames(lIdx) = LimitTable(Index).ParamSyntax.TestName
        
        stScaleTypes(lIdx) = ByTNum_GetFlowParameters(mTNumKey, gut_ScaleType)
        utUnitTypes(lIdx) = ByTNum_GetFlowParameters(mTNumKey, gut_UnitType)
    Else
       TheExec.AddOutput ("Get test limit details failed: Test number '" & lTestNumber & "' does not exist in test limit array check your test number aginst the test numbers in limit table")
    End If
Next lIdx

Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UpdateTestNameSheet(ByVal mS_Sheet As String, _
                                    ByVal mTName As String, _
                                    ByVal mLLimit As Variant, _
                                    ByVal mHLimit As Variant, _
                                    ByVal mTNum As Long, _
                                    Optional ByVal mUserTName As String = "", _
                                    Optional ByVal mUnitType As UnitType, _
                                    Optional ByVal mScaleType As tlScaleType, _
                                    Optional ByVal mUnitCustomStr As String, _
                                    Optional ByVal index_i As String) As Boolean
                                    
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "UpdateTestNameSheet"
 
Dim mCol_TName      As Long

Dim mCol_Run        As Long
Dim mCol_flowtable  As Long
Dim mCol_instance   As Long
Dim mCol_testname   As Long
Dim mCol_testnumber As Long
Dim mCol_UserTName  As Long
Dim mCol_highcompsign As Long
Dim mCol_lolim()    As Long
Dim mCol_hilim()    As Long
Dim mCol_scale      As Long
Dim mCol_units      As Long
Dim mCol_format     As Long
Dim mCol_passsort   As Long
Dim mCol_failsort   As Long
Dim mCol_passbin    As Long
Dim mCol_failbin    As Long
Dim mCol_result     As Long
Dim mCol_passaction As Long
Dim mCol_failaction As Long
Dim mCol_assume     As Long
Dim mCol_sites      As Long
Dim mCol_comment    As Long


Dim mVal_Run        As Variant
Dim mVal_flowtable  As Variant
Dim mVal_instance   As Variant
Dim mVal_testname   As Variant
Dim mVal_testnumber As Variant
Dim mVal_UserTName    As Variant
Dim mVal_highcompsign   As Variant
Dim mVal_lolim()    As Variant
Dim mVal_hilim()    As Variant
Dim mVal_scale  As Variant
Dim mVal_units  As Variant
Dim mVal_format As Variant
Dim mVal_passsort   As Variant
Dim mVal_failsort   As Variant
Dim mVal_passbin    As Variant
Dim mVal_failbin    As Variant
Dim mVal_result As Variant

Dim mS_Job()            As String
Dim mL_JobNum           As Long
Dim mRow_JobList        As Long
Dim mCol_JobListStart   As Long
Dim mRow_Header   As Long

Dim mSheet As Worksheet
Dim FlowLimitsInfo As IFlowLimitsInfo
Dim i           As Long
Dim mL_Row      As Long



Dim Index As Long
Dim mCell As Object
Dim mTestMode As gtlTestMode
 
mTestMode = TheExec.Datalog.Setup.LotSetup.TESTMODE
If mTestMode <> gtl_Engineeringmode Then Exit Function
If mUserTName = "" Then mUserTName = mTName

mS_Job = Split(gS_JobList, ",")
mL_JobNum = UBound(mS_Job) + 1
ReDim mCol_lolim(mL_JobNum - 1)
ReDim mCol_hilim(mL_JobNum - 1)
ReDim mVal_lolim(mL_JobNum - 1)
ReDim mVal_hilim(mL_JobNum - 1)

'Define the Row and Col.
mRow_JobList = 2
mRow_Header = 3

mCol_Run = 2
mCol_flowtable = 3
mCol_instance = 4
mCol_testname = 5
mCol_UserTName = 6
mCol_testnumber = 7

mCol_scale = 8
mCol_units = 9
        
For i = 0 To mL_JobNum - 1
    mCol_lolim(i) = 10 + 2 * i
    mCol_hilim(i) = 11 + 2 * i
Next i



Set mSheet = Application.ActiveWorkbook.Sheets(mS_Sheet)
mVal_flowtable = TheExec.Flow.CurrentFlowSheetName
mVal_instance = TheExec.DataManager.InstanceName
mVal_testname = mTName
mVal_testnumber = mTNum
If mTNum = 0 Then mVal_testnumber = TheExec.Datalog.LastTestNumLogged


For i = 0 To mL_JobNum - 1
mVal_lolim(i) = mLLimit
mVal_hilim(i) = mHLimit
Next i


If UCase(mS_Sheet) Like "*UPDATE*" Then
    If UBound(Split(gC_Sheet_UpdateLimit, ",")) > 0 Then
        
        If gL_UpdateTestMulti(index_i) = 0 Then
        TheExec.Datalog.WriteComment vbCrLf & UCase("Please update the limit sheet.")
        ResetSheet mS_Sheet
        End If
        gL_UpdateTestMulti(index_i) = gL_UpdateTestMulti(index_i) + 1
        i = 3 + gL_UpdateTestMulti(index_i)
        
    Else
        If gL_UpdateTest = 0 Then
        TheExec.Datalog.WriteComment vbCrLf & UCase("Please update the limit sheet.")
        ResetSheet mS_Sheet
        End If
        gL_UpdateTest = gL_UpdateTest + 1
        i = 3 + gL_UpdateTest
    End If
Else
    If gL_CurrTest = 0 Then
    ResetSheet mS_Sheet
    End If
    gL_CurrTest = gL_CurrTest + 1
    i = 3 + gL_CurrTest
End If

'------------------------------------------
mSheet.Cells(i, mCol_flowtable) = mVal_flowtable
mSheet.Cells(i, mCol_instance) = mVal_instance
mSheet.Cells(i, mCol_testname) = mVal_testname
mSheet.Cells(i, mCol_testnumber) = mVal_testnumber

mSheet.Cells(i, mCol_scale) = ToGetFlowParameters(gut_ScaleType, mUnitType, mScaleType, mUnitCustomStr)
mSheet.Cells(i, mCol_units) = ToGetFlowParameters(gut_UnitType, mUnitType, mScaleType, mUnitCustomStr)


         
If (UCase(mS_Sheet) Like "*UPDATE*") Or bEnableLogAllJob = True Then
        For Index = 0 To mL_JobNum - 1
            mSheet.Cells(i, mCol_lolim(Index)) = mVal_lolim(Index)
            mSheet.Cells(i, mCol_hilim(Index)) = mVal_hilim(Index)
        Next Index
Else
     For Index = 0 To mL_JobNum - 1
            mRow_JobList = 2
         If UCase(mSheet.Cells(mRow_JobList, mCol_lolim(Index))) = (TheExec.CurrentJob) Then
            mSheet.Cells(i, mCol_lolim(Index)) = mVal_lolim(Index)
            mSheet.Cells(i, mCol_hilim(Index)) = mVal_hilim(Index)
        End If
    Next Index

End If
'------------------------------------------
If (UCase(mS_Sheet) Like "*UPDATE*") Or bEnableLogAllJob = True Then
        mVal_Run = "Run"
        mSheet.Cells(i, mCol_Run) = mVal_Run
Else
     If GetRun(mTName, False) = True Then
        mVal_Run = "Run"
        mSheet.Cells(i, mCol_Run) = mVal_Run
        If mTName <> mUserTName Then
         mSheet.Cells(i, mCol_UserTName) = mUserTName
        End If
   End If
End If

Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UpdateTestNameTextFile(ByVal mF_File As String, _
                                    ByVal mTName As String, _
                                    ByVal mLLimit As Variant, _
                                    ByVal mHLimit As Variant, _
                                    ByVal mTNum As Long, _
                                    Optional ByVal mUserTName As String = "", _
                                    Optional ByVal mUnitType As UnitType, _
                                    Optional ByVal mScaleType As tlScaleType, _
                                    Optional ByVal mUnitCustomStr As String, _
                                    Optional ByVal index_i As String) As Boolean
                                    
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "UpdateTestNameSheet"
 
Dim mCol_TName      As Long

Dim mCol_Run        As Long
Dim mCol_flowtable  As Long
Dim mCol_instance   As Long
Dim mCol_testname   As Long
Dim mCol_testnumber As Long
Dim mCol_UserTName  As Long
Dim mCol_highcompsign As Long
Dim mCol_lolim()    As Long
Dim mCol_hilim()    As Long
Dim mCol_scale      As Long
Dim mCol_units      As Long
Dim mCol_format     As Long
Dim mCol_passsort   As Long
Dim mCol_failsort   As Long
Dim mCol_passbin    As Long
Dim mCol_failbin    As Long
Dim mCol_result     As Long
Dim mCol_passaction As Long
Dim mCol_failaction As Long
Dim mCol_assume     As Long
Dim mCol_sites      As Long
Dim mCol_comment    As Long



Dim mVal_Run        As Variant
Dim mVal_flowtable  As Variant
Dim mVal_instance   As Variant
Dim mVal_testname   As Variant
Dim mVal_testnumber As Variant
Dim mVal_UserTName    As Variant
Dim mVal_highcompsign   As Variant
Dim mVal_lolim()    As Variant
Dim mVal_hilim()    As Variant
Dim mVal_scale  As Variant
Dim mVal_units  As Variant
Dim mVal_format As Variant
Dim mVal_passsort   As Variant
Dim mVal_failsort   As Variant
Dim mVal_passbin    As Variant
Dim mVal_failbin    As Variant
Dim mVal_result As Variant
Dim mTXT_input()    As String ''''''''''''''''''''''''''''''''20180927
Dim mCurrentTXTinput As String
Dim DeleteFileName  As String

Dim mS_Job()            As String
Dim mL_JobNum           As Long
Dim mRow_JobList        As Long
Dim mCol_JobListStart   As Long
Dim mRow_Header   As Long

Dim mSheet As Worksheet
Dim FlowLimitsInfo As IFlowLimitsInfo
Dim i           As Long
Dim mL_Row      As Long



Dim Index As Long
Dim mCell As Object
Dim mTestMode As gtlTestMode

mF_File = mF_File & ".txt" '- ----------------------------------------------------20180927

 
mTestMode = TheExec.Datalog.Setup.LotSetup.TESTMODE
If mTestMode <> gtl_Engineeringmode Then Exit Function
If mUserTName = "" Then mUserTName = mTName

mS_Job = Split(gS_JobList, ",")
mL_JobNum = UBound(mS_Job) + 1
ReDim mCol_lolim(mL_JobNum - 1)
ReDim mCol_hilim(mL_JobNum - 1)
ReDim mVal_lolim(mL_JobNum - 1)
ReDim mVal_hilim(mL_JobNum - 1)

'Define the Row and Col.
mRow_JobList = 2
mRow_Header = 3

mCol_Run = 2
mCol_flowtable = 3
mCol_instance = 4
mCol_testname = 5
mCol_testnumber = 6
mCol_UserTName = 7
mCol_highcompsign = 8

For i = 0 To mL_JobNum - 1
    mCol_lolim(i) = 9 + 2 * i
    mCol_hilim(i) = 10 + 2 * i
    If i = (mL_JobNum - 1) Then
        mCol_scale = mCol_hilim(i) + 1
        mCol_units = mCol_hilim(i) + 2
        mCol_format = mCol_hilim(i) + 3
        mCol_passsort = mCol_hilim(i) + 4
        mCol_failsort = mCol_hilim(i) + 5
        mCol_passbin = mCol_hilim(i) + 6
        mCol_failbin = mCol_hilim(i) + 7
        mCol_result = mCol_hilim(i) + 8
    End If
Next i



'Set mSheet = Application.ActiveWorkbook.Sheets(mS_Sheet)
mVal_flowtable = TheExec.Flow.CurrentFlowSheetName
mVal_instance = TheExec.DataManager.InstanceName
mVal_testname = mTName
mVal_testnumber = mTNum
If mTNum = 0 Then mVal_testnumber = TheExec.Datalog.LastTestNumLogged


For i = 0 To mL_JobNum - 1
mVal_lolim(i) = mLLimit
mVal_hilim(i) = mHLimit
Next i



    
Dim file_temp As Object
Dim FileExists As Boolean
DeleteFileName = TheExec.TestProgram.Path & "\" & mF_File
Set file_temp = CreateObject("Scripting.filesystemobject")
FileExists = file_temp.FileExists(DeleteFileName)


If UCase(mF_File) Like "*UPDATE*" Then
    If UBound(Split(gC_Sheet_UpdateLimit, ",")) > 0 Then
        
        If gL_UpdateTestMulti(index_i) = 0 Then
        TheExec.Datalog.WriteComment vbCrLf & UCase("Please update the limit sheet.")
        'ResetSheet mS_Sheet
            If FileExists Then
                Kill DeleteFileName
                LL.LimitTextHeader gS_JobList, mF_File
            End If
        End If
        gL_UpdateTestMulti(index_i) = gL_UpdateTestMulti(index_i) + 1
        i = 3 + gL_UpdateTestMulti(index_i)
        
    Else
        If gL_UpdateTest = 0 Then
        TheExec.Datalog.WriteComment vbCrLf & UCase("Please update the limit sheet.")
        'ResetSheet mS_Sheet
            If FileExists Then
                Kill DeleteFileName
                LL.LimitTextHeader gS_JobList, mF_File
            End If
        End If
        gL_UpdateTest = gL_UpdateTest + 1
        i = 3 + gL_UpdateTest
    End If
Else
    If gL_CurrTest = 0 Then
    'ResetSheet mS_Sheet
            If FileExists Then
                Kill DeleteFileName
                LL.LimitTextHeader gS_JobList, mF_File
            Else
                LL.LimitTextHeader gS_JobList, mF_File
            End If
    End If
    gL_CurrTest = gL_CurrTest + 1
    i = 3 + gL_CurrTest
End If

'------------------------------------------
'mSheet.Cells(i, mCol_flowtable) = mVal_flowtable
'mSheet.Cells(i, mCol_instance) = mVal_instance
'mSheet.Cells(i, mCol_testname) = mVal_testname
'mSheet.Cells(i, mCol_instance) = mVal_instance
'mSheet.Cells(i, mCol_testnumber) = mVal_testnumber
'
'mSheet.Cells(i, mCol_scale) = ToGetFlowParameters(gut_ScaleType, mUnitType, mScaleType, mUnitCustomStr)
'mSheet.Cells(i, mCol_units) = ToGetFlowParameters(gut_UnitType, mUnitType, mScaleType, mUnitCustomStr)

'----------------------------------------------------------------------------------------------------20180927
ReDim mTXT_input(mCol_result - 1)
mTXT_input(mCol_flowtable - 1) = mVal_flowtable
mTXT_input(mCol_instance - 1) = mVal_instance
mTXT_input(mCol_testname - 1) = mVal_testname
mTXT_input(mCol_instance - 1) = mVal_instance
mTXT_input(mCol_testnumber - 1) = mVal_testnumber
mTXT_input(mCol_scale - 1) = ToGetFlowParameters(gut_ScaleType, mUnitType, mScaleType, mUnitCustomStr)
mTXT_input(mCol_units - 1) = ToGetFlowParameters(gut_UnitType, mUnitType, mScaleType, mUnitCustomStr)

         
If (UCase(mF_File) Like "*UPDATE*") Or bEnableLogAllJob = True Then
        For Index = 0 To mL_JobNum - 1
'            mSheet.Cells(i, mCol_lolim(Index)) = mVal_lolim(Index)
'            mSheet.Cells(i, mCol_hilim(Index)) = mVal_hilim(Index)
            mTXT_input(mCol_lolim(Index) - 1) = mVal_lolim(Index)
            mTXT_input(mCol_hilim(Index) - 1) = mVal_hilim(Index)
        Next Index
Else
     For Index = 0 To mL_JobNum - 1
            mRow_JobList = 2
         'If UCase(mSheet.Cells(mRow_JobList, mCol_lolim(Index))) = (TheExec.CurrentJob) Then
                  
        If g_mTXT_input_JobList(mCol_lolim(Index) - 1) = (TheExec.CurrentJob) Then              '---------------------------------------------------------------20180929
         
'            mSheet.Cells(i, mCol_lolim(Index)) = mVal_lolim(Index)
'            mSheet.Cells(i, mCol_hilim(Index)) = mVal_hilim(Index)
            mTXT_input(mCol_lolim(Index) - 1) = mVal_lolim(Index)
            mTXT_input(mCol_hilim(Index) - 1) = mVal_hilim(Index)
        End If
    Next Index

End If
'------------------------------------------
If (UCase(mF_File) Like "*UPDATE*") Or bEnableLogAllJob = True Then
        mVal_Run = "Run"
        'mSheet.Cells(i, mCol_Run) = mVal_Run
        mTXT_input(mCol_Run - 1) = mVal_Run
Else
     If GetRun(mTName, False) = True Then
        mVal_Run = "Run"
        'mSheet.Cells(i, mCol_Run) = mVal_Run
        If mTName <> mUserTName Then
         'mSheet.Cells(i, mCol_UserTName) = mUserTName
         mTXT_input(mCol_UserTName - 1) = mUserTName
        End If
   End If
End If


mCurrentTXTinput = Join(mTXT_input, vbTab)
SaveStringtoFile mF_File, mCurrentTXTinput



Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function ResetSheet(mS_Sheet As String) As Boolean
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "ResetSheet"

Dim mSheet As Worksheet
Set mSheet = Application.ActiveWorkbook.Sheets(mS_Sheet)
 
'mSheet.Range("A4", "AT60000").Clear

Dim usedrow, UsedColumn As Long
usedrow = mSheet.UsedRange.Rows.Count
UsedColumn = mSheet.UsedRange.Columns.Count
 
mSheet.Activate

ActiveSheet.Range(Cells(4, 1), Cells(usedrow, UsedColumn)).Clear

LL.SheetDataList "AD4:AD60000", SupportUnitType, mS_Sheet 'Limitsheet:units==>"V,A,Ohm,,DB,Hz,S,V"
   
   
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function CreatMainSheet_form_Multisubsheet(Optional RunTimes As Long = 0) As Boolean
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "CreatMainSheet_form_Multisubsheet"
 
Dim mSheet          As Worksheet
Dim mCol_TName      As Long
Dim mCol_Tnum       As Long
Dim mCol_HiLimit   As Long
Dim mCol_LowLimit   As Long

Dim mCol_Run        As Long
Dim mCol_flowtable  As Long
Dim mCol_instance   As Long
Dim mCol_testname   As Long
Dim mCol_testnumber As Long
Dim mCol_UserTName  As Long
Dim mCol_highcompsign As Long
Dim mCol_lolim()    As Long
Dim mCol_hilim()    As Long
Dim mCol_scale      As Long
Dim mCol_units      As Long
Dim mCol_format     As Long
Dim mCol_passsort   As Long
Dim mCol_failsort   As Long
Dim mCol_passbin    As Long
Dim mCol_failbin    As Long
Dim mCol_result     As Long
Dim mCol_passaction As Long
Dim mCol_failaction As Long
Dim mCol_assume     As Long
Dim mCol_sites      As Long
Dim mCol_comment    As Long


Dim mS_Job()            As String
Dim mL_JobNum           As Long
Dim mRow_JobList        As Long
Dim mCol_JobListStart   As Long
Dim mRow_Header         As Long


Dim mCellName           As String
Dim mCheckString        As String
Dim mCheckFalg          As Boolean

Dim i           As Long
Dim Index       As Long

Dim mKey        As String
Dim mVal        As Variant
Dim mL_Row      As Long
Dim mL_Col      As Long
Dim mL_Range    As Long


Dim mCell As Object

Dim Allnumber_index As Double
Dim All_Sheet_arr() As String
Dim All_Sheet_number As String
Dim PorjectSheetName As String
All_Sheet_arr = Split(gC_ProjectLimitSheet, ",")
All_Sheet_number = UBound(All_Sheet_arr) + 1
Dim dd As Variant

        



    If CheckSheet(gC_Final_PM_Want_LimitSheet) = False Then
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = gC_Final_PM_Want_LimitSheet
        LL.LimitSheetHeader gS_JobList, gC_Final_PM_Want_LimitSheet
    End If
    
    Index = 4




    For Allnumber_index = 0 To All_Sheet_number - 1 Step 1
        PorjectSheetName = All_Sheet_arr(Allnumber_index)
        Set mSheet = Application.ActiveWorkbook.Sheets(PorjectSheetName)
        Worksheets(PorjectSheetName).Select
        For i = 4 To 65535 Step 1
            If UCase(mSheet.Cells(i, 5)) = "" Then Exit For
        Next i
        
'        Range("b" & Sheets(5).Range("b4").Value & ":d" & Sheets(5).Range("c4").Value & "").Select
        mSheet.Range(mSheet.Cells(4, 1), mSheet.Cells(i - 1, 8 + gL_JobNum * 2 + 13 + 1)).Select '"1:8+gL_JobNum*2 +13+1"
        Selection.Copy
        Worksheets(gC_Final_PM_Want_LimitSheet).Select
        Set mSheet = Application.ActiveWorkbook.Sheets(gC_Final_PM_Want_LimitSheet)
        Range(mSheet.Cells(Index, 1), mSheet.Cells(Index - 4 + i - 1, 8 + gL_JobNum * 2 + 13 + 1)).Select
        ActiveSheet.Paste
        
        Index = Index + i - 4
    Next Allnumber_index
    


    CreatMainSheet_form_Multisubsheet = True
    

Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function





'20190222
'___Highligh duplicated TName and TNum in LimitSheet in "Pink"
Public Function Dupl_TName_Check(mKey As String)
On Error GoTo ErrHandler
Dim funcName As String: funcName = "Dupl_TName_Check"
Dim sheet_index As Double
Dim All_Sheet_arr() As String
Dim All_Sheet_number As String
Dim PorjectSheetName As String
Dim Sht As Worksheet
Dim LookupTrue As Boolean
All_Sheet_arr = Split(gC_ProjectLimitSheet, ",")
All_Sheet_number = UBound(All_Sheet_arr) + 1
    
'___Lookup the duplicated item ("mkey") in each Project Limit sheet
    For sheet_index = 0 To All_Sheet_number - 1 Step 1
    LookupTrue = False
        PorjectSheetName = All_Sheet_arr(sheet_index)
        Set Sht = Application.ActiveWorkbook.Sheets(PorjectSheetName)
        Sht.Activate
        Call ShtLoc(PorjectSheetName, mKey, LookupTrue, "testname")
        
     Next sheet_index
     
     
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190222
'___Highligh duplicated TNum in LimitSheet in "Pink"
Public Function Dupl_TNum_Check(mKey_TNum As Double)
On Error GoTo ErrHandler
Dim funcName As String: funcName = "Dupl_TNum_Check"
Dim sheet_index As Double
Dim All_Sheet_arr() As String
Dim All_Sheet_number As String
Dim PorjectSheetName As String
Dim Sht As Worksheet
Dim LookupTrue As Boolean
All_Sheet_arr = Split(gC_ProjectLimitSheet, ",")
All_Sheet_number = UBound(All_Sheet_arr) + 1
    
'___Lookup the duplicated item ("mkey") in each Project Limit sheet
    For sheet_index = 0 To All_Sheet_number - 1 Step 1
    LookupTrue = False
        PorjectSheetName = All_Sheet_arr(sheet_index)
        Set Sht = Application.ActiveWorkbook.Sheets(PorjectSheetName)
        Sht.Activate
        Call ShtLoc(PorjectSheetName, CStr(mKey_TNum), LookupTrue, "testnumber")
        
     Next sheet_index
     
     
Exit Function
ErrHandler:
    RunTimeError funcName
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function ShtLoc(SheetName As String, Keyword As String, ByRef LookupTrue As Boolean, NameOrNum As String)
On Error GoTo ErrHandler
Dim funcName As String:: funcName = "ShtLoc"

Dim Col_Rng As Long
Dim end_row As Long
Dim Rng As Range
Dim i As Long

'___Locate the Start and End range
    '-------------------------------------------------------
        Worksheets(SheetName).Activate

     For Each Rng In Range("a1:z100")
         If Rng = NameOrNum Then
            Col_Rng = Rng.Column
            
            Cells(3, Col_Rng).Select
            Selection.End(xlDown).Select
            end_row = Selection.Row
          End If
      Next
            
            
      For i = 4 To end_row
         If UCase(Worksheets(SheetName).Cells(i, Col_Rng).Value) Like UCase(Keyword) Then
           
            LookupTrue = True
            
            '___Highlight in Pink
            Cells(i, Col_Rng).Select
            Selection.Interior.ColorIndex = 7
         Else
            LookupTrue = False
         End If
      Next i
         

   '-------------------------------------------------------
Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function





