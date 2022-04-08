Attribute VB_Name = "LIB_HardIP"
Option Explicit
Public RegDict As New Dictionary
Public gl_UseLimitCheck_Counter As Long
Public gl_CZ_FlowTestName_Counter As Long
Public ByPassTestLimit As Boolean
Public gl_GetInstrument_Dic As New Dictionary
Public gl_GetInstrumentType_Dic As New Dictionary
Public glb_Disable_CurrRangeSetting_Print As Boolean
'For ELB/ILB/TMPS HardIP Results update, 20190813
Public BV_Pass As New SiteBoolean

Public gl_FlowForLoop_DigSrc_SweepCode_Dec As String
Type String_Equation
    Name As String
    value_string As String
    Value As Long
End Type

Type DUTConditions
    PinName As String
    CurrentRange As Double
    FV_Val As Double
    FI_Val As Double
End Type

Dim G_MeasI_DIFF_1ST(50) As New PinListData

Public SrcBitBinaryString_MSB_LSB As String
Public SrcBitIndex As Long

Dim Min_Freq As New SiteDouble
Dim Max_freq As New SiteDouble

Dim MeasDelta_V1 As New PinListData
Dim MeasDelta_V2 As New PinListData
Dim MeasDelta_V As New PinListData
Public Delta_count As Integer
Public Delta_count1 As Integer
Public Count As Integer
Dim Arr(63) As New SiteDouble

Public LPDP_Sweep_EYE_POS(64, 15) As New SiteBoolean
Public LPDP_Sweep_EYE_NEG(64, 15) As New SiteBoolean
Public Counter_Y_Array As New SiteLong
'20170510 Eye Diagram Variable
Public Eye_Diagram_Binary(62) As New SiteVariant
'====================================================
'   Define the variables for TTR 20151225
'====================================================
Public Range_Check_Enable_Word As Boolean
Public CurrentJobName_L As String
Public CurrentJobName_U As String
Public EyeWf_sgmt0 As New DSPWave
Public EyeWf_sgmt1 As New DSPWave
Public DSSCSrcBitRecord As String '20180655 TER
        
''20160914 - Type for calculate equation pinlistdata of VFI
Type CALC_EQUATION_PLD
    FirstPinName As String
    FirstDictKey As String
    b_FirstConstant As Boolean
    SecondPinName As String
    SecondDictKey As String
    Operator As String
    b_SecondConstant As Boolean
End Type

'-------------------------Oscar 180523 For RegAssign Sheet Read
Type RegAssign
    RegName As String
    RegAssignByModeA As String
    RegAssignByModeB As String
End Type

Type ByTest
    testName As String
    RegAssign() As RegAssign
    RtnByModeA As String
    RtnByModeB As String
End Type

Type RegAssignInfo
    ByTest() As ByTest
End Type

Public RegAssignInfo As RegAssignInfo
'--------------------------Oscar

''20170809 - Cyprus Gross Fine Sweep
Public Gross_Counter As Integer
Public Fine_Counter As Integer

Public Instance_Data As Instance_Type
Public TestConditionSeqData() As MeasSeq_Type

Public Const SplitSeq_Pin = "+"
Public Const SplitSeq_WaitTime = "+" ''Carter, 20190604
Public Const SplitSeq_ForceVal = "|"
Public Const SplitRange = "+"
Public Const SplitSeqSweep = ":"
Public Const SplitJob = "=" ' ":" use for sweep
Public Const SplitPin = ","
Public Const SplitByPinForceVal = ","  ' merge
Public Const SplitMeasZ_ForceVal = "&"

Public Const DC_Spec_Var = "_VAR"

Type SaveCondition_Type
    Pin As String
    SourceFlodLimit As Double
    SinkFoldLimit As Double
    FilterValue As Double
    SrcCurrentRange As Double
    IfPowerPin As Boolean
    current As Double
End Type

Type TypePin_Type
    PPMU As String
    UVS256 As String
    HexVS As String
    VSM As String
    UVI80 As String
End Type

Type Meas_ByPin_Type
    Pin As String
    Pin_Diff_L As String
    Meas_Range As String
    ForceValue1 As String
    ForceValue2 As String
    UVI80_POWER_Flag As Boolean
End Type

Type Pin_Type
    PPMU As Meas_ByPin_Type
    UVS256 As Meas_ByPin_Type
    HexVS As Meas_ByPin_Type
    VSM As Meas_ByPin_Type
    UVI80 As Meas_ByPin_Type
End Type

Type ByType_Setup_Type
    PPMU_Flag As Boolean
    PPMU() As Meas_ByPin_Type
    UVI80_Flag As Boolean
    UVI80() As Meas_ByPin_Type
    UVS256_Flag As Boolean
    UVS256() As Meas_ByPin_Type
    HexVS_Flag As Boolean
    HexVS() As Meas_ByPin_Type
    VSM_Flag As Boolean
    VSM() As Meas_ByPin_Type
End Type

Type Meas_Type
    Pins As TypePin_Type
    WaitTime As TypePin_Type
    Setup_ByType As Pin_Type
    Setup_ByTypeByPin_Flag As Boolean
    Setup_ByTypeByPin As ByType_Setup_Type
    ForceValueDic As New Scripting.Dictionary
    ForceValueDic_HWCom As New Scripting.Dictionary ''Carter, 20190503
    MeasCurRangeDic As New Scripting.Dictionary
    SaveCondition() As SaveCondition_Type
    DiffMeter_Flag As Boolean
End Type

Type MeasF_Type
    Pins As String
    WaitTime As Double
    Interval As Double
    VT_Mode_Flag As Boolean
    Differential_Flag As Boolean
    MeasureThreshold_Flag As Boolean
    ThresholdPercentage As String
    EnableVtMode_Flag As Boolean
    WalkingStrobe_Flag As Boolean
    EventSource As FreqCtrEventSrcSel
End Type

Type MeasSeq_Type
    MeasCase As String   'MeasCase
    MeasV() As Meas_Type
    MeasVdiff() As Meas_Type
    MeasI() As Meas_Type
    measf() As MeasF_Type
    MeasFdiff() As MeasF_Type
    MeasR() As Meas_Type
    MeasZ() As Meas_Type
    Meas_StoreDicName() As String
End Type

Type Instance_Type
    patset As String
    TestSequence As String
    CPUA_Flag_In_Pat As Boolean
    DisableComparePins As String
    DisableConnectPins As String
    DisableFRC As Boolean
    FRCPortName As String
    
    MeasV_Pins As String
    MeasV_WaitTime_UVI80 As String

    MeasF_PinS_SingleEnd As String
    MeasF_Interval As String
    MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode
    MeasF_Flag_MeasureThreshold As Boolean
    MeasF_ThresholdPercentage As Double
    MeasF_WaitTime As String
    MeasF_EventSource As FreqCtrEventSrcSel ''Carter, 20190611
    MeasF_EnableVtMode_Flag As Boolean ''Carter, 20190611
    
    MeasI_pinS As String
    MeasI_Range As String
    MeasI_WaitTime As String
    
    DigCap_Pin As String
    DigCap_DataWidth As Long
    DigCap_Sample_Size As Long
    DigSrc_pin As String
    DigSrc_DataWidth As Long
    DigSrc_Sample_Size As Long
    DigSrc_Equation As String
    DigSrc_Assignment As String
    DigSrc_FlowForLoopIntegerName As String
    SpecialCalcValSetting As CalculateMethodSetup
    InstSpecialSetting As InstrumentSpecialSetup
    
    CUS_Str_MainProgram As String
    CUS_Str_DigCapData As String
    CUS_Str_DigSrcData As String
    
    Flag_SingleLimit As Boolean
    ForceFunctional_Flag As Boolean
    
    MeasF_PinS_Differential As String
    MeasF_WalkingStrobe_Flag As Boolean
    MeasF_WalkingStrobe_StartV As Double
    MeasF_WalkingStrobe_EndV As Double
    MeasF_WalkingStrobe_StepVoltage As Double
    MeasF_WalkingStrobe_BothVohVolDiffV As Double
    MeasF_WalkingStrobe_interval As Double
    MeasF_WalkingStrobe_miniFreq As Double
    Meas_StoreName As String
    
    Calc_Eqn As String
    Interpose_PrePat As String
    Interpose_PreMeas As String
    Interpose_PostTest As String
    
    CharSetName As String
    ForceV_Val As String
    ForceI_Val As String
    
    RAK_Flag As Enum_RAK
    WaitTime_VFIRZ As String
    Tname() As String
    
    LowLimit() As String
    HiLimit() As String
    
    Sweep_Info() As Power_Sweep
    Sweep_Enable As Boolean
    Sweep_CUS_Str_DigCapData As String
    Sweep_Calc_Eqn_Arg_Name As String
    
    TestSeqNum As Long
    TestSeqSweepNum As Long
    Is_PreCheck_Func As Boolean
    
    MergeDigSrcEquation As String
    DigSrcCheckCorrect As Boolean
    DigSrcEquationSampleSize() As Long
    StoreDictName As String
End Type
Enum Inst_Type
    PPMU = 1
    HexVS = 2
    UVI80 = 3
    UVS256 = 4
    VSM = 5 ''Carter, 20190412
End Enum


''Carter, 20190410
Public gCMError As New PinListData   ' save common mode error before DAC trim
Public DiffPinGroup As String

''20190107 - Global name for saving Customize Subblock name
Public gl_Current_Instance_Tname As String
Public gl_Current_Instance_Tname_subblock As String

Public Meas_StoreName_Flag As Boolean ''Carter, 20190521

Public gl_Sweep_vt As String
Public Function HardIP_waittime_trimming(MeasV As Meas_Type, MeasureVolt As PinListData)
    
    Dim i As Double
    Dim measureV As New PinListData
    Dim st_waittime As New SiteDouble: st_waittime = 0
    Dim temp_measureV As New PinListData: temp_measureV.AddPin (MeasV.Pins.UVI80)
    Dim site As Variant
    Dim st_boo As New SiteBoolean: st_boo = False
    Dim overhead_time As Double
    
    Dim start_time1 As Double, end_time1 As Double
    
    start_time1 = TheExec.Timer
        For Each site In TheExec.sites
            temp_measureV.Pins(MeasV.Pins.UVI80).Value = MeasureVolt.Pins(MeasV.Pins.UVI80).Value
        Next
    end_time1 = TheExec.Timer(start_time1)

    For i = 0 To 99 Step 1
        If st_boo.All(True) Then Exit For
        Dim start_time As Double
        Dim end_time As Double
        start_time = TheExec.Timer
        TheHdw.Wait 0.001
        measureV = TheHdw.DCVI.Pins(MeasV.Pins.UVI80).Meter.Read(tlStrobe, 2)
        TheExec.Datalog.WriteComment "=====================Wait time is " + CStr(1 + i) + "ms"
        TheExec.Flow.TestLimit measureV, , , , , scaleNone, unitVolt
        end_time = TheExec.Timer(start_time)
        overhead_time = overhead_time + end_time
        For Each site In TheExec.sites
            If temp_measureV.Pins(MeasV.Pins.UVI80).Subtract(measureV.Pins(MeasV.Pins.UVI80)).Abs.Divide(temp_measureV.Pins(MeasV.Pins.UVI80)).Value < 0.001 Then
                If st_boo = False Then
                    st_waittime(site) = overhead_time * 1000 + end_time1
                    st_boo = True
                End If
            End If
        Next site
        temp_measureV = measureV
    Next
    
    TheExec.Flow.TestLimit st_waittime

End Function

Public Function IO_Power_Split(TestPinArrayIV() As String, TestSeqNumIdx As Long, TempArr3() As String, TempStr As String) As Double
    If InStr(TestPinArrayIV(TestSeqNumIdx), ":") > 0 Then
        'Dim TempStr As String
        Dim TempArr1() As String
        Dim TempArr2() As String
        Dim TempArr4() As String
        Dim TempStrPin() As String
        Dim PinCount As Long
        Dim index As Variant
        Dim j As Integer
        ReDim TempArr3(0) As String
        TempStr = TestPinArrayIV(TestSeqNumIdx)
        TempArr2 = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        TestPinArrayIV(TestSeqNumIdx) = ""
        For Each index In TempArr2
            TempStrPin = Split(index, ":")
            If TestPinArrayIV(TestSeqNumIdx) <> "" Then
                TestPinArrayIV(TestSeqNumIdx) = TestPinArrayIV(TestSeqNumIdx) + "," + TempStrPin(0)
            Else
                TestPinArrayIV(TestSeqNumIdx) = TempStrPin(0)
            End If
            Call TheExec.DataManager.DecomposePinList(TempStrPin(0), TempArr4, PinCount)
            For j = 0 To UBound(TempArr4)
                TempArr3(UBound(TempArr3)) = TempArr4(j) + ":" + TempStrPin(1)
                ReDim Preserve TempArr3(UBound(TempArr3) + 1)
            Next j
        Next index
    End If
End Function

Public Function PATT_ExculdePath(Pat As Variant) As String
Dim patt_ary_temp() As String
    patt_ary_temp = Split(Pat, "\")
    PATT_ExculdePath = patt_ary_temp(UBound(patt_ary_temp))

End Function

'===========================================================================================
' Check if pattern name provided is a pattern set
' NOTE: If pattern set is true and count > 1 then the elements returned may still be
'          nested pattern sets.  The calling function should recursively call this to
'          ensure that the returned Names resolve to individual patterns
'===========================================================================================
Public Function PATT_GetPatListFromPatternSet(TestPat As String, _
                              rtnPatNames() As String, _
                              rtnPatCnt As Long) As Boolean

    Dim PatCnt As Long                          '<- Number of patterns in set
    Dim RawNameData() As String                 '<- Raw pattern name data
    Dim rtnPatNames1() As String
    Dim rtnPatNames2() As String
    Dim i As Long, j As Long
    '___ Init _____________________________________________________________________________
    On Error GoTo errHandler
    
    '___ Check the name ___________________________________________________________________
    '    Individual pattern name or non-pattern string returns an error - thus false
    '--------------------------------------------------------------------------------------
    rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(TestPat, PatCnt)
    If (UBound(rtnPatNames) > 0) Then
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            PATT_GetPatListFromPatternSet = True
            rtnPatCnt = UBound(rtnPatNames) + 1
        Else
            rtnPatCnt = 0
            For i = 0 To UBound(rtnPatNames)
                rtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames(i), PatCnt)
                rtnPatCnt = rtnPatCnt + UBound(rtnPatNames2) + 1
            Next i
            rtnPatNames1 = TheExec.DataManager.Raw.GetPatternsInSet(TestPat, PatCnt)
            ReDim rtnPatNames(rtnPatCnt - 1)    ' modify 827 j
            rtnPatCnt = 0
            For i = 0 To UBound(rtnPatNames1)
                rtnPatNames2 = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames1(i), PatCnt)
                For j = 0 To UBound(rtnPatNames2)
                    If LCase(rtnPatNames2(j)) Like "*.pat*" Then
                        rtnPatNames(rtnPatCnt) = rtnPatNames2(j)
                    Else
                        TheExec.ErrorLogMessage TestPat & " in more than 2 level of pattern set"
                    End If
                    rtnPatCnt = rtnPatCnt + 1
                Next j
            Next i
            PATT_GetPatListFromPatternSet = True
        End If
    Else
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            PATT_GetPatListFromPatternSet = True
            rtnPatCnt = 1
        Else
            rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(rtnPatNames(0), PatCnt)
            rtnPatCnt = UBound(rtnPatNames) + 1
            For j = 0 To UBound(rtnPatNames)
                If LCase(rtnPatNames(j)) Like "*.pat*" Then
                Else
                    TheExec.ErrorLogMessage TestPat & " in more than 2 level of pattern set"
                End If
            Next j
        End If
    End If
    
    Exit Function
    
errHandler:
    PATT_GetPatListFromPatternSet = False
    rtnPatCnt = -1
    Exit Function

End Function

Public Function Decide_Measure_Pin(TestSeqNum As Integer, MeasPinAry() As String, ByRef Measure_Pin As PinList, Optional k As Long)
    Dim TestSeqNumIdx As Integer
    TestSeqNumIdx = TestSeqNum
    
    If TestSeqNum > 0 Then
        If UBound(MeasPinAry) = 0 Then TestSeqNumIdx = 0
    End If
    If UBound(MeasPinAry) >= 0 Then
''        If InStr(LCase(MeasPinAry(TestSeqNumIdx)), "idx") <> 0 Then TestSeqNumIdx = Int(Mid(MeasPinAry(TestSeqNumIdx), 4, 1))
        Measure_Pin = MeasPinAry(TestSeqNumIdx)
    End If
End Function

Public Function Decide_MeasureI_CurrentRange(TestSeqNum As Integer, MeasPinAry_IRange() As String, ByRef MeasureI_CurrentRange As String, Optional k As Long)
    Dim TestSeqNumIdx As Integer
    TestSeqNumIdx = TestSeqNum
    
    If TestSeqNum > 0 Then
        If UBound(MeasPinAry_IRange) = 0 Then TestSeqNumIdx = 0
    End If
    If UBound(MeasPinAry_IRange) >= 0 Then
        '' 20150605 - Check with CC
''        If InStr(LCase(MeasPinAry_IRange(TestSeqNumIdx)), "idx") <> 0 Then TestSeqNumIdx = Int(Mid(MeasPinAry_IRange(TestSeqNumIdx), 4, 1))
        If InStr(MeasPinAry_IRange(TestSeqNumIdx), ":") <> 0 Then
            If (UBound(Split(MeasPinAry_IRange(TestSeqNumIdx), ":")) >= (k - 1)) Then
                MeasureI_CurrentRange = Split(MeasPinAry_IRange(TestSeqNumIdx), ":")(k - 1)
            Else
                MeasureI_CurrentRange = Split(MeasPinAry_IRange(TestSeqNumIdx), ":")(0)
            End If
        Else
        MeasureI_CurrentRange = MeasPinAry_IRange(TestSeqNumIdx)
    End If
        
    End If
End Function

Public Function HardIP_WriteFuncResult(Optional SpecialReserve As String = "", Optional CodeSearchPatternResult As SiteBoolean, Optional m_testName As String = "") As Long
    Dim site As Variant
    Dim TestNumber As Long
    Dim FailCount As New PinListData
    Dim allpins As PinList
    Dim Pin As Variant
    Dim Pins() As String
    Dim Pin_Cnt As Long
    
    '' 20150604: Need to modify "All_Digital" to the parameter.
    TheExec.DataManager.DecomposePinList "All_Digital", Pins(), Pin_Cnt
    
    If SpecialReserve <> "" Then
        If SpecialReserve = "DSSC_CODESEARCH" Then
            For Each site In TheExec.sites
                TestNumber = TheExec.sites.Item(site).TestNumber
                If CodeSearchPatternResult(site) Then
                    If TheExec.DevChar.Setups.IsRunning = True Then TheExec.sites.Item(site).testResult = sitePass

                    ''''20151106 update
                    If (m_testName <> "") Then
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass, , m_testName)
                    Else
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                    End If
                Else
                    ''''20151106 update
                    If (m_testName <> "") Then
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail, , m_testName)
                    Else
                        Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                    End If
                    '' 20160218 - Modify sequence to let TestResult after WriteFunctionalResult to cover test number increment 2 issue if souce sink time out alarm happen.
                    TheExec.sites.Item(site).testResult = siteFail
                End If
                TheExec.sites.Item(site).TestNumber = TestNumber + 1
            Next site
        End If
    Else

        Dim patPassed As New SiteBoolean
        patPassed = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        For Each site In TheExec.sites
            TestNumber = TheExec.sites.Item(site).TestNumber
            Exit For
        Next site
        
        For Each site In TheExec.sites
            If patPassed Then
                TheExec.sites.Item(site).testResult = sitePass
                If TheExec.DevChar.Setups.IsRunning = True Then TheExec.sites.Item(site).testResult = sitePass
                ''''20151106 update
                If (m_testName <> "") Then
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass, , m_testName)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
                End If
            Else
                ''''20151106 update
                If (m_testName <> "") Then
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail, , m_testName)
                Else
                    Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
                End If
                '' 20160218 - Modify sequence to let TestResult after WriteFunctionalResult to cover test number increment 2 issue if souce sink time out alarm happen.
                TheExec.sites.Item(site).testResult = siteFail
                
            End If
            
            If TheExec.DevChar.Setups.IsRunning = False Then TheExec.sites.Item(site).TestNumber = TestNumber + 1
            
            '20180607 TER **************************************************************************************
            If TheExec.DevChar.Setups.IsRunning = True Then
                Dim SetupName As String
                
                SetupName = TheExec.DevChar.Setups.ActiveSetupName
                If Not ((TheExec.DevChar.Results(SetupName).startTime Like "1/1/0001*" Or TheExec.DevChar.Results(SetupName).startTime Like "0001/1/1*")) Then
                    With TheExec.DevChar.Setups(SetupName)
                        If .Shmoo.Axes.Count > 1 Then
                                XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                                YVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
                                If gl_flag_end_shmoo = True Then
                                    TheExec.sites.Item(site).TestNumber = TestNumber + 1
                                End If
                        Else
                                XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                                If gl_flag_end_shmoo = True Then
                                    TheExec.sites.Item(site).TestNumber = TestNumber + 1
                                End If
                        End If
                    End With
                End If
            End If
            '*****************************************************************************************************
        Next site
    End If
    Call Update_BC_PassFail_Flag(True)
End Function


Public Function GetFlowSingleUseLimit(ByRef d_HighLimitVal() As Double, ByRef d_LowLimitVal() As Double) As Double
    ' Get the limits info
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Dim HighLimitValArray() As String
    Dim LowLimitValArray() As String
    Dim HighLimitArraySize As Long
    Dim LowLimitArraySize As Long
    Dim i As Integer
    
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    If FlowLimitsInfo Is Nothing Then
        ReDim d_HighLimitVal(0) As Double
        ReDim d_LowLimitVal(0) As Double
        ReDim HighLimitValArray(0) As String
        ReDim LowLimitValArray(0) As String
        d_HighLimitVal(0) = 0
        d_LowLimitVal(0) = 0
        HighLimitValArray(0) = 0
        LowLimitValArray(0) = 0
    Else
        Call FlowLimitsInfo.GetHighLimits(HighLimitValArray())
        Call FlowLimitsInfo.GetLowLimits(LowLimitValArray())
        HighLimitArraySize = UBound(HighLimitValArray)
        ReDim d_HighLimitVal(HighLimitArraySize) As Double
        LowLimitArraySize = UBound(LowLimitValArray)
        ReDim d_LowLimitVal(LowLimitArraySize) As Double
    End If
    For i = 0 To HighLimitArraySize
        If (HighLimitValArray(i)) = "" Then HighLimitValArray(i) = 0
        d_HighLimitVal(i) = CDbl(HighLimitValArray(i))
    Next i
    For i = 0 To LowLimitArraySize
        If LowLimitValArray(i) = "" Then LowLimitValArray(i) = 0
        d_LowLimitVal(i) = CDbl(LowLimitValArray(i))
    Next i
End Function
Public Function Find_Assignement(name_str As String, str_eq_ary() As String_Equation, Optional DigSrcPrint As String = "", Optional NumberPins As Long = 1, Optional MSB_First_Flag As Boolean) As String
    Dim i As Long
    Dim printstring As String
    For i = 0 To UBound(str_eq_ary)
        If name_str Like Trim(str_eq_ary(i).Name) Then
            Find_Assignement = str_eq_ary(i).value_string
            printstring = str_eq_ary(i).value_string '=======180410 Added by Oscar
            If NumberPins > 1 Then printstring = CStr(Dec2BinStr32Bit(NumberPins, CLng(printstring))) '=======180410 Added by Oscar
            If MSB_First_Flag Then '''MSB first
                printstring = StrReverse(printstring)
                Find_Assignement = StrReverse(Find_Assignement)
            End If
            DigSrcPrint = printstring & "(" & name_str & ")" '=======180410 Added by Oscar
            Exit Function
        End If
    Next i
End Function
Public Function HardIP_MeasureVolt() As PinListData

    Dim site As Variant
    Dim p As Long
    Dim TempMeasVal_PerPin(100) As New PinListData
    Dim MeasureV_Pin_IO As String
    Dim MeasureV_Pin_UVI80 As String
    Dim MeasV_INstType_Num As Integer
    Dim MeasureV_Average As New PinListData
    Dim OutputTname_format() As String
    Dim DUT_TestConditions() As DUTConditions
    
    Dim MeasureV_pin As PinList
    Dim TestLimitByPin_VFI As String
    Dim TestSeqNum As Integer
    Dim k As Long
    Dim Pat As Variant
    Dim Flag_SingleLimit As Boolean
    Dim HighLimitVal As Double
    Dim LowLimitVal As Double
    Dim InstSpecialSetting As InstrumentSpecialSetup
    Dim CUS_Str_MainProgram As String
    Dim SpecialCalcValSetting As CalculateMethodSetup
    Dim Rtn_MeasVolt As PinListData
    Dim Rtn_SweepTestName As String
    Dim MeasurePin_ForceI_Val As String
    Dim UVI80_MeasV_WaitTime As String
    Dim OutputTname As String
    Dim MeasDiff_V As New PinListData
    Dim MeasVOD_V As New PinListData
    Dim MeasVOCM_V As New PinListData
    Dim TestNameInput As String
    Dim MipiResult As New PinListData
    
    Dim Temp_index As Long
    
    Dim MeasV As Meas_Type
    Dim MeasVoltage(0 To 1) As New PinListData
    Dim DicStoreName As String
    
    MeasV = TestConditionSeqData(Instance_Data.TestSeqNum).MeasV(Instance_Data.TestSeqSweepNum)
    If Meas_StoreName_Flag Then DicStoreName = TestConditionSeqData(Instance_Data.TestSeqNum).Meas_StoreDicName(Instance_Data.TestSeqSweepNum)
    
    Dim index As Integer
    index = 0
    If (MeasV.Pins.UVI80 <> "") Then
        If Not TestConditionSeqData(Instance_Data.TestSeqNum).MeasV(Instance_Data.TestSeqSweepNum).DiffMeter_Flag Then
            Call HardIP_SetupAndMeasureVolt_UVI80(MeasVoltage(index))
        Else
            Call HardIP_SetupAndMeasureVolt_UVI80_Diff(MeasVoltage(index))
        End If
        index = index + 1
    End If
    If (MeasV.Pins.PPMU <> "") Then
        If Instance_Data.InstSpecialSetting = PPMU_SerialMeasurement Then
            Call HardIP_SetupAndMeasureVolt_PPMU_BySerial(MeasVoltage(index))
        Else
            Call HardIP_SetupAndMeasureVolt_PPMU(MeasVoltage(index))
        End If
        index = index + 1
    End If
    ''''---------Offline---------
    If TheExec.TesterMode = testModeOffline Then
        Dim Pin As Variant
        Dim Dummy_Value As New SiteDouble: Dummy_Value = 0.000000001
        For Each Pin In MeasVoltage(index - 1).Pins
            MeasVoltage(index - 1).Pins(Pin) = Dummy_Value
        Next Pin
    End If
    ''''---------Offline---------
    Dim MeasV_ToTestLimit As New PinListData
    If index = 1 Then
        Set MeasV_ToTestLimit = MeasVoltage(0)
    Else
        Call MergePinListData(index, MeasVoltage, MeasV_ToTestLimit)
    End If
    If Instance_Data.SpecialCalcValSetting = Average_voltage Then
        MeasureV_Average.AddPin (MeasV_ToTestLimit.Pins(0).Name)
        MeasureV_Average.Pins(MeasV_ToTestLimit.Pins(0).Name) = MeasV_ToTestLimit.Analyze.mean
    End If

    If Instance_Data.SpecialCalcValSetting = DIFF_PN Then
        For p = 0 To MeasV_ToTestLimit.Pins.Count - 1 Step 2
            MeasDiff_V.AddPin (MeasV_ToTestLimit.Pins(p + 1).Name)
            For Each site In TheExec.sites.Active
                MeasDiff_V.Pins(MeasV_ToTestLimit.Pins(p + 1).Name).Value = Abs(MeasV_ToTestLimit.Pins(p).Value - MeasV_ToTestLimit.Pins(p + 1).Value)
            Next site
        Next p
    ElseIf Instance_Data.SpecialCalcValSetting = CalculateMethodSetup.VIR_VOD_VOCM_PN Then 'add by JiYi 20160721 for refbuf
        For p = 0 To MeasV_ToTestLimit.Pins.Count - 1 Step 2
            MeasVOD_V.AddPin (MeasV_ToTestLimit.Pins(p + 1).Name)
            MeasVOCM_V.AddPin (MeasV_ToTestLimit.Pins(p + 1).Name)
            
            For Each site In TheExec.sites.Active
                MeasVOD_V.Pins(MeasV_ToTestLimit.Pins(p + 1).Name).Value = MeasV_ToTestLimit.Pins(p + 1).Value - MeasV_ToTestLimit.Pins(p).Value
                MeasVOCM_V.Pins(MeasV_ToTestLimit.Pins(p + 1).Name).Value = 0.5 * (MeasV_ToTestLimit.Pins(p + 1).Value + MeasV_ToTestLimit.Pins(p).Value)
            Next site
        Next p
    End If
    
    If Not ByPassTestLimit Then
        If Instance_Data.SpecialCalcValSetting = DIFF_DCO Then
            TestNameInput = Report_TName_From_Instance("V", MeasV_ToTestLimit.Pins(p), "Src" & TheExec.Flow.var("SrcCodeIndex").Value, CInt(Instance_Data.TestSeqNum), 0)
            TheExec.Flow.TestLimit MeasV_ToTestLimit.Pins(p), , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
        ElseIf Instance_Data.SpecialCalcValSetting = CalculateMethodSetup.VIR_VOD_VOCM_PN Then 'add by JiYi 20160721 for refbuf
            For p = 0 To MeasV_ToTestLimit.Pins.Count - 1
                If Instance_Data.CUS_Str_MainProgram = "TTR" Then
                    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
                Else
                TestNameInput = Report_TName_From_Instance("V", MeasV_ToTestLimit.Pins(p), "Pin", CInt(Instance_Data.TestSeqNum), p)
                TheExec.Flow.TestLimit MeasV_ToTestLimit.Pins(p), , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
                End If
            Next p
            
            For p = 1 To (MeasV_ToTestLimit.Pins.Count) - 1 Step 2
                TestNameInput = Report_TName_From_Instance("V", MeasV_ToTestLimit.Pins(p), "vdiff", CInt(Instance_Data.TestSeqNum), p)
                TheExec.Flow.TestLimit MeasVOD_V.Pins((p - 1) / 2), , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next p
            For p = 1 To (MeasV_ToTestLimit.Pins.Count) - 1 Step 2
                TestNameInput = Report_TName_From_Instance("V", MeasV_ToTestLimit.Pins(p), "vocm", CInt(Instance_Data.TestSeqNum), p)
                TheExec.Flow.TestLimit MeasVOCM_V.Pins((p - 1) / 2), , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next p
        ElseIf Instance_Data.SpecialCalcValSetting = CalculateMethodSetup.VIR_DDIO Then
            'not printing anymore
        Else
            Call ProsscessTestLimit(MeasV_ToTestLimit, "V", CInt(Instance_Data.TestSeqNum))
        End If
        If Instance_Data.SpecialCalcValSetting = DIFF_PN Then
            Call ProsscessTestLimit(MeasDiff_V, "V", CInt(Instance_Data.TestSeqNum), "Vdiff")
        End If
        If Instance_Data.SpecialCalcValSetting = Average_voltage Then
            TestNameInput = Report_TName_From_Instance("V", MeasureV_Average.Pins(0), "AvgV", CInt(Instance_Data.TestSeqNum), 0)
            TheExec.Flow.TestLimit MeasureV_Average, , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If
    End If
       
    ''Start ---- Carter, 20190521
    If Meas_StoreName_Flag Then
        If DicStoreName <> "" Then Call AddStoredMeasurement(DicStoreName, MeasV_ToTestLimit)
    End If
    ''End ---- Carter, 20190521
    Set HardIP_MeasureVolt = MeasV_ToTestLimit
    
End Function

Public Function DSSCSrcBitFromFlowForLoop(FlowForLoopIntegerName As String, DigSrc_DataWidth As Long, DigSrc_Equation As String, ByRef DigSrc_Assignment As String, _
    Optional CUS_Str_DigSrcData As String, Optional ByRef Rtn_SweepTestName As String)
    
    Dim SrcBitIndex As Long
    Dim i As Long, j As Long, k As Long, z As Long
    
    Dim ForLoopIntegerNameSplitBySemi() As String
    Dim ForLoopIntegerNameSplitByLine() As String
    Dim ForLoopIntegerNameSplitByColon() As String
    Dim BinaryBitWidth As Long
    Dim DigSrc_Assignment_Update As String
    
    Dim SplitBySemi() As String
    Dim SplitByEqual() As String
    Dim ProcessCopyString As String
    Dim ProcessCopyStartPoint As Long
    Dim splitbyand() As String
    Dim b_BinaryFormat As Boolean
    Dim ProcessBinString As String
    
    Dim inputDecimal As Long
    Dim m_binarr() As Long
    Dim Temp_SweepTestName
    Dim SweepCodeLength As Long
    Dim b_OnlySrcBitIndexStr As Boolean
    Dim ForRepeat_SrcBitBinaryString As String
    
    
    'ex: FlowForLoopIntegerName=SrcCodeIndx;sdll_sel_dck:7:0:1|SrcCodeIndx1;sdll_sel_ck:10:800:1
                                     ' Outer for loop              Inter for loop
    ForLoopIntegerNameSplitByLine = Split(FlowForLoopIntegerName, "|")  ' Add for more than one loop sweep
    
    Dim Temp_SweepTestName_byline() As String
    ReDim Temp_SweepTestName_byline(UBound(ForLoopIntegerNameSplitByLine)) As String
    
    
    For z = 0 To UBound(ForLoopIntegerNameSplitByLine)
    ForLoopIntegerNameSplitBySemi = Split(ForLoopIntegerNameSplitByLine(z), ";")
    SrcBitIndex = TheExec.Flow.var(ForLoopIntegerNameSplitBySemi(0)).Value
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("SrcBitIndex" & z & "value = " & SrcBitIndex)
    
    If UBound(ForLoopIntegerNameSplitBySemi) = 0 Then
        b_OnlySrcBitIndexStr = True
    Else
        b_OnlySrcBitIndexStr = False
    End If
    If b_OnlySrcBitIndexStr Then
        SplitBySemi = Split(DigSrc_Assignment, ";")
        If InStr(LCase(SplitBySemi(0)), "copy") <> 0 Then
            SweepCodeLength = DigSrc_DataWidth / CDbl(Right(SplitBySemi(0), 1))
        Else
            SweepCodeLength = DigSrc_DataWidth
        End If
        ReDim m_binarr(SweepCodeLength - 1) As Long
        inputDecimal = SrcBitIndex
        Call Dec2Bin(Abs(inputDecimal), m_binarr)
        
        '' GrayCode trasfer
        ''===============================================================
        If UCase(CUS_Str_DigSrcData) = UCase("BinToGray") Then
            ReDim m_binarr_graycode(SweepCodeLength - 1) As Long
            For j = 0 To UBound(m_binarr) Step 1
                If j = 0 Then
                    m_binarr_graycode(j) = m_binarr(j)
                Else
                    If m_binarr(j - 1) = m_binarr(j) Then
                        m_binarr_graycode(j) = 0
                    Else
                        m_binarr_graycode(j) = 1
                    End If
                End If
            Next j
            
            For j = 0 To UBound(m_binarr) Step 1
                m_binarr(j) = m_binarr_graycode(j)
            Next j
        End If
        ''===============================================================
        '' 20170418 - SignGrayCode trasfer
        If UCase(CUS_Str_DigSrcData) = UCase("BinToGray_Sign") Then
            ReDim m_binarr_graycode(SweepCodeLength - 1) As Long
            If SrcBitIndex < 0 Then
                m_binarr_graycode(0) = 1
            Else
                m_binarr_graycode(0) = 0
            End If
            Call Dec2Bin(Abs(SrcBitIndex), m_binarr)
            For j = 1 To UBound(m_binarr) Step 1
                If j = 0 Then
                    m_binarr_graycode(j) = m_binarr(j)
                Else
                    If m_binarr(j - 1) = m_binarr(j) Then
                        m_binarr_graycode(j) = 0
                    Else
                        m_binarr_graycode(j) = 1
                    End If
                End If
            Next j
            For j = 0 To UBound(m_binarr) Step 1
                 m_binarr(j) = m_binarr_graycode(j)
            Next j
        End If
        ''===============================================================
        
        '' 20150811 - Reverse order of binary string
        For j = UBound(m_binarr) To 0 Step -1
            If j = UBound(m_binarr) Then
                ForRepeat_SrcBitBinaryString = m_binarr(j)
            Else
                ForRepeat_SrcBitBinaryString = ForRepeat_SrcBitBinaryString & m_binarr(j)
            End If
        Next j
        
        SplitByEqual = Split(SplitBySemi(0), "=")
        If InStr(LCase(SplitByEqual(1)), "copy") <> 0 Then
            ProcessCopyStartPoint = InStr(LCase(SplitByEqual(1)), "copy")
            ProcessCopyString = ForRepeat_SrcBitBinaryString & Right(SplitByEqual(1), Len(SplitByEqual(1)) - ProcessCopyStartPoint + 2)
            DigSrc_Assignment_Update = ProcessCopyString
        Else
            DigSrc_Assignment_Update = ForRepeat_SrcBitBinaryString
        End If
        DigSrc_Assignment = "repeat=" & DigSrc_Assignment_Update
        Temp_SweepTestName = ForRepeat_SrcBitBinaryString
    Else
    
        Dim RegNum As Long
        RegNum = UBound(ForLoopIntegerNameSplitBySemi) - 1
        
        ReDim RegReplaceName(RegNum) As String
        ReDim RegBinBitWidth(RegNum) As Long
        ReDim RegStartPoint(RegNum) As Long
        ReDim RegStepSize(RegNum) As Long
        ReDim SrcBitBinaryString(RegNum) As String
        
        For i = 0 To RegNum
            ForLoopIntegerNameSplitByColon = Split(ForLoopIntegerNameSplitBySemi(i + 1), ":")
            ''20161025-Start point from 0 and step 1, if input not specify
            If UBound(ForLoopIntegerNameSplitByColon) = 1 Then
                RegReplaceName(i) = ForLoopIntegerNameSplitByColon(0)
                RegBinBitWidth(i) = ForLoopIntegerNameSplitByColon(1)
                RegStartPoint(i) = 0
                RegStepSize(i) = 1
            Else
            
                RegReplaceName(i) = ForLoopIntegerNameSplitByColon(0)
                RegBinBitWidth(i) = ForLoopIntegerNameSplitByColon(1)
                RegStartPoint(i) = ForLoopIntegerNameSplitByColon(2)
                RegStepSize(i) = ForLoopIntegerNameSplitByColon(3)
            End If
        Next i
    
        
    
        For i = 0 To RegNum
            ReDim m_binarr(RegBinBitWidth(i) - 1) As Long
    
            inputDecimal = RegStartPoint(i) + (SrcBitIndex * RegStepSize(i))
            
            If z = 0 Then
                gl_FlowForLoop_DigSrc_SweepCode_Dec = inputDecimal '20190613 CT add for Decimal value printing
            Else
                gl_FlowForLoop_DigSrc_SweepCode_Dec = gl_FlowForLoop_DigSrc_SweepCode_Dec & "&" & inputDecimal '20190613 CT add for Decimal value printing
            End If
            
            Call Dec2Bin(Abs(inputDecimal), m_binarr)
             Temp_SweepTestName_byline(z) = inputDecimal  ' For Two "For loop" Value
                        '' Cyprus Gross Fine Sweep 20170809
            ''===============================================================
            If UCase(CUS_Str_DigSrcData) Like UCase("*GrossFineSweep*") Then
                'Dim Fine_Counter As Integer
                Dim Gross_Num As Integer
                Dim Fine_Num As Integer
                Dim m_binarr_gross() As Long
                Dim m_binarr_fine() As Long
                Dim GrossFineSplitByComma() As String
                GrossFineSplitByComma = Split(CUS_Str_DigSrcData, ",")
                Gross_Num = GrossFineSplitByComma(1)
                Fine_Num = GrossFineSplitByComma(2)
                
                ''-simulation-
'                Gross_Num = 4
'                Fine_Num = 8
                ''------------
                
                ReDim m_binarr_gross(Gross_Num - 1) As Long ''4bit
                ReDim m_binarr_fine(Fine_Num - 1) As Long ''8bit
                
                ''Counter reset
                If inputDecimal = 0 Then
                    Fine_Counter = 0
                    Gross_Counter = 0
                End If
                
                
                
                If Fine_Counter = 0 Then    ''8'b00000000
                    Fine_Counter = Fine_Counter + 1
                    
                    Call Dec2Bin(Abs(Gross_Counter), m_binarr_gross)

                    For j = 0 To Fine_Num - 1 Step 1
                        m_binarr_fine(j) = 0
                    Next j
                    
                    '' Combine
                    For j = 0 To UBound(m_binarr) Step 1
                        If j < Gross_Num Then
                            m_binarr(j) = m_binarr_gross(Gross_Num - 1 - j)
                        Else
                            m_binarr(j) = m_binarr_fine(j - Gross_Num)
                        End If
                    Next j
                    
                ElseIf Fine_Counter = 1 Then    ''8'b10000000
                    Fine_Counter = Fine_Counter + 1
                    Call Dec2Bin(Abs(Gross_Counter), m_binarr_gross)
                    For j = 0 To Fine_Num - 1 Step 1
                        m_binarr_fine(j) = 0
                    Next j
                    m_binarr_fine(0) = 1
                    
                    '' Combine
                    For j = 0 To UBound(m_binarr) Step 1
                        If j < Gross_Num Then
                            m_binarr(j) = m_binarr_gross(Gross_Num - 1 - j)
                        Else
                            m_binarr(j) = m_binarr_fine(j - Gross_Num)
                        End If
                    Next j
                    
                ElseIf Fine_Counter = 2 Then    ''8'b11111111
                    Fine_Counter = 0
                    Call Dec2Bin(Abs(Gross_Counter), m_binarr_gross)
                    Gross_Counter = Gross_Counter + 1
                    For j = 0 To Fine_Num - 1 Step 1
                        m_binarr_fine(j) = 1
                    Next j
                    
                    '' Combine
                    For j = 0 To UBound(m_binarr) Step 1
                        If j < Gross_Num Then
                            m_binarr(j) = m_binarr_gross(Gross_Num - 1 - j)
                        Else
                            m_binarr(j) = m_binarr_fine(j - Gross_Num)
                        End If
                    Next j

                End If
            End If
            ''===============================================================
            
            
            '' GrayCode trasfer
            ''===============================================================
            If UCase(CUS_Str_DigSrcData) = UCase("BinToGray") Then
                ReDim m_binarr_graycode(RegBinBitWidth(i) - 1) As Long
                For j = 0 To UBound(m_binarr) Step 1
                    If j = 0 Then
                        m_binarr_graycode(j) = m_binarr(j)
                    Else
                        If m_binarr(j - 1) = m_binarr(j) Then
                            m_binarr_graycode(j) = 0
                        Else
                            m_binarr_graycode(j) = 1
                        End If
                    End If
                Next j
                
                For j = 0 To UBound(m_binarr) Step 1
                    m_binarr(j) = m_binarr_graycode(j)
                Next j
            End If
            ''===============================================================

            '' 20170418 - SignGrayCode trasfer
            If UCase(CUS_Str_DigSrcData) = UCase("BinToGray_Sign") Then
                ReDim m_binarr_graycode(RegBinBitWidth(i) - 1) As Long
                If SrcBitIndex < 0 Then
                    m_binarr_graycode(0) = 1
                Else
                    m_binarr_graycode(0) = 0
                End If
                Call Dec2Bin(Abs(SrcBitIndex), m_binarr)
                For j = 1 To UBound(m_binarr) Step 1
                    If j = 0 Then
                        m_binarr_graycode(j) = m_binarr(j)
                    Else
                        If m_binarr(j - 1) = m_binarr(j) Then
                            m_binarr_graycode(j) = 0
                        Else
                            m_binarr_graycode(j) = 1
                        End If
                    End If
                Next j
                For j = 0 To UBound(m_binarr) Step 1
                     m_binarr(j) = m_binarr_graycode(j)
                Next j
            End If
            ''===============================================================
            
            
            '' 20150811 - Reverse order of binary string
            For j = UBound(m_binarr) To 0 Step -1
                If j = UBound(m_binarr) Then
                    SrcBitBinaryString(i) = m_binarr(j)
                Else
                    SrcBitBinaryString(i) = SrcBitBinaryString(i) & m_binarr(j)
                End If
            Next j
        Next i
        
        '' 20160829 - Replace reg content to sweep code

        SplitBySemi = Split(DigSrc_Assignment, ";")
        DigSrc_Assignment_Update = ""
        ProcessBinString = ""
        For i = 0 To UBound(SplitBySemi)
            SplitByEqual = Split(SplitBySemi(i), "=")
    
                For j = 0 To RegNum
                    If InStr(LCase(SplitByEqual(1)), "copy") <> 0 Then
                        ProcessCopyStartPoint = InStr(LCase(SplitByEqual(1)), "copy")
                        
                        If SplitByEqual(0) = RegReplaceName(j) Then
                            ProcessCopyString = SrcBitBinaryString(j) & Right(SplitByEqual(1), Len(SplitByEqual(1)) - ProcessCopyStartPoint + 2)
                            SplitByEqual(1) = ProcessCopyString
                        End If
                    ''20160912-Check & format at DigSrc_Assignment, ==> 110&CAL_A&011
                    ElseIf InStr(LCase(SplitByEqual(1)), "&") <> 0 Then
                        splitbyand = Split(SplitByEqual(1), "&")
                        For k = 0 To UBound(splitbyand)
                            b_BinaryFormat = Checker_ConstantBinary(splitbyand(k))
                            If k = 0 Then
                                If b_BinaryFormat Then
                                    ProcessBinString = splitbyand(k)
                                Else
    ''                                SplitByColon = Split(SplitByAnd(k), ":")
                                    If SplitByEqual(0) = RegReplaceName(j) Then
                                        ProcessBinString = SrcBitBinaryString(j)
                                    Else
                                        ProcessBinString = splitbyand(k)
                                    End If
                                End If
                            Else
                                If b_BinaryFormat Then
                                    ProcessBinString = ProcessBinString & "&" & splitbyand(k)
                                Else
    ''                                SplitByColon = Split(SplitByAnd(k), ":")
                                    If SplitByEqual(0) = RegReplaceName(j) Then
                                        ProcessBinString = ProcessBinString & "&" & SrcBitBinaryString(j)
                                    Else
                                        ProcessBinString = ProcessBinString & "&" & splitbyand(k)
                                    End If
                                End If
                            End If
                        Next k
                        SplitByEqual(1) = ProcessBinString
                    Else

                        If SplitByEqual(0) = RegReplaceName(j) Then
                            SplitByEqual(1) = SrcBitBinaryString(j)
                            '20180523 TER
                            If gl_FlowForLoop_DigSrc_SweepCode <> "" Then
                                gl_FlowForLoop_DigSrc_SweepCode = gl_FlowForLoop_DigSrc_SweepCode & "&" & SplitByEqual(1)
                            Else
                                gl_FlowForLoop_DigSrc_SweepCode = SplitByEqual(1)
                            End If
                        End If
                    End If
                Next j
                DigSrc_Assignment_Update = DigSrc_Assignment_Update & SplitByEqual(0) & "=" & SplitByEqual(1) & ";"
        Next i
        
        DigSrc_Assignment_Update = Left(DigSrc_Assignment_Update, Len(DigSrc_Assignment_Update) - 1)
        DigSrc_Assignment = DigSrc_Assignment_Update
    
    End If
    
    Next z
    
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Modified DigSrc_Assignment = " & DigSrc_Assignment)

    If b_OnlySrcBitIndexStr Then
       ' If TPModeAsCharz_GLB = True Then
       '     Rtn_SweepTestName = CStr(Temp_SweepTestName) & "_"
       ' Else
            Rtn_SweepTestName = CStr(SrcBitIndex) & "_" & CStr(Temp_SweepTestName)
        'End If
    Else
        'If TPModeAsCharz_GLB = True Then
        '    Rtn_SweepTestName = SrcBitBinaryString(0)
        'Else
        If UBound(ForLoopIntegerNameSplitByLine) = 0 Then
            Temp_SweepTestName = "_"
            For i = 0 To UBound(RegReplaceName)
                Temp_SweepTestName = Temp_SweepTestName & RegReplaceName(i) & "_" & SrcBitBinaryString(i) & "_"
            Next i
            Temp_SweepTestName = Left(Temp_SweepTestName, Len(Temp_SweepTestName) - 1)
            Rtn_SweepTestName = CStr(SrcBitIndex) & Temp_SweepTestName
            
            Else
             '////////////////////////////  Print out for Loop to loop case/////////
                For z = 0 To UBound(ForLoopIntegerNameSplitByLine)
                    If z = 0 Then
                        Rtn_SweepTestName = Temp_SweepTestName_byline(z)
                    Else
                        Rtn_SweepTestName = "SDLLROCounter" & "-" & Temp_SweepTestName_byline(z) & "-" & Rtn_SweepTestName
        End If
                    gl_FlowForLoop_DigSrc_SweepCode = Rtn_SweepTestName
                Next z
              '////////////////////////////////////////////////////////////////////
             End If
        End If
End Function

Public Function EndSetupForMeasureVoltPins(MeasureV_Pin_PPMU As String, MeasureV_Pin_UVI80 As String) As String
    If MeasureV_Pin_PPMU <> "" Then TheHdw.PPMU.Pins(MeasureV_Pin_PPMU).Disconnect
    If MeasureV_Pin_UVI80 <> "" Then TheHdw.DCVI.Pins(MeasureV_Pin_UVI80).Disconnect tlDCVIConnectDefault
End Function

Public Function PinNameDuplicateCheck(PinName As String) As String
    Dim i As Integer
    Dim j As Integer
    Dim Pin_SplitArrayNum As Integer
    Dim Pin_SplitArray() As String
    Dim PinStringTemp As String
    Dim PinStringAfterCheck As String

    
    Pin_SplitArray = Split(PinName, ",")
    Pin_SplitArrayNum = UBound(Pin_SplitArray)
    ReDim b_FlagDuplicate(Pin_SplitArrayNum) As Boolean
    Dim b_FlagFirstTime As Boolean
    b_FlagFirstTime = False
    
    For i = 0 To Pin_SplitArrayNum
        PinStringTemp = Pin_SplitArray(i)
        For j = i + 1 To Pin_SplitArrayNum
            If Pin_SplitArray(j) = PinStringTemp Then
                b_FlagDuplicate(j) = True
            End If
        Next j
    Next i
    
    For i = 0 To Pin_SplitArrayNum
        If b_FlagDuplicate(i) = False Then
            If b_FlagFirstTime = False Then
                PinStringAfterCheck = Pin_SplitArray(i)
                b_FlagFirstTime = True
            Else
                PinStringAfterCheck = PinStringAfterCheck & "," & Pin_SplitArray(i)
            End If
            
        End If
    Next i
    PinNameDuplicateCheck = PinStringAfterCheck
End Function

Public Function Printout_DigSrc(DigSrc_array() As Long, DigSrc_Sample_Size As Long, Optional DigSrc_Width As Long = 4, Optional NumberPins As Long = 1, Optional site As Variant) As Long
    
    Dim i As Long, j As Long, cnt As Long
    Dim out_str As String

    out_str = "" & vbCrLf ' modified by Oscar 180523
''    Cnt = 0
If gl_Disable_HIP_debug_log = False Then

    '' 20160727 - Avoid DigSrc_Width divide by 0, DigSrc_Width only have relation with display format
        If NumberPins > 1 Then
            DigSrc_Width = NumberPins
        Else
            If DigSrc_Width = 0 Then
                DigSrc_Width = 4
            End If
        End If
    
        For i = 0 To DigSrc_Sample_Size - 1
            '' 20151230 - Modify to mod DigSrc_Width to let read more easy, default value is 4
            If (i Mod DigSrc_Width) = 0 And (Not i = 0) Then
                out_str = out_str & " "
                If (i Mod (DigSrc_Width * 5)) = 0 And (Not i = UBound(DigSrc_array)) Then out_str = out_str & vbCrLf ''added by Oscar 180523
            End If
            out_str = out_str & DigSrc_array(i)
        Next i
    
    If NumberPins > 1 And gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment "Site [" & site & "], Parallel to Binary of Src Bits = " & DigSrc_Sample_Size & ",Output String [ LSB(L) ==> MSB(R) ]: " & out_str
    ElseIf gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment "Site [" & site & "], Src Bits = " & DigSrc_Sample_Size & ",Output String [ LSB(L) ==> MSB(R) ]: " & out_str
    End If
End If
''    TheExec.Datalog.WriteComment ""
End Function

Public Function Printout_DigSrc_Newformat(DigSrcPrint() As String, DigSrc_Sample_Size As Long, Optional DigSrc_Width As Long = 4, Optional NumberPins As Long = 1, Optional site As Variant) As Long

    Dim i As Long, j As Long, cnt As Long
    Dim out_str As String
    out_str = "" & vbCrLf
'==========================180410 Added by Oscar
    If gl_Disable_HIP_debug_log = False Then
        For i = 0 To UBound(DigSrcPrint)
            out_str = out_str & DigSrcPrint(i) & ","
            If ((i + 1) Mod 5) = 0 And (Not i = UBound(DigSrcPrint)) Then out_str = out_str & vbCrLf
        Next i
'==========================180410 Added by Oscar
        
        If NumberPins > 1 Then
            TheExec.Datalog.WriteComment "Site [" & site & "] (Parellel)" & out_str
        Else
            TheExec.Datalog.WriteComment "Site [" & site & "]" & out_str  '========180410 Modified by Oscar
        End If
    End If
End Function

Public Function Create_DigSrc_Data(DigSrc_pin As PinList, DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, _
                        DigSrc_Equation As String, ByVal DigSrc_Assignment As String, InDspWav As DSPWave, site As Variant, Optional CUS_Str_DigSrcData As String = "", _
                        Optional NumberPins As Long = 1, Optional MSB_First_Flag As Boolean) As Long
''                        Optional InDSPWave_Parallel As DSPWave)

    Dim str_eq_ary(1000) As String_Equation
    Dim Assignment_ary() As String
    Dim Eq_ary() As String
    Dim i As Long, j As Long, k As Long
    Dim DigSrc_array() As Long
    Dim Str As String
    Dim idx As Long
    Dim Ary() As String
    Dim RdIn() As String
    
    Dim RdIn_tmp() As String
    Dim Rd_Fix_data As String

    Dim TempString_Repeat As String

    ''20161121-According to sample size and pin number to create data array size
    ReDim DigSrc_array(DigSrc_Sample_Size - 1)
    InDspWav.CreateConstant 0, DigSrc_Sample_Size
    
    ''20160824
    Dim b_WithDictionary As Boolean
    Dim SrcDspWave As New DSPWave
    Dim Ary_Src_DSPWave() As Long
    Dim b_Pre_data As Boolean
    
    Dim b_Append_Data As Boolean
    
    Dim PrePostData_SplitByAnd() As String
    Dim PrePostData_BinaryConstant As String
    
    ''20160909 - Append pre and post data  as 111&DictionA&000
    Dim b_AppendPrePostData As Boolean
    Dim PreBinDataString As String
    Dim PostBinDataString As String
    
    ''20160824-With "Repeat" keyword of DigSrc_Assignment
    If DigSrc_Assignment <> "" And InStr(LCase(DigSrc_Assignment), "repeat") = 0 Then
        Assignment_ary = Split(DigSrc_Assignment, ";")
        idx = 0
        
        For i = 0 To UBound(Assignment_ary)
           Ary = Split(Assignment_ary(i), "=")
          
           If UBound(Ary) > 1 Then
           Else
                str_eq_ary(idx).Name = Ary(0)
                
                ''20160825 - Pre/Post data as 111&DictionA or DictionA&101
                b_Pre_data = False
                b_Append_Data = False

                ''20160909 - Append pre and post data  as 111&DictionA&000
                b_AppendPrePostData = False
                PreBinDataString = ""
                PostBinDataString = ""
                PrePostData_SplitByAnd = Split(Ary(1), "&")
                
                If UBound(PrePostData_SplitByAnd) > 0 Then
                    If UBound(PrePostData_SplitByAnd) = 2 Then
                        b_AppendPrePostData = True
                        PreBinDataString = PrePostData_SplitByAnd(0)
                        Ary(1) = PrePostData_SplitByAnd(1)
                        PostBinDataString = PrePostData_SplitByAnd(2)
                    Else
                        b_Append_Data = True
                        
                        If Checker_ConstantBinary(PrePostData_SplitByAnd(0)) Then
                            b_Pre_data = True
                            Ary(1) = PrePostData_SplitByAnd(1)
                            PrePostData_BinaryConstant = PrePostData_SplitByAnd(0)
                        Else
                            b_Pre_data = False
                            Ary(1) = PrePostData_SplitByAnd(0)
                            PrePostData_BinaryConstant = PrePostData_SplitByAnd(1)
                        End If
                    End If
                End If
                ''20160825 - Check segment content whether Directionary
                ''20161014 - Check Dictionary whether need to calculation. EX: wdr0_4=010&[CAL_A-1]:0:3&110;wdr1_4=010&[CAL_A+1]:0:3&110
                RdIn = Split(Ary(1), ":")
                b_WithDictionary = Checker_WithDictionary(RdIn(0))

                If b_WithDictionary Then
                    
                    Dim b_IsDictNeedToCalc As Boolean
                    b_IsDictNeedToCalc = Checker_DictCalculated(RdIn(0))
                    
                    
                    
                    If b_IsDictNeedToCalc Then
                        Call AnalyzeDictCalculatedContent(RdIn(0), SrcDspWave)
                    Else
                        SrcDspWave = GetStoredCaptureData(RdIn(0))
                    End If
                
                    SrcDspWave = SrcDspWave.ConvertDataTypeTo(DspLong)
                    Ary_Src_DSPWave = SrcDspWave(site).Data
                    str_eq_ary(idx).value_string = ""

                    If UBound(RdIn) > 0 Then
                        
                        Dim StartNum As Long
                        Dim EndNum As Long
                        Dim StepNum As Long
                        StartNum = Int(RdIn(1))
                        EndNum = Int(RdIn(2))
                        If StartNum > EndNum Then
                            StepNum = -1
                        Else
                            StepNum = 1
                        End If
''                        For j = Int(RdIn(1)) To Int(RdIn(2))
                        For j = StartNum To EndNum Step StepNum
                            If UBound(RdIn) > 2 Then
                                If LCase(RdIn(3)) = "copy" Then
                                    For k = 0 To RdIn(4) - 1
                                        str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & Ary_Src_DSPWave(j)
                                    Next k
                                End If
                            Else
                                str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & Ary_Src_DSPWave(j)
                            End If
                        Next j
                    Else
                        For j = 0 To UBound(Ary_Src_DSPWave)
                            str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & Ary_Src_DSPWave(j)
                        Next j
                    End If
                    ''20160825 - Pre/Post data as 111&DictionA or DictionA&101
                    ''==========================================================================
                    If b_Append_Data Then
                        If b_Pre_data = False Then
                            str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & PrePostData_BinaryConstant
                        Else
                            str_eq_ary(idx).value_string = PrePostData_BinaryConstant & str_eq_ary(idx).value_string
                        End If
                    End If
                    ''==========================================================================
                    ''20160909 - Append pre and post data  as 111&DictionA&000
                    If b_AppendPrePostData Then
                        str_eq_ary(idx).value_string = PreBinDataString & str_eq_ary(idx).value_string & PostBinDataString
                    End If
                Else
                    If UBound(RdIn) > 0 Then
                        If LCase(RdIn(1)) = "copy" Then
                            For j = 0 To Len(RdIn(0)) - 1
                                For k = 0 To RdIn(2) - 1
                                    str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & Mid(RdIn(0), j + 1, 1)
                                Next k
                            Next j
                        End If
                    Else
                        str_eq_ary(idx).value_string = Ary(1)
                    End If
                    
                    ''20160825 - Pre/Post data as 111&DictionA or DictionA&101
                    ''==========================================================================
                    If b_Append_Data Then
                        If b_Pre_data = False Then
                            str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & PrePostData_BinaryConstant
                        Else
                            str_eq_ary(idx).value_string = PrePostData_BinaryConstant & str_eq_ary(idx).value_string
                        End If
                    End If
                    ''==========================================================================
                    ''20160909 - Append pre and post data  as 111&DictionA&000
                    If b_AppendPrePostData Then
                        str_eq_ary(idx).value_string = PreBinDataString & str_eq_ary(idx).value_string & PostBinDataString
                    End If
                End If
                idx = idx + 1
            End If
        Next i
    End If

    For i = 0 To idx - 1
        Str = Str & str_eq_ary(i).Name & ":" & str_eq_ary(i).value_string & ","
    Next i

    'TheExec.Datalog.WriteComment "Site [" & site & "] " & Str
            
    If DigSrc_Equation <> "" Then
        Eq_ary = Split(DigSrc_Equation, "+")
        idx = 0
        

        Dim StringDecomposeEqual() As String
        Dim StringDecomposeColon() As String
        Dim SelectedSrcDSPWave As New DSPWave
        
        ''20160830
        Dim b_StrInvolveCopy As Boolean
        '' 20160122 - Modify rule to soruce assigned bit from rd to replace souce all of bits, EX: repeat=rd:1:3
        Dim StartBit As Long
        Dim EndBit As Long
        
        Dim b_SelectallBits As Boolean
        Dim DataCopyTimes As Long
        Dim LoopIndex As Long
        
        If InStr(LCase(DigSrc_Assignment), "repeat") <> 0 Then
            StringDecomposeEqual = Split(DigSrc_Assignment, "=")
            
            ''20160830-Check DigSrc_Assignment has "copy" or not
            If InStr(LCase(DigSrc_Assignment), "copy") <> 0 Then
                b_StrInvolveCopy = True
            Else
                b_StrInvolveCopy = False
            End If
                        
            StringDecomposeColon = Split(LCase(StringDecomposeEqual(1)), ":")
            b_WithDictionary = Checker_WithDictionary(StringDecomposeColon(0))
            
            If b_WithDictionary Then
                SelectedSrcDSPWave = GetStoredCaptureData(StringDecomposeColon(0))
                
                    If UBound(StringDecomposeColon) = 4 Then
                        ''CAL_A:1:3:copy:2
                    StartBit = StringDecomposeColon(1)
                    EndBit = StringDecomposeColon(2)
                    DataCopyTimes = StringDecomposeColon(4)
                
                ElseIf UBound(StringDecomposeColon) = 2 Then
                    If b_StrInvolveCopy = True Then ''CAL_A:copy:2
                        StartBit = 0
                        EndBit = SelectedSrcDSPWave.SampleSize - 1
                        DataCopyTimes = StringDecomposeColon(2)
                    
                    Else                                        ''CAL_A:1:3
                        StartBit = StringDecomposeColon(1)
                        EndBit = StringDecomposeColon(2)
                        DataCopyTimes = 1
                    End If

                ElseIf UBound(StringDecomposeColon) = 0 Then
                    ''CAL_A
                    StartBit = 0
                    EndBit = SelectedSrcDSPWave.SampleSize - 1
                    DataCopyTimes = 1
                End If
            Else

                If b_StrInvolveCopy Then
                    ''1101:copy:2
                    DataCopyTimes = StringDecomposeColon(2)
                Else                                ''1101
                    DataCopyTimes = 1
                End If
            End If
            
            LoopIndex = 0
            If b_WithDictionary Then
            
                Dim StepSize As Long
                If StartBit > EndBit Then
                    StepSize = -1
                Else
                    StepSize = 1
                End If
                
                For i = StartBit To EndBit Step StepSize
                    For j = 0 To DataCopyTimes - 1
                        If LoopIndex = 0 Then
                            TempString_Repeat = SelectedSrcDSPWave(site).Element(i)
                        Else
                            TempString_Repeat = TempString_Repeat & SelectedSrcDSPWave(site).Element(i)
                        End If
                        LoopIndex = LoopIndex + 1
                    Next j
                Next i
                
            Else
                For i = 0 To Len(StringDecomposeColon(0)) - 1
                    For j = 0 To DataCopyTimes - 1
                        TempString_Repeat = TempString_Repeat & Mid(StringDecomposeColon(0), i + 1, 1)
                    Next j
                Next i
            End If
            TempString_Repeat = "repeat=" & TempString_Repeat
            DigSrc_Assignment = TempString_Repeat

        End If
        
        ''20160824-Final process, Analyze DigSrc_Assignment to create InDspWav
        '' Number of equation segment
        Dim DigSrcPrint() As String '======Added by Oscar 180523
        For i = 0 To UBound(Eq_ary)
            ReDim Preserve DigSrcPrint(i) As String '===================180416 Modified by Oscar
            If InStr(LCase(DigSrc_Assignment), "repeat") <> 0 Then
                If InStr(DigSrc_Assignment, "=") <> 0 Then RdIn = Split(DigSrc_Assignment, "=")
                If InStr(DigSrc_Assignment, ",") <> 0 Then RdIn = Split(DigSrc_Assignment, ",")
                Str = Trim(RdIn(UBound(RdIn)))
                DigSrcPrint(i) = Str & "(" & Eq_ary(i) & ")" '===================180424 Modified by Oscar
            Else
                Str = Trim(Find_Assignement(Eq_ary(i), str_eq_ary, DigSrcPrint(i), NumberPins, MSB_First_Flag))
            End If
                
            '' Number of DigSrc_Assignment content
            For j = 1 To Len(Str)
                DigSrc_array(idx) = Val(Mid(Str, j, 1))  ''InDspWav(Site).Element(idx) = DigSrc_array(idx)
                idx = idx + 1
            Next j
        Next i
    
        InDspWav(site).Data = DigSrc_array

        If idx <> DigSrc_Sample_Size And gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Num of bits in digsrc equation(" & idx & ") is not the same as DigSrc_SampleSize(" & DigSrc_Sample_Size & ")"
    End If
    
    Call Printout_DigSrc_Newformat(DigSrcPrint, DigSrc_Sample_Size, DigSrc_DataWidth, NumberPins, site)

End Function

Public Function SetupDigSrcDspWave(patt As String, DigSrcPin As PinList, SignalName As String, SegmentSize As Long, InWave As DSPWave)

    Dim site As Variant
    Dim WaveDef As String
    
    WaveDef = "WaveDef" & SignalName
    ''20150708 - Comment program load
    TheHdw.Patterns(patt).Load  ' 20151211: addedback to fix error re: pattern not being loaded
    
'**************************************************
    'SeaHawk Edited by 20190606
    Dim DebugStr As String              'Avoid IGXL Bug
    TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Reinitialize
    For Each site In TheExec.sites
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & site & "_" & CStr(DigSrcPin), InWave, True
        TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.Add SignalName & "_" & CStr(DigSrcPin)
        With TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals(SignalName & "_" & CStr(DigSrcPin))
            .WaveDefinitionName = WaveDef & site & "_" & CStr(DigSrcPin)
            .SampleSize = SegmentSize
            .Amplitude = pc_Def_DSSC_Amplitude
            .LoadSamples
            .LoadSettings
        End With
        TheHdw.DSSC.Pins(DigSrcPin).Pattern(patt).Source.Signals.DefaultSignal = SignalName & "_" & CStr(DigSrcPin)
    Next site
    
End Function


Public Function DigCapSetup(patt As String, DigCapPin As PinList, SignalName As String, SampleSize As Long, ByRef DspWav As DSPWave)
    
    '' 20150813 - Need to put because need to double confirm the pattern whether loaded before
    TheHdw.Patterns(patt).Load

'**************************************************
'SeaHawk Edited by 20190606
    Dim DebugStr As String
    TheHdw.DSSC.Pins(DigCapPin).Pattern(patt).Capture.Signals.Reinitialize
        '' 20150812-Modify program to process multiply dig cap pins
        With TheHdw.DSSC.Pins(DigCapPin).Pattern(patt).Capture.Signals
            .Add (SignalName & SampleSize & "_" & DigCapPin)
            With .Item(SignalName & SampleSize & "_" & DigCapPin)
                .SampleSize = SampleSize    'CaptureCyc * OneCycle
                .LoadSettings
            End With
        End With
        DspWav = TheHdw.DSSC.Pins(DigCapPin).Pattern(patt).Capture.Signals(SignalName & SampleSize & "_" & DigCapPin).DSPWave
'**************************************************
    
    'Create capture waveform
    
    '' 20150813 - Assign WaveName to the DSPWave to do recognition of post process.
    Dim site As Variant
    For Each site In TheExec.sites
        DspWav(site).Info.WaveName = DigCapPin
    Next site
    
   ' TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug ' use defaut as automatic
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
End Function


Public Function MergePinListData(Meas_INstType_Num As Integer, Measurement() As PinListData, ByRef MergedData As PinListData) As Long
    '' 20150608 - Merge Measurement to the same pin list data if instrument over 1 type.
    Dim index As Integer
    Dim p As Integer
    For index = 0 To Meas_INstType_Num - 1
        For p = 0 To Measurement(index).Pins.Count - 1
            MergedData.AddPin (Measurement(index).Pins(p))
            MergedData.Pins(Measurement(index).Pins(p)) = Measurement(index).Pins(p)
        Next p
    Next index

End Function
                       
Public Function Freq_PPMU_Meas_VOH(Meas_Pin As PinList, percentage As Double, _
        Optional VtMode As Boolean = False, _
        Optional MeasF_EventSource As FreqCtrEventSrcSel)  'percentage = 0.01 ~ 0.99
    
    Dim PinArr() As String
    Dim PinCount As Long
    Dim Pin As Variant
    Dim meas_value As New PinListData
    Dim Val_VOH As Double
    Dim i As Long
    Dim org_VOH As New PinListData
    Dim site As Variant
    Dim org_VOL As New PinListData
    Dim Val_VOL As Double
    
    TheExec.DataManager.DecomposePinList Meas_Pin, PinArr, PinCount
    
    For i = 0 To PinCount - 1
        org_VOH.AddPin (PinArr(i))
        org_VOL.AddPin (PinArr(i))
        
        For Each site In TheExec.sites
            
            org_VOH.Pins(PinArr(i)).Value(site) = TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVoh)
            
            '' 20150806 - Store original VOL value if EventSource = BOTH
            If MeasF_EventSource = BOTH Then
                org_VOL.Pins(PinArr(i)).Value(site) = TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVol)
            End If
        Next site
    Next i

    TheHdw.Digital.Pins(Meas_Pin).Disconnect
    With TheHdw.PPMU.Pins(Meas_Pin)   '' make sure which pins
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
    End With

    TheHdw.Wait (1 * ms)
    DebugPrintFunc_PPMU CStr(Meas_Pin)
    meas_value = TheHdw.PPMU.Pins(Meas_Pin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint_FreqDC)

    For i = 0 To PinCount - 1
        For Each site In TheExec.sites
            If MeasF_EventSource = VOH Then
                If meas_value.Pins(PinArr(i)).Value(site) < org_VOH.Pins(PinArr(i)).Value(site) * (1 + percentage) And meas_value.Pins(PinArr(i)).Value(site) > org_VOH.Pins(PinArr(i)).Value(site) * (1 - percentage) Then
                    '' 20150806 - Comment it because it to do the same thing
                    TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVoh) = meas_value.Pins(PinArr(i)).Value(site)
                    
                    Val_VOH = meas_value.Pins(PinArr(i)).Value(site)
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & PinArr(i) & " , " & "VOH= " & Format(Val_VOH, "0.000") & " V"
                
                Else
                    Val_VOH = org_VOH.Pins(PinArr(i)).Value(site)
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & PinArr(i) & " , " & "VOH= " & Format(Val_VOH, "0.000") & " V, search fail, use defalut value"
                End If
                
                If VtMode = True Then
                    'Vt Mode
                    TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVt) = Val_VOH
                    TheHdw.Digital.Pins(PinArr(i)).Levels.DriverMode = tlDriverModeVt
                End If
                
            ElseIf MeasF_EventSource = BOTH Then
                '' 20150806 - When "BOTH" is chosen, set VOH=ppmu_meas*(1+factor) and VOL=ppmu_meas*(1-factor).
                Dim VOH_ThresholdValue As Double
                Dim VOL_ThresholdValue As Double
                VOH_ThresholdValue = meas_value.Pins(PinArr(i)).Value(site) * (1 + percentage)
                VOL_ThresholdValue = meas_value.Pins(PinArr(i)).Value(site) * (1 - percentage)
                
                If VOH_ThresholdValue > 6 Then
                    TheExec.Datalog.WriteComment ("VOH_ThresholdValue = " & VOH_ThresholdValue & "V, Over spec and clamp to 6V")
                    VOH_ThresholdValue = 6
                End If
                If VOH_ThresholdValue < -1 Then
                    TheExec.Datalog.WriteComment ("VOH_ThresholdValue = " & VOH_ThresholdValue & "V, less than spec and clamp to -1V")
                    VOH_ThresholdValue = -1
                End If
                If VOL_ThresholdValue > 6 Then
                    TheExec.Datalog.WriteComment ("VOL_ThresholdValue = " & VOL_ThresholdValue & "V, Over spec and clamp to 6V")
                    VOL_ThresholdValue = 6
                End If
                If VOL_ThresholdValue < -1 Then
                    TheExec.Datalog.WriteComment ("VOL_ThresholdValue = " & VOL_ThresholdValue & "V, less than spec and clamp to -1V")
                    VOL_ThresholdValue = -1
                End If
                
                TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVoh) = VOH_ThresholdValue
                TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVol) = VOL_ThresholdValue
                Val_VOH = TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVoh)
                Val_VOL = TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVol)
                
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & PinArr(i) & " , " & "VOH= " & Format(Val_VOH, "0.000") & " V"
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & PinArr(i) & " , " & "VOL= " & Format(Val_VOL, "0.000") & " V"
                
                If VtMode = True Then
                    'Vt Mode
                                        If TheExec.TesterMode = testModeOffline Then
                            If meas_value.Pins(PinArr(i)).Value(site) > 6 Then
                                        TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVt) = 6
                            ElseIf meas_value.Pins(PinArr(i)).Value(site) < -1 Then
                                        TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVt) = -1
                            Else
                                        TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVt) = meas_value.Pins(PinArr(i)).Value(site)
                            End If
                                        Else
                                                TheHdw.Digital.Pins(PinArr(i)).Levels.Value(chVt) = meas_value.Pins(PinArr(i)).Value(site)
                                        End If
                    TheHdw.Digital.Pins(PinArr(i)).Levels.DriverMode = tlDriverModeVt
                End If
            End If
        Next site
    Next i
    
    '' 20150615 - Add gate off
    With TheHdw.PPMU.Pins(Meas_Pin)
        .Gate = tlOff
        .Disconnect
    End With
    TheHdw.Wait (1 * us)
    TheHdw.Digital.Pins(Meas_Pin).Connect
End Function
Public Function IO_HardIP_PPMU_Measure_V(TestPinArrayIV() As String, TestSeqNum As Integer, TestSeqNumIdx As Long, ForceSequenceArray() As String, _
    k As Long, Pat As Variant, Flag_SingleLimit As Boolean, HighLimitVal As Double, LowLimitVal As Double, TestLimitPerPin_VIR As String, ByRef ReturnMeasVolt As PinListData, _
    FlowTestNme() As String, _
    Optional SpecialCalcValSetting As CalculateMethodSetup = 0, _
    Optional InstSpecialSetting As InstrumentSpecialSetup = 0, Optional RAK_Flag As Enum_RAK = 0, _
    Optional CUS_Str_MainProgram As String = "", Optional Rtn_SweepTestName As String, Optional OutputTname As String, Optional WaitTime_V As String) As Long

    Dim MeasureValue As New PinListData
    Dim Force_idx As Integer
    Dim site As Variant
    Dim TestNum As Long
    Dim Pin  As Variant
    
    Dim p As Long
    Dim ForceV  As Double
    Dim ForceByPin() As String
    Dim ForceValByPin() As String
    Dim ForceValIdx As Integer
    Dim IdxV As Integer
    Dim MeasurePinStr As String
    Dim Temp_index As Long
    Dim OutputTname_format() As String
    
    ''========================================================================================
    '' 20150202 - Range Check
    Dim TempMeasVal_PerPin(100) As New PinListData
    If UBound(ForceSequenceArray) = 0 Then
        ForceValByPin = Split(ForceSequenceArray(0), ",")
    Else
        If (UBound(ForceSequenceArray) >= TestSeqNumIdx) Then
            If ForceSequenceArray(TestSeqNumIdx) = "" Then ForceSequenceArray(TestSeqNumIdx) = 0
            ForceValByPin = Split(ForceSequenceArray(TestSeqNumIdx), ",")
        End If
    End If
    ForceValIdx = 0
    
    If UBound(TestPinArrayIV) = 0 Then
        ForceByPin = Split(TestPinArrayIV(0), ",")
         MeasurePinStr = TestPinArrayIV(0)       '20160224 add to allow every seq with the same pins
        'TheHdw.Digital.Pins(TestPinArrayIV(0)).Disconnect
    Else
        ForceByPin = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        MeasurePinStr = TestPinArrayIV(TestSeqNumIdx)       '20160224 add to allow every seq with the same pins
        'TheHdw.Digital.Pins(TestPinArrayIV(TestSeqNumIdx)).Disconnect
    End If
    '' 20150721 - Apply force I value from Stored_MeasI_PPMU,
    ''                - Can not coexist between stored value and force value at the same sequence
    
    Dim b_IsNumeral As Boolean
    Dim b_UseStoredForceVal As Boolean
    
    b_IsNumeral = ContentIsNumeral(ForceValByPin(0))
    If b_IsNumeral Then
        b_UseStoredForceVal = False
    Else
        b_UseStoredForceVal = True
    End If
    Dim ForceValI As Double
    If b_UseStoredForceVal = False Then
        '' 20150721 - Normal usage
        For Each Pin In ForceByPin
        
            With TheHdw.PPMU.Pins(Pin)
                If InStr(CUS_Str_MainProgram, "PCIE_Init0") <> 0 Then
                   .ForceI 0, 0
                Else
                '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
                   .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                End If
                .Connect
                .Gate = tlOn

                If InstSpecialSetting = InstrumentSpecialSetup.PPMU_SerialMeasurement Then
                    '20160121, update by flag control
                    'don't force current when do serial measure V
                ElseIf UBound(ForceValByPin) = 0 Then
                    .ForceI ForceValByPin(0), ForceValByPin(0)
                    ForceValI = ForceValByPin(0)
                ElseIf ForceValByPin(ForceValIdx) <> "" Then
                    .ForceI ForceValByPin(ForceValIdx), ForceValByPin(ForceValIdx)
                     ForceValI = ForceValByPin(ForceValIdx)
                Else:
                    .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_InitialValue_FI_Range
                     ForceValI = 0
                End If
            End With
            
            ForceValIdx = ForceValIdx + 1
        Next Pin
    
    Else
        '' 20150721 - Apply stored value
        Dim AfterformulaVal_PPMU As New PinListData
''        Call CUS_FormulaCalc(Stored_MeasI_PPMU, AfterformulaVal_PPMU)
        
        '' 20150721 - Store ForceValue for each site for test limit use.
        Dim TestPinMaxNum As Integer
        TestPinMaxNum = UBound(ForceByPin)
        ReDim StoreForceI(TestPinMaxNum) As New SiteDouble
        
        For Each Pin In ForceByPin
            For Each site In TheExec.sites.Active
                With TheHdw.PPMU.Pins(Pin)

                    If InstSpecialSetting = InstrumentSpecialSetup.PPMU_SerialMeasurement Then
                        '20160121, update by flag control
                      'do nothing
                    ElseIf UBound(ForceValByPin) = 0 Then
                        .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(0)).Value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(0)).Value(site)
                        StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site)
                    ElseIf ForceValByPin(ForceValIdx) <> "" Then
                        .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site)
                        StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).Value(site)
                    Else
                        .ForceI pc_Def_PPMU_InitialValue_FI
                    End If
                End With
            Next site
            ForceValIdx = ForceValIdx + 1
        Next Pin
    End If
        
    If UBound(ForceSequenceArray) <> 0 Then
        If ForceSequenceArray(TestSeqNum) = "" Then
            ForceSequenceArray(TestSeqNum) = 0
        End If
    End If

    For Each site In TheExec.sites.Active
        TestNum = TheExec.sites.Item(site).TestNumber
    Next site

    If WaitTime_V = "" Then
        TheHdw.Wait (1 * ms)
    Else
        TheHdw.Wait CDbl(WaitTime_V)
    End If
 
    If InstSpecialSetting = InstrumentSpecialSetup.DigitalConnectPPMU2 Then
        TheHdw.PPMU.AllowPPMUFuncRelayConnection (True)
        TheHdw.PPMU.Pins(MeasurePinStr).ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_FI_Range_200uA '20160224 add to allow every seq with the same pins
        TheHdw.Digital.Pins(MeasurePinStr).Connect
    End If
    If InstSpecialSetting = InstrumentSpecialSetup.PPMU_SerialMeasurement Then
        Dim PinArr() As String
        Dim PinCount As Long
        Dim temp_MeasureValue As New PinListData '''''20180623 serial measure print force condition
                
        TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
        TheHdw.PPMU.Pins(MeasurePinStr).ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
        ForceValI = ForceValByPin(0)
        For Each Pin In PinArr
            MeasureValue.AddPin (Pin)
            TheHdw.PPMU.Pins(Pin).ForceI ForceValByPin(0), Abs(ForceValByPin(0))
            TheHdw.Wait 0.001
            DebugPrintFunc_PPMU CStr(Pin)
            MeasureValue.Pins(Pin) = TheHdw.PPMU.Pins(Pin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                        '''''20180623 serial measure print force condition
            temp_MeasureValue.AddPin(Pin) = MeasureValue.Pins(Pin)
            Call Print_Force_Condition("v", temp_MeasureValue)  ''''20180623 position check
            Set temp_MeasureValue = Nothing
            '=====================================================
            TheHdw.PPMU.Pins(Pin).ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
            TheHdw.PPMU.Pins(Pin).Gate = tlOff
            TheHdw.PPMU.Pins(Pin).Disconnect
            TheHdw.Digital.Pins(Pin).Connect
        Next Pin
    Else
        DebugPrintFunc_PPMU CStr(MeasurePinStr)
        MeasureValue = TheHdw.PPMU.Pins(MeasurePinStr).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
        Call Print_Force_Condition("v", MeasureValue) ''''20180623 position check
    End If
    
    '' Calculate RAK
    'Dim RakV() As Double
    If RAK_Flag = R_TraceOnly Then
        For Each site In TheExec.sites
            For Each Pin In MeasureValue.Pins
                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(Pin, Site)
                MeasureValue.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site) - ForceValI * (CurrentJob_Card_RAK.Pins(Pin).Value(site))
            Next Pin
        Next site
    ElseIf RAK_Flag = R_PathWithContact Then
         For Each site In TheExec.sites
            For Each Pin In MeasureValue.Pins
                  MeasureValue.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site) - ForceValI * R_Path_PLD.Pins(Pin).Value(site)
             Next Pin
        Next site
    ElseIf InStr(UCase(CUS_Str_MainProgram), UCase("RREF_RAK_CALC")) <> 0 Then
        Call CUS_RREF_Rak_Calc(MeasureValue)
    End If
    
    '' 20170710 - Midify hard code to use enum >> SpecialCalcValSetting = VIR_DDIO
''    If (UCase(TheExec.DataManager.InstanceName) Like UCase("*_DDIO_MEA*VMX*_T*_F*_*V*") And UCase(CUS_Str_MainProgram) Like UCase("*VOL,VOH*")) Or UCase(TheExec.DataManager.InstanceName) Like UCase("*_AJ00_*DIO_*33_*") Or UCase(TheExec.DataManager.InstanceName) Like UCase("*_AJ00_*DIO_*35_*") Then
    If SpecialCalcValSetting = CalculateMethodSetup.VIR_DDIO Then
        Call CUS_DDR_Emulate_Const_Res_Loading(MeasureValue, ForceValByPin(), CUS_Str_MainProgram, TestSeqNum, RAK_Flag)
    End If
    
    ''20170107-Add SpecialCalcValSetting for VIR_DIFF_PN = 2 and VIR_VOD_VOCM_PN = 3
    Dim MeasDiff_V As New PinListData
    Dim MeasVOD_V As New PinListData
    Dim MeasVOCM_V As New PinListData

    If SpecialCalcValSetting = CalculateMethodSetup.VIR_DIFF_PN Then
        For p = 0 To MeasureValue.Pins.Count - 1 Step 2
            MeasDiff_V.AddPin (MeasureValue.Pins(p + 1).Name)
            For Each site In TheExec.sites.Active
                MeasDiff_V.Pins(MeasureValue.Pins(p + 1).Name).Value = (MeasureValue.Pins(p).Value - MeasureValue.Pins(p + 1).Value)
            Next site
        Next p
    ElseIf SpecialCalcValSetting = CalculateMethodSetup.VIR_DIFF_PN_ABS Then 'ABS 20180110 backup
        For p = 0 To MeasureValue.Pins.Count - 1 Step 2
            MeasDiff_V.AddPin (MeasureValue.Pins(p + 1).Name)
            For Each site In TheExec.sites.Active
                MeasDiff_V.Pins(MeasureValue.Pins(p + 1).Name).Value = Abs(MeasureValue.Pins(p).Value - MeasureValue.Pins(p + 1).Value)
            Next site
        Next p
    ElseIf SpecialCalcValSetting = CalculateMethodSetup.VIR_VOD_VOCM_PN Then 'add by JiYi 20160721 for refbuf
        For p = 0 To MeasureValue.Pins.Count - 1 Step 2
            MeasVOD_V.AddPin (MeasureValue.Pins(p + 1).Name)
            MeasVOCM_V.AddPin (MeasureValue.Pins(p + 1).Name)
            
            For Each site In TheExec.sites.Active
                MeasVOD_V.Pins(MeasureValue.Pins(p + 1).Name).Value = MeasureValue.Pins(p + 1).Value - MeasureValue.Pins(p).Value
                MeasVOCM_V.Pins(MeasureValue.Pins(p + 1).Name).Value = 0.5 * (MeasureValue.Pins(p + 1).Value + MeasureValue.Pins(p).Value)
            Next site
        Next p
    End If
    
    '' 20150728 - Add return measure volt to main function.

    If SpecialCalcValSetting = CalculateMethodSetup.VIR_DIFF_PN Or SpecialCalcValSetting = CalculateMethodSetup.VIR_DIFF_PN_ABS Then
        ''20180110 ''update ABS.
        ReturnMeasVolt = MeasDiff_V
    Else
        ReturnMeasVolt = MeasureValue
    End If
    
    Force_idx = TestSeqNum
    If UBound(ForceSequenceArray) = 0 Then
        Force_idx = 0
    End If

    Dim TestNameInput As String

    'If TPModeAsCharz_GLB = True Then
    '    If SpecialCalcValSetting = VIR_VOD_VOCM_PN Or VIR_DIFF_PN Then
    '        gl_CZ_FlowTestName_Counter = 0
    '    Else
    '        TestNameInput = FlowTestNme(gl_CZ_FlowTestName_Counter)
    '        gl_CZ_FlowTestName_Counter = gl_CZ_FlowTestName_Counter + 1
    '    End If
    'Else
        TestNameInput = "Volt_meas_" + CStr(TestSeqNum)
        If Rtn_SweepTestName <> "" Then
            TestNameInput = TestNameInput & "_" & Rtn_SweepTestName
        End If
    'End If
        
    If CUS_Str_MainProgram <> "" Then

        If LCase(CUS_Str_MainProgram) Like LCase("tname*") Then
            ' Modify Tname that refer to CUS_Str_MainProgram . Ex: Tname:VOL,VOH,VOL,VOH
            Dim Temp_Input() As String
            Temp_Input() = Split(CUS_Str_MainProgram, ":")
            Temp_Input() = Split(Temp_Input(1), ",")
            TestNameInput = Temp_Input(TestSeqNum) + "_" + CStr(TestSeqNum)
        End If
    End If
        
    '''20151103 print force condition
'    If SpecialCalcValSetting <> VIR_DDIO Then Call Print_Force_Condition("v", MeasureValue) ''20180619 update slip print DDR IO force condition by normal format
    
    '' 20150721 - Test limit for force stored value
    Dim ForceIndex As Integer
    ForceIndex = 0
    If b_UseStoredForceVal = True Then
        For Each Pin In ForceByPin
            'For Each Site In TheExec.Sites
                TestNameInput = Report_TName_From_Instance("V", MeasureValue.Pins(Pin), "VOD", TestSeqNum, CLng(ForceIndex))
                TheExec.Flow.TestLimit MeasureValue.Pins(Pin), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=StoreForceI(ForceIndex).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            'Next Site
            ForceIndex = ForceIndex + 1
        Next Pin
      '  Exit Function
    'End If
    ElseIf InStr(CUS_Str_MainProgram, "DDR_VOHL") <> 0 Then
         
        Dim HiLimitVal As Integer
        Dim LoLimitVal As Integer
        HiLimitVal = 0: LoLimitVal = 0
        If CUS_Str_MainProgram = "DDR_VOHL_1" Then
            HiLimitVal = 132: LoLimitVal = 108
        ElseIf CUS_Str_MainProgram = "DDR_VOHL_2" Then
            If TestSeqNumIdx = 0 Then HiLimitVal = 42: LoLimitVal = 38
            If TestSeqNumIdx = 1 Then HiLimitVal = 176: LoLimitVal = 144
        ElseIf CUS_Str_MainProgram = "DDR_VOHL_3" Then
            HiLimitVal = 264: LoLimitVal = 216
        Else
            If TestSeqNumIdx = 0 Then HiLimitVal = 132: LoLimitVal = 108
        End If
    
        For p = 0 To MeasureValue.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("V", MeasureValue.Pins(p), , TestSeqNum, p)
            If TestSeqNumIdx = 0 Then
                TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
                TheExec.Flow.TestLimit MeasureValue.Pins(p).Divide(ForceValI), LoLimitVal, HiLimitVal, scaletype:=scaleNone, Unit:=unitCustom, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI, customUnit:="ohm" 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            Else
                TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
                TheExec.Flow.TestLimit MeasureValue.Pins(p).Subtract(1.1).Divide(ForceValI).Abs, LoLimitVal, HiLimitVal, scaletype:=scaleNone, Unit:=unitCustom, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI, customUnit:="ohm" 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            End If
        Next p
    ElseIf Flag_SingleLimit = True Then
        If SpecialCalcValSetting = CalculateMethodSetup.VIR_DDIO Then
        Else
                        Temp_index = TheExec.Flow.TestLimitIndex
                For p = 0 To MeasureValue.Pins.Count - 1
                TheExec.Flow.TestLimitIndex = Temp_index
                    TestNameInput = Report_TName_From_Instance("V", MeasureValue.Pins(p), , TestSeqNum, p)
                TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=ForceSequenceArray(Force_idx), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Next p
                End If
    Else
        ''20170106-Add TestLimit by SpecialCalcValSetting with different result value
        If SpecialCalcValSetting = CalculateMethodSetup.VIR_DIFF_PN Or SpecialCalcValSetting = CalculateMethodSetup.VIR_DIFF_PN_ABS Then
            For p = 0 To MeasDiff_V.Pins.Count - 1
                TestNameInput = Report_TName_From_Instance("V", MeasDiff_V.Pins(p), "Vdiff", TestSeqNum, p)
                TheExec.Flow.TestLimit MeasDiff_V.Pins(p), , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next p
        
        ElseIf SpecialCalcValSetting = CalculateMethodSetup.VIR_VOD_VOCM_PN Then 'add by JiYi 20160721 for refbuf
            Dim t As Long
            
            For t = 0 To MeasureValue.Pins.Count - 1
                If CUS_Str_MainProgram = "TTR" Then
                    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
                Else
                TestNameInput = Report_TName_From_Instance("V", MeasureValue.Pins(t), "Pin", TestSeqNum, t)
                TheExec.Flow.TestLimit MeasureValue.Pins(t), , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
                End If
            Next t
            
            For t = 1 To (MeasureValue.Pins.Count) - 1 Step 2
                TestNameInput = Report_TName_From_Instance("V", MeasureValue.Pins(t), "vdiff", TestSeqNum, t)
                TheExec.Flow.TestLimit MeasVOD_V.Pins((t - 1) / 2), , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next t
            For t = 1 To (MeasureValue.Pins.Count) - 1 Step 2
                TestNameInput = Report_TName_From_Instance("V", MeasureValue.Pins(t), "vocm", TestSeqNum, t)
                TheExec.Flow.TestLimit MeasVOCM_V.Pins((t - 1) / 2), , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next t
        Else
            If Mid(TestLimitPerPin_VIR, 1, 1) = "F" And UBound(ForceValByPin) = 0 Then
                
            For t = 0 To (MeasureValue.Pins.Count / 2) - 1
                    TestNameInput = Report_TName_From_Instance("V", MeasureValue.Pins(t), , TestSeqNum, t)
                    TheExec.Flow.TestLimit MeasureValue.Pins(t), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=ForceSequenceArray(Force_idx), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Next t
            ElseIf Mid(TestLimitPerPin_VIR, 1, 1) = "T" And UBound(ForceByPin) = MeasureValue.Pins.Count - 1 Then
                IdxV = 0

                     For p = 0 To MeasureValue.Pins.Count - 1
                        TestNameInput = Report_TName_From_Instance("V", MeasureValue.Pins(p), "VOCM", TestSeqNum, p)
                        TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=ForceValByPin(IdxV), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                        
                        If UBound(ForceValByPin) = 0 Then
                        IdxV = 0
                    Else
                        IdxV = IdxV + 1
                    End If
                    Next p
               ' End If
                
            ElseIf Mid(TestLimitPerPin_VIR, 1, 1) = "T" And UBound(ForceByPin) <> MeasureValue.Pins.Count - 1 Then
                '' 20170710 - Midify hard code to use enum >> SpecialCalcValSetting = VIR_DDIO
''                If (UCase(TheExec.DataManager.InstanceName) Like UCase("*_DDIO_MEA*VMX*_T*_F*_*V*") And UCase(CUS_Str_MainProgram) Like UCase("*VOL,VOH*")) Then
                If SpecialCalcValSetting = CalculateMethodSetup.VIR_DDIO Then
                Else
                    Dim dataStr() As String
                    ReDim dataStr(UBound(ForceByPin)) As String
                    Dim i As Long
                    Dim DS_ForceVal() As String
                    ReDim DS_ForceVal(MeasureValue.Pins.Count - 1) As String
                    Dim DePins() As String
                    Dim NumberPins As Long
                    For i = 0 To UBound(ForceByPin)
                        dataStr(i) = ""
                        Call TheExec.DataManager.DecomposePinList(ForceByPin(i), DePins(), NumberPins)
                        For Each Pin In DePins
                            dataStr(i) = dataStr(i) & "," & Pin
                        Next Pin
                    Next i
                        
                    IdxV = 0
                    For p = 0 To MeasureValue.Pins.Count - 1
                        For i = 0 To UBound(ForceByPin)
                            If InStr(LCase(dataStr(i)), LCase(MeasureValue.Pins(p))) <> 0 Then
                                DS_ForceVal(IdxV) = ForceValByPin(i)
                                TestNameInput = Report_TName_From_Instance("V", MeasureValue.Pins(p), , TestSeqNum, p)
                                TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=DS_ForceVal(IdxV), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                                IdxV = IdxV + 1
                            End If
                        Next i
                    Next p
                End If
                
            Else
            End If
        End If
    End If
       
    If InstSpecialSetting = InstrumentSpecialSetup.DigitalConnectPPMU2 Then
        '20160204
        TheHdw.Digital.Pins(TestPinArrayIV(TestSeqNumIdx)).Disconnect ' 20160217 fixed LPDP_RX vil/vih failure
        TheHdw.PPMU.AllowPPMUFuncRelayConnection (False)
    End If
    
    If SpecialCalcValSetting = CalculateMethodSetup.VIR_VOD_VOCM_XI0_Off Then
        For Each site In TheExec.sites.Active
            TheHdw.Digital.Pins("XI0").Levels.Value(chVih) = TheExec.specs.DC("Pins_1p8v_Vih_VAR_H").CurrentValue
            TheHdw.Digital.Pins("XI0").Levels.Value(chVil) = TheExec.specs.DC("Pins_1p8v_Vil_VAR_H").CurrentValue
        Next site
    End If
       
    ' 20160105: Steph added for Refbuf test (Autogen) --- start
''    Call CUS_VFI_MeasureVolt(CUS_Str_MainProgram, MeasureValue, TestSeqNum, Pat)
    ' 20160105: Steph added for Refbuf test (Autogen) --- end
    
End Function
Public Function IO_HardIP_PPMU_Measure_I(TestPinArrayIV() As String, TestSeqNum As Integer, TestSeqNumIdx As Long, ForceSequenceArray() As String, _
                        k As Long, Pat As Variant, Flag_SingleLimit As Boolean, HighLimitVal As Double, LowLimitVal As Double, TestLimitPerPin_VIR As String, TestIrange() As String, _
                        FlowTestNme() As String, _
                        Optional CUS_Str_MainProgram As String, Optional SpecialCalcValSetting As CalculateMethodSetup = 0, _
                        Optional ByRef Rtn_MeasCurr As PinListData, Optional ByRef Rtn_SweepTestName As String, _
                        Optional InstSpecialSetting As InstrumentSpecialSetup = 0, Optional OutputTname As String, Optional WaitTime_I As String) As Long

    Dim MeasureValue As New PinListData
    Dim Force_idx As Integer
    Dim site As Variant
    Dim TestNum As Long
    Dim Pin  As Variant
    Dim p As Long
    Dim ForceV  As Double
    Dim OutputTname_format() As String
    Dim TempMeasVal_PerPin(100) As New PinListData
    ''=========================================================================================================
    '' 20160108 - Add rule to cover force value is different
    Dim ForceByPin() As String
    Dim ForceValByPin() As String
''    Dim ForceValIdx As Integer
    Dim Measure_I_Range() As String
    Dim MeasurePin As String
    Dim MI_Range_Index As Long
    Dim i As Long
    Dim Temp_index As Long
    
    MI_Range_Index = 0
    
    '' Force Pin
    If UBound(TestPinArrayIV) = 0 Then
        ForceByPin = Split(TestPinArrayIV(0), ",")
        MeasurePin = TestPinArrayIV(0)
        TheHdw.Digital.Pins(TestPinArrayIV(0)).Disconnect
    Else
        ForceByPin = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        MeasurePin = TestPinArrayIV(TestSeqNumIdx)
        TheHdw.Digital.Pins(TestPinArrayIV(TestSeqNumIdx)).Disconnect
    End If
    
    '' Force Volt value
    If UBound(ForceSequenceArray) = 0 Then
        ForceValByPin = Split(ForceSequenceArray(0), ",")
    Else
        ForceValByPin = Split(ForceSequenceArray(TestSeqNumIdx), ",")
    End If
    
    '' Measure Current range
    If UBound(TestIrange) = 0 Then
        Measure_I_Range = Split(TestIrange(0), ",")
        Dim MeasurePinArry() As String
        MeasurePinArry = Split(MeasurePin, ",")    'expend I range for all point
        If UBound(Measure_I_Range) = 0 Then
            ReDim Preserve Measure_I_Range(UBound(MeasurePinArry))
''            ReDim Measure_I_Range(UBound(MeasurePinArry))
            For i = 0 To UBound(MeasurePinArry)
                Measure_I_Range(i) = Measure_I_Range(0)
            Next i
        End If
    Else
        Measure_I_Range = Split(TestIrange(TestSeqNumIdx), ",")

        If (InStr(TestIrange(TestSeqNumIdx), ":") <> 0) Then
            'Add 20180104 Roger, current range for sweep condition

            For p = 0 To UBound(Measure_I_Range)
                If (InStr(Measure_I_Range(p), ":") <> 0) Then
                    If (UBound(Split(Measure_I_Range(p), ":")) >= (k - 1)) Then
                        Measure_I_Range(p) = Split(Measure_I_Range(p))(k - 1)
                    Else
                        Measure_I_Range(p) = Split(Measure_I_Range(p))(0)
    End If
                End If
            Next p
        End If
    End If
    
    ''=========================================================================================================
    '' 20150108 - Check number whether differrent between measure current range and force pin, add defalut value to let input number are the same.
    Call VIR_CheckTestCondition_Measure_I_R_Z("I", ForceByPin, Measure_I_Range)
     '' 20150111 - Check force value is the same or different
   ' Dim i As Long
    Dim b_ForceDiffVolt As Boolean
    Dim PastVal As Double
    b_ForceDiffVolt = False
    For i = 0 To UBound(ForceValByPin)
        If (InStr(ForceValByPin(i), ":") <> 0) Then
            If (UBound(Split(ForceValByPin(i), ":")) < k - 1 Or UBound(Split(ForceValByPin(i), ":")) = 0) Then
                ForceValByPin(i) = Split(ForceValByPin(i), ":")(0)
            Else
                ForceValByPin(i) = Split(ForceValByPin(i), ":")(k - 1)
            End If
        End If
        If i <> 0 Then
            If ForceValByPin(i) <> PastVal Then
                b_ForceDiffVolt = True
                Exit For
            End If
        End If
        PastVal = ForceValByPin(i)
    Next i
    ''=========================================================================================================
    ''20170126-Add serial measure method
    If InstSpecialSetting = InstrumentSpecialSetup.PPMU_SerialMeasurement Then
        Call PPMU_SerialMeasureCurr(ForceByPin(), ForceValByPin(), Measure_I_Range(), MeasureValue, b_ForceDiffVolt)
    Else
        For Each Pin In ForceByPin
        
            With TheHdw.PPMU.Pins(Pin)
                If InStr(CUS_Str_MainProgram, "PCIE_Init0") <> 0 Then
                    .ForceI 0, 0
                    '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
                Else
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                End If
                .Connect
                .Gate = tlOn
                If b_ForceDiffVolt = False Then
                    .ForceV ForceValByPin(0), Measure_I_Range(MI_Range_Index)
                Else
                    .ForceV ForceValByPin(MI_Range_Index), Measure_I_Range(MI_Range_Index)
                End If
                '' 20160108 - Only keep 1 force value but current range can be different for force pin
            End With
            
    ''        ForceValIdx = ForceValIdx + 1
            MI_Range_Index = MI_Range_Index + 1
'            TheExec.Datalog.WriteComment "Pin = " & (Pin & " Measure Current Range = " & TheHdw.PPMU.Pins(Pin).MeasureCurrentRange)
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & Pin & " =" & TheHdw.PPMU.Pins(Pin).MeasureCurrentRange)
            
        Next Pin
        
        If WaitTime_I = "" Then
            TheHdw.Wait (100 * us)
        Else
            TheHdw.Wait CDbl(WaitTime_I)
        End If
        DebugPrintFunc_PPMU CStr(MeasurePin)
        MeasureValue = TheHdw.PPMU.Pins(MeasurePin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    End If
    
    ''20161011 - Return MeasI_ToTestLimit to main program
    Rtn_MeasCurr = MeasureValue
    
    ''20150721 - DCWP used, store current value to apply next item, force the stored current to measure volt.
    If SpecialCalcValSetting = CalculateMethodSetup.PPMU_STORE_I Then
        Stored_MeasI_PPMU = MeasureValue
    End If
    
    ''CUS_Str for abs valuse by SP
    If CUS_Str_MainProgram <> "" And UCase(CUS_Str_MainProgram) Like "*MEAS_I_ABS*" Then
        Call MEAS_I_ABS(MeasureValue)
    End If
    
    Dim TestNameInput As String
        

    TestNameInput = "Curr_meas_" + CStr(TestSeqNum)
    If Rtn_SweepTestName <> "" Then
        TestNameInput = TestNameInput & "_" & Rtn_SweepTestName
    End If
    
    ''20151103 print force condition
    Call Print_Force_Condition("i", MeasureValue)
    Temp_index = TheExec.Flow.TestLimitIndex
    If SpecialCalcValSetting = PPMU_TestLimit_TTR Then
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "TTR:no report!"
    ElseIf CUS_Str_MainProgram = "TTR" Then
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
    Else
        For p = 0 To MeasureValue.Pins.Count - 1
            If Flag_SingleLimit = True Then TheExec.Flow.TestLimitIndex = Temp_index
            TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(p), , TestSeqNum, p)
            If b_ForceDiffVolt = False Then
                TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, Unit:=unitAmp, Tname:=TestNameInput, ForceVal:=ForceValByPin(0), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            Else
                TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, Unit:=unitAmp, Tname:=TestNameInput, ForceVal:=FormatNumber(TheHdw.PPMU(MeasureValue.Pins(p)).Voltage.Value, 3), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            End If
        Next p
    End If
End Function

Public Function IO_HardIP_PPMU_Measure_R(TestPinArrayIV() As String, TestSeqNum As Integer, TestSeqNumIdx As Long, ForceSequenceArray() As String, _
                        k As Long, Pat As Variant, Flag_SingleLimit As Boolean, HighLimitVal As Double, LowLimitVal As Double, TestLimitPerPin_VIR As String, _
                        TestIrange() As String, _
                        FlowTestNme() As String, _
                        Optional RAK_Flag As Enum_RAK = 0, _
                        Optional Rtn_SweepTestName As String, _
                        Optional CUS_Str_MainProgram As String, Optional OutputTname As String, Optional WaitTime_R As String, Optional SpecialCalcValSetting As CalculateMethodSetup = 0) As Long

    Dim MeasureValue As New PinListData
    Dim Force_idx As Integer
    Dim site As Variant
    Dim TestNum As Long
    
    Dim Imped As New PinListData
    Dim Pin  As Variant
    Dim RAK_Pin As String
    Dim GetRakVal As Double
    Dim p As Long
    Dim ForceVal_Volt  As Double

    Dim MeasCurr1 As New PinListData
    Dim MeasCurr2 As New PinListData

    '' 20160108 - Add rule to cover force value is different
    Dim ForceByPin() As String
    Dim ForceValByPin() As String
    Dim Measure_I_Range() As String
    Dim MeasurePin As String
    Dim MI_Range_Index As Long
    
    Dim OutputTname_format() As String
    MI_Range_Index = 0
    Dim Temp_index As Long
    '' Force Pin
    If UBound(TestPinArrayIV) = 0 Then
        ForceByPin = Split(TestPinArrayIV(0), ",")
        MeasurePin = TestPinArrayIV(0)
        TheHdw.Digital.Pins(TestPinArrayIV(0)).Disconnect
    Else
        ForceByPin = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        MeasurePin = TestPinArrayIV(TestSeqNumIdx)
        TheHdw.Digital.Pins(TestPinArrayIV(TestSeqNumIdx)).Disconnect
    End If
    
    '' Force Volt value
    If UBound(ForceSequenceArray) = 0 Then
        ForceValByPin = Split(ForceSequenceArray(0), ",")
    Else
        ForceValByPin = Split(ForceSequenceArray(TestSeqNumIdx), ",")
    End If
    
    '' Measure Current range
    If UBound(TestIrange) = 0 Then
        Measure_I_Range = Split(TestIrange(0), ",")
    Else
        Measure_I_Range = Split(TestIrange(TestSeqNumIdx), ",")
    End If
    
    ''=========================================================================================================
    '' 20150108 - Check number whether differrent between measure current range and force pin, add defalut value to let input number are the same.
    Call VIR_CheckTestCondition_Measure_I_R_Z("R", ForceByPin, Measure_I_Range)
    
    '' 20150111 - Check force value is the same or different
    Dim i As Long
    Dim b_ForceDiffVolt As Boolean
    Dim PastVal As Double
    b_ForceDiffVolt = False
    For i = 0 To UBound(ForceValByPin)
        If i <> 0 Then
            If ForceValByPin(i) <> PastVal Then
            b_ForceDiffVolt = True
            Exit For
            End If
        End If
        PastVal = ForceValByPin(i)
    Next i
    ''=========================================================================================================
    
    For Each Pin In ForceByPin
    
        With TheHdw.PPMU.Pins(Pin)
            '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
''            .ForceI 0, 0
            .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
            .Connect
            .Gate = tlOn
            If b_ForceDiffVolt = False Then
                .ForceV ForceValByPin(0), Measure_I_Range(MI_Range_Index)
            Else
                .ForceV ForceValByPin(MI_Range_Index), Measure_I_Range(MI_Range_Index)
            End If
            '' 20160108 - Only keep 1 force value but current range can be different for force pin
''            If UBound(ForceValByPin) = 0 Then
''                .ForceV ForceValByPin(0), Measure_I_Range(0)
''                ForceVal_Volt = ForceValByPin(0)
''            ElseIf ForceValByPin(ForceValIdx) <> "" Then
''                .ForceV ForceValByPin(ForceValIdx), Measure_I_Range(ForceValIdx)
''                 ForceVal_Volt = ForceValByPin(ForceValIdx)
''            Else:
''                .ForceV 0
''                ForceVal_Volt = 0
''            End If
''''           .Connect
''''           .Gate = tlOn
        End With
        
''        ForceValIdx = ForceValIdx + 1
        MI_Range_Index = MI_Range_Index + 1
'        TheExec.Datalog.WriteComment "Pin = " & (Pin & " Measure Current Range = " & TheHdw.PPMU.Pins(Pin).MeasureCurrentRange)
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & Pin & " =" & TheHdw.PPMU.Pins(Pin).MeasureCurrentRange)
        
    Next Pin
    
    If UBound(ForceSequenceArray) <> 0 Then
        If ForceSequenceArray(TestSeqNum) = "" Then
            ForceSequenceArray(TestSeqNum) = 0
        End If
    End If
    
    For Each site In TheExec.sites.Active
        TestNum = TheExec.sites.Item(site).TestNumber
    Next site
    
    If WaitTime_R = "" Then
        TheHdw.Wait 0.001
    Else
        TheHdw.Wait CDbl(WaitTime_R)
    End If
    
''    MeasureValue = TheHdw.PPMU.Pins(TestPinArrayIV(TestSeqNumIdx)).Read(tlPPMUReadMeasurements, 10)
    DebugPrintFunc_PPMU CStr(MeasurePin)
''    MeasureValue = TheHdw.PPMU.Pins(MeasurePin).Read(tlPPMUReadMeasurements, 10)
    MeasureValue = TheHdw.PPMU.Pins(MeasurePin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For Each Pin In MeasureValue.Pins
                If MeasureValue.Pins(Pin).Value(site) = 0 Then
                    MeasureValue.Pins(Pin).Value(site) = 1
                End If
            Next Pin
        Next site
    End If
    
    For Each Pin In MeasureValue.Pins
        For Each site In TheExec.sites
        If MeasureValue.Pins(Pin).Value(site) = 0 Then MeasureValue.Pins(Pin).Value(site) = 0.000000000001
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site :" & site & ",Pin : " & Pin & ",Measure current : " & MeasureValue.Pins(Pin).Value(site)
        Next site
    Next Pin
 
    ''20151103 print force condition
    Call Print_Force_Condition("r", MeasureValue)
 
    '' 20160111- Impedence measurement, force the same or different voltage to do calculation
    If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
       ' If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
            Call AnalyzeCusStrToCalcR(CUS_Str_MainProgram, TestSeqNum, ForceSequenceArray, MeasureValue, Imped)
      '  End If
    Else
        If b_ForceDiffVolt = False Then
            Imped = MeasureValue.Math.Invert.Multiply(ForceValByPin(0)).Abs
        Else
            Call MeasureR_ForceDifferentVolt(MeasureValue, Imped)
        End If
    End If
    
    '' Compensate resistance after Kelvin for path resistance considerations
    If RAK_Flag = R_TraceOnly Then
        'Dim RakV() As Double
''        Call AddRXRAkPinValue
        For Each Pin In Imped.Pins
            For Each site In TheExec.sites
                RAK_Pin = CStr(Pin)
                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(RAK_Pin, Site)
''                If InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
''                  GetRakVal = RakV(0) + FT_Card_RAK.Pins(Pin).Value(Site)
''                Else
''                    GetRakVal = RakV(0) + CP_Card_RAK.Pins(Pin).Value(Site)
''                End If
                GetRakVal = CurrentJob_Card_RAK.Pins(Pin).Value(site)
                
                If SpecialCalcValSetting <> PPMU_TestLimit_TTR And gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment Pin & " = " & Imped.Pins.Item(Pin).Value(site) & ", RAK val = " & CStr(GetRakVal)
                Imped.Pins.Item(Pin).Value(site) = Imped.Pins.Item(Pin).Value(site) - GetRakVal
            Next site
        Next Pin
    ElseIf RAK_Flag = R_PathWithContact Then
       For Each Pin In Imped.Pins
            For Each site In TheExec.sites
                If SpecialCalcValSetting <> PPMU_TestLimit_TTR And gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment Pin & " = " & Imped.Pins.Item(Pin).Value(site) & ", R_Path val = " & R_Path_PLD.Pins(Pin).Value(site)
                Imped.Pins.Item(Pin).Value(site) = Imped.Pins.Item(Pin).Value(site) - R_Path_PLD.Pins(Pin).Value(site)
            Next site
        Next Pin
    End If

    Dim TestNameInput As String

    'If TPModeAsCharz_GLB = True Then
    '    TestNameInput = FlowTestNme(gl_CZ_FlowTestName_Counter)
    '    'gl_CZ_FlowTestName_Counter = gl_CZ_FlowTestName_Counter + 1
    'Else
        TestNameInput = "Imp_meas_" + CStr(TestSeqNum)
        If Rtn_SweepTestName <> "" Then
            TestNameInput = TestNameInput & "_" & Rtn_SweepTestName
        End If
    'End If
    
    ''20160112 - Force value index for test limit if force voltage value is different
    Dim ForceVal_Index As Long
    ForceVal_Index = 0
    Temp_index = TheExec.Flow.TestLimitIndex
    
    If SpecialCalcValSetting = PPMU_TestLimit_TTR Then
        Dim Lowlimitval_temp As Double
        Dim Hilimitval_temp As Double
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        Lowlimitval_temp = GetLowLimitFromFlow
        Hilimitval_temp = GetHiLimitFromFlow
        If TheExec.EnableWord("HIP_TTR_FailResultOnly") = True Then
            For Each site In TheExec.sites.Active
                For p = 0 To Imped.Pins.Count - 1
                    If Imped.Pins(p).Value > Hilimitval_temp Or Imped.Pins(p).Value < Lowlimitval_temp Then
                        
                        If RAK_Flag = R_TraceOnly And gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment Imped.Pins(p) & " = " & Imped.Pins(p).Value & ", RAK val = " & GetRakVal
                        ElseIf RAK_Flag = R_PathWithContact And gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment Imped.Pins(p) & " = " & Imped.Pins(p).Value & ", R_Path val = " & R_Path_PLD.Pins(p).Value
                        End If
                        
                        TestNameInput = Report_TName_From_Instance("R", Imped.Pins(p), , TestSeqNum, p)
                        TheExec.Flow.TestLimit Imped.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitAmp, , Tname:=TestNameInput, ForceResults:=tlForceNone
                    End If
                Next p
            Next site
        Else
            For p = 0 To Imped.Pins.Count - 1
                'If Imped.Pins(p).Value > Hilimitval_temp Or Imped.Pins(p).Value < Lowlimitval_temp Then
                    
                    If RAK_Flag = R_TraceOnly And gl_Disable_HIP_debug_log = False Then
                        TheExec.Datalog.WriteComment Imped.Pins(p) & " = " & Imped.Pins(p).Value & ", RAK val = " & GetRakVal
                    ElseIf RAK_Flag = R_PathWithContact And gl_Disable_HIP_debug_log = False Then
                        TheExec.Datalog.WriteComment Imped.Pins(p) & " = " & Imped.Pins(p).Value & ", R_Path val = " & R_Path_PLD.Pins(p).Value
                    End If
                    
                    TestNameInput = Report_TName_From_Instance("R", Imped.Pins(p), , TestSeqNum, p)
                    TheExec.Flow.TestLimit Imped.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceNone
                'End If
            Next p
        End If
    Else
        
        For p = 0 To Imped.Pins.Count - 1
            If Flag_SingleLimit = True Then TheExec.Flow.TestLimitIndex = Temp_index
            TestNameInput = Report_TName_From_Instance("R", Imped.Pins(p), , TestSeqNum, p)
            If b_ForceDiffVolt = False Then
                TheExec.Flow.TestLimit Imped.Pins(p), , , , , , unitCustom, , TestNameInput, , , ForceValByPin(0), unitVolt, " ohm", , ForceResults:=tlForceFlow
            Else
                TheExec.Flow.TestLimit Imped.Pins(p), , , , , , unitCustom, , TestNameInput, , , FormatNumber(TheHdw.PPMU(MeasureValue.Pins(p)).Voltage.Value, 3), unitVolt, " ohm", , ForceResults:=tlForceFlow
            End If
        Next p
    End If

End Function

Public Function HardIP_SetupAndMeasureR() As Long

    Dim i As Long
    Dim Pins() As Variant
    Dim Pin As Variant
    Dim MeasR As Meas_Type
    Dim measureCurrent As New PinListData
    Dim Imped As New PinListData
    Dim DicStoreName As String
    Dim Temp_index As Long
    Dim site As Variant
    Dim p As Long
    Dim GetRakVal As Double
    Dim TestNameInput As String
    Dim GetRakVal_PinList As New PinListData
    
    MeasR = TestConditionSeqData(Instance_Data.TestSeqNum).MeasR(Instance_Data.TestSeqSweepNum)
    If Meas_StoreName_Flag Then DicStoreName = TestConditionSeqData(Instance_Data.TestSeqNum).Meas_StoreDicName(Instance_Data.TestSeqSweepNum)

    TheHdw.Digital.Pins(MeasR.Pins.PPMU).Disconnect
    If MeasR.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.PPMU.Pins(MeasR.Pins.PPMU)
            .Gate = tlOff
            .ForceI pc_Def_PPMU_InitialValue_FI
            .ForceV CDbl(MeasR.Setup_ByType.PPMU.ForceValue1), CDbl(MeasR.Setup_ByType.PPMU.Meas_Range)
            .Connect
            .Gate = tlOn
        End With
        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasR.Pins.PPMU & " =" & TheHdw.PPMU.Pins(SplitInputCondition(MeasR.Pins.PPMU, ",", 0)).MeasureCurrentRange)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasR.Pins.PPMU & " =" & MeasR.Setup_ByType.PPMU.Meas_Range)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasR.Pins.PPMU & " =" & MeasR.WaitTime.PPMU)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasR.Pins.PPMU & " =" & MeasR.Setup_ByType.PPMU.ForceValue1)
        End If
    Else
        For i = 0 To UBound(MeasR.Setup_ByTypeByPin.PPMU)
            With TheHdw.PPMU.Pins(MeasR.Setup_ByTypeByPin.PPMU(i).Pin)
                .Gate = tlOff
                .ForceI pc_Def_PPMU_InitialValue_FI
                .ForceV CDbl(MeasR.Setup_ByTypeByPin.PPMU(i).ForceValue1), CDbl(MeasR.Setup_ByTypeByPin.PPMU(i).Meas_Range)
                .Connect
                .Gate = tlOn
            End With
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasR.Setup_ByTypeByPin.PPMU(i).Pin & " =" & TheHdw.PPMU.Pins(MeasR.Setup_ByTypeByPin.PPMU(i).Pin).MeasureCurrentRange)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasR.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasR.Setup_ByTypeByPin.PPMU(i).Meas_Range)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasR.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasR.WaitTime.PPMU)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasR.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasR.Setup_ByTypeByPin.PPMU(i).ForceValue1)
            End If
        Next i
    End If
    
    TheHdw.Wait CDbl(MeasR.WaitTime.PPMU)
    
    DebugPrintFunc_PPMU CStr(MeasR.Pins.PPMU)
    measureCurrent = TheHdw.PPMU.Pins(MeasR.Pins.PPMU).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    If gl_Disable_HIP_debug_log = False Then Call ForceVal_Compare(MeasR, MeasR.Pins.PPMU, "R", "PPMU") ''Current1 Force Condition check - Carter, 20190521
    
    ''''---------Offline---------
    If TheExec.TesterMode = testModeOffline Then
        For Each Pin In measureCurrent.Pins
            measureCurrent.Pins(Pin).Value = measureCurrent.Pins(Pin).Add(0.0001)
        Next Pin
    End If
    ''''---------Offline---------
    
    If MeasR.Setup_ByTypeByPin_Flag = False Then
        Imped = measureCurrent.Math.Invert.Multiply(CDbl(MeasR.Setup_ByType.PPMU.ForceValue1)).Abs
    Else
        For i = 0 To UBound(MeasR.Setup_ByTypeByPin.PPMU)
            Imped.AddPin (MeasR.Setup_ByTypeByPin.PPMU(i).Pin)
            Imped.Pins(MeasR.Setup_ByTypeByPin.PPMU(i).Pin).Value = measureCurrent.Pins(MeasR.Setup_ByTypeByPin.PPMU(i).Pin).Invert.Multiply(CDbl(MeasR.Setup_ByTypeByPin.PPMU(i).ForceValue1)).Abs
        Next i
    End If

    With TheHdw.PPMU.Pins(MeasR.Pins.PPMU)
            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
            .Disconnect
            .Gate = tlOff
    End With

    TheHdw.Digital.Pins(MeasR.Pins.PPMU).Connect ''Carter, 20190521

    ''''---------Offline---------
    If TheExec.TesterMode = testModeOffline Then
        For Each Pin In measureCurrent.Pins
            measureCurrent.Pins(Pin).Value = measureCurrent.Pins(Pin).Add(0.0001)
        Next Pin
    End If
    ''''---------Offline---------
    If Instance_Data.RAK_Flag = R_TraceOnly Then
        For Each Pin In Imped.Pins
            GetRakVal_PinList.AddPin (CStr(Pin))
            GetRakVal_PinList.Pins(Pin) = CurrentJob_Card_RAK.Pins(Pin)
            If gl_Disable_HIP_debug_log = False Then
                For Each site In TheExec.sites.Active
                    TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " = " & Imped.Pins(Pin).Value(site) & ", RAK val = " & GetRakVal_PinList.Pins(Pin).Value(site)
                Next site
            End If
        Next Pin
        Imped = Imped.Math.Subtract(GetRakVal_PinList)
    
    ElseIf Instance_Data.RAK_Flag = R_PathWithContact Then
        For Each Pin In Imped.Pins
            GetRakVal_PinList.AddPin (CStr(Pin))
            GetRakVal_PinList.Pins(Pin) = R_Path_PLD.Pins(Pin)
            If gl_Disable_HIP_debug_log = False Then
                For Each site In TheExec.sites
                    TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " = " & Imped.Pins.Item(Pin).Value(site) & ", R_Path val = " & GetRakVal_PinList.Pins(Pin).Value(site) & ",Current = " & measureCurrent.Pins(Pin).Value(site)
                Next site
            End If
        Next Pin
        Imped = Imped.Math.Subtract(GetRakVal_PinList)
    
    End If

    '---------------------------------------------------------------------------------------------------------------
    Temp_index = TheExec.Flow.TestLimitIndex
    
    If Instance_Data.SpecialCalcValSetting = PPMU_TestLimit_TTR Then
        Dim Lowlimitval_temp As Double
        Dim Hilimitval_temp As Double
        
        If TheExec.EnableWord("HIP_TTR_FailResultOnly") = True Then
            
            For Each site In TheExec.sites.Active
                For p = 0 To Imped.Pins.Count - 1
                    TheExec.Flow.TestLimitIndex = Temp_index
                    If Imped.Pins(p).Value > CDbl(Instance_Data.HiLimit(Temp_index)) Or Imped.Pins(p).Value < CDbl(Instance_Data.LowLimit(Temp_index)) Then
                        
                        If Instance_Data.RAK_Flag = R_TraceOnly And gl_Disable_HIP_debug_log = False Then
                            
                            TheExec.Datalog.WriteComment "Site[" & site & "]," & Imped.Pins(p) & " = " & Imped.Pins(p).Value & ", RAK val = " & CStr(GetRakVal)
                        ElseIf Instance_Data.RAK_Flag = R_PathWithContact And gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment "Site[" & site & "]," & Imped.Pins(p) & " = " & Imped.Pins(p).Value & ", R_Path val = " & R_Path_PLD.Pins(p).Value
                        End If
                        
                        TestNameInput = Report_TName_From_Instance("R", Imped.Pins(p), , CInt(Instance_Data.TestSeqNum), p, , , , tlForceNone)
                        TheExec.Flow.TestLimit Imped.Pins(p), , , , , , unitCustom, , Tname:=TestNameInput, ForceResults:=tlForceFlow, customUnit:="ohm"
                    End If
                Next p
            Next site
        Else
            For p = 0 To Imped.Pins.Count - 1
                TheExec.Flow.TestLimitIndex = Temp_index
                If Instance_Data.RAK_Flag = R_TraceOnly And gl_Disable_HIP_debug_log = False Then
                    TheExec.Datalog.WriteComment Imped.Pins(p) & " = " & Imped.Pins(p).Value & ", RAK val = " & CStr(GetRakVal)
                ElseIf Instance_Data.RAK_Flag = R_PathWithContact And gl_Disable_HIP_debug_log = False Then
                    TheExec.Datalog.WriteComment Imped.Pins(p) & " = " & Imped.Pins(p).Value & ", R_Path val = " & R_Path_PLD.Pins(p).Value
                End If
                
                TestNameInput = Report_TName_From_Instance("R", Imped.Pins(p), , CInt(Instance_Data.TestSeqNum), p)
                TheExec.Flow.TestLimit Imped.Pins(p), , , , , , unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next p
        End If
    Else
    For Each site In TheExec.sites.Active
        For p = 0 To Imped.Pins.Count - 1
        If Instance_Data.RAK_Flag = R_TraceOnly And gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "Site[" & site & "]," & Imped.Pins(p) & " = " & Imped.Pins(p).Add(GetRakVal_PinList.Pins(p).Value).Value & ", RAK val = " & CStr(GetRakVal_PinList.Pins(p).Value)
        End If
    Next p
    Next site
        Call ProsscessTestLimit(Imped, "R", CInt(Instance_Data.TestSeqNum))
        
    End If
    '---------------------------------------------------------------------------------------------------------------
   
    ''Start ---- Carter, 20190521
    If Meas_StoreName_Flag Then
        If DicStoreName <> "" Then Call AddStoredMeasurement(DicStoreName, Imped)
    End If
    ''End ---- Carter, 20190521

End Function

Public Function Meas_VIR_IO_PreSetupBeforeMeasurement(TestPinArrayIV() As String, TestSeq As Long) As Long
    Dim TempStr As String
    Dim TempArr1() As String
    Dim TempArr2() As String
    Dim TempArr3() As String
    Dim TempStrPin() As String
    Dim index As Variant
    
     If UBound(TestPinArrayIV) = 0 Then
        If InStr(TestPinArrayIV(0), ":") > 0 Then
            TempStr = TestPinArrayIV(0)
            TempArr2 = Split(TestPinArrayIV(0), ",")
            TestPinArrayIV(0) = ""
            For Each index In TempArr2
                TempStrPin = Split(index, ":")
                TestPinArrayIV(0) = TestPinArrayIV(0) + "," + TempStrPin(0)
            Next index
        End If
     Else
        If (UBound(TestPinArrayIV) >= TestSeq) Then
            If InStr(TestPinArrayIV(TestSeq), ":") > 0 Then
                TempStr = TestPinArrayIV(TestSeq)
                TempArr2 = Split(TestPinArrayIV(TestSeq), ",")
                TestPinArrayIV(TestSeq) = ""
                For Each index In TempArr2
                    TempStrPin = Split(index, ":")
                    TestPinArrayIV(TestSeq) = TestPinArrayIV(TestSeq) + "," + TempStrPin(0)
                Next index
            End If
        End If
    End If

    If UBound(TestPinArrayIV) = 0 Then
        TheHdw.Digital.Pins(TestPinArrayIV(0)).Disconnect
        With TheHdw.PPMU.Pins(TestPinArrayIV(0))
            .Gate = tlOff
            .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_InitialValue_FV_Range
            .Connect
            .Gate = tlOn
        End With
    Else
        If (UBound(TestPinArrayIV) >= TestSeq) Then
            TheHdw.Digital.Pins(TestPinArrayIV(TestSeq)).Disconnect
            With TheHdw.PPMU.Pins(TestPinArrayIV(TestSeq))
                .Gate = tlOff
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_InitialValue_FV_Range
                .Connect
                 .Gate = tlOn
            End With
        End If
    End If
    
    If TempStr <> "" Then
        TestPinArrayIV(TestSeq) = TempStr
    End If
    
End Function
Public Function Meas_VIR_IO_PostSetupAfterMeasurement(TestPinArrayIV() As String, TestSeq As Long) As Long

    If UBound(TestPinArrayIV) = 0 Then

        With TheHdw.PPMU.Pins(TestPinArrayIV(0))
            .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_VIR_MeasCurrRange
            .Gate = tlOff
            .Disconnect
        End With

        TheHdw.Digital.Pins(TestPinArrayIV(0)).Connect

    Else

        With TheHdw.PPMU.Pins(TestPinArrayIV(TestSeq))
            .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_VIR_MeasCurrRange
            .Gate = tlOff
            .Disconnect
        End With

        TheHdw.Digital.Pins(TestPinArrayIV(TestSeq)).Connect
    End If

End Function

Public Function HardIP_MeasureCurrent()
    
    Dim index As Long
    Dim MeasCurrent(0 To 3) As New PinListData
    Dim MeasI_INstType_Num As Long
    Dim MeasI As Meas_Type
    Dim TestNameInput As String
    Dim site As Variant
    Dim p As Long
    Dim DicStoreName As String
    Dim Temp_LimitIndex As Long
    
    MeasI = TestConditionSeqData(Instance_Data.TestSeqNum).MeasI(Instance_Data.TestSeqSweepNum)
    If Meas_StoreName_Flag Then DicStoreName = TestConditionSeqData(Instance_Data.TestSeqNum).Meas_StoreDicName(Instance_Data.TestSeqSweepNum)

    index = 0
    
    If (MeasI.Pins.UVI80 <> "") Or (MeasI.Setup_ByTypeByPin.UVI80_Flag = True) Then
        Call HardIP_SetupAndMeasureCurrent_UVI80(10, MeasCurrent(index))
        index = index + 1
    End If
    
    If (MeasI.Pins.HexVS <> "") Or (MeasI.Setup_ByTypeByPin.HexVS_Flag = True) Then
        Call HardIP_SetupAndMeasureCurrent_HexVS(10, MeasCurrent(index))
        index = index + 1
    End If
    
    If (MeasI.Pins.UVS256 <> "") Or (MeasI.Setup_ByTypeByPin.UVS256_Flag = True) Then
        Call HardIP_SetupAndMeasureCurrent_UVS256(64, MeasCurrent(index))
        index = index + 1
    End If
    
    If (MeasI.Pins.PPMU <> "") Or (MeasI.Setup_ByTypeByPin.PPMU_Flag = True) Then
        If Instance_Data.InstSpecialSetting = PPMU_SerialMeasurement Then
            Call HardIP_SetupAndMeasureCurrent_PPMU_BySerial(10, MeasCurrent(index))
        Else
            Call HardIP_SetupAndMeasureCurrent_PPMU(10, MeasCurrent(index))
        End If
        index = index + 1
    End If

    If (MeasI.Pins.VSM <> "") Or (MeasI.Setup_ByTypeByPin.VSM_Flag = True) Then
        Call HardIP_SetupAndMeasureCurrent_VSM(10, MeasCurrent(index))
        index = index + 1
    End If
    ''''---------Offline---------
    If TheExec.TesterMode = testModeOffline Then
        Dim Pin As Variant
        Dim Dummy_Value As New SiteDouble: Dummy_Value = 0.0001
        For Each Pin In MeasCurrent(index - 1).Pins
            MeasCurrent(index - 1).Pins(Pin) = Dummy_Value
        Next Pin
    End If
    ''''---------Offline---------
    Dim MeasI_ToTestLimit As New PinListData
    If index = 1 Then
        Set MeasI_ToTestLimit = MeasCurrent(0)
    Else
        Call MergePinListData(CInt(index), MeasCurrent, MeasI_ToTestLimit)
    End If
    
    
    '-------[ 20190531  CT add for GPIO fuse]-------------------------------------------------------
    
    If (TheExec.Flow.EnableWord("HardIP_CZ") = False) Then
        
        If InStr(UCase(Instance_Data.CUS_Str_MainProgram), UCase("GPIO_Fuse")) <> 0 Then   'i.e.  CUS_Str_MainProgram = "GPIO_Fuse,SPI1_SCLK,gpio_iol_2,CFG,1"
        
            
            Dim Pin_SplitArray() As String
            Pin_SplitArray = Split(Instance_Data.CUS_Str_MainProgram, ",")
            
            Dim Pin_for_Fuse As String
            Dim Name_for_Fuse As String
            Dim Type_for_Fuse As String
            Dim TestSequence_for_Fuse As String
            Pin_for_Fuse = Pin_SplitArray(1)
            Name_for_Fuse = Pin_SplitArray(2)
            Type_for_Fuse = Pin_SplitArray(3)
            TestSequence_for_Fuse = Pin_SplitArray(4)
            
            Dim InstanceName_Str As String
            InstanceName_Str = TheExec.DataManager.instanceName
     
            '' Remove for RF doesn't have auto_eFuse_SetWriteDecimal function
            ''If Instance_Data.TestSeqNum = CLng(TestSequence_for_Fuse) Then
            '''If Instance_Data.TestSeqSweepNum = 1 Then
            ''
            ''    If UCase(InstanceName_Str) Like "*GPIO_T10SIDS2V1818_PP_TURA0_S_FULP_IO_GPIO_BSR_JIO_DIO_V1818_SI_SIDS2_T10_NV*" Then
            ''
            ''        theexec.Datalog.WriteComment "***The Value drop to FuseStructure***"
            ''        theexec.Datalog.WriteComment "   Pin: " & Pin_for_Fuse & ";  FuseName: " & Name_for_Fuse & ";  FuseType: " & Type_for_Fuse & ";  TestSequence: " & TestSequence_for_Fuse & ";"
            ''
            ''        For Each Site In theexec.sites
            ''            Call auto_eFuse_SetWriteDecimal(UCase(Type_for_Fuse), LCase(Name_for_Fuse), FormatNumber(2 * 1000 * MeasI_ToTestLimit.Pins(Pin_for_Fuse).Value(Site), 0), True)
            ''        Next Site
            ''    ElseIf UCase(InstanceName_Str) Like "*GPIO_T10SIDS4V1818_PP_TURA0_S_FULP_IO_GPIO_BSR_JIO_DIO_V1818_SI_SIDS4_T10_NV" Then
            ''
            ''        theexec.Datalog.WriteComment "***The Value drop to FuseStructure***"
            ''        theexec.Datalog.WriteComment "   Pin: " & Pin_for_Fuse & ";  FuseName: " & Name_for_Fuse & ";  FuseType: " & Type_for_Fuse & ";  TestSequence: " & TestSequence_for_Fuse & ";"
            ''
            ''        For Each Site In theexec.sites
            ''            Call auto_eFuse_SetWriteDecimal(UCase(Type_for_Fuse), LCase(Name_for_Fuse), FormatNumber(2 * 1000 * MeasI_ToTestLimit.Pins(Pin_for_Fuse).Value(Site), 0), True)
            ''        Next Site
            ''    End If
            ''
            ''End If
            
        End If
        
    End If

    
    '--------------------------------------------------------------
    
'-----------------------------------------------------------------------------------------------------------------------------------
    If Not ByPassTestLimit Then
        If Instance_Data.SpecialCalcValSetting = PPMU_TestLimit_TTR And MeasI.Pins.PPMU <> "" Then
            If TheExec.EnableWord("HIP_TTR_FailResultOnly") = True Then
                Temp_LimitIndex = TheExec.Flow.TestLimitIndex
                For Each site In TheExec.sites.Active
                    For p = 0 To MeasI_ToTestLimit.Pins.Count - 1
                        TheExec.Flow.TestLimitIndex = Temp_LimitIndex
                        If MeasI_ToTestLimit.Pins(p).Value > CDbl(Instance_Data.HiLimit(TheExec.Flow.TestLimitIndex)) Or MeasI_ToTestLimit.Pins(p).Value < CDbl(Instance_Data.LowLimit(TheExec.Flow.TestLimitIndex)) Then
                            TestNameInput = Report_TName_From_Instance("I", MeasI_ToTestLimit.Pins(p), , CInt(Instance_Data.TestSeqNum), p)
                            TheExec.Flow.TestLimit MeasI_ToTestLimit.Pins(p), CDbl(Instance_Data.LowLimit(TheExec.Flow.TestLimitIndex)), CDbl(Instance_Data.HiLimit(TheExec.Flow.TestLimitIndex)), _
                            , , , unitAmp, , Tname:=TestNameInput, ForceResults:=tlForceFlow, ForceVal:=MeasI.ForceValueDic_HWCom(UCase(MeasI_ToTestLimit.Pins(p)))
                        End If
                    Next p
                Next site
            Else
                Call ProsscessTestLimit(MeasI_ToTestLimit, "I", CInt(Instance_Data.TestSeqNum))
            End If
        Else
            Call ProsscessTestLimit(MeasI_ToTestLimit, "I", CInt(Instance_Data.TestSeqNum))
        End If
        
        If Instance_Data.SpecialCalcValSetting = DIFF_1ST Then
            G_MeasI_DIFF_1ST(Instance_Data.TestSeqNum) = MeasI_ToTestLimit
        End If
    
        If Instance_Data.SpecialCalcValSetting = DIFF_2ND Then
            MeasI_ToTestLimit = G_MeasI_DIFF_1ST(Instance_Data.TestSeqNum).Math.Subtract(MeasI_ToTestLimit)
            Call ProsscessTestLimit(MeasI_ToTestLimit, "I", CInt(Instance_Data.TestSeqNum), "Idiff")
        End If
    
        If Instance_Data.SpecialCalcValSetting = DIFF_PT12 And Instance_Data.TestSeqNum = 0 Then
            G_MeasI_DIFF_1ST(Instance_Data.TestSeqNum) = MeasI_ToTestLimit
        End If
    
        If Instance_Data.SpecialCalcValSetting = DIFF_PT12 And Instance_Data.TestSeqNum = 1 Then
            MeasI_ToTestLimit = MeasI_ToTestLimit.Math.Subtract(G_MeasI_DIFF_1ST(0))
            Call ProsscessTestLimit(MeasI_ToTestLimit, "I", CInt(Instance_Data.TestSeqNum), "Idiff", CInt(p))
        End If
    End If
       
    ''Start ---- Carter, 20190521
    If Meas_StoreName_Flag Then
        If DicStoreName <> "" Then Call AddStoredMeasurement(DicStoreName, MeasI_ToTestLimit)
    End If
    ''End ---- Carter, 20190521
    
End Function

Public Function PPMU_Reset(measureCase As String, meas As Meas_Type)
    Select Case measureCase
        Case "I"
            With TheHdw.PPMU.Pins(meas.Pins.PPMU)
                .ForceV pc_Def_PPMU_InitialValue_FV     ', pc_Def_PPMU_Max_InitialValue_FI_Range
                .Disconnect
                .Gate = tlOff
            End With
        Case "V"
            With TheHdw.PPMU.Pins(meas.Pins.PPMU)
                .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
                .Disconnect
                .Gate = tlOff
            End With
    End Select
    TheHdw.Digital.Pins(meas.Pins.PPMU).Connect
End Function


Public Function UVI80_Reset(measureCase As String, Pin As String)
    Select Case measureCase
        Case "I"
            With TheHdw.DCVI.Pins(Pin)
                .SetCurrentAndRange 0.2, 0.2
                .CurrentRange.Autorange = True
                .Voltage = 0
                .Gate = False
                .Disconnect
            End With
        Case "V"
           With TheHdw.DCVI.Pins(Pin)
                .SetCurrentAndRange 0.2, 0.2
                .current = 0
                .CurrentRange.Autorange = True
                .Voltage = 0
                .Gate(tlDCVIGateHiZ) = False
                .BleederResistor = tlDCVIBleederResistorAuto
                .Disconnect
                .mode = tlDCVIModeCurrent
            End With
    End Select
End Function



Public Function HardIP_SetupAndMeasureVolt_UVI80_old(MI_TestCond_UVI80() As DUTConditions, ByRef MeasureVolt As PinListData, Optional UVI80_MeasV_WaitTime As String = "") As Long
    ''Optional b_HighImpedenceMode As Boolean = True
    
    '' 20160419 - Debug Alarm off
     If TheExec.EnableWord("HardIP_Alarm_off") = True Then
        TheHdw.DCVI.Pins("analogmux_out").Alarm(tlDCVIAlarmAll) = tlAlarmOff
        TheHdw.DCVI.Pins("analogmux_out").Alarm(tlDCVIAlarmDGS) = tlAlarmOff
    End If
    ''20170526 - Add FI condition
    Dim i As Integer
    
    
    For i = 0 To UBound(MI_TestCond_UVI80)

        With TheHdw.DCVI.Pins(MI_TestCond_UVI80(i).PinName)
            .Gate = False
            
            '' High impedence mode
            If MI_TestCond_UVI80(i).FI_Val = 0 Then
                '' 20150612 - High impedence mode
                .Disconnect tlDCVIConnectDefault ' Only required if force was previously connected
                .mode = tlDCVIModeHighImpedance ' Program the DCVI mapped to MyPin to high impedance mode
                .Connect tlDCVIConnectHighSense ' Connect only the sense to use with high impedance mode
                .Meter.mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
                .BleederResistor = tlDCVIBleederResistorOff
                .current = 0
            Else
                .mode = tlDCVIModeCurrent
                .Connect tlDCVIConnectDefault
                .Voltage = pc_Def_VFI_UVI80_VoltCalmp
                .Meter.mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
                .CurrentRange.Autorange = True ''20170526-Add FI condition
                .current = MI_TestCond_UVI80(i).FI_Val
            End If
            .VoltageRange.Autorange = True
            .Gate = True
        End With
    Next i
    
    TheHdw.Wait (1 * ms)
    
    Dim MeasureV_Pin_UVI80 As String
    For i = 0 To UBound(MI_TestCond_UVI80)
        If i = 0 Then
            MeasureV_Pin_UVI80 = MI_TestCond_UVI80(i).PinName
        Else
            MeasureV_Pin_UVI80 = MeasureV_Pin_UVI80 & "," & MI_TestCond_UVI80(i).PinName
        End If
    Next i
    
    If UVI80_MeasV_WaitTime <> "" Then
        TheHdw.Wait (CDbl(UVI80_MeasV_WaitTime))
    End If
    
    MeasureVolt = TheHdw.DCVI.Pins(MeasureV_Pin_UVI80).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
    
    'For i = 0 To UBound(MI_TestCond_UVI80)
         '' 20150703 If use HiZ mode to measure volt that have to gate off HiZ and change to mode current
        'If MI_TestCond_UVI80(i).FI_Val = 0 Then
            With TheHdw.DCVI.Pins(MeasureV_Pin_UVI80)
                .current = 0
                .Voltage = 0
                .Gate(tlDCVIGateHiZ) = False
                .BleederResistor = tlDCVIBleederResistorAuto
                .Disconnect
                .mode = tlDCVIModeCurrent
            End With
        'End If
    'Next i
End Function

Public Function HardIP_SetupAndMeasureVolt_UVI80(ByRef MeasureVolt As PinListData) As Long

    Dim MeasV As Meas_Type
    Dim i As Long
    Dim VoltRangeList() As Double
    Dim CurrentRangeList() As Double
    Dim Pins() As String
    Dim Pin_Cnt As Long
    
    MeasV = TestConditionSeqData(Instance_Data.TestSeqNum).MeasV(Instance_Data.TestSeqSweepNum)
    'If TheHdw.DCVI.Pins(MeasV.Pins.UVI80).Gate <> True Then TheHdw.DCVI.Pins(MeasV.Pins.UVI80).Gate = False
    If TheHdw.DCVI.Pins(MeasV.Pins.UVI80).Gate <> True Then TheHdw.DCVI.Pins(MeasV.Pins.UVI80).Disconnect
    TheHdw.DCVI.Pins(MeasV.Pins.UVI80).VoltageRange.Autorange = True
    TheHdw.DCVI.Pins(MeasV.Pins.UVI80).CurrentRange.Autorange = True
    If MeasV.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.DCVI.Pins(MeasV.Pins.UVI80)
            If (CDbl(MeasV.Setup_ByType.UVI80.ForceValue1) = 0) And TheHdw.DCVI.Pins(MeasV.Pins.UVI80).Gate <> True Then
                .mode = tlDCVIModeHighImpedance
                .Voltage = pc_Def_VFI_UVI80_VoltCalmp
                .BleederResistor = tlDCVIBleederResistorOff
                .current = 0
                .Connect tlDCVIConnectHighSense
            Else
                .mode = tlDCVIModeCurrent
                .Voltage = pc_Def_VFI_UVI80_VoltCalmp
                .current = CDbl(MeasV.Setup_ByType.UVI80.ForceValue1)
                .Connect tlDCVIConnectDefault
            End If
            '/*** added by Kaino on 2019/06/19 for Turks : fixed Mode alarm ***/
            If CDbl(MeasV.Setup_ByType.UVI80.ForceValue1) <> 0 Then
                                'UCase(Instance_Data.CUS_Str_MainProgram) Like UCase("*DCVI_OFFHIZ_FIRST_AFTER_ForceI*") Then
                .Gate(tlDCVIGateHiZ) = False
            End If
            '/*** added by Kaino on 2019/06/19 for Turks : fixed Mode alarm ***/
            
            .Gate = True
            .Meter.mode = tlDCVIMeterVoltage
        End With
        
        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
            If (CDbl(MeasV.Setup_ByType.UVI80.ForceValue1) = 0) Then
                 TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> " & MeasV.Pins.UVI80 & " = High impedance mode  ")
            Else
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas Force Current value, " & MeasV.Pins.UVI80 & " =" & MeasV.Setup_ByType.UVI80.ForceValue1)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas UVI80 WaitTime, " & MeasV.Pins.UVI80 & " =" & MeasV.WaitTime.UVI80)
                End If
        End If
    Else
        For i = 0 To UBound(MeasV.Setup_ByTypeByPin.UVI80)
            With TheHdw.DCVI.Pins(MeasV.Setup_ByTypeByPin.UVI80(i).Pin)
                If (CDbl(MeasV.Setup_ByTypeByPin.UVI80(i).ForceValue1) = 0) And TheHdw.DCVI.Pins(MeasV.Setup_ByTypeByPin.UVI80(i).Pin).Gate <> True Then
                    .mode = tlDCVIModeHighImpedance
                    .Voltage = pc_Def_VFI_UVI80_VoltCalmp
                    .BleederResistor = tlDCVIBleederResistorOff
                    .current = 0
                    .Connect tlDCVIConnectHighSense
                Else
                    .mode = tlDCVIModeCurrent
                    .Voltage = pc_Def_VFI_UVI80_VoltCalmp
                    .current = CDbl(MeasV.Setup_ByTypeByPin.UVI80(i).ForceValue1)
                    .Connect tlDCVIConnectDefault
                End If
                .Gate = True
                .Meter.mode = tlDCVIMeterVoltage
            End With
            
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas Force Current value, " & MeasV.Setup_ByTypeByPin.UVI80(i).Pin & " =" & MeasV.Setup_ByTypeByPin.UVI80(i).ForceValue1)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas UVI80 WaitTime, " & MeasV.Setup_ByTypeByPin.UVI80(i).Pin & " =" & MeasV.WaitTime.UVI80)
            End If
        Next i
    End If
'    TheExec.Datalog.WriteComment "UVI80_MeasV_Wait_Time" & ":" & MeasV.WaitTime.UVI80
    If Not gl_Disable_HIP_debug_log Then: TheExec.Datalog.WriteComment "UVI80_MeasV_Wait_Time" & ":" & MeasV.WaitTime.UVI80
    TheHdw.Wait CDbl(MeasV.WaitTime.UVI80)
    MeasureVolt = TheHdw.DCVI.Pins(MeasV.Pins.UVI80).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
    
    If gl_Disable_HIP_debug_log = False Then Call ForceVal_Compare(MeasV, MeasV.Pins.UVI80, "V", "DCVI")
    
'    If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
'        TheExec.DataManager.DecomposePinList MeasV.Pins.UVI80, Pins, Pin_cnt
'        For i = 0 To Pin_cnt - 1
'            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> VoltageMeas AutoRange Irange setting, " & Pins(i) & " =" & TheHdw.DCVI.Pins(Pins(i)).CurrentRange)
'        Next i
'    End If
    
    With TheHdw.DCVI.Pins(MeasV.Pins.UVI80)
        '/*** added by Kaino on 2019/06/25 for Turks : fixed Mode alarm ***/
        If CDbl(MeasV.Setup_ByType.UVI80.ForceValue1) <> 0 Then
            TheExec.Datalog.WriteComment " *** UVI80 Force Conditon Rollback ***  " & "Gate: ON -> HiZOFF -> OFF"
            .Gate(tlDCVIGateHiZ) = False
            .Gate = False
        End If
        '/*** added by Kaino on 2019/06/25 for Turks : fixed Mode alarm ***/
        
        .current = 0
        .Voltage = 0
        .Gate(tlDCVIGateHiZ) = False
        .BleederResistor = tlDCVIBleederResistorAuto
        .Disconnect
        .mode = tlDCVIModeCurrent
    End With
End Function

Public Function HardIP_Characterization_ShmooEndFunction(argc As Long, argv() As String)
    gl_flag_end_shmoo = True
End Function


Public Function HardIP_SetupAndMeasureVolt_UVI80_Diff(ByRef MeasureVolt As PinListData) As Long
        
        Dim H_Pin As String
        Dim L_Pin As String
        
        H_Pin = TestConditionSeqData(Instance_Data.TestSeqNum).MeasV(Instance_Data.TestSeqSweepNum).Setup_ByType.UVI80.Pin
        L_Pin = TestConditionSeqData(Instance_Data.TestSeqNum).MeasV(Instance_Data.TestSeqSweepNum).Setup_ByType.UVI80.Pin_Diff_L
        
        Call UVI80_DIFFMETER_SETUP(H_Pin, L_Pin)
        
        If UCase(Instance_Data.CUS_Str_MainProgram) = "SPOTCAL_PERSTEP" Then
                        
                        ' perform spot cal for each DAC step
            Call UVI80_DCDIFFMETER_SPOTCAL(H_Pin, L_Pin, MeasureVolt)

        ElseIf UCase(Instance_Data.CUS_Str_MainProgram) = "SPOTCAL_PRETEST" Then
                        
                        ' must execute UVI80_DCDIFFMETER_SPOTCAL once before "SPOTCAL_PRETEST"
                Call UVI80_DCDIFFMETER(H_Pin, MeasureVolt)
                Set MeasureVolt = MeasureVolt.Math.Subtract(gCMError)
                        
        Else
                Call UVI80_DCDIFFMETER(H_Pin, MeasureVolt)
        End If
        
        If H_Pin <> "" Then Call UVI80_DIFFMETER_RELEASE(H_Pin, L_Pin)
        
End Function

Public Function HardIP_SetupAndMeasureVolt_PPMU(ByRef MeasureVolt As PinListData) As Long
    
    Dim i As Long
    Dim Pins() As Variant
    Dim MeasV As Meas_Type
    Dim measureCurrent As New PinListData
    MeasV = TestConditionSeqData(Instance_Data.TestSeqNum).MeasV(Instance_Data.TestSeqSweepNum)
    If Instance_Data.InstSpecialSetting = DigitalConnectPPMU2 Or Instance_Data.InstSpecialSetting = DigitalConnectPPMU Then ' = DigitalConnectPPMU Then
        TheHdw.PPMU.AllowPPMUFuncRelayConnection True, False
        TheHdw.PPMU.Pins(MeasV.Pins.PPMU).ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Digital_MaxCurrRange
        'thehdw.Digital.Pins(MeasV.Pins.PPMU).Connect
    Else
        TheHdw.Digital.Pins(MeasV.Pins.PPMU).Disconnect
    End If
    If MeasV.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.PPMU.Pins(MeasV.Pins.PPMU)
            .Gate = tlOff
                        If Instance_Data.InstSpecialSetting = PPMU_AccurateMeasurement Then
            .ForceI CDbl(MeasV.Setup_ByType.PPMU.ForceValue1), CDbl(MeasV.Setup_ByType.PPMU.ForceValue1)
            ElseIf Instance_Data.InstSpecialSetting = DigitalConnectPPMU Then
                .ForceI CDbl(MeasV.Setup_ByType.PPMU.ForceValue1), pc_Def_PPMU_FI_Range_200uA
            ElseIf Instance_Data.InstSpecialSetting = PPMU_2mA_Force_I_Range Then
                .ForceI CDbl(MeasV.Setup_ByType.PPMU.ForceValue1), 0.002
            ElseIf Instance_Data.InstSpecialSetting = PPMU_200uA_Force_I_Range Then
                .ForceI CDbl(MeasV.Setup_ByType.PPMU.ForceValue1), 0.0002
            ElseIf Instance_Data.InstSpecialSetting = PPMU_20uA_Force_I_Range Then
                .ForceI CDbl(MeasV.Setup_ByType.PPMU.ForceValue1), 0.00002
            Else
                .ForceI CDbl(MeasV.Setup_ByType.PPMU.ForceValue1), CDbl(MeasV.Setup_ByType.PPMU.Meas_Range)
            End If
            .Connect
            If Instance_Data.InstSpecialSetting <> DigitalConnectPPMU2 And Instance_Data.InstSpecialSetting <> DigitalConnectPPMU Then .Gate = tlOn
        End With
        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas Force Current value, " & MeasV.Pins.PPMU & " =" & MeasV.Setup_ByType.PPMU.ForceValue1)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas PPMU WaitTime, " & MeasV.Pins.PPMU & " =" & MeasV.WaitTime.PPMU)
                                If Instance_Data.InstSpecialSetting = PPMU_AccurateMeasurement Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByType.PPMU.Pin & " =" & CDbl(MeasV.Setup_ByType.PPMU.ForceValue1))
                ElseIf Instance_Data.InstSpecialSetting = DigitalConnectPPMU Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByType.PPMU.Pin & " =" & pc_Def_PPMU_FI_Range_200uA)
                ElseIf Instance_Data.InstSpecialSetting = PPMU_2mA_Force_I_Range Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByType.PPMU.Pin & " =" & "0.002")
                ElseIf Instance_Data.InstSpecialSetting = PPMU_200uA_Force_I_Range Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByType.PPMU.Pin & " =" & "0.0002")
                ElseIf Instance_Data.InstSpecialSetting = PPMU_20uA_Force_I_Range Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByType.PPMU.Pin & " =" & "0.00002")
                Else
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByType.PPMU.Pin & " =" & MeasV.Setup_ByType.PPMU.Meas_Range)
                End If
        End If
    Else
        For i = 0 To UBound(MeasV.Setup_ByTypeByPin.PPMU)
            With TheHdw.PPMU.Pins(MeasV.Setup_ByTypeByPin.PPMU(i).Pin)
                .Gate = tlOff
                If Instance_Data.InstSpecialSetting = PPMU_AccurateMeasurement Then
                                        .ForceI CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1), CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1)
                ElseIf Instance_Data.InstSpecialSetting = PPMU_2mA_Force_I_Range Then
                    .ForceI CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1), 0.002
                ElseIf Instance_Data.InstSpecialSetting = PPMU_200uA_Force_I_Range Then
                    .ForceI CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1), 0.0002
                ElseIf Instance_Data.InstSpecialSetting = PPMU_20uA_Force_I_Range Then
                    .ForceI CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1), 0.00002
                ElseIf Instance_Data.InstSpecialSetting = EUSB_T10T11_Split_Force_I_Range Then
                  If CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1) = 0 Then .ForceI CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1), 0.0002
                  If CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1) <> 0 Then .ForceI CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1), 0.05
                Else
                    .ForceI CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1), CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).Meas_Range)
                End If
                .Connect
                If Instance_Data.InstSpecialSetting <> DigitalConnectPPMU2 And Instance_Data.InstSpecialSetting <> DigitalConnectPPMU Then .Gate = tlOn
            End With
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas Force Current value, " & MeasV.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas PPMU WaitTime, " & MeasV.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasV.WaitTime.PPMU)
                                        If Instance_Data.InstSpecialSetting = PPMU_AccurateMeasurement Then
                        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByTypeByPin.PPMU(i).Pin & " =" & CStr(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1))
                        ElseIf Instance_Data.InstSpecialSetting = PPMU_2mA_Force_I_Range Then
                        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByTypeByPin.PPMU(i).Pin & " =" & "0.002")
                        ElseIf Instance_Data.InstSpecialSetting = PPMU_200uA_Force_I_Range Then
                        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByTypeByPin.PPMU(i).Pin & " =" & "0.0002")
                        ElseIf Instance_Data.InstSpecialSetting = PPMU_20uA_Force_I_Range Then
                        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByTypeByPin.PPMU(i).Pin & " =" & "0.00002")
                        ElseIf Instance_Data.InstSpecialSetting = EUSB_T10T11_Split_Force_I_Range Then
                                If CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1) = 0 Then TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByTypeByPin.PPMU(i).Pin & " =" & "0.0002")
                                If CDbl(MeasV.Setup_ByTypeByPin.PPMU(i).ForceValue1) <> 0 Then TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByTypeByPin.PPMU(i).Pin & " =" & "0.05")
                        Else
                        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Force I range setting, " & MeasV.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasV.Setup_ByTypeByPin.PPMU(i).Meas_Range)
                        End If
            End If
        Next i
    End If
    
    TheHdw.Wait CDbl(MeasV.WaitTime.PPMU)
    DebugPrintFunc_PPMU CStr(MeasV.Pins.PPMU)
    MeasureVolt = TheHdw.PPMU.Pins(MeasV.Pins.PPMU).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)

    If gl_Disable_HIP_debug_log = False Then Call ForceVal_Compare(MeasV, MeasV.Pins.PPMU, "V", "PPMU") ''Carter, 20190507
    If Instance_Data.InstSpecialSetting = DigitalConnectPPMU Then
        TheHdw.Digital.Pins(MeasV.Pins.PPMU).Disconnect
        TheHdw.PPMU.AllowPPMUFuncRelayConnection False, False
    End If
    
    With TheHdw.PPMU.Pins(MeasV.Pins.PPMU)
            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
            .Disconnect
            .Gate = tlOff
    End With
    
    'If Instance_Data.InstSpecialSetting = DigitalConnectPPMU Then
        'thehdw.Digital.Pins(MeasV.Pins.PPMU).Disconnect
    '    thehdw.PPMU.AllowPPMUFuncRelayConnection False, False
    'Else
        TheHdw.Wait (0.1 * ms)
        TheHdw.Digital.Pins(MeasV.Pins.PPMU).Connect
    'End If
    
    Dim Pin As Variant
    Dim GetRakVal_PinList As New PinListData
    Dim DiffVolt_Pinlist As New PinListData
    If Instance_Data.RAK_Flag = R_TraceOnly Then
        For Each Pin In MeasureVolt.Pins
            GetRakVal_PinList.AddPin (CStr(Pin))
            DiffVolt_Pinlist.AddPin (CStr(Pin))
            GetRakVal_PinList.Pins(Pin) = CurrentJob_Card_RAK.Pins(Pin)
            If MeasV.Setup_ByTypeByPin_Flag = False Then
                DiffVolt_Pinlist.Pins(Pin) = GetRakVal_PinList.Pins(Pin).Multiply(CDbl(MeasV.Setup_ByType.PPMU.ForceValue1))
            Else
                DiffVolt_Pinlist.Pins(Pin) = GetRakVal_PinList.Pins(Pin).Multiply(CDbl(MeasV.ForceValueDic_HWCom(Pin)))
            End If
            
            If gl_Disable_HIP_debug_log = False Then
                For Each site In TheExec.sites.Active
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Voltage = " & MeasureVolt.Pins(Pin).Value(site) & ", RAK val = " & GetRakVal_PinList.Pins(Pin).Value(site)
                Next site
            End If
            
        Next Pin
        MeasureVolt = MeasureVolt.Math.Subtract(DiffVolt_Pinlist)
        
    ElseIf Instance_Data.RAK_Flag = R_PathWithContact Then
        For Each Pin In MeasureVolt.Pins
            GetRakVal_PinList.AddPin (CStr(Pin))
            DiffVolt_Pinlist.AddPin (CStr(Pin))
            GetRakVal_PinList.Pins(Pin) = R_Path_PLD.Pins(Pin)
            If MeasV.Setup_ByTypeByPin_Flag = False Then
                DiffVolt_Pinlist.Pins(Pin) = GetRakVal_PinList.Pins(Pin).Multiply(CDbl(MeasV.Setup_ByType.PPMU.ForceValue1))
            Else
                DiffVolt_Pinlist.Pins(Pin) = GetRakVal_PinList.Pins(Pin).Multiply(CDbl(MeasV.ForceValueDic_HWCom(Pin)))
            End If
            
            If gl_Disable_HIP_debug_log = False Then
                For Each site In TheExec.sites.Active
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Voltage = " & MeasureVolt.Pins(Pin).Value(site) & ", R_Path val = " & GetRakVal_PinList.Pins(Pin).Value(site)
                Next site
            End If
    
        Next Pin
        MeasureVolt = MeasureVolt.Math.Subtract(DiffVolt_Pinlist)
        
    ElseIf InStr(UCase(Instance_Data.CUS_Str_MainProgram), UCase("RREF_RAK_CALC")) <> 0 Then
        Call CUS_RREF_Rak_Calc(MeasureVolt)
    End If

    Dim ForceValByPin() As String
    ForceValByPin = Split(MeasV.Setup_ByType.PPMU.ForceValue1, ",")
    Select Case Instance_Data.SpecialCalcValSetting
        Case CalculateMethodSetup.VIR_DDIO:
            Call CUS_DDR_Emulate_Const_Res_Loading(MeasureVolt, ForceValByPin(), Instance_Data.CUS_Str_MainProgram, CInt(Instance_Data.TestSeqNum), Instance_Data.RAK_Flag)
    End Select
End Function

Public Function HardIP_SetupAndMeasureCurrent_UVI80(SampleSize As Long, ByRef measureCurrent As PinListData)

    Dim i As Integer
    Dim WaitTime As Double
    Dim MaxWaitTime As Double
    Dim Pins() As String
    Dim Pin_Cnt As Long
    Dim PinType As String
    Dim factor As Long
    Dim Irange As Double
    
    Dim MeasI As Meas_Type
    MeasI = TestConditionSeqData(Instance_Data.TestSeqNum).MeasI(Instance_Data.TestSeqSweepNum)
    
    Call HardIP_DCVI_MI_StoreAndRestoreCondition(MeasI, UVI80, True)
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Pin As Variant
    Pins = Split(MeasI.Pins.UVI80, ",")
    For Each Pin In Pins
        If UCase(TheExec.DataManager.PinType(Pin)) = UCase("Power") Then
            If MeasI.ForceValueDic_HWCom.Exists(Pin) Then
                MeasI.ForceValueDic_HWCom.Remove (Pin)
            End If
            MeasI.ForceValueDic_HWCom.Add Pin, CStr(FormatNumber(TheHdw.DCVI.Pins(Pin).Voltage, 3))
        End If
    Next Pin
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If MeasI.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.DCVI.Pins(MeasI.Pins.UVI80)
            If sweep_power_val_per_loop_count <> "" Then
                'do not gate off
            ElseIf UCase(TheExec.DataManager.PinType(MeasI.Pins.UVI80)) <> UCase("Power") Then
                .Gate = False
                .mode = tlDCVIModeVoltage
                .Voltage = CDbl(MeasI.Setup_ByType.UVI80.ForceValue1)
            End If
            .VoltageRange.Autorange = True
            .current = pc_Def_UVI80_Init_MeasCurrRange
            .CurrentRange.Autorange = True
            .Connect tlDCVIConnectDefault
            If .Gate = False Then                   '-+
                .Gate(tlDCVIGateHiZ) = False        ' +---Added by Kaino on 2019/09/02 for Mode alarm
            End If                                  '-+
            .Gate = True
        End With
        
        TheHdw.Wait 0.001
        
        With TheHdw.DCVI.Pins(MeasI.Pins.UVI80)
            .Meter.mode = tlDCVIMeterCurrent
            .SetCurrentAndRange CDbl(MeasI.Setup_ByType.UVI80.Meas_Range), CDbl(MeasI.Setup_ByType.UVI80.Meas_Range)
            '.Meter.CurrentRange.Value = CDbl(MeasI.Setup_ByType.UVI80.Meas_Range)
            .CurrentRange.Autorange = True
        End With
        
        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
            If sweep_power_val_per_loop_count <> "" Then
                MeasI.Setup_ByType.UVI80.ForceValue1 = sweep_power_val_per_loop_count
                If UCase(TheExec.DataManager.PinType(MeasI.Pins.UVI80)) <> UCase("Power") Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasI.Pins.UVI80 & " =" & MeasI.Setup_ByType.UVI80.ForceValue1)
                Else
                    'TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Force Volt value, " & MeasI.Pins.UVI80 & " =" & FormatNumber(thehdw.DCVI.Pins(MeasI.Pins.UVI80).Voltage, 3))
                End If
                'TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Force Volt value, " & MeasI.Pins.UVI80 & " =" & MeasI.Setup_ByType.UVI80.ForceValue1)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range, " & MeasI.Pins.UVI80 & " = " & MeasI.Setup_ByType.UVI80.Meas_Range)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasI.Pins.UVI80 & " =" & MeasI.WaitTime.UVI80)
            Else
                If UCase(TheExec.DataManager.PinType(MeasI.Pins.UVI80)) <> UCase("Power") Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasI.Pins.UVI80 & " =" & MeasI.Setup_ByType.UVI80.ForceValue1)
                Else
                    'TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Force Volt value, " & MeasI.Pins.UVI80 & " =" & FormatNumber(thehdw.DCVI.Pins(MeasI.Pins.UVI80).Voltage, 3))
                End If
                'TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Force Volt value, " & MeasI.Pins.UVI80 & " =" & MeasI.Setup_ByType.UVI80.ForceValue1)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range, " & MeasI.Pins.UVI80 & " = " & MeasI.Setup_ByType.UVI80.Meas_Range)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasI.Pins.UVI80 & " =" & MeasI.WaitTime.UVI80)
            End If
                
        End If
    Else
        For i = 0 To UBound(MeasI.Setup_ByTypeByPin.UVI80)
            
            With TheHdw.DCVI.Pins(MeasI.Setup_ByTypeByPin.UVI80(i).Pin)
                If sweep_power_val_per_loop_count <> "" Then
                    'do not gate off
                ElseIf UCase(TheExec.DataManager.PinType(MeasI.Setup_ByTypeByPin.UVI80(i).Pin)) <> UCase("Power") Then
                    .Gate = False
                    .mode = tlDCVIModeVoltage
                .Voltage = CDbl(MeasI.Setup_ByTypeByPin.UVI80(i).ForceValue1)
                End If
                .VoltageRange.Autorange = True
                .current = pc_Def_UVI80_Init_MeasCurrRange
                .CurrentRange.Autorange = True
                .Connect tlDCVIConnectDefault
                .Gate = True
            End With
            
            TheHdw.Wait 0.001
            
            With TheHdw.DCVI.Pins(MeasI.Setup_ByTypeByPin.UVI80(i).Pin)
                .Meter.mode = tlDCVIMeterCurrent
                .SetCurrentAndRange CDbl(MeasI.Setup_ByTypeByPin.UVI80(i).Meas_Range), CDbl(MeasI.Setup_ByTypeByPin.UVI80(i).Meas_Range)
                .CurrentRange.Autorange = True
            End With
                
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                    If UCase(TheExec.DataManager.PinType(MeasI.Setup_ByTypeByPin.UVI80(i).Pin)) <> UCase("Power") Then ''20191217
                        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasI.Pins.UVI80 & " =" & MeasI.Setup_ByTypeByPin.UVI80(i).ForceValue1)
                    Else
                        'TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Force Volt value, " & MeasI.Pins.UVI80 & " =" & FormatNumber(thehdw.DCVI.Pins(MeasI.Pins.UVI80).Voltage, 3))
                    End If
                    'TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Force Volt value, " & MeasI.Setup_ByTypeByPin.UVI80(i).Pin & " =" & MeasI.Setup_ByTypeByPin.UVI80(i).ForceValue1)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range, " & MeasI.Setup_ByTypeByPin.UVI80(i).Pin & " = " & MeasI.Setup_ByTypeByPin.UVI80(i).Meas_Range)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasI.Setup_ByTypeByPin.UVI80(i).Pin & " =" & MeasI.WaitTime.UVI80)
            End If
        Next i
    End If
    TheHdw.Wait CDbl(MeasI.WaitTime.UVI80)
    measureCurrent = TheHdw.DCVI.Pins(MeasI.Pins.UVI80).Meter.Read(tlStrobe, SampleSize)
    
    If gl_Disable_HIP_debug_log = False Then Call ForceVal_Compare(MeasI, MeasI.Pins.UVI80, "I", "DCVI")
    
    Call HardIP_DCVI_MI_StoreAndRestoreCondition(MeasI, UVI80, False)
    
End Function


Public Function HardIP_SetupAndMeasureCurrent_HexVS(SampleSize As Long, ByRef measureCurrent As PinListData)

    '' 20150623 - Suggest use for single pin it can mapping expected current range,
    ''                  if use for pin group it will refer the same current range to pin group by your specified.
    Dim i As Integer

    Dim Pins_MeasureI_Together As String
    Dim MeasI As Meas_Type
    MeasI = TestConditionSeqData(Instance_Data.TestSeqNum).MeasI(Instance_Data.TestSeqSweepNum)
    Call HardIP_DCVS_MI_StoreAndRestoreCondition(MeasI, HexVS, True)
    If MeasI.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.DCVS.Pins(MeasI.Pins.HexVS)
            .Meter.mode = tlDCVSMeterCurrent
            If MeasI.Setup_ByType.HexVS.Meas_Range <> "0" Then .SetCurrentRanges CDbl(MeasI.Setup_ByType.HexVS.Meas_Range), CDbl(MeasI.Setup_ByType.HexVS.Meas_Range)
            .Gate = True
        End With

        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasI.Pins.HexVS & " =" & TheHdw.DCVS.Pins(SplitInputCondition(MeasI.Pins.HexVS, ",", 0)).Meter.CurrentRange.Value)
            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasI.Pins.HexVS & " =" & MeasI.Setup_ByType.HexVS.Meas_Range)
            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasI.Pins.HexVS & " =" & MeasI.WaitTime.HexVS)
        End If
    Else
        For i = 0 To UBound(MeasI.Setup_ByTypeByPin.HexVS)
            With TheHdw.DCVS.Pins(MeasI.Setup_ByTypeByPin.HexVS(i).Pin)
                .Meter.mode = tlDCVSMeterCurrent
                If MeasI.Setup_ByTypeByPin.HexVS(i).Meas_Range <> "0" Then .SetCurrentRanges CDbl(MeasI.Setup_ByTypeByPin.HexVS(i).Meas_Range), CDbl(MeasI.Setup_ByTypeByPin.HexVS(i).Meas_Range)
                .Gate = True
            End With
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasI.Setup_ByTypeByPin.HexVS(i).Pin & " =" & TheHdw.DCVS.Pins(MeasI.Setup_ByTypeByPin.HexVS(i).Pin).Meter.CurrentRange.Value)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasI.Setup_ByTypeByPin.HexVS(i).Pin & " =" & MeasI.Setup_ByTypeByPin.HexVS(i).Meas_Range)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range WaitTime, " & MeasI.Setup_ByTypeByPin.HexVS(i).Pin & " =" & MeasI.WaitTime.HexVS)
            End If
        Next i
    End If
    TheHdw.Wait CDbl(MeasI.WaitTime.HexVS)
    measureCurrent = TheHdw.DCVS.Pins(MeasI.Pins.HexVS).Meter.Read(tlStrobe, SampleSize, pc_Def_HexVS_ReadPoint, tlDCVSMeterReadingFormatAverage)
    
    '''--------->Save the interpose PrePat and interpose PreMeas force condition---------
    Dim num_pins As Long
    Dim instr_pins() As String
    Dim DCVS_HW_Value As String
    Call TheExec.DataManager.DecomposePinList(MeasI.Pins.HexVS, instr_pins(), num_pins)
    For i = 0 To num_pins - 1
        DCVS_HW_Value = CStr(FormatNumber(TheHdw.DCVS.Pins(instr_pins(i)).Voltage.Value, 3))
        If MeasI.ForceValueDic_HWCom.Exists(UCase(instr_pins(i))) Then
            MeasI.ForceValueDic_HWCom(UCase(instr_pins(i))) = DCVS_HW_Value
        Else
            MeasI.ForceValueDic_HWCom.Add UCase(instr_pins(i)), DCVS_HW_Value
        End If
    Next i
    '''--------->Save the interpose PrePat and interpose PreMeas force condition---------
    
    If gl_Disable_HIP_debug_log = False Then Call ForceVal_Compare(MeasI, MeasI.Pins.HexVS, "I", "DCVS") ''Carter, 20190624
    
    Call HardIP_DCVS_MI_StoreAndRestoreCondition(MeasI, HexVS, False)

End Function

Public Function HardIP_SetupAndMeasureCurrent_VSM(SampleSize As Long, ByRef measureCurrent As PinListData)

    '' 20150623 - Suggest use for single pin it can mapping expected current range,
    ''                  if use for pin group it will refer the same current range to pin group by your specified.
    Dim i As Integer

    Dim Pins_MeasureI_Together As String
    Dim MeasI As Meas_Type
    MeasI = TestConditionSeqData(Instance_Data.TestSeqNum).MeasI(Instance_Data.TestSeqSweepNum)
    Call HardIP_DCVS_MI_StoreAndRestoreCondition(MeasI, VSM, True)
    If MeasI.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.DCVS.Pins(MeasI.Pins.VSM)
            .Meter.mode = tlDCVSMeterCurrent
            If MeasI.Setup_ByType.VSM.Meas_Range <> "0" Then .SetCurrentRanges CDbl(MeasI.Setup_ByType.VSM.Meas_Range), CDbl(MeasI.Setup_ByType.VSM.Meas_Range)
            .Gate = True
        End With

        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasI.Pins.VSM & " =" & TheHdw.DCVS.Pins(SplitInputCondition(MeasI.Pins.VSM, ",", 0)).Meter.CurrentRange.Value)
            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasI.Pins.VSM & " =" & MeasI.Setup_ByType.VSM.Meas_Range)
            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasI.Pins.VSM & " =" & MeasI.WaitTime.VSM)
        End If
    Else
        For i = 0 To UBound(MeasI.Setup_ByTypeByPin.VSM)
            With TheHdw.DCVS.Pins(MeasI.Setup_ByTypeByPin.VSM(i).Pin)
                .Meter.mode = tlDCVSMeterCurrent
                If MeasI.Setup_ByTypeByPin.VSM(i).Meas_Range <> "0" Then .SetCurrentRanges CDbl(MeasI.Setup_ByTypeByPin.VSM(i).Meas_Range), CDbl(MeasI.Setup_ByTypeByPin.VSM(i).Meas_Range)
                .Gate = True
            End With
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasI.Setup_ByTypeByPin.VSM(i).Pin & " =" & TheHdw.DCVS.Pins(MeasI.Setup_ByTypeByPin.VSM(i).Pin).Meter.CurrentRange.Value)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasI.Setup_ByTypeByPin.VSM(i).Pin & " =" & MeasI.Setup_ByTypeByPin.VSM(i).Meas_Range)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range WaitTime, " & MeasI.Setup_ByTypeByPin.VSM(i).Pin & " =" & MeasI.WaitTime.VSM)
            End If
        Next i
    End If
    TheHdw.Wait CDbl(MeasI.WaitTime.VSM)
    measureCurrent = TheHdw.DCVS.Pins(MeasI.Pins.VSM).Meter.Read(tlStrobe, SampleSize, pc_Def_HexVS_ReadPoint, tlDCVSMeterReadingFormatAverage)
    Call HardIP_DCVS_MI_StoreAndRestoreCondition(MeasI, VSM, False)

End Function

Public Function HardIP_SetupAndMeasureCurrent_UVS256(SampleSize As Long, ByRef measureCurrent As PinListData)

    '' 20150623 - Suggest use for single pin it can mapping expected current range,
    ''                  if use for pin group it will refer the same current range to pin group by your specified.
    Dim i As Integer
    Dim Irange As Double
    Dim StoreSourceFoldLimit() As Double
    Dim StoreSinkFoldLimit() As Double
    Dim StoreFilterSetting() As Double
    Dim StoreSrcCurrentRange() As Double
    Dim PinsMaxNum As Long
    ReDim StoreSourceFoldLimit(PinsMaxNum) As Double
    ReDim StoreSinkFoldLimit(PinsMaxNum) As Double
    ReDim StoreFilterSetting(PinsMaxNum) As Double
    ReDim StoreSrcCurrentRange(PinsMaxNum) As Double
    Dim Pins_MeasureI_Together As String
    Dim MeasI As Meas_Type
    MeasI = TestConditionSeqData(Instance_Data.TestSeqNum).MeasI(Instance_Data.TestSeqSweepNum)
    Call HardIP_DCVS_MI_StoreAndRestoreCondition(MeasI, UVS256, True)
    If MeasI.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.DCVS.Pins(MeasI.Pins.UVS256)
            .Meter.mode = tlDCVSMeterCurrent
            If MeasI.Setup_ByType.UVS256.Meas_Range <> "0" Then .SetCurrentRanges CDbl(MeasI.Setup_ByType.UVS256.Meas_Range), CDbl(MeasI.Setup_ByType.UVS256.Meas_Range)
            .Gate = True
        End With

        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasI.Pins.UVS256 & " =" & TheHdw.DCVS.Pins(SplitInputCondition(MeasI.Pins.UVS256, ",", 0)).Meter.CurrentRange.Value)
            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasI.Pins.UVS256 & " =" & MeasI.Setup_ByType.UVS256.Meas_Range)
            TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasI.Pins.UVS256 & " =" & MeasI.WaitTime.UVS256)
        End If
    Else
        For i = 0 To UBound(MeasI.Setup_ByTypeByPin.UVS256)
            With TheHdw.DCVS.Pins(MeasI.Setup_ByTypeByPin.UVS256(i).Pin)
                .Meter.mode = tlDCVSMeterCurrent
                If MeasI.Setup_ByTypeByPin.UVS256(i).Meas_Range <> "0" Then .SetCurrentRanges CDbl(MeasI.Setup_ByTypeByPin.UVS256(i).Meas_Range), CDbl(MeasI.Setup_ByTypeByPin.UVS256(i).Meas_Range)
                .Gate = True
            End With
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasI.Setup_ByTypeByPin.UVS256(i).Pin & " =" & TheHdw.DCVS.Pins(MeasI.Setup_ByTypeByPin.UVS256(i).Pin).Meter.CurrentRange.Value)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasI.Setup_ByTypeByPin.UVS256(i).Pin & " =" & MeasI.Setup_ByTypeByPin.UVS256(i).Meas_Range)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasI.Setup_ByTypeByPin.UVS256(i).Pin & " =" & MeasI.WaitTime.UVS256)
            End If
        Next i
    End If
    
    ' HardIP BySite Autorange
'    If TheExec.EnableWord("HardIP_Autorange") = True Then
'        Dim uvs256_waittime As Double: uvs256_waittime = CDbl(MeasI.WaitTime.UVS256)
'        ReDim StoreSourceFoldLimit(UBound(MeasI.SaveCondition)) As Double
'        For i = 0 To UBound(MeasI.SaveCondition)
'            StoreSourceFoldLimit(i) = CDbl(MeasI.SaveCondition(i).SourceFlodLimit)
'        Next
'        Call HardIP_UVS256_autorange(MeasI.Pins.UVS256, StoreSourceFoldLimit, uvs256_waittime)
'        MeasI.WaitTime.UVS256 = CStr(uvs256_waittime)
'    End If
    
    TheHdw.Wait CDbl(MeasI.WaitTime.UVS256)
    measureCurrent = TheHdw.DCVS.Pins(MeasI.Pins.UVS256).Meter.Read(tlStrobe, pc_Def_UVS256_ReadPoint)
    
    '''--------->Save the interpose PrePat and interpose PreMeas force condition---------
    Dim num_pins As Long
    Dim instr_pins() As String
    Dim DCVS_HW_Value As String
    Call TheExec.DataManager.DecomposePinList(MeasI.Pins.UVS256, instr_pins(), num_pins)
    For i = 0 To num_pins - 1
        DCVS_HW_Value = CStr(FormatNumber(TheHdw.DCVS.Pins(instr_pins(i)).Voltage.Value, 3))
        If MeasI.ForceValueDic_HWCom.Exists(UCase(instr_pins(i))) Then
            MeasI.ForceValueDic_HWCom(UCase(instr_pins(i))) = DCVS_HW_Value
        Else
            MeasI.ForceValueDic_HWCom.Add UCase(instr_pins(i)), DCVS_HW_Value
        End If
    Next i
    '''--------->Save the interpose PrePat and interpose PreMeas force condition---------
    
    If gl_Disable_HIP_debug_log = False Then Call ForceVal_Compare(MeasI, MeasI.Pins.UVS256, "I", "DCVS") ''Carter, 20190624
    
    Call HardIP_DCVS_MI_StoreAndRestoreCondition(MeasI, UVS256, False)
End Function

Public Function HardIP_UVS256_autorange(PinName As String, Fold() As Double, WaitTime As Double)
        Dim range_siteaware As New SiteDouble
        Dim range_siteaware_temp_pld As New PinListData
        Dim range_siteaware_pld As New PinListData
        Dim site As Variant
        Dim measure_value As New PinListData
        Dim Pin As String
        Dim i As Integer
        Dim waittime_siteaware As New SiteDouble: waittime_siteaware = WaitTime
        Dim pinName_array() As String: pinName_array = Split(PinName, ",")
        Dim initcurrent_siteaware As New SiteDouble
        Dim continued As Boolean: continued = True
        Dim boo As New SiteBoolean
        
        TheExec.Datalog.WriteComment "========Autorange Start========"
        ' Initial test measure with iFold
        For i = 0 To UBound(pinName_array)
            initcurrent_siteaware = Fold(i)
            Call UVS256_setup_range_n_time(initcurrent_siteaware, range_siteaware, waittime_siteaware, pinName_array(i), boo)
            range_siteaware_pld.AddPin(pinName_array(i)).Value = range_siteaware
            range_siteaware_temp_pld = range_siteaware_pld 'add
            If WaitTime < return_max(waittime_siteaware.Abs) Then WaitTime = return_max(waittime_siteaware.Abs)
        Next i
        TheHdw.Wait WaitTime
        measure_value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, pc_Def_UVS256_ReadPoint)
        For i = 0 To UBound(pinName_array)
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "Site" & site & " " & pinName_array(i) & " Change curr_range to " & range_siteaware_pld.Pins(pinName_array(i)).Value
                TheExec.Datalog.WriteComment "Site" & site & " " & pinName_array(i) & " Measurement after changing curr_range is " & measure_value.Pins(pinName_array(i)).Value
            Next site
            range_siteaware_temp_pld.Pins(pinName_array(i)) = range_siteaware_pld.Pins(pinName_array(i))
            Call UVS256_next_range(range_siteaware_temp_pld.Pins(pinName_array(i)), pinName_array(i), measure_value)
            continued = (continued And measure_value.Pins(pinName_array(i)).Abs.Subtract(range_siteaware_temp_pld.Pins(pinName_array(i))).compare(GreaterThanOrEqualTo, 0).All(True))
        Next i
             
        ' Do until range found
        Do Until continued
            continued = True
            For i = 0 To UBound(pinName_array)
                Call UVS256_setup_range_n_time(range_siteaware_temp_pld.Pins(pinName_array(i)), range_siteaware_pld.Pins(pinName_array(i)), waittime_siteaware, pinName_array(i), measure_value.Pins(pinName_array(i)).Abs.compare(GreaterThan, range_siteaware_temp_pld.Pins(pinName_array(i))))
                If WaitTime < return_max(waittime_siteaware.Abs) Then WaitTime = return_max(waittime_siteaware.Abs)
            Next
            TheHdw.Wait WaitTime
            measure_value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, pc_Def_UVS256_ReadPoint)
            For i = 0 To UBound(pinName_array)
                For Each site In TheExec.sites
                    TheExec.Datalog.WriteComment "Site" & site & " " & pinName_array(i) & " Change curr_range to " & CStr(TheHdw.DCVS.Pins(pinName_array(i)).Meter.CurrentRange)
                    TheExec.Datalog.WriteComment "Site" & site & " " & pinName_array(i) & " Measurement after changing curr_range is " & measure_value.Pins(pinName_array(i)).Value
                Next site
                range_siteaware_temp_pld.Pins(pinName_array(i)) = range_siteaware_pld.Pins(pinName_array(i))
                Call UVS256_next_range(range_siteaware_temp_pld.Pins(pinName_array(i)), pinName_array(i), measure_value)
                continued = continued And measure_value.Pins(pinName_array(i)).Abs.Subtract(range_siteaware_temp_pld.Pins(pinName_array(i))).compare(GreaterThanOrEqualTo, 0).All(True)
            Next i
        Loop
        TheExec.Datalog.WriteComment "========Autorange Done========"
End Function
Public Function UVS256_setup_range_n_time(current As SiteDouble, range As SiteDouble, WaitTime As SiteDouble, PinName As String, boo As SiteBoolean)
    'setup UVS256 current range and test time : input current
    Dim site As Variant
    For Each site In TheExec.sites
        If boo Then GoTo casebottom
        Select Case current
            Case Is > 2.8
                range = 5.6
                WaitTime = 30 * us
            Case Is > 1.4
                range = 2.8
                WaitTime = 45 * us
            Case Is > 0.8
                range = 1.4
                WaitTime = 50 * us
            Case Is > 0.7
                range = 0.8
                WaitTime = 100 * us
            Case Is > 0.4
                range = 0.7
                WaitTime = 100 * us
            Case Is > 0.2
                range = 0.4
                WaitTime = 90 * us
            Case Is > 0.04
                range = 0.2
                WaitTime = 210 * us
            Case Is > 0.02
                range = 0.04
                WaitTime = 260 * us
            Case Is > 0.002
                range = 0.02
                WaitTime = 540 * us
            Case Is > 0.0002
                range = 0.002
                WaitTime = 3.5 * ms
            Case Is > 0.00002
                range = 0.0002
                WaitTime = 210 * us
            Case Is > 0.000004
                range = 0.00002
                WaitTime = 4 * ms
            Case 0.000004
                range = 0
                WaitTime = 18 * ms
            Case Else
                range = 0.000004
                WaitTime = 18 * ms
        End Select
casebottom:
    Next site
    For Each site In TheExec.sites
        ' Hardware setup
        If boo Then GoTo setupbottom
        With TheHdw.DCVS.Pins(PinName)
            .Meter.mode = tlDCVSMeterCurrent
            .SetCurrentRanges range, range
            .Gate = True
        End With
setupbottom:
    Next site
End Function
Public Function UVS256_next_range(range As SiteDouble, Optional PinName As String, Optional MeasVal As PinListData)
        Dim site As Variant
        Dim chtype As String
        chtype = TheExec.DataManager.ChannelType(PinName)
        
        If chtype = "DCVS" Then GoTo DCVSrange
        If chtype = "DCVSMerged4" Then GoTo DCVSrange4
        If chtype = "DCVSMerged8" Then GoTo DCVSrange8
        
        For Each site In TheExec.sites
            Select Case range
                Case 5.6
                    range = 2.8
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.007 + 0.02
                Case 2.8
                    range = 1.4
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.007 + 0.01
                Case 1.4
                    range = 0.8
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.006
                Case 0.8
                    range = 0.7
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.006
                Case 0.7
                    range = 0.4
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.007 + 0.0024
                Case 0.4
                    range = 0.2
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.0012
                Case 0.2
                    range = 0.04
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.007 + 0.00024
                Case 0.04
                    range = 0.02
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.00012
                Case 0.02
                    range = 0.002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.000012
                Case 0.002
                    range = 0.0002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.0000012
                Case 0.0002
                    range = 0.00002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.00000012
                Case 0.00002
'                    range = 0.000004
'                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.000000036
'                Case 0.000004
                    range = -1
            End Select
        Next site
        Exit Function
        
DCVSrange:
        
        For Each site In TheExec.sites
            Select Case range
                Case 0.8
                    range = 0.7
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.006
                Case 0.7
                    range = 0.2
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.0012
                Case 0.2
                    range = 0.02
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.00012
                Case 0.02
                    range = 0.002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.000012
                Case 0.002
                    range = 0.0002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.0000012
                Case 0.0002
                    range = 0.00002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.00000012
                Case 0.00002
'                    range = 0.000004
'                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.000000036
'                Case 0.000004
                    range = -1
            End Select
        Next site
        Exit Function
        
DCVSrange4:
        For Each site In TheExec.sites
            Select Case range
                Case 2.8
                    range = 0.7
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.006
                Case 0.7
                    range = 0.2
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.0012
                Case 0.2
                    range = 0.02
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.00012
                Case 0.02
                    range = 0.002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.000012
                Case 0.002
                    range = 0.0002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.0000012
                Case 0.0002
                    range = 0.00002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.00000012
                Case 0.00002
'                    range = 0.000004
'                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.000000036
'                Case 0.000004
                    range = -1
            End Select
        Next site
        Exit Function
DCVSrange8:
        For Each site In TheExec.sites
            Select Case range
                Case 5.6
                    range = 0.7
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.006
                Case 0.7
                    range = 0.2
                     MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.0012
                Case 0.2
                    range = 0.02
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.00012
                Case 0.02
                    range = 0.002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.000012
                Case 0.002
                    range = 0.0002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.0000012
                Case 0.0002
                    range = 0.00002
                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.00000012
                Case 0.00002
'                    range = 0.000004
'                    MeasVal.Pins(PinName).Value = MeasVal.Pins(PinName).Abs.Value + range * 0.005 + 0.000000036
'                Case 0.000004
                    range = -1
            End Select
        Next site
        Exit Function
        
End Function
Public Function return_max(Measure As SiteDouble) As Double
    Dim site As Variant
    For Each site In TheExec.sites
        If Measure(site) > return_max Then return_max = Measure(site)
    Next site
End Function
Public Function HardIP_Freq_MeasFreqStart(Pin As PinList, Interval As Double, ByRef freq As PinListData, Optional CustomizeWaitTime As String)
    
    '' 20150918 - Check whether have duplicated pins, change measure mrthod from parallel to serial if pin name duplicate.
    Dim FlagDuplicatePins As Boolean
    FlagDuplicatePins = CheckDuplicateInputPins(CStr(Pin))
    
    
    Dim CounterValue As New PinListData
    Dim site As Variant
    
    ''20150623 - Add CustomizeWaitTime
    If CustomizeWaitTime <> "" Then
        TheHdw.Wait (CDbl(CustomizeWaitTime))
    End If
    
    TheHdw.Digital.Pins(Pin).FreqCtr.Clear
    TheHdw.Digital.Pins(Pin).FreqCtr.start
    
    '' 20150918 - Check whether have duplicated pins, change measure mrthod from parallel to serial if pin name duplicate.
    Dim i As Long
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
End Function

Public Function ContentIsNumeral(FoceValue As String) As Boolean
    
On Error GoTo ChkFalse

    Dim d_FoceValue As Double
    d_FoceValue = CDbl(FoceValue)
    
    ContentIsNumeral = True
    Exit Function
    
ChkFalse:
    ContentIsNumeral = False
    
End Function

Public Function IO_HardIP_PPMU_Measure_Z(TestPinArrayIV() As String, TestSeqNum As Integer, TestSeqNumIdx As Long, ForceSequenceArray() As String, _
                        k As Long, Pat As Variant, Flag_SingleLimit As Boolean, HighLimitVal As Double, LowLimitVal As Double, TestLimitPerPin_VIR As String, TestIrange() As String, _
                        FlowTestNme() As String, _
                        Optional RAK_Flag As Enum_RAK = 0, _
                        Optional Rtn_SweepTestName As String, Optional OutputTname As String, Optional WaitTime_Z As String) As Long

    Dim MeasureValue As New PinListData
    Dim Force_idx As Integer
    Dim site As Variant
    Dim TestNum As Long
    
    Dim Imped As New PinListData
    Dim Pin  As Variant
    Dim RAK_Pin As String
    Dim GetRakVal As Double
    Dim p As Long
    Dim ForceV  As Double

    Dim MeasCurr1 As New PinListData
    Dim MeasCurr2 As New PinListData
    ''==========================================================================
    Dim MeasurePinAry() As String
    Dim ForceVoltAry() As String
    Dim MeasurePin As String
    Dim OutputTname_format() As String
    
    Dim ForceSequenceArrayByPin() As String
    Dim ForceSequenceArraybySweep() As String
    Dim i As Long
    Dim Sweep_Flag As Boolean: Sweep_Flag = False
    Dim Sweep_Name As String
    
    
    If (InStr(ForceSequenceArray(TestSeqNum), "sweep") <> 0) Then
        Sweep_Flag = True
        ForceSequenceArray(TestSeqNum) = Replace(ForceSequenceArray(TestSeqNum), ";sweep", "")
    End If
    
    
    
    '' Force Pin
    If UBound(TestPinArrayIV) = 0 Then
        MeasurePinAry = Split(TestPinArrayIV(0), ",")
        MeasurePin = TestPinArrayIV(0)
        TheHdw.Digital.Pins(TestPinArrayIV(0)).Disconnect
    Else
        MeasurePinAry = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        MeasurePin = TestPinArrayIV(TestSeqNumIdx)
        TheHdw.Digital.Pins(TestPinArrayIV(TestSeqNumIdx)).Disconnect
    End If
        
    
'    If InStr(ForceSequenceArray(TestSeqNumIdx), "&") = 0 Then
'        ForceSequenceArrayByPin = Split(ForceSequenceArray(TestSeqNumIdx), ",")
'    End If
    
    
    For i = 0 To UBound(MeasurePinAry)
        
        If InStr(ForceSequenceArray(TestSeqNumIdx), "&") <> 0 Then
            
            ForceSequenceArrayByPin = Split(ForceSequenceArray(TestSeqNumIdx), ",")
            
            If (UBound(ForceSequenceArrayByPin) >= i) Then
                ForceSequenceArraybySweep = Split(ForceSequenceArrayByPin(i), ":")
            Else
                ForceSequenceArraybySweep = Split(ForceSequenceArrayByPin(0), ":")
            End If
            
            If UBound(ForceSequenceArraybySweep) >= k - 1 Then
                Sweep_Name = ForceSequenceArraybySweep(k - 1)
                ForceVoltAry = Split(ForceSequenceArraybySweep(k - 1), "&")
            Else
                Sweep_Name = ForceSequenceArraybySweep(0)
                ForceVoltAry = Split(ForceSequenceArraybySweep(0), "&")
            End If
        Else
            ForceVoltAry = Split(ForceSequenceArray(TestSeqNumIdx), ",")
        End If
        Dim DiffVolt As Double
        
        '' 20150727 - Use low voltage for first measurement
        Dim TempSwitchValue
        If ForceVoltAry(0) > ForceVoltAry(1) Then
            TempSwitchValue = ForceVoltAry(1)
            ForceVoltAry(1) = ForceVoltAry(0)
            ForceVoltAry(0) = TempSwitchValue
        End If
        DiffVolt = ForceVoltAry(1) - ForceVoltAry(0)
        
        '' 20150112 - Check number whether differrent between measure current range and force pin, add defalut value to let input number are the same.
        '' Measure Current range
        Dim Measure_I_Range() As String
        Dim Measure_I_Range_Index As Long
        If UBound(TestIrange) = 0 Then
            Measure_I_Range = Split(TestIrange(0), ",")
        Else
            Measure_I_Range = Split(TestIrange(TestSeqNumIdx), ",")
        End If
        Call VIR_CheckTestCondition_Measure_I_R_Z("Z", MeasurePinAry, Measure_I_Range)
        
        Dim TestNameInput As String
        
        'If TPModeAsCharz_GLB = True Then
        '    TestNameInput = FlowTestNme(gl_CZ_FlowTestName_Counter)
        '    'gl_CZ_FlowTestName_Counter = gl_CZ_FlowTestName_Counter + 1
        'Else
            TestNameInput = "2_Point_Imp_meas_"
            If Rtn_SweepTestName <> "" Then
                TestNameInput = TestNameInput & "_" & Rtn_SweepTestName
            End If
        'End If
        
        If (InStr(Measure_I_Range(Measure_I_Range_Index), ":") <> 0) Then
            If (UBound(Split(Measure_I_Range(Measure_I_Range_Index), ":"))) >= k Then
                Measure_I_Range(Measure_I_Range_Index) = Split(Measure_I_Range(Measure_I_Range_Index), ":")(k)
            Else
                Measure_I_Range(Measure_I_Range_Index) = Split(Measure_I_Range(Measure_I_Range_Index), ":")(0)
            End If
        End If

    'For i = 0 To UBound(MeasurePinAry)
        With TheHdw.PPMU.Pins(MeasurePinAry(i))
            .ForceV ForceVoltAry(0), Measure_I_Range(Measure_I_Range_Index)
            '.Connect
            '.Gate = tlOn
        End With
        
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Pin = " & MeasurePinAry(i) & " Force Volt = " & ForceVoltAry(0) & " Measure Current Limit =" & Measure_I_Range(Measure_I_Range_Index) & " =====> Curr_meas Meter I range setting, " & MeasurePinAry(i) & " =" & TheHdw.PPMU.Pins(MeasurePinAry(i)).MeasureCurrentRange
        
        If WaitTime_Z = "" Then
            TheHdw.Wait (2 * ms)
        Else
            TheHdw.Wait CDbl(WaitTime_Z)
        End If
        
        DebugPrintFunc_PPMU CStr(MeasurePinAry(i))
''        MeasCurr1 = TheHdw.PPMU.Pins(MeasurePinAry(i)).Read(tlPPMUReadMeasurements, 10)
        MeasCurr1 = TheHdw.PPMU.Pins(MeasurePinAry(i)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            
        If gl_Disable_HIP_debug_log = False Then
            For Each Pin In MeasCurr1.Pins
                For Each site In TheExec.sites
                    TheExec.Datalog.WriteComment "Site " & site & " Force Volt = " & ForceVoltAry(0) & ", Measure Current Pin = " & Pin & ", Value = " & MeasCurr1.Pins.Item(Pin).Value(site)
                Next site
            Next Pin
        End If
        
        '20151103 print force condition
        Call Print_Force_Condition("Z", MeasCurr1)
        
        
        With TheHdw.PPMU.Pins(MeasurePinAry(i))
            .ForceV ForceVoltAry(1), Measure_I_Range(Measure_I_Range_Index)
            '.Connect
            '.Gate = tlOn
        End With
        
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Pin = " & MeasurePinAry(i) & " Force Volt = " & ForceVoltAry(1) & " and Measure Current Range = " & TheHdw.PPMU.Pins(MeasurePinAry(i)).MeasureCurrentRange
        
        If WaitTime_Z = "" Then
            TheHdw.Wait (2 * ms)
        Else
            TheHdw.Wait CDbl(WaitTime_Z)
        End If
        
        DebugPrintFunc_PPMU CStr(MeasurePinAry(i))
''        MeasCurr2 = TheHdw.PPMU.Pins(MeasurePinAry(i)).Read(tlPPMUReadMeasurements, 10)
        MeasCurr2 = TheHdw.PPMU.Pins(MeasurePinAry(i)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
        
        If gl_Disable_HIP_debug_log = False Then
            For Each Pin In MeasCurr2.Pins
                For Each site In TheExec.sites
                     TheExec.Datalog.WriteComment "Site " & site & " Force Volt = " & ForceVoltAry(1) & ", Measure Current Pin = " & Pin & ", Value = " & MeasCurr2.Pins.Item(Pin).Value(site)
                Next site
            Next Pin
        End If

        '20151103 print force condition
        Call Print_Force_Condition("Z", MeasCurr2)
        
        For Each Pin In MeasCurr2.Pins
            For Each site In TheExec.sites      '''Offline force current different
                If MeasCurr2.Pins(Pin).Value = MeasCurr1.Pins(Pin).Value Then
                    MeasCurr1.Pins(Pin).Value = MeasCurr1.Pins(Pin).Value + 0.000000001
                End If
            Next site
        Next Pin
        
        Imped = MeasCurr2.Math.Subtract(MeasCurr1).Invert.Multiply(DiffVolt).Abs
        
        If RAK_Flag = R_TraceOnly Then
            'Dim RakV() As Double
    ''        Call AddRXRAkPinValue
            For Each Pin In Imped.Pins
                For Each site In TheExec.sites
                    RAK_Pin = CStr(Pin)
                    'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(RAK_Pin, Site)
''                    If InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
''                      GetRakVal = RakV(0) + FT_Card_RAK.Pins(Pin).Value(Site)
''                    Else
''                        GetRakVal = RakV(0) + CP_Card_RAK.Pins(Pin).Value(Site)
''                    End If
                    GetRakVal = CurrentJob_Card_RAK.Pins(Pin).Value(site)
                    
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment Pin & " = " & Imped.Pins.Item(Pin).Value(site) & ", RAK val = " & GetRakVal
                    Imped.Pins.Item(Pin).Value(site) = Imped.Pins.Item(Pin).Value(site) - GetRakVal
                Next site
            Next Pin
        ElseIf RAK_Flag = R_PathWithContact Then
             For Each Pin In Imped.Pins
                For Each site In TheExec.sites
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment Pin & " = " & Imped.Pins.Item(Pin).Value(site) & ", R_Path val = " & R_Path_PLD.Pins(Pin).Value(site)
                    Imped.Pins.Item(Pin).Value(site) = Imped.Pins.Item(Pin).Value(site) - R_Path_PLD.Pins(Pin).Value(site)
                Next site
            Next Pin
        End If
        
        For p = 0 To Imped.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("Z", Imped.Pins(p), , TestSeqNum, k, Sweep_Name)
            TheExec.Flow.TestLimit Imped.Pins(p), , , , , , unitCustom, , TestNameInput, , , , unitVolt, " ohm", , ForceResults:=tlForceFlow
        Next p
        Measure_I_Range_Index = Measure_I_Range_Index + 1
        
    Next i
    
End Function

Public Function CheckDuplicateInputPins(CheckPins As String) As Boolean

    '' 20150918 - Check whether have duplicated pins, change measure mrthod from parallel to serial if pin name duplicate.
    Dim ActualPins() As String
    Dim ActtualNumberPins As Long
    Call TheExec.DataManager.DecomposePinList(CheckPins, ActualPins(), ActtualNumberPins)
      
    Dim i As Long
    Dim InputPins() As String
    Dim Pins() As String
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
End Function

Public Function Print_Force_Condition(MeasType As String, MeasureValue As PinListData)

    '' 20151103 - Print force condition
    Dim PrintPinName As String
    Dim PrintForceVal As Double
    Dim All_Force_Condition As String
    Dim p As Integer
    
    If gl_Disable_HIP_debug_log = False Then
        If (LCase(MeasType) = "v") Then
            For p = 0 To MeasureValue.Pins.Count - 1
                PrintPinName = MeasureValue.Pins(p)
                PrintForceVal = FormatNumber(TheHdw.PPMU(PrintPinName).current.Value, 7)
                
                If (p = 0) Then
                    All_Force_Condition = "force condition: " & PrintPinName & ": " & PrintForceVal & "A"
                Else
                    All_Force_Condition = All_Force_Condition & ", " & PrintPinName & ": " & PrintForceVal & "A"
                End If
            Next p
        Else
            For p = 0 To MeasureValue.Pins.Count - 1
                PrintPinName = MeasureValue.Pins(p)
                PrintForceVal = FormatNumber(TheHdw.PPMU(PrintPinName).Voltage.Value, 3)
                
                If (p = 0) Then
                    All_Force_Condition = "force condition: " & PrintPinName & ": " & PrintForceVal & "V"
                Else
                    All_Force_Condition = All_Force_Condition & ", " & PrintPinName & ": " & PrintForceVal & "V"
                End If
            Next p
        End If
        
        TheExec.Datalog.WriteComment All_Force_Condition
    End If
End Function





Public Function Print_Force_Condition_I(MeasType As String, Save_force_data As DSPWave, MeasureValue As PinListData)

    '' 20180620 - Print force condition by CS
    Dim PrintPinName As String
    Dim PrintForceVal As Double
    Dim All_Force_Condition As String
    Dim p As Integer
    
    If gl_Disable_HIP_debug_log = False Then
        If (LCase(MeasType) = "v") Then
            For p = 0 To MeasureValue.Pins.Count - 1
                PrintPinName = MeasureValue.Pins(p)
                PrintForceVal = Round(TheHdw.PPMU(PrintPinName).current.Value, 7)
                
                If (p = 0) Then
                    All_Force_Condition = "force condition: " & PrintPinName & ": " & PrintForceVal & "A"
                Else
                    All_Force_Condition = All_Force_Condition & ", " & PrintPinName & ": " & PrintForceVal & "A"
                End If
            Next p
        Else
            For p = 0 To MeasureValue.Pins.Count - 1
                PrintPinName = MeasureValue.Pins(p)
                PrintForceVal = Round(Save_force_data.Element(p), 3)
                
                If (p = 0) Then
                    All_Force_Condition = "force condition: " & PrintPinName & ": " & PrintForceVal & "V"
                Else
                    All_Force_Condition = All_Force_Condition & ", " & PrintPinName & ": " & PrintForceVal & "V"
                End If
            Next p
        End If
        
        TheExec.Datalog.WriteComment All_Force_Condition
    End If
End Function

Public Function IPF_CZ_PrintFreq(argc As Integer, argv() As String) As Long

    '' 20151114 - Print Freq measurement during shmoo
    Dim site As Variant
    Dim i As Long
    Dim X_SetupName As String
    Dim Y_SetupName As String
    Dim Volt_pointval As Double
    Dim FRC_pointval As Double

''    If UCase(argv(2)) = UCase("PrintFreq") Then
        X_SetupName = TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Item(tlDevCharShmooAxis_X).StepName
        Y_SetupName = TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Item(tlDevCharShmooAxis_Y).StepName
        For Each site In TheExec.sites
            Volt_pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
            FRC_pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
            If gl_Disable_HIP_debug_log = False Then
                For i = 0 To G_MeasFreqForCZ.Pins.Count - 1
                    TheExec.Datalog.WriteComment ("Site = " & site & ",  " & X_SetupName & "=" & Volt_pointval & " V,  " & Y_SetupName & "=" & FRC_pointval & "Hz,  Pin name = " & G_MeasFreqForCZ.Pins(i) & " , Frequency value is " & G_MeasFreqForCZ.Pins(i).Value(site))
                Next i
            End If
        Next site
''    End If
End Function
Public Function Freq_ProcessEventSourceTerminationMode(MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, ByRef MeasF_EventSource As FreqCtrEventSrcSel, ByRef MeasF_EnableVtMode As Boolean)

    Select Case MeasF_EventSourceWithTerminationMode
        Case 5:
            MeasF_EventSource = VOH
            MeasF_EnableVtMode = False
        Case 6:
            MeasF_EventSource = vol
            MeasF_EnableVtMode = False
        Case 4:
            MeasF_EventSource = BOTH
            MeasF_EnableVtMode = False
        Case 2:
            MeasF_EventSource = VOH
            MeasF_EnableVtMode = True
        Case 3:
            MeasF_EventSource = vol
            MeasF_EnableVtMode = True
        Case 1:
            MeasF_EventSource = BOTH
            MeasF_EnableVtMode = True
        Case Else:
            MeasF_EventSource = BOTH
            MeasF_EnableVtMode = True
    End Select
    
End Function

Public Function VIR_CheckTestCondition_Measure_I_R_Z(InputSequence As String, ForceByPin() As String, ByRef Measure_I_Range() As String)
'' 20150108 - MI is OK but MR and MZ, need the same force condition in the same sequence because apply pin in the same time but pinlistdata(I) is difficult to Multiply different voltage value
Dim Diff_Index As Long
Dim i As Long
Dim TempString_Measure_I_Range As String
Dim Default_MeasureCurrRange As Double
''Default_MeasureCurrRange = 0.05
Default_MeasureCurrRange = pc_Def_VIR_MeasCurrRange

Dim b_CurrRangeNoInput As Boolean

Select Case UCase(InputSequence)
    Case "R", "I", "Z":
        If UBound(Measure_I_Range) = -1 Then
            Diff_Index = UBound(ForceByPin)
            For i = 0 To Diff_Index
                If i = 0 Then
                    TempString_Measure_I_Range = Default_MeasureCurrRange
                Else
                    TempString_Measure_I_Range = TempString_Measure_I_Range & "," & Default_MeasureCurrRange
                End If
            Next i
            Measure_I_Range = Split(TempString_Measure_I_Range, ",")
            
        Else
            If UBound(ForceByPin) - UBound(Measure_I_Range) <> 0 Then
                
                Diff_Index = UBound(ForceByPin) - UBound(Measure_I_Range)
                
                For i = 0 To UBound(Measure_I_Range)
                    If i = 0 Then
                        TempString_Measure_I_Range = Measure_I_Range(i)
                    Else
                        TempString_Measure_I_Range = TempString_Measure_I_Range & "," & Measure_I_Range(i)
                    End If
                Next i
                
                ' TempString_Measure_I_Range = TempString_Measure_I_Range
                
                For i = 0 To Diff_Index - 1
                    TempString_Measure_I_Range = TempString_Measure_I_Range & "," & Measure_I_Range(0)
                Next i
                
                Measure_I_Range = Split(TempString_Measure_I_Range, ",")
            End If
            
            '' 20160111 - Add default range if input argement only specify ","
            If Measure_I_Range(0) = "" Then Measure_I_Range(0) = Default_MeasureCurrRange
            
            If UBound(Measure_I_Range) > 0 Then
                For i = 1 To UBound(Measure_I_Range)
                    If Measure_I_Range(i) = "" Then
                        Measure_I_Range(i) = Measure_I_Range(0)
                    End If
                Next
            End If
        End If
        
    Case Else:
End Select
End Function

Public Function MeasureR_ForceDifferentVolt(MeasureValue As PinListData, Impedence As PinListData) As Long
    Dim site As Variant
    Dim PinName As String
    Dim ForceVal As Double
    Dim All_Force_Condition As String
    Dim p As Integer
    
    Dim AddPinName_Switch As Boolean
    AddPinName_Switch = False
    For p = 0 To MeasureValue.Pins.Count - 1
        For Each site In TheExec.sites
            PinName = MeasureValue.Pins(p)
            If AddPinName_Switch = False Then
                Impedence.AddPin (PinName)
                AddPinName_Switch = True
            End If
            ForceVal = TheHdw.PPMU(PinName).Voltage.Value
            
            Impedence.Pins(PinName).Value(site) = MeasureValue.Pins(PinName).Invert.Multiply(ForceVal)(site)
        Next site
        AddPinName_Switch = False
    Next p
End Function


Public Function HardIP_OnProgramStarted_Process() As Long

    On Error GoTo err
    If TheExec.EnableWord("HIP_TTR_FailResultOnly") = True Then
        gl_Disable_HIP_debug_log = True
    Else
            gl_Disable_HIP_debug_log = False
    End If

    gl_Tname_Alg_Index = 0   '''clear global variable from touchdown to touchdown

    HardIP_RAK_Init

'init job string  for TTR 20151225

   CurrentJobName_L = LCase(TheExec.CurrentJob)
   CurrentJobName_U = UCase(TheExec.CurrentJob)

   Range_Check_Enable_Word = Range_Check_Enable_Word

        TheHdw.Digital.Patgen.TimeoutEnable = True
''        TheHdw.Digital.TimeOut = 10
        TheHdw.Digital.Patgen.TimeOut = pc_Def_HardIP_PatGenTimeout
        
    Call GetInstTypToDic("all_dcvi,all_dcvi_analog,all_hexvs,all_uvs,all_digital")
'    theexec.DataLog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
'    theexec.DataLog.Setup.DatalogSetup.DisablePinNameInPTR = True
'    theexec.DataLog.ApplySetup
    Exit Function
err:
    Stop
    Resume Next
End Function


                                                                                                                                                        
Public Function GetInstrument(PinList As String, site As Variant) As String
    Dim chanString As String
    Dim PinName() As String
    Dim NumberPins As Long
    Call TheExec.DataManager.DecomposePinList(PinList, PinName(), NumberPins)
    Call TheExec.DataManager.GetChannelStringFromPinAndSite(PinName(0), site, chanString)
    Dim slotstr() As String
    Dim slot As Long
    
    If chanString = "" Then
        TheExec.Datalog.WriteComment ("Warnning : Please check pin type of  " & PinList & " in channel map")
    Else
        slotstr = Split(chanString, ".")
        slot = CLng(slotstr(0))
        GetInstrument = TheHdw.config.Slots(slot).Type
    End If
End Function

Public Function SrcVoltFromFlowForLoop(FlowForLoopIntegerName As String, powerPin As PinList, StartStopStep As String) As Long
    
    
    Dim StepIndex As Long
    Dim i As Long
    
    StepIndex = TheExec.Flow.var(FlowForLoopIntegerName).Value
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Get index = " & StepIndex)
    
    Dim StartVolt As Double
    Dim StopVolt As Double
    Dim StepVolt As Double
    Dim ForceVolt As Double
    
    Dim DecomposeString() As String
    
    DecomposeString = Split(StartStopStep, ",")
    StartVolt = CDbl(DecomposeString(0))
    StopVolt = CDbl(DecomposeString(1))
    StepVolt = CDbl(DecomposeString(2))
    
    ForceVolt = StartVolt + StepIndex * StepVolt
    
    If ForceVolt > StopVolt Then
        TheExec.Datalog.WriteComment ("Warning !! Force Volt over Stop Volt")
        Exit Function
    End If
    
    Dim InstName As String
    Dim Pins() As String
    Dim NumberPins As Long
    Call TheExec.DataManager.DecomposePinList(powerPin, Pins(), NumberPins)
     
    InstName = GetInstrument(Pins(0), 0)

    Select Case InstName
        Case "DC-07"
            TheHdw.DCVI.Pins(powerPin).Voltage = ForceVolt
        
        Case "VHDVS"
            TheHdw.DCVS.Pins(powerPin).Voltage.Value = ForceVolt
            
        Case "HexVS"
            TheHdw.DCVS.Pins(powerPin).Voltage.Value = ForceVolt
            
        Case Else
        
    End Select
    
    TheHdw.Wait (1 * ms)
    
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment (powerPin.Value & " output voltage = " & ForceVolt)


End Function

Public Function Rev_BinArray(m_binarr() As Long) As String
    Dim i As Long
    Dim SrcBitBinaryString  As String
    '' 20150811 - Reverse order of binary string
    For i = UBound(m_binarr) To 0 Step -1
        If i = UBound(m_binarr) Then
            SrcBitBinaryString = m_binarr(i)
        Else
            SrcBitBinaryString = SrcBitBinaryString & m_binarr(i)
        End If
        Rev_BinArray = SrcBitBinaryString
    Next i
End Function

Public Function CheckInputStringByAt(ByRef InPutString As String) As String
    If Left(InPutString, 1) <> "@" And InPutString <> "" Then
        InPutString = "@" & InPutString
    End If
    If InPutString <> "" Then
        InPutString = Mid(InPutString, 2, Len(InPutString) - 1)
    End If
    CheckInputStringByAt = InPutString
End Function

Public Function DigCapDataProcessByDSP(CUS_Str_DigCapData As String, OutDspWave As DSPWave, DigCap_Sample_Size As Long, DigCap_DataWidth As Long, Optional CUS_Str_MainProgram As String, _
                        Optional BypassAllDigCapTestLimit As Boolean = False, Optional DigCap_PinName As String, Optional Tname As String, Optional MSB_First_Flag As Boolean)
    
    If ByPassTestLimit Then: BypassAllDigCapTestLimit = True
    
    Dim site As Variant
    Dim i As Long, j As Long
    Dim Str_PrintBinary As New SiteVariant
    Dim ConvertedDataWf As New DSPWave
    Dim SourceBitStrmWf As New DSPWave
    Dim NoOfSamples As New SiteLong
    
    Dim FlexibleConvertedDataWf() As New DSPWave

    '' 20160328
    Dim TestLimitWithTestName As New PinListData
    
    Dim TestInstanceName As String
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    '20191107 CT add
    Dim dspBinStr As String
    Dim dspBinStr_counter As New SiteLong: dspBinStr_counter = 0
    On Error GoTo err
    TestInstanceName = TheExec.DataManager.instanceName

    ''20190604 -- Update Sweep Flow Var TName (Support 1D Only)
    
    If InStr(Instance_Data.Interpose_PrePat, "[") > 0 And InStr(Instance_Data.Interpose_PrePat, "]") > 0 Then ''e.g. PinA:V:0.5*[SrcCodeIndx]
        Dim TempInterposeStr As String
        Dim FlowVarStr As String

        Dim TempIdx As Long
        Dim TempIdxSemi As Long
        Dim SplitLeftBigColon() As String
        Dim SplitSemiColon() As String
        Dim SplitColon() As String
        Dim SweepFlowVarTNameStr As String
        Dim SweepFlowVarEqn As String
        SweepFlowVarTNameStr = ""

        TempInterposeStr = Instance_Data.Interpose_PrePat
        SplitSemiColon = Split(TempInterposeStr, ";")
        For TempIdxSemi = 0 To UBound(SplitSemiColon)
            If InStr(SplitSemiColon(TempIdxSemi), "[") > 0 Then
                SplitColon = Split(SplitSemiColon(TempIdxSemi), ":")
                SweepFlowVarEqn = SplitColon(UBound(SplitColon))
                
                SplitLeftBigColon = Split(SweepFlowVarEqn, "[")
                For TempIdx = 0 To UBound(SplitLeftBigColon)
                    If InStr(SplitLeftBigColon(TempIdx), "]") > 0 Then
                        FlowVarStr = Split(SplitLeftBigColon(TempIdx), "]")(0)
                        SweepFlowVarEqn = Replace(SweepFlowVarEqn, "[" & FlowVarStr & "]", CStr(TheExec.Flow.var(FlowVarStr).Value))
                        SweepFlowVarTNameStr = SweepFlowVarTNameStr & CStr(Evaluate(SweepFlowVarEqn))
                        If CDbl(SweepFlowVarTNameStr) <= 1 Then SweepFlowVarTNameStr = CStr(CDbl(SweepFlowVarTNameStr) * 1000) ''May Need Update
                        SweepFlowVarTNameStr = Replace(SweepFlowVarTNameStr, ".", "p")
                        Exit For
                    End If
                Next TempIdx
            End If
        Next TempIdxSemi
            
 
    End If
    
    ''20190604 -- Update Sweep Flow Var TName (END) ----------------------------------------
    
    ''20170418 - Move out from If DigCap_DataWidth <> 0 And InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") = 0 And CUS_Str_MainProgram = "" Then
    Dim p As Long
    Dim PinName As String
    Dim DigCapValue As New PinListData
    Dim b_FirstTimeSwitch As Boolean
    
    '' 20160211 - Process format by DigCap_DataWidth, capture word size is fixed
    If DigCap_DataWidth <> 0 And InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") = 0 Then
    
            Dim CalcOutputDSPWave As New DSPWave
            Dim CalcEyeWidth As New SiteLong
            Dim FinalEyeOutBitNum As Long
            Dim TestLimitForEyeSweep As New DSPWave
        
        ''20170811 - EyeSweep for LPDPRX
        If UCase(CUS_Str_MainProgram) = UCase("LPDPRX_EyeSweep") Then
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (DigCap_DataWidth) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
               If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
            SourceBitStrmWf = OutDspWave
        
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            
            FinalEyeOutBitNum = DigCap_Sample_Size / 32
            rundsp.LPDPRX_EyeSweep ConvertedDataWf, FinalEyeOutBitNum, CalcOutputDSPWave, CalcEyeWidth
            
            For Each site In TheExec.sites
                Str_PrintBinary(site) = ""
                For i = 1 To CalcOutputDSPWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & CalcOutputDSPWave(site).Element(i - 1)
                Next i
                If gl_Disable_HIP_debug_log = False And gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") Output eye bits = " & FinalEyeOutBitNum & ", Binary string = " & Str_PrintBinary(site))
                ''20170510 Store Binary String for Eye Diagram
                Eye_Diagram_Binary(TheExec.Flow.var("SrcCodeIndx").Value + 31)(site) = Str_PrintBinary(site)
            Next site
'            Dim TestLimitForEyeSweep As New DSPWave
            rundsp.BitWf2Arry CalcOutputDSPWave, DigCap_DataWidth, NoOfSamples, TestLimitForEyeSweep
    
            b_FirstTimeSwitch = True
            For Each site In TheExec.sites
                PinName = "EyeCapWord_"
                Exit For
            Next site
            For Each site In TheExec.sites
                For i = 1 To TestLimitForEyeSweep.SampleSize
                    If b_FirstTimeSwitch Then
                        DigCapValue.AddPin (PinName & CStr(i - 1))
                    End If
                    DigCapValue.Pins(PinName & CStr(i - 1)).Value(site) = TestLimitForEyeSweep(site).Element(i - 1)
                Next i
                b_FirstTimeSwitch = False
            Next site
    
            If BypassAllDigCapTestLimit = False Then
                For p = 0 To DigCapValue.Pins.Count - 1
                    TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, "EyeCpatureCode" & p, CInt(p), p)
                    TheExec.Flow.TestLimit DigCapValue.Pins(p), 0, 2 ^ DigCap_DataWidth - 1, Tname:=TestNameInput, PinName:="EyeCpatureCode_" & p, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                Next p
                
                TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, "EyeWidth", 0, 0)
                TheExec.Flow.TestLimit resultVal:=CalcEyeWidth, Tname:=TestNameInput, ForceResults:=tlForceFlow

            End If
        '20170811 PCIE Eye Sweep
        ElseIf UCase(CUS_Str_MainProgram) = UCase("PCIE_EyeSweep") Then
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (DigCap_DataWidth) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
            SourceBitStrmWf = OutDspWave
        
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            
'            Dim CalcOutputDSPWave As New DSPWave
'            Dim CalcEyeWidth As New SiteLong
'            Dim FinalEyeOutBitNum As Long
            FinalEyeOutBitNum = DigCap_Sample_Size / 20
            rundsp.PCIE_EyeSweep ConvertedDataWf, FinalEyeOutBitNum, CalcOutputDSPWave, CalcEyeWidth
            
            For Each site In TheExec.sites
                Str_PrintBinary(site) = ""
                For i = 1 To CalcOutputDSPWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & CalcOutputDSPWave(site).Element(i - 1)
    ''                If i Mod (DigCap_DataWidth) = 0 Then
    ''                    Str_PrintBinary(Site) = Str_PrintBinary(Site) & ","
    ''                End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") Output eye bits = " & FinalEyeOutBitNum & ", Binary string = " & Str_PrintBinary(site))
                ''20170510 Store Binary String for Eye Diagram
                Eye_Diagram_Binary(TheExec.Flow.var("SrcCodeIndx").Value + 31)(site) = Str_PrintBinary(site)
            Next site
'            Dim TestLimitForEyeSweep As New DSPWave
            rundsp.BitWf2Arry CalcOutputDSPWave, DigCap_DataWidth, NoOfSamples, TestLimitForEyeSweep
    
            b_FirstTimeSwitch = True
            For Each site In TheExec.sites
                PinName = "EyeCapWord_"
                Exit For
            Next site
            For Each site In TheExec.sites
                For i = 1 To TestLimitForEyeSweep.SampleSize
                    If b_FirstTimeSwitch Then
                        DigCapValue.AddPin (PinName & CStr(i - 1))
                    End If
                    DigCapValue.Pins(PinName & CStr(i - 1)).Value(site) = TestLimitForEyeSweep(site).Element(i - 1)
                Next i
                b_FirstTimeSwitch = False
            Next site
    
            If BypassAllDigCapTestLimit = False Then
                For p = 0 To DigCapValue.Pins.Count - 1
                    
                    TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, "EyeCpatureCode" & p, CInt(p), p)
                    TheExec.Flow.TestLimit DigCapValue.Pins(p), 0, 2 ^ DigCap_DataWidth - 1, Tname:=TestNameInput, PinName:="EyeCpatureCode_" & p, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                Next p
                
                TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, "EyeWidth", 0, p)
                TheExec.Flow.TestLimit resultVal:=CalcEyeWidth, Tname:=TestInstanceName & "_EyeWidth", ForceResults:=tlForceFlow
                
            End If
        Else
    
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (DigCap_DataWidth) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
            SourceBitStrmWf = OutDspWave
        
            
            If MSB_First_Flag Then
                rundsp.BitWf2Arry_MSB1st SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            Else
                rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            End If
            
            '' 20160211 - Get pin name from dsp wave
    ''        Dim p As Long
    ''        Dim PinName As String
    ''        Dim DigCapValue As New PinListData
    ''        Dim b_FirstTimeSwitch As Boolean
            b_FirstTimeSwitch = True
            For Each site In TheExec.sites
                PinName = OutDspWave(site).Info.WaveName & "_DigCapWord_"
                Exit For
            Next site
            For Each site In TheExec.sites
                For i = 1 To ConvertedDataWf.SampleSize
                    If b_FirstTimeSwitch Then
                        DigCapValue.AddPin (PinName & CStr(i - 1))
                    End If
                    DigCapValue.Pins(PinName & CStr(i - 1)).Value(site) = ConvertedDataWf(site).Element(i - 1)
                Next i
                b_FirstTimeSwitch = False
            Next site
            
            If BypassAllDigCapTestLimit = False Then
                For p = 0 To DigCapValue.Pins.Count - 1
                    TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, "CpatureCode" & p, CInt(p), p)
                    TheExec.Flow.TestLimit DigCapValue.Pins(p), 0, 2 ^ DigCap_DataWidth - 1, Tname:=TestNameInput, PinName:="CpatureCode_" & p, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                Next p
            End If
    ''        Call CUS_VFI_MainProgram_ECID(CUS_Str_MainProgram, DigCapValue)

        End If


        
    '' 20160212 - Process format by DSSC_OUT, capture word size is flexible, also parse with/without test name.
    ElseIf InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") <> 0 Then

        Dim Split_Num() As String
        Dim StartNum As Long
        '' 20151231 - Add rule to check new format that include test name and parse bits
        Dim ParseStringByBits As String
        Dim ParseStringForTestName As String
        Dim DSSC_Out_DecompseByComma() As String
        Dim DSSC_Out_DecompseByColon() As String
        Dim b_DSSC_Out_InvolveTestName As Boolean
        ParseStringByBits = ""
        ParseStringForTestName = ""
        b_DSSC_Out_InvolveTestName = False
        Dim DecomposeTestName() As String
        Dim DecomposeParseDigCapBit() As String
        
        ''20160807 - Add directionary to store DigCap DSPwave
        Dim ParseStringForDirectionary As String
        Dim DecomposeDirectionary() As String
        Dim b_ParseForDirectionary_Switch As Boolean
        b_ParseForDirectionary_Switch = False
        
        Dim b_ParseForGrayCode_Switch As Boolean
        b_ParseForGrayCode_Switch = False
        Dim ParseStringForGrayCode As String
        Dim DecomposeGrayCode() As String
        
        
        ''''''''''''''''''''20191028 CT debug
        If UCase(CUS_Str_MainProgram) Like UCase("*PRINT_DIGCAP*") Then
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (OutDspWave.SampleSize) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
        End If
        ''''''''''''''''''''20191028 CT debug
        
        If InStr(UCase(CUS_Str_DigCapData), ":") <> 0 Then
            b_DSSC_Out_InvolveTestName = True
            DSSC_Out_DecompseByComma = Split(CUS_Str_DigCapData, ",")
            For i = 0 To UBound(DSSC_Out_DecompseByComma)
                DSSC_Out_DecompseByColon = Split(DSSC_Out_DecompseByComma(i), ":")
                If UBound(DSSC_Out_DecompseByColon) > 0 Then
                    If ParseStringByBits = "" And ParseStringForTestName = "" Then
                        ParseStringByBits = DSSC_Out_DecompseByColon(0)
                        ParseStringForTestName = DSSC_Out_DecompseByColon(1)
                        If UBound(DSSC_Out_DecompseByColon) = 2 Then    '' Dictionary
                            ParseStringForDirectionary = DSSC_Out_DecompseByColon(2) & ","
                            ParseStringForGrayCode = ","
                        ElseIf UBound(DSSC_Out_DecompseByColon) = 3 Then    '' Dictionary and GrayCode
                            ParseStringForDirectionary = DSSC_Out_DecompseByColon(2) & ","
                            ParseStringForGrayCode = DSSC_Out_DecompseByColon(3) & ","
                        Else
                            ParseStringForDirectionary = ","
                            ParseStringForGrayCode = ","
                        End If
                    Else
                        ParseStringByBits = ParseStringByBits & "," & DSSC_Out_DecompseByColon(0)
                        ParseStringForTestName = ParseStringForTestName & "," & DSSC_Out_DecompseByColon(1)
                        
                        If b_ParseForDirectionary_Switch = False Then
                            b_ParseForDirectionary_Switch = True
                        Else
                            ParseStringForDirectionary = ParseStringForDirectionary & ","
                        End If
                        
                        If b_ParseForGrayCode_Switch = False Then
                            b_ParseForGrayCode_Switch = True
                        Else
                            ParseStringForGrayCode = ParseStringForGrayCode & ","
                        End If
                        
                        If UBound(DSSC_Out_DecompseByColon) = 2 Then    '' Dictionary
                            ParseStringForDirectionary = ParseStringForDirectionary & DSSC_Out_DecompseByColon(2)
                        End If
                        
                        If UBound(DSSC_Out_DecompseByColon) = 3 Then    '' Dictionary
                            ParseStringForDirectionary = ParseStringForDirectionary & DSSC_Out_DecompseByColon(2)
                            ParseStringForGrayCode = ParseStringForGrayCode & DSSC_Out_DecompseByColon(3)
                        End If
                    
                    End If
                End If
            Next i
            
            ''20161220-Remove comma in the last of string
            If Right(ParseStringForTestName, 1) = "," Then
                ParseStringForTestName = Left(ParseStringForTestName, (Len(ParseStringForTestName) - 1))
            End If
            If Right(ParseStringForDirectionary, 1) = "," Then
                ParseStringForDirectionary = Left(ParseStringForDirectionary, (Len(ParseStringForDirectionary) - 1))
            End If
            
            ParseStringByBits = "DSSC_OUT," & ParseStringByBits
            DecomposeTestName = Split(ParseStringForTestName, ",")
            DecomposeDirectionary = Split(ParseStringForDirectionary, ",")
            DecomposeGrayCode = Split(ParseStringForGrayCode, ",")
        Else
            ParseStringByBits = CUS_Str_DigCapData
        End If
        
        If Right(ParseStringByBits, 1) = "," Then
            ParseStringByBits = Left(ParseStringByBits, (Len(ParseStringByBits) - 1))
        End If
        DecomposeParseDigCapBit = Split(ParseStringByBits, ",")
        Dim StrParseDigCapBit As String
        
        For i = 1 To UBound(DecomposeParseDigCapBit)
            If i = 1 Then
                StrParseDigCapBit = DecomposeParseDigCapBit(i)
            Else
                StrParseDigCapBit = StrParseDigCapBit & "," & DecomposeParseDigCapBit(i)
            End If
        Next i
        DecomposeParseDigCapBit = Split(StrParseDigCapBit, ",")
        
        ReDim FlexibleConvertedDataWf(UBound(DecomposeParseDigCapBit)) As New DSPWave
        ''20160823-Store binary dsp wave after processed by DSSC_OUT
        Dim DSPWave_Binary() As New DSPWave
        ReDim DSPWave_Binary(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim DSPWave_GrayCode() As New DSPWave
        ReDim DSPWave_GrayCode(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim DSPWave_GrayCodeDec() As New DSPWave
        ReDim DSPWave_GrayCodeDec(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim StartIndex As Long
        StartIndex = 0
        
        ''20161230-Add copy and site loop to pass data
        For Each site In TheExec.sites
            SourceBitStrmWf = OutDspWave.Copy
        Next site
        Dim width_Wf  As New DSPWave, OutWf As New DSPWave ', OutBinWf() As New DSPWave
       ' ReDim OutBinWf(UBound(DecomposeParseDigCapBit))
        width_Wf.CreateConstant 0, UBound(DecomposeParseDigCapBit) + 1   'Create space for DSP
        'OutWf.CreateConstant 0, UBound(DecomposeParseDigCapBit) + 1
        Dim DecomposeParseDigCapBit_long() As Long
        ReDim DecomposeParseDigCapBit_long(UBound(DecomposeParseDigCapBit))
        
        
        For i = 0 To UBound(DecomposeParseDigCapBit)
            DecomposeParseDigCapBit_long(i) = CLng(DecomposeParseDigCapBit(i))  'deliver data to dsp array
        Next i
        
        For Each site In TheExec.sites
            width_Wf.Data = DecomposeParseDigCapBit_long  'deliver data to dsp array
        Next site
        
        
        '///////////////////////////////////// for Sicily///////////////
        
        ' Specail calculation for skua ddr cz mdll
'        If UCase(CUS_Str_MainProgram) = "CZ_MDLL" Then
'            Dim sl_decrease As New SiteLong
'            Dim sl_unique As New SiteLong
'            Dim sl_maxdiff As New SiteLong
'            rundsp.Split_Dspwave_CZ_MDLL SourceBitStrmWf, width_Wf, OutWf, sl_decrease, sl_unique, sl_maxdiff
'            GoTo skip
'        End If
        
        If UCase(CUS_Str_MainProgram) Like "*TRIM_CODE_MODE*" Then
            Dim CUS_Str_Split() As String
            Dim Target As Long
            Dim trim_name As String
            Dim calc_data As New DSPWave
            Dim delta_value As New DSPWave
            Dim delta_value2 As New DSPWave
            Dim target_var As New SiteDouble
            Dim target_var2 As New SiteDouble
            Dim F_cal_binary As New DSPWave
            Dim bit_number As Long
            calc_data.CreateConstant 0, 64, DspDouble
            delta_value.CreateConstant 0, 64, DspDouble
            delta_value2.CreateConstant 0, 64, DspDouble
            
            If InStr(CUS_Str_MainProgram, ")") <> 0 Then
                CUS_Str_Split = Split(Split(CUS_Str_MainProgram, "(")(0), "_")
            Else
                CUS_Str_Split = Split(CUS_Str_MainProgram, "_")
            End If
            
            If UBound(CUS_Str_Split) > 5 Then
                Dim Target_low As Long
                Dim Target_high As Long
                Target_low = CUS_Str_Split(5)
                Target_high = CUS_Str_Split(6)
            Else
                Target = CUS_Str_Split(5)
            End If
            
            trim_name = CUS_Str_Split(4)
            bit_number = CUS_Str_Split(3)
            Dim storeDSP As New DSPWave
            Dim storeDSP1 As New DSPWave
            storeDSP.CreateConstant 0, bit_number, DspDouble
            storeDSP1.CreateConstant 0, bit_number, DspDouble
            Dim sda_measuredata(26) As New SiteDouble
            Dim sda_measuredata1(26) As New SiteDouble
            
            
            
            If InStr(UCase(trim_name), UCase("PCIEREFPLL")) <> 0 Or InStr(UCase(trim_name), UCase("PCIERPLL")) <> 0 Then
                rundsp.Split_Dspwave_PCIEREFPLL SourceBitStrmWf, width_Wf, OutWf, calc_data, delta_value, target_var, Target, 1, storeDSP                                           ', OutBinWf
            ElseIf InStr(UCase(trim_name), UCase("PCIEPLL")) <> 0 Then
                calc_data.CreateConstant 0, 128, DspDouble
                rundsp.Split_Dspwave_PCIETXPLL SourceBitStrmWf, width_Wf, OutWf, calc_data, delta_value, target_var, Target, 1, storeDSP, delta_value2, target_var2, storeDSP1                                               ', OutBinWf
            ElseIf InStr(UCase(trim_name), UCase("CIO")) <> 0 Then
                
                Dim target_var3 As New SiteDouble
                Dim target_var4 As New SiteDouble
                Dim storeDSP2 As New DSPWave
                Dim storeDSP3 As New DSPWave
                Dim StoreDSP4 As New DSPWave
                Dim calc_data128 As New DSPWave
                Dim delta_value128 As New DSPWave
                calc_data128.CreateConstant 0, 128, DspDouble
                delta_value128.CreateConstant 0, 128, DspDouble
                storeDSP2.CreateConstant 0, bit_number, DspDouble
                storeDSP3.CreateConstant 0, bit_number, DspDouble
                StoreDSP4.CreateConstant 0, bit_number, DspDouble
                rundsp.Split_Dspwave_CIOPLL SourceBitStrmWf, width_Wf, OutWf, calc_data128, delta_value128, target_var, target_var2, target_var3, target_var4, Target_low, Target_high, 1, storeDSP, storeDSP2, storeDSP3, StoreDSP4
            ElseIf InStr(UCase(trim_name), UCase("CAUSPLL")) <> 0 Then
                
                Dim Outwf_T125 As New DSPWave
                Dim Outwf_T225 As New DSPWave
                Dim buf_caus1 As New SiteDouble
                Dim unbuf_caus1 As New SiteDouble
                Dim Diff_buf_caus1 As New SiteDouble
                Dim buf_caus2 As New SiteDouble
                Dim unbuf_caus2 As New SiteDouble
                Dim Diff_buf_caus2 As New SiteDouble
                
                buf_caus1 = GetStoredMeasurement("V1")
                unbuf_caus1 = GetStoredMeasurement("V2")
                Diff_buf_caus1 = buf_caus1.Subtract(unbuf_caus1)

                buf_caus2 = GetStoredMeasurement("V28")
                unbuf_caus2 = GetStoredMeasurement("V29")
                Diff_buf_caus2 = buf_caus2.Subtract(unbuf_caus2)
                
                For i = 3 To UBound(sda_measuredata)
                    sda_measuredata(i) = GetStoredMeasurement("V" + CStr(i + 1))
                    sda_measuredata(i) = sda_measuredata(i).Subtract(Diff_buf_caus1)
                Next
                For i = 3 To UBound(sda_measuredata1)
                    sda_measuredata1(i) = GetStoredMeasurement("V" + CStr(i + 28))
                    sda_measuredata1(i) = sda_measuredata1(i).Subtract(Diff_buf_caus2)
                Next
                
                rundsp.Split_Dspwave_CIO SourceBitStrmWf, width_Wf, OutWf, Outwf_T125, Outwf_T225
                rundsp.Split_Dspwave_CIOCALC Outwf_T125, sda_measuredata(3), sda_measuredata(4), sda_measuredata(5), sda_measuredata(6), sda_measuredata(7), sda_measuredata(8), _
                                                    sda_measuredata(9), sda_measuredata(10), sda_measuredata(11), sda_measuredata(12), sda_measuredata(13), sda_measuredata(14), sda_measuredata(15), _
                                                    sda_measuredata(16), sda_measuredata(17), sda_measuredata(18), sda_measuredata(19), sda_measuredata(20), sda_measuredata(21), sda_measuredata(22), _
                                                    sda_measuredata(23), sda_measuredata(24), sda_measuredata(25), sda_measuredata(26), storeDSP
                rundsp.Split_Dspwave_CIOCALC Outwf_T225, sda_measuredata1(3), sda_measuredata1(4), sda_measuredata1(5), sda_measuredata1(6), sda_measuredata1(7), sda_measuredata1(8), _
                                                    sda_measuredata1(9), sda_measuredata1(10), sda_measuredata1(11), sda_measuredata1(12), sda_measuredata1(13), sda_measuredata1(14), sda_measuredata1(15), _
                                                    sda_measuredata1(16), sda_measuredata1(17), sda_measuredata1(18), sda_measuredata1(19), sda_measuredata1(20), sda_measuredata1(21), sda_measuredata1(22), _
                                                    sda_measuredata1(23), sda_measuredata1(24), sda_measuredata1(25), sda_measuredata1(26), storeDSP1

            
            ElseIf InStr(UCase(trim_name), UCase("AUS")) <> 0 Then
                Dim stra_storename() As String
                Dim Outwf_T1 As New DSPWave
                
                Dim buf As New SiteDouble
                Dim unbuf As New SiteDouble
                Dim Diff_buf As New SiteDouble
                

                
                stra_storename = Split(Left(Split(CUS_Str_MainProgram, "(")(1), Len(Split(CUS_Str_MainProgram, "(")(1)) - 1), ",")
                'Dim sda_measuredata(26) As New SiteDouble
                buf = GetStoredMeasurement(stra_storename(0))
                unbuf = GetStoredMeasurement(stra_storename(1))
                Diff_buf = buf.Subtract(unbuf)
                
                For i = 3 To UBound(sda_measuredata)
                    sda_measuredata(i) = GetStoredMeasurement(stra_storename(i))
                    sda_measuredata(i) = sda_measuredata(i).Subtract(Diff_buf)
                Next

                rundsp.Split_Dspwave_AUS SourceBitStrmWf, width_Wf, OutWf, Outwf_T1, sda_measuredata(3), sda_measuredata(4), sda_measuredata(5), sda_measuredata(6), sda_measuredata(7), sda_measuredata(8), _
                                                    sda_measuredata(9), sda_measuredata(10), sda_measuredata(11), sda_measuredata(12), sda_measuredata(13), sda_measuredata(14), sda_measuredata(15), _
                                                    sda_measuredata(16), sda_measuredata(17), sda_measuredata(18), sda_measuredata(19), sda_measuredata(20), sda_measuredata(21), sda_measuredata(22), _
                                                    sda_measuredata(23), sda_measuredata(24), sda_measuredata(25), sda_measuredata(26), storeDSP

            End If
        ElseIf UCase(CUS_Str_MainProgram) Like "*SIGNEDGRAY*UNSIGNEDGRAY*2SCOMPLEMENT*SIGNEDBIN*" Then
            Call Split_GrayDSP_2sComplementDSP_to_Dec(CUS_Str_MainProgram, DecomposeParseDigCapBit, DecomposeTestName, SourceBitStrmWf, width_Wf, OutWf)
        ElseIf UCase(CUS_Str_MainProgram) = "2SCOMPLEMENT" Then
            rundsp.Split_2SComplementDSPWave_To_SignDec SourceBitStrmWf, width_Wf, OutWf

        Else
          rundsp.Split_Dspwave SourceBitStrmWf, width_Wf, OutWf                    ', OutBinWf
        End If
        
skip:
        '/////////////////////////////////////////////////////////////////////////////////
        
        Dim Dec_val As New SiteDouble
        
        Dim Dec_val_ary() As New SiteDouble
        ReDim Dec_val_ary(UBound(DecomposeParseDigCapBit)) As New SiteDouble
        For i = 0 To UBound(DecomposeParseDigCapBit)
            For Each site In TheExec.sites
                FlexibleConvertedDataWf(i).CreateConstant 0, 1
                FlexibleConvertedDataWf(i).Element(0) = OutWf.ElementLite(i)
                
                If ParseStringForDirectionary <> "" Then                        ' save in dictionary for CalcEqn
                    If UBound(DecomposeDirectionary) >= i Then
                        If DecomposeDirectionary(i) <> "" Then
                            Dec_val_ary(i) = FlexibleConvertedDataWf(i).ElementLite(0)
                        End If
                    End If
                End If

            Next site
            If ParseStringForDirectionary <> "" Then                        ' save in dictionary for CalcEqn
                If UBound(DecomposeDirectionary) >= i Then
                    If DecomposeDirectionary(i) <> "" Then
                        Call AddStoredData(DecomposeDirectionary(i) & "_para", Dec_val_ary(i))
                    End If
                End If
            End If
            
            
        Next i
       
        ''20160823-Modify dsp function to add one input argument to process DSPwave with binary format and use Directionary to store it.
        For i = 0 To UBound(DecomposeParseDigCapBit)
'            rundsp.FlexibleBitWf2Arry SourceBitStrmWf, StartIndex, CLng(DecomposeParseDigCapBit(i)), FlexibleConvertedDataWf(i), DSPWave_Binary(i)
            
            ''20160823-Store binary DSP wave by using Directionary
''            If DecomposeDirectionary(i) <> "" Then
''                Call AddStoredCaptureData(DecomposeDirectionary(i), DSPWave_Binary(i))
''            End If
            
            If UCase(DecomposeGrayCode(i)) = "GRAYCODE" Then
''                DSPWave_GrayCode(i).CreateConstant 0, DecomposeParseDigCapBit(i), DspLong
''                DSPWave_GrayCodeDec(i).CreateConstant 0, 1, DspLong
                
                Call rundsp.Transfer2GrayCode(DSPWave_Binary(i), DSPWave_GrayCode(i), DSPWave_GrayCodeDec(i))
                
            End If
            
            StartIndex = StartIndex + DecomposeParseDigCapBit(i)
        Next i
        
        ''20161215-Check the dictionary name, re-combine them to one dsp wave and store to dictionary if there have the same dictionary name cross multi-segment (over 24 bit).
        '' Separate dsp wave to different segment if over 24bits, this is for cover STDF display truncation issue.
        Dim CombineDSPBit2Dict As New Dictionary
        Dim KeyName As String
        
        If ParseStringForDirectionary <> "" Then
            CombineDSPBit2Dict.RemoveAll
            For i = 0 To UBound(DecomposeDirectionary)
                If DecomposeDirectionary(i) = "" Then
                    KeyName = "EMPTYSPACE_DICT_" & i
                Else
                    KeyName = LCase(DecomposeDirectionary(i))
                End If
                If i = 0 Then
                    CombineDSPBit2Dict.Add KeyName, CLng(DecomposeParseDigCapBit(i))
    
                Else
                    If CombineDSPBit2Dict.Exists(KeyName) Then
                        CombineDSPBit2Dict.Item(KeyName) = CombineDSPBit2Dict.Item(KeyName) + CLng(DecomposeParseDigCapBit(i))
                    Else
                        CombineDSPBit2Dict.Add KeyName, CLng(DecomposeParseDigCapBit(i))
                    End If
                End If
            Next i
            
            Dim CombineKeys() As Variant
            CombineKeys() = CombineDSPBit2Dict.Keys()
            
            StartIndex = 0
            ReDim AddToDict_DSP_Dec(CombineDSPBit2Dict.Count - 1) As New DSPWave
            ReDim AddToDict_DSP_Bin(CombineDSPBit2Dict.Count - 1) As New DSPWave
            Dim FinalLength As Long
            
            For i = 0 To CombineDSPBit2Dict.Count - 1
                FinalLength = CombineDSPBit2Dict.Item(CombineKeys(i))
''                rundsp.FlexibleBitWf2Arry SourceBitStrmWf, StartIndex, FinalLength, AddToDict_DSP_Dec(i), AddToDict_DSP_Bin(i)
                If InStr(CombineKeys(i), "EMPTYSPACE_DICT_") <> 0 Then
                Else
                    For Each site In TheExec.sites
                        AddToDict_DSP_Bin(i) = OutDspWave.Select(StartIndex, , FinalLength).Copy '.ConvertStreamTo(tldspSerial, FinalLength, 0, Bit0IsMsb)
                    Next site
                    Call AddStoredCaptureData(CStr(CombineKeys(i)), AddToDict_DSP_Bin(i))
                End If
                StartIndex = StartIndex + FinalLength
            Next i
        End If
        
        '' Debug use
        Dim BinaryCodeString As String
        Dim GrayCodeString As String
''        Dim j As Long
        For Each site In TheExec.sites
            For i = 0 To UBound(DecomposeParseDigCapBit)
                If UCase(DecomposeGrayCode(i)) = "GRAYCODE" Then
                    BinaryCodeString = ""
                    GrayCodeString = ""
                    For j = 0 To DSPWave_Binary(i).SampleSize - 1
                        BinaryCodeString = BinaryCodeString & DSPWave_Binary(i)(site).Element(j)
                        GrayCodeString = GrayCodeString & DSPWave_GrayCode(i)(site).Element(j)
                    Next j
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " DSSC_OUT part " & i & " binary code = " & BinaryCodeString)
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " DSSC_OUT part " & i & "   gray code = " & GrayCodeString)
                End If
            Next i
        Next site
                
        '' 20160317 - Test limit for DSSC_OUT
        If b_DSSC_Out_InvolveTestName = True Then '' Test limit with test name
            For i = 0 To UBound(DecomposeTestName)
                If LCase(DecomposeTestName(i)) = "skip" Then
                Else
                    TestLimitWithTestName.AddPin (DecomposeTestName(i) & "_" & i)
                    If UCase(DecomposeGrayCode(i)) = "GRAYCODE" Then
                        TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i).Value = DSPWave_GrayCodeDec(i).Element(0)
                    Else
                        TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i).Value = FlexibleConvertedDataWf(i).Element(0)
                    End If
                    If BypassAllDigCapTestLimit = False Then
                        If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
                            TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, DecomposeTestName(i), CInt(i), 0)
                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%.0f"
                    
                        ElseIf MTR_CusDigCap <> "" And UCase(MTR_CusDigCap) = "CUS_DIGCAP_VIN" Then
                            
                            TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, DecomposeTestName(i), CInt(i), 0)
                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=TestInstanceName & "_" & DecomposeTestName(i) & "_" & MTR_VIN & "_" & i, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                            MTR_CusDigCap = ""
                        
                            '  ElseIf TPModeAsCharz_GLB = True Then
                            ' 'CZ TP name force flow
                      '      TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                        
                        'ElseIf CUS_Str_MainProgram = "TMPS_BV" Then  ''TMPS_BV
                            
                            'TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, DecomposeTestName(i), CInt(i), 0)
                            'TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 1552, 2 ^ DecomposeParseDigCapBit(i) - 1, TName:=TestNameInput, ForceResults:=tlForceNone, ScaleType:=scaleNoScaling, FormatStr:="%.0f"
                        ElseIf (CUS_Str_MainProgram = "TTROUT" Or InStr(LCase(CUS_Str_MainProgram), "trim_code_mode") <> 0) And gl_Disable_HIP_debug_log = True Then    ''TMPS_BV trim_code_mode
                            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
                        
                        Else
                            TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, DecomposeTestName(i), CInt(i), 0, , , , tlForceNone)
                            
                            If UCase(CUS_Str_MainProgram) Like UCase("*PRINT_DIGCAP_BINARY_CODE*") Then
                                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
                            Else
                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                        End If
                            
                            ''''''''''''''''''''20191101 Ryder/CT debug
                            If UCase(CUS_Str_MainProgram) Like UCase("*PRINT_DIGCAP_BINARY_CODE*") Then
                                
'                                Dim dspBin As New DSPWave
'                                Dim dspBinStr As String
'                                Dim dspBinStr_counter As New SiteLong: dspBinStr_counter = 0
                                For Each site In TheExec.sites
                                    dspBinStr = ""
                                    
'                                    dspBin = FlexibleConvertedDataWf(i).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, DecomposeParseDigCapBit_long(i), 0, Bit0IsMsb)
'                                    For j = 0 To DecomposeParseDigCapBit_long(i) - 1
'                                         dspBinStr = dspBinStr & CStr(dspBin.Element(j))
'                                    Next j
                                    
                                    For j = 0 To DecomposeParseDigCapBit_long(i) - 1
                                        dspBinStr = dspBinStr & OutDspWave(site).Element(dspBinStr_counter(site))
                                        dspBinStr_counter(site) = dspBinStr_counter(site) + 1
                                    Next j
                                                        
                                    
                                    OutputTname_format = Split(TestNameInput, "_")
                                    OutputTname_format(7) = "Binary"
                                    TestNameInput = Merge_TName(OutputTname_format)
                                    TestNameInput = Mid(TestNameInput, 1, Len(TestNameInput) - 1)
                                    
                                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " " & TestNameInput & ", BinVal[Lsb:Msb] = " & dspBinStr)
                                    'If gl_Disable_HIP_debug_log = False Then theexec.Datalog.WriteComment ("Site_" & Site & " " & TestNameInput & ", DecVal = " & FlexibleConvertedDataWf(i).Element(0) & ", BinVal[Lsb:Msb] = " & dspBinStr)
                                Next site
'                                theexec.Flow.TestLimit CLng(dspBinStr), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%.0f"
                            End If
                            ''''''''''''''''''''20191101 Ryder/CT debug
                        
                        
                        End If
                    End If
                End If
                                Call Update_BC_PassFail_Flag
            Next i
        Else
            If BypassAllDigCapTestLimit = False Then
                For i = 0 To UBound(FlexibleConvertedDataWf)
                    TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, DecomposeTestName(i), CInt(i), 0, , , , tlForceNone)
                    TheExec.Flow.TestLimit FlexibleConvertedDataWf(i).Element(0), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:="DSSC_OUT_Code_" & i, Tname:="Digcap_" & i & "_DSSC_OUT_" & CStr(i - 1), ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                Next
            End If
        End If
    End If
    
    
    If UCase(CUS_Str_MainProgram) Like "*TRIM_CODE_MODE*" Then
    
        If InStr(UCase(trim_name), UCase("PCIEREFPLL")) <> 0 Or InStr(UCase(trim_name), UCase("PCIERPLL")) <> 0 Then
            Call PLL_calibration_calc_MOD_PCIEREFPLL(CUS_Str_MainProgram, calc_data, target_var, storeDSP)
        ElseIf InStr(UCase(trim_name), UCase("PCIEPLL")) <> 0 Then
            Call PLL_calibration_calc_MOD_PCIETXPLL(CUS_Str_MainProgram, calc_data, target_var, storeDSP, target_var2, storeDSP1)
        ElseIf InStr(UCase(trim_name), UCase("CAUSPLL")) <> 0 Then
            Call PLL_calibration_calc_MOD_AUS(CUS_Str_MainProgram, Outwf_T125, sda_measuredata, storeDSP, 0)
            Call PLL_calibration_calc_MOD_AUS(CUS_Str_MainProgram, Outwf_T225, sda_measuredata1, storeDSP1, 1)
        ElseIf InStr(UCase(trim_name), UCase("CIO")) <> 0 Then
            Call PLL_calibration_calc_MOD_CIO(CUS_Str_MainProgram, calc_data128, target_var, target_var2, storeDSP, storeDSP2, target_var3, target_var4, storeDSP3, StoreDSP4)
        ElseIf InStr(UCase(trim_name), UCase("AUS")) <> 0 Then
            Call PLL_calibration_calc_MOD_AUS(CUS_Str_MainProgram, Outwf_T1, sda_measuredata, storeDSP)
        End If
    ElseIf UCase(TestInstanceName) Like "*OFFCAL*" Then
        'Call MTRTMPS_OffSet_Cal(OutWf, "MTR_TSENSE_OFFSET_" & gl_FlowForLoop_DigSrc_SweepCode, 18)
    ElseIf UCase(TestInstanceName) Like "*GAINCAL*" Then
'        Call MTRTMPS_Gain_AVG(OutWf, "MTR_TSENSE_GAIN_" & gl_FlowForLoop_DigSrc_SweepCode, 18)
        'Call MTRTMPS_DSSCOUT_AVG(OutWf, "MTR_TSENSE_GAIN_" & gl_FlowForLoop_DigSrc_SweepCode)
'        Call MTRTMPS_Gain_Cal(OutWf, "MTRTMPS_OFFSET_MEAN", "MTR_TSENSE_GAIN")
    ElseIf UCase(CUS_Str_MainProgram) Like "DSSC_OUT_AVG*" Then
        Dim Dictionary_Name As String: Dictionary_Name = Split(CUS_Str_MainProgram, ":")(1)
        'Call MTRTMPS_DSSCOUT_AVG(OutWf, Dictionary_Name)
'    End If
'
'    If UCase(CUS_Str_MainProgram) = "CZ_MDLL" Then
'        Call ddr_cz_mdll_testlimit(sl_decrease, sl_unique, sl_maxdiff)
    End If
'
    Call CZ_TNum_Increment

    'Call CZ_TNum_Increment
    
    Exit Function
err:
        If AbortTest Then Exit Function Else Resume Next

End Function
Public Function PLL_calibration_calc(ByRef calib_code() As DSPWave, CUS_Str_MainProgram As String) As Long


Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim temp1_dict As New DSPWave
Dim temp2_dict As New DSPWave
Dim calc_data() As New SiteDouble
Dim temp_testname_bin As Long
Dim temp_testname_dec As Long
Dim testname_str() As String
Dim delta_value() As New SiteDouble
Dim target_var As New SiteDouble
Dim temp_delta_value As Integer
Dim temp_cal_code As New DSPWave
Dim RefPLL_calibration_code As New DSPWave
Dim TXPLL_calibration_code As New DSPWave
Dim DPTXPLL_calibration_code As New DSPWave
Dim calibration_target() As String
Dim calibration_target_value As Long

Dim OutputTname_format() As String
Dim TestNameInput As String
Dim site As Variant
calibration_target = Split(CUS_Str_MainProgram, "_")
calibration_target_value = CLng(calibration_target(4))


ReDim testname_str(5)
ReDim calc_data(31)
ReDim delta_value(31)

        ''''calc and print in datalog
        temp1_dict.CreateConstant 0, 1, DspLong
        temp2_dict.CreateConstant 0, 1, DspLong
        RefPLL_calibration_code.CreateConstant 0, 5, DspLong
        temp_cal_code.CreateConstant 0, 1, DspLong
        
            For i = 0 To (UBound(calib_code()) + 1) / 2 - 1
                    For Each site In TheExec.sites
                        temp1_dict.Element(0) = calib_code(i).Element(0)
                        temp2_dict.Element(0) = calib_code(i + 32).Element(0)
                        calc_data(i) = (temp2_dict.Element(0) + temp1_dict.Element(0)) / 2
                    Next site
    
                '''''dec to bin testname
                    temp_testname_dec = i
                        For j = 0 To 4
                          temp_testname_bin = temp_testname_dec Mod 2
                          temp_testname_dec = Fix(temp_testname_dec / 2)
                          testname_str(j) = CStr(temp_testname_bin)
                        Next j
                        testname_str(5) = testname_str(4) & testname_str(3) & testname_str(2) & testname_str(1) & testname_str(0)
                    
                        TestNameInput = Report_TName_From_Instance("C", "", "F_" & testname_str(5), i, 0)
                
                If gl_Disable_HIP_debug_log = False Then
                    TheExec.Flow.TestLimit resultVal:=calc_data(i), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%.1f"
                End If
            Next i

            '''' compare the target
            For Each site In TheExec.sites
             temp_delta_value = 5000
                 For k = 0 To (UBound(calib_code()) + 1) / 2 - 1
                    If UCase(TheExec.DataManager.instanceName) Like "*PCIEREFPLL*" Then
                    delta_value(k) = Abs(calibration_target_value - calc_data(k))
                    ElseIf UCase(TheExec.DataManager.instanceName) Like "*PCIETXPLL*" Then
                    delta_value(k) = Abs(calibration_target_value - calc_data(k))
                    ElseIf UCase(TheExec.DataManager.instanceName) Like "*DPTXPLL*" Then
                    delta_value(k) = Abs(calibration_target_value - calc_data(k))
                    End If
                'search min delta
                    If delta_value(k) < temp_delta_value Then
                        temp_delta_value = delta_value(k)
                        target_var = k
                    End If
                 Next k
             Next site
             
             
            If UCase(TheExec.DataManager.instanceName) Like "*PCIEREFPLL*" Then
                
                TestNameInput = Report_TName_From_Instance("C", "", "PCIE_REFPLL_Fcal_code", 0, 0)
                TheExec.Flow.TestLimit resultVal:=target_var, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                
                 'store to dictionary
                    For Each site In TheExec.sites
                    temp_cal_code.Element(0) = target_var
                    Next site
                 Call HardIP_Dec2Bin(RefPLL_calibration_code, temp_cal_code, 5)
                 Call AddStoredCaptureData("PCIE_REFPLL_FCAL_BYPASS", RefPLL_calibration_code)
                 
            ElseIf UCase(TheExec.DataManager.instanceName) Like "*PCIETXPLL*" Then
            
                TestNameInput = Report_TName_From_Instance("C", "", "PCIE_TXPLL_Fcal_code", 0, 0)
                
                TheExec.Flow.TestLimit resultVal:=target_var, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                
                 'store to dictionary
                    For Each site In TheExec.sites
                    temp_cal_code.Element(0) = target_var
                    Next site
                 Call HardIP_Dec2Bin(TXPLL_calibration_code, temp_cal_code, 5)
                 Call AddStoredCaptureData("PCIE_TXPLL_FCAL_BYPASS", TXPLL_calibration_code)
                 
            ElseIf UCase(TheExec.DataManager.instanceName) Like "*DPTXPLL*" Then
            
                TestNameInput = Report_TName_From_Instance("C", "", "DP_TXPLL_Fcal_code", 0, 0)
                TheExec.Flow.TestLimit resultVal:=target_var, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                
                 'store to dictionary
                    For Each site In TheExec.sites
                    temp_cal_code.Element(0) = target_var
                    Next site
                 Call HardIP_Dec2Bin(DPTXPLL_calibration_code, temp_cal_code, 5)
                 Call AddStoredCaptureData("DPTX_PCIEPLL_FCAL_BYPASS", DPTXPLL_calibration_code)
            End If
   
End Function
    
Public Function PLL_calibration_calc_MOD_PCIETXPLL(CUS_Str_MainProgram As String, calc_data As DSPWave, target_var As SiteDouble, storeDSP As DSPWave, target_var2 As SiteDouble, storeDSP2 As DSPWave) As Long


Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim temp1_dict As New DSPWave
Dim temp2_dict As New DSPWave

Dim temp_testname_bin As Long
Dim temp_testname_dec As Long
Dim testname_str() As String
Dim delta_value() As New SiteDouble
Dim temp_delta_value As Integer
Dim temp_cal_code As New DSPWave
Dim RefPLL_calibration_code As New DSPWave
Dim TXPLL_calibration_code As New DSPWave
Dim DPTXPLL_calibration_code As New DSPWave
Dim calibration_target() As String
Dim calibration_target_value As Long

Dim OutputTname_format() As String
Dim TestNameInput As String
Dim bitnumber As Long
Dim calibration_name As String

calibration_target = Split(CUS_Str_MainProgram, "_")
calibration_target_value = CLng(calibration_target(5))
calibration_name = calibration_target(4)
bitnumber = calibration_target(3)


ReDim testname_str(bitnumber)

        ''''calc and print in datalog
        temp1_dict.CreateConstant 0, 1, DspLong
        temp2_dict.CreateConstant 0, 1, DspLong
        RefPLL_calibration_code.CreateConstant 0, 5, DspLong
        temp_cal_code.CreateConstant 0, 1, DspLong
        
            Dim active_site As Variant
            Dim site As Variant
            For Each site In TheExec.sites.Active
                active_site = site
            Next site

            For i = 0 To (calc_data(active_site).SampleSize) - 2 'TY add 2018/08/29

            '''''dec to bin testname
                temp_testname_dec = i + 1
                    For j = 0 To bitnumber - 1
                      temp_testname_bin = temp_testname_dec Mod 2
                      temp_testname_dec = Fix(temp_testname_dec / 2)
                      testname_str(j) = CStr(temp_testname_bin)
                    Next j
                    testname_str(bitnumber) = ""
                    For j = 1 To bitnumber
                        testname_str(bitnumber) = testname_str(bitnumber) & testname_str(bitnumber - j)
                    Next j
                    'TestNameInput = Report_TName_From_Instance("C", "X", "F" & testname_str(bitnumber), i, 1)
                
                   If gl_Disable_HIP_debug_log = False Then
                    
                        For Each site In TheExec.sites.Active
                            TheExec.Datalog.WriteComment "Site = " & site & "    " & "F_" & testname_str(bitnumber) & "========>" & calc_data.Element(i)
                        Next site
                        
                    End If
                    
                    If i = ((calc_data(active_site).SampleSize) - 2) / 2 - 1 Then TheExec.Datalog.WriteComment ""
                    If i = ((calc_data(active_site).SampleSize) - 2) / 2 - 1 Then i = i + 1
            Next i
             
            TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
            
            If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=target_var, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
            If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=target_var2, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
            
            
             'store to dictionary
            Dim storename() As String
             'storename = LCase(calibration_name) & "_cal"

            storename = Split(Split(Instance_Data.CUS_Str_DigSrcData, ";")(0), ":")

            If LCase(storename(0)) Like "trimcodestorename" And UBound(storename) = 1 Then
                Call AddStoredCaptureData(storename(1), storeDSP)
            End If


            storename = Split(Split(Instance_Data.CUS_Str_DigSrcData, ";")(1), ":")

            If LCase(storename(0)) Like "trimcodestorename" And UBound(storename) = 1 Then
                Call AddStoredCaptureData(storename(1), storeDSP2)
            End If
               
            
   
End Function
Public Function PLL_calibration_calc_MOD_PCIEREFPLL(CUS_Str_MainProgram As String, calc_data As DSPWave, target_var As SiteDouble, storeDSP As DSPWave) As Long


Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim temp1_dict As New DSPWave
Dim temp2_dict As New DSPWave

Dim temp_testname_bin As Long
Dim temp_testname_dec As Long
Dim testname_str() As String
Dim delta_value() As New SiteDouble
Dim temp_delta_value As Integer
Dim temp_cal_code As New DSPWave
Dim RefPLL_calibration_code As New DSPWave
Dim TXPLL_calibration_code As New DSPWave
Dim DPTXPLL_calibration_code As New DSPWave
Dim calibration_target() As String
Dim calibration_target_value As Long

Dim OutputTname_format() As String
Dim TestNameInput As String
Dim bitnumber As Long
Dim calibration_name As String

calibration_target = Split(CUS_Str_MainProgram, "_")
calibration_target_value = CLng(calibration_target(5))
calibration_name = calibration_target(4)
bitnumber = calibration_target(3)


ReDim testname_str(bitnumber)

        ''''calc and print in datalog
        temp1_dict.CreateConstant 0, 1, DspLong
        temp2_dict.CreateConstant 0, 1, DspLong
        RefPLL_calibration_code.CreateConstant 0, 5, DspLong
        temp_cal_code.CreateConstant 0, 1, DspLong
        
            Dim active_site As Variant
            Dim site As Variant
            For Each site In TheExec.sites.Active
                active_site = site
            Next site

            For i = 0 To (calc_data(active_site).SampleSize) - 2 'TY add 2018/08/29

            '''''dec to bin testname
                temp_testname_dec = i + 1
                    For j = 0 To bitnumber - 1
                      temp_testname_bin = temp_testname_dec Mod 2
                      temp_testname_dec = Fix(temp_testname_dec / 2)
                      testname_str(j) = CStr(temp_testname_bin)
                    Next j
                    testname_str(bitnumber) = ""
                    For j = 1 To bitnumber
                        testname_str(bitnumber) = testname_str(bitnumber) & testname_str(bitnumber - j)
                    Next j
                    'TestNameInput = Report_TName_From_Instance("C", "X", "F" & testname_str(bitnumber), i, 1)
                
                   If gl_Disable_HIP_debug_log = False Then
                    
                        For Each site In TheExec.sites.Active
                            TheExec.Datalog.WriteComment "Site = " & site & "    " & "F_" & testname_str(bitnumber) & "========>" & calc_data.Element(i)
                        Next site
                        
                    End If
            Next i
             
            TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
            
            If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=target_var, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
            
             'store to dictionary
             Dim storename() As String
             'storename = LCase(calibration_name) & "_cal"
            If InStr(LCase(Instance_Data.CUS_Str_DigSrcData), "trimcodestorename") <> 0 And InStr(Instance_Data.CUS_Str_DigSrcData, ":") = 1 Then
             
             storename = Split(Instance_Data.CUS_Str_DigSrcData, ":")
             'If LCase(StoreName(0)) Like "trimcodestorename" And UBound(StoreName) = 1 Then
                Call AddStoredCaptureData(storename(1), storeDSP)
            End If
   
End Function

Public Function PLL_calibration_calc_MOD_AUS(CUS_Str_MainProgram As String, calc_data As DSPWave, MeasValue() As SiteDouble, storeDSP As DSPWave, Optional y As Integer = 2) As Long


Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim temp1_dict As New DSPWave
Dim temp2_dict As New DSPWave

Dim temp_testname_bin As Long
Dim temp_testname_dec As Long
Dim testname_str() As String
Dim delta_value() As New SiteDouble
Dim temp_delta_value As Integer
Dim temp_cal_code As New DSPWave
Dim RefPLL_calibration_code As New DSPWave
Dim TXPLL_calibration_code As New DSPWave
Dim DPTXPLL_calibration_code As New DSPWave
Dim calibration_target() As String
Dim calibration_target_value As Long

Dim OutputTname_format() As String
Dim TestNameInput As String
Dim bitnumber As Long
Dim calibration_name As String

calibration_target = Split(Split(CUS_Str_MainProgram, "(")(0), "_")
calibration_target_value = CLng(calibration_target(5))
calibration_name = calibration_target(4)
bitnumber = calibration_target(3)


ReDim testname_str(bitnumber)

        ''''calc and print in datalog

            Dim active_site As Variant
            Dim site As Variant
            For Each site In TheExec.sites.Active
                active_site = site
            Next site

            For i = 0 To (calc_data(active_site).SampleSize) - 2 'TY add 2018/08/29

            '''''dec to bin testname
                temp_testname_dec = i
                If i > 15 Then
                    temp_testname_dec = i + 8
                End If
                    For j = 0 To bitnumber - 1
                      temp_testname_bin = temp_testname_dec Mod 2
                      temp_testname_dec = Fix(temp_testname_dec / 2)
                      testname_str(j) = CStr(temp_testname_bin)
                    Next j
                    testname_str(bitnumber) = ""
                    For j = 1 To bitnumber
                        testname_str(bitnumber) = testname_str(bitnumber) & testname_str(bitnumber - j)
                    Next j
                
                    'TestNameInput = Report_TName_From_Instance("C", "", "F_" & testname_str(5), i, 1)
                
                    If gl_Disable_HIP_debug_log = False Then
                    
                        For Each site In TheExec.sites.Active
                            TheExec.Datalog.WriteComment "Site = " & site & "    " & "F_" & testname_str(bitnumber) & " = " & calc_data.Element(i) & ", MeasureVoltage = " & MeasValue(i + 3)
                        Next site
                        
                    End If
            Next i
             
              TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
            
            If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=calc_data.Element(24), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
            
             'store to dictionary
             Dim storename() As String
             'storename = LCase(calibration_name) & "_cal"
            If InStr(Instance_Data.CUS_Str_DigSrcData, "TrimCodeStoreName") <> 0 And InStr(Instance_Data.CUS_Str_DigSrcData, ":") = 1 Then
             If y <> 2 Then
                storename = Split(Split(Instance_Data.CUS_Str_DigSrcData, ";")(y), ":")
             Else
                storename = Split(Instance_Data.CUS_Str_DigSrcData, ":")
             End If
             'If LCase(StoreName(0)) Like "trimcodestorename" And UBound(StoreName) = 1 Then
                Call AddStoredCaptureData(storename(1), storeDSP)
            End If

   
End Function
Public Function PLL_calibration_calc_MOD_CIO(CUS_Str_MainProgram As String, calc_data As DSPWave, _
                                    target_var_low As SiteDouble, target_var_high As SiteDouble, storeDSP As DSPWave, storeDSP2 As DSPWave, _
                                    target_var_low2 As SiteDouble, target_var_high2 As SiteDouble, storeDSP3 As DSPWave, StoreDSP4 As DSPWave) As Long


Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim temp1_dict As New DSPWave
Dim temp2_dict As New DSPWave

Dim temp_testname_bin As Long
Dim temp_testname_dec As Long
Dim testname_str() As String
Dim delta_value() As New SiteDouble
Dim temp_delta_value As Integer
Dim temp_cal_code As New DSPWave
Dim RefPLL_calibration_code As New DSPWave
Dim TXPLL_calibration_code As New DSPWave
Dim DPTXPLL_calibration_code As New DSPWave
Dim calibration_target() As String
Dim calibration_target_value As Long

Dim OutputTname_format() As String
Dim TestNameInput As String
Dim bitnumber As Long
Dim calibration_name As String

calibration_target = Split(CUS_Str_MainProgram, "_")
calibration_name = calibration_target(4)
bitnumber = calibration_target(3)


ReDim testname_str(bitnumber)

        ''''calc and print in datalog
            Dim active_site As Variant
            
            For Each site In TheExec.sites.Active
                active_site = site
            Next site

            For i = 0 To (calc_data(active_site).SampleSize) - 2 'TY add 2018/08/29

            '''''dec to bin testname
                temp_testname_dec = i + 1
                    For j = 0 To bitnumber - 1
                      temp_testname_bin = temp_testname_dec Mod 2
                      temp_testname_dec = Fix(temp_testname_dec / 2)
                      testname_str(j) = CStr(temp_testname_bin)
                    Next j
                    testname_str(bitnumber) = ""
                    For j = 1 To bitnumber
                        testname_str(bitnumber) = testname_str(bitnumber) & testname_str(bitnumber - j)
                    Next j
                    
                    If gl_Disable_HIP_debug_log = False Then
                    
                        For Each site In TheExec.sites.Active
                            TheExec.Datalog.WriteComment "Site = " & site & "    " & "F_" & testname_str(bitnumber) & "========>" & calc_data.Element(i)
                        Next site
                        
                    End If
                    'TestNameInput = Report_TName_From_Instance("C", "X", "F" & testname_str(bitnumber), i, 1)
                
                'If gl_Disable_HIP_debug_log = False Then
                    'TheExec.Flow.TestLimit resultVal:=calc_data.Element(i), Tname:=TestNameInput, ForceResults:=tlForceNone, ScaleType:=scaleNoScaling, formatStr:="%.1f"
              '  Else
             '       TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
                'End If
                
                    If i = ((calc_data(active_site).SampleSize) - 2) / 2 - 1 Then TheExec.Datalog.WriteComment ""
                    If i = ((calc_data(active_site).SampleSize) - 2) / 2 - 1 Then i = i + 1
                
            Next i
             
            TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
            
            If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=target_var_low, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
            
            TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
            
            If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=target_var_high, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
            
            TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
            
            If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=target_var_low2, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
            
            TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
            
            If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=target_var_high2, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

             'store to dictionary
             Dim storename As String
             storename = "cio3pll0_efuse0_fcal_ate"
             'storename = LCase(calibration_name) & "callow"
             
             Call AddStoredCaptureData(storename, storeDSP)
             storename = "cio3pll0_efuse1_fcal_ate"
             'storename = LCase(calibration_name) & "calhigh"
             
             Call AddStoredCaptureData(storename, storeDSP2)
             
             
            storename = "cio3pll1_efuse0_fcal_ate"
             'storename = LCase(calibration_name) & "callow"
             
             Call AddStoredCaptureData(storename, storeDSP3)
             storename = "cio3pll1_efuse1_fcal_ate"
             'storename = LCase(calibration_name) & "calhigh"
             
             Call AddStoredCaptureData(storename, StoreDSP4)
                
   
End Function

Public Function CreateSimulateDataDSPWave(OutDspWave As DSPWave, DigCap_Sample_Size As Long, DigCap_DataWidth As Long)
    Dim site As Variant
    Dim i As Integer
    Dim TempStr_DSP As New SiteVariant
    
    If TheExec.TesterMode = testModeOffline Then
        If DigCap_DataWidth <> 0 Then
            For Each site In TheExec.sites
                For i = 0 To DigCap_Sample_Size - 1
                    If site = 0 Then
                        If i Mod 3 = 0 Then
                            OutDspWave(site).Element(i) = 1
                        Else
                            OutDspWave(site).Element(i) = 0
                        End If
                    Else
                        If i Mod 2 = 0 Then
                            OutDspWave(site).Element(i) = 0
                        Else
                            OutDspWave(site).Element(i) = 1
                        End If

                    End If
                Next i
            Next site
        Else
            For Each site In TheExec.sites
                If site = 0 Then
                    OutDspWave(site).Element(0) = 1
                    OutDspWave(site).Element(1) = 1
                    OutDspWave(site).Element(2) = 1
                    OutDspWave(site).Element(3) = 1
                    OutDspWave(site).Element(4) = 1
                    OutDspWave(site).Element(5) = 0

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
                    OutDspWave(site).Element(0) = 1
                    OutDspWave(site).Element(1) = 0
                    OutDspWave(site).Element(2) = 0
                    OutDspWave(site).Element(3) = 0
                    OutDspWave(site).Element(4) = 0
                    OutDspWave(site).Element(5) = 1
                    
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
            Next site
        End If
        
        For Each site In TheExec.sites
            For i = 0 To OutDspWave(site).SampleSize - 1
            TempStr_DSP(site) = TempStr_DSP(site) & CStr(OutDspWave(site).Element(i))
            Next i
            
        Next site
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Site_" & site & " simulate data = " & TempStr_DSP(site))
            Next site
        End If
        
    End If
End Function

Public Function Freq_WalkingStrobe_Meas_VOHVOL(MeasureF_Pin_SingleEnd As PinList, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, _
    Optional MeasF_WalkingStrobe_StepVoltage As Double, Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, _
    Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double, _
    Optional DictKey_VT As String)
    'Frequency WalkingStrobe created by JT 2016/03/01
    Dim site As Variant
    Dim MeasF_WalkingStrobe_Step As Long
    MeasF_WalkingStrobe_Step = (MeasF_WalkingStrobe_EndV - MeasF_WalkingStrobe_StartV) / MeasF_WalkingStrobe_StepVoltage + 1
    
    Dim MeasFreq_WKStrobe() As New PinListData
    ReDim MeasFreq_WKStrobe(MeasF_WalkingStrobe_Step) As New PinListData
    Dim WalkStrobe_i As Long
    Dim WalkStrobe_j As Long
    ''setup and measure Freq base on VOL and VOH setting.
    Dim WalkingStrobe_stepV As Double
    WalkingStrobe_stepV = (MeasF_WalkingStrobe_EndV - MeasF_WalkingStrobe_StartV) / MeasF_WalkingStrobe_Step
        For WalkStrobe_i = 0 To MeasF_WalkingStrobe_Step
            TheHdw.Digital.Pins(MeasureF_Pin_SingleEnd).Levels.Value(chVoh) = MeasF_WalkingStrobe_StartV + WalkStrobe_i * WalkingStrobe_stepV + MeasF_WalkingStrobe_BothVohVolDiffV
            TheHdw.Digital.Pins(MeasureF_Pin_SingleEnd).Levels.Value(chVol) = MeasF_WalkingStrobe_StartV + WalkStrobe_i * WalkingStrobe_stepV
            Call Freq_MeasFreqSetup(MeasureF_Pin_SingleEnd, MeasF_WalkingStrobe_interval, BOTH)
            Call HardIP_Freq_MeasFreqStart(MeasureF_Pin_SingleEnd, MeasF_WalkingStrobe_interval, MeasFreq_WKStrobe(WalkStrobe_i), 0)
        Next WalkStrobe_i
        
    ''analyze measurement data to decide which VOH/VOL level shiuld be used for measurement.
    Dim PinArr_WK() As String
    Dim PinCount_WK As Long
    TheExec.DataManager.DecomposePinList MeasureF_Pin_SingleEnd, PinArr_WK, PinCount_WK
    
    Dim Record_Temp_VOL As Double
    Dim Record_Min_VOL As Double
    Dim Record_Max_VOL As Double
    Dim Record_Mid_VOL As Double
    
    ''20170322 - Store mid value to dictionary
    Dim StoreMidValue As New SiteDouble
    For Each site In TheExec.sites
    
        For WalkStrobe_j = 0 To PinCount_WK - 1
            Record_Min_VOL = 9999
            Record_Max_VOL = -9999
            For WalkStrobe_i = 0 To MeasF_WalkingStrobe_Step
                If MeasFreq_WKStrobe(WalkStrobe_i).Pins(PinArr_WK(WalkStrobe_j)).Value(site) > MeasF_WalkingStrobe_miniFreq Then
                    Record_Temp_VOL = MeasF_WalkingStrobe_StartV + WalkStrobe_i * WalkingStrobe_stepV
                    If Record_Temp_VOL > Record_Max_VOL Then Record_Max_VOL = Record_Temp_VOL
                    If Record_Temp_VOL < Record_Min_VOL Then Record_Min_VOL = Record_Temp_VOL
                End If
            Next WalkStrobe_i
            
            If Record_Min_VOL <> 9999 Then
                Record_Mid_VOL = (Record_Max_VOL + Record_Min_VOL) / 2
                TheHdw.Digital.Pins(PinArr_WK(WalkStrobe_j)).Levels.Value(chVoh) = Record_Mid_VOL + MeasF_WalkingStrobe_BothVohVolDiffV
                TheHdw.Digital.Pins(PinArr_WK(WalkStrobe_j)).Levels.Value(chVol) = Record_Mid_VOL
                TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & PinArr_WK(WalkStrobe_j) & " , " & "V_Max= " & Format(Record_Max_VOL, "0.000")
                TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & PinArr_WK(WalkStrobe_j) & " , " & "V_Min= " & Format(Record_Min_VOL, "0.000")
                TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & PinArr_WK(WalkStrobe_j) & " , " & "V_Mid= " & Format(Record_Mid_VOL, "0.000")
                StoreMidValue(site) = Record_Mid_VOL
''                TheExec.Datalog.WriteComment "Site= " & Site & " , " & " Pin= " & PinArr_WK(WalkStrobe_j) & " , " & "VOH= " & Format((Record_Mid_VOL + MeasF_WalkingStrobe_BothVohVolDiffV), "0.000") & " V"
''                TheExec.Datalog.WriteComment "Site= " & Site & " , " & " Pin= " & PinArr_WK(WalkStrobe_j) & " , " & "VOL= " & Format(Record_Mid_VOL, "0.000") & " V"
            Else
                TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & PinArr_WK(WalkStrobe_j) & " , " & "VOH= default, search fail"
                TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & PinArr_WK(WalkStrobe_j) & " , " & "VOL= default, search fail"
    
            End If
        Next WalkStrobe_j
        
    Next site
    If DictKey_VT <> "" Then
        Call AddStoredMeasurement(DictKey_VT, StoreMidValue)
    End If
End Function

Public Function SimulateOutputFreq(MeasureF_Pin As PinList, ByRef MeasureFreq As PinListData) As Long
    Dim site As Variant
    For Each site In TheExec.sites.Active
        If site = 0 Then
            MeasureFreq.Pins(MeasureF_Pin).Value(site) = 1001000
        ElseIf site = 1 Then
            MeasureFreq.Pins(MeasureF_Pin).Value(site) = 991000
        ElseIf site = 2 Then
''            MeasureFreq.Pins(MeasureF_Pin).Value(Site) = 992000
        ElseIf site = 3 Then
''            MeasureFreq.Pins(MeasureF_Pin).Value(Site) = 1002000
        End If
    Next site
End Function

Public Function IPF_CZ_PrintDigCapInfo(argc As Integer, argv() As String) As Long

    '' 20151114 - Print DigCap info during CZ
    Dim site As Variant
    Dim i As Long
    Dim X_SetupName As String
    Dim Y_SetupName As String
    Dim Volt_pointval As Double
    Dim FRC_pointval As Double

''    If UCase(argv(2)) = UCase("PrintFreq") Then
        X_SetupName = TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Item(tlDevCharShmooAxis_X).StepName
''        Y_SetupName = TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Item(tlDevCharShmooAxis_Y).StepName
        For Each site In TheExec.sites
            Volt_pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
''            FRC_pointval = TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
            For i = 0 To G_pld_DigCapInfo(0).Pins.Count - 1
                TheExec.Datalog.WriteComment ("Site = " & site & ",  " & X_SetupName & "=" & Volt_pointval & " V,   DSSC Code " & G_pld_DigCapInfo(0).Pins(i) & " = " & G_pld_DigCapInfo(0).Pins(i).Value(site))
            Next i
        Next site
''    End If
End Function
Public Function CheckRangesAndClamps(pld_MeasureResult As PinListData, s_CheckType As String, HighUseLimitVal As Double, LowUseLimitVal As Double) As Long

    Dim DUTPin As String
    Dim i As Long
    
    Dim Power_Volt As Double
    Dim Power_Current As Double
    Dim Power_SourceCurrentClamp As Double
    Dim Power_SourceCurrentRange As Double
    Dim Power_MeterCurrentRange As Double
    
    Dim PPMU_Volt As Double
    Dim PPMU_V_ClampHi As Double
    Dim PPMU_V_ClampLo As Double
    Dim PPMU_ForceCurrentRange As Double
    Dim PPMU_MeasureCurrentRange As Double
    Dim PPMU_Mode As Long
    
    Dim site As Variant
    Dim CheckDone As Boolean
        
    Dim NumTypes As Long
    Dim ThisPinType As String
    Dim PowerType() As String
    
    Dim MeasValue As Double
    Dim MeasPin As String
    Dim Meas_VS_MeterRange_DiffPercent As Double
    Dim DiffPercentage As Double
    DiffPercentage = 0.1
    
    Dim InstName As String
    Dim SuggestMeterCurrentRange As Double
    ''==========================================================================
    CheckDone = False
    
    Dim PinCountMax As Long
    PinCountMax = pld_MeasureResult.Pins.Count
    
    For Each site In TheExec.sites.Selected
    
        If Not (CheckDone) Then
        
            For i = 0 To PinCountMax - 1
                
                DUTPin = pld_MeasureResult.Pins.Item(i).Name
            
                ThisPinType = TheExec.DataManager.PinType(DUTPin)
                
                MeasValue = pld_MeasureResult.Pins.Item(i).Value
                
                InstName = GetInstrument(DUTPin, site)
                
                If ThisPinType = "Power" Then
                    Call TheExec.DataManager.GetChannelTypes(DUTPin, NumTypes, PowerType())
                    
                    If PowerType(0) = "DCVS" Or PowerType(0) = "DCVSMerged2" Or PowerType(0) = "DCVSMerged4" Or PowerType(0) = "DCVSMerged6" Or PowerType(0) = "DCVSMerged8" Then
                        Power_Volt = TheHdw.DCVS.Pins(DUTPin).Voltage.Value
                        Power_SourceCurrentClamp = TheHdw.DCVS.Pins(DUTPin).CurrentLimit.Source.FoldLimit.Level.Value
                        Power_SourceCurrentRange = TheHdw.DCVS.Pins(DUTPin).CurrentRange.Value
                        Power_MeterCurrentRange = TheHdw.DCVS.Pins(DUTPin).Meter.CurrentRange.Value
                                                
                    ElseIf PowerType(0) = "DCVI" Or PowerType(0) = "DCVIMerged" Then
                        Power_Volt = TheHdw.DCVI.Pins(DUTPin).Voltage
                        Power_Current = TheHdw.DCVI.Pins(DUTPin).current
                        Power_SourceCurrentRange = TheHdw.DCVI.Pins(DUTPin).CurrentRange.Value
                        Power_MeterCurrentRange = TheHdw.DCVI.Pins(DUTPin).Meter.CurrentRange.Value
                    End If
                    
                    If UCase(s_CheckType) = "I" Then
                        If Power_SourceCurrentClamp < HighUseLimitVal Then
                            TheExec.Datalog.WriteComment ("Range Check: " & UCase(DUTPin) & " type is " & PowerType(0) & ", Current clamp is " & Power_SourceCurrentClamp & " less than High Use-Limit value " & HighUseLimitVal)
                        Else
                        End If
                        
                        If Power_SourceCurrentRange < HighUseLimitVal Then
                            TheExec.Datalog.WriteComment ("Range Check: " & UCase(DUTPin) & " type is " & PowerType(0) & ", Source current range is " & Power_SourceCurrentRange & " less than High Use-Limit value " & HighUseLimitVal)
                        Else
                        End If
                        
                        If Power_MeterCurrentRange < HighUseLimitVal Then
                            TheExec.Datalog.WriteComment ("Range Check: " & UCase(DUTPin) & " type is " & PowerType(0) & ", Meter current range is " & Power_MeterCurrentRange & " less than High Use-Limit value " & HighUseLimitVal)
                        Else
                            Call GetAppropriateMeterCurrentRange(DUTPin, InstName, HighUseLimitVal, SuggestMeterCurrentRange)
                            
                            If Power_MeterCurrentRange > SuggestMeterCurrentRange Then
                                TheExec.Datalog.WriteComment ("Range Check: " & UCase(DUTPin) & " type is " & PowerType(0) & ", Meter current range current setup is " & Power_MeterCurrentRange & " Suggested meter current range is " & SuggestMeterCurrentRange)
                            End If
                        End If
                
                    ElseIf UCase(s_CheckType) = "V" Then
                        '' Only for UVI80, it can compare source current and source current range, have to pass Power_Current
                        InstName = GetInstrument(DUTPin, site)
                        
                        If InstName = "DC-07" Then
'''                            Call GetAppropriateMeterCurrentRange(DUTPin, InstName, HighUseLimitVal, SuggestMeterCurrentRange)
                        End If
'''                    ElseIf UCase(s_CheckType) = "R" Then
'''                        Meas_VS_MeterRange_DiffPercent = Abs((MeasValue - Power_MeterCurrentRange) / Power_MeterCurrentRange)
'''                        If Meas_VS_MeterRange_DiffPercent < DiffPercentage Then
'''                            TheExec.Datalog.WriteComment ("R Test, " & DUTPin & " type is " & PowerType(0) & ", Meter current range is " & Power_MeterCurrentRange & " less than 10% by compare with measure value")
'''                        End If
                    End If
                    
                Else
                    PPMU_Volt = TheHdw.PPMU.Pins(DUTPin).Voltage.Value
                    PPMU_V_ClampHi = TheHdw.PPMU.Pins(DUTPin).ClampVHi.Value
                    PPMU_V_ClampLo = TheHdw.PPMU.Pins(DUTPin).ClampVLo.Value
                   
                    PPMU_ForceCurrentRange = TheHdw.PPMU.Pins(DUTPin).ForceCurrentRange
                    PPMU_MeasureCurrentRange = TheHdw.PPMU.Pins(DUTPin).MeasureCurrentRange
                    PPMU_Mode = TheHdw.PPMU.Pins(DUTPin).mode       ' forceV=2, forceI=1
                    
                    If UCase(s_CheckType) = "I" Then
                    
                       If PPMU_MeasureCurrentRange < HighUseLimitVal Then
                            TheExec.Datalog.WriteComment ("Range Check: " & UCase(DUTPin) & " type is PPMU " & ", Measure current range is " & PPMU_MeasureCurrentRange & " less than High Use-Limit value " & HighUseLimitVal)
                       Else
                            Call GetAppropriateMeterCurrentRange(DUTPin, InstName, HighUseLimitVal, SuggestMeterCurrentRange)
                            
                            If PPMU_MeasureCurrentRange > SuggestMeterCurrentRange Then
                                TheExec.Datalog.WriteComment ("Range Check: " & UCase(DUTPin) & " type is PPMU" & ", Measure current range current setup is " & PPMU_MeasureCurrentRange & " Suggested meter current range is " & SuggestMeterCurrentRange)
                            End If
                       End If
                
                    ElseIf UCase(s_CheckType) = "V" Then
                         '' Maybe can add foce current to compare force current range
                         
                        If PPMU_V_ClampHi < HighUseLimitVal Then
                            TheExec.Datalog.WriteComment ("Range Check: " & UCase(DUTPin) & " type is PPMU " & ", Voltage clamp high is " & PPMU_V_ClampHi & " less than High Use-Limit value " & HighUseLimitVal)
                        End If
                        
                        If PPMU_V_ClampLo > LowUseLimitVal Then
                            TheExec.Datalog.WriteComment ("Range Check: " & UCase(DUTPin) & " type is PPMU " & ", Voltage clamp low is " & PPMU_V_ClampLo & " over than Low Use-Limit value " & LowUseLimitVal)
                        End If
                
'''                    ElseIf UCase(s_CheckType) = "R" Then
'''                        Meas_VS_MeterRange_DiffPercent = Abs((MeasValue - PPMU_MeasureCurrentRange) / PPMU_MeasureCurrentRange)
'''
'''                        If Meas_VS_MeterRange_DiffPercent < DiffPercentage Then
'''                            TheExec.Datalog.WriteComment ("R Test, " & DUTPin & " type is PPMU " & ", Meter current range is " & PPMU_MeasureCurrentRange & " less than 10% by compare with measure value")
'''                        End If
                    End If
                End If
            Next i
            
            CheckDone = True        ' only has to be done foe 1 Site
            
        End If
    
    Next site
    
End Function



'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Public Function GetAppropriateMeterCurrentRange(DUTPin As String, InstName As String, HighUseLimitVal As Double, ByRef SuggestMeterCurrentRange) As Double
    Dim NumTypes As Long
    Dim PowerType() As String
    Dim factor As Long
           
    Select Case InstName
        
        Case "DC-07"
            Call TheExec.DataManager.GetChannelTypes(DUTPin, NumTypes, PowerType())
            
            Select Case PowerType(0)
                Case "DCVI"
                    factor = 1
                    
                Case "DCVIMerged"
                    factor = 2
                    
                Case Else
            End Select
            
            If HighUseLimitVal > 2 * factor Then
                TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 2A of DCVI")
            
            ElseIf HighUseLimitVal > 1 * factor Then
                 SuggestMeterCurrentRange = 2 * factor
            
            ElseIf HighUseLimitVal > 0.2 * factor Then
                  SuggestMeterCurrentRange = 1 * factor
            
            ElseIf HighUseLimitVal > 0.02 * factor Then
                 SuggestMeterCurrentRange = 0.2 * factor
            
            ElseIf HighUseLimitVal > 0.002 * factor Then
                 SuggestMeterCurrentRange = 0.02 * factor
            
            ElseIf HighUseLimitVal > 0.0002 * factor Then
                 SuggestMeterCurrentRange = 0.002 * factor
            
            ElseIf HighUseLimitVal > 0.00002 * factor Then
                 SuggestMeterCurrentRange = 0.0002 * factor
                 
            Else
                 SuggestMeterCurrentRange = 0.00002 * factor
            
            End If
        
        Case "VHDVS"
            Call TheExec.DataManager.GetChannelTypes(DUTPin, NumTypes, PowerType())
                
            Select Case PowerType(0)
                Case "DCVS"
                
                    If HighUseLimitVal > 0.8 Then
                        TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 0.8A of DCVS-UVS256")
                    
                    ElseIf HighUseLimitVal > 0.5 Then
                         SuggestMeterCurrentRange = 0.8
                    
                    ElseIf HighUseLimitVal > 0.2 Then
                          SuggestMeterCurrentRange = 0.5
                    
                    ElseIf HighUseLimitVal > 0.02 Then
                         SuggestMeterCurrentRange = 0.2
                    
                    ElseIf HighUseLimitVal > 0.002 Then
                         SuggestMeterCurrentRange = 0.02
                    
                    ElseIf HighUseLimitVal > 0.0002 Then
                         SuggestMeterCurrentRange = 0.002
                    
                    ElseIf HighUseLimitVal > 0.00002 Then
                         SuggestMeterCurrentRange = 0.0002
                         
                    Else
                         SuggestMeterCurrentRange = 0.000004
                    
                    End If
                                        
                Case "DCVSMerged2"
                
                    If HighUseLimitVal > 1 Then
                        TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 0.8A of DCVSMerged2-UVS256")
                    
                    ElseIf HighUseLimitVal > 0.5 Then
                         SuggestMeterCurrentRange = 1
                    
                    ElseIf HighUseLimitVal > 0.4 Then
                         SuggestMeterCurrentRange = 0.5
                         
                    ElseIf HighUseLimitVal > 0.2 Then
                          SuggestMeterCurrentRange = 0.4
                    
                    ElseIf HighUseLimitVal > 0.04 Then
                         SuggestMeterCurrentRange = 0.2
                    
                    ElseIf HighUseLimitVal > 0.02 Then
                         SuggestMeterCurrentRange = 0.04
                         
                    ElseIf HighUseLimitVal > 0.002 Then
                         SuggestMeterCurrentRange = 0.02
                    
                    ElseIf HighUseLimitVal > 0.0002 Then
                         SuggestMeterCurrentRange = 0.002
                    
                    ElseIf HighUseLimitVal > 0.00002 Then
                         SuggestMeterCurrentRange = 0.0002
                         
                    Else
                         SuggestMeterCurrentRange = 0.000004
                    
                    End If
                
                Case "DCVSMerged4"

                    If HighUseLimitVal > 2 Then
                        TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 0.8A of DCVSMerged4-UVS256")
                    
                    ElseIf HighUseLimitVal > 0.5 Then
                         SuggestMeterCurrentRange = 2
                    
                    ElseIf HighUseLimitVal > 0.2 Then
                          SuggestMeterCurrentRange = 0.5
                    
                    ElseIf HighUseLimitVal > 0.02 Then
                         SuggestMeterCurrentRange = 0.2
                    
                    ElseIf HighUseLimitVal > 0.002 Then
                         SuggestMeterCurrentRange = 0.02
                    
                    ElseIf HighUseLimitVal > 0.0002 Then
                         SuggestMeterCurrentRange = 0.002
                    
                    ElseIf HighUseLimitVal > 0.00002 Then
                         SuggestMeterCurrentRange = 0.0002
                         
                    Else
                         SuggestMeterCurrentRange = 0.000004
                    
                    End If

                Case "DCVSMerged8"
                
                    If HighUseLimitVal > 4 Then
                        TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 0.8A of DCVSMerged8-UVS256")
                    
                    ElseIf HighUseLimitVal > 0.5 Then
                         SuggestMeterCurrentRange = 4
                    
                    ElseIf HighUseLimitVal > 0.2 Then
                          SuggestMeterCurrentRange = 0.5
                    
                    ElseIf HighUseLimitVal > 0.02 Then
                         SuggestMeterCurrentRange = 0.2
                    
                    ElseIf HighUseLimitVal > 0.002 Then
                         SuggestMeterCurrentRange = 0.02
                    
                    ElseIf HighUseLimitVal > 0.0002 Then
                         SuggestMeterCurrentRange = 0.002
                    
                    ElseIf HighUseLimitVal > 0.00002 Then
                         SuggestMeterCurrentRange = 0.0002
                         
                    Else
                         SuggestMeterCurrentRange = 0.000004
                    
                    End If
                    
                Case Else
                
            End Select
                
        Case "HexVS"
            Call TheExec.DataManager.GetChannelTypes(DUTPin, NumTypes, PowerType())
            
            Select Case PowerType(0)
                Case "DCVS"
                
                    If HighUseLimitVal > 15 Then
                        TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 15A of DCVS-HexVS")
                    
                    ElseIf HighUseLimitVal > 1 Then
                         SuggestMeterCurrentRange = 15
                    
                    ElseIf HighUseLimitVal > 0.1 Then
                          SuggestMeterCurrentRange = 1
                    
                    ElseIf HighUseLimitVal > 0.01 Then
                         SuggestMeterCurrentRange = 0.1
                         
                    Else
                         SuggestMeterCurrentRange = 0.01
                    
                    End If
                
                Case "DCVSMerged2"
                
                    If HighUseLimitVal > 30 Then
                        TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 30A of DCVSMerged2-HexVS")
                    
                    ElseIf HighUseLimitVal > 15 Then
                         SuggestMeterCurrentRange = 30
                         
                    ElseIf HighUseLimitVal > 1 Then
                         SuggestMeterCurrentRange = 15
                    
                    ElseIf HighUseLimitVal > 0.1 Then
                          SuggestMeterCurrentRange = 1
                    
                    ElseIf HighUseLimitVal > 0.01 Then
                         SuggestMeterCurrentRange = 0.1
                         
                    Else
                         SuggestMeterCurrentRange = 0.01
                    
                    End If
                
                Case "DCVSMerged4"
                
                    If HighUseLimitVal > 60 Then
                        TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 60A of DCVSMerged4-HexVS")
                    
                    ElseIf HighUseLimitVal > 30 Then
                         SuggestMeterCurrentRange = 60
                         
                    ElseIf HighUseLimitVal > 15 Then
                         SuggestMeterCurrentRange = 30
                         
                    ElseIf HighUseLimitVal > 1 Then
                         SuggestMeterCurrentRange = 15
                    
                    ElseIf HighUseLimitVal > 0.1 Then
                          SuggestMeterCurrentRange = 1
                    
                    ElseIf HighUseLimitVal > 0.01 Then
                         SuggestMeterCurrentRange = 0.1
                         
                    Else
                         SuggestMeterCurrentRange = 0.01
                    
                    End If
                    
                Case "DCVSMerged6"
            
                    If HighUseLimitVal > 60 Then
                        TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 60A of DCVSMerged4-HexVS")
                    
                    ElseIf HighUseLimitVal > 30 Then
                         SuggestMeterCurrentRange = 60
                         
                    ElseIf HighUseLimitVal > 15 Then
                         SuggestMeterCurrentRange = 30
                         
                    ElseIf HighUseLimitVal > 1 Then
                         SuggestMeterCurrentRange = 15
                    
                    ElseIf HighUseLimitVal > 0.1 Then
                          SuggestMeterCurrentRange = 1
                    
                    ElseIf HighUseLimitVal > 0.01 Then
                         SuggestMeterCurrentRange = 0.1
                         
                    Else
                         SuggestMeterCurrentRange = 0.01
                    
                    End If
                    
                Case Else
                    
            End Select
        
        Case "HSD - U"
        
            If HighUseLimitVal > 0.05 Then
                TheExec.Datalog.WriteComment ("Wrong setting - High use-limit over spec 0.05A of HSD-U")
            
            ElseIf HighUseLimitVal > 0.002 Then
                 SuggestMeterCurrentRange = 0.05
                 
            ElseIf HighUseLimitVal > 0.0002 Then
                 SuggestMeterCurrentRange = 0.002
                 
            ElseIf HighUseLimitVal > 0.00002 Then
                 SuggestMeterCurrentRange = 0.0002
            
            ElseIf HighUseLimitVal > 0.000002 Then
                  SuggestMeterCurrentRange = 0.00002
              
            Else
                 SuggestMeterCurrentRange = 0.000002
            
            End If
        
        Case Else
        
    End Select

End Function

Public Function VFI_AnalyzedInputStringByAt(ByRef MeasV_Pins As String, ByRef MeasF_PinS_SingleEnd As String, ByRef MeasI_pinS As String, ByRef MeasI_Range As String, ByRef MeasF_PinS_Differential As String, _
    ByRef ForceV_Val As String, ByRef ForceI_Val As String) As Long
    '' 20160201 - Check input argumenets whether have "@" in the first character. Add it If no "@" in the beginning. Then remove it to process fomat.
    '' The purpose is to cover import issue. Ex:++
    
    'Call CheckInputStringByAt(MeasV_PinS)
    'Call CheckInputStringByAt(MeasF_PinS_SingleEnd)
    'Call CheckInputStringByAt(MeasI_pinS)
    'Call CheckInputStringByAt(MeasI_Range)
    'Call CheckInputStringByAt(MeasF_PinS_Differential)
    
    'Call CheckInputStringByAt(ForceV_Val)
    'Call CheckInputStringByAt(ForceI_Val)
    
    
    MeasV_Pins = Replace(MeasV_Pins, "@", "")
    MeasF_PinS_SingleEnd = Replace(MeasF_PinS_SingleEnd, "@", "")
    MeasI_pinS = Replace(MeasI_pinS, "@", "")
    MeasI_Range = Replace(MeasI_Range, "@", "")
    MeasF_PinS_Differential = Replace(MeasF_PinS_Differential, "@", "")
    
    ForceV_Val = Replace(ForceV_Val, "@", "")
    ForceI_Val = Replace(ForceI_Val, "@", "")
    
End Function

Public Function AnalyzeDigSrcEquationAssignmentContent(DigSrc_Equation As String, DigSrc_Assignment As String) As Long


    ''20160824 - Check0: Show error log if DigSrc_Equation has content but DigSrc_Assignment doesn't has content, vice versa
    ''20160804 - Check1: Check DigSrc_Assignment whether have duplicate segment
    ''20160805 - Check2: Check DigSrc_Equation segment name whether reference DigSrc_Assignment segment name
    Dim splitbysemicolon() As String
    Dim SplitByEqual() As String
    Dim Assignment_SegmentName As String
    Assignment_SegmentName = ""
        
        
    ''20160824 - Check0: Show error log if DigSrc_Equation has content but DigSrc_Assignment doesn't has content, vice versa
    If DigSrc_Equation <> "" And DigSrc_Assignment = "" Then
        Call TheExec.ErrorLogMessage("DigSrc_Equation has content but DigSrc_Assignment doesn't has")
        End
    ElseIf DigSrc_Equation = "" And DigSrc_Assignment <> "" Then
        Call TheExec.ErrorLogMessage("DigSrc_Assignment has content but DigSrc_Equation doesn't has")
        End
    End If
        
    Dim i As Long
    splitbysemicolon = Split(DigSrc_Assignment, ";")
    
    For i = 0 To UBound(splitbysemicolon)
        SplitByEqual = Split(splitbysemicolon(i), "=")
        If i = 0 Then
            Assignment_SegmentName = SplitByEqual(0)
        Else
            Assignment_SegmentName = Assignment_SegmentName & "," & SplitByEqual(0)
        End If
    Next i
    
    Dim SplitByComma() As String
    Dim TheSameNameIndex As String
''    TheSameNameIndex = 0
    SplitByComma = Split(Assignment_SegmentName, ",")
    Dim Segment As Variant
    ''20160804 - Check1: Check DigSrc_Assignment whether have duplicate segment
    For Each Segment In SplitByComma
        TheSameNameIndex = 0
        For i = 0 To UBound(SplitByComma)
            If Segment = SplitByComma(i) Then
                TheSameNameIndex = TheSameNameIndex + 1
            End If
        Next i
        If TheSameNameIndex > 1 Then
            Call TheExec.ErrorLogMessage("DigSrc_Assignment duplicate segment name, please modify segment name as unique")
            End
        End If
    Next Segment
    ''20160805 - Check2: Check DigSrc_Equation segment name whether reference DigSrc_Assignment segment name
    Dim AryEquationSegName() As String
    Dim AryAssignmentSegName() As String
    Dim j As Long
    Dim GetReferenceIndex As Long
    
    AryEquationSegName = Split(DigSrc_Equation, "+")
    AryAssignmentSegName = Split(Assignment_SegmentName, ",")
    
    For i = 0 To UBound(AryEquationSegName)
        GetReferenceIndex = 0
        For j = 0 To UBound(AryAssignmentSegName)
            If LCase(AryEquationSegName(i)) = LCase(AryAssignmentSegName(j)) Or LCase(AryAssignmentSegName(j)) = "repeat" Then
                GetReferenceIndex = GetReferenceIndex + 1
            End If
        Next j

        If GetReferenceIndex = 0 Then
            Call TheExec.ErrorLogMessage("DigSrc_Equation segment name """ & AryEquationSegName(i) & """ doesn't have reference with DigSrc_Assignment segment name")
            End
        End If

    Next i
End Function

Public Function Checker_WithDictionary(InPutString As String) As Boolean
    Dim SplitBy1() As String
    Dim SplitBy0() As String
    Dim NumOf_1 As Long
    Dim NumOf_0 As Long
    SplitBy1 = Split(InPutString, "1")
    SplitBy0 = Split(InPutString, "0")
    
    NumOf_1 = UBound(SplitBy1)
    NumOf_0 = UBound(SplitBy0)
    
    If Len(InPutString) = NumOf_1 + NumOf_0 Then
        Checker_WithDictionary = False
    Else
        Checker_WithDictionary = True
    End If
End Function

Public Function Checker_ConstantBinary(InPutString As String) As Boolean
    Dim SplitBy1() As String
    Dim SplitBy0() As String
    Dim NumOf_1 As Long
    Dim NumOf_0 As Long
    SplitBy1 = Split(InPutString, "1")
    SplitBy0 = Split(InPutString, "0")
    
    NumOf_1 = UBound(SplitBy1)
    NumOf_0 = UBound(SplitBy0)
    
    If Len(InPutString) = NumOf_1 + NumOf_0 Then
        Checker_ConstantBinary = True
    Else
        Checker_ConstantBinary = False
    End If
End Function

Public Function Checker_StoreDigCapAllToDictionary(ByRef CUS_Str_DigCapData As String, OutDspWave As DSPWave, Optional NumberPins As Long = 1) As Long

    Dim splitbyand() As String
    Dim b_StoreToDictionary As Boolean
    splitbyand = Split(CUS_Str_DigCapData, "&")
    
    Dim OutDspWave_Binary As New DSPWave
    Dim tmp_dsp As New DSPWave
    
    If NumberPins > 1 Then
        
''        thehdw.DSP.ExecutionMode = tlDSPModeAutomatic
        Dim site As Variant
        For Each site In TheExec.sites
            tmp_dsp = OutDspWave.Copy
        Next site
        Call rundsp.DSPWaveDecToBinary(tmp_dsp, NumberPins, OutDspWave_Binary)
        
        If UBound(splitbyand) > 0 Then
            b_StoreToDictionary = True
            CUS_Str_DigCapData = splitbyand(1)
            Call AddStoredCaptureData(splitbyand(0), OutDspWave_Binary)
            
        ElseIf UBound(splitbyand) = 0 And InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") = 0 Then
            Call AddStoredCaptureData(splitbyand(0), OutDspWave_Binary)
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("DigCap data store in dictionary " & "<<" & splitbyand(0) & ">>")
        End If
    
    Else
        If UBound(splitbyand) > 0 Then
            b_StoreToDictionary = True
            CUS_Str_DigCapData = splitbyand(1)
            Call AddStoredCaptureData(splitbyand(0), OutDspWave)
            
        ElseIf UBound(splitbyand) = 0 And InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") = 0 Then
            Call AddStoredCaptureData(splitbyand(0), OutDspWave)
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("DigCap data store in dictionary " & "<<" & splitbyand(0) & ">>")
        End If
    End If

End Function

Public Function Checker_StoreWholeDigSrcInDict(ByRef DigSrc_Assignment As String, ByRef DigSrcWholeDictName As String) As Boolean
    Dim SplitByLeftPara() As String
    SplitByLeftPara = Split(DigSrc_Assignment, "(")
    If UBound(SplitByLeftPara) > 0 Then
        DigSrcWholeDictName = SplitByLeftPara(0)
        Checker_StoreWholeDigSrcInDict = True
        DigSrc_Assignment = Left(SplitByLeftPara(1), Len(SplitByLeftPara(1)) - 1)
    End If
End Function

Public Function Checker_DigSrcFromDict(DigSrc_Assignment As String) As Boolean
    If InStr(DigSrc_Assignment, "=") <> 0 Then
        Checker_DigSrcFromDict = False
    Else
        Checker_DigSrcFromDict = True
    End If
End Function

Public Function HardIP_InitialSetupForPatgen() As Long

    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If

    TheHdw.Digital.Patgen.Halt
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
End Function

Public Function AnalyzePatName(Pat As String, ByRef Str_FinalPatName As String) As Long
    Dim Str_Before_UnderLine As String, Str_After_UnderLine As String
    Dim pat_name() As String
    Dim pat_name_module() As String
    Dim Pat_name1() As String
     
    pat_name_module = Split(Pat, ":")
    pat_name = Split(pat_name_module(0), "\")
    
    pat_name(0) = pat_name(UBound(pat_name))
    pat_name(0) = Replace(pat_name(0), ".", "_")
    Pat_name1 = Split(TheExec.DataManager.instanceName, "_")

    Str_Before_UnderLine = pat_name(0)
    Str_After_UnderLine = Pat_name1(UBound(Pat_name1))
    
    Str_FinalPatName = Str_Before_UnderLine & "_" & Str_After_UnderLine
    
End Function

Public Function ProcessCalcEquation(Calc_Eqn As String) As Long
    
    Dim SplitBySemi() As String, SplitByColon() As String, SplitByLeftPara() As String
    Dim i As Long, j As Long, p As Long
    Dim KeyWord_Calc As String
    Dim ALG_InterPoseFuncName As String
    Dim ALG_InterPoseArgcName As String
    Dim testName As String
    Dim Operator As String
    Dim SplitByKeyWord() As String
    Dim ReturnDSPWave As New DSPWave
    Dim TestNameInput As String

    '' 20160914 - Equation pins for V,F,I
''    Dim EquationPins As String
    Dim CalcEquationPLD As CALC_EQUATION_PLD
    Dim ReturnPLD As New PinListData
    Dim OutputTname_format() As String
    
    'ReturnDSPWave.CreateConstant 0, 1, DspDouble
    
    SplitBySemi = Split(Calc_Eqn, ";")
    
    Dim KeyWord_Description As String
    
    For i = 0 To UBound(SplitBySemi)
                
                Set ReturnPLD = Nothing
                Set ReturnDSPWave = Nothing
                ReturnDSPWave.CreateConstant 0, 1, DspDouble
        
        SplitByColon = Split(SplitBySemi(i), ":")
        KeyWord_Calc = Left(UCase(SplitByColon(0)), 1)
        
        KeyWord_Description = SplitByColon(0)
        
        If UCase(KeyWord_Description) = UCase("ALG") Then
            KeyWord_Calc = UCase(KeyWord_Description)
        End If
        
        testName = SplitByColon(1)
        
        ''20160607 - Store return value in the Dictionary
        Dim DictKeyName As String
        If UBound(SplitByColon) = 3 Then
            DictKeyName = SplitByColon(3)
        End If
        
        Select Case KeyWord_Calc
            Case "V"
                Call ProcessStringForCalType(SplitByColon(2), CalcEquationPLD)
                Call StandardCalcuation(CalcEquationPLD, ReturnPLD)
                
                If UBound(SplitByColon) = 3 Then
                    Call AddStoredMeasurement(DictKeyName, ReturnPLD)
                End If
                
                If Not ByPassTestLimit Then
                    If InStr(UCase(KeyWord_Description), "SKIPUNIT") <> 0 Then
                        For p = 0 To ReturnPLD.Pins.Count - 1
                            TestNameInput = Report_TName_From_Instance("CalcV", ReturnPLD.Pins(p), , 0)
                            TheExec.Flow.TestLimit resultVal:=ReturnPLD.Pins(p), Tname:=TestNameInput, ForceResults:=tlForceFlow
                        Next p
                    Else
                        For p = 0 To ReturnPLD.Pins.Count - 1
                            TestNameInput = Report_TName_From_Instance("CalcV", ReturnPLD.Pins(p), , 0)
                            TheExec.Flow.TestLimit resultVal:=ReturnPLD.Pins(p), Unit:=unitVolt, Tname:=TestNameInput, ForceResults:=tlForceFlow
                        Next p
                    End If
                End If
                
            Case "F"
                Call ProcessStringForCalType(SplitByColon(2), CalcEquationPLD)
                Call StandardCalcuation(CalcEquationPLD, ReturnPLD)
                
                If UBound(SplitByColon) = 3 Then
                    Call AddStoredMeasurement(DictKeyName, ReturnPLD)
                End If
                
                If Not ByPassTestLimit Then
                    For p = 0 To ReturnPLD.Pins.Count - 1
                        TestNameInput = Report_TName_From_Instance("CalcF", ReturnPLD.Pins(p), SplitByColon(1), 0)
                        TheExec.Flow.TestLimit resultVal:=ReturnPLD.Pins(p), Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
                    Next p
                End If
                
            Case "I"
                Call ProcessStringForCalType(SplitByColon(2), CalcEquationPLD)
                Call StandardCalcuation(CalcEquationPLD, ReturnPLD)
                
                If UBound(SplitByColon) = 3 Then
                    Call AddStoredMeasurement(DictKeyName, ReturnPLD)
                End If
                
                If Not ByPassTestLimit Then
                    For p = 0 To ReturnPLD.Pins.Count - 1
                        TestNameInput = Report_TName_From_Instance("CalcI", ReturnPLD.Pins(p), SplitByColon(1), 0)
                        TheExec.Flow.TestLimit resultVal:=ReturnPLD.Pins(p), Unit:=unitAmp, Tname:=TestNameInput, ForceResults:=tlForceFlow
                    Next p
                End If
                
            Case "C"
                Operator = GetSplitKeyWord(SplitByColon(2))
                SplitByKeyWord = Split(SplitByColon(2), Operator)
                Call ProcessDSPCalculation(SplitByKeyWord, Operator, ReturnDSPWave)
                
                If UBound(SplitByColon) = 3 Then
                    Call AddStoredCaptureData(DictKeyName, ReturnDSPWave)
                End If
                
                If Not ByPassTestLimit Then
                    
                    TestNameInput = Report_TName_From_Instance("CalcC", "", 0)
                    TheExec.Flow.TestLimit resultVal:=ReturnDSPWave.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
                End If
                
            Case "ALG"
                SplitByLeftPara = Split(SplitByColon(2), "(")
                
                ALG_InterPoseFuncName = SplitByLeftPara(0)
                ALG_InterPoseArgcName = Replace(SplitByLeftPara(1), ")", "")
                
''                TheExec.Flow.SetInterpose 1, ALG_InterPoseFuncName, ALG_InterPoseArgcName    ' key, func name, arguments
''                TheExec.Flow.ExecuteInterpose 1

                Call Interpose(ALG_InterPoseFuncName, ALG_InterPoseArgcName)
            Case Else
            
        End Select
    Next i
End Function

Public Function IsConstant(InputArg As Variant) As Boolean
    
    On Error GoTo ChkFalse
    
    If InputArg / InputArg = 1 Then
        IsConstant = True
        Exit Function
    End If
    
ChkFalse:
    IsConstant = False
End Function

Public Function GetSplitKeyWord(InputFormula As String) As String

    If InStr(InputFormula, "+") > 0 Then
        GetSplitKeyWord = "+"
    ElseIf InStr(InputFormula, "-") > 0 Then
        GetSplitKeyWord = "-"
    ElseIf InStr(InputFormula, "*") > 0 Then
        GetSplitKeyWord = "*"
    ElseIf InStr(InputFormula, "/") > 0 Then
        GetSplitKeyWord = "/"
    End If
End Function

Public Function ProcessDSPCalculation(InputArg() As String, Operator As String, RTN_DSPWave As DSPWave) As Long
    Dim i As Long
    Dim Val_TempDSP As New DSPWave
    Dim Val_SerialDSP As New DSPWave
    Dim Val_ParallelDSP As New DSPWave
    Dim Val_TempConstant As Double
    Dim b_IsConstant As Boolean
    Dim WordWidth As Long
    Dim b_FirstIsConstant As Boolean
    Dim b_SecondIsConstant As Boolean
    b_FirstIsConstant = False
    b_SecondIsConstant = False
    
    Dim site As Variant
    
    For i = 0 To UBound(InputArg)
        If InStr(UCase(InputArg(i)), DC_Spec_Var) <> 0 Then
''            InputArg(i) = Spec_Evaluate_DC(InputArg(i))
''            InputArg(i) = Evaluate(InputArg(i))
            InputArg(i) = EvaluateForDCSpec(InputArg(i))
        End If
        
        b_IsConstant = IsConstant(InputArg(i))
        If b_IsConstant Then
            If i = 0 Then
                Val_TempConstant = CDbl(InputArg(i))
                b_FirstIsConstant = True
            Else
                b_SecondIsConstant = True
                
                If b_FirstIsConstant = True Then
                    Select Case Operator
                        Case "+"
                             Val_TempConstant = Val_TempConstant + CDbl(InputArg(i))
                        Case "-"
                             Val_TempConstant = Val_TempConstant - CDbl(InputArg(i))
                        Case "*"
                            Val_TempConstant = Val_TempConstant * CDbl(InputArg(i))
                        Case "/"
                            Val_TempConstant = Val_TempConstant / CDbl(InputArg(i))
                    End Select
                Else
                    Val_TempConstant = CDbl(InputArg(i))
                End If
                
            End If
        Else
'            Val_SerialDSP = GetStoredCaptureData(InputArg(i))
'
'            For Each site In theexec.sites
'                WordWidth = Val_SerialDSP(site).SampleSize
'
'                If WordWidth <> 0 Then
'                    Exit For
'                End If
'            Next site
'
'            Call rundsp.ConvertToLongAndSerialToParrel(Val_SerialDSP, WordWidth, Val_ParallelDSP)
            
            Dim val_dec As New SiteDouble         'Roger
            Dim val_dec_Temp As New SiteDouble    'Roger
            
            val_dec = GetStoredData(InputArg(i) & "_para")  'Roger , get decial resutl directly
            
            If i = 0 Then
                val_dec_Temp = val_dec  'Roger
                
'                Val_TempDSP = Val_ParallelDSP
                b_FirstIsConstant = False
            Else
                b_SecondIsConstant = False
                
                If b_FirstIsConstant = True Then
''                    Val_TempDSP.CreateConstant 0, 1
                    val_dec_Temp = val_dec      'Roger
'                    Val_TempDSP = Val_ParallelDSP
                Else
                    Select Case Operator
                        Case "+"
                             val_dec_Temp = val_dec_Temp.Add(val_dec)       'Roger
                             
'                             Call rundsp.DSP_Add(Val_TempDSP, Val_ParallelDSP)
                        Case "-"
                             val_dec_Temp = val_dec_Temp.Subtract(val_dec)     'Roger
                             
'                             Call rundsp.DSP_Subtract(Val_TempDSP, Val_ParallelDSP)
                        Case "*"
                            val_dec_Temp = val_dec_Temp.Multiply(val_dec)     'Roger
'                            Call rundsp.DSP_Multiply(Val_TempDSP, Val_ParallelDSP)
                        Case "/"
                            val_dec_Temp = val_dec_Temp.Divide(val_dec)     'Roger
'                            Call rundsp.DSP_Divide(Val_TempDSP, Val_ParallelDSP)
                    End Select
                End If
            End If
        End If
    Next i
    

    For Each site In TheExec.sites
        If b_FirstIsConstant Then
            If b_SecondIsConstant Then
                RTN_DSPWave(site).Element(0) = Val_TempConstant
            Else
                Select Case Operator

                    Case "+"
                        RTN_DSPWave(site).Element(0) = Val_TempConstant + val_dec_Temp
                        
                    Case "-"
                        RTN_DSPWave(site).Element(0) = Val_TempConstant - val_dec_Temp
                        
                    Case "*"
                        RTN_DSPWave(site).Element(0) = Val_TempConstant * val_dec_Temp
                        
                    Case "/"
                        RTN_DSPWave(site).Element(0) = Val_TempConstant / val_dec_Temp
                        
                End Select
            End If
            
        Else
            If b_SecondIsConstant Then
                Select Case Operator

                    Case "+"
                        RTN_DSPWave(site).Element(0) = val_dec_Temp + Val_TempConstant
                        
                    Case "-"
                        RTN_DSPWave(site).Element(0) = val_dec_Temp - Val_TempConstant
                        
                    Case "*"
                        RTN_DSPWave(site).Element(0) = val_dec_Temp * Val_TempConstant
                        
                    Case "/"
                       RTN_DSPWave(site).Element(0) = val_dec_Temp / Val_TempConstant
                        
                End Select
            
            Else
                RTN_DSPWave(site).Element(0) = val_dec_Temp
            End If
        End If
    Next site
    
End Function
Public Function ProcessStringForCalType(CalcEquationInput As String, ByRef CalcEquationPLD As CALC_EQUATION_PLD)
    Dim i As Long, j As Long
    CalcEquationPLD.Operator = GetSplitKeyWord(CalcEquationInput)
    Dim SplitByOperator() As String
    Dim SplitByLeftPara() As String
    SplitByOperator = Split(CalcEquationInput, CalcEquationPLD.Operator)
    
    For i = 0 To UBound(SplitByOperator)
        If InStr(SplitByOperator(i), "(") > 0 Then
            SplitByLeftPara = Split(SplitByOperator(i), "(")
            
            If i = 0 Then
                CalcEquationPLD.FirstPinName = Trim(SplitByLeftPara(0))
                CalcEquationPLD.FirstDictKey = Trim(Replace(SplitByLeftPara(1), ")", ""))
                CalcEquationPLD.b_FirstConstant = False
            Else
                CalcEquationPLD.SecondPinName = Trim(SplitByLeftPara(0))
                CalcEquationPLD.SecondDictKey = Trim(Replace(SplitByLeftPara(1), ")", ""))
                CalcEquationPLD.b_SecondConstant = False
            End If
        Else
            If i = 0 Then
                If InStr(UCase(SplitByOperator(i)), DC_Spec_Var) Then
                    SplitByOperator(i) = EvaluateForDCSpec(SplitByOperator(i))

                End If
                
                CalcEquationPLD.FirstPinName = SplitByOperator(i)
                CalcEquationPLD.FirstDictKey = ""
                CalcEquationPLD.b_FirstConstant = True
            Else
                If InStr(UCase(SplitByOperator(i)), DC_Spec_Var) Then
                    SplitByOperator(i) = EvaluateForDCSpec(SplitByOperator(i))
                End If
                
                CalcEquationPLD.SecondPinName = SplitByOperator(i)
                CalcEquationPLD.SecondDictKey = ""
                CalcEquationPLD.b_SecondConstant = True
            End If
        End If
    Next i
    
End Function

Public Function StandardCalcuation(ByRef CalcEquationPLD As CALC_EQUATION_PLD, ByRef RTN_Result As PinListData) As Long
    
    Dim DictKeyVal_First As New PinListData
    Dim DictKeyVal_Second As New PinListData
    Dim PinName_First As String
    Dim PinName_Second As String
    Dim TestLimitPinName As String
    TestLimitPinName = "'"
    Dim DifferentTestPinVal As New PinListData
    Dim TempTestPinVal As New PinListData
    Dim i As Long, j As Long
    Dim Pins_First() As String, NumberPins_First As Long
    Dim Pins_Second() As String, NumberPins_Second As Long
    Dim DummyPLD As New PinListData
    
    
    
    If CalcEquationPLD.b_FirstConstant = True Then
        ''EX: 10+11
        If CalcEquationPLD.b_SecondConstant = True Then
        
            TheExec.Datalog.WriteComment ("No pins information, only 2 constant to do calculation")
        
        ''EX: 10+PinA
        Else
            DictKeyVal_Second = GetStoredMeasurement(CalcEquationPLD.SecondDictKey)
            PinName_Second = CalcEquationPLD.SecondPinName
            
            ''20160922-Analyze pin structure, decompose them if structure is pin group,
            Call TheExec.DataManager.DecomposePinList(PinName_Second, Pins_Second(), NumberPins_Second)
            For i = 0 To NumberPins_Second - 1
                TempTestPinVal.AddPin(Pins_Second(i)) = DictKeyVal_Second.Copy(Pins_Second(i))
                
                Select Case CalcEquationPLD.Operator
                    Case "+"
                         TempTestPinVal.Pins(Pins_Second(i)) = DictKeyVal_Second.Pins(Pins_Second(i)).Add(CDbl(CalcEquationPLD.FirstPinName))
                    Case "-"
                         TempTestPinVal.Pins(Pins_Second(i)) = DictKeyVal_Second.Pins(Pins_Second(i)).Add(0 - CDbl(CalcEquationPLD.FirstPinName))
                    Case "*"
                        TempTestPinVal.Pins(Pins_Second(i)) = DictKeyVal_Second.Pins(Pins_Second(i)).Multiply(CDbl(CalcEquationPLD.FirstPinName))
                    Case "/"
                        If CDbl(CalcEquationPLD.FirstPinName) = 0 Then
                            Call TheExec.ErrorLogMessage("Divide 0 is not reasonable")
                            CalcEquationPLD.FirstPinName = 1
''                            TempTestPinVal.Pins(Pins_Second(i)) = DictKeyVal_Second.Pins(Pins_Second(i)).Multiply((1 / CDbl(CalcEquationPLD.FirstPinName)))
                            TempTestPinVal.Pins(Pins_Second(i)) = DictKeyVal_Second.Pins(Pins_Second(i)).Invert.Multiply(CDbl(CalcEquationPLD.FirstPinName))
                        Else
''                            TempTestPinVal.Pins(Pins_Second(i)) = DictKeyVal_Second.Pins(Pins_Second(i)).Multiply((1 / CDbl(CalcEquationPLD.FirstPinName)))
                            TempTestPinVal.Pins(Pins_Second(i)) = DictKeyVal_Second.Pins(Pins_Second(i)).Invert.Multiply(CDbl(CalcEquationPLD.FirstPinName))
                        End If
                End Select
            Next i
            RTN_Result = TempTestPinVal
        End If
        
    Else
        ''EX: PinA+12
        If CalcEquationPLD.b_SecondConstant = True Then
            DictKeyVal_First = GetStoredMeasurement(CalcEquationPLD.FirstDictKey)
            PinName_First = CalcEquationPLD.FirstPinName
            
            ''20160922-Analyze pin structure, decompose them if structure is pin group,
            Call TheExec.DataManager.DecomposePinList(PinName_First, Pins_First(), NumberPins_First)
            
            For i = 0 To NumberPins_First - 1
                TempTestPinVal.AddPin(Pins_First(i)) = DictKeyVal_First.Copy(Pins_First(i))
                
                Select Case CalcEquationPLD.Operator
                    Case "+"
                         TempTestPinVal.Pins(Pins_First(i)) = DictKeyVal_First.Pins(Pins_First(i)).Add(CDbl(CalcEquationPLD.SecondPinName))
                    Case "-"
                         TempTestPinVal.Pins(Pins_First(i)) = DictKeyVal_First.Pins(Pins_First(i)).Subtract(CDbl(CalcEquationPLD.SecondPinName))
                    Case "*"
                        TempTestPinVal.Pins(Pins_First(i)) = DictKeyVal_First.Pins(Pins_First(i)).Multiply(CDbl(CalcEquationPLD.SecondPinName))
                    Case "/"
                        If CDbl(CalcEquationPLD.SecondPinName) = 0 Then
                            Call TheExec.ErrorLogMessage("Divide 0 is not reasonable")
                            CalcEquationPLD.SecondPinName = 1
                            TempTestPinVal.Pins(Pins_First(i)) = DictKeyVal_First.Pins(Pins_First(i)).Divide(CDbl(CalcEquationPLD.SecondPinName))
                        Else
                            TempTestPinVal.Pins(Pins_First(i)) = DictKeyVal_First.Pins(Pins_First(i)).Divide(CDbl(CalcEquationPLD.SecondPinName))
                        End If
                End Select
            Next i
            RTN_Result = TempTestPinVal
            
        '' PinA(In5)-PinA(In6) or PinA(In5) - PinB(In5)
        Else
            DictKeyVal_First = GetStoredMeasurement(CalcEquationPLD.FirstDictKey)
            DictKeyVal_Second = GetStoredMeasurement(CalcEquationPLD.SecondDictKey)
            PinName_First = CalcEquationPLD.FirstPinName
            PinName_Second = CalcEquationPLD.SecondPinName
            
            ''20160922-Analyze pin structure, decompose them if structure is pin group,
            ''               Only process PinName_First, not allow different pin group to do calculation.
            Call TheExec.DataManager.DecomposePinList(PinName_First, Pins_First(), NumberPins_First)
            Call TheExec.DataManager.DecomposePinList(PinName_Second, Pins_Second(), NumberPins_Second)
            
            If PinName_First <> PinName_Second And NumberPins_First > 1 And NumberPins_Second > 1 Then
                Call TheExec.ErrorLogMessage("Not allow different pin group to do calculation")
                Exit Function
            End If
            
            If PinName_First = PinName_Second Then
                For i = 0 To NumberPins_First - 1
               If TheExec.DataManager.ChannelType(Pins_First(i)) = "N/C" Then
                    Else
                    TempTestPinVal.AddPin(Pins_First(i)) = DictKeyVal_First.Copy(Pins_First(i))
                
                    Select Case CalcEquationPLD.Operator
                        Case "+"
                            TempTestPinVal.Pins(Pins_First(i)) = DictKeyVal_First.Pins(Pins_First(i)).Add(DictKeyVal_Second.Pins(Pins_Second(i)))
                        Case "-"
                            TempTestPinVal.Pins(Pins_First(i)) = DictKeyVal_First.Pins(Pins_First(i)).Subtract(DictKeyVal_Second.Pins(Pins_Second(i)))
                        Case "*"
                            TempTestPinVal.Pins(Pins_First(i)) = DictKeyVal_First.Pins(Pins_First(i)).Multiply(DictKeyVal_Second.Pins(Pins_Second(i)))
                        Case "/"
                            TempTestPinVal.Pins(Pins_First(i)) = DictKeyVal_First.Pins(Pins_First(i)).Divide(DictKeyVal_Second.Pins(Pins_Second(i)))
                    End Select
                    End If
                Next i
                RTN_Result = TempTestPinVal
                
            Else
                
                TestLimitPinName = PinName_First & CalcEquationPLD.Operator & PinName_Second
                DifferentTestPinVal.AddPin (TestLimitPinName)
                Select Case CalcEquationPLD.Operator
                    Case "+"
                        DifferentTestPinVal.Pins(TestLimitPinName) = DictKeyVal_First.Pins(PinName_First).Add(DictKeyVal_Second.Pins(PinName_Second))
                    Case "-"
                        DifferentTestPinVal.Pins(TestLimitPinName) = DictKeyVal_First.Pins(PinName_First).Subtract(DictKeyVal_Second.Pins(PinName_Second))
                    Case "*"
                        DifferentTestPinVal.Pins(TestLimitPinName) = DictKeyVal_First.Pins(PinName_First).Multiply(DictKeyVal_Second.Pins(PinName_Second))
                    Case "/"
                        DifferentTestPinVal.Pins(TestLimitPinName) = DictKeyVal_First.Pins(PinName_First).Divide(DictKeyVal_Second.Pins(PinName_Second))
                End Select
                RTN_Result = DifferentTestPinVal
            End If
        End If
    
    End If
    
End Function


Public Function CreateSimulateDataDSPWave_Parallel(OutDspWave As DSPWave, DigCap_Sample_Size As Long, Optional DigCap_DataWidth As Long, Optional InitialDecVal As Double = 32)
    Dim site As Variant
    Dim i As Integer
    Dim TempDecVal As Double
    TempDecVal = InitialDecVal
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For i = 0 To DigCap_Sample_Size - 1
                If site = 0 Then
                    OutDspWave(site).Element(i) = TempDecVal + i
                Else
                    OutDspWave(site).Element(i) = TempDecVal + i + 128
                End If
            Next i
        Next site
    End If
End Function

Public Function DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData As String, OutDspWave As DSPWave, DigCap_Sample_Size As Long, Optional NumberPins As Long, Optional CUS_Str_MainProgram As String, Optional DigCap_Name As String, Optional Tname)

    Dim i As Long
    Dim SourceBitStrmWf As New DSPWave
    Dim NoOfSamples As New SiteLong
    Dim FlexibleConvertedDataWf() As New DSPWave
    Dim TestLimitWithTestName As New PinListData
    
    Dim Split_Num() As String
    Dim StartNum As Long
    
    '' 20151231 - Add rule to check new format that include test name and parse bits
    Dim ParseStringByBits As String
    Dim ParseStringForTestName As String
    Dim DSSC_Out_DecompseByComma() As String
    Dim DSSC_Out_DecompseByColon() As String
    Dim b_DSSC_Out_InvolveTestName As Boolean
    ParseStringByBits = ""
    ParseStringForTestName = ""
    b_DSSC_Out_InvolveTestName = False
    Dim DecomposeTestName() As String
    Dim DecomposeParseDigCapBit() As String
    
    ''20160807 - Add directionary to store DigCap DSPwave
    Dim ParseStringForDirectionary As String
    Dim DecomposeDirectionary() As String
    Dim b_ParseForDirectionary_Switch As Boolean
    b_ParseForDirectionary_Switch = False
    
    
    '' 20160212 - Process format by DSSC_OUT, capture word size is flexible, also parse with/without test name.
    If InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") <> 0 Then
        If InStr(UCase(CUS_Str_DigCapData), ":") <> 0 Then
            b_DSSC_Out_InvolveTestName = True
            DSSC_Out_DecompseByComma = Split(CUS_Str_DigCapData, ",")
            For i = 0 To UBound(DSSC_Out_DecompseByComma)
                DSSC_Out_DecompseByColon = Split(DSSC_Out_DecompseByComma(i), ":")
                If UBound(DSSC_Out_DecompseByColon) > 0 Then
                    If ParseStringByBits = "" And ParseStringForTestName = "" Then
                        ParseStringByBits = DSSC_Out_DecompseByColon(0)
                        ParseStringForTestName = DSSC_Out_DecompseByColon(1)
                        If UBound(DSSC_Out_DecompseByColon) = 2 Then
                            ParseStringForDirectionary = DSSC_Out_DecompseByColon(2) & ","
                        Else
                            ParseStringForDirectionary = ","
                        End If
                    Else
                        ParseStringByBits = ParseStringByBits & "," & DSSC_Out_DecompseByColon(0)
                        ParseStringForTestName = ParseStringForTestName & "," & DSSC_Out_DecompseByColon(1)
                        If b_ParseForDirectionary_Switch = False Then
                            b_ParseForDirectionary_Switch = True
                        Else
                            ParseStringForDirectionary = ParseStringForDirectionary & ","
                        End If
                        
                        If UBound(DSSC_Out_DecompseByColon) = 2 Then
                            ParseStringForDirectionary = ParseStringForDirectionary & DSSC_Out_DecompseByColon(2)
                        End If
                    
                    End If
                End If
            Next i
            ParseStringByBits = "DSSC_OUT," & ParseStringByBits
            DecomposeTestName = Split(ParseStringForTestName, ",")
            DecomposeDirectionary = Split(ParseStringForDirectionary, ",")
        Else
            ParseStringByBits = CUS_Str_DigCapData
        End If
        If Right(ParseStringByBits, 1) = "," Then
            ParseStringByBits = Left(ParseStringByBits, (Len(ParseStringByBits) - 1))
        End If
        DecomposeParseDigCapBit = Split(ParseStringByBits, ",")
        Dim StrParseDigCapBit As String
        
        For i = 1 To UBound(DecomposeParseDigCapBit)
            If i = 1 Then
                StrParseDigCapBit = DecomposeParseDigCapBit(i)
            Else
                StrParseDigCapBit = StrParseDigCapBit & "," & DecomposeParseDigCapBit(i)
            End If
        Next i
        DecomposeParseDigCapBit = Split(StrParseDigCapBit, ",")
        
        ReDim FlexibleConvertedDataWf(UBound(DecomposeParseDigCapBit)) As New DSPWave
        ''20160823-Store binary dsp wave after processed by DSSC_OUT
        Dim DSPWave_Binary() As New DSPWave
        ReDim DSPWave_Binary(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim StartIndex As Long
        StartIndex = 0
        SourceBitStrmWf = OutDspWave
        
        Dim TempDSPWaveForDict As New DSPWave
        TempDSPWaveForDict.CreateConstant 0, 1, DspLong
        
        Dim TempDSPWaveBinaryForDict() As New DSPWave
        ReDim TempDSPWaveBinaryForDict(UBound(DecomposeDirectionary)) As New DSPWave
        
        Dim site As Variant
        
        If InStr(UCase(CUS_Str_DigCapData), ":") <> 0 Then
            For i = 0 To UBound(DecomposeDirectionary)
                ''20160823-Store binary DSP wave by using Directionary
                If DecomposeDirectionary(i) <> "" Then
                    ''20161122-Store parallel DigCap value to Dictionary and have to create DSPWave to store element value
                    For Each site In TheExec.sites
                        TempDSPWaveForDict(site).Element(0) = SourceBitStrmWf(site).Element(i)
                    Next site
                    
                    ''20161124-Change to binary format and store to dictionary, the purpose is to match DigSrc fomat between serial and parallel
                    Call rundsp.DSPWaveDecToBinary(TempDSPWaveForDict, NumberPins, TempDSPWaveBinaryForDict(i))
                    Call AddStoredCaptureData(DecomposeDirectionary(i), TempDSPWaveBinaryForDict(i))
                End If
            Next i
        End If

        Dim TestNameInput As String
        Dim OutputTname_format() As String
        
        '' 20160317 - Test limit for DSSC_OUT
        If b_DSSC_Out_InvolveTestName = True Then '' Test limit with test name
            For i = 0 To UBound(DecomposeTestName)
                If LCase(DecomposeTestName(i)) = "skip" Then
                Else
                TestLimitWithTestName.AddPin (DecomposeTestName(i))
                TestLimitWithTestName.Pins((DecomposeTestName(i))).Value = SourceBitStrmWf.Element(i)
                
                TestNameInput = Report_TName_From_Instance("C", TestLimitWithTestName.Pins(DecomposeTestName(i)), DecomposeTestName(i), CInt(i), 0)
                TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i)), 0, 2 ^ NumberPins - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow
                End If
            Next i
        Else
            For i = 0 To DigCap_Sample_Size - 1
                TestNameInput = Report_TName_From_Instance("C", "", DecomposeTestName(i), CInt(i), 0)
                TheExec.Flow.TestLimit SourceBitStrmWf.Element(i), 0, 2 ^ NumberPins - 1, PinName:="DSSC_OUT_Code_" & i, Tname:=TheExec.DataManager.instanceName & "_DSSC_OUT_" & CStr(i), ForceResults:=tlForceFlow
            Next i
        End If
    End If
End Function

Public Function VIR_AnalyzedInputStringByAt(ByRef Measure_Pin_PPMU As String, ByRef ForceV As String, ByRef ForceI As String, ByRef MeasureI_Range As String) As Long
    '' 20160201 - Check input argumenets whether have "@" in the first character. Add it If no "@" in the beginning. Then remove it to process fomat.
    '' The purpose is to cover import issue. Ex:++
    Call CheckInputStringByAt(Measure_Pin_PPMU)
    Call CheckInputStringByAt(ForceV)
    Call CheckInputStringByAt(ForceI)
    Call CheckInputStringByAt(MeasureI_Range)
End Function

Public Function Checker_DictCalculated(InPutString As String) As Boolean
    Dim SplitByBracket() As String
    SplitByBracket = Split(InPutString, "[")
    If UBound(SplitByBracket) > 0 Then
        Checker_DictCalculated = True
    Else
        Checker_DictCalculated = False
    End If
End Function

Public Function AnalyzeDictCalculatedContent(InputContent As String, SrcDspWave As DSPWave)
    Dim site As Variant
    
    Dim TrimInputContent As String
    Dim SplitByOperator() As String
    Dim DictName As String
    Dim Operator As String
    Dim Data As Long
    Dim TempSrcDspWave As New DSPWave
    Dim TempSrcDspWaveDec As New DSPWave
    Dim TempSrcDspWaveSerial As New DSPWave
    Dim SampleSize As Long
    Dim DictNamelen As String
   
    TrimInputContent = Replace(InputContent, "[", "")
    TrimInputContent = Replace(TrimInputContent, "]", "")
    
    If InStr(TrimInputContent, "+") <> 0 Then
        Operator = "+"
    ElseIf InStr(TrimInputContent, "*") <> 0 Then
        Operator = "*"
    ElseIf InStr(TrimInputContent, "/") <> 0 Then
        Operator = "/"
    ElseIf InStr(TrimInputContent, "-") <> 0 Then
        Operator = "-"
        
    End If
    
    
    DictNamelen = InStrRev(TrimInputContent, Operator)
    
    
    SplitByOperator = Split(TrimInputContent, Operator)
    'DictName = SplitByOperator(0)
     DictName = Left(TrimInputContent, DictNamelen - 1)
    'Data = CLng(SplitByOperator(1))
     Data = SplitByOperator(UBound(SplitByOperator))
    
    TempSrcDspWave = GetStoredCaptureData(DictName)
    SampleSize = TempSrcDspWave.SampleSize
    Call rundsp.ConvertToLongAndSerialToParrel(TempSrcDspWave, SampleSize, TempSrcDspWaveDec)
    TempSrcDspWaveDec = TempSrcDspWaveDec.ConvertDataTypeTo(DspDouble)
    
    If Operator = "+" Then
        TempSrcDspWaveDec = TempSrcDspWaveDec.Add(Data)
    ElseIf Operator = "-" Then
        TempSrcDspWaveDec = TempSrcDspWaveDec.Subtract(Data)
        
    
    ElseIf Operator = "*" Then
        For Each site In TheExec.sites.Active
            TempSrcDspWaveDec.Element(0) = FormatNumber((TempSrcDspWaveDec.Element(0) * Data), 0)
        Next site
    ElseIf Operator = "/" Then
        For Each site In TheExec.sites.Active
            TempSrcDspWaveDec.Element(0) = FormatNumber((TempSrcDspWaveDec.Element(0) / Data), 0)
        Next site
    End If
    TempSrcDspWaveDec = TempSrcDspWaveDec.ConvertDataTypeTo(DspLong)
    Call rundsp.DSPWaveDecToBinary(TempSrcDspWaveDec, SampleSize, TempSrcDspWaveSerial)
    SrcDspWave = TempSrcDspWaveSerial
    SrcDspWave = SrcDspWave.ConvertDataTypeTo(DspLong)
End Function


Public Function GeneralDigSrcSetting(Pat As String, DigSrc_pin As PinList, DigSrc_Sample_Size As Long, DigSrc_DataWidth As Long, _
DigSrc_Equation As String, DigSrc_Assignment As String, _
DigSrc_FlowForLoopIntegerName As String, CUS_Str_DigSrcData As String, ByRef InDSPwave As DSPWave, _
Optional ByRef Rtn_SweepTestName As String, Optional MSB_First_Flag As Boolean = False) As Long

    Dim i, j, k As Long
    Dim b_StoreWholeDigSrc As Boolean, b_DigSrcFromDict As Boolean
    Dim DigSrcWholeDictName As String
    Dim DigSrc_Ary() As Long
    Dim site As Variant
    
    ''20161121- Check DigSrc type is serial or parallel
    Dim InDSPWave_Parallel As New DSPWave
    Dim DigSrcPinAry() As String, NumberPins As Long
    Dim b_SrcTypeIsParallel As Boolean
    Dim NoOfSamples As New SiteLong
    Dim CreateDigSrcDataSize As Long
    
    Dim CheckResult As Boolean
    Dim MergeEquationString As String
    Dim Equation_sample_size() As Long
    Dim StoreDictName As String
    
    Dim DigSrcInDspWave As New DSPWave

    On Error GoTo err
'==========================================================================================================================================================================
    If False Then
        DigSrc_Equation = LCase(DigSrc_Equation)
        DigSrc_Assignment = LCase(DigSrc_Assignment)
        
        
        If DigSrc_Sample_Size = 0 Then Exit Function
        If Instance_Data.Is_PreCheck_Func Then
            CheckResult = Instance_Data.DigSrcCheckCorrect
            Equation_sample_size = Instance_Data.DigSrcEquationSampleSize
        Else
            CheckResult = CheckDigSrcEquationAssignment(DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, Equation_sample_size)
        End If
        'If TheExec.DataManager.instanceName = "ADCLKG_T10_MN_CEBA0_C_FULP_AN_AA04_FRQ_JTG_VMX_ALLFV_SI_ADCLKGFX_T10_NV" Then Stop
        If CheckResult = True Then
            If Instance_Data.Is_PreCheck_Func Then
                CreateDigSrcDSPWave Instance_Data.MergeDigSrcEquation, DigSrcInDspWave, DigSrc_Sample_Size
            Else
                MergeEquationString = MergeDigSrcEquationAssignment(DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, StoreDictName)
                CreateDigSrcDSPWave MergeEquationString, DigSrcInDspWave, DigSrc_Sample_Size
            End If
        Else
            'Create Dummy DSPWave
            DigSrcInDspWave.CreateConstant 0, DigSrc_Sample_Size, DspLong
        End If
        
        If CheckResult = True Then
            PrintDigSrcEquationAssignment DigSrcInDspWave, Equation_sample_size, DigSrc_Equation, DigSrc_Assignment, DigSrc_pin.Value, DigSrc_Sample_Size, MSB_First_Flag
        Else
            TheExec.Datalog.WriteComment "[CheckDigSrcEquationAssignment something Error!!!!]"
        End If
        
        If StoreDictName <> "" Then AddStoredCaptureData StoreDictName, DigSrcInDspWave
        
        Call SetupDigSrcDspWave(Pat, DigSrc_pin, "Meas_Src", DigSrc_Sample_Size, DigSrcInDspWave)
        
        Exit Function
    End If
'==========================================================================================================================================================================
    ' For Multi-DigSrc format
    ' DigSrc_pin   : M1_JTAG_TDI+S1_JTAG_TDI+S2_JTAG_TDI;M2_JTAG_TDI+S3_JTAG_TDI+S4_JTAG_TDI
    ' DigAssigment : M_D2D_ZCPU=ZCPUM;M_D2D_ZCPD=ZCPDM;S_D2D_ZCPU=ZCPUM;S_D2D_ZCPD=ZCPDM;
    ' DigEquation  : D2D_ZCPU+D2D_ZCPD+D2D_ZCPU+D2D_ZCPD;D2D_ZCPD+D2D_ZCPD+D2D_ZCPD+D2D_ZCPD
    Dim AssemblyEquation As String
    Dim AssemblyEquationTemp As String
    Dim AssemblyEquation_Split() As String
    Dim DigSrcEquation_Split() As String
    Dim DigSrcPinsByEqu_Split() As String
    Dim DigSrcPinsByPin_Split() As String
    Dim DigSrc_pin_PinList As New PinList
    DigSrcEquation_Split = Split(DigSrc_Equation, ";")
    DigSrcPinsByEqu_Split = Split(DigSrc_pin, ";")

    For i = 0 To UBound(DigSrcPinsByEqu_Split)
        DigSrcPinsByPin_Split = Split(DigSrcPinsByEqu_Split(i), ",")
        For j = 0 To UBound(DigSrcPinsByPin_Split)
            DigSrc_pin_PinList.Value = DigSrcPinsByPin_Split(j)
            
            ' For Multi DigSrc Assembly
            If UBound(DigSrcPinsByPin_Split) <> 0 Then
                If UBound(DigSrcEquation_Split) <> UBound(DigSrcPinsByEqu_Split) Then
                    TheExec.AddOutput "Error : DigSrcEquation_Split mismatch with DigSrcPinsByEqu_Split !!"
                    On Error GoTo err
                Else
                    AssemblyEquation = ""
                    AssemblyEquationTemp = ""
                    AssemblyEquation_Split = Split(DigSrcEquation_Split(i), "+")
                    For k = 0 To UBound(AssemblyEquation_Split)
                        AssemblyEquationTemp = Left(DigSrcPinsByPin_Split(j), InStr(DigSrcPinsByPin_Split(j), "_")) + AssemblyEquation_Split(k)
                        If k = 0 Then
                            AssemblyEquation = AssemblyEquationTemp
                        Else
                            AssemblyEquation = AssemblyEquation + "+" + AssemblyEquationTemp
                        End If
                    Next k
                End If
            ElseIf UBound(DigSrcEquation_Split) <> -1 Then
                AssemblyEquation = DigSrcEquation_Split(0)
            End If
            
            Call TheExec.DataManager.DecomposePinList(DigSrc_pin_PinList, DigSrcPinAry(), NumberPins)
            If NumberPins > 1 Then
                b_SrcTypeIsParallel = True
        CreateDigSrcDataSize = DigSrc_Sample_Size * NumberPins
    Else
        b_SrcTypeIsParallel = False
        CreateDigSrcDataSize = DigSrc_Sample_Size
    End If

    If DigSrc_Sample_Size <> 0 Then
        If gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment ("======== Setup Dig Src Test Start ========")
            '=========================================180410 added by Oscar
            TheExec.Datalog.WriteComment "Src Bits = " & DigSrc_Sample_Size
                    TheExec.Datalog.WriteComment "SrcPin = " & DigSrc_pin_PinList.Value
        End If
        
        Dim Original_Assignment As String
        Dim DigSrc_Equation_Print As String
        Dim DigSrc_Assignment_Print As String
        Dim RegAssignChecker As Boolean
        
        Original_Assignment = DigSrc_Assignment
        If InStr(DigSrc_Assignment, "(") Then Call Checker_StoreWholeDigSrcInDict(DigSrc_Assignment, DigSrcWholeDictName) 'split dict(test1_ModeA) into dictName and test1_ModeA for replacement
        If InStr(DigSrc_Assignment, "=") = 0 Then Call GetRegFromDictByTestByMode(DigSrc_Assignment, RegAssignChecker)
        b_DigSrcFromDict = Checker_DigSrcFromDict(DigSrc_Assignment)
        
'''''                DigSrc_Equation_Print = vbCrLf & Print_NewFormat(DigSrcEquation_Split(i)) & vbCrLf ''''''''''process new format for printing
        DigSrc_Assignment_Print = Print_NewFormat(DigSrc_Assignment)

        If RegAssignChecker Then
            'Register1 = ABC_CA, repeat = 1001, Dict
            DigSrc_Assignment_Print = vbCrLf & Original_Assignment & "(" & DigSrc_Assignment_Print & ")" & vbCrLf 'test1_ModeA then use the original assignment & (replaced assignment)
                Else
                    DigSrc_Assignment_Print = vbCrLf & Original_Assignment & vbCrLf 'repeat=1000, register=1000, Dict(A=ABC_A) keeps the original
                End If
                '=========================================180410 added by Oscar
                If b_DigSrcFromDict Then
                    InDSPwave = GetStoredCaptureData(DigSrc_Assignment)

                    '*******************************New Feature for trimcode table*******************************
                    'Added by  20190509
                    For Each site In TheExec.sites.Active
                        If InDSPwave.DataSize <> DigSrc_Sample_Size Then
                            TheExec.AddOutput "Error : DigSrc Samplesize is mismatch"
                            GoTo err:
                        End If
                    Next site
                    '********************************************************************************************

                    If gl_Disable_HIP_debug_log = False Then
                        If UBound(DigSrcEquation_Split) <> -1 Then
                            TheExec.Datalog.WriteComment "DataSequence:" & DigSrcEquation_Split(i)
                        End If
                        TheExec.Datalog.WriteComment "Assignment:" & DigSrc_Assignment
                TheExec.Datalog.WriteComment "Output String [ LSB(L) ==> MSB(R) ]:"
            End If
            
            For Each site In TheExec.sites.Active
                        DigSrc_Ary = InDSPwave(site).Data
                       ' TheExec.Datalog.WriteComment ("[Site " & site & "]")
                        Call Printout_DigSrc(DigSrc_Ary, DigSrc_Sample_Size, DigSrc_DataWidth, , site)
                    Next site

                    Call SetupDigSrcDspWave(Pat, DigSrc_pin_PinList, "Meas_Src", DigSrc_Sample_Size, InDSPwave) ''Setup DSPWave, 20190508

                    Dim Table_Decvalue As String

                    If InStr(DigSrc_Assignment, "digsrctable") <> 0 Then
                          Table_Decvalue = ""
                          gl_SweepNum = ""
                           For k = 0 To DigSrc_Sample_Size - 1 Step 1
                              Table_Decvalue = Table_Decvalue & InDSPwave.Element(i)
                           Next k
                           gl_SweepNum = BinStr2HexStr(Table_Decvalue, (DigSrc_Sample_Size \ 4) + 1)
                    End If
                Else
                   ''20160826 - Check DigSrc_Assignment's content to decide whether to store InDSPWave to the Dictionary
            b_StoreWholeDigSrc = Checker_StoreWholeDigSrcInDict(DigSrc_Assignment, DigSrcWholeDictName)
            
            '----------------Added by Oscar 0424------------
            If InStr(Original_Assignment, "(") <> 0 Then b_StoreWholeDigSrc = True 'Test1_ModeA has already split but the flag was turn false on the last func.
                    '----------------Added by Oscar 0424------------

                    ''20160805 - Analyze DigSrc_Equation and DigSrc_Assignment inout format whether violate design rule
                    Call AnalyzeDigSrcEquationAssignmentContent(AssemblyEquation, DigSrc_Assignment)

                    If DigSrc_FlowForLoopIntegerName <> "" Then
                        Call DSSCSrcBitFromFlowForLoop(DigSrc_FlowForLoopIntegerName, DigSrc_DataWidth, AssemblyEquation, DigSrc_Assignment, CUS_Str_DigSrcData, Rtn_SweepTestName)
                    End If
            
            If CUS_Str_DigSrcData <> "" And UCase(CUS_Str_DigSrcData) Like UCase("*VOLH_Sweep*") Then
                        Call VOLH_Sweep(CUS_Str_DigSrcData, DigSrc_Assignment)
                    End If

                    ' CS Edited by 20191755
                    ' For Multi DigSrc Function
                    If gl_Disable_HIP_debug_log = False And Len(AssemblyEquation) < 8000 Then
                        TheExec.Datalog.WriteComment "DataSequence:" & AssemblyEquation
                        TheExec.Datalog.WriteComment "Assignment:" & DigSrc_Assignment
                        If MSB_First_Flag Then
                            TheExec.Datalog.WriteComment "Output String [ MSB(L) ==> LSB(R) ]:"
                        Else
                            TheExec.Datalog.WriteComment "Output String [ LSB(L) ==> MSB(R) ]:"
                        End If
                    End If


                    For Each site In TheExec.sites.Active
                        Call Create_DigSrc_Data(DigSrc_pin_PinList, DigSrc_DataWidth, CreateDigSrcDataSize, AssemblyEquation, DigSrc_Assignment, InDSPwave, site, , NumberPins, MSB_First_Flag)
                    Next site

                    ''20161121-Setup DigSrc by serial or parallel
                    If b_SrcTypeIsParallel Then
                        rundsp.BitWf2Arry InDSPwave, NumberPins, NoOfSamples, InDSPWave_Parallel
                        Call SetupDigSrcDspWave(Pat, DigSrc_pin_PinList, "Meas_Src_Parallel", DigSrc_Sample_Size, InDSPWave_Parallel)
                    Else
            
                If b_StoreWholeDigSrc Then
                    Call AddStoredCaptureData(DigSrcWholeDictName, InDSPwave)
                End If
                        Call SetupDigSrcDspWave(Pat, DigSrc_pin_PinList, "Meas_Src", DigSrc_Sample_Size, InDSPwave)
            End If
        End If
                'If gl_Disable_HIP_debug_log = False Then theexec.DataLog.WriteComment ("Src Pin =" & DigSrc_pin.Value)
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======== Setup Dig Src Test End   ========")
            End If
        Next j
    Next i
    
    Exit Function
err:
    TheExec.AddOutput "GeneralDigSrcSetting : Error"
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function GeneralDigCapSetting(Pat As String, DigCap_Pin As PinList, DigCap_Sample_Size As Long, ByRef OutDspWave As DSPWave) As Long
    Dim i As Long
    
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

End Function

Public Function HardIP_Alarm_off()
    If TheExec.Flow.EnableWord("HardIPAlarm") = True Then
        Dim i As Long, j As Long, p As Long
    
         Dim PinAry() As String, PinCnt As Long
    
 '   If TheExec.Flow.EnableWord("HardIPAlarm") = True Then
    
         TheExec.DataManager.DecomposePinList "All_UVS256,VDD_Warm", PinAry(), PinCnt
            For i = 0 To PinCnt - 1
            
         For Each site In TheExec.sites
            TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmOpenKelvinDUT) = tlAlarmOff
            TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmFoldCurrentLimitTimeout) = tlAlarmOff
            TheHdw.DCVS.Pins(PinAry(i)).Alarm(tlDCVSAlarmSourceFoldCurrentLimitTimeout) = tlAlarmOff
       Next site
       Next i
       
              TheExec.DataManager.DecomposePinList "VDDIO18_AOP,ANALOGMUX_OUT", PinAry(), PinCnt
            For i = 0 To PinCnt - 1
            
        ' For Each Site In TheExec.sites
            TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVSAlarmOpenKelvinDUT) = tlAlarmOff
            TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVIAlarmOpenKelvin) = tlAlarmOff
            TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVIAlarmDGS) = tlAlarmOff
           ' TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVIAlarmGuard) = tlAlarmOff
            TheHdw.DCVI.Pins(PinAry(i)).Alarm(tlDCVIAlarmMode) = tlAlarmOff
      ' Next Site
       Next i
    TheHdw.DIB.LeavePowerOn = False
    TheHdw.DCVI.Pins("All_DCVI").Alarm(tlDCVIAlarmDGS) = tlAlarmOff
    End If
End Function

Public Function VIR_ProcessInputString(TestSequence As String, ForceI As String, ForceV As String, Measure_Pin_PPMU As String, MeasureI_Range As String, Meas_StoreName As String, Interpose_PreMeas As String, _
                                                           ByRef TestSequenceArray() As String, ByRef ForceISequenceArray() As String, ByRef ForceVSequenceArray() As String, ByRef TestPinArrayIV() As String, _
                                                           ByRef TestIrange() As String, ByRef MeasStoreName_Ary() As String, ByRef Interpose_PreMeas_Ary() As String) As Long
                                                           
    Dim i As Long
    
    TestSequenceArray = Split(TestSequence, ",")
    
    ''20170706-Check "&" firstly for VAR_H, process "+" if "&" not exist
    ''20170706-Check "&" firstly for VAR_H, process "+" if "&" not exist
''    ForceISequenceArray = Split(ForceI, "+")
    If InStr(ForceI, "|") <> 0 Then
        ForceISequenceArray = Split(ForceI, "|")
    ElseIf InStr(UCase(ForceI), DC_Spec_Var) <> 0 Then
        ForceISequenceArray = Split(ForceI, "|")
    Else
        ForceISequenceArray = Split(ForceI, "|")
    End If


    ''20170110-Check "&" firstly for VAR_H, process "+" if "&" not exist
    If InStr(ForceV, "|") <> 0 Then
        ForceVSequenceArray = Split(ForceV, "|")
    ElseIf InStr(UCase(ForceV), DC_Spec_Var) <> 0 Then
       ForceVSequenceArray = Split(ForceV, "|")
    Else
        ForceVSequenceArray = Split(ForceV, "|")
    End If
    
    TestPinArrayIV = Split((Measure_Pin_PPMU), "+")
    TestIrange = Split(MeasureI_Range, "+")
    ''20160906 - Analyze Meas_StoreName and store the measurement for futher use.
    MeasStoreName_Ary = Split(Meas_StoreName, "+")
    ''20160923 - Analyze Interpose_PreMeas to force setting with different sequence.
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    
End Function

Public Function GrayCode2Bin_TTR(ByVal IsUnsigned As Boolean, ByVal InWf As DSPWave, ByRef OutWf As DSPWave, ByRef OutWf_Dec As DSPWave) As Long
    Dim site As Variant
    For Each site In TheExec.sites
        OutWf.CreateConstant 0, InWf.SampleSize, DspLong
        OutWf_Dec.CreateConstant 0, 1, DspLong
    Next site
    Dim i As Long
    Dim MSB_ElementNumForSignUnsign As Long
    Dim SignUnsignDiffBit As Long
    If IsUnsigned Then
        SignUnsignDiffBit = 1
    Else
        SignUnsignDiffBit = 2
    End If
    Dim index As Long
    
    For Each site In TheExec.sites
        MSB_ElementNumForSignUnsign = InWf(site).SampleSize - 1
        index = 0
        For i = InWf(site).SampleSize - SignUnsignDiffBit To 0 Step -1
            If index = 0 Then
                OutWf(site).Element(i) = InWf(site).Element(i)
            Else
                If InWf(site).Element(i) = OutWf(site).Element(i + 1) Then
                    OutWf(site).Element(i) = 0
                Else
                    OutWf(site).Element(i) = 1
                End If
            End If
            index = index + 1
        Next i
        OutWf_Dec(site) = OutWf(site).ConvertStreamTo(tldspParallel, OutWf(site).SampleSize, 0, Bit0IsMsb)
    
        If IsUnsigned = True Then
        Else
            If InWf(site).Element(MSB_ElementNumForSignUnsign) = 1 Then
                OutWf_Dec(site).Element(0) = OutWf_Dec(site).Element(0) * -1
            Else
            End If
        End If
    Next site

End Function


Public Function DSSC_Search_par_run(Pat As String, srcPin As String, code As SiteLong, MeasPin As String, Res As SiteDouble, TrimCodeSize As Long, Optional TrimRepeat As Long = 1)
    
    Dim sigName As String, srcWave As New DSPWave, site As Variant
    
    Dim DigSrcCodeSize As Long
    Dim i As Long, j As Long
    DigSrcCodeSize = TrimCodeSize * TrimRepeat

    sigName = "DSSC_Search_Code"
    TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals.Add sigName
    srcWave.CreateConstant 0, DigSrcCodeSize, DspLong
    For Each site In TheExec.sites
''        srcwave.Element(0) = code And 1
''        srcwave.Element(1) = (code And 2) \ 2
''        srcwave.Element(2) = (code And 4) \ 4
''        srcwave.Element(3) = (code And 8) \ 8
        
        For i = 0 To TrimCodeSize - 1
            If i = 0 Then
                srcWave.Element(i) = code And 1
            Else
                srcWave.Element(i) = (code And (2 ^ i)) \ (2 ^ i)
            End If
        Next i
        
        For i = 0 To TrimRepeat - 1
            For j = 0 To TrimCodeSize - 1
                srcWave.Element(i * TrimCodeSize + j) = srcWave.Element(j)
            Next j
        Next i
        
        TheExec.WaveDefinitions.CreateWaveDefinition "WaveDef" & site, srcWave, True
        With TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals(sigName)
            .WaveDefinitionName = "WaveDef" & site
            .SampleSize = DigSrcCodeSize
            .Amplitude = 1
            .LoadSamples
            .LoadSettings
        End With
    Next site
    TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals.DefaultSignal = sigName
    
    TheHdw.Patterns(Pat).start
    TheHdw.Digital.Patgen.FlagWait cpuA, 0
    TheHdw.Wait 10 * ms
    
    Call DebugPrintFunc_PPMU("")
    
    Res = TheHdw.DCVI.Pins(MeasPin).Meter.Read(tlStrobe, 10)
    
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites.Active
            TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Voltage = " & Res(site)
        Next site
    End If
    
    TheHdw.Digital.Patgen.Continue 0, cpuA
    TheHdw.Digital.Patgen.HaltWait
End Function

Public Function DSSC_Special_Str_Filter(DSSC_Str As String, Special_Str As String, DSPWave_Org As DSPWave, _
                                 ByRef DSSC_Sub_Str As String, ByRef DSPWave_Special As DSPWave, ByRef DSPWave_Other As DSPWave) As Long   '' TYCHENGG
                                 
    Dim SplitByComma() As String
    Dim SplitByColumn() As String
    Dim NumOfReg As Integer
    Dim ListOfSpecialCase As String '' ex. "0,1,3,7...,"
    Dim ArrayOfSpecialCase() As String
    Dim NumOfSpecialCase As Integer
    Dim NumOfSpecialBits As Integer
    Dim NumOfOtherBits As Integer
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim DigCapIndex As Integer
    Dim SpecialBitsIndex As Integer
    Dim OtherBitsIndex As Integer
    Dim StartIndex As Integer
    Dim TempStrArr() As String
        

    SplitByComma = Split(DSSC_Str, ",")
    NumOfReg = UBound(SplitByComma) - 1 ''discard the first one and the last one
    ReDim SplitByColumn(NumOfReg - 1, 1)
    For i = 0 To NumOfReg - 1
        TempStrArr = Split(SplitByComma(i + 1), ":")
        For j = 0 To 1
            SplitByColumn(i, j) = TempStrArr(j)
        Next j
        
    Next i
    ListOfSpecialCase = ""
    For i = 0 To NumOfReg - 1
        If LCase(SplitByColumn(i, 1)) Like "*" & LCase(Special_Str) & "*" Then
            ListOfSpecialCase = ListOfSpecialCase & CStr(i) & ","
        End If
        
    Next i
    ArrayOfSpecialCase = Split(ListOfSpecialCase, ",")
    NumOfSpecialCase = UBound(ArrayOfSpecialCase) '' discard the last one
    
    '' Initialize
    DSSC_Sub_Str = SplitByComma(0) & "," '' (rdx_)DSSC_OUT,
    
    NumOfSpecialBits = 0
    NumOfOtherBits = 0
    StartIndex = 0
    For i = 0 To NumOfReg - 1
        
        If LCase(SplitByColumn(i, 1)) Like "*" & LCase(Special_Str) & "*" Then
            NumOfSpecialBits = NumOfSpecialBits + SplitByColumn(i, 0)
                
        Else
            NumOfOtherBits = NumOfOtherBits + SplitByColumn(i, 0)
            DSSC_Sub_Str = DSSC_Sub_Str & SplitByComma(i + 1) & "," '' (rdx_)DSSC_OUT,7:ddr0_bistsweep...,
        End If
        
        
    Next i
        
    For Each site In TheExec.sites.Active '' Create DSPWave space for both site
        DSPWave_Special.CreateConstant 0, NumOfSpecialBits
        DSPWave_Other.CreateConstant 0, NumOfOtherBits
    Next site

    ''-------------
    DigCapIndex = 0
    SpecialBitsIndex = 0
    OtherBitsIndex = 0
    StartIndex = 0
    For i = 0 To NumOfReg - 1
      
        If LCase(SplitByColumn(i, 1)) Like "*" & LCase(Special_Str) & "*" Then
            For Each site In TheExec.sites.Active
                For k = 0 To SplitByColumn(i, 0) - 1
                    DSPWave_Special.Element(SpecialBitsIndex + k) = DSPWave_Org.Element(DigCapIndex + k)  '20160407
                Next k
            Next site
            SpecialBitsIndex = SpecialBitsIndex + SplitByColumn(i, 0)
                    
        Else
            For Each site In TheExec.sites.Active
                For k = 0 To SplitByColumn(i, 0) - 1
                    DSPWave_Other.Element(OtherBitsIndex + k) = DSPWave_Org.Element(DigCapIndex + k)
                Next k
            Next site
            OtherBitsIndex = OtherBitsIndex + SplitByColumn(i, 0)
        End If
        DigCapIndex = DigCapIndex + SplitByColumn(i, 0)
       
    Next i
    
End Function

Public Function SimulatePreCheckOutputFreq(MeasureF_Pin As PinList, ByRef MeasureFreq As PinListData) As Long
    Dim site As Variant
    For Each site In TheExec.sites.Active
        If site = 0 Then
            MeasureFreq.Pins(MeasureF_Pin).Value(site) = 1000000
        ElseIf site = 1 Then
            MeasureFreq.Pins(MeasureF_Pin).Value(site) = 992000
        ElseIf site = 2 Then
''            MeasureFreq.Pins(MeasureF_Pin).Value(Site) = 992000
        ElseIf site = 3 Then
''            MeasureFreq.Pins(MeasureF_Pin).Value(Site) = 1002000
        End If
    Next site
End Function

Public Function GetFlowTestName(ByRef FlowTestNme() As String) As Long
    ''Get the limits info
    Dim FlowLimitsInfo As IFlowLimitsInfo
''    Dim TNamesVals() As String
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    
    If FlowLimitsInfo Is Nothing Then
        Exit Function
    Else
        
        
        Call FlowLimitsInfo.GetTNames(FlowTestNme)
    End If
    
End Function

Public Function VIR_CheckForceVal(ByRef ForceI As String, ByRef ForceV As String) As Long

If InStr(UCase(ForceI), "CP") <> 0 Or InStr(UCase(ForceI), "FT") <> 0 Then
    Call VIR_ProcessForceVal(ForceI)
Else
    ForceI = ForceI
End If

If InStr(UCase(ForceV), "CP") <> 0 Or InStr(UCase(ForceV), "FT") <> 0 Then
    Call VIR_ProcessForceVal(ForceV)
Else
    ForceV = ForceV
End If

End Function

Public Function VIR_ProcessForceVal(ByRef ForceVal As String) As Long

    Dim i As Long, j, k As Long
    Dim SplitByAdd() As String
    Dim SplitByComma() As String
    Dim splitbyand() As String
    Dim SplitByColon() As String
    Dim TempForceVal As String
    Dim FinalTempForceVal As String
    TempForceVal = ""
    FinalTempForceVal = ""
    If InStr(ForceVal, "+") <> 0 Then
        SplitByAdd = Split(ForceVal, "+")
        For k = 0 To UBound(SplitByAdd)
            SplitByComma = Split(SplitByAdd(k), ",")
            
            For i = 0 To UBound(SplitByComma)
                SplitByColon = Split(SplitByComma(i), ":")
                If UBound(SplitByColon) > 0 Then
                    If UCase(SplitByColon(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Or UCase(SplitByColon(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
                    Else
                        If UCase(SplitByColon(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
                            TempForceVal = TempForceVal & "," & SplitByColon(1)
                        ElseIf UCase(SplitByColon(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
                            TempForceVal = TempForceVal & "," & SplitByColon(1)
                        Else

                        End If
                    End If
                Else
                    TempForceVal = TempForceVal & "," & SplitByColon(0)
                End If
            Next i
            
            If Left(TempForceVal, 1) = "," Then
                TempForceVal = Right(TempForceVal, Len(TempForceVal) - 1)
            End If
            If k = 0 Then
                If TempForceVal <> "" Then
                    FinalTempForceVal = TempForceVal
                Else
                End If
            Else
                FinalTempForceVal = FinalTempForceVal & "+" & TempForceVal
            End If
            
            TempForceVal = ""
        Next k
        ForceVal = FinalTempForceVal
    Else
        SplitByComma = Split(ForceVal, ",")
        For i = 0 To UBound(SplitByComma)
            SplitByColon = Split(SplitByComma(i), ":")
            If UBound(SplitByColon) > 0 Then
                If UCase(SplitByColon(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Or UCase(SplitByColon(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
                Else
                    If UCase(SplitByColon(0)) = "CP" And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
                        TempForceVal = TempForceVal & "," & SplitByColon(1)
                    ElseIf UCase(SplitByColon(0)) = "FT" And InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
                        TempForceVal = TempForceVal & "," & splitbyand(1)
                    Else

                    End If
                End If
            Else
                TempForceVal = TempForceVal & "," & SplitByColon(0)
            End If

        Next i
        If Left(TempForceVal, 1) = "," Then
            TempForceVal = Right(TempForceVal, Len(TempForceVal) - 1)
        End If
        If Right(TempForceVal, 1) = "," Then
            TempForceVal = Left(TempForceVal, Len(TempForceVal) - 1)
        End If
        ForceVal = TempForceVal
    End If
    

End Function

Public Function SimulateFlowForSweep(FlowVal As String) As Long
    
    If TheExec.TesterMode = testModeOffline Then
        If (TheExec.DataManager.instanceName Like "*VIR_FlowForSweepVoltage*") = True Then
            Dim index As Long
            index = TheExec.Flow.var("ForIndex").Value
            FlowVal = CStr((index + 1) * 100)
        End If
    End If
End Function

Public Function PrLoadPattern(PatName As String) As Long
    
    Dim PattArray() As String
    Dim patt As Variant
    Dim Pat As String
    Dim PatCount As Long
    
    If PatName = "" Then Exit Function
    
    ' Run validation
    Call ValidatePattern(PatName)
    Call PATT_GetPatListFromPatternSet(PatName, PattArray, PatCount)

    For Each patt In PattArray
        Pat = CStr(patt)
        Call ValidatePattern(Pat)
    Next patt

End Function


Public Function Freq_WalkingStrobe_Meas_VOD_Diff(MeasureF_Pin_Differential As PinList, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, _
    Optional MeasF_WalkingStrobe_StepVoltage As Double, Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, _
    Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double)
    
    Dim site As Variant
    Dim MeasF_WalkingStrobe_Step As Long
    MeasF_WalkingStrobe_Step = (MeasF_WalkingStrobe_EndV - MeasF_WalkingStrobe_StartV) / MeasF_WalkingStrobe_StepVoltage + 1
    
    Dim MeasFreq_WKStrobe() As New PinListData
    ReDim MeasFreq_WKStrobe(MeasF_WalkingStrobe_Step) As New PinListData
    Dim WalkStrobe_i As Long
    Dim WalkStrobe_j As Long
    ''setup and measure Freq base on VOL and VOH setting.
    Dim WalkingStrobe_stepV As Double
    WalkingStrobe_stepV = (MeasF_WalkingStrobe_EndV - MeasF_WalkingStrobe_StartV) / MeasF_WalkingStrobe_Step
    
    Dim DiffPinGroup() As String
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim i As Long, j As Long, k As Long
    DiffPinGroup = Split(MeasureF_Pin_Differential, ",")
    
    Dim DiffPinGroupPinList As New PinList
    
    Dim MeasurePin As String
    Dim MeasurePin_Opposite As String
    Dim FreqAccessPin As String
    Dim Record_Final_Mid_VOD() As New SiteDouble
    Dim b_UpdateVOD_Flag() As New SiteBoolean
    Dim Val_UpdateToVt As Double
    Dim Default_VOD As Double
    
    
    For i = 0 To UBound(DiffPinGroup)
        TheExec.DataManager.DecomposePinList DiffPinGroup(i), Pin_Ary, Pin_Cnt
        DiffPinGroupPinList.Value = DiffPinGroup(i)
        Default_VOD = TheHdw.Digital.Pins(DiffPinGroupPinList).DifferentialLevels.Value(chVod)
        ReDim Record_Final_Mid_VOD(Pin_Cnt - 1) As New SiteDouble
        ReDim b_UpdateVOD_Flag(Pin_Cnt - 1) As New SiteBoolean

        For j = 0 To Pin_Cnt - 1
            
            MeasurePin = Pin_Ary(j)
            If InStr(UCase(MeasurePin), "_P") <> 0 Then
                MeasurePin_Opposite = Replace(UCase(MeasurePin), "_P", "_N")
            ElseIf InStr(UCase(MeasurePin), "_N") <> 0 Then
                MeasurePin_Opposite = Replace(UCase(MeasurePin), "_N", "_P")
            End If
            
            If InStr(UCase(MeasurePin), "_N") <> 0 Then
                FreqAccessPin = Replace(UCase(MeasurePin), "_N", "_P")
            Else
                FreqAccessPin = MeasurePin
            End If
            
            TheHdw.Digital.Pins(MeasurePin_Opposite).Disconnect
            TheHdw.Wait 10 * us
            With TheHdw.PPMU.Pins(MeasurePin_Opposite)
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                .Connect
                .Gate = tlOn
                .ForceV 0, 0
            End With
            
            For WalkStrobe_i = 0 To MeasF_WalkingStrobe_Step
                TheHdw.Digital.Pins(DiffPinGroupPinList).DifferentialLevels.Value(chVod) = MeasF_WalkingStrobe_StartV + WalkStrobe_i * WalkingStrobe_stepV
                
                Call Freq_MeasFreqSetup(DiffPinGroupPinList, MeasF_WalkingStrobe_interval, VOH)
                Call HardIP_Freq_MeasFreqStart(DiffPinGroupPinList, MeasF_WalkingStrobe_interval, MeasFreq_WKStrobe(WalkStrobe_i), 0)
            Next WalkStrobe_i
                
            ''analyze measurement data to decide which VOH/VOL level shiuld be used for measurement.
            Dim Record_Temp_VOD As Double
            Dim Record_Min_VOD As Double
            Dim Record_Max_VOD As Double
            Dim Record_Mid_VOD As Double
            
            For Each site In TheExec.sites
            
                Record_Min_VOD = 9999
                Record_Max_VOD = -9999
                For WalkStrobe_i = 0 To MeasF_WalkingStrobe_Step
                    If MeasFreq_WKStrobe(WalkStrobe_i).Pins(FreqAccessPin).Value(site) > MeasF_WalkingStrobe_miniFreq Then
                        Record_Temp_VOD = MeasF_WalkingStrobe_StartV + WalkStrobe_i * WalkingStrobe_stepV
                        If Record_Temp_VOD > Record_Max_VOD Then Record_Max_VOD = Record_Temp_VOD
                        If Record_Temp_VOD < Record_Min_VOD Then Record_Min_VOD = Record_Temp_VOD
                    End If
                Next WalkStrobe_i
                
                If TheExec.TesterMode = testModeOffline Then
                    Record_Max_VOD = 0.5 + i * 0.1 + j * 0.1 + site * 0.01
                    Record_Min_VOD = 0.5 - i * 0.1 - j * 0.1 - site * 0.02
                End If
                
                If Record_Min_VOD <> 9999 Then
                    Record_Mid_VOD = (Record_Max_VOD + Record_Min_VOD) / 2
                    Record_Final_Mid_VOD(j).Value(site) = Record_Mid_VOD
                    b_UpdateVOD_Flag(j)(site) = True
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & MeasurePin & " , " & " Record VOD = " & Record_Mid_VOD & " V"
                Else
                    b_UpdateVOD_Flag(j)(site) = False
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & MeasurePin & " , " & " Record VOD = Default, search fail"
                End If
            Next site
            
            TheHdw.PPMU.Pins(MeasurePin_Opposite).Disconnect
            TheHdw.Digital.Pins(MeasurePin_Opposite).Connect
        Next j
        
        For Each site In TheExec.sites
            If b_UpdateVOD_Flag(0)(site) = True And b_UpdateVOD_Flag(1)(site) = True Then
                Val_UpdateToVt = (Record_Final_Mid_VOD(0).Value(site) + Record_Final_Mid_VOD(1).Value(site)) / 2
                TheHdw.Digital.Pins(DiffPinGroupPinList).DifferentialLevels.Value(chDiff_Vt) = Val_UpdateToVt
                TheHdw.Digital.Pins(DiffPinGroupPinList).DifferentialLevels.Value(chVod) = Default_VOD
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & DiffPinGroupPinList & " , " & " Update Differential Vt = " & Val_UpdateToVt & " V"
            ElseIf gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment "Site= " & site & " , " & " Pin= " & DiffPinGroupPinList & " , " & " Vt = Default, search fail"
            End If
        
        Next site
        
    Next i
End Function


Public Function UP1600_PPMU_Measure_R_SE(MeasurePin As String, ForceVoltStr As String, MeasureCurrRange As Double, Optional RAK_Flag As Enum_RAK = 0, _
                                                                                 Optional ByRef RTN_Imped_Val As PinListData, Optional b_PD_Mode As Boolean = True) As Long

    Dim MeasureValue As New PinListData
    Dim Imped As New PinListData
    Dim Pin  As Variant
    Dim site As Variant

    
    Dim ForceVoltVal As Double
    ForceVoltVal = CDbl(ForceVoltStr)
    
    TheHdw.Digital.Pins(MeasurePin).Disconnect
    TheHdw.Wait 10 * us
    
    '' Initial force I to 0 and force V by your specified
    With TheHdw.PPMU.Pins(MeasurePin)
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
        .ForceV ForceVoltVal, MeasureCurrRange
    End With
    
    TheHdw.Wait 1 * ms
    
    DebugPrintFunc_PPMU CStr(MeasurePin)
    
    MeasureValue = TheHdw.PPMU.Pins(MeasurePin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    '' Avoid divide 0
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For Each Pin In MeasureValue.Pins
                If MeasureValue.Pins(Pin).Value(site) = 0 Then
                    MeasureValue.Pins(Pin).Value(site) = 1
                End If
            Next Pin
        Next site
    End If
    
    For Each Pin In MeasureValue.Pins
        For Each site In TheExec.sites
            If MeasureValue.Pins(Pin).Value(site) = 0 Then
                MeasureValue.Pins(Pin).Value(site) = 0.000000000001
            End If
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & ", Pin : " & Pin & ", Measure Current = " & MeasureValue.Pins(Pin).Value(site))
        Next site
    Next Pin
 
    '' Print force condition
    Call Print_Force_Condition("r", MeasureValue)
    
    Dim PowerVal As Double
    '' Impedence measurement
    If b_PD_Mode Then
        Imped = MeasureValue.Math.Invert.Multiply(ForceVoltVal).Abs
    Else
        PowerVal = TheHdw.DCVS.Pins("VDDQL_DDR0").Voltage.Value
        Imped = MeasureValue.Math.Invert.Multiply(PowerVal - ForceVoltVal).Abs
    End If
    
    If TheExec.TesterMode = testModeOffline Then
        Call SimulateOutputImped(MeasurePin, Imped)
    Else
        If RAK_Flag = R_PathWithContact Then
            '' Compensate resistance after Kelvin for path resistance considerations
            For Each Pin In Imped.Pins
                For Each site In TheExec.sites
                    Imped.Pins.Item(Pin).Value(site) = Imped.Pins.Item(Pin).Value(site) - R_Path_PLD.Pins.Item(Pin).Value(site)
                Next site
            Next Pin
        End If
    End If
    
    TheHdw.PPMU.Pins(MeasurePin).Disconnect
    TheHdw.Digital.Pins(MeasurePin).Connect
    
    RTN_Imped_Val = Imped
    
End Function

Public Function SubMeasR(CPUA_Flag_In_Pat As Boolean, Pin As String, ForceVolt As String, ByRef RTN_Imped_Val As PinListData, Optional b_IsDifferential As Boolean, _
                                              Optional b_PD_Mode As Boolean = True)

    If (CPUA_Flag_In_Pat) Then
        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    Else
        Call TheHdw.Digital.Patgen.HaltWait
    End If
    
    Dim Diff_P_Pin As String, Diff_N_Pin As String
    Dim P_Pin_ForceV As Double, N_Pin_ForceV As Double
    Dim SplitForceVolt() As String
    
    If b_IsDifferential = False Then
        Call UP1600_PPMU_Measure_R_SE(Pin, ForceVolt, 50 * mA, R_PathWithContact, RTN_Imped_Val, b_PD_Mode)
    Else
        Diff_P_Pin = UCase(Pin)
        Diff_N_Pin = Replace(UCase(Diff_P_Pin), "_P", "_N")
        SplitForceVolt = Split(ForceVolt, ",")
        P_Pin_ForceV = CDbl(SplitForceVolt(0))
        N_Pin_ForceV = CDbl(SplitForceVolt(1))
        
        Call UP1600_PPMU_Measure_R_DI(Diff_P_Pin, Diff_N_Pin, P_Pin_ForceV, N_Pin_ForceV, 50 * mA, R_PathWithContact, RTN_Imped_Val)
    End If
    
    If (CPUA_Flag_In_Pat) Then
        Call TheHdw.Digital.Patgen.Continue(0, cpuA)
    Else
        TheHdw.Digital.Patgen.HaltWait
    End If
    
    TheHdw.Digital.Patgen.HaltWait

End Function

Public Function SimulateOutputImped(MeasureR_Pin As String, ByRef MeasureImped As PinListData) As Long
    Dim site As Variant
    For Each site In TheExec.sites.Active
        If site = 0 Then
            MeasureImped.Pins(MeasureR_Pin).Value(site) = 46
        ElseIf site = 1 Then
            MeasureImped.Pins(MeasureR_Pin).Value(site) = 53
        ElseIf site = 2 Then
''            MeasureImped.Pins(MeasureR_Pin).Value(Site) = 49
        ElseIf site = 3 Then
''            MeasureImped.Pins(MeasureR_Pin).Value(Site) = 52
        End If
    Next site
End Function

Public Function EvaluateForDCSpec(InputStrVal As String) As String

    InputStrVal = Trim(InputStrVal)
    Dim Temp_InputStrVal As String
    If Left(InputStrVal, 1) = "_" Then
        Temp_InputStrVal = Right(InputStrVal, Len(InputStrVal) - 1)
    End If
    
    EvaluateForDCSpec = CStr(TheExec.specs.DC.Item(Temp_InputStrVal).ContextValue)
End Function

Public Function ProcessEvaluateDCSpec(Pin_info As String) As String

    Dim temp_pin_info As String
    Dim temp_pin_name, calc_info As String
    Dim temp_pininfo_arr() As String
    Dim i As Integer
    Dim Update_calc_info As String
    Dim site As Variant
    temp_pin_info = Pin_info
    Update_calc_info = Pin_info
    temp_pin_info = Replace(temp_pin_info, "(", "")
    temp_pin_info = Replace(temp_pin_info, ")", "")
    temp_pin_info = Replace(temp_pin_info, "+", "~")
    temp_pin_info = Replace(temp_pin_info, "-", "~")
    temp_pin_info = Replace(temp_pin_info, "*", "~")
    temp_pin_info = Replace(temp_pin_info, "/", "~")
                   
    temp_pininfo_arr = Split(temp_pin_info, "~")
    
    For i = 0 To UBound(temp_pininfo_arr)
        If InStr(Left(temp_pininfo_arr(i), 1), "_") <> 0 Then
            temp_pin_name = temp_pininfo_arr(i)
            For Each site In TheExec.sites.Active
                temp_pininfo_arr(i) = CStr(TheExec.specs.DC.Item(Mid(temp_pininfo_arr(i), 2)).CurrentValue(site))
                Exit For
            Next site
            calc_info = Replace(Update_calc_info, temp_pin_name, temp_pininfo_arr(i))
            Update_calc_info = calc_info
        Else
            temp_pin_name = temp_pininfo_arr(i)
            'If IsNumeric(temp_pin_name) = False And temp_pin_name <> "" Then
            If InStr(temp_pin_name, DC_Spec_Var) <> 0 Then
                For Each site In TheExec.sites.Active
                    temp_pininfo_arr(i) = CStr(TheExec.specs.DC.Item(temp_pininfo_arr(i)).CurrentValue(site)) ''Carter, 20190506
                    Exit For
                Next site
            ElseIf InStr(temp_pin_name, "_") <> 0 Then
                For Each site In TheExec.sites.Active
                    temp_pininfo_arr(i) = CStr(TheExec.specs.DC.Item(temp_pininfo_arr(i) & DC_Spec_Var).CurrentValue(site))
                    Exit For
                Next site
            End If
            calc_info = Replace(Update_calc_info, temp_pin_name, temp_pininfo_arr(i))
            Update_calc_info = calc_info
        End If
    Next i
               
    ProcessEvaluateDCSpec = CStr(Evaluate(Update_calc_info))
End Function

Public Function HIP_Evaluate_ForceVal(ByRef ForceVSequenceArray() As String) As Long
    Dim EvalIndex As Integer
    Dim SplitByComma() As String
    Dim TempEvalStr As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim SplitByColon() As String
    Dim TempEvalFinalStr As String
    
    Dim splitbyand() As String
    Dim TempEvalAndStr As String
    
    For EvalIndex = 0 To UBound(ForceVSequenceArray)   ' Evaluate can use for equation calculation
        If InStr(ForceVSequenceArray(EvalIndex), "_") > 0 Or InStr(ForceVSequenceArray(EvalIndex), "+") > 0 Or InStr(ForceVSequenceArray(EvalIndex), "-") > 0 Or InStr(ForceVSequenceArray(EvalIndex), "*") > 0 Or InStr(ForceVSequenceArray(EvalIndex), "/") > 0 Then ' can not evaluate if only with  single number
            If InStr(ForceVSequenceArray(EvalIndex), ",") Then
                SplitByComma = Split(ForceVSequenceArray(EvalIndex), ",")
                For i = 0 To UBound(SplitByComma)
                    SplitByColon = Split(SplitByComma(i), ":")
                    For j = 0 To UBound(SplitByColon)
                        If InStr(SplitByColon(j), "_") Then
                            If j = 0 Then
                                If InStr(SplitByColon(j), "&") Then
                                    splitbyand = Split(SplitByColon(j), "&")
                                    For k = 0 To UBound(splitbyand)
                                        If k = 0 Then
                                            SplitByColon(j) = ProcessEvaluateDCSpec(splitbyand(k))
                                        Else
                                            SplitByColon(j) = SplitByColon(j) & "&" & ProcessEvaluateDCSpec(splitbyand(k))
                                        End If
                                    Next k
                                    TempEvalStr = SplitByColon(j)
                                Else
                                    TempEvalStr = ProcessEvaluateDCSpec(SplitByColon(j))
                                End If
                            Else
                                TempEvalStr = TempEvalStr & ":" & ProcessEvaluateDCSpec(SplitByColon(j))
                            End If
                        Else
                            If j = 0 Then
                                TempEvalStr = SplitByColon(j)
                            Else
                                TempEvalStr = TempEvalStr & ":" & SplitByColon(j)
                            End If
                        End If
                    Next j
                    If i = 0 Then
                        TempEvalFinalStr = TempEvalStr
                    Else
                        TempEvalFinalStr = TempEvalFinalStr & "," & TempEvalStr
                    End If
                Next i
                ForceVSequenceArray(EvalIndex) = TempEvalFinalStr
            Else
                If InStr(ForceVSequenceArray(EvalIndex), ":") Then
                    'SplitByComma = Split(ForceVSequenceArray(EvalIndex), ",")
                    SplitByColon = Split(ForceVSequenceArray(EvalIndex), ":")
                    For j = 0 To UBound(SplitByColon)
                        If InStr(SplitByColon(j), "_") Then
                            If j = 0 Then
                                TempEvalStr = ProcessEvaluateDCSpec(SplitByColon(j))
                            Else
                                TempEvalStr = TempEvalStr & ":" & ProcessEvaluateDCSpec(SplitByColon(j))
                            End If
                        Else
                            If j = 0 Then
                                TempEvalStr = SplitByColon(j)
                            Else
                                TempEvalStr = TempEvalStr & ":" & SplitByColon(j)
                            End If
                        End If
                    Next j
                    ForceVSequenceArray(EvalIndex) = TempEvalStr
                Else
                    ForceVSequenceArray(EvalIndex) = ProcessEvaluateDCSpec(ForceVSequenceArray(EvalIndex))
                End If

            End If
        End If
    Next EvalIndex
End Function

Public Function HEX_to_BIN(ByVal Hex As String) As String
    Dim i As Long
    Dim B As String
    
    Hex = UCase(Hex)
    For i = 1 To Len(Hex)
        Select Case Mid(Hex, i, 1)
            Case "0": B = B & "0000"
            Case "1": B = B & "0001"
            Case "2": B = B & "0010"
            Case "3": B = B & "0011"
            Case "4": B = B & "0100"
            Case "5": B = B & "0101"
            Case "6": B = B & "0110"
            Case "7": B = B & "0111"
            Case "8": B = B & "1000"
            Case "9": B = B & "1001"
            Case "A": B = B & "1010"
            Case "B": B = B & "1011"
            Case "C": B = B & "1100"
            Case "D": B = B & "1101"
            Case "E": B = B & "1110"
            Case "F": B = B & "1111"
        End Select
    Next i
'    While Left(b, 1) = "0"                 ''ZB correct  for  Cyprus AMP dqpi, capi total binary bits is fixed - 20170905
'        b = Right(b, Len(b) - 1)
'    Wend
    HEX_to_BIN = B
End Function
Public Function UP1600_PPMU_Measure_R_DI(Measure_P_Pin As String, Measure_N_Pin As String, P_ForceVolt As Double, N_ForceVolt As Double, _
MeasureCurrRange As Double, Optional RAK_Flag As Enum_RAK = 0, Optional ByRef RTN_Imped_Val As PinListData) As Long

    Dim MeasureValue As New PinListData
    Dim Imped As New PinListData
    Dim Pin  As Variant
    Dim site As Variant
    
    TheHdw.Digital.Pins(Measure_P_Pin & "," & Measure_N_Pin).Disconnect
    TheHdw.Wait 10 * us
    
    '' Initial force I to 0 and force V by your specified
    With TheHdw.PPMU.Pins(Measure_P_Pin)
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
        .ForceV P_ForceVolt, MeasureCurrRange
    End With
    
    With TheHdw.PPMU.Pins(Measure_N_Pin)
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
        .ForceV N_ForceVolt, MeasureCurrRange
    End With

    TheHdw.Wait 1 * ms
    
    DebugPrintFunc_PPMU CStr(Measure_P_Pin)
    
    MeasureValue = TheHdw.PPMU.Pins(Measure_P_Pin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    '' Avoid divide 0
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For Each Pin In MeasureValue.Pins
                If MeasureValue.Pins(Pin).Value(site) = 0 Then
                    MeasureValue.Pins(Pin).Value(site) = 1
                End If
            Next Pin
        Next site
    End If
    
    For Each Pin In MeasureValue.Pins
        For Each site In TheExec.sites
            If MeasureValue.Pins(Pin).Value(site) = 0 Then
                MeasureValue.Pins(Pin).Value(site) = 0.000000000001
            End If
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & ", Pin : " & Pin & ", Measure Current = " & MeasureValue.Pins(Pin).Value(site))
        Next site
    Next Pin
 
    '' Print force condition
    Call Print_Force_Condition("r", MeasureValue)
 
    '' Impedence measurement
    Imped = MeasureValue.Math.Invert.Multiply(P_ForceVolt - N_ForceVolt)
    
    Dim RAK_Pin_N As String
    
    If TheExec.TesterMode = testModeOffline Then
        Call SimulateOutputImped(Measure_P_Pin, Imped)
    Else
        If RAK_Flag = R_PathWithContact Then
            '' Compensate resistance after Kelvin for path resistance considerations
            For Each Pin In Imped.Pins
                RAK_Pin_N = Replace(UCase(Pin), "_P", "_N")
                For Each site In TheExec.sites
                    Imped.Pins.Item(Pin).Value(site) = Imped.Pins.Item(Pin).Value(site) - R_Path_PLD.Pins.Item(Pin).Value(site) - R_Path_PLD.Pins.Item(RAK_Pin_N).Value(site)
                Next site
            Next Pin
        End If
    End If
    
    TheHdw.PPMU.Pins(Measure_P_Pin & "," & Measure_N_Pin).Disconnect
    TheHdw.Digital.Pins(Measure_P_Pin & "," & Measure_N_Pin).Connect
    
    RTN_Imped_Val = Imped
    
End Function


Public Function CreateSimulateMDLL_Data(argc As Integer, argv() As String) As Long
    Dim i As Long, j As Long
    Dim site As Variant
    
    Dim DSPWaveLength As Long
    If TheExec.TesterMode = testModeOffline Then
    
        DSPWaveLength = 16
        Dim SimulateDSPWaveBin() As New DSPWave
        ReDim SimulateDSPWaveBin(argc - 1) As New DSPWave
        
        
        For i = 1 To argc - 1
            SimulateDSPWaveBin(i).CreateConstant 0, DSPWaveLength, DspLong
            For Each site In TheExec.sites
                For j = 0 To DSPWaveLength - 1
                    If site = 0 Then
                        If i < 4 Then
                            SimulateDSPWaveBin(i)(site).Element(j) = 0
                        Else
                            If j < 2 Then
                                SimulateDSPWaveBin(i)(site).Element(j) = 0
                            Else
                                SimulateDSPWaveBin(i)(site).Element(j) = 1
                            End If
                        End If
                    Else
                        If i < 3 Then
                            SimulateDSPWaveBin(i)(site).Element(j) = 1
                        ElseIf i >= 3 And i <= 5 Then
                            If j < 2 Then
                                SimulateDSPWaveBin(i)(site).Element(j) = 0
                            Else
                                SimulateDSPWaveBin(i)(site).Element(j) = 1
                            End If
                        Else
                            SimulateDSPWaveBin(i)(site).Element(j) = 0
                        End If
                        
                    End If
                Next j
                
            Next site
            Call AddStoredCaptureData(argv(i), SimulateDSPWaveBin(i))
        Next i
    End If

    Dim Displaystring As String

     For Each site In TheExec.sites
            For i = 1 To argc - 1
            Displaystring = ""
            For j = 0 To DSPWaveLength - 1
                If j = 0 Then
                    Displaystring = SimulateDSPWaveBin(i)(site).Element(j)
                Else
                    Displaystring = Displaystring & SimulateDSPWaveBin(i)(site).Element(j)
                End If
            Next j
           If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Simulate data Site_" & site & " Dict = " & argv(i) & " Content = " & Displaystring)
        Next i
     Next site

End Function

Public Function IPF_Connect_PPMU_ForceV(argc As Long, argv() As String)

    Dim ForceValStr(0) As String
    ForceValStr(0) = argv(1)
    Call HIP_Evaluate_ForceVal(ForceValStr())
    
    argv(0) = Replace(argv(0), "+", ",")
    TheHdw.Digital.Pins(argv(0)).Disconnect
    With TheHdw.PPMU.Pins(argv(0))
        .ForceV CDbl(ForceValStr(0)), 0.05
        .Connect
        .Gate = tlOn
    End With
                                                                                                                                                                                                                                                               
   If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Pin = " & argv(0) & ", Force V = " & argv(1) & ", Measure Current Range = " & TheHdw.PPMU.Pins(argv(0)).MeasureCurrentRange
End Function

Public Function IPF_Connect_Digital(argc As Long, argv() As String)
    argv(0) = Replace(argv(0), "+", ",")
    
    With TheHdw.PPMU.Pins(argv(0))
        .ForceV 0, 0
        .Disconnect
        .Gate = tlOff
    End With
    TheHdw.Digital.Pins(argv(0)).Connect
End Function

Public Function PPMU_SerialMeasureCurr(ForceByPin() As String, ForceValByPin() As String, Measure_I_Range() As String, ByRef MeasureValue As PinListData, Optional b_ForceDiffVolt As Boolean) As Long
    Dim InputPin As Variant
    Dim i As Long, j As Long, p As Long
    
    Dim PinAry() As String, NumberPins As Long
    Dim MeasPin As String
    
    Dim InputPin_Index As Long
    InputPin_Index = 0
    
    Dim TempMeasVal() As New PinListData
    Dim Save_force_data As New DSPWave
    
    
    ReDim TempMeasVal(UBound(ForceByPin)) As New PinListData
    Dim MergeValPinListData As New PinListData
    
    
    
    For Each InputPin In ForceByPin
        Call TheExec.DataManager.DecomposePinList(CStr(InputPin), PinAry(), NumberPins)
        
        For i = 0 To NumberPins - 1
        
        
        If i = 0 Then
        
         Save_force_data.CreateConstant 0, NumberPins
         
         End If
         
     
            MeasPin = PinAry(i)
            TheHdw.Digital.Pins(MeasPin).Disconnect
            With TheHdw.PPMU.Pins(MeasPin)
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                .Connect
                .Gate = tlOn
                If b_ForceDiffVolt = False Then
                    .ForceV ForceValByPin(0), Measure_I_Range(InputPin_Index)
                Else
                    .ForceV ForceValByPin(InputPin_Index), Measure_I_Range(InputPin_Index)
                End If
            End With
            
'            Call TheExec.Datalog.WriteComment("Pin = " & MeasPin & " Measure Current Range = " & TheHdw.PPMU.Pins(MeasPin).MeasureCurrentRange)
            If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment(TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasPin & " =" & TheHdw.PPMU.Pins(MeasPin).MeasureCurrentRange)
            TheHdw.Wait (100 * us)
            DebugPrintFunc_PPMU CStr(MeasPin)
            
            TempMeasVal(InputPin_Index).AddPin (MeasPin)
            TempMeasVal(InputPin_Index).Pins(MeasPin) = TheHdw.PPMU.Pins(MeasPin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            
            ' Add for print force CSHO 20180620
'            Call Print_Force_Condition("i", MeasureValue)


            Save_force_data.Element(i) = TheHdw.PPMU(MeasPin).Voltage.Value
            
            With TheHdw.PPMU.Pins(MeasPin)
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                .Disconnect
                .Gate = tlOff
            End With
            TheHdw.Digital.Pins(MeasPin).Connect
        Next i
        InputPin_Index = InputPin_Index + 1
     Next InputPin
    
    For i = 0 To UBound(TempMeasVal)
        For p = 0 To TempMeasVal(i).Pins.Count - 1
            MergeValPinListData.AddPin (TempMeasVal(i).Pins(p))
            MergeValPinListData.Pins(TempMeasVal(i).Pins(p)) = TempMeasVal(i).Pins(p)
        Next p
    Next i
    MeasureValue = MergeValPinListData
    
    '//////////////////////////////// for meas I print////////////////////////////////////////  csho
    
    Call Print_Force_Condition_I("I", Save_force_data, MeasureValue)
        
End Function

Public Function DictDSPToSiteLong(DictName As String, ByRef RTN_sl_Val As SiteLong, DictTrimFuseName As String) As Long
    Dim DSP_Val_Dec As New DSPWave
    Dim site As Variant
    DSP_Val_Dec = GetStoredCaptureData(DictName)
    
    Call AddStoredCaptureData(DictTrimFuseName, DSP_Val_Dec)
    
    For Each site In TheExec.sites
        RTN_sl_Val(site) = DSP_Val_Dec(site).Element(0)
    Next site
    
End Function

Public Function HardIP_Duty_Frequency(FreqMeasPins As PinList, IsDifferentialPin As Boolean, TestSeqNum As Integer, d_FreqMeasInterval As Double, _
        Optional Rtn_MeasFreq As PinListData, Optional b_TestLimitPerPin As Boolean = False, Optional b_SkipTestLimit As Boolean = True)
    
    Dim site As Variant
    Dim p As Long
    Dim MeasFreq As New PinListData
    
    Call Freq_MeasFreqSetup(FreqMeasPins, d_FreqMeasInterval)  '' 20150621 - default d_FreqMeasInterval = 0.001
    '' 20150623 - Add Customize Wait Time
    Call HardIP_Freq_MeasFreqStart(FreqMeasPins, d_FreqMeasInterval, MeasFreq)       '' 20150621 - default d_FreqMeasInterval = 0.001

    Dim TestNameInput As String
    

    TestNameInput = "Freq_Meas_" & CStr(TestSeqNum)


    If Not b_SkipTestLimit Then
        If IsDifferentialPin = True Then
            For p = 0 To MeasFreq.Pins.Count - 1 Step 2 ' freq counter result of differential pins is stored in positive pin
                TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p + 1), Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceNone
            Next p
        Else
            If b_TestLimitPerPin = True Then
                For p = 0 To MeasFreq.Pins.Count - 1
                    TheExec.Flow.TestLimit resultVal:=MeasFreq.Pins(p), Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceNone
                Next p
            Else
                TheExec.Flow.TestLimit resultVal:=MeasFreq, Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceNone
            End If
        End If
    End If

        
    '' 20151224 - Merge print measured frequency during shmoo if need
    G_MeasFreqForCZ = MeasFreq
    
    ''20160906 - Return MeasFreq to main program
    Rtn_MeasFreq = MeasFreq
    
End Function

Public Function Dec2BinStr32Bit_Rev(ByVal Nbit As Long, ByVal num As Long) As String
    ' 2'complement: invert the number's bits and then add 1
    'Dec2BinStr32Bit 32, -65525
    '1111111111111110000000000001011    -65525
    '0000000000000001111111111110101     65525
    Dim i As Integer, j As Integer
    Dim Element_Amount As Integer
    Dim Count As Integer
    Dim BinStr As String
    ' MSB "010101" LSB
    
    BinStr = ""
    If Nbit < 1 Then MsgBox ("Warning(Dec2BinStr32Bit)!!! Decimal Number or number of Bit is wrong")
    If Nbit = 32 Then
        Nbit = 30
        If num < 0 Then
            BinStr = "1"
        Else
            BinStr = "0"
        End If
    End If
    For i = Nbit To 0 Step -1
        If num And (2 ^ i) Then
            BinStr = BinStr & "1"
        Else
            BinStr = BinStr & "0"
        End If
    Next
    Dim FinalStr As String
    Dim ExactStr As String
    For i = 1 To Len(BinStr)
        ExactStr = Mid(BinStr, i, 1)
        If i = 1 Then
            FinalStr = ExactStr
        Else
            FinalStr = ExactStr & FinalStr
        End If
    Next i
    
    Dec2BinStr32Bit_Rev = FinalStr
'    Debug.Print BinStr
End Function


Public Function DisplayForLoopFuncResult_EndOfTest(CUS_Str_DigSrcData As String, Rtn_SweepTestName As String, CPUA_Flag_In_Pat As Boolean, DigSrc_FlowForLoopIntegerName As String)
        ''20170405-Record all functional test result from flow for loop opcode, use global string to store them
        Dim sb_FuncTestResult As New SiteBoolean
        Dim StrTestResult As String
        Dim StrTestResultPerTestInstance As String
        Dim StrCodeDisplay As String
        Dim SplitUnderLine() As String
        Dim SplitFlowForLoopByColon() As String
        Dim TestResultSiteIndex As Long
        Dim MaxForNum As Long
        Dim site As Variant
        TestResultSiteIndex = 0
        If CUS_Str_DigSrcData <> "" And UCase(CUS_Str_DigSrcData) = UCase("BinToGray") Then
            If CPUA_Flag_In_Pat = False Then
                If Left(Rtn_SweepTestName, 1) = "0" Then
                    gs_RecordGrayCodeTestResult = ""
                End If
                sb_FuncTestResult = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
                SplitUnderLine = Split(Rtn_SweepTestName, "_")
                StrCodeDisplay = "Dec = " & SplitUnderLine(0) & " , Gray Code = " & SplitUnderLine(UBound(SplitUnderLine))
                
                StrTestResultPerTestInstance = ""
                
                For Each site In TheExec.sites
                    If sb_FuncTestResult(site) = True Then
                        If TestResultSiteIndex = 0 Then
                            StrTestResultPerTestInstance = StrCodeDisplay & " Site " & site & " Test Pass" & vbCrLf
                        Else
                            StrTestResultPerTestInstance = StrTestResultPerTestInstance & StrCodeDisplay & " Site " & site & " Test Pass" & vbCrLf
                        End If
                    Else
                        If TestResultSiteIndex = 0 Then
                            StrTestResultPerTestInstance = StrCodeDisplay & " Site " & site & " Test Fail" & vbCrLf
                        Else
                            StrTestResultPerTestInstance = StrTestResultPerTestInstance & StrCodeDisplay & " Site " & site & " Test Fail" & vbCrLf
                        End If
                    End If
                    TestResultSiteIndex = TestResultSiteIndex + 1
                Next site
                
                If CDbl(SplitUnderLine(0)) = 0 Then
                    gs_RecordGrayCodeTestResult = StrTestResultPerTestInstance
                Else
                    gs_RecordGrayCodeTestResult = gs_RecordGrayCodeTestResult & StrTestResultPerTestInstance
                End If
                SplitFlowForLoopByColon = Split(DigSrc_FlowForLoopIntegerName, ":")
                MaxForNum = 2 ^ CDbl(SplitFlowForLoopByColon(UBound(SplitFlowForLoopByColon))) - 1
                If CDbl(SplitUnderLine(0)) = MaxForNum And DebugPrintEnable = True Then
                    TheExec.Datalog.WriteComment (gs_RecordGrayCodeTestResult)
                End If
            End If
        End If
End Function

Public Function PrintDigCapSetting(DigCap_Pin As PinList, DigCap_Sample_Size As Long, CUS_Str_DigCapData As String)
    If DigCap_Pin <> "" And DigCap_Sample_Size <> 0 And gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment (CUS_Str_DigCapData)
        TheExec.Datalog.WriteComment ("Cap Bits = " & DigCap_Sample_Size)
        TheExec.Datalog.WriteComment ("Cap Pin = " & DigCap_Pin)
        TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test End   ========")
    End If
End Function

Public Function Select_MeasIRange(src_string As String, Job As String) As String

    Dim src_str_array() As String

    If Job = "" Then Job = "CP1"
    
    If Not (((UCase(src_string) Like "*CP*") Or (UCase(src_string) Like "*FT*"))) Then Select_MeasIRange = src_string
    
    If Not (((UCase(Job) Like "*CP*") Or (UCase(Job) Like "*FT*"))) Then Job = "CP1"
    
    
    src_str_array = Split(src_string, ";")
    
    Dim var As Variant
    
    For Each var In src_str_array
        If (UCase(var) Like "*" & UCase(Job) & "*") Then
            
            Select_MeasIRange = Split(var, "=")(1)
            Exit For
            
        End If
    Next var

End Function

Public Function HardIP_Bin2Dec(ByRef DataOut_85C As DSPWave, Optional DSPWave_Dict As DSPWave) As Long
Dim site As Variant
Dim i As Integer
Dim Data_Temp As String
    For Each site In TheExec.sites
        For i = 0 To (DSPWave_Dict(site).SampleSize - 1)
            Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(i))
        Next i
            DataOut_85C(site).Element(0) = Bin2Dec_rev(Data_Temp)
            Data_Temp = ""
    Next site

End Function


Public Function HardIP_Dec2Bin(ByRef Read_Code As DSPWave, Optional DSPWave_Dict As DSPWave, Optional dspwavesize) As Long
Dim TempVal As Long
Read_Code.CreateConstant 0, dspwavesize

Dim i As Integer
Dim Data_Temp As String
Dim site As Variant

    For Each site In TheExec.sites
        TempVal = DSPWave_Dict(site).Element(0)
        For i = 0 To dspwavesize - 1
            Read_Code.Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i
    Next site
End Function

Public Function TMPS_Coeff_Calculation(ByRef Coeff_A0 As DSPWave, ByRef Coeff_A1 As DSPWave, ByRef Coeff_A2 As DSPWave, ByRef Coeff_A3 As DSPWave, ByRef Coeff_A4 As DSPWave, Optional DataOut_85C As DSPWave, Optional DataOut_25C As DSPWave) As Long

    Dim site As Variant
    Dim a3 As Double
    Dim a2 As Double
    Dim a1 As Double
    Dim A0 As Double
    Dim b1 As Double
    Dim b0 As Double
    Dim Sensitivity_25C As Double
    Dim Sensitivity_85C As Double
    Dim Error_25C As Double
    Dim Error_85C As Double
    Dim DataOut_Ideal_85C As Integer
    Dim DataOut_Ideal_25C As Integer
    Dim M As Double
    Dim k As Double
    Dim AA0 As New SiteDouble
    Dim AA1 As New SiteDouble
    Dim AA2 As New SiteDouble
    Dim AA3 As New SiteDouble
    Dim AA4 As New SiteDouble
    Dim TestNameInput As String
    Dim OutputTname_format() As String
   
        a3 = -2.3122
        a2 = 2.1532
        a1 = -59.793
        A0 = 216.28
        b1 = -3.5425
        b0 = 19.744
        
        'Sensitivity_25C = 10.1515
        'Sensitivity_85C = 12.5541
        
        Sensitivity_25C = 10.389 ' for Turks
        Sensitivity_85C = 12.4897 ' for Turks
        
        DataOut_Ideal_85C = 2023
        DataOut_Ideal_25C = 2700

       For Each site In TheExec.sites
            Error_85C = (DataOut_85C(site).Element(0) - DataOut_Ideal_85C) / Sensitivity_85C
            Error_25C = (DataOut_25C(site).Element(0) - DataOut_Ideal_25C) / Sensitivity_25C
            M = (Error_85C - Error_25C) / (85 - 25)
            k = Error_25C - M * 25
        

            AA0(site) = M * A0 * b0 + k * b0
            AA1(site) = M * A0 * b1 + M * a1 * b0 + k * b1
            AA2(site) = M * a1 * b1 + M * a2 * b0
            AA3(site) = M * a2 * b1 + M * a3 * b0
            AA4(site) = M * a3 * b1
        Next site
            
            Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A0, AA0, 11, 4)
    Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A1, AA1, 10, 4)
    Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A2, AA2, 7, 5)
    Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A3, AA3, 5, 5)
    Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A4, AA4, 3, 8)

    'TestNameInput = Report_TName_From_Instance("C", "X", "", 0)
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    'TheExec.Flow.TestLimit resultVal:=AA0, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=AA0, ForceResults:=tlForceFlow, Tname:=TestNameInput


    'TestNameInput = Report_TName_From_Instance("C", "X", "", 0)
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    'TheExec.Flow.TestLimit resultVal:=AA1, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=AA1, ForceResults:=tlForceFlow, Tname:=TestNameInput

    'TestNameInput = Report_TName_From_Instance("C", "X", "", 0)
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    'TheExec.Flow.TestLimit resultVal:=AA2, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=AA2, ForceResults:=tlForceFlow, Tname:=TestNameInput

    'TestNameInput = Report_TName_From_Instance("C", "X", "", 0)
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    'TheExec.Flow.TestLimit resultVal:=AA3, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=AA3, ForceResults:=tlForceFlow, Tname:=TestNameInput

    'TestNameInput = Report_TName_From_Instance("C", "X", "", 0)
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    'TheExec.Flow.TestLimit resultVal:=AA4, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=AA4, ForceResults:=tlForceFlow, Tname:=TestNameInput

End Function


Public Function TMPS_Coeff_Calculation_1point(ByRef Coeff_A0 As DSPWave, ByRef Coeff_A1 As DSPWave, ByRef Coeff_A2 As DSPWave, ByRef Coeff_A3 As DSPWave, ByRef Coeff_A4 As DSPWave, Optional DataOut_25C As DSPWave) As Long

    Dim site As Variant
    Dim a3 As Double
    Dim a2 As Double
    Dim a1 As Double
    Dim A0 As Double
    Dim b1 As Double
    Dim b0 As Double
    Dim Sensitivity_25C As Double
'    Dim Sensitivity_85C As Double
    Dim Error_25C As Double
'    Dim Error_85C As Double
'    Dim DataOut_Ideal_85C As Integer
    Dim DataOut_Ideal_25C As Integer
    Dim M As Double
    Dim k As Double
    Dim AA0 As New SiteDouble
    Dim AA1 As New SiteDouble
    Dim AA2 As New SiteDouble
    Dim AA3 As New SiteDouble
    Dim AA4 As New SiteDouble
    Dim TestNameInput As String
    Dim OutputTname_format() As String
   
        a3 = -2.3122
        a2 = 2.1532
        a1 = -59.793
        A0 = 216.28
        b1 = -3.5425
        b0 = 19.744
        
        'Sensitivity_25C = 10.1515
        'Sensitivity_85C = 12.5541
        
        Sensitivity_25C = 10.389 ' for Turks
'        Sensitivity_85C = 12.4897 ' for Turks
        
'        DataOut_Ideal_85C = 2023
        DataOut_Ideal_25C = 2700

       For Each site In TheExec.sites
'            Error_85C = (DataOut_85C(Site).Element(0) - DataOut_Ideal_85C) / Sensitivity_85C
            Error_25C = (DataOut_25C(site).Element(0) - DataOut_Ideal_25C) / Sensitivity_25C
'            M = (Error_85C - Error_25C) / (85 - 25)
'            k = Error_25C - M * 25
            k = Error_25C

'            AA0(Site) = M * a0 * b0 + k * b0
'            AA1(Site) = M * a0 * b1 + M * a1 * b0 + k * b1
'            AA2(Site) = M * a1 * b1 + M * a2 * b0
'            AA3(Site) = M * a2 * b1 + M * a3 * b0
'            AA4(Site) = M * a3 * b1
            
            AA0(site) = k * b0
            AA1(site) = k * b1
            AA2(site) = 0
            AA3(site) = 0
            AA4(site) = 0
            
            
        Next site
            
            Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A0, AA0, 11, 4)
            Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A1, AA1, 10, 4)
            Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A2, AA2, 7, 5)
            Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A3, AA3, 5, 5)
            Call TMPS_2s_Complement_Fractional_Conversion(Coeff_A4, AA4, 3, 8)


    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    TheExec.Flow.TestLimit resultVal:=AA0, ForceResults:=tlForceFlow, Tname:=TestNameInput


    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    TheExec.Flow.TestLimit resultVal:=AA1, ForceResults:=tlForceFlow, Tname:=TestNameInput

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    TheExec.Flow.TestLimit resultVal:=AA2, ForceResults:=tlForceFlow, Tname:=TestNameInput

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    TheExec.Flow.TestLimit resultVal:=AA3, ForceResults:=tlForceFlow, Tname:=TestNameInput

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "", 0)
    TheExec.Flow.TestLimit resultVal:=AA4, ForceResults:=tlForceFlow, Tname:=TestNameInput

End Function

Public Function TMPS_2s_Complement_Fractional_Conversion(ByRef Coeff_Dict As DSPWave, Coeff As SiteDouble, Integer_Bit As Integer, Fractional_Bit As Integer) As Long
Dim site As Variant
Dim High_limit As Double: High_limit = Bin2Dec_rev(String(Integer_Bit - 1, "1")) + Bin2Dec_rev_Fractional(String(Fractional_Bit, "1"))
Dim Low_limit As Double: Low_limit = -2 ^ (Integer_Bit - 1)

    For Each site In TheExec.sites
        If Coeff(site) <= Low_limit Then
            Coeff_Dict(site).Element(0) = 2 ^ (Integer_Bit + Fractional_Bit) + FormatNumber(2 ^ Fractional_Bit * Low_limit)
        ElseIf Coeff(site) > Low_limit And Coeff(site) < 0 Then
            Coeff_Dict(site).Element(0) = 2 ^ (Integer_Bit + Fractional_Bit) + FormatNumber(2 ^ Fractional_Bit * Coeff(site))
        ElseIf Coeff(site) < High_limit And Coeff(site) >= 0 Then
            Coeff_Dict(site).Element(0) = FormatNumber(2 ^ Fractional_Bit * Coeff(site))
        Else
            Coeff_Dict(site).Element(0) = FormatNumber(2 ^ Fractional_Bit * High_limit)
        End If
    Next site
End Function

Public Function TMPS_Temperature2iEDA(Dict_Str As String, Temperature As SiteDouble)
Dim Split_Name() As String: Split_Name = Split(Dict_Str, "_")
Dim site As Variant
For Each site In TheExec.sites
    If UCase(Split_Name(2)) = "25C" Then
        If UCase(Split_Name(1)) = "SOC0" Then
            gS_TMPS1_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "SOC1" Then
            gS_TMPS2_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "GFX" Then
            gS_TMPS3_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "SOC3" Then
            gS_TMPS4_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "ANE" Then
            gS_TMPS5_Untrim(site) = CStr(FormatNumber(Temperature(site)))
       ElseIf UCase(Split_Name(1)) = "PCPU0" Then
            gS_TMPS6_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU1" Then
            gS_TMPS7_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU2" Then
            gS_TMPS8_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU3" Then
            gS_TMPS9_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU4" Then
            gS_TMPS10_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU5" Then
            gS_TMPS11_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU6" Then
            gS_TMPS12_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU7" Then
            gS_TMPS13_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "ECPU0" Then
            gS_TMPS14_Untrim(site) = CStr(FormatNumber(Temperature(site)))
        End If
    ElseIf InStr(UCase(Dict_Str), "85C") <> 0 Then
        If UCase(Split_Name(1)) = "SOC0" Then
            gS_TMPS1_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "SOC1" Then
            gS_TMPS2_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "GFX" Then
            gS_TMPS3_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "SOC3" Then
            gS_TMPS4_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "ANE" Then
            gS_TMPS5_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU0" Then
            gS_TMPS6_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU1" Then
            gS_TMPS7_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU2" Then
            gS_TMPS8_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU3" Then
            gS_TMPS9_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU4" Then
            gS_TMPS10_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU5" Then
            gS_TMPS11_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU6" Then
            gS_TMPS12_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "PCPU7" Then
            gS_TMPS13_Trim(site) = CStr(FormatNumber(Temperature(site)))
        ElseIf UCase(Split_Name(1)) = "ECPU0" Then
            gS_TMPS14_Trim(site) = CStr(FormatNumber(Temperature(site)))
        End If
    End If
Next site
End Function

Public Function TrimUVI80_Meas_VFI(Pat As String, TestSequenceArray() As String, srcPin As PinList, code As SiteLong, _
MeasV_Pin As String, MeasValue As SiteDouble, MeasI_Pin As PinList, MeasureI_Range As Double, _
MeasF_PinS_SingleEnd As PinList, MeasF_Interval As String, MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, _
DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, DigSrc_Equation As String, DigSrc_Assignment As String, _
DigCap_Pin As PinList, DigCap_DataWidth As Long, DigCap_Sample_Size As Long, CUS_Str_DigCapData As String, OutDSP As DSPWave, _
TrimCodeSize As Long, Trimname As String, Meas_StoreName As String, Cal_Eqn As String, TrimCal_Name As String, CPUA_Flag_In_Pat As Boolean, Optional Final_Calc As Boolean, Optional b_Trimfinish As Boolean = False, Optional MSB_First_Flag As Boolean = False)

    On Error GoTo err
    
    Dim sigName As String, srcWave As New DSPWave, site As Variant
    
    Dim DigSrcCodeSize As Long
    Dim i As Long, j As Long
    Dim code_bin() As String
    Dim Ts As Variant
    Dim Str_FinalPatName As String
    Dim temp_assignment As String
    Dim cal As New SiteDouble
    Dim out_str() As String
    Dim MeasStoreName_Ary() As String
    Dim TestSeqNum As Long
    ReDim code_bin(TheExec.sites.Existing.Count)
    ReDim out_str(TheExec.sites.Existing.Count)
    Dim TrimCal_value As New PinListData
    Dim TrimCalCap_value As New DSPWave
    Dim TrimCal_Name_array() As String
    Dim MeasV_Pin_split() As String
    TrimCal_Name_array = Split(TrimCal_Name, ":")
    MeasV_Pin_split = Split(MeasV_Pin, "+")
    ''''''''''''''''''''''''''''''''setup store name'''''''''''''''''''''''''''''.
    MeasStoreName_Ary = Split(Meas_StoreName, "+")
    ReDim Preserve MeasStoreName_Ary(UBound(TestSequenceArray))
    Dim Rtn_Meas As New PinListData
    Dim Store_Rtn_Meas() As New PinListData
    Dim SoreMaxNum As Long
    Dim StoreIndex As Long
    ''20170123-Get how many store name in MeasStoreName_Ary
    If Meas_StoreName <> "" Then
        SoreMaxNum = 0
        For i = 0 To UBound(MeasStoreName_Ary)
            If MeasStoreName_Ary(i) <> "" Then
                SoreMaxNum = SoreMaxNum + 1
            End If
        Next i
         ReDim Store_Rtn_Meas(SoreMaxNum - 1) As New PinListData
         StoreIndex = 0
    End If
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    temp_assignment = DigSrc_Assignment
    
    sigName = "DSSC_Search_Code"
    TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals.Add sigName
    srcWave.CreateConstant 0, DigSrc_Sample_Size, DspLong
    For Each site In TheExec.sites
        

        DigSrc_Assignment = temp_assignment
        code_bin(site) = ""
        
        For i = 0 To TrimCodeSize - 1
            If i = 0 Then
                code_bin(site) = CStr(code(site) And 1)
            Else
                code_bin(site) = code_bin(site) & CStr((code(site) And (2 ^ i)) \ (2 ^ i))
            End If
        Next i
        
        
        
        DigSrc_Assignment = Replace(DigSrc_Assignment, Trimname, code_bin(site))
        
        
        Call Create_DigSrc_Data_Trim(srcPin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, srcWave, site, MSB_First_Flag:=MSB_First_Flag)
        
        If DigSrc_DataWidth = 0 Then
            DigSrc_DataWidth = 4
        End If
        
        out_str(site) = ""
        
        For i = 0 To DigSrc_Sample_Size - 1
        
            If (i Mod DigSrc_DataWidth) = 0 Then
                out_str(site) = out_str(site) & " "
            End If
                out_str(site) = out_str(site) & srcWave.Element(i)

        Next i
        
        'theexec.Datalog.WriteComment "Site " & Site & ",Code " & code_bin(Site) & ",Src Code = " & out_str(Site)
'        For i = 0 To TrimRepeat - 1
'            For j = 0 To TrimCodeSize - 1
'                srcwave.Element(i * TrimCodeSize + j) = srcwave.Element(j)
'            Next j
'        Next i
        
        TheExec.WaveDefinitions.CreateWaveDefinition "WaveDef" & site, srcWave, True
        With TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals(sigName)
            .WaveDefinitionName = "WaveDef" & site
            .SampleSize = DigSrc_Sample_Size
            .Amplitude = 1
            .LoadSamples
            .LoadSettings
        End With

    Next site
    TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals.DefaultSignal = sigName
    
    If DigCap_Sample_Size <> 0 Then
    
        Call AnalyzePatName(Pat, Str_FinalPatName)
        
        '' 20150812-Modify program to process multiply dig cap pins
        With TheHdw.DSSC.Pins(DigCap_Pin).Pattern(Pat).Capture.Signals
            .Add (Str_FinalPatName & DigCap_Sample_Size & "_" & DigCap_Pin)
            With .Item(Str_FinalPatName & DigCap_Sample_Size & "_" & DigCap_Pin)
                .SampleSize = DigCap_Sample_Size    'CaptureCyc * OneCycle
                .LoadSettings
            End With
        End With
        
        'Create capture waveform
        OutDSP = TheHdw.DSSC.Pins(DigCap_Pin).Pattern(Pat).Capture.Signals(Str_FinalPatName & DigCap_Sample_Size & "_" & DigCap_Pin).DSPWave
        
        '' 20150813 - Assign WaveName to the DSPWave to do recognition of post process.
        For Each site In TheExec.sites
            OutDSP(site).Info.WaveName = DigCap_Pin
        Next site
        
        ''TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug ''20180827 -- TYCHENGG -- use defaut as automatic
        TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
    End If
    
  
    
        
    TheHdw.Patterns(Pat).start
    
    
    
    'Call DebugPrintFunc_PPMU("")
    Dim MeasValue_Temp As New SiteDouble
    Set MeasValue_Temp = MeasValue_Temp.Add(10000000000000#)
    
    Dim MeasV_Flag As Boolean: MeasV_Flag = False
    
    For Each Ts In TestSequenceArray
        If CPUA_Flag_In_Pat = True Then
        TheHdw.Digital.Patgen.FlagWait cpuA, 0
        'thehdw.Wait 10 * ms
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
        Select Case UCase(Ts)
            Case "V"
                MeasV_Pin = CheckAndReturnArrayData(MeasV_Pin_split, TestSeqNum)
                Call Trim_SetupandmeasureV_UVI80(MeasV_Pin, MeasValue, code, code_bin, out_str, b_Trimfinish)
                If Meas_StoreName <> "" Then
                    If MeasStoreName_Ary(TestSeqNum) <> "" Then
                        Store_Rtn_Meas(StoreIndex).AddPin (MeasV_Pin)
                        Store_Rtn_Meas(StoreIndex).Pins(MeasV_Pin) = MeasValue
                        Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                        StoreIndex = StoreIndex + 1
                    End If
                End If
                For Each site In TheExec.sites.Active
                    If (MeasValue_Temp > MeasValue) Then
                        MeasValue_Temp = MeasValue
                    End If
                Next site
                MeasV_Flag = True
            Case "I"
            
            
                Call Trim_SetupandmeasureI_UVI80(MeasI_Pin, MeasValue, MeasureI_Range, code, code_bin, out_str, b_Trimfinish)
                If Meas_StoreName <> "" Then
                    If MeasStoreName_Ary(TestSeqNum) <> "" Then
                        Rtn_Meas.AddPin (MeasI_Pin)
                        For Each site In TheExec.sites
                            Rtn_Meas.Pins(MeasI_Pin).Value(site) = MeasValue(site)
                        Next site
                        Store_Rtn_Meas(StoreIndex) = Rtn_Meas
                        Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                        StoreIndex = StoreIndex + 1
                    End If
                End If
            
               
                 
            Case "F"
            
                If MeasF_Interval = "" Then
                    MeasF_Interval = 0.001
                End If
            
                Call Trim_SetupandmeasureF_UVI80(MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, MeasValue, code, code_bin, out_str, b_Trimfinish)
                
                If Meas_StoreName <> "" Then
                    If MeasStoreName_Ary(TestSeqNum) <> "" Then
                        Rtn_Meas.AddPin (MeasF_PinS_SingleEnd)
                        For Each site In TheExec.sites
                            Rtn_Meas.Pins(MeasF_PinS_SingleEnd).Value(site) = MeasValue(site)
                        Next site
                        Store_Rtn_Meas(StoreIndex) = Rtn_Meas
                        Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                        StoreIndex = StoreIndex + 1
                    End If
                End If
                
                
                

            Case "C"
                Dim OutDSP2 As New DSPWave
                Dim OutDSP_Temp As New DSPWave
            
                If TheExec.TesterMode = testModeOffline Then
                    For Each site In TheExec.sites.Active
                        OutDSP.CreateRandom 0, 1, DigCap_Sample_Size, , DspLong
                    Next site
                End If
                
                
                'Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP)
                'Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
                Call Addstorecapture_Trim(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
                If TrimCal_Name <> "" And InStr(TrimCal_Name, "C:") = 0 Then
                    OutDSP_Temp = GetStoredCaptureData(TrimCal_Name)
                Else
                    OutDSP_Temp = OutDSP
                End If
                
                For Each site In TheExec.sites.Active
                    OutDSP_Temp = OutDSP_Temp.ConvertDataTypeTo(DspLong)
                    OutDSP2 = OutDSP_Temp.ConvertStreamTo(tldspParallel, OutDSP_Temp.SampleSize, 0, Bit0IsMsb) '' Convert BInary to Desimal
                    MeasValue(site) = OutDSP2.Element(0)
                   If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Code(Decimal) = " & MeasValue(site)
                Next site
            
            Case "N"
            
        End Select
        If CPUA_Flag_In_Pat = True Then
        TheHdw.Digital.Patgen.Continue 0, cpuA
        
        TestSeqNum = TestSeqNum + 1
        End If
    Next Ts
    
    If MeasV_Flag = True Then
        Set MeasValue = MeasValue_Temp
    End If
        
    TheHdw.Digital.Patgen.HaltWait
    
    If Final_Calc <> True Then
    If TrimCal_Name <> "" Then
        If Cal_Eqn <> "" Then
            Call ProcessCalcEquation(Cal_Eqn)
            If TrimCal_Name_array(0) = "C" Then
                TrimCalCap_value = GetStoredCaptureData(TrimCal_Name_array(1))
                For Each site In TheExec.sites.Active
                    MeasValue(site) = TrimCalCap_value.Element(0)
                   If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", " & TrimCal_Name_array(1) & " =" & MeasValue(site)
                Next site
                
            Else
                TrimCal_value = GetStoredMeasurement(TrimCal_Name)
                For Each site In TheExec.sites.Active
                    MeasValue(site) = TrimCal_value.Pins(0).Value(site)
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", " & TrimCal_Name & " =" & MeasValue(site)
                Next site
            End If
        End If
                
    End If
    End If
    
    Exit Function
err:
    Stop
    Resume Next
End Function

Public Function Trim_CIOTXbyTable() As Long
'Added  20190509

    Dim RegStr As String
    Dim PatStr() As String
    Dim trimvalue As String
    Dim i, j, k As Integer
    Dim TempCnt As Integer
    Dim RegWidth As Integer
    Dim RegSweeps() As Integer
    Dim OutWf() As New DSPWave
    ReDim OutWf(0) As New DSPWave
    Dim WorkBookName As Workbook
    Dim WorkSheetName As Worksheet
    
    Dim ciotx_sheetnumber As Integer
    Dim sheet_Cnt As Integer
    Dim SheetExist As Boolean: SheetExist = False
    sheet_Cnt = ActiveWorkbook.Sheets.Count
    For i = 1 To sheet_Cnt
        RegStr = LCase(Sheets(i).Name)
        If LCase(Sheets(i).Name) Like "*ciotx_trimtable*" Then
            SheetExist = True
            ciotx_sheetnumber = i
            Exit For
        End If
    Next i
    
    
    If SheetExist = True Then
        Set WorkBookName = Application.ActiveWorkbook
        'Set WorkSheetName = WorkBookName.Sheets("CIOTX_TrimTable")
         Set WorkSheetName = WorkBookName.Sheets(UCase(Sheets(ciotx_sheetnumber).Name))
        For i = 1 To CLng(WorkSheetName.UsedRange.Rows.Count)
            If CStr(WorkSheetName.Cells(i, 1)) <> "" Then
                If CStr(WorkSheetName.Cells(i, 1)) Like "*Pat*" Then                        ' Record each register parameter (size/width) from trim table
                    RegWidth = 0
                    PatStr = Split(CStr(WorkSheetName.Cells(i, 1)), ":")
                    ReDim RegSweeps(WorkSheetName.Cells(i, 1).End(xlToRight).Column - 2)
                    For j = 0 To WorkSheetName.Cells(i, 1).End(xlToRight).Column - 2
                        RegStr = Mid(CStr(WorkSheetName.Cells(i, j + 2)), InStr(1, CStr(WorkSheetName.Cells(i, j + 2)), "["))
                        RegStr = WorksheetFunction.Substitute(WorksheetFunction.Substitute(RegStr, "[", ""), "]", "")
                        RegSweeps(j) = CInt(RegStr)                                         ' Each regsiter size
                        RegWidth = RegWidth + CInt(RegStr)                                  ' Each sweep registers width
                    Next j
                Else
                    OutWf(UBound(OutWf)).CreateConstant 0, CLng(RegWidth), DspLong
                    For j = 0 To WorkSheetName.Cells(i, 1).End(xlToRight).Column - 2        ' Record each register value from trim table
                        trimvalue = CStr(WorkSheetName.Cells(i, j + 2))
                        If trimvalue Like "*x*" Or trimvalue Like "*X*" Then                ' Avoid format error , 0xA0 --> 0A0 , 0X55 --> 055
                            trimvalue = Replace(trimvalue, "x", "")
                            trimvalue = Replace(trimvalue, "X", "")
                        End If
                        For k = 0 To RegSweeps(j) - 1                                       ' Format LSB ----> MSB
                            OutWf(UBound(OutWf)).Element(TempCnt) = CLng((CInt(WorksheetFunction.Hex2Dec(CStr(trimvalue))) And 2 ^ k) / 2 ^ k)
                            TempCnt = TempCnt + 1
                        Next k
                    Next j
                    TempCnt = 0
                    AddStoredCaptureData PatStr(1) & "_DigSrcTable_" & CStr(WorkSheetName.Cells(i, 1)), OutWf(UBound(OutWf))
                    ReDim Preserve OutWf(UBound(OutWf) + 1)
                End If
            End If
        Next i
    End If
    
End Function

Public Function Trim_SetupandmeasureF_UVI80(MeasF_PinS_SingleEnd As PinList, MeasF_Interval As String, MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, MeasValue As SiteDouble, code As SiteLong, code_bin() As String, out_str() As String, b_Trimfinish As Boolean)
    Dim site As Variant
    Dim MeasF_EventSource As FreqCtrEventSrcSel
    Dim MeasF_EnableVtMode As Boolean
    
    Call Freq_ProcessEventSourceTerminationMode(MeasF_EventSourceWithTerminationMode, MeasF_EventSource, MeasF_EnableVtMode)
    
    ''''''''''''''''''''''''''''setup measure F'''''''''''''''''''''''''''''''''
    With TheHdw.Digital.Pins(MeasF_PinS_SingleEnd).FreqCtr
        .EventSource = MeasF_EventSource '' VOH
        .EventSlope = Positive
        .Interval = MeasF_Interval
        .Enable = IntervalEnable
        .Clear
    End With
     
   
    
    Dim CounterValue As New SiteDouble
    
    TheHdw.Digital.Pins(MeasF_PinS_SingleEnd).FreqCtr.Clear
    TheHdw.Digital.Pins(MeasF_PinS_SingleEnd).FreqCtr.start
    
'    If CustomizeWaitTime <> "" Then
'        thehdw.Wait (CDbl(CustomizeWaitTime))
'    End If
        
    
''        freq = CounterValue.Math.Divide(interval)
    
    ''''''''''''''''''''''''''offline''''''''''''''''''''''''''''''''''''''''''''
    If TheExec.TesterMode = testModeOffline Then
        'Dim Pin As Variant
        '900000+code*100000

        MeasValue = code.Multiply(-50000).Add(1100000)     '0.000015 * -1 + code * 0.000001
        
        If b_Trimfinish = False And gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "Trimming"
        ElseIf gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "TrimResult"
        End If
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Frequency = " & MeasValue(site)
            Next site
        End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
    
        CounterValue = TheHdw.Digital.Pins(MeasF_PinS_SingleEnd).FreqCtr.Read
        MeasValue = CounterValue.Divide(MeasF_Interval)
        
        If b_Trimfinish = False And gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "Trimming"
        ElseIf gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "TrimResult"
        End If
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Frequency = " & MeasValue(site)
            Next site
        End If
    End If
    
End Function

Public Function Trim_SetupandmeasureI_UVI80(MeasI_Pin As PinList, MeasValue As SiteDouble, MeasureI_Range As Double, code As SiteLong, code_bin() As String, out_str() As String, b_Trimfinish As Boolean)
 '''''''''''''''''setup UVI80 for measI''''''''''''''''''''''''''''''''''
   Dim factor As Long
   Dim WaitTime As Double
   Dim site As Variant
   factor = 1
   
    If MeasureI_Range > 2 * factor Then
        MeasureI_Range = 2 * factor
        WaitTime = 1.6 * ms
    ElseIf MeasureI_Range > 1 * factor Then
        MeasureI_Range = 2 * factor
        WaitTime = 1.6 * ms
    ElseIf MeasureI_Range > 0.2 * factor Then
        MeasureI_Range = 1 * factor
        WaitTime = 1.6 * ms
    ElseIf MeasureI_Range > 0.02 * factor Then
        MeasureI_Range = 0.2 * factor
        WaitTime = 260 * us
    ElseIf MeasureI_Range > 0.002 * factor Then
        MeasureI_Range = 0.02 * factor
        WaitTime = 1.5 * ms
    ElseIf MeasureI_Range > 0.0002 * factor Then
        MeasureI_Range = 0.002 * factor
        WaitTime = 11 * ms
    ElseIf MeasureI_Range > 0.00002 * factor Then
        MeasureI_Range = 0.0002 * factor
        WaitTime = 1.4 * ms
    Else
        MeasureI_Range = 0.00002 * factor
        WaitTime = 6 * ms
    End If
      
    
    With TheHdw.DCVI.Pins(MeasI_Pin)
        .Gate = False
        .mode = tlDCVIModeVoltage
        .Voltage = 0
        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
        .SetCurrentAndRange MeasureI_Range, MeasureI_Range
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    
    With TheHdw.DCVI.Pins(MeasI_Pin)
        .Meter.mode = tlDCVIMeterCurrent
        .Meter.CurrentRange.Value = MeasureI_Range
    End With
    
    TheHdw.Wait (WaitTime)
     '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
     
    ''''''''''''''''''offline simulate''''''''''''''''''''''''''''''
    If TheExec.TesterMode = testModeOffline Then
        'Dim Pin As Variant
        
        MeasValue = code.Multiply(0.000001).Add(-0.000015)  '0.000015 * -1 + code * 0.000001
        
        
        If gl_Disable_HIP_debug_log = False Then
            If b_Trimfinish = False Then
                TheExec.Datalog.WriteComment "Trimming"
            Else
                TheExec.Datalog.WriteComment "TrimResult"
            End If
                
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Current = " & MeasValue(site)
            Next site
        End If
        
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
        
        MeasValue = TheHdw.DCVI.Pins(MeasI_Pin.Value).Meter.Read(tlStrobe, 10)
        
        If gl_Disable_HIP_debug_log = False Then

            If b_Trimfinish = False Then
                TheExec.Datalog.WriteComment "Trimming"
            Else
                TheExec.Datalog.WriteComment "TrimResult"
            End If
            
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Current = " & MeasValue(site)
            Next site
        End If
    
    End If





End Function

Public Function Trim_SetupandmeasureV_UVI80(MeasV_Pin As String, MeasValue As SiteDouble, code As SiteLong, code_bin() As String, out_str() As String, b_Trimfinish As Boolean)
        Dim site As Variant
 '''''''''''''''setup UVI80 for meas V''''''''''''''''''
    With TheHdw.DCVI.Pins(MeasV_Pin)
        .Gate = False
        .Disconnect tlDCVIConnectDefault
        .mode = tlDCVIModeHighImpedance
        .Connect tlDCVIConnectHighSense
        .Voltage = 6
        .current = 0
         'thehdw.Wait 1 * ms
        .Gate = True
    End With
    
    With TheHdw.DCVI.Pins(MeasV_Pin)
        .Meter.mode = tlDCVIMeterVoltage
    End With
    TheHdw.Wait 1 * ms
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim cal As New SiteDouble
    
    ''''''''''''''''''offline simulate''''''''''''''''''''''''''''''
'    If theexec.TesterMode = testModeOffline Then
'        'Dim Pin As Variant
'
'        For Each Site In theexec.sites.Active
'            MeasValue(Site) = code * 0.1
'        Next Site
''        If InStr(TheExec.DataManager.InstanceName, "MTRGR_T4P2") <> 0 Then
''            TheExec.Datalog.WriteComment "trimming"
''            cal = MeasValue.Subtract(0.4).Divide(0.7975).Subtract(1)
''            For Each Site In TheExec.sites.Active
''                TheExec.Datalog.WriteComment "Site " & Site & ",Code " & code_bin(Site) & ", Src_code = " & out_str(Site) & ", Gain_error = " & cal(Site) & ", Voltage = " & MeasValue(Site)
''            Next Site
''            MeasValue = cal
''        Else
'        If gl_Disable_HIP_debug_log = False Then
'
'            If b_Trimfinish = False Then
'                theexec.Datalog.WriteComment "Trimming"
'            Else
'                theexec.Datalog.WriteComment "TrimResult"
'            End If
'
'            For Each Site In theexec.sites.Active
'                theexec.Datalog.WriteComment "Site " & Site & ",Code " & code_bin(Site) & ", Src_code = " & out_str(Site) & ", Voltage = " & MeasValue(Site)
'            Next Site
'        End If
''        End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Else
        Dim MeasValue_PinListData As New PinListData: MeasValue_PinListData = TheHdw.DCVI.Pins(MeasV_Pin).Meter.Read(tlStrobe, 10)
        For Each site In TheExec.sites.Active
        
            If UCase(TheExec.DataManager.instanceName) Like "*BGTRIM*" Or UCase(TheExec.DataManager.instanceName) Like "*BGMEAS*" Then
                Dim PinsArray() As String: PinsArray = Split(MeasV_Pin, ",")
                MeasValue = Abs(MeasValue_PinListData.Pins(PinsArray(0)).Value - MeasValue_PinListData.Pins(PinsArray(1)).Value)
            Else
                MeasValue = MeasValue_PinListData.Pins(MeasV_Pin).Value
            End If
            
        Next site
'        If InStr(TheExec.DataManager.InstanceName, "MTRGR_T4P2") <> 0 Then
'            TheExec.DataLog.WriteComment "trimming"
'            cal = MeasValue.Subtract(0.4).Divide(0.7975).Subtract(1)
'            For Each site In TheExec.sites.Active
'                TheExec.DataLog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Gain_error = " & cal(site) & ", Voltage = " & MeasValue(site)
'            Next site
'            MeasValue = cal
'        Else
        If gl_Disable_HIP_debug_log = False Then

            If b_Trimfinish = False Then
                TheExec.Datalog.WriteComment "Trimming"
            Else
                TheExec.Datalog.WriteComment "TrimResult"
            End If

            For Each site In TheExec.sites.Active
                If UCase(TheExec.DataManager.instanceName) Like "*BGTRIM*" Or UCase(TheExec.DataManager.instanceName) Like "*BGMEAS*" Then
                    TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site)
                    TheExec.Datalog.WriteComment PinsArray(0) & " Voltage = " & MeasValue_PinListData.Pins(PinsArray(0)).Value
                    TheExec.Datalog.WriteComment PinsArray(1) & " Voltage = " & MeasValue_PinListData.Pins(PinsArray(1)).Value
                    TheExec.Datalog.WriteComment "Difference Voltage = " & MeasValue
                Else
                    TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Voltage = " & MeasValue(site)
                End If
            Next site
        End If
'        End If
    
'    End If

    With TheHdw.DCVI.Pins(MeasV_Pin)
        .Gate(tlDCVIGateHiZ) = False
        .Disconnect
        .mode = tlDCVIModeCurrent
    End With




End Function


Public Function Addstorecapture_Trim(CUS_Str_DigCapData As String, OutDspWave As DSPWave, DigCap_Sample_Size As Long, DigCap_DataWidth As Long, Optional CUS_Str_MainProgram As String, _
                        Optional BypassAllDigCapTestLimit As Boolean = False)
    
    Dim site As Variant
    Dim i As Long, j As Long
    Dim Str_PrintBinary As New SiteVariant
    Dim ConvertedDataWf As New DSPWave
    Dim SourceBitStrmWf As New DSPWave
    Dim NoOfSamples As New SiteLong
    
    Dim FlexibleConvertedDataWf() As New DSPWave

    '' 20160328
    Dim TestLimitWithTestName As New PinListData
    
    Dim TestInstanceName As String
    TestInstanceName = TheExec.DataManager.instanceName

    ''20170418 - Move out from If DigCap_DataWidth <> 0 And InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") = 0 And CUS_Str_MainProgram = "" Then
    Dim p As Long
    Dim PinName As String
    Dim DigCapValue As New PinListData
    Dim b_FirstTimeSwitch As Boolean
    
    '' 20160211 - Process format by DigCap_DataWidth, capture word size is fixed
    If DigCap_DataWidth <> 0 And InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") = 0 Then
    
            Dim CalcOutputDSPWave As New DSPWave
            Dim CalcEyeWidth As New SiteLong
            Dim FinalEyeOutBitNum As Long
            Dim TestLimitForEyeSweep As New DSPWave
        
        ''20170811 - EyeSweep for LPDPRX
        If UCase(CUS_Str_MainProgram) = UCase("LPDPRX_EyeSweep") Then
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (DigCap_DataWidth) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
            SourceBitStrmWf = OutDspWave
        
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            
'            Dim CalcOutputDSPWave As New DSPWave
'            Dim CalcEyeWidth As New SiteLong
'            Dim FinalEyeOutBitNum As Long
            FinalEyeOutBitNum = DigCap_Sample_Size / 32
            rundsp.LPDPRX_EyeSweep ConvertedDataWf, FinalEyeOutBitNum, CalcOutputDSPWave, CalcEyeWidth
            
            For Each site In TheExec.sites
                Str_PrintBinary(site) = ""
                For i = 1 To CalcOutputDSPWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & CalcOutputDSPWave(site).Element(i - 1)
    ''                If i Mod (DigCap_DataWidth) = 0 Then
    ''                    Str_PrintBinary(Site) = Str_PrintBinary(Site) & ","
    ''                End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") Output eye bits = " & FinalEyeOutBitNum & ", Binary string = " & Str_PrintBinary(site))
                ''20170510 Store Binary String for Eye Diagram
                Eye_Diagram_Binary(TheExec.Flow.var("SrcCodeIndx").Value + 31)(site) = Str_PrintBinary(site)
            Next site
'            Dim TestLimitForEyeSweep As New DSPWave
            rundsp.BitWf2Arry CalcOutputDSPWave, DigCap_DataWidth, NoOfSamples, TestLimitForEyeSweep
    
            b_FirstTimeSwitch = True
            For Each site In TheExec.sites
                PinName = "EyeCapWord_"
                Exit For
            Next site
            For Each site In TheExec.sites
                For i = 1 To TestLimitForEyeSweep.SampleSize
                    If b_FirstTimeSwitch Then
                        DigCapValue.AddPin (PinName & CStr(i - 1))
                    End If
                    DigCapValue.Pins(PinName & CStr(i - 1)).Value(site) = TestLimitForEyeSweep(site).Element(i - 1)
                Next i
                b_FirstTimeSwitch = False
            Next site
    
            If BypassAllDigCapTestLimit = False Then
                For p = 0 To DigCapValue.Pins.Count - 1
                    'TheExec.Flow.TestLimit DigCapValue.Pins(p), 0, 2 ^ DigCap_DataWidth - 1, Tname:=TestInstanceName & "_EyeCpatureCode_" & p, PinName:="EyeCpatureCode_" & p, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
                Next p
                'TheExec.Flow.TestLimit resultVal:=CalcEyeWidth, Tname:=TestInstanceName & "_EyeWidth", ForceResults:=tlForceFlow
            End If
        '20170811 PCIE Eye Sweep
        ElseIf UCase(CUS_Str_MainProgram) = UCase("PCIE_EyeSweep") Then
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (DigCap_DataWidth) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
            SourceBitStrmWf = OutDspWave
        
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            
'            Dim CalcOutputDSPWave As New DSPWave
'            Dim CalcEyeWidth As New SiteLong
'            Dim FinalEyeOutBitNum As Long
            FinalEyeOutBitNum = DigCap_Sample_Size / 20
            rundsp.PCIE_EyeSweep ConvertedDataWf, FinalEyeOutBitNum, CalcOutputDSPWave, CalcEyeWidth
            
            For Each site In TheExec.sites
                Str_PrintBinary(site) = ""
                For i = 1 To CalcOutputDSPWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & CalcOutputDSPWave(site).Element(i - 1)
    ''                If i Mod (DigCap_DataWidth) = 0 Then
    ''                    Str_PrintBinary(Site) = Str_PrintBinary(Site) & ","
    ''                End If
                Next i
               If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") Output eye bits = " & FinalEyeOutBitNum & ", Binary string = " & Str_PrintBinary(site))
                ''20170510 Store Binary String for Eye Diagram
                Eye_Diagram_Binary(TheExec.Flow.var("SrcCodeIndx").Value + 31)(site) = Str_PrintBinary(site)
            Next site
'            Dim TestLimitForEyeSweep As New DSPWave
            rundsp.BitWf2Arry CalcOutputDSPWave, DigCap_DataWidth, NoOfSamples, TestLimitForEyeSweep
    
            b_FirstTimeSwitch = True
            For Each site In TheExec.sites
                PinName = "EyeCapWord_"
                Exit For
            Next site
            For Each site In TheExec.sites
                For i = 1 To TestLimitForEyeSweep.SampleSize
                    If b_FirstTimeSwitch Then
                        DigCapValue.AddPin (PinName & CStr(i - 1))
                    End If
                    DigCapValue.Pins(PinName & CStr(i - 1)).Value(site) = TestLimitForEyeSweep(site).Element(i - 1)
                Next i
                b_FirstTimeSwitch = False
            Next site
    
            If BypassAllDigCapTestLimit = False Then
                For p = 0 To DigCapValue.Pins.Count - 1
                    'TheExec.Flow.TestLimit DigCapValue.Pins(p), 0, 2 ^ DigCap_DataWidth - 1, Tname:=TestInstanceName & "_EyeCpatureCode_" & p, PinName:="EyeCpatureCode_" & p, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
                Next p
                'TheExec.Flow.TestLimit resultVal:=CalcEyeWidth, Tname:=TestInstanceName & "_EyeWidth", ForceResults:=tlForceFlow
            End If
        Else
    
            For Each site In TheExec.sites
                For i = 1 To OutDspWave.SampleSize
                    Str_PrintBinary(site) = Str_PrintBinary(site) & OutDspWave(site).Element(i - 1)
                    If i Mod (DigCap_DataWidth) = 0 Then
                        Str_PrintBinary(site) = Str_PrintBinary(site) & ","
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ") DigCap Bit Size = " & DigCap_Sample_Size & ", Data Width = " & DigCap_DataWidth & ", Binary string = " & Str_PrintBinary(site))
            
            Next site
            SourceBitStrmWf = OutDspWave
        
            rundsp.BitWf2Arry SourceBitStrmWf, DigCap_DataWidth, NoOfSamples, ConvertedDataWf
            
            
            '' 20160211 - Get pin name from dsp wave
    ''        Dim p As Long
    ''        Dim PinName As String
    ''        Dim DigCapValue As New PinListData
    ''        Dim b_FirstTimeSwitch As Boolean
            b_FirstTimeSwitch = True
            For Each site In TheExec.sites
                PinName = OutDspWave(site).Info.WaveName & "_DigCapWord_"
                Exit For
            Next site
            For Each site In TheExec.sites
                For i = 1 To ConvertedDataWf.SampleSize
                    If b_FirstTimeSwitch Then
                        DigCapValue.AddPin (PinName & CStr(i - 1))
                    End If
                    DigCapValue.Pins(PinName & CStr(i - 1)).Value(site) = ConvertedDataWf(site).Element(i - 1)
                Next i
                b_FirstTimeSwitch = False
            Next site
            
            If BypassAllDigCapTestLimit = False Then
                For p = 0 To DigCapValue.Pins.Count - 1
                    'TheExec.Flow.TestLimit DigCapValue.Pins(p), 0, 2 ^ DigCap_DataWidth - 1, Tname:=TestInstanceName & "_CpatureCode_" & p, PinName:="CpatureCode_" & p, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
                Next p
            End If
    ''        Call CUS_VFI_MainProgram_ECID(CUS_Str_MainProgram, DigCapValue)

        End If


        
    '' 20160212 - Process format by DSSC_OUT, capture word size is flexible, also parse with/without test name.
    ElseIf InStr(UCase(CUS_Str_DigCapData), "DSSC_OUT") <> 0 Then

        Dim Split_Num() As String
        Dim StartNum As Long
        '' 20151231 - Add rule to check new format that include test name and parse bits
        Dim ParseStringByBits As String
        Dim ParseStringForTestName As String
        Dim DSSC_Out_DecompseByComma() As String
        Dim DSSC_Out_DecompseByColon() As String
        Dim b_DSSC_Out_InvolveTestName As Boolean
        ParseStringByBits = ""
        ParseStringForTestName = ""
        b_DSSC_Out_InvolveTestName = False
        Dim DecomposeTestName() As String
        Dim DecomposeParseDigCapBit() As String
        
        ''20160807 - Add directionary to store DigCap DSPwave
        Dim ParseStringForDirectionary As String
        Dim DecomposeDirectionary() As String
        Dim b_ParseForDirectionary_Switch As Boolean
        b_ParseForDirectionary_Switch = False
        
        Dim b_ParseForGrayCode_Switch As Boolean
        b_ParseForGrayCode_Switch = False
        Dim ParseStringForGrayCode As String
        Dim DecomposeGrayCode() As String
        
        If InStr(UCase(CUS_Str_DigCapData), ":") <> 0 Then
            b_DSSC_Out_InvolveTestName = True
            DSSC_Out_DecompseByComma = Split(CUS_Str_DigCapData, ",")
            For i = 0 To UBound(DSSC_Out_DecompseByComma)
                DSSC_Out_DecompseByColon = Split(DSSC_Out_DecompseByComma(i), ":")
                If UBound(DSSC_Out_DecompseByColon) > 0 Then
                    If ParseStringByBits = "" And ParseStringForTestName = "" Then
                        ParseStringByBits = DSSC_Out_DecompseByColon(0)
                        ParseStringForTestName = DSSC_Out_DecompseByColon(1)
                        If UBound(DSSC_Out_DecompseByColon) = 2 Then    '' Dictionary
                            ParseStringForDirectionary = DSSC_Out_DecompseByColon(2) & ","
                            ParseStringForGrayCode = ","
                        ElseIf UBound(DSSC_Out_DecompseByColon) = 3 Then    '' Dictionary and GrayCode
                            ParseStringForDirectionary = DSSC_Out_DecompseByColon(2) & ","
                            ParseStringForGrayCode = DSSC_Out_DecompseByColon(3) & ","
                        Else
                            ParseStringForDirectionary = ","
                            ParseStringForGrayCode = ","
                        End If
                    Else
                        ParseStringByBits = ParseStringByBits & "," & DSSC_Out_DecompseByColon(0)
                        ParseStringForTestName = ParseStringForTestName & "," & DSSC_Out_DecompseByColon(1)
                        
                        If b_ParseForDirectionary_Switch = False Then
                            b_ParseForDirectionary_Switch = True
                        Else
                            ParseStringForDirectionary = ParseStringForDirectionary & ","
                        End If
                        
                        If b_ParseForGrayCode_Switch = False Then
                            b_ParseForGrayCode_Switch = True
                        Else
                            ParseStringForGrayCode = ParseStringForGrayCode & ","
                        End If
                        
                        If UBound(DSSC_Out_DecompseByColon) = 2 Then    '' Dictionary
                            ParseStringForDirectionary = ParseStringForDirectionary & DSSC_Out_DecompseByColon(2)
                        End If
                        
                        If UBound(DSSC_Out_DecompseByColon) = 3 Then    '' Dictionary
                            ParseStringForDirectionary = ParseStringForDirectionary & DSSC_Out_DecompseByColon(2)
                            ParseStringForGrayCode = ParseStringForGrayCode & DSSC_Out_DecompseByColon(3)
                        End If
                    
                    End If
                End If
            Next i
            
            ''20161220-Remove comma in the last of string
            If Right(ParseStringForTestName, 1) = "," Then
                ParseStringForTestName = Left(ParseStringForTestName, (Len(ParseStringForTestName) - 1))
            End If
            If Right(ParseStringForDirectionary, 1) = "," Then
                ParseStringForDirectionary = Left(ParseStringForDirectionary, (Len(ParseStringForDirectionary) - 1))
            End If
            
            ParseStringByBits = "DSSC_OUT," & ParseStringByBits
            DecomposeTestName = Split(ParseStringForTestName, ",")
            DecomposeDirectionary = Split(ParseStringForDirectionary, ",")
            DecomposeGrayCode = Split(ParseStringForGrayCode, ",")
        Else
            ParseStringByBits = CUS_Str_DigCapData
        End If
        
        If Right(ParseStringByBits, 1) = "," Then
            ParseStringByBits = Left(ParseStringByBits, (Len(ParseStringByBits) - 1))
        End If
        DecomposeParseDigCapBit = Split(ParseStringByBits, ",")
        Dim StrParseDigCapBit As String
        
        For i = 1 To UBound(DecomposeParseDigCapBit)
            If i = 1 Then
                StrParseDigCapBit = DecomposeParseDigCapBit(i)
            Else
                StrParseDigCapBit = StrParseDigCapBit & "," & DecomposeParseDigCapBit(i)
            End If
        Next i
        DecomposeParseDigCapBit = Split(StrParseDigCapBit, ",")
        
        ReDim FlexibleConvertedDataWf(UBound(DecomposeParseDigCapBit)) As New DSPWave
        ''20160823-Store binary dsp wave after processed by DSSC_OUT
        Dim DSPWave_Binary() As New DSPWave
        ReDim DSPWave_Binary(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim DSPWave_GrayCode() As New DSPWave
        ReDim DSPWave_GrayCode(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim DSPWave_GrayCodeDec() As New DSPWave
        ReDim DSPWave_GrayCodeDec(UBound(DecomposeParseDigCapBit)) As New DSPWave
        
        Dim StartIndex As Long
        StartIndex = 0
        
        ''20161230-Add copy and site loop to pass data
        For Each site In TheExec.sites
            SourceBitStrmWf = OutDspWave.Copy
        Next site
        
        
        Dim width_Wf  As New DSPWave, OutWf As New DSPWave ', OutBinWf() As New DSPWave
       ' ReDim OutBinWf(UBound(DecomposeParseDigCapBit))
        width_Wf.CreateConstant 0, UBound(DecomposeParseDigCapBit) + 1   'Create space for DSP
        'OutWf.CreateConstant 0, UBound(DecomposeParseDigCapBit) + 1
        
        For Each site In TheExec.sites
            For i = 0 To UBound(DecomposeParseDigCapBit)
                width_Wf.ElementLite(i) = CLng(DecomposeParseDigCapBit(i))  'deliver data to dsp array
            Next i
        Next site
        
        rundsp.Split_Dspwave SourceBitStrmWf, width_Wf, OutWf                   ', OutBinWf
        
        For Each site In TheExec.sites
            For i = 0 To UBound(DecomposeParseDigCapBit)
                FlexibleConvertedDataWf(i).CreateConstant 0, 1
                FlexibleConvertedDataWf(i).Element(0) = OutWf.ElementLite(i)
            Next i
        Next site
       
        ''20160823-Modify dsp function to add one input argument to process DSPwave with binary format and use Directionary to store it.
        For i = 0 To UBound(DecomposeParseDigCapBit)
'            rundsp.FlexibleBitWf2Arry SourceBitStrmWf, StartIndex, CLng(DecomposeParseDigCapBit(i)), FlexibleConvertedDataWf(i), DSPWave_Binary(i)
            
            ''20160823-Store binary DSP wave by using Directionary
''            If DecomposeDirectionary(i) <> "" Then
''                Call AddStoredCaptureData(DecomposeDirectionary(i), DSPWave_Binary(i))
''            End If
            
            If UCase(DecomposeGrayCode(i)) = "GRAYCODE" Then
''                DSPWave_GrayCode(i).CreateConstant 0, DecomposeParseDigCapBit(i), DspLong
''                DSPWave_GrayCodeDec(i).CreateConstant 0, 1, DspLong
                
                Call rundsp.Transfer2GrayCode(DSPWave_Binary(i), DSPWave_GrayCode(i), DSPWave_GrayCodeDec(i))
                
            End If
            
            StartIndex = StartIndex + DecomposeParseDigCapBit(i)
        Next i
        
        ''20161215-Check the dictionary name, re-combine them to one dsp wave and store to dictionary if there have the same dictionary name cross multi-segment (over 24 bit).
        '' Separate dsp wave to different segment if over 24bits, this is for cover STDF display truncation issue.
        Dim CombineDSPBit2Dict As New Dictionary
        Dim KeyName As String
        If ParseStringForDirectionary <> "" Then
            CombineDSPBit2Dict.RemoveAll
            For i = 0 To UBound(DecomposeDirectionary)
                If DecomposeDirectionary(i) = "" Then
                    KeyName = "EMPTYSPACE_DICT_" & i
                Else
                    KeyName = LCase(DecomposeDirectionary(i))
                End If
                If i = 0 Then
                    CombineDSPBit2Dict.Add KeyName, CLng(DecomposeParseDigCapBit(i))
    
                Else
                    If CombineDSPBit2Dict.Exists(KeyName) Then
                        CombineDSPBit2Dict.Item(KeyName) = CombineDSPBit2Dict.Item(KeyName) + CLng(DecomposeParseDigCapBit(i))
                    Else
                        CombineDSPBit2Dict.Add KeyName, CLng(DecomposeParseDigCapBit(i))
                    End If
                End If
            Next i
            
            Dim CombineKeys() As Variant
            CombineKeys() = CombineDSPBit2Dict.Keys()
            
            StartIndex = 0
            ReDim AddToDict_DSP_Dec(CombineDSPBit2Dict.Count - 1) As New DSPWave
            ReDim AddToDict_DSP_Bin(CombineDSPBit2Dict.Count - 1) As New DSPWave
            Dim FinalLength As Long
            
            For i = 0 To CombineDSPBit2Dict.Count - 1
                FinalLength = CombineDSPBit2Dict.Item(CombineKeys(i))
''                rundsp.FlexibleBitWf2Arry SourceBitStrmWf, StartIndex, FinalLength, AddToDict_DSP_Dec(i), AddToDict_DSP_Bin(i)
                If InStr(CombineKeys(i), "EMPTYSPACE_DICT_") <> 0 Then
                Else
                    For Each site In TheExec.sites
                        AddToDict_DSP_Bin(i) = SourceBitStrmWf.Select(StartIndex, , FinalLength).Copy '.ConvertStreamTo(tldspSerial, FinalLength, 0, Bit0IsMsb)
                    Next site
                    Call AddStoredCaptureData(CStr(CombineKeys(i)), AddToDict_DSP_Bin(i))
                End If
                StartIndex = StartIndex + FinalLength
            Next i
        End If
        
        
        '' Debug use
        Dim BinaryCodeString As String
        Dim GrayCodeString As String
''        Dim j As Long
        For Each site In TheExec.sites
            For i = 0 To UBound(DecomposeParseDigCapBit)
                If UCase(DecomposeGrayCode(i)) = "GRAYCODE" Then
                    BinaryCodeString = ""
                    GrayCodeString = ""
                    For j = 0 To DSPWave_Binary(i).SampleSize - 1
                        BinaryCodeString = BinaryCodeString & DSPWave_Binary(i)(site).Element(j)
                        GrayCodeString = GrayCodeString & DSPWave_GrayCode(i)(site).Element(j)
                    Next j
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " DSSC_OUT part " & i & " binary code = " & BinaryCodeString)
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " DSSC_OUT part " & i & "   gray code = " & GrayCodeString)
                End If
            Next i
        Next site
                
        '' 20160317 - Test limit for DSSC_OUT
'        If b_DSSC_Out_InvolveTestName = True Then '' Test limit with test name
'            For i = 0 To UBound(DecomposeTestName)
'                If LCase(DecomposeTestName(i)) = "skip" Then
'                Else
'                    TestLimitWithTestName.AddPin (DecomposeTestName(i) & "_" & i)
'                    If UCase(DecomposeGrayCode(i)) = "GRAYCODE" Then
'                        TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i).Value = DSPWave_GrayCodeDec(i).Element(0)
'                    Else
'                        TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i).Value = FlexibleConvertedDataWf(i).Element(0)
'                    End If
'                    If BypassAllDigCapTestLimit = False Then
'                        If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=TheExec.DataManager.InstanceName & "_" & DecomposeTestName(i) & "_" & i, ForceResults:=tlForceNone, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'
'                        ElseIf MTR_CusDigCap <> "" And UCase(MTR_CusDigCap) = "CUS_DIGCAP_VIN" Then
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=TheExec.DataManager.InstanceName & "_" & DecomposeTestName(i) & "_" & MTR_VIN & "_" & i, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'                            MTR_CusDigCap = ""
'
'                        ElseIf TPModeAsCharz_GLB = True Then  ''CZ TP name force flow
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'
'                        ElseIf CUS_Str_MainProgram = "TMPS_BV" Then  ''TMPS_BV
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=TheExec.DataManager.InstanceName & "_" & DecomposeTestName(i) & "_" & i, ForceResults:=tlForceNone, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'                        Else
'                            TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(i) & "_" & i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, Tname:=TheExec.DataManager.InstanceName & "_" & DecomposeTestName(i) & "_" & i, ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'                        End If
'                    End If
'                End If
'            Next i
'
'        Else
'            If BypassAllDigCapTestLimit = False Then
'                For i = 0 To UBound(FlexibleConvertedDataWf)
'                    TheExec.Flow.TestLimit FlexibleConvertedDataWf(i).Element(0), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:="DSSC_OUT_Code_" & i, Tname:=TheExec.DataManager.InstanceName & "_DSSC_OUT_" & CStr(i - 1), ForceResults:=tlForceFlow, ScaleType:=scaleNoScaling, formatstr:="%.0f"
'                Next
'            End If
'        End If
    End If
End Function


Public Function Create_DigSrc_Data_Trim(DigSrc_pin As PinList, DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, _
                        DigSrc_Equation As String, ByVal DigSrc_Assignment As String, InDspWav As DSPWave, site As Variant, Optional CUS_Str_DigSrcData As String = "", _
                        Optional NumberPins As Long = 1, Optional MSB_First_Flag As Boolean = False) As Long
''                        Optional InDSPWave_Parallel As DSPWave)

    Dim str_eq_ary(1000) As String_Equation
    Dim Assignment_ary() As String
    Dim Eq_ary() As String
    Dim i As Long, j As Long, k As Long
    Dim DigSrc_array() As Long
    Dim Str As String
    Dim idx As Long
    Dim Ary() As String
    Dim RdIn() As String
    
    Dim RdIn_tmp() As String
    Dim Rd_Fix_data As String

    Dim TempString_Repeat As String

    ''20161121-According to sample size and pin number to create data array size
    ReDim DigSrc_array(DigSrc_Sample_Size - 1)
    InDspWav.CreateConstant 0, DigSrc_Sample_Size
    
    'TheExec.Datalog.WriteComment "DataSequence:" & DigSrc_Equation
    'TheExec.Datalog.WriteComment "Assignment:" & DigSrc_Assignment
    
    ''20160824
    Dim b_WithDictionary As Boolean
    Dim SrcDspWave As New DSPWave
    Dim Ary_Src_DSPWave() As Long
    Dim b_Pre_data As Boolean
    
    Dim b_Append_Data As Boolean
    
    Dim PrePostData_SplitByAnd() As String
    Dim PrePostData_BinaryConstant As String
    
    ''20160909 - Append pre and post data  as 111&DictionA&000
    Dim b_AppendPrePostData As Boolean
    Dim PreBinDataString As String
    Dim PostBinDataString As String
    
    ''20160824-With "Repeat" keyword of DigSrc_Assignment
    If DigSrc_Assignment <> "" And InStr(LCase(DigSrc_Assignment), "repeat") = 0 Then
        Assignment_ary = Split(DigSrc_Assignment, ";")
        idx = 0
        
        For i = 0 To UBound(Assignment_ary)
            Ary = Split(Assignment_ary(i), "=")

            If UBound(Ary) > 1 Then           'check multi source with same datta
''                For j = 0 To UBound(Ary) - 1
''                    str_eq_ary(idx).Name = Ary(j)
''                    str_eq_ary(idx).value_string = Ary(UBound(Ary))
''                    idx = idx + 1
''                Next j
            Else
                str_eq_ary(idx).Name = Ary(0)
                
                ''20160825 - Pre/Post data as 111&DictionA or DictionA&101
                b_Pre_data = False
                b_Append_Data = False

                ''20160909 - Append pre and post data  as 111&DictionA&000
                b_AppendPrePostData = False
                PreBinDataString = ""
                PostBinDataString = ""
                PrePostData_SplitByAnd = Split(Ary(1), "&")
                
                If UBound(PrePostData_SplitByAnd) > 0 Then
                    If UBound(PrePostData_SplitByAnd) = 2 Then
                        b_AppendPrePostData = True
                        PreBinDataString = PrePostData_SplitByAnd(0)
                        Ary(1) = PrePostData_SplitByAnd(1)
                        PostBinDataString = PrePostData_SplitByAnd(2)
                    Else
                        b_Append_Data = True
                        
                        If Checker_ConstantBinary(PrePostData_SplitByAnd(0)) Then
                            b_Pre_data = True
                            Ary(1) = PrePostData_SplitByAnd(1)
                            PrePostData_BinaryConstant = PrePostData_SplitByAnd(0)
                        Else
                            b_Pre_data = False
                            Ary(1) = PrePostData_SplitByAnd(0)
                            PrePostData_BinaryConstant = PrePostData_SplitByAnd(1)
                        End If
                    End If
                End If
                ''20160825 - Check segment content whether Directionary
                ''20161014 - Check Dictionary whether need to calculation. EX: wdr0_4=010&[CAL_A-1]:0:3&110;wdr1_4=010&[CAL_A+1]:0:3&110
                RdIn = Split(Ary(1), ":")
                b_WithDictionary = Checker_WithDictionary(RdIn(0))

                If b_WithDictionary Then
                    
                    Dim b_IsDictNeedToCalc As Boolean
                    b_IsDictNeedToCalc = Checker_DictCalculated(RdIn(0))
                    
                    
                    
                    If b_IsDictNeedToCalc Then
                        Call AnalyzeDictCalculatedContent(RdIn(0), SrcDspWave)
                    Else
                        SrcDspWave = GetStoredCaptureData(RdIn(0))
                    End If
                
                    SrcDspWave = SrcDspWave.ConvertDataTypeTo(DspLong)
                    Ary_Src_DSPWave = SrcDspWave(site).Data
                    str_eq_ary(idx).value_string = ""

                    If UBound(RdIn) > 0 Then
                        
                        Dim StartNum As Long
                        Dim EndNum As Long
                        Dim StepNum As Long
                        StartNum = Int(RdIn(1))
                        EndNum = Int(RdIn(2))
                        If StartNum > EndNum Then
                            StepNum = -1
                        Else
                            StepNum = 1
                        End If
''                        For j = Int(RdIn(1)) To Int(RdIn(2))
                        For j = StartNum To EndNum Step StepNum
                            If UBound(RdIn) > 2 Then
                                If LCase(RdIn(3)) = "copy" Then
                                    For k = 0 To RdIn(4) - 1
                                        str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & Ary_Src_DSPWave(j)
                                    Next k
                                End If
                            Else
                                str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & Ary_Src_DSPWave(j)
                            End If
                        Next j
                    Else
                        For j = 0 To UBound(Ary_Src_DSPWave)
                            str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & Ary_Src_DSPWave(j)
                        Next j
                    End If
                    ''20160825 - Pre/Post data as 111&DictionA or DictionA&101
                    ''==========================================================================
                    If b_Append_Data Then
                        If b_Pre_data = False Then
                            str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & PrePostData_BinaryConstant
                        Else
                            str_eq_ary(idx).value_string = PrePostData_BinaryConstant & str_eq_ary(idx).value_string
                        End If
                    End If
                    ''==========================================================================
                    ''20160909 - Append pre and post data  as 111&DictionA&000
                    If b_AppendPrePostData Then
                        str_eq_ary(idx).value_string = PreBinDataString & str_eq_ary(idx).value_string & PostBinDataString
                    End If
                Else
                    If UBound(RdIn) > 0 Then
                        If LCase(RdIn(1)) = "copy" Then
                            For j = 0 To Len(RdIn(0)) - 1
                                For k = 0 To RdIn(2) - 1
                                    str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & Mid(RdIn(0), j + 1, 1)
                                Next k
                            Next j
                        End If
                    Else
                        str_eq_ary(idx).value_string = Ary(1)
                    End If
                    
                    ''20160825 - Pre/Post data as 111&DictionA or DictionA&101
                    ''==========================================================================
                    If b_Append_Data Then
                        If b_Pre_data = False Then
                            str_eq_ary(idx).value_string = str_eq_ary(idx).value_string & PrePostData_BinaryConstant
                        Else
                            str_eq_ary(idx).value_string = PrePostData_BinaryConstant & str_eq_ary(idx).value_string
                        End If
                    End If
                    ''==========================================================================
                    ''20160909 - Append pre and post data  as 111&DictionA&000
                    If b_AppendPrePostData Then
                        str_eq_ary(idx).value_string = PreBinDataString & str_eq_ary(idx).value_string & PostBinDataString
                    End If
                End If
                idx = idx + 1
            End If
        Next i
    End If

    For i = 0 To idx - 1
        Str = Str & str_eq_ary(i).Name & ":" & str_eq_ary(i).value_string & ","
    Next i

    'TheExec.Datalog.WriteComment "Site [" & Site & "] " & Str
            
    If DigSrc_Equation <> "" Then
        Eq_ary = Split(DigSrc_Equation, "+")
        idx = 0
        

        Dim StringDecomposeEqual() As String
        Dim StringDecomposeColon() As String
        Dim SelectedSrcDSPWave As New DSPWave
        
        ''20160830
        Dim b_StrInvolveCopy As Boolean
        '' 20160122 - Modify rule to soruce assigned bit from rd to replace souce all of bits, EX: repeat=rd:1:3
        Dim StartBit As Long
        Dim EndBit As Long
        
        Dim b_SelectallBits As Boolean
        Dim DataCopyTimes As Long
        Dim LoopIndex As Long
        
        If InStr(LCase(DigSrc_Assignment), "repeat") <> 0 Then
            StringDecomposeEqual = Split(DigSrc_Assignment, "=")
            
            ''20160830-Check DigSrc_Assignment has "copy" or not
            If InStr(LCase(DigSrc_Assignment), "copy") <> 0 Then
                b_StrInvolveCopy = True
            Else
                b_StrInvolveCopy = False
            End If
                        
            StringDecomposeColon = Split(LCase(StringDecomposeEqual(1)), ":")
            b_WithDictionary = Checker_WithDictionary(StringDecomposeColon(0))
            
            If b_WithDictionary Then
                SelectedSrcDSPWave = GetStoredCaptureData(StringDecomposeColon(0))
                
                If UBound(StringDecomposeColon) = 4 Then ''CAL_A:1:3:copy:2
                    StartBit = StringDecomposeColon(1)
                    EndBit = StringDecomposeColon(2)
                    DataCopyTimes = StringDecomposeColon(4)
                
                ElseIf UBound(StringDecomposeColon) = 2 Then
                    If b_StrInvolveCopy = True Then ''CAL_A:copy:2
                        StartBit = 0
                        EndBit = SelectedSrcDSPWave.SampleSize - 1
                        DataCopyTimes = StringDecomposeColon(2)
                    
                    Else                                        ''CAL_A:1:3
                        StartBit = StringDecomposeColon(1)
                        EndBit = StringDecomposeColon(2)
                        DataCopyTimes = 1
                    End If
                ElseIf UBound(StringDecomposeColon) = 0 Then ''CAL_A
                    StartBit = 0
                    EndBit = SelectedSrcDSPWave.SampleSize - 1
                    DataCopyTimes = 1
                End If
            Else
                If b_StrInvolveCopy Then    ''1101:copy:2
                    DataCopyTimes = StringDecomposeColon(2)
                Else                                ''1101
                    DataCopyTimes = 1
                End If
            End If
            
            LoopIndex = 0
            If b_WithDictionary Then
            
                Dim StepSize As Long
                If StartBit > EndBit Then
                    StepSize = -1
                Else
                    StepSize = 1
                End If
                
                For i = StartBit To EndBit Step StepSize
                    For j = 0 To DataCopyTimes - 1
                        If LoopIndex = 0 Then
                            TempString_Repeat = SelectedSrcDSPWave(site).Element(i)
                        Else
                            TempString_Repeat = TempString_Repeat & SelectedSrcDSPWave(site).Element(i)
                        End If
                        LoopIndex = LoopIndex + 1
                    Next j
                Next i
                
            Else
                For i = 0 To Len(StringDecomposeColon(0)) - 1
                    For j = 0 To DataCopyTimes - 1
                        TempString_Repeat = TempString_Repeat & Mid(StringDecomposeColon(0), i + 1, 1)
                    Next j
                Next i
            End If
            TempString_Repeat = "repeat=" & TempString_Repeat
            DigSrc_Assignment = TempString_Repeat

        End If
        
        ''20160824-Final process, Analyze DigSrc_Assignment to create InDspWav
        '' Number of equation segment
        For i = 0 To UBound(Eq_ary)
            If InStr(LCase(DigSrc_Assignment), "repeat") <> 0 Then
                If InStr(DigSrc_Assignment, "=") <> 0 Then RdIn = Split(DigSrc_Assignment, "=")
                If InStr(DigSrc_Assignment, ",") <> 0 Then RdIn = Split(DigSrc_Assignment, ",")
                Str = Trim(RdIn(UBound(RdIn)))
            Else
                Str = Trim(Find_Assignement(Eq_ary(i), str_eq_ary, , , MSB_First_Flag))
            End If
                
            '' Number of DigSrc_Assignment content
            For j = 1 To Len(Str)
                DigSrc_array(idx) = Val(Mid(Str, j, 1))
    ''           InDspWav(Site).Element(idx) = DigSrc_array(idx)
                idx = idx + 1
            Next j
        Next i
    
        InDspWav(site).Data = DigSrc_array

        If idx <> DigSrc_Sample_Size Then TheExec.Datalog.WriteComment "Num of bits in digsrc equation(" & idx & ") is not the same as DigSrc_SampleSize(" & DigSrc_Sample_Size & ")"
    End If
    
    'Call Printout_DigSrc(DigSrc_array, DigSrc_Sample_Size, DigSrc_DataWidth, NumberPins,site)

End Function

Public Function SetForceSweepVoltAndTName(ByRef Sweep_Info() As Power_Sweep, ByRef cap_data As String, Loop_count As Long)
    Dim sweep_volt As String: sweep_volt = ""
    Dim i As Long
    Dim Split_cap_data() As String
    Dim temp_str As String
    Dim temp_dict_name As String
    Dim index_name As String
    
    For i = 0 To UBound(Sweep_Info)
        index_name = Sweep_Info(i).Loop_Index_Name
        Call SetForceCondition(Sweep_Info(i).PinName & ":V:" & CStr(CDbl(Sweep_Info(i).from) + CDbl(Sweep_Info(i).step) * Loop_count))
        
        If (sweep_volt = "") Then
            sweep_volt = Replace(CStr(CDbl(Sweep_Info(i).from) + CDbl(Sweep_Info(i).step) * Loop_count), ".", "p")
        Else
            sweep_volt = sweep_volt & "_" & Replace(CStr(CDbl(Sweep_Info(i).from) + CDbl(Sweep_Info(i).step) * Loop_count), ".", "p")
        End If
    Next i
    
    If (cap_data <> "" And Sweep_Info(0).Key <> "") Then
        
        Split_cap_data = Split(cap_data, ",")
        If (UCase(Split_cap_data(0)) = "DSSC_OUT") Then
            For i = 1 To UBound(Split_cap_data)
                If Split_cap_data(i) <> "" Then
                    temp_str = Split(Split_cap_data(i), ":")(2)
                    temp_dict_name = temp_str & "_" & sweep_volt & "_" & CStr(TheExec.Flow.var(index_name).Value)
                    cap_data = Replace(cap_data, temp_str, temp_dict_name)
                End If
            Next i
        End If
    End If
    
    
    
    
End Function

Public Function SortSweepInfo(ByRef Sweep_Info() As Power_Sweep, ByRef Interpose_PrePat As String)

    Dim s As Variant
    Dim InterPre_split_array() As String
    Dim sweep_str() As String
    
    Dim Sweep_Count As Long: Sweep_Count = -1
    Dim sweep_pin As String: sweep_pin = ""
    Dim sweep_from As String: sweep_from = ""
    Dim sweep_stop As String: sweep_stop = ""
    Dim sweep_step As String: sweep_step = ""
    Dim aveg_flag As Boolean: aveg_flag = False
    Dim sweep_step_temp As Double
    
    ' Sweep:pinA:V:0.3:1.5:0.005:SrcCodeIndx:10
    

    ReDim Preserve Sweep_Info(0) As Power_Sweep

    
    InterPre_split_array = Split(Interpose_PrePat, ";")
    For Each s In InterPre_split_array
        If InStr(1, LCase(s), "sweep") <> 0 Then
            If (Sweep_Info(0).PinName = "") Then
                ReDim Sweep_Info(0) As Power_Sweep
            Else
                ReDim Preserve Sweep_Info(UBound(Sweep_Info) + 1) As Power_Sweep
            End If
            Dim Sweep_index As Integer: Sweep_index = UBound(Sweep_Info)
            sweep_str = Split(s, ":")
            Sweep_Info(Sweep_index).Loop_count = CLng(sweep_str(7))
            Sweep_Info(Sweep_index).PinName = sweep_str(1)
            Sweep_Info(Sweep_index).from = sweep_str(3)
            Sweep_Info(Sweep_index).stop = sweep_str(4)
            sweep_step_temp = CDbl(sweep_str(4)) - CDbl(sweep_str(3))
            Sweep_Info(Sweep_index).step = sweep_str(5)
            If (sweep_step_temp < 0) Then Sweep_Info(Sweep_index).step = "-" & Sweep_Info(Sweep_index).step
            Sweep_Info(Sweep_index).Count = Abs(CLng((CDbl(Sweep_Info(Sweep_index).stop) - CDbl(Sweep_Info(Sweep_index).from)) / CDbl(Sweep_Info(Sweep_index).step))) + 1
            Sweep_Info(Sweep_index).Loop_Index_Name = sweep_str(6)
            Sweep_Info(Sweep_index).Key = sweep_str(8)
            Interpose_PrePat = Replace(Interpose_PrePat, s, "")
'            AddStoredSweepInfo Sweep_Info(sweep_index).key, Sweep_Info(sweep_index)
            Sweep_index = Sweep_index + 1
        End If
        
        
    Next s
        
End Function





Public Function Merge_TName(TName_Ary() As String) As String

Dim i As Integer
Dim Tname As String: Tname = ""

For i = 0 To UBound(TName_Ary)
    If TName_Ary(i) = "" Then TName_Ary(i) = "X"
    TName_Ary(i) = Replace(TName_Ary(i), "_", "")
    
    If Tname <> "" Then
        Tname = Tname & TName_Ary(i) & "_"
    Else
        Tname = TName_Ary(i) & "_"
    End If
Next i


'Merge_TName = Tname
Merge_TName = Tname '& "(" & Replace(gl_TName_Pat, "_", "-") & ")"

    'If AbortTest Then Exit Function Else Resume Next
End Function





'   Segment:        0     1        2       3      4        5        6        7        8        9
'   TestName:      HAC_USERVAR2_UserVar3_Group_CATEGORY_USERVAR4_USERVAR5_USERVAR6_UserVar7_USERVAR8_
'   Meaning:       HAC_[Meas? ]_[H/N/L ]_SubB1_[Block ]_[ Pin  ]_[SubB2 ]_[  X   ]_[  X   ]_subr-seq_
'   X:                                    X1                       X2        X3       X4       X5

'               Instance name:    [Block]_[X1]_{patset}_[X2]_[X3]_[HV/NV/LV]
'
'               Test name:        HAC_[Meas type]_[HV/NV/LV]_[X1]_[Block]_[Pin name]_[X2]_[X3]_[X4]_[X5]_
'
'               01: HAC                 :FIXED
'               02: [Meas?]             :Measurement Tyep       :from VBT,  (? = V, I, F, C, or X )     EX:  MeasV/MeasI/MeasF/MeasC/MeasX
'               03: [HV/NV/LV]          :Power Condition        :from instance name,                    EX:  HV/NV/LV/MeasC/MeasX
'               04: [X1]                :from instance name
'               05: [Block]             :from instance name
'               06: [Pin name]          :from VBT (pinlistdata)
'               07: [X2]                :from                   :(1) flow table / (2)instance name  :(priority)
'               08: [X3]                :from instance name     :DSSC Segment name
'               09: [X4]                :from instance name     :DSSC Register
'               10: [X5]                :from [subr-seq]

Public Function Report_TName_From_Instance(MeasType As String, PinName As String, Optional Tname As String = "", Optional TestSeqNum As Integer = 0, Optional k As Long = 0, Optional SpecifyTname As String = "", Optional Sweep_Name As String, Optional SweepY_Name As String, Optional ForceResult As tlLimitForceResults = tlForceFlow) As String

        'Modify from M9 module
        Dim instanceName As String
        Dim InstanceName_WO_Pset As String
        Dim InstNameSegs() As String
        Dim TNameSeg(9) As String
        Dim site As Variant
        Dim i As Long
        Dim tempAry() As String
        Dim SubInstNameSegs() As String
        instanceName = UCase(TheExec.DataManager.instanceName)

        If gl_Current_Instance_Tname <> instanceName Then
            gl_Current_Instance_Tname = instanceName
            gl_Current_Instance_Tname_subblock = Application.Worksheets(TheExec.Flow.Raw.SheetInRun).range("AM" & CStr(TheExec.Flow.Raw.GetCurrentLineNumber + 5)).Value
        End If
'        PatSetName = Trim(PatSetName)
'        InstanceName_WO_Pset = Replace(InstanceName, UCase(PatSetName), "")
'        InstanceName_WO_Pset = Replace(InstanceName_WO_Pset, "__", "_")

        ' At Head   :   "DCTEST"
        If InstanceName_WO_Pset Like "DCTEST_*" Then
            instanceName = Replace(InstanceName_WO_Pset, "DCTEST_", "")
        End If
        ' All places:   "VIR"
        instanceName = Replace(instanceName, "_VIR_", "_")

        InstNameSegs = Split(instanceName, "_")

        'Instance name:    [Block]_[X1]_{patset}_[X2]_[X3]_[HV/NV/LV]
        'Test name:        HAC_____[Meas type]_[HV/NV/LV]_[X1]_[Block]_[Pin name]_[X2]_[X3]_[X4]_[X5]_

        TNameSeg(0) = "HAC"
        TNameSeg(1) = "Meas"
        TNameSeg(2) = InstNameSegs(UBound(InstNameSegs))                '[HV/NV/LV]
        TNameSeg(3) = "x"                                               '[X1] : sub-block-name-1
        TNameSeg(4) = InstNameSegs(0)                                   '[Block]
        TNameSeg(5) = "{pinname}"                                       '[Pin-name]
        TNameSeg(6) = "x"                                               '[X2] : sub-block-name-2
        TNameSeg(7) = "x"                                               '[X3] : X3 / DSSC Segment name
        TNameSeg(8) = "x"                                               '[X4] :    / DSSC Register
        TNameSeg(9) = "x"                                               '[X5] : subr-seq#

        '[H/N/L]
        TNameSeg(2) = Replace(UCase(TNameSeg(2)), "V", "")

        '[X1]
        If UBound(InstNameSegs) >= 2 Then
            If gl_Current_Instance_Tname_subblock <> "" Then
                TNameSeg(3) = gl_Current_Instance_Tname_subblock
            Else
                TNameSeg(3) = InstNameSegs(1)
            End If
        End If

        '20180314 For CZ Sweep Tname
        Dim SetupName As String
        Dim X_StepName As String
        Dim Y_StepName As String
        Dim X_ApplyToPin As String
        Dim Y_ApplyToPin As String
        Dim X_RangeFrom As Double
        Dim Y_RangeFrom As Double
        Dim X_CurrentPointVal_Str As String
        Dim Y_CurrentPointVal_Str As String

        'Call CZ_TNum_Increment       '20180604 TER

        If TheExec.DevChar.Setups.IsRunning = True Then
            SetupName = TheExec.DevChar.Setups.ActiveSetupName
            With TheExec.DevChar.Setups(SetupName)
                If .Shmoo.Axes.Count > 1 Then
                    X_StepName = .Shmoo.Axes(tlDevCharShmooAxis_X).StepName
                    Y_StepName = .Shmoo.Axes(tlDevCharShmooAxis_Y).StepName
                    X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
                    Y_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.range.from
                    X_ApplyToPin = .Shmoo.Axes.Item(tlDevCharShmooAxis_X).ApplyTo.Pins
                    Y_ApplyToPin = .Shmoo.Axes.Item(tlDevCharShmooAxis_Y).ApplyTo.Pins
                Else
                    X_StepName = .Shmoo.Axes(tlDevCharShmooAxis_X).StepName
                    X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range.from
                    X_ApplyToPin = .Shmoo.Axes.Item(tlDevCharShmooAxis_X).ApplyTo.Pins
                End If
            End With

            If Not ((TheExec.DevChar.Results(SetupName).startTime Like "1/1/0001*" Or TheExec.DevChar.Results(SetupName).startTime Like "0001/1/1*")) Then

                With TheExec.DevChar.Setups(SetupName)
                    If .Shmoo.Axes.Count > 1 Then
                        For Each site In TheExec.sites.Active
                            XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                            YVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
                        Next site
                        If XVal = X_RangeFrom And YVal = Y_RangeFrom Then
                            gl_flag_end_shmoo = False
                        End If
                        If gl_flag_end_shmoo = False Then

                            If InStr(CStr(XVal), ".") > 0 Then
                                X_CurrentPointVal_Str = Replace(CStr(XVal), ".", "P")
                            Else
                                X_CurrentPointVal_Str = CStr(XVal) & "P0"
                            End If
                            If .Shmoo.Axes(tlDevCharShmooAxis_X).TrackingParameters.Count > 0 Then
                                TNameSeg(7) = "MULTI&" & X_CurrentPointVal_Str
                            Else
                                TNameSeg(7) = X_ApplyToPin & "&" & X_CurrentPointVal_Str
                            End If

                            If InStr(CStr(YVal), ".") > 0 Then
                                Y_CurrentPointVal_Str = Replace(CStr(YVal), ".", "P")
                            Else
                                Y_CurrentPointVal_Str = CStr(YVal) & "P0"
                            End If

                            If .Shmoo.Axes(tlDevCharShmooAxis_Y).TrackingParameters.Count > 0 Then
                                TNameSeg(7) = TNameSeg(7) & "&MULTI&" & Y_CurrentPointVal_Str
                            Else
                                TNameSeg(7) = TNameSeg(7) & "&" & Y_ApplyToPin & "&" & Y_CurrentPointVal_Str
                            End If
                        End If
                    Else
                        For Each site In TheExec.sites.Active
                            XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                        Next site
                        If XVal = X_RangeFrom Then
                            gl_flag_end_shmoo = False
                        End If
                        If gl_flag_end_shmoo = False Then
                            If InStr(CStr(XVal), ".") > 0 Then
                                X_CurrentPointVal_Str = Replace(CStr(XVal), ".", "P")
                            Else
                                X_CurrentPointVal_Str = CStr(XVal) & "P0"
                            End If
                            If .Shmoo.Axes(tlDevCharShmooAxis_X).TrackingParameters.Count > 0 Then
                                TNameSeg(7) = "MULTI&" & X_CurrentPointVal_Str
                            Else
                                TNameSeg(7) = Replace(X_ApplyToPin, "_", "") & "&" & X_CurrentPointVal_Str 'Remove "_" in pin name @190612 CWCIOU
                            End If
                        End If
                    End If
                End With
            Else
                If gl_flag_CZ_Nominal_Measured_1st_Point Then
                Else
                    'Call CZ_TNum_Decrement       '20180613 TER
                    gl_flag_CZ_Nominal_Measured_1st_Point = True
                End If
            End If
        End If


        If gl_FlowForLoop_DigSrc_SweepCode <> "" Then
            TNameSeg(7) = gl_FlowForLoop_DigSrc_SweepCode
            TNameSeg(7) = gl_FlowForLoop_DigSrc_SweepCode_Dec '20190613 CT add for Decimal value printing
        End If
        
        TNameSeg(9) = CStr(TestSeqNum)
        TNameSeg(5) = Replace(PinName, "_", "")


        If ForceResult = tlForceFlow Then
            Dim TestLimitIndex As Long
                        Dim Tname_LimitIndex As String
                        Dim TName_Ary() As String
            If Instance_Data.Is_PreCheck_Func = True Then
                If UBound(Instance_Data.Tname) < TheExec.Flow.TestLimitIndex Then
                    TestLimitIndex = 0
                    TheExec.Datalog.WriteComment "Error : Test Limit Index More then Flow Table in " & instanceName
                Else
                    TestLimitIndex = TheExec.Flow.TestLimitIndex
                End If
                ''------> Start - Flow's Tname has special char "__" - Carter, 20190627
                                Tname_LimitIndex = Instance_Data.Tname(TestLimitIndex)
                If InStr(Tname_LimitIndex, "__") <> 0 Then
                    TName_Ary = Split(Tname_LimitIndex, "__")
                    TNameSeg(7) = TName_Ary(0)
                    Tname_LimitIndex = TName_Ary(1)
                End If
                ''------> End - Flow's Tname has special char "__" - Carter, 20190627
                If Tname <> "" Then
                    TNameSeg(6) = Replace(Tname_LimitIndex, "_", "-") & Replace(Tname, "_", "-")
                Else
                    TNameSeg(6) = Replace(Tname_LimitIndex, "_", "-")
                End If
            Else
                If UBound(gl_Tname_Meas_FromFlow) < TheExec.Flow.TestLimitIndex Then
                    TestLimitIndex = 0
                    TheExec.Datalog.WriteComment "Error : Test Limit Index More then Flow Table in " & instanceName
                Else
                    TestLimitIndex = TheExec.Flow.TestLimitIndex
                End If
                ''------> Start - Flow's Tname has special char "__" - Carter, 20190701
                                Tname_LimitIndex = gl_Tname_Meas_FromFlow(TestLimitIndex)
                If InStr(Tname_LimitIndex, "__") <> 0 Then
                    TName_Ary = Split(Tname_LimitIndex, "__")
                    TNameSeg(7) = TName_Ary(0)
                    Tname_LimitIndex = TName_Ary(1)
                End If
                ''------> End - Flow's Tname has special char "__" - Carter, 20190701
                If Tname <> "" Then
                    TNameSeg(6) = Replace(Tname_LimitIndex, "_", "-") & Replace(Tname, "_", "-") ''Carter, 20190523
                Else
                    TNameSeg(6) = Replace(Tname_LimitIndex, "_", "-") ''Carter, 20190523
'                    TNameSeg(6) = gl_Tname_Meas_FromFlow(TestLimitIndex)
                End If
            End If

        Else
            TNameSeg(6) = Tname
        End If
        
         TNameSeg(7) = Replace(TNameSeg(7), "_", "-") ''------> TNameSeg(7) has special char "_" - Carter, 20190701
        
        
        If TNameSeg(6) = "" Then TNameSeg(6) = "X"
        
        If (InStr(Tname, ",") <> 0) Then
            SubInstNameSegs = Split(Tname, ",")
            If (UBound(SubInstNameSegs) < k) Then
                TNameSeg(6) = SubInstNameSegs(UBound(SubInstNameSegs))
            Else
                TNameSeg(6) = SubInstNameSegs(k)
            End If
        End If

        If LCase(MeasType) = "calc" Then
                TNameSeg(1) = "Calc"
        Else
                TNameSeg(1) = TNameSeg(1) & MeasType
        End If

        If (Sweep_Name <> "") Then TNameSeg(8) = Replace(Sweep_Name, ".", "p")
        If gl_Sweep_Name <> "" Then

            If instanceName Like "*MTRGR*" Then
            
                For Each site In TheExec.sites
                    TNameSeg(8) = CStr(Replace(Format(TheExec.specs.DC.Item(gl_Sweep_Name).CurrentValue, "0.000"), ".", "p"))
                        If (InStr(LCase(TNameSeg(8)), "-") <> 0) Then
                            TNameSeg(8) = Replace(TNameSeg(8), "-", "N")
                        End If
                
                    Exit For
                Next
            
                GoTo skiptohere
            End If

            If (InStr(LCase(gl_Sweep_Name), "vdd_") <> 0) Then
                Dim sweep_name_temp As String
                Dim PinType As String
                Dim power_value As String

                PinType = GetInstrument(gl_Sweep_Name, 0)

                Select Case PinType
                    Case "DC-07"

                        TNameSeg(8) = Replace(gl_Sweep_Name, "_", "") & Replace(Format(TheHdw.DCVI.Pins(gl_Sweep_Name).Voltage, "0.000"), ".", "p")

                    Case "VHDVS"
                        TNameSeg(8) = Replace(gl_Sweep_Name, "_", "") & Replace(Format(TheHdw.DCVS.Pins(gl_Sweep_Name).Voltage.Value, "0.000"), ".", "p")

                    Case "HexVS"
                        TNameSeg(8) = Replace(gl_Sweep_Name, "_", "") & Replace(Format(TheHdw.DCVS.Pins(gl_Sweep_Name).Voltage.Main.Value, "0.000"), ".", "p")

                    Case Else

                End Select
                
            Else
                If (TNameSeg(8) <> "") Then
                    TNameSeg(8) = TheExec.Flow.var(gl_Sweep_Name).Value
                Else
                    TNameSeg(8) = TNameSeg(8) & "&" & CStr(TheExec.Flow.var(gl_Sweep_Name).Value)
                End If
                
            End If
        End If
skiptohere:
        TNameSeg(6) = Replace(TNameSeg(6), " ", "")
      
        If LCase(TNameSeg(4)) = "pp" Or LCase(TNameSeg(4)) = "dd" Or LCase(TNameSeg(4)) = "dp" Or LCase(TNameSeg(4)) = "cz" Or LCase(TNameSeg(4)) = "ht" Then TNameSeg(4) = "X"
        If LCase(TNameSeg(3)) = "pp" Or LCase(TNameSeg(3)) = "dd" Or LCase(TNameSeg(3)) = "dp" Or LCase(TNameSeg(3)) = "cz" Or LCase(TNameSeg(3)) = "ht" Then TNameSeg(3) = "X"

''' *********   {20190523 mask special case Spliting Naming rule in Central VBT for turks}  ******
'''
'''        If InStr(LCase(InstanceName), "lapll") <> 0 Or InStr(LCase(InstanceName), "usb2") <> 0 Or InStr(LCase(InstanceName), "mipi") <> 0 Then
'''            If InStr(TNameSeg(3), "-") = 0 Then
'''            TNameSeg(3) = UCase(TNameSeg(3))
'''                If InStr(LCase(TNameSeg(3)), "v") <> 0 Then    'MPCLV2T2
'''                TNameSeg(7) = Split(TNameSeg(3), "V")(0)
'''                TNameSeg(3) = "V" & Split(TNameSeg(3), "V")(1)
'''                ElseIf InStr(LCase(TNameSeg(3)), "t") <> 0 Then   ' DDRPLLT2P3   ' C00T9P1
'''                TNameSeg(7) = Split(TNameSeg(3), "T")(0)
'''                TNameSeg(3) = "T" & Split(TNameSeg(3), "T")(1)
'''            End If
'''        End If
'''        End If
'''
'''        If InStr(LCase(InstanceName), "lpdprx") <> 0 And InStr(TNameSeg(3), "-") = 0 Then
'''            TNameSeg(3) = UCase(TNameSeg(3))
'''            TNameSeg(6) = UCase(TNameSeg(6))
'''
'''            If LCase(TNameSeg(3)) Like "rx2*" And InStr(TNameSeg(3), "L") <> 0 Then
'''                TNameSeg(7) = "L" & Split(TNameSeg(3), "L")(1)
'''
'''                If UCase(TNameSeg(6)) Like "LN*" Then
'''                    TNameSeg(6) = Replace(UCase(TNameSeg(6)), "LN" & Split(TNameSeg(3), "L")(1), "")
'''            End If
'''
'''                TNameSeg(3) = Split(TNameSeg(3), "L")(0)
'''            End If
'''        End If
'''        If InStr(LCase(InstanceName), "pcie") <> 0 Then
'''
'''                If UCase(TNameSeg(6)) Like "LN*" Then
'''                    TNameSeg(6) = UCase(Replace(TNameSeg(6), "_", ""))
'''                    TNameSeg(7) = UCase(Mid(TNameSeg(6), 1, 3))
'''                    TNameSeg(6) = UCase(Mid(TNameSeg(6), 4, Len(TNameSeg(6)) - 3))
'''                End If
'''
'''        End If
'''
'''        If InStr(LCase(InstanceName), "amp") <> 0 And LCase(TNameSeg(6)) Like "ddr*" And InStr(TNameSeg(3), "-") = 0 Then
'''            TNameSeg(3) = UCase(TNameSeg(3))
'''
'''                TNameSeg(6) = Replace(TNameSeg(6), "_", "")
'''                TNameSeg(7) = UCase(Mid(TNameSeg(6), 1, 4))
'''                TNameSeg(6) = UCase(Mid(TNameSeg(6), 5, Len(TNameSeg(6)) - 4))
'''            End If
'''
'''        If InstNameSegs(0) = "PCIEREFBUF" Then
'''            If Right(TNameSeg(5), 1) = "p" Or Right(TNameSeg(5), 1) = "n" Then
'''            TNameSeg(6) = Tname & Right(TNameSeg(5), 1)
'''            End If
'''        End If
        If TNameSeg(5) = "" Then: TNameSeg(5) = "X"
'''        '-------------------------------Pin Split--------------------------------------------------------
'''        If InStr(LCase(InstanceName), "amp") <> 0 And LCase(TNameSeg(5)) Like "ddr*" Then
'''                TNameSeg(5) = Replace(TNameSeg(5), "_", "")
'''                TNameSeg(7) = UCase(Mid(TNameSeg(5), 1, 4))
'''                TNameSeg(5) = UCase(Mid(TNameSeg(5), 5, Len(TNameSeg(5)) - 4))
'''            End If
'''    '-------------------------------Pin Split--------------------------------------------------------
    '--------------SpecifyTname-------------------------------------------------------------------------------------
    'Ex:1=block;2=subblock
    If SpecifyTname <> "" Then
        If InStr(LCase(SpecifyTname), "replace;") <> 0 Then
            SpecifyTname = Replace(SpecifyTname, "replace;", "")
            tempAry = Split(SpecifyTname, ";")
            For i = 0 To UBound(tempAry)
                If InStr(tempAry(i), "=") <> 0 Then
                    TNameSeg(CInt(Split(tempAry(i), "=")(0))) = Split(tempAry(i), "=")(1)
                End If
            Next i
        
        Else
            tempAry = Split(SpecifyTname, ";")
            For i = 0 To UBound(tempAry)
                If InStr(tempAry(i), "=") <> 0 Then
                    TNameSeg(CInt(Split(tempAry(i), "=")(0))) = TNameSeg(CInt(Split(tempAry(i), "=")(0))) & Split(tempAry(i), "=")(1)
                End If
            Next i
        End If
    End If
    '--------------SpecifyTname-------------------------------------------------------------------------------------
    For i = 0 To UBound(TNameSeg) ''-- Fix Blank Tname Issue
        If (TNameSeg(i)) = "" Then TNameSeg(i) = "X"
    Next i    ''-- Fix Blank Tname Issue (END)
    Report_TName_From_Instance = Join(TNameSeg, "_")

    Call SetupDatalogFormat(80, 100)

End Function

Public Function GetFlowTName() As Double
'---------20180523-------------
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Dim testName() As String
    Dim i As Integer
    
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
       
    If FlowLimitsInfo Is Nothing Then
         ReDim testName(0) As String
    Else
        Call FlowLimitsInfo.GetTNames(gl_Tname_Meas_FromFlow)
    End If
End Function

Public Function SetupDatalogFormat(TestNameW As Integer, PatternW As Integer)
    'Init_Datalog_Setup
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SetupDatalogFormat_Test"
    
    If TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width < TestNameW Then
        TheExec.Datalog.Setup.Shared.Ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Parametric.testName.Width = TestNameW    '70
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.testName.Width = TestNameW    '70
        TheExec.Datalog.Setup.Shared.Ascii.Columns.Functional.Pattern.Width = PatternW      '102
        TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
        
    End If
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function HardIP_SetupAndMeasureCurrent_PPMU(SampleSize As Long, ByRef measureCurrent As PinListData)
    
    Dim i As Long
'    Dim Pins() As Variant
    Dim MeasI As Meas_Type
    
    Dim b_ForceDiffVolt As Boolean
    Dim PastVal As Double
    b_ForceDiffVolt = False
    Dim Pins() As String
    Dim Pin_Cnt As Long
    Dim var As Variant
    
    
    
    MeasI = TestConditionSeqData(Instance_Data.TestSeqNum).MeasI(Instance_Data.TestSeqSweepNum)
    TheHdw.Digital.Pins(MeasI.Pins.PPMU).Disconnect
    If MeasI.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.PPMU.Pins(MeasI.Pins.PPMU)
            .Gate = tlOff
            .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
            .ForceV CDbl(MeasI.Setup_ByType.PPMU.ForceValue1), CDbl(MeasI.Setup_ByType.PPMU.Meas_Range)
            .Connect
            .Gate = tlOn
        End With
        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasI.Pins.PPMU & " =" & TheHdw.PPMU.Pins(SplitInputCondition(MeasI.Pins.PPMU, ",", 0)).MeasureCurrentRange)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasI.Pins.PPMU & " =" & MeasI.Setup_ByType.PPMU.Meas_Range)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasI.Pins.PPMU & " =" & MeasI.WaitTime.PPMU)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasI.Pins.PPMU & " =" & MeasI.Setup_ByType.PPMU.ForceValue1)
        End If
    Else
        For i = 0 To UBound(MeasI.Setup_ByTypeByPin.PPMU)
            With TheHdw.PPMU.Pins(MeasI.Setup_ByTypeByPin.PPMU(i).Pin)
                .Gate = tlOff
                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                .ForceV CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1), CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).Meas_Range)
                .Connect
                .Gate = tlOn
            End With
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasI.Setup_ByTypeByPin.PPMU(i).Pin & " =" & TheHdw.PPMU.Pins(MeasI.Setup_ByTypeByPin.PPMU(i).Pin).MeasureCurrentRange)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasI.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasI.Setup_ByTypeByPin.PPMU(i).Meas_Range)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasI.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasI.WaitTime.PPMU)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasI.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1)
            End If
        Next i
        
        For i = 0 To UBound(MeasI.Setup_ByTypeByPin.PPMU)
            If i <> 0 Then
                If MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1 <> PastVal Then
                    b_ForceDiffVolt = True
                    Exit For
    End If
            End If
        
            PastVal = MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1
        Next i
        
    End If
    TheHdw.Wait CDbl(MeasI.WaitTime.PPMU)
    DebugPrintFunc_PPMU CStr(MeasI.Pins.PPMU)
    measureCurrent = TheHdw.PPMU.Pins(MeasI.Pins.PPMU).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
       
    If gl_Disable_HIP_debug_log = False Then Call ForceVal_Compare(MeasI, MeasI.Pins.PPMU, "I", "PPMU")

    With TheHdw.PPMU.Pins(MeasI.Pins.PPMU)
            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range ''FVMI - Carter, 20190503
            .Disconnect
            .Gate = tlOff
    End With
    TheHdw.Digital.Pins(MeasI.Pins.PPMU).Connect ''Connect Digital pins after measurement - Carter, 20190503



''''20190527  CT add for GPIO ForceV then MeasI(de-embedded Trace effect)
If Not (UCase(TheExec.DataManager.instanceName) Like "*FAILSAFE*") Then
    If UCase(TheExec.DataManager.instanceName) Like "*GPIO*" Then
    
        Dim Pin As Variant
        Dim GetRakVal_PinList As New PinListData
        Dim DiffVolt_Pinlist As New PinListData
        Dim Vdiff As Double
        If Instance_Data.RAK_Flag = R_TraceOnly Then
        
            If MeasI.Setup_ByTypeByPin_Flag = False Then
                If CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) > 0.8 Then
                    'If LCase(Pin) Like "*1p2*" Then
                    '    Vdiff = TheHdw.DCVS.Pins(split(CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) ''vddio12_grp!!!
                    'Else
                        Vdiff = TheHdw.DCVS.Pins(Split(Instance_Data.CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) ''vddio18_grp!!!
                    'End If
                Else
                    Vdiff = CDbl(MeasI.Setup_ByType.PPMU.ForceValue1)
                End If
                
                For Each Pin In measureCurrent.Pins
                
                    If gl_Disable_HIP_debug_log = False Then
                        For Each site In TheExec.sites.Active
                             TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Current = " & measureCurrent.Pins(Pin).Value(site) & ", RAK val = " & CurrentJob_Card_RAK.Pins(Pin).Value(site)
                        Next site
                    End If
                
                    measureCurrent.Pins(Pin) = measureCurrent.Pins(Pin).Multiply(CurrentJob_Card_RAK.Pins(Pin)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(Pin))
                                        
                Next Pin
            Else
                If b_ForceDiffVolt = False Then
                    If CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1) > 0.8 Then
                        'If LCase(Pin) Like "*1p2*" Then
                        '    Vdiff = TheHdw.DCVS.Pins(split(CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio12_grp!!!
                        'Else
                            Vdiff = TheHdw.DCVS.Pins(Split(Instance_Data.CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1) ''vddio18_grp!!!
                        'End If
                    Else
                        Vdiff = CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1)
                    End If
                    
                    For Each Pin In measureCurrent.Pins
                    
                        If gl_Disable_HIP_debug_log = False Then
                            For Each site In TheExec.sites.Active
                                 TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Current = " & measureCurrent.Pins(Pin).Value(site) & ", RAK val = " & CurrentJob_Card_RAK.Pins(Pin).Value(site)
                            Next site
                        End If
                    
                        measureCurrent.Pins(Pin) = measureCurrent.Pins(Pin).Multiply(CurrentJob_Card_RAK.Pins(Pin)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(Pin))
                        
                    Next Pin
                Else
                    For i = 0 To UBound(MeasI.Setup_ByTypeByPin.PPMU)
                        If CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) > 0.8 Then
                            'If LCase(Pin) Like "*1p2*" Then
                            '    Vdiff = TheHdw.DCVS.Pins(split(CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio12_grp!!!
                            'Else
                                Vdiff = TheHdw.DCVS.Pins(Split(Instance_Data.CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio18_grp!!!
                            'End If
                        Else
                            Vdiff = CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1)
                        End If
                        
'                        Dim Pins() As String
'                        Dim Pin_Cnt As Long
'                        Dim var As Variant
                        TheExec.DataManager.DecomposePinList MeasI.Setup_ByTypeByPin.PPMU(i).Pin, Pins, Pin_Cnt
                        For Each var In Pins
                        
                            If gl_Disable_HIP_debug_log = False Then
                                For Each site In TheExec.sites.Active
                                     TheExec.Datalog.WriteComment "Site[" & site & "]," & var & " Current = " & measureCurrent.Pins(var).Value(site) & ", RAK val = " & CurrentJob_Card_RAK.Pins(var).Value(site)
                                Next site
                            End If
                        
                            measureCurrent.Pins(var) = measureCurrent.Pins(var).Multiply(CurrentJob_Card_RAK.Pins(var)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(var))
                            
                        Next var
                    Next i
                
                End If
            End If
            
        ElseIf Instance_Data.RAK_Flag = R_PathWithContact Then
        
            If MeasI.Setup_ByTypeByPin_Flag = False Then
                If CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) > 0.8 Then
                    'If LCase(Pin) Like "*1p2*" Then
                    '    Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) ''vddio12_grp!!!
                    'Else
                        Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) ''vddio18_grp!!!
                    'End If
                Else
                    Vdiff = CDbl(MeasI.Setup_ByType.PPMU.ForceValue1)
                End If
                For Each Pin In measureCurrent.Pins
                
                    If gl_Disable_HIP_debug_log = False Then
                        For Each site In TheExec.sites
                            TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Current = " & measureCurrent.Pins(Pin).Value(site) & ", R_Path val = " & R_Path_PLD.Pins(Pin).Value(site)
                        Next site
                    End If
                
                    measureCurrent.Pins(Pin) = measureCurrent.Pins(Pin).Multiply(R_Path_PLD.Pins(Pin)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(Pin))
                Next Pin
                                        
            Else
                If b_ForceDiffVolt = False Then
                    If CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1) > 0.8 Then
                        'If LCase(Pin) Like "*1p2*" Then
                        '    Vdiff = TheHdw.DCVS.Pins(split(CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio12_grp!!!
                        'Else
                            Vdiff = TheHdw.DCVS.Pins(Split(Instance_Data.CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1) ''vddio18_grp!!!
                        'End If
                    Else
                        Vdiff = CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1)
                    End If
                    For Each Pin In measureCurrent.Pins
                    
                        If gl_Disable_HIP_debug_log = False Then
                            For Each site In TheExec.sites
                                TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Current = " & measureCurrent.Pins(Pin).Value(site) & ", R_Path val = " & R_Path_PLD.Pins(Pin).Value(site)
                            Next site
                        End If
                    
                        measureCurrent.Pins(Pin) = measureCurrent.Pins(Pin).Multiply(R_Path_PLD.Pins(Pin)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(Pin))
                    Next Pin
                Else
                    For i = 0 To UBound(MeasI.Setup_ByTypeByPin.PPMU)
                        If CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) > 0.8 Then
                            'If LCase(Pin) Like "*1p2*" Then
                            '    Vdiff = TheHdw.DCVS.Pins(split(CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio12_grp!!!
                            'Else
                                Vdiff = TheHdw.DCVS.Pins(Split(Instance_Data.CUS_Str_MainProgram)(1)).Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio18_grp!!!
                            'End If
                        Else
                            Vdiff = CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1)
                        End If
'                        Dim Pins() As String
'                        Dim Pin_Cnt As Long
'                        Dim var As Variant
                        TheExec.DataManager.DecomposePinList MeasI.Setup_ByTypeByPin.PPMU(i).Pin, Pins, Pin_Cnt
                        For Each var In Pins
                        
                            If gl_Disable_HIP_debug_log = False Then
                                For Each site In TheExec.sites
                                    TheExec.Datalog.WriteComment "Site[" & site & "]," & var & " Current = " & measureCurrent.Pins(var).Value(site) & ", R_Path val = " & R_Path_PLD.Pins(var).Value(site)
                                Next site
                            End If
                        
                            measureCurrent.Pins(var) = measureCurrent.Pins(var).Multiply(R_Path_PLD.Pins(var)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(var))
                        Next var
                    Next i
                End If
            End If
        
        End If


    End If

End If


''    '' 20150623 - Suggest use for single pin it can mapping expected current range,
''    ''                  if use for pin group it will refer the same current range to pin group by your specified.
''    Dim i As Integer
''    Dim WaitTime As Double
''    Dim MaxWaitTime As Double
''
''    WaitTime = pc_Def_VFI_MI_WaitTime
''    MaxWaitTime = 0
''
''    Dim Pins() As String
''    Dim NumberPins As Long
''    Dim NumTypes As Long
''    Dim PowerType() As String
''    Dim Factor As Long
''    Dim Pins_MeasureI_Together As String
''
''    '' 20160419 - Debug Alarm off
'''    If TheExec.EnableWord("HardIP_Alarm_off") = True Then
'''        thehdw.DCVI.Pins("analogmux_out").Alarm(tlDCVIAlarmAll) = tlAlarmOff
'''    End If
''    '' 20150616 - Findout the expected range and wait time also think for merge Mode
''
''
''    On Error GoTo err
''
''
''    For i = 0 To UBound(MI_TestCond_PPMU)
''        MaxWaitTime = WaitTime
''
''        Call TheExec.DataManager.DecomposePinList(MI_TestCond_PPMU(i).pinName, Pins(), NumberPins)
''        Call TheExec.DataManager.GetChannelTypes(Pins(0), NumTypes, PowerType())
''
''
''
''        If MI_TestCond_PPMU(i).CurrentRange > 0.002 Then
''            MI_TestCond_PPMU(i).CurrentRange = 0.05
''        ElseIf MI_TestCond_PPMU(i).CurrentRange > 0.0002 Then
''            MI_TestCond_PPMU(i).CurrentRange = 0.002
''        ElseIf MI_TestCond_PPMU(i).CurrentRange > 0.00002 Then
''            MI_TestCond_PPMU(i).CurrentRange = 0.0002
''        ElseIf MI_TestCond_PPMU(i).CurrentRange > 0.000002 Then
''            MI_TestCond_PPMU(i).CurrentRange = 0.00002
''        Else
''            MI_TestCond_PPMU(i).CurrentRange = 0.000002
''        End If
''
''        If WaitTime > MaxWaitTime Then
''            MaxWaitTime = WaitTime
''        End If
''
''        TheHdw.Digital.Pins(MI_TestCond_PPMU(i).pinName).Disconnect
''        With TheHdw.PPMU.Pins(MI_TestCond_PPMU(i).pinName)
''            .ForceI pc_Def_PPMU_InitialValue_FI, MI_TestCond_PPMU(i).CurrentRange
''            .ForceV MI_TestCond_PPMU(i).FV_Val, MI_TestCond_PPMU(i).CurrentRange
''            .Connect
''            .Gate = tlOn
''        End With
''
''        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Meter I range setting, " & MI_TestCond_PPMU(i).pinName & " =" & TheHdw.PPMU.Pins(MI_TestCond_PPMU(i).pinName).MeasureCurrentRange.Value)
''        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Force Volt value, " & MI_TestCond_PPMU(i).pinName & " =" & Format(TheHdw.PPMU.Pins(MI_TestCond_PPMU(i).pinName).Voltage, "0.000"))
''
''        If i = 0 Then
''            Pins_MeasureI_Together = MI_TestCond_PPMU(i).pinName
''        Else
''            Pins_MeasureI_Together = Pins_MeasureI_Together & "," & MI_TestCond_PPMU(i).pinName
''        End If
''    Next i
''    '' 20150623 - Convert customize wait time type string to double if MeasCurrWaitTime specified
''    If CustomizeWaitTime <> "" Then
''        MaxWaitTime = CDbl(CustomizeWaitTime)
''    End If
''    TheHdw.Wait (MaxWaitTime)
''
''    '' 20150615 - Current measurement
''
''    MeasureCurrent = TheHdw.PPMU.Pins(Pins_MeasureI_Together).Read(tlPPMUReadMeasurements)
''
''
''    With TheHdw.PPMU.Pins(Pins_MeasureI_Together)
''            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
''            .Disconnect
''            .Gate = tlOff
''    End With


Exit Function

err:

If AbortTest Then Exit Function Else Resume Next

End Function


Public Function ProcessTestNameInputString(ByRef OutputTname() As String, TS_count As Long) As Long
    On Error GoTo err

'Roger add ========================================================================================
    
    Dim OutputTnameBy_Comma() As String
    Dim i As Integer
    
    
    OutputTname = Split(gl_Tname_Meas, "+")
    
    If UBound(OutputTname) < TS_count Then
        ReDim Preserve OutputTname(TS_count) As String
        For i = 0 To UBound(OutputTname)
            If OutputTname(i) = "" Then OutputTname(i) = "X"
        Next i
    End If
        
'=================================================================================================

    Exit Function
err:

If AbortTest Then Exit Function Else Resume Next
End Function

Public Function HIP_Evaluate_ForceVal_New(ByRef ForceVSequenceArray As String) As Long
    Dim EvalIndex As Integer
    Dim TempForceSeq As String
    Dim SplitArray() As String
    Dim i As Long
   
    Do
       
        If InStr(ForceVSequenceArray, "_") > 0 Or InStr(ForceVSequenceArray, "+") > 0 Or InStr(ForceVSequenceArray, "-") > 0 Or InStr(ForceVSequenceArray, "*") > 0 Or InStr(ForceVSequenceArray, "/") > 0 Then ' can not evaluate if only with  single number
            TempForceSeq = ForceVSequenceArray
            TempForceSeq = Replace(TempForceSeq, "|", "~")
            TempForceSeq = Replace(TempForceSeq, ",", "~")
            TempForceSeq = Replace(TempForceSeq, ":", "~")
            TempForceSeq = Replace(TempForceSeq, "&", "~")
            SplitArray = Split(TempForceSeq, "~")
            
            'Fix error by Kaino 2019/05/21
            For i = 0 To UBound(SplitArray)
                'Fix error by Kaino 2019/05/21
                If SplitArray(i) <> "" Then
                'If InStr(SplitArray(i), "_") <> 0 Then
                    ForceVSequenceArray = Replace(ForceVSequenceArray, SplitArray(i), ProcessEvaluateDCSpec(SplitArray(i)))
                                        
'                    Debug.Print ForceVSequenceArray(EvalIndex)
                End If
            Next i
        End If
        
    Loop While (ForceVSequenceArray Like "*#[*/+-][.0123456789]*")  'Fix bug by Kaino 2019-05-22
    
End Function

Public Function GetHiLimitFromFlow() As Double

    Dim FlowLimitsInfo As IFlowLimitsInfo
    Dim hi_limit() As String
    
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    
    If FlowLimitsInfo Is Nothing Then
        TheExec.AddOutput "Could not get the limits info", vbRed, True
        Exit Function
    End If
    
    Call FlowLimitsInfo.GetHighLimits(hi_limit)
    If hi_limit(TheExec.Flow.TestLimitIndex - 1) = "" Then
        GetHiLimitFromFlow = 0
    Else
        GetHiLimitFromFlow = CDbl(hi_limit(TheExec.Flow.TestLimitIndex - 1))
    End If

End Function
Public Function GetLowLimitFromFlow() As Double

    Dim FlowLimitsInfo As IFlowLimitsInfo
    Dim Low_limit() As String
    
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    
    If FlowLimitsInfo Is Nothing Then
        TheExec.AddOutput "Could not get the limits info", vbRed, True
        Exit Function
    End If
    
    Call FlowLimitsInfo.GetLowLimits(Low_limit)
    If Low_limit(TheExec.Flow.TestLimitIndex - 1) = "" Then
        GetLowLimitFromFlow = 0
    Else
        GetLowLimitFromFlow = CDbl(Low_limit(TheExec.Flow.TestLimitIndex - 1))
    End If

End Function


Public Function ProsscessTestLimit(result As Object, MeasType As String, TestSeqNum As Integer, Optional SpecifyTestName As String = "", Optional TestSeqSweepNum As Integer, Optional SpecScaleType As tlScaleType = scaleNone, Optional LimitForceMode As tlLimitForceResults = tlForceFlow) As Long
    Dim p As Long

    Dim Temp_index As Long
    Dim TestNameInput As String
    Dim Measure As Meas_Type
    Dim Unit As UnitType
    Dim ForceUnit As UnitType
    Dim CusUnit As String
    
    Select Case UCase(MeasType)
        Case "V":
            Measure = TestConditionSeqData(Instance_Data.TestSeqNum).MeasV(Instance_Data.TestSeqSweepNum)
            Unit = unitVolt
            ForceUnit = unitAmp
            CusUnit = ""
        Case "I":
            Measure = TestConditionSeqData(Instance_Data.TestSeqNum).MeasI(Instance_Data.TestSeqSweepNum)
            Unit = unitAmp
            ForceUnit = unitVolt
            CusUnit = ""
        Case "F":
            Unit = unitHz
            ForceUnit = unitNone
        Case "R":
            Measure = TestConditionSeqData(Instance_Data.TestSeqNum).MeasR(Instance_Data.TestSeqSweepNum)
            Unit = unitCustom
            ForceUnit = unitVolt
            CusUnit = "ohm"
        Case "Z":
            Unit = unitCustom
            ForceUnit = unitVolt
            CusUnit = "ohm"
    End Select

    Temp_index = TheExec.Flow.TestLimitIndex
    
    If Instance_Data.SpecialCalcValSetting = CalculateMethodSetup.VIR_DDIO Then Exit Function   '''20190509
   
    If TypeName(result) = "IPinListData" Then
        For p = 0 To result.Pins.Count - 1
            If Instance_Data.Flag_SingleLimit = True Then TheExec.Flow.TestLimitIndex = Temp_index
            TestNameInput = Report_TName_From_Instance(MeasType, UCase(result.Pins(p)), SpecifyTestName, TestSeqNum) '''20190531
            If MeasType = "V" Or MeasType = "I" Or MeasType = "R" Then
                TheExec.Flow.TestLimit result.Pins(p), , , , , SpecScaleType, Unit, Tname:=TestNameInput, _
                                ForceVal:=CDbl(Measure.ForceValueDic_HWCom(UCase(result.Pins(p)))), ForceUnit:=ForceUnit, _
                                customUnit:=CusUnit, ForceResults:=LimitForceMode
            Else
                TheExec.Flow.TestLimit result.Pins(p), , , , , SpecScaleType, Unit, Tname:=TestNameInput, _
                                ForceUnit:=ForceUnit, customUnit:=CusUnit, ForceResults:=LimitForceMode
            End If
        Next p
    ElseIf TypeName(result) = "IDspWave_i" Then
        TestNameInput = Report_TName_From_Instance("C", "", SpecifyTestName, TestSeqNum)
        TheExec.Flow.TestLimit result.Element(0), , , , , , Tname:=TestNameInput, ForceResults:=LimitForceMode
    Else
        TheExec.AddOutput "Unknow Limit Type : " & TypeName(result)
    End If
End Function
Public Function ProcessSweepString(assignment As String, customerstring As String, ByRef sweepsrcarray() As String, ByRef loopmax As Long) As Long
                                 
    Dim SplitByColon() As String
    Dim splitbysemicolon() As String
    Dim SplitByEqual() As String
    Dim splitbycondon() As String
    Dim splitbycondon2() As String
    Dim splitbyand() As String
    Dim srcname() As String
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim loopinit As Long
    'Dim loopmax As Long
    Dim SrcBits() As String
    Dim withandstring As String
    SplitByColon = Split(customerstring, ":")
    splitbysemicolon = Split(SplitByColon(1), ";")
    Dim tempassign As String
    Dim replaceindex As Long
    'Dim sweepsrcarray() As String
    
    ReDim srcname(UBound(splitbysemicolon))
    ReDim SrcBits(UBound(splitbysemicolon))
    loopmax = 1
    loopinit = 0
    
    For i = 0 To UBound(splitbysemicolon)
        If splitbysemicolon(i) Like "*|*" Then
            SplitByEqual = Split(splitbysemicolon(i), "=")
            srcname(i) = SplitByEqual(0)
            SrcBits(i) = SplitByEqual(1)
            loopmax = loopmax * (UBound(Split(Split(SplitByEqual(1), "|")(0), ",")) + 1)
        Else
            SplitByEqual = Split(splitbysemicolon(i), "=")
            srcname(i) = SplitByEqual(0)
            If InStr(SplitByEqual(1), "&") Then
                withandstring = ""
                splitbyand = Split(SplitByEqual(1), "&")
                splitbycondon = Split(splitbyand(0), ",")
                splitbycondon2 = Split(splitbyand(1), ",")
                For j = 0 To UBound(splitbycondon)
                    For k = 0 To UBound(splitbycondon2)
                        withandstring = withandstring & splitbycondon(j) & splitbycondon2(k) & ","
                    Next k
                Next j
                withandstring = Left(withandstring, Len(withandstring) - 1)
                SrcBits(i) = withandstring
            Else
                SrcBits(i) = SplitByEqual(1)
            End If
            loopmax = loopmax * (Len(SrcBits(i)) - Len(Replace(SrcBits(i), ",", "")) + 1)
        End If
    Next i
    ReDim sweepsrcarray(loopmax - 1)
    tempassign = assignment
    Dim assignstring As String
    assignstring = ""
    replaceindex = 0
    Dim SplitByVerticalBar_SrcName() As String
    Dim SplitByVerticalBar_SrcBit() As String
    For i = 0 To UBound(srcname)
        If srcname(i) Like "*|*" Then
            If i = 0 Then   'for DDR
                Dim sweepsrcarray_temp() As String
                ReDim sweepsrcarray_temp(UBound(sweepsrcarray))
                splitbycondon = Split(SrcBits(i), ",")
                SplitByVerticalBar_SrcName = Split(srcname(i), "|")
                SplitByVerticalBar_SrcBit = Split(SrcBits(i), "|")
                   For k = 0 To UBound(SplitByVerticalBar_SrcName)
                    replaceindex = 0
                    For j = 0 To UBound(sweepsrcarray)
                        If k = 0 Then
                            sweepsrcarray(j) = Replace(assignment, SplitByVerticalBar_SrcName(k), Split(SplitByVerticalBar_SrcBit(k), ",")(replaceindex))
                        Else
                            sweepsrcarray(j) = Replace(sweepsrcarray(j), SplitByVerticalBar_SrcName(k), Split(SplitByVerticalBar_SrcBit(k), ",")(replaceindex))
                        End If
                        replaceindex = replaceindex + 1
                    Next j
                Next k
            Else
                SplitByVerticalBar_SrcName = Split(srcname(i), "|")
                SplitByVerticalBar_SrcBit = Split(SrcBits(i), "|")
                For k = 0 To UBound(SplitByVerticalBar_SrcName)
                    replaceindex = 0
                    For j = 0 To UBound(sweepsrcarray)
    '                    If j = 21 Then: Stop
                        sweepsrcarray(j) = Replace(sweepsrcarray(j), SplitByVerticalBar_SrcName(k), Split(SplitByVerticalBar_SrcBit(k), ",")(replaceindex))
                        replaceindex = replaceindex + 1
                        If replaceindex > UBound(Split(SplitByVerticalBar_SrcBit(k), ",")) Then: replaceindex = 0
                    Next j
                Next k
            End If
        Else
            splitbycondon = Split(SrcBits(i), ",")
            replaceindex = 0
            Dim stringadclk As String
            For j = 0 To UBound(sweepsrcarray)
         
                If i = 0 Then
                    If InStr(assignment, "ADCLK__CLKMON_SELECT") <> 0 Then
                        assignment = tempassign
                        Call Dec2Bin_str(j, stringadclk, 3)
                        assignment = Replace(assignment, "adclksel", stringadclk)
                    End If
                   sweepsrcarray(j) = Replace(assignment, srcname(i), splitbycondon(replaceindex))
                    If (j + 1) Mod loopmax / (UBound(splitbycondon) + 1) = 0 Then
                        replaceindex = replaceindex + 1
                    End If
    
                    
                Else
                    If InStr(sweepsrcarray(j), "ADCLK__CLKMON_SELECT") <> 0 Then
                        assignment = tempassign
                        Call Dec2Bin_str(j, stringadclk, 3)
                        assignment = Replace(assignment, "adclksel", stringadclk)
                    End If
                    sweepsrcarray(j) = Replace(sweepsrcarray(j), srcname(i), splitbycondon(replaceindex))
                    replaceindex = replaceindex + 1
                    If replaceindex > UBound(splitbycondon) Then
                        replaceindex = 0
                    End If
                End If
                
            Next j
        End If
    Next i
End Function
Public Function ProcessSweepString_specailcase(assignment As String, customerstring As String, ByRef sweepsrcarray() As String, ByRef loopmax As Long) As Long
                                 
    Dim i As Long
    Dim j As Long
    Dim sdline_sel0 As String
    Dim sdline_sel13 As String
    Dim sdline_sel4158 As String: sdline_sel4158 = String$(155, "0")
    ReDim sweepsrcarray(623) As String

    For i = 0 To loopmax
        If i Mod 8 < 4 Then sdline_sel0 = "0"
        If i Mod 8 >= 4 Then sdline_sel0 = "1"
        If i Mod 4 = 0 Then sdline_sel13 = "111"
        If i Mod 4 = 1 Then sdline_sel13 = "011"
        If i Mod 4 = 2 Then sdline_sel13 = "001"
        If i Mod 4 = 3 Then sdline_sel13 = "000"
        If i Mod 4 = 0 Then sdline_sel4158 = "1" + Left(sdline_sel4158, Len(sdline_sel4158) - 1)
        If i = 0 Then sdline_sel4158 = String$(155, "0")
        sweepsrcarray(i) = "sdline_sel=" + sdline_sel0 + sdline_sel13 + sdline_sel4158
    Next
    
End Function
Public Function DSSC_Search_par_run_LDO(Pat As String, srcPin As PinList, code As SiteLong, MeasPin As PinList, Res As SiteDouble, TrimCodeSize As Long, NumberOfMeasV As Integer, ByRef MeasV_Name_Array() As String, ByRef MeasValue_Array() As SiteDouble, ByRef TrimPoint() As Long, DigSrc_Sample_Size As Long, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName As String, MeasV_WaitTime As String)
    Dim sigName As String, srcWave As New DSPWave, site As Variant ': Site = theexec.sites.SiteNumber
    Dim InDSPwave As New DSPWave
'    Dim MeasV_Name_Array() As String: MeasV_Name_Array = Split(MeasV_Name, "+")
    Dim MeasValue As New SiteDouble
    Dim i As Long, j As Long
    Dim Rtn_MeasVolt As New PinListData
'    ByPassTestLimit = True
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array() As Long: ReDim srcwave_array(TrimCodeSize - 1)
    srcWave.CreateConstant 0, TrimCodeSize, DspLong
    Dim Previous_ByPassTestLimit_Flag As Boolean: Previous_ByPassTestLimit_Flag = ByPassTestLimit
    Dim Previous_Disable_CurrRangeSetting_Print_Flag As Boolean: Previous_Disable_CurrRangeSetting_Print_Flag = glb_Disable_CurrRangeSetting_Print
    

    ByPassTestLimit = True
    glb_Disable_CurrRangeSetting_Print = True
    For Each site In TheExec.sites
        For i = 0 To TrimCodeSize - 1
            If i = 0 Then
                srcwave_array(i) = code And 1
            Else
                srcwave_array(i) = (code And (2 ^ i)) \ (2 ^ i)
            End If
        Next i
    srcWave.Data = srcwave_array
    Next site
    
    Call AddStoredCaptureData(TrimStoreName, srcWave)
    Call GeneralDigSrcSetting_LDO(Pat, srcPin, DigSrc_Sample_Size, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, "", "", InDSPwave, "")
    Dim MinTrimIndex As Long
    For i = 0 To UBound(TrimPoint)
        If TrimPoint(i) = 1 Then
            MinTrimIndex = i
            Exit For
        End If
    Next i
    TheHdw.Patterns(Pat).start
    
    For i = 0 To NumberOfMeasV - 1
        TheHdw.Digital.Patgen.FlagWait cpuA, 0

'        Call HardIP_MeasureVolt(MeasPin, "FFF", NumberOfMeasV, 1, Pat, False, HighLimitVal(0), LowLimitVal(0), FlowTestNme, , "LDO_Trim", , Rtn_MeasVolt, , , MeasV_WaitTime)
        Rtn_MeasVolt = HardIP_MeasureVolt
        Call DebugPrintFunc_PPMU("")

        For Each site In TheExec.sites
            MeasValue = Rtn_MeasVolt.Pins(MeasPin.Value).Value
            MeasValue_Array(i) = Rtn_MeasVolt.Pins(MeasPin.Value).Value

'            If i = 0 And TrimPoint(i) Then
            If i = MinTrimIndex Then
                Res = Rtn_MeasVolt.Pins(MeasPin.Value).Value
            ElseIf MeasValue.compare(LessThan, Res) And TrimPoint(i) Then
                Res = Rtn_MeasVolt.Pins(MeasPin.Value).Value
            End If
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Meas Name : " & MeasV_Name_Array(i) & ", Voltage = " & Rtn_MeasVolt.Pins(MeasPin.Value).Value
        Next site

        TheHdw.Digital.Patgen.Continue 0, cpuA
    Next i
    TheHdw.Digital.Patgen.HaltWait
    
    ByPassTestLimit = Previous_ByPassTestLimit_Flag
    glb_Disable_CurrRangeSetting_Print = Previous_Disable_CurrRangeSetting_Print_Flag
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search_par_run_LDO: " & err.Description)
    DSSC_Search_par_run_LDO = TL_ERROR
End Function
Public Function GeneralDigSrcSetting_LDO(Pat As String, DigSrc_pin As PinList, DigSrc_Sample_Size As Long, DigSrc_DataWidth As Long, _
DigSrc_Equation As String, DigSrc_Assignment As String, _
DigSrc_FlowForLoopIntegerName As String, CUS_Str_DigSrcData As String, ByRef InDSPwave As DSPWave, _
Optional ByRef Rtn_SweepTestName As String) As Long

    Dim b_StoreWholeDigSrc As Boolean, b_DigSrcFromDict As Boolean
    Dim DigSrcWholeDictName As String
    Dim DigSrc_Ary() As Long
    Dim site As Variant ': Site = theexec.sites.SiteNumber
    
    ''20161121- Check DigSrc type is serial or parallel
    Dim InDSPWave_Parallel As New DSPWave
    Dim DigSrcPinAry() As String, NumberPins As Long
    Dim b_SrcTypeIsParallel As Boolean
    Dim NoOfSamples As New SiteLong
    Dim CreateDigSrcDataSize As Long
    Call TheExec.DataManager.DecomposePinList(DigSrc_pin, DigSrcPinAry(), NumberPins)
    If NumberPins > 1 Then
        b_SrcTypeIsParallel = True
        CreateDigSrcDataSize = DigSrc_Sample_Size * NumberPins
    Else
        b_SrcTypeIsParallel = False
        CreateDigSrcDataSize = DigSrc_Sample_Size
    End If

    If DigSrc_Sample_Size <> 0 Then
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======== Setup Dig Src Test Start ========")
        
        b_DigSrcFromDict = Checker_DigSrcFromDict(DigSrc_Assignment)
        If b_DigSrcFromDict Then
            InDSPwave = GetStoredCaptureData(DigSrc_Assignment)
            For Each site In TheExec.sites.Active
                DigSrc_Ary = InDSPwave(site).Data
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("[Site " & site & "]")
                If gl_Disable_HIP_debug_log = False Then Call Printout_DigSrc(DigSrc_Ary, DigSrc_Sample_Size, DigSrc_DataWidth)
            Next site
        Else
           ''20160826 - Check DigSrc_Assignment's content to decide whether to store InDSPWave to the Dictionary
            b_StoreWholeDigSrc = Checker_StoreWholeDigSrcInDict(DigSrc_Assignment, DigSrcWholeDictName)
            
            ''20160805 - Analyze DigSrc_Equation and DigSrc_Assignment inout format whether violate design rule
            Call AnalyzeDigSrcEquationAssignmentContent(DigSrc_Equation, DigSrc_Assignment)
            
            If DigSrc_FlowForLoopIntegerName <> "" Then
                Call DSSCSrcBitFromFlowForLoop(DigSrc_FlowForLoopIntegerName, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, CUS_Str_DigSrcData, Rtn_SweepTestName)
            End If
            ''20170905 Cyprus AMP sweep with two partial src code
             If CUS_Str_DigSrcData <> "" And UCase(CUS_Str_DigSrcData) Like UCase("*VOLH_Sweep*") Then
                Call VOLH_Sweep(CUS_Str_DigSrcData, DigSrc_Assignment)
            End If

            For Each site In TheExec.sites.Active
                Call Create_DigSrc_Data(DigSrc_pin, DigSrc_DataWidth, CreateDigSrcDataSize, DigSrc_Equation, DigSrc_Assignment, InDSPwave, site, , NumberPins)
            Next site
            
            ''20161121-Setup DigSrc by serial or parallel
            If b_SrcTypeIsParallel Then
                rundsp.BitWf2Arry InDSPwave, NumberPins, NoOfSamples, InDSPWave_Parallel
                Call SetupDigSrcDspWave(Pat, DigSrc_pin, "Meas_Src_Parallel", DigSrc_Sample_Size, InDSPWave_Parallel)
            Else
            
                If b_StoreWholeDigSrc Then
                    Call AddStoredCaptureData(DigSrcWholeDictName, InDSPwave)
                End If
                Call SetupDigSrcDspWave(Pat, DigSrc_pin, "Meas_Src", DigSrc_Sample_Size, InDSPwave)
            End If
        End If
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Src Pin =" & DigSrc_pin.Value)
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======== Setup Dig Src Test End   ========")
    End If
End Function
    





Public Function Print_NewFormat(ProcessString As String) As String

    Dim Separator As String
    If InStr(ProcessString, ";") <> 0 Then Separator = ";"
    If InStr(ProcessString, "+") <> 0 Then Separator = "+"
    Dim NewFormatArray() As String
    Dim NewFormat_Print As String
    NewFormat_Print = ""
    Dim i As Long
    NewFormatArray() = Split(ProcessString, Separator)
    For i = 0 To UBound(NewFormatArray)
        If i = 0 Then
            Print_NewFormat = NewFormatArray(0)
        ElseIf (i Mod 5) = 0 Then
            Print_NewFormat = Print_NewFormat & Separator & vbCrLf & NewFormatArray(i)
        Else
            Print_NewFormat = Print_NewFormat & Separator & NewFormatArray(i)
        End If
    Next i
End Function

Public Function ProcessAssignment()

    Dim i As Long
    Dim TestNumber As Long
    Dim RegisterNumber As Long
    Dim RegAssignSheet As Worksheet
    Dim MaxRow As Long
    Dim RegAssignArr() As Variant

    i = 2
    TestNumber = 0
    'RegisterNumber = 0

    'Dim RegDict As Object

    'Set RegDict = CreateObject(Scripting.Dictionary)

    On Error GoTo errHandler

    If RegAssignChecker(MaxRow) = True Then
    
        RegAssignArr() = Worksheets("Reg_Assign").range(Cells(1, 1), Cells(MaxRow + 1, 5)).Value
        
        While (RegAssignArr(i, 3) <> "") Or (RegAssignArr(i, 4) <> "") Or (RegAssignArr(i, 5) <> "")
            If RegAssignArr(i, 1) <> "" Then
                RegisterNumber = 0
                ReDim Preserve RegAssignInfo.ByTest(TestNumber) As ByTest
                ReDim Preserve RegAssignInfo.ByTest(TestNumber).RegAssign(RegisterNumber) As RegAssign
                RegAssignInfo.ByTest(TestNumber).testName = RegAssignArr(i, 1)
                RegAssignInfo.ByTest(TestNumber).RegAssign(RegisterNumber).RegName = RegAssignArr(i, 3)
                RegAssignInfo.ByTest(TestNumber).RegAssign(RegisterNumber).RegAssignByModeA = RegAssignArr(i, 4)
                RegAssignInfo.ByTest(TestNumber).RegAssign(RegisterNumber).RegAssignByModeB = RegAssignArr(i, 5)
                If RegAssignArr(1, 4) <> "" Then
                    RegAssignInfo.ByTest(TestNumber).RtnByModeA = _
                    RegAssignInfo.ByTest(TestNumber).RegAssign(RegisterNumber).RegName & "=" & RegAssignInfo.ByTest(TestNumber).RegAssign(RegisterNumber).RegAssignByModeA
                End If
                If RegAssignArr(1, 5) <> "" Then
                    RegAssignInfo.ByTest(TestNumber).RtnByModeB = _
                    RegAssignInfo.ByTest(TestNumber).RegAssign(RegisterNumber).RegName & "=" & RegAssignInfo.ByTest(TestNumber).RegAssign(RegisterNumber).RegAssignByModeB
                End If
                TestNumber = TestNumber + 1
                RegisterNumber = RegisterNumber + 1
            Else
                ReDim Preserve RegAssignInfo.ByTest(TestNumber - 1).RegAssign(RegisterNumber) As RegAssign
                RegAssignInfo.ByTest(TestNumber - 1).RegAssign(RegisterNumber).RegName = RegAssignArr(i, 3)
                RegAssignInfo.ByTest(TestNumber - 1).RegAssign(RegisterNumber).RegAssignByModeA = RegAssignArr(i, 4)
                RegAssignInfo.ByTest(TestNumber - 1).RegAssign(RegisterNumber).RegAssignByModeB = RegAssignArr(i, 5)
                
                If RegAssignArr(1, 4) <> "" Then
                    RegAssignInfo.ByTest(TestNumber - 1).RtnByModeA = _
                    RegAssignInfo.ByTest(TestNumber - 1).RtnByModeA & ";" & RegAssignInfo.ByTest(TestNumber - 1).RegAssign(RegisterNumber).RegName & "=" & RegAssignInfo.ByTest(TestNumber - 1).RegAssign(RegisterNumber).RegAssignByModeA
                End If
                If RegAssignArr(1, 5) <> "" Then
                    RegAssignInfo.ByTest(TestNumber - 1).RtnByModeB = _
                    RegAssignInfo.ByTest(TestNumber - 1).RtnByModeB & ";" & RegAssignInfo.ByTest(TestNumber - 1).RegAssign(RegisterNumber).RegName & "=" & RegAssignInfo.ByTest(TestNumber - 1).RegAssign(RegisterNumber).RegAssignByModeB
                End If
                RegisterNumber = RegisterNumber + 1
            End If
            i = i + 1
        Wend
        
    Call StoredRegAssign


    Else

        TheExec.Datalog.WriteComment "There is no 'Reg_Assign' Sheet in this workbook"

    End If
    
    Exit Function
errHandler:
    
    TheExec.ErrorLogMessage "Error on Register Assignment sheet processing, Please check the contents"

End Function


Public Function ConcatenateDSP_TTR(ByVal DSPWave_First As DSPWave, ByVal First_StartElement As Long, ByVal First_EndElement As Long, _
                               ByVal DSPWave_Second As DSPWave, ByVal Second_StartElement As Long, ByVal Second_EndElement As Long, _
                               ByRef DSPWave_Combine As DSPWave) As Long

    Dim FinalLength As Long
    Dim i As Long
    FinalLength = Abs(First_EndElement - First_StartElement) + Abs(Second_EndElement - Second_StartElement) + 2
    DSPWave_Combine.CreateConstant 0, FinalLength
    
    Dim b_MinToMax_First As Boolean
    Dim b_MinToMax_Second As Boolean
    Dim Step_First As Integer
    Dim Step_Second As Integer
    Dim counter As Long
    Dim site As Variant
    counter = 0
    If First_EndElement - First_StartElement > 0 Then
        b_MinToMax_First = True
        Step_First = 1
    Else
        b_MinToMax_First = False
        Step_First = -1
    End If
    
    If Second_EndElement - Second_StartElement > 0 Then
        b_MinToMax_Second = True
        Step_Second = 1
    Else
        b_MinToMax_Second = False
        Step_Second = -1
    End If
    
    For Each site In TheExec.sites
        For i = First_StartElement To First_EndElement Step Step_First
            DSPWave_Combine(site).Element(counter) = DSPWave_First(site).Element(i)
            counter = counter + 1
        Next i
        
        For i = Second_StartElement To Second_EndElement Step Step_Second
            DSPWave_Combine(site).Element(counter) = DSPWave_Second(site).Element(i)
            counter = counter + 1
        Next i
        counter = 0
    Next site
End Function


Public Function TrimUVI80_Meas_VFI_ADC(Pat As String, TestSequenceArray() As String, srcPin As PinList, code() As SiteLong, _
MeasV_Pin As PinList, MeasValue As SiteDouble, MeasI_Pin As PinList, MeasureI_Range As Double, _
MeasF_PinS_SingleEnd As PinList, MeasF_Interval As String, MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, _
DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, DigSrc_Equation As String, DigSrc_Assignment As String, _
DigCap_Pin As PinList, DigCap_DataWidth As Long, DigCap_Sample_Size As Long, CUS_Str_DigCapData As String, OutDSP As DSPWave, _
TrimCodeSize As Long, Trimname() As String, Meas_StoreName As String, Cal_Eqn As String, TrimCal_Name As String, CPUA_Flag_In_Pat As Boolean, MeasSeqAry() As SiteDouble, Optional Final_Calc As Boolean, Optional b_Trimfinish As Boolean = False)

    On Error GoTo err1:
    Dim sigName As String, srcWave As New DSPWave, site As Variant
    
    Dim DigSrcCodeSize As Long
    Dim i As Long, j As Long
    Dim code_bin() As String
    Dim Ts As Variant
    Dim Str_FinalPatName As String
    Dim temp_assignment As String
    Dim cal As New SiteDouble
    Dim out_str() As String
    Dim MeasStoreName_Ary() As String
    Dim TestSeqNum As Long
    ReDim code_bin(TheExec.sites.Existing.Count)
    ReDim out_str(TheExec.sites.Existing.Count)
    Dim TrimCal_value As New PinListData
    Dim TrimCalCap_value As New DSPWave
    Dim TrimCal_Name_array() As String
    TrimCal_Name_array = Split(TrimCal_Name, ":")
    
    ''''''''''''''''''''''''''''''''setup store name'''''''''''''''''''''''''''''.
    MeasStoreName_Ary = Split(Meas_StoreName, "+")
    Dim Rtn_Meas As New PinListData
    Dim Store_Rtn_Meas() As New PinListData
    Dim SoreMaxNum As Long
    Dim StoreIndex As Long
    Dim MeasSeqAry_temp() As New SiteDouble
    ReDim MeasSeqAry(0) As SiteDouble
    ReDim MeasSeqAry_temp(0) As New SiteDouble
    ''20170123-Get how many store name in MeasStoreName_Ary
    If Meas_StoreName <> "" Then
        SoreMaxNum = 0
        For i = 0 To UBound(MeasStoreName_Ary)
            If MeasStoreName_Ary(i) <> "" Then
                SoreMaxNum = SoreMaxNum + 1
            End If
        Next i
         ReDim Store_Rtn_Meas(SoreMaxNum - 1) As New PinListData
         StoreIndex = 0
    End If
    
    '''''''''''''''''''''''''''''''''''''Setup SrcCode'''''''''''''''''''''''''''''''''''''''''''
    
    temp_assignment = DigSrc_Assignment
    
    sigName = "DSSC_Search_Code"
    TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals.Add sigName
    srcWave.CreateConstant 0, DigSrc_Sample_Size, DspLong
    For Each site In TheExec.sites
        

        DigSrc_Assignment = temp_assignment
        code_bin(site) = ""
        For j = 0 To UBound(Trimname)
            For i = 0 To TrimCodeSize - 1
                If i = 0 Then
                    code_bin(site) = CStr(code(j)(site) And 1)
                Else
                    code_bin(site) = code_bin(site) & CStr((code(j)(site) And (2 ^ i)) \ (2 ^ i))
                End If
            Next i
       
        DigSrc_Assignment = Replace(DigSrc_Assignment, Trimname(j), code_bin(site))
        Next j
        
        Call Create_DigSrc_Data_Trim(srcPin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, srcWave, site)
        
        If DigSrc_DataWidth = 0 Then
            DigSrc_DataWidth = 4
        End If
        
        out_str(site) = ""
        
        For i = 0 To DigSrc_Sample_Size - 1
        
            If (i Mod DigSrc_DataWidth) = 0 Then
                out_str(site) = out_str(site) & " "
            End If
                out_str(site) = out_str(site) & srcWave.Element(i)

        Next i
        
        
        TheExec.WaveDefinitions.CreateWaveDefinition "WaveDef" & site, srcWave, True
        With TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals(sigName)
            .WaveDefinitionName = "WaveDef" & site
            .SampleSize = DigSrc_Sample_Size
            .Amplitude = 1
            .LoadSamples
            .LoadSettings
        End With

    Next site
    TheHdw.DSSC.Pins(srcPin).Pattern(Pat).Source.Signals.DefaultSignal = sigName
    
 ''''''''''''''''''''''''''''''''''''''''''''''''''''Setup DigCap'''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If DigCap_Sample_Size <> 0 Then
    
        Call AnalyzePatName(Pat, Str_FinalPatName)
        
        '' 20150812-Modify program to process multiply dig cap pins
        With TheHdw.DSSC.Pins(DigCap_Pin).Pattern(Pat).Capture.Signals
            .Add (Str_FinalPatName & DigCap_Sample_Size & "_" & DigCap_Pin)
            With .Item(Str_FinalPatName & DigCap_Sample_Size & "_" & DigCap_Pin)
                .SampleSize = DigCap_Sample_Size    'CaptureCyc * OneCycle
                .LoadSettings
            End With
        End With
        
        'Create capture waveform
        OutDSP = TheHdw.DSSC.Pins(DigCap_Pin).Pattern(Pat).Capture.Signals(Str_FinalPatName & DigCap_Sample_Size & "_" & DigCap_Pin).DSPWave
        
        '' 20150813 - Assign WaveName to the DSPWave to do recognition of post process.
        For Each site In TheExec.sites
            OutDSP(site).Info.WaveName = DigCap_Pin
        Next site
        
        TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
        TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
    End If
    
  
    
        
    TheHdw.Patterns(Pat).start
    
    
    
    'Call DebugPrintFunc_PPMU("")
    Dim MeasValue_Temp As New SiteDouble
    Set MeasValue_Temp = MeasValue_Temp.Add(10000000000000#)
    
    Dim MeasV_Flag As Boolean: MeasV_Flag = False
    
 '''''''''''''''''''''''''''''''''''''''''''''''''''''Start Measure''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    For Each Ts In TestSequenceArray
        If CPUA_Flag_In_Pat = True Then
        TheHdw.Digital.Patgen.FlagWait cpuA, 0
        'thehdw.Wait 10 * ms
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
        Select Case UCase(Ts)
            Case "V"
                
                Call Trim_SetupandmeasureV_UVI80_ADC(MeasV_Pin, MeasValue, code, code_bin, out_str, b_Trimfinish)
                If Meas_StoreName <> "" Then
                    If MeasStoreName_Ary(TestSeqNum) <> "" Then
                        Rtn_Meas.AddPin (MeasV_Pin)
                        For Each site In TheExec.sites
                            Rtn_Meas.Pins(MeasV_Pin).Value(site) = MeasValue(site)
                        Next site
                        Store_Rtn_Meas(StoreIndex) = Rtn_Meas
                        Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                        StoreIndex = StoreIndex + 1
                    End If
                End If
                For Each site In TheExec.sites.Active
                    If (MeasValue_Temp > MeasValue) Then
                        MeasValue_Temp = MeasValue
                    End If
                Next site
                MeasV_Flag = True
                                            
            
            Case "N"
            
        End Select
        If CPUA_Flag_In_Pat = True Then
        TheHdw.Digital.Patgen.Continue 0, cpuA
        
        If UCase(Ts) = "V" Or UCase(Ts) = "I" Or UCase(Ts) = "F" Then
            If TestSeqNum <> 0 Then
                ReDim Preserve MeasSeqAry(UBound(MeasSeqAry) + 1) As SiteDouble
                ReDim Preserve MeasSeqAry_temp(UBound(MeasSeqAry_temp) + 1) As New SiteDouble
            End If
            MeasSeqAry_temp(UBound(MeasSeqAry_temp)) = MeasValue
            Set MeasSeqAry(UBound(MeasSeqAry)) = MeasSeqAry_temp(UBound(MeasSeqAry_temp))
        End If
        TestSeqNum = TestSeqNum + 1
        
        
        End If
    Next Ts
    
    If MeasV_Flag = True Then
        Set MeasValue = MeasValue_Temp
    End If
        
    TheHdw.Digital.Patgen.HaltWait
    
    '''''''''''''''''''''''''''''''''''''''''''''''''Calc Equation at trim step or final ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Final_Calc <> True Then
        If TrimCal_Name <> "" Then
            If Cal_Eqn <> "" Then
                Call ProcessCalcEquation(Cal_Eqn)
                If TrimCal_Name_array(0) = "C" Then
                    TrimCalCap_value = GetStoredCaptureData(TrimCal_Name_array(1))
                    For Each site In TheExec.sites.Active
                        MeasValue(site) = TrimCalCap_value.Element(0)
                        TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", " & TrimCal_Name_array(1) & " =" & MeasValue(site)
                    Next site
                    
                Else
                    TrimCal_value = GetStoredMeasurement(TrimCal_Name)
                    For Each site In TheExec.sites.Active
                        MeasValue(site) = TrimCal_value.Pins(0).Value(site)
                        TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", " & TrimCal_Name & " =" & MeasValue(site)
                    Next site
                End If
            End If
                    
        End If
    End If
      
    Exit Function
err1:
    Debug.Print "error"
    Resume Next
    
End Function


Public Function Trim_SetupandmeasureV_UVI80_ADC(MeasV_Pin As PinList, MeasValue As SiteDouble, code() As SiteLong, code_bin() As String, out_str() As String, b_Trimfinish As Boolean)
    Dim i As Integer
    Dim site As Variant
 '''''''''''''''setup UVI80 for meas V''''''''''''''''''
    With TheHdw.DCVI.Pins(MeasV_Pin)
        .Gate = False
        .Disconnect tlDCVIConnectDefault
        .mode = tlDCVIModeHighImpedance
        .Connect tlDCVIConnectHighSense
        .Voltage = 6
        .current = 0
         'thehdw.Wait 1 * ms
        .Gate = True
    End With
    
    With TheHdw.DCVI.Pins(MeasV_Pin)
        .Meter.mode = tlDCVIMeterVoltage
    End With
    TheHdw.Wait 1 * ms
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim cal As New SiteDouble
    
    ''''''''''''''''''offline simulate''''''''''''''''''''''''''''''
    If TheExec.TesterMode = testModeOffline Then
        'Dim Pin As Variant
        
        For Each site In TheExec.sites.Active
            MeasValue(site) = code(0) * 0.1
        Next site

        If gl_Disable_HIP_debug_log = False Then

            If b_Trimfinish = False Then
                TheExec.Datalog.WriteComment "Trimming"
            Else
                TheExec.Datalog.WriteComment "TrimResult"
            End If
            
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code_bin(site) & ", Src_code = " & out_str(site) & ", Voltage = " & MeasValue(site)
            Next site
        End If
'        End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Else
    
        MeasValue = TheHdw.DCVI.Pins(MeasV_Pin.Value).Meter.Read(tlStrobe, 10)
        
        If gl_Disable_HIP_debug_log = False Then
        For i = 0 To 0
            If b_Trimfinish = False Then
                TheExec.Datalog.WriteComment "Trimming_" & "VoltageSeq_" '& CStr(i + 1)
            Else
                TheExec.Datalog.WriteComment "TrimResult"
            End If
        
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(i)(site) & ", Src_code = " & out_str(site) & ", Voltage = " & MeasValue(site)
            Next site
            Next i
        End If

'        End If
    
    End If

    With TheHdw.DCVI.Pins(MeasV_Pin)
        .Gate(tlDCVIGateHiZ) = False
        .Disconnect
        .mode = tlDCVIModeCurrent
    End With

End Function

Public Function RegAssignChecker(MaxRow As Long) As Boolean

    Dim USsheet() As String

    Dim indx As Long

    RegAssignChecker = False
    
        #If IGXL8p30 Then
        
                Dim RegAssignSheet As Worksheet
                        
                For Each RegAssignSheet In Worksheets
                
                        If LCase(RegAssignSheet.Name) Like "reg_assign" Then
                
                        Worksheets("reg_assign").Activate
                        
                        MaxRow = Worksheets(USsheet(indx)).UsedRange.Rows.Count
                        
                        If MaxRow > 1 Then RegAssignChecker = True
                        
                        End If
                        
                Next RegAssignSheet

        #Else
        
            USsheet() = TheExec.Job.GetSheetNamesOfType(DMGR_SHEET_TYPE_USER)

            For indx = 0 To UBound(USsheet)
        
                If LCase(USsheet(indx)) Like "reg_assign" Then
                    
                    Worksheets("reg_assign").Activate
        
                    MaxRow = Worksheets(USsheet(indx)).UsedRange.Rows.Count
                    
                    If MaxRow > 1 Then RegAssignChecker = True
        
                End If
        
            Next indx
            
        #End If
        
End Function

Public Function ProcessInputToGLB(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean = False, Optional DisableComparePins As PinList = "", Optional DisableConnectPins As PinList = "", Optional _
DisableFRC As Boolean = False, Optional FRCPortName As String, Optional MeasV_Pins As String, Optional MeasF_PinS_SingleEnd As String, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, Optional _
MeasF_Flag_MeasureThreshold As Boolean, Optional MeasF_ThresholdPercentage As Double, Optional MeasF_WaitTime As String, Optional MeasI_pinS As String, Optional MeasI_Range As String, Optional MeasI_WaitTime As String, Optional _
DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional _
DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String, Optional SpecialCalcValSetting As CalculateMethodSetup, Optional InstSpecialSetting As InstrumentSpecialSetup, Optional _
CUS_Str_MainProgram As String, Optional CUS_Str_DigCapData As String, Optional CUS_Str_DigSrcData As String, Optional Flag_SingleLimit As Boolean, Optional TestLimitPerPin_VFI As String, Optional MeasF_PinS_Differential As String, Optional ForceFunctional_Flag As Boolean, Optional MeasF_WalkingStrobe_Flag As Boolean, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, Optional _
MeasF_WalkingStrobe_StepVoltage As Double, Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double, Optional Meas_StoreName As String, Optional _
Calc_Eqn As String, Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, Optional CharSetName As String, Optional ForceV_Val As String, Optional ForceI_Val As String, Optional _
UVI80_MeasV_WaitTime As String, Optional RAK_Flag As Enum_RAK, Optional WaitTime_VIRZ As String) As Long
    
    Dim Instance_Data_temp() As Instance_Type
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    Dim SeqIdx, SweepIdx, k, PinIdx As Long
    Dim Temp_Array() As String
    Dim temp_string As String
    Dim tt As Long
    Dim TempPin As String
    Dim TempArray() As Variant
    
    Dim tempMeas_Pin As String
    Dim tempMeas_Range As String
    Dim tempMeas_ForceValue As String
    Dim tempMeas_WaitTime As String
    
    Dim TestSeqSplitMeasCase() As String
    Dim TestSeqSplitStoreName() As String
    
    Dim TestSeqSplitPin() As String
    Dim TestSeqSplitDiffPin() As String
    Dim TestSeqSplitRange() As String
    Dim TestSeqSplitForceValue() As String
    Dim TestSeqSplitWaitTime() As String
    
    Dim MeasV As Meas_Type
    Dim MeasI As Meas_Type
    Dim MeasR As Meas_Type
    Dim MeasZ As Meas_Type
    Dim MeasVdiff As Meas_Type
    
    
    
    ReDim Instance_Data_temp(0)

    On Error GoTo err:
    
    If gl_GetInstrumentType_Dic.Count < 10 Or gl_GetInstrument_Dic.Count < 10 Then
        Call GetInstTypToDic("All_DCVI,All_DCVI_Analog,All_HexVS,All_UVS256,All_Digital")
    End If

    Instance_Data = Instance_Data_temp(0)
    
    Instance_Data.patset = patset.Value
    Instance_Data.TestSequence = TestSequence
    Instance_Data.CPUA_Flag_In_Pat = CPUA_Flag_In_Pat
    If Not (DisableComparePins Is Nothing) Then Instance_Data.DisableComparePins = DisableComparePins
    If Not (DisableConnectPins Is Nothing) Then Instance_Data.DisableConnectPins = DisableConnectPins
    Instance_Data.DisableFRC = DisableFRC
    Instance_Data.FRCPortName = FRCPortName
    Instance_Data.MeasV_Pins = CheckInputStringByAt(MeasV_Pins)
    Instance_Data.MeasF_PinS_SingleEnd = CheckInputStringByAt(MeasF_PinS_SingleEnd)
    Instance_Data.MeasF_Interval = MeasF_Interval
    Instance_Data.MeasF_EventSourceWithTerminationMode = MeasF_EventSourceWithTerminationMode
    Instance_Data.MeasF_Flag_MeasureThreshold = MeasF_Flag_MeasureThreshold
    Instance_Data.MeasF_ThresholdPercentage = MeasF_ThresholdPercentage
    Instance_Data.MeasF_WaitTime = MeasF_WaitTime
    Instance_Data.MeasF_PinS_Differential = CheckInputStringByAt(MeasF_PinS_Differential)
    Instance_Data.MeasI_pinS = CheckInputStringByAt(MeasI_pinS)
    Instance_Data.MeasI_Range = CheckInputStringByAt(MeasI_Range)
    Instance_Data.MeasI_WaitTime = MeasI_WaitTime
    If Not (DigCap_Pin Is Nothing) Then Instance_Data.DigCap_Pin = DigCap_Pin
    Instance_Data.DigCap_DataWidth = DigCap_DataWidth
    Instance_Data.DigCap_Sample_Size = DigCap_Sample_Size
    If Not (DigSrc_pin Is Nothing) Then Instance_Data.DigSrc_pin = DigSrc_pin
    Instance_Data.DigSrc_DataWidth = DigSrc_DataWidth
    Instance_Data.DigSrc_Sample_Size = DigSrc_Sample_Size
    Instance_Data.DigSrc_Equation = DigSrc_Equation
    Instance_Data.DigSrc_Assignment = DigSrc_Assignment
    Instance_Data.DigSrc_FlowForLoopIntegerName = DigSrc_FlowForLoopIntegerName
    Instance_Data.SpecialCalcValSetting = SpecialCalcValSetting
    Instance_Data.InstSpecialSetting = InstSpecialSetting
    Instance_Data.CUS_Str_MainProgram = CUS_Str_MainProgram
    Instance_Data.CUS_Str_DigCapData = CUS_Str_DigCapData
    Instance_Data.CUS_Str_DigSrcData = CUS_Str_DigSrcData
    Instance_Data.Flag_SingleLimit = Flag_SingleLimit
    Instance_Data.ForceFunctional_Flag = ForceFunctional_Flag
    Instance_Data.MeasF_WalkingStrobe_Flag = MeasF_WalkingStrobe_Flag
    Instance_Data.MeasF_WalkingStrobe_StartV = MeasF_WalkingStrobe_StartV
    Instance_Data.MeasF_WalkingStrobe_EndV = MeasF_WalkingStrobe_EndV
    Instance_Data.MeasF_WalkingStrobe_StepVoltage = MeasF_WalkingStrobe_StepVoltage
    Instance_Data.MeasF_WalkingStrobe_BothVohVolDiffV = MeasF_WalkingStrobe_BothVohVolDiffV
    Instance_Data.MeasF_WalkingStrobe_interval = MeasF_WalkingStrobe_interval
    Instance_Data.MeasF_WalkingStrobe_miniFreq = MeasF_WalkingStrobe_miniFreq
    Instance_Data.Meas_StoreName = Meas_StoreName
    Instance_Data.Calc_Eqn = Calc_Eqn
    Instance_Data.Interpose_PrePat = Interpose_PrePat
    Instance_Data.Interpose_PreMeas = Interpose_PreMeas
    Instance_Data.Interpose_PostTest = Interpose_PostTest
    Instance_Data.CharSetName = CharSetName
    Instance_Data.ForceV_Val = CheckInputStringByAt(ForceV_Val)
    Instance_Data.ForceI_Val = CheckInputStringByAt(ForceI_Val)
    Instance_Data.RAK_Flag = RAK_Flag
    Instance_Data.WaitTime_VFIRZ = WaitTime_VIRZ
    Instance_Data.Is_PreCheck_Func = True
    Instance_Data.DigSrcCheckCorrect = True
    Instance_Data.MeasV_WaitTime_UVI80 = UVI80_MeasV_WaitTime
    
    If (UCase(Instance_Data.MeasI_Range) Like "*CP*=*" Or UCase(Instance_Data.MeasI_Range) Like "*FT*=*") Then Instance_Data.MeasI_Range = Select_MeasIRange(Instance_Data.MeasI_Range, CurrentJobName_U)
    If (UCase(Instance_Data.MeasV_Pins) Like "*CP*=*" Or UCase(Instance_Data.MeasV_Pins) Like "*FT*=*") Then Instance_Data.MeasV_Pins = Select_MeasIRange(Instance_Data.MeasV_Pins, CurrentJobName_U)
    If (UCase(Instance_Data.MeasI_pinS) Like "*CP*=*" Or UCase(Instance_Data.MeasI_pinS) Like "*FT*=*") Then Instance_Data.MeasI_pinS = Select_MeasIRange(Instance_Data.MeasI_pinS, CurrentJobName_U)
    


    If InStr(LCase(Interpose_PrePat), "sweep:") <> 0 Then
        Call SortSweepInfo(Instance_Data.Sweep_Info, Interpose_PrePat)
        Instance_Data.Sweep_Enable = True
    End If

    If FlowLimitsInfo Is Nothing Then
         ReDim Instance_Data.Tname(0) As String
    Else
        Call FlowLimitsInfo.GetTNames(Instance_Data.Tname)
        Call FlowLimitsInfo.GetLowLimits(Instance_Data.LowLimit)
        Call FlowLimitsInfo.GetHighLimits(Instance_Data.HiLimit)
    End If
    
    If False Then
        If Instance_Data.DigSrc_Equation <> "" Or Instance_Data.DigSrc_Assignment <> "" Then
            Instance_Data.DigSrcCheckCorrect = CheckDigSrcEquationAssignment(Instance_Data.DigSrc_Sample_Size, Instance_Data.DigSrc_DataWidth, Instance_Data.DigSrc_Equation, Instance_Data.DigSrc_Assignment, Instance_Data.DigSrcEquationSampleSize)
            Instance_Data.MergeDigSrcEquation = MergeDigSrcEquationAssignment(DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, Instance_Data.StoreDictName)
        End If
    End If
    

    '=======================================================================================================
    
    
    If ForceFunctional_Flag = True Then
        TestSequence = Replace(TestSequence, "V", "N")
        TestSequence = Replace(TestSequence, "F", "N")
        TestSequence = Replace(TestSequence, "I", "N")
        TestSequence = Replace(TestSequence, "R", "N")
        TestSequence = Replace(TestSequence, "Z", "N")
        
        While (InStr(TestSequence, "NN") <> 0)
            TestSequence = Replace(TestSequence, "NN", "N")
        Wend
    End If
    
    'If theexec.DataManager.InstanceName = "MIPI_T1PRO_PP_CEBA0_V_FULP_AN_MIPI_MEA_JTG_IDS_ALLFV_SI_SHUT_T1PRO_NV" Then Stop
    
    If Instance_Data.ForceI_Val = "" Then Instance_Data.ForceI_Val = "0"
    If Instance_Data.ForceV_Val = "" Then Instance_Data.ForceV_Val = "0"
    If Instance_Data.Meas_StoreName <> "" Then Meas_StoreName_Flag = True
    
    If (TestSequence <> "") Then
        TestSeqSplitMeasCase = SplitInputCondition(Instance_Data.TestSequence, ",")
        If Meas_StoreName_Flag Then TestSeqSplitStoreName = SplitInputCondition(Instance_Data.Meas_StoreName, "+")
        Call HIP_Evaluate_ForceVal_New(Instance_Data.ForceI_Val)
        Call HIP_Evaluate_ForceVal_New(Instance_Data.ForceV_Val)
        ReDim TestConditionSeqData(UBound(TestSeqSplitMeasCase))
        ''Start-------Integrate WaiTime - Carter, 20190610
        Dim WaitTime_Temp() As String
        ReDim WaitTime_Temp(UBound(TestSeqSplitMeasCase))
        Call WaitTime_Check(TestSeqSplitMeasCase, WaitTime_Temp)
        Instance_Data.WaitTime_VFIRZ = Join(WaitTime_Temp, "+")
        ''End-------Integrate WaiTime - Carter, 20190610
        For SeqIdx = 0 To UBound(TestSeqSplitMeasCase)
            TestSeqSplitMeasCase(SeqIdx) = UCase(TestSeqSplitMeasCase(SeqIdx))
            TestConditionSeqData(SeqIdx).MeasCase = TestSeqSplitMeasCase(SeqIdx)
            
            ''Start-------Define store name for sweep
            Dim TestSeqSweepSplitStoreName() As String
            If Meas_StoreName_Flag Then
                If TestSeqSplitStoreName(SeqIdx) <> "" Then
                    TestSeqSweepSplitStoreName = SplitInputCondition(TestSeqSplitStoreName(SeqIdx), ":")
                    ReDim TestConditionSeqData(SeqIdx).Meas_StoreDicName(UBound(TestSeqSweepSplitStoreName))
                Else
                    ReDim TestConditionSeqData(SeqIdx).Meas_StoreDicName(0)
                End If
            End If
            ''End-------Define store name for sweep
            
'===============================================================================================================
            If (InStr(TestSeqSplitMeasCase(SeqIdx), "V") <> 0) Then
            
                TestSeqSplitPin = SplitInputCondition(Instance_Data.MeasV_Pins, SplitSeq_Pin) 'SplitSeq_Pin = "+"
                TestSeqSplitRange = SplitInputCondition(Instance_Data.MeasI_Range, SplitRange) 'SplitRange = "+"
                TestSeqSplitForceValue = SplitInputCondition(Instance_Data.ForceI_Val, SplitSeq_ForceVal) 'SplitSeq_ForceVal = "|"
                TestSeqSplitWaitTime = SplitInputCondition(Instance_Data.WaitTime_VFIRZ, SplitSeq_WaitTime) 'SplitSeq_WaitTime = "+"
                
                ReDim TestConditionSeqData(SeqIdx).MeasV(Len(TestConditionSeqData(SeqIdx).MeasCase) - 1)
                
                For SweepIdx = 0 To UBound(TestConditionSeqData(SeqIdx).MeasV)
                    
                    Call ParseData(TestConditionSeqData(SeqIdx).MeasV(SweepIdx), "V", _
                                CheckAndReturnArrayData(TestSeqSplitPin, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitRange, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitForceValue, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitWaitTime, CLng(SeqIdx)), _
                                CLng(SweepIdx))
                                
                    If Meas_StoreName_Flag Then
                        If TestSeqSplitStoreName(SeqIdx) <> "" Then TestConditionSeqData(SeqIdx).Meas_StoreDicName(SweepIdx) = CheckAndReturnArrayData(TestSeqSweepSplitStoreName, CLng(SweepIdx))
                    End If
                Next SweepIdx
                
'===============================================================================================================
            ElseIf (InStr(TestSeqSplitMeasCase(SeqIdx), "I") <> 0) Or (InStr(TestSeqSplitMeasCase(SeqIdx), "P") <> 0) Then
                
                ReDim TestConditionSeqData(SeqIdx).MeasI(Len(TestConditionSeqData(SeqIdx).MeasCase) - 1)
                
                TestSeqSplitPin = SplitInputCondition(Instance_Data.MeasI_pinS, SplitSeq_Pin)
                TestSeqSplitRange = SplitInputCondition(Instance_Data.MeasI_Range, SplitRange)
                TestSeqSplitForceValue = SplitInputCondition(Instance_Data.ForceV_Val, SplitSeq_ForceVal)
                TestSeqSplitWaitTime = SplitInputCondition(Instance_Data.WaitTime_VFIRZ, SplitSeq_WaitTime)
                
                For SweepIdx = 0 To UBound(TestConditionSeqData(SeqIdx).MeasI)
                    Call ParseData(TestConditionSeqData(SeqIdx).MeasI(SweepIdx), "I", _
                                CheckAndReturnArrayData(TestSeqSplitPin, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitRange, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitForceValue, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitWaitTime, CLng(SeqIdx)), _
                                CLng(SweepIdx))
                                
                    If Meas_StoreName_Flag Then
                        If TestSeqSplitStoreName(SeqIdx) <> "" Then TestConditionSeqData(SeqIdx).Meas_StoreDicName(SweepIdx) = CheckAndReturnArrayData(TestSeqSweepSplitStoreName, CLng(SweepIdx))
                    End If
                Next SweepIdx
            '===============================================================================================================
            ElseIf (InStr(TestSeqSplitMeasCase(SeqIdx), "F") <> 0) Then
                
                Dim TestSeqInterval() As String
                ReDim TestConditionSeqData(SeqIdx).measf(Len(TestConditionSeqData(SeqIdx).MeasCase) - 1)
                
                Call Freq_ProcessEventSourceTerminationMode(Instance_Data.MeasF_EventSourceWithTerminationMode, Instance_Data.MeasF_EventSource, Instance_Data.MeasF_EnableVtMode_Flag)
                
                TestSeqSplitPin = SplitInputCondition(Instance_Data.MeasF_PinS_SingleEnd, SplitSeq_Pin)
                TestSeqSplitDiffPin = SplitInputCondition(Instance_Data.MeasF_PinS_Differential, SplitSeq_Pin)
                TestSeqInterval = SplitInputCondition(Instance_Data.MeasF_Interval, SplitSeq_Pin)
                ''-----------Carter, 20190611
                TestSeqSplitWaitTime = SplitInputCondition(Instance_Data.WaitTime_VFIRZ, SplitSeq_WaitTime)
                ''-----------Carter, 20190611

                For SweepIdx = 0 To UBound(TestConditionSeqData(SeqIdx).measf)
                    ''-----------Carter, 20190611
                    Call ParseData_Freq(TestConditionSeqData(SeqIdx).measf(SweepIdx), "F", _
                                    CheckAndReturnArrayData(TestSeqSplitPin, CLng(SeqIdx)), _
                                    CheckAndReturnArrayData(TestSeqSplitDiffPin, CLng(SeqIdx)), _
                                    CheckAndReturnArrayData(TestSeqInterval, CLng(SeqIdx)), _
                                    CheckAndReturnArrayData(TestSeqSplitWaitTime, CLng(SeqIdx)), _
                                    CLng(SweepIdx))
                    ''-----------Carter, 20190611

                    If Meas_StoreName_Flag Then
                        If TestSeqSplitStoreName(SeqIdx) <> "" Then TestConditionSeqData(SeqIdx).Meas_StoreDicName(SweepIdx) = CheckAndReturnArrayData(TestSeqSweepSplitStoreName, CLng(SweepIdx))
                    End If
                Next SweepIdx
            '===============================================================================================================
            ElseIf (InStr(TestSeqSplitMeasCase(SeqIdx), "Z") <> 0) Then
                ReDim TestConditionSeqData(SeqIdx).MeasZ(Len(TestConditionSeqData(SeqIdx).MeasCase) - 1)
                TestSeqSplitPin = SplitInputCondition(Instance_Data.MeasI_pinS, SplitSeq_Pin)
                TestSeqSplitRange = SplitInputCondition(Instance_Data.MeasI_Range, SplitRange)
                TestSeqSplitForceValue = SplitInputCondition(Instance_Data.ForceV_Val, SplitSeq_ForceVal)
                TestSeqSplitWaitTime = SplitInputCondition(Instance_Data.WaitTime_VFIRZ, SplitSeq_WaitTime)
                
                For SweepIdx = 0 To UBound(TestConditionSeqData(SeqIdx).MeasZ)
                    Call ParseData(TestConditionSeqData(SeqIdx).MeasZ(SweepIdx), "Z", _
                                CheckAndReturnArrayData(TestSeqSplitPin, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitRange, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitForceValue, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitWaitTime, CLng(SeqIdx)), _
                                CLng(SweepIdx))
                                
                    If Meas_StoreName_Flag Then
                        If TestSeqSplitStoreName(SeqIdx) <> "" Then TestConditionSeqData(SeqIdx).Meas_StoreDicName(SweepIdx) = CheckAndReturnArrayData(TestSeqSweepSplitStoreName, CLng(SweepIdx))
                    End If
                Next SweepIdx
            '===============================================================================================================
            ElseIf (InStr(TestSeqSplitMeasCase(SeqIdx), "R") <> 0) Then
                ReDim TestConditionSeqData(SeqIdx).MeasR(Len(TestConditionSeqData(SeqIdx).MeasCase) - 1)
                TestSeqSplitPin = SplitInputCondition(Instance_Data.MeasI_pinS, SplitSeq_Pin)
                TestSeqSplitRange = SplitInputCondition(Instance_Data.MeasI_Range, SplitRange)
                TestSeqSplitForceValue = SplitInputCondition(Instance_Data.ForceV_Val, SplitSeq_ForceVal)
                TestSeqSplitWaitTime = SplitInputCondition(Instance_Data.WaitTime_VFIRZ, SplitSeq_WaitTime)
                
                For SweepIdx = 0 To UBound(TestConditionSeqData(SeqIdx).MeasR)
                    Call ParseData(TestConditionSeqData(SeqIdx).MeasR(SweepIdx), "R", _
                                CheckAndReturnArrayData(TestSeqSplitPin, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitRange, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitForceValue, CLng(SeqIdx)), _
                                CheckAndReturnArrayData(TestSeqSplitWaitTime, CLng(SeqIdx)), _
                                CLng(SweepIdx))
                                
                    If Meas_StoreName_Flag Then
                        If TestSeqSplitStoreName(SeqIdx) <> "" Then TestConditionSeqData(SeqIdx).Meas_StoreDicName(SweepIdx) = CheckAndReturnArrayData(TestSeqSweepSplitStoreName, CLng(SweepIdx))
                    End If
                Next SweepIdx
                '===============================================================================================================
            End If
            
        Next SeqIdx
    End If
    Exit Function
err:
    TheExec.Datalog.WriteComment "<Error> " + "ProcessInputToGLB" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function SplitInputCondition(InputData As String, SplitChar As String, Optional index As Long = -1) As Variant

    Dim Temp_RT_String() As String
    If InputData = "" Then
        If index = -1 Then
            ReDim Temp_RT_String(0): Temp_RT_String(0) = "" ''Carter, 20190605
'            Temp_RT_String = Split("", ",")
            SplitInputCondition = Temp_RT_String
        Else
            SplitInputCondition = ""
        End If
        Exit Function
    End If
    If InStr(InputData, SplitChar) <> 0 Then
        Temp_RT_String = Split(InputData, SplitChar)
        
    Else
        ReDim Temp_RT_String(0)
        Temp_RT_String(0) = InputData
    End If
    SplitInputCondition = Temp_RT_String
    
    If index = -1 Then
        SplitInputCondition = Temp_RT_String
    Else
        If index > UBound(Temp_RT_String) Then
            SplitInputCondition = Temp_RT_String(0)
        Else
            SplitInputCondition = Temp_RT_String(index)
        End If
    End If

End Function
Public Function SortPinInstrument(Pin As String) As String
    Dim Pin_Cnt As Long
    Dim Pins() As String
    Dim pin_instrument As String
    
    Call TheExec.DataManager.DecomposePinList(Pin, Pins, Pin_Cnt)
    If Pin_Cnt <> 0 Then
        pin_instrument = gl_GetInstrument_Dic(LCase(Pins(0)))
        SortPinInstrument = pin_instrument
    Else
        SortPinInstrument = ""
    End If

End Function
Public Function SortPinChannelType(Pin As String) As String
    Dim Pin_Cnt As Long
    Dim Pins() As String
    Dim pin_instrument As String
    
    Call TheExec.DataManager.DecomposePinList(Pin, Pins, Pin_Cnt)
    If Pin_Cnt <> 0 Then
        pin_instrument = gl_GetInstrumentType_Dic(LCase(Pins(0)))
        SortPinChannelType = pin_instrument
    Else
        SortPinChannelType = ""
    End If

End Function

Public Function GetInstTypToDic(pin_grp As String) As Long
    Dim Pins() As String
    Dim Pin_Cnt As Long
    'Dim pin_grp As String
    Dim var As Variant
    Dim PinName As String
    Dim NumTypes As Long
    Dim PowerType() As String
    
    gl_GetInstrument_Dic.RemoveAll
    gl_GetInstrumentType_Dic.RemoveAll
    
    TheExec.DataManager.DecomposePinList pin_grp, Pins, Pin_Cnt
    
    If Pin_Cnt = 0 Then Exit Function
    
    For Each var In Pins
        PinName = LCase(var)
        
        If gl_GetInstrumentType_Dic.Exists(PinName) Then
        Else
            Call TheExec.DataManager.GetChannelTypes(PinName, NumTypes, PowerType())
            Call gl_GetInstrumentType_Dic.Add(PinName, PowerType(0))
        End If
        
        If LCase(PowerType(0)) <> "n/c" Then
            If gl_GetInstrument_Dic.Exists(PinName) Then
            
            Else
                Call gl_GetInstrument_Dic.Add(PinName, GetInstrument(PinName, 0))
            End If
        End If
    Next var
End Function


Public Function Rtn_Dic_count(in_dic As Scripting.Dictionary) As Long
    On Error GoTo err

    Rtn_Dic_count = in_dic.Count
    Exit Function
err:
    Rtn_Dic_count = 0
End Function

Public Function ParseData(ByRef Measure As Meas_Type, MeasType As String, TestSeqSweepSplitPin As String, TestSeqSweepSplitRange As String, TestSeqSweepSplitForceValue As String, TestSeqSweepSplitWaitTime As String, TestSeqSweepIndex As Long) As Meas_Type
    
    Dim i As Integer: i = 0
    Dim num_pins As Long
    Dim instr_pins() As String
    Dim DCVS_HW_Value As String
    
    Dim TestSeqSweepLoopPinSplitPinAry() As String
    Dim TestSeqSweepLoopPinSplitRangeAry() As String
    Dim TestSeqSweepLoopPinSplitWaitTimeAry() As String ''Carter, 20190610
    Dim TestSeqSweepLoopPinSplitForceValueAry() As String
    
    Dim TestSeqSweepLoopPinSplitPin As String
    Dim TestSeqSweepLoopPinSplitRange As String
    Dim TestSeqSweepLoopPinSplitWaitTime As String
    Dim TestSeqSweepLoopPinSplitForceValue As String

    Dim PinIdx As Long
    Dim temp_string As String
    Dim GetPinCHType As String
    Dim Remove_PowerPin_Flag As Boolean ''Default is false
    On Error GoTo err:
    
    If MeasType = "I" And TestSeqSweepIndex > 0 Then Remove_PowerPin_Flag = True
    
    If TestSeqSweepSplitRange = "" Then TestSeqSweepSplitRange = "0"
    If TestSeqSweepSplitForceValue = "" Then TestSeqSweepSplitForceValue = "0"
    
    ''Comment - SplitByPinForceVal = ","
    If ((InStr(TestSeqSweepSplitForceValue, SplitByPinForceVal) <> 0) Or (InStr(TestSeqSweepSplitRange, SplitByPinForceVal) <> 0)) Then
        Measure.Setup_ByTypeByPin_Flag = True
    End If
    TestSeqSweepLoopPinSplitPinAry = SplitInputCondition(TestSeqSweepSplitPin, ",")
    TestSeqSweepLoopPinSplitRangeAry = SplitInputCondition(TestSeqSweepSplitRange, ",")
    TestSeqSweepLoopPinSplitWaitTimeAry = SplitInputCondition(TestSeqSweepSplitWaitTime, ",") ''Carter, 20190610
    TestSeqSweepLoopPinSplitForceValueAry = SplitInputCondition(TestSeqSweepSplitForceValue, ",")
    
    For PinIdx = 0 To UBound(TestSeqSweepLoopPinSplitPinAry)
        
        TestSeqSweepLoopPinSplitPin = SplitInputCondition(CheckAndReturnArrayData(TestSeqSweepLoopPinSplitPinAry, PinIdx), ":", TestSeqSweepIndex)
        TestSeqSweepLoopPinSplitRange = SplitInputCondition(CheckAndReturnArrayData(TestSeqSweepLoopPinSplitRangeAry, PinIdx), ":", TestSeqSweepIndex)
        TestSeqSweepLoopPinSplitWaitTime = CheckAndReturnArrayData(TestSeqSweepLoopPinSplitWaitTimeAry, PinIdx) ''Carter, 20190610
        TestSeqSweepLoopPinSplitForceValue = SplitInputCondition(CheckAndReturnArrayData(TestSeqSweepLoopPinSplitForceValueAry, PinIdx), ":", TestSeqSweepIndex)

        If TestSeqSweepLoopPinSplitForceValue = "" Then TestSeqSweepLoopPinSplitForceValue = "0"
        If TestSeqSweepLoopPinSplitRange = "" Then TestSeqSweepLoopPinSplitRange = pc_Def_Default_Range_By_Instrument
        temp_string = SortPinInstrument(CStr(SplitInputCondition(TestSeqSweepLoopPinSplitPinAry(PinIdx), ":", CLng(TestSeqSweepIndex))))
                
        If (temp_string = "HSD-U") Then
            If Measure.Setup_ByTypeByPin_Flag = True Then
                If Measure.Setup_ByTypeByPin.PPMU_Flag = False Then
                    ReDim Measure.Setup_ByTypeByPin.PPMU(0)
                    Measure.Setup_ByTypeByPin.PPMU_Flag = True
                Else
                    ReDim Preserve Measure.Setup_ByTypeByPin.PPMU(UBound(Measure.Setup_ByTypeByPin.PPMU) + 1)
                End If
                Measure.Setup_ByTypeByPin.PPMU(UBound(Measure.Setup_ByTypeByPin.PPMU)).Pin = TestSeqSweepLoopPinSplitPin
                '''---------Check PPMU/Digital---------
                If Instance_Data.InstSpecialSetting = DigitalConnectPPMU Then ''Carter, 20190531 , PPMU max = 50mA, Digital channel Max = 20mA, when connect both, max is only 20mA.
                    If (TestSeqSweepLoopPinSplitRange = pc_Def_Default_Range_By_Instrument) Or (CDbl(TestSeqSweepLoopPinSplitRange) >= pc_Def_PPMU_Digital_MaxCurrRange) Then
                        TestSeqSweepLoopPinSplitRange = pc_Def_PPMU_Digital_MaxCurrRange
                    End If
                Else
                    If TestSeqSweepLoopPinSplitRange = pc_Def_Default_Range_By_Instrument Then TestSeqSweepLoopPinSplitRange = pc_Def_PPMU_Max_InitialValue_FI_Range
                End If
                '''---------Check PPMU/Digital---------
                Measure.Setup_ByTypeByPin.PPMU(UBound(Measure.Setup_ByTypeByPin.PPMU)).Meas_Range = TestSeqSweepLoopPinSplitRange
                
                If MeasType <> "Z" Then
                    Measure.Setup_ByTypeByPin.PPMU(UBound(Measure.Setup_ByTypeByPin.PPMU)).ForceValue1 = TestSeqSweepLoopPinSplitForceValue
                Else
                    Measure.Setup_ByTypeByPin.PPMU(UBound(Measure.Setup_ByTypeByPin.PPMU)).ForceValue1 = SplitInputCondition(TestSeqSweepLoopPinSplitForceValue, "&", 0)
                    Measure.Setup_ByTypeByPin.PPMU(UBound(Measure.Setup_ByTypeByPin.PPMU)).ForceValue2 = SplitInputCondition(TestSeqSweepLoopPinSplitForceValue, "&", 1)
                End If
            Else
                If MeasType <> "Z" Then
                    Measure.Setup_ByType.PPMU.ForceValue1 = TestSeqSweepLoopPinSplitForceValue
                Else
                    Measure.Setup_ByType.PPMU.ForceValue1 = SplitInputCondition(TestSeqSweepLoopPinSplitForceValue, "&", 0)
                    Measure.Setup_ByType.PPMU.ForceValue2 = SplitInputCondition(TestSeqSweepLoopPinSplitForceValue, "&", 1)
                End If
                '''---------Check PPMU/Digital---------
                If Instance_Data.InstSpecialSetting = DigitalConnectPPMU Then
                    If (TestSeqSweepLoopPinSplitRange = pc_Def_Default_Range_By_Instrument) Or (CDbl(TestSeqSweepLoopPinSplitRange) >= pc_Def_PPMU_Digital_MaxCurrRange) Then
                        TestSeqSweepLoopPinSplitRange = pc_Def_PPMU_Digital_MaxCurrRange
                    End If
                Else
                    If TestSeqSweepLoopPinSplitRange = pc_Def_Default_Range_By_Instrument Then TestSeqSweepLoopPinSplitRange = pc_Def_PPMU_Max_InitialValue_FI_Range
                End If
                '''---------Check PPMU/Digital---------
                Measure.Setup_ByType.PPMU.Meas_Range = TestSeqSweepLoopPinSplitRange
            End If
            Measure.Pins.PPMU = Measure.Pins.PPMU & "," & TestSeqSweepLoopPinSplitPin
            Measure.WaitTime.PPMU = CStr(Return_RangeAndMaxWaitTime(Measure, TestSeqSweepLoopPinSplitWaitTime, "PPMU")) ''Carter, 20190610
            
        ElseIf (temp_string = "DC-07") Then
            GetPinCHType = SortPinChannelType(CStr(SplitInputCondition(TestSeqSweepLoopPinSplitPinAry(PinIdx), ":", CLng(TestSeqSweepIndex))))
            If GetPinCHType = "DCDiffMeter" Then
                Dim HighSidePin As String
                Dim LowSidePin As String
                Measure.DiffMeter_Flag = True
                Call UVI80_DIFFMETER_INIT(TestSeqSweepLoopPinSplitPin, HighSidePin, LowSidePin)
                Measure.Setup_ByType.UVI80.Pin = HighSidePin
                Measure.Setup_ByType.UVI80.Pin_Diff_L = LowSidePin
            Else
                   If Measure.Setup_ByTypeByPin_Flag = True Then
                        If Measure.Setup_ByTypeByPin.UVI80_Flag = False Then
                            ReDim Measure.Setup_ByTypeByPin.UVI80(0)
                            Measure.Setup_ByTypeByPin.UVI80_Flag = True
                        Else
                            ReDim Preserve Measure.Setup_ByTypeByPin.UVI80(UBound(Measure.Setup_ByTypeByPin.UVI80) + 1)
                        End If
                        Measure.Setup_ByTypeByPin.UVI80(UBound(Measure.Setup_ByTypeByPin.UVI80)).Pin = TestSeqSweepLoopPinSplitPin
                        Measure.Setup_ByTypeByPin.UVI80(UBound(Measure.Setup_ByTypeByPin.UVI80)).Meas_Range = TestSeqSweepLoopPinSplitRange
                        Measure.Setup_ByTypeByPin.UVI80(UBound(Measure.Setup_ByTypeByPin.UVI80)).ForceValue1 = TestSeqSweepLoopPinSplitForceValue
                    Else
                        Measure.Setup_ByType.UVI80.ForceValue1 = TestSeqSweepLoopPinSplitForceValue
                        Measure.Setup_ByType.UVI80.Meas_Range = TestSeqSweepLoopPinSplitRange
                    End If
            End If
            Measure.Pins.UVI80 = Measure.Pins.UVI80 & "," & TestSeqSweepLoopPinSplitPin
            Measure.WaitTime.UVI80 = CStr(Return_RangeAndMaxWaitTime(Measure, TestSeqSweepLoopPinSplitWaitTime, "UVI80")) ''Carter, 20190610
            
        ElseIf (temp_string = "HexVS") And (Not Remove_PowerPin_Flag) Then
            If Measure.Setup_ByTypeByPin_Flag = True Then
                If Measure.Setup_ByTypeByPin.HexVS_Flag = False Then
                    ReDim Measure.Setup_ByTypeByPin.HexVS(0)
                    Measure.Setup_ByTypeByPin.HexVS_Flag = True
                Else
                    ReDim Preserve Measure.Setup_ByTypeByPin.HexVS(UBound(Measure.Setup_ByTypeByPin.HexVS) + 1)
                End If
                Measure.Setup_ByTypeByPin.HexVS(UBound(Measure.Setup_ByTypeByPin.HexVS)).Pin = TestSeqSweepLoopPinSplitPin
                Measure.Setup_ByTypeByPin.HexVS(UBound(Measure.Setup_ByTypeByPin.HexVS)).Meas_Range = TestSeqSweepLoopPinSplitRange
                Measure.Setup_ByTypeByPin.HexVS(UBound(Measure.Setup_ByTypeByPin.HexVS)).ForceValue1 = TestSeqSweepLoopPinSplitForceValue
            Else
                Measure.Setup_ByType.HexVS.ForceValue1 = TestSeqSweepLoopPinSplitForceValue
                Measure.Setup_ByType.HexVS.Meas_Range = TestSeqSweepLoopPinSplitRange
            End If
            Measure.Pins.HexVS = Measure.Pins.HexVS & "," & TestSeqSweepLoopPinSplitPin
            Measure.WaitTime.HexVS = CStr(Return_RangeAndMaxWaitTime(Measure, TestSeqSweepLoopPinSplitWaitTime, "HexVS")) ''Carter, 20190610
            
        ElseIf (temp_string = "VHDVS") And (Not Remove_PowerPin_Flag) Then
            If Measure.Setup_ByTypeByPin_Flag = True Then
                If Measure.Setup_ByTypeByPin.UVS256_Flag = False Then
                    ReDim Measure.Setup_ByTypeByPin.UVS256(0)
                    Measure.Setup_ByTypeByPin.UVS256_Flag = True
                Else
                    ReDim Preserve Measure.Setup_ByTypeByPin.UVS256(UBound(Measure.Setup_ByTypeByPin.UVS256) + 1)
                End If
                Measure.Setup_ByTypeByPin.UVS256(UBound(Measure.Setup_ByTypeByPin.UVS256)).Pin = TestSeqSweepLoopPinSplitPin
                Measure.Setup_ByTypeByPin.UVS256(UBound(Measure.Setup_ByTypeByPin.UVS256)).Meas_Range = TestSeqSweepLoopPinSplitRange
                Measure.Setup_ByTypeByPin.UVS256(UBound(Measure.Setup_ByTypeByPin.UVS256)).ForceValue1 = TestSeqSweepLoopPinSplitForceValue
            Else
                Measure.Setup_ByType.UVS256.ForceValue1 = TestSeqSweepLoopPinSplitForceValue
                Measure.Setup_ByType.UVS256.Meas_Range = TestSeqSweepLoopPinSplitRange
            End If
            Measure.Pins.UVS256 = Measure.Pins.UVS256 & "," & TestSeqSweepLoopPinSplitPin
            Measure.WaitTime.UVS256 = CStr(Return_RangeAndMaxWaitTime(Measure, TestSeqSweepLoopPinSplitWaitTime, "UVS256")) ''Carter, 20190610
            
        ElseIf (temp_string = "VSM") And (Not Remove_PowerPin_Flag) Then
            If Measure.Setup_ByTypeByPin_Flag = True Then
                If Measure.Setup_ByTypeByPin.VSM_Flag = False Then
                    ReDim Measure.Setup_ByTypeByPin.VSM(0)
                    Measure.Setup_ByTypeByPin.VSM_Flag = True
                Else
                    ReDim Preserve Measure.Setup_ByTypeByPin.VSM(UBound(Measure.Setup_ByTypeByPin.VSM) + 1)
                End If
                Measure.Setup_ByTypeByPin.VSM(UBound(Measure.Setup_ByTypeByPin.VSM)).Pin = TestSeqSweepLoopPinSplitPin
                Measure.Setup_ByTypeByPin.VSM(UBound(Measure.Setup_ByTypeByPin.VSM)).Meas_Range = TestSeqSweepLoopPinSplitRange
                Measure.Setup_ByTypeByPin.VSM(UBound(Measure.Setup_ByTypeByPin.VSM)).ForceValue1 = TestSeqSweepLoopPinSplitForceValue
            Else
                Measure.Setup_ByType.VSM.ForceValue1 = TestSeqSweepLoopPinSplitForceValue
                Measure.Setup_ByType.VSM.Meas_Range = TestSeqSweepLoopPinSplitRange
            End If
            Measure.Pins.VSM = Measure.Pins.VSM & "," & TestSeqSweepLoopPinSplitPin
            Measure.WaitTime.VSM = CStr(Return_RangeAndMaxWaitTime(Measure, TestSeqSweepLoopPinSplitWaitTime, "VSM")) ''Carter, 20190610
        End If
        
        If Not Measure.ForceValueDic.Exists(UCase(TestSeqSweepLoopPinSplitPin)) Then
            Measure.ForceValueDic.Add UCase(TestSeqSweepLoopPinSplitPin), TestSeqSweepLoopPinSplitForceValue
        End If
        '''---------Compare with HW value and Print to data log---------
        Call TheExec.DataManager.DecomposePinList(TestSeqSweepLoopPinSplitPin, instr_pins(), num_pins)
        For i = 0 To num_pins - 1
            If Not Measure.ForceValueDic_HWCom.Exists(UCase(instr_pins(i))) Then
                If (temp_string = "HexVS") Or (temp_string = "VHDVS") Or (temp_string = "VSM") Then
                    '''---------Check Force Voltage for DCVS---------
                    DCVS_HW_Value = CStr(FormatNumber(TheHdw.DCVS.Pins(instr_pins(i)).Voltage.Value, 3))
                    Measure.ForceValueDic_HWCom.Add UCase(instr_pins(i)), DCVS_HW_Value
                    
'                    If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
'                        TheExec.Datalog.WriteComment (" HIP Parse Data=====> DCVS Force Volt value, " & instr_pins(i) & " =" & DCVS_HW_Value)
'                    End If
                    '''---------Check Force Voltage for DCVS---------
                Else
                    
                    If sweep_power_val_per_loop_count <> "" Then
                        Measure.ForceValueDic_HWCom.Add UCase(instr_pins(i)), sweep_power_val_per_loop_count
                    Else
                        Measure.ForceValueDic_HWCom.Add UCase(instr_pins(i)), TestSeqSweepLoopPinSplitForceValue
                    End If
                    
                    If Instance_Data.InstSpecialSetting = PPMU_SerialMeasurement Then
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If TestSeqSweepLoopPinSplitRange > 0.002 Then
                            TestSeqSweepLoopPinSplitRange = 0.05
                        ElseIf TestSeqSweepLoopPinSplitRange > 0.0002 Then
                            TestSeqSweepLoopPinSplitRange = 0.002
                        ElseIf TestSeqSweepLoopPinSplitRange > 0.00002 Then
                            TestSeqSweepLoopPinSplitRange = 0.0002
                        ElseIf TestSeqSweepLoopPinSplitRange > 0.000002 Then
                            TestSeqSweepLoopPinSplitRange = 0.00002
                        Else
                            TestSeqSweepLoopPinSplitRange = 0.000002
                        End If
                        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        Measure.MeasCurRangeDic.Add UCase(instr_pins(i)), TestSeqSweepLoopPinSplitRange
'                        TheExec.Datalog.WriteComment instr_pins(i) & ": FV = " & TestSeqSweepLoopPinSplitForceValue
'                        TheExec.Datalog.WriteComment instr_pins(i) & ": CR = " & TestSeqSweepLoopPinSplitRange
                    End If
                End If
            End If
        Next i
        '''---------Compare with HW value and Print to data log---------
    Next PinIdx
    
    Measure.Pins.PPMU = Replace(Measure.Pins.PPMU, ",", "", 1, 1)
    Measure.Pins.UVI80 = Replace(Measure.Pins.UVI80, ",", "", 1, 1)
    Measure.Pins.HexVS = Replace(Measure.Pins.HexVS, ",", "", 1, 1)
    Measure.Pins.UVS256 = Replace(Measure.Pins.UVS256, ",", "", 1, 1)
    Measure.Pins.VSM = Replace(Measure.Pins.VSM, ",", "", 1, 1)
    Remove_PowerPin_Flag = False
    
'    Measure.WaitTime = tempMeas_WaitTime
'    ParseData = Measure
    Exit Function
err:
    TheExec.Datalog.WriteComment "<Error> " + "ParseData" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckAndReturnArrayData(Ary() As String, index As Long, Optional MissMatchReturnChar As String = "N") As String
    
    If UBound(Ary) <> -1 Then
        If UBound(Ary) < index Then
            If MissMatchReturnChar = "N" Then
                CheckAndReturnArrayData = Ary(0)
            Else
                CheckAndReturnArrayData = MissMatchReturnChar
            End If
        Else
            CheckAndReturnArrayData = Ary(index)
        End If
    Else
        If MissMatchReturnChar = "N" Then
            CheckAndReturnArrayData = ""
        Else
            CheckAndReturnArrayData = MissMatchReturnChar
        End If
    End If
End Function

Public Function HardIP_MeasureFreq()

    Dim measf As MeasF_Type
    Dim DictKey_StoreVT As String
    Dim SplitFreqVtValue() As String
    Dim Rtn_MeasureResult As New PinListData
    Dim PinList() As String
    Dim cus_pins_flag As Boolean
    Dim DicStoreName As String
    Dim TempStr_Semi() As String
    Dim TestNameInput_TempArr() As String
    Dim idx As Integer
    Dim Temp_LimitIndex As Long
    Dim Pin As New PinList
    Dim p As Long
    Dim TestNameInput As String
    Dim site As Variant
    
    measf = TestConditionSeqData(Instance_Data.TestSeqNum).measf(Instance_Data.TestSeqSweepNum)
    If Meas_StoreName_Flag Then DicStoreName = TestConditionSeqData(Instance_Data.TestSeqNum).Meas_StoreDicName(Instance_Data.TestSeqSweepNum)

    Pin.Value = measf.Pins
    If measf.MeasureThreshold_Flag = True Then
        Call Freq_PPMU_Meas_VOH(Pin, CDbl(measf.ThresholdPercentage), measf.EnableVtMode_Flag, measf.EventSource)
    End If
    If measf.EnableVtMode_Flag = True Then
        TheHdw.Digital.Pins(measf.Pins).Levels.DriverMode = tlDriverModeVt
    End If
    If measf.WalkingStrobe_Flag = True Then
        If Instance_Data.CUS_Str_DigSrcData <> "" Then
            SplitFreqVtValue = Split(Instance_Data.CUS_Str_DigSrcData, ":")
            If UCase(SplitFreqVtValue(0)) = "STORE_VT" Then
                DictKey_StoreVT = SplitFreqVtValue(1)
            End If
        End If
        Call Freq_WalkingStrobe_Meas_VOHVOL(Pin, Instance_Data.MeasF_WalkingStrobe_StartV, Instance_Data.MeasF_WalkingStrobe_EndV, Instance_Data.MeasF_WalkingStrobe_StepVoltage, Instance_Data.MeasF_WalkingStrobe_BothVohVolDiffV, Instance_Data.MeasF_WalkingStrobe_interval, Instance_Data.MeasF_WalkingStrobe_miniFreq, DictKey_StoreVT)
    End If
    If GetInstrument(CStr(Split(measf.Pins, ",")(0)), 0) = "HSD-U" Then
        Call HardIP_FrequencyMeasure(Rtn_MeasureResult)
    Else
        Call HardIP_FrequencyMeasure_Dctime(Rtn_MeasureResult)
    End If
'--------------------------------------------------------------------------------------------------------------------------------------------------
    If Not ByPassTestLimit Then
        If measf.Differential_Flag = True Then
            If Instance_Data.CUS_Str_MainProgram <> "" And InStr(UCase(Instance_Data.CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
                    For p = 0 To Rtn_MeasureResult.Pins.Count - 1
                        If InStr(UCase(Rtn_MeasureResult.Pins(p)), "_P") Then
                            TestNameInput = Report_TName_From_Instance("F", Rtn_MeasureResult.Pins(p), , CInt(Instance_Data.TestSeqNum), p)
                            TheExec.Flow.TestLimit resultVal:=Rtn_MeasureResult.Pins(p), Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
                        End If
                    Next p
                 Else
                 '=======================================================================================================
                   If LCase(Instance_Data.CUS_Str_MainProgram) Like "*cus_pin_list*" Then
                        PinList = Split(Instance_Data.CUS_Str_MainProgram, ";")
                        For p = 0 To UBound(PinList)
                                If LCase(PinList(p)) Like "*cus_pin_list*" Then
                                    PinList = Split(PinList(p), ":")
                                    PinList = Split(PinList(1), ",")
                                    Exit For
                                End If
                        Next p
                    End If

                    For p = 0 To Rtn_MeasureResult.Pins.Count - 1 Step 2 ' freq counter result of differential pins is stored in positive pin
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        If InStr(UCase(Instance_Data.CUS_Str_MainProgram), UCase(Rtn_MeasureResult.Pins(p).Name)) <> 0 And Instance_Data.CUS_Str_MainProgram <> "" And LCase(Instance_Data.CUS_Str_MainProgram) Like "*cus_pin_list*" Then
                            cus_pins_flag = True
                            Exit For
                       End If
                       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        TestNameInput = Report_TName_From_Instance("F", Rtn_MeasureResult.Pins(p + 1), , CInt(Instance_Data.TestSeqNum), p + 1)
                        TheExec.Flow.TestLimit resultVal:=Rtn_MeasureResult.Pins(p + 1), Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
                    Next p
                    '-----------------------------------------------------------------------------------------------------
                    If cus_pins_flag = True Then
                        For p = 0 To UBound(PinList) Step 2 ' freq counter result of differential pins is stored in positive pin
                            TestNameInput = Report_TName_From_Instance("F", Rtn_MeasureResult.Pins(PinList(p + 1)), , CInt(Instance_Data.TestSeqNum), p + 1)
                            TheExec.Flow.TestLimit resultVal:=Rtn_MeasureResult.Pins(PinList(p + 1)), Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
                        Next p
                    End If
                    '=======================================================================================================
                End If
        ElseIf InStr(UCase(Instance_Data.CUS_Str_MainProgram), UCase("Calc_Freq_SDLL_SWP")) <> 0 Then
            Call ProsscessTestLimit(Rtn_MeasureResult, "F", CInt(Instance_Data.TestSeqNum), , , , tlForceNone)
        Else
            Call ProsscessTestLimit(Rtn_MeasureResult, "F", CInt(Instance_Data.TestSeqNum))
        End If

        If Instance_Data.SpecialCalcValSetting = RATIO_FREQ Then
            If Instance_Data.TestSeqNum = 0 Then
                For Each site In TheExec.sites.Active
                    Min_Freq(site) = Rtn_MeasureResult.Pins(0).Value(site)
                    Max_freq(site) = Rtn_MeasureResult.Pins(0).Value(site)
                Next site
            Else
                For Each site In TheExec.sites.Active
                    If Rtn_MeasureResult.Pins(0).Value(site) < Min_Freq(site) Then Min_Freq(site) = Rtn_MeasureResult.Pins(0).Value(site)
                    If Rtn_MeasureResult.Pins(0).Value(site) > Max_freq(site) Then Max_freq(site) = Rtn_MeasureResult.Pins(0).Value(site)
                Next site
            End If
            
            If Instance_Data.TestSeqNum = 23 Then
                For Each site In TheExec.sites.Active
                    If Min_Freq(site) = 0 Then Min_Freq(site) = 1
                    TestNameInput = Report_TName_From_Instance("F", Rtn_MeasureResult.Pins(0), "Freq_ratio", CInt(Instance_Data.TestSeqNum), 0)
                    TheExec.Flow.TestLimit resultVal:=Max_freq(site) / Min_Freq(site), lowVal:=1.6, hiVal:=2, Tname:=TestNameInput & "_" & "Freq_ratio", ForceResults:=tlForceNone
                Next site
            End If
        End If
        G_MeasFreqForCZ = Rtn_MeasureResult
        If InStr(UCase(Instance_Data.CUS_Str_MainProgram), UCase("Calc_Freq_SDLL_SWP")) <> 0 Then
            If TestNameInput_TempArr(3) <> 0 Then
                Dim Extra_TName_Str As String
                Extra_TName_Str = TestNameInput_TempArr(1) & ":" & TestNameInput_TempArr(2) & ":" & TestNameInput_TempArr(3) & ":" & TestNameInput_TempArr(4)
                Call CUS_AMP_SDLL_SWP(Rtn_MeasureResult, Extra_TName_Str)
            End If
        End If
    End If
    
    ''Start ---- Carter, 20190521
    If Meas_StoreName_Flag Then
        If DicStoreName <> "" Then Call AddStoredMeasurement(DicStoreName, Rtn_MeasureResult)
    End If
    ''End ---- Carter, 20190521
    
End Function

Public Function HardIP_SetupAndMeasureZ() As Long

    Dim i As Long
    Dim Pins() As Variant
    Dim MeasZ As Meas_Type
    Dim measureCurrent As New PinListData
    Dim Imped As New PinListData
    Dim MeasCurr1 As New PinListData
    Dim MeasCurr2 As New PinListData
    Dim GetRakVal As Double
    Dim Pin As Variant
    Dim DiffVolt As Double
    Dim DiffVolt_Pinlist As New PinListData
    Dim DicStoreName As String
    Dim site As Variant
    
    
    
    ''added by Kaino for 2 iRange for MeasZ
    ''------------------------------------------------------------------------
    Dim Irange() As String
    Dim iRangeStr As String
    Dim CUS_Str() As String
    Dim kk  As Integer
    
    'Instance_Data.CUS_Str_MainProgram = "xxx;MeasZby2iRange(0.002,0.05);xxx"
    If Instance_Data.CUS_Str_MainProgram Like "*MeasZby2iRange(*,*)*" Then
        CUS_Str = Split(Instance_Data.CUS_Str_MainProgram, ";")
        
        For kk = 0 To UBound(CUS_Str())
            If Trim(CUS_Str(kk)) Like "MeasZby2iRange(*,*)" Then
                iRangeStr = CUS_Str(kk)
                iRangeStr = Replace(iRangeStr, "(", ",")
                iRangeStr = Replace(iRangeStr, ")", ",")
                Irange = Split(iRangeStr, ",")
            End If
        Next kk
    Else
        iRangeStr = ""
        ReDim Irange(3)
    End If
    ''------------------------------------------------------------------------
    
    
    MeasZ = TestConditionSeqData(Instance_Data.TestSeqNum).MeasZ(Instance_Data.TestSeqSweepNum)
    If Meas_StoreName_Flag Then DicStoreName = TestConditionSeqData(Instance_Data.TestSeqNum).Meas_StoreDicName(Instance_Data.TestSeqSweepNum)

    ''added by Kaino for 2 iRange for MeasZ
    ''------------------------------------------------------------------------
    If Irange(1) <> "" Then
        TheExec.Datalog.WriteComment " *** set iRange for MeasZ *** " & CDbl(Irange(1) * 1000) & " mA"
        MeasZ.Setup_ByType.PPMU.Meas_Range = Irange(1)        'added by Kaino for 2 iRange for MeasZ
    End If
    ''------------------------------------------------------------------------
    
    TheHdw.Digital.Pins(MeasZ.Pins.PPMU).Disconnect
    If MeasZ.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.PPMU.Pins(MeasZ.Pins.PPMU)
            .Gate = tlOff
            .ForceI pc_Def_PPMU_InitialValue_FI
            .ForceV CDbl(MeasZ.Setup_ByType.PPMU.ForceValue1), CDbl(MeasZ.Setup_ByType.PPMU.Meas_Range)
            .Connect
            .Gate = tlOn
        End With
        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment ("Force V1, Test sequence: " & Instance_Data.TestSeqNum)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasZ.Pins.PPMU & " =" & TheHdw.PPMU.Pins(SplitInputCondition(MeasZ.Pins.PPMU, ",", 0)).MeasureCurrentRange)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasZ.Pins.PPMU & " =" & MeasZ.Setup_ByType.PPMU.Meas_Range)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasZ.Pins.PPMU & " =" & MeasZ.WaitTime.PPMU)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasZ.Pins.PPMU & " =" & MeasZ.Setup_ByType.PPMU.ForceValue1)
        End If
    Else
        For i = 0 To UBound(MeasZ.Setup_ByTypeByPin.PPMU)
            With TheHdw.PPMU.Pins(MeasZ.Setup_ByTypeByPin.PPMU(i).Pin)
                .Gate = tlOff
                .ForceI pc_Def_PPMU_InitialValue_FI
                .ForceV CDbl(MeasZ.Setup_ByTypeByPin.PPMU(i).ForceValue1), CDbl(MeasZ.Setup_ByTypeByPin.PPMU(i).Meas_Range)
                .Connect
                .Gate = tlOn
            End With
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                    TheExec.Datalog.WriteComment ("Force V1, Test sequence: " & Instance_Data.TestSeqNum)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasZ.Setup_ByTypeByPin.PPMU(i).Pin & " =" & TheHdw.PPMU.Pins(MeasZ.Setup_ByTypeByPin.PPMU(i).Pin).MeasureCurrentRange)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasZ.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasZ.Setup_ByTypeByPin.PPMU(i).Meas_Range)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasZ.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasZ.WaitTime.PPMU)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasZ.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasZ.Setup_ByTypeByPin.PPMU(i).ForceValue1)
            End If
        Next i
    End If
    TheHdw.Wait CDbl(MeasZ.WaitTime.PPMU)
    DebugPrintFunc_PPMU CStr(MeasZ.Pins.PPMU)
    MeasCurr1 = TheHdw.PPMU.Pins(MeasZ.Pins.PPMU).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    If gl_Disable_HIP_debug_log = False Then Call ForceVal_Compare(MeasZ, MeasZ.Pins.PPMU, "Z", "PPMU") ''Current1 Force Condition check - Carter, 20190507
    
    ''added by Kaino for 2 iRange for MeasZ
    ''------------------------------------------------------------------------
    If Irange(2) <> "" Then
        TheExec.Datalog.WriteComment " *** set iRange for MeasZ *** " & CDbl(Irange(2) * 1000) & " mA"
        MeasZ.Setup_ByType.PPMU.Meas_Range = Irange(2)        'added by Kaino for 2 iRange for MeasZ
        End If
    ''------------------------------------------------------------------------
    
    If MeasZ.Setup_ByTypeByPin_Flag = False Then
        With TheHdw.PPMU.Pins(MeasZ.Pins.PPMU)
            .ForceV CDbl(MeasZ.Setup_ByType.PPMU.ForceValue2), CDbl(MeasZ.Setup_ByType.PPMU.Meas_Range)
        End With
        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment ("Force V2, Test sequence: " & Instance_Data.TestSeqNum)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasZ.Pins.PPMU & " =" & TheHdw.PPMU.Pins(SplitInputCondition(MeasZ.Pins.PPMU, ",", 0)).MeasureCurrentRange)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasZ.Pins.PPMU & " =" & MeasZ.Setup_ByType.PPMU.Meas_Range)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasZ.Pins.PPMU & " =" & MeasZ.WaitTime.PPMU)
                TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasZ.Pins.PPMU & " =" & MeasZ.Setup_ByType.PPMU.ForceValue2)
        End If
    Else
        For i = 0 To UBound(MeasZ.Setup_ByTypeByPin.PPMU)
            With TheHdw.PPMU.Pins(MeasZ.Setup_ByTypeByPin.PPMU(i).Pin)
                .ForceV CDbl(MeasZ.Setup_ByTypeByPin.PPMU(i).ForceValue2), CDbl(MeasZ.Setup_ByTypeByPin.PPMU(i).Meas_Range)
            End With
            If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                    TheExec.Datalog.WriteComment ("Force V2, Test sequence: " & Instance_Data.TestSeqNum)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & MeasZ.Setup_ByTypeByPin.PPMU(i).Pin & " =" & TheHdw.PPMU.Pins(MeasZ.Setup_ByTypeByPin.PPMU(i).Pin).MeasureCurrentRange)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & MeasZ.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasZ.Setup_ByTypeByPin.PPMU(i).Meas_Range)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & MeasZ.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasZ.WaitTime.PPMU)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & MeasZ.Setup_ByTypeByPin.PPMU(i).Pin & " =" & MeasZ.Setup_ByTypeByPin.PPMU(i).ForceValue2)
            End If
        Next i
    End If
    DebugPrintFunc_PPMU CStr(MeasZ.Pins.PPMU)
    MeasCurr2 = TheHdw.PPMU.Pins(MeasZ.Pins.PPMU).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
    
    If gl_Disable_HIP_debug_log = False Then Call ForceVal_Compare(MeasZ, MeasZ.Pins.PPMU, "Z", "PPMU", 1)
    
    With TheHdw.PPMU.Pins(MeasZ.Pins.PPMU)
            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
            .Disconnect
            .Gate = tlOff
    End With

    TheHdw.Digital.Pins(MeasZ.Pins.PPMU).Connect ''Carter, 20190521
    
    ''''---------Offline---------
    If TheExec.TesterMode = testModeOffline Then
        For Each Pin In MeasCurr1.Pins
            MeasCurr1.Pins(Pin).Value = MeasCurr1.Pins(Pin).Add(0.0001)
            MeasCurr2.Pins(Pin).Value = MeasCurr2.Pins(Pin).Add(0.0002)
        Next Pin
    Else
        For Each Pin In MeasCurr1.Pins
            For Each site In TheExec.sites.Active
                If MeasCurr2.Pins(Pin).Value = MeasCurr1.Pins(Pin).Value Then
                    MeasCurr2.Pins(Pin).Value = MeasCurr1.Pins(Pin).Value + 0.0001
                End If
            Next site
        Next Pin
    End If
    ''''---------Offline---------
    
    If MeasZ.Setup_ByTypeByPin_Flag = False Then
        DiffVolt = CDbl(MeasZ.Setup_ByType.PPMU.ForceValue2) - CDbl(MeasZ.Setup_ByType.PPMU.ForceValue1)
        Imped = MeasCurr2.Math.Subtract(MeasCurr1).Invert.Multiply(DiffVolt).Abs
    Else
        For i = 0 To UBound(MeasZ.Setup_ByTypeByPin.PPMU)
            Imped.AddPin (MeasZ.Setup_ByTypeByPin.PPMU(i).Pin)
            DiffVolt_Pinlist.AddPin (MeasZ.Setup_ByTypeByPin.PPMU(i).Pin)
            DiffVolt_Pinlist.Pins(MeasZ.Setup_ByTypeByPin.PPMU(i).Pin).Value = CDbl(MeasZ.Setup_ByTypeByPin.PPMU(i).ForceValue2) - CDbl(MeasZ.Setup_ByTypeByPin.PPMU(i).ForceValue1)
        Next i
        Imped = MeasCurr2.Math.Subtract(MeasCurr1).Invert.Multiply(DiffVolt_Pinlist).Abs
    End If
    
    Dim GetRakVal_PinList As New PinListData
        
    If Instance_Data.RAK_Flag = R_TraceOnly Then
        For Each Pin In Imped.Pins
            GetRakVal_PinList.AddPin (CStr(Pin))
            GetRakVal_PinList.Pins(Pin) = CurrentJob_Card_RAK.Pins(Pin)
            If gl_Disable_HIP_debug_log = False Then
                For Each site In TheExec.sites.Active
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Curren1= " & MeasCurr1.Pins(Pin).Value(site)
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Curren2= " & MeasCurr2.Pins(Pin).Value(site)
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " = " & Imped.Pins(Pin).Value(site) & ", RAK val = " & GetRakVal_PinList.Pins(Pin).Value(site)
                Next site
            End If
        Next Pin
        Imped = Imped.Math.Subtract(GetRakVal_PinList)
            
    ElseIf Instance_Data.RAK_Flag = R_PathWithContact Then
        For Each Pin In Imped.Pins
            GetRakVal_PinList.AddPin (CStr(Pin))
            GetRakVal_PinList.Pins(Pin) = R_Path_PLD.Pins(Pin)
            If gl_Disable_HIP_debug_log = False Then
                For Each site In TheExec.sites
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Curren1= " & MeasCurr1.Pins(Pin).Value(site)
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Curren2= " & MeasCurr2.Pins(Pin).Value(site)
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " = " & Imped.Pins.Item(Pin).Value(site) & ", R_Path val = " & R_Path_PLD.Pins(Pin).Value(site)
                
                     'TheExec.Datalog.WriteComment Pin & " = " & Imped.Pins.Item(Pin).Value(Site) & ", R_Path val = " & R_Path_PLD.Pins(Pin).Value(Site)
                Next site
            End If
        Next Pin
        Imped = Imped.Math.Subtract(GetRakVal_PinList)
        
    End If
'----------------------------------------------------------------------------------------------------------------------------------------------
    If gl_SweepNum <> "" Then                                            'Edited for CIOTX  20190522
        Call ProsscessTestLimit(Imped, "Z", (CInt(TheExec.Flow.var("SrcCodeIndx").Value) - 1))
    Else
        Call ProsscessTestLimit(Imped, "Z", CInt(Instance_Data.TestSeqNum), CInt(Instance_Data.TestSeqSweepNum))
    End If
'----------------------------------------------------------------------------------------------------------------------------------------------
    ''Start ---- Carter, 20190521
    If Meas_StoreName_Flag Then
        If DicStoreName <> "" Then Call AddStoredMeasurement(DicStoreName, Imped)
    End If
    ''End ---- Carter, 20190521
    
End Function

Public Function Return_RangeAndMaxWaitTime(Measure As Meas_Type, SpecifyWaitTime As String, Optional InstType As String = "") As Double
    
    Dim WaitTime As Double
    Dim MaxWaitTime As Double
    Dim factor As Double
    Dim i As Long
    Dim PinType As String
    Dim SpecifyWaitTime_double As Double
    
    On Error GoTo err
    
    If SpecifyWaitTime = "" Then
        SpecifyWaitTime_double = 0
    Else
        SpecifyWaitTime_double = CDbl(SpecifyWaitTime)
    End If
    
    factor = 1
    
    If LCase(InstType) = "ppmu" Then
        MaxWaitTime = pc_Def_VFI_MI_WaitTime_PPMU
    End If
        
    If SpecifyWaitTime_double > MaxWaitTime Then MaxWaitTime = SpecifyWaitTime_double
    
    If Measure.Setup_ByTypeByPin_Flag = False Then
        
        If (Measure.Setup_ByType.HexVS.Meas_Range <> "") Then
                
            If CDbl(Measure.Setup_ByType.HexVS.Meas_Range) = 0 Then
                Measure.Setup_ByType.HexVS.Meas_Range = 15
                WaitTime = 100 * us
            ElseIf CDbl(Measure.Setup_ByType.HexVS.Meas_Range) > 60 Then
                Measure.Setup_ByType.HexVS.Meas_Range = 90
                WaitTime = 100 * us
            ElseIf CDbl(Measure.Setup_ByType.HexVS.Meas_Range) > 30 Then
                Measure.Setup_ByType.HexVS.Meas_Range = 60
                WaitTime = 100 * us
            ElseIf CDbl(Measure.Setup_ByType.HexVS.Meas_Range) > 15 Then
                Measure.Setup_ByType.HexVS.Meas_Range = 30
                WaitTime = 100 * us
            ElseIf CDbl(Measure.Setup_ByType.HexVS.Meas_Range) > 1 Then
                Measure.Setup_ByType.HexVS.Meas_Range = 15
                WaitTime = 100 * us
            ElseIf CDbl(Measure.Setup_ByType.HexVS.Meas_Range) > 0.1 Then
                Measure.Setup_ByType.HexVS.Meas_Range = 1
                WaitTime = 1 * ms
            ElseIf CDbl(Measure.Setup_ByType.HexVS.Meas_Range) > 0.01 Then
                Measure.Setup_ByType.HexVS.Meas_Range = 0.1
                WaitTime = 10 * ms
            Else
                Measure.Setup_ByType.HexVS.Meas_Range = 0.01
                WaitTime = 100 * ms
            End If
            If WaitTime > MaxWaitTime Then MaxWaitTime = WaitTime
        ElseIf (Measure.Setup_ByType.UVI80.Meas_Range <> "") Then
            
'            For i = 0 To UBound(Split(Measure.Pins.UVI80, ","))
'                PinType = gl_GetInstrumentType_Dic(CStr(Split(Measure.Pins.UVI80, ",")(i)))
                PinType = gl_GetInstrumentType_Dic(LCase(SplitInputCondition(Measure.Pins.UVI80, ",", 1)))
                Select Case PinType
                    Case "DCVI"
                        factor = 1
                    Case "DCVIMerged"
                        factor = 2
                    Case Else
                End Select
                If CDbl(Measure.Setup_ByType.UVI80.Meas_Range) = 0 Then
                    Measure.Setup_ByType.UVI80.Meas_Range = 0.2
                    WaitTime = 1.5 * ms
                ElseIf CDbl(Measure.Setup_ByType.UVI80.Meas_Range) > 1 * factor Then
                    Measure.Setup_ByType.UVI80.Meas_Range = 2 * factor
                    WaitTime = 1.6 * ms
                ElseIf CDbl(Measure.Setup_ByType.UVI80.Meas_Range) > 0.2 * factor Then
                    Measure.Setup_ByType.UVI80.Meas_Range = 1 * factor
                    WaitTime = 1.6 * ms
                ElseIf CDbl(Measure.Setup_ByType.UVI80.Meas_Range) > 0.02 * factor Then
                    Measure.Setup_ByType.UVI80.Meas_Range = 0.2 * factor
                    WaitTime = 1.5 * ms
                ElseIf CDbl(Measure.Setup_ByType.UVI80.Meas_Range) > 0.002 * factor Then
                    Measure.Setup_ByType.UVI80.Meas_Range = 0.02 * factor
                    WaitTime = 1.5 * ms
                ElseIf CDbl(Measure.Setup_ByType.UVI80.Meas_Range) > 0.0002 * factor Then
                    Measure.Setup_ByType.UVI80.Meas_Range = 0.002 * factor
                    WaitTime = 3 * ms
                ElseIf CDbl(Measure.Setup_ByType.UVI80.Meas_Range) > 0.00002 * factor Then
                    Measure.Setup_ByType.UVI80.Meas_Range = 0.0002 * factor
                    WaitTime = 4 * ms
                Else
                    Measure.Setup_ByType.UVI80.Meas_Range = 0.00002
                    WaitTime = 6 * ms
                End If
                If WaitTime > MaxWaitTime Then MaxWaitTime = WaitTime
'            Next i
        ElseIf (Measure.Setup_ByType.UVS256.Meas_Range <> "") Then
            
            If CDbl(Measure.Setup_ByType.UVS256.Meas_Range) = 0 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.2
                WaitTime = 210 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 2.8 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 5.6
                WaitTime = 30 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 1.4 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 2.8
                WaitTime = 45 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.8 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 1.4
                WaitTime = 50 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.7 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.8
                WaitTime = 100 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.4 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.7
                WaitTime = 100 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.2 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.4
                WaitTime = 90 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.04 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.2
                WaitTime = 210 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.02 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.04
                WaitTime = 260 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.002 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.02
                WaitTime = 540 * us
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.0002 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.002
                WaitTime = 3.5 * ms
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.00002 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.0002
                WaitTime = 4 * ms
            ElseIf CDbl(Measure.Setup_ByType.UVS256.Meas_Range) > 0.000002 Then
                Measure.Setup_ByType.UVS256.Meas_Range = 0.00002
                WaitTime = 4 * ms
            Else
                Measure.Setup_ByType.UVS256.Meas_Range = 0.000004
                WaitTime = 4 * ms
            End If
            If WaitTime > MaxWaitTime Then MaxWaitTime = WaitTime
        ElseIf Measure.Setup_ByType.PPMU.Meas_Range <> "" Then

            If Measure.Setup_ByType.PPMU.Meas_Range > 0.002 Then
                Measure.Setup_ByType.PPMU.Meas_Range = 0.05
            ElseIf Measure.Setup_ByType.PPMU.Meas_Range > 0.0002 Then
                Measure.Setup_ByType.PPMU.Meas_Range = 0.002
            ElseIf Measure.Setup_ByType.PPMU.Meas_Range > 0.00002 Then
                Measure.Setup_ByType.PPMU.Meas_Range = 0.0002
            ElseIf Measure.Setup_ByType.PPMU.Meas_Range > 0.000002 Then
                Measure.Setup_ByType.PPMU.Meas_Range = 0.00002
            Else
                Measure.Setup_ByType.PPMU.Meas_Range = 0.000002
            End If
        ElseIf (Measure.Setup_ByType.VSM.Meas_Range <> "") Then
        
        End If
    Else
        If Measure.Setup_ByTypeByPin.HexVS_Flag = True Then
            For i = 0 To UBound(Measure.Setup_ByTypeByPin.HexVS)
                
                If CDbl(Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range) = 0 Then
                    Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range = 15
                    WaitTime = 100 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range) > 60 Then
                    Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range = 90
                    WaitTime = 100 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range) > 30 Then
                    Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range = 60
                    WaitTime = 100 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range) > 15 Then
                    Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range = 30
                    WaitTime = 100 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range) > 1 Then
                    Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range = 15
                    WaitTime = 100 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range) > 0.1 Then
                    Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range = 1
                    WaitTime = 1 * ms
                ElseIf CDbl(Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range) > 0.01 Then
                    Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range = 0.1
                    WaitTime = 10 * ms
                Else
                    Measure.Setup_ByTypeByPin.HexVS(i).Meas_Range = 0.01
                    WaitTime = 100 * ms
                End If
                If WaitTime > MaxWaitTime Then MaxWaitTime = WaitTime
            Next i
        ElseIf Measure.Setup_ByTypeByPin.UVI80_Flag = True Then
            For i = 0 To UBound(Measure.Setup_ByTypeByPin.UVI80)
                PinType = gl_GetInstrumentType_Dic(LCase(Measure.Setup_ByTypeByPin.UVI80(i).Pin))
                Select Case PinType
                    Case "DCVI"
                        factor = 1
                    Case "DCVIMerged"
                        factor = 2
                    Case Else
                End Select
                
                If CDbl(Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range) = 0 Then
                    Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range = 0.2
                    WaitTime = 1.5 * ms
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range) > 1 * factor Then
                    Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range = 2 * factor
                    WaitTime = 1.6 * ms
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range) > 0.2 * factor Then
                    Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range = 1 * factor
                    WaitTime = 1.6 * ms
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range) > 0.02 * factor Then
                    Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range = 0.2 * factor
                    WaitTime = 1.5 * ms
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range) > 0.002 * factor Then
                    Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range = 0.02 * factor
                    WaitTime = 1.5 * ms
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range) > 0.0002 * factor Then
                    Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range = 0.002 * factor
                    WaitTime = 3 * ms
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range) > 0.00002 * factor Then
                    Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range = 0.0002 * factor
                    WaitTime = 4 * ms
                Else
                    Measure.Setup_ByTypeByPin.UVI80(i).Meas_Range = 0.00002
                    WaitTime = 6 * ms
                End If
                If WaitTime > MaxWaitTime Then MaxWaitTime = WaitTime
            Next i
        ElseIf Measure.Setup_ByTypeByPin.UVS256_Flag = True Then
            For i = 0 To UBound(Measure.Setup_ByTypeByPin.UVS256)
            
                If CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) = 0 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.2
                    WaitTime = 210 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 2.8 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 5.6
                    WaitTime = 30 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 1.4 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 2.8
                    WaitTime = 45 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.8 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 1.4
                    WaitTime = 50 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.7 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.8
                    WaitTime = 100 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.4 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.7
                    WaitTime = 100 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.2 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.4
                    WaitTime = 90 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.04 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.2
                    WaitTime = 210 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.02 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.04
                    WaitTime = 260 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.002 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.02
                    WaitTime = 540 * us
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.0002 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.002
                    WaitTime = 3.5 * ms
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.00002 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.0002
                    WaitTime = 4 * ms
                ElseIf CDbl(Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range) > 0.000002 Then
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.00002
                    WaitTime = 4 * ms
                Else
                    Measure.Setup_ByTypeByPin.UVS256(i).Meas_Range = 0.000004
                    WaitTime = 4 * ms
                End If
                If WaitTime > MaxWaitTime Then MaxWaitTime = WaitTime
            Next i
        ElseIf Measure.Setup_ByTypeByPin.PPMU_Flag = True Then
            For i = 0 To UBound(Measure.Setup_ByTypeByPin.PPMU)

                If Measure.Setup_ByTypeByPin.PPMU(i).Meas_Range > 0.002 Then
                    Measure.Setup_ByTypeByPin.PPMU(i).Meas_Range = 0.05
                ElseIf Measure.Setup_ByTypeByPin.PPMU(i).Meas_Range > 0.0002 Then
                    Measure.Setup_ByTypeByPin.PPMU(i).Meas_Range = 0.002
                ElseIf Measure.Setup_ByTypeByPin.PPMU(i).Meas_Range > 0.00002 Then
                    Measure.Setup_ByTypeByPin.PPMU(i).Meas_Range = 0.0002
                ElseIf Measure.Setup_ByTypeByPin.PPMU(i).Meas_Range > 0.000002 Then
                    Measure.Setup_ByTypeByPin.PPMU(i).Meas_Range = 0.00002
                Else
                    Measure.Setup_ByTypeByPin.PPMU(i).Meas_Range = 0.000002
                End If
            Next i
            
        ElseIf Measure.Setup_ByTypeByPin.VSM_Flag = True Then
            
        End If
    End If

    Return_RangeAndMaxWaitTime = MaxWaitTime
    Exit Function
err:
    Stop
    Resume Next
End Function

Public Function HardIP_FrequencyMeasure(ByRef Rtn_MeasFreq As PinListData, Optional ByRef Rtn_SweepTestName As String)
    
    Dim MeasFreq As New PinListData
    Dim measf As MeasF_Type
    Dim CounterValue As New PinListData
    Dim Pin As New PinList
    measf = TestConditionSeqData(Instance_Data.TestSeqNum).measf(Instance_Data.TestSeqSweepNum)
    
    With TheHdw.Digital.Pins(measf.Pins).FreqCtr
        .EventSource = measf.EventSource
        .EventSlope = Positive
        .Interval = CDbl(measf.Interval)
        .Enable = IntervalEnable
        .Clear
        TheHdw.Wait CDbl(measf.WaitTime)
        .start
        CounterValue = .Read()
    End With
    
    If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Freq_meas Pin setting = " & measf.Pins)
        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Freq_meas Interval setting, " & measf.Pins & " = " & measf.Interval)
        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Freq_meas Waittime Setting, " & measf.Pins & " = " & measf.WaitTime)
    End If
    
    Rtn_MeasFreq = CounterValue.Math.Divide(CDbl(measf.Interval))
    
End Function


Public Function HardIP_FrequencyMeasure_Dctime(Rtn_MeasResult As PinListData)

    Dim pld1 As New PinListData
    Dim pld2 As New PinListData
    
    Dim measf As MeasF_Type
    
    measf = TestConditionSeqData(Instance_Data.TestSeqNum).measf(Instance_Data.TestSeqSweepNum)
    
    TheHdw.DCTime.Pins(measf.Pins).Connect

    TheHdw.DCTime.Pins(measf.Pins).mode = tlDCTimeModeStamper
    TheHdw.DCTime.Pins(measf.Pins).Interleave = tlDCTimeInterleaveOff
    TheHdw.DCTime.Pins(measf.Pins).Measurement.Frequency.start.SetInput tlDCTimeStartInputOnTrigger
    With TheHdw.DCTime.Pins(measf.Pins).Measurement.Frequency
        .SetVoltageRange 7, tlDCTimeImpedanceHiZ
        .Hysteresis = tlDCTimeHysteresisOn
        .Threshold = 1
        .SampleSize = 100
        .Slope = tlSlopePositive
    End With
    
    If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Freq_meas Pin setting = " & measf.Pins)
    End If
    
    TheHdw.DCTime.Pins(measf.Pins).TimeStamps.Trigger ("my_capture")

    TheHdw.DCTime.Pins(measf.Pins).TimeStamps("my_capture").GetWaves pld1, pld2

    Rtn_MeasResult = TheHdw.DCTime.Pins(measf.Pins).Measurement.Frequency.CalculatedAverageResults(pld1)

    TheHdw.DCTime.Pins(measf.Pins).Disconnect
    
End Function


Public Function HardIP_DCVS_MI_StoreAndRestoreCondition(Measure As Meas_Type, Instrument As Inst_Type, SaveCondition As Boolean) As Long
    
    Dim MI_Pin As String
    Dim PinName() As String
    Dim TypeName As String
    Dim i As Long
    
    
    If Instrument = HexVS Then
        MI_Pin = Measure.Pins.HexVS
    ElseIf Instrument = UVS256 Then
        MI_Pin = Measure.Pins.UVS256
    ElseIf Instrument = VSM Then ''CArter, 20190412
        MI_Pin = Measure.Pins.VSM
    End If
    
    PinName = Split(MI_Pin, ",")
    
    ReDim Preserve Measure.SaveCondition(UBound(PinName)) 'For restore

    For i = 0 To UBound(PinName)
        TypeName = SortPinInstrument(PinName(i))
        If SaveCondition = True Then
            Measure.SaveCondition(i).Pin = PinName(i)
            Measure.SaveCondition(i).SourceFlodLimit = FormatNumber(TheHdw.DCVS.Pins(PinName(i)).CurrentLimit.Source.FoldLimit.Level.Value, 3)
            Measure.SaveCondition(i).SinkFoldLimit = FormatNumber(TheHdw.DCVS.Pins(PinName(i)).CurrentLimit.Sink.FoldLimit.Level.Value, 3)
            If TypeName = "VHDVS" Then Measure.SaveCondition(i).FilterValue = FormatNumber(TheHdw.DCVS.Pins(PinName(i)).Meter.Filter.Value, 3)
            Measure.SaveCondition(i).SrcCurrentRange = FormatNumber(TheHdw.DCVS.Pins(PinName(i)).CurrentRange.Value, 3)
        Else
            TheHdw.DCVS.Pins(Measure.SaveCondition(i).Pin).CurrentRange.Value = CDbl(Measure.SaveCondition(i).SrcCurrentRange) ''Move here for avoid Flodlimit Alarm - Carter, 20190507
            TheHdw.DCVS.Pins(Measure.SaveCondition(i).Pin).CurrentLimit.Source.FoldLimit.Level.Value = CDbl(Measure.SaveCondition(i).SourceFlodLimit)
            TheHdw.DCVS.Pins(Measure.SaveCondition(i).Pin).CurrentLimit.Sink.FoldLimit.Level.Value = CDbl(Measure.SaveCondition(i).SinkFoldLimit)
            If TypeName = "VHDVS" Then TheHdw.DCVS.Pins(Measure.SaveCondition(i).Pin).Meter.Filter.Value = CDbl(Measure.SaveCondition(i).FilterValue)
        End If
    Next i
    
End Function


Public Function HardIP_BySeqCurrentProfile()
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim p As Variant
    Dim DSPW As New DSPWave
    Dim Label As String
    Dim FileName As String
    Dim IsIO As Boolean
    Dim SampleSize As Long
    Dim SampleRate As Long
    Dim MeasI As Meas_Type
    Dim Profile_Pin As String
    Dim i As Long
    Dim MeasureI_Pin_CurrentRange As String
    
    MeasI = TestConditionSeqData(Instance_Data.TestSeqNum).MeasI(Instance_Data.TestSeqSweepNum)
    Profile_Pin = MeasI.Pins.HexVS & "," & MeasI.Pins.UVS256 & "," & MeasI.Pins.UVI80
    Profile_Pin = Replace(Profile_Pin, ",,", ",")
    'Profile_Pin = Replace(Profile_Pin, ",", "", Len(Profile_Pin), 1)
    If Right(Profile_Pin, 1) = "," Then Profile_Pin = Left(Profile_Pin, Len(Profile_Pin) - 1)
    If Left(Profile_Pin, 1) = "," Then Profile_Pin = Right(Profile_Pin, Len(Profile_Pin) - 1)
    
    If InStr(LCase(TheExec.DataManager.instanceName), "lpro") <> 0 Then
        SampleSize = 16000
        SampleRate = 12500
    Else
        SampleSize = 16000
        SampleRate = 25000
    End If
    
    TheExec.Datalog.WriteComment "[Current_Profile] " & TheExec.DataManager.instanceName
    
    Call TheExec.DataManager.DecomposePinList(Profile_Pin, Pin_Ary(), Pin_Cnt)
    IsIO = False
    For Each p In Pin_Ary
        
        If gl_GetInstrumentType_Dic(p) = "N/C" Or gl_GetInstrumentType_Dic(p) = "I/O" Then IsIO = True
    Next p
    
    If IsIO = False Then
        For i = 0 To Pin_Cnt - 1
            Do While TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.IsRunning = True
            Loop

            TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.Signals.Add Pin_Ary(i) & "_Signal"
            TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.Signals.DefaultSignal = Pin_Ary(i) & "_Signal"
            
            With TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.Signals.Item(Pin_Ary(i) & "_Signal")
                .Reinitialize
                .mode = tlDCVSMeterCurrent
                If MeasureI_Pin_CurrentRange = "" Then
                    .range = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max
                ElseIf i <= UBound(Split(MeasureI_Pin_CurrentRange, ",")) Then
                    .range = CDbl(Split(MeasureI_Pin_CurrentRange, ",")(i))
                    If CDbl(Split(MeasureI_Pin_CurrentRange, ",")(i)) < 0.2 Then
                        TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange = CDbl(Split(MeasureI_Pin_CurrentRange, ",")(i)) * 10
                    Else
                        TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange = CDbl(Split(MeasureI_Pin_CurrentRange, ",")(i))
                    End If
                Else
                    .range = CDbl(Split(MeasureI_Pin_CurrentRange, ",")(0))
                    If CDbl(Split(MeasureI_Pin_CurrentRange, ",")(i)) < 0.2 Then
                        TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange = CDbl(Split(MeasureI_Pin_CurrentRange, ",")(0)) * 10
                    Else
                        TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange = CDbl(Split(MeasureI_Pin_CurrentRange, ",")(0))
                    End If
                End If
                .SampleRate = SampleRate
                .SampleSize = SampleSize
            End With
            TheHdw.DCVS.Pins(Pin_Ary(i)).Capture.Signals.Item(Pin_Ary(i) & "_Signal").LoadSettings
        Next i

        For Each p In Pin_Ary
            TheHdw.DCVS.Pins(p).Capture.Signals.Item(p & "_Signal").Trigger
        Next p

'        Do While thehdw.DCVS.Pins(MeasureI_pin).Capture.IsRunning = True
'        Loop
    
        ' Get the captured samples from the instrument
        
        Dim sampleR As String
        Dim inst_ary() As String
        inst_ary = Split(TheExec.DataManager.instanceName, "_")
        
        For Each p In Pin_Ary
            If gl_GetInstrumentType_Dic(p) <> "N/C" Then
                
                DSPW = TheHdw.DCVS.Pins(p).Capture.Signals(p & "_Signal").DSPWave
                
                For Each site In TheExec.sites
                    sampleR = CStr(TheHdw.DCVS.Pins(p).Capture.SampleRate)
                    'If thehdw.DCVS.Pins(p).Meter.mode = thehdw.DCVS.Pins(p).CurrentRange.Max Then

                    Label = "Current Profile for Site: " & site & " " & " " & p & "_Signal" & "Pin :" & " " & p
                    FileName = "CurrentProfile" & "-Site" & site & "_" & inst_ary(0) & "_" & inst_ary(1) & "_" & p & "_" & Instance_Data.TestSeqNum & "-" & p & "-" & sampleR & "-" & TheExec.DataManager.instanceName & "_" & TheHdw.DCVS.Pins(p).Capture.Signals.Item(p & "_Signal").range & ".txt"
                    
'                                            If LCase(p) = LCase("VDD_FIXED_PCIE_REFBUF") Then
'                                                FileName = Replace(FileName, "_", "")
'                                            End If
                    
                    If TheExec.TesterMode = testModeOffline Then
                    
                        Call DSPW.CreateRandom(TheHdw.DCVS.Pins(p).CurrentRange * 0.9, TheHdw.DCVS.Pins(p).CurrentRange * 1.1, SampleSize, 0, DspDouble)
                        
                        'DSPW.Plot Label
                    End If
                    'DSPW.Plot Label
'                                        If True Then DSPW.Plot Label   'for pliot
                    If True Then
                        Dim TempStr As String
                        TempStr = "D:\" & p
                        Dim fso As New FileSystemObject
                         
                        If Dir(TempStr, vbDirectory) = Empty Then
                            MkDir TempStr
                        End If

                        DSPW.FileExport TempStr & "\" & FileName, File_txt
                    End If
                    If LCase(gl_GetInstrument_Dic(CStr(p))) <> "hexvs" Then DSPW.Clear
                Next site
            End If
        Next p
    End If
End Function

Public Function RegExCheckStringIsNumber(inString As String) As Boolean
    Dim RegEx As Object
    Set RegEx = CreateObject("vbscript.regexp")
    With RegEx
        .Global = True
        .Pattern = "\D"         'none number
        .IgnoreCase = True
        RegExCheckStringIsNumber = Not (.Test(inString))
        
'        If Not (RegExPattern = "\D") Then RegExCheckStringIsNumber = Not RegExCheckStringIsNumber
        
    End With
End Function

Public Function CheckDigSrcEquationAssignment(SampleSize As Long, DataWidth As Long, Equation As String, assignment As String, ByRef Equation_Size() As Long) As Boolean

    Dim SplitEquationAry() As String
    Dim SplitAssignmentAry() As String
    Dim splitbyand() As String
    Dim Str As Variant
    Dim Dict_Assignment As New Dictionary
    Dim dict_key As String
    Dim dict_value As String
    Dim DspWav As DSPWave
    Dim CheckResult As Boolean: CheckResult = True
    
    Dim Calc_Sample_Size As Long
    Dim IsRepeat As Boolean
    Dim str2 As Variant
    Dim index As Long
    Dim temp_Calc_Sample_Size As Long
    Dim site As Variant
    Dim TestInstanceName As String
    
    On Error GoTo err
    Equation = LCase(Equation)
    assignment = LCase(assignment)
    TestInstanceName = TheExec.DataManager.instanceName
    
    '------------------dummy, debug
    Dim temp_dw As New DSPWave


    SplitEquationAry = Split(Equation, "+")
    SplitAssignmentAry = Split(assignment, ";")
    ReDim Equation_Size(UBound(SplitEquationAry))
'==============================check Equation,Assignment Empty
    If gl_Disable_HIP_debug_log = False Then
        If Equation = "" Then
            TheExec.AddOutput "DigSrc Equation is Empty in " & TestInstanceName, vbRed, True
            CheckDigSrcEquationAssignment = False
            'MsgBox "DigSrc Equation is Empty"
            TheExec.Datalog.WriteComment "DigSrc Equation is Empty"
            Exit Function
        ElseIf assignment = "" Then
            TheExec.AddOutput "DigSrc Assignment is Empty in " & TestInstanceName, vbRed, True
            CheckDigSrcEquationAssignment = False
            'MsgBox "DigSrc Assignment is Empty"
            TheExec.Datalog.WriteComment "DigSrc Assignment is Empty"
            Exit Function
        End If
    End If
'==============================check equation need to store dspwave
    If InStr(Equation, "(") <> 0 Then
        Equation = Split(Split(Equation, "(")(1), ")")(0)
    End If
'==============================check assignment Reproduce
    If gl_Disable_HIP_debug_log = False Then
        For Each Str In SplitAssignmentAry
            Dim AssignmentName As String
            Dim AssignmentValue As String
            If InStr(Str, "=") <> 0 Then
                AssignmentName = Split(Str, "=")(0)
                AssignmentValue = Split(Str, "=")(1)
            Else
                AssignmentName = "LastAssignmentName"
                AssignmentValue = Str
            End If
            
            If Dict_Assignment.Exists(AssignmentName) Then
                TheExec.AddOutput "DigSrc Assignment [" & AssignmentName & "] Reproduce in " & TestInstanceName, vbRed, True
                CheckResult = CheckResult And False
                'MsgBox "DigSrc Assignment [" & AssignmentName & "] Reproduce"
                TheExec.Datalog.WriteComment "DigSrc Assignment [" & AssignmentName & "] Reproduce"
            Else
                If Not Dict_Assignment.Exists(AssignmentName) Then Dict_Assignment.Add AssignmentName, AssignmentValue
            End If
            
        Next Str
    End If
'==============================check assignment dictionary exist
    If gl_Disable_HIP_debug_log = False Then
        For Each Str In Dict_Assignment.Keys
            dict_key = Str
            dict_value = Dict_Assignment(dict_key)
            If Not RegExCheckStringIsNumber(dict_value) Then
                If Not gDictDSPWaves.Exists(dict_value) Then
                    TheExec.AddOutput "DigSrc Assignment Dictionary Name [" & dict_value & "] Not Exist in " & TestInstanceName, vbRed, True
                    CheckResult = CheckResult And False
                    'MsgBox "DigSrc Assignment Dictionary Name [" & dict_value & "] Not Exist"
                    TheExec.Datalog.WriteComment "DigSrc Assignment Dictionary Name [" & dict_value & "] Not Exist"
                End If
            End If
        Next Str
    End If
'==============================check equation in assignment
    If gl_Disable_HIP_debug_log = False Then
        If InStr(LCase(assignment), "repeat") = 0 Then
            For Each Str In SplitEquationAry
                If Not Dict_Assignment.Exists(Str) Then
                    TheExec.AddOutput "DigSrc Equation [" & Str & "] Not Exist in Assignment in " & TestInstanceName, vbRed, True
                    CheckResult = CheckResult And False
                    'MsgBox "DigSrc Equation [" & str & "] Not Exist in Assignment"
                    TheExec.Datalog.WriteComment "DigSrc Equation [" & Str & "] Not Exist in Assignment"
                End If
            Next Str
        End If
    End If
'==============================check sample size
    Dim CheckSampleSizeDspWave As New DSPWave
    If gl_Disable_HIP_debug_log = False Then
        If CheckResult = True Then
            temp_Calc_Sample_Size = 0
            Calc_Sample_Size = 0
            index = 0
            
            If InStr(LCase(assignment), "repeat") <> 0 Then
                CheckSampleSizeDspWave = gDictDSPWaves(Split(assignment, "=")(1))
                For Each site In TheExec.sites.Selected
                    temp_Calc_Sample_Size = CheckSampleSizeDspWave.SampleSize
                    Exit For
                Next site
                If SampleSize Mod temp_Calc_Sample_Size = 0 Then
                    Calc_Sample_Size = SampleSize
                End If
            Else
                For Each Str In SplitEquationAry
                    
                    temp_Calc_Sample_Size = Calc_Sample_Size
                    If RegExCheckStringIsNumber(Dict_Assignment(Str)) Then
                        'number only
                        Calc_Sample_Size = Calc_Sample_Size + Len(Dict_Assignment(Str))
                    Else
                        If InStr(Dict_Assignment(Str), "&") <> 0 Then
                            splitbyand = Split(Dict_Assignment(Str), "&")
                            For Each str2 In splitbyand
                                If RegExCheckStringIsNumber(CStr(str2)) Then
                                    Calc_Sample_Size = Calc_Sample_Size + Len(str2)
                                Else
                                    Calc_Sample_Size = Calc_Sample_Size + gDictDSPWaves(str2).SampleSize
                                End If
                            Next str2
                        ElseIf InStr(Dict_Assignment(Str), ":") <> 0 Then
                            splitbyand = Split(Dict_Assignment(Str), ":")
                            Calc_Sample_Size = Calc_Sample_Size + (splitbyand(1) - splitbyand(2) + 1)
                        Else
                            CheckSampleSizeDspWave = gDictDSPWaves(Dict_Assignment(Str))
                            For Each site In TheExec.sites.Active
                                Calc_Sample_Size = Calc_Sample_Size + CheckSampleSizeDspWave.SampleSize
                                Exit For
                            Next site
                        End If
                    End If
                    Equation_Size(index) = Calc_Sample_Size - temp_Calc_Sample_Size
                    index = index + 1
                Next Str
            End If
            
            If SampleSize <> Calc_Sample_Size Then
                TheExec.AddOutput "DigSrc Sample Size not match in " & TheExec.DataManager.instanceName, vbRed, True
                CheckResult = CheckResult And False
                'MsgBox "DigSrc Sample Size not match"
                TheExec.Datalog.WriteComment "DigSrc Sample Size not match"
                Stop
            End If
        End If
    End If
    CheckDigSrcEquationAssignment = CheckResult
    Exit Function
err:
    Stop
    Resume Next
End Function

Public Function CreateDigSrcDSPWave(MergeDigSrc_Equation As String, ByRef InDSPwave As DSPWave, sample_size As Long) As String

    Dim SplitAry() As String
    Dim DspWaveData() As Long
    Dim Temp_DspWaveData() As Long
    Dim temp_string As String
    Dim Str As Variant
    Dim index As Long
    Dim site As Variant
    Dim temp_dspwave As DSPWave
    Dim TempDictionary As New Dictionary
    
    Dim MyDspWave As New DSPWave
    Dim TempWave As New DSPWave
    Dim Dummy1 As New SiteLong
    Dim Dummy2 As New SiteLong

    On Error GoTo err
    
    InDSPwave.CreateConstant 0, 0

    If InStr(LCase(MergeDigSrc_Equation), "repeat") <> 0 Then
        temp_string = Split(MergeDigSrc_Equation, "=")(1)
        If gDictDSPWaves.Exists(temp_string) Then
            Set temp_dspwave = gDictDSPWaves(temp_string)
            rundsp.DspWaveMergeRepeat InDSPwave, temp_dspwave, sample_size
        Else
        
        End If
        
        Exit Function
    End If
    
    
    ReDim DspWaveData(sample_size - 1)
    MergeDigSrc_Equation = LCase(MergeDigSrc_Equation)
    temp_string = Replace(MergeDigSrc_Equation, "+", "")
    If RegExCheckStringIsNumber(temp_string) Then
        For index = 0 To UBound(DspWaveData)
            DspWaveData(index) = CLng(Mid(temp_string, index + 1, 1))
        Next index
        For Each site In TheExec.sites.Active
            InDSPwave.Data = DspWaveData
        Next site
    Else
        SplitAry = Split(MergeDigSrc_Equation, "+")
        For Each Str In SplitAry
            Set temp_dspwave = New DSPWave
            If RegExCheckStringIsNumber(CStr(Str)) And Not TempDictionary.Exists(Str) Then
                ReDim Temp_DspWaveData(Len(Str) - 1)
                
                For index = 0 To UBound(Temp_DspWaveData)
                    Temp_DspWaveData(index) = CLng(Mid(Str, index + 1, 1))
                Next index
                
                For Each site In TheExec.sites.Active
                    temp_dspwave.Data = Temp_DspWaveData
                Next site
                TempDictionary.Add Str, temp_dspwave
            Else
                If TempDictionary.Exists(Str) Then
                    temp_dspwave = TempDictionary(Str)
                Else
                    temp_dspwave = gDictDSPWaves(LCase(Str))
                End If
            End If
            
'            Dim InDspWave_temp As DSPWave
            
'            For Each site In TheExec.sites.Active
'                If InDspWave.SampleSize = 0 Then
'
'                    InDspWave = temp_dspwave.Copy
'                Else
'                    InDspWave = InDspWave.ConvertDataTypeTo(DspLong)
'                    temp_dspwave = temp_dspwave.ConvertDataTypeTo(DspLong)
'
'                    InDspWave = InDspWave.Concatenate(temp_dspwave)
'
'                End If
'            Next site

            'thehdw.DSP.ExecutionMode = tlDSPModeForceAutomatic
                                                                                                                                                                Dim LIB_HardIP_ProfileMark_14735 As Long: LIB_HardIP_ProfileMark_14735 = ProfileMarkEnter(2, TheExec.DataManager.instanceName & "_" & "DigSrcDSP&Module=LIB_HardIP&ProcName=CreateDigSrcDSPWave&LineNumber=14731")    ' Profile Mark
            
            rundsp.DSPWf_Concatenate InDSPwave, temp_dspwave, Dummy1
            
            Dummy2 = Dummy1
                                                                                                                                                                ProfileMarkLeave LIB_HardIP_ProfileMark_14735    ' Profile Mark

        Next Str
        

    End If
    
    Exit Function
err:
    'Stop
    
    'Resume Next
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MergeDigSrcEquationAssignment(SampleSize As Long, DataWidth As Long, Equation As String, assignment As String, ByRef StoreDictName As String) As String
    Dim SplitEquationAry() As String
    Dim SplitAssignmentAry() As String
    Dim splitbyand() As String
    Dim Str As Variant
    Dim Dict_Assignment As New Dictionary
    Dim Dict_Equation As New Dictionary
    Dim dict_key As String
    Dim dict_value As String
    Dim DspWav As DSPWave
    Dim IsRepeat As Boolean
    Dim str2 As Variant
    Dim Assignment_temp As String
    Dim index As Long
    
    On Error GoTo err
    SplitEquationAry = Split(Equation, "+")
    SplitAssignmentAry = Split(assignment, ";")
    
    If InStr(assignment, "=") = 0 Then
        Exit Function
    End If
    
    If InStr(Equation, "(") <> 0 Then
        StoreDictName = Split(Equation, "(")(0)
        Equation = Split(Split(Equation, "(")(1), ")")(0)
    Else
        StoreDictName = ""
    End If
    
    
    If UBound(SplitEquationAry) > 0 And UBound(SplitAssignmentAry) = 0 Then
        For Each Str In SplitEquationAry
            If Not Dict_Equation.Exists(Str) Then
                Dict_Equation.Add Str, ""
            End If
        Next Str
        If Dict_Equation.Count = 1 Then
            assignment = Replace(assignment, SplitEquationAry(0), "reapeat")
        End If
    End If
    
    If InStr(LCase(assignment), "repeat") <> 0 Then
        MergeDigSrcEquationAssignment = LCase(assignment)
        Exit Function
    End If
    
    
    
    For Each Str In SplitAssignmentAry
        Call Dict_Assignment.Add(Split(Str, "=")(0), Split(Str, "=")(1))
    Next Str
    
    If InStr(LCase(assignment), "repeat") <> 0 Then
        Assignment_temp = Split(SplitAssignmentAry(0), "=")(1)
        Assignment_temp = Replace(Assignment_temp, "&", "+")
        For index = 0 To UBound(SplitEquationAry)
            SplitEquationAry(index) = Assignment_temp
        Next index
    Else
        For index = 0 To UBound(SplitEquationAry)
            Assignment_temp = Dict_Assignment(SplitEquationAry(index))
            Assignment_temp = Replace(Assignment_temp, "&", "+")
            SplitEquationAry(index) = Assignment_temp
        Next index
        
    End If
    
    MergeDigSrcEquationAssignment = Join(SplitEquationAry, "+")
    
    Exit Function
err:
    Stop
    Resume Next
End Function

Public Function PrintDigSrcEquationAssignment(InDSPwave As DSPWave, EquationSampleSizeAry() As Long, Equation As String, assignment As String, DigSrc_pin As String, DigSrc_Sample_Size, Optional MSB_First_Flag As Boolean = False) As Long
    
    Dim site As Variant
    Dim temp_string As Variant
    Dim dspwave_index As Long
    Dim index, index2 As Long
    Dim assignment_Dict As New Dictionary
    Dim SplitEquationAry() As String: SplitEquationAry = Split(Equation, "+")
    Dim SplitAssignmentAry() As String: SplitAssignmentAry = Split(assignment, ";")
    Dim EquationName As String
    Dim EquationValue As String
    Dim TotalEquationValue As String
    
    On Error GoTo err
'    If gl_Disable_HIP_debug_log = False Then
    Equation = LCase(Equation)
    assignment = LCase(assignment)
    
    TheExec.Datalog.WriteComment ("======== Setup Dig Src Test Start ========")
    TheExec.Datalog.WriteComment "Src Bits = " & DigSrc_Sample_Size
    TheExec.Datalog.WriteComment "SrcPin = " & DigSrc_pin
    TheExec.Datalog.WriteComment "DataSequence:" & Equation
    TheExec.Datalog.WriteComment "Assignment:" & assignment
    
    If MSB_First_Flag Then
        TheExec.Datalog.WriteComment "Output String [ MSB(L) ==> LSB(R) ]:"
    Else
        TheExec.Datalog.WriteComment "Output String [ LSB(L) ==> MSB(R) ]:"
    End If
    
    For Each site In TheExec.sites.Active
        dspwave_index = 0
        assignment_Dict.RemoveAll
        TotalEquationValue = ""
        For index = 0 To UBound(EquationSampleSizeAry)
            
            EquationName = SplitEquationAry(index)
            EquationValue = ""
            
            For index2 = dspwave_index To EquationSampleSizeAry(index) - 1
                EquationValue = EquationValue & CStr(InDSPwave.Element(index2))
            Next index2
            
            If Not assignment_Dict.Exists(EquationName) Then
                assignment_Dict.Add EquationName, EquationValue
            End If
            
            If index Mod 5 = 0 Then TotalEquationValue = TotalEquationValue & Chr(10)
            
            TotalEquationValue = TotalEquationValue & " " & EquationValue & "(" & EquationName & ")"
            
        Next index
        TheExec.Datalog.WriteComment "Site " & site & " : " & TotalEquationValue
'        For Each temp_string In assignment_Dict.Keys
'            TheExec.DataLog.WriteComment temp_string & "(" & assignment_Dict(temp_string) & ")"
'        Next temp_string
        
    Next site
    Exit Function
err:
    Stop
    Resume Next
    
End Function

Public Function HardIP_RAK_Init() As Long
    '' 20151029 - Compensate DIB impedence for RAK
    Dim i As Long
    Dim j As Long
    Dim ws_def_CP As Worksheet
    Dim ws_def_FT As Worksheet
    Dim ws_def_WLFT1 As Worksheet
    Dim wb As Workbook
    Dim SiteIndex As Integer
    Dim Start_SiteNum As Long
    Dim Stop_SiteNum As Long
                                                                                                                                                        
    Set wb = Application.ActiveWorkbook
    Set ws_def_CP = wb.Sheets("RAK_CP")
    Set ws_def_FT = wb.Sheets("RAK_FT2")
    Set ws_def_WLFT1 = wb.Sheets("RAK_WLFT1")
                                                                                                                                                        
''''''ChannelMap_[CP/FT]_X[1/2]_Site[0_2]
                                                                                                                                                        
        Start_SiteNum = 0
        Stop_SiteNum = TheExec.sites.Existing.Count - 1
                                                                                                                                                        
                                                                                                                                                        
    Dim CP_PinsNum As Long
    Dim FT_PinsNum As Long
    Dim WLFT1_PinsNum As Long
                                                                                                                                                        
    Dim StartRows As Long
    Dim EndRows As Long
    Dim RakV() As Double
    Dim Factor_Idx As Double
                                                                                                                                                        
    CP_PinsNum = ws_def_CP.Cells(Rows.Count, 1).End(xlUp).row - 1
    FT_PinsNum = ws_def_FT.Cells(Rows.Count, 1).End(xlUp).row - 1
    WLFT1_PinsNum = ws_def_WLFT1.Cells(Rows.Count, 1).End(xlUp).row - 1
                                                                                                                                                        
                                                                                                                                                        
     If ws_def_CP.Cells(2, 2).Value > 100 Then
        Factor_Idx = 0.001
     Else
        Factor_Idx = 1
     End If
                                                                                                                                                        
    If CP_Card_RAK.Pins.Count <> CP_PinsNum And InStr(UCase(TheExec.CurrentChanMap), "CP") <> 0 Then
        Set CurrentJob_Card_RAK = Nothing
        Set CP_Card_RAK = Nothing
        
        StartRows = 2
        EndRows = CP_PinsNum + 1
        For i = StartRows To EndRows
         CP_Card_RAK.AddPin (ws_def_CP.Cells(i, 1).Value)
         SiteIndex = 0
         If InStr(ws_def_CP.Cells(i, 1), "_SENSE") = 0 And TheExec.DataManager.ChannelType(ws_def_CP.Cells(i, 1)) <> "N/C" Then
            If TheExec.DataManager.PinType(ws_def_CP.Cells(i, 1).Value) = "I/O" Then
                RakV = TheHdw.PPMU.ReadRakValuesByPinnames(ws_def_CP.Cells(i, 1).Value, -1)
            Else
                ReDim RakV(Stop_SiteNum - Start_SiteNum)
            End If
         End If

         For j = Start_SiteNum To Stop_SiteNum
               If UBound(RakV) > 0 Then
                   CP_Card_RAK.Pins(ws_def_CP.Cells(i, 1).Value).Value(SiteIndex) = RakV(SiteIndex)
               Else
                   CP_Card_RAK.Pins(ws_def_CP.Cells(i, 1).Value).Value(SiteIndex) = RakV(0)
               End If
               SiteIndex = SiteIndex + 1
        Next j
        Next i
    
        SiteIndex = 0
                                                                                                                                                        
        For j = Start_SiteNum To Stop_SiteNum
            For i = StartRows To EndRows
                CP_Card_RAK.Pins(ws_def_CP.Cells(i, 1).Value).Value(SiteIndex) = ws_def_CP.Cells(i, 2 + j).Value * Factor_Idx + CP_Card_RAK.Pins(ws_def_CP.Cells(i, 1).Value).Value(SiteIndex)
            Next i
            SiteIndex = SiteIndex + 1
        Next j
        
        'CurrentJob_Card_RAK = CP_Card_RAK
        Set CurrentJob_Card_RAK = CP_Card_RAK
                                                                                                                                                        
    ElseIf FT_Card_RAK.Pins.Count <> FT_PinsNum And InStr(UCase(TheExec.CurrentChanMap), "FT2") <> 0 Then
        Set CurrentJob_Card_RAK = Nothing
        Set FT_Card_RAK = Nothing
        
        StartRows = 2
        EndRows = FT_PinsNum + 1
                                                                                                                                                        
        For i = StartRows To EndRows
            FT_Card_RAK.AddPin (ws_def_FT.Cells(i, 1).Value)
         SiteIndex = 0
         If InStr(ws_def_FT.Cells(i, 1), "_SENSE") = 0 And TheExec.DataManager.ChannelType(ws_def_FT.Cells(i, 1)) <> "N/C" Then
            If TheExec.DataManager.PinType(ws_def_FT.Cells(i, 1).Value) = "I/O" Then
                RakV = TheHdw.PPMU.ReadRakValuesByPinnames(ws_def_FT.Cells(i, 1).Value, -1)
            Else
                ReDim RakV(Stop_SiteNum - Start_SiteNum)
            End If
         End If

           For j = Start_SiteNum To Stop_SiteNum
               If UBound(RakV) > 0 Then
                   FT_Card_RAK.Pins(ws_def_FT.Cells(i, 1).Value).Value(SiteIndex) = RakV(SiteIndex)
                Else
                    FT_Card_RAK.Pins(ws_def_FT.Cells(i, 1).Value).Value(SiteIndex) = RakV(0)
                End If
            SiteIndex = SiteIndex + 1
           Next j

        Next i
                                                                                                                                                        
        SiteIndex = 0
                                                                                                                                                        
        For j = Start_SiteNum To Stop_SiteNum
            For i = StartRows To EndRows
                                                                         
                FT_Card_RAK.Pins(ws_def_FT.Cells(i, 1).Value).Value(SiteIndex) = ws_def_FT.Cells(i, 2 + j).Value * Factor_Idx + FT_Card_RAK.Pins(ws_def_FT.Cells(i, 1).Value).Value(SiteIndex)

                                                                                                                                                        
            Next i
            SiteIndex = SiteIndex + 1
        Next j
        
        'CurrentJob_Card_RAK = FT_Card_RAK
        Set CurrentJob_Card_RAK = FT_Card_RAK
                                                                                                                                                        
    ElseIf WLFT1_Card_RAK.Pins.Count <> WLFT1_PinsNum And InStr(UCase(TheExec.CurrentChanMap), "FT1") <> 0 Then
        Set CurrentJob_Card_RAK = Nothing
        Set WLFT1_Card_RAK = Nothing
        
        StartRows = 2
        EndRows = WLFT1_PinsNum + 1
                                                                                                                                                        
        For i = StartRows To EndRows
            WLFT1_Card_RAK.AddPin (ws_def_WLFT1.Cells(i, 1).Value)
         SiteIndex = 0
         If InStr(ws_def_WLFT1.Cells(i, 1), "_SENSE") = 0 And TheExec.DataManager.ChannelType(ws_def_WLFT1.Cells(i, 1)) <> "N/C" Then
            If TheExec.DataManager.PinType(ws_def_WLFT1.Cells(i, 1).Value) = "I/O" Then
                RakV = TheHdw.PPMU.ReadRakValuesByPinnames(ws_def_WLFT1.Cells(i, 1).Value, -1)
            Else
                ReDim RakV(Stop_SiteNum - Start_SiteNum)
            End If
         End If

           For j = Start_SiteNum To Stop_SiteNum
            If UBound(RakV) = Stop_SiteNum Then 'modfiy for turks WLFT 190821 by CW
            'If UBound(RakV) > 0 Then
                WLFT1_Card_RAK.Pins(ws_def_WLFT1.Cells(i, 1).Value).Value(SiteIndex) = RakV(SiteIndex)
            Else
                WLFT1_Card_RAK.Pins(ws_def_WLFT1.Cells(i, 1).Value).Value(SiteIndex) = RakV(0)
            End If
            SiteIndex = SiteIndex + 1
           Next j

        Next i
                                                                                                                                                        
        SiteIndex = 0
                                                                                                                                                        
        For j = Start_SiteNum To Stop_SiteNum
            For i = StartRows To EndRows
                WLFT1_Card_RAK.Pins(ws_def_WLFT1.Cells(i, 1).Value).Value(SiteIndex) = ws_def_WLFT1.Cells(i, 2 + j).Value * Factor_Idx + WLFT1_Card_RAK.Pins(ws_def_WLFT1.Cells(i, 1).Value).Value(SiteIndex)
            Next i
            SiteIndex = SiteIndex + 1
        Next j
        
        'CurrentJob_Card_RAK = WLFT1_Card_RAK
        Set CurrentJob_Card_RAK = WLFT1_Card_RAK
        
    End If
End Function


Public Function UVI80_DIFFMETER_INIT(MeasV_Pins As String, HighSidePin As String, LowSidePin As String) As Long
Dim MeasVPin() As String
Dim PinCount As Long
Dim H_Pin As String
Dim L_Pin As String
Dim i, j As Integer
Dim ReplacePin As String
Dim ThisPinType As String
Dim DiffVMGroup() As String

HighSidePin = ""
LowSidePin = ""
'PinGroup = ""

DiffVMGroup = Split(MeasV_Pins, ",")

For j = 0 To UBound(DiffVMGroup)
Call TheExec.DataManager.DecomposePinList(DiffVMGroup(j), MeasVPin, PinCount)

        For i = 0 To PinCount - 1
            If InStr(1, MeasVPin(i), "_P") <> 0 Then
                H_Pin = MeasVPin(i)
                If HighSidePin = "" Then
                    HighSidePin = MeasVPin(i)
                Else
                    HighSidePin = HighSidePin & "," & MeasVPin(i)
                End If
    '            ThisPinType = GetInstrument(MeasVPin(i), 0)
    '            MeasVPin(i) = Replace(MeasVPin(i), "_DIFFVMETER_H", "", 1)
        
            ElseIf InStr(1, MeasVPin(i), "_N") <> 0 Then
                L_Pin = MeasVPin(i)
                If LowSidePin = "" Then
                LowSidePin = MeasVPin(i)
            Else
                LowSidePin = LowSidePin & "," & MeasVPin(i)
            End If
'            ThisPinType = GetInstrument(MeasVPin(i), 0)
'            MeasVPin(i) = Replace(MeasVPin(i), "_DIFFVMETER_L", "", 1)
        
        End If
    Next i
'    If UBound(MeasVPin) > 0 Then
'        If MeasVPin(0) = MeasVPin(1) Then
'            PinGroup = MeasVPin(0)
'        Else
'            PinGroup = Join(MeasVPin, ",")
'        End If
'    Else
'        PinGroup = Join(MeasVPin, ",")
'    End If
'    PinGroup = MeasV_Pins ''Carter, 20190408
'
'    If ReplacePin = "" Then ReplacePin = PinGroup Else ReplacePin = ReplacePin & "," & PinGroup
Next j

'    UVI80DIFFMETER_H = HighSidePin
'    UVI80DIFFMETER_L = LowSidePin
'    MeasV_Pins = ReplacePin

End Function

Public Function UVI80_DIFFMETER_SETUP(HighSidePin As String, LowSidePin As String, Optional V_range As Double = 1.4, Optional H_average As Double = 32) As Long
Dim H_Pin() As String
Dim L_Pin() As String
Dim i As Integer
Dim ReplacePin As String
Dim ThisPinType As String
Dim DiffVMGroup() As String
Dim PinCount_H As Long
Dim PinCount_L As Long
Call TheExec.DataManager.DecomposePinList(HighSidePin, H_Pin, PinCount_H)
Call TheExec.DataManager.DecomposePinList(LowSidePin, L_Pin, PinCount_L)

If PinCount_H = PinCount_L Then
    For i = 0 To PinCount_H - 1
        With TheHdw.DCDiffMeter.Pins(H_Pin(i))
            .VoltageRange = V_range
            .LowSide.Pins = (L_Pin(i))
            .HardwareAverage = H_average
            .MeterMode = tlDCDiffMeterModeHighAccuracy '(24bit mode)
            'TheHdw.Wait (0.05)
            .Connect tlDCDiffMeterConnectDefault
            'TheHdw.Wait (0.05)
        End With
    Next i
Else
    For i = 0 To PinCount_H - 1
        LowSidePin = Mid(H_Pin(i), 1, InStr(H_Pin(i), "DIFFVMETER_H") - 1) & "DGS"
        
        With TheHdw.DCDiffMeter.Pins(H_Pin(i))
            .VoltageRange = V_range
            .LowSide.Pins = LowSidePin ''ADC_DIGTST0_UVI80
            .HardwareAverage = H_average
            .MeterMode = tlDCDiffMeterModeHighAccuracy '(24bit mode)
            'TheHdw.Wait (0.05)
            .Connect tlDCDiffMeterConnectDefault
            'TheHdw.Wait (0.05)
        End With
    Next i
    If PinCount_L > 0 Then TheExec.Datalog.WriteComment "The number of HighSidePin does not align with the LowSidePin"
End If

End Function


' measure VDM with error-free per DAC steps
Function UVI80_DCDIFFMETER_SPOTCAL(pin_H As String, pin_L As String, MeasureVolt As PinListData, Optional V_range As Double = 1.4, Optional H_average As Double = 32) As Long
        
        Dim CMError As New PinListData
        Dim measNorm As New PinListData
        Dim measHigh As New PinListData
        Dim measLow As New PinListData
        Dim samples As Long: samples = 1
        
        
        'do spotcal @ 0.7V common mode
        With TheHdw.DCDiffMeter.Pins(pin_H)
            .MeterMode = tlDCDiffMeterModeHighAccuracy
            .VoltageRange = V_range * v
            .HardwareAverage = H_average
        End With
        
        'notes: Cannot program the LowSide pin if the DCDiffMeter is already connected.
        
        '______ norm meas ______
        With TheHdw.DCDiffMeter.Pins(pin_H)
            
            .Connect
            TheHdw.Wait 0.2 * ms
            measNorm = .Read(tlStrobe, samples, 100 * khz, tlDCDiffMeterReadingFormatAverage)
            .Disconnect
        End With
      
        '______ high meas ______
        With TheHdw.DCDiffMeter.Pins(pin_H)
            .LowSide = pin_H
            .Connect
            TheHdw.Wait 0.2 * ms
            measHigh = .Read(tlStrobe, samples, 100 * khz, tlDCDiffMeterReadingFormatAverage)
            .Disconnect
        End With
        
        '______ low meas ______
        With TheHdw.DCDiffMeter.Pins(pin_L)
            .LowSide = pin_L
            .Connect
            TheHdw.Wait 0.2 * ms
            measLow = .Read(tlStrobe, samples, 100 * khz, tlDCDiffMeterReadingFormatAverage)
            .Disconnect
        End With
        

        ' restore high/low side connection
        With TheHdw.DCDiffMeter.Pins(pin_H)
            .LowSide = pin_L
            .Connect
        End With



    If TheExec.TesterMode = testModeOffline Then
         For Each site In TheExec.sites.Selected
                 measHigh.Pins(0) = 1.1 * Rnd()
                 measLow.Pins(0) = 1.05 * Rnd()
                 measNorm.Pins(0) = 1.2 * Rnd()
                 
        Next site
    End If

        
        ' ______ Null out common mode error ______
        Set CMError = measHigh.Math.Add(measLow).Divide(2)
        Set MeasureVolt = measNorm.Math.Subtract(CMError)
        Set gCMError = CMError
    
End Function


Public Function UVI80_DCDIFFMETER(H_Pin As String, ByRef MeasureVolt As PinListData) As Long
    TheHdw.Wait 0.0005
    MeasureVolt = TheHdw.DCDiffMeter.Pins(H_Pin).Read(tlStrobe, 5, 1000, tlDCDiffMeterReadingFormatAverage)
End Function

Public Function UVI80_DIFFMETER_RELEASE(H_Pin As String, L_Pin As String) As Long
    
    TheHdw.DCDiffMeter.Pins(H_Pin).Disconnect tlDCDiffMeterConnectDefault
    TheHdw.DCDiffMeter.Pins(L_Pin).Disconnect tlDCDiffMeterConnectDefault
    
End Function


Public Function ForceVal_Compare(InputType As Meas_Type, InputPin As String, MeasCase As String, Instr_Type As String, Optional MeasZ_idx As Integer = 0) ''Carter, 20190506
    Dim instr_pins() As String
    Dim num_pins As Long
    Dim ForceVal As Double
    Dim ForceHWVal As Double
    Dim i As Integer
    Call TheExec.DataManager.DecomposePinList(InputPin, instr_pins(), num_pins)
    
    For i = 0 To num_pins - 1
    
        If MeasCase = "I" Or MeasCase = "R" Or MeasCase = "Z" Then
            If Instr_Type = "PPMU" Then
                ForceHWVal = FormatNumber(TheHdw.PPMU.Pins(instr_pins(i)).Voltage.Value, 3)
            ElseIf Instr_Type = "DCVS" Then
                ForceHWVal = FormatNumber(TheHdw.DCVS.Pins(instr_pins(i)).Voltage.Value, 3)
            ElseIf Instr_Type = "DCVI" Then
                ForceHWVal = FormatNumber(TheHdw.DCVI.Pins(instr_pins(i)).Voltage, 3)
            End If
        ElseIf MeasCase = "V" Then
            If Instr_Type = "PPMU" Then
                ForceHWVal = FormatNumber(TheHdw.PPMU.Pins(instr_pins(i)).current.Value, 3)
            ElseIf Instr_Type = "DCVI" Then
                ForceHWVal = FormatNumber(TheHdw.DCVI.Pins(instr_pins(i)).current, 3)
            End If
        End If
        
        If MeasCase = "V" Or MeasCase = "I" Or MeasCase = "F" Or MeasCase = "R" Then
            ForceVal = FormatNumber(CDbl(InputType.ForceValueDic_HWCom(UCase(instr_pins(i)))), 3)
        Else
            Dim temp_ary() As String
            temp_ary = SplitInputCondition(InputType.ForceValueDic_HWCom(UCase(instr_pins(i))), "&")
            If MeasZ_idx > UBound(temp_ary) Then MeasZ_idx = temp_ary(0)
            ForceVal = FormatNumber(CDbl(temp_ary(MeasZ_idx)), 3)
        End If
                            
        If Not ForceVal = ForceHWVal Then
            If TheExec.TesterMode <> testModeOffline Then
                TheExec.Datalog.WriteComment "Error - PinName: " & instr_pins(i) & " - The Force Value setup is differ from HW by Pin setting, Please check"
            End If
        End If
            
    Next i
End Function
    

Public Function Opt_Input_Parsing(DSSC_Str As String, Special_Str As String, ByRef DSP_EYE_StartBit As DSPWave, ByRef DSP_EYE_BitLength As DSPWave) As Long
                                 
    Dim SplitByComma() As String
    Dim NumOfReg As Integer
    Dim i As Integer
    Dim TempStrArr() As String
    Dim array_count As Integer: array_count = 0
    Dim Temp_StartBit As Integer: Temp_StartBit = 0
    
    SplitByComma = Split(DSSC_Str, ",")
    NumOfReg = UBound(SplitByComma) - 1

    ReDim EYE_StartBit(NumOfReg) As Long
    ReDim EYE_BitLength(NumOfReg) As Long

    For i = 0 To NumOfReg - 1
        TempStrArr = Split(SplitByComma(i + 1), ":")
  
        'If LCase(TempStrArr(1) Like "*" & LCase(Special_Str) & "*") Then
        If InStr(LCase(TempStrArr(1)), LCase(Special_Str)) <> 0 Then
            EYE_StartBit(array_count) = Temp_StartBit
            EYE_BitLength(array_count) = TempStrArr(0)
            array_count = array_count + 1
        End If
         Temp_StartBit = Temp_StartBit + TempStrArr(0)
        
    Next i
    
    ReDim Preserve EYE_StartBit(array_count - 1)
    ReDim Preserve EYE_BitLength(array_count - 1)
    
    DSP_EYE_StartBit.Data = EYE_StartBit
    DSP_EYE_BitLength.Data = EYE_BitLength
    
End Function
Public Function Split_GrayDSP_2sComplementDSP_to_Dec(CUS_Str_MainProgram As String, DecomposeParseDigCapBit() As String, DecomposeTestName() As String, SourceBitStrmWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
            Dim i As Long
            Dim Index_SignedGray As Long
            Dim Index_UnSignedGray As Long
            Dim Index_2sComplement As Long
            Dim Index_SignedBin As Long
            Dim DSSC_SplitBySemiColon() As String: DSSC_SplitBySemiColon = Split(CUS_Str_MainProgram, ";")
            Dim DSSC_SignedGray() As String
            Dim DSSC_UnSignedGray() As String
            Dim DSSC_2sComplement() As String
            Dim DSSC_SignedBin() As String
            Dim DSPSignedGray_StartBit As New DSPWave
            Dim DSPUnSignedGray_StartBit As New DSPWave
            Dim DSP2sComplement_StartBit As New DSPWave
            Dim DSPSignedBin_StartBit As New DSPWave
            Dim DSPSignedGray_StartBit_Array() As Long
            Dim DSPUnSignedGray_StartBit_Array() As Long
            Dim DSP2sComplement_StartBit_Array() As Long
            Dim DSPSignedBin_StartBit_Array() As Long
            Dim AccumulateParseDigCapBit() As Long: ReDim AccumulateParseDigCapBit(UBound(DecomposeParseDigCapBit)) As Long
            
            For i = 0 To UBound(DSSC_SplitBySemiColon)
                If UCase(DSSC_SplitBySemiColon(i)) Like "SIGNEDGRAY*" Then
                    DSSC_SignedGray = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                ElseIf UCase(DSSC_SplitBySemiColon(i)) Like "UNSIGNEDGRAY*" Then
                    DSSC_UnSignedGray = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                ElseIf UCase(DSSC_SplitBySemiColon(i)) Like "2SCOMPLEMENT*" Then
                    DSSC_2sComplement = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                ElseIf UCase(DSSC_SplitBySemiColon(i)) Like "SIGNEDBIN*" Then
                    DSSC_SignedBin = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                End If
            Next i
            
            ReDim DSPSignedGray_StartBit_Array(UBound(DSSC_SignedGray)) As Long
            ReDim DSPUnSignedGray_StartBit_Array(UBound(DSSC_UnSignedGray)) As Long
            ReDim DSP2sComplement_StartBit_Array(UBound(DSSC_2sComplement)) As Long
            ReDim DSPSignedBin_StartBit_Array(UBound(DSSC_SignedBin)) As Long
            
            For i = 0 To UBound(DecomposeTestName)
                If i = 0 Then
                    AccumulateParseDigCapBit(i) = DecomposeParseDigCapBit(i)
                Else
                    AccumulateParseDigCapBit(i) = AccumulateParseDigCapBit(i - 1) + DecomposeParseDigCapBit(i)
                End If
                
                If DecomposeTestName(i) = DSSC_SignedGray(Index_SignedGray) Then
                    DSPSignedGray_StartBit_Array(Index_SignedGray) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
                    If Index_SignedGray <> UBound(DSSC_SignedGray) Then: Index_SignedGray = Index_SignedGray + 1
                ElseIf DecomposeTestName(i) = DSSC_UnSignedGray(Index_UnSignedGray) Then
                    DSPUnSignedGray_StartBit_Array(Index_UnSignedGray) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
                    If Index_UnSignedGray <> UBound(DSSC_UnSignedGray) Then: Index_UnSignedGray = Index_UnSignedGray + 1
                ElseIf DecomposeTestName(i) = DSSC_2sComplement(Index_2sComplement) Then
                    DSP2sComplement_StartBit_Array(Index_2sComplement) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
                    If Index_2sComplement <> UBound(DSSC_2sComplement) Then: Index_2sComplement = Index_2sComplement + 1
                ElseIf DecomposeTestName(i) = DSSC_SignedBin(Index_SignedBin) Then
                    DSPSignedBin_StartBit_Array(Index_SignedBin) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
                    If Index_SignedBin <> UBound(DSSC_SignedBin) Then: Index_SignedBin = Index_SignedBin + 1
                End If
            Next i
            If UBound(DSSC_SignedGray) = 0 And LCase(DSSC_SignedGray(0)) = "nouse" Then DSPSignedGray_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            If UBound(DSSC_UnSignedGray) = 0 And LCase(DSSC_UnSignedGray(0)) = "nouse" Then DSPUnSignedGray_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            If UBound(DSSC_2sComplement) = 0 And LCase(DSSC_2sComplement(0)) = "nouse" Then DSP2sComplement_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            If UBound(DSSC_SignedBin) = 0 And LCase(DSSC_SignedBin(0)) = "nouse" Then DSPSignedBin_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            DSPSignedGray_StartBit.Data = DSPSignedGray_StartBit_Array
            DSPUnSignedGray_StartBit.Data = DSPUnSignedGray_StartBit_Array
            DSP2sComplement_StartBit.Data = DSP2sComplement_StartBit_Array
            DSPSignedBin_StartBit.Data = DSPSignedBin_StartBit_Array
            rundsp.Split_Gray_2sComplementDSPWave_to_Dec DSPSignedGray_StartBit, DSPUnSignedGray_StartBit, DSP2sComplement_StartBit, DSPSignedBin_StartBit, SourceBitStrmWf, width_Wf, OutWf
End Function
Public Function Special_DigSrc_Collection() As Long
'**************************************************
'SeaHawk Edited by 20190606
'**************************************************
    
    Dim WorkBookName As Workbook
    Dim WorkSheetName As Worksheet
    
    Dim i, j, k As Integer
    Dim StringSplit() As String
    Dim PatNameStr As String
    Dim SweepStartPoint As Integer
    Dim DSPWave_Capture As New DSPWave
    Dim DSPWave_Calculate As New DSPWave
    Dim CUS_Str_DigCapSrcData As String
    
    Set WorkBookName = Application.ActiveWorkbook
    Set WorkSheetName = WorkBookName.Sheets("Special_DigCapSrcTable")
    
    For i = 1 To CLng(WorkSheetName.UsedRange.Rows.Count)                                     ' Search InstanceName for test item
        PatNameStr = CStr(WorkSheetName.Cells(i, 1))
        If PatNameStr Like "*Pat*" Then
            SweepStartPoint = i
            Debug.Print PatNameStr
            CUS_Str_DigCapSrcData = ""
            For j = SweepStartPoint + 1 To CLng(WorkSheetName.UsedRange.Rows.Count)           ' Addition two is means skip string InstanceNam
                If CStr(WorkSheetName.Cells(j, 1)) = "" Then
                    Exit For
                End If
                If CStr(WorkSheetName.Cells(i, 2)) Like "*DigSrc*" Then                       ' For CUS_DSSC_Source
                    CUS_Str_DigCapSrcData = CUS_Str_DigCapSrcData & CStr(WorkSheetName.Cells(j, 1))
                    CUS_Str_DigCapSrcData = CUS_Str_DigCapSrcData & "+"
                Else                                                                          ' For CUS_DSSC_Capture
                    CUS_Str_DigCapSrcData = CUS_Str_DigCapSrcData & CStr(WorkSheetName.Cells(j, 1))
                    CUS_Str_DigCapSrcData = CUS_Str_DigCapSrcData & ","
                End If
            Next j
            If CStr(WorkSheetName.Cells(i, 2)) Like "*DigSrc*" Then
                PatNameStr = Replace(PatNameStr, "Pat:", "") & "_SpecialDigSrc"
            Else
                PatNameStr = Replace(PatNameStr, "Pat:", "") & "_SpecialDigCap"
            End If
            Public_AddStoredString PatNameStr, CUS_Str_DigCapSrcData
        End If
    Next i
    
End Function




Public Function WaitTime_Check(TestCase() As String, SeqSplit_WaitTime() As String) As Variant
    
    Dim Seq_Index As Long
    Dim Sweep_index As Long
    Dim Max_WaitTime As String
    Dim Max_WaitTime_Temp() As String
    
    Dim Seq_WaitTime_VFIRZ As String
    Dim Seq_WaitTime_MeasI As String
    Dim Seq_WaitTime_MeasF As String
    Dim Seq_WaitTime_MeasV_UVI80 As String
    
    Dim Sweep_WaitTime_VFIRZ As String
    Dim Sweep_WaitTime_MeasI As String
    Dim Sweep_WaitTime_MeasF As String
    Dim Sweep_WaitTime_MeasV_UVI80 As String
    
    Dim SeqSplit_WaitTime_VFIRZ() As String
    Dim SeqSplit_WaitTime_MeasI() As String
    Dim SeqSplit_WaitTime_MeasF() As String
    Dim SeqSplit_WaitTime_MeasV_UVI80() As String
    
    Dim SweepSplit_WaitTime_VFIRZ() As String
    Dim SweepSplit_WaitTime_MeasI() As String
    Dim SweepSplit_WaitTime_MeasF() As String
    Dim SweepSplit_WaitTime_MeasV_UVI80() As String
     
    On Error GoTo err:
    
    SeqSplit_WaitTime_VFIRZ = SplitInputCondition(Instance_Data.WaitTime_VFIRZ, SplitSeq_WaitTime)
    SeqSplit_WaitTime_MeasI = SplitInputCondition(Instance_Data.MeasI_WaitTime, SplitSeq_WaitTime)
    SeqSplit_WaitTime_MeasF = SplitInputCondition(Instance_Data.MeasF_WaitTime, SplitSeq_WaitTime)
    SeqSplit_WaitTime_MeasV_UVI80 = SplitInputCondition(Instance_Data.MeasV_WaitTime_UVI80, SplitSeq_WaitTime)
    
    If UBound(SeqSplit_WaitTime) <> UBound(SeqSplit_WaitTime_VFIRZ) Then ReDim Preserve SeqSplit_WaitTime_VFIRZ(UBound(SeqSplit_WaitTime))
    If UBound(SeqSplit_WaitTime) <> UBound(SeqSplit_WaitTime_MeasI) Then ReDim Preserve SeqSplit_WaitTime_MeasI(UBound(SeqSplit_WaitTime))
    If UBound(SeqSplit_WaitTime) <> UBound(SeqSplit_WaitTime_MeasF) Then ReDim Preserve SeqSplit_WaitTime_MeasF(UBound(SeqSplit_WaitTime))
    If UBound(SeqSplit_WaitTime) <> UBound(SeqSplit_WaitTime_MeasV_UVI80) Then ReDim Preserve SeqSplit_WaitTime_MeasV_UVI80(UBound(SeqSplit_WaitTime))
    
    For Seq_Index = 0 To UBound(SeqSplit_WaitTime)
    
    ''''------Check Duplicate WaitTime------
        If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
            If SeqSplit_WaitTime_MeasI(Seq_Index) <> "" And SeqSplit_WaitTime_MeasF(Seq_Index) <> "" Then
                TheExec.Datalog.WriteComment (" =====> Error: Argument WaitTime is duplicate in MeasI_WaitTime and MeasF_WaitTime")
            ElseIf SeqSplit_WaitTime_MeasI(Seq_Index) <> "" And SeqSplit_WaitTime_MeasV_UVI80(Seq_Index) <> "" Then
                TheExec.Datalog.WriteComment (" =====> Error: Argument WaitTime is duplicate in MeasI_WaitTime and WaitTime_MeasV_UVI80")
            ElseIf SeqSplit_WaitTime_MeasF(Seq_Index) <> "" And SeqSplit_WaitTime_MeasV_UVI80(Seq_Index) <> "" Then
                TheExec.Datalog.WriteComment (" =====> Error: Argument WaitTime is duplicate in WaitTime_MeasF and WaitTime_MeasV_UVI80")
            End If
        End If
    ''''------Check Duplicate WaitTime------
        If Seq_Index = 0 Then
            SeqSplit_WaitTime_VFIRZ(Seq_Index) = IIf((SeqSplit_WaitTime_VFIRZ(Seq_Index) = ""), "0", SeqSplit_WaitTime_VFIRZ(0))
            SeqSplit_WaitTime_MeasI(Seq_Index) = IIf((SeqSplit_WaitTime_MeasI(Seq_Index) = ""), "0", SeqSplit_WaitTime_MeasI(0))
            SeqSplit_WaitTime_MeasF(Seq_Index) = IIf((SeqSplit_WaitTime_MeasF(Seq_Index) = ""), "0", SeqSplit_WaitTime_MeasF(0))
            SeqSplit_WaitTime_MeasV_UVI80(Seq_Index) = IIf((SeqSplit_WaitTime_MeasV_UVI80(Seq_Index) = ""), "0", SeqSplit_WaitTime_MeasV_UVI80(0))
            
        Else
            SeqSplit_WaitTime_VFIRZ(Seq_Index) = IIf((SeqSplit_WaitTime_VFIRZ(Seq_Index) = ""), SeqSplit_WaitTime_VFIRZ(Seq_Index - 1), SeqSplit_WaitTime_VFIRZ(Seq_Index))
            SeqSplit_WaitTime_MeasI(Seq_Index) = IIf((SeqSplit_WaitTime_MeasI(Seq_Index) = ""), SeqSplit_WaitTime_MeasI(Seq_Index - 1), SeqSplit_WaitTime_MeasI(Seq_Index))
            SeqSplit_WaitTime_MeasF(Seq_Index) = IIf((SeqSplit_WaitTime_MeasF(Seq_Index) = ""), SeqSplit_WaitTime_MeasF(Seq_Index - 1), SeqSplit_WaitTime_MeasF(Seq_Index))
            SeqSplit_WaitTime_MeasV_UVI80(Seq_Index) = IIf((SeqSplit_WaitTime_MeasV_UVI80(Seq_Index) = ""), SeqSplit_WaitTime_MeasV_UVI80(Seq_Index - 1), SeqSplit_WaitTime_MeasV_UVI80(Seq_Index))
        End If

        SweepSplit_WaitTime_VFIRZ = SplitInputCondition(SeqSplit_WaitTime_VFIRZ(Seq_Index), SplitPin)
        If UBound(SweepSplit_WaitTime_VFIRZ) <> 0 Then
            TheExec.Datalog.WriteComment (" =====> Error: Do not support the waittime setting for each pin!!!")
                            
        Else
            Max_WaitTime = SeqSplit_WaitTime_VFIRZ(Seq_Index)
            Select Case UCase(Mid(TestCase(Seq_Index), 1, 1))
            Case "I":
                ''------Check MeasI_WaitTime------
                If SeqSplit_WaitTime_MeasI(Seq_Index) <> "0" Then Max_WaitTime = SweepSplit_WaitTime(SeqSplit_WaitTime_MeasI(Seq_Index), Max_WaitTime)
            Case "F":
                ''------Check MeasF_WaitTime------
                If SeqSplit_WaitTime_MeasF(Seq_Index) <> "0" Then Max_WaitTime = SweepSplit_WaitTime(SeqSplit_WaitTime_MeasF(Seq_Index), Max_WaitTime)
            Case "V":
                ''------Check MeasV_WaitTime_UVI80------
                If SeqSplit_WaitTime_MeasV_UVI80(Seq_Index) <> "0" Then Max_WaitTime = SweepSplit_WaitTime(SeqSplit_WaitTime_MeasV_UVI80(Seq_Index), Max_WaitTime, "0.001")
            End Select
            SeqSplit_WaitTime(Seq_Index) = Max_WaitTime
        End If
        
    Next Seq_Index
    
        
'    For Seq_Index = 0 To UBound(SeqSplit_WaitTime)
'
'        Max_WaitTime = SeqSplit_WaitTime_VFIRZ(Seq_Index)
'        If SeqSplit_WaitTime_MeasI(Seq_Index) <> "0" Then
'            If CDbl(SeqSplit_WaitTime_MeasI(Seq_Index)) > CDbl(SeqSplit_WaitTime_VFIRZ(Seq_Index)) Then Max_WaitTime = SeqSplit_WaitTime_MeasI(Seq_Index)
'
'        ElseIf SeqSplit_WaitTime_MeasF(Seq_Index) <> "0" Then
'            If CDbl(SeqSplit_WaitTime_MeasF(Seq_Index)) > CDbl(SeqSplit_WaitTime_VFIRZ(Seq_Index)) Then Max_WaitTime = SeqSplit_WaitTime_MeasF(Seq_Index)
'
'        ElseIf SeqSplit_WaitTime_MeasV_UVI80(Seq_Index) <> "0" Then
'            If CDbl(SeqSplit_WaitTime_MeasV_UVI80(Seq_Index)) > CDbl(SeqSplit_WaitTime_VFIRZ(Seq_Index)) Then Max_WaitTime = SeqSplit_WaitTime_MeasV_UVI80(Seq_Index)
'
'        End If
'        SeqSplit_WaitTime(Seq_Index) = Max_WaitTime
'
'    Next Seq_Index
Exit Function
err:
    TheExec.Datalog.WriteComment "<Error> " + "WaitTime_Check" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SweepSplit_WaitTime(Seq_WaitTime As String, Max_WaitTime As String, Optional Default_WaitTime As String = "0") As Variant
    
    Dim Sweep_index As Long
    Dim Sweep_WaitTime() As String
    Dim Max_WaitTime_Temp() As String
    
    On Error GoTo err:
    
    Sweep_WaitTime = SplitInputCondition(Seq_WaitTime, SplitPin)
    
    If UBound(Sweep_WaitTime) <> 0 Then
        ReDim Max_WaitTime_Temp(UBound(Sweep_WaitTime))
        For Sweep_index = 0 To UBound(Sweep_WaitTime)
            Sweep_WaitTime(Sweep_index) = IIf((Sweep_WaitTime(Sweep_index) = ""), "0", Sweep_WaitTime(Sweep_index))
            If CDbl(Sweep_WaitTime(Sweep_index)) > CDbl(Max_WaitTime) Then
                Max_WaitTime_Temp(Sweep_index) = CStr(Evaluate(Sweep_WaitTime(Sweep_index) & "+" & Default_WaitTime))
            Else
                Max_WaitTime_Temp(Sweep_index) = CStr(Evaluate(Max_WaitTime & "+" & Default_WaitTime))
            End If
        Next Sweep_index
        SweepSplit_WaitTime = Join(Max_WaitTime_Temp, ",")
    Else
        If CDbl(Seq_WaitTime) > CDbl(Max_WaitTime) Then
            SweepSplit_WaitTime = CStr(Evaluate(Seq_WaitTime & "+" & Default_WaitTime))
        Else
            SweepSplit_WaitTime = CStr(Evaluate(Max_WaitTime & "+" & Default_WaitTime))
        End If
        
    End If
    Exit Function
err:
    TheExec.Datalog.WriteComment "<Error> " + "SweepSplit_WaitTime" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ParseData_Freq(ByRef Measure As MeasF_Type, MeasType As String, TestSeqSweepSplitSingPin As String, TestSeqSweepSplitDiffPin As String, TestSeqSweepSplitInterval As String, TestSeqSweepSplitWaitTime As String, TestSeqSweepIndex As Long) As MeasF_Type

    Dim i As Integer
    Dim PinIdx As Long
    Dim Maxubound As Double
    Dim TestSeqSweepLoopPinSplitPinAry() As String
    Dim TestSeqSweepLoopPinSplitSingPinAry() As String
    Dim TestSeqSweepLoopPinSplitDiffPinAry() As String
    Dim TestSeqSweepLoopPinSplitIntervalAry() As String
    Dim TestSeqSweepLoopPinSplitWaitTimeAry() As String
    Dim TestSeqSweepLoopPinSplitDiffFlagAry() As Boolean
    
    Dim TestSeqSweepLoopPinSplitPin As String
    Dim TestSeqSweepLoopPinSplitInterval As String
    Dim TestSeqSweepLoopPinSplitWaitTime As String
    Dim TestSeqSweepLoopPinSplitDiffFlag As Boolean

    On Error GoTo err:
    
    If Instance_Data.MeasF_ThresholdPercentage = 0 Then Instance_Data.MeasF_ThresholdPercentage = pc_Def_VFI_FreqThresholdPercentage
    
'    TestSeqSweepLoopPinSplitPinAry = SplitInputCondition(TestSeqSweepSplitPin, ",")

    If TestSeqSweepSplitSingPin <> "" And TestSeqSweepSplitDiffPin <> "" Then
        TestSeqSweepLoopPinSplitSingPinAry = SplitInputCondition(TestSeqSweepSplitSingPin, ",")
        TestSeqSweepLoopPinSplitDiffPinAry = SplitInputCondition(TestSeqSweepSplitDiffPin, ",")
        Maxubound = max(UBound(TestSeqSweepLoopPinSplitSingPinAry), UBound(TestSeqSweepLoopPinSplitDiffPinAry))
        ReDim TestSeqSweepLoopPinSplitPinAry(Maxubound)
        ReDim TestSeqSweepLoopPinSplitDiffFlagAry(Maxubound)
        ReDim Preserve TestSeqSweepLoopPinSplitSingPinAry(Maxubound)
        ReDim Preserve TestSeqSweepLoopPinSplitDiffPinAry(Maxubound)
        For PinIdx = 0 To Maxubound
            If TestSeqSweepLoopPinSplitSingPinAry(PinIdx) <> "" And TestSeqSweepLoopPinSplitDiffPinAry(PinIdx) <> "" Then
                TestSeqSweepLoopPinSplitPinAry(PinIdx) = TestSeqSweepLoopPinSplitSingPinAry(PinIdx)
                TestSeqSweepLoopPinSplitDiffFlagAry(PinIdx) = False
                TheExec.Datalog.WriteComment ("Warning: Same sequence include both Single and differeitnal pin, use single pin only")
                
            ElseIf TestSeqSweepLoopPinSplitSingPinAry(PinIdx) <> "" And TestSeqSweepLoopPinSplitDiffPinAry(PinIdx) = "" Then
                TestSeqSweepLoopPinSplitPinAry(PinIdx) = TestSeqSweepLoopPinSplitSingPinAry(PinIdx)
                TestSeqSweepLoopPinSplitDiffFlagAry(PinIdx) = False
                
            ElseIf TestSeqSweepLoopPinSplitSingPinAry(PinIdx) = "" And TestSeqSweepLoopPinSplitDiffPinAry(PinIdx) <> "" Then
                TestSeqSweepLoopPinSplitPinAry(PinIdx) = TestSeqSweepLoopPinSplitDiffPinAry(PinIdx)
                TestSeqSweepLoopPinSplitDiffFlagAry(PinIdx) = True
                
            End If
        Next PinIdx
        
    ElseIf TestSeqSweepSplitSingPin <> "" And TestSeqSweepSplitDiffPin = "" Then
        TestSeqSweepLoopPinSplitPinAry = SplitInputCondition(TestSeqSweepSplitSingPin, ",")
        ReDim TestSeqSweepLoopPinSplitDiffFlagAry(UBound(TestSeqSweepLoopPinSplitPinAry))
        
    ElseIf TestSeqSweepSplitSingPin = "" And TestSeqSweepSplitDiffPin <> "" Then
        TestSeqSweepLoopPinSplitPinAry = SplitInputCondition(TestSeqSweepSplitDiffPin, ",")
        ReDim TestSeqSweepLoopPinSplitDiffFlagAry(UBound(TestSeqSweepLoopPinSplitPinAry))
        For i = 0 To UBound(TestSeqSweepLoopPinSplitPinAry)
            TestSeqSweepLoopPinSplitDiffFlagAry(i) = True
        Next i
        
    Else
        TheExec.Datalog.WriteComment (" =====> Error, No Single and Diff Pin when MeasF")
        Exit Function
    End If
        
    Measure.MeasureThreshold_Flag = Instance_Data.MeasF_Flag_MeasureThreshold
    Measure.ThresholdPercentage = Instance_Data.MeasF_ThresholdPercentage
    Measure.WalkingStrobe_Flag = Instance_Data.MeasF_WalkingStrobe_Flag
    Measure.EventSource = Instance_Data.MeasF_EventSource
    Measure.EnableVtMode_Flag = Instance_Data.MeasF_EnableVtMode_Flag
    
    TestSeqSweepLoopPinSplitIntervalAry = SplitInputCondition(TestSeqSweepSplitInterval, ",")
    TestSeqSweepLoopPinSplitWaitTimeAry = SplitInputCondition(TestSeqSweepSplitWaitTime, ",")
    
    For PinIdx = 0 To UBound(TestSeqSweepLoopPinSplitPinAry)
        
        TestSeqSweepLoopPinSplitPin = SplitInputCondition(CheckAndReturnArrayData(TestSeqSweepLoopPinSplitPinAry, PinIdx), ":", TestSeqSweepIndex)
        TestSeqSweepLoopPinSplitInterval = SplitInputCondition(CheckAndReturnArrayData(TestSeqSweepLoopPinSplitIntervalAry, PinIdx), ":", TestSeqSweepIndex)
        TestSeqSweepLoopPinSplitWaitTime = SplitInputCondition(CheckAndReturnArrayData(TestSeqSweepLoopPinSplitWaitTimeAry, PinIdx), ":", TestSeqSweepIndex)
        
        If TestSeqSweepLoopPinSplitInterval <> "" Then
            Measure.Interval = CDbl(TestSeqSweepLoopPinSplitInterval)
        Else
            Measure.Interval = CDbl(pc_Def_VFI_FreqInterval)
        End If
        
        Measure.Pins = Measure.Pins & "," & TestSeqSweepLoopPinSplitPin
        Measure.Differential_Flag = TestSeqSweepLoopPinSplitDiffFlagAry(PinIdx)
        Measure.WaitTime = TestSeqSweepLoopPinSplitWaitTime
        
    Next PinIdx
    Measure.Pins = Replace(Measure.Pins, ",", "", 1, 1)
    
    Exit Function
err:
    TheExec.Datalog.WriteComment "<Error> " + "ParseData_Freq" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ParseData_InterPose(Interpose_Ary() As String, TestSeq_Ary() As String) As Variant
    
    Dim MaxLen As Integer
    Dim AllTesSeq As Integer
    Dim TestSeqNum As Integer
    Dim TestOptLen As Integer
    Dim TestSeqSweepNum As Integer
    Dim TestSeqLoopPreMeas As String
    Dim TestSeqSweepLoopPreMeas_Temp() As String
    Dim TestSeqSweepLoopPreMeasArray() As String
    
    On Error GoTo err:
    AllTesSeq = UBound(TestSeq_Ary)
    '''-------Start - Interpose_premeas assigns value by sequence - Carter, 20190616-------
    '''-----------> testsequence     : V,VV
    '''-----------> interpose_premeas:  |VDD_A:V:0.01`VDD_A:V:0.02;VDD_B:V:0.03
    If UBound(Interpose_Ary) <> AllTesSeq Then
        ReDim Preserve Interpose_Ary(AllTesSeq)
        If AllTesSeq <> 0 Then
            For TestSeqNum = 1 To AllTesSeq
                If Interpose_Ary(TestSeqNum) = "" Then Interpose_Ary(TestSeqNum) = Interpose_Ary(TestSeqNum - 1)  '''Use the previous setting for the current sequence
            Next TestSeqNum
        End If
    End If
    '''-------End - Interpose_premeas assigns value by sequence - Carter, 20190616-------
    For TestSeqNum = 0 To AllTesSeq
        TestOptLen = Len(TestSeq_Ary(TestSeqNum))
        MaxLen = max(CDbl(MaxLen), CDbl(TestOptLen))
    Next TestSeqNum
    
    ReDim TestSeqSweepLoopPreMeasArray(TestSeqNum - 1, MaxLen - 1)
    
    For TestSeqNum = 0 To AllTesSeq
        TestSeqLoopPreMeas = Interpose_Ary(TestSeqNum)
        '''-------Start - Add per sweep feature for interpose_premeas - Carter, 20190614-------
        '''-------example -> testsequence     : V,VVV,V
        '''-------example -> interpose_premeas: VDD_A:V:0.01,VDD_B:V:0.02|VDD_A:V:0.03,VDD_B:V:0.04
        TestSeqSweepLoopPreMeas_Temp = SplitInputCondition(TestSeqLoopPreMeas, "`")
        If MaxLen <> UBound(TestSeqSweepLoopPreMeas_Temp) + 1 Then ReDim Preserve TestSeqSweepLoopPreMeas_Temp(MaxLen - 1)
        
        For TestSeqSweepNum = 0 To MaxLen - 1
            If TestSeqSweepNum = 0 Then
                TestSeqSweepLoopPreMeas_Temp(0) = IIf((TestSeqSweepLoopPreMeas_Temp(0) = ""), "", TestSeqSweepLoopPreMeas_Temp(0))
            Else
                TestSeqSweepLoopPreMeas_Temp(TestSeqSweepNum) = IIf((TestSeqSweepLoopPreMeas_Temp(TestSeqSweepNum) = ""), TestSeqSweepLoopPreMeas_Temp(TestSeqSweepNum - 1), TestSeqSweepLoopPreMeas_Temp(TestSeqSweepNum))
            End If
            TestSeqSweepLoopPreMeasArray(TestSeqNum, TestSeqSweepNum) = TestSeqSweepLoopPreMeas_Temp(TestSeqSweepNum)
        
        Next TestSeqSweepNum
        
    Next TestSeqNum
    
    ParseData_InterPose = TestSeqSweepLoopPreMeasArray
    
    Exit Function
err:
    TheExec.Datalog.WriteComment "<Error> " + "ParseData_InterPose" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function StoredRegAssign()
Dim i As Long
RegDict.RemoveAll
For i = 0 To UBound(RegAssignInfo.ByTest)
    If RegDict.Exists(RegAssignInfo.ByTest(i).testName & "_" & "ModeA") Or RegDict.Exists(RegAssignInfo.ByTest(i).testName & "_" & "ModeB") Then
    TheExec.ErrorLogMessage "Duplicate Assign on" & RegAssignInfo.ByTest(i).testName
    Else
    RegDict.Add RegAssignInfo.ByTest(i).testName & "_" & "ModeA", RegAssignInfo.ByTest(i).RtnByModeA
    RegDict.Add RegAssignInfo.ByTest(i).testName & "_" & "ModeB", RegAssignInfo.ByTest(i).RtnByModeB
    End If
Next i

End Function
Public Function Split_GrayDSP_to_Dec(CUS_Str_MainProgram As String, DecomposeParseDigCapBit() As String, DecomposeTestName() As String, SourceBitStrmWf As DSPWave, width_Wf As DSPWave, OutWf As DSPWave) As Long
            Dim i As Long
            Dim Index_SignedGray As Long
            Dim Index_UnSignedGray As Long
            Dim Index_2sComplement As Long
            Dim DSSC_SplitBySemiColon() As String: DSSC_SplitBySemiColon = Split(CUS_Str_MainProgram, ";")
            Dim DSSC_SignedGray() As String
            Dim DSSC_UnSignedGray() As String
            Dim DSSC_2sComplement() As String
            Dim DSPSignedGray_StartBit As New DSPWave
            Dim DSPUnSignedGray_StartBit As New DSPWave
            Dim DSP2sComplement_StartBit As New DSPWave
            Dim DSPSignedGray_StartBit_Array() As Long
            Dim DSPUnSignedGray_StartBit_Array() As Long
            Dim DSP2sComplement_StartBit_Array() As Long
            Dim AccumulateParseDigCapBit() As Long: ReDim AccumulateParseDigCapBit(UBound(DecomposeParseDigCapBit)) As Long
            
            For i = 0 To UBound(DSSC_SplitBySemiColon)
                If UCase(DSSC_SplitBySemiColon(i)) Like "SIGNEDGRAY*" Then
                    DSSC_SignedGray = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                ElseIf UCase(DSSC_SplitBySemiColon(i)) Like "UNSIGNEDGRAY*" Then
                    DSSC_UnSignedGray = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                ElseIf UCase(DSSC_SplitBySemiColon(i)) Like "2SCOMPLEMENT*" Then
                    'DSSC_2sComplement = Split(Split(DSSC_SplitBySemiColon(i), ":")(1), ",")
                End If
            Next i
            
            ReDim DSPSignedGray_StartBit_Array(UBound(DSSC_SignedGray)) As Long
            ReDim DSPUnSignedGray_StartBit_Array(UBound(DSSC_UnSignedGray)) As Long
            'ReDim DSP2sComplement_StartBit_Array(UBound(DSSC_2sComplement)) As Long
            
            For i = 0 To UBound(DecomposeTestName)
                If i = 0 Then
                    AccumulateParseDigCapBit(i) = DecomposeParseDigCapBit(i)
                Else
                    AccumulateParseDigCapBit(i) = AccumulateParseDigCapBit(i - 1) + DecomposeParseDigCapBit(i)
                End If
                
                If DecomposeTestName(i) = DSSC_SignedGray(Index_SignedGray) Then
                    DSPSignedGray_StartBit_Array(Index_SignedGray) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
                    If Index_SignedGray <> UBound(DSSC_SignedGray) Then: Index_SignedGray = Index_SignedGray + 1
                ElseIf DecomposeTestName(i) = DSSC_UnSignedGray(Index_UnSignedGray) Then
                    DSPUnSignedGray_StartBit_Array(Index_UnSignedGray) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
                    If Index_UnSignedGray <> UBound(DSSC_UnSignedGray) Then: Index_UnSignedGray = Index_UnSignedGray + 1
'                ElseIf DecomposeTestName(i) = DSSC_2sComplement(Index_2sComplement) Then
'                    'DSP2sComplement_StartBit_Array(Index_2sComplement) = AccumulateParseDigCapBit(i) - DecomposeParseDigCapBit(i)
'                    'If Index_2sComplement <> UBound(DSSC_2sComplement) Then: Index_2sComplement = Index_2sComplement + 1
                End If
            Next i
            If UBound(DSSC_SignedGray) = 0 And LCase(DSSC_SignedGray(0)) = "nouse" Then DSPSignedGray_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            If UBound(DSSC_UnSignedGray) = 0 And LCase(DSSC_UnSignedGray(0)) = "nouse" Then DSPUnSignedGray_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            'If UBound(DSSC_2sComplement) = 0 And LCase(DSSC_2sComplement(0)) = "nouse" Then DSP2sComplement_StartBit_Array(0) = Instance_Data.DigCap_Sample_Size + 1
            
            DSPSignedGray_StartBit.Data = DSPSignedGray_StartBit_Array
            DSPUnSignedGray_StartBit.Data = DSPUnSignedGray_StartBit_Array
            'DSP2sComplement_StartBit.Data = DSP2sComplement_StartBit_Array
            rundsp.Split_Gray_to_Dec DSPSignedGray_StartBit, DSPUnSignedGray_StartBit, SourceBitStrmWf, width_Wf, OutWf
            'rundsp.Split_Gray_2sComplementDSPWave_to_Dec DSPSignedGray_StartBit, DSPUnSignedGray_StartBit, DSP2sComplement_StartBit, SourceBitStrmWf, width_Wf, OutWf
End Function

Public Function LDO_Measurement_Process(Pat As String, srcPin As PinList, code() As SiteLong, ByRef Res() As SiteDouble, TrimCodeSize As Long, NumberOfMeasV As Integer, ByRef Rtn_MeasVolt() As PinListData, DigSrc_Sample_Size As Long, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName() As String, MeasV_WaitTime As String)
    Dim srcWave() As New DSPWave: ReDim srcWave(UBound(code))
    Dim site As Variant
    Dim InDSPwave As New DSPWave
    Dim i As Long, j As Long
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array() As Long: ReDim srcwave_array(TrimCodeSize - 1)
    
    ByPassTestLimit = True
    glb_Disable_CurrRangeSetting_Print = True

    For i = 0 To UBound(code)
    srcWave(i).CreateConstant 0, TrimCodeSize, DspLong
        For Each site In TheExec.sites
            For j = 0 To TrimCodeSize - 1
                If j = 0 Then
                    srcwave_array(j) = code(i) And 1
                Else
                    srcwave_array(j) = (code(i) And (2 ^ j)) \ (2 ^ j)
                End If
            Next j
        srcWave(i).Data = srcwave_array
        Next site
        Call AddStoredCaptureData(TrimStoreName(i), srcWave(i))
    Next i
    Call GeneralDigSrcSetting(Pat, srcPin, DigSrc_Sample_Size, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, "", "", InDSPwave, "")
    
    TheHdw.Patterns(Pat).start
    
    For i = 0 To NumberOfMeasV - 1
        TheHdw.Digital.Patgen.FlagWait cpuA, 0
        Rtn_MeasVolt(i) = HardIP_MeasureVolt
        Call DebugPrintFunc_PPMU("")
        For Each site In TheExec.sites
            Res(i) = Rtn_MeasVolt(i).Pins(0).Value
            For j = 0 To Rtn_MeasVolt(i).Pins.Count - 1
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(i) & ", Pin : " & Rtn_MeasVolt(i).Pins(j) & ", Voltage = " & Rtn_MeasVolt(i).Pins(j).Value
            Next j
        Next site
        TheHdw.Digital.Patgen.Continue 0, cpuA
    Next i
    TheHdw.Digital.Patgen.HaltWait
    ByPassTestLimit = False
    glb_Disable_CurrRangeSetting_Print = False

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in LDO_Measurement_Process"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Sub SWEEP_VT(SweepVtStr As String, Interpose_PrePat As String)
    Dim SplitByColon() As String
    Dim SourceIndexStr As String, SourceIndex As Long
    Dim StartVal As Double, StepVal As Double, FinalVal As Double
    Dim ReplaceStr() As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    Dim p As Integer
    Dim TempStr() As String



    If SweepVtStr <> "" Then
        SplitByColon = Split(SweepVtStr, ":")
        SourceIndexStr = SplitByColon(0)
        SourceIndex = TheExec.Flow.var(SourceIndexStr).Value
        StartVal = SplitByColon(1)
        StepVal = SplitByColon(2)
        FinalVal = StartVal + SourceIndex * StepVal
        
        If Abs(FinalVal) <= 0.00000001 Then
            
            FinalVal = 0
        End If
            If InStr(UCase(Interpose_PrePat), ":VT:") <> 0 Then

                ''''' Purpose to only update VT value and keep the other interpose setting the same
                ReplaceStr = Split(Interpose_PrePat, "VT")
                If InStr(ReplaceStr(1), ";") Then
                    TempStr = Split(ReplaceStr(1), ";")

                    For p = 0 To UBound(TempStr)
                        If p = 0 Then
                            Interpose_PrePat = ReplaceStr(0) & "VT:" & CStr(FinalVal)
                        Else
                            Interpose_PrePat = Interpose_PrePat & ";" & TempStr(p)
                        End If
                    Next p
                Else
                    Interpose_PrePat = ReplaceStr(0) & "VT:" & CStr(FinalVal)
                End If

            End If
        End If
    gl_Sweep_vt = "sweepVT_" & Replace(Replace(FinalVal, "-", "m"), ".", "p") & "V"
End Sub


Public Sub SWEEP_V(SweepVtStr As String, Interpose_PrePat As String)
    '/*** ADDED by Kaino on 2019/06/27***/
    Dim SplitByColon() As String
    Dim SourceIndexStr As String, SourceIndex As Long
    Dim StartVal As Double, StepVal As Double, FinalVal As Double
    Dim ReplaceStr() As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    Dim p As Integer
    Dim TempStr() As String
    
    
    Dim tmp1() As String
    Dim tmp2() As String
    Dim i As Integer
    Dim j As Integer
    Dim ForceStr As String
    
    
    'V_Sweep:ANI_VREF:V:_VDDIO_NAND_VAR_H*0.5-0.02:_VDDIO_NAND_VAR_H*0.5+0.02:0.005
    
    
   
    
    
    tmp1() = Split(SweepVtStr, ";")
    
    For i = 0 To UBound(tmp1)
        If InStr(UCase(tmp1(i)), UCase("V_Sweep")) > 0 Then
        
            tmp2 = Split(tmp1(i), ":")
            SourceIndexStr = tmp2(0)
            
            SourceIndex = TheExec.Flow.var(SourceIndexStr).Value
            
            Call HIP_Evaluate_ForceVal_New(tmp2(3))
            Call HIP_Evaluate_ForceVal_New(tmp2(4))
            Call HIP_Evaluate_ForceVal_New(tmp2(5))
            
            
            FinalVal = CDbl(tmp2(3)) + CDbl(tmp2(5)) * SourceIndex
  
        
            ForceStr = tmp2(1) & ":V:" & CStr(FinalVal)
            
            TheExec.Datalog.WriteComment " ***pin:" & tmp2(1) & " sweep V from " & tmp2(3) & " to " & tmp2(4) & " step " & tmp2(5) & " current " & CStr(FinalVal)
            
            
        
            Interpose_PrePat = Interpose_PrePat & ";" & ForceStr
            
        End If
    
    
    Next i
    
    sweep_power_val_per_loop_count = CStr(FinalVal) '20190814
    gl_Sweep_Name = tmp2(1)
    gl_Sweep_vt = "sweepVT_" & Replace(Replace(FinalVal, "-", "m"), ".", "p") & "V"
End Sub


Public Function HardIP_SetupAndMeasureVolt_PPMU_BySerial(ByRef MeasureVolt As PinListData) As Long
    
    Dim MeasV As Meas_Type
    MeasV = TestConditionSeqData(Instance_Data.TestSeqNum).MeasV(Instance_Data.TestSeqSweepNum)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim PinArray() As String
    Dim PinArrayCnt As Long
    Dim TempCnt_Pin As Long
    Dim PinGroupCnt As Long
    Dim TempCnt_PinGroup As Long
    Dim TempCnt As Long
    Dim TempPinGroupString As String
    Dim TempMeasureVolt As New PinListData
    
    Dim PinNumPerMeas As Integer
    PinNumPerMeas = 1
    
    'If Instance_Data.InstSpecialSetting = PPMU_SerialMeasurement Then
    
        TheHdw.Digital.Pins(MeasV.Pins.PPMU).Disconnect
        
        Call TheExec.DataManager.DecomposePinList(MeasV.Pins.PPMU, PinArray(), PinArrayCnt)
        PinGroupCnt = PinArrayCnt \ PinNumPerMeas
        
        For TempCnt_PinGroup = 0 To PinGroupCnt
        
            TempPinGroupString = ""
            For TempCnt_Pin = 0 To PinNumPerMeas - 1
                TempCnt = TempCnt_PinGroup * PinNumPerMeas + TempCnt_Pin
                If TempCnt >= PinArrayCnt Then
                    Exit For
                End If
                'TheExec.Datalog.WriteComment ("TempCnt_PinGroup = " & TempCnt_PinGroup & ", TempCnt_Pin = " & TempCnt_Pin & ", TempCnt = " & TempCnt)
                With TheHdw.PPMU.Pins(PinArray(TempCnt))
                    .Gate = tlOff
                    .ForceI CDbl(MeasV.ForceValueDic_HWCom(PinArray(TempCnt))), CDbl(MeasV.ForceValueDic_HWCom(PinArray(TempCnt)))
                    .Connect
                    .Gate = tlOn
                End With
                If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas Force Current value, " & PinArray(TempCnt) & " =" & MeasV.ForceValueDic_HWCom(PinArray(TempCnt)))
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Volt_meas PPMU WaitTime, " & PinArray(TempCnt) & " =" & MeasV.WaitTime.PPMU) ''''
                End If
                TempPinGroupString = TempPinGroupString & "," & PinArray(TempCnt)
            Next TempCnt_Pin
            
            If TempPinGroupString <> "" Then
                TempPinGroupString = Right(TempPinGroupString, Len(TempPinGroupString) - 1)
                TheHdw.Wait CDbl(MeasV.WaitTime.PPMU)
                TempMeasureVolt = TheHdw.PPMU.Pins(TempPinGroupString).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            End If
            
            For TempCnt_Pin = 0 To PinNumPerMeas - 1
                TempCnt = TempCnt_PinGroup * PinNumPerMeas + TempCnt_Pin
                If TempCnt >= PinArrayCnt Then
                    Exit For
                End If
                MeasureVolt.AddPin (PinArray(TempCnt))
                MeasureVolt.Pins(PinArray(TempCnt)) = TempMeasureVolt.Pins(PinArray(TempCnt))
                With TheHdw.PPMU.Pins(PinArray(TempCnt))
                    .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
                    .Disconnect
                    .Gate = tlOff
                End With
            Next TempCnt_Pin
        
        Next TempCnt_PinGroup
        
        TheHdw.Wait (0.1 * ms)
        TheHdw.Digital.Pins(MeasV.Pins.PPMU).Connect
    
    'End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Pin As Variant
    Dim GetRakVal_PinList As New PinListData
    Dim DiffVolt_Pinlist As New PinListData
    If Instance_Data.RAK_Flag = R_TraceOnly Then
        For Each Pin In MeasureVolt.Pins
            GetRakVal_PinList.AddPin (CStr(Pin))
            DiffVolt_Pinlist.AddPin (CStr(Pin))
            GetRakVal_PinList.Pins(Pin) = CurrentJob_Card_RAK.Pins(Pin)
            If MeasV.Setup_ByTypeByPin_Flag = False Then
                DiffVolt_Pinlist.Pins(Pin) = GetRakVal_PinList.Pins(Pin).Multiply(CDbl(MeasV.Setup_ByType.PPMU.ForceValue1))
            Else
                DiffVolt_Pinlist.Pins(Pin) = GetRakVal_PinList.Pins(Pin).Multiply(CDbl(MeasV.ForceValueDic_HWCom(Pin)))
            End If
            
            If gl_Disable_HIP_debug_log = False Then
                For Each site In TheExec.sites.Active
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Voltage = " & MeasureVolt.Pins(Pin).Value(site) & ", RAK val = " & GetRakVal_PinList.Pins(Pin).Value(site)
                Next site
            End If
            
        Next Pin
        MeasureVolt = MeasureVolt.Math.Subtract(DiffVolt_Pinlist)
        
    ElseIf Instance_Data.RAK_Flag = R_PathWithContact Then
        For Each Pin In MeasureVolt.Pins
            GetRakVal_PinList.AddPin (CStr(Pin))
            DiffVolt_Pinlist.AddPin (CStr(Pin))
            GetRakVal_PinList.Pins(Pin) = R_Path_PLD.Pins(Pin)
            If MeasV.Setup_ByTypeByPin_Flag = False Then
                DiffVolt_Pinlist.Pins(Pin) = GetRakVal_PinList.Pins(Pin).Multiply(CDbl(MeasV.Setup_ByType.PPMU.ForceValue1))
            Else
                DiffVolt_Pinlist.Pins(Pin) = GetRakVal_PinList.Pins(Pin).Multiply(CDbl(MeasV.ForceValueDic_HWCom(Pin)))
            End If
            
            If gl_Disable_HIP_debug_log = False Then
                For Each site In TheExec.sites.Active
                     TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Voltage = " & MeasureVolt.Pins(Pin).Value(site) & ", R_Path val = " & GetRakVal_PinList.Pins(Pin).Value(site)
                Next site
            End If
    
        Next Pin
        MeasureVolt = MeasureVolt.Math.Subtract(DiffVolt_Pinlist)
        
    ElseIf InStr(UCase(Instance_Data.CUS_Str_MainProgram), UCase("RREF_RAK_CALC")) <> 0 Then
        Call CUS_RREF_Rak_Calc(MeasureVolt)
    End If

End Function



Public Function HardIP_SetupAndMeasureCurrent_PPMU_BySerial(SampleSize As Long, ByRef measureCurrent As PinListData)
    
    Dim MeasI As Meas_Type
    
    Dim Pins() As String
    Dim Pin_Cnt As Long
    Dim var As Variant
    
    MeasI = TestConditionSeqData(Instance_Data.TestSeqNum).MeasI(Instance_Data.TestSeqSweepNum)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim PinArray() As String
    Dim PinArrayCnt As Long
    Dim TempCnt_Pin As Long
    Dim PinGroupCnt As Long
    Dim TempCnt_PinGroup As Long
    Dim TempCnt As Long
    Dim TempPinGroupString As String
    Dim TempMeasureCurrent As New PinListData
    
    Dim PinNumPerMeas As Integer
    PinNumPerMeas = 1
    
    'If Instance_Data.InstSpecialSetting = PPMU_SerialMeasurement Then
    
        TheHdw.Digital.Pins(MeasI.Pins.PPMU).Disconnect
        
        Call TheExec.DataManager.DecomposePinList(MeasI.Pins.PPMU, PinArray(), PinArrayCnt)
        PinGroupCnt = PinArrayCnt \ PinNumPerMeas
        
        For TempCnt_PinGroup = 0 To PinGroupCnt
        
            TempPinGroupString = ""
            For TempCnt_Pin = 0 To PinNumPerMeas - 1
                TempCnt = TempCnt_PinGroup * PinNumPerMeas + TempCnt_Pin
                If TempCnt >= PinArrayCnt Then
                    Exit For
                End If
                'TheExec.Datalog.WriteComment ("TempCnt_PinGroup = " & TempCnt_PinGroup & ", TempCnt_Pin = " & TempCnt_Pin & ", TempCnt = " & TempCnt)
                With TheHdw.PPMU.Pins(PinArray(TempCnt))
                    .Gate = tlOff
                    .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                    .ForceV CDbl(MeasI.ForceValueDic_HWCom(PinArray(TempCnt))), CDbl(MeasI.MeasCurRangeDic(PinArray(TempCnt)))
                    .Connect
                    .Gate = tlOn
                End With
                If glb_Disable_CurrRangeSetting_Print = False And gl_Disable_HIP_debug_log = False Then
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & PinArray(TempCnt) & " =" & TheHdw.PPMU.Pins(PinArray(TempCnt)).MeasureCurrentRange)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & PinArray(TempCnt) & " =" & MeasI.MeasCurRangeDic(PinArray(TempCnt)))
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I WaitTime, " & PinArray(TempCnt) & " =" & MeasI.WaitTime.PPMU)
                    TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Force Volt value, " & PinArray(TempCnt) & " =" & MeasI.ForceValueDic_HWCom(PinArray(TempCnt)))
                End If
                TempPinGroupString = TempPinGroupString & "," & PinArray(TempCnt)
            Next TempCnt_Pin
            
            If TempPinGroupString <> "" Then
                TempPinGroupString = Right(TempPinGroupString, Len(TempPinGroupString) - 1)
                TheHdw.Wait CDbl(MeasI.WaitTime.PPMU)
                DebugPrintFunc_PPMU CStr(MeasI.Pins.PPMU)
                TempMeasureCurrent = TheHdw.PPMU.Pins(TempPinGroupString).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            End If
            
            For TempCnt_Pin = 0 To PinNumPerMeas - 1
                TempCnt = TempCnt_PinGroup * PinNumPerMeas + TempCnt_Pin
                If TempCnt >= PinArrayCnt Then
                    Exit For
                End If
                measureCurrent.AddPin (PinArray(TempCnt))
                measureCurrent.Pins(PinArray(TempCnt)) = TempMeasureCurrent.Pins(PinArray(TempCnt))
                With TheHdw.PPMU.Pins(PinArray(TempCnt))
                    .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range
                    TheHdw.Wait (0.3 * ms) '20191002 CT add to solve MeasI clamp for GPIO DS tests
                    .Disconnect
                    .Gate = tlOff
                End With
            Next TempCnt_Pin
        
        Next TempCnt_PinGroup
        
        TheHdw.Wait (0.1 * ms)
        TheHdw.Digital.Pins(MeasI.Pins.PPMU).Connect
    
    'End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    Dim i As Long
    Dim PastVal As Double
    Dim b_ForceDiffVolt As Boolean
    
    b_ForceDiffVolt = False
    If MeasI.Setup_ByTypeByPin_Flag = True Then
        For i = 0 To UBound(MeasI.Setup_ByTypeByPin.PPMU)
            If i <> 0 Then
                If MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1 <> PastVal Then
                    b_ForceDiffVolt = True
                    Exit For
                End If
            End If
            PastVal = MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1
        Next i
    End If
    
''''20190527  CT add for GPIO ForceV then MeasI(de-embedded Trace effect)
If Not (UCase(TheExec.DataManager.instanceName) Like "*FAILSAFE*") Then
    If UCase(TheExec.DataManager.instanceName) Like "*GPIO*" Then
    
        Dim Pin As Variant
        Dim GetRakVal_PinList As New PinListData
        Dim DiffVolt_Pinlist As New PinListData
        Dim Vdiff As Double
        If Instance_Data.RAK_Flag = R_TraceOnly Then
        
            If MeasI.Setup_ByTypeByPin_Flag = False Then
                If CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) > 0.8 Then
                    'If LCase(Pin) Like "*1p2*" Then
                    '    Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) ''vddio12_grp!!!
                    'Else
                        Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) ''vddio18_grp!!!
                    'End If
                Else
                    Vdiff = CDbl(MeasI.Setup_ByType.PPMU.ForceValue1)
                End If
                
                For Each Pin In measureCurrent.Pins
                
                    If gl_Disable_HIP_debug_log = False Then
                        For Each site In TheExec.sites.Active
                             TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Current = " & measureCurrent.Pins(Pin).Value(site) & ", RAK val = " & CurrentJob_Card_RAK.Pins(Pin).Value(site)
                        Next site
                    End If
                
                    measureCurrent.Pins(Pin) = measureCurrent.Pins(Pin).Multiply(CurrentJob_Card_RAK.Pins(Pin)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(Pin))
                                        
                Next Pin
            Else
                If b_ForceDiffVolt = False Then
                    If CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1) > 0.8 Then
                        'If LCase(Pin) Like "*1p2*" Then
                        '    Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio12_grp!!!
                        'Else
                            Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1) ''vddio18_grp!!!
                        'End If
                    Else
                        Vdiff = CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1)
                    End If
                    
                    For Each Pin In measureCurrent.Pins
                    
                        If gl_Disable_HIP_debug_log = False Then
                            For Each site In TheExec.sites.Active
                                 TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Current = " & measureCurrent.Pins(Pin).Value(site) & ", RAK val = " & CurrentJob_Card_RAK.Pins(Pin).Value(site)
                            Next site
                        End If
                    
                        measureCurrent.Pins(Pin) = measureCurrent.Pins(Pin).Multiply(CurrentJob_Card_RAK.Pins(Pin)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(Pin))
                        
                    Next Pin
                Else
                    For i = 0 To UBound(MeasI.Setup_ByTypeByPin.PPMU)
                        If CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) > 0.8 Then
                            'If LCase(Pin) Like "*1p2*" Then
                            '    Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio12_grp!!!
                            'Else
                                Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio18_grp!!!
                            'End If
                        Else
                            Vdiff = CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1)
                        End If
                        
'                        Dim Pins() As String
'                        Dim Pin_Cnt As Long
'                        Dim var As Variant
                        TheExec.DataManager.DecomposePinList MeasI.Setup_ByTypeByPin.PPMU(i).Pin, Pins, Pin_Cnt
                        For Each var In Pins
                        
                            If gl_Disable_HIP_debug_log = False Then
                                For Each site In TheExec.sites.Active
                                     TheExec.Datalog.WriteComment "Site[" & site & "]," & var & " Current = " & measureCurrent.Pins(var).Value(site) & ", RAK val = " & CurrentJob_Card_RAK.Pins(var).Value(site)
                                Next site
                            End If
                        
                            measureCurrent.Pins(var) = measureCurrent.Pins(var).Multiply(CurrentJob_Card_RAK.Pins(var)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(var))
                            
                        Next var
                    Next i
                
                End If
            End If
            
        ElseIf Instance_Data.RAK_Flag = R_PathWithContact Then
        
            If MeasI.Setup_ByTypeByPin_Flag = False Then
                If CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) > 0.8 Then
                    'If LCase(Pin) Like "*1p2*" Then
                    '    Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) ''vddio12_grp!!!
                    'Else
                        Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByType.PPMU.ForceValue1) ''vddio18_grp!!!
                    'End If
                Else
                    Vdiff = CDbl(MeasI.Setup_ByType.PPMU.ForceValue1)
                End If
                For Each Pin In measureCurrent.Pins
                
                    If gl_Disable_HIP_debug_log = False Then
                        For Each site In TheExec.sites
                            TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Current = " & measureCurrent.Pins(Pin).Value(site) & ", R_Path val = " & R_Path_PLD.Pins(Pin).Value(site)
                        Next site
                    End If
                
                    measureCurrent.Pins(Pin) = measureCurrent.Pins(Pin).Multiply(R_Path_PLD.Pins(Pin)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(Pin))
                Next Pin
                                        
            Else
                If b_ForceDiffVolt = False Then
                    If CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1) > 0.8 Then
                        'If LCase(Pin) Like "*1p2*" Then
                        '    Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio12_grp!!!
                        'Else
                            Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1) ''vddio18_grp!!!
                        'End If
                    Else
                        Vdiff = CDbl(MeasI.Setup_ByTypeByPin.PPMU(0).ForceValue1)
                    End If
                    For Each Pin In measureCurrent.Pins
                    
                        If gl_Disable_HIP_debug_log = False Then
                            For Each site In TheExec.sites
                                TheExec.Datalog.WriteComment "Site[" & site & "]," & Pin & " Current = " & measureCurrent.Pins(Pin).Value(site) & ", R_Path val = " & R_Path_PLD.Pins(Pin).Value(site)
                            Next site
                        End If
                    
                        measureCurrent.Pins(Pin) = measureCurrent.Pins(Pin).Multiply(R_Path_PLD.Pins(Pin)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(Pin))
                    Next Pin
                Else
                    For i = 0 To UBound(MeasI.Setup_ByTypeByPin.PPMU)
                        If CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) > 0.8 Then
                            'If LCase(Pin) Like "*1p2*" Then
                            '    Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio12_grp!!!
                            'Else
                                Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP0_1").Voltage.Value - CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1) ''vddio18_grp!!!
                            'End If
                        Else
                            Vdiff = CDbl(MeasI.Setup_ByTypeByPin.PPMU(i).ForceValue1)
                        End If
'                        Dim Pins() As String
'                        Dim Pin_Cnt As Long
'                        Dim var As Variant
                        TheExec.DataManager.DecomposePinList MeasI.Setup_ByTypeByPin.PPMU(i).Pin, Pins, Pin_Cnt
                        For Each var In Pins
                        
                            If gl_Disable_HIP_debug_log = False Then
                                For Each site In TheExec.sites
                                    TheExec.Datalog.WriteComment "Site[" & site & "]," & var & " Current = " & measureCurrent.Pins(var).Value(site) & ", R_Path val = " & R_Path_PLD.Pins(var).Value(site)
                                Next site
                            End If
                        
                            measureCurrent.Pins(var) = measureCurrent.Pins(var).Multiply(R_Path_PLD.Pins(var)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(measureCurrent.Pins(var))
                        Next var
                    Next i
                End If
            End If
        
        End If


    End If
End If

Exit Function

err:

If AbortTest Then Exit Function Else Resume Next

End Function


Public Function HardIP_DCVI_MI_StoreAndRestoreCondition(Measure As Meas_Type, Instrument As Inst_Type, SaveCondition As Boolean) As Long
    
    Dim MI_Pin As String
    Dim PinName() As String
    Dim TypeName As String
    Dim i As Long

    MI_Pin = Measure.Pins.UVI80

    PinName = Split(MI_Pin, ",")
    
    ReDim Preserve Measure.SaveCondition(UBound(PinName)) 'For restore

    For i = 0 To UBound(PinName)
        TypeName = SortPinInstrument(PinName(i))
        If SaveCondition = True Then
            Measure.SaveCondition(i).Pin = PinName(i)
            Measure.SaveCondition(i).current = FormatNumber(TheHdw.DCVI.Pins(PinName(i)).current, 3)
            Measure.SaveCondition(i).SrcCurrentRange = FormatNumber(TheHdw.DCVI.Pins(PinName(i)).CurrentRange.Value, 3)
            If UCase(TheExec.DataManager.PinType(PinName(i))) = UCase("Power") Then Measure.SaveCondition(i).IfPowerPin = True
        Else
            If Measure.SaveCondition(i).IfPowerPin = True Then
                TheHdw.DCVI.Pins(PinName(i)).CurrentRange.Value = CDbl(Measure.SaveCondition(i).SrcCurrentRange)
                TheHdw.DCVI.Pins(PinName(i)).current = Measure.SaveCondition(i).current
            Else
                With TheHdw.DCVI.Pins(PinName(i))
                    .Voltage = 0
                    .current = pc_Def_UVI80_Init_MeasCurrRange ''Init the source's current after measurment- Carter, 20190503
                    .Gate = False
                    .Disconnect
                End With
            End If
        End If
    Next i
    
End Function

Public Function ReDefineDigSrcForCharacterization(ByRef DigSrc_Assignment As String) As String
    ' Dylan Edited by 20190726
    Dim i, j As Integer
    Dim StringTemp As String
    Dim StringDict() As String
    Dim StringSplit() As String
    
    Dim OriDigSrcSplit() As String
    Dim NewDigSrcSplit() As String
    
    OriDigSrcSplit = Split(DigSrc_Assignment, ";")
    
    If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Type.Value = "Global Spec" And _
    TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Name Like "*DigSrc*" Then
        StringTemp = TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).InterposeFunctions.PrePoint.Arguments
    Else
        StringTemp = TheExec.DevChar.Setups.Item(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_Y).InterposeFunctions.PrePoint.Arguments
    End If
    
    StringTemp = Replace(StringTemp, "[", ":")
    StringTemp = Replace(StringTemp, "]", "")
    StringSplit = Split(StringTemp, ":")
    NewDigSrcSplit = Split(StringSplit(3), ";")

      
    For i = 0 To UBound(NewDigSrcSplit)
         For j = 0 To UBound(OriDigSrcSplit)
             StringDict = Split(OriDigSrcSplit(j), "=")
             If StringDict(0) = NewDigSrcSplit(i) Then
                 StringDict(1) = "=" & StringSplit(2)
                 OriDigSrcSplit(j) = StringDict(0) & StringDict(1)
             End If
         Next j
     Next i
     
     For i = 0 To UBound(OriDigSrcSplit)
         If i = 0 Then
             DigSrc_Assignment = OriDigSrcSplit(i)
         Else
             DigSrc_Assignment = DigSrc_Assignment & ";" & OriDigSrcSplit(i)
         End If
     Next i
    
End Function


'20190710, Modified for BV pass/fail flag Merging into HardIP function: TMPS, ELB, Oscar
Public Function Update_BC_PassFail_Flag(Optional PaternStartOnly As Boolean = False)


Dim inst_name As String
Dim Temp_Result As New SiteLong


inst_name = TheExec.DataManager.instanceName

If (inst_name Like "*BV") Then

    If PaternStartOnly Then
        Temp_Result = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        
        BV_Pass = BV_Pass.LogicalAnd(Temp_Result)
    Else
        Temp_Result = TheExec.Flow.LastIndividualTestResult
        
        Temp_Result = Temp_Result.Negate.Add(2)
        
        BV_Pass = BV_Pass.LogicalAnd(Temp_Result)
    End If
End If


End Function
