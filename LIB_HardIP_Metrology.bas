Attribute VB_Name = "LIB_HardIP_Metrology"
Type MTRSNS_Matrix
    ROT_Matrix() As Double
    ROV_Matrix() As Double
    ROT_a_max_min_Matrix() As Double
    ROV_a_max_min_Matrix() As Double
End Type
Public MetrologySense_Matrix() As MTRSNS_Matrix
Public Flag_TMPS_1st_Run As Boolean

Type DDDDDD
     MTR_DDDDD() As New DSPWave
     MTR_AAAAA() As New DSPWave
     
End Type

Public Try_Dictemp() As DDDDDD

Public Function MetrologyTMPS_Measurement_Process(Pat As String, srcPin As PinList, code() As SiteLong, ByRef Res() As SiteDouble, TrimCodeSize As Long, NumberOfMeasV As Integer, ByRef Rtn_MeasVolt() As PinListData, DigSrc_Sample_Size As Long, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName() As String, MeasV_WaitTime As String)
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
            Res(i) = Abs(Rtn_MeasVolt(i).Pins(0).Value - Rtn_MeasVolt(i).Pins(1).Value)
            For j = 0 To Rtn_MeasVolt(i).Pins.Count - 1
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(i) & ", Pin : " & Rtn_MeasVolt(i).Pins(j) & ", Voltage = " & Rtn_MeasVolt(i).Pins(j).Value
            Next j
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(i) & ", Voltage Difference = " & Res(i)
        Next site
        TheHdw.Digital.Patgen.Continue 0, cpuA
    Next i
    TheHdw.Digital.Patgen.HaltWait
    ByPassTestLimit = False
    glb_Disable_CurrRangeSetting_Print = False

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in MetrologyTMPS_Measurement_Process"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function MTRSNS_IDLoading()
Dim Max_Rows_Count As Long
Dim Max_Columns_Count As Long
Dim MTRSNS_Matrix_Range As Variant
Dim MTRSNS_Matrix_Sheet As Worksheet: Set MTRSNS_Matrix_Sheet = Sheets("MTRSNS_Matrix")
Dim i As Integer: i = 0
Dim MTX_ID_FIX As Long
Dim MTR_Version As Long
Dim First_ID As Long
Dim MTX_ID_VAR As Long
Dim MTR_T1_Version As Long

With MTRSNS_Matrix_Sheet
    Max_Rows_Count = .UsedRange.Rows.Count
    Max_Columns_Count = .UsedRange.Columns.Count
    MTRSNS_Matrix_Range = .range(.Cells(1, 1), .Cells(Max_Rows_Count, Max_Columns_Count))
End With



For i = 1 To Max_Columns_Count

    If Cells(i, 1).Value Like "*Matrix ID*" Then
       MTX_ID_FIX = Cells(i, 3).Value
       MTR_Version = Cells(i + 1, 3).Value
       First_ID = i + 1
       Exit For
    End If
    
Next i


For i = First_ID To Max_Columns_Count

    If Cells(i, 1).Value Like "*Matrix ID*" Then
       MTX_ID_VAR = Cells(i, 3).Value
       MTR_T1_Version = Cells(i + 1, 3).Value
       First_ID = i + 1
       Exit For
    End If

Next i

End Function

Public Function MTRTMPS_Gain_AVG(InWf As DSPWave, StoreName_GainMean As String, Integer_Bit As Long) As Long
Dim DSP_Gain_Mean As New DSPWave
Dim DSP_Gain_Mean_Array(0) As Double
Dim DSP_Gain_Mean_Fuse As New DSPWave
Dim DSP_Gain_Mean_Fuse_Array(0) As Double
Dim TestNameInput As String
'Dim High_limit As Double: High_limit = Bin2Dec_rev(String(Integer_Bit - 1, "1"))
'Dim Low_limit As Double: Low_limit = -2 ^ (Integer_Bit - 1)

    For Each site In TheExec.sites.Active
        DSP_Gain_Mean_Array(0) = InWf.CalcMean
        DSP_Gain_Mean.Data = DSP_Gain_Mean_Array
'        If DSP_OffSet_Mean_Array(0) < Low_limit Then
'            DSP_OffSet_Mean_Fuse_Array(0) = 2 ^ (Integer_Bit) + FormatNumber(Low_limit)
'        ElseIf DSP_OffSet_Mean_Array(0) >= Low_limit And DSP_OffSet_Mean_Array(0) < 0 Then
'            DSP_OffSet_Mean_Fuse_Array(0) = 2 ^ (Integer_Bit) + FormatNumber(DSP_OffSet_Mean_Array(0))
'        ElseIf DSP_OffSet_Mean_Array(0) < High_limit And DSP_OffSet_Mean_Array(0) >= 0 Then
'            DSP_OffSet_Mean_Fuse_Array(0) = FormatNumber(DSP_OffSet_Mean_Array(0))
'        Else
'            DSP_OffSet_Mean_Fuse_Array(0) = FormatNumber(High_limit)
'        End If
'        DSP_OffSet_Mean_Fuse.Data = DSP_OffSet_Mean_Fuse_Array
    Next site
    
    TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
    TheExec.Flow.TestLimit resultVal:=DSP_Gain_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
'    Call AddStoredCaptureData(StoreName_OffSetMean, DSP_OffSet_Mean_Fuse)
End Function
Public Function Calc_MetrologyTMPS_OffSet(argc As Integer, argv() As String) As Long
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")
Dim DSP_OffSet_Mean As New DSPWave
Dim DSP_OffSet_Mean_Array(0) As Double
Dim DSP_OffSet_Mean_eFuse As New DSPWave
'Dim DSP_OffSet_Mean_Fuse_Array(0) As Double
Dim TestNameInput As String
Dim i As Long
    For i = 0 To UBound(InWf_Array)
        InWf_SiteDouble(i) = GetStoredData(InWf_Split(i) & "_para")
    Next i
    For Each site In TheExec.sites.Active
        For i = 0 To UBound(InWf_Array)
            InWf_Array(i) = InWf_SiteDouble(i)
        Next i
        InWf.Data = InWf_Array
        DSP_OffSet_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0)
        DSP_OffSet_Mean.Data = DSP_OffSet_Mean_Array
    Next site
    
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.Flow.TestLimit resultVal:=DSP_OffSet_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
    Call AddStoredCaptureData(argv(1) & "_" & CStr(TheExec.Flow.var(argv(2)).Value), DSP_OffSet_Mean)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_OffSet_Mean_eFuse, DSP_OffSet_Mean, 18, 0)
    Call AddStoredCaptureData(argv(1) & "_eFuse_" & CStr(TheExec.Flow.var(argv(2)).Value), DSP_OffSet_Mean_eFuse)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyTMPS_OffSet"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologyTMPS_Gain(argc As Integer, argv() As String) As Long
Dim InWf As New DSPWave
Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")

Dim DSP_Gain_Mean As New DSPWave
Dim DSP_Gain_Mean_Array(0) As Double
Dim DSP_OffSet_Mean As New DSPWave
Dim DSP_OffSet_Mean_Array() As Double
Dim DSP_Gain_Mean_Final As New DSPWave
Dim DSP_Gain_Mean_Final_Array(0) As Double
Dim TestNameInput As String
Dim i As Long

For i = 0 To UBound(InWf_Array)
    InWf_SiteDouble(i) = GetStoredData(InWf_Split(i) & "_para")
Next i

DSP_OffSet_Mean = GetStoredCaptureData(argv(1) & "_" & CStr(TheExec.Flow.var(argv(2)).Value))
For Each site In TheExec.sites.Active
    DSP_OffSet_Mean_Array = DSP_OffSet_Mean.Data
    For i = 0 To UBound(InWf_Array)
        InWf_Array(i) = InWf_SiteDouble(i)
    Next i
    InWf.Data = InWf_Array
    DSP_Gain_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0) - DSP_OffSet_Mean_Array(0)
    If DSP_Gain_Mean_Array(0) < 0 Then: DSP_Gain_Mean_Array(0) = 0
    DSP_Gain_Mean.Data = DSP_Gain_Mean_Array
Next site

TestNameInput = Report_TName_From_Instance("CalcC", "X", Replace(argv(3), "_", ""), 0, 0)
TheExec.Flow.TestLimit resultVal:=DSP_Gain_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

Call AddStoredCaptureData(argv(3) & "_" & CStr(TheExec.Flow.var(argv(2)).Value), DSP_Gain_Mean)

If TheExec.Flow.var(argv(2)).Value = argv(5) Then
    For Each site In TheExec.sites.Active
        InWf.Clear
        For i = argv(4) To argv(5)
            If i = argv(4) Then
                InWf = GetStoredCaptureData(argv(3) & "_" & i)
            Else
                InWf = InWf.Concatenate(GetStoredCaptureData(argv(3) & "_" & i))
            End If
        Next i
        
    DSP_Gain_Mean_Final_Array(0) = FormatNumber(InWf.CalcMean, 0)
    DSP_Gain_Mean_Final.Data = DSP_Gain_Mean_Final_Array
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", "", Replace(argv(3) & "_AVG", "_", ""))
    TheExec.Flow.TestLimit resultVal:=DSP_Gain_Mean_Final.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    Call AddStoredCaptureData(argv(3), DSP_Gain_Mean_Final)
End If

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyTMPS_Gain"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function MTRTMPS_DSSCOUT_AVG(InWf As DSPWave, StoreName_DSSC_OUT_Mean As String) As Long
Dim DSP_DSSCOUT_Mean As New DSPWave
Dim DSP_DSSCOUT_Mean_Array(0) As Double
Dim TestNameInput As String

For Each site In TheExec.sites.Active
    DSP_DSSCOUT_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0)
    DSP_DSSCOUT_Mean.Data = DSP_DSSCOUT_Mean_Array
Next site

TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
TheExec.Flow.TestLimit resultVal:=DSP_DSSCOUT_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
Call AddStoredCaptureData(StoreName_DSSC_OUT_Mean, DSP_DSSCOUT_Mean)

End Function
Public Function MetrologyGR_Measurement_Process(Pat As String, srcPin As PinList, code As SiteLong, Res As SiteDouble, TrimCodeSize As Long, NumberOfMeasV As Integer, ByRef Rtn_MeasVolt() As PinListData, DigSrc_Sample_Size As Long, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName As String, MeasV_WaitTime As String)
    Dim sigName As String, srcWave As New DSPWave, site As Variant
    Dim InDSPwave As New DSPWave
    Dim i As Long, j As Long
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array() As Long: ReDim srcwave_array(TrimCodeSize - 1)
    srcWave.CreateConstant 0, TrimCodeSize, DspLong
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
    Call GeneralDigSrcSetting(Pat, srcPin, DigSrc_Sample_Size, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, "", "", InDSPwave, "")
    
    TheHdw.Patterns(Pat).start
    
    For i = 0 To NumberOfMeasV - 1
        Instance_Data.TestSeqNum = i
        TheHdw.Digital.Patgen.FlagWait cpuA, 0

        Rtn_MeasVolt(i) = HardIP_MeasureVolt
        Call DebugPrintFunc_PPMU("")

        For Each site In TheExec.sites

            If i = NumberOfMeasV - 1 Then Res = Abs(Rtn_MeasVolt(i).Pins(0).Value - Rtn_MeasVolt(i - 1).Pins(0).Value)
            For j = 0 To Rtn_MeasVolt(i).Pins.Count - 1
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Pin : " & Rtn_MeasVolt(i).Pins(j) & ", Voltage = " & Rtn_MeasVolt(i).Pins(j).Value
            Next j
            If gl_Disable_HIP_debug_log = False Then If i = NumberOfMeasV - 1 Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Voltage Difference = " & Res
        Next site

        TheHdw.Digital.Patgen.Continue 0, cpuA
    Next i
    TheHdw.Digital.Patgen.HaltWait
    ByPassTestLimit = False
    glb_Disable_CurrRangeSetting_Print = False

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in MetrologyGR_Measurement_Process"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologyGR_TRC(argc As Integer, argv() As String) As Long
    Dim DSP_APM_Count_Out_0 As New DSPWave: DSP_APM_Count_Out_0 = GetStoredCaptureData(argv(0))
    Dim DSP_APM_Count_Out_1 As New DSPWave: DSP_APM_Count_Out_1 = GetStoredCaptureData(argv(1))
    Dim DictionaryName As String: DictionaryName = argv(2)
    Dim DSP_T1 As New DSPWave
    Dim DSP_T2 As New DSPWave
    Dim DSP_TRC As New DSPWave
    Dim DSP_TRC_ERR As New DSPWave
    Dim DSP_TRC_ERR_eFuse As New DSPWave
    Dim TestNameInput As String
    Dim site As Variant
    For Each site In TheExec.sites
        DSP_T1 = DSP_APM_Count_Out_0.ConvertStreamTo(tldspParallel, DSP_APM_Count_Out_0.SampleSize, 0, Bit0IsMsb).Multiply(375000).Add(0.001).Reciprocate
        DSP_T2 = DSP_APM_Count_Out_1.ConvertStreamTo(tldspParallel, DSP_APM_Count_Out_1.SampleSize, 0, Bit0IsMsb).Multiply(375000).Add(0.001).Reciprocate
        DSP_TRC = DSP_T2.Subtract(DSP_T1)
        DSP_TRC_ERR = DSP_TRC.Subtract(0.0000000017).Divide(0.0000000017).Multiply(100).Divide(0.5)
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", "")
    TheExec.Flow.TestLimit resultVal:=DSP_T1.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", "")
    TheExec.Flow.TestLimit resultVal:=DSP_T2.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", "")
    TheExec.Flow.TestLimit resultVal:=DSP_TRC.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", "")
    TheExec.Flow.TestLimit resultVal:=DSP_TRC_ERR.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_TRC_ERR_eFuse, DSP_TRC_ERR, 8, 0)
    Call AddStoredCaptureData(DictionaryName, DSP_TRC_ERR_eFuse)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyGR_TRC"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologySense_Frequency(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim TestNameInput As String
    Dim site As Variant
    Dim MetrologySense_Frequency() As New DSPWave: ReDim MetrologySense_Frequency(argc - 1)
    Dim SubBlockName As String: SubBlockName = Split(TheExec.DataManager.instanceName, "_")(1)
    For i = 0 To argc - 1
        For Each site In TheExec.sites
            DSPWave_Dict = GetStoredCaptureData(argv(i))
            MetrologySense_Frequency(i) = DSPWave_Dict.ConvertStreamTo(tldspParallel, 21, 0, Bit0IsMsb).Multiply(50000)
        Next
        Call AddStoredCaptureData(SubBlockName & "-" & Instance_Data.Tname(TheExec.Flow.TestLimitIndex), MetrologySense_Frequency(i))
        TestNameInput = Report_TName_From_Instance("CalcF", "")
        TheExec.Flow.TestLimit resultVal:=MetrologySense_Frequency(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
If UCase(TheExec.DataManager.instanceName) = "MTRSNS_ASGMTRT1P3VDDCPUSRAMV1P250VDDPCPUV0P750VDDECPUV0P750_PP_SCYA0_C_FULP_AN_MT03_DLL_JTG_COD_ALLFV_SI_ASGMTR_T1P3_LV" Then Flag_TMPS_1st_Run = False
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologySense_Frequency"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologyTMPS_Temperature(argc As Integer, argv() As String) As Long
    
    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim Temperature_Array(0) As Double
    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim Temperature_Dictionary() As String
    For i = 0 To UBound(Temperature_Sensor)
        Temperature = GetStoredData(Temperature_Sensor(i) + "_para")
        For Each site In TheExec.sites
            Temperature_Array(0) = Temperature / 64
            DSP_Temperature(i).Data = Temperature_Array
        Next site
    Next i
    If argc = 2 Then
        Temperature_Dictionary = Split(argv(1), "+")
        For i = 0 To UBound(Temperature_Sensor)
            For Each site In TheExec.sites
                DSP_MTRSNS_Temperature(i) = DSP_Temperature(i).Multiply(8)
            Next site
            'If UCase(TheExec.CurrentJob) = "" Then '<512
                'Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_MTRSNS_Temperature_eFuse(i), DSP_MTRSNS_Temperature(i), 10, 0)
            'Else
                Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_MTRSNS_Temperature_eFuse(i), DSP_MTRSNS_Temperature(i), 11, 0)
            'End If
            Call AddStoredCaptureData(Temperature_Dictionary(i), DSP_MTRSNS_Temperature_eFuse(i))
        Next i
    End If
    
    For i = 0 To UBound(Temperature_Sensor)
        TestNameInput = Report_TName_From_Instance("CalcC", "")
        TheExec.Flow.TestLimit resultVal:=DSP_Temperature(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    If TheExec.Flow.EnableWord("TMPS_Monitor") = True Then
        TheHdw.Pins("all_digital").Digital.InitState = chInitLo
        TheHdw.DCVS.Pins("VDD_SRAM_GPU,VDD_GPU,VDD_ECPU,VDD_PCPU,VDD_CPU_SRAM").Voltage.Main.Value = 0.5
        If Not (Flag_TMPS_1st_Run) Then
            TheHdw.Wait 5
            Flag_TMPS_1st_Run = True
        Else
            TheHdw.Wait 1
        End If
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyTMPS_Temperature"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologyTMPS_Vref_Error(argc As Integer, argv() As String) As Long
    Dim Str_Vref As String: Str_Vref = argv(0)
    Dim Fuse_BitCount As Long: Fuse_BitCount = CLng(argv(1))
    
    Dim SiteDbl_Vref As New SiteDouble: SiteDbl_Vref = GetStoredMeasurement(Str_Vref)
    Dim SiteDbl_Vref_Error As New SiteDouble: SiteDbl_Vref_Error = SiteDbl_Vref.Divide(0.8).Subtract(1).Multiply(100).Divide(0.125)
    Dim High_limit As Double: High_limit = Bin2Dec_rev(String(Fuse_BitCount - 1, "1"))
    Dim Low_limit As Double: Low_limit = -2 ^ (Fuse_BitCount - 1)
    Dim site As Variant
    For Each site In TheExec.sites
        SiteDbl_Vref_Error = Floor(SiteDbl_Vref_Error)
    Next site
    Dim TestNameInput As String: TestNameInput = Report_TName_From_Instance("CalcC", "X", "vreferr", 0, 0)
    
    TheExec.Flow.TestLimit resultVal:=SiteDbl_Vref_Error, lowVal:=Low_limit, hiVal:=High_limit, Tname:=TestNameInput, ForceResults:=tlForceNone

    Dim SiteDbl_Vref_Error_Fuse As New DSPWave
    Dim SiteDbl_Vref_Error_Fuse_Array(0) As Long

    For Each site In TheExec.sites
        SiteDbl_Vref_Error = FormatNumber(SiteDbl_Vref_Error, 0)
        If SiteDbl_Vref_Error < Low_limit Then
            SiteDbl_Vref_Error_Fuse_Array(0) = 2 ^ (Fuse_BitCount) + FormatNumber(Low_limit)
        ElseIf SiteDbl_Vref_Error >= Low_limit And SiteDbl_Vref_Error < 0 Then
            SiteDbl_Vref_Error_Fuse_Array(0) = 2 ^ (Fuse_BitCount) + FormatNumber(SiteDbl_Vref_Error)
        ElseIf SiteDbl_Vref_Error < High_limit And SiteDbl_Vref_Error >= 0 Then
            SiteDbl_Vref_Error_Fuse_Array(0) = FormatNumber(SiteDbl_Vref_Error)
        Else
            SiteDbl_Vref_Error_Fuse_Array(0) = FormatNumber(High_limit)
        End If
        SiteDbl_Vref_Error_Fuse.Data = SiteDbl_Vref_Error_Fuse_Array
    Next site
    
    Call AddStoredCaptureData(argv(argc - 1), SiteDbl_Vref_Error_Fuse)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyTMPS_Vref_Error"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_MetrologyTMPS_Coefficients(argc As Integer, argv() As String) As Long
    Dim InWf As New DSPWave
    Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
    Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
    Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")
    Dim TempVal As Long
    Dim DSP_Offset As New DSPWave: DSP_Offset = GetStoredCaptureData(argv(1))
    Dim DSP_Gain As New DSPWave: DSP_Gain = GetStoredCaptureData(argv(2))
    Dim Integer_Bit As Long: Integer_Bit = argv(4)
    Dim Fractional_Bit As Long: Fractional_Bit = argv(5)
    Dim DSP_ADC_Temperature_Sensor_Raw_Data As New DSPWave
    Dim ADC_Temperature_Sensor_Raw_Data_Array(0) As Double
    Dim DSP_V0 As New DSPWave
    Dim DSP_V1 As New DSPWave
    Dim DSP_X0A As New DSPWave
    Dim DSP_X0 As New DSPWave
    Dim DSP_a0Cal As New DSPWave
    Dim A0 As Double
    Dim a1 As Double
    Dim a2 As Double
    Dim a3 As Double
    Dim DSP_Coefficient_C0 As New DSPWave
    Dim DSP_Coefficient_C1 As New DSPWave
    Dim DSP_Coefficient_C2 As New DSPWave
    Dim DSP_Coefficient_C3 As New DSPWave
    Dim DSP_Coefficient_C0_eFuse As New DSPWave
    Dim DSP_Coefficient_C1_eFuse As New DSPWave
    Dim DSP_Coefficient_C2_eFuse As New DSPWave
    Dim DSP_Coefficient_C3_eFuse As New DSPWave
    Dim Array_Coefficient_C0_eFuse() As Double
    Dim Array_Coefficient_C1_eFuse() As Double
    Dim Array_Coefficient_C2_eFuse() As Double
    Dim Array_Coefficient_C3_eFuse() As Double
    Dim DSP_Coefficient_C0_Source As New DSPWave
    Dim DSP_Coefficient_C1_Source As New DSPWave
    Dim DSP_Coefficient_C2_Source As New DSPWave
    Dim DSP_Coefficient_C3_Source As New DSPWave
    Dim Array_Coefficient_C0_Source() As Long: ReDim Array_Coefficient_C0_Source(Integer_Bit + Fractional_Bit - 1)
    Dim Array_Coefficient_C1_Source() As Long: ReDim Array_Coefficient_C1_Source(Integer_Bit + Fractional_Bit - 1)
    Dim Array_Coefficient_C2_Source() As Long: ReDim Array_Coefficient_C2_Source(Integer_Bit + Fractional_Bit - 1)
    Dim Array_Coefficient_C3_Source() As Long: ReDim Array_Coefficient_C3_Source(Integer_Bit + Fractional_Bit - 1)
    Dim BKM_Decode As String
    For i = 0 To UBound(InWf_Array)
        InWf_SiteDouble(i) = GetStoredData(InWf_Split(i) & "_para")
    Next i
    For Each site In TheExec.sites.Active
        For i = 0 To UBound(InWf_Array)
            InWf_Array(i) = InWf_SiteDouble(i)
        Next i
        InWf.Data = InWf_Array
        ADC_Temperature_Sensor_Raw_Data_Array(0) = FormatNumber(InWf.CalcMean, 0)
        DSP_ADC_Temperature_Sensor_Raw_Data.Data = ADC_Temperature_Sensor_Raw_Data_Array
    Next site
    For Each site In TheExec.sites
        BKM_Decode = gS_BKM_IEDA
        Exit For
    Next site
    
    If BKM_Decode = "0" Then
        A0 = -17.95433315
        a1 = 423.91403942
        a2 = -128.66284472
        a3 = 15.66328344
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM4.2 ) **********************"
        TheExec.Datalog.WriteComment "  a0 = -17.95433315(0xFFDC176)"
        TheExec.Datalog.WriteComment "  a1 = 423.91403942(0x34FD40)"
        TheExec.Datalog.WriteComment "  a2 = -128.66284472(0xFEFEACA)"
        TheExec.Datalog.WriteComment "  a3 = 15.66328344(0x1F53A)"
        TheExec.Datalog.WriteComment "*************************************************************************************"
'    ElseIf BKM_Decode = "1" Then
'        A0 = -17.95433315
'        a1 = 423.91403942
'        a2 = -128.66284472
'        a3 = 15.66328344
'        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM4.5 ) **********************"
'        TheExec.Datalog.WriteComment "  a0 = -17.95433315(0xFFDC176)"
'        TheExec.Datalog.WriteComment "  a1 = 423.91403942(0x34FD40)"
'        TheExec.Datalog.WriteComment "  a2 = -128.66284472(0xFEFEACA)"
'        TheExec.Datalog.WriteComment "  a3 = 15.66328344(0x1F53A)"
'        TheExec.Datalog.WriteComment "*************************************************************************************"
'    ElseIf BKM_Decode = "2" Then
'        A0 = -17.95433315
'        a1 = 423.91403942
'        a2 = -128.66284472
'        a3 = 15.66328344
'        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( BKM4.6 ) **********************"
'        TheExec.Datalog.WriteComment "  a0 = -17.95433315(0xFFDC176)"
'        TheExec.Datalog.WriteComment "  a1 = 423.91403942(0x34FD40)"
'        TheExec.Datalog.WriteComment "  a2 = -128.66284472(0xFEFEACA)"
'        TheExec.Datalog.WriteComment "  a3 = 15.66328344(0x1F53A)"
'        TheExec.Datalog.WriteComment "*************************************************************************************"
    Else
        TheExec.Datalog.WriteComment "********************** Coefficients for Calibration ( Unknown BKM ) **********************"
'        DSP_Coefficient_C0.CreateConstant 0, 1, DspDouble
'        DSP_Coefficient_C1.CreateConstant 0, 1, DspDouble
'        DSP_Coefficient_C2.CreateConstant 0, 1, DspDouble
'        DSP_Coefficient_C3.CreateConstant 0, 1, DspDouble
'        GoTo Unknown_BKM
        
        For Each site In TheExec.sites
            TheExec.sites.Item(site).FlagState("F_HardIP_MTRTSNS_Unknown_BKM_Flag") = logicTrue
        Next site
        Exit Function
    End If

    
    
    
    For Each site In TheExec.sites
        DSP_V0 = DSP_Offset.Divide(2 ^ 13)
        DSP_V1 = DSP_Gain.Divide(2 ^ 14)
        DSP_X0A = DSP_ADC_Temperature_Sensor_Raw_Data.Divide(2 ^ 13)
        DSP_X0 = DSP_X0A.Subtract(DSP_V0).Divide(DSP_V1.Add(0.0000000001)) 'x0=(x0a-v0)/v1;
        DSP_a0Cal = DSP_X0.Multiply(-a1).Add(DSP_X0.Square.Multiply(-a2)).Add(DSP_X0.Square.Multiply(DSP_X0).Multiply(-a3)).Add(25).Add(273.15) 'a0cal=25+273.15-a1*x0-a2*x0^2-a3*x0^3
        DSP_Coefficient_C0 = DSP_a0Cal.Subtract(DSP_V0.Divide(DSP_V1.Add(0.0000000001)).Multiply(a1)) 'coef_c0=a0cal-a1*v0/v1;
        DSP_Coefficient_C1 = DSP_V1.Add(0.0000000001).Reciprocate.Multiply(a1).Subtract(DSP_V0.Divide(DSP_V1.Add(0.0000000001).Square).Multiply(2).Multiply(a2)) 'coef_c1=a1/v1-2*a2*v0/v1^2;
        DSP_Coefficient_C2 = DSP_V1.Add(0.0000000001).Square.Reciprocate.Multiply(a2).Subtract(DSP_V0.Divide(DSP_V1.Add(0.0000000001).Square.Multiply(DSP_V1.Add(0.0000000001))).Multiply(3).Multiply(a3)) 'coef_c2=a2/v1^2-3*a3*v0/v1^3;
        DSP_Coefficient_C3 = DSP_V1.Add(0.0000000001).Square.Multiply(DSP_V1.Add(0.0000000001)).Reciprocate.Multiply(a3) 'coef_c3=a3/v1^3
    Next site
Unknown_BKM:
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_Coefficient_C0_eFuse, DSP_Coefficient_C0, Integer_Bit, Fractional_Bit)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_Coefficient_C1_eFuse, DSP_Coefficient_C1, Integer_Bit, Fractional_Bit)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_Coefficient_C2_eFuse, DSP_Coefficient_C2, Integer_Bit, Fractional_Bit)
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_Coefficient_C3_eFuse, DSP_Coefficient_C3, Integer_Bit, Fractional_Bit)
    For Each site In TheExec.sites
        Array_Coefficient_C0_eFuse = DSP_Coefficient_C0_eFuse.Data
        TempVal = Array_Coefficient_C0_eFuse(0)
            For i = 0 To Integer_Bit + Fractional_Bit - 1
                Array_Coefficient_C0_Source(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        DSP_Coefficient_C0_Source.Data = Array_Coefficient_C0_Source
        Array_Coefficient_C1_eFuse = DSP_Coefficient_C1_eFuse.Data
        TempVal = Array_Coefficient_C1_eFuse(0)
            For i = 0 To Integer_Bit + Fractional_Bit - 1
                Array_Coefficient_C1_Source(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        DSP_Coefficient_C1_Source.Data = Array_Coefficient_C1_Source
        Array_Coefficient_C2_eFuse = DSP_Coefficient_C2_eFuse.Data
        TempVal = Array_Coefficient_C2_eFuse(0)
            For i = 0 To Integer_Bit + Fractional_Bit - 1
                Array_Coefficient_C2_Source(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        DSP_Coefficient_C2_Source.Data = Array_Coefficient_C2_Source
        Array_Coefficient_C3_eFuse = DSP_Coefficient_C3_eFuse.Data
        TempVal = Array_Coefficient_C3_eFuse(0)
            For i = 0 To Integer_Bit + Fractional_Bit - 1
                Array_Coefficient_C3_Source(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        DSP_Coefficient_C3_Source.Data = Array_Coefficient_C3_Source
    Next site
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.Flow.TestLimit resultVal:=DSP_ADC_Temperature_Sensor_Raw_Data.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    TestNameInput = Report_TName_From_Instance("CalcC", "", Replace(argv(3) & "_coeff0", "_", ""))
    TheExec.Flow.TestLimit resultVal:=DSP_Coefficient_C0.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", "", Replace(argv(3) & "_coeff1", "_", ""))
    TheExec.Flow.TestLimit resultVal:=DSP_Coefficient_C1.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", "", Replace(argv(3) & "_coeff2", "_", ""))
    TheExec.Flow.TestLimit resultVal:=DSP_Coefficient_C2.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", "", Replace(argv(3) & "_coeff3", "_", ""))
    TheExec.Flow.TestLimit resultVal:=DSP_Coefficient_C3.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow

    Call AddStoredCaptureData(argv(3) & "_coeff0", DSP_Coefficient_C0_eFuse)
    Call AddStoredCaptureData(argv(3) & "_coeff1", DSP_Coefficient_C1_eFuse)
    Call AddStoredCaptureData(argv(3) & "_coeff2", DSP_Coefficient_C2_eFuse)
    Call AddStoredCaptureData(argv(3) & "_coeff3", DSP_Coefficient_C3_eFuse)
    Call AddStoredCaptureData(argv(3) & "_coeff0_src", DSP_Coefficient_C0_Source)
    Call AddStoredCaptureData(argv(3) & "_coeff1_src", DSP_Coefficient_C1_Source)
    Call AddStoredCaptureData(argv(3) & "_coeff2_src", DSP_Coefficient_C2_Source)
    Call AddStoredCaptureData(argv(3) & "_coeff3_src", DSP_Coefficient_C3_Source)

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyTMPS_Coefficients"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function MTRSNS_Matrix_Loading()
Dim MTRSNS_Matrix_Sheet As Worksheet: Set MTRSNS_Matrix_Sheet = Sheets("MTRSNS_Matrix")
Dim Column_Index As Long: Column_Index = 1
Dim Row_Index As Long: Row_Index = 1
Dim Matrix_Index As Long
Dim MetrologySense_Matrix_Index As Long: MetrologySense_Matrix_Index = 0
Dim MTRSNS_Matrix_Range As Variant
Dim Max_Rows_Count As Long
Dim Max_Columns_Count As Long

With MTRSNS_Matrix_Sheet
    Max_Rows_Count = .UsedRange.Rows.Count
    Max_Columns_Count = .UsedRange.Columns.Count
    MTRSNS_Matrix_Range = .range(.Cells(1, 1), .Cells(Max_Rows_Count, Max_Columns_Count))
End With

While Row_Index <> Max_Rows_Count
    ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index)
    If UCase(MTRSNS_Matrix_Range(Row_Index, 1)) Like "*ROT*" Then
        Row_Index = Row_Index + 1
        Matrix_Index = 0
        While MTRSNS_Matrix_Range(Row_Index, 1) <> ""
            Column_Index = 1
            Do While Column_Index <= Max_Columns_Count
                If MTRSNS_Matrix_Range(Row_Index, Column_Index) = "" Then Exit Do
                ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index).ROT_Matrix(Matrix_Index)
                MetrologySense_Matrix(MetrologySense_Matrix_Index).ROT_Matrix(Matrix_Index) = MTRSNS_Matrix_Range(Row_Index, Column_Index)
                Column_Index = Column_Index + 1
                Matrix_Index = Matrix_Index + 1
            Loop
            Row_Index = Row_Index + 1
        Wend
        Do
            Row_Index = Row_Index + 1
        Loop Until UCase(MTRSNS_Matrix_Range(Row_Index, 1)) Like "*A_MAX*"
        Row_Index = Row_Index + 1
        Matrix_Index = 0
        While MTRSNS_Matrix_Range(Row_Index, 1) <> ""
            Column_Index = 1
            Do While Column_Index <= Max_Columns_Count
                If MTRSNS_Matrix_Range(Row_Index, Column_Index) = "" Then Exit Do
                ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index).ROT_a_max_min_Matrix(Matrix_Index)
                MetrologySense_Matrix(MetrologySense_Matrix_Index).ROT_a_max_min_Matrix(Matrix_Index) = MTRSNS_Matrix_Range(Row_Index, Column_Index)
                Column_Index = Column_Index + 1
                Matrix_Index = Matrix_Index + 1
            Loop
            Row_Index = Row_Index + 1
        Wend
    ElseIf UCase(MTRSNS_Matrix_Range(Row_Index, 1)) Like "*ROV*" Then
        Row_Index = Row_Index + 1
        Matrix_Index = 0
        While MTRSNS_Matrix_Range(Row_Index, 1) <> ""
            Column_Index = 1
            Do While Column_Index <= Max_Columns_Count
                If MTRSNS_Matrix_Range(Row_Index, Column_Index) = "" Then Exit Do
                ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index).ROV_Matrix(Matrix_Index)
                MetrologySense_Matrix(MetrologySense_Matrix_Index).ROV_Matrix(Matrix_Index) = MTRSNS_Matrix_Range(Row_Index, Column_Index)
                Column_Index = Column_Index + 1
                Matrix_Index = Matrix_Index + 1
            Loop
            Row_Index = Row_Index + 1
        Wend
        Do
            Row_Index = Row_Index + 1
        Loop Until UCase(MTRSNS_Matrix_Range(Row_Index, 1)) Like "*A_MAX*"
        Row_Index = Row_Index + 1
        Matrix_Index = 0
        While MTRSNS_Matrix_Range(Row_Index, 1) <> ""
            Column_Index = 1
            Do While Column_Index <= Max_Columns_Count
                If MTRSNS_Matrix_Range(Row_Index, Column_Index) = "" Then Exit Do
                ReDim Preserve MetrologySense_Matrix(MetrologySense_Matrix_Index).ROV_a_max_min_Matrix(Matrix_Index)
                MetrologySense_Matrix(MetrologySense_Matrix_Index).ROV_a_max_min_Matrix(Matrix_Index) = MTRSNS_Matrix_Range(Row_Index, Column_Index)
                Column_Index = Column_Index + 1
                Matrix_Index = Matrix_Index + 1
            Loop
            Row_Index = Row_Index + 1
        Wend
    ElseIf Replace(UCase(MTRSNS_Matrix_Range(Row_Index, 1)), " ", "") Like "*MATRIXID*" Then
        MetrologySense_Matrix_Index = MetrologySense_Matrix_Index + 1
    End If
    Row_Index = Row_Index + 1
Wend
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in MTRSNS_Matrix_Loading"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologySense_Compression(argc As Integer, argv() As String) As Long
If UCase(TheExec.CurrentJob) = "CP1" Or UCase(TheExec.CurrentJob) = "FT1" Then Exit Function 'Tonga
Dim SweepCondition_Split() As String: SweepCondition_Split = Split(argv(0), "+")
Dim Sensor As String: Sensor = argv(1)
Dim MetrologySense_ROT_Frequency As New DSPWave
Dim MetrologySense_ROV_Frequency As New DSPWave
Dim MTRSNS_Matrix_Index As Long: MTRSNS_Matrix_Index = argv(2)
Dim MTRSNS_Matrix_ROT_Row As Long: MTRSNS_Matrix_ROT_Row = argv(3)
Dim MTRSNS_Matrix_ROT_Column As Long: MTRSNS_Matrix_ROT_Column = argv(4)
Dim MTRSNS_Matrix_ROV_Row As Long: MTRSNS_Matrix_ROV_Row = argv(5)
Dim MTRSNS_Matrix_ROV_Column As Long: MTRSNS_Matrix_ROV_Column = argv(6)
Dim a1 As New DSPWave
Dim a2 As New DSPWave
Dim a1_Compression As New DSPWave
Dim a2_Compression As New DSPWave
Dim a1_Compression_eFuse As New DSPWave
Dim a2_Compression_eFuse As New DSPWave
Dim Array_a1_Compression_eFuse() As Double
Dim Array_a2_Compression_eFuse() As Double
Dim a1_Compression_eFuse_Store() As New DSPWave: ReDim a1_Compression_eFuse_Store(MTRSNS_Matrix_ROT_Row - 1)
Dim a2_Compression_eFuse_Store() As New DSPWave: ReDim a2_Compression_eFuse_Store(MTRSNS_Matrix_ROV_Row - 1)
Dim a1_LogicalCompare As New DSPWave
Dim a1_LogicalCompare_Array() As Double
Dim a2_LogicalCompare As New DSPWave
Dim a2_LogicalCompare_Array() As Double
Dim DSP_MetrologySense_ROT_Matrix As New DSPWave: DSP_MetrologySense_ROT_Matrix.Data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROT_Matrix
Dim DSP_MetrologySense_ROV_Matrix As New DSPWave: DSP_MetrologySense_ROV_Matrix.Data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROV_Matrix
Dim DSP_ROT_a_max_min As New DSPWave: DSP_ROT_a_max_min.Data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROT_a_max_min_Matrix
Dim DSP_ROV_a_max_min As New DSPWave: DSP_ROV_a_max_min.Data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROV_a_max_min_Matrix
Dim a1_max() As Double
Dim a1_min() As Double
Dim a2_max() As Double
Dim a2_min() As Double
Dim i As Long
Dim site As Variant

For i = 0 To UBound(SweepCondition_Split)
    If i = 0 Then
        MetrologySense_ROT_Frequency = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT")
        MetrologySense_ROV_Frequency = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV")
    Else
        For Each site In TheExec.sites
            MetrologySense_ROT_Frequency = MetrologySense_ROT_Frequency.Concatenate(GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT"))
            MetrologySense_ROV_Frequency = MetrologySense_ROV_Frequency.Concatenate(GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV"))
        Next site
    End If
Next i
For Each site In TheExec.sites
    MetrologySense_ROT_Frequency = MetrologySense_ROT_Frequency.Divide(10 ^ 9)
    MetrologySense_ROV_Frequency = MetrologySense_ROV_Frequency.Divide(10 ^ 9)
    a1 = DSP_MetrologySense_ROT_Matrix.MatrixMultiply(MTRSNS_Matrix_ROT_Row, MTRSNS_Matrix_ROT_Column, MetrologySense_ROT_Frequency)
    a2 = DSP_MetrologySense_ROV_Matrix.MatrixMultiply(MTRSNS_Matrix_ROV_Row, MTRSNS_Matrix_ROV_Column, MetrologySense_ROV_Frequency)
    a1_LogicalCompare_Array = a1.Data
    a1_max = DSP_ROT_a_max_min.Select(0, 2, MTRSNS_Matrix_ROT_Row).Data
    a1_min = DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row).Data
    a2_LogicalCompare_Array = a2.Data
    a2_max = DSP_ROV_a_max_min.Select(0, 2, MTRSNS_Matrix_ROV_Row).Data
    a2_min = DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row).Data
'    For i = 0 To UBound(a1_LogicalCompare_Array)
'        If a1_LogicalCompare_Array(i) > a1_max(i) Then
'            a1_LogicalCompare_Array(i) = a1_max(i)
'        ElseIf a1_LogicalCompare_Array(i) < a1_min(i) Then
'            a1_LogicalCompare_Array(i) = a1_min(i)
'        End If
'    Next i
'    For i = 0 To UBound(a2_LogicalCompare_Array)
'        If a2_LogicalCompare_Array(i) > a2_max(i) Then
'            a2_LogicalCompare_Array(i) = a2_max(i)
'        ElseIf a2_LogicalCompare_Array(i) < a2_min(i) Then
'            a2_LogicalCompare_Array(i) = a2_min(i)
'        End If
'    Next i
    a1_LogicalCompare.Data = a1_LogicalCompare_Array
    a2_LogicalCompare.Data = a2_LogicalCompare_Array
    a1_Compression = a1_LogicalCompare.Subtract(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row)).Divide(DSP_ROT_a_max_min.Select(0, 2, MTRSNS_Matrix_ROT_Row).Subtract(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row)))
    a2_Compression = a2_LogicalCompare.Subtract(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row)).Divide(DSP_ROV_a_max_min.Select(0, 2, MTRSNS_Matrix_ROV_Row).Subtract(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row)))
    Array_a1_Compression_eFuse = a1_Compression.Data
    Array_a2_Compression_eFuse = a2_Compression.Data
    For i = 0 To UBound(Array_a1_Compression_eFuse)
        If i = 0 Then
            If Array_a1_Compression_eFuse(i) >= 1 Then Array_a1_Compression_eFuse(i) = 2 ^ 15 - 1 Else Array_a1_Compression_eFuse(i) = FormatNumber(Array_a1_Compression_eFuse(i) * 2 ^ 15, 0)
        Else
            If Array_a1_Compression_eFuse(i) >= 1 Then Array_a1_Compression_eFuse(i) = 2 ^ 14 - 1 Else Array_a1_Compression_eFuse(i) = FormatNumber(Array_a1_Compression_eFuse(i) * 2 ^ 14, 0)
        End If
    Next i
    For i = 0 To UBound(Array_a2_Compression_eFuse)
        If i = 0 Then
            If Array_a2_Compression_eFuse(i) >= 1 Then Array_a2_Compression_eFuse(i) = 2 ^ 15 - 1 Else Array_a2_Compression_eFuse(i) = FormatNumber(Array_a2_Compression_eFuse(i) * 2 ^ 15, 0)
        Else
            If Array_a2_Compression_eFuse(i) >= 1 Then Array_a2_Compression_eFuse(i) = 2 ^ 14 - 1 Else Array_a2_Compression_eFuse(i) = FormatNumber(Array_a2_Compression_eFuse(i) * 2 ^ 14, 0)
        End If
    Next i
    a1_Compression_eFuse.Data = Array_a1_Compression_eFuse
    a2_Compression_eFuse.Data = Array_a2_Compression_eFuse
Next site

For i = 0 To MTRSNS_Matrix_ROT_Row - 1
    TestNameInput = Report_TName_From_Instance("CalcC", "", , CInt(i))
    TheExec.Flow.TestLimit resultVal:=a1.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", "", , CInt(i))
    TheExec.Flow.TestLimit resultVal:=a1_Compression.Element(i), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Tname:=TestNameInput, ForceResults:=tlForceFlow
    For Each site In TheExec.sites
        a1_Compression_eFuse_Store(i) = a1_Compression_eFuse.Select(i, , 1).Copy
    Next site
    Call AddStoredCaptureData("mtr_" & Sensor & "_t1_a1_" & CStr(i + 1), a1_Compression_eFuse_Store(i))
Next i
For i = 0 To MTRSNS_Matrix_ROV_Row - 1
    TestNameInput = Report_TName_From_Instance("CalcC", "", , CInt(i))
    TheExec.Flow.TestLimit resultVal:=a2.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance("CalcC", "", , CInt(i))
    TheExec.Flow.TestLimit resultVal:=a2_Compression.Element(i), lowCompareSign:=tlSignGreater, highCompareSign:=tlSignLess, Tname:=TestNameInput, ForceResults:=tlForceFlow
    For Each site In TheExec.sites
        a2_Compression_eFuse_Store(i) = a2_Compression_eFuse.Select(i, , 1).Copy
    Next site
    Call AddStoredCaptureData("mtr_" & Sensor & "_t1_a2_" & CStr(i + 1), a2_Compression_eFuse_Store(i))
Next i

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologySense_Compression"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function MetrologyTMPS_2s_Complement_Fractional_Conversion(ByRef Coeff_Dict As DSPWave, Coeff As DSPWave, Integer_Bit As Long, Fractional_Bit As Long) As Long
Dim site As Variant
Dim Array_Coeff_Dict() As Double
Dim High_limit As Double: High_limit = Bin2Dec_rev(String(Integer_Bit - 1, "1")) + Bin2Dec_rev_Fractional(String(Fractional_Bit, "1"))
Dim Low_limit As Double: Low_limit = -2 ^ (Integer_Bit - 1)

    For Each site In TheExec.sites
        Array_Coeff_Dict = Coeff.Data
        If Array_Coeff_Dict(0) < Low_limit Then
            Array_Coeff_Dict(0) = 2 ^ (Integer_Bit + Fractional_Bit) + FormatNumber(2 ^ Fractional_Bit * Low_limit, 0)
        ElseIf Array_Coeff_Dict(0) >= Low_limit And Array_Coeff_Dict(0) < 0 Then
            If FormatNumber(2 ^ Fractional_Bit * Array_Coeff_Dict(0), 0) = 0 Then Array_Coeff_Dict(0) = 0 Else Array_Coeff_Dict(0) = 2 ^ (Integer_Bit + Fractional_Bit) + FormatNumber(2 ^ Fractional_Bit * Array_Coeff_Dict(0), 0)
        ElseIf Array_Coeff_Dict(0) < High_limit And Array_Coeff_Dict(0) >= 0 Then
            Array_Coeff_Dict(0) = FormatNumber(2 ^ Fractional_Bit * Array_Coeff_Dict(0), 0)
        Else
            Array_Coeff_Dict(0) = FormatNumber(2 ^ Fractional_Bit * High_limit, 0)
        End If
        Coeff_Dict.Data = Array_Coeff_Dict
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in MetrologyTMPS_2s_Complement_Fractional_Conversion"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_MetrologySense_DeCompression(argc As Integer, argv() As String) As Long
        If UCase(TheExec.CurrentJob) = "CP1" Or UCase(TheExec.CurrentJob) = "FT1" Then Exit Function 'For Tonga
    Dim SweepCondition_Split() As String: SweepCondition_Split = Split(argv(0), "+")
    Dim Sensor As String: Sensor = argv(1)
    Dim MTRSNS_Matrix_Index As Long: MTRSNS_Matrix_Index = argv(2)
    Dim MTRSNS_Matrix_ROT_Row As Long: MTRSNS_Matrix_ROT_Row = argv(3)
    Dim MTRSNS_Matrix_ROT_Column As Long: MTRSNS_Matrix_ROT_Column = argv(4)
    Dim MTRSNS_Matrix_ROV_Row As Long: MTRSNS_Matrix_ROV_Row = argv(5)
    Dim MTRSNS_Matrix_ROV_Column As Long: MTRSNS_Matrix_ROV_Column = argv(6)
    Dim MetrologySense_ROT_Frequency As New DSPWave
    Dim MetrologySense_ROV_Frequency As New DSPWave
    Dim MetrologySense_ROT_Frequency_DeCompression As New DSPWave
    Dim MetrologySense_ROV_Frequency_DeCompression As New DSPWave
'    Dim MetrologySense_ROT_Frequency_DeCompression_Store() As New DSPWave: ReDim MetrologySense_ROT_Frequency_DeCompression_Store(UBound(SweepCondition_Split))
'    Dim MetrologySense_ROV_Frequency_DeCompression_Store() As New DSPWave: ReDim MetrologySense_ROV_Frequency_DeCompression_Store(UBound(SweepCondition_Split))
    Dim DSP_MetrologySense_ROT_Matrix As New DSPWave: DSP_MetrologySense_ROT_Matrix.Data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROT_Matrix
    Dim DSP_MetrologySense_ROV_Matrix As New DSPWave: DSP_MetrologySense_ROV_Matrix.Data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROV_Matrix
    Dim a1_Compression_eFuse As New DSPWave
    Dim a2_Compression_eFuse As New DSPWave
    Dim a1_Compression As New DSPWave
    Dim a2_Compression As New DSPWave
    Dim Array_a1_Compression() As Double
    Dim Array_a2_Compression() As Double
    Dim Array_a1_Compression_Str() As String: ReDim Array_a1_Compression_Str(MTRSNS_Matrix_ROT_Row - 1)
    Dim Array_a2_Compression_Str() As String: ReDim Array_a2_Compression_Str(MTRSNS_Matrix_ROV_Row - 1)
    Dim a1_DeCompression As New DSPWave
    Dim a2_DeCompression As New DSPWave
    Dim DSP_ROT_a_max_min As New DSPWave: DSP_ROT_a_max_min.Data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROT_a_max_min_Matrix
    Dim DSP_ROV_a_max_min As New DSPWave: DSP_ROV_a_max_min.Data = MetrologySense_Matrix(MTRSNS_Matrix_Index).ROV_a_max_min_Matrix
    Dim MetrologySense_ROT_Frequency_Error As New DSPWave
    Dim MetrologySense_ROV_Frequency_Error As New DSPWave
'    Dim FlowLimitObj As IFlowLimitsInfo: Call TheExec.Flow.GetTestLimits(FlowLimitObj)
'    Dim FlowTestName() As String: Call FlowLimitObj.GetTNames(FlowTestName)
'    Dim TestLimitIndex As Long: TestLimitIndex = TheExec.Flow.TestLimitIndex
    Dim i As Long
    Dim site As Variant
    
    For i = 0 To UBound(SweepCondition_Split)
        If i = 0 Then
            MetrologySense_ROT_Frequency = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT")
            MetrologySense_ROV_Frequency = GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV")
        Else
            For Each site In TheExec.sites
                MetrologySense_ROT_Frequency = MetrologySense_ROT_Frequency.Concatenate(GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT"))
                MetrologySense_ROV_Frequency = MetrologySense_ROV_Frequency.Concatenate(GetStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV"))
            Next site
        End If
    Next i
    For i = 0 To MTRSNS_Matrix_ROT_Row - 1
        If i = 0 Then
            a1_Compression_eFuse = GetStoredCaptureData("mtr_" & Sensor & "_t1_a1_" & CStr(i + 1))
        Else
            For Each site In TheExec.sites
                a1_Compression_eFuse = a1_Compression_eFuse.Concatenate(GetStoredCaptureData("mtr_" & Sensor & "_t1_a1_" & CStr(i + 1)))
            Next site
        End If
    Next i
    For i = 0 To MTRSNS_Matrix_ROV_Row - 1
        If i = 0 Then
            a2_Compression_eFuse = GetStoredCaptureData("mtr_" & Sensor & "_t1_a2_" & CStr(i + 1))
        Else
            For Each site In TheExec.sites
                a2_Compression_eFuse = a2_Compression_eFuse.Concatenate(GetStoredCaptureData("mtr_" & Sensor & "_t1_a2_" & CStr(i + 1)))
            Next site
        End If
    Next i
    For Each site In TheExec.sites
        Array_a1_Compression = a1_Compression_eFuse.Data
        Array_a2_Compression = a2_Compression_eFuse.Data
        For i = 0 To UBound(Array_a1_Compression)
            If i = 0 Then
                Call Dec2Bin_str(Array_a1_Compression(i), Array_a1_Compression_Str(i), 14)
            Else
                Call Dec2Bin_str(Array_a1_Compression(i), Array_a1_Compression_Str(i), 13)
            End If
            Array_a1_Compression(i) = Bin2Dec_rev_Fractional(Array_a1_Compression_Str(i))
        Next i
        For i = 0 To UBound(Array_a2_Compression)
            If i = 0 Then
                Call Dec2Bin_str(Array_a2_Compression(i), Array_a2_Compression_Str(i), 14)
            Else
                Call Dec2Bin_str(Array_a2_Compression(i), Array_a2_Compression_Str(i), 13)
            End If
            Array_a2_Compression(i) = Bin2Dec_rev_Fractional(Array_a2_Compression_Str(i))
        Next i
        a1_Compression.Data = Array_a1_Compression
        a2_Compression.Data = Array_a2_Compression
        a1_DeCompression = a1_Compression.Multiply(DSP_ROT_a_max_min.Select(0, 2, MTRSNS_Matrix_ROT_Row).Subtract(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row))).Add(DSP_ROT_a_max_min.Select(1, 2, MTRSNS_Matrix_ROT_Row))
        a2_DeCompression = a2_Compression.Multiply(DSP_ROV_a_max_min.Select(0, 2, MTRSNS_Matrix_ROV_Row).Subtract(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row))).Add(DSP_ROV_a_max_min.Select(1, 2, MTRSNS_Matrix_ROV_Row))
        MetrologySense_ROT_Frequency_DeCompression = DSP_MetrologySense_ROT_Matrix.MatrixTranspose(MTRSNS_Matrix_ROT_Row).MatrixMultiply(MTRSNS_Matrix_ROT_Column, MTRSNS_Matrix_ROT_Row, a1_DeCompression)
        MetrologySense_ROV_Frequency_DeCompression = DSP_MetrologySense_ROV_Matrix.MatrixTranspose(MTRSNS_Matrix_ROV_Row).MatrixMultiply(MTRSNS_Matrix_ROV_Column, MTRSNS_Matrix_ROV_Row, a2_DeCompression)
        MetrologySense_ROT_Frequency_DeCompression = MetrologySense_ROT_Frequency_DeCompression.Multiply(10 ^ 9)
        MetrologySense_ROV_Frequency_DeCompression = MetrologySense_ROV_Frequency_DeCompression.Multiply(10 ^ 9)
        MetrologySense_ROT_Frequency_Error = MetrologySense_ROT_Frequency_DeCompression.Subtract(MetrologySense_ROT_Frequency).Divide(MetrologySense_ROT_Frequency_DeCompression)
        MetrologySense_ROV_Frequency_Error = MetrologySense_ROV_Frequency_DeCompression.Subtract(MetrologySense_ROV_Frequency).Divide(MetrologySense_ROV_Frequency_DeCompression)
    Next site
'    If UCase(TheExec.CurrentJob) = "FT2" Then
'        For i = 0 To UBound(SweepCondition_Split)
'            For Each site In TheExec.sites
'                MetrologySense_ROT_Frequency_DeCompression_Store(i) = MetrologySense_ROT_Frequency_DeCompression.Select(i, , 1)
'                MetrologySense_ROV_Frequency_DeCompression_Store(i) = MetrologySense_ROV_Frequency_DeCompression.Select(i, , 1)
'            Next site
'            Call AddStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROT-DeCompression", MetrologySense_ROT_Frequency_DeCompression_Store(i))
'            Call AddStoredCaptureData(SweepCondition_Split(i) & "-Freq-" & Sensor & "-sensor-ROV-DeCompression", MetrologySense_ROV_Frequency_DeCompression_Store(i))
'        Next i
'    End If
'    Do While Not (LCase(FlowTestName(TestLimitIndex)) Like "*decompression*")
'        TestLimitIndex = TestLimitIndex + 1
'    Loop
'    TheExec.Flow.TestLimitIndex = TestLimitIndex
    For i = 0 To MTRSNS_Matrix_ROT_Column - 1
        TestNameInput = Report_TName_From_Instance("CalcF", "", , CInt(i))
        TheExec.Flow.TestLimit resultVal:=MetrologySense_ROT_Frequency_DeCompression.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcF", "", , CInt(i))
        TheExec.Flow.TestLimit resultVal:=MetrologySense_ROV_Frequency_DeCompression.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcC", "", , CInt(i))
        TheExec.Flow.TestLimit resultVal:=MetrologySense_ROT_Frequency_Error.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance("CalcC", "", , CInt(i))
        TheExec.Flow.TestLimit resultVal:=MetrologySense_ROV_Frequency_Error.Element(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologySense_DeCompression"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologyGR_Offset(argc As Integer, argv() As String) As Long
    Dim MTRGR_Offset As New SiteDouble: MTRGR_Offset = GetStoredData(argv(0) & "_para")
    Dim Integer_Bit As Long: Integer_Bit = argv(1)
    Dim Dictionary_Name As String: Dictionary_Name = argv(2)
    Dim MTRGR_Offset_Array(0) As Double
    Dim DSP_MTRGR_Offset As New DSPWave
    Dim DSP_MTRGR_Offset_eFuse As New DSPWave
    
    For Each site In TheExec.sites
        MTRGR_Offset_Array(0) = MTRGR_Offset
        DSP_MTRGR_Offset.Data = MTRGR_Offset_Array
    Next site
    Call MetrologyTMPS_2s_Complement_Fractional_Conversion(DSP_MTRGR_Offset_eFuse, DSP_MTRGR_Offset, Integer_Bit, 0)
    Call AddStoredCaptureData(Dictionary_Name, DSP_MTRGR_Offset_eFuse)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyGR_Offset"
    If AbortTest Then Exit Function Else Resume Next
End Function
Function Calc_MetrologyGR_Gain(argc As Integer, argv() As String) As Long
    Dim MTRGR_Gain_Split() As String: MTRGR_Gain_Split = Split(argv(0), "+")
    Dim MTRGR_Gain() As New SiteDouble: ReDim MTRGR_Gain(UBound(MTRGR_Gain_Split))
    Dim MTRGR_Offset_Dictionary As String: MTRGR_Offset_Dictionary = argv(1) & "_para"
    Dim DictionaryName As String: DictionaryName = argv(2)
    Dim MTRGR_Offset As New SiteDouble: MTRGR_Offset = GetStoredData(MTRGR_Offset_Dictionary)
    Dim DSP_MTRGR_Gain_eFuse As New DSPWave
    Dim MTRGR_Gain_eFuse_Array(0) As Double
    Dim i As Long
    Dim TestNameInput As String
    
    For i = 0 To UBound(MTRGR_Gain_Split)
        MTRGR_Gain(i) = GetStoredData(MTRGR_Gain_Split(i) & "_para")
    Next i
    For Each site In TheExec.sites
        MTRGR_Gain_eFuse_Array(0) = 0
        For i = 0 To UBound(MTRGR_Gain_Split)
            MTRGR_Gain_eFuse_Array(0) = MTRGR_Gain_eFuse_Array(0) + MTRGR_Gain(i)
        Next i
        MTRGR_Gain_eFuse_Array(0) = MTRGR_Gain_eFuse_Array(0) - 8 * MTRGR_Offset
        DSP_MTRGR_Gain_eFuse.Data = MTRGR_Gain_eFuse_Array
    Next site

    TestNameInput = Report_TName_From_Instance("CalcC", "")
    TheExec.Flow.TestLimit resultVal:=DSP_MTRGR_Gain_eFuse.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Call AddStoredCaptureData(DictionaryName, DSP_MTRGR_Gain_eFuse)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyGR_Gain"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MetrologyGR_RPSRSPARE0(argc As Integer, argv() As String) As Long
    Dim i As Long
    Dim MeasValue() As New PinListData: ReDim MeasValue(argc - 1)
    Dim RPSR_SPARE0 As New SiteDouble
    Dim site As Variant
    For i = 0 To argc - 1
        MeasValue(i) = GetStoredMeasurement(argv(i))
    Next i
    For Each site In TheExec.sites
        RPSR_SPARE0 = (MeasValue(0).Pins(0).Value - MeasValue(1).Pins(0).Value) / 0.000001
    Next site
    TestNameInput = Report_TName_From_Instance("CalcR", "")
    TheExec.Flow.TestLimit resultVal:=RPSR_SPARE0, Tname:=TestNameInput, ForceResults:=tlForceFlow
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyGR_RPSRSPARE0"
    If AbortTest Then Exit Function Else Resume Next
End Function


