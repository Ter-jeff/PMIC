Attribute VB_Name = "LIB_HardIP_Calc_Turks"
Option Explicit
Dim DigCapStrs As New SiteVariant

Dim DigCapStrsCompeA As New SiteVariant
Dim DigCapStrsCompeB As New SiteVariant


Dim PVTx_1to0  As New SiteVariant
Dim CompeVref As New SiteVariant
Dim CompeDigcapSwap As New SiteVariant
Public K_InstanceName As String
Public K_PatSetName As String
Public K_InstanceName_WO_Pset As String
Public InstNameSegs() As String
Public TNameSeg(10) As String
Public sweep_power_val_per_loop_count As String

Public Function Calc_PVTx(argc As Long, argv() As String) As Long

    Dim site As Variant
    Dim k As Integer
    
    Dim DigCapWave As New DSPWave
    Dim DigSrcWave As New DSPWave
    Dim DigCapKey As String
    Dim Savekeyname As String
    Dim PS As New DSPWave
    Dim SweepFrom As Integer
    Dim SweepTo As Integer
    
    DigCapKey = argv(0)
    Savekeyname = argv(1)
    
    SweepFrom = argv(2)
    SweepTo = argv(3)
    
    DigCapWave = GetStoredCaptureData(DigCapKey)

    

    k = TheExec.Flow.var("SrcCodeIndx").Value


    If k = SweepFrom Then
            'Set ParallelStream = New DSPWave
            


        For Each site In TheExec.sites
            DigCapStrs = Str(DigCapWave.Element(0))
        Next site
        PVTx_1to0 = -1

    ElseIf (SweepFrom < SweepTo And k <= SweepTo) Or (SweepFrom > SweepTo And k >= SweepTo) Then


        For Each site In TheExec.sites
            If Right(DigCapStrs, 1) = "1" And DigCapWave.Element(0) = 0 Then
                PVTx_1to0 = k - (SweepTo - SweepFrom) / Abs(SweepTo - SweepFrom)
                

            End If
            DigCapStrs = DigCapStrs & Trim(Str(DigCapWave.Element(0)))

        Next site


        If k = SweepTo Then
            Dim TestNameInput As String
            Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
            
            
            PS.CreateConstant 0, 1, DspLong
            TheExec.Datalog.WriteComment " *** PVTX-SEARCH (1->0) ***"
            If SweepTo = 63 Then
                TheExec.Datalog.WriteComment "         0         1         2         3         4         5         6"
            ElseIf SweepTo = 0 Then
                TheExec.Datalog.WriteComment "        63  6         5         4         3         2         1         0"
            End If
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrs(site)
                
                If PVTx_1to0 >= 0 Then
                PS.Element(0) = PVTx_1to0
                Else
                    PS.Element(0) = 0
                    'Stop
                End If
                
                DigSrcWave = PS.ConvertStreamTo(tldspSerial, 6, 0, Bit0IsMsb)
            Next site

            AddStoredCaptureData Savekeyname, DigSrcWave
            
            
            
            gl_FlowForLoop_DigSrc_SweepCode_temp = gl_FlowForLoop_DigSrc_SweepCode
            gl_FlowForLoop_DigSrc_SweepCode = ""
            TestNameInput = Report_TName_From_Instance(CalcC, "X")
            gl_FlowForLoop_DigSrc_SweepCode = gl_FlowForLoop_DigSrc_SweepCode_temp
            
            TheExec.Flow.TestLimit PVTx_1to0, Tname:=TestNameInput, ForceResults:=tlForceFlow

        End If
    Else

    End If

    
    
End Function


Public Function Calc_MDLL_BIST(argc As Integer, argv() As String) As Long
    'Calc_ONE_MAX_COUNT
    Dim InputKey() As String
  
    Dim site As Variant
    Dim arg As Long
    Dim i As Integer
    Dim WaveeStr As String
    
    Dim Input_Dspwave() As New DSPWave
    Dim SampleSize As Integer
    
    Dim Temp_ContinuousOne As New SiteLong
    Dim Max_ContinuousOne As New SiteLong
    Dim MDLL_Calc As New SiteDouble
    Dim EachRCapDspWave As New DSPWave
    
    '/* ------------------------------ */
    ReDim Input_Dspwave(argc)
    ReDim InputKey(argc)
    'NAND_T9BISTWRV1818_PP_TURA0_S_FULP_AN_AN01_PFF_JTG_CAL_V1818_SI_BISTWR_T9_HV
    '(0) leading
    '(1) trailing
    '(2) nis_bist_wr_bitmap
    'NAND_T10BISTRDV1818_PP_TURA0_S_FULP_AN_AN01_PFF_JTG_CAL_V1818_SI_BISTRD_T10_HV
    '(2) nis_bist_rd_pos_bitmap
    '(2) nis_bist_rd_neg_bitmap
    
    For arg = 0 To argc - 1
        InputKey(arg) = LCase(argv(arg))
        Input_Dspwave(arg) = GetStoredCaptureData(InputKey(arg))
        For Each site In TheExec.sites
        Call Wave2Str_Single(Input_Dspwave(arg), WaveeStr)
        TheExec.Datalog.WriteComment "Site(" & site & "):" & InputKey(arg) & " = " & WaveeStr
        Next site
    Next arg
    

    For arg = 1 To argc - 1 Step 3
        Call rundsp.ConvertToLongAndSerialToParrel(Input_Dspwave(arg - 1), 9, EachRCapDspWave)
            For Each site In TheExec.sites
                Temp_ContinuousOne = 0
                Max_ContinuousOne = 0
                For i = 0 To Input_Dspwave(arg + 1).SampleSize - 1
                    If Input_Dspwave(arg + 1).Element(i) = 1 Then
                        Temp_ContinuousOne = Temp_ContinuousOne + 1
                    Else
                        If Temp_ContinuousOne > Max_ContinuousOne Then
                            Max_ContinuousOne = Temp_ContinuousOne
                        End If
                        Temp_ContinuousOne = 0
                    End If
                Next i
                If Temp_ContinuousOne > Max_ContinuousOne Then
                        Max_ContinuousOne = Temp_ContinuousOne
                        Temp_ContinuousOne = 0
                End If
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Decimal of Leading" & " = " & EachRCapDspWave.Element(0)
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Max Continuous One of Bitmap" & " = " & Max_ContinuousOne
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Formula" & " = " & Max_ContinuousOne & "/" & EachRCapDspWave.Element(0) & "*" & 360
             If EachRCapDspWave.Element(0) = 0 Then
                TheExec.Datalog.WriteComment "Site(" & site & "):" & "Decimal of Leading" & " = 0, Set Decimal of Leading=99999999"
                EachRCapDspWave.Element(0) = 99999999
             End If
             MDLL_Calc = Round((Max_ContinuousOne * 360) / EachRCapDspWave.Element(0), 5)
         Next site
        'Report_TestLimit_by_CZ_Format resultVal:=Max_ContinuousOne, MeasType:="C", UserVar5:="MaxOneCount", UserVar7:=InputKey(arg + 1), scaletype:=scaleNoScaling, ForceResults:=tlForceNone
        'Report_TestLimit_by_CZ_Format resultVal:=MDLL_Calc, MeasType:="C", UserVar7:=InputKey(arg + 1), scaletype:=scaleNoScaling, ForceResults:=tlForceFlow
   
         
        Dim TestNameInput As String
        Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
        gl_FlowForLoop_DigSrc_SweepCode_temp = gl_FlowForLoop_DigSrc_SweepCode
        
        gl_FlowForLoop_DigSrc_SweepCode = Replace(InputKey(arg + 1), "_", "")
        TestNameInput = Report_TName_From_Instance(CalcC, PinName:="MaxOneCount")
        
            
        TheExec.Flow.TestLimit resultVal:=Max_ContinuousOne, Tname:=TestNameInput, ForceResults:=tlForceNone
        TestNameInput = Report_TName_From_Instance(CalcC, "Degrees")
        TheExec.Flow.TestLimit resultVal:=MDLL_Calc, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        gl_FlowForLoop_DigSrc_SweepCode = gl_FlowForLoop_DigSrc_SweepCode_temp
   
   Next arg
   
End Function

Public Function Calc_NAND_PHY_MDLL(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim Result_Ratio As New SiteDouble
    Dim DSPWave_Binary() As New DSPWave
    ReDim DSPWave_Binary(argc - 1) As New DSPWave
    
    Dim DSPWave_Combine_Dec() As New DSPWave
    ReDim DSPWave_Combine_Dec(argc - 1) As New DSPWave
    
    Dim DSPWave_Result As New DSPWave
    DSPWave_Result.CreateConstant 0, 3, DspLong

    Dim DSPWave_Result_K As New DSPWave
    DSPWave_Result_K.CreateConstant 0, 1, DspLong
        
        
    If TheExec.TesterMode = testModeOnline Then
    
        If TheExec.Flow.EnableWord("TTR") = True Then
            For i = 0 To 3
                DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
            Next i
            Call rundsp.Calc_NAND_PHY_MDLL_DSP(DSPWave_Binary(0), DSPWave_Binary(1), DSPWave_Binary(2), DSPWave_Binary(3), DSPWave_Result, Result_Ratio)
        Else
        
            For i = 0 To 3
                DSPWave_Combine_Dec(i).CreateConstant 0, 1, DspLong
                DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
                Call rundsp.ConvertToLongAndSerialToParrel(DSPWave_Binary(i), 9, DSPWave_Combine_Dec(i))
            Next i
            
             For Each site In TheExec.sites
                DSPWave_Result.Element(0) = DSPWave_Combine_Dec(0).Element(0) - DSPWave_Combine_Dec(1).Element(0)
                DSPWave_Result.Element(1) = DSPWave_Combine_Dec(2).Element(0) - DSPWave_Combine_Dec(3).Element(0)
        '        DSPWave_Result.Element(2) = DSPWave_Result.Element(1) - DSPWave_Result.Element(0)
                If DSPWave_Result.Element(0) = 0 Then DSPWave_Result.Element(0) = 99999999
                Result_Ratio = DSPWave_Result.Element(1) / DSPWave_Result.Element(0)
            Next site
        End If
    
    'If TheExec.TesterMode = testModeOffline Then
    Else
    'testModeOffline
        For i = 0 To 3
            DSPWave_Combine_Dec(i).CreateConstant 0, 1, DspLong
            DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
            Call rundsp.ConvertToLongAndSerialToParrel(DSPWave_Binary(i), 9, DSPWave_Combine_Dec(i))
        Next i
        
         For Each site In TheExec.sites
            DSPWave_Result.Element(0) = DSPWave_Combine_Dec(0).Element(0) - DSPWave_Combine_Dec(1).Element(0)
            DSPWave_Result.Element(1) = DSPWave_Combine_Dec(2).Element(0) - DSPWave_Combine_Dec(3).Element(0)
    '        DSPWave_Result.Element(2) = DSPWave_Result.Element(1) - DSPWave_Result.Element(0)
            If DSPWave_Result.Element(0) = 0 Then DSPWave_Result.Element(0) = 99999999
            Result_Ratio = DSPWave_Result.Element(1) / DSPWave_Result.Element(0)
        Next site
    
        For Each site In TheExec.sites
            DSPWave_Result.Element(1) = DSPWave_Combine_Dec(2).Element(0) - DSPWave_Combine_Dec(3).Element(0) + 400 + (site * 10)
        Next site
    End If
    
    
    If TheExec.Flow.EnableWord("CZ2_PRINT_EN") = False Then
        TheExec.Flow.TestLimit DSPWave_Result.Element(0), , , , ForceResults:=tlForceFlow  'chyehq
        TheExec.Flow.TestLimit DSPWave_Result.Element(1), , , , ForceResults:=tlForceFlow  'chyehq
        TheExec.Flow.TestLimit Result_Ratio, , , , ForceResults:=tlForceFlow  'chyehq
    Else
         
        Dim TestNameInput As String
        Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
        gl_FlowForLoop_DigSrc_SweepCode_temp = gl_FlowForLoop_DigSrc_SweepCode
        
        
        
        
    
        'Report_TestLimit_by_CZ_Format DSPWave_Result.Element(0), , , , ForceResults:=tlForceFlow, MeasType:="C"
        'Report_TestLimit_by_CZ_Format DSPWave_Result.Element(1), , , , ForceResults:=tlForceFlow, MeasType:="C"
        'Report_TestLimit_by_CZ_Format Result_Ratio, , , ForceResults:=tlForceFlow, MeasType:="C", UserVar5:="MDLL", UserVar6:="CAL", UserVar7:="RATIO"
        
        TestNameInput = Report_TName_From_Instance(CalcC, "", ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit DSPWave_Result.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        TestNameInput = Report_TName_From_Instance(CalcC, "", ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit DSPWave_Result.Element(1), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        TestNameInput = Report_TName_From_Instance(CalcC, PinName:="MDLL", ForceResult:=tlForceFlow)
        'Stop
        
        TheExec.Flow.TestLimit Result_Ratio, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    End If
    
    
'     For Each Site In TheExec.sites
'         If DSPWave_Result.Element(1) < 1 Then DSPWave_Result.Element(1) = 1
'
'         If TheExec.Flow.EnableWord("CZ2_PRINT_EN") = False Then
'            TheExec.Flow.TestLimit DSPWave_Result.Element(2), 1, DSPWave_Result.Element(1), Tname:="MDLLCALDIFF", ForceResults:=tlForceNone  'chyehq
'        Else
'            Report_TestLimit_by_CZ_Format DSPWave_Result.Element(2), 1, DSPWave_Result.Element(1), ForceResults:=tlForceNone, MeasType:="C", UserVar5:="MDLL", UserVar6:="CAL", UserVar7:="DIFF"
'        End If
'    Next Site
    
End Function

Public Function Wave2Str_Single(InDSPwave As DSPWave, outstr As Variant, Optional sp As Integer = 0, Optional EP As Integer = 0) As Long
    Dim i As Integer
    outstr = ""
    If sp + EP > 0 Then
        If EP > InDSPwave.SampleSize - 1 Then EP = InDSPwave.SampleSize - 1
        For i = sp To EP
            outstr = outstr & CStr(InDSPwave.Element(i))
        Next i
    Else
        For i = 0 To InDSPwave.SampleSize - 1
            outstr = outstr & CStr(InDSPwave.Element(i))
        Next i
    End If
End Function
Public Function Calc_PVTP(argc As Long, argv() As String) As Long
    
    Dim site As Variant
    Dim PVTPNR_int As New SiteLong
    Dim i, j, k As Integer
    Dim WaveeStr As String
    Dim dataArray(31) As Long
    Dim SimStr(7) As String
    Dim SimArray() As Long
        
    Dim InDSPwave As New DSPWave
    Dim PVTP_Wave As New DSPWave
    Dim PVTP_PLUS_1_Wave As New DSPWave
    Dim PVTP_PLUS_2_Wave As New DSPWave
    Dim PVTP_PLUS_3_Wave As New DSPWave
    Dim PVTP_PLUS_4_Wave As New DSPWave
    Dim PVTP_MINUS_1_Wave As New DSPWave
    
    Dim ParallelStream As New DSPWave ''''20190604Error
    Dim PVT_1to0 As Integer ''''20190604Error
    
    Dim PVTx As New SiteDouble
    Dim PVTx_Stored As New SiteDouble
    
    
    Dim AnySite_1to0_Check As New SiteBoolean
    'Dim All_1_to_0 As new SiteBoolean
    
    
    
    InDSPwave = GetStoredCaptureData(argv(0))
    
    Dim KeyName As String
    
    KeyName = argv(1)
    
    k = TheExec.Flow.var("SrcCodeIndx").Value
    
    PVTPNR_int = InDSPwave.Element(0)
        
    If k = 0 Then
        ParallelStream.CreateConstant 0, 1, DspLong
        
        For Each site In TheExec.sites
            DigCapStrs = Str(InDSPwave.Element(0))
        Next site
        PVT_1to0 = 999
        
    ElseIf k <= 32 Then
    
   
        For Each site In TheExec.sites
            If Right(DigCapStrs, 1) = "1" And InDSPwave.Element(0) = 0 Then
                PVT_1to0 = k
                ParallelStream.Element(0) = k
                
            End If
            DigCapStrs = DigCapStrs & Trim(Str(InDSPwave.Element(0)))
            
        Next site
        
        
        If k = 32 Then
            TheExec.Flow.TestLimit PVT_1to0, ForceResults:=tlForceNone
            
            TheExec.Datalog.WriteComment "         0         1         2         3"
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrs(site)
                PVTP_Wave = ParallelStream.ConvertStreamTo(tldspSerial, 6, 0, Bit0IsMsb)
            Next site
            
            Stop
            
            AddStoredCaptureData KeyName, PVTP_Wave
            
        End If
    Else
    
            
    End If
    Exit Function
    
    PVTx = PVTPNR_int.compare(EqualTo, 0)
    PVTx = PVTx.Abs.Multiply(k)
    
    
    If k = 32 Then
        
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrs(site)
        Next site
        Stop

        
    Else
    
    
        'Stop
    End If
    
    
    
    
    

    
    
    

End Function
Public Function Calc_PVTN(argc As Long, argv() As String) As Long
    
    
    Dim site As Variant
    Dim PVTPNR_int As New SiteLong
    Dim i, j, k As Integer
    Dim WaveeStr As String
    Dim dataArray(31) As Long
    Dim SimStr(7) As String
    Dim SimArray() As Long
     
    Dim InDSPwave As New DSPWave
    Dim PVTN_Wave As New DSPWave
    Dim PVTN_PLUS_1_Wave As New DSPWave
    Dim PVTN_PLUS_2_Wave As New DSPWave
    Dim PVTN_PLUS_3_Wave As New DSPWave
    Dim PVTN_PLUS_4_Wave As New DSPWave
    Dim PVTN_MINUS_1_Wave As New DSPWave

    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For i = 0 To UBound(dataArray)
                dataArray(i) = 0
                If i > 15 + site Then dataArray(i) = 1
            Next i
            InDSPwave.Data = dataArray
        Next site
    Else
        InDSPwave = GetStoredCaptureData(argv(0))
    End If
    
    Call rundsp.NAND_PVTN(InDSPwave, PVTPNR_int, PVTN_Wave, PVTN_PLUS_1_Wave, PVTN_PLUS_2_Wave, PVTN_PLUS_3_Wave, PVTN_PLUS_4_Wave, PVTN_MINUS_1_Wave)
    
    AddStoredCaptureData argv(1), PVTN_Wave
    AddStoredCaptureData "PVTN_PLUS_1", PVTN_PLUS_1_Wave
    AddStoredCaptureData "PVTN_PLUS_2", PVTN_PLUS_2_Wave
    AddStoredCaptureData "PVTN_PLUS_3", PVTN_PLUS_3_Wave
    AddStoredCaptureData "PVTN_PLUS_4", PVTN_PLUS_4_Wave
    AddStoredCaptureData "PVTN_MINUS_1", PVTN_MINUS_1_Wave
    
    'If NDLog = False Then
        For Each site In TheExec.sites
            Call Wave2Str_Single(PVTN_Wave, WaveeStr)
            TheExec.Datalog.WriteComment "Site: " & site & " The transition is point " & PVTPNR_int
            TheExec.Datalog.WriteComment "Site: " & site & " The PVTN Binary Code =  " & WaveeStr
    
            Call Wave2Str_Single(InDSPwave, WaveeStr)
            TheExec.Datalog.WriteComment "Site: " & site & ", Capture bits " & InDSPwave.SampleSize & " = " & WaveeStr
        Next site
    'End If
    
    ''2017/08/11 , updated by Kaino for CZ2 naming
    'TheExec.Flow.TestLimit PVTPNR_int, , , , , , , , , , , , , , , ForceResults:=tlForceFlow
    If TheExec.Flow.EnableWord("CZ2_PRINT_EN") = False Then
        TheExec.Flow.TestLimit PVTPNR_int, , , , , , , , , , , , , , , ForceResults:=tlForceFlow
    Else
        Report_TestLimit_by_CZ_Format resultVal:=PVTPNR_int, ForceResults:=tlForceFlow, MeasType:="C", PinName:="JTAG_TDO"
    End If
        
    'Exit Function

End Function


Public Function Calc_ONE_MAX_COUNT(argc As Integer, argv() As String) As Long
    'Calc_ONE_MAX_COUNT
    Dim InputKey() As String
  
    Dim site As Variant
    Dim arg As Long
    Dim i As Integer
    Dim WaveeStr As String
    
    Dim Input_Dspwave() As New DSPWave
    Dim SampleSize As Integer
    
    Dim Temp_ContinuousOne As New SiteLong
    Dim Max_ContinuousOne As New SiteLong
    Dim MDLL_Calc As New SiteDouble
    Dim EachRCapDspWave As New DSPWave
    
    '/* ------------------------------ */
    ReDim Input_Dspwave(argc)
    ReDim InputKey(argc)
    'NAND_T9BISTWRV1818_PP_TURA0_S_FULP_AN_AN01_PFF_JTG_CAL_V1818_SI_BISTWR_T9_HV
    '(0) leading
    '(1) trailing
    '(2) nis_bist_wr_bitmap
    'NAND_T10BISTRDV1818_PP_TURA0_S_FULP_AN_AN01_PFF_JTG_CAL_V1818_SI_BISTRD_T10_HV
    '(2) nis_bist_rd_pos_bitmap
    '(2) nis_bist_rd_neg_bitmap
    
    For arg = 0 To argc - 1
        InputKey(arg) = LCase(argv(arg))
        Input_Dspwave(arg) = GetStoredCaptureData(InputKey(arg))
        For Each site In TheExec.sites
            Call Wave2Str_Single(Input_Dspwave(arg), WaveeStr)
            TheExec.Datalog.WriteComment "Site(" & site & "):" & InputKey(arg) & " = " & WaveeStr
        Next site
    Next arg
    
    Dim leading As Integer
    Dim nis_bist_x_bitmap As Integer
    'Stop
'    If arg = 2 Then
'        nis_bist_x_bitmap = 0
'        leading = 1
'
'    ElseIf arg = 3 Then
'        nis_bist_x_bitmap = 0
'        leading = 2
'    Else
'        Stop
'    End If
    
    leading = argc - 1
    

    For nis_bist_x_bitmap = 0 To argc - 2
        'Stop
        Call rundsp.ConvertToLongAndSerialToParrel(Input_Dspwave(leading), 9, EachRCapDspWave)  'leading
            For Each site In TheExec.sites
                Temp_ContinuousOne = 0
                Max_ContinuousOne = 0
                For i = 0 To Input_Dspwave(nis_bist_x_bitmap).SampleSize - 1
                    If Input_Dspwave(nis_bist_x_bitmap).Element(i) = 1 Then
                        Temp_ContinuousOne = Temp_ContinuousOne + 1
                    Else
                        If Temp_ContinuousOne > Max_ContinuousOne Then
                            Max_ContinuousOne = Temp_ContinuousOne
                        End If
                        Temp_ContinuousOne = 0
                    End If
                Next i
                If Temp_ContinuousOne > Max_ContinuousOne Then
                        Max_ContinuousOne = Temp_ContinuousOne
                        Temp_ContinuousOne = 0
                End If
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Decimal of Leading" & " = " & EachRCapDspWave.Element(0)
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Max Continuous One of Bitmap (" & InputKey(nis_bist_x_bitmap) & ") = " & Max_ContinuousOne
             TheExec.Datalog.WriteComment "Site(" & site & "):" & "Formula" & " = " & Max_ContinuousOne & "/" & EachRCapDspWave.Element(0) & "*" & 360
             If EachRCapDspWave.Element(0) = 0 Then
                TheExec.Datalog.WriteComment "Site(" & site & "):" & "Decimal of Leading" & " = 0, Set Decimal of Leading=99999999"
                EachRCapDspWave.Element(0) = 99999999
             End If
             MDLL_Calc = Round((Max_ContinuousOne * 360) / EachRCapDspWave.Element(0), 5)
         Next site
        'Report_TestLimit_by_CZ_Format resultVal:=Max_ContinuousOne, MeasType:="C", UserVar5:="MaxOneCount", UserVar7:=InputKey(arg + 1), scaletype:=scaleNoScaling, ForceResults:=tlForceNone
        'Report_TestLimit_by_CZ_Format resultVal:=MDLL_Calc, MeasType:="C", UserVar7:=InputKey(arg + 1), scaletype:=scaleNoScaling, ForceResults:=tlForceFlow
   
         
        Dim TestNameInput As String
        Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
        gl_FlowForLoop_DigSrc_SweepCode_temp = gl_FlowForLoop_DigSrc_SweepCode
        
        gl_FlowForLoop_DigSrc_SweepCode = Replace(InputKey(nis_bist_x_bitmap), "_", "")
        TestNameInput = Report_TName_From_Instance(CalcC, "", Tname:="MaxOneCount", ForceResult:=tlForceNone)
        
            
        TheExec.Flow.TestLimit resultVal:=Max_ContinuousOne, Tname:=TestNameInput, ForceResults:=tlForceNone
        TestNameInput = Report_TName_From_Instance(CalcC, "", ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=MDLL_Calc, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        gl_FlowForLoop_DigSrc_SweepCode = gl_FlowForLoop_DigSrc_SweepCode_temp
   
   Next nis_bist_x_bitmap
   
End Function



Public Function Calc_Abs_Res(argc As Long, argv() As String)

    Dim Calc_Operand1 As Double
    Dim Calc_Operand2 As New PinListData
    Dim Calc_Res As New PinListData
    Dim TestNameInput As String
    Dim p As Variant
    Dim Temp_index As Long
        
    If InStr(UCase(argv(0)), UCase("VDD")) <> 0 Then
        'Calc_Operand1 = TheExec.Specs.DC.Item(Mid(argv(0), 2)).ContextValue
        Call HIP_Evaluate_ForceVal_New(argv(0)) 'add for +-*/ calc. by CW 190816
        Calc_Operand1 = argv(0)
        'Calc_Operand1 = TheExec.specs.DC.Item(argv(0) & "_VAR_H").ContextValue
    Else
        Calc_Operand1 = CDbl(argv(0))
    End If
    Calc_Operand2 = GetStoredMeasurement(argv(1))
    
    Calc_Res = Calc_Operand2.Copy
    Calc_Res = Calc_Operand2.Math.Abs.Invert.Multiply(Calc_Operand1)
    
    

    Temp_index = TheExec.Flow.TestLimitIndex
    
    For p = 0 To Calc_Res.Pins.Count - 1
        
        
        TestNameInput = Report_TName_From_Instance(CalcC, Calc_Res.Pins(p), , 0)
        'TestNameInput = Report_TName_From_Instance("Calc", Calc_Res.Pins(p), , 0)
        TheExec.Flow.TestLimit Calc_Res.Pins(p), , , , , , unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        
        If argv(1) <> "ip0" Then TheExec.Flow.TestLimitIndex = Temp_index 'modify for AMPH T58
        
    Next p
    
    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1

End Function



Public Function Calc_ROSC_Freq(argc As Long, argv() As String)
    
'''''''''''''''''''''input Variable info'''''''''''''''''
'Calc:Calc_ROSC_Freq;
'CalcArg:24000000,ringclk_count_val,refclk_count_val;
'
'24MHz             value to calc    > argv(0)
'ringclk_count_val dictionary name  > argv(1)
'refclk_count_val  dictionary name  > argv(2)
'
'
''''''''''''''''''''' Algorithm'''''''''''''''''
'Calculate 24Mhz*(ringclk_count_val/refclk_count_val)
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim Input_Freq As Long
    Dim Input_ringclk_count_val As New DSPWave
    Dim Input_refclk_count_val As New DSPWave
    
    Dim Input_ringclk_count_val_Dec As New DSPWave
    Dim Input_refclk_count_val_Dec As New DSPWave
    
    Dim Output_Calc_Freq As New DSPWave
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim site As Variant

    Input_Freq = CLng(argv(0))
    Input_ringclk_count_val = GetStoredCaptureData(argv(1))
    Input_refclk_count_val = GetStoredCaptureData(argv(2))
    
    Input_ringclk_count_val_Dec.CreateConstant 0, 1, DspDouble
    Input_refclk_count_val_Dec.CreateConstant 0, 1, DspDouble
    Output_Calc_Freq.CreateConstant 0, 1, DspDouble
    
    'Convert DSPwave from Binary value to Decimal value
    Call HardIP_Bin2Dec(Input_ringclk_count_val_Dec, Input_ringclk_count_val)
    Call HardIP_Bin2Dec(Input_refclk_count_val_Dec, Input_refclk_count_val)
        
    For Each site In TheExec.sites
'        If TheExec.TesterMode = testModeOffline Then
'            If Site Mod 2 = 0 Then
'                Output_Calc_Freq.Element(0) = 1000000 * CLng(TheExec.Flow.var("SrcCodeIndx").Value)
'            Else
'                Output_Calc_Freq.Element(0) = 2000000 * CLng(TheExec.Flow.var("SrcCodeIndx").Value)
'            End If
'        Else
            If Input_refclk_count_val_Dec.Element(0) <= 0 Then
                Output_Calc_Freq.Element(0) = -1
            Else
                Output_Calc_Freq.Element(0) = Input_Freq * (Input_ringclk_count_val_Dec.Element(0) / Input_refclk_count_val_Dec.Element(0))
            End If
'        End If
    Next site


    TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0)

     
    TheExec.Flow.TestLimit resultVal:=Output_Calc_Freq.Element(0), Tname:=TestNameInput, Unit:=unitHz, ForceResults:=tlForceFlow
        

End Function












Public Function Calc_ANP_Impedance_UP(argc As Integer, argv() As String) As Long
    'Calc:Calc_ANP_Impedance_UP_Sweep;
    'CalcArg:Curr0,VDDIO_ANP0,ANI0_ZQ;
    
    '****description****
    'argv(0) => to get measured value
    'argv(1) => to get "VDD power"
    'argv(2) => to get "current sweep power"
        
        
    '****algorithm****
    'Calc "Impedance", (VDDIO_ANP0 - ForcedV) / ANI0_ZQ(Curr0)
    
        
    
    Dim meas_value As New PinListData
    Dim VDD_Power_Current As Double
    Dim VDD_Power As Double
    Dim Calc_value As New PinListData
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    meas_value = GetStoredMeasurement(argv(0))
    
    VDD_Power = TheExec.specs.DC.Item(argv(1) & "_VAR_H").ContextValue
    
    If UCase(TheExec.DataManager.ChannelType(argv(2))) Like "*DCVS*" Then
        VDD_Power_Current = TheHdw.DCVS.Pins(argv(2)).Voltage.Main.Value
    ElseIf UCase(TheExec.DataManager.ChannelType(argv(2))) Like "*DCVI*" Then
        VDD_Power_Current = Val(sweep_power_val_per_loop_count) '/ 10000
        'VDD_Power_Current = thehdw.DCVI.Pins(argv(2)).Voltage
        'sweep_power_val_per_loop_count = "" '20190814
    End If
    
    
    Calc_value = meas_value.Math.Invert.Multiply(VDD_Power - VDD_Power_Current).Abs
    
    
    
    If gl_UseStandardTestName_Flag = True Then
        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
'''            If gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex) <> "" Then
'''                OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
'''            End If
'''        TestNameInput = Merge_TName(OutputTname_format)
        TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , sweep_power_val_per_loop_count)
        TheExec.Flow.TestLimit resultVal:=Calc_value, Tname:=TestNameInput, ForceResults:=tlForceFlow ', forceVal:=VDD_Power_Current
    Else
        TheExec.Flow.TestLimit resultVal:=Calc_value, ForceResults:=tlForceFlow
    End If
    

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ANP_Impedance_UP"
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function Calc_ANP_Impedance_DOWN(argc As Integer, argv() As String) As Long
    'Calc:Calc_ANP_Impedance_DOWN_Sweep;
    'CalcArg:Curr0,VDDIO_ANP0,ANI0_ZQ;
        
    '****description****
    'argv(0) => to get measured value
    'argv(1) => to get "VDD power"
    'argv(2) => to get "current sweep power"
        
    '****algorithm****
    'Calc "Impedance", ForcedV / ANI0_ZQ(Curr0)
    
        
    
    Dim meas_value As New PinListData
    Dim VDD_Power_Current As Double
    Dim VDD_Power As Double
    Dim Calc_value As New PinListData
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    meas_value = GetStoredMeasurement(argv(0))
    
    'VDD_Power = TheExec.specs.DC.Item(argv(1) & "_VAR_H").ContextValue
    
    If UCase(TheExec.DataManager.ChannelType(argv(2))) Like "*DCVS*" Then
        VDD_Power_Current = TheHdw.DCVS.Pins(argv(2)).Voltage.Main.Value
    ElseIf UCase(TheExec.DataManager.ChannelType(argv(2))) Like "*DCVI*" Then
        VDD_Power_Current = Val(sweep_power_val_per_loop_count) '/ 10000
        'VDD_Power_Current = thehdw.DCVI.Pins(argv(2)).Voltage
        'sweep_power_val_per_loop_count = "" '20190814
    End If

    
    Calc_value = meas_value.Math.Invert.Multiply(VDD_Power_Current).Abs
    
    
    
    If gl_UseStandardTestName_Flag = True Then
        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
'''            If gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex) <> "" Then
'''                OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
'''            End If
'''        TestNameInput = Merge_TName(OutputTname_format)
        TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , sweep_power_val_per_loop_count)
        TheExec.Flow.TestLimit resultVal:=Calc_value, Tname:=TestNameInput, ForceResults:=tlForceFlow ', forceVal:=VDD_Power_Current
    Else
        TheExec.Flow.TestLimit resultVal:=Calc_value, ForceResults:=tlForceFlow
    End If
    

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ANP_Impedance_DOWN"
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function Calc_ANP_VrefREAL_and_DeltaFromIdeal(argc As Long, argv() As String)
    ''''''''''''''Input format'''''''''''''''
    'Calc:Calc_ANP_VrefREAL_and_DeltaFromIdeal;
    'CalcArg:VREF0,VDDIO_ANP0,256;
    
    ''''''''''''''Algorithm''''''''''''''''''
    'MeasV  Pin = ANI0_VREF (VREF0)  "STEP3VREF"
    'Calculate delta_from_ideal = Vref - VDDIO_ANP * vref_wr_val_sel / 256
    
    
    
    
    '***for DNL&INL calculations please use the following formulas:
    'DNL:
    '[{result(n+1) - result(n)}*256*k / VDDIO_ANP] -1
    'VREFWR: k = 1
    'VREFRD (&IO) @ MS418: k=4.5
    'VREFRD (&IO) @ MS512: k=3
    '
    'limits for DNL function [-1,1]
    '
    '
    'INL:
    '''(wrong)''''''''[{result(n+1) - result(0)}*256*k*(n+1) / VDDIO_ANP] -1
    'INL=[{result(n) - result(0)}*256*k / VDDIO_ANP] - n
    '
    'if (T71 tests)  then
    'INL=[{result(n) - result(0)}*256*k / VDDIO_ANP] - n + 50
    'end if


    Dim Volt_Vref As New PinListData
    Dim Volt_VrefvsIdeal As New PinListData
    Dim Volt_VDD As Double
    Dim Volt_VDD_calc As Double
    
    '20190610 for DNL calc
    Dim Volt_Vref__N As New PinListData
    Dim Volt_Vref__Nplus1 As New PinListData
    Dim Volt_Vref__CalcDNL As New PinListData
    '20190611 for INL calc
    Dim Volt_Vref__CalcINL As New PinListData
    Dim Volt_Vref__ORG As New PinListData

    Volt_Vref = GetStoredMeasurement(argv(0))

    Volt_VDD = TheExec.specs.DC.Item(argv(1) & "_VAR_H").ContextValue

    'Volt_VDD_calc = Volt_VDD * Val(TheExec.Flow.var("SrcCodeIndx").Value) / Val(argv(2))
    Volt_VDD_calc = Volt_VDD * Val(gl_FlowForLoop_DigSrc_SweepCode_Dec) / Val(argv(2)) ' 180306 for S5E

    Volt_VrefvsIdeal = Volt_Vref.Math.Subtract(Volt_VDD_calc)
    Dim k As Single
    Dim nn As Integer
    
    k = Val(argv(2)) / 256
    nn = Val(gl_FlowForLoop_DigSrc_SweepCode_Dec)
    

    TheExec.Datalog.WriteComment "*** VREF-VS-IDEAL Calculate : delta_from_ideal=Vref-VDDIO_NAND*vrefint_sel / (K *256) ***   [  vrefint_sel = " & nn & "  ;  " & "k  = " & k & "  ;  " & argv(1) & " = " & Volt_VDD & "(V)  ]" & "   *** " & TheExec.DataManager.instanceName
    
    '20190610 for DNL calc
    Call AddStoredMeasurement(argv(0) & "_" & CStr(gl_FlowForLoop_DigSrc_SweepCode_Dec), Volt_Vref)
    
    '20190611 for INL calc
    If Val(TheExec.Flow.var("SrcCodeIndx").Value) = 0 Then
        Call AddStoredMeasurement(argv(0) & "_ORG", Volt_Vref)
    End If
  
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    
'''    If gl_UseStandardTestName_Flag = True Then
'''        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
''''''            If gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex) <> "" Then
''''''                OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
''''''            End If
'''        TestNameInput = Merge_TName(OutputTname_format)
'''        TestNameInput = Report_TName_From_Instance(CalcStr, "", ForceResult:=tlForceFlow)
'''        TheExec.Flow.TestLimit resultVal:=Volt_VrefvsIdeal, Tname:=TestNameInput, ForceResults:=tlForceFlow
'''    Else
'''        TheExec.Flow.TestLimit resultVal:=Volt_VrefvsIdeal, ForceResults:=tlForceFlow
'''    End If


'''    '20190610 for DNL calc
'''    If Val(TheExec.Flow.var("SrcCodeIndx").Value) <> 0 Then
'''
'''
'''        Volt_Vref__N = GetStoredMeasurement(argv(0) & "_" & CStr(Val(gl_FlowForLoop_DigSrc_SweepCode_Dec) - 1))
'''        Volt_Vref__CalcDNL = Volt_Vref.Math.Subtract(Volt_Vref__N).Multiply(Val(argv(2))).Divide(Volt_VDD).Subtract(1)
'''
'''        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
'''
'''        OutputTname_format(6) = UCase(argv(0)) & "-Calc-MeasV-DNL"
'''
'''        TestNameInput = Merge_TName(OutputTname_format)
'''
'''        TestNameInput = Report_TName_From_Instance(CalcStr, "", ForceResult:=tlForceFlow)
'''
'''        TheExec.Flow.TestLimit resultVal:=Volt_Vref__CalcDNL, Tname:=TestNameInput, ForceResults:=tlForceFlow, lowVal:=-1, hiVal:=1
'''    Else
'''        'added by Kaino on 2019/06/18 for Turks
'''        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
'''
'''    End If
'''
'''    '20190611 for INL calc
'''    If Val(TheExec.Flow.var("SrcCodeIndx").Value) <> 0 Then
'''
'''        Volt_Vref__ORG = GetStoredMeasurement(argv(0) & "_ORG")
'''        If UCase(TheExec.DataManager.InstanceName) Like UCase("*T17*") Then
'''            Volt_Vref__CalcINL = Volt_Vref.Math.Subtract(Volt_Vref__ORG).Multiply(Val(argv(2))).Divide(Volt_VDD).Subtract(Val(gl_FlowForLoop_DigSrc_SweepCode_Dec)).Add(50)
'''        Else
'''            Volt_Vref__CalcINL = Volt_Vref.Math.Subtract(Volt_Vref__ORG).Multiply(Val(argv(2))).Divide(Volt_VDD).Subtract(Val(gl_FlowForLoop_DigSrc_SweepCode_Dec))
'''        End If
'''        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
'''
'''        OutputTname_format(6) = UCase(argv(0)) & "-Calc-MeasV-INL"
'''
'''        TestNameInput = Merge_TName(OutputTname_format)
'''
'''
'''        TestNameInput = Report_TName_From_Instance(CalcStr, "", ForceResult:=tlForceFlow)
'''
'''
'''        TheExec.Flow.TestLimit resultVal:=Volt_Vref__CalcINL, Tname:=TestNameInput, ForceResults:=tlForceFlow
'''     Else
'''        'added by Kaino on 2019/06/18 for Turks
'''        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
'''
'''    End If


    '---------------------------------------
    'Added by Kaino for INL formula update
    '---------------------------------------
    Dim n  As Integer
    Dim Key As String
    n = Val(TheExec.Flow.var("SrcCodeIndx").Value)
    Key = LCase("vref_vs_ideal_" & Trim(CStr(n)))
    Call AddStoredMeasurement(Key, Volt_VrefvsIdeal)
    '---------------------------------------
    
    If gl_UseStandardTestName_Flag = True Then
        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
'''            If gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex) <> "" Then
'''                OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
'''            End If
        TestNameInput = Merge_TName(OutputTname_format)
        TestNameInput = Report_TName_From_Instance(CalcC, "", ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=Volt_VrefvsIdeal, Tname:=TestNameInput, ForceResults:=tlForceFlow
    Else
        TheExec.Flow.TestLimit resultVal:=Volt_VrefvsIdeal, ForceResults:=tlForceFlow
    End If

    
End Function


Public Function Calc_ANP_DeltaFromNoLoad(argc As Long, argv() As String)
    ''''''''''''''Input format'''''''''''''''
    'Calc:Calc_ANP_DeltaFromNoLoad;
    'CalcArg:VREF0SRC,VREF0;
    
    ''''''''''''''Algorithm''''''''''''''''''
    'Calculate delta_from_no_load=VrefSRC-Vref


    Dim Volt_Vref_NoLoad As New PinListData
    Dim Volt_Vref_ILoad As New PinListData
    Dim Volt_VrefDelta As New PinListData
    
    Volt_Vref_ILoad = GetStoredMeasurement(argv(0))
    Volt_Vref_NoLoad = GetStoredMeasurement(argv(1))
    
    Volt_VrefDelta = Volt_Vref_ILoad.Math.Subtract(Volt_Vref_NoLoad)
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
  
    If gl_UseStandardTestName_Flag = True Then
        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
'''            If gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex) <> "" Then
'''                OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
'''            End If
        TestNameInput = Merge_TName(OutputTname_format)
        TestNameInput = Report_TName_From_Instance(CalcC, "")
        TheExec.Flow.TestLimit resultVal:=Volt_VrefDelta, Tname:=TestNameInput, ForceResults:=tlForceFlow
    Else
        TheExec.Flow.TestLimit resultVal:=Volt_VrefDelta, ForceResults:=tlForceFlow
    End If
    
    
End Function


Public Function Calc_DNL(argc As Long, argv() As String)
    'Calculate DNL= {[result(n+1)-result(n)]*k*256/VDDIO_NAND} -1   ;   k=3
    'Calculate INL= {[result(n)-result(0)]*k*256/VDDIO_NAND} -n   ;   k=3
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    Dim Volt_Vref As New PinListData
    Dim Volt_VDD As Double


    '20190610 for DNL calc
    Dim Volt_Vref__N As New PinListData
    Dim Volt_Vref__CalcDNL As New PinListData
    '20190611 for INL calc
    Dim Volt_Vref__CalcINL As New PinListData
    Dim Volt_Vref__ORG As New PinListData

    
    Dim PowerPinName As String
    Dim k As Single
    Dim n As Integer
    'Dim Kx256 As Integer
    k = Val(argv(2))
    'Kx256 = k * 256
    
    
    Volt_Vref = GetStoredMeasurement(argv(0))
    Volt_VDD = TheExec.specs.DC.Item(argv(1) & "_VAR_H").ContextValue
    
    If Val(TheExec.Flow.var("SrcCodeIndx").Value) > 0 Then
        'theexec.Datalog.WriteComment "*** DNL/DNI: k = " & k & ", k*256 = " & Kx256 & " ***"
        PowerPinName = argv(1)
        n = CInt(TheExec.Flow.var("SrcCodeIndx").Value)
        
        TheExec.Datalog.WriteComment "*** DNL Calculate : DNL(n) = {[result(n) - result(n-1)] * k * 256 / " & PowerPinName & "} - 1 ***   [  n  = " & n & "  ;  " & "k  = " & k & "  ;  " & PowerPinName & " = " & Volt_VDD & "(V)  ]" & "   *** " & TheExec.DataManager.instanceName
    End If
    
    


  
    '20190610 for DNL calc
    If Val(TheExec.Flow.var("SrcCodeIndx").Value) <> 0 Then
    
    
        Volt_Vref__N = GetStoredMeasurement(argv(0) & "_" & CStr(Val(gl_FlowForLoop_DigSrc_SweepCode_Dec) - 1))
        Volt_Vref__CalcDNL = Volt_Vref.Math.Subtract(Volt_Vref__N).Multiply(k * 256).Divide(Volt_VDD).Subtract(1)
        
'''        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
'''        OutputTname_format(6) = UCase(argv(0)) & "-Calc-MeasV-DNL"
'''        TestNameInput = Merge_TName(OutputTname_format)

        
        TestNameInput = Report_TName_From_Instance(CalcC, "", ForceResult:=tlForceFlow)
        
        TheExec.Flow.TestLimit resultVal:=Volt_Vref__CalcDNL, Tname:=TestNameInput, ForceResults:=tlForceFlow, lowVal:=-1, hiVal:=1
    Else
        'added by Kaino on 2019/06/18 for Turks
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        
    End If

''' Masked by Kaino for INL formula update
'''    '--------------------------------------------------------------------------------------------------------
'''    ' added by Kaino on 2019/08/16 for Turks
'''    ' INL(n) = vref_vs_ideal(n)*256*k/VDDIO_NAND
'''    '--------------------------------------------------------------------------------------------------------
'''    Dim Vref_log As String
'''    Dim VDD_log As String
'''
'''    Vref_log = UCase(argv(0))   ' for showing INL formula
'''    VDD_log = UCase(argv(1))    ' for showing INL formula
'''    '--------------------------------------------------------------------------------------------------------
'''
'''
'''    '20190611 for INL calc
'''    If Val(TheExec.Flow.var("SrcCodeIndx").Value) <> 0 Then
'''
'''        Volt_Vref__ORG = GetStoredMeasurement(argv(0) & "_ORG")
'''        If UCase(TheExec.DataManager.InstanceName) Like UCase("*T17*") Then
'''            Volt_Vref__CalcINL = Volt_Vref.Math.Subtract(Volt_Vref__ORG).Multiply(Kx256).Divide(Volt_VDD).Subtract(Val(gl_FlowForLoop_DigSrc_SweepCode_Dec)).Add(50)
'''        Else
'''            'Volt_Vref__CalcINL = Volt_Vref.Math.Subtract(Volt_Vref__ORG).Multiply(Kx256).Divide(Volt_VDD).Subtract(Val(gl_FlowForLoop_DigSrc_SweepCode_Dec))
'''
'''            '--------------------------------------------------------------------------------------------------------
'''            ' added by Kaino on 2019/08/16 for Turks
'''            ' INL(n) = vref_vs_ideal(n)*256*k/VDDIO_NAND
'''            '--------------------------------------------------------------------------------------------------------
'''            TheExec.Datalog.WriteComment "*** INL = [" & Vref_log & "(n)-" & Vref_log & "(0)]*256*K/" & VDD_log
'''            Volt_Vref__CalcINL = Volt_Vref.Math.Subtract(Volt_Vref__ORG).Multiply(Kx256).Divide(Volt_VDD)
'''            '--------------------------------------------------------------------------------------------------------
'''        End If
''''''        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
''''''        OutputTname_format(6) = UCase(argv(0)) & "-Calc-MeasV-INL"
''''''        TestNameInput = Merge_TName(OutputTname_format)
'''
'''
'''        TestNameInput = Report_TName_From_Instance(CalcStr, "", ForceResult:=tlForceFlow)
'''
'''
'''        TheExec.Flow.TestLimit resultVal:=Volt_Vref__CalcINL, Tname:=TestNameInput, ForceResults:=tlForceFlow
'''     Else
'''        'added by Kaino on 2019/06/18 for Turks
'''        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
'''
'''    End If
'''
'''
'''
End Function




Public Function Calc_SweepV_COMPE(argc As Long, argv() As String) As Long
    Dim site As Variant
    Dim k As Integer
    
    Dim DigSrcWave As New DSPWave
    
    
    
    Dim FlowSweepVar As String              '0
    Dim FlowFrom As String                  '1
    Dim FlowTo  As String                   '2
    
    Dim PinName As String                   '3
    Dim VFrom As String                    '4
    Dim VTo    As String                   '5
    Dim Vstep   As String                   '6
    
    Dim DigCapKey1 As String                '7
    Dim DigCapKey2 As String                '8
    
    Dim DigCapWave1 As New DSPWave          '7
    Dim DigCapWave2 As New DSPWave          '8
    
    Dim DatalogArg As String                '9
    
    Dim DatalogOut As New SiteDouble
    
    
    Dim Savekeyname As String
    Dim PS As New DSPWave
    Dim SweepFrom As Integer
    Dim SweepTo As Integer
    
    
    
    FlowSweepVar = argv(0)
    FlowFrom = argv(1)
    FlowTo = argv(2)
    PinName = argv(3)
    VFrom = argv(4)
    VTo = argv(5)
    Vstep = argv(6)
    
    DigCapKey1 = argv(7)
    DigCapKey2 = argv(8)
    DatalogArg = argv(9)
    
'    DigCapKey1 = "comp1"
'    DigCapKey2 = "comp2"
    
    DigCapWave1 = GetStoredCaptureData(DigCapKey1)
    DigCapWave2 = GetStoredCaptureData(DigCapKey2)
    
    Call HIP_Evaluate_ForceVal_New(VFrom)
    Call HIP_Evaluate_ForceVal_New(VTo)
    Call HIP_Evaluate_ForceVal_New(Vstep)
    Call HIP_Evaluate_ForceVal_New(DatalogArg)
        

    k = TheExec.Flow.var(FlowSweepVar).Value


    If k = FlowFrom Then

        For Each site In TheExec.sites
            DigCapStrsCompeA = Str(DigCapWave1.Element(0))
            DigCapStrsCompeB = Str(DigCapWave2.Element(0))
        Next site
        CompeVref = -99.9
        CompeDigcapSwap = -1

    ElseIf (FlowFrom < FlowTo And k <= FlowTo) Or (FlowFrom > FlowTo And k >= FlowTo) Then


        For Each site In TheExec.sites
            If Right(DigCapStrsCompeA, 1) = "0" And DigCapWave1.Element(0) = 1 Then
               ' Stop
                
                CompeVref = CDbl(VFrom) + k * CDbl(Vstep)
                CompeDigcapSwap = k
                
            ElseIf Right(DigCapStrsCompeB, 1) = "0" And DigCapWave2.Element(0) = 1 Then
                'Stop
                CompeVref = CDbl(VFrom) + k * CDbl(Vstep)
                CompeDigcapSwap = k

            End If
            
            
            
'            DigCapStrs = DigCapStrs & Trim(Str(DigCapWave.Element(0)))
            
            
            
            DigCapStrsCompeA = DigCapStrsCompeA & Trim(Str(DigCapWave1.Element(0)))
            DigCapStrsCompeB = DigCapStrsCompeB & Trim(Str(DigCapWave2.Element(0)))

        Next site


        If k = FlowTo Then
            
            

            
            Dim TestNameInput As String
            Dim gl_FlowForLoop_DigSrc_SweepCode_temp As String
            
            
            PS.CreateConstant 0, 1, DspLong

            If FlowTo > FlowFrom Then
                TheExec.Datalog.WriteComment "DigCap:  0         1         2         3         4         5         6         7         8"
            ElseIf SweepTo = 0 Then
                TheExec.Datalog.WriteComment "DigCap:  8         7         6         5         4         3         2         1         0"
            End If
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrsCompeA(site)
                TheExec.Datalog.WriteComment "site(" & site & "):" & DigCapStrsCompeB(site)
                
                
   
            Next site

           'AddStoredCaptureData Savekeyname, DigSrcWave
            
            
            
            
            gl_Sweep_Name = ""
            TestNameInput = Report_TName_From_Instance(CalcC, "X")
            
            TheExec.Flow.TestLimit CompeVref, Tname:=TestNameInput, Unit:=unitVolt, ForceResults:=tlForceFlow
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X")
            DatalogOut = CompeVref.Subtract(CDbl(DatalogArg))
            
            For Each site In TheExec.sites
                If CompeVref = -99.9 Then
                    DatalogOut = -99.9
                End If
            Next site
            TheExec.Flow.TestLimit DatalogOut, Tname:=TestNameInput, Unit:=unitVolt, ForceResults:=tlForceFlow
           
        End If
    Else

    End If

    
    
End Function


'
''''Public Function Calc_DNL_T10(argc As Long, argv() As String)
''''    'Calculate DNL= {[result(n+1)-result(n)]*k*256/VDDIO_NAND} -1   ;   k=3
''''    'Calculate INL= {[result(n)-result(0)]*k*256/VDDIO_NAND} -n   ;   k=3
''''    Dim TestNameInput As String
''''    Dim OutputTname_format() As String
''''
''''    Dim Volt_Vref As New PinListData
''''    Dim Volt_VDD As Double
''''
''''
''''    '20190610 for DNL calc
''''    Dim Volt_Vref__N As New PinListData
''''    Dim Volt_Vref__CalcDNL As New PinListData
''''    '20190611 for INL calc
''''    Dim Volt_Vref__CalcINL As New PinListData
''''    Dim Volt_Vref__ORG As New PinListData
''''
''''    Dim k As Integer
''''    Dim Kx256 As Integer
''''    k = Val(argv(2))
''''    Kx256 = k * 256
''''
''''
''''    If Val(TheExec.Flow.var("SrcCodeIndx").Value) > 0 Then
''''        TheExec.Datalog.WriteComment "*** DNL/DNI: k = " & k & ", k*256 = " & Kx256 & " ***"
''''    End If
''''
''''
''''    Volt_Vref = GetStoredMeasurement(argv(0))
''''    Volt_VDD = TheExec.specs.DC.Item(argv(1) & "_VAR_H").ContextValue
''''
''''
''''    '20190610 for DNL calc
''''    If Val(TheExec.Flow.var("SrcCodeIndx").Value) <> 0 Then
''''
''''
''''        Volt_Vref__N = GetStoredMeasurement(argv(0) & "_" & CStr(Val(gl_FlowForLoop_DigSrc_SweepCode_Dec) - 1))
''''        Volt_Vref__CalcDNL = Volt_Vref.Math.Subtract(Volt_Vref__N).Multiply(Kx256).Divide(Volt_VDD).Subtract(1)
''''
'''''''        Call Report_ALG_TName_From_Instance(OutputTname_format, "V", "X", "", 0)
'''''''        OutputTname_format(6) = UCase(argv(0)) & "-Calc-MeasV-DNL"
'''''''        TestNameInput = Merge_TName(OutputTname_format)
''''
''''
''''        TestNameInput = Report_TName_From_Instance(CalcStr, "", ForceResult:=tlForceFlow)
''''
''''        TheExec.Flow.TestLimit resultVal:=Volt_Vref__CalcDNL, Tname:=TestNameInput, ForceResults:=tlForceFlow, lowVal:=-1, hiVal:=1
''''    Else
''''        'added by Kaino on 2019/06/18 for Turks
''''        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
''''
''''    End If
''''
''''
''''End Function
''''
Public Function Calc_INL(argc As Long, argv() As String)
    '/* --------------------------------------------------------------------------------- */
    '/* create by Kaino on 2019/08/21                                                     */
    '/* --------------------------------------------------------------------------------- */
    '/*     argv(0) = vref_vs_ideal
    '/*     argv(1) = VDDIO_NAND
    '/*     argv(2) = k ;
    '/* --------------------------------------------------------------------------------- */
    '/* INL(n) = vref_vs_ideal(n)*256*k/VDDIO_NAND   ;   k=3 for V1212; k=4.5 for V1818
    '/* --------------------------------------------------------------------------------- */

    '/* Alg::Calc_INL(vref_vs_ideal,VDDIO_NAND,3)
    '/* Alg::Calc_INL(vref_vs_ideal,VDDIO_NAND,4.5)


    Dim INL As New PinListData
    Dim vref_vs_ideal As New PinListData
    Dim n As Integer
    Dim k As Single
    
    Dim VDDIO_NAND As Double
    Dim PowerPinName As String

    Dim Key As String
    
    Key = LCase(Trim(argv(0)) & Trim(CStr(n)))

    
    n = CInt(TheExec.Flow.var("SrcCodeIndx").Value)
    PowerPinName = Trim(argv(1))
    k = Val(argv(2))
    Key = LCase(Trim(argv(0)) & "_" & Trim(CStr(n)))
    vref_vs_ideal = GetStoredMeasurement(Key)


    VDDIO_NAND = TheExec.specs.DC.Item(PowerPinName & "_VAR_H").ContextValue




    TheExec.Datalog.WriteComment "*** INL Calculate : INL(n) = vref_vs_ideal(n) * 256 * k / " & PowerPinName & " ***   [  n  = " & n & "  ;  " & "k  = " & k & "  ;  " & PowerPinName & " = " & VDDIO_NAND & "(V)  ]" & "   *** " & TheExec.DataManager.instanceName




    INL = vref_vs_ideal.Math.Multiply(256 * k).Divide(VDDIO_NAND)

     Dim TestNameInput As String
    TestNameInput = Report_TName_From_Instance(CalcC, "", ForceResult:=tlForceFlow)
    TheExec.Flow.TestLimit resultVal:=INL, Tname:=TestNameInput, ForceResults:=tlForceFlow


End Function


Public Function Calc_ADDRIO_Find_Closest_Result_To_0_Alg1(argc As Integer, argv() As String) As Long
    
    On Error GoTo errHandler
    

'Alg::Calc_ADDRIO_Find_Closest_Result_To_0(sn1,ADDR_P2M_CK_P,ADDRIO_Norm_Y_T1,50,0,8,ADDR_RX_ZCPU,ADDR_TX_ZCPU)

''''''input argument'''''''''
'MeasR result: sn1
'PinName : ADDR_P2M_CK_P
'polynom_result(each code):ADDRIO_Norm_Y_T1()
'Offset:50
'Closest to:0
'StoreDictionary_Bits:8
'StoreDictionary_Name: ADDR_RX_ZCPU,ADDR_TX_ZCPU

    
'1.  Use T1P1 pattern, set code as 35 then measure R from ADDR_P2M_CK_P pins, Get R1 result
'2.  UseT1 coefficient and calculate R1 result * Norm Y (trim code ) -50 , Total 64 result.
'3.  Find the result closets to 0 then the index as the trim code. Need to fuse in FT2.
    

    Dim site As Variant
    Dim i As Long
    Dim MeasR_Results As New PinListData:: MeasR_Results = GetStoredMeasurement(argv(0))
    Dim PinName As String:: PinName = argv(1)
    Dim Offset_Value As Long:: Offset_Value = CDbl(argv(3))
    Dim Closest_to_What_Value As Long:: Closest_to_What_Value = CDbl(argv(4))
    
    Dim Temp_Value_Array() As New SiteVariant
    
    
    If UCase(argv(2)) Like "*T1*" Then
        ReDim Temp_Value_Array(UBound(ADDRIO_Norm_Y_T1))
        
        For i = 0 To UBound(ADDRIO_Norm_Y_T1)
            Temp_Value_Array(i) = MeasR_Results.Math.Multiply(ADDRIO_Norm_Y_T1(i)).Subtract(Offset_Value)
        Next i
        
        
    
    ElseIf UCase(argv(2)) Like "*T2*" Then
        ReDim Temp_Value_Array(UBound(ADDRIO_Norm_Y_T2))
        
        For i = 0 To UBound(ADDRIO_Norm_Y_T2)
            Temp_Value_Array(i) = MeasR_Results.Math.Multiply(ADDRIO_Norm_Y_T2(i)).Subtract(Offset_Value)
        Next i
    
    End If
    
    Dim Temp_Closest_Previous As Double
    Dim Temp_Closest As Double
    Dim Closest_Code As Double
    Dim Closest_Code_Final As New DSPWave
    Dim Closest_Code_Final_Bin As New DSPWave
    
    Closest_Code_Final.CreateConstant 0, 1, DspLong
    For Each site In TheExec.sites.Active
        Temp_Closest_Previous = 999999
        Closest_Code = 999999
        For i = 0 To UBound(Temp_Value_Array)
        
            Temp_Closest = Abs(Temp_Value_Array(i) - Closest_to_What_Value)

            If Temp_Closest <= Temp_Closest_Previous Then
                Temp_Closest_Previous = Temp_Closest
                Closest_Code = i
            End If

        Next i
        Closest_Code_Final.Element(0) = Closest_Code
        Closest_Code_Final_Bin = Closest_Code_Final.ConvertStreamTo(tldspSerial, argv(5), 0, Bit0IsMsb)
        
        If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ")  Closest Code = " & Closest_Code)
        
    Next site
    
    Call AddStoredCaptureData(argv(6), Closest_Code_Final_Bin)
    Call AddStoredCaptureData(argv(7), Closest_Code_Final_Bin)
    
   
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ADDRIO_Find_Closest_Result_To_0"
    If AbortTest Then Exit Function Else Resume Next
    
End Function




Public Function Calc_ADDRIO_Find_Closest_Result_To_0(argc As Integer, argv() As String) As Long
    
    On Error GoTo errHandler
    

'Alg::Calc_ADDRIO_Find_Closest_Result_To_0(sn1,ADDR_P2M_CK_P,ADDRIO_Norm_Y_T1,50,0,8,ADDR_RX_ZCPU,ADDR_TX_ZCPU)

''''''input argument'''''''''
'MeasR result: sn1
'PinName : ADDR_P2M_CK_P
'polynom_result(each code):ADDRIO_Norm_Y_T1()
'Offset:50
'Closest to:0
'StoreDictionary_Bits:8
'StoreDictionary_Name: ADDR_RX_ZCPU,ADDR_TX_ZCPU

    
'1.  Use T1P1 pattern, set code as 35 then measure R from ADDR_P2M_CK_P pins, Get R1 result
'2.  UseT1 coefficient and calculate R1 result * Norm Y (trim code ) -50 , Total 64 result.
'3.  Find the result closets to 0 then the index as the trim code. Need to fuse in FT2.
    

    Dim site As Variant
    Dim i As Long
    Dim MeasR_Results As New PinListData:: MeasR_Results = GetStoredMeasurement(argv(0))
    Dim PinName As String:: PinName = argv(1)
    Dim Offset_Value As Long:: Offset_Value = CDbl(argv(3))
    Dim Closest_to_What_Value As Long:: Closest_to_What_Value = CDbl(argv(4))
    
    Dim Temp_Value_Array() As New SiteVariant
    
    
    If UCase(argv(2)) Like "*T1*" Then
        ReDim Temp_Value_Array(UBound(ADDRIO_Norm_Y_T1))
        
        For i = 0 To UBound(ADDRIO_Norm_Y_T1)
            Temp_Value_Array(i) = MeasR_Results.Math.Multiply(ADDRIO_Norm_Y_T1(i)).Subtract(Offset_Value)
        Next i
        
        
    
    ElseIf UCase(argv(2)) Like "*T2*" Then
        ReDim Temp_Value_Array(UBound(ADDRIO_Norm_Y_T2))
        
        For i = 0 To UBound(ADDRIO_Norm_Y_T2)
            Temp_Value_Array(i) = MeasR_Results.Math.Multiply(ADDRIO_Norm_Y_T2(i)).Subtract(Offset_Value)
        Next i
    
    End If
    
    Dim Temp_Closest_Previous As Double
    Dim Temp_Closest As Double
    Dim Closest_Code As Double
    Dim Closest_Code_Final As New DSPWave
    Dim Closest_Code_Final_Bin As New DSPWave
    
    Closest_Code_Final.CreateConstant 0, 1, DspLong
    For Each site In TheExec.sites.Active
        Temp_Closest_Previous = 999999
        Closest_Code = 999999
        For i = 0 To UBound(Temp_Value_Array)
            Temp_Closest = Temp_Value_Array(i) - Closest_to_What_Value

            If Temp_Closest < 0 Then ' the first result below 0
                Closest_Code = i
                Exit For
            End If


'            Temp_Closest = Abs(Temp_Value_Array(i) - Closest_to_What_Value)
'
'            If Temp_Closest <= Temp_Closest_Previous Then
'                Temp_Closest_Previous = Temp_Closest
'                Closest_Code = i
'            End If
        Next i
        Closest_Code_Final.Element(0) = Closest_Code
        Closest_Code_Final_Bin = Closest_Code_Final.ConvertStreamTo(tldspSerial, argv(5), 0, Bit0IsMsb)
        
        If gl_Disable_HIP_debug_log = False Then Call TheExec.Datalog.WriteComment("Site(" & site & ")  Closest Code = " & Closest_Code)
        
    Next site
    
    Call AddStoredCaptureData(argv(6), Closest_Code_Final_Bin)
    Call AddStoredCaptureData(argv(7), Closest_Code_Final_Bin)
    
   
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ADDRIO_Find_Closest_Result_To_0"
    If AbortTest Then Exit Function Else Resume Next
    
End Function



Public Sub Report_TestLimit_by_CZ_Format(ByVal resultVal As Variant, Optional ByVal lowVal As String, Optional ByVal hiVal As String, _
                                                Optional lowCompareSign As tlCompareSign = tlSignGreaterEqual, Optional highCompareSign As tlCompareSign = tlSignLessEqual, _
                                                Optional scaletype As tlScaleType = scaleNone, Optional Unit As UnitType = unitNone, Optional ByVal formatStr As String = "", _
                                                Optional Tname As String = "", Optional compareMode As tlLimitCompareType, Optional ByVal PinName As String = "x", _
                                                Optional ForceVal As Variant, Optional ForceUnit As UnitType = unitNone, _
                                                Optional customUnit As String, Optional customForceunit As String, _
                                                Optional ForceResults As tlLimitForceResults, _
                                                Optional ByVal MeasType As String = "x", _
                                                Optional ByVal UserVar3 As String = "x", _
                                                Optional ByVal UserVar4 As String = "x", _
                                                Optional ByVal UserVar5 As String = "x", _
                                                Optional ByVal UserVar6 As String = "x", _
                                                Optional ByVal UserVar7 As String = "x", _
                                                Optional ByVal UserVar8 As String = "x", _
                                                Optional ByVal TestSeqNum As Integer = -1, Optional ByVal Tail As String, Optional TNGroup As String = "x")
        '
        '   Segment:        0     1        2       3      4        5        6        7        8        9
        '   TestName:      HAC_USERVAR2_UserVar3_Group_CATEGORY_USERVAR4_USERVAR5_USERVAR6_UserVar7_USERVAR8_
        '   Test name:     HAC_[Meas type]_[HV/NV/LV]_[X1]_[Block]_[Pin name]_[X2]_[X3]_[X4]_[X5]_
        '
        '
        '       Segment:    TestName    :   Meaning         :   arg(priority)   arg(1)                  arg(2)                  arg(3)
        '       ------------------------------------------------------------------------------------------------------------------------------------
        '       01  :       HAC         :   HAC             :
        '       02  :       UserVar2    :   [Meas?]         :                   MeasType                unit
        '       03  :       UserVar3    :   [HV/NV/LV]      :                   UserVar3                (instance name)
        '       04  :       Group       :   [X1]            :                   (instance name)
        '       05  :       CATEGORY    :   [Block]         :                   (instance name)
        '       06  :       UserVar4    :   [Pin name]      :                   UserVar4                PinName                 resultVal
        '       07  :       UserVar5    :   [X2]            :                   (flow table TName)      (instance name)         UserVar5
        '       08  :       UserVar6    :   [X3]            :                   (flow table TName)      (instance name)         UserVar6
        '       09  :       UserVar7    :   [X4]            :                   (flow table TName)      UserVar7                (instance name)
        '       10  :       UserVar8    :   [X5]            :                   (flow table TName)      TestSeqNum >= 0         UserVar8

        'This is v2
        ' 2017-02-02 by Kaino Test Name for PP
        ' 2017-06-02 by Kaion update for auto-gen lib
        ' 2017-08-23 by Kaino : overwrite pin-name in test-name by UserVar4
        ' 2017-10-11 by Kaino : update forceVal

        Dim LoLimitStr As String
        Dim HiLimitStr As String
        Dim testName As String
        Dim LoLimitVal As Double
        Dim HiLimitVal As Double
        Dim p As Long
        Dim ForceValue As Double
        Dim ScaleInFlowTable As tlScaleType
        Dim DisplayFormatInFlowTable As String
    


        'Dim forceStr As String

        Dim MeasureType As String
        Dim TestNameInput As String

        Dim TNameSegments() As String
        Dim TestNameInFlow As String
        
        
        '/* ---
        Dim uv3 As String
        Dim uv4 As String
        Dim uv5 As String
        Dim uv6 As String
        Dim uv7 As String
        Dim uv8 As String
        
        '/* --- 2017-10-11 --- */
        Dim ChannelTypesCount As Long
        Dim ChannelTypesString() As String
        Dim PinOnProcess As String
                '/* --- ---------- --- */

        'Dim resultVal_tmp As Variant
        '---------------------------------------------------------------------------------------
        ' 2017-06-02 by Kaino
        '---------------------------------------------------------------------------------------
        ' for debug by Kaino
        
        '/* update by Kaino on 2019/07/02 */
        If K_InstanceName = "" Or K_InstanceName <> TheExec.DataManager.instanceName Then
            Call CZ_Style_TName_InstanceInfo_Reg("")
            '/* update by Kaino on 2019/07/02 */'TheExec.Datalog.WriteComment " ***  ---" & TheExec.DataManager.InstanceName & " : VBT must call CZ_Style_TName_InstanceInfo_Reg( PatSetName ) in the beginning --- ***"
            '/* update by Kaino on 2019/07/02 */ 'TheExec.Datalog.WriteComment " ***  ---" & TheExec.DataManager.InstanceName & " : VBT must call CZ_Style_TName_InstanceInfo_Clear( )          in the endding   --- ***"
            'Stop
        End If


        uv4 = UserVar4

        'UserVar8 = Trim(Replace(UserVar8, "_", ""))
        If TestSeqNum >= 0 Then
                UserVar8 = Trim(Str(TestSeqNum))
        Else
                UserVar8 = Trim(Replace(UserVar8, "_", ""))
        End If

        '---------------------------------------------------------------------------------------
        '[Meas?]
        If MeasType = "x" Then
            Select Case Unit
                Case unitAmp
                        MeasureType = "I"
                Case unitVolt
                        MeasureType = "V"
                Case unitHz
                        MeasureType = "F"
                Case unitOhm
                        MeasureType = "R"
                Case unitCustom
                    'Stop
                    If Trim(LCase(customUnit)) = "ohm" Then
                            MeasureType = "R"
                    Else
                        TheExec.Datalog.WriteComment " ***  ---" & TheExec.DataManager.instanceName & " should update VBT for CZ Naming by Kaino --- ***"
                        Stop
                    End If
                Case Else
                    MeasureType = "X"
                    TheExec.Datalog.WriteComment " ***  ---" & TheExec.DataManager.instanceName & " should update VBT for CZ Naming by Kaino --- ***"
                    Stop
            End Select
            MeasType = MeasureType
        ElseIf MeasType = "C" Then
            scaletype = scaleNoScaling
            'for debugging
            'Stop
        ElseIf MeasType = "X" Then
            MeasType = "Calc"
            'for debugging
            'TheExec.Datalog.WriteComment " ***  ---" & TheExec.DataManager.InstanceName & " should update VBT for CZ Naming by Kaino --- ***"
            'Stop
        Else
            'MeasType = trim(MeasType)
        End If

        '---------------------------------------------------------------------------------------
        ' Get Limit
        If ForceResults = tlForceNone Then

                LoLimitStr = lowVal
                HiLimitStr = hiVal

        ElseIf ForceResults = tlForceFlow Then

                Call GetTestInfo_single(LoLimitStr, HiLimitStr, TestNameInFlow, ScaleInFlowTable:=ScaleInFlowTable, DisplayFormat:=DisplayFormatInFlowTable)
                
                If DisplayFormatInFlowTable <> "" Then
                    formatStr = DisplayFormatInFlowTable
                End If

                If ScaleInFlowTable <> scaleNone Then
                    scaletype = ScaleInFlowTable
                End If

        Else
            ForceResults = tlForceNone
            LoLimitStr = lowVal
            HiLimitStr = hiVal
        End If


        If LoLimitStr = "" Then
                lowCompareSign = tlSignNone
        Else
                lowCompareSign = tlSignGreaterEqual
                LoLimitVal = Val(LoLimitStr)                    '2016-06-05 by Kaino
        End If

        If HiLimitStr = "" Then
                highCompareSign = tlSignNone
        Else
                highCompareSign = tlSignLessEqual
                HiLimitVal = Val(HiLimitStr)                    '2016-06-05 by Kaino
        End If
        '---------------------------------------------------------------







        If TypeName(resultVal) = "IPinListData" Then
                For p = 0 To resultVal.Pins.Count - 1

                    'forceValue
                    If TypeName(ForceVal) = "IPinListData" Then
                    
                        '/*--- 2017-10-11 ---*/
                        PinOnProcess = resultVal.Pins(p).Name
                        
                        Call TheExec.DataManager.GetChannelTypes(ForceVal.Pins(PinOnProcess), ChannelTypesCount, ChannelTypesString())
                        
                        
                        'Debug.Print ChannelTypesString(0)
                        If ChannelTypesString(0) Like "DCVS*" Then
                            
                            ForceValue = Round(TheHdw.DCVS.Pins(PinOnProcess).Voltage.Main.Value, 3)
                            
                        ElseIf ChannelTypesString(0) = "I/O" Then
                        '/*--- ---------- ---*/
                                
                                If ForceUnit = unitVolt Then
                                        ForceValue = Round(TheHdw.PPMU(ForceVal.Pins(p)).Voltage.Value, 3)
                                ElseIf ForceUnit = unitAmp Then
                                        ForceValue = Round(TheHdw.PPMU(ForceVal.Pins(p)).current.Value, 3)
                                End If
                            
                        Else
                            ForceValue = 0
                        End If
                    ElseIf IsNumeric(ForceVal) Then
                            ForceValue = Val(ForceVal)
                    Else
                            ForceValue = 0
                    End If


                        PinName = resultVal.Pins(p).Name

                        '2017/08/23 : updated by Kaino :  'overwrit pinname in testname by UserVar4
                        If UserVar4 = "x" Then
                            uv4 = PinName
                        Else

                        End If


                        '2017/07/25 : updated by Kaino
                        'If TestNameInFlow = "" Then
                        'TestNameInput = Report_TName(Group:=TNGroup, UserVar2:=MeasType, UserVar4:=PinName, UserVar7:=UserVar7, UserVar8:=UserVar8, TestNameInFlow:=TestNameInFlow)
                        '2017/08/23 : updated by Kaino :  'overwrit pinname in testname by UserVar4
                        TestNameInput = Report_TName(Group:=TNGroup, UserVar2:=MeasType, UserVar4:=uv4, UserVar5:=UserVar5, UserVar6:=UserVar6, UserVar7:=UserVar7, UserVar8:=UserVar8, TestNameInFlow:=TestNameInFlow)

                        'End If

                        TestNameInput = TestNameInput & Tail

                        
                        '/* --------------------------------------------------------------- */
                        '/* --- overwrite pinname ("{pinname}" in TName from Flow Table)--- */
                        '/* --------------------------------------------------------------- */
                        '  moved to Report_TName()
                        '/* --------------------------------------------------------------- */
                        
                        'tlForceNone
                        'TheExec.Flow.TestLimit resultVal, lowVal, hiVal, lowCompareSign, highCompareSign, scaletype, unit:=unitHz, TName:=TestNameInput, forceVal:=forceValue, forceunit:=forceunit, ForceResults:=ForceResults
                        'tlForceFlow
                        '/* ------------------------------------------------------------------------------ */
                        '/* --------------------------------- 2017-11-15 --------------------------------- */
                        'TheExec.Flow.TestLimit resultVal.Pins(p), LoLimitVal, HiLimitVal, lowCompareSign, highCompareSign, scaletype:=scaletype, unit:=unit, Tname:=TestNameInput, forceVal:=ForceValue, customUnit:=customUnit, forceunit:=forceunit, ForceResults:=tlForceNone, formatStr:=formatStr

                        If ForceResults = tlForceNone Then
                            TheExec.Flow.TestLimit resultVal.Pins(p), LoLimitVal, HiLimitVal, lowCompareSign, highCompareSign, scaletype:=scaletype, Unit:=Unit, Tname:=TestNameInput, ForceVal:=ForceValue, customUnit:=customUnit, ForceUnit:=ForceUnit, ForceResults:=ForceResults, formatStr:=formatStr
                        ElseIf ForceResults = tlForceFlow Then
                            TheExec.Flow.TestLimit resultVal.Pins(p), Unit:=Unit, scaletype:=scaletype, Tname:=TestNameInput, ForceVal:=ForceValue, customUnit:=customUnit, ForceUnit:=ForceUnit, ForceResults:=ForceResults, formatStr:=formatStr

                            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
                        End If
                        '/* --------------------------------- 2017-11-15 --------------------------------- */
                        '/* ------------------------------------------------------------------------------ */
                Next p
                '/* ------------------------------------------------------------------------------ */
                '/* --------------------------------- 2017-11-15 --------------------------------- */
                If ForceResults = tlForceNone Then
                    ' do nothing
                ElseIf ForceResults = tlForceFlow Then
                    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
                End If
                '/* --------------------------------- 2017-11-15 --------------------------------- */
                '/* ------------------------------------------------------------------------------ */

        ElseIf TypeName(resultVal) = "IPinData" Or TypeName(resultVal) = "ISiteDouble" Or TypeName(resultVal) = "ISiteLong" Or IsNumeric(resultVal) Or TypeName(resultVal) = "ISiteComplexDouble" Or TypeName(resultVal) = "ISiteVariant" Then

                Dim result_value As New SiteVariant

                result_value = resultVal

            'forceValue
            If TypeName(ForceVal) = "IPinListData" Then

                    '/*--- 2017-10-11 ---*/
                    PinOnProcess = resultVal.Name
                        
                    Call TheExec.DataManager.GetChannelTypes(ForceVal.Pins(PinOnProcess), ChannelTypesCount, ChannelTypesString())
                    
                    Debug.Print ChannelTypesString(0)
                    If ChannelTypesString(0) Like "DCVS*" Then
                        
                        ForceValue = Round(TheHdw.DCVS.Pins(PinOnProcess).Voltage.Main.Value, 3)
                        
                    ElseIf ChannelTypesString(0) = "I/O" Then
                    
                    '/*--- ---------- ---*/
                
                        If ForceUnit = unitVolt Then
                                ForceValue = Round(TheHdw.PPMU(ForceVal.Pins(p)).Voltage.Value, 3)
                        ElseIf ForceUnit = unitAmp Then
                                ForceValue = Round(TheHdw.PPMU(ForceVal.Pins(p)).current.Value, 3)
                        End If
                    Else
                        ForceValue = 0
                        
                    End If
            ElseIf IsNumeric(ForceVal) Then
                    ForceValue = Val(ForceVal)
            Else
                    ForceValue = 0
            End If


                If PinName = "x" Then
                    If TypeName(resultVal) = "IPinData" Then
                        PinName = resultVal.Name
                    End If
                End If

                '2017/08/23 : updated by Kaino : overwrite pinname in testname by UserVar4
                If UserVar4 = "x" Then
                    uv4 = PinName
                Else
                '2017/08/23 : updated by Kaino : do not show "x" in column "pin" in datalog
                End If
                If PinName = "x" Then
                    PinName = ""
                End If


                '2017/07/25 : updated by Kaino
                'If TestNameInFlow = "" Then
                'TestNameInput = Report_TName(Group:=TNGroup, UserVar2:=MeasType, UserVar4:=PinName, UserVar5:=UserVar5, UserVar6:=UserVar6, UserVar7:=UserVar7, UserVar8:=UserVar8, TestNameInFlow:=TestNameInFlow)
                '2017/08/23 : updated by Kaino :  'overwrit pinname in testname by UserVar4
                TestNameInput = Report_TName(Group:=TNGroup, UserVar2:=MeasType, UserVar3:=UserVar3, UserVar4:=uv4, UserVar5:=UserVar5, UserVar6:=UserVar6, UserVar7:=UserVar7, UserVar8:=UserVar8, TestNameInFlow:=TestNameInFlow)
                'End If

                TestNameInput = TestNameInput & Tail

                '/* --------------------------------------------------------------- */
                '/* --- overwrite pinname ("{pinname}" in TName from Flow Table)--- */
                '/* --------------------------------------------------------------- */
                '  moved to Report_TName()
                '/* --------------------------------------------------------------- */


                '/* ------------------------------------------------------------------------------ */
                '/* --------------------------------- 2017-11-15 --------------------------------- */
                'TheExec.Flow.TestLimit result_value, LoLimitVal, HiLimitVal, lowCompareSign, highCompareSign, scaletype:=scaletype, unit:=unit, Tname:=TestNameInput, forceVal:=ForceValue, customUnit:=customUnit, forceunit:=forceunit, ForceResults:=tlForceNone, PinName:=PinName, formatStr:=formatStr
                If ForceResults = tlForceNone Then
                    TheExec.Flow.TestLimit result_value, LoLimitVal, HiLimitVal, lowCompareSign, highCompareSign, scaletype:=scaletype, Unit:=Unit, Tname:=TestNameInput, ForceVal:=ForceValue, customUnit:=customUnit, ForceUnit:=ForceUnit, ForceResults:=ForceResults, PinName:=PinName, formatStr:=formatStr
                    
                ElseIf ForceResults = tlForceFlow Then
                    TheExec.Flow.TestLimit result_value, Unit:=Unit, scaletype:=scaletype, Tname:=TestNameInput, ForceVal:=ForceValue, customUnit:=customUnit, ForceUnit:=ForceUnit, ForceResults:=ForceResults, PinName:=PinName, formatStr:=formatStr
                End If
                '/* --------------------------------- 2017-11-15 --------------------------------- */
                '/* ------------------------------------------------------------------------------ */

        ElseIf IsNumeric(resultVal) Then
                TheExec.Datalog.WriteComment " ***  ---" & TheExec.DataManager.instanceName & " should update VBT for CZ Naming by Kaino --- ***"
                Stop
        Else
                Debug.Print TypeName(resultVal)
                TheExec.Datalog.WriteComment " ***  ---" & TheExec.DataManager.instanceName & " should update VBT for CZ Naming by Kaino --- ***"
                Stop
        End If


End Sub


Public Function CZ_Style_TName_InstanceInfo_Reg(ByVal PatSetName As String)
        ' V0: 2017-06-02 by Kaino
        K_InstanceName = TheExec.DataManager.instanceName

        K_PatSetName = Trim(PatSetName)
        K_InstanceName_WO_Pset = Replace(K_InstanceName, K_PatSetName, "", , , vbTextCompare)
        K_InstanceName_WO_Pset = Replace(K_InstanceName_WO_Pset, "__", "_", , , vbTextCompare)


        '/*
        '/*  Remove Useless Word from intance-name
        '/*

        ' At Head   :   "DCTEST"
        If K_InstanceName_WO_Pset Like "DCTEST_*" Then
            K_InstanceName_WO_Pset = Replace(K_InstanceName_WO_Pset, "DCTEST_", "", , , vbTextCompare)
        End If
        ' All places:   "VIR"
        K_InstanceName_WO_Pset = Replace(K_InstanceName_WO_Pset, "_VIR_", "_", , , vbTextCompare)


        InstNameSegs() = Split(K_InstanceName_WO_Pset, "_")





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
        TNameSeg(2) = Replace(TNameSeg(2), "v", "", compare:=vbTextCompare)

        '[X1]
        If UBound(InstNameSegs) >= 2 Then
            TNameSeg(3) = InstNameSegs(1)
        End If

        '[X2]_[X3]_[X4]
        If UBound(InstNameSegs) >= 5 Then
                TNameSeg(6) = InstNameSegs(UBound(InstNameSegs) - 3)        '[X2]
                TNameSeg(7) = InstNameSegs(UBound(InstNameSegs) - 2)        '[X3]
                TNameSeg(8) = InstNameSegs(UBound(InstNameSegs) - 1)        '[X4]
        ElseIf UBound(InstNameSegs) >= 4 Then
                TNameSeg(6) = InstNameSegs(UBound(InstNameSegs) - 2)        '[X2]
                TNameSeg(7) = InstNameSegs(UBound(InstNameSegs) - 1)        '[X3]
        ElseIf UBound(InstNameSegs) >= 3 Then
                TNameSeg(6) = InstNameSegs(UBound(InstNameSegs) - 1)        '[X2]
        End If

        'Call SetupDatalogFormat(TestNameW:=90, PatternW:=100)
        
        '/* update by Kaino on 2019/07/02 */
        Dim SubBlockName1_from_FlowSheet As String
        SubBlockName1_from_FlowSheet = Trim(Application.Worksheets(TheExec.Flow.Raw.SheetInRun).range("AM" & CStr(TheExec.Flow.Raw.GetCurrentLineNumber + 5)).Value)
        
        If SubBlockName1_from_FlowSheet <> "" Then
            TNameSeg(3) = Trim(Application.Worksheets(TheExec.Flow.Raw.SheetInRun).range("AM" & CStr(TheExec.Flow.Raw.GetCurrentLineNumber + 5)).Value)
        End If
End Function

Public Sub GetTestInfo_single(LowLimit As String, HighLimit As String, Optional testName As String, Optional ScaleInFlowTable As tlScaleType, Optional DisplayFormat As String)
        Dim FlowLimitsInfo As IFlowLimitsInfo

        Dim TestNames() As String
        Dim HighLimits() As String
        Dim LowLimits() As String
        Dim ScaleInFlowTables() As tlScaleType
        Dim DisplayFormats() As String                                  'updated by Kaino on 2017/09/21
        
        Dim LimitIndex As Long


        Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)

        If FlowLimitsInfo Is Nothing Then       '/* --- Kaino 2017/08/25 --- */
        
            TheExec.Datalog.WriteComment " ***  --- " & TheExec.DataManager.instanceName & " ---- Use-Limit should be updated in Flow Table: " & TheExec.Flow.CurrentFlowSheetName & " : for CZ Naming --- ***"
    
        Else
            Call FlowLimitsInfo.GetTNames(TestNames())
            Call FlowLimitsInfo.GetHighLimits(HighLimits())
            Call FlowLimitsInfo.GetLowLimits(LowLimits())
            Call FlowLimitsInfo.GetScales(ScaleInFlowTables())
            Call FlowLimitsInfo.GetFormats(DisplayFormats)              'updated by Kaino on 2017/09/21
    
            LimitIndex = TheExec.Flow.TestLimitIndex
    
            'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1      'Marked by Kaino on 2017/11/15
    
            If LimitIndex <= UBound(TestNames) Then
                testName = TestNames(LimitIndex)
                HighLimit = HighLimits(LimitIndex)
                LowLimit = LowLimits(LimitIndex)
                ScaleInFlowTable = ScaleInFlowTables(LimitIndex)
                DisplayFormat = DisplayFormats(LimitIndex)       'updated by Kaino on 2017/09/21
            Else
                'TheExec.Datalog.WriteComment " *** CZ Test Name ERROR! *** Please Add more Use-Limit in Flow Table : " & TheExec.Flow.CurrentFlowSheetName & " for test"
                TheExec.Datalog.WriteComment " "
                TheExec.Datalog.WriteComment " ***  --- " & TheExec.DataManager.instanceName & " ---- Use-Limit should be updated in Flow Table: " & TheExec.Flow.CurrentFlowSheetName & " : for CZ Naming --- ***"
                TheExec.Datalog.WriteComment " "
                'Stop
            End If
        End If

End Sub
Public Function Report_TName(ByVal Group As String, Optional ByVal UserVar2 As String = "x", Optional ByVal UserVar3 As String = "x", Optional ByVal UserVar4 As String = "x", _
                                           Optional ByVal UserVar5 As String = "x", Optional ByVal UserVar6 As String = "x", Optional ByVal UserVar7 As String = "x", _
                                           Optional ByVal UserVar8 As String = "x", Optional ByVal TestNameInFlow As String = "") As String
        ' This is V3
        ' 2017/06/02 : by Kaino
        ' 2017/08/25
        
        
        Dim TNameInFlowSegs() As String
        Dim UserVarInFlow(5 To 8) As String
        Dim i As Integer
        'Dim UserVar3 As String
        Dim Category As String
        Dim Tname As String
        
        

        '   Segment:        0     1        2       3      4        5        6        7        8        9
        '   TestName:      HAC_USERVAR2_UserVar3_Group_CATEGORY_USERVAR4_USERVAR5_USERVAR6_UserVar7_USERVAR8_
        '   Meaning:       HAC_[Meas? ]_[H/N/L ]_SubB1_[Block ]_[ Pin  ]_[SubB2 ]_[  X   ]_[  X   ]_subr-seq_
        '   X:                                    X1                       X2        X3       X4       X5


        '0  HAC
        '1  [Meas?]
                If Len(UserVar2) = 1 Then   '/* update by Kaino on 2019/07/02 */
                    UserVar2 = "Meas" & UCase(UserVar2)
                End If
        '2  [HV/LV/NV]
                'TNameSeg(2) = InstNameSegs(UBound(InstNameSegs))               'init at CZ_Style_TName_InstanceInfo_Reg()
                If UserVar3 = "x" Then
                    UserVar3 = TNameSeg(2)
                End If
        '3  [X1]
                'sub-block-name-1
                If Group = "x" Then
                        Group = TNameSeg(3)
                End If
        '4  [Block]
                'TNameSeg(4) = InstNameSegs(0)                                  'init at CZ_Style_TName_InstanceInfo_Reg()
                Category = TNameSeg(4)

        '5  [Pin name]
                'UserVar4 =" {pinname}"
                UserVar4 = Replace(UserVar4, "_", "")
        '6  [X2]
                If UserVar5 = "x" Then
                   UserVar5 = TNameSeg(6)
                Else
                   UserVar5 = Trim(Replace(UserVar5, "_", ""))
                End If
        '7  [X3]

                If TNameSeg(7) <> "x" Then
                   UserVar6 = TNameSeg(7)
                Else
                   UserVar6 = Trim(Replace(UserVar6, "_", ""))
                End If

'                    If UserVar6 = "x" Then
'                        UserVar6 = TNameSeg(7)
'                    Else
'                        UserVar6 = Trim(Replace(UserVar6, "_", ""))
'                    End If

        '8  [X4]
                If TNameSeg(8) <> "x" Then
                   UserVar7 = TNameSeg(8)
                Else
                   UserVar7 = Replace(UserVar7, "_", "")
                End If

        '9  [X5]
                UserVar8 = Replace(UserVar8, "_", "")


        
        '/* ------------------------------------- */
        '/* --- overwrite TName by Flow Table --- */
        '/* ------------------------------------- */

        If TestNameInFlow = "" Then

            'TName = "HAC_" & UserVar2 & "_" & UCase(UserVar3 & "_" & Group & "_" & Category & "_" & UserVar4 & "_" & UserVar5 & "_" & UserVar6 & "_" & UserVar7 & "_" & UserVar8 & "_")
            Tname = "HAC_" & UserVar2 & "_" & UCase(UserVar3 & "_" & Group & "_" & Category & "_") & UserVar4 & UCase("_" & UserVar5 & "_" & UserVar6 & "_" & UserVar7 & "_" & UserVar8)
        Else
            
            

            If TestNameInFlow Like "HAC*" Then
                Tname = TestNameInFlow
            
            Else
            
                TNameInFlowSegs = Split(TestNameInFlow, "_")
                
                For i = 5 To 8
                    UserVarInFlow(i) = "x"
                    
                    If i - 5 <= UBound(TNameInFlowSegs) Then
                        UserVarInFlow(i) = TNameInFlowSegs(i - 5)
                    End If
                
                Next i
                
            
                If UBound(TNameInFlowSegs) >= 0 Then

                    '/* TestNameInFlow = "U5_U6_U7_U8"
                    If UserVarInFlow(5) <> "x" Then UserVar5 = UserVarInFlow(5)
                    
                    'remove reg name(UseVar7)  when  UserVar5 = UseVar7
                    If UserVar5 = UserVar7 Then
                        UserVar7 = "X"
                    End If

                    
                    If UserVarInFlow(6) <> "x" Then UserVar6 = UserVarInFlow(6)
                    If UserVarInFlow(7) <> "x" Then UserVar7 = UserVarInFlow(7)
                    If UserVarInFlow(8) <> "x" Then UserVar8 = UserVarInFlow(8)
                                   
                
                Else
                    TheExec.Datalog.WriteComment " ***  ---" & TheExec.DataManager.instanceName & " should update VBT for CZ Naming by Kaino --- ***"
                    Stop
                    'please call Kaino to check,thanks
                    
                    UserVar5 = Trim(Replace(TestNameInFlow, "_", ""))
                    If UserVar5 = UserVar7 Then
                        UserVar7 = "X"
                    End If
                End If
                
                
                
                
                
                
                'TName = "HAC_" & UserVar2 & "_" & UCase(UserVar3 & "_" & Group & "_" & Category & "_" & UserVar4 & "_" & UserVar5 & "_" & UserVar6 & "_" & UserVar7 & "_" & UserVar8 & "_")
                Tname = "HAC_" & UserVar2 & "_" & UCase(UserVar3 & "_" & Group & "_" & Category & "_") & UserVar4 & UCase("_" & UserVar5 & "_" & UserVar6 & "_" & UserVar7 & "_" & UserVar8)

            End If


            '/* ------------------------------------------------------------- */
            '/* --- overwrite PinName ("pinname" in TName from Flow Table)--- */
            '/* ------------------------------------------------------------- */

            If InStr(1, Tname, "{pinname}", vbTextCompare) > 0 Then
                Tname = Replace(Tname, "{pinname}", UserVar4, 1, compare:=vbTextCompare)
            End If

            '/* --------------------------------------------- */


        End If
        '-------------------------------------------------------------------------------------------------
        'Tname = "HAC_" & TNameSeg(1) & "_" & TNameSeg(2) & "_" & TNameSeg(3) & "_" & TNameSeg(4) & "_" & TNameSeg(5) & "_" & TNameSeg(6) & "_" & TNameSeg(7) & "_" & TNameSeg(8) & "_" & TNameSeg(9) & "_"



        Report_TName = Tname


End Function

