Attribute VB_Name = "VBT_LIB_HardIP_AP"
Option Explicit
Public Function HIP_eFuse_Read_TMPS_Coeff(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, dspwavesize As Long, Optional Efuse_Read_Dec_Flag As Boolean = False, Optional Dict_Store_Dec_Name As String = "") As Long

    ' Parameter : eFuse Block , eFuse Variable , data , Data Width
    ' Create dictionary , if exist then remove and re-create
    ' MUST :  if necessary , we can set limit if read out value = 0 then bin out .

    Dim site As Variant
    Dim Read_Code As New DSPWave
    Dim Read_Value As New DSPWave
    Dim Efuse_Value As New SiteLong
    Dim TempVal As Long
    Dim Efuse_Value_Chk As New SiteVariant
    Dim i As Long

    On Error GoTo errHandler

    Read_Code.CreateConstant 0, dspwavesize

    If Efuse_Read_Dec_Flag = True Then
        Read_Value.CreateConstant 0, 1
    End If

    For Each site In TheExec.sites

        Efuse_Value(site) = auto_eFuse_GetReadDecimal(FuseType, m_catename, True)
'''''        Efuse_Value(Site) = CLng(Site) + 8

        If Efuse_Read_Dec_Flag = True Then
            Read_Value.Element(0) = Efuse_Value(site)
        End If

        TempVal = Efuse_Value(site)
        For i = 0 To dspwavesize - 1
            Read_Code.Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i

        If Efuse_Value(site) = 0 Then
        'If Read out value = 0 then bin out
            Efuse_Value_Chk(site) = 0
        Else
            Efuse_Value_Chk(site) = 1
        End If

    Next site

    'TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=1, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceNone

    Call AddStoredCaptureData(Dict_Store_Code_Name, Read_Code)

    If Efuse_Read_Dec_Flag = True Then
        Call AddStoredCaptureData(Dict_Store_Dec_Name, Read_Value)
    End If

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Read"
    If AbortTest Then Exit Function Else Resume Next

End Function


''
''Public Function pll_read() As Long
''
''    Call HIP_eFuse_Read(A, b, c)
''    Call HIP_eFuse_Read
''
''End Function


Public Function TMPS(patset As Pattern, CPUA_Flag_In_Pat As Boolean, DigSrc_pin As PinList, DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, DigSrc_Equation As String, DigSrc_Assignment As String, _
                           Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
                            Optional CUS_Str_DigSrcData As String, Optional CUS_Str_DigCapData As String = "", Optional TMPS_Fuse_string As String, Optional Validating_ As Boolean) As Long

'' Step 1 : trim code is 8 bit, show out measured volt and trimed code, target volt is 0.9v
'' Step 2 : start from 0x8 and add algorithm to decide +/- direction
'' while decimal < 2 ^ DigSrc_Sample_Size
'' convert decimal to binary reverse
'' input the binary reverse data to digSrc_assignment

    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If
    Dim X As Long
    Dim InDSPwave As New DSPWave
    Dim OutDspWave As New DSPWave
    Dim CapOut As String
    Dim SrcOut As String
    Dim site As Variant
    Dim Pat As String
    Dim i As Integer
    Dim ShowDec As String
    Dim ShowOut As String
    Dim TrimBits As String
    Dim b_TestDone As Boolean
    Dim code(7) As Integer
    Dim SourceNum As Integer
    Dim k As Integer
    Dim CtrlBits As New DSPWave
    Dim MinValue As New DSPWave
    Dim Data_Array(7) As New SiteLong
    Dim PassFlag_TMPS As New SiteBoolean
    On Error GoTo errHandler

    b_TestDone = False
    SourceNum = 0
    CtrlBits.CreateConstant 0, 2 * DigSrc_Sample_Size / DigSrc_DataWidth
    MinValue.CreateConstant 0, DigSrc_Sample_Size / DigSrc_DataWidth

    If DigSrc_Sample_Size = 0 Then
        TheExec.Datalog.WriteComment ("Error!! - Please check input argument DigSrc_Sample_Size")
        Exit Function
    End If

    TheHdw.Digital.Patgen.Halt
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Patterns(patset).Load

    Dim PattArray() As String
    Dim PatCount As Long

    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode

    Do While b_TestDone = False
        For Each site In TheExec.sites.Active

          ''  theexec.Datalog.WriteComment ("======== Start Dig Src setup =======")
            If SourceNum = 0 Then
                Call Create_DigSrc_Data(DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPwave, site)
            End If

            '' Processing Captured Code and Printing the Values
            If SourceNum > 0 Then
                For k = 0 To DigCap_Sample_Size / DigCap_DataWidth - 1
                code(k) = 0
                    For i = 0 To DigCap_DataWidth - 1
                        code(k) = code(k) + OutDspWave(site).Element(k * DigCap_DataWidth + i) * (2 ^ i)
                    Next i
                ''   Code(k) = 3770 + Int(Rnd * 20)    '' Random Codes for Offline testing
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Capture data_ " & k & "_Value =" & code(k)
                Next k
            End If

            If SourceNum > 0 And SourceNum < 6 Then
                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                    If code(k) > 3780 Then
                        For X = 0 To 3
                            InDSPwave(site).Element(k * 28 + 4 * (7 - SourceNum) + X) = 0
                        Next X
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 1) = 0
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 2) = 0
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 3) = 0
                    Else
                        For X = 0 To 3
                            InDSPwave(site).Element(k * 28 + 4 * (7 - SourceNum) + X) = 1
                        Next X
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 1) = 1
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 2) = 1
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 3) = 1
                    End If

                    If SourceNum < 5 Then
                        For X = 0 To 3
                            InDSPwave(site).Element(k * 28 + 4 * (6 - SourceNum) + X) = 1
                        Next X
                        'InDSPWave(site).Element(k * 28 + 4 * (6 - SourceNum) + 1) = 1
                        'InDSPWave(site).Element(k * 28 + 4 * (6 - SourceNum) + 2) = 1
                        'InDSPWave(site).Element(k * 28 + 4 * (6 - SourceNum) + 3) = 1
                    End If
                Next k

            End If

            ''' Trim Control(first 2) Bits

            If SourceNum > 4 And SourceNum < 9 Then
                Select Case SourceNum - 5

                    Case 0
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            For X = 0 To 7
                                InDSPwave(site).Element(k * 28 + X) = 0
                            Next X
                            'InDSPWave(site).Element(k * 28 + 1) = 0
                            'InDSPWave(site).Element(k * 28 + 2) = 0
                            'InDSPWave(site).Element(k * 28 + 3) = 0
                            'InDSPWave(site).Element(k * 28 + 4) = 0
                            'InDSPWave(site).Element(k * 28 + 5) = 0
                            'InDSPWave(site).Element(k * 28 + 6) = 0
                            'InDSPWave(site).Element(k * 28 + 7) = 0
                        Next k


                    Case 1
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            MinValue(site).Element(k) = code(k)
                            For i = 0 To 15
                                CtrlBits(site).Element(i) = 0
                            Next i

                            For X = 0 To 3
                                InDSPwave(site).Element(k * 28 + X) = 0
                            Next X
                            'InDSPWave(site).Element(k * 28 + 1) = 0
                            'InDSPWave(site).Element(k * 28 + 2) = 0
                            'InDSPWave(site).Element(k * 28 + 3) = 0
                            For X = 4 To 7
                                InDSPwave(site).Element(k * 28 + X) = 1
                            Next X
                            'InDSPWave(site).Element(k * 28 + 5) = 1
                            'InDSPWave(site).Element(k * 28 + 6) = 1
                            'InDSPWave(site).Element(k * 28 + 7) = 1
                        Next k


                    Case 2
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            If code(k) < MinValue(site).Element(k) Then
                                MinValue(site).Element(k) = code(k)
                                CtrlBits(site).Element(2 * k) = 0
                                CtrlBits(site).Element(2 * k + 1) = 1
                            End If

                            For X = 0 To 3

                                InDSPwave(site).Element(k * 28 + X) = 1
                            Next
                            'InDSPWave(site).Element(k * 28 + 1) = 1
                            'InDSPWave(site).Element(k * 28 + 2) = 1
                            'InDSPWave(site).Element(k * 28 + 3) = 1
                            For X = 4 To 7
                                InDSPwave(site).Element(k * 28 + X) = 0
                            Next X
                            'InDSPWave(site).Element(k * 28 + 5) = 0
                            'InDSPWave(site).Element(k * 28 + 6) = 0
                            'InDSPWave(site).Element(k * 28 + 7) = 0
                        Next k


                    Case 3
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            If code(k) < MinValue(site).Element(k) Then
                                MinValue(site).Element(k) = code(k)
                                CtrlBits(site).Element(2 * k) = 1
                                CtrlBits(site).Element(2 * k + 1) = 0
                            End If

                            For X = 0 To 7
                                InDSPwave(site).Element(k * 28 + X) = 1
                            Next X
                            'InDSPWave(site).Element(k * 28 + 1) = 1
                            'InDSPWave(site).Element(k * 28 + 2) = 1
                            'InDSPWave(site).Element(k * 28 + 3) = 1
                            'InDSPWave(site).Element(k * 28 + 4) = 1
                            'InDSPWave(site).Element(k * 28 + 5) = 1
                            'InDSPWave(site).Element(k * 28 + 6) = 1
                            'InDSPWave(site).Element(k * 28 + 7) = 1
                        Next k

                End Select
            End If


            ''''' Sourcing Trimmed Data Bits

            If SourceNum = 9 Then

                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                    If code(k) < MinValue(site).Element(k) Then
                        MinValue(site).Element(k) = code(k)
                        CtrlBits(site).Element(2 * k) = 1
                        CtrlBits(site).Element(2 * k + 1) = 1
                    End If

                Next k

                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1
                    For X = 0 To 3
                        InDSPwave(site).Element(k * 28 + X) = CtrlBits(site).Element(2 * k)
                    Next X

                    'InDSPWave(site).Element(k * 28 + 1) = CtrlBits(site).Element(2 * k)
                    'InDSPWave(site).Element(k * 28 + 2) = CtrlBits(site).Element(2 * k)
                    'InDSPWave(site).Element(k * 28 + 3) = CtrlBits(site).Element(2 * k)
                    For X = 4 To 7
                        InDSPwave(site).Element(k * 28 + X) = CtrlBits(site).Element(2 * k + 1)
                    Next X
                    'InDSPWave(site).Element(k * 28 + 5) = CtrlBits(site).Element(2 * k + 1)
                    'InDSPWave(site).Element(k * 28 + 6) = CtrlBits(site).Element(2 * k + 1)
                    'InDSPWave(site).Element(k * 28 + 7) = CtrlBits(site).Element(2 * k + 1)
                Next k

                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1
                    Data_Array(k) = 0
                    For i = 0 To DigSrc_DataWidth / 4 - 1
                        Data_Array(k) = Data_Array(k) + InDSPwave(site).Element(k * DigSrc_DataWidth + 4 * i) * (2 ^ i)
                    Next i
                Next k

            End If



        Next site


        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "Meas_Src", DigSrc_Sample_Size, InDSPwave)

        If SourceNum = 9 Then
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "==========SOURCING TRIMMED DATA BITS==============="
        End If

        For Each site In TheExec.sites.Active
            If SourceNum > 0 Then
                SrcOut = ""
                For i = 0 To DigSrc_Sample_Size - 1
                    SrcOut = SrcOut & InDSPwave(site).Element(i)
                    If i Mod DigSrc_DataWidth = DigSrc_DataWidth - 1 Then
                        SrcOut = SrcOut & ", "
                    ElseIf i Mod 4 = 3 Then
                        SrcOut = SrcOut & " "
                    End If

                Next i
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Source Data=" & SrcOut
            End If
        Next site

        If DigCap_Sample_Size <> 0 Then

           TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test Start ========")
           Call DigCapSetup(PattArray(0), DigCap_Pin, "Meas_cap", DigCap_Sample_Size, OutDspWave)

        End If

        SourceNum = SourceNum + 1

        Call TheHdw.Patterns(PattArray(0)).start

        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end

        PatCount = PatCount + 1
        CapOut = ""


        TheHdw.Digital.Patgen.HaltWait   '' Haltwait at patten end

        If SourceNum = 10 Then
            b_TestDone = True
        End If

    Loop

    For Each site In TheExec.sites.Active
        For k = 0 To DigCap_Sample_Size / DigCap_DataWidth - 1
        code(k) = 0
            For i = 0 To DigCap_DataWidth - 1
                code(k) = code(k) + OutDspWave(site).Element(k * 12 + i) * (2 ^ i)
            Next i
         If CurrentJobName_L Like "*char*" Then
         Disable_Inst_pinname_in_PTR
            TheExec.Flow.TestLimit resultVal:=code(k), Unit:=unitNone, ForceResults:=tlForceFlow
         Enable_Inst_pinname_in_PTR
         Else
            TheExec.Flow.TestLimit resultVal:=code(k), Unit:=unitNone, Tname:="Capture Data_" & k, ForceResults:=tlForceFlow
            End If
        Next k

'' Add for TMPS fusing 20160526
       If TheHdw.Digital.Patgen.PatternBurstPassed(site) = False Then 'Pattern Fail
           PassFlag_TMPS(site) = False
       Else
           PassFlag_TMPS(site) = True
        End If

        If LCase(TMPS_Fuse_string) Like "*t3_fuse_25c*" Then
            For X = 1 To 7
                Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp" & X, PassFlag_TMPS(site))
            Next X
            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp2", PassFlag_TMPS(site))
            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp3", PassFlag_TMPS(site))
            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp4", PassFlag_TMPS(site))
            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp5", PassFlag_TMPS(site))
            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp6", PassFlag_TMPS(site))
            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp7", PassFlag_TMPS(site))
            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp8", PassFlag_TMPS(site))

            For X = 0 To 7
                Call auto_eFuse_SetWriteDecimal("ECID", "Temp" & X + 1, Data_Array(X))
            Next X
            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp2", Data_Array(1))
            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp3", Data_Array(2))
            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp4", Data_Array(3))
            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp5", Data_Array(4))
            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp6", Data_Array(5))
            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp7", Data_Array(6))
            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp8", Data_Array(7))

        ElseIf LCase(TMPS_Fuse_string) Like "*t5_fuse_85c*" Then

            Call auto_eFuse_SetPatTestPass_Flag("CFG", "TS_TEMP_REF_CTRL0", True)
            Call auto_eFuse_SetPatTestPass_Flag("MON", "THERMAL_PARAM_TS_REFERENCE_CTRL", True)
            Call auto_eFuse_SetPatTestPass_Flag("CFG", "TS_TEMP_REF_CTRL1", True)
            Call auto_eFuse_SetPatTestPass_Flag("CFG", "TS_TEMP_REF_CTRL2", True)
            Call auto_eFuse_SetPatTestPass_Flag("CFG", "TS_TEMP_REF_CTRL3", True)
            Call auto_eFuse_SetPatTestPass_Flag("UDR", "Temp_sensor3_tTRIM", True)
            Call auto_eFuse_SetPatTestPass_Flag("UDR", "Temp_sensor2_tTRIM", True)
            Call auto_eFuse_SetPatTestPass_Flag("UDR", "Temp_sensor1_tTRIM", True)
            Call auto_eFuse_SetPatTestPass_Flag("UDR", "Temp_sensor0_tTRIM", True)

            If PassFlag_TMPS(site) = True Then
                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL0", Data_Array(0))
                Call auto_eFuse_SetWriteDecimal("MON", "THERMAL_PARAM_TS_REFERENCE_CTRL", Data_Array(0))
                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL1", Data_Array(1))
                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL2", Data_Array(2))
                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL3", Data_Array(3))
                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor3_tTRIM", Data_Array(4))
                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor2_tTRIM", Data_Array(5))
                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor1_tTRIM", Data_Array(6))
                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor0_tTRIM", Data_Array(7))

            Else  ' if fail then burn the default code

                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL0", 64)
                Call auto_eFuse_SetWriteDecimal("MON", "THERMAL_PARAM_TS_REFERENCE_CTRL", 64)
                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL1", 64)
                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL2", 64)
                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL3", 64)
                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor3_tTRIM", 64)
                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor2_tTRIM", 64)
                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor1_tTRIM", 64)
                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor0_tTRIM", 64)

            End If
        End If

    Next site

    Call HardIP_WriteFuncResult

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in TMPS function"
    If AbortTest Then Exit Function Else Resume Next

End Function






Public Function HIP_eFuse_Write(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, Flag_Name As String, Optional Efuse_Binary_Write_Flag As Boolean = False, _
                                Optional Calc_code As String) As Long

    ' Parameter : eFuse Block , eFuse Variable , data
    ' Call auto_eFuse_SetPatTestPass_Flag("CFG", "LPDP_C_RX", TheHdw.Digital.Patgen.PatternBurstPassed(Site))
    ' Call auto_eFuse_SetWriteDecimal("CFG", "LPDP_C_RX", BestCode(Site))

    Dim site As Variant
    Dim DSPWave_Dict As New DSPWave
    Dim Data_Temp As String
    Dim m_value As New SiteDouble
    Dim j As Integer
    Dim Pass_Fail_Flag As New SiteBoolean
    Dim Flag_Name_Split() As String: Flag_Name_Split = Split(Flag_Name, ",")
    Dim i As Long
    Dim fusetype_org As String
    Dim m_dlogstr As String
    On Error GoTo errHandler

    DSPWave_Dict = GetStoredCaptureData(Dict_Store_Code_Name)

        For Each site In TheExec.sites

                If Efuse_Binary_Write_Flag Then
                                For j = 0 To (DSPWave_Dict(site).SampleSize - 1)
                                        Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(j))
                                Next j
                                m_value(site) = Bin2Dec_rev(Data_Temp)
                                Data_Temp = ""
                Else
                        m_value(site) = DSPWave_Dict(site).Element(0)
                End If
        Next site
'''----------cal write fused code
    If Calc_code <> "" Then
    'Calc_code = "add,100"
        If Split(Calc_code, ",")(0) = "add" Then
            m_value = m_value.Add(Split(Calc_code, ",")(1))
        End If
    End If
'''----------cal write fused code
    If UBound(Flag_Name_Split) > 0 Then
        For Each site In TheExec.sites
            For i = 0 To UBound(Flag_Name_Split)
                If i = 0 Then
                    Pass_Fail_Flag = TheExec.Flow.SiteFlag(site, Flag_Name_Split(i))
                Else
                    Pass_Fail_Flag = Pass_Fail_Flag Or TheExec.Flow.SiteFlag(site, Flag_Name_Split(i))
                End If
            Next i
                If Pass_Fail_Flag = 1 Then
                    Pass_Fail_Flag(site) = False
                ElseIf Pass_Fail_Flag = 0 Then
                    Pass_Fail_Flag(site) = True
                Else
                    Pass_Fail_Flag(site) = False
                    TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
                End If
'            Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catename, Pass_Fail_Flag(site), True)
'            Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, m_value(site), True)
            If (True) Then
                fusetype_org = ""
                m_dlogstr = ""
                fusetype_org = FormatNumeric(FuseType, 4)
                m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + fusetype_org + FormatNumeric("Fuse SetWriteVariable_SiteAware", -35)
                m_dlogstr = m_dlogstr + FormatNumeric(m_catename, Len(fusetype_org)) + " = " + FormatNumeric(m_value, -10)
                TheExec.Datalog.WriteComment m_dlogstr
            End If
        Next site
    Else
        For Each site In TheExec.sites
            If TheExec.Flow.SiteFlag(site, Flag_Name) = 1 Then
                Pass_Fail_Flag(site) = False
            ElseIf TheExec.Flow.SiteFlag(site, Flag_Name) = 0 Then
                Pass_Fail_Flag(site) = True
            Else
                Pass_Fail_Flag(site) = False
                TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
            End If
'            Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catename, Pass_Fail_Flag(site), True)
'            Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, m_value(site), True)
            If (True) Then
                fusetype_org = ""
                m_dlogstr = ""
                fusetype_org = FormatNumeric(FuseType, 4)
                m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + fusetype_org + FormatNumeric("Fuse SetWriteVariable_SiteAware", -35)
                m_dlogstr = m_dlogstr + FormatNumeric(m_catename, Len(fusetype_org)) + " = " + FormatNumeric(m_value, -10)
                TheExec.Datalog.WriteComment m_dlogstr
            End If
        Next site
    End If
    
    Call auto_eFuse_SetPatTestPass_Flag_SiteAware(FuseType, m_catename, Pass_Fail_Flag, True)
    Call auto_eFuse_SetWriteVariable_SiteAware(FuseType, m_catename, m_value, False)
    
    ' Check implicit alarms
    TheHdw.Alarms.Check

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Write"
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function ADC_Trim(patset As Pattern, CPUA_Flag_In_Pat As Boolean, _
    MeasureV_PinS As String, _
    DigSrc_pin As PinList, DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, _
    DigSrc_Equation As String, DigSrc_Assignment As String, _
    Optional TargetValue_Volt As Double, Optional CUS_Str As String, Optional Validating_ As Boolean) As Long

'' Step 1 : trim code is 32 bit, show out measured volt and trimed code, target volt is 1.1v
'' Step 2 : start from 0x8 and add algorithm to decide +/- direction
'' while decimal < 2 ^ DigSrc_Sample_Size
'' convert decimal to binary reverse
'' input the binary reverse data to digSrc_assignment


    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen
    Dim X As Long
    Dim InDSPwave As New DSPWave
    Dim SrcOut As String
    Dim site As Variant
    Dim Pat As String
    Dim i As Integer
    Dim ShowDec As String
    Dim ShowOut As String
    Dim TrimBits As String
    Dim b_TestDone As Boolean
    Dim SourceNum As Integer
    Dim k As Integer
    Dim MeasureVoltage As New PinListData
    Dim Data As Integer
    Dim PassFlag_ADC As New SiteBoolean

    gl_TName_Pat = patset.Value

    On Error GoTo errHandler
      For Each site In TheExec.sites.Active
            Src_DSPWave.CreateConstant 0, DigSrc_Sample_Size
    Next site

    b_TestDone = False
    SourceNum = 0

    If DigSrc_Sample_Size = 0 Then
        TheExec.Datalog.WriteComment ("Error!! - Please check input argument DigSrc_Sample_Size")
        Exit Function
    End If

    TheHdw.Digital.Patgen.Halt
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Patterns(patset).Load

    Dim PattArray() As String
    Dim PatCount As Long

    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode

    Do While b_TestDone = False
        For Each site In TheExec.sites.Active

            ''  theexec.Datalog.WriteComment ("======== Start Dig Src setup =======")
            If SourceNum = 0 Then
                Call Create_DigSrc_Data(DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPwave, site)
            End If

            If SourceNum > 0 And SourceNum < 9 Then

                If MeasureVoltage.Pins(0).Value(site) > TargetValue_Volt Then
                    For X = 0 To 3
                        InDSPwave(site).Element(4 * (8 - SourceNum) + X) = 0
                    Next X
                    'InDSPWave(site).Element(4 * (8 - SourceNum) + 1) = 0
                    'InDSPWave(site).Element(4 * (8 - SourceNum) + 2) = 0
                    'InDSPWave(site).Element(4 * (8 - SourceNum) + 3) = 0
                End If

                If SourceNum < 8 Then
                    For X = 0 To 3
                        InDSPwave(site).Element(4 * (7 - SourceNum) + X) = 1
                    Next X
                    'InDSPWave(site).Element(4 * (7 - SourceNum) + 1) = 1
                    'InDSPWave(site).Element(4 * (7 - SourceNum) + 2) = 1
                    'InDSPWave(site).Element(4 * (7 - SourceNum) + 3) = 1
                End If

                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Measured Voltage: " & MeasureVoltage.Pins(0).Value(site)
            End If

        Next site

        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "Meas_Src", DigSrc_Sample_Size, InDSPwave)

        If SourceNum = 8 Then
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "==========SOURCING TRIMMED DATA BITS==============="
        End If

        For Each site In TheExec.sites.Active
            If SourceNum > 0 Then
                SrcOut = ""
                For i = 0 To DigSrc_Sample_Size - 1
                    SrcOut = SrcOut & InDSPwave(site).Element(i)
                    If i Mod DigSrc_DataWidth = DigSrc_DataWidth - 1 Then
                        SrcOut = SrcOut & ", "
                    ElseIf i Mod 4 = 3 Then
                        SrcOut = SrcOut & " "
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Source Data=" & SrcOut
            End If
        Next site

        Call TheHdw.Patterns(PattArray(0)).start

        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)    '' Meas during CPUA loop
        Else
            Call TheHdw.Digital.Patgen.HaltWait '' Pattern without CPUA loop
        End If

''        Call HardIP_SetupAndMeasureVolt_UVI80(MeasureV_PinS, MeasureVoltage, True)
        ''20170621
        Dim MV_TestCond_UVI80(0) As DUTConditions
        MV_TestCond_UVI80(0).PinName = MeasureV_PinS
        Call HardIP_SetupAndMeasureVolt_UVI80_old(MV_TestCond_UVI80, MeasureVoltage)

        SourceNum = SourceNum + 1

        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)    '' Jump out CPUA loop
        End If

        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end

        PatCount = PatCount + 1

        If SourceNum = 9 Then
            b_TestDone = True
        End If
    Loop

    TheExec.Flow.TestLimit resultVal:=MeasureVoltage, Unit:=unitVolt, Tname:="Volt_meas_ADC_Trim", ForceResults:=tlForceNone

    Call HardIP_WriteFuncResult

    For Each site In TheExec.sites.Active
        For i = 0 To DigSrc_Sample_Size - 1
            Src_DSPWave(site).Element(i) = InDSPwave(site).Element(i)
        Next i
    Next site

    For Each site In TheExec.sites.Active
        SrcOut = ""
        For i = 0 To DigSrc_Sample_Size - 1
            SrcOut = SrcOut & Src_DSPWave(site).Element(i)
        Next i
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Stored Data=" & SrcOut

        '''''''''''''''''''''eFUSE
        Data = 0
        For i = 0 To DigSrc_DataWidth / 4 - 1
            Data = Data + InDSPwave(site).Element(4 * i) * (2 ^ i)
        Next i
        If CurrentJobName_U Like "*FT*" Then
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ""
            Data = auto_eFuse_GetReadDecimal("UDR", "ADC_vTRIM", True)
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ""
            For i = 0 To DigSrc_Sample_Size / 4 - 1
                For X = 0 To 3
                    Src_DSPWave(site).Element(4 * i + X) = Data Mod 2
                Next X
                'Src_DSPWave(site).Element(4 * i + 1) = Data Mod 2
                'Src_DSPWave(site).Element(4 * i + 2) = Data Mod 2
                'Src_DSPWave(site).Element(4 * i + 3) = Data Mod 2
                Data = Data \ 2
            Next i
        Else
            If TheHdw.Digital.Patgen.PatternBurstPassed(site) = False Then 'Pattern Fail
                PassFlag_ADC(site) = False
            Else
                PassFlag_ADC(site) = True
            End If
            If CUS_Str = "ADC_VTRIM" Then
                Call auto_eFuse_SetPatTestPass_Flag("UDR", "ADC_vTRIM", PassFlag_ADC(site))
                Call auto_eFuse_SetWriteDecimal("UDR", "ADC_vTRIM", Data)
            End If
        End If

    Next site

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in ADC_Trim function"
    If AbortTest Then Exit Function Else Resume Next

End Function
    




Public Function MTR_Sense_Calibration_Coeff_Verification(SensorArray As String, Temperature As String, FusedCoeffDicName_1 As String, FusedCoeffDicName_2 As String, _
Optional MTRMatricesSheet As String, Optional SensorCalculate As String, Optional SweepVArryDic As String, Optional Validating_ As Boolean) As Long
'MTR Record
If Validating_ Then
    
    Exit Function    ' Exit after validation
End If
    
    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If
    
    HIP_Init_Datalog_Setup


    Dim site As Variant
    Dim PowerPinsGroup() As String
    Dim PowerPinsCounter As Long
    Dim PowerPinLevelsGroup() As String
    Dim PowerPinLevelsCounter As Long
    Dim SensorsGroup() As String
    Dim SensorsCounter As Long
    Dim ScanSeson() As String
    Dim ScanSesonCounter As Long
    Dim MatrixStr As String
    Dim VoltageScaner As New SiteLong
    Dim VoltageScanerStr As String
    Dim FullGroup() As String
    Dim FullLevels() As String
    
    Dim SetInformation() As New DSPWave
    Dim Aininformation() As New DSPWave
    Dim Aixinformation() As New DSPWave
    Dim PiUInformation() As New DSPWave
    
    Dim Fused_ROT_Decimal_Vector As New DSPWave
    Dim Fused_ROV_Decimal_Vector As New DSPWave
    Fused_ROT_Decimal_Vector.CreateConstant 0, 4, DspDouble
    Fused_ROV_Decimal_Vector.CreateConstant 0, 3, DspDouble
            
    Dim Output_ROT_Freq_Vector As New DSPWave
    Dim Output_ROV_Freq_Vector As New DSPWave
    Output_ROT_Freq_Vector.CreateConstant 0, 8, DspDouble
    Output_ROV_Freq_Vector.CreateConstant 0, 8, DspDouble
        
    Dim Difference_ROT_Freq_Vector As New DSPWave
    Dim Difference_ROV_Freq_Vector As New DSPWave
    Difference_ROT_Freq_Vector.CreateConstant 0, 8, DspDouble
    Difference_ROV_Freq_Vector.CreateConstant 0, 8, DspDouble
    
    
    FullGroup = Split(SweepVArryDic, ";")
    PowerPinsGroup = Split(SensorCalculate, ";")
    For PowerPinsCounter = 0 To UBound(PowerPinsGroup)
        FullLevels = Split(FullGroup(PowerPinsCounter), ",")
        PowerPinLevelsGroup = Split(PowerPinsGroup(PowerPinsCounter), ",")
        ReDim SetInformation(UBound(PowerPinLevelsGroup))
        ReDim Aininformation(UBound(PowerPinLevelsGroup))
        ReDim Aixinformation(UBound(PowerPinLevelsGroup))
        ReDim PiUInformation(UBound(PowerPinLevelsGroup))
        For PowerPinLevelsCounter = 0 To UBound(PowerPinLevelsGroup)
            SensorsGroup = Split(SensorArray, ";")
            For SensorsCounter = 0 To UBound(SensorsGroup)
                ScanSeson = Split(SensorsGroup(SensorsCounter), ",")
                For ScanSesonCounter = 0 To UBound(ScanSeson)
                    If CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_GPU" Then
                        If ScanSeson(ScanSesonCounter) Like "GPU*" Then
                            VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_1150_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_1150_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_1150_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_1150_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_1150_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    ElseIf CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_ECPU" Then
                        If ScanSeson(ScanSesonCounter) Like "ECPU*" Then
                            VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_1150_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_1150_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_1150_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_1150_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_1150_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    ElseIf CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_PCPU" Then
                        If ScanSeson(ScanSesonCounter) Like "ANE*" Or ScanSeson(ScanSesonCounter) Like "PCPU*" Then
                        VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_1150_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_1150_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_1150_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_1150_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_1150_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    ElseIf CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_AVE" Then
                        If ScanSeson(ScanSesonCounter) Like "AVE*" Then
                            VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_950_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_950_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_950_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_950_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_950_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    ElseIf CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_SOC" Then
                        If ScanSeson(ScanSesonCounter) Like "SOC*" Then
                            VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_950_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_950_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_950_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_950_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_950_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    End If
                Next ScanSesonCounter
            Next SensorsCounter
        Next PowerPinLevelsCounter
    Next PowerPinsCounter
'    thehdw.DSP.ExecutionMode = tlDSPModeHostDebug
End Function



Public Function Metrology_CAL_eFuse_Write(FuseType As String, SensorArray As String, m_catename As String, Dict_Store_Code_Name As String, Flag_Name As String, Hex_BitSize As String, Optional Temperature As String, _
                                          Optional Calculate_Group As String, Optional Efuse_Hex_Write_Flag As Boolean = True) As Long

    ' Parameter : eFuse Block , eFuse Variable , data
    ' Call auto_eFuse_SetPatTestPass_Flag("CFG", "LPDP_C_RX", TheHdw.Digital.Patgen.PatternBurstPassed(Site))
    ' Call auto_eFuse_SetWriteDecimal("CFG", "LPDP_C_RX", BestCode(Site))

    Dim site As Variant
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_DictTemp As DSPWave
    Dim Data_Temp As String
    Dim m_value As New SiteVariant
    Dim i, j, k As Integer
    Dim Pass_Fail_Flag As New SiteBoolean
    On Error GoTo errHandler

    Dim SizeCounter As Long
    Dim SizeCounterTemp As Long
    Dim CalculateSize As String
    Dim CalculateArray() As String
    Dim CalculateSplit() As String
    Dim m_catenameTemp() As String
    Dim m_catenameCombination() As String
    Dim SensorArrayTemp() As String
    Dim Dict_Store_Code_NameTemp As String

    CalculateArray = Split(Calculate_Group, ",")
    m_catenameTemp = Split(m_catename, ",")
    m_catenameCombination = m_catenameTemp
    SensorArrayTemp = Split(SensorArray, ",")

    For i = 0 To UBound(SensorArrayTemp)
        SizeCounter = 1
        Dict_Store_Code_NameTemp = Dict_Store_Code_Name + "_" + SensorArrayTemp(i) + "_" + Temperature + "c"
        DSPWave_Dict = GetStoredCaptureData(Dict_Store_Code_NameTemp)

        For j = 0 To UBound(CalculateArray)
            CalculateSplit = Split(CalculateArray(j), ":")
            CalculateSize = CLng(CalculateSplit(1))

            For Each site In TheExec.sites
                If Efuse_Hex_Write_Flag Then
                    For k = 0 To (DSPWave_Dict(site).SampleSize - 1)
                        Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(k))
                    Next k
                    Data_Temp = StrReverse(Data_Temp)
                    Data_Temp = Mid(Data_Temp, SizeCounter, CalculateSize)
                    m_value(site) = "0x" + Calc_MTR_BinStr2HexStr(Data_Temp, CLng(Hex_BitSize))
                    Data_Temp = ""
                Else
                    m_value(site) = DSPWave_Dict(site).Element(0)
                End If
            Next site
            SizeCounter = CalculateSize + SizeCounter

            m_catenameCombination(i) = m_catenameTemp(i) + "_" + "t" + Temperature + "_" + CalculateSplit(0)
            For Each site In TheExec.sites
                If TheExec.Flow.SiteFlag(site, Flag_Name) = 1 Then
                    Pass_Fail_Flag(site) = False
                    'm_value(site) = 0   'Cebu MTRGSNS fuse 0 when MTRGSNS fail
                ElseIf TheExec.Flow.SiteFlag(site, Flag_Name) = 0 Then
                    Pass_Fail_Flag(site) = True
                Else
                    Pass_Fail_Flag(site) = False
                    TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
                End If
                    Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catenameCombination(i), Pass_Fail_Flag(site), True)
                    Call auto_eFuse_SetWriteDecimal(FuseType, m_catenameCombination(i), m_value(site), True)
            Next site

''''''''''            For Each site In TheExec.sites
''''''''''                If Efuse_Hex_Write_Flag Then
''''''''''                    For k = 0 To (DSPWave_Dict(site).SampleSize - 1)
''''''''''                        Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(k))
''''''''''                    Next k
''''''''''                    Data_Temp = StrReverse(Data_Temp)
''''''''''                    m_value(site) = "0x" + Calc_MTR_BinStr2HexStr(Data_Temp, CLng(Hex_BitSize))
''''''''''                    Data_Temp = ""
''''''''''                Else
''''''''''                    m_value(site) = DSPWave_Dict(site).Element(0)
''''''''''                End If
''''''''''            Next site
''''''''''            For Each site In TheExec.sites
''''''''''                If TheExec.Flow.SiteFlag(site, flag_name) = 1 Then
''''''''''                    Pass_Fail_Flag(site) = False
''''''''''                ElseIf TheExec.Flow.SiteFlag(site, flag_name) = 0 Then
''''''''''                    Pass_Fail_Flag(site) = True
''''''''''                Else
''''''''''                    Pass_Fail_Flag(site) = False
''''''''''                    TheExec.DataLog.WriteComment ("Error! " & flag_name & "(" & site & ")" & " status is Clear !")
''''''''''                End If
''''''''''                    Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catenameTemp(i), Pass_Fail_Flag(site), True)
''''''''''                    Call auto_eFuse_SetWriteDecimal(FuseType, m_catenameTemp(i), m_value(site), True)
''''''''''            Next site

        Next j
    Next i
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Write"
    If AbortTest Then Exit Function Else Resume Next

End Function

'''''
'''''Public Function pll_read() As Long
'''''
'''''    Call HIP_eFuse_Read(A, b, c)
'''''    Call HIP_eFuse_Read
'''''
'''''End Function

Public Function HIP_eFuse_Write_by_MTRGSNS(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, Flag_Name As String, Optional Efuse_Binary_Write_Flag As Boolean = False, _
                                Optional Calc_code As String) As Long

    ' Parameter : eFuse Block , eFuse Variable , data
    ' Call auto_eFuse_SetPatTestPass_Flag("CFG", "LPDP_C_RX", TheHdw.Digital.Patgen.PatternBurstPassed(Site))
    ' Call auto_eFuse_SetWriteDecimal("CFG", "LPDP_C_RX", BestCode(Site))

    Dim site As Variant
    Dim DSPWave_Dict As New DSPWave
    Dim Data_Temp As String
    Dim m_value As New SiteDouble
    Dim j As Integer
    Dim Pass_Fail_Flag As New SiteBoolean
    On Error GoTo errHandler

    DSPWave_Dict = GetStoredCaptureData(Dict_Store_Code_Name)

        For Each site In TheExec.sites

                If Efuse_Binary_Write_Flag Then
                                For j = 0 To (DSPWave_Dict(site).SampleSize - 1)
                                        Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(j))
                                Next j
                                m_value(site) = Bin2Dec_rev(Data_Temp)
                                Data_Temp = ""
                Else
                        m_value(site) = DSPWave_Dict(site).Element(0)
                End If
        Next site
'''----------cal write fused code
    If Calc_code <> "" Then
    'Calc_code = "add,100"
        If Split(Calc_code, ",")(0) = "add" Then
            m_value = m_value.Add(Split(Calc_code, ",")(1))
        End If
    End If
'''----------cal write fused code
    For Each site In TheExec.sites
        If TheExec.Flow.SiteFlag(site, Flag_Name) = 1 Then
            Pass_Fail_Flag(site) = True
            m_value(site) = 0   'Cebu MTRGSNS fuse 0 when MTRGSNS fail
        ElseIf TheExec.Flow.SiteFlag(site, Flag_Name) = 0 Then
            Pass_Fail_Flag(site) = True
        Else
            Pass_Fail_Flag(site) = False
            TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
        End If
        Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catename, Pass_Fail_Flag(site), True)
        Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, m_value(site), True)
    Next site

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Write"
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Check_MTRGSNS_25C_Fuse()
    On Error GoTo errHandler
    Dim site As Variant
    For Each site In TheExec.sites
        If CFGFuse.Category(CFGIndex("mtr_sense_vt_ts3i_t25_a2_3")).Read.Decimal(site) <> 0 Then
            TheExec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Check_MTRGSNS_25C_Fuse"
        Else
            TheExec.Flow.TestLimit 0, lowVal:=1, hiVal:=1, Tname:="Check_MTRGSNS_25C_Fuse"
        End If
    Next site

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in Check_MTRGSNS_25C_Fuse"
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function MTR_REL_Fuse_Calc_Verification(Optional Adc_Offset As String, Optional Adc_Gain As String, Optional Cap_Code_Temp1 As String, Optional Encoded_Temp1 As String, Optional Cap_Code_Temp2 As String, Optional Temp2 As String, Optional Fuse_Calc_Temp1 As String, Optional Fuse_Calc_Temp2 As String, Optional Store_Coeff_c0 As String, Optional c0_size As String, Optional Store_Coeff_c1 As String, Optional c1_size As String, Optional Store_Coeff_c2 As String, Optional c2_size As String, Optional Store_Coeff_c3 As String, Optional c3_size As String, Optional Validating_ As Boolean) As Long

    
    Dim site As Variant

    Dim Dict_name_tfe_vol_x1 As String
    Dim cal_tfe_vol_y1 As String

    Dim fuse_read_tfe_vol_0 As String
    Dim fuse_read_tfe_vol_1 As String
    Dim fuse_read_tfe_x0 As String
    Dim fuse_read_tfe_y0 As String

    Dim fuse_write_tfe_temp_0 As String
    Dim fuse_write_tfe_temp_1 As String

    Dim tfe_y0_decimal As String
    Dim tfe_y1_decimal As String


    Dim actual_Temp_CP2 As Double


Dim Store_C0_Coeff_Dic As String
 Dim Store_C1_Coeff_Dic As String
 Dim Store_C2_Coeff_Dic As String
 Dim Store_C3_Coeff_Dic As String
 
 Dim C0_Reg_Size As Long
 Dim C1_Reg_Size As Long
 Dim C2_Reg_Size As Long
 Dim C3_Reg_Size As Long
 
 
 
    
    fuse_read_tfe_vol_0 = Adc_Offset
    fuse_read_tfe_vol_1 = Adc_Gain

    fuse_read_tfe_x0 = Cap_Code_Temp1
    Dict_name_tfe_vol_x1 = Cap_Code_Temp2

    fuse_read_tfe_y0 = Encoded_Temp1
    cal_tfe_vol_y1 = Temp2

    fuse_write_tfe_temp_0 = Fuse_Calc_Temp1
    fuse_write_tfe_temp_1 = Fuse_Calc_Temp2

    Store_C0_Coeff_Dic = Store_Coeff_c0
Store_C1_Coeff_Dic = Store_Coeff_c1
Store_C2_Coeff_Dic = Store_Coeff_c2
Store_C3_Coeff_Dic = Store_Coeff_c3

 C0_Reg_Size = CLng(c0_size)
 C1_Reg_Size = CLng(c1_size)
 C2_Reg_Size = CLng(c2_size)
 C3_Reg_Size = CLng(c3_size)


    'Get Cap data for t5p2 at 85C and Fuse Data for offset,gain and x0 at 25C
    Dim DSP_tfe_vol_x1_binary As New DSPWave
    Dim DSP_fuse_read_tfe_vol_0_2S_binary As New DSPWave
    Dim DSP_fuse_read_tfe_vol_1_binary As New DSPWave
    Dim DSP_fuse_read_tfe_x0_binary As New DSPWave

    Dim Dsp_tfe_temp0_in_binary As New DSPWave
    Dim Dsp_tfe_temp1_in_binary As New DSPWave




    DSP_tfe_vol_x1_binary = GetStoredCaptureData(Dict_name_tfe_vol_x1)
    DSP_fuse_read_tfe_vol_0_2S_binary = GetStoredCaptureData(fuse_read_tfe_vol_0)
    DSP_fuse_read_tfe_vol_1_binary = GetStoredCaptureData(fuse_read_tfe_vol_1)
    DSP_fuse_read_tfe_x0_binary = GetStoredCaptureData(fuse_read_tfe_x0)
    Dsp_tfe_temp0_in_binary = GetStoredCaptureData(fuse_write_tfe_temp_0)
    Dsp_tfe_temp1_in_binary = GetStoredCaptureData(fuse_write_tfe_temp_1)




    ' y0 in decimal for 25C
    Dim DSP_fuse_read_tfe_y0_in_double As New DSPWave
    Dim decoded_Dic_tfe_y0_in_double As String
    decoded_Dic_tfe_y0_in_double = "decoded_Dic_tfe_y0_in_double"

    Dim call_decode_argv(2) As String
    call_decode_argv(0) = fuse_read_tfe_y0
    call_decode_argv(1) = decoded_Dic_tfe_y0_in_double
    Dim call_decodeActualTemp As Long
    call_decodeActualTemp = Calc_Metrology_DecodeActualTemp(1, call_decode_argv)

    DSP_fuse_read_tfe_y0_in_double = GetStoredCaptureData(decoded_Dic_tfe_y0_in_double)



    ' y1 in decimal for 85C .. for now..will be changed in future
  '  If cal_tfe_vol_y1 Like "CP2" Then

        actual_Temp_CP2 = CDbl(cal_tfe_vol_y1)

  '  End If

    Dim DSP_tfe_y1_in_double As New DSPWave

    DSP_tfe_y1_in_double.CreateConstant 0, 1, DspDouble

    For Each site In TheExec.sites

    DSP_tfe_y1_in_double(site).Element(0) = actual_Temp_CP2

    Next site

    'Start the algo


    'Define Constants

    Dim A0 As Double
    Dim a1 As Double
    Dim a2 As Double
    Dim a3 As Double

    'Values for Constants

    A0 = CDbl("-21.5822184999726")
    a1 = CDbl("428.0092266096283") 'truncated one digit
    a2 = CDbl("-133.4543109228228") 'truncated one digit
    a3 = CDbl("19.0485545665615")




    'Convert x1 to decimal

    Dim DSP_tfe_vol_x1_in_decimal As New DSPWave

    Call rundsp.BinToDec(DSP_tfe_vol_x1_binary, DSP_tfe_vol_x1_in_decimal)



    'Convert x0 to decimal

    Dim DSP_fuse_read_tfe_x0_in_decimal As New DSPWave

    Call rundsp.BinToDec(DSP_fuse_read_tfe_x0_binary, DSP_fuse_read_tfe_x0_in_decimal)



    'Convert vol_0 2S to decimal
     Dim DSP_fuse_read_tfe_vol_0_in_decimal As New DSPWave
     Dim SL_BitWidth As New SiteLong
     For Each site In TheExec.sites
            SL_BitWidth(site) = 18

    Next site

    DSP_fuse_read_tfe_vol_0_in_decimal.CreateConstant 0, 1, DspLong



    Call rundsp.DSP_2S_Complement_To_SignDec(DSP_fuse_read_tfe_vol_0_2S_binary, SL_BitWidth, DSP_fuse_read_tfe_vol_0_in_decimal)



    'Convert vol_1 to Decimal
    Dim DSP_fuse_read_tfe_vol_1_in_decimal As New DSPWave

    Call rundsp.BinToDec(DSP_fuse_read_tfe_vol_1_binary, DSP_fuse_read_tfe_vol_1_in_decimal)




    Dim Dsp_tfe_temp0_in_decimal As New DSPWave
    Dim Dsp_tfe_temp1_in_decimal As New DSPWave

    Dsp_tfe_temp0_in_decimal.CreateConstant 0, 1, DspDouble
     Dsp_tfe_temp1_in_decimal.CreateConstant 0, 1, DspDouble

     Set SL_BitWidth = Nothing

     For Each site In TheExec.sites
            SL_BitWidth(site) = 28

    Next site

     Call rundsp.DSP_2S_Complement_To_SignDec(Dsp_tfe_temp0_in_binary, SL_BitWidth, Dsp_tfe_temp0_in_decimal)
    Set SL_BitWidth = Nothing

     For Each site In TheExec.sites
            SL_BitWidth(site) = 28

    Next site
 Call rundsp.DSP_2S_Complement_To_SignDec(Dsp_tfe_temp1_in_binary, SL_BitWidth, Dsp_tfe_temp1_in_decimal)


    Dim detA0 As New SiteDouble
    Dim detA1 As New SiteDouble
    Dim Offset As New SiteDouble
    Dim Gain As New SiteDouble

    Dim C0 As New SiteDouble
    Dim c1 As New SiteDouble
    Dim C2 As New SiteDouble
    Dim C3 As New SiteDouble

    Dim C0_coeff_Val As New SiteDouble
    Dim C1_coeff_Val As New SiteDouble
    Dim C2_coeff_Val As New SiteDouble
    Dim C3_coeff_Val As New SiteDouble

    Dim Coeff_C0_in_decimal As New DSPWave
    Dim Coeff_C1_in_decimal As New DSPWave
    Dim Coeff_C2_in_decimal As New DSPWave
    Dim Coeff_C3_in_decimal As New DSPWave
    
     Coeff_C0_in_decimal.CreateConstant 0, 1, DspDouble
     Coeff_C1_in_decimal.CreateConstant 0, 1, DspDouble
    Coeff_C2_in_decimal.CreateConstant 0, 1, DspDouble
     Coeff_C3_in_decimal.CreateConstant 0, 1, DspDouble
    

    Dim Coeff_C0_in_binary As New DSPWave
    Dim Coeff_C1_in_binary As New DSPWave
    Dim Coeff_C2_in_binary As New DSPWave
    Dim Coeff_C3_in_binary As New DSPWave


    For Each site In TheExec.sites

                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0), Tname:="tfe_vol_0", ForceResults:=tlForceNone

                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0), Tname:="tfe_vol_1", ForceResults:=tlForceNone


                TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp0_in_decimal(site).Element(0), Tname:="temp_0", ForceResults:=tlForceNone

                TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp1_in_decimal(site).Element(0), Tname:="temp_1", ForceResults:=tlForceNone



                If (DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0) = 0) Then

                         C0_coeff_Val(site) = 178956970

                        C1_coeff_Val(site) = 178956970


                         C2_coeff_Val(site) = 178956970

                        C3_coeff_Val(site) = 178956970


                            Coeff_C0_in_decimal(site).Element(0) = C0_coeff_Val(site)

                        Coeff_C1_in_decimal(site).Element(0) = C1_coeff_Val(site)

                        Coeff_C2_in_decimal(site).Element(0) = C2_coeff_Val(site)

                        Coeff_C3_in_decimal(site).Element(0) = C3_coeff_Val(site)


                        TheExec.Flow.TestLimit resultVal:=Coeff_C0_in_decimal(site).Element(0), Tname:="Error_coeff_c0", ForceResults:=tlForceNone

                    TheExec.Flow.TestLimit resultVal:=Coeff_C1_in_decimal(site).Element(0), Tname:="Error_coeff_c1", ForceResults:=tlForceNone

                    TheExec.Flow.TestLimit resultVal:=Coeff_C2_in_decimal(site).Element(0), Tname:="Error_coeff_c2", ForceResults:=tlForceNone

                    TheExec.Flow.TestLimit resultVal:=Coeff_C3_in_decimal(site).Element(0), Tname:="Error_coeff_c3", ForceResults:=tlForceNone

                Else




                    detA0(site) = Dsp_tfe_temp0_in_decimal(site).Element(0) / (2 ^ 13)
                    detA1(site) = Dsp_tfe_temp1_in_decimal(site).Element(0) / (2 ^ 13)
                    Offset(site) = DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0) / (2 ^ 13)
                    Gain(site) = DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0) / (2 ^ 17)

                    C0(site) = A0 + detA0(site) - (a1 + detA1(site)) * (Offset(site) / Gain(site))
                    c1(site) = ((a1 + detA1(site)) / Gain(site)) - 2 * a2 * (Offset(site) / (Gain(site) ^ 2))
                    C2(site) = (a2 / (Gain(site) ^ 2)) - 3 * a3 * (Offset(site) / (Gain(site) ^ 3))
                    C3(site) = a3 / (Gain(site) ^ 3)


                   TheExec.Flow.TestLimit resultVal:=detA0(site), Tname:="detA0", ForceResults:=tlForceNone

                   TheExec.Flow.TestLimit resultVal:=detA1(site), Tname:="detA1", ForceResults:=tlForceNone



                   TheExec.Flow.TestLimit resultVal:=Offset(site), Tname:="Offset", ForceResults:=tlForceNone

                   TheExec.Flow.TestLimit resultVal:=Gain(site), Tname:="Gain", ForceResults:=tlForceNone

                    TheExec.Flow.TestLimit resultVal:=C0(site), Tname:="C0", ForceResults:=tlForceNone

                   TheExec.Flow.TestLimit resultVal:=c1(site), Tname:="C1", ForceResults:=tlForceNone



                   TheExec.Flow.TestLimit resultVal:=C2(site), Tname:="C2", ForceResults:=tlForceNone

                   TheExec.Flow.TestLimit resultVal:=C3(site), Tname:="C3", ForceResults:=tlForceNone

            C0_coeff_Val(site) = FormatNumber(C0(site) * 2 ^ 13)
            C1_coeff_Val(site) = FormatNumber(c1(site) * 2 ^ 13)
            C2_coeff_Val(site) = FormatNumber(C2(site) * 2 ^ 13)
            C3_coeff_Val(site) = FormatNumber(C3(site) * 2 ^ 13)








                    If (C0_coeff_Val(site) > 42949672945#) Then


                         TheExec.Flow.TestLimit resultVal:=C0_coeff_Val(site), Tname:="UpperLimit_Reached_c0_coeff", ForceResults:=tlForceNone


                        C0_coeff_Val(site) = 4294967295#



                    End If
                     If (C1_coeff_Val(site) > 4294967295#) Then


                         TheExec.Flow.TestLimit resultVal:=C1_coeff_Val(site), Tname:="UpperLimit_Reached_c1_coeff", ForceResults:=tlForceNone


                        C1_coeff_Val(site) = 4294967295#



                    End If
                     If (C2_coeff_Val(site) > 42949672945#) Then


                         TheExec.Flow.TestLimit resultVal:=C2_coeff_Val(site), Tname:="UpperLimit_Reached_c2_coeff", ForceResults:=tlForceNone


                        C2_coeff_Val(site) = 42949672945#



                    End If
                     If (C3_coeff_Val(site) > 42949672945#) Then


                         TheExec.Flow.TestLimit resultVal:=C3_coeff_Val(site), Tname:="UpperLimit_Reached_c3_coeff", ForceResults:=tlForceNone


                        C3_coeff_Val(site) = 42949672945#



                    End If

                Coeff_C0_in_decimal(site).Element(0) = C0_coeff_Val(site)

                Coeff_C1_in_decimal(site).Element(0) = C1_coeff_Val(site)
                Coeff_C2_in_decimal(site).Element(0) = C2_coeff_Val(site)

                Coeff_C3_in_decimal(site).Element(0) = C3_coeff_Val(site)



                    TheExec.Flow.TestLimit resultVal:=Coeff_C0_in_decimal(site).Element(0), Tname:="C0_coeff_Val", ForceResults:=tlForceNone

                   TheExec.Flow.TestLimit resultVal:=Coeff_C1_in_decimal(site).Element(0), Tname:="C1_coeff_Val", ForceResults:=tlForceNone



                   TheExec.Flow.TestLimit resultVal:=Coeff_C2_in_decimal(site).Element(0), Tname:="C2_coeff_Val", ForceResults:=tlForceNone

                   TheExec.Flow.TestLimit resultVal:=Coeff_C3_in_decimal(site).Element(0), Tname:="C3_coeff_Val", ForceResults:=tlForceNone



            End If

    Next site





    Call rundsp.DSPWf_Dec2Binary(Coeff_C0_in_decimal, C0_Reg_Size, Coeff_C0_in_binary)

    Call rundsp.DSPWf_Dec2Binary(Coeff_C1_in_decimal, C1_Reg_Size, Coeff_C1_in_binary)

        Call rundsp.DSPWf_Dec2Binary(Coeff_C2_in_decimal, C2_Reg_Size, Coeff_C2_in_binary)

    Call rundsp.DSPWf_Dec2Binary(Coeff_C3_in_decimal, C3_Reg_Size, Coeff_C3_in_binary)

    'Algo end


Store_C0_Coeff_Dic = Store_Coeff_c0
Store_C1_Coeff_Dic = Store_Coeff_c1
Store_C2_Coeff_Dic = Store_Coeff_c2
Store_C3_Coeff_Dic = Store_Coeff_c3


    Call AddStoredCaptureData(Store_C0_Coeff_Dic, Coeff_C0_in_binary)
    Call AddStoredCaptureData(Store_C1_Coeff_Dic, Coeff_C1_in_binary)

    Call AddStoredCaptureData(Store_C2_Coeff_Dic, Coeff_C2_in_binary)
    Call AddStoredCaptureData(Store_C3_Coeff_Dic, Coeff_C3_in_binary)


End Function

Public Function MTR_Sense_Alignment_Calc(Optional FreqSensorArray As String, Optional SensorArray As String, Optional SweepVArray As String, Optional Temperature As String) As Long
                                                                                                                                                                                                                                                             
    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If

Dim i, j, k As Long
Dim site As Variant
Dim SweepTempName As String
Dim SensorTempName As String
Dim SweepOriginTempName As String
Dim SensorOriginTempName As String


Dim SplitSensorArray() As String
Dim SplitSweepVArray() As String
Dim SplitSensorGroup() As String
Dim SplitSweepVGroup() As String
Dim SplitOriginSensorArray() As String
Dim SplitOriginSensorGroup() As String
SplitSensorGroup = Split(FreqSensorArray, ";")
SplitSweepVGroup = Split(SweepVArray, ";")
SplitOriginSensorGroup = Split(SensorArray, ";")
Dim TempDSPWave As New DSPWave
Dim RotRovMatrix As New DSPWave
Dim RotRovOriginMatrix As New DSPWave
Dim DicCounterbyGroup As Long
Dim DicCounterbyGroup_temp As Long
RotRovMatrix.CreateConstant 0, 16, DspDouble
RotRovOriginMatrix.CreateConstant 0, 16, DspDouble
DicCounterbyGroup = -1
For i = 0 To UBound(SplitSensorGroup)
    SplitSensorArray = Split(SplitSensorGroup(i), ",")
    For j = 0 To UBound(SplitSensorArray)
        DicCounterbyGroup = DicCounterbyGroup + 1
    Next j
Next i

DicCounterbyGroup_temp = 0
For i = 0 To UBound(SplitSensorGroup)
    SplitSensorArray = Split(SplitSensorGroup(i), ",")
    SplitSweepVArray = Split(SplitSweepVGroup(i), ",")
    SplitOriginSensorArray = Split(SplitOriginSensorGroup(i), ",")
    
    For j = 0 To UBound(SplitSensorArray)
        Dim RotRovMatrix_temp() As New DSPWave
        ReDim RotRovMatrix_temp(DicCounterbyGroup)
        Dim RotRovOriginMatrix_temp() As New DSPWave
        ReDim RotRovOriginMatrix_temp(DicCounterbyGroup)
        
        For Each site In TheExec.sites
            RotRovMatrix(site).CreateConstant 0, UBound(SplitSweepVArray) + 1, DspDouble
            RotRovOriginMatrix(site).CreateConstant 0, UBound(SplitSweepVArray) + 1, DspDouble
        Next site
        
        SensorTempName = SplitSensorArray(j) + "_" + Temperature + "C"
        SensorOriginTempName = SplitOriginSensorArray(j) + "_" + Temperature + "C"
        
        For k = 0 To UBound(SplitSweepVArray)
            
            
            Set TempDSPWave = Nothing
            Dim DPSWaveConvert As New SiteDouble
            SweepTempName = SplitSensorArray(j) + "_" + SplitSweepVArray(k)
            SweepOriginTempName = SplitOriginSensorArray(j) + "_" + SplitSweepVArray(k)
            
            TempDSPWave = GetStoredCaptureData(SweepTempName)
            DPSWaveConvert = GetStoredData(SweepOriginTempName & "_para")
            
            For Each site In TheExec.sites
                RotRovMatrix(site).Element(k) = TempDSPWave(site).Element(0)
                RotRovOriginMatrix(site).Element(k) = CDbl(DPSWaveConvert)
            Next site
            
        Next k
    
        RotRovMatrix_temp(DicCounterbyGroup_temp) = RotRovMatrix
        RotRovOriginMatrix_temp(DicCounterbyGroup_temp) = RotRovOriginMatrix
        Call AddStoredCaptureData(SensorTempName, RotRovMatrix_temp(DicCounterbyGroup_temp))
        Call AddStoredCaptureData(SensorOriginTempName, RotRovOriginMatrix_temp(DicCounterbyGroup_temp))
        DicCounterbyGroup_temp = DicCounterbyGroup_temp + 1
        
    Next j
Next i
End Function


Public Function MTR_Sense_Calibration_Coeff_Calc(Optional FreqSensorAlignment As String, Optional SensorAlignment As String, _
Optional SweepVArrayDic As String, Optional SweepVArrayValue As String, Optional Temperature As String, Optional FuseSize_1 As String, _
Optional StoreFuseDicName_1 As String, Optional FuseSize_2 As String, Optional StoreFuseDicName_2 As String, Optional SensorCalculate As String, _
Optional MTR_CAL_Sheet As String, Optional Validating_ As Boolean) As Long '
'MTR Record
    
    If Validating_ Then
        Exit Function    ' Exit after validation
    End If
    
    


    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If
    
    Call MTR_Sense_Alignment_Calc(FreqSensorAlignment, SensorAlignment, SweepVArrayDic, Temperature)

'===================================================================================================================================================='
'Read bincut voltage from eFuse
'===================================================================================================================================================='
'    If LCase(TheExec.CurrentJob) <> "cp1" Then Call Read_DVFM_To_GradeVDD
    Dim PrintStr As String
    Dim i, j, k As Integer
    Dim CP_GB_Record As Double
    Dim SensorSplitStr() As String
    Dim eFuseValueOnly() As New SiteDouble
    Dim eFuseCPGBOnly() As New SiteDouble
    Dim eFuseValueLowest() As New SiteDouble
    Dim eFuseValueHighest() As New SiteDouble
    Dim CurrentPassBinCutNum_MTR As New SiteLong
    SensorSplitStr = Split(SensorCalculate, ",")
    
    ReDim eFuseValueOnly(UBound(SensorSplitStr))
    ReDim eFuseCPGBOnly(UBound(SensorSplitStr))
    ReDim eFuseValueLowest(UBound(SensorSplitStr))
    ReDim eFuseValueHighest(UBound(SensorSplitStr))

    For Each site In TheExec.sites
        For i = 0 To UBound(SensorSplitStr)
            If LCase(SensorSplitStr(i)) Like "vdd_ecpu" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_pcpu" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_gpu" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_soc" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_ave" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            End If
        Next i
    Next site

    For Each site In TheExec.sites
        For i = 0 To UBound(SensorSplitStr)
            If LCase(SensorSplitStr(i)) Like "vdd_ecpu" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_pcpu" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_gpu" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_soc" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_ave" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            End If
        Next i
    Next site

''===================================================================================================================================================='
''Calculate measurement point is greater or not
''===================================================================================================================================================='
    Dim Catchdone() As Boolean
    Dim SensorVoltage() As String
    Dim VoltageMark() As New SiteDouble
    Dim VoltageLevelCNT() As New SiteLong
    Dim VoltageLevelOffset() As New SiteLong
    ReDim Catchdone(UBound(SensorSplitStr))
    ReDim VoltageMark(UBound(SensorSplitStr))
    ReDim VoltageLevelCNT(UBound(SensorSplitStr))
    ReDim VoltageLevelOffset(UBound(SensorSplitStr))
    
    SensorVoltage = Split(SweepVArrayValue, ",")
    For Each site In TheExec.sites
        For i = 0 To UBound(SensorSplitStr)
            VoltageLevelCNT(i) = UBound(SensorVoltage)
        Next i
    Next site
    
    For Each site In TheExec.sites
        For i = 0 To UBound(SensorSplitStr)
            Catchdone(i) = True
            VoltageLevelOffset(i) = 0
        Next i
        For i = 0 To UBound(SensorSplitStr)
            For j = 0 To UBound(SensorVoltage)
                If Catchdone(i) = True Then
                    If CDbl(SensorVoltage(j)) < eFuseValueLowest(i) Then
                        VoltageLevelOffset(i) = VoltageLevelOffset(i) + 1
                        VoltageMark(i) = CDbl(SensorVoltage(j))
                    ElseIf (CDbl(SensorVoltage(j)) >= eFuseValueLowest(i)) And Catchdone(i) = True And VoltageLevelOffset(i) <> 0 Then
                        If VoltageMark(i) < eFuseValueLowest(i) Then
                            VoltageLevelOffset(i) = VoltageLevelOffset(i) - 1
'                            PrintStr = "Site" + CStr(site) + "_" + CStr(SensorSplitStr(i)) + "_" + "LowestMode(eFuse-CPBG) : " + CStr(eFuseValueLowest(i))
'                            TheExec.Datalog.WriteComment (PrintStr)
                            Catchdone(i) = False
                        End If
                    End If
                End If
            Next j
            If VoltageLevelOffset(i) > 4 Then
                VoltageLevelOffset(i) = 4
                VoltageLevelCNT(i) = VoltageLevelCNT(i) - VoltageLevelOffset(i)
            Else
                VoltageLevelCNT(i) = VoltageLevelCNT(i) - VoltageLevelOffset(i)
            End If
        Next i
    Next site

    For i = 0 To UBound(SensorSplitStr)
        TheExec.Flow.TestLimit resultVal:=eFuseValueOnly(i), Tname:=CStr(SensorSplitStr(i)) + "_Lowest_eFuse", ForceResults:=tlForceNone, scaletype:=scaleNone
        TheExec.Flow.TestLimit resultVal:=eFuseCPGBOnly(i), Tname:=CStr(SensorSplitStr(i)) + "_Lowest_CPGB", ForceResults:=tlForceNone, scaletype:=scaleNone
        TheExec.Flow.TestLimit resultVal:=eFuseValueLowest(i), Tname:=CStr(SensorSplitStr(i)) + "_Lowest(eFuse - CPGB)", ForceResults:=tlForceNone, scaletype:=scaleNone
    Next i

'    For i = 0 To UBound(SensorSplitStr)
'        PrintStr = CStr(SensorSplitStr(i)) + "_" + "VoltageLevelCNT"
'        Call AddStoredMeasurement(PrintStr, VoltageLevelCNT(i))
'    Next i
    Dim DicEfuseName As String
    Dim DicEfuseRecord() As New DSPWave
    ReDim DicEfuseRecord(UBound(SensorSplitStr))
    For i = 0 To UBound(SensorSplitStr)
        DicEfuseRecord(i).CreateConstant 0, 1, DspLong
        PrintStr = CStr(SensorSplitStr(i)) + "_" + "VoltageLevelCNT"
        Call AddStoredMeasurement(PrintStr, VoltageLevelCNT(i))
        If SensorSplitStr(i) = "VDD_ECPU" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_ecpu"
        ElseIf SensorSplitStr(i) = "VDD_PCPU" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_pcpu"
        ElseIf SensorSplitStr(i) = "VDD_GPU" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_gpu"
        ElseIf SensorSplitStr(i) = "VDD_SOC" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_soc"
        ElseIf SensorSplitStr(i) = "VDD_AVE" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_ave"
        End If
        For Each site In TheExec.sites
            If UBound(SensorVoltage) > 13 Then
                If VoltageLevelCNT(i) = UBound(SensorVoltage) Then
                    DicEfuseRecord(i).Element(0) = 5
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 1 Then
                    DicEfuseRecord(i).Element(0) = 4
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 2 Then
                    DicEfuseRecord(i).Element(0) = 3
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 3 Then
                    DicEfuseRecord(i).Element(0) = 2
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 4 Then
                    DicEfuseRecord(i).Element(0) = 1
                End If
            Else
                If VoltageLevelCNT(i) = UBound(SensorVoltage) Then
                    DicEfuseRecord(i).Element(0) = 13
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 1 Then
                    DicEfuseRecord(i).Element(0) = 12
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 2 Then
                    DicEfuseRecord(i).Element(0) = 11
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 3 Then
                    DicEfuseRecord(i).Element(0) = 10
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 4 Then
                    DicEfuseRecord(i).Element(0) = 9
                End If
            End If
        Next site
        Call AddStoredCaptureData(DicEfuseName, DicEfuseRecord(i))
    Next i
    
'===================================================================================================================================================='
'DataProcess for Matrix calculrate
'===================================================================================================================================================='
    Dim SenName As String
    Dim MatrixMaxValue As New SiteLong
    Dim MatrixMinValue As New SiteLong
    Dim ThisPinType As String
    Dim SetInformation() As New DSPWave
    Dim Aininformation() As New DSPWave
    Dim Aixinformation() As New DSPWave
    Dim PiUInformation() As New DSPWave
    ReDim SetInformation(UBound(SensorSplitStr))
    ReDim Aininformation(UBound(SensorSplitStr))
    ReDim Aixinformation(UBound(SensorSplitStr))
    ReDim PiUInformation(UBound(SensorSplitStr))
    Dim ProcessLock As Boolean
    Dim RuleMonitor() As New SiteLong
'    Dim MaxEndPointMonitor_First() As New SiteLong
'    Dim MaxEndPointMonitor_Second() As New SiteLong
'    Dim MinEndPointMonitor_First() As New SiteLong
'    Dim MinEndPointMonitor_Second() As New SiteLong
    Dim FinalCheck As New SiteLong
    Dim Coeff_DspWave_a1 As New DSPWave
    Dim Coeff_DspWave_a2 As New DSPWave
    Dim SensorTempName_rot As String
    Dim SensorTempName_rov As String
    Dim sizeOfFuse_1 As Long
    sizeOfFuse_1 = CLng(FuseSize_1)
    Dim OutFuseDspWave_1 As New DSPWave
    OutFuseDspWave_1.CreateConstant 0, sizeOfFuse_1, DspLong
    Dim sizeOfFuse_2 As Long
    sizeOfFuse_2 = CLng(FuseSize_2)
    Dim OutFuseDspWave_2 As New DSPWave
    OutFuseDspWave_2.CreateConstant 0, sizeOfFuse_2, DspLong
    Dim DicCounterbyMain As Long
    Dim DicCounterbyMain_temp As Long
    Dim SensorGroup_main() As String
    Dim SensorArray_main() As String
    Dim StoreFuseDicNameTemp_1 As String
    Dim StoreFuseDicNameTemp_2 As String
    DicCounterbyMain = -1
    DicCounterbyMain_temp = 0
    SensorGroup_main = Split(FreqSensorAlignment, ";")
        
    For i = 0 To UBound(SensorGroup_main)
        SensorArray_main = Split(SensorGroup_main(i), ",")
        For j = 0 To UBound(SensorArray_main)
            DicCounterbyMain = DicCounterbyMain + 1
        Next j
    Next i
    DicCounterbyMain = (DicCounterbyMain + 1) / 2
        
    For i = 0 To UBound(SensorGroup_main)
        SensorArray_main = Split(SensorGroup_main(i), ",")
        ReDim RuleMonitor(UBound(SensorArray_main))
'        ReDim MaxEndPointMonitor_First(UBound(SensorArray_main))
'        ReDim MinEndPointMonitor_First(UBound(SensorArray_main))
'        ReDim MaxEndPointMonitor_Second(UBound(SensorArray_main))
'        ReDim MinEndPointMonitor_Second(UBound(SensorArray_main))
        For j = 0 To UBound(SensorSplitStr)
            For k = 0 To UBound(SensorArray_main)
'''''                If (k + 1) Mod 2 = 0 Then
'''''                    If CStr(SensorSplitStr(j)) = "VDD_GPU" Then
'''''                        If SensorArray_main(k) Like "*GPU*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    ElseIf CStr(SensorSplitStr(j)) = "VDD_ECPU" Then
'''''                        If SensorArray_main(k) Like "*ECPU*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    ElseIf CStr(SensorSplitStr(j)) = "VDD_PCPU" Then
'''''                        If SensorArray_main(k) Like "*ANE*" Or SensorArray_main(k) Like "*PCPU*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    ElseIf CStr(SensorSplitStr(j)) = "VDD_AVE" Then
'''''                        If SensorArray_main(k) Like "*AVE*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    ElseIf CStr(SensorSplitStr(j)) = "VDD_SOC" Then
'''''                        If SensorArray_main(k) Like "*SOC*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    Else
'''''                        ProcessLock = False
'''''                    End If
                
'''''                    If ProcessLock = True Then
'''''                        For Each site In TheExec.sites
'''''                            SetInformation(j).CreateConstant 0, 300, DspLong
'''''                            Aininformation(j).CreateConstant 0, 300, DspLong
'''''                            Aixinformation(j).CreateConstant 0, 300, DspLong
'''''                            PiUInformation(j).CreateConstant 0, 300, DspLong
'''''                            If MTR_CAL_Sheet Like "*1150*" Then
'''''                                MatrixMaxValue = 1150
'''''                            Else
'''''                                MatrixMaxValue = 950
'''''                            End If
'''''                            If VoltageLevelCNT(j) = UBound(SensorVoltage) Then
'''''                                ThisPinType = MTR_CAL_Sheet + "450"
'''''                                MatrixMinValue = 450
'''''                            ElseIf VoltageLevelCNT(j) = UBound(SensorVoltage) - 1 Then
'''''                                ThisPinType = MTR_CAL_Sheet + "475"
'''''                                MatrixMinValue = 475
'''''                            ElseIf VoltageLevelCNT(j) = UBound(SensorVoltage) - 2 Then
'''''                                ThisPinType = MTR_CAL_Sheet + "500"
'''''                                MatrixMinValue = 500
'''''                            ElseIf VoltageLevelCNT(j) = UBound(SensorVoltage) - 3 Then
'''''                                ThisPinType = MTR_CAL_Sheet + "525"
'''''                                MatrixMinValue = 525
'''''                            ElseIf VoltageLevelCNT(j) = UBound(SensorVoltage) - 4 Then
'''''                                ThisPinType = MTR_CAL_Sheet + "550"
'''''                                MatrixMinValue = 550
'''''                            End If
'''''                            SetInformation(j)(site) = Public_GetStoredCaptureData(ThisPinType + "_" + "size")
'''''                            Aixinformation(j)(site) = Public_GetStoredCaptureData(ThisPinType + "_" + "Aix")
'''''                            Aininformation(j)(site) = Public_GetStoredCaptureData(ThisPinType + "_" + "Ain")
'''''                            PiUInformation(j)(site) = Public_GetStoredCaptureData(ThisPinType + "_" + "PiUInfo")
'''''                            SenName = Left(SensorArray_main(k), (InStr(SensorArray_main(k), "_") - 1))
''''''                            PrintStr = "Site" + CStr(site) + "_" + CStr(SensorSplitStr(j)) + "_" + SenName + "_" + ThisPinType
''''''                            TheExec.Datalog.WriteComment (PrintStr)
'''''                        Next site
'''''
'''''                        TheExec.Flow.TestLimit resultVal:=MatrixMinValue, TName:=SenName + "_" + "Matrix_StartPoint", ForceResults:=tlForceNone
'''''                        TheExec.Flow.TestLimit resultVal:=MatrixMaxValue, TName:=SenName + "_" + "Matrix_EndPoint", ForceResults:=tlForceNone
'''''
''''''    Call MTR_LinearRegression(SensorArray_main(k - 1), SensorArray_main(k), SweepVArrayValue, SweepVArrayDic, Temperature, VoltageLevelCNT(j))
'''''
'''''                        Dim OutFuseDspWave_1_temp() As New DSPWave
'''''                        Dim OutFuseDspWave_2_temp() As New DSPWave
'''''                        ReDim OutFuseDspWave_1_temp(DicCounterbyMain)
'''''                        ReDim OutFuseDspWave_2_temp(DicCounterbyMain)
'''''
'''''                        SensorTempName_rot = SensorArray_main(k - 1) + "_" + Temperature + "C"
'''''                        SensorTempName_rov = SensorArray_main(k) + "_" + Temperature + "C"
'''''
'''''                        Call Calc_FromLoad_MTR_SE_CAL_Coeff(SensorTempName_rot, SensorTempName_rov, Temperature, sizeOfFuse_1, sizeOfFuse_2, SetInformation(j), PiUInformation(j), _
'''''                        Aixinformation(j), Aininformation(j), OutFuseDspWave_1, OutFuseDspWave_2, CLng(UBound(SensorVoltage)), VoltageLevelCNT(j), RuleMonitor(k - 1), RuleMonitor(k))
'''''
'''''                        StoreFuseDicNameTemp_1 = StoreFuseDicName_1
'''''                        StoreFuseDicNameTemp_2 = StoreFuseDicName_2
'''''
'''''                        OutFuseDspWave_1_temp(DicCounterbyMain_temp) = OutFuseDspWave_1
'''''                        OutFuseDspWave_2_temp(DicCounterbyMain_temp) = OutFuseDspWave_2
'''''
'''''                        StoreFuseDicNameTemp_1 = StoreFuseDicNameTemp_1 + "_" + SensorTempName_rot
'''''                        Call AddStoredCaptureData(StoreFuseDicNameTemp_1, OutFuseDspWave_1_temp(DicCounterbyMain_temp))
'''''                        StoreFuseDicNameTemp_2 = StoreFuseDicNameTemp_2 + "_" + SensorTempName_rov
'''''                        Call AddStoredCaptureData(StoreFuseDicNameTemp_2, OutFuseDspWave_2_temp(DicCounterbyMain_temp))
'''''                        DicCounterbyMain_temp = DicCounterbyMain_temp + 1
'''''                    End If
'''''                End If
            Next k
        Next j
    Next i
'''''
'''''    For i = 0 To UBound(SensorArray_main)
'''''        TheExec.Flow.TestLimit resultVal:=RuleMonitor(i), ForceResults:=tlForceFlow
'''''    Next i
End Function



Public Function MTRG_t5p2a_DigSrc_Coefficient_PreCalc(Optional v0 As String, Optional v1 As String, Optional x0a As String, Optional x1a As String, Optional c0_DictName As String, Optional c1_DictName As String, Optional c2_DictName As String, Optional c3_DictName As String, Optional Validating_ As Boolean) As Long
    If Validating_ Then
        Exit Function    ' Exit after validation
    End If
     Dim site As Variant
     
    On Error GoTo errHandler
    
    Dim DSPWave_v0 As New DSPWave
    DSPWave_v0.CreateConstant 0, 1
    Dim DSPWave_v1 As New DSPWave
    DSPWave_v1.CreateConstant 0, 1
    Dim DSPWave_x0a As New DSPWave
    DSPWave_x0a.CreateConstant 0, 1
    Dim DSPWave_x1a As New DSPWave
    DSPWave_x1a.CreateConstant 0, 1
    
    Dim DSPWave_Binary_v0 As New DSPWave
    DSPWave_v0.CreateConstant 0, 18
    Dim DSPWave_Binary_v1 As New DSPWave
    DSPWave_v1.CreateConstant 0, 18
    Dim DSPWave_Binary_x0a As New DSPWave
    DSPWave_x0a.CreateConstant 0, 18
    Dim DSPWave_Binary_x1a As New DSPWave
    DSPWave_x1a.CreateConstant 0, 18
    
    Dim DSPWave_x0 As New DSPWave
    DSPWave_x0.CreateConstant 0, 1
    Dim DSPWave_x1 As New DSPWave
    DSPWave_x1.CreateConstant 0, 1
    Dim DSPWave_y0 As New DSPWave
    DSPWave_y0.CreateConstant 0, 1
    Dim DSPWave_y1 As New DSPWave
    DSPWave_y1.CreateConstant 0, 1
    
    Dim DSPWave_a0cal As New DSPWave
    DSPWave_a0cal.CreateConstant 0, 1
    Dim DSPWave_a1cal As New DSPWave
    DSPWave_a1cal.CreateConstant 0, 1
    
    Dim DSPWave_c0 As New DSPWave
    DSPWave_c0.CreateConstant 0, 1
    Dim DSPWave_c1 As New DSPWave
    DSPWave_c1.CreateConstant 0, 1
    Dim DSPWave_c2 As New DSPWave
    DSPWave_c2.CreateConstant 0, 1
    Dim DSPWave_c3 As New DSPWave
    DSPWave_c3.CreateConstant 0, 1
    
    Dim DSPWave_Binary_c0 As New DSPWave
    DSPWave_Binary_c0.CreateConstant 0, 32
    Dim DSPWave_Binary_c1 As New DSPWave
    DSPWave_Binary_c1.CreateConstant 0, 32
    Dim DSPWave_Binary_c2 As New DSPWave
    DSPWave_Binary_c2.CreateConstant 0, 32
    Dim DSPWave_Binary_c3 As New DSPWave
    DSPWave_Binary_c3.CreateConstant 0, 32
    Dim C_BitWidth As New SiteLong
    For Each site In TheExec.sites
        C_BitWidth(site) = DSPWave_Binary_c0(site).SampleSize
    Next site
    
    Dim A0 As Double: A0 = -21.5822184999726
    Dim a1 As Double: a1 = 428.009226609628
    Dim a2 As Double: a2 = -133.454310922823
    Dim a3 As Double: a3 = 19.0485545665615
    
    
    DSPWave_Binary_v0 = GetStoredCaptureData(v0)
    DSPWave_Binary_v1 = GetStoredCaptureData(v1)
    DSPWave_Binary_x0a = GetStoredCaptureData(x0a)
    DSPWave_Binary_x1a = GetStoredCaptureData(x1a)     '25C and 85C have same dictionary name.
    
    Dim i As Integer
    If currentJobName Like "cp1" Then   'work around 25C and 85C have same dictionary name.
        For Each site In TheExec.sites
            For i = 0 To DSPWave_Binary_x1a(site).SampleSize - 1
                DSPWave_Binary_x1a(site).Element(i) = 0
            Next i
        Next site
    End If
    
    Dim SL_BitWidth As New SiteLong
    For Each site In TheExec.sites
        SL_BitWidth(site) = DSPWave_Binary_v0(site).SampleSize
    Next site
    
    Call rundsp.DSP_2S_Complement_To_SignDec(DSPWave_Binary_v0, SL_BitWidth, DSPWave_v0)
    Call rundsp.DSP_DivideConstant(DSPWave_v0, 2 ^ 13)
    Call rundsp.BinToDec(DSPWave_Binary_v1, DSPWave_v1)
    Call rundsp.DSP_DivideConstant(DSPWave_v1, 2 ^ 17)
    Call rundsp.BinToDec(DSPWave_Binary_x0a, DSPWave_x0a)
    Call rundsp.DSP_DivideConstant(DSPWave_x0a, 2 ^ 13)
    Call rundsp.BinToDec(DSPWave_Binary_x1a, DSPWave_x1a)
    Call rundsp.DSP_DivideConstant(DSPWave_x1a, 2 ^ 13)
    
'    DSPWave_x0 = DSPWave_x0a
'    Call rundsp.DSP_Subtract(DSPWave_x0, DSPWave_v0)
'    Call rundsp.DSP_Divide(DSPWave_x0, DSPWave_v1)
'
'    DSPWave_x1 = DSPWave_x1a
'    Call rundsp.DSP_Subtract(DSPWave_x1, DSPWave_v0)
'    Call rundsp.DSP_Divide(DSPWave_x1, DSPWave_v1)
    
    For Each site In TheExec.sites
        If DSPWave_v1(site).Element(0) = 0 Then DSPWave_v1(site).Element(0) = 0.0000000001
        DSPWave_x0(site).Element(0) = (DSPWave_x0a(site).Element(0) - DSPWave_v0(site).Element(0)) / DSPWave_v1(site).Element(0)
        DSPWave_x1(site).Element(0) = (DSPWave_x1a(site).Element(0) - DSPWave_v0(site).Element(0)) / DSPWave_v1(site).Element(0)
        DSPWave_y0(site).Element(0) = 25 + 273.15 - a2 * DSPWave_x0(site).Element(0) ^ 2 - a3 * DSPWave_x0(site).Element(0) ^ 3
        DSPWave_y1(site).Element(0) = 85 + 273.15 - a2 * DSPWave_x1(site).Element(0) ^ 2 - a3 * DSPWave_x1(site).Element(0) ^ 3
        If DSPWave_x1(site).Element(0) = DSPWave_x0(site).Element(0) Then DSPWave_x1(site).Element(0) = DSPWave_x1(site).Element(0) + 0.0000000001
        DSPWave_a0cal(site).Element(0) = (DSPWave_x1(site).Element(0) * DSPWave_y0(site).Element(0) - DSPWave_x0(site).Element(0) * DSPWave_y1(site).Element(0)) / (DSPWave_x1(site).Element(0) - DSPWave_x0(site).Element(0))
        DSPWave_a1cal(site).Element(0) = (DSPWave_y1(site).Element(0) - DSPWave_y0(site).Element(0)) / (DSPWave_x1(site).Element(0) - DSPWave_x0(site).Element(0)) / (DSPWave_x1(site).Element(0) - DSPWave_x0(site).Element(0))
        DSPWave_c0(site).Element(0) = DSPWave_a0cal(site).Element(0) - DSPWave_a1cal(site).Element(0) * DSPWave_v0(site).Element(0) / DSPWave_v1(site).Element(0)
        DSPWave_c1(site).Element(0) = DSPWave_a1cal(site).Element(0) / DSPWave_v1(site).Element(0) - 2 * a2 * DSPWave_v0(site).Element(0) / DSPWave_v1(site).Element(0) ^ 2
        DSPWave_c2(site).Element(0) = a2 / DSPWave_v1(site).Element(0) ^ 2 - 3 * a3 * DSPWave_v0(site).Element(0) / DSPWave_v1(site).Element(0) ^ 3
        DSPWave_c3(site).Element(0) = a3 / DSPWave_v1(site).Element(0) ^ 3
        DSPWave_c0(site).Element(0) = FormatNumber(DSPWave_c0(site).Element(0) * 2 ^ 13, 0)
        DSPWave_c1(site).Element(0) = FormatNumber(DSPWave_c1(site).Element(0) * 2 ^ 13, 0)
        DSPWave_c2(site).Element(0) = FormatNumber(DSPWave_c2(site).Element(0) * 2 ^ 13, 0)
        DSPWave_c3(site).Element(0) = FormatNumber(DSPWave_c3(site).Element(0) * 2 ^ 13, 0)
    Next site
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_x0.Element(0), Tname:="x0", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=DSPWave_x1.Element(0), Tname:="x1", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=DSPWave_y0.Element(0), Tname:="y0", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=DSPWave_y1.Element(0), Tname:="y1", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=DSPWave_c0.Element(0), Tname:="c0", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=DSPWave_c1.Element(0), Tname:="c1", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=DSPWave_c2.Element(0), Tname:="c2", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=DSPWave_c3.Element(0), Tname:="c3", ForceResults:=tlForceNone
    
    'Call rundsp.DSP_Convert_2S_Complement(DSPWave_c0, C_BitWidth, DSPWave_Binary_c0)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c0, C_BitWidth, DSPWave_Binary_c0)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c1, C_BitWidth, DSPWave_Binary_c1)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c2, C_BitWidth, DSPWave_Binary_c2)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c3, C_BitWidth, DSPWave_Binary_c3)
    
    Call AddStoredCaptureData(c0_DictName, DSPWave_Binary_c0)
    Call AddStoredCaptureData(c1_DictName, DSPWave_Binary_c1)
    Call AddStoredCaptureData(c2_DictName, DSPWave_Binary_c2)
    Call AddStoredCaptureData(c3_DictName, DSPWave_Binary_c3)
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error in MTRG_t5p2a_PreCalculation"
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function MTRG_t6p3abc_DigSrc_Coefficient_PreCalc(Optional v0 As String, Optional v1 As String, Optional c0_DictName As String, Optional c1_DictName As String, Optional Validating_ As Boolean) As Long
    If Validating_ Then
        Exit Function    ' Exit after validation
    End If
    On Error GoTo errHandler
    Dim site As Variant
    
    Dim DSPWave_v0 As New DSPWave
    DSPWave_v0.CreateConstant 0, 1
    Dim DSPWave_v1 As New DSPWave
    DSPWave_v1.CreateConstant 0, 1
    
    Dim DSPWave_Binary_v0 As New DSPWave
    DSPWave_v0.CreateConstant 0, 18
    Dim DSPWave_Binary_v1 As New DSPWave
    DSPWave_v1.CreateConstant 0, 18
    
    Dim DSPWave_c0 As New DSPWave
    DSPWave_c0.CreateConstant 0, 1
    Dim DSPWave_c1 As New DSPWave
    DSPWave_c1.CreateConstant 0, 1
    
    Dim DSPWave_Binary_c0 As New DSPWave
    DSPWave_Binary_c0.CreateConstant 0, 32
    Dim DSPWave_Binary_c1 As New DSPWave
    DSPWave_Binary_c1.CreateConstant 0, 32
    
    Dim Vref As Double: Vref = 10
    
    DSPWave_Binary_v0 = GetStoredCaptureData(v0)
    DSPWave_Binary_v1 = GetStoredCaptureData(v1)
    
    Dim SL_BitWidth As New SiteLong
    For Each site In TheExec.sites
        SL_BitWidth(site) = DSPWave_Binary_v0(site).SampleSize
    Next site
    Dim C_BitWidth As New SiteLong
    For Each site In TheExec.sites
        C_BitWidth(site) = DSPWave_Binary_c0(site).SampleSize
    Next site
    
    Call rundsp.DSP_2S_Complement_To_SignDec(DSPWave_Binary_v0, SL_BitWidth, DSPWave_v0)
    Call rundsp.DSP_DivideConstant(DSPWave_v0, 2 ^ 13)
    Call rundsp.BinToDec(DSPWave_Binary_v1, DSPWave_v1)
    Call rundsp.DSP_DivideConstant(DSPWave_v1, 2 ^ 17)
    
    For Each site In TheExec.sites
        If DSPWave_v1(site).Element(0) = 0 Then DSPWave_v1(site).Element(0) = 0.0000000001
        DSPWave_c0(site).Element(0) = 273.15 - Vref * DSPWave_v0(site).Element(0) / DSPWave_v1(site).Element(0)
        DSPWave_c1(site).Element(0) = Vref / DSPWave_v1(site).Element(0)
        DSPWave_c0(site).Element(0) = FormatNumber(DSPWave_c0(site).Element(0) * 2 ^ 13, 0)
        DSPWave_c1(site).Element(0) = FormatNumber(DSPWave_c1(site).Element(0) * 2 ^ 13, 0)
    Next site
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_c0.Element(0), Tname:="c0", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=DSPWave_c1.Element(0), Tname:="c1", ForceResults:=tlForceNone
    
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c0, C_BitWidth, DSPWave_Binary_c0)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c1, C_BitWidth, DSPWave_Binary_c1)
    
    Call AddStoredCaptureData(c0_DictName, DSPWave_Binary_c0)
    Call AddStoredCaptureData(c1_DictName, DSPWave_Binary_c1)
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error in MTRG_t5p2a_PreCalculation"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DSSC_Search(Optional Pat As String, Optional MeasureV_pin As PinList, Optional DigSrc_pin As PinList, _
    Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimFormat As String, Optional TrimRepeat As Long, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Validating_ As Boolean)
''    Dim PatCount As Long
    Dim i As Integer
    Dim pats() As String
    Dim code As New SiteLong
    Dim vout As New SiteDouble
    Dim BestCode As New SiteLong, BestVal As New SiteDouble, verr As New SiteDouble, Temp As New SiteLong
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim trace As Boolean
    Dim site As Variant
    
    Dim PatCount As Long, PattArray() As String
    
    If Validating_ Then
        Call PrLoadPattern(Pat)
        Exit Function    ' Exit after validation
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    On Error GoTo errHandler
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    PATT_GetPatListFromPatternSet Pat, pats, PatCount
    With TheHdw.DCVI.Pins(MeasureV_pin)
            .Gate = False
            .Disconnect tlDCVIConnectDefault
            .mode = tlDCVIModeHighImpedance
            .Connect tlDCVIConnectHighSense
            .Voltage = 6
            .current = 0
             TheHdw.Wait 0.5 * ms
            .Gate = True
    End With
    
    code = TrimStart
    
    Dim SplitByEqual() As String, SplitByColon() As String, TrimCodeSize As Long
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Mid As Long, TrimCodeValue_Max As Long
    SplitByEqual = Split(TrimFormat, "=")
    SplitByColon = Split(SplitByEqual(1), ":")
    TrimCodeSize = SplitByColon(0) + 1
    TrimCodeValue_Min = 0
    TrimCodeValue_Mid = (2 ^ TrimCodeSize) / 2
    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
    
    Call DSSC_Search_par_run(pats(0), DigSrc_pin.Value, code, MeasureV_pin.Value, vout, TrimCodeSize, TrimRepeat)
    If trace Then TheExec.Flow.TestLimit code
    If trace Then TheExec.Flow.TestLimit vout
    First = vout.compare(LessThan, TrimTarget)
    If trace Then TheExec.Flow.TestLimit First, , , , , , , , "first"
    BestCode = code
    BestVal = vout
    verr = vout.Subtract(TrimTarget).Abs

    For i = 0 To TrimCodeValue_Mid - 1
        If trace Then TheExec.Datalog.WriteComment ("i = " & i)
        Temp = vout.compare(LessThan, TrimTarget)
        code = code.Add(Temp.Multiply(-2).Subtract(1))  ' If vout < TrimTarget Then code++ Else code--
        
        For Each site In TheExec.sites.Active
            If code(site) > TrimCodeValue_Max Then: code(site) = code(site) - 1: ''GoTo EndTrim
            If code(site) < TrimCodeValue_Min Then: code(site) = code(site) + 1: ''GoTo EndTrim
        Next site
      
        Call DSSC_Search_par_run(pats(0), DigSrc_pin.Value, code, MeasureV_pin.Value, vout, TrimCodeSize, TrimRepeat)
        If trace Then TheExec.Flow.TestLimit code
        If trace Then TheExec.Flow.TestLimit vout
        For Each site In TheExec.sites
            If Abs(vout - TrimTarget) < verr Then
                BestCode = code
                BestVal = vout
                verr = Abs(vout - TrimTarget)
            End If
        Next site

        Done = Done.LogicalOr(First.LogicalXor(vout.compare(LessThan, TrimTarget)))
        If trace Then TheExec.Flow.TestLimit Done, , , , , , , , "done"
        If Done.All(True) Then Exit For
    Next i
    
EndTrim:
    TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , "Volt_meas_LPDPRX_LDO", , MeasureV_pin, , , , , tlForceFlow
    TheExec.Flow.TestLimit BestCode, 0, 15, , , , , , "LPDPRX_LDO_Trim", , , , , , , tlForceFlow
    HardIP_WriteFuncResult
    
    Dim TempVal As Integer
    Dim FinalTrimCode As New DSPWave
    
    FinalTrimCode.CreateConstant 0, TrimCodeSize
    
    For Each site In TheExec.sites
        TempVal = BestCode(site)
        For i = 0 To TrimCodeSize - 1
            FinalTrimCode(site).Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i
    Next site
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, FinalTrimCode)
    End If
    
    '' 20170704 - Add write efuse function
''    Dim sl_Fuse_Val As SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else
''        For Each Site In TheExec.sites
''            sl_Fuse_Val(Site) = BestCode(Site)
''        Next Site
        
        If TrimFuseName <> "" And TrimFuseTypeName <> "" Then
            ''Call HIP_eFuse_Write(TrimFuseTypeName, TrimFuseName, BestCode) ''set fuse information from flow
        End If
    End If
        
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search: " & err.Description)
    DSSC_Search = TL_ERROR
End Function




Public Function DSSC_Search_LDO(Optional Pat As Pattern, Optional MeasureV_pin As PinList, Optional MeasV_Name As String, Optional MeasV_Name_Trim As String, Optional MeaV_WaitTime As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional TrimMethod As String, Optional TrimStepSize As Double, Optional Validating_ As Boolean)
   Dim site As Variant
    Dim i As Integer
    Dim pats() As String
    Dim code As New SiteLong: code = TrimStart
    Dim vout As New SiteDouble
    Dim NumberOfMeasV As Integer: NumberOfMeasV = UBound(Split(MeasV_Name, "+")) + 1
    Dim BestCode As New SiteLong, BestVal() As New SiteDouble, Temp As New SiteLong
    ReDim BestVal(NumberOfMeasV - 1) As New SiteDouble
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim PreviousNegative As New SiteBoolean
    Dim PreviousPositive As New SiteBoolean
    Dim step As New SiteBoolean
    Dim DecideTrim As New SiteBoolean
        For Each site In TheExec.sites.Active
            PreviousNegative = False
            PreviousPositive = False
            DecideTrim = False
        Next site
    Dim blockName() As String: blockName = Split(TheExec.DataManager.instanceName, "_")
    Dim MeasValue() As New SiteDouble: ReDim MeasValue(NumberOfMeasV - 1)
    Dim PreviousMeasValue() As New SiteDouble: ReDim PreviousMeasValue(NumberOfMeasV - 1) As New SiteDouble
    Dim StepCount As Long: StepCount = 0
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim MeasV_Name_Array() As String: MeasV_Name_Array = Split(MeasV_Name, "+")
    Dim MeasV_Name_Trim_Array() As String: MeasV_Name_Trim_Array = Split(MeasV_Name_Trim, "+")
    Dim TrimPoint() As Long: ReDim TrimPoint(UBound(Split(MeasV_Name, "+")))
    
    Call ProcessInputToGLB(Pat, "V", False, , , , , MeasureV_pin.Value, , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , MeaV_WaitTime)
     
    For i = 0 To NumberOfMeasV - 1
        For Each site In TheExec.sites.Active
            PreviousMeasValue(i).Value = 0
            MeasValue(i).Value = 0
        Next site
        If MeasV_Name_Array(i) = MeasV_Name_Trim_Array(i) Then: TrimPoint(i) = 1
    Next i
    
    Call GetFlowTName

    If Validating_ Then
        Call PrLoadPattern(Pat.Value)
        Exit Function    ' Exit after validation
    End If
    
    On Error GoTo errHandler
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    
    If TheExec.DevChar.Setups.IsRunning Then
        If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Contains(tlDevCharShmooAxis_Y) Then
            If gl_Flag_HardIP_Trim_Set_PrePoint And Not (gl_Flag_HardIP_Characterization_1stRun) Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_0_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value)
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_1_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value)
                Call TheExec.Overlays.ApplyUniformSpecToHW("PCIE_XI0_0_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value)
            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_0_Shmoo_Freq_VAR", 24000000#)
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_1_Shmoo_Freq_VAR", 24000000#)
                Call TheExec.Overlays.ApplyUniformSpecToHW("PCIE_XI0_0_Shmoo_Freq_VAR", 24000000#)
            Else
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            End If
        End If
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet Pat.Value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Mid As Long, TrimCodeValue_Max As Long
    TrimCodeValue_Min = 0
    TrimCodeValue_Mid = (2 ^ TrimCodeSize) / 2
    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
  

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
    Call DSSC_Search_par_run_LDO(pats(0), DigSrc_pin, code, MeasureV_pin, vout, TrimCodeSize, NumberOfMeasV, MeasV_Name_Array(), MeasValue(), TrimPoint(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName, MeaV_WaitTime)

    First = vout.compare(LessThan, TrimTarget)
    BestCode = code
    For Each site In TheExec.sites.Active
        For i = 0 To NumberOfMeasV - 1
            BestVal(i) = MeasValue(i)
        Next i
    Next site
    
    If TrimMethod = "LinearSearch" Then
        For Each site In TheExec.sites.Active
            If vout.compare(LessThan, TrimTarget) Then
                code = code + 1
            ElseIf vout.compare(GreaterThan, TrimTarget) Then
                code = code - 1
            End If
            step = True
        Next site
    Else
        For Each site In TheExec.sites.Active
            code = code + Fix((TrimTarget - vout) / TrimStepSize)
            If Fix((TrimTarget - vout) / TrimStepSize) <> 0 Then: step = True
        Next site
    End If
    
StartTrim:
        If step.Any(True) Then
            StepCount = StepCount + 1
            
            If gl_Disable_HIP_debug_log = False Then
                If Right(CStr(StepCount), 1) = "1" Then
                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "st Trim Process ****************")
                ElseIf Right(CStr(StepCount), 1) = "2" Then
                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "nd Trim Process ****************")
                ElseIf Right(CStr(StepCount), 1) = "3" Then
                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "rd Trim Process ****************")
                Else
                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "th Trim Process ****************")
                End If
            End If
            
            
            Call DSSC_Search_par_run_LDO(pats(0), DigSrc_pin, code, MeasureV_pin, vout, TrimCodeSize, NumberOfMeasV, MeasV_Name_Array(), MeasValue(), TrimPoint(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName, MeaV_WaitTime)
        End If

        For Each site In TheExec.sites.Active
        If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
            BestCode = code
            For i = 0 To NumberOfMeasV - 1
                BestVal(i) = MeasValue(i)
            Next i
            DecideTrim = False
        ElseIf code.compare(GreaterThan, TrimCodeValue_Max) Then
            code = TrimCodeValue_Max
            DecideTrim = True

        ElseIf code.compare(LessThan, TrimCodeValue_Min) Then
            code = TrimCodeValue_Min
            DecideTrim = True

        ElseIf code.compare(EqualTo, TrimCodeValue_Max) Or code.compare(EqualTo, TrimCodeValue_Min) Then
            BestCode = code
            For i = 0 To NumberOfMeasV - 1
                BestVal(i) = MeasValue(i)
            Next i
            DecideTrim = False
        ElseIf vout.compare(LessThan, TrimTarget) And PreviousPositive Then
            BestCode = code + 1
            For i = 0 To NumberOfMeasV - 1
                BestVal(i) = PreviousMeasValue(i)
            Next i
            DecideTrim = False
        ElseIf vout.compare(LessThan, TrimTarget) And Not (PreviousPositive) Then
            code(site) = code(site) + 1
            PreviousNegative = True
            For i = 0 To NumberOfMeasV - 1
                PreviousMeasValue(i) = MeasValue(i)
            Next i
            step = True
            DecideTrim = True

        ElseIf vout.compare(GreaterThan, TrimTarget) And PreviousNegative Then
            BestCode = code
            For i = 0 To NumberOfMeasV - 1
                BestVal(i) = MeasValue(i)
            Next i
            DecideTrim = False
        ElseIf vout.compare(GreaterThan, TrimTarget) And Not (PreviousNegative) Then
            code(site) = code(site) - 1
            PreviousPositive = True
            For i = 0 To NumberOfMeasV - 1
                PreviousMeasValue(i) = MeasValue(i)
            Next i
            step = True
            DecideTrim = True

        End If
        Next site
        
    If DecideTrim.Any(True) Then GoTo StartTrim
        
    For i = 0 To NumberOfMeasV - 1
        TestNameInput = Report_TName_From_Instance("V", MeasureV_pin.Value, "", i, 0)
        If Not ByPassTestLimit Then: TheExec.Flow.TestLimit BestVal(i), , , , , , unitVolt, , TestNameInput, , MeasureV_pin, , , , , tlForceFlow
    Next i
    
    TestNameInput = Report_TName_From_Instance("C", "", , 0, 0)
    If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=BestCode, Tname:=TestNameInput, ForceResults:=tlForceFlow
    

    
    Dim TempVal As Integer
    Dim FinalTrimCode As New DSPWave
    
    FinalTrimCode.CreateConstant 0, TrimCodeSize
    
    For Each site In TheExec.sites
        TempVal = BestCode(site)
        For i = 0 To TrimCodeSize - 1
            FinalTrimCode(site).Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i
    Next site
    
    If TrimStoreName <> "" Then
        Call AddStoredCaptureData(TrimStoreName, FinalTrimCode)
    End If
    DebugPrintFunc Pat.Value
           
    If InStr(UCase(TheExec.DataManager.instanceName), UCase("VREGDT")) <> 0 Or InStr(UCase(TheExec.DataManager.instanceName), UCase("DTVREG")) Then
        Dim FinalTrimCode_DCVREG As New DSPWave
        Dim BestCode_DCVREG As New SiteLong
        FinalTrimCode_DCVREG.CreateConstant 0, 3
        
        
        For Each site In TheExec.sites
            If BestCode(site) = 0 Then BestCode_DCVREG(site) = 0
            If BestCode(site) <= 6 And BestCode(site) >= 1 Then BestCode_DCVREG(site) = 1
            If BestCode(site) <= 12 And BestCode(site) >= 7 Then BestCode_DCVREG(site) = 2
            If BestCode(site) <= 17 And BestCode(site) >= 13 Then BestCode_DCVREG(site) = 3
            If BestCode(site) <= 23 And BestCode(site) >= 18 Then BestCode_DCVREG(site) = 4
            If BestCode(site) <= 29 And BestCode(site) >= 24 Then BestCode_DCVREG(site) = 5
            If BestCode(site) <= 31 And BestCode(site) >= 30 Then BestCode_DCVREG(site) = 6
            
            TempVal = BestCode_DCVREG(site)
            For i = 0 To 2
                FinalTrimCode_DCVREG(site).Element(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        Next site
        TestNameInput = Report_TName_From_Instance("C", "", , 0, 0)
        If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=BestCode_DCVREG, Tname:=TestNameInput, ForceResults:=tlForceFlow
        If UCase(Inst_Name_Str) Like "CIO*NV" Then: Call AddStoredCaptureData("CIO_LDO_TRIM_VREGDC", FinalTrimCode_DCVREG)
        If UCase(Inst_Name_Str) Like "PCIE*NV" Then: Call AddStoredCaptureData("PCIE_LDO_TRIM_VREGDC", FinalTrimCode_DCVREG)
        If UCase(Inst_Name_Str) Like "LPDPRX*NV" Then: Call AddStoredCaptureData("LPDPRX_LDO_TRIM_DCVREG", FinalTrimCode_DCVREG)
    End If
    
    Call HardIP_WriteFuncResult(, , Inst_Name_Str)
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search_LDO: " & err.Description)
    DSSC_Search = TL_ERROR
End Function

Public Function TrimUVI80Code_VFI_2sComplement(TwoS_Complement As Boolean, Optional Pat As String, Optional TestSequence As String, Optional MeasV_Pins As PinList, Optional MeasI_pinS As PinList, Optional MeasI_Range As Double, _
    Optional MeasF_PinS_SingleEnd As PinList, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, _
    Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimFormat As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_DigCapData As String = "", _
    Optional TrimStoreName As String, Optional Diffaccuracy As Double, Optional TrimMethod As Long, _
    Optional Meas_StoreName As String, Optional Calc_Eqn As String, Optional TrimCal_Name As String, Optional Antitrim As Boolean = False, Optional Validating_ As Boolean, Optional Interpose_PrePat As String, Optional CPUA_Flag_In_Pat As Boolean = True, Optional Interpose_PostTest As String, _
    Optional Final_Calc As Boolean = False)
''    Dim PatCount As Long
    Dim i As Integer
    Dim pats() As String
    Dim code As New SiteLong
    Dim MeasValue As New SiteDouble
    Dim BestCode As New SiteLong, BestVal As New SiteDouble, verr As New SiteDouble, Temp As New SiteLong
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim trace As Boolean
    Dim site As Variant
    Dim Ts As Variant
    Dim ADCOUT As New SiteBoolean
    Dim TrimStep As Long
    Dim OutDSP As New DSPWave
    Dim PatCount As Long, PattArray() As String
    Dim TestSequence_array() As String
    Dim doallFlag As Boolean
    Dim finalflag As Boolean
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim TestNameInput As String
    Dim ReCalc As New SiteDouble
    Dim DoneEndTrim As New SiteBoolean
    gl_TName_Pat = Pat
    
    Call GetFlowTName
    
    If Validating_ Then
        Call PrLoadPattern(Pat)
        Exit Function    ' Exit after validation
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    TestSequence_array = Split(TestSequence, ",")
    
    On Error GoTo errHandler
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    PATT_GetPatListFromPatternSet Pat, pats, PatCount
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    
    Dim SplitByEqual() As String, SplitByColon() As String, TrimCodeSize As Long
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    Dim Trimname As String
    SplitByEqual = Split(TrimFormat, "=")
    SplitByColon = Split(SplitByEqual(1), ":")
    Trimname = SplitByEqual(0)
    TrimCodeSize = SplitByColon(0) + 1
    
    If TwoS_Complement = True Then
        TrimCodeValue_Min = -(2 ^ (TrimCodeSize - 1))
        TrimCodeValue_Max = (2 ^ (TrimCodeSize - 1)) - 1
    Else
        TrimCodeValue_Min = 0
        TrimCodeValue_Max = (2 ^ TrimCodeSize) - 1
    End If
    
    Dim binaryFlag As Boolean
    Dim temp_assignment As String
    temp_assignment = DigSrc_Assignment
     
    '''''''''''''''''''''''''''''''''''Process Trim Method'''''''''''''''''''''''''''''''''''''
    If TrimMethod = 0 Then
        binaryFlag = False
        doallFlag = False
    ElseIf TrimMethod = 2 Then
        binaryFlag = False
        doallFlag = True
        
    Else
        binaryFlag = True
    End If
    
    '''''''''''''''''''''''''''''''''''Linear Search'''''''''''''''''''''''''''''''''''''''''''
    If binaryFlag = False Then
    
        code = TrimStart
        'If doallFlag = True Then
        TrimStep = TrimCodeValue_Max - TrimCodeValue_Min
        'Else
            'TrimStep = TrimCodeValue_Max - TrimCodeValue_Min
        'End If
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site" & site & " decimal code is " & code
        Next site
        
        Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, CStr(MeasV_Pins), MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc)
        
        
        If trace Then TheExec.Flow.TestLimit code
        If trace Then TheExec.Flow.TestLimit MeasValue
        First = MeasValue.compare(LessThan, TrimTarget)
        If trace Then TheExec.Flow.TestLimit First, , , , , , , , "first"
        BestCode = code
        BestVal = MeasValue
        verr = MeasValue.Subtract(TrimTarget).Abs
    
        For i = 0 To TrimStep - 1
            'DoneEndTrim = False
            DigSrc_Assignment = temp_assignment
            If trace Then TheExec.Datalog.WriteComment ("i = " & i)
            If doallFlag = True Then
            
            Else
                ADCOUT = verr.compare(LessThan, Diffaccuracy)
                If ADCOUT.All(True) Then Exit For
            End If
            
            
            If doallFlag = True Then
                code = code.Add(1) 'do all should set start value to smallest value
            Else
                Temp = MeasValue.compare(LessThan, TrimTarget)
                If Antitrim = True Then
                    code = code.Add(Temp.Multiply(2).Add(1))
                Else
                    code = code.Add(Temp.Multiply(-2).Subtract(1))  ' If MeasValue < TrimTarget Then code++ Else code--
                End If
                
            End If
            
            For Each site In TheExec.sites.Active
                If code(site) > TrimCodeValue_Max Then
                    code(site) = code(site) - 1
                    DoneEndTrim = True
                End If ''GoTo EndTrim
                If code(site) < TrimCodeValue_Min Then
                    code(site) = code(site) + 1
                    DoneEndTrim = True
                End If ''GoTo EndTrim
            Next site
          
          
          
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site" & site & " decimal code is " & code
            Next site
            Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, CStr(MeasV_Pins), MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc)
            
            If Diffaccuracy <> 0 And TheExec.TesterMode = testModeOffline And i = 3 Then
                For Each site In TheExec.sites
                    MeasValue(site) = TrimTarget - Diffaccuracy + 0.0003
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Voltage = " & MeasValue(site)
                Next site
            End If
    
            If trace Then TheExec.Flow.TestLimit code
            If trace Then TheExec.Flow.TestLimit MeasValue
            For Each site In TheExec.sites
                If Abs(MeasValue - TrimTarget) < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = Abs(MeasValue - TrimTarget)
                End If
            Next site
            
            If doallFlag = True Then
            Else
                Done = Done.LogicalOr(First.LogicalXor(MeasValue.compare(LessThan, TrimTarget)))
                If trace Then TheExec.Flow.TestLimit Done, , , , , , , , "done"
                DoneEndTrim = DoneEndTrim.LogicalOr(Done)
                'If Done.All(True) Then Exit For
                If DoneEndTrim.All(True) Then Exit For
            End If
            
        Next i
    
    '''''''''''''''''''''''''''''''''''Binary Search'''''''''''''''''''''''''''''''''''''''''''
    Else
        Dim counter As Long
        Dim trimmax As New SiteLong
        Dim trimmin As New SiteLong
        
        trimmax = TrimCodeValue_Max
        trimmin = TrimCodeValue_Min
        
        counter = 0
        code = (trimmax + trimmin) / 2
        
        Do While counter < TrimCodeSize
            DigSrc_Assignment = temp_assignment
            
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site" & site & " decimal code is " & code
            Next site
            Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, CStr(MeasV_Pins), MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc)
            If counter = 0 Then
            verr = MeasValue.Subtract(TrimTarget).Abs
            BestCode = code
            BestVal = MeasValue
            End If

            
            For Each site In TheExec.sites
                If MeasValue.Subtract(TrimTarget).Abs < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = MeasValue.Subtract(TrimTarget).Abs
                End If
            Next site
            
            For Each site In TheExec.sites
                If Antitrim = True Then
                    If MeasValue(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site) - 1
                        Else
                            trimmax(site) = code(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            'code(Site) = code(Site)
                            code(site) = code(site) + 1
                        Else
                            trimmin(site) = code(site)
                        End If
                    End If
                Else
                    If MeasValue(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            'code(Site) = code(Site)
                            code(site) = code(site) + 1
                        Else
                            trimmin(site) = code(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site) - 1
                        Else
                            trimmax(site) = code(site)
                        End If
                    End If
                End If
            Next site
            
            If counter = TrimCodeSize - 1 Then
            Else
                code = trimmax.Add(trimmin).Divide(2)
            End If
            
        
            counter = counter + 1
        Loop
        DigSrc_Assignment = temp_assignment
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site" & site & " decimal code is " & code
        Next site
        Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, CStr(MeasV_Pins), MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc)
        
            For Each site In TheExec.sites
                If MeasValue.Subtract(TrimTarget).Abs < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = MeasValue.Subtract(TrimTarget).Abs
                End If
            Next site

    End If
    
    
    finalflag = True
    
    
EndTrim:

    If Interpose_PostTest <> "" Then
        Call SetForceCondition(Interpose_PostTest & ";STOREPREPAT")
    End If
    

    
    If MeasV_Pins <> "" Then
        TestNameInput = Report_TName_From_Instance("V", MeasV_Pins.Value, "TrimmedVoltage", 0, 0)
        If InStr(TheExec.DataManager.instanceName, "T4P2") <> 0 Then
            ReCalc = BestVal.Add(1).Multiply(0.7975).Add(0.4)
            OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit ReCalc, , , , , , unitVolt, , TestNameInput, , MeasV_Pins.Value, , , , , tlForceFlow
            OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , TestNameInput, , , , , , , tlForceFlow
        Else
            TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , TestNameInput, , MeasV_Pins.Value, , , , , tlForceFlow
        End If
    ElseIf MeasI_pinS <> "" Then
        TestNameInput = Report_TName_From_Instance("I", MeasI_pinS.Value, "TrimmedVoltage", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitAmp, , TestNameInput, , MeasI_pinS.Value, , , , , tlForceFlow
    ElseIf MeasF_PinS_SingleEnd <> "" Then
        TestNameInput = Report_TName_From_Instance("F", MeasF_PinS_SingleEnd.Value, "TrimmedFrequency", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitHz, , TestNameInput, , MeasI_pinS.Value, , , , , tlForceFlow
    Else
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.Value, "TrimmedCode(Decimal)", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitNone, , TestNameInput, , DigCap_Pin, , , , , tlForceFlow
    End If
    TestNameInput = Report_TName_From_Instance("C", "x", , 0, 0)
    TheExec.Flow.TestLimit BestCode, TrimCodeValue_Min, TrimCodeValue_Max, , , , , , TestNameInput, , , , , , , tlForceNone
    
    
    
    Dim TempVal As Integer
    Dim FinalTrimCode As New DSPWave
    Dim Binary_FinalTrimCode As String
    FinalTrimCode.CreateConstant 0, TrimCodeSize
    
    For Each site In TheExec.sites
'        TempVal = BestCode(Site)
        TheExec.Datalog.WriteComment "site" & site & " decimal best code is " & BestCode
        Binary_FinalTrimCode = ""
        For i = 0 To TrimCodeSize - 1
            FinalTrimCode(site).Element(i) = ((BestCode(site) And (2 ^ i)) \ (2 ^ i))
            Binary_FinalTrimCode = Binary_FinalTrimCode & FinalTrimCode(site).Element(i)
'            If i = 0 Then
'                code_bin(Site) = CStr(code(Site) And 1)
'            Else
'                code_bin(Site) = code_bin(Site) & CStr((code(Site) And (2 ^ i)) \ (2 ^ i))
'            End If
'            FinalTrimCode(Site).Element(i) = TempVal Mod 2
'            TempVal = TempVal \ 2
            
        Next i
        TheExec.Datalog.WriteComment "site" & site & " Binary best code is " & Binary_FinalTrimCode
    Next site
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, FinalTrimCode)
    End If
    
    If TrimCal_Name <> "" Then
        If Final_Calc = True Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    Else
        If Calc_Eqn <> "" Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    End If
    
    HardIP_WriteFuncResult
    
    DebugPrintFunc Pat
    

    If TheExec.TesterMode = testModeOffline Then
    Else
    End If
        
    Exit Function
    
errHandler:
'       ByPassTestLimit = False
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search: " & err.Description)
    'TrimUVI80Code_VFI = TL_ERROR
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function TrimUVI80Code_VFI(Optional Pat As String, Optional TestSequence As String, Optional MeasV_Pins As String, Optional MeasI_pinS As PinList, Optional MeasI_Range As Double, _
    Optional MeasF_PinS_SingleEnd As PinList, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, _
    Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimFormat As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_DigCapData As String = "", _
    Optional TrimStoreName As String, Optional Diffaccuracy As Double, Optional TrimMethod As Long, _
    Optional Meas_StoreName As String, Optional Calc_Eqn As String, Optional TrimCal_Name As String, Optional Antitrim As Boolean = False, Optional Interpose_PrePat As String, Optional CPUA_Flag_In_Pat As Boolean = True, Optional Interpose_PostTest As String, _
    Optional Final_Calc As Boolean = False, Optional BestMeasVal_StoreName As String, Optional Final_Calc_Eqn As String, Optional MSB_First_Flag As Boolean = False, Optional Validating_ As Boolean)
''    Dim PatCount As Long
    Dim i As Integer
    Dim pats() As String
    Dim code As New SiteLong
    Dim MeasValue As New SiteDouble
    Dim BestCode As New SiteLong, BestVal As New SiteDouble, verr As New SiteDouble, Temp As New SiteLong
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim trace As Boolean
    Dim site As Variant
    Dim Ts As Variant
    Dim ADCOUT As New SiteBoolean
    Dim TrimStep As Long
    Dim OutDSP As New DSPWave
    Dim PatCount As Long, PattArray() As String
    Dim TestSequence_array() As String
    Dim doallFlag As Boolean
    Dim finalflag As Boolean
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim TestNameInput As String
    Dim ReCalc As New SiteDouble
    
    gl_TName_Pat = Pat
    
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    Call GetFlowTName
    
'    If InStr(TheExec.DataManager.InstanceName, "MTRGR_T4P2") <> 0 Then
'        Cal_Eqn = "ALG::Calc_Metrology_GainError(GainError,V1)"
'        Meas_StoreName = "V1+"
'        TrimCal_Name = "GainError"
'    End If
    'ByPassTestLimit = True
    If Validating_ Then
        Call PrLoadPattern(Pat)
        Exit Function    ' Exit after validation
    End If
    
    
    Call HardIP_InitialSetupForPatgen

    TestSequence_array = Split(TestSequence, ",")
    
    
    On Error GoTo errHandler
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    PATT_GetPatListFromPatternSet Pat, pats, PatCount
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    
  '  If (UBound(TestSequence_array) > UBound(TName_Ary)) Then
  '      ReDim Preserve TName_Ary(UBound(TestSequence_array)) As String
  '
  '  End If
    
    
'    '''''''''''''''setup UVI80 for meas V''''''''''''''''''
'    With thehdw.DCVI.Pins(MeasV_PinS)
'            .Gate = False
'            .Disconnect tlDCVIConnectDefault
'            .mode = tlDCVIModeHighImpedance
'            .Connect tlDCVIConnectHighSense
'            .Voltage = 6
'            .Current = 0
'             thehdw.Wait 0.5 * ms
'            .Gate = True
'    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''Setup UVI180 for meas I'''''''''''''''''
    
    Dim SplitByEqual() As String, SplitByColon() As String, TrimCodeSize As Long
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Mid As Long, TrimCodeValue_Max As Long
    Dim Trimname As String
    SplitByEqual = Split(TrimFormat, "=")
    SplitByColon = Split(SplitByEqual(1), ":")
    Trimname = SplitByEqual(0)
    TrimCodeSize = SplitByColon(0) + 1
    TrimCodeValue_Min = 0
    TrimCodeValue_Mid = (2 ^ TrimCodeSize) / 2
    If SplitByColon(1) = 0 Then
        TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
    Else
        TrimCodeValue_Max = SplitByColon(1)
    End If
    Dim binaryFlag As Boolean
    Dim temp_assignment As String
    temp_assignment = DigSrc_Assignment
     
    '''''''''''''''''''''''''''''''''''Process Trim Method'''''''''''''''''''''''''''''''''''''
    If TrimMethod = 0 Then
        binaryFlag = False
        doallFlag = False
    ElseIf TrimMethod = 2 Then
        binaryFlag = False
        doallFlag = True
        
    Else
        binaryFlag = True
    End If
    
    '''''''''''''''''''''''''''''''''''Linear Search'''''''''''''''''''''''''''''''''''''''''''
    If binaryFlag = False Then
    
        code = TrimStart
        If doallFlag = True Then
            TrimStep = TrimCodeValue_Max - TrimStart
        Else
            TrimStep = TrimCodeValue_Max
        End If
        
        
        Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_Pins, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc, MSB_First_Flag:=MSB_First_Flag)
        
        
    '    If Diffaccuracy <> 0 And TheExec.TesterMode = testModeOffline Then
    '        For Each Site In TheExec.sites
    '            MeasValue(Site) = TrimTarget + Diffaccuracy - 0.0003
    '            TheExec.Datalog.WriteComment "Site " & Site & ",Code " & code(Site) & ", Voltage = " & MeasValue(Site)
    '        Next Site
    '    End If
        If trace Then TheExec.Flow.TestLimit code
        If trace Then TheExec.Flow.TestLimit MeasValue
        First = MeasValue.compare(LessThan, TrimTarget)
        If trace Then TheExec.Flow.TestLimit First, , , , , , , , "first"
        BestCode = code
        BestVal = MeasValue
        verr = MeasValue.Subtract(TrimTarget).Abs
    
        For i = 0 To TrimStep - 1
            DigSrc_Assignment = temp_assignment
            If trace Then TheExec.Datalog.WriteComment ("i = " & i)
            If doallFlag = True Then
            
            Else
                ADCOUT = verr.compare(LessThan, Diffaccuracy)
                If ADCOUT.All(True) Then Exit For
            End If
            
            If doallFlag = True Then
                code = code.Add(1)
            Else
                Temp = MeasValue.compare(LessThan, TrimTarget)
                If Antitrim = True Then
                    code = code.Add(Temp.Multiply(2).Add(1))
                Else
                    code = code.Add(Temp.Multiply(-2).Subtract(1))  ' If MeasValue < TrimTarget Then code++ Else code--
                End If
                
            End If
            
            For Each site In TheExec.sites.Active
                If code(site) > TrimCodeValue_Max Then: code(site) = code(site) - 1: ''GoTo EndTrim
                If code(site) < TrimCodeValue_Min Then: code(site) = code(site) + 1: ''GoTo EndTrim
            Next site
          
            Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_Pins, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc, MSB_First_Flag:=MSB_First_Flag)
            
            If Diffaccuracy <> 0 And TheExec.TesterMode = testModeOffline And i = 3 Then
                For Each site In TheExec.sites
                    MeasValue(site) = TrimTarget - Diffaccuracy + 0.0003
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Voltage = " & MeasValue(site)
                Next site
            End If
    
            If trace Then TheExec.Flow.TestLimit code
            If trace Then TheExec.Flow.TestLimit MeasValue
            For Each site In TheExec.sites
                If Abs(MeasValue - TrimTarget) < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = Abs(MeasValue - TrimTarget)
                End If
            Next site
            
            If doallFlag = True Then
            Else
                Done = Done.LogicalOr(First.LogicalXor(MeasValue.compare(LessThan, TrimTarget)))
                If trace Then TheExec.Flow.TestLimit Done, , , , , , , , "done"
                If Done.All(True) Then Exit For
            End If
        Next i
    
    '''''''''''''''''''''''''''''''''''Binary Search'''''''''''''''''''''''''''''''''''''''''''
    Else
        Dim counter As Long
        Dim trimmax As New SiteLong
        Dim trimmin As New SiteLong
        
        trimmax = TrimCodeValue_Max
        trimmin = TrimCodeValue_Min
        
        counter = 0
        code = (trimmax + trimmin) / 2
        
        Do While counter < TrimCodeSize
            DigSrc_Assignment = temp_assignment
            
            Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_Pins, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc, MSB_First_Flag:=MSB_First_Flag)
            If counter = 0 Then
            verr = MeasValue.Subtract(TrimTarget).Abs
            BestCode = code
            BestVal = MeasValue
            End If

            
            
            For Each site In TheExec.sites
                If MeasValue.Subtract(TrimTarget).Abs < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = MeasValue.Subtract(TrimTarget).Abs
                End If
            Next site
            
            For Each site In TheExec.sites
                If Antitrim = True Then
                    If MeasValue(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site) - 1
                        Else
                            trimmax(site) = code(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site)
                        Else
                            trimmin(site) = code(site)
                        End If
                    End If
                Else
                    If MeasValue(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site)
                        Else
                            trimmin(site) = code(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site) - 1
                        Else
                            trimmax(site) = code(site)
                        End If
                    End If
                End If
            Next site
            
            If counter = TrimCodeSize - 1 Then
            Else
                code = trimmax.Add(trimmin).Divide(2)
            End If
            
        
            counter = counter + 1
        Loop
        DigSrc_Assignment = temp_assignment
        Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_Pins, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc, MSB_First_Flag:=MSB_First_Flag)
        
            For Each site In TheExec.sites
                If MeasValue.Subtract(TrimTarget).Abs < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = MeasValue.Subtract(TrimTarget).Abs
                End If
            Next site

    End If
    
    
    finalflag = True
    
    'Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, BestCode, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodesize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, finalflag)
    
    
EndTrim:

    If Interpose_PostTest <> "" Then
        Call SetForceCondition(Interpose_PostTest & ";STOREPREPAT")
    End If
    
    Dim MeasV_Split() As String
    MeasV_Split = Split(MeasV_Pins, "+")
    
    If MeasV_Pins <> "" Then
        TestNameInput = Report_TName_From_Instance("V", "", "TrimmedVoltage", 0, 0)
        If InStr(TheExec.DataManager.instanceName, "T4P2") <> 0 Then
            ReCalc = BestVal.Add(1).Multiply(0.7975).Add(0.4)
            'OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            'TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit ReCalc, , , , , , unitVolt, , TestNameInput, , MeasV_Split(0), , , , , tlForceFlow
            'OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            'TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , TestNameInput, , , , , , , tlForceFlow
        Else
            TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , TestNameInput, , MeasV_Split(0), , , , , tlForceFlow
        End If
    ElseIf MeasI_pinS <> "" Then
        TestNameInput = Report_TName_From_Instance("I", "", "TrimmedVoltage", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitAmp, , TestNameInput, , MeasI_pinS.Value, , , , , tlForceFlow
    ElseIf MeasF_PinS_SingleEnd <> "" Then
        TestNameInput = Report_TName_From_Instance("F", "", "TrimmedFrequency", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitHz, , TestNameInput, , MeasI_pinS.Value, , , , , tlForceFlow
    Else
        TestNameInput = Report_TName_From_Instance("C", "", "TrimmedCode(Decimal)", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitNone, , TestNameInput, , DigCap_Pin, , , , , tlForceFlow
    End If
        
    If BestMeasVal_StoreName <> "" Then         'Cebu MTRG GR t1p1 store best meas value 20180806
        Call AddStoredMeasurement(BestMeasVal_StoreName, BestVal)
    End If
    
    TestNameInput = Report_TName_From_Instance("C", "", "TrimmedCode", 0, 0)
    TheExec.Flow.TestLimit BestCode, 0, 2 ^ TrimCodeSize, , , , , , TestNameInput, , , , , , , tlForceNone
    'ByPassTestLimit = False
    If DigCap_Sample_Size <> 0 Then
        Dim DigCapPinAry() As String, NumberPins As Long
        
        'Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, BestCode, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodesize, Trimname, Meas_StoreName, Cal_Eqn, TrimCal_Name)
        Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
        
        If NumberPins > 1 Then
            'Call CreateSimulateDataDSPWave_Parallel(OutDSP, DigCap_Sample_Size)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP, NumberPins)
            Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, NumberPins)

        ElseIf NumberPins = 1 Then
            'Call CreateSimulateDataDSPWave(OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP, NumberPins)
            Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
        End If
    End If
    
    
    Dim TempVal As Integer
    Dim FinalTrimCode As New DSPWave
    Dim trimvalue As String
    
    FinalTrimCode.CreateConstant 0, TrimCodeSize
    
    For Each site In TheExec.sites
        TempVal = BestCode(site)
        trimvalue = ""
        For i = 0 To TrimCodeSize - 1
            FinalTrimCode(site).Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
            trimvalue = trimvalue & CStr(FinalTrimCode(site).Element(i))
            
        Next i
        'theexec.Datalog.WriteComment "Site==> " & Site & ", Trimed code==>" & trimvalue
        
        If UCase(TheExec.DataManager.instanceName) Like "*MTRGRT1P1*" And FinalTrimCode(site).Element(TrimCodeSize - 1) = 0 Then
                FinalTrimCode(site).Element(TrimCodeSize - 1) = 1
        ElseIf UCase(TheExec.DataManager.instanceName) Like "*MTRGRT1P1*" And FinalTrimCode(site).Element(TrimCodeSize - 1) = 1 Then
                FinalTrimCode(site).Element(TrimCodeSize - 1) = 0
        End If
    Next site
    
    If TrimStoreName <> "" Then
        
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, FinalTrimCode)
        
        Dim negate_FinalTrimCode As New DSPWave
        negate_FinalTrimCode = FinalTrimCode
        For Each site In TheExec.sites
            If negate_FinalTrimCode.Element(0) = 1 Then
                negate_FinalTrimCode.Element(0) = 0
            Else
                negate_FinalTrimCode.Element(0) = 1
            End If
        Next site
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName + "_negate", FinalTrimCode)

    End If
    
    If TrimCal_Name <> "" Then
        If Final_Calc = True Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    Else
        If Calc_Eqn <> "" Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    End If
    
    If Final_Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Final_Calc_Eqn)
    End If
    
    Call HardIP_WriteFuncResult(, , Inst_Name_Str)
    
    DebugPrintFunc Pat
    
    '' 20170704 - Add write efuse function
''    Dim sl_Fuse_Val As SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else
''        For Each Site In TheExec.sites
''            sl_Fuse_Val(Site) = BestCode(Site)
''        Next Site
        
'        If TrimFuseName <> "" And TrimFuseTypeName <> "" Then
'            Call HIP_eFuse_Write(TrimFuseTypeName, TrimFuseName, BestCode)
'        End If
    End If
        
    Exit Function
    
errHandler:
'       ByPassTestLimit = False
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search: " & err.Description)
    'TrimUVI80Code_VFI = TL_ERROR
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TrimUVI80Code_VFI_ADC(Optional Pat As String, Optional TestSequence As String, Optional MeasV_Pins As PinList, Optional MeasI_pinS As PinList, Optional MeasI_Range As Double, _
    Optional MeasF_PinS_SingleEnd As PinList, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, _
    Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimFormat As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_DigCapData As String = "", _
    Optional TrimStoreName As String, Optional Diffaccuracy As Double, Optional TrimMethod As Long, _
    Optional Meas_StoreName As String, Optional Calc_Eqn As String, Optional TrimCal_Name As String, Optional Antitrim As Boolean = False, Optional Validating_ As Boolean, Optional Interpose_PrePat As String, Optional CPUA_Flag_In_Pat As Boolean = True, Optional Interpose_PostTest As String, _
    Optional Final_Calc As Boolean = False, Optional TrimPredictStep As Double, Optional NolessthanTarget As Boolean = False)
''    Dim PatCount As Long
    Dim i, j As Integer
    Dim pats() As String
    Dim code() As New SiteLong
    Dim MeasValue As New SiteDouble
    Dim BestCode() As New SiteLong, BestVal() As New SiteDouble, verr() As New SiteDouble, Temp As New SiteLong
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim trace As Boolean
    Dim site As Variant
    Dim Ts As Variant
    Dim ADCOUT As New SiteBoolean
    Dim TrimStep As Long
    Dim OutDSP As New DSPWave
    Dim PatCount As Long, PattArray() As String
    Dim TestSequence_array() As String
    Dim doallFlag As Boolean
    Dim finalflag As Boolean
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim TestNameInput As String
    Dim ReCalc As New SiteDouble
    Dim predictFlag As Boolean
    Dim linearFlag As Boolean
    Dim Firstvoltage As New SiteDouble
    Dim lessboolen As New SiteBoolean
    Dim lessboolen_need_change As New SiteBoolean
    
    Dim Temp_sub As Double
    Dim MeasSeqAry() As New SiteDouble
    Dim MeasSeqAry_Best() As New SiteDouble
    
    gl_TName_Pat = Pat

    Call GetFlowTName

    If Validating_ Then
        Call PrLoadPattern(Pat)
        Exit Function    ' Exit after validation
    End If
    Call HardIP_InitialSetupForPatgen
    TestSequence_array = Split(TestSequence, ",")
    
    
    On Error GoTo errHandler
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    PATT_GetPatListFromPatternSet Pat, pats, PatCount
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    

    
    Dim SplitByEqual() As String, SplitByColon() As String, TrimCodeSize As Long
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Mid As Long, TrimCodeValue_Max As Long
    Dim Trimname() As String
    Dim SplitByComma() As String
    SplitByComma = Split(TrimFormat, ";")
    ReDim Trimname(UBound(SplitByComma))
    ReDim code(UBound(SplitByComma))
    ReDim BestCode(UBound(SplitByComma))
    ReDim BestVal(UBound(SplitByComma))
    ReDim verr(UBound(SplitByComma))
    TrimCodeValue_Min = 0
    For i = 0 To UBound(SplitByComma)
        SplitByEqual = Split(SplitByComma(i), "=")
        SplitByColon = Split(SplitByEqual(1), ":")
        Trimname(i) = SplitByEqual(0)
        TrimCodeSize = SplitByColon(0) + 1
        TrimCodeValue_Mid = (2 ^ TrimCodeSize) / 2
        If SplitByColon(1) = 0 Then
            TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
        Else
            TrimCodeValue_Max = SplitByColon(1)
        End If
    Next i
    Dim binaryFlag As Boolean
    Dim temp_assignment As String
    temp_assignment = DigSrc_Assignment
     
    '''''''''''''''''''''''''''''''''''Process Trim Method'''''''''''''''''''''''''''''''''''''
    Select Case TrimMethod
        Case 0
            linearFlag = True
            doallFlag = False
        Case 1
            binaryFlag = True
        Case 2
            linearFlag = True
            doallFlag = True
        Case 3
            predictFlag = True
            linearFlag = True
    End Select
    
    
    '''''''''''''''''''''''''''''''''''Binary Search'''''''''''''''''''''''''''''''''''''''''''
    If binaryFlag = True Then
        Dim counter As Long
        Dim trimmax() As New SiteLong
        Dim trimmin() As New SiteLong
        ReDim trimmax(UBound(SplitByComma))
        ReDim trimmin(UBound(SplitByComma))
        For i = 0 To UBound(SplitByComma)
        trimmax(i) = TrimCodeValue_Max
        trimmin(i) = TrimCodeValue_Min
        
        counter = 0
        code(i) = (trimmax(i) + trimmin(i)) / 2
        Next i
        Do While counter < TrimCodeSize
            DigSrc_Assignment = temp_assignment
            
            Call TrimUVI80_Meas_VFI_ADC(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_Pins, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, MeasSeqAry, Final_Calc)
            For i = 0 To UBound(SplitByComma)
            
            If counter = 0 Then
            verr(i) = MeasSeqAry(i).Subtract(TrimTarget).Abs
            BestCode(i) = code(i)
            BestVal(i) = MeasSeqAry(i)
            End If
            
            For Each site In TheExec.sites
                If MeasSeqAry(i).Subtract(TrimTarget).Abs < verr(i) Then
                    BestCode(i) = code(i)
                    BestVal(i) = MeasSeqAry(i)
                    verr(i) = MeasSeqAry(i).Subtract(TrimTarget).Abs
'                    For j = 0 To UBound(MeasSeqAry)
'                        MeasSeqAry_Best(j) = MeasSeqAry(j)
'                    Next j
                    'TheExec.DataLog.WriteComment "Site : " & Site & ",Best Value : " & BestVal(i)
                    'TheExec.DataLog.WriteComment "Site : " & Site & ",Best Code : " & BestCode(i)
                End If
            Next site
            
            For Each site In TheExec.sites
                If Antitrim = True Then
                    If MeasSeqAry(i)(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(i)(site) = code(i)(site) - 1
                        Else
                            trimmax(i)(site) = code(i)(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(i)(site) = code(i)(site)
                        Else
                            trimmin(i)(site) = code(i)(site)
                        End If
                    End If
                Else
                    If MeasSeqAry(i)(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(i)(site) = code(i)(site)
                        Else
                            trimmin(i)(site) = code(i)(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(i)(site) = code(i)(site) - 1
                        Else
                            trimmax(i)(site) = code(i)(site)
                        End If
                    End If
                End If
            Next site
            
            
            If counter = TrimCodeSize - 1 Then
            Else
                code(i) = trimmax(i).Add(trimmin(i)).Divide(2)
            End If
            
            
            Next i
            counter = counter + 1
        Loop
        
        DigSrc_Assignment = temp_assignment
        Call TrimUVI80_Meas_VFI_ADC(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_Pins, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, MeasSeqAry, Final_Calc)
        For i = 0 To UBound(SplitByComma)
            For Each site In TheExec.sites
                    If MeasSeqAry(i).Subtract(TrimTarget).Abs < verr(i) Then
                        BestCode(i) = code(i)
                        BestVal(i) = MeasSeqAry(i)
                        verr(i) = MeasSeqAry(i).Subtract(TrimTarget).Abs
                    End If
                Next site
        Next i
            
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
    
    
    finalflag = True
    
    'Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, BestCode, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodesize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, finalflag)
    
    
EndTrim:

    If Interpose_PostTest <> "" Then
        Call SetForceCondition(Interpose_PostTest & ";STOREPREPAT")
    End If
    
    For i = 0 To UBound(SplitByComma)
    If MeasV_Pins <> "" Then

        TestNameInput = Report_TName_From_Instance("V", "", "TrimmedVoltage", 0, CLng(i), "3=" & CStr(i))

        If InStr(TheExec.DataManager.instanceName, "T4P2") <> 0 Then
            ReCalc = BestVal(i).Add(1).Multiply(0.7975).Add(0.4)
            'OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit ReCalc, , , , , , unitVolt, , TestNameInput, , MeasV_Pins.Value, , , , , tlForceFlow
            'OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit BestVal(i), , , , , , unitVolt, , TestNameInput, , , , , , , tlForceFlow
        Else
            TheExec.Flow.TestLimit BestVal(i), , , , , , unitVolt, , TestNameInput, , MeasV_Pins.Value, , , , , tlForceFlow
        End If
    ElseIf MeasI_pinS <> "" Then

        TestNameInput = Report_TName_From_Instance("I", "", "TrimmedVoltage", 0, CLng(i), "3=" & CStr(i))

        TheExec.Flow.TestLimit BestVal(i), , , , , , unitAmp, , TestNameInput, , MeasI_pinS.Value, , , , , tlForceFlow
    ElseIf MeasF_PinS_SingleEnd <> "" Then

        TestNameInput = Report_TName_From_Instance("F", "", "TrimmedFrequency", 0, CLng(i), "3=" & CStr(i))
        TheExec.Flow.TestLimit BestVal(i), , , , , , unitHz, , TestNameInput, , MeasI_pinS.Value, , , , , tlForceFlow
    Else

        TestNameInput = Report_TName_From_Instance("C", "", "TrimmedCode(Decimal)", 0, CLng(i), "3=" & CStr(i))

        TheExec.Flow.TestLimit BestVal, , , , , , unitNone, , TestNameInput, , DigCap_Pin, , , , , tlForceFlow
    End If

        TestNameInput = Report_TName_From_Instance("C", "", "TrimmedCode", 0, CLng(i), "3=" & CStr(i))

    TheExec.Flow.TestLimit BestCode(i), 0, 2 ^ TrimCodeSize, , , , , , TestNameInput, , , , , , , tlForceNone
    'ByPassTestLimit = False
    Next i
    If DigCap_Sample_Size <> 0 Then
        Dim DigCapPinAry() As String, NumberPins As Long
        
        'Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, BestCode, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodesize, Trimname, Meas_StoreName, Cal_Eqn, TrimCal_Name)
        Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
        
        If NumberPins > 1 Then
            'Call CreateSimulateDataDSPWave_Parallel(OutDSP, DigCap_Sample_Size)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP, NumberPins)
            Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, NumberPins)

        ElseIf NumberPins = 1 Then
            'Call CreateSimulateDataDSPWave(OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP, NumberPins)
            Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
        End If
    End If
    

    Dim TempVal As Integer
    Dim FinalTrimCode() As New DSPWave
    Dim SplitTrimStoreName() As String
    SplitTrimStoreName = Split(TrimStoreName, "+")
    
    'FinalTrimCode.CreateConstant 0, TrimCodeSize
    ReDim FinalTrimCode(UBound(SplitTrimStoreName)) As New DSPWave
    For i = 0 To UBound(SplitTrimStoreName)
    FinalTrimCode(i).CreateConstant 0, TrimCodeSize
    For Each site In TheExec.sites
        TempVal = BestCode(i)(site)
        For j = 0 To TrimCodeSize - 1
            FinalTrimCode(i).Element(j) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next j
    Next site
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(SplitTrimStoreName(i), FinalTrimCode(i))
    End If
    
    Next i
    If TrimCal_Name <> "" Then
        If Final_Calc = True Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    Else
        If Calc_Eqn <> "" Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    End If
    
    
    HardIP_WriteFuncResult
    DebugPrintFunc Pat
    
    '' 20170704 - Add write efuse function
''    Dim sl_Fuse_Val As SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else
''        For Each Site In TheExec.sites
''            sl_Fuse_Val(Site) = BestCode(Site)
''        Next Site
        
'        If TrimFuseName <> "" And TrimFuseTypeName <> "" Then
'            Call HIP_eFuse_Write(TrimFuseTypeName, TrimFuseName, BestCode)
'        End If
    End If
    
    
    
        
    Exit Function
    
errHandler:
'       ByPassTestLimit = False
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search: " & err.Description)
    'TrimUVI80Code_VFI = TL_ERROR
    If AbortTest Then Exit Function Else Resume Next
End Function





Public Function TrimCodeFreq_new(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double = 1000000, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PreMeas As String, Optional Validating_ As Boolean, Optional Interpose_PrePat As String) As Long
    
    Dim PatCount As Long, PattArray() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen
    Dim Ts As Variant, TestSequenceArray() As String
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPwave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long, p As Long
    
    Dim MeasureFreq As New PinListData, MeasureFreq_F1 As New PinListData, MeasureFreq_F2 As New PinListData
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName

    On Error GoTo ErrorHandler
         
    Call GetFlowTName
         
    ''Update Interpose_PreMeas 20170801
    Dim Interpose_PreMeas_Ary() As String
    ''20160923 - Analyze Interpose_PreMeas to force setting with different sequence.
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    If (UBound(TestSequenceArray) > UBound(TName_Ary)) Then
        ReDim Preserve TName_Ary(UBound(TestSequenceArray)) As String
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    TheHdw.Patterns(patset).Load
    gl_TName_Pat = patset.Value

    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)

    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    Dim d_MeasF_Interval  As Double
    d_MeasF_Interval = 0.01
    
    Dim b_HighThanTargetFreq As New SiteBoolean
    b_HighThanTargetFreq = False
    
    Dim OutputTrimCode As String
    Dim TestLimitIndex As Long, LastSectionF1F2_Index As Long
    LastSectionF1F2_Index = 0


    ''==================================================================================================
'    Dim TrimStart_1st() As String
    Dim Dec_TrimStart_1st As Long
    
    '' 20160706 Create value for final frequency
    Dim b_DefineFinalFreq As New SiteBoolean
    Dim FinalFreq As New PinListData
    
    ''20160712 - If match taget freq just store the trim code
    Dim b_MatchTagetFreq As New SiteBoolean
    Dim b_DisplayFreq As New SiteBoolean
    Dim StoredTargetTrimCode As New DSPWave
    b_MatchTagetFreq = False
    b_DisplayFreq = False
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    Dim StoreEachTrimFreq() As New PinListData
    Dim StoreEachTrimCode() As New DSPWave
    ReDim StoreEachTrimFreq(DigSrc_Sample_Size + 1) As New PinListData
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i

    ''20170721-Updated the TrimStart when the first bit is zero and seperate with "&"
    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", "")
    End If
    
'    Dim z As Integer
'        For z = 0 To 1
'
'        ''    TrimStart_1st = TrimStart
'            If z = 0 Then
'                Dec_TrimStart_1st = Bin2Dec(TrimStart)
'            Else
'                Dec_TrimStart_1st = Dec_TrimStart_1st + 4080
'            End If
'            InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong
'
'            Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, DigSrc_Sample_Size, InDSPWave)
'
'            Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPWave)
'
'            If gl_Disable_HIP_debug_log = False Then theexec.Datalog.WriteComment ("First Time Setup")
'            '' Debug use
'            For Each site In theexec.sites
'                OutputTrimCode = ""
'                For k = 0 To InDSPWave(site).SampleSize - 1
'                    OutputTrimCode = OutputTrimCode & CStr(InDSPWave(site).Element(k))
'                Next k
'
'                If gl_Disable_HIP_debug_log = False Then theexec.Datalog.WriteComment ("Site_" & site & " Coarse Output Trim Code = " & OutputTrimCode)
'            Next site
'
'            For Each site In theexec.sites
'                StoreEachTrimCode(0)(site).Data = InDSPWave(site).Data
'            Next site
'
'            Call thehdw.Patterns(PattArray(0)).start
'
            ''Update Interpose_PreMeas 20170801
            Dim TestSeqNum As Integer
            TestSeqNum = 0
'
'            For Each Ts In TestSequenceArray
'                If (CPUA_Flag_In_Pat) Then
'                    Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)
'                Else
'                    Call thehdw.Digital.Patgen.HaltWait
'                End If
'
'                ''Update Interpose_PreMeas 20170801
'                ''20160923 - Add Interpose_PreMeas entry point by each sequence
'                If Interpose_PreMeas <> "" Then
'                    If UBound(Interpose_PreMeas_Ary) = 0 Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
'                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
'                    End If
'                End If
'
'                If UCase(Ts) = "F" Then
'                    Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
'                    Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
'
'                    If theexec.TesterMode = testModeOffline Then
'                        Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
'                    End If
'                End If
'
'                ''Update Interpose_PreMeas 20170801
'                ''20161206-Restore force condiction after measurement
'                ''Call SetForceCondition("RESTORE")
'                If Interpose_PreMeas <> "" Then
'                    If UBound(Interpose_PreMeas_Ary) = 0 Then
'                        Call SetForceCondition("RESTOREPREMEAS")
'                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                        Call SetForceCondition("RESTOREPREMEAS")
'                    End If
'                End If
'
'                TestSeqNum = TestSeqNum + 1
'
'                If (CPUA_Flag_In_Pat) Then
'                    Call thehdw.Digital.Patgen.Continue(0, cpuA)
'                Else
'                    thehdw.Digital.Patgen.HaltWait
'                End If
'            Next Ts
'            thehdw.Digital.Patgen.HaltWait
'
'            StoreEachTrimFreq(z) = MeasureFreq
'
'            b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
'            PastDSPWave = InDSPWave
'
'            TestNameInput = "Freq_meas_"
'            TestLimitIndex = 0
'
'            '' 20160712 - Modify to use WriteComment to display output frequency.
'            If gl_Disable_HIP_debug_log = False Then
'                For Each site In theexec.sites
'                        theexec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
'                Next site
'            End If
'            '' 20160712 - Compare Measure Frequency whether match target Freq
'            b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
'
'            b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
'            For Each site In theexec.sites
'                If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
'                    StoredTargetTrimCode(site).Data = InDSPWave(site).Data
'                    b_StopTrimCodeProcess(site) = True
'                End If
'            Next site
'            If gl_Disable_HIP_debug_log = False Then theexec.Datalog.WriteComment ("======================================================================================")
'
'   Next z
            ''========================================================================================
    '20161128 Pre check Min/Max trim code process
    Dim b_KeepGoing As New SiteBoolean
    Dim PreviousFreq As New PinListData
'    If PreCheckMinMaxTrimCode = True Then
'        PreviousFreq = MeasureFreq
'        Call rundsp.PreCheckMinMaxTrimCode(b_HighThanTargetFreq, InDSPWave)
'        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPWave)
'
'        ''Update Interpose_PreMeas 20170801
'        TestSeqNum = 0
'
'        For Each Ts In TestSequenceArray
'            If (CPUA_Flag_In_Pat) Then
'                Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)
'            Else
'                Call thehdw.Digital.Patgen.HaltWait
'            End If
'
'            ''Update Interpose_PreMeas 20170801
'            ''20160923 - Add Interpose_PreMeas entry point by each sequence
'            If Interpose_PreMeas <> "" Then
'                If UBound(Interpose_PreMeas_Ary) = 0 Then
'                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
'                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
'                End If
'            End If
'
'            If UCase(Ts) = "F" Then
'                Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
'                Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
'
'                If theexec.TesterMode = testModeOffline Then
'                    Call SimulatePreCheckOutputFreq(MeasureF_Pin, MeasureFreq)
'                End If
'            End If
'
'            ''Update Interpose_PreMeas 20170801
'            ''20161206-Restore force condiction after measurement
'            ''Call SetForceCondition("RESTORE")
'            If Interpose_PreMeas <> "" Then
'                If UBound(Interpose_PreMeas_Ary) = 0 Then
'                    Call SetForceCondition("RESTOREPREMEAS")
'                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                    Call SetForceCondition("RESTOREPREMEAS")
'                End If
'            End If
'
'            TestSeqNum = TestSeqNum + 1
'
'            If (CPUA_Flag_In_Pat) Then
'                Call thehdw.Digital.Patgen.Continue(0, cpuA)
'            Else
'                thehdw.Digital.Patgen.HaltWait
'            End If
'        Next Ts
'
'        thehdw.Digital.Patgen.HaltWait
'
'        For Each site In theexec.sites
'            OutputTrimCode = ""
'            For k = 0 To InDSPWave(site).SampleSize - 1
'                OutputTrimCode = OutputTrimCode & CStr(InDSPWave(site).Element(k))
'            Next k
'            If gl_Disable_HIP_debug_log = False Then theexec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
'        Next site
'
'        If gl_Disable_HIP_debug_log = False Then
'            For Each site In theexec.sites
'                theexec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
'            Next site
'        End If
'        For Each site In theexec.sites
'            If b_HighThanTargetFreq(site) = True Then
'                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(LessThan, 0)
'            Else
'                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(GreaterThan, 0)
'            End If
'        Next site
'
'        Dim PreCheckBinStr As String, PreCheckDecVal As Double
'        For Each site In theexec.sites
'            If b_KeepGoing(site) = False Then
'                b_StopTrimCodeProcess(site) = True
'                PreCheckBinStr = ""
'                StoredTargetTrimCode(site).Data = InDSPWave(site).Data
'                For i = 0 To StoredTargetTrimCode(site).SampleSize - 1
'                    PreCheckBinStr = PreCheckBinStr & StoredTargetTrimCode.Element(i)
'                Next i
'                PreCheckDecVal = Bin2Dec_rev_Double(PreCheckBinStr)
'                ''TheExec.Flow.TestLimit PreCheckDecVal, 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TheExec.DataManager.InstanceName & "_TrimCode_Decimal", ForceResults:=tlForceNone
'            End If
'        Next site
'    End If

    ''========================================================================================
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    StoreEachIndex = 0
    
    Dim b_pos_trim As New SiteBoolean
    Dim b_neg_trim As New SiteBoolean
    Dim b_temp_0x00 As New SiteLong
    Dim b_temp_0xff As New SiteLong
    Dim fine_trim_flag As New SiteBoolean
    Dim coarse_trim_flag As New SiteBoolean
    Dim trim_store_temp As New DSPWave
    Dim InitialDSPWave_0x00 As New DSPWave
    Dim InitialDSPWave_0xff As New DSPWave
        InitialDSPWave_0x00.CreateConstant 0, 8, DspLong
        InitialDSPWave_0xff.CreateConstant 1, 8, DspLong
    fine_trim_flag = False
    
    ''20170103-Setup b_KeepGoing to true if PreCheckMinMaxTrimCode=false
    If PreCheckMinMaxTrimCode = False Then
        b_KeepGoing = True
    End If
    If b_KeepGoing.All(False) Then
    Else

        For i = 0 To ExecutionMax
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If
            StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
            StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
            SrcStartBit = StrSeparatebyColon(0)
            SrcEndBit = StrSeparatebyColon(1)
            If i = 0 Then
                b_FirstExecution = True
            Else
                b_FirstExecution = False
                'SrcStartBit = SrcStartBit + 1
            End If
            
                For j = 3 To 0 Step -1
                  If fine_trim_flag.All(True) Then Exit For
                    Dim z As Integer
                    For z = 0 To 1
                        If z = 0 Then
                            For Each site In TheExec.sites
                                InDSPwave.CreateConstant 0, 0, DspLong
                                If b_FirstExecution = True Then
                                    b_temp_0x00 = 8
                                ElseIf b_temp_0x00 = 15 Then
                                     fine_trim_flag = True
                                     
                                ElseIf b_FirstExecution = False And b_pos_trim(site) = True Then
                                    b_temp_0x00 = b_temp_0x00 + (2) ^ (j)
                                ElseIf b_FirstExecution = False And b_neg_trim(site) = True Then
                                    b_temp_0x00 = b_temp_0x00 - (2) ^ (j)
                                End If
                                InitialDSPWave.CreateConstant b_temp_0x00, 1, DspLong
                                Dim temp_bin() As String
                                 If b_StopTrimCodeProcess(site) = True Then
                                    'do nothing
                                 Else
                                  TheExec.Datalog.WriteComment ("Setting site=" & site & "_coarse trim=" & b_temp_0x00)
                                End If
                            Next site
                            If fine_trim_flag.All(True) Then Exit For
                                Call rundsp.CreateFlexibleDSPWave_lpro(InitialDSPWave, 4, InDSPwave, InitialDSPWave_0x00)
                                For Each site In TheExec.sites
                                    trim_store_temp.Data = InDSPwave.Data
                                Next site
                         Else
                            For Each site In TheExec.sites
                                InDSPwave.CreateConstant 0, 0, DspLong
                                If b_FirstExecution = True Then
                                    b_temp_0xff = b_temp_0x00
                                ElseIf b_temp_0xff = 15 Then
                                     fine_trim_flag = True
                                     
                                ElseIf b_FirstExecution = False And b_pos_trim = True Then
                                    b_temp_0xff = b_temp_0xff + (2) ^ (j)
                                ElseIf b_FirstExecution = False And b_neg_trim = True Then
                                    b_temp_0xff = b_temp_0xff - (2) ^ (j)
                                End If
                                InitialDSPWave.CreateConstant b_temp_0xff, 1, DspLong
                                If b_StopTrimCodeProcess(site) = True Then
                                    'do  nothing
                                 Else
                                   TheExec.Datalog.WriteComment ("Setting site=" & site & "_coarse trim=" & b_temp_0xff)
                                End If
                           Next site
                                Call rundsp.CreateFlexibleDSPWave_lpro(InitialDSPWave, 4, InDSPwave, InitialDSPWave_0xff)
                        End If
                        
                            Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
            
                                    
                        
                                    
                        For Each site In TheExec.sites
                            StoreEachTrimCode(StoreEachIndex)(site).Data = InDSPwave(site).Data
                        Next site

                            For Each site In TheExec.sites
                ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                                If b_StopTrimCodeProcess(site) = False Then
                                    OutputTrimCode = ""

                                    For k = 0 To InDSPwave(site).SampleSize - 1
                                        OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
                                    Next k

                                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
                                End If
                ''                End If
                            Next site
                            '' ==============================================================================================
    
                            Call TheHdw.Patterns(PattArray(0)).start
                            
                            ''Update Interpose_PreMeas 20170801
                            TestSeqNum = 0
                            
                            For Each Ts In TestSequenceArray
                                If (CPUA_Flag_In_Pat) Then
                                    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
                                Else
                                    Call TheHdw.Digital.Patgen.HaltWait
                                End If
                                
                                ''Update Interpose_PreMeas 20170801
                                ''20160923 - Add Interpose_PreMeas entry point by each sequence
                                If Interpose_PreMeas <> "" Then
                                    If UBound(Interpose_PreMeas_Ary) = 0 Then
                                        Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                                        Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                                    End If
                                End If
                                
                                
                                If UCase(Ts) = "F" Then
                                    Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                                    TheHdw.Digital.Pins(MeasureF_Pin).Levels.DriverMode = tlDriverModeVt
                                    Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
                                    
                                    If z = 0 Then
                                        MeasureFreq_F1 = MeasureFreq
                                    ElseIf z = 1 Then
                                        MeasureFreq_F2 = MeasureFreq
                                    End If
            
                                End If
                                
                                ''Update Interpose_PreMeas 20170801
                                ''20161206-Restore force condiction after measurement
                                ''Call SetForceCondition("RESTORE")
                                If Interpose_PreMeas <> "" Then
                                    If UBound(Interpose_PreMeas_Ary) = 0 Then
                                        Call SetForceCondition("RESTOREPREMEAS")
                                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                                        Call SetForceCondition("RESTOREPREMEAS")
                                    End If
                                End If
                                TestSeqNum = TestSeqNum + 1
                                
                                If (CPUA_Flag_In_Pat) Then
                                    Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                                Else
                                    TheHdw.Digital.Patgen.HaltWait
                                End If
                            Next Ts
                            
                            TheHdw.Digital.Patgen.HaltWait
                            StoreEachTrimFreq(StoreEachIndex) = MeasureFreq
                            StoreEachIndex = StoreEachIndex + 1
                                    
                            If gl_Disable_HIP_debug_log = False Then
                                For Each site In TheExec.sites
                    ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                                     If b_StopTrimCodeProcess(site) = False Then
                                        TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                                    End If
                    ''                End If
                                Next site
                            End If
    
                            If TrimPrcocessAll = False Then
                                If b_StopTrimCodeProcess.All(True) Then
                                    Exit For
                                End If
                            End If
                            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
                        Next z
                b_FirstExecution = False
                For Each site In TheExec.sites
                    If MeasureFreq_F1.Pins(0).Subtract(1000000).Value(site) <= 0 And MeasureFreq_F2.Pins(0).Subtract(1000000).Value(site) >= 0 Then
                         fine_trim_flag(site) = True
                         '---provide 0 or 1
                                b_HighThanTargetFreq = False
                                b_HighThanTargetFreq = MeasureFreq_F2.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureFreq_F1.Math.Subtract(TrimTarget).Abs)
                                PastDSPWave.Data = trim_store_temp.Data
                    ElseIf MeasureFreq_F1.Pins(0).Subtract(1000000).Value(site) >= 0 And MeasureFreq_F2.Pins(0).Subtract(1000000).Value(site) >= 0 Then
                        b_neg_trim(site) = True
                        b_pos_trim(site) = False
                        PastDSPWave.Data = trim_store_temp.Data
                    ElseIf MeasureFreq_F1.Pins(0).Subtract(1000000).Value(site) <= 0 And MeasureFreq_F2.Pins(0).Subtract(1000000).Value(site) <= 0 Then
                        b_pos_trim(site) = True
                        b_neg_trim(site) = False
                        PastDSPWave.Data = trim_store_temp.Data
                    End If
                Next site
                Next j
            '------------------------------fine trim
            
            StoreEachIndex = 0
            For j = SrcStartBit To SrcEndBit Step -1
                If i = 0 Then Exit For
               
                If b_FirstExecution = True Then
                    b_ControlNextBit = True
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                Else
                ''20160716-Control next bit to 1 no matter first or last progress
                    b_ControlNextBit = True
    ''                b_ControlNextBit = False
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                End If

                If b_FirstExecution = True And j = SrcEndBit Then
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPwave)
                Else
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                End If
                
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
                
                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site).Data = InDSPwave(site).Data
                Next site
            
                '' Debug use
                '' ==============================================================================================
                '' 20160716 - Modify trim code rule
                
                If gl_Disable_HIP_debug_log = False Then
                    If b_FirstExecution = True Then
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    Else
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j)
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    End If
                End If
                Dim fine_wave As New DSPWave
                Dim OutputTrimCode_fine As String
                Dim OutputTrimCode_fine_dec As New DSPWave
                    
                For Each site In TheExec.sites
                 fine_wave = InDSPwave.Select(0, , 8).Copy
    ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                    OutputTrimCode_fine_dec = fine_wave.ConvertStreamTo(tldspParallel, 8, 0, Bit0IsMsb)
                    If b_KeepGoing(site) = True Then
                        OutputTrimCode = ""
                        OutputTrimCode_fine = ""
                        For k = 0 To InDSPwave(site).SampleSize - 1
                            OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
                            If k < 8 Then
                            OutputTrimCode_fine = OutputTrimCode_fine & CStr(fine_wave(site).Element(k))
                            End If
                        Next k
                        TheExec.Datalog.WriteComment ("Site_" & site & "  Fine Trim Code = " & OutputTrimCode_fine & ", fine decimal= " & OutputTrimCode_fine_dec(site).Element(0))
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output All Code = " & OutputTrimCode)
                    End If
    ''                End If
                Next site
                
'
'                For Each site In TheExec.sites
'    ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
'                    If b_KeepGoing(site) = True Then
'                        OutputTrimCode = ""
'                        For k = 0 To InDSPWave(site).SampleSize - 1
'                            OutputTrimCode = OutputTrimCode & CStr(InDSPWave(site).Element(k))
'                        Next k
'
'                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output Fine Trim Code = " & OutputTrimCode)
'                    End If
'    ''                End If
'                Next site
'                '' ==============================================================================================
'
                Call TheHdw.Patterns(PattArray(0)).start
                
                ''Update Interpose_PreMeas 20170801
                TestSeqNum = 0
                
                For Each Ts In TestSequenceArray
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
                    Else
                        Call TheHdw.Digital.Patgen.HaltWait
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20160923 - Add Interpose_PreMeas entry point by each sequence
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                        End If
                    End If
                    
                    
                    If UCase(Ts) = "F" Then
                        Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                        Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
                        
                        '--------------- off line mode data --------
                        If TheExec.TesterMode = testModeOffline Then
                            Dim SimuIndex As Long
                            SimuIndex = TestLimitIndex
                            If SimuIndex >= 8 Then
                                SimuIndex = 8
                            End If
                            Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
                            MeasureFreq.Pins(MeasureF_Pin).Value(0) = MeasureFreq.Pins(MeasureF_Pin).Value(0) - (SimuIndex * 1000)
                           ' MeasureFreq.Pins(MeasureF_Pin).Value(1) = MeasureFreq.Pins(MeasureF_Pin).Value(1) + (SimuIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(2) = MeasureFreq.Pins(MeasureF_Pin).Value(2) + (TestLimitIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(3) = MeasureFreq.Pins(MeasureF_Pin).Value(3) - (TestLimitIndex * 1000)
                        End If
                        '--------------------------------------------
                        
                        If j = SrcEndBit + 1 Then
                            MeasureFreq_F1 = MeasureFreq
                        ElseIf j = SrcEndBit Then
                            MeasureFreq_F2 = MeasureFreq
                        End If
                    Else
                        '' Do nothing
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20161206-Restore force condiction after measurement
                    ''Call SetForceCondition("RESTORE")
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        End If
                    End If
            
                    TestSeqNum = TestSeqNum + 1
                    
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                    Else
                        TheHdw.Digital.Patgen.HaltWait
                    End If
                Next Ts
                
                TheHdw.Digital.Patgen.HaltWait
                
                StoreEachTrimFreq(StoreEachIndex) = MeasureFreq
                StoreEachIndex = StoreEachIndex + 1
                
                If j = SrcEndBit Then
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq_F1.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureFreq_F2.Math.Subtract(TrimTarget).Abs)
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                    PastDSPWave = InDSPwave
                Else
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                    PastDSPWave = InDSPwave
                End If
    
                TestLimitIndex = TestLimitIndex + 1
                
                '' 20160712 - Modify to use WriteComment to display output frequency.
                
                If gl_Disable_HIP_debug_log = False Then
                
                    For Each site In TheExec.sites
        ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        If b_KeepGoing(site) = True Then
                            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                        End If
        ''                End If
                    Next site
                End If
                
                ''20160716 - Modify display info sequence when source bit in the section end
                If j = SrcEndBit Then
                    For Each site In TheExec.sites
    ''                    If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        
                        If b_KeepGoing(site) = True And gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 1 & " Output Frequency = " & FormatNumber((MeasureFreq_F1.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 2 & " Output Frequency = " & FormatNumber((MeasureFreq_F2.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                        End If
    ''                    End If
                    Next site
                    LastSectionF1F2_Index = LastSectionF1F2_Index + 2
                End If
                
                '' 20160712 - Compare Measure Frequency whether match target Freq
                b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
                        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                            StoredTargetTrimCode(site).Data = InDSPwave(site).Data
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                Next site
                ''20161128-Stop trim code process if found out match code of all site
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
            Next j

            
        Next i
    End If
    
    
    
    
    ''============================================================================
    ''20161128 Findout mimiumn trim code
    Dim CloseTargetFreq As New PinListData
    Dim DiffValue As New SiteLong, PreviousDiffValue As New SiteLong, CloseIndex As New SiteLong
    
    Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteLong
    Dim BinStr As String
    Dim CloseTargetTrimCode As New DSPWave
    Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
    Dim b_FirstTimeSwitch As Boolean
    
    If b_KeepGoing.All(False) Then
    Else
''        If TrimPrcocessAll = True Then
            CloseTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If StoredTargetTrimCode(site).CalcSum = 0 Then
                        b_UseMinTrim(site) = True
                    End If
                End If
            Next site
            
            If UseMinimumTrimCode = True Then
                b_UseMinTrim = True
            End If
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_UseMinTrim(site) = True Then
                        '' Findout minimum difference value
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            DiffValue(site) = Abs(StoreEachTrimFreq(i).Pins(0).Value(site) - TrimTarget)
                            If DiffValue(site) <= PreviousDiffValue(site) Then
                                CloseIndex(site) = i
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                            If i = 0 Then
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                        Next i
                        '' Transfer to decimal value to findout minimum code
                        PreviousDecVal = 0
                        DecVal = 0
                        b_FirstTimeSwitch = False
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            BinStr = ""
                            If Abs(StoreEachTrimFreq(i).Pins(0).Value(site) - TrimTarget) = MinDiffVal(site) Then
                                For j = 0 To StoreEachTrimCode(i)(site).SampleSize - 1
                                    BinStr = BinStr & StoreEachTrimCode(i)(site).Element(j)
                                Next j
                                DecVal = Bin2Dec_rev_Double(BinStr)
                               
                                If DecVal < PreviousDecVal Then
                                    MinDecVal = DecVal
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                End If
                                PreviousDecVal = DecVal
                                If b_FirstTimeSwitch = False Then
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                    b_FirstTimeSwitch = True
                                End If
                            End If
                        Next i
                    End If
                End If
            Next site
''        End If
    End If
    
    For Each site In TheExec.sites
        If b_KeepGoing(site) = True Then
            If b_UseMinTrim(site) = True Then
                StoredTargetTrimCode(site).Data = CloseTargetTrimCode(site).Data
            Else
                StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
            End If
        Else
            StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
        End If
    Next site
    ''============================================================================
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, StoredTargetTrimCode)
    End If
    
    
    
    Call HardIP_WriteFuncResult(, , Inst_Name_Str)

    For Each site In TheExec.sites
        OutputTrimCode = ""
        For k = 0 To StoredTargetTrimCode(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(StoredTargetTrimCode(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Final Output Trim Code = " & OutputTrimCode)
    Next site
    
    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel StoredTargetTrimCode, DigSrc_Sample_Size, ConvertedDataWf
    
    TestNameInput = Report_TName_From_Instance("C", DigSrc_pin.Value, "", 0, 0)
    
    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, StoredTargetTrimCode)

    Call TheHdw.Patterns(PattArray(0)).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
    
        For Each Ts In TestSequenceArray
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
            Else
                Call TheHdw.Digital.Patgen.HaltWait
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                End If
            End If
            
            If UCase(Ts) = "F" Then
                Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20161206-Restore force condiction after measurement
            ''Call SetForceCondition("RESTORE")
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                End If
            End If
    
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA)
            Else
                TheHdw.Digital.Patgen.HaltWait
            End If
        Next Ts
    
    TheHdw.Digital.Patgen.HaltWait
    
    If TPModeAsCharz_GLB Then
        Disable_Inst_pinname_in_PTR
        TheExec.Flow.TestLimit resultVal:=MeasureFreq, Unit:=unitHz, ForceResults:=tlForceFlow
        Enable_Inst_pinname_in_PTR
    Else
        For p = 0 To MeasureFreq.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("F", MeasureFreq.Pins(p), "", CInt(p))
            TheExec.Flow.TestLimit resultVal:=MeasureFreq, Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next p
    End If
    
    Dim sl_FUSE_Val As New SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else

    End If
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    DebugPrintFunc patset.Value
    
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeFreq function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function

Public Function TrimCodeFreq_new_0828(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double = 1000000, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PreMeas As String, Optional Validating_ As Boolean, Optional Interpose_PrePat As String) As Long
    
    Dim PatCount As Long, PattArray() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen
    Dim Ts As Variant, TestSequenceArray() As String
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPwave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long, p As Long
    
    Dim MeasureFreq As New PinListData, MeasureFreq_F1 As New PinListData, MeasureFreq_F2 As New PinListData
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName

    On Error GoTo ErrorHandler
         
    Call GetFlowTName
         
    ''Update Interpose_PreMeas 20170801
    Dim Interpose_PreMeas_Ary() As String
    ''20160923 - Analyze Interpose_PreMeas to force setting with different sequence.
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    If (UBound(TestSequenceArray) > UBound(TName_Ary)) Then
        ReDim Preserve TName_Ary(UBound(TestSequenceArray)) As String
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    TheHdw.Patterns(patset).Load
    gl_TName_Pat = patset.Value

    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)

    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    Dim d_MeasF_Interval  As Double
    d_MeasF_Interval = 0.01
    
    Dim b_HighThanTargetFreq As New SiteBoolean
    b_HighThanTargetFreq = False
    
    Dim OutputTrimCode As String
    Dim TestLimitIndex As Long, LastSectionF1F2_Index As Long
    LastSectionF1F2_Index = 0


    ''==================================================================================================
'    Dim TrimStart_1st() As String
    Dim Dec_TrimStart_1st As Long
    
    '' 20160706 Create value for final frequency
    Dim b_DefineFinalFreq As New SiteBoolean
    Dim FinalFreq As New PinListData
    
    ''20160712 - If match taget freq just store the trim code
    Dim b_MatchTagetFreq As New SiteBoolean
    Dim b_DisplayFreq As New SiteBoolean
    Dim StoredTargetTrimCode As New DSPWave
    b_MatchTagetFreq = False
    b_DisplayFreq = False
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    Dim StoreEachTrimFreq() As New PinListData
    Dim StoreEachTrimCode() As New DSPWave
    ReDim StoreEachTrimFreq(DigSrc_Sample_Size + 1) As New PinListData
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i

    ''20170721-Updated the TrimStart when the first bit is zero and seperate with "&"
    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", "")
    End If
    
'    Dim z As Integer
'        For z = 0 To 1
'
'        ''    TrimStart_1st = TrimStart
'            If z = 0 Then
'                Dec_TrimStart_1st = Bin2Dec(TrimStart)
'            Else
'                Dec_TrimStart_1st = Dec_TrimStart_1st + 4080
'            End If
'            InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong
'
'            Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, DigSrc_Sample_Size, InDSPWave)
'
'            Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPWave)
'
'            If gl_Disable_HIP_debug_log = False Then theexec.Datalog.WriteComment ("First Time Setup")
'            '' Debug use
'            For Each site In theexec.sites
'                OutputTrimCode = ""
'                For k = 0 To InDSPWave(site).SampleSize - 1
'                    OutputTrimCode = OutputTrimCode & CStr(InDSPWave(site).Element(k))
'                Next k
'
'                If gl_Disable_HIP_debug_log = False Then theexec.Datalog.WriteComment ("Site_" & site & " Coarse Output Trim Code = " & OutputTrimCode)
'            Next site
'
'            For Each site In theexec.sites
'                StoreEachTrimCode(0)(site).Data = InDSPWave(site).Data
'            Next site
'
'            Call thehdw.Patterns(PattArray(0)).start
'
            ''Update Interpose_PreMeas 20170801
            Dim TestSeqNum As Integer
            TestSeqNum = 0
'
'            For Each Ts In TestSequenceArray
'                If (CPUA_Flag_In_Pat) Then
'                    Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)
'                Else
'                    Call thehdw.Digital.Patgen.HaltWait
'                End If
'
'                ''Update Interpose_PreMeas 20170801
'                ''20160923 - Add Interpose_PreMeas entry point by each sequence
'                If Interpose_PreMeas <> "" Then
'                    If UBound(Interpose_PreMeas_Ary) = 0 Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
'                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
'                    End If
'                End If
'
'                If UCase(Ts) = "F" Then
'                    Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
'                    Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
'
'                    If theexec.TesterMode = testModeOffline Then
'                        Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
'                    End If
'                End If
'
'                ''Update Interpose_PreMeas 20170801
'                ''20161206-Restore force condiction after measurement
'                ''Call SetForceCondition("RESTORE")
'                If Interpose_PreMeas <> "" Then
'                    If UBound(Interpose_PreMeas_Ary) = 0 Then
'                        Call SetForceCondition("RESTOREPREMEAS")
'                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                        Call SetForceCondition("RESTOREPREMEAS")
'                    End If
'                End If
'
'                TestSeqNum = TestSeqNum + 1
'
'                If (CPUA_Flag_In_Pat) Then
'                    Call thehdw.Digital.Patgen.Continue(0, cpuA)
'                Else
'                    thehdw.Digital.Patgen.HaltWait
'                End If
'            Next Ts
'            thehdw.Digital.Patgen.HaltWait
'
'            StoreEachTrimFreq(z) = MeasureFreq
'
'            b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
'            PastDSPWave = InDSPWave
'
'            TestNameInput = "Freq_meas_"
'            TestLimitIndex = 0
'
'            '' 20160712 - Modify to use WriteComment to display output frequency.
'            If gl_Disable_HIP_debug_log = False Then
'                For Each site In theexec.sites
'                        theexec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
'                Next site
'            End If
'            '' 20160712 - Compare Measure Frequency whether match target Freq
'            b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
'
'            b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
'            For Each site In theexec.sites
'                If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
'                    StoredTargetTrimCode(site).Data = InDSPWave(site).Data
'                    b_StopTrimCodeProcess(site) = True
'                End If
'            Next site
'            If gl_Disable_HIP_debug_log = False Then theexec.Datalog.WriteComment ("======================================================================================")
'
'   Next z
            ''========================================================================================
    '20161128 Pre check Min/Max trim code process
    Dim b_KeepGoing As New SiteBoolean
    Dim PreviousFreq As New PinListData
'    If PreCheckMinMaxTrimCode = True Then
'        PreviousFreq = MeasureFreq
'        Call rundsp.PreCheckMinMaxTrimCode(b_HighThanTargetFreq, InDSPWave)
'        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPWave)
'
'        ''Update Interpose_PreMeas 20170801
'        TestSeqNum = 0
'
'        For Each Ts In TestSequenceArray
'            If (CPUA_Flag_In_Pat) Then
'                Call thehdw.Digital.Patgen.FlagWait(cpuA, 0)
'            Else
'                Call thehdw.Digital.Patgen.HaltWait
'            End If
'
'            ''Update Interpose_PreMeas 20170801
'            ''20160923 - Add Interpose_PreMeas entry point by each sequence
'            If Interpose_PreMeas <> "" Then
'                If UBound(Interpose_PreMeas_Ary) = 0 Then
'                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
'                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
'                End If
'            End If
'
'            If UCase(Ts) = "F" Then
'                Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
'                Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
'
'                If theexec.TesterMode = testModeOffline Then
'                    Call SimulatePreCheckOutputFreq(MeasureF_Pin, MeasureFreq)
'                End If
'            End If
'
'            ''Update Interpose_PreMeas 20170801
'            ''20161206-Restore force condiction after measurement
'            ''Call SetForceCondition("RESTORE")
'            If Interpose_PreMeas <> "" Then
'                If UBound(Interpose_PreMeas_Ary) = 0 Then
'                    Call SetForceCondition("RESTOREPREMEAS")
'                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                    Call SetForceCondition("RESTOREPREMEAS")
'                End If
'            End If
'
'            TestSeqNum = TestSeqNum + 1
'
'            If (CPUA_Flag_In_Pat) Then
'                Call thehdw.Digital.Patgen.Continue(0, cpuA)
'            Else
'                thehdw.Digital.Patgen.HaltWait
'            End If
'        Next Ts
'
'        thehdw.Digital.Patgen.HaltWait
'
'        For Each site In theexec.sites
'            OutputTrimCode = ""
'            For k = 0 To InDSPWave(site).SampleSize - 1
'                OutputTrimCode = OutputTrimCode & CStr(InDSPWave(site).Element(k))
'            Next k
'            If gl_Disable_HIP_debug_log = False Then theexec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
'        Next site
'
'        If gl_Disable_HIP_debug_log = False Then
'            For Each site In theexec.sites
'                theexec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
'            Next site
'        End If
'        For Each site In theexec.sites
'            If b_HighThanTargetFreq(site) = True Then
'                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(LessThan, 0)
'            Else
'                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(GreaterThan, 0)
'            End If
'        Next site
'
'        Dim PreCheckBinStr As String, PreCheckDecVal As Double
'        For Each site In theexec.sites
'            If b_KeepGoing(site) = False Then
'                b_StopTrimCodeProcess(site) = True
'                PreCheckBinStr = ""
'                StoredTargetTrimCode(site).Data = InDSPWave(site).Data
'                For i = 0 To StoredTargetTrimCode(site).SampleSize - 1
'                    PreCheckBinStr = PreCheckBinStr & StoredTargetTrimCode.Element(i)
'                Next i
'                PreCheckDecVal = Bin2Dec_rev_Double(PreCheckBinStr)
'                ''TheExec.Flow.TestLimit PreCheckDecVal, 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TheExec.DataManager.InstanceName & "_TrimCode_Decimal", ForceResults:=tlForceNone
'            End If
'        Next site
'    End If

    ''========================================================================================
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    StoreEachIndex = 0
    
    Dim b_pos_trim As New SiteBoolean
    Dim b_neg_trim As New SiteBoolean
    Dim b_temp_0x00 As New SiteLong
    Dim b_temp_0xff As New SiteLong
    Dim fine_trim_flag As New SiteBoolean
    Dim coarse_trim_flag As New SiteBoolean
    Dim trim_store_temp As New DSPWave
    Dim InitialDSPWave_0x00 As New DSPWave
    Dim InitialDSPWave_0xff As New DSPWave
        InitialDSPWave_0x00.CreateConstant 0, 8, DspLong
        InitialDSPWave_0xff.CreateConstant 1, 8, DspLong
    fine_trim_flag = False
    
    ''20170103-Setup b_KeepGoing to true if PreCheckMinMaxTrimCode=false
    If PreCheckMinMaxTrimCode = False Then
        b_KeepGoing = True
    End If
    If b_KeepGoing.All(False) Then
    Else

        For i = 0 To ExecutionMax
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If
            StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
            StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
            SrcStartBit = StrSeparatebyColon(0)
            SrcEndBit = StrSeparatebyColon(1)
            If i = 0 Then
                b_FirstExecution = True
            Else
                b_FirstExecution = False
                SrcStartBit = SrcStartBit + 1
            End If
            
                For j = 3 To 0 Step -1
                  If fine_trim_flag.All(True) Then Exit For
                    Dim z As Integer
                    For z = 0 To 1
                        If z = 0 Then
                            For Each site In TheExec.sites
                                InDSPwave.CreateConstant 0, 0, DspLong
                                If b_FirstExecution = True Then
                                    b_temp_0x00 = 8
                                ElseIf b_temp_0x00 = 15 Then
                                     fine_trim_flag = True
                                ElseIf b_FirstExecution = False And b_pos_trim(site) = True Then
                                    b_temp_0x00 = b_temp_0x00 + (2) ^ (j)
                                ElseIf b_FirstExecution = False And b_neg_trim(site) = True Then
                                    b_temp_0x00 = b_temp_0x00 - (2) ^ (j)
                                End If
                                InitialDSPWave.CreateConstant b_temp_0x00, 1, DspLong
                                Dim temp_bin() As String
                                 If b_StopTrimCodeProcess(site) = True Then
                                    'do nothing
                                 Else
                                  TheExec.Datalog.WriteComment ("Setting site=" & site & "_coarse trim=" & b_temp_0x00)
                                End If
                            Next site
                            If fine_trim_flag.All(True) Then Exit For
                                Call rundsp.CreateFlexibleDSPWave_lpro(InitialDSPWave, 4, InDSPwave, InitialDSPWave_0x00)
                                For Each site In TheExec.sites
                                    trim_store_temp.Data = InDSPwave.Data
                                Next site
                         Else
                            For Each site In TheExec.sites
                                InDSPwave.CreateConstant 0, 0, DspLong
                                If b_FirstExecution = True Then
                                    b_temp_0xff = b_temp_0x00
                                ElseIf b_temp_0x00 = 15 Then
                                     fine_trim_flag = True
                                ElseIf b_FirstExecution = False And b_pos_trim = True Then
                                    b_temp_0xff = b_temp_0xff + (2) ^ (j)
                                ElseIf b_FirstExecution = False And b_neg_trim = True Then
                                    b_temp_0xff = b_temp_0xff - (2) ^ (j)
                                End If
                                InitialDSPWave.CreateConstant b_temp_0xff, 1, DspLong
                                If b_StopTrimCodeProcess(site) = True Then
                                    'do  nothing
                                 Else
                                   TheExec.Datalog.WriteComment ("Setting site=" & site & "_coarse trim=" & b_temp_0xff)
                                End If
                           Next site
                                Call rundsp.CreateFlexibleDSPWave_lpro(InitialDSPWave, 4, InDSPwave, InitialDSPWave_0xff)
                        End If
                        
                            Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
            
                                    
                        For Each site In TheExec.sites
                            StoreEachTrimCode(StoreEachIndex)(site).Data = InDSPwave(site).Data
                        Next site
                        
                            For Each site In TheExec.sites
                ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                                If b_StopTrimCodeProcess(site) = False Then
                                    OutputTrimCode = ""

                                    For k = 0 To InDSPwave(site).SampleSize - 1
                                        OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
                                    Next k
                                    
                                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
                                End If
                ''                End If
                            Next site
                            '' ==============================================================================================
    
                            Call TheHdw.Patterns(PattArray(0)).start
                            
                            ''Update Interpose_PreMeas 20170801
                            TestSeqNum = 0
                            
                            For Each Ts In TestSequenceArray
                                If (CPUA_Flag_In_Pat) Then
                                    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
                                Else
                                    Call TheHdw.Digital.Patgen.HaltWait
                                End If
                                
                                ''Update Interpose_PreMeas 20170801
                                ''20160923 - Add Interpose_PreMeas entry point by each sequence
                                If Interpose_PreMeas <> "" Then
                                    If UBound(Interpose_PreMeas_Ary) = 0 Then
                                        Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                                        Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                                    End If
                                End If
                                
                                
                                If UCase(Ts) = "F" Then
                                    Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                                    Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
                                    
                                    If z = 0 Then
                                        MeasureFreq_F1 = MeasureFreq
                                    ElseIf z = 1 Then
                                        MeasureFreq_F2 = MeasureFreq
                                    End If
            
                                End If
                                
                                ''Update Interpose_PreMeas 20170801
                                ''20161206-Restore force condiction after measurement
                                ''Call SetForceCondition("RESTORE")
                                If Interpose_PreMeas <> "" Then
                                    If UBound(Interpose_PreMeas_Ary) = 0 Then
                                        Call SetForceCondition("RESTOREPREMEAS")
                                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                                        Call SetForceCondition("RESTOREPREMEAS")
                                    End If
                                End If
                                TestSeqNum = TestSeqNum + 1
                                
                                If (CPUA_Flag_In_Pat) Then
                                    Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                                Else
                                    TheHdw.Digital.Patgen.HaltWait
                                End If
                            Next Ts
                            
                            TheHdw.Digital.Patgen.HaltWait
                            StoreEachTrimFreq(StoreEachIndex) = MeasureFreq
                            StoreEachIndex = StoreEachIndex + 1
                                    
                            If gl_Disable_HIP_debug_log = False Then
                                For Each site In TheExec.sites
                    ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                                     If b_StopTrimCodeProcess(site) = False Then
                                        TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                                    End If
                    ''                End If
                                Next site
                            End If
    
                            If TrimPrcocessAll = False Then
                                If b_StopTrimCodeProcess.All(True) Then
                                    Exit For
                                End If
                            End If
                            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
                        Next z
                b_FirstExecution = False
                For Each site In TheExec.sites
                    If MeasureFreq_F1.Pins(0).Subtract(1000000).Value(site) <= 0 And MeasureFreq_F2.Pins(0).Subtract(1000000).Value(site) >= 0 Then
                         fine_trim_flag(site) = True
'                         '---provide 0 or 1
'                                b_HighThanTargetFreq = False
'                                b_HighThanTargetFreq = MeasureFreq_F2.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureFreq_F1.Math.Subtract(TrimTarget).Abs)
                            If trim_store_temp.Element(8) = 1 Then
                                b_HighThanTargetFreq = False
                            Else
                                b_HighThanTargetFreq = True
                            End If
                                PastDSPWave.Data = trim_store_temp.Data
                    ElseIf MeasureFreq_F1.Pins(0).Subtract(1000000).Value(site) >= 0 And MeasureFreq_F2.Pins(0).Subtract(1000000).Value(site) >= 0 Then
                        b_neg_trim(site) = True
                        b_pos_trim(site) = False
                        PastDSPWave.Data = trim_store_temp.Data
                    ElseIf MeasureFreq_F1.Pins(0).Subtract(1000000).Value(site) <= 0 And MeasureFreq_F2.Pins(0).Subtract(1000000).Value(site) <= 0 Then
                        b_pos_trim(site) = True
                        b_neg_trim(site) = False
                        PastDSPWave.Data = trim_store_temp.Data
                    End If
                Next site
                Next j
            '------------------------------fine trim
            
            StoreEachIndex = 0
            For j = SrcStartBit To SrcEndBit Step -1
                If i = 0 Then Exit For
               
                If b_FirstExecution = True Then
                    b_ControlNextBit = True
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                Else
                ''20160716-Control next bit to 1 no matter first or last progress
                    b_ControlNextBit = True
    ''                b_ControlNextBit = False
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                End If

                If b_FirstExecution = True And j = SrcEndBit Then
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPwave)
                Else
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                End If
                
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
                
                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site).Data = InDSPwave(site).Data
                Next site
            
                '' Debug use
                '' ==============================================================================================
                '' 20160716 - Modify trim code rule
                
                If gl_Disable_HIP_debug_log = False Then
                    If b_FirstExecution = True Then
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    Else
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j)
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    End If
                End If
                
                Dim fine_wave As New DSPWave
                Dim OutputTrimCode_fine As String
                Dim OutputTrimCode_fine_dec As New DSPWave
                    
                    
                For Each site In TheExec.sites
                 fine_wave = InDSPwave.Select(0, , 8).Copy
    ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                    OutputTrimCode_fine_dec = fine_wave.ConvertStreamTo(tldspParallel, 8, 0, Bit0IsMsb)
                    If b_KeepGoing(site) = True Then
                        OutputTrimCode = ""
                        OutputTrimCode_fine = ""
                        For k = 0 To InDSPwave(site).SampleSize - 1
                            OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
                            If k < 8 Then
                            OutputTrimCode_fine = OutputTrimCode_fine & CStr(fine_wave(site).Element(k))
                            End If
                        Next k
                        TheExec.Datalog.WriteComment ("Site_" & site & "  Fine Trim Code = " & OutputTrimCode_fine & ", fine decimal= " & OutputTrimCode_fine_dec(site).Element(0))
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output All Code = " & OutputTrimCode)
                    End If
    ''                End If
                Next site
                '' ==============================================================================================
                
                Call TheHdw.Patterns(PattArray(0)).start
                
                ''Update Interpose_PreMeas 20170801
                TestSeqNum = 0
                
                For Each Ts In TestSequenceArray
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
                    Else
                        Call TheHdw.Digital.Patgen.HaltWait
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20160923 - Add Interpose_PreMeas entry point by each sequence
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                        End If
                    End If
                    
                    
                    If UCase(Ts) = "F" Then
                        Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                        TheHdw.Digital.Pins(MeasureF_Pin).Levels.DriverMode = tlDriverModeVt
                        Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
                        
                        '--------------- off line mode data --------
                        If TheExec.TesterMode = testModeOffline Then
                            Dim SimuIndex As Long
                            SimuIndex = TestLimitIndex
                            If SimuIndex >= 8 Then
                                SimuIndex = 8
                            End If
                            Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
                            MeasureFreq.Pins(MeasureF_Pin).Value(0) = MeasureFreq.Pins(MeasureF_Pin).Value(0) - (SimuIndex * 1000)
                           ' MeasureFreq.Pins(MeasureF_Pin).Value(1) = MeasureFreq.Pins(MeasureF_Pin).Value(1) + (SimuIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(2) = MeasureFreq.Pins(MeasureF_Pin).Value(2) + (TestLimitIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(3) = MeasureFreq.Pins(MeasureF_Pin).Value(3) - (TestLimitIndex * 1000)
                        End If
                        '--------------------------------------------
                        
                        If j = SrcEndBit + 1 Then
                            MeasureFreq_F1 = MeasureFreq
                        ElseIf j = SrcEndBit Then
                            MeasureFreq_F2 = MeasureFreq
                        End If
                    Else
                        '' Do nothing
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20161206-Restore force condiction after measurement
                    ''Call SetForceCondition("RESTORE")
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        End If
                    End If
            
                    TestSeqNum = TestSeqNum + 1
                    
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                    Else
                        TheHdw.Digital.Patgen.HaltWait
                    End If
                Next Ts
                
                TheHdw.Digital.Patgen.HaltWait
                
                StoreEachTrimFreq(StoreEachIndex) = MeasureFreq
                StoreEachIndex = StoreEachIndex + 1
                
                If j = SrcEndBit Then
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq_F1.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureFreq_F2.Math.Subtract(TrimTarget).Abs)
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                    PastDSPWave = InDSPwave
                Else
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                    PastDSPWave = InDSPwave
                End If
    
                TestLimitIndex = TestLimitIndex + 1
                
                '' 20160712 - Modify to use WriteComment to display output frequency.
                
                If gl_Disable_HIP_debug_log = False Then
                
                    For Each site In TheExec.sites
        ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        If b_KeepGoing(site) = True Then
                            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                        End If
        ''                End If
                    Next site
                End If
                
                ''20160716 - Modify display info sequence when source bit in the section end
                If j = SrcEndBit Then
                    For Each site In TheExec.sites
    ''                    If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        
                        If b_KeepGoing(site) = True And gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 1 & " Output Frequency = " & FormatNumber((MeasureFreq_F1.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 2 & " Output Frequency = " & FormatNumber((MeasureFreq_F2.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                        End If
    ''                    End If
                    Next site
                    LastSectionF1F2_Index = LastSectionF1F2_Index + 2
                End If
                
                '' 20160712 - Compare Measure Frequency whether match target Freq
                b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
                        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                            StoredTargetTrimCode(site).Data = InDSPwave(site).Data
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                Next site
                ''20161128-Stop trim code process if found out match code of all site
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
            Next j

            
        Next i
    End If
    
    
    
    
    ''============================================================================
    ''20161128 Findout mimiumn trim code
    Dim CloseTargetFreq As New PinListData
    Dim DiffValue As New SiteLong, PreviousDiffValue As New SiteLong, CloseIndex As New SiteLong
    
    Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteLong
    Dim BinStr As String
    Dim CloseTargetTrimCode As New DSPWave
    Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
    Dim b_FirstTimeSwitch As Boolean
    
    If b_KeepGoing.All(False) Then
    Else
''        If TrimPrcocessAll = True Then
            CloseTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If StoredTargetTrimCode(site).CalcSum = 0 Then
                        b_UseMinTrim(site) = True
                    End If
                End If
            Next site
            
            If UseMinimumTrimCode = True Then
                b_UseMinTrim = True
            End If
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_UseMinTrim(site) = True Then
                        '' Findout minimum difference value
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            DiffValue(site) = Abs(StoreEachTrimFreq(i).Pins(0).Value(site) - TrimTarget)
                            If DiffValue(site) <= PreviousDiffValue(site) Then
                                CloseIndex(site) = i
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                            If i = 0 Then
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                        Next i
                        '' Transfer to decimal value to findout minimum code
                        PreviousDecVal = 0
                        DecVal = 0
                        b_FirstTimeSwitch = False
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            BinStr = ""
                            If Abs(StoreEachTrimFreq(i).Pins(0).Value(site) - TrimTarget) = MinDiffVal(site) Then
                                For j = 0 To StoreEachTrimCode(i)(site).SampleSize - 1
                                    BinStr = BinStr & StoreEachTrimCode(i)(site).Element(j)
                                Next j
                                DecVal = Bin2Dec_rev_Double(BinStr)
                               
                                If DecVal < PreviousDecVal Then
                                    MinDecVal = DecVal
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                End If
                                PreviousDecVal = DecVal
                                If b_FirstTimeSwitch = False Then
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                    b_FirstTimeSwitch = True
                                End If
                            End If
                        Next i
                    End If
                End If
            Next site
''        End If
    End If
    
    For Each site In TheExec.sites
        If b_KeepGoing(site) = True Then
            If b_UseMinTrim(site) = True Then
                StoredTargetTrimCode(site).Data = CloseTargetTrimCode(site).Data
            Else
                StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
            End If
        Else
            StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
        End If
    Next site
    ''============================================================================
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, StoredTargetTrimCode)
    End If
    
    
    
    Call HardIP_WriteFuncResult(, , Inst_Name_Str)

    For Each site In TheExec.sites
        OutputTrimCode = ""
        For k = 0 To StoredTargetTrimCode(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(StoredTargetTrimCode(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Final Output Trim Code = " & OutputTrimCode)
    Next site
    
    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel StoredTargetTrimCode, DigSrc_Sample_Size, ConvertedDataWf
    
    TestNameInput = Report_TName_From_Instance("C", DigSrc_pin.Value, "", 0, 0)
    
    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, StoredTargetTrimCode)

    Call TheHdw.Patterns(PattArray(0)).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
    
        For Each Ts In TestSequenceArray
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
            Else
                Call TheHdw.Digital.Patgen.HaltWait
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                End If
            End If
            
            If UCase(Ts) = "F" Then
                Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20161206-Restore force condiction after measurement
            ''Call SetForceCondition("RESTORE")
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                End If
            End If
    
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA)
            Else
                TheHdw.Digital.Patgen.HaltWait
            End If
        Next Ts
    
    TheHdw.Digital.Patgen.HaltWait
    
    If TPModeAsCharz_GLB Then
        Disable_Inst_pinname_in_PTR
        TheExec.Flow.TestLimit resultVal:=MeasureFreq, Unit:=unitHz, ForceResults:=tlForceFlow
        Enable_Inst_pinname_in_PTR
    Else
        For p = 0 To MeasureFreq.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("F", MeasureFreq.Pins(p), "", CInt(p))
            TheExec.Flow.TestLimit resultVal:=MeasureFreq, Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next p
    End If
    
    Dim sl_FUSE_Val As New SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else

    End If
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    DebugPrintFunc patset.Value
    
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeFreq function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function


Public Function TrimCodeFreq(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double = 1000000, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PreMeas As String, Optional Validating_ As Boolean, Optional Interpose_PrePat As String) As Long
    
    Dim PatCount As Long, PattArray() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen
    Dim Ts As Variant, TestSequenceArray() As String
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPwave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long, p As Long
    
    Dim MeasureFreq As New PinListData, MeasureFreq_F1 As New PinListData, MeasureFreq_F2 As New PinListData
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName

    On Error GoTo ErrorHandler
         
    Call GetFlowTName
         
    ''Update Interpose_PreMeas 20170801
    Dim Interpose_PreMeas_Ary() As String
    ''20160923 - Analyze Interpose_PreMeas to force setting with different sequence.
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    If (UBound(TestSequenceArray) > UBound(TName_Ary)) Then
        ReDim Preserve TName_Ary(UBound(TestSequenceArray)) As String
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    TheHdw.Patterns(patset).Load
    gl_TName_Pat = patset.Value

    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)

    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    Dim d_MeasF_Interval  As Double
    d_MeasF_Interval = 0.001
    
    Dim b_HighThanTargetFreq As New SiteBoolean
    b_HighThanTargetFreq = False
    
    Dim OutputTrimCode As String
    Dim TestLimitIndex As Long, LastSectionF1F2_Index As Long
    LastSectionF1F2_Index = 0
    
    ''==================================================================================================
'    Dim TrimStart_1st() As String
    Dim Dec_TrimStart_1st As Long
    
    '' 20160706 Create value for final frequency
    Dim b_DefineFinalFreq As New SiteBoolean
    Dim FinalFreq As New PinListData
    
    ''20160712 - If match taget freq just store the trim code
    Dim b_MatchTagetFreq As New SiteBoolean
    Dim b_DisplayFreq As New SiteBoolean
    Dim StoredTargetTrimCode As New DSPWave
    b_MatchTagetFreq = False
    b_DisplayFreq = False
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    Dim StoreEachTrimFreq() As New PinListData
    Dim StoreEachTrimCode() As New DSPWave
    ReDim StoreEachTrimFreq(DigSrc_Sample_Size + 1) As New PinListData
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i

    ''20170721-Updated the TrimStart when the first bit is zero and seperate with "&"
    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", "")
    End If

''    TrimStart_1st = TrimStart
    Dec_TrimStart_1st = Bin2Dec(TrimStart)
    
    InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong

    Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, DigSrc_Sample_Size, InDSPwave)
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("First Time Setup")
    '' Debug use
    For Each site In TheExec.sites
        OutputTrimCode = ""
        For k = 0 To InDSPwave(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
    Next site
    
    For Each site In TheExec.sites
        StoreEachTrimCode(0)(site).Data = InDSPwave(site).Data
    Next site
    
    Call TheHdw.Patterns(PattArray(0)).start
    
    ''Update Interpose_PreMeas 20170801
    Dim TestSeqNum As Integer
    TestSeqNum = 0

    For Each Ts In TestSequenceArray
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20160923 - Add Interpose_PreMeas entry point by each sequence
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            End If
        End If
        
        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
            TheHdw.Digital.Pins(MeasureF_Pin).Levels.DriverMode = tlDriverModeVt
            Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
            
            If TheExec.TesterMode = testModeOffline Then
                Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
            End If
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20161206-Restore force condiction after measurement
        ''Call SetForceCondition("RESTORE")
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            End If
        End If

        TestSeqNum = TestSeqNum + 1

        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts
    TheHdw.Digital.Patgen.HaltWait
    
    StoreEachTrimFreq(0) = MeasureFreq
    
    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
    PastDSPWave = InDSPwave
    
    TestNameInput = "Freq_meas_"
    TestLimitIndex = 0
    
    '' 20160712 - Modify to use WriteComment to display output frequency.
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
        Next site
    End If
    '' 20160712 - Compare Measure Frequency whether match target Freq
    b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
    
    b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
    For Each site In TheExec.sites
        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
            StoredTargetTrimCode(site).Data = InDSPwave(site).Data
            b_StopTrimCodeProcess(site) = True
        End If
    Next site
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
    
    
    ''========================================================================================
    ''20161128 Pre check Min/Max trim code process
    Dim b_KeepGoing As New SiteBoolean
    Dim PreviousFreq As New PinListData
    If PreCheckMinMaxTrimCode = True Then
        PreviousFreq = MeasureFreq
        Call rundsp.PreCheckMinMaxTrimCode(b_HighThanTargetFreq, InDSPwave)
        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
        
        ''Update Interpose_PreMeas 20170801
        TestSeqNum = 0
        
        For Each Ts In TestSequenceArray
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
            Else
                Call TheHdw.Digital.Patgen.HaltWait
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                End If
            End If
            
            If UCase(Ts) = "F" Then
                Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
                
                If TheExec.TesterMode = testModeOffline Then
                    Call SimulatePreCheckOutputFreq(MeasureF_Pin, MeasureFreq)
                End If
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20161206-Restore force condiction after measurement
            ''Call SetForceCondition("RESTORE")
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                End If
            End If
    
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA)
            Else
                TheHdw.Digital.Patgen.HaltWait
            End If
        Next Ts
        
        TheHdw.Digital.Patgen.HaltWait
        
        For Each site In TheExec.sites
            OutputTrimCode = ""
            For k = 0 To InDSPwave(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
            Next k
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
        Next site
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
            Next site
        End If
        For Each site In TheExec.sites
            If b_HighThanTargetFreq(site) = True Then
                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(LessThan, 0)
            Else
                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(GreaterThan, 0)
            End If
        Next site

        Dim PreCheckBinStr As String, PreCheckDecVal As Double
        For Each site In TheExec.sites
            If b_KeepGoing(site) = False Then
                b_StopTrimCodeProcess(site) = True
                PreCheckBinStr = ""
                StoredTargetTrimCode(site).Data = InDSPwave(site).Data
                For i = 0 To StoredTargetTrimCode(site).SampleSize - 1
                    PreCheckBinStr = PreCheckBinStr & StoredTargetTrimCode.Element(i)
                Next i
                PreCheckDecVal = Bin2Dec_rev_Double(PreCheckBinStr)
                ''TheExec.Flow.TestLimit PreCheckDecVal, 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TheExec.DataManager.InstanceName & "_TrimCode_Decimal", ForceResults:=tlForceNone
            End If
        Next site
    End If
    
    ''========================================================================================
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    StoreEachIndex = 1
    
    ''20170103-Setup b_KeepGoing to true if PreCheckMinMaxTrimCode=false
    If PreCheckMinMaxTrimCode = False Then
        b_KeepGoing = True
    End If
    
    If b_KeepGoing.All(False) Then
    Else

        For i = 0 To ExecutionMax
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If
            StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
            StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
            SrcStartBit = StrSeparatebyColon(0)
            SrcEndBit = StrSeparatebyColon(1)
            
            If i = 0 Then
                b_FirstExecution = True
            Else
                b_FirstExecution = False
                SrcStartBit = SrcStartBit + 1
            End If
            
            For j = SrcStartBit To SrcEndBit Step -1
            
                If b_FirstExecution = True Then
                    b_ControlNextBit = True
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                Else
                ''20160716-Control next bit to 1 no matter first or last progress
                    b_ControlNextBit = True
    ''                b_ControlNextBit = False
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                End If
    
                If b_FirstExecution = True And j = SrcEndBit Then
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPwave)
    ''            ElseIf b_FirstExecution = False And j = SrcStartBit Then
    ''                j = SrcStartBit + 1
                Else
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                End If
                
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
                
                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site).Data = InDSPwave(site).Data
                Next site
            
                '' Debug use
                '' ==============================================================================================
                '' 20160716 - Modify trim code rule
                
                If gl_Disable_HIP_debug_log = False Then
                    If b_FirstExecution = True Then
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    Else
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j)
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    End If
                End If
                
                For Each site In TheExec.sites
    ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                    If b_KeepGoing(site) = True Then
                        OutputTrimCode = ""
                        For k = 0 To InDSPwave(site).SampleSize - 1
                            OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
                        Next k
                        
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
                    End If
    ''                End If
                Next site
                '' ==============================================================================================
                
                Call TheHdw.Patterns(PattArray(0)).start
                
                ''Update Interpose_PreMeas 20170801
                TestSeqNum = 0
                
                For Each Ts In TestSequenceArray
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
                    Else
                        Call TheHdw.Digital.Patgen.HaltWait
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20160923 - Add Interpose_PreMeas entry point by each sequence
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                        End If
                    End If
                    
                    
                    If UCase(Ts) = "F" Then
                        Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                        Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
                        
                        '--------------- off line mode data --------
                        If TheExec.TesterMode = testModeOffline Then
                            Dim SimuIndex As Long
                            SimuIndex = TestLimitIndex
                            If SimuIndex >= 8 Then
                                SimuIndex = 8
                            End If
                            Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
                            MeasureFreq.Pins(MeasureF_Pin).Value(0) = MeasureFreq.Pins(MeasureF_Pin).Value(0) - (SimuIndex * 1000)
                           ' MeasureFreq.Pins(MeasureF_Pin).Value(1) = MeasureFreq.Pins(MeasureF_Pin).Value(1) + (SimuIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(2) = MeasureFreq.Pins(MeasureF_Pin).Value(2) + (TestLimitIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(3) = MeasureFreq.Pins(MeasureF_Pin).Value(3) - (TestLimitIndex * 1000)
                        End If
                        '--------------------------------------------
                        
                        If j = SrcEndBit + 1 Then
                            MeasureFreq_F1 = MeasureFreq
                        ElseIf j = SrcEndBit Then
                            MeasureFreq_F2 = MeasureFreq
                        End If
                    Else
                        '' Do nothing
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20161206-Restore force condiction after measurement
                    ''Call SetForceCondition("RESTORE")
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        End If
                    End If
            
                    TestSeqNum = TestSeqNum + 1
                    
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                    Else
                        TheHdw.Digital.Patgen.HaltWait
                    End If
                Next Ts
                
                TheHdw.Digital.Patgen.HaltWait
                
                StoreEachTrimFreq(StoreEachIndex) = MeasureFreq
                StoreEachIndex = StoreEachIndex + 1
                
                If j = SrcEndBit Then
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq_F1.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureFreq_F2.Math.Subtract(TrimTarget).Abs)
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                    PastDSPWave = InDSPwave
                Else
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                    PastDSPWave = InDSPwave
                End If
    
                TestLimitIndex = TestLimitIndex + 1
                
                '' 20160712 - Modify to use WriteComment to display output frequency.
                
                If gl_Disable_HIP_debug_log = False Then
                
                    For Each site In TheExec.sites
        ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        If b_KeepGoing(site) = True Then
                            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                        End If
        ''                End If
                    Next site
                End If
                
                ''20160716 - Modify display info sequence when source bit in the section end
                If j = SrcEndBit Then
                    For Each site In TheExec.sites
    ''                    If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        
                        If b_KeepGoing(site) = True And gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 1 & " Output Frequency = " & FormatNumber((MeasureFreq_F1.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 2 & " Output Frequency = " & FormatNumber((MeasureFreq_F2.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                        End If
    ''                    End If
                    Next site
                    LastSectionF1F2_Index = LastSectionF1F2_Index + 2
                End If
                
                '' 20160712 - Compare Measure Frequency whether match target Freq
                b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
                        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                            StoredTargetTrimCode(site).Data = InDSPwave(site).Data
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                Next site
                ''20161128-Stop trim code process if found out match code of all site
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
            Next j
        Next i
    End If
    
    ''============================================================================
    ''20161128 Findout mimiumn trim code
    Dim CloseTargetFreq As New PinListData
    Dim DiffValue As New SiteLong, PreviousDiffValue As New SiteLong, CloseIndex As New SiteLong
    
    Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteLong
    Dim BinStr As String
    Dim CloseTargetTrimCode As New DSPWave
    Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
    Dim b_FirstTimeSwitch As Boolean
    
    If b_KeepGoing.All(False) Then
    Else
''        If TrimPrcocessAll = True Then
            CloseTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If StoredTargetTrimCode(site).CalcSum = 0 Then
                        b_UseMinTrim(site) = True
                    End If
                End If
            Next site
            
            If UseMinimumTrimCode = True Then
                b_UseMinTrim = True
            End If
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_UseMinTrim(site) = True Then
                        '' Findout minimum difference value
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            DiffValue(site) = Abs(StoreEachTrimFreq(i).Pins(0).Value(site) - TrimTarget)
                            If DiffValue(site) <= PreviousDiffValue(site) Then
                                CloseIndex(site) = i
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                            If i = 0 Then
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                        Next i
                        '' Transfer to decimal value to findout minimum code
                        PreviousDecVal = 0
                        DecVal = 0
                        b_FirstTimeSwitch = False
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            BinStr = ""
                            If Abs(StoreEachTrimFreq(i).Pins(0).Value(site) - TrimTarget) = MinDiffVal(site) Then
                                For j = 0 To StoreEachTrimCode(i)(site).SampleSize - 1
                                    BinStr = BinStr & StoreEachTrimCode(i)(site).Element(j)
                                Next j
                                DecVal = Bin2Dec_rev_Double(BinStr)
                               
                                If DecVal < PreviousDecVal Then
                                    MinDecVal = DecVal
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                End If
                                PreviousDecVal = DecVal
                                If b_FirstTimeSwitch = False Then
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                    b_FirstTimeSwitch = True
                                End If
                            End If
                        Next i
                    End If
                End If
            Next site
''        End If
    End If
    
    For Each site In TheExec.sites
        If b_KeepGoing(site) = True Then
            If b_UseMinTrim(site) = True Then
                StoredTargetTrimCode(site).Data = CloseTargetTrimCode(site).Data
            Else
                StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
            End If
        Else
            StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
        End If
    Next site
    ''============================================================================
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, StoredTargetTrimCode)
    End If
    
    
    
    Call HardIP_WriteFuncResult(, , Inst_Name_Str)

    For Each site In TheExec.sites
        OutputTrimCode = ""
        For k = 0 To StoredTargetTrimCode(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(StoredTargetTrimCode(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Final Output Trim Code = " & OutputTrimCode)
    Next site
    
    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel StoredTargetTrimCode, DigSrc_Sample_Size, ConvertedDataWf
    
    TestNameInput = Report_TName_From_Instance("C", DigSrc_pin.Value, "TrimCode(Decimal)", 0, 0)
    
    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, StoredTargetTrimCode)

    Call TheHdw.Patterns(PattArray(0)).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
    
    For Each Ts In TestSequenceArray
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20160923 - Add Interpose_PreMeas entry point by each sequence
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            End If
        End If
        
        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
            Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20161206-Restore force condiction after measurement
        ''Call SetForceCondition("RESTORE")
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            End If
        End If

        TestSeqNum = TestSeqNum + 1
        
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts
    
    TheHdw.Digital.Patgen.HaltWait
    
    If TPModeAsCharz_GLB Then
        Disable_Inst_pinname_in_PTR
        TheExec.Flow.TestLimit resultVal:=MeasureFreq, Unit:=unitHz, ForceResults:=tlForceFlow
        Enable_Inst_pinname_in_PTR
    Else
        For p = 0 To MeasureFreq.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("F", MeasureFreq.Pins(p), "Final", CInt(p))
            TheExec.Flow.TestLimit resultVal:=MeasureFreq, Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next p
    End If
    
    Dim sl_FUSE_Val As New SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else

    End If
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    DebugPrintFunc patset.Value
    
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
    
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeFreq function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function


Public Function TrimCodeImpedence(Optional patset As Pattern, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasR_Pins_SingleEnd As PinList, Optional MeasR_Pins_Differential As PinList, Optional StrForceVolt As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double = 50, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, _
    Optional Fixed_DigSrc_DataWidth As Long, Optional Fixed_DigSrc_Sample_Size As Long, Optional Fixed_DigSrc_Equation As String, Optional Fixed_DigSrc_Assignment As String, _
    Optional Calc_Eqn As String, Optional GetbitNumber As Long, Optional b_PD_Mode As Boolean = True, Optional Validating_ As Boolean, Optional Interpose_PrePat As String) As Long

    Dim PatCount As Long, PattArray() As String
    Dim TrimmedImpedance() As New SiteDouble
    Dim TrimCode() As New SiteLong
    Dim X As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPwave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long
    
    Dim MeasureImped As New PinListData, MeasureImped_F1 As New PinListData, MeasureImped_F2 As New PinListData

    On Error GoTo ErrorHandler

    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Calc_Eqn)
    End If
    TName_Ary = Split(gl_Tname_Meas, "+")
    
    Call HardIP_InitialSetupForPatgen
    
    TheHdw.Patterns(patset).Load

    gl_TName_Pat = patset.Value

    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    '' 20161230 - To check "Fixed" key word to decide trim code process
    Dim b_ExecuteTrimCode() As Boolean
    ReDim b_ExecuteTrimCode(ExecutionMax) As Boolean
    Dim b_HasFIXED_InTrimFormat As Boolean
    
    For i = 0 To ExecutionMax
        If InStr(UCase(StrSeparatebyComma(i)), "FIXED") <> 0 Then
            b_ExecuteTrimCode(i) = False
            b_HasFIXED_InTrimFormat = True
        Else
            b_ExecuteTrimCode(i) = True

        End If
    Next i
    
    Dim Pin_Ary() As String, Pin_Cnt As Long, Pin As Variant
    Dim StorePerPinFinalTrimCode() As New DSPWave
    Dim StorePerPinFinalTrimCode_Dec() As New DSPWave
    Dim b_IsDifferential As Boolean
    Dim TempPins As String
    If MeasR_Pins_SingleEnd <> "" Then
        TheExec.DataManager.DecomposePinList MeasR_Pins_SingleEnd, Pin_Ary, Pin_Cnt
        b_IsDifferential = False
    ElseIf MeasR_Pins_Differential <> "" Then
        TheExec.DataManager.DecomposePinList MeasR_Pins_Differential, Pin_Ary, Pin_Cnt
        
        For i = 0 To Pin_Cnt - 1
            If InStr(UCase(Pin_Ary(i)), "_P") <> 0 Then
                If i = 0 Then
                    TempPins = Pin_Ary(i)
                Else
                    TempPins = TempPins & "," & Pin_Ary(i)
                End If
            End If
        Next i
        Pin_Cnt = Pin_Cnt / 2
        ReDim Pin_Ary(Pin_Cnt) As String
        Pin_Ary = Split(TempPins, ",")
        b_IsDifferential = True
    End If
    
    ReDim StorePerPinFinalTrimCode(Pin_Cnt - 1) As New DSPWave
    ReDim StorePerPinFinalTrimCode_Dec(Pin_Cnt - 1) As New DSPWave
        
    For i = 0 To Pin_Cnt - 1
        StorePerPinFinalTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
        StorePerPinFinalTrimCode_Dec(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i
    
    Dim b_HighThanTargetImped As New SiteBoolean
    b_HighThanTargetImped = False
    
    Dim OutputTrimCode As String
    Dim TestLimitIndex As Long, LastSectionF1F2_Index As Long
    LastSectionF1F2_Index = 0
    
    Dim Dec_TrimStart_1st As Long
    
    '' 20160706 Create value for final Impeduency
    Dim b_DefineFinalImped As New SiteBoolean, FinalImped As New PinListData
    
    ''20160712 - If match taget Imped just store the trim code
    Dim b_MatchTagetImped As New SiteBoolean, b_DisplayImped As New SiteBoolean, StoredTargetTrimCode As New DSPWave
    
    b_MatchTagetImped = False
    b_DisplayImped = False
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    
    Dim StoreEachTrimImped() As New PinListData
    Dim StoreEachTrimCode() As New DSPWave
    ReDim StoreEachTrimImped(DigSrc_Sample_Size + 1) As New PinListData
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i
    
    ''20161230-Add loop to trim code for each pin
    Dim PerPinIndex As Long
    PerPinIndex = 0
    
    ''20170117 - For long length fixed code
    Dim InitialTrimDSPWave As New DSPWave
    Dim InitialFixedDSPWave As New DSPWave
    Dim Trim_DigSrc_Sample_Size As Long
    
    '' 20170117-Evaluate for ForceVolt
    Dim SplitForceVolt() As String
    SplitForceVolt = Split(StrForceVolt, ",")
    Dim ForceVolt As String
    Call HIP_Evaluate_ForceVal(SplitForceVolt)
    For i = 0 To UBound(SplitForceVolt)
        If i = 0 Then
            ForceVolt = SplitForceVolt(i)
        Else
            ForceVolt = ForceVolt & "," & SplitForceVolt(i)
        End If
    Next i
    
    TrimStart = Replace(TrimStart, "&", "")
    ReDim TrimCode(Pin_Cnt)
    ReDim TrimmedImpedance(Pin_Cnt)
    Dim PinNumber As Long
    ''20170210-Specified item to store trim code to fuse by 1 pin
    Trim_DigSrc_Sample_Size = DigSrc_Sample_Size - Fixed_DigSrc_Sample_Size
    Dim wkds_StoreTrimCodeToDict_DEC As New DSPWave
    Dim wkds_StoreTrimCodeToDict_BIN As New DSPWave
    wkds_StoreTrimCodeToDict_DEC.CreateConstant 0, 1, DspLong
    
    If Fixed_DigSrc_Equation <> "" Then
        wkds_StoreTrimCodeToDict_BIN.CreateConstant 0, Trim_DigSrc_Sample_Size, DspLong
    Else
        wkds_StoreTrimCodeToDict_BIN.CreateConstant 0, DigSrc_Sample_Size, DspLong
    End If
    
    Dim b_wkds_Store_Flag As Boolean
    b_wkds_Store_Flag = False

    For Each Pin In Pin_Ary

        If LCase(TrimStoreName) = LCase("wkds_1") And LCase(Pin) = LCase("DDR0_ADDR_SOP_P0") Then
            b_wkds_Store_Flag = True
        ElseIf LCase(TrimStoreName) = LCase("wkds_2") And LCase(Pin) = LCase("DDR0_ADDR_SOP_P1") Then
            b_wkds_Store_Flag = True
        ElseIf LCase(TrimStoreName) = LCase("wkds_3") And LCase(Pin) = LCase("DDR1_ADDR_SOP_P0") Then
            b_wkds_Store_Flag = True
        ElseIf LCase(TrimStoreName) = LCase("wkds_4") And LCase(Pin) = LCase("DDR1_ADDR_SOP_P1") Then
            b_wkds_Store_Flag = True
        End If
        
        If Fixed_DigSrc_Equation <> "" Then
            For Each site In TheExec.sites.Active
                Call Create_DigSrc_Data(DigSrc_pin, Fixed_DigSrc_DataWidth, Fixed_DigSrc_Sample_Size, Fixed_DigSrc_Equation, Fixed_DigSrc_Assignment, InitialFixedDSPWave, site)
            Next site
            Dec_TrimStart_1st = Bin2Dec(TrimStart)
            
            InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong
''            Trim_DigSrc_Sample_Size = DigSrc_Sample_Size - Fixed_DigSrc_Sample_Size
            
            Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, Trim_DigSrc_Sample_Size, InitialTrimDSPWave)
           
           Call rundsp.CombineDSPWave(InitialFixedDSPWave, InitialTrimDSPWave, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, InDSPwave)
        Else
            Dec_TrimStart_1st = Bin2Dec(TrimStart)
            
            InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong
        
            Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, DigSrc_Sample_Size, InDSPwave)
        End If
        
        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPwave)

        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("First Time Setup")
        '' Debug use
        For Each site In TheExec.sites
            OutputTrimCode = ""
            For k = 0 To InDSPwave(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
            Next k
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
        Next site
        
        For Each site In TheExec.sites
            StoreEachTrimCode(0)(site).Data = InDSPwave(site).Data
        Next site
        
        Call TheHdw.Patterns(PattArray(0)).start
        
        Call SubMeasR(CPUA_Flag_In_Pat, CStr(Pin), ForceVolt, MeasureImped, b_IsDifferential, b_PD_Mode)
        
        StoreEachTrimImped(0) = MeasureImped
        
        b_HighThanTargetImped = MeasureImped.Math.Subtract(TrimTarget).compare(LessThan, 0)
        PastDSPWave = InDSPwave
        
        TestNameInput = "Imped"
        TestLimitIndex = 0
        
        '' 20160712 - Modify to use WriteComment to display output Impeduency.
        For Each site In TheExec.sites
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Pin " & CStr(Pin) & " Impedence = " & FormatNumber((MeasureImped.Pins(0).Value(site)), 3) & " Ohm")
            
            TrimmedImpedance(PinNumber) = MeasureImped.Pins(0).Value(site)
            TrimCode(PinNumber) = 0
             For X = 0 To GetbitNumber - 1
                TrimCode(PinNumber) = InDSPwave.Element(InDSPwave.SampleSize - X - 1) * 2 ^ (GetbitNumber - 1 - X) + TrimCode(PinNumber)
            Next X
            
        Next site
        
        '' 20160712 - Compare Measure Impeduency whether match target Imped
        b_MatchTagetImped = MeasureImped.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
        
        b_DisplayImped = b_DisplayImped.LogicalOr(b_MatchTagetImped)
        For Each site In TheExec.sites
            If b_MatchTagetImped(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                StoredTargetTrimCode(site).Data = InDSPwave(site).Data
                b_StopTrimCodeProcess(site) = True
            End If
        Next site
       If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
        
        ''========================================================================================
        ''20161128 Pre check Min/Max trim code process
        Dim b_KeepGoing As New SiteBoolean
        Dim PreviousImped As New PinListData
        If PreCheckMinMaxTrimCode = True Then
            PreviousImped = MeasureImped
            Call rundsp.PreCheckMinMaxTrimCode(b_HighThanTargetImped, InDSPwave)
            Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPwave)
    
            Call TheHdw.Patterns(PattArray(0)).start
        
            Call SubMeasR(CPUA_Flag_In_Pat, CStr(Pin), ForceVolt, MeasureImped, b_IsDifferential, b_PD_Mode)

            If TheExec.TesterMode = testModeOffline Then
                MeasureImped.Pins(CStr(Pin)).Value(0) = MeasureImped.Pins(CStr(Pin)).Value(0) + 10
                MeasureImped.Pins(CStr(Pin)).Value(1) = MeasureImped.Pins(CStr(Pin)).Value(1) - 10
            End If
            
            For Each site In TheExec.sites
                OutputTrimCode = ""
                For k = 0 To InDSPwave(site).SampleSize - 1
                    OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
                Next k
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
            Next site
            
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site " & site & " Pin " & CStr(Pin) & " Impedence = " & FormatNumber((MeasureImped.Pins(0).Value(site)), 3) & " Ohm")
                If Abs(TrimmedImpedance(PinNumber) - TrimTarget) > Abs(MeasureImped.Pins(0).Value(site) - TrimTarget) Then
                    TrimmedImpedance(PinNumber) = MeasureImped.Pins(0).Value(site)
                    
                     TrimCode(PinNumber) = 0
                     For X = 0 To GetbitNumber - 1
                        TrimCode(PinNumber) = InDSPwave.Element(InDSPwave.SampleSize - X - 1) * 2 ^ (GetbitNumber - 1 - X) + TrimCode(PinNumber)
                    Next X
                End If
            
            Next site
            
            For Each site In TheExec.sites
                If b_HighThanTargetImped(site) = True Then
                    b_KeepGoing(site) = MeasureImped.Math.Subtract(PreviousImped).compare(LessThan, 0)
                Else
                    b_KeepGoing(site) = MeasureImped.Math.Subtract(PreviousImped).compare(GreaterThan, 0)
                End If
            Next site
    
            Dim PreCheckBinStr As String, PreCheckDecVal As Double
            For Each site In TheExec.sites
                If b_KeepGoing(site) = False Then
                    b_StopTrimCodeProcess(site) = True
                    PreCheckBinStr = ""
                    StoredTargetTrimCode(site).Data = InDSPwave(site).Data
                    For i = 0 To StoredTargetTrimCode(site).SampleSize - 1
                        PreCheckBinStr = PreCheckBinStr & StoredTargetTrimCode.Element(i)
                    Next i
                    PreCheckDecVal = Bin2Dec_rev_Double(PreCheckBinStr)
    
                End If
            Next site
        End If
        
        ''========================================================================================
        Dim b_ControlNextBit As Boolean
        b_ControlNextBit = False
        Dim b_FirstExecution As Boolean
        b_FirstExecution = False
        StoreEachIndex = 1

        ''20170103-Setup b_KeepGoing to true if PreCheckMinMaxTrimCode=false
        If PreCheckMinMaxTrimCode = False Then
            b_KeepGoing = True
        End If
        
        If b_KeepGoing.All(False) Then
        Else
    
            For i = 0 To ExecutionMax
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
                StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
                SrcStartBit = StrSeparatebyColon(0)
                SrcEndBit = StrSeparatebyColon(1)
                If b_HasFIXED_InTrimFormat Then
                    If UBound(StrSeparatebyColon) = 1 Then
                        If i = 0 Then
                            b_FirstExecution = True
                        Else
                            b_FirstExecution = False
                            SrcStartBit = SrcStartBit
                        End If
                    End If
                
                Else
                    If i = 0 Then
                        b_FirstExecution = True
                    Else
                        b_FirstExecution = False
                        SrcStartBit = SrcStartBit + 1
                    End If
                End If
                If b_ExecuteTrimCode(i) = True Then
                    For j = SrcStartBit To SrcEndBit Step -1
                    
                        If b_FirstExecution = True Then
                            b_ControlNextBit = True
                            If j = SrcEndBit Then
                                b_ControlNextBit = False
                            End If
                        Else
                        ''20160716-Control next bit to 1 no matter first or last progress
                            b_ControlNextBit = True
            ''                b_ControlNextBit = False
                            If j = SrcEndBit Then
                                b_ControlNextBit = False
                            End If
                        End If
            
''                        If b_FirstExecution = True And j = SrcEndBit Then
                        If j = SrcEndBit Then
                            Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPwave)
            ''            ElseIf b_FirstExecution = False And j = SrcStartBit Then
            ''                j = SrcStartBit + 1
                        Else
                            Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetImped, j, b_ControlNextBit, InDSPwave)
                        End If
                        
                        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPwave)
                        
                        For Each site In TheExec.sites
                            StoreEachTrimCode(StoreEachIndex)(site).Data = InDSPwave(site).Data
                        Next site
                    
                        '' Debug use
                        '' ==============================================================================================
                        '' 20160716 - Modify trim code rule
                        
                        If b_FirstExecution = True And gl_Disable_HIP_debug_log = False Then
                            If j = SrcEndBit Then
                                TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                            Else
                                TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                            End If
                        ElseIf gl_Disable_HIP_debug_log = False Then
                            If j = SrcEndBit Then
                                TheExec.Datalog.WriteComment ("Setup Bit " & j)
                            Else
                                TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                            End If
                        End If
                        
                        For Each site In TheExec.sites
        
                            If b_KeepGoing(site) = True Then
                                OutputTrimCode = ""
                                For k = 0 To InDSPwave(site).SampleSize - 1
                                    OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
                                Next k
                                
                                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
                            End If
        
                        Next site
                        '' ==============================================================================================
                        
                        Call TheHdw.Patterns(PattArray(0)).start
        
                        Call SubMeasR(CPUA_Flag_In_Pat, CStr(Pin), ForceVolt, MeasureImped, b_IsDifferential, b_PD_Mode)
                        
                        If TheExec.TesterMode = testModeOffline Then
                            Dim SimuIndex As Long
                            SimuIndex = TestLimitIndex
                            If SimuIndex >= 3 Then
                                SimuIndex = 3
                            End If
    
                            MeasureImped.Pins(CStr(Pin)).Value(0) = MeasureImped.Pins(CStr(Pin)).Value(0) + (SimuIndex * 1) + ((PerPinIndex + 1) * 1 / (PerPinIndex + 1))
                            MeasureImped.Pins(CStr(Pin)).Value(1) = MeasureImped.Pins(CStr(Pin)).Value(1) - (SimuIndex * 1) + ((PerPinIndex + 1) * 1 / (PerPinIndex + 1))
    
                        End If
                        
                        If j = SrcEndBit + 1 Then
                            MeasureImped_F1 = MeasureImped
                        ElseIf j = SrcEndBit Then
                            MeasureImped_F2 = MeasureImped
                        End If
                        
                        StoreEachTrimImped(StoreEachIndex) = MeasureImped
                        StoreEachIndex = StoreEachIndex + 1
                        
                        If j = SrcEndBit Then
                            b_HighThanTargetImped = False
                            b_HighThanTargetImped = MeasureImped_F1.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureImped_F2.Math.Subtract(TrimTarget).Abs)
''                            Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetImped, j, b_ControlNextBit, InDSPWave)
                            PastDSPWave = InDSPwave
                        Else
                            b_HighThanTargetImped = False
                            b_HighThanTargetImped = MeasureImped.Math.Subtract(TrimTarget).compare(LessThan, 0)
                            PastDSPWave = InDSPwave
                        End If
            
                        TestLimitIndex = TestLimitIndex + 1
                        
                        '' 20160712 - Modify to use WriteComment to display output Impeduency.
                        For Each site In TheExec.sites
                            If b_KeepGoing(site) = True Then
                                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Pin " & CStr(Pin) & " Impedence = " & FormatNumber((MeasureImped.Pins(0).Value(site)), 3) & " Ohm")
                                If Abs(TrimmedImpedance(PinNumber) - TrimTarget) > Abs(MeasureImped.Pins(0).Value(site) - TrimTarget) Then
                                    
                                    TrimCode(PinNumber) = 0
                                    TrimmedImpedance(PinNumber) = MeasureImped.Pins(0).Value(site)
                                     For X = 0 To GetbitNumber - 1
                                        TrimCode(PinNumber) = InDSPwave.Element(InDSPwave.SampleSize - X - 1) * 2 ^ (GetbitNumber - 1 - X) + TrimCode(PinNumber)
                                    Next X
                                End If
                            End If
                        Next site
                        
                        ''20160716 - Modify display info sequence when source bit in the section end
                        If j = SrcEndBit Then
                            For Each site In TheExec.sites
        
                                If b_KeepGoing(site) = True And gl_Disable_HIP_debug_log = False Then
                                    TheExec.Datalog.WriteComment ("Site " & site & " Pin " & CStr(Pin) & " R" & LastSectionF1F2_Index + 1 & " Impedence = " & FormatNumber((MeasureImped_F1.Pins(0).Value(site)), 3) & " Ohm")
                                    TheExec.Datalog.WriteComment ("Site " & site & " Pin " & CStr(Pin) & " R" & LastSectionF1F2_Index + 2 & " Impedence = " & FormatNumber((MeasureImped_F2.Pins(0).Value(site)), 3) & " Ohm")
                                End If
        
                            Next site
                            LastSectionF1F2_Index = LastSectionF1F2_Index + 2
                        End If
                        
                        '' 20160712 - Compare Measure Impeduency whether match target Imped
                        b_MatchTagetImped = MeasureImped.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                        b_DisplayImped = b_DisplayImped.LogicalOr(b_MatchTagetImped)
                        For Each site In TheExec.sites
                            If b_KeepGoing(site) = True Then
                                If b_MatchTagetImped(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                                    StoredTargetTrimCode(site).Data = InDSPwave(site).Data
                                    b_StopTrimCodeProcess(site) = True
                                End If
                            End If
                        Next site
                        ''20161128-Stop trim code process if found out match code of all site
                        If TrimPrcocessAll = False Then
                            If b_StopTrimCodeProcess.All(True) Then
                                Exit For
                            End If
                        End If
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
                    Next j
                End If
            Next i
        End If
    
        ''============================================================================
        ''20161128 Findout mimiumn trim code
        Dim CloseTargetImped As New PinListData
        Dim DiffValue As New SiteDouble, PreviousDiffValue As New SiteDouble, CloseIndex As New SiteLong
        
        Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteDouble
        Dim BinStr As String
        Dim CloseTargetTrimCode As New DSPWave
        Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
        Dim b_FirstTimeSwitch As Boolean
        
        
        If b_KeepGoing.All(False) Then
        Else
''            If PerPinIndex = 0 Then
                Set CloseTargetTrimCode = Nothing
''                If Fixed_DigSrc_Equation <> "" Then
''                    CloseTargetTrimCode.CreateConstant 0, Trim_DigSrc_Sample_Size, DspLong
''                Else
                    CloseTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
''                End If
''            Else
''                For Each Site In TheExec.Sites
''                    CloseTargetTrimCode.Clear
''                Next Site
''            End If
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If StoredTargetTrimCode(site).CalcSum = 0 Then
                        b_UseMinTrim(site) = True
                    End If
                End If
            Next site
                
            If UseMinimumTrimCode = True Then
                b_UseMinTrim = True
            End If
                
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_UseMinTrim(site) = True Then
                        '' Findout minimum difference value
    
                        For i = 0 To StoreEachIndex - 1
                            DiffValue(site) = Abs(StoreEachTrimImped(i).Pins(0).Value(site) - TrimTarget)
                            If DiffValue(site) <= PreviousDiffValue(site) Then
                                CloseIndex(site) = i
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                            If i = 0 Then
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                        Next i
                        '' Transfer to decimal value to findout minimum code
                        PreviousDecVal = 0
                        DecVal = 0
                        b_FirstTimeSwitch = False
    
                        For i = 0 To StoreEachIndex - 1
                            BinStr = ""
                            If Abs(StoreEachTrimImped(i).Pins(0).Value(site) - TrimTarget) <= MinDiffVal(site) Then
                                For j = 0 To StoreEachTrimCode(i)(site).SampleSize - 1
                                    BinStr = BinStr & StoreEachTrimCode(i)(site).Element(j)
                                Next j
                                DecVal = Bin2Dec_rev_Double(BinStr)
                               
                                If DecVal < PreviousDecVal Then
                                    MinDecVal = DecVal
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                End If
                                PreviousDecVal = DecVal
                                If b_FirstTimeSwitch = False Then
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                    b_FirstTimeSwitch = True
                                End If
                            End If
                        Next i
                    End If
                End If
            Next site
        End If
        
        For Each site In TheExec.sites
            If b_KeepGoing(site) = True Then
                If b_UseMinTrim(site) = True Then
                    StoredTargetTrimCode(site).Data = CloseTargetTrimCode(site).Data
                Else
                    StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
                End If
            Else
                StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
            End If
        Next site
        
        For Each site In TheExec.sites
            OutputTrimCode = ""
            For k = 0 To StoredTargetTrimCode(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(StoredTargetTrimCode(site).Element(k))
            Next k
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Pin " & CStr(Pin) & " Final Output Trim Code = " & OutputTrimCode)
        Next site
        
        ''20170103-Store per pin final trim code
        StorePerPinFinalTrimCode(PerPinIndex) = StoredTargetTrimCode
        PerPinIndex = PerPinIndex + 1
        LastSectionF1F2_Index = 0
        
        PinNumber = PinNumber + 1
        
        ''20170210-Specified item to store trim code to fuse by 1 pin
        If b_wkds_Store_Flag Then
            If Fixed_DigSrc_Equation <> "" Then
                Call rundsp.SelectCertainBitsToDec(StoredTargetTrimCode, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, wkds_StoreTrimCodeToDict_DEC)
                Call rundsp.DSPWaveDecToBinary(wkds_StoreTrimCodeToDict_DEC, Trim_DigSrc_Sample_Size, wkds_StoreTrimCodeToDict_BIN)
                b_wkds_Store_Flag = False
            End If
        End If
    Next Pin
    ''============================================================================
    ''20161230 - Average per pin trim code and store it
    Dim TrimCodeForTotalPin_Dec As New DSPWave
    Dim TrimCodeForTotalPin_Bin As New DSPWave
    TrimCodeForTotalPin_Dec.CreateConstant 0, 1, DspLong
    If Fixed_DigSrc_Equation <> "" Then
        TrimCodeForTotalPin_Bin.CreateConstant 0, Trim_DigSrc_Sample_Size, DspLong
    Else
        TrimCodeForTotalPin_Bin.CreateConstant 0, DigSrc_Sample_Size, DspLong
    End If
    For i = 0 To UBound(StorePerPinFinalTrimCode_Dec)
    
        If Fixed_DigSrc_Equation <> "" Then
            '' 20170117-only select trim code,  not select all to convert decimal
            Call rundsp.SelectCertainBitsToDec(StorePerPinFinalTrimCode(i), Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, StorePerPinFinalTrimCode_Dec(i))
        Else
            Call rundsp.ConvertToLongAndSerialToParrel(StorePerPinFinalTrimCode(i), DigSrc_Sample_Size, StorePerPinFinalTrimCode_Dec(i))
        End If
        Call rundsp.DSP_Add(TrimCodeForTotalPin_Dec, StorePerPinFinalTrimCode_Dec(i))
    Next i
    Dim DenominatorConstant As Double
    DenominatorConstant = UBound(StorePerPinFinalTrimCode_Dec) + 1
    If DenominatorConstant = 0 Then
        TheExec.Datalog.WriteComment ("Error! Divide 0.")
        Exit Function
    End If
''    Call rundsp.DSP_DivideConstant(TrimCodeForTotalPin_Dec, DenominatorConstant)
    For Each site In TheExec.sites
        TrimCodeForTotalPin_Dec(site).Element(0) = Int(CDbl(TrimCodeForTotalPin_Dec(site).Element(0) / DenominatorConstant) + 0.5)
    Next site
    
''    For Each Site In TheExec.Sites
''        If TrimCodeForTotalPin_Dec(Site).Element(0) > 192 Then
''            TrimCodeForTotalPin_Dec(Site).Element(0) = TrimCodeForTotalPin_Dec(Site).Element(0) - 192
''        End If
''    Next Site
    
    If Fixed_DigSrc_Equation <> "" Then
        Call rundsp.DSPWaveDecToBinary(TrimCodeForTotalPin_Dec, Trim_DigSrc_Sample_Size, TrimCodeForTotalPin_Bin)
        ''20170214 - If trim code size = 8 , set element 7 and 6 to 0.
        For Each site In TheExec.sites
            If TrimCodeForTotalPin_Bin(site).SampleSize = 8 Then
                TrimCodeForTotalPin_Bin(site).Element(6) = 0
                TrimCodeForTotalPin_Bin(site).Element(7) = 0
            End If
        Next site
    Else
        Call rundsp.DSPWaveDecToBinary(TrimCodeForTotalPin_Dec, DigSrc_Sample_Size, TrimCodeForTotalPin_Bin)
    End If
    ''============================================================================
    ''20170104 - TestLimit for found code and measured R for each pin
    i = 0
    Dim TrimCodeDSP_DEC() As New DSPWave
    ReDim TrimCodeDSP_DEC(Pin_Cnt) As New DSPWave
    Dim TrimCodeDSP_BIN() As New DSPWave
    ReDim TrimCodeDSP_BIN(Pin_Cnt) As New DSPWave
    
    For Each Pin In Pin_Ary
        Set TrimCodeDSP_DEC(i) = Nothing
        TrimCodeDSP_DEC(i).CreateConstant 0, 1, DspLong
        For Each site In TheExec.sites
''            If TrimCode(i) > 192 Then
''                TrimCode(i) = TrimCode(i) - 192
''            End If

''          ''20170214 - If trim code size = 8 , set element 7 and 6 to 0.
            TrimCodeDSP_DEC(i)(site).Element(0) = TrimCode(i)(site)
        Next site
        Call rundsp.DSPWaveDecToBinary(TrimCodeDSP_DEC(i), GetbitNumber, TrimCodeDSP_BIN(i))
        If GetbitNumber = 8 Then
            For Each site In TheExec.sites
                TrimCodeDSP_BIN(i)(site).Element(6) = 0
                TrimCodeDSP_BIN(i)(site).Element(7) = 0
            Next site
        End If
        Call rundsp.BinToDec(TrimCodeDSP_BIN(i), TrimCodeDSP_DEC(i))
''        TheExec.Flow.TestLimit TrimCode(i), , , , , , , , CStr(pin), , , , , , , tlForceNone
        
        TestNameInput = Report_TName_From_Instance("C", "", "", 0)
        TheExec.Flow.TestLimit TrimCodeDSP_DEC(i).Element(0), , , , , , , , TestNameInput, , , , , , , tlForceNone

        
        TestNameInput = Report_TName_From_Instance("R", "", "", 0)
        TheExec.Flow.TestLimit TrimmedImpedance(i), , , , , , unitCustom, , TestNameInput, , , , , "ohm", , tlForceNone

        i = i + 1
    Next Pin
    ''============================================================================
    If TrimStoreName <> "" Then
        ''20170210-Specified item to store trim code to fuse by 1 pin
        If LCase(TrimStoreName) Like LCase("wkds_*") Then

            Call Checker_StoreDigCapAllToDictionary(TrimStoreName, wkds_StoreTrimCodeToDict_BIN)
        Else
            Call Checker_StoreDigCapAllToDictionary(TrimStoreName, TrimCodeForTotalPin_Bin)
        End If
    End If
    
    Call HardIP_WriteFuncResult



''    For Each site In TheExec.Sites
''        If TrimCodeForTotalPin_Dec(site).Element(0) > 192 Then
''            TrimCodeForTotalPin_Dec(site).Element(0) = TrimCodeForTotalPin_Dec(site).Element(0) - 192
''        End If
''    Next site
    ''20170214 - If trim code size = 8 , set element 7 and 6 to 0.
    Dim TEMP_TrimCodeForTotalPin_Bin As New DSPWave
    Call rundsp.DSPWaveDecToBinary(TrimCodeForTotalPin_Dec, Trim_DigSrc_Sample_Size, TEMP_TrimCodeForTotalPin_Bin)
    For Each site In TheExec.sites
        If TEMP_TrimCodeForTotalPin_Bin(site).SampleSize = 8 Then
            TEMP_TrimCodeForTotalPin_Bin(site).Element(6) = 0
            TEMP_TrimCodeForTotalPin_Bin(site).Element(7) = 0
        End If
    Next site
    Call rundsp.BinToDec(TEMP_TrimCodeForTotalPin_Bin, TrimCodeForTotalPin_Dec)
    
    TestNameInput = Report_TName_From_Instance("C", "", "AverageTrimCode", 0)
    TheExec.Flow.TestLimit TrimCodeForTotalPin_Dec.Element(0), 0, 2 ^ Trim_DigSrc_Sample_Size - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    ''20170213-Source average trim code to measure impedence again
    If Fixed_DigSrc_Equation <> "" Then
        Call rundsp.CombineDSPWave(InitialFixedDSPWave, InitialTrimDSPWave, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, InDSPwave)
    End If

    '' 20170210 - Source final code to do re-measurement
''    If Fixed_DigSrc_Equation <> "" Then
''        Call rundsp.CombineDSPWave(InitialFixedDSPWave, TrimCodeForTotalPin_Bin, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, InDSPWave)
''    End If
''    For Each pin In pin_ary
''        Call SetupDigSrcDspWave(PattArray(0), DigSrc_Pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPWave)
''
''        Call thehdw.Patterns(PattArray(0)).start
''        Call SubMeasR(CPUA_Flag_In_Pat, CStr(pin), ForceVolt, MeasureImped, b_IsDifferential, b_PD_Mode)
''        TheExec.Flow.TestLimit resultVal:=MeasureImped, unit:=unitCustom, customUnit:="ohm", Tname:="SourceAverCode" & "_Pin_" & pin, ForceResults:=tlForceFlow
''
''    Next pin
    
    Dim SplitByComma() As String
    Dim DictName_FUSE As String
    Dim sl_FUSE_Val As New SiteLong
    '' 20170119 - Process calculate equation by dictionary.
    If Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Calc_Eqn)
        If UCase(Calc_Eqn) Like UCase("*ADDRIO_TrimCodeAverage*") Then
            SplitByComma() = Split(Calc_Eqn, ",")
            DictName_FUSE = SplitByComma(UBound(SplitByComma))
            DictName_FUSE = Replace(DictName_FUSE, ")", "")
            Call DictDSPToSiteLong(DictName_FUSE, sl_FUSE_Val, TrimFuseName)
        End If
    End If
    '' 20170704- Comment this
''    If TrimFuseName <> "" And TrimFuseName <> "addr-wkds_u" Then
''        Call HIP_eFuse_Write("ECID", TrimFuseName, sl_FUSE_Val)
''    End If
    If TheExec.TesterMode = testModeOffline Then
    Else
        If TrimFuseName <> "" And TrimFuseTypeName <> "" Then
            ''Call HIP_eFuse_Write(TrimFuseTypeName, TrimFuseName, sl_FUSE_Val) ''set fuse information from flow
        End If
    End If
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    '' NOTE : Efuse write TrimCodeForTotalPin_Dec ( DSPWave )
    
''    Dim ConvertedDataWf As New DSPWave
''    rundsp.ConvertToLongAndSerialToParrel StoredTargetTrimCode, DigSrc_Sample_Size, ConvertedDataWf

''    Call SetupDigSrcDspWave(PattArray(0), DigSrc_Pin, "TrimCodeImped", DigSrc_Sample_Size, TrimCodeForTotalPin_Bin)
''
''    Call TheHdw.Patterns(PattArray(0)).start
''
''    Call SubMeasR(CPUA_Flag_In_Pat, MeasR_Pins_SingleEnd.Value, ForceVolt, MeasureImped, b_IsDifferential)
''
''    TheExec.Flow.TestLimit resultVal:=MeasureImped, unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput & "Final", ForceResults:=tlForceFlow
    
    
''    If TheExec.TesterMode = testModeOffline Then
''    Else
''        '' eFUSE
''        For Each Site In TheExec.Sites.Active
''            Dim PassFlag_LPRO As New SiteBoolean
''
''            If CurrentJobName_U Like "*FT*" Then
''                TheExec.Datalog.WriteComment ""
''                ConvertedDataWf(Site).Element(0) = auto_eFuse_GetReadDecimal("ECID", "OSC", True)
''                TheExec.Datalog.WriteComment ""
''                For i = 0 To DigSrc_Sample_Size - 1
''                    '' 20161110 - Hint! Need to check Src_DSPWave
''                    Src_DSPWave(Site).Element(i) = ConvertedDataWf(Site).Element(0) Mod 2
''                    ConvertedDataWf(Site).Element(0) = ConvertedDataWf(Site).Element(0) \ 2
''                Next i
''            Else
''                If TheHdw.Digital.Patgen.PatternBurstPassed(Site) = False Then 'Pattern Fail
''                    PassFlag_LPRO(Site) = False
''                Else
''                    PassFlag_LPRO(Site) = True
''                End If
''
''                If UCase(TrimFuseName) = "OSC" Then
''                    Call auto_eFuse_SetPatTestPass_Flag("ECID", "OSC", PassFlag_LPRO(Site))
''                    Call auto_eFuse_SetWriteDecimal("ECID", "OSC", ConvertedDataWf(Site).Element(0))
''                End If
''            End If
''        Next Site
''    End If
    
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeImped function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function


Public Function HIP_eFuse_Read(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, dspwavesize As Long, Optional Efuse_Read_Dec_Flag As Boolean = False, Optional Dict_Store_Dec_Name As String = "", _
                                Optional Calc_code As String = "") As Long

    ' Parameter : eFuse Block , eFuse Variable , data , Data Width
    ' Create dictionary , if exist then remove and re-create
    ' MUST :  if necessary , we can set limit if read out value = 0 then bin out .

    Dim site As Variant
    Dim Read_Code As New DSPWave
    Dim Read_Value As New DSPWave
    Dim Efuse_Value As New SiteLong
    Dim TempVal As Long
    Dim Efuse_Value_Chk As New SiteVariant
    Dim i As Long

    On Error GoTo errHandler

    Read_Code.CreateConstant 0, dspwavesize

    If Efuse_Read_Dec_Flag = True Then
        Read_Value.CreateConstant 0, 1
    End If

    For Each site In TheExec.sites

        Efuse_Value(site) = auto_eFuse_GetReadDecimal(FuseType, m_catename, True)
'''''        Efuse_Value(Site) = CLng(Site) + 8
'''----------cal get fused code
        If Calc_code <> "" Then
        'Calc_code = "minus,100"
            If Split(Calc_code, ",")(0) = "minus" Then
                Efuse_Value = Efuse_Value.Subtract(Split(Calc_code, ",")(1))
            End If
        End If
'''----------cal get fused code
        If Efuse_Read_Dec_Flag = True Then
            Read_Value.Element(0) = Efuse_Value(site)
        End If

        TempVal = Efuse_Value(site)
        For i = 0 To dspwavesize - 1
            Read_Code.Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i

        If UCase(TheExec.DataManager.instanceName) Like "MTRGR_VREF*" Then
            Dim Read_Code_Inverse As New DSPWave: Read_Code_Inverse = Read_Code
            If Read_Code.Element(dspwavesize - 1) = 0 Then
                Read_Code_Inverse.Element(dspwavesize - 1) = 1
            Else
                Read_Code_Inverse.Element(dspwavesize - 1) = 0
            End If
        End If

        If Efuse_Value(site) = 0 Then
        'If Read out value = 0 then bin out
            Efuse_Value_Chk(site) = 0
        Else
            Efuse_Value_Chk(site) = 1
        End If

    Next site

    TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=1, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceNone

    Call AddStoredCaptureData(Dict_Store_Code_Name, Read_Code)
    If UCase(TheExec.DataManager.instanceName) Like "MTRGR_VREF*" Then: Call AddStoredCaptureData("var_bg_rref", Read_Code_Inverse)
    If Efuse_Read_Dec_Flag = True Then
        Call AddStoredCaptureData(Dict_Store_Dec_Name, Read_Value)
    End If

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Read"
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function MTR_Verification_Calculate(SensorName As String, Temperature As String, FusedCoeffDicName_1 As String, FusedCoeffDicName_2 As String, SetInformation As DSPWave, _
                                        Aininformation As DSPWave, Aixinformation As DSPWave, PiUInformation As DSPWave, LevelsRecord As String) As Long
 
        Dim i As Long
        Dim row As Long
        Dim col As Long
        Dim bitIter As Long
        Dim tmpCount As Long
        Dim testName As String
        Dim ScanOffset As Integer
        Dim ElementCNT As Integer
        Dim PiUCntTempA As Integer
        Dim PiUCntTempB As Integer
        Dim VoltageName() As String
        Dim tmpValue_ROT  As Double
        Dim tmpValue_ROV  As Double
        Dim InheritanceCnt As Integer
        Dim temp_rowVal_f As New SiteDouble
        Dim decimalPlaces As Long: decimalPlaces = 4
        VoltageName = Split(LevelsRecord, ",")
        
        Dim actualROTMatrix As New DSPWave
        Dim actualROVMatrix As New DSPWave
        Dim RotMatrixDicName As String
        Dim RovMatrixDicName As String
    
        Dim FusedCoeffDicNameTemp_1 As String
        Dim FusedCoeffDicNameTemp_2 As String
        Dim readFusedCoeffDspWave1 As New DSPWave
        Dim readFusedCoeffDspWave2 As New DSPWave
        
        Dim Fused_ROT_Decimal_Vector As New DSPWave
        Dim Fused_ROV_Decimal_Vector As New DSPWave
        Fused_ROT_Decimal_Vector.CreateConstant 0, 4, DspDouble
        Fused_ROV_Decimal_Vector.CreateConstant 0, 3, DspDouble
                
        Dim Output_ROT_Freq_Vector As New DSPWave
        Dim Output_ROV_Freq_Vector As New DSPWave
        Output_ROT_Freq_Vector.CreateConstant 0, 20, DspDouble
        Output_ROV_Freq_Vector.CreateConstant 0, 20, DspDouble
            
        Dim Difference_ROT_Freq_Vector As New DSPWave
        Dim Difference_ROV_Freq_Vector As New DSPWave
        Difference_ROT_Freq_Vector.CreateConstant 0, 20, DspDouble
        Difference_ROV_Freq_Vector.CreateConstant 0, 20, DspDouble
    
        RotMatrixDicName = "Freq" + "_" + SensorName + "_" + "rot" + "_" + Temperature + "c"
        RovMatrixDicName = "Freq" + "_" + SensorName + "_" + "rov" + "_" + Temperature + "c"
    
        actualROTMatrix = GetStoredCaptureData(RotMatrixDicName)
        actualROVMatrix = GetStoredCaptureData(RovMatrixDicName)
    
        FusedCoeffDicNameTemp_1 = FusedCoeffDicName_1 + "_" + "Freq" + "_" + SensorName + "_" + "rot" + "_" + Temperature + "c"
        FusedCoeffDicNameTemp_2 = FusedCoeffDicName_2 + "_" + "Freq" + "_" + SensorName + "_" + "rov" + "_" + Temperature + "c"
        readFusedCoeffDspWave1 = GetStoredCaptureData(FusedCoeffDicNameTemp_1)
        readFusedCoeffDspWave2 = GetStoredCaptureData(FusedCoeffDicNameTemp_2)
        
        If Temperature = 25 Then
            For Each site In TheExec.sites
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
                For bitIter = 0 To readFusedCoeffDspWave1.SampleSize - 1
                    If bitIter = 14 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(0) = tmpValue_ROT * (Aixinformation.Element(0) - Aininformation.Element(0)) + Aininformation.Element(0)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 28 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(1) = tmpValue_ROT * (Aixinformation.Element(1) - Aininformation.Element(1)) + Aininformation.Element(1)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 42 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(2) = tmpValue_ROT * (Aixinformation.Element(2) - Aininformation.Element(2)) + Aininformation.Element(2)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 56 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(3) = tmpValue_ROT * (Aixinformation.Element(3) - Aininformation.Element(3)) + Aininformation.Element(3)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    Else
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        tmpCount = tmpCount + 1
                    End If
                Next bitIter
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
                For bitIter = 0 To readFusedCoeffDspWave2.SampleSize - 1
                    If bitIter = 14 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(0) = tmpValue_ROV * (Aixinformation.Element(4) - Aininformation.Element(4)) + Aininformation.Element(4)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    ElseIf bitIter = 28 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(1) = tmpValue_ROV * (Aixinformation.Element(5) - Aininformation.Element(5)) + Aininformation.Element(5)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    ElseIf bitIter = 42 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(2) = tmpValue_ROV * (Aixinformation.Element(6) - Aininformation.Element(6)) + Aininformation.Element(6)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    Else
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        tmpCount = tmpCount + 1
                    End If
                Next bitIter
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
            Next site
        ElseIf Temperature = 85 Then
            For Each site In TheExec.sites
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
                For bitIter = 0 To readFusedCoeffDspWave1.SampleSize - 1
                    If bitIter = 14 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(0) = tmpValue_ROT * (Aixinformation.Element(7) - Aininformation.Element(7)) + Aininformation.Element(7)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 28 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(1) = tmpValue_ROT * (Aixinformation.Element(8) - Aininformation.Element(8)) + Aininformation.Element(8)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 42 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(2) = tmpValue_ROT * (Aixinformation.Element(9) - Aininformation.Element(9)) + Aininformation.Element(9)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 56 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(3) = tmpValue_ROT * (Aixinformation.Element(10) - Aininformation.Element(10)) + Aininformation.Element(10)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    Else
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        tmpCount = tmpCount + 1
                    End If
                Next bitIter
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
                For bitIter = 0 To readFusedCoeffDspWave2.SampleSize - 1
                    If bitIter = 14 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(0) = tmpValue_ROV * (Aixinformation.Element(11) - Aininformation.Element(11)) + Aininformation.Element(11)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    ElseIf bitIter = 28 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(1) = tmpValue_ROV * (Aixinformation.Element(12) - Aininformation.Element(12)) + Aininformation.Element(12)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    ElseIf bitIter = 42 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(2) = tmpValue_ROV * (Aixinformation.Element(13) - Aininformation.Element(13)) + Aininformation.Element(13)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    Else
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        tmpCount = tmpCount + 1
                    End If
                Next bitIter
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
            Next site
        End If
        
        If Temperature = 25 Then
            For Each site In TheExec.sites
                temp_rowVal_f(site) = 0
                ElementCNT = SetInformation.Element(0)
                ScanOffset = (UBound(VoltageName) + 1) - ElementCNT
                For col = 0 To ElementCNT - 1
                    temp_rowVal_f(site) = 0
                    For row = 0 To 3
                        PiUCntTempA = (row * ElementCNT) + col
                        temp_rowVal_f(site) = temp_rowVal_f(site) + PiUInformation.Element(PiUCntTempA) * Fused_ROT_Decimal_Vector(site).Element(row)
                    Next row
                    Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) = temp_rowVal_f(site)
    '                    TestName = "Compression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=actualROTMatrix(site).Element(col + ScanOffset), Tname:=TestName, ForceResults:=tlForceNone
    '                    TestName = "Decompression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceNone
                    If (actualROTMatrix(site).Element(col + CLng(ScanOffset)) = 0) Then
                        Difference_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) = 100 * (Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) - 0.0001) / 0.0001
                    Else
                        Difference_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) = 100 * (Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) - actualROTMatrix(site).Element(col + CLng(ScanOffset))) / actualROTMatrix(site).Element(col + CLng(ScanOffset))
                    End If
                Next col
                temp_rowVal_f(site) = 0
                ElementCNT = SetInformation.Element(0)
                InheritanceCnt = SetInformation.Element(4)
                For col = 0 To ElementCNT - 1
                    temp_rowVal_f(site) = 0
                    For row = 0 To 2
                        PiUCntTempB = (row * ElementCNT) + col
                        temp_rowVal_f(site) = temp_rowVal_f(site) + PiUInformation.Element(InheritanceCnt + PiUCntTempB) * Fused_ROV_Decimal_Vector(site).Element(row)
                    Next row
                    Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) = temp_rowVal_f(site)
    '                    TestName = "Compression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=actualROVMatrix(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceNone
    '                    TestName = "Decompression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceNone
                    If (actualROVMatrix(site).Element(col + CLng(ScanOffset)) = 0) Then
                        Difference_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) = 100 * (Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) - 0.0001) / 0.0001
                    Else
                        Difference_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) = 100 * (Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) - actualROVMatrix(site).Element(col + CLng(ScanOffset))) / actualROVMatrix(site).Element(col + CLng(ScanOffset))
                    End If
                Next col
            Next site
    
    
            For col = 0 To UBound(VoltageName)
                    testName = "Compression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col))
                    TheExec.Flow.TestLimit resultVal:=actualROTMatrix.Element(col), Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                    testName = "Decompression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col))
                    TheExec.Flow.TestLimit resultVal:=Output_ROT_Freq_Vector.Element(col), Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                    testName = "Compression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col))
                    TheExec.Flow.TestLimit resultVal:=actualROVMatrix.Element(col), Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                    testName = "Decompression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col))
                    TheExec.Flow.TestLimit resultVal:=Output_ROV_Freq_Vector.Element(col), Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
            Next col
    
            For col = 0 To UBound(VoltageName)
                testName = "PercentageDiff_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col))
                TheExec.Flow.TestLimit resultVal:=Difference_ROT_Freq_Vector.Element(col), hiVal:=1, lowVal:=-1, Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
            Next col
            
            For col = 0 To UBound(VoltageName)
                testName = "PercentageDiff_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col))
                TheExec.Flow.TestLimit resultVal:=Difference_ROV_Freq_Vector.Element(col), hiVal:=1, lowVal:=-1, Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
            Next col
                
    '            For Each site In TheExec.sites
    '                For col = 0 To ElementCNT - 1
    '                    TestName = "PercentageDiff_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Difference_ROT_Freq_Vector.Element(col + CLng(ScanOffset)), hival:=1, lowval:=-1, Tname:=TestName, ForceResults:=tlForceNone
    '                Next col
                
    '                For col = 0 To ElementCNT - 1
    '                    TestName = "PercentageDiff_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Difference_ROV_Freq_Vector.Element(col + CLng(ScanOffset)), hival:=1, lowval:=-1, Tname:=TestName, ForceResults:=tlForceNone
    '                Next col
    '            Next site
            
        ElseIf Temperature = 85 Then
            For Each site In TheExec.sites
                temp_rowVal_f(site) = 0
                ElementCNT = SetInformation.Element(0)
                ScanOffset = (UBound(VoltageName) + 1) - ElementCNT
                InheritanceCnt = SetInformation.Element(4) + SetInformation.Element(5)
                For col = 0 To ElementCNT - 1
                    temp_rowVal_f(site) = 0
                    For row = 0 To 3
                        PiUCntTempA = (row * ElementCNT) + col
                        temp_rowVal_f(site) = temp_rowVal_f(site) + PiUInformation.Element(InheritanceCnt + PiUCntTempA) * Fused_ROT_Decimal_Vector(site).Element(row)
                    Next row
                    Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) = temp_rowVal_f(site)
    '                    TestName = "Compression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=actualROTMatrix(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceNone
    '                    TestName = "Decompression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceNone
                    If (actualROTMatrix(site).Element(col + CLng(ScanOffset)) = 0) Then
                        Difference_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) = 100 * (Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) - 0.0001) / 0.0001
                    Else
                        Difference_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) = 100 * (Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)) - actualROTMatrix(site).Element(col + CLng(ScanOffset))) / actualROTMatrix(site).Element(col + CLng(ScanOffset))
                    End If
                Next col
                temp_rowVal_f(site) = 0
                ElementCNT = SetInformation.Element(0)
                InheritanceCnt = SetInformation.Element(4) + SetInformation.Element(5) + SetInformation.Element(6)
                For col = 0 To ElementCNT - 1
                    temp_rowVal_f(site) = 0
                    For row = 0 To 2
                        PiUCntTempB = (row * ElementCNT) + col
                        temp_rowVal_f(site) = temp_rowVal_f(site) + PiUInformation.Element(InheritanceCnt + PiUCntTempB) * Fused_ROV_Decimal_Vector(site).Element(row)
                    Next row
                    Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) = temp_rowVal_f(site)
    '                    TestName = "Compression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=actualROVMatrix(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceNone
    '                    TestName = "Decompression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col))
    '                    TheExec.Flow.TestLimit resultval:=Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceNone
                    If (actualROVMatrix(site).Element(col + CLng(ScanOffset)) = 0) Then
                        Difference_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) = 100 * (Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) - 0.0001) / 0.0001
                    Else
                        Difference_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) = 100 * (Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)) - actualROVMatrix(site).Element(col + CLng(ScanOffset))) / actualROVMatrix(site).Element(col + CLng(ScanOffset))
                    End If
                Next col
            Next site
            
            For col = 0 To UBound(VoltageName)
                    testName = "Compression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col))
                    TheExec.Flow.TestLimit resultVal:=actualROTMatrix.Element(col), Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                    testName = "Decompression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col))
                    TheExec.Flow.TestLimit resultVal:=Output_ROT_Freq_Vector.Element(col), Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                    testName = "Compression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col))
                    TheExec.Flow.TestLimit resultVal:=actualROVMatrix.Element(col), Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                    testName = "Decompression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col))
                    TheExec.Flow.TestLimit resultVal:=Output_ROV_Freq_Vector.Element(col), Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
            Next col
            
             For col = 0 To UBound(VoltageName)
                testName = "PercentageDiff_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col))
                TheExec.Flow.TestLimit resultVal:=Difference_ROT_Freq_Vector.Element(col), hiVal:=1, lowVal:=-1, Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
            Next col
            
            For col = 0 To UBound(VoltageName)
                testName = "PercentageDiff_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col))
                TheExec.Flow.TestLimit resultVal:=Difference_ROV_Freq_Vector.Element(col), hiVal:=1, lowVal:=-1, Tname:=testName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
            Next col
            
    
    '            For Each site In TheExec.sites
    '                For col = 0 To ElementCNT - 1
    '                    TestName = "PercentageDiff_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Difference_ROT_Freq_Vector.Element(col + CLng(ScanOffset)), hival:=1, lowval:=-1, Tname:=TestName, ForceResults:=tlForceNone
    '                Next col
    '                For col = 0 To ElementCNT - 1
    '                    TestName = "PercentageDiff_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Difference_ROV_Freq_Vector.Element(col + CLng(ScanOffset)), hival:=1, lowval:=-1, Tname:=TestName, ForceResults:=tlForceNone
    '                Next col
    '            Next site
        End If
End Function


Public Function TrimCodeFreq_New_ALG(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double = 1000000, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PreMeas As String, _
    Optional TrimCodeRepeat As Integer, _
    Optional Interpose_PrePat As String, Optional Validating_ As Boolean) As Long
    
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    Dim PatCount As Long, PattArray() As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    
    Call GetFlowTName
    Call HardIP_InitialSetupForPatgen
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Patterns(patset).Load
    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If TrimCodeRepeat <= 0 Then TrimCodeRepeat = 1
    
    ''
    Dim i As Long, j As Long, k As Long, p As Long
    Dim b_MatchTagetFreq As New SiteBoolean, b_DisplayFreq As New SiteBoolean
    Dim StoredTargetTrimCode As New DSPWave, StoreEachTrimCode() As New DSPWave
    Dim StoreEachTrimFreq() As New PinListData
    ReDim StoreEachTrimFreq(DigSrc_Sample_Size + 1) As New PinListData
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    b_MatchTagetFreq = False
    b_DisplayFreq = False
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i
    Dim StoreEachIndex As Long
    Dim PastDSPWave As New DSPWave, InDSPwave As New DSPWave
    Dim ExecutionMax As Long, SrcStartBit As Long, SrcEndBit As Long
    Dim d_MeasF_Interval  As Double
    d_MeasF_Interval = 0.001 ''20190903 0.001
    'd_MeasF_Interval = 0.01 ''20190903 0.001
    Dim b_HighThanTargetFreq As New SiteBoolean
    b_HighThanTargetFreq = False
    Dim TestLimitIndex As Long, LastSectionF1F2_Index As Long
    LastSectionF1F2_Index = 0
    Dim b_DefineFinalFreq As New SiteBoolean
    Dim FinalFreq As New PinListData
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    Dim Interpose_PreMeas_Ary() As String, TestSequenceArray() As String
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    TestSequenceArray = Split(TestSequence, ",")
    Dim MeasureFreq_F1 As New PinListData, MeasureFreq_F2 As New PinListData
    Dim Ts As Variant
    Dim TestNameInput As String
    Dim b_KeepGoing As New SiteBoolean
    Dim PreviousFreq As New PinListData
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    Dim CloseTargetFreq As New PinListData
    Dim DiffValue As New SiteLong, PreviousDiffValue As New SiteLong, CloseIndex As New SiteLong
    
    Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteLong
    Dim BinStr As String
    Dim CloseTargetTrimCode As New DSPWave
    Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
    Dim b_FirstTimeSwitch As Boolean
    ''
    
    Dim site As Variant
    Dim MeasureFreq As New PinListData
    Dim OutputTrimCode As String
    Dim InitialDSPWave As New DSPWave, TrimStartDSPWave As New DSPWave, InDSPwave_Repeat As New DSPWave
    Dim StrSeparatebyComma() As String, StrSeparatebyEqual() As String, StrSeparatebyColon() As String
    
    Dim StartBit_CoarseTuning As Long, EndBit_CoarseTuning As Long
    Dim TrimStart_Dec As Long
    Dim TempCnt As Long
    Dim MaskDSPWave0 As New DSPWave, MaskDSPWave1 As New DSPWave
    Dim bTrimCodeFreqContinue As New SiteBoolean, bCoarseTuningContinue As New SiteBoolean, bCheckAverageContinue As New SiteBoolean
    Dim InitialDSPWave_F0 As New DSPWave, InitialDSPWave_F1 As New DSPWave
    Dim StoreEachTrimFreq0() As New PinListData, StoreEachTrimFreq1() As New PinListData
    Dim StoreTrimFreq0 As New PinListData, StoreTrimFreq1 As New PinListData
    ReDim StoreEachTrimFreq0(DigSrc_Sample_Size)
    ReDim StoreEachTrimFreq1(DigSrc_Sample_Size)
    Dim bF0HigherThanTargetFreq As New SiteBoolean, bF1HigherThanTargetFreq As New SiteBoolean
    Dim MeasureFreq0 As New PinListData, MeasureFreq1 As New PinListData
    Dim FinalDSPCoarseTuning As New DSPWave
    FinalDSPCoarseTuning.CreateConstant 0, DigSrc_Sample_Size, DspLong
    Dim FreqAVG01 As New PinListData, FreqAVG23 As New PinListData
    Dim bF2HigherThanTargetFreq As New SiteBoolean, bF3HigherThanTargetFreq As New SiteBoolean
    Dim bFAVG01HigherThanTargetFreq As New SiteBoolean, bFAVG23HigherThanTargetFreq As New SiteBoolean
    Dim bFAVG01EqualToTargetFreq As New SiteBoolean
    Dim bTempSiteSelected As New SiteBoolean

    StrSeparatebyComma = Split(TrimFormat, ";")
    StrSeparatebyEqual = Split(StrSeparatebyComma(0), "=")
    StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
    StartBit_CoarseTuning = StrSeparatebyColon(0)
    EndBit_CoarseTuning = StrSeparatebyColon(1)
    
    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", "")
    End If
    TrimStart_Dec = Bin2Dec(TrimStart)
    TrimStartDSPWave.CreateConstant TrimStart_Dec, 1, DspLong
    Call rundsp.CreateFlexibleDSPWave(TrimStartDSPWave, DigSrc_Sample_Size, InitialDSPWave)
    For Each site In TheExec.sites
        InitialDSPWave(site) = InitialDSPWave(site).ConvertDataTypeTo(DspLong)
    Next site
    TheExec.Datalog.WriteComment ("Initial Trim Code")
    Call TrimCodeFreq_WriteComment_DspTrimCode(InitialDSPWave)
        
    'Coarse Tuning
    TheExec.Datalog.WriteComment ("********** Coarse Tuning")
    bCoarseTuningContinue = True
    bTrimCodeFreqContinue = True
    
    MaskDSPWave0.CreateConstant 0, DigSrc_Sample_Size, DspLong
    MaskDSPWave1.CreateConstant 1, DigSrc_Sample_Size, DspLong
    For TempCnt = StartBit_CoarseTuning To EndBit_CoarseTuning Step -1
        MaskDSPWave0.Element(TempCnt) = 1 '1111 0000 0000
        MaskDSPWave1.Element(TempCnt) = 0 '0000 1111 1111
    Next TempCnt
    
    Dim TempCnt_TrimStep As Long
    For TempCnt_TrimStep = StartBit_CoarseTuning To EndBit_CoarseTuning Step -1
        TheExec.Datalog.WriteComment ("======================================================================================")
        TheExec.Datalog.WriteComment ("BIT" & TempCnt_TrimStep & "=1")
        For Each site In TheExec.sites
            InitialDSPWave_F0(site) = InitialDSPWave(site).bitwiseand(MaskDSPWave0(site))
            InDSPwave_Repeat(site) = InitialDSPWave_F0(site).repeat(TrimCodeRepeat)
        Next site
        Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)
        Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        StoreEachTrimFreq0(TempCnt_TrimStep) = MeasureFreq
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency0 = " & FormatNumber((StoreEachTrimFreq0(TempCnt_TrimStep).Pins(0).Value(site) / 1000000), 6) & "M Hz")
        Next site
        
        For Each site In TheExec.sites
            InitialDSPWave_F1(site) = InitialDSPWave(site).BitwiseOr(MaskDSPWave1(site))
            InDSPwave_Repeat(site) = InitialDSPWave_F1(site).repeat(TrimCodeRepeat)
        Next site
        Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)
        Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        StoreEachTrimFreq1(TempCnt_TrimStep) = MeasureFreq
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency1 = " & FormatNumber((StoreEachTrimFreq1(TempCnt_TrimStep).Pins(0).Value(site) / 1000000), 6) & "M Hz")
        Next site
        
        bF0HigherThanTargetFreq = StoreEachTrimFreq0(TempCnt_TrimStep).Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        bF1HigherThanTargetFreq = StoreEachTrimFreq1(TempCnt_TrimStep).Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        
        For Each site In TheExec.sites
            If (bF0HigherThanTargetFreq(site) = False) And (bF1HigherThanTargetFreq(site) = False) Then
                If TempCnt_TrimStep = EndBit_CoarseTuning Then
                    TheExec.Datalog.WriteComment ("Site " & site & ": EXIT FUNCTION !!!")
                    ''20190903
                    StoredTargetTrimCode(site) = InitialDSPWave(site)
                    bCoarseTuningContinue(site) = False
                    bTrimCodeFreqContinue(site) = False
                End If
                InitialDSPWave(site).Element(TempCnt_TrimStep - 1) = 1
            ElseIf bF0HigherThanTargetFreq(site) = True Then
                If TempCnt_TrimStep = EndBit_CoarseTuning Then
                    TheExec.Datalog.WriteComment ("Site " & site & ": EXIT FUNCTION !!!")
                    ''20190903
                    StoredTargetTrimCode(site) = InitialDSPWave(site)
                    bCoarseTuningContinue(site) = False
                    bTrimCodeFreqContinue(site) = False
                End If
                InitialDSPWave(site).Element(TempCnt_TrimStep) = 0
                InitialDSPWave(site).Element(TempCnt_TrimStep - 1) = 1
            ElseIf (bF0HigherThanTargetFreq(site) = False) And (bF1HigherThanTargetFreq(site) = True) Then
                bCoarseTuningContinue(site) = False
                StoreTrimFreq0 = StoreEachTrimFreq0(TempCnt_TrimStep)
                StoreTrimFreq1 = StoreEachTrimFreq1(TempCnt_TrimStep)
                TheExec.Datalog.WriteComment ("Site " & site & ": FO<Target AND F1>Target, Go Check Average")
            End If
        Next site
        If bCoarseTuningContinue.All(False) = True Then
            Exit For
        End If
        TheExec.sites.Selected = bCoarseTuningContinue
    Next TempCnt_TrimStep

    'Check Average
    TheExec.Datalog.WriteComment ("********** Check Average")
    ''TEST
    ''bTrimCodeFreqContinue(5) = False
    'TheExec.sites.Selected = True
    TheExec.sites.Selected = bTrimCodeFreqContinue
    bCheckAverageContinue = bTrimCodeFreqContinue
    FinalDSPCoarseTuning = InitialDSPWave
    
    bTempSiteSelected = TheExec.sites.Selected
    If bTempSiteSelected.Any(True) = True Then
        TheExec.Datalog.WriteComment ("Recode F_Coarse")
        
        FreqAVG01 = StoreTrimFreq0.Math.Add(StoreTrimFreq1).Divide(2)
        
        bFAVG01HigherThanTargetFreq = FreqAVG01.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        bFAVG01EqualToTargetFreq = FreqAVG01.Math.Subtract(TrimTarget).compare(EqualTo, 0) ''20190903
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " AVG01(F0,F1) = " & FormatNumber((FreqAVG01.Pins(0).Value(site) / 1000000), 6) & "M Hz")
        Next site
        Dim TempDSP As New DSPWave
        TempDSP.CreateConstant 0, 1, DspLong
        For Each site In TheExec.sites
            TempDSP(site) = InitialDSPWave(site).ConvertStreamTo(tldspParallel, DigSrc_Sample_Size, 0, Bit0IsMsb)
            ''
            If TempDSP(site).Element(0) < 2 ^ (EndBit_CoarseTuning) Or TempDSP(site).Element(0) > (2 ^ (StartBit_CoarseTuning + 1) - 1 - 2 ^ (EndBit_CoarseTuning)) Then ''256 to 2 ^ (EndBit_CoarseTuning + 1), 3839
                TheExec.Datalog.WriteComment ("Site " & site & " : ERROR")
            End If
            ''
            If bFAVG01HigherThanTargetFreq = True Then
                TempDSP(site) = TempDSP(site).Add(-1 * 2 ^ (EndBit_CoarseTuning)).ConvertDataTypeTo(DspLong) 'watch out 0000XXXXXXXX
                InitialDSPWave(site) = TempDSP(site).ConvertStreamTo(tldspSerial, DigSrc_Sample_Size, 0, Bit0IsMsb)
                InitialDSPWave(site) = InitialDSPWave(site).ConvertDataTypeTo(DspLong)
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01>Target, CurrentTrimCode-1")
            ElseIf bFAVG01HigherThanTargetFreq = False And bFAVG01EqualToTargetFreq = False Then ''20190903
                TempDSP(site) = TempDSP(site).Add(2 ^ (EndBit_CoarseTuning)).ConvertDataTypeTo(DspLong) 'watch out 1111XXXXXXXX
                InitialDSPWave(site) = TempDSP(site).ConvertStreamTo(tldspSerial, DigSrc_Sample_Size, 0, Bit0IsMsb)
                InitialDSPWave(site) = InitialDSPWave(site).ConvertDataTypeTo(DspLong)
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01<Target, CurrentTrimCode+1")
            ElseIf bFAVG01EqualToTargetFreq = True Then ''20190903
                bCheckAverageContinue(site) = False
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01=Target, Go Fine Tuning")
            End If
        Next site
    End If
    
    'TheExec.sites.Selected = True
    TheExec.sites.Selected = bTrimCodeFreqContinue.LogicalAnd(bCheckAverageContinue)
    TheExec.Datalog.WriteComment ("********** Check Average : F2 F3")
    bTempSiteSelected = TheExec.sites.Selected
    If bTempSiteSelected.Any(True) = True Then
        For Each site In TheExec.sites
            InitialDSPWave_F0(site) = InitialDSPWave(site).bitwiseand(MaskDSPWave0(site))
            InDSPwave_Repeat(site) = InitialDSPWave_F0(site).repeat(TrimCodeRepeat)
        Next site
        Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)
        Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        StoreEachTrimFreq0(TempCnt_TrimStep) = MeasureFreq
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency2 = " & FormatNumber((StoreEachTrimFreq0(TempCnt_TrimStep).Pins(0).Value(site) / 1000000), 6) & "M Hz")
        Next site
        
        For Each site In TheExec.sites
            InitialDSPWave_F1(site) = InitialDSPWave(site).BitwiseOr(MaskDSPWave1(site))
            InDSPwave_Repeat(site) = InitialDSPWave_F1(site).repeat(TrimCodeRepeat)
        Next site
        Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)
        Call TrimCodeFreq_RunPat_and_MeasF(PattArray(0), TestSequence, CPUA_Flag_In_Pat, DigSrc_pin, DigSrc_Sample_Size * TrimCodeRepeat, MeasureF_Pin, MeasureFreq, InDSPwave_Repeat, Interpose_PreMeas)
        StoreEachTrimFreq1(TempCnt_TrimStep) = MeasureFreq
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency3 = " & FormatNumber((StoreEachTrimFreq1(TempCnt_TrimStep).Pins(0).Value(site) / 1000000), 6) & "M Hz")
        Next site
        
        bF0HigherThanTargetFreq = StoreEachTrimFreq0(TempCnt_TrimStep).Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        bF1HigherThanTargetFreq = StoreEachTrimFreq1(TempCnt_TrimStep).Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        
        For Each site In TheExec.sites
            If bF0HigherThanTargetFreq(site) = True Or bF1HigherThanTargetFreq(site) = False Then
                bCheckAverageContinue(site) = False
                TheExec.Datalog.WriteComment ("Site " & site & " : F2>Target OR F3<Target, Go Fine Tuning")
            End If
        Next site
    End If
    
    TheExec.Datalog.WriteComment ("********** Check Average : FreqAVG01 vs FreqAVG23")
    'TheExec.sites.Selected = True
    TheExec.sites.Selected = bTrimCodeFreqContinue.LogicalAnd(bCheckAverageContinue)
    bTempSiteSelected = TheExec.sites.Selected
    If bTempSiteSelected.Any(True) = True Then
        FreqAVG23 = StoreEachTrimFreq0(TempCnt_TrimStep).Math.Add(StoreEachTrimFreq1(TempCnt_TrimStep)).Divide(2)
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment ("Site " & site & " AVG23(F2,F3) = " & FormatNumber((FreqAVG23.Pins(0).Value(site) / 1000000), 6) & "M Hz")
        Next site
        FreqAVG01 = FreqAVG01.Math.Subtract(TrimTarget).Abs
        FreqAVG23 = FreqAVG23.Math.Subtract(TrimTarget).Abs
        For Each site In TheExec.sites
            If FreqAVG01.Pins(0).Value(site) > FreqAVG23.Pins(0).Value(site) Then ''20190903upadte, from "<" to ">"
                FinalDSPCoarseTuning(site) = InitialDSPWave(site)
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01>FreqAVG23,         Update F_Coarse, Go Fine Tuning")
            Else
                TheExec.Datalog.WriteComment ("Site " & site & " : FreqAVG01<=FreqAVG23, Do not Update F_Coarse, Go Fine Tuning")
            End If
        Next site
    End If
    
    'Go Fine Tuning
    
    'Fine Tuning
    TheExec.Datalog.WriteComment ("********** Fine Tuning")
    TheExec.sites.Selected = True
    If bTrimCodeFreqContinue.All(False) = True Then
        TheExec.Datalog.WriteComment ("********** All Site Fail !!!, Jump to Final")
        GoTo TrimCodeTTR
    Else
        TheExec.sites.Selected = bTrimCodeFreqContinue
    End If


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    TrimFormat = StrSeparatebyComma(1)
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    For Each site In TheExec.sites
        InDSPwave(site) = FinalDSPCoarseTuning(site)
        InDSPwave_Repeat(site) = InDSPwave(site).repeat(TrimCodeRepeat)
    Next site
    Call TrimCodeFreq_WriteComment_DspTrimCode(InDSPwave_Repeat)
    ''
    
    'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
    
    '/* ------ Added by Kaino on 2019/06/12 ------- */
    For Each site In TheExec.sites
        InDSPwave_Repeat = InDSPwave.repeat(TrimCodeRepeat)
    Next site
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
    '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */
  

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("First Time Setup")
    '' Debug use
    For Each site In TheExec.sites
        OutputTrimCode = ""
   
        If gl_Disable_HIP_debug_log = False Then
            'For k = 0 To InDSPwave(Site).SampleSize - 1
            '    OutputTrimCode = OutputTrimCode & CStr(InDSPwave(Site).Element(k))
            'Next k
        
            '/* ------ Added by Kaino on 2019/06/12 ------- */
            For k = 0 To InDSPwave_Repeat(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPwave_Repeat(site).Element(k))
               If (k + 1) Mod DigSrc_Sample_Size = 0 And k < InDSPwave_Repeat(site).SampleSize - 1 Then
                   OutputTrimCode = OutputTrimCode & ","
               End If
            Next k
        '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */
        
            TheExec.Datalog.WriteComment ("Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
                End If
    Next site
    
    For Each site In TheExec.sites
        StoreEachTrimCode(0)(site).Data = InDSPwave(site).Data
    Next site
    
    Call TheHdw.Patterns(PattArray(0)).start
    
    ''Update Interpose_PreMeas 20170801
    Dim TestSeqNum As Integer
    TestSeqNum = 0

    For Each Ts In TestSequenceArray
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20160923 - Add Interpose_PreMeas entry point by each sequence
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            End If
        End If
        
        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
            Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
            
            If TheExec.TesterMode = testModeOffline Then
                Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
            End If
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20161206-Restore force condiction after measurement
        ''Call SetForceCondition("RESTORE")
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            End If
        End If

        TestSeqNum = TestSeqNum + 1

        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts
    TheHdw.Digital.Patgen.HaltWait
    
    StoreEachTrimFreq(0) = MeasureFreq
    
    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
    PastDSPWave = InDSPwave
    
    TestNameInput = "Freq_meas_"
    TestLimitIndex = 0
    
    '' 20160712 - Modify to use WriteComment to display output frequency.
    If gl_Disable_HIP_debug_log = False Then
        For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
        Next site
    End If
    '' 20160712 - Compare Measure Frequency whether match target Freq
    b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
    
    b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
    For Each site In TheExec.sites
        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
            StoredTargetTrimCode(site).Data = InDSPwave(site).Data
            b_StopTrimCodeProcess(site) = True
        End If
    Next site
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
    
    
    ''========================================================================================
    ''20161128 Pre check Min/Max trim code process
'    Dim b_KeepGoing As New SiteBoolean
'    Dim PreviousFreq As New PinListData
    If PreCheckMinMaxTrimCode = True Then
        PreviousFreq = MeasureFreq
        Call rundsp.PreCheckMinMaxTrimCode(b_HighThanTargetFreq, InDSPwave)
        
        
        'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
        '/* ------ Added by Kaino on 2019/06/12 ------- */
        For Each site In TheExec.sites
            InDSPwave_Repeat = InDSPwave.repeat(TrimCodeRepeat)
        Next site
        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
        '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */
        
        ''Update Interpose_PreMeas 20170801
        Call TheHdw.Patterns(PattArray(0)).start
        
        TestSeqNum = 0
        
        For Each Ts In TestSequenceArray
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
            Else
                Call TheHdw.Digital.Patgen.HaltWait
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                End If
            End If
            
            If UCase(Ts) = "F" Then
                Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
                
                If TheExec.TesterMode = testModeOffline Then
                    Call SimulatePreCheckOutputFreq(MeasureF_Pin, MeasureFreq)
                End If
            End If
            
            ''Update Interpose_PreMeas 20170801
            ''20161206-Restore force condiction after measurement
            ''Call SetForceCondition("RESTORE")
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                End If
            End If
    
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA)
            Else
                TheHdw.Digital.Patgen.HaltWait
            End If
        Next Ts
        
        TheHdw.Digital.Patgen.HaltWait
        
        For Each site In TheExec.sites
            OutputTrimCode = ""
            For k = 0 To InDSPwave(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
            Next k
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
        Next site
        
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
            Next site
        End If
        For Each site In TheExec.sites
            If b_HighThanTargetFreq(site) = True Then
                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(LessThan, 0)
            Else
                b_KeepGoing(site) = MeasureFreq.Math.Subtract(PreviousFreq).compare(GreaterThan, 0)
            End If
        Next site

        Dim PreCheckBinStr As String, PreCheckDecVal As Double
        For Each site In TheExec.sites
            If b_KeepGoing(site) = False Then
                b_StopTrimCodeProcess(site) = True
                PreCheckBinStr = ""
                StoredTargetTrimCode(site).Data = InDSPwave(site).Data
                For i = 0 To StoredTargetTrimCode(site).SampleSize - 1
                    PreCheckBinStr = PreCheckBinStr & StoredTargetTrimCode.Element(i)
                Next i
                PreCheckDecVal = Bin2Dec_rev_Double(PreCheckBinStr)
                ''TheExec.Flow.TestLimit PreCheckDecVal, 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TheExec.DataManager.InstanceName & "_TrimCode_Decimal", ForceResults:=tlForceNone
            End If
        Next site
    End If
    
    ''========================================================================================
    

    

'    Dim b_ControlNextBit As Boolean
'    b_ControlNextBit = False
'    Dim b_FirstExecution As Boolean
'    b_FirstExecution = False
    StoreEachIndex = 1
    
    ''20170103-Setup b_KeepGoing to true if PreCheckMinMaxTrimCode=false
    If PreCheckMinMaxTrimCode = False Then
        b_KeepGoing = True
    End If
    
    If b_KeepGoing.All(False) Then
    Else

        For i = 0 To ExecutionMax
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If
            StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
            StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
            SrcStartBit = StrSeparatebyColon(0)
            SrcEndBit = StrSeparatebyColon(1)
            
            If i = 0 Then
                b_FirstExecution = True
            Else
                b_FirstExecution = False
                SrcStartBit = SrcStartBit + 1
            End If
            
            For j = SrcStartBit To SrcEndBit Step -1
            
                If b_FirstExecution = True Then
                    b_ControlNextBit = True
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                Else
                ''20160716-Control next bit to 1 no matter first or last progress
                    b_ControlNextBit = True
    ''                b_ControlNextBit = False
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                End If
    
                If b_FirstExecution = True And j = SrcEndBit Then
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPwave)
    ''            ElseIf b_FirstExecution = False And j = SrcStartBit Then
    ''                j = SrcStartBit + 1
                Else
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                End If
                
                
                'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, InDSPwave)
                '/* ------ Added by Kaino on 2019/06/12 ------- */
                For Each site In TheExec.sites
                    InDSPwave_Repeat = InDSPwave.repeat(TrimCodeRepeat)
                Next site
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
                '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */
                
                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site).Data = InDSPwave(site).Data
                Next site
            
                '' Debug use
                '' ==============================================================================================
                '' 20160716 - Modify trim code rule
                
                If gl_Disable_HIP_debug_log = False Then
                    If b_FirstExecution = True Then
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    Else
                        If j = SrcEndBit Then
                            TheExec.Datalog.WriteComment ("Setup Bit " & j)
                        Else
                            TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                        End If
                    End If
                End If
                
                For Each site In TheExec.sites
    ''              If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                    If b_KeepGoing(site) = True Then
                                                If gl_Disable_HIP_debug_log = False Then
                                OutputTrimCode = ""
                                'For k = 0 To InDSPwave(Site).SampleSize - 1
                                '    OutputTrimCode = OutputTrimCode & CStr(InDSPwave(Site).Element(k))
                                'Next k
                                '/* ------ Added by Kaino on 2019/06/12 ------- */
                                For k = 0 To InDSPwave_Repeat(site).SampleSize - 1
                                    OutputTrimCode = OutputTrimCode & CStr(InDSPwave_Repeat(site).Element(k))
                                    If (k + 1) Mod DigSrc_Sample_Size = 0 And k < InDSPwave_Repeat(site).SampleSize - 1 Then
                                        OutputTrimCode = OutputTrimCode & ","
                                    End If
                                Next k
                                '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */
                                TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
                                                End If
                    End If
    ''              End If
                Next site
                '' ==============================================================================================
                
                Call TheHdw.Patterns(PattArray(0)).start
                
                ''Update Interpose_PreMeas 20170801
                TestSeqNum = 0
                
                For Each Ts In TestSequenceArray
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
                    Else
                        Call TheHdw.Digital.Patgen.HaltWait
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20160923 - Add Interpose_PreMeas entry point by each sequence
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                        End If
                    End If
                    
                    
                    If UCase(Ts) = "F" Then
                        Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
                        Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
                        
                        '--------------- off line mode data --------
                        If TheExec.TesterMode = testModeOffline Then
                            Dim SimuIndex As Long
                            SimuIndex = TestLimitIndex
                            If SimuIndex >= 8 Then
                                SimuIndex = 8
                            End If
                            Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
                            MeasureFreq.Pins(MeasureF_Pin).Value(0) = MeasureFreq.Pins(MeasureF_Pin).Value(0) - (SimuIndex * 1000)
                           ' MeasureFreq.Pins(MeasureF_Pin).Value(1) = MeasureFreq.Pins(MeasureF_Pin).Value(1) + (SimuIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(2) = MeasureFreq.Pins(MeasureF_Pin).Value(2) + (TestLimitIndex * 1000)
    ''                        MeasureFreq.Pins(MeasureF_Pin).Value(3) = MeasureFreq.Pins(MeasureF_Pin).Value(3) - (TestLimitIndex * 1000)
                        End If
                        '--------------------------------------------
                        
                        If j = SrcEndBit + 1 Then
                            MeasureFreq_F1 = MeasureFreq
                        ElseIf j = SrcEndBit Then
                            MeasureFreq_F2 = MeasureFreq
                        End If
                    Else
                        '' Do nothing
                    End If
                    
                    ''Update Interpose_PreMeas 20170801
                    ''20161206-Restore force condiction after measurement
                    ''Call SetForceCondition("RESTORE")
                    If Interpose_PreMeas <> "" Then
                        If UBound(Interpose_PreMeas_Ary) = 0 Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        End If
                    End If
            
                    TestSeqNum = TestSeqNum + 1
                    
                    If (CPUA_Flag_In_Pat) Then
                        Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                    Else
                        TheHdw.Digital.Patgen.HaltWait
                    End If
                Next Ts
                
                TheHdw.Digital.Patgen.HaltWait
                
                StoreEachTrimFreq(StoreEachIndex) = MeasureFreq
                StoreEachIndex = StoreEachIndex + 1
                
                If j = SrcEndBit Then
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq_F1.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureFreq_F2.Math.Subtract(TrimTarget).Abs)
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetFreq, j, b_ControlNextBit, InDSPwave)
                    PastDSPWave = InDSPwave
                Else
                    b_HighThanTargetFreq = False
                    b_HighThanTargetFreq = MeasureFreq.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                    PastDSPWave = InDSPwave
                End If
    
                TestLimitIndex = TestLimitIndex + 1
                
                '' 20160712 - Modify to use WriteComment to display output frequency.
                
                If gl_Disable_HIP_debug_log = False Then
                
                    For Each site In TheExec.sites
        ''                If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        If b_KeepGoing(site) = True Then
                            TheExec.Datalog.WriteComment ("Site " & site & " Output Frequency = " & FormatNumber((MeasureFreq.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                        End If
        ''                End If
                    Next site
                End If
                
                ''20160716 - Modify display info sequence when source bit in the section end
                If j = SrcEndBit Then
                    For Each site In TheExec.sites
    ''                    If b_MatchTagetFreq(Site) = False And b_DisplayFreq(Site) = False Then
                        
                        If b_KeepGoing(site) = True And gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 1 & " Output Frequency = " & FormatNumber((MeasureFreq_F1.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                            TheExec.Datalog.WriteComment ("Site " & site & " F" & LastSectionF1F2_Index + 2 & " Output Frequency = " & FormatNumber((MeasureFreq_F2.Pins(0).Value(site) / 1000000), 6) & "M Hz")
                        End If
    ''                    End If
                    Next site
                    LastSectionF1F2_Index = LastSectionF1F2_Index + 2
                End If
                
                '' 20160712 - Compare Measure Frequency whether match target Freq
                b_MatchTagetFreq = MeasureFreq.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                b_DisplayFreq = b_DisplayFreq.LogicalOr(b_MatchTagetFreq)
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
                        If b_MatchTagetFreq(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                            StoredTargetTrimCode(site).Data = InDSPwave(site).Data
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                Next site
                ''20161128-Stop trim code process if found out match code of all site
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
            Next j
        Next i
    End If
    
    ''============================================================================
    ''20161128 Findout mimiumn trim code
'    Dim CloseTargetFreq As New PinListData
'    Dim DiffValue As New SiteLong, PreviousDiffValue As New SiteLong, CloseIndex As New SiteLong
'
'    Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteLong
'    Dim BINstr As String
'    Dim CloseTargetTrimCode As New DSPWave
'    Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
'    Dim b_FirstTimeSwitch As Boolean
'    TheExec.sites.Selected = True
    If b_KeepGoing.All(False) Then
    Else
''        If TrimPrcocessAll = True Then
            CloseTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If StoredTargetTrimCode(site).CalcSum = 0 Then
                        b_UseMinTrim(site) = True
                    End If
                End If
            Next site
            
            If UseMinimumTrimCode = True Then
                b_UseMinTrim = True
            End If
            ''TEST
'            TheExec.sites.Selected = True
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_UseMinTrim(site) = True Then
                        '' Findout minimum difference value
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            DiffValue(site) = Abs(StoreEachTrimFreq(i).Pins(0).Value(site) - TrimTarget)
                            If DiffValue(site) <= PreviousDiffValue(site) Then
                                CloseIndex(site) = i
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                            If i = 0 Then
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                        Next i
                        '' Transfer to decimal value to findout minimum code
                        PreviousDecVal = 0
                        DecVal = 0
                        b_FirstTimeSwitch = False
''                        For i = 0 To UBound(StoreEachTrimFreq)
                        For i = 0 To StoreEachIndex - 1
                            BinStr = ""
                            If Abs(StoreEachTrimFreq(i).Pins(0).Value(site) - TrimTarget) = MinDiffVal(site) Then
                                For j = 0 To StoreEachTrimCode(i)(site).SampleSize - 1
                                    BinStr = BinStr & StoreEachTrimCode(i)(site).Element(j)
                                Next j
                                DecVal = Bin2Dec_rev_Double(BinStr)
                               
                                If DecVal < PreviousDecVal Then
                                    MinDecVal = DecVal
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                End If
                                PreviousDecVal = DecVal
                                If b_FirstTimeSwitch = False Then
                                    CloseTargetTrimCode(site).Data = StoreEachTrimCode(i)(site).Data
                                    b_FirstTimeSwitch = True
                                End If
                            End If
                        Next i
                    End If
                End If
            Next site
''        End If
    End If
    
    For Each site In TheExec.sites
        If b_KeepGoing(site) = True Then
            If b_UseMinTrim(site) = True Then
                StoredTargetTrimCode(site).Data = CloseTargetTrimCode(site).Data
            Else
                StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
            End If
        Else
            StoredTargetTrimCode(site).Data = StoredTargetTrimCode(site).Data
        End If
    Next site
    ''============================================================================
    
    ''
TrimCodeTTR:
    TheExec.sites.Selected = True
'    Dim ZeroData() As Long
'    ReDim ZeroData(DigSrc_Sample_Size - 1)
'
'
'    For Each Site In TheExec.sites
'        If bTrimCodeFreqContinue(Site) = False Then
'            StoredTargetTrimCode(Site).Data = ZeroData
'        End If
'    Next Site
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, StoredTargetTrimCode)
    End If
    
    
    
    Call HardIP_WriteFuncResult(, , Inst_Name_Str)

    For Each site In TheExec.sites
        OutputTrimCode = ""
        For k = 0 To StoredTargetTrimCode(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(StoredTargetTrimCode(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Final Output Trim Code = " & OutputTrimCode)
    Next site
    
    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel StoredTargetTrimCode, DigSrc_Sample_Size, ConvertedDataWf
    
    TestNameInput = Report_TName_From_Instance("C", DigSrc_pin.Value, "TrimCode(Decimal)", 0, 0)
    
    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TestNameInput, ForceResults:=tlForceNone  'for Turks
    'TheExec.Flow.TestLimit ConvertedDataWf.Element(0), 0, 2 ^ DigSrc_Sample_Size - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    
    'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size, StoredTargetTrimCode)
    '/* ------ Added by Kaino on 2019/06/12 ------- */
    For Each site In TheExec.sites
        InDSPwave_Repeat = StoredTargetTrimCode.repeat(TrimCodeRepeat)
                If gl_Disable_HIP_debug_log = False Then
                OutputTrimCode = ""
                For k = 0 To InDSPwave_Repeat(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPwave_Repeat(site).Element(k))
                    If (k + 1) Mod DigSrc_Sample_Size = 0 And k < InDSPwave_Repeat(site).SampleSize - 1 Then
                    OutputTrimCode = OutputTrimCode & ","
                    End If
            Next k
                TheExec.Datalog.WriteComment ("Site_" & site & " Final Test Trim Code = " & OutputTrimCode)
                End If

    Next site
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeFreq", DigSrc_Sample_Size * TrimCodeRepeat, InDSPwave_Repeat)
    '/* ^^^^^^ Added by Kaino on 2019/06/12 ^^^^^^ */

    Call TheHdw.Patterns(PattArray(0)).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
    
    For Each Ts In TestSequenceArray
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20160923 - Add Interpose_PreMeas entry point by each sequence
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            End If
        End If
        
        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
            Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
        End If
        
        ''Update Interpose_PreMeas 20170801
        ''20161206-Restore force condiction after measurement
        ''Call SetForceCondition("RESTORE")
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            End If
        End If

        TestSeqNum = TestSeqNum + 1
        
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts
    
    TheHdw.Digital.Patgen.HaltWait
    
    If TPModeAsCharz_GLB Then
        Disable_Inst_pinname_in_PTR
        TheExec.Flow.TestLimit resultVal:=MeasureFreq, Unit:=unitHz, ForceResults:=tlForceFlow
        Enable_Inst_pinname_in_PTR
    Else
        For p = 0 To MeasureFreq.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("F", MeasureFreq.Pins(p), "Final", CInt(p))
            TheExec.Flow.TestLimit resultVal:=MeasureFreq, Unit:=unitHz, Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next p
    End If
    
    Dim sl_FUSE_Val As New SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else

    End If
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    DebugPrintFunc patset.Value
    
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeFreq_New_ALG function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function













Public Function TrimCodeFreq_RunPat_and_MeasF(Pat As String, TestSequence As String, CPU_Flag_In_Pat As Boolean, _
    DigSrc_pin As PinList, DigSrc_SampleSize As Long, MeasureF_Pin As PinList, MeasureFreq As PinListData, _
    InDSPwave As DSPWave, Interpose_PreMeas As String)
    'TrimCodeFreq_RunPat_and_MeasF(Pat,TestSequence,CPU_Flag_In_Pat,DigSrc_Pin,DigSrc_SampleSize,MeasureF_Pin,MeasureFreq,InDSPwave,Interpose_PreMeas)
    
    Dim Ts As Variant
    Dim TestSeqNum As Long
    Dim TestSequenceArray() As String
    Dim Interpose_PreMeas_Ary() As String
    Dim d_MeasF_Interval  As Double
    
    TestSeqNum = 0
    TestSequenceArray = Split(TestSequence, ",")
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    d_MeasF_Interval = 0.001 ''20190903 0.001
    'd_MeasF_Interval = 0.01 ''20190903 0.001
        
    Call SetupDigSrcDspWave(Pat, DigSrc_pin, "TrimCodeFreq", DigSrc_SampleSize, InDSPwave)
    Call TheHdw.Patterns(Pat).start
    
    For Each Ts In TestSequenceArray
        If (CPU_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
        Else
            Call TheHdw.Digital.Patgen.HaltWait
        End If
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
            End If
        End If
        
        If UCase(Ts) = "F" Then
            Call Freq_MeasFreqSetup(MeasureF_Pin, d_MeasF_Interval)
            Call HardIP_Freq_MeasFreqStart(MeasureF_Pin, d_MeasF_Interval, MeasureFreq, 0.001)
            If TheExec.TesterMode = testModeOffline Then
                Call SimulateOutputFreq(MeasureF_Pin, MeasureFreq)
            End If
        End If
        
        If Interpose_PreMeas <> "" Then
            If UBound(Interpose_PreMeas_Ary) = 0 Then
                Call SetForceCondition("RESTOREPREMEAS")
            ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                Call SetForceCondition("RESTOREPREMEAS")
            End If
        End If
        TestSeqNum = TestSeqNum + 1
        If (CPU_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Else
            TheHdw.Digital.Patgen.HaltWait
        End If
    Next Ts
    TheHdw.Digital.Patgen.HaltWait
    
End Function


Public Function TrimCodeFreq_WriteComment_DspTrimCode(InDsp As DSPWave)
    Dim OutputTrimCode As String
    Dim TempCnt As Long
    Dim site As Variant
    
    For Each site In TheExec.sites
        OutputTrimCode = ""
        For TempCnt = 0 To InDsp(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(InDsp(site).Element(TempCnt))
            If TempCnt = 11 Then ''11
                OutputTrimCode = OutputTrimCode & ","
            End If
        Next TempCnt
        TheExec.Datalog.WriteComment ("Site " & site & " Output Trim Code = " & OutputTrimCode)
    Next site
End Function


