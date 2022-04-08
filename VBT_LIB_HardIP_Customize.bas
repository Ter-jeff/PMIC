Attribute VB_Name = "VBT_LIB_HardIP_Customize"
Option Explicit
Enum VIR_EntryPoint ''VBT_CUSTOMIZE_HardIP
    VIR_MI_AFTER_MEASUREMENT = 0
    VIR_MI_AFTER_TESTLIMIT = 1
End Enum

Public AMP_EYE_VT_CZ_Flag As Boolean

Public Function AMP_EYE_VT_Setup(Char_Flag As Boolean)

    If Char_Flag = True Then
        AMP_EYE_VT_CZ_Flag = True
    Else
        AMP_EYE_VT_CZ_Flag = False
    End If
    
End Function

Public Function CUS_AMP_SDLL_SWP_Init(Loop_count As Long, Loop_Init As Long, Loop_Idx As Long, CUS_Str_MainProgram As String, Ori_CUST_Str_MainProgram As String) As Long
    
    If Loop_count = Loop_Init Then
        Ori_CUST_Str_MainProgram = CUS_Str_MainProgram
    End If
    
    CUS_Str_MainProgram = Ori_CUST_Str_MainProgram
    CUS_Str_MainProgram = Replace(UCase(CUS_Str_MainProgram), UCase("Loop_Idx"), CStr(Loop_Idx))
    CUS_Str_MainProgram = Replace(UCase(CUS_Str_MainProgram), UCase("HexSrcCode"), CStr(Loop_count))
    ''                                    CUS_Str_MainProgram = Replace(UCase(CUS_Str_MainProgram), UCase("HexSrcStep"), CStr(Loop_Step))
     Loop_Idx = Loop_Idx + 1
     
End Function

Public Function CUS_VIR_MainProgram_MeasV_CalR(TestPinArrayIV() As String, TestSeqNum As Integer, _
                            CUS_CalR_Seq() As String, ForceI() As String, MeasV As PinListData, CUS_VDD As Double) As Long ''VBT_CUSTOMIZE_HardIP
    
    Dim ForceI_Ary() As String
    Dim ForceISeqIndexPerSeq As Long ''Split by "+"
                                                                                                                                                                                                                   
    Dim R As New PinListData
    Dim R_AK() As Double
                                                                                                                                                                                                                   
    Dim v As Double ''voltage
    Dim i As Double ''current
    Dim p As Long
    Dim PinAry() As String
    Dim FirstSite As Integer
                                                                                                                                                                                                                   
    If (UBound(ForceI) <> 0) Then
        ForceISeqIndexPerSeq = TestSeqNum
    Else
        ForceISeqIndexPerSeq = 0
    End If
                                                                                                                                                                                                                   
    ForceI_Ary = Split(ForceI(ForceISeqIndexPerSeq), ",")
                                                                                                                                                                                                                   
    PinAry = Split(TestPinArrayIV(TestSeqNum), ",")
                                                                                                                                                                                                                   
    FirstSite = 0
    
    If (UBound(ForceI_Ary) = 0) Then
                                                                                                                                                                                                                   
        For Each site In TheExec.sites
                                                                                                                                                                                                                   
            For p = 0 To MeasV.Pins.Count - 1
                                                                                                                                                                                                                   
                If (FirstSite = 0) Then
                    R.AddPin (MeasV.Pins(p))
                End If
                                                                                                                                                                                                                   
                If (UCase(CUS_CalR_Seq(TestSeqNum)) Like "*RVOH*") Then
                                                                                                                                                                                                                   
                    R.Pins(p).Value(site) = CUS_VDD - MeasV.Pins(p).Value(site)
                    R.Pins(p).Value(site) = R.Pins(p).Divide(CDbl(ForceI_Ary(0)))
                                                                                                                                                                                                                   
                ElseIf (UCase(CUS_CalR_Seq(TestSeqNum)) Like "*RVOL*") Then
                                                                                                                                                                                                                   
                    R.Pins(p).Value(site) = MeasV.Pins(p).Value(site)
                    R.Pins(p).Value(site) = R.Pins(p).Divide(CDbl(ForceI_Ary(0)))
                End If
                                                                                                                                                                                                                   
                'R_AK = TheHdw.PPMU.ReadRakValuesByPinnames(MeasV.Pins(p), site)  ''Get instrament impedance
                                                                                                                                                                                                                   
''                If InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then                          ''Get DIB impedance
''                    R_AK(0) = R_AK(0) + FT_Card_RAK.Pins(MeasV.Pins(p)).Value(Site)
''                Else
''                    R_AK(0) = R_AK(0) + CP_Card_RAK.Pins(MeasV.Pins(p)).Value(Site)
''                End If
                R_AK(0) = CurrentJob_Card_RAK.Pins(MeasV.Pins(p)).Value(site)
                                                                                                                                                                                                                   
                R.Pins(p).Value(site) = R.Pins(p).Value(site) - R_AK(0)
                                                                                                                                                                                                                   
            Next p
            FirstSite = FirstSite + 1
        Next site
                                                                                                                                                                                                                   
        TheExec.Flow.TestLimit resultVal:=R, Unit:=unitCustom, Tname:="Calculate_" + CUS_CalR_Seq(TestSeqNum), customUnit:="ohm", ForceResults:=tlForceNone
                                                                                                                                                                                                                   
    End If

End Function

Public Function AnalyzeCusStrToCalcR(CUS_Str_MainProgram As String, TestSeqNum As Integer, ForceSequenceArray() As String, MeasCurr As PinListData, RTN_Imped As PinListData) As Long
    
    ''ex.  CUS_Str_MainProgram = "CalR;_VDDQL_DDR0_VAR_H;RVOH,RVOL,RVOH,RVOL"
    ''========================================================================================
    Dim CUS_CalR_VDD As Double
    Dim CUS_CalR_Seq_Ary() As String
    Dim CUS_CalR_Arg_Ary() As String
    Dim Ary_str(0) As String
    
    CUS_CalR_Arg_Ary = Split(CUS_Str_MainProgram, ";")
    Ary_str(0) = CUS_CalR_Arg_Ary(1)
    Call HIP_Evaluate_ForceVal(Ary_str)
    CUS_CalR_VDD = CDbl(Ary_str(0))
    CUS_CalR_Seq_Ary = Split(CUS_CalR_Arg_Ary(2), ",")
    ''========================================================================================
    Dim p As Long
    Dim b_FirstTime As Boolean
    b_FirstTime = True
    Dim ForceVolt As Double
    
    If UBound(ForceSequenceArray) = 0 Then
        ForceVolt = ForceSequenceArray(0)
    Else
        ForceVolt = ForceSequenceArray(TestSeqNum)
    End If
    
    For Each site In TheExec.sites
        For p = 0 To MeasCurr.Pins.Count - 1
                                                                                                                                                                                                               
            If b_FirstTime = True Then
                RTN_Imped.AddPin (MeasCurr.Pins(p))
            End If
                                                                                                                                                                                                               
            If (UCase(CUS_CalR_Seq_Ary(TestSeqNum)) Like "*RVOH*") Then
                                                                                                                                                                                                               
                RTN_Imped.Pins(p).Value(site) = CUS_CalR_VDD - ForceVolt
                RTN_Imped.Pins(p).Value(site) = RTN_Imped.Pins(p).Divide(MeasCurr.Pins(p)).Multiply(-1)
                                                                                                                                                                                                               
            ElseIf (UCase(CUS_CalR_Seq_Ary(TestSeqNum)) Like "*RVOL*") Then
                                                                                                                                                                                                               
                RTN_Imped.Pins(p).Value(site) = ForceVolt
                RTN_Imped.Pins(p).Value(site) = RTN_Imped.Pins(p).Divide(MeasCurr.Pins(p))
            End If
        Next p
        b_FirstTime = False
    Next site
End Function

Public Function Cust_Sweep_V()
    Dim StepIndex_Val As Long
    Dim StartVolt_Val_1 As Double, StepVolt_Val_1 As Double, ForceVolt_Val_1 As Double
    Dim StartVolt_Val_2 As Double, StepVolt_Val_2 As Double, ForceVolt_Val_2 As Double
    Dim Force_Pins_1 As String, Force_Pins_2 As String
    StartVolt_Val_1 = 0.99
    StartVolt_Val_2 = 0.82
    StepVolt_Val_1 = 0.02
    StepVolt_Val_2 = 0.02
    Force_Pins_1 = "VDDIO11_RET_DDR0,VDDIO11_RET_DDR1"
    Force_Pins_2 = "VDD_DCS_DDR0,VDD_DCS_DDR1"
    
    StepIndex_Val = CDbl(Val(TheExec.Flow.var("SrcCodeIndx").Value))
    ForceVolt_Val_1 = StartVolt_Val_1 + StepIndex_Val * StepVolt_Val_1
    ForceVolt_Val_2 = StartVolt_Val_2 + StepIndex_Val * StepVolt_Val_2

    TheHdw.DCVS.Pins(Force_Pins_1).Voltage.Value = ForceVolt_Val_1
    TheHdw.DCVS.Pins(Force_Pins_2).Voltage.Value = ForceVolt_Val_2
    
    TheExec.Datalog.WriteComment ("Force Pin " & Force_Pins_1 & " Value = " & ForceVolt_Val_1)
    TheExec.Datalog.WriteComment ("Force Pin " & Force_Pins_2 & " Value = " & ForceVolt_Val_2)
End Function
Public Function VOLH_Sweep(CUS_Str_DigSrcData As String, DigSrc_Assignment As String) As Long
    Dim SrcCode_Initial_Dec As Long
    Dim SrcCode_Target_Dec As Long
    Dim SrcCode_Target_Bin() As Long
    ReDim SrcCode_Target_Bin(9) As Long
    Dim SrcCode_Target_Bin_One As String: SrcCode_Target_Bin_One = ""
    Dim SrcCode_Target_Bin_Two As String: SrcCode_Target_Bin_Two = ""
    Dim i As Long
    Dim SplitArray() As String
    Dim ReplaceTarget_1 As String
    Dim ReplaceTarget_2 As String
    
    '' Split with comma
    SplitArray = Split(CUS_Str_DigSrcData, ",")
    ReplaceTarget_1 = SplitArray(1)
    ReplaceTarget_2 = SplitArray(2)
    
    SrcCode_Target_Dec = TheExec.Flow.var("SrcCodeIndx").Value
    Call Dec2Bin(Abs(SrcCode_Target_Dec), SrcCode_Target_Bin)
    
    For i = 0 To 9
        If i < 3 Then
            SrcCode_Target_Bin_One = SrcCode_Target_Bin(i) & SrcCode_Target_Bin_One
        Else
            SrcCode_Target_Bin_Two = SrcCode_Target_Bin(i) & SrcCode_Target_Bin_Two
        End If
    Next i
    SrcCode_Target_Bin_One = SrcCode_Target_Bin_One & "0"
    SrcCode_Target_Bin_Two = SrcCode_Target_Bin_Two & "0"
    DigSrc_Assignment = Replace(DigSrc_Assignment, ReplaceTarget_1, SrcCode_Target_Bin_One)
    DigSrc_Assignment = Replace(DigSrc_Assignment, ReplaceTarget_2, SrcCode_Target_Bin_Two)
End Function

Public Function MTR_UVI80_Setup()
    ' 20170227 - Set current range only for osprey Metrology 20170227
            TheHdw.DCVI.Pins("mtr_analog_test_p").CurrentRange = 0.002
            TheHdw.DCVI.Pins("mtr_analog_test_p").current = 0.002

End Function


Public Function CUS_DDR_Emulate_Const_Res_Loading(MeasureValue As PinListData, ForceValByPin() As String, CUS_Str_MainProgram As String, TestSeqNum As Integer, _
    Optional RAK_Flag As Enum_RAK = 0) As Long

    Dim R0_Value As New PinListData
    Dim Final_Volt_Value As New PinListData
    Dim Final_Force_Curr_Value As New PinListData
    Dim R1_Value As New PinListData
    Dim Adjust_I_Value As New PinListData
    Dim site As Variant
    Dim Pin  As Variant
    Dim Temp_Input() As String
    
''''''''''    Dim Pwr_Voltage As Double: Pwr_Voltage = TheHdw.DCVS.Pins("VDDQL_DDR").Voltage.Value
    '   Hardcode for debug
    Dim Pwr_Voltage As Double: Pwr_Voltage = TheHdw.DCVS.Pins("VDDA_LP5_LP5").Voltage.Value
    
    Dim Target_Resistance As Double
    Dim Flag_1 As Boolean: Flag_1 = False
    Dim Flag_2 As Boolean: Flag_2 = False
    Dim Initial_Setting_Flag As Boolean: Initial_Setting_Flag = True
    Dim Counter_Meas As Integer: Counter_Meas = 1
    Dim Counter_End As Integer: Counter_End = 50
    Dim ForceValue As Double: ForceValue = Abs(ForceValByPin(0))
    Dim p As Integer
    Dim RakV() As Double
    Dim Ary_str(0) As String
    Dim HiLimit As Double
    Dim LoLimit As Double
    Dim Pin_Diff As String
    Dim error_flag As Boolean: error_flag = False
    Dim inst_name As String: inst_name = TheExec.DataManager.instanceName
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    
''''''''''    Ary_str(0) = "_VDDQL_DDR_VAR_H"
    ' Edied by Dyaln 20190909 , Hardcode for debug
    Ary_str(0) = "VDDA_LP5_LP5_VAR_H"


    Call HIP_Evaluate_ForceVal(Ary_str)
    
    Temp_Input() = Split(CUS_Str_MainProgram, ":")     ' CUS_Str_MainProgram . Ex: Tname:VOL,VOH,VOL,VOH,48
    Temp_Input() = Split(Temp_Input(1), ",")
    Target_Resistance = Temp_Input(UBound(Temp_Input))
    
    If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
        HiLimit = CDbl(Ary_str(0))
        LoLimit = 0.65 * CDbl(Ary_str(0))
    Else
        HiLimit = 0.489 * CDbl(Ary_str(0))
        LoLimit = 0
    End If
    
    
    For p = 0 To MeasureValue.Pins.Count - 1
        Flag_1 = False
        Pin = MeasureValue.Pins(p).Name
        
        For Each site In TheExec.sites
            Initial_Setting_Flag = True
            error_flag = False
            
                For Counter_Meas = 1 To Counter_End
                    If Flag_1 = False Then
                        R0_Value.AddPin (Pin)
                        Final_Volt_Value.AddPin (Pin)
                        Final_Force_Curr_Value.AddPin (Pin)
                        R1_Value.AddPin (Pin)
                        Adjust_I_Value.AddPin (Pin)
                        Flag_1 = True
                    End If
                    
                    If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                        If Initial_Setting_Flag = True Then
                            R0_Value.Pins(Pin) = MeasureValue.Pins(Pin).Value(site) / ForceValue
                        Else
                            R0_Value.Pins(Pin) = MeasureValue.Pins(Pin).Value(site) / Abs(Adjust_I_Value.Pins(Pin).Value(site))
                        End If
                    ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                        If Initial_Setting_Flag = True Then
                            R0_Value.Pins(Pin) = (Pwr_Voltage - MeasureValue.Pins(Pin).Value(site)) / ForceValue
                        Else
                            R0_Value.Pins(Pin) = (Pwr_Voltage - MeasureValue.Pins(Pin).Value(site)) / Adjust_I_Value.Pins(Pin).Value(site)
                        End If
                    End If
                    
                    If ((Abs(R0_Value.Pins(Pin).Value(site) - Target_Resistance)) / Target_Resistance) < 0.1 Then
                        Final_Volt_Value.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site)
                        
                        If Counter_Meas = 1 Then
                            If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                                Final_Force_Curr_Value.Pins(Pin).Value(site) = -1 * ForceValue
                            ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                                Final_Force_Curr_Value.Pins(Pin).Value(site) = ForceValue
                            End If
                        Else
                            Final_Force_Curr_Value.Pins(Pin).Value(site) = Adjust_I_Value.Pins(Pin).Value(site)
                        End If
                        Counter_Meas = Counter_End + 1 ' Iteration search done , Exit from Loop
                        
                        ' 20171204 Update latest R1_Value
                        If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                            R1_Value.Pins(Pin).Value(site) = (Pwr_Voltage - Final_Volt_Value.Pins(Pin).Value(site)) / Abs(Final_Force_Curr_Value.Pins(Pin).Value(site))
                        ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                            R1_Value.Pins(Pin).Value(site) = Final_Volt_Value.Pins(Pin).Value(site) / Final_Force_Curr_Value.Pins(Pin).Value(site)
                        End If
                        
                        If R1_Value.Pins(Pin).Value(site) = 0 Then
                            TheExec.Datalog.WriteComment ("Final " & "Site_" & site & " pin = " & Pin & " R0 value = " & Format(R0_Value.Pins(Pin).Value(site), "0.000") & " R1 value = " & "NA" & _
                                                                            " Meas Volt = " & Format(Final_Volt_Value.Pins(Pin).Value(site), "0.0000"))
                        Else
                            TheExec.Datalog.WriteComment ("Final " & "Site_" & site & " pin = " & Pin & " R0 value = " & Format(R0_Value.Pins(Pin).Value(site), "0.000") & " R1 value = " & Format(R1_Value.Pins(Pin).Value(site), "0.000") & _
                                                                            " Meas Volt = " & Format(Final_Volt_Value.Pins(Pin).Value(site), "0.0000"))
                        End If
                    Else
                                                            
                        If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                            If Initial_Setting_Flag = True Then
                                R1_Value.Pins(Pin).Value(site) = (Pwr_Voltage - MeasureValue.Pins(Pin).Value(site)) / ForceValue
                            Else
                                R1_Value.Pins(Pin).Value(site) = (Pwr_Voltage - MeasureValue.Pins(Pin).Value(site)) / Abs(Adjust_I_Value.Pins(Pin).Value(site))
                            End If
                       ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                            If Initial_Setting_Flag = True Then
                                R1_Value.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site) / ForceValue
                            Else
                                R1_Value.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site) / Adjust_I_Value.Pins(Pin).Value(site)
                            End If
                        End If
                        
                        If (R1_Value.Pins(Pin).Value(site) + Target_Resistance) = 0 Then
                            Adjust_I_Value.Pins(Pin).Value(site) = 0
                            TheExec.Datalog.WriteComment (" Error : Denominator = 0 !  ")
                            error_flag = True
                        Else
                            Adjust_I_Value.Pins(Pin).Value(site) = Pwr_Voltage / (R1_Value.Pins(Pin).Value(site) + Target_Resistance)
                        End If
                        
                        If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                            Adjust_I_Value.Pins(Pin).Value(site) = (-1) * Adjust_I_Value.Pins(Pin).Value(site)
                        End If
                        
                        Initial_Setting_Flag = False    ' False means start to use Adjust I Value for next Iteration search
                        
                        If error_flag = False Then
                            ' Update Force Condition and Measure Voltage
                            TheHdw.Digital.Pins(Pin).Disconnect
                            TheHdw.PPMU.Pins(Pin).ForceI 0, 0.002
                            TheHdw.PPMU.Pins(Pin).Connect
                            TheHdw.PPMU.Pins(Pin).Gate = tlOn
                         
                             'if PPMU > 50 mA set Warning and set PPMU = 50 mA
                             If Abs(Adjust_I_Value.Pins(Pin).Value(site)) <= 50 * mA Then
                                TheHdw.PPMU.Pins(Pin).ForceI Adjust_I_Value.Pins(Pin).Value(site), Abs(Adjust_I_Value.Pins(Pin).Value(site))
                               
                                'NEW 20170728
                                If UCase(Pin) Like UCase("DDR*_P*") Or UCase(Pin) Like UCase("DDR*_N*") Then  ' Differential pair needs to force opposite current
                                    If UCase(Pin) Like ("DDR*_DQS_P*") Then             'DDR0_DQS_P0
                                        Pin_Diff = Replace(UCase(Pin), "DQS_P", "DQS_N")
                                    ElseIf UCase(Pin) Like ("DDR*_DQS_N*") Then
                                        Pin_Diff = Replace(UCase(Pin), "DQS_N", "DQS_P")
                                    ElseIf UCase(Pin) Like ("DDR*_CK*_P*") Then          'DDR0_CK_P
                                        Pin_Diff = Replace(UCase(Pin), "_P", "_N")
                                    ElseIf UCase(Pin) Like ("DDR*_CK*_N*") Then
                                        Pin_Diff = Replace(UCase(Pin), "_N", "_P")
                                    End If
                                    TheHdw.PPMU.Pins(Pin_Diff).ForceI (-1) * Adjust_I_Value.Pins(Pin).Value(site), Abs(Adjust_I_Value.Pins(Pin).Value(site))
                                    TheExec.Datalog.WriteComment ("Pin Diff: " & Pin_Diff & " Site (" & site & ")" & " , Force Value : " & Format((-1) * Adjust_I_Value.Pins(Pin).Value(site) * 1000, "0.000") & "mA")
                                End If
                                
                            Else
                                TheExec.Datalog.WriteComment (" Error : Irange >= 50mA , Bypass Pin " & Pin & " Measurement ")
                                MeasureValue.Pins(Pin).Value(site) = 0
                                Final_Volt_Value.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site)
                                Final_Force_Curr_Value.Pins(Pin).Value(site) = Adjust_I_Value.Pins(Pin).Value(site)
                                TheHdw.PPMU.Pins(Pin).Gate = tlOff
                                TheHdw.PPMU.Pins(Pin).Disconnect
                                TheHdw.Digital.Pins(Pin).Connect
                                Exit For
                            End If
                        
                            TheHdw.Wait 0.002
                             MeasureValue.Pins(Pin).Value(site) = TheHdw.PPMU.Pins(Pin).Read(tlPPMUReadMeasurements, 10)
                            
                            TheHdw.PPMU.Pins(Pin).ForceI 0, 0
                            TheHdw.PPMU.Pins(Pin).Gate = tlOff
                            TheHdw.PPMU.Pins(Pin).Disconnect
                            TheHdw.Digital.Pins(Pin).Connect
                         
                            '' Calculate RAK
                            If RAK_Flag = R_TraceOnly Then
                                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(pin, Site)
                                
''                                If InStr(TheExec.CurrentChanMap, "CP") <> 0 Then
''                                    MeasureValue.Pins(Pin).Value(Site) = MeasureValue.Pins(Pin).Value(Site) - Adjust_I_Value.Pins(Pin).Value(Site) * (CP_Card_RAK.Pins(Pin).Value(Site) + RakV(0))
''                                Else
''                                    MeasureValue.Pins(Pin).Value(Site) = MeasureValue.Pins(Pin).Value(Site) - Adjust_I_Value.Pins(Pin).Value(Site) * (FT_Card_RAK.Pins(Pin).Value(Site) + RakV(0))  ' + TheHdw.PPMU.ReadRakValuesByPinnames(FT_Card_RAK.Pins(pin).Name, Site))
''                                End If
                                MeasureValue.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site) - Adjust_I_Value.Pins(Pin).Value(site) * (CurrentJob_Card_RAK.Pins(Pin).Value(site))
                            
                            ElseIf RAK_Flag = R_PathWithContact Then
                                MeasureValue.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site) - Adjust_I_Value.Pins(Pin).Value(site) * R_Path_PLD.Pins(Pin).Value(site)
                            End If
                       Else
                            MeasureValue.Pins(Pin).Value(site) = 0
                            Final_Volt_Value.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site)
                            Final_Force_Curr_Value.Pins(Pin).Value(site) = Adjust_I_Value.Pins(Pin).Value(site)
                            Counter_Meas = Counter_End + 1 'Exit from loop
                       End If
                       
                        If R1_Value.Pins(Pin).Value(site) = 0 Then
                            TheExec.Datalog.WriteComment ("Adjust " & "Site_" & site & " pin = " & Pin & " R0 value = " & Format(R0_Value.Pins(Pin).Value(site), "0.000") & " R1 value = " & "NA" & _
                                                                            " Meas Volt = " & Format(MeasureValue.Pins(Pin).Value(site), "0.0000"))
                        Else
                            TheExec.Datalog.WriteComment ("Adjust " & "Site_" & site & " pin = " & Pin & " R0 value = " & Format(R0_Value.Pins(Pin).Value(site), "0.000") & " R1 value = " & Format(R1_Value.Pins(Pin).Value(site), "0.000") & _
                                                                            " Meas Volt = " & Format(MeasureValue.Pins(Pin).Value(site), "0.0000"))
                        End If
                        
                       If Counter_Meas = Counter_End Then
                            Final_Volt_Value.Pins(Pin).Value(site) = MeasureValue.Pins(Pin).Value(site)
                            Final_Force_Curr_Value.Pins(Pin).Value(site) = Adjust_I_Value.Pins(Pin).Value(site)
                       End If
                        
                    End If
                    
               Next Counter_Meas
        Next site
    Next p


    Dim testName As String
    Dim Temp_index
    
    

    Temp_index = TheExec.Flow.TestLimitIndex

    For Each Pin In Final_Volt_Value.Pins
        
        TheExec.Flow.TestLimitIndex = Temp_index
        testName = Report_TName_From_Instance(CalcC, CStr(Pin))
        'TestName = Report_TName_From_Instance("V", CStr(pin))
        For Each site In TheExec.sites
             If UCase(inst_name) Like "*VOLH_SWEEP*LOOP*" Then
                 TheExec.Flow.TestLimit Final_Volt_Value.Pins(Pin).Value(site), PinName:=Final_Volt_Value.Pins(Pin).Name, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=CStr(Temp_Input(TestSeqNum)) & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(Pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
             ElseIf UCase(inst_name) Like "*VOLH_SWEEP*" Then       ' 20170912 Used for VOLH_SWEEP Average ZCAL test
                Dim ZCAL_Testname As String: ZCAL_Testname = "Average_ZCAL"
                TheExec.Flow.TestLimit Final_Volt_Value.Pins(Pin).Value(site), lowVal:=LoLimit, hiVal:=HiLimit, PinName:=Final_Volt_Value.Pins(Pin).Name, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=ZCAL_Testname & "_" & CStr(Temp_Input(TestSeqNum)) & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(Pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
                Else
                'TheExec.Flow.TestLimit Final_Volt_Value.Pins(Pin).Value(Site), lowVal:=LoLimit, hiVal:=HiLimit, PinName:=Final_Volt_Value.Pins(Pin).Name, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestName & "_" & CStr(TestSeqNum), forceVal:=Final_Force_Curr_Value.Pins(Pin).Value(Site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimit Final_Volt_Value.Pins(Pin).Value(site), lowVal:=LoLimit, hiVal:=HiLimit, PinName:=Final_Volt_Value.Pins(Pin).Name, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=testName, ForceVal:=Final_Force_Curr_Value.Pins(Pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                'TheExec.Flow.TestLimit Final_Volt_Value.Pins(pin).Value(site), lowval:=LoLimit, hival:=HiLimit, PinName:=Final_Volt_Value.Pins(pin).name, ScaleType:=scaleNone, Unit:=unitVolt, FormatStr:="%.3f", TName:=CStr(Temp_Input(TestSeqNum)) & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
                End If
             TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
        Next site
                    
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
    Next Pin
                    
    '20171204 Print Out Final Pull Up / Pull Down resistance
    Dim Res_Tname As String: Res_Tname = ""
    Dim Target_Resistance_Hi_Lim As Double: Target_Resistance_Hi_Lim = Target_Resistance * 1.1
    Dim Target_Resistance_Lo_Lim As Double: Target_Resistance_Lo_Lim = Target_Resistance * 0.9
                
    'Output Datalog PU/PD Info
    TheExec.Datalog.WriteComment ("")
             
    For Each Pin In Final_Volt_Value.Pins
        For Each site In TheExec.sites
            If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                Res_Tname = "R_Pull_Up"
            ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                Res_Tname = "R_Pull_Down"
                End If
                
            If UCase(inst_name) Like "*VOLH_SWEEP*LOOP*" Then
                TheExec.Flow.TestLimit R1_Value.Pins(Pin).Value(site), PinName:=Final_Volt_Value.Pins(Pin).Name, scaletype:=scaleNone, Unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=Res_Tname & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(Pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
            ElseIf UCase(inst_name) Like "*VOLH_SWEEP*" Then
               TheExec.Flow.TestLimit R1_Value.Pins(Pin).Value(site), lowVal:=Target_Resistance_Lo_Lim, hiVal:=Target_Resistance_Hi_Lim, PinName:=Final_Volt_Value.Pins(Pin).Name, scaletype:=scaleNone, Unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=ZCAL_Testname & "_" & Res_Tname & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(Pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
                Else
               TheExec.Flow.TestLimit R1_Value.Pins(Pin).Value(site), lowVal:=Target_Resistance_Lo_Lim, hiVal:=Target_Resistance_Hi_Lim, PinName:=Final_Volt_Value.Pins(Pin).Name, scaletype:=scaleNone, Unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=Res_Tname & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(Pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
                End If
        Next site
    Next Pin
    
End Function

Public Function CUS_DDR_DCS_PrintOut() As Long

    TheExec.Datalog.WriteComment ("")
    TheExec.Datalog.WriteComment (" VDD_DCS_DDR Level = " & FormatNumber(TheHdw.DCVS.Pins("VDD_DCS_DDR").Voltage.Value, 3) & " v")
    TheExec.Datalog.WriteComment ("")

End Function
Public Function MEAS_I_ABS(MeasureValue As PinListData) As Long
Dim p As Long

For p = 0 To MeasureValue.Pins.Count - 1
    MeasureValue.Pins(p) = MeasureValue.Pins(p).Abs
Next p

End Function

Public Function CUS_RREF_Rak_Calc(ByRef MeasureVolt As PinListData) As Long

    Dim Ary_str(0) As String
    Dim VDDQL_Val As Double
    Dim MeasV As Double
    Dim pin_name As String
    Dim GetRakVal As Double
    Dim p As Long
    'Dim RakV() As Double
    
    Ary_str(0) = "_VDDQL_DDR_VAR"
    Call HIP_Evaluate_ForceVal(Ary_str)
    VDDQL_Val = CDbl(Ary_str(0))

    For p = 0 To MeasureVolt.Pins.Count - 1 Step 1
          pin_name = MeasureVolt.Pins(p)
          
          For Each site In TheExec.sites.Active
                MeasV = MeasureVolt.Pins(pin_name).Value(site)
                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(pin_name, site)
                GetRakVal = CurrentJob_Card_RAK.Pins(pin_name).Value(site)
                MeasureVolt.Pins(pin_name).Value(site) = MeasV - GetRakVal * (VDDQL_Val - MeasV) / 240
          Next site
          
    Next p
        
End Function
Public Function CUS_AMP_SDLL_SWP(MeasFreq As PinListData, Extra_TName As String) As Long
   
    Dim Dict_Freq_PLD As New PinListData
    Dim Step_Freq_PLD As New PinListData
    Dim pin_name As Variant
    Dim Extra_TName_StrArr() As String
    Dim Dict_idx As Long
    Dim Step_idx As Long
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    Dict_Freq_PLD = GetStoredMeasurement("Freq_SDLL_SWP")
    
    Dim p As Long
     
     For p = 0 To MeasFreq.Pins.Count - 1
         If InStr(UCase(MeasFreq.Pins(p)), "_P") Then
            pin_name = MeasFreq.Pins(p).Name
         
             For Dict_idx = 0 To Dict_Freq_PLD.Pins.Count - 1
                 If (pin_name = Dict_Freq_PLD.Pins(Dict_idx).Name) Then
                     Step_Freq_PLD.AddPin (pin_name)
                     
                     For Each site In TheExec.sites.Active
                         If MeasFreq.Pins(p).Value(site) = 0 Or Dict_Freq_PLD.Pins(Dict_idx).Value(site) = 0 Then
                            Step_Freq_PLD.Pins(pin_name).Value(site) = 999
                            TheExec.Datalog.WriteComment (" Freq measurement = 0 , Denominator = 0  ! ")
                         Else
                            Step_Freq_PLD.Pins(pin_name).Value(site) = 0.5 * ((1 / MeasFreq.Pins(p).Value(site)) - (1 / Dict_Freq_PLD.Pins(Dict_idx).Value(site)))
                         End If
                     Next site
                     
                     Exit For
                 End If
             Next Dict_idx
         End If
     Next p

     Extra_TName_StrArr = Split(Extra_TName, ":")
              
    For Step_idx = 0 To Step_Freq_PLD.Pins.Count - 1
            pin_name = Step_Freq_PLD.Pins(Step_idx).Name
            pin_name = Replace(LCase(pin_name), "ddr", "ch")
            pin_name = Replace(LCase(pin_name), "dqs_p", "core")
            
            '' Extra_TName_StrArr (0) : Frequency
            '' Extra_TName_StrArr (1) : Octant
            '' Extra_TName_StrArr (2) : Loop_Idx
            '' Extra_TName_StrArr (3) : Sweep_Name (LSW_0X or MSW_0X)
           TestNameInput = Report_TName_From_Instance("F", Step_Freq_PLD.Pins(Step_idx), "Step_SDLL_SWP", 0, Step_idx)
           '''         OutputTname_format(8) = CStr(gl_Tname_Alg_Index)           TheExec.Flow.TestLimit resultVal:=Step_Freq_PLD.Pins(Step_idx), Unit:=unitTime, Tname:=TestNameInput, ForceResults:=tlForceNone

    Next Step_idx
  
End Function
Public Function ADCLK_Matrix_Loading()
Dim ADCLK_Matrix_Sheet As Worksheet: Set ADCLK_Matrix_Sheet = Sheets("Flow_HARDIP_ADCLK")
Dim Column_Index As Long: Column_Index = 1
Dim Row_Index As Long: Row_Index = 1
Dim Matrix_Index As Long
Dim ADCLK_Matrix_Index As Long: ADCLK_Matrix_Index = 0
Dim ADCLK_Matrix_Range As Variant
Dim Max_Rows_Count As Long
Dim Max_Columns_Count As Long

With ADCLK_Matrix_Sheet
    Max_Rows_Count = .UsedRange.Rows.Count
    Max_Columns_Count = .UsedRange.Columns.Count
    ADCLK_Matrix_Range = .range(.Cells(5, 1), .Cells(Max_Rows_Count, Max_Columns_Count))
End With

Dim add_Matrix_Sheet As Worksheet: Set add_Matrix_Sheet = Sheets("add")
Dim add_Matrix_Index As Long: add_Matrix_Index = 0
Dim add_Matrix_Range As Variant
Dim Max_Rows_Count_A As Long
Dim Max_Columns_Count_A As Long
Dim Column_Index_A As Long: Column_Index_A = 1
Dim Row_Index_A As Long: Row_Index_A = 1


With add_Matrix_Sheet
    Max_Rows_Count_A = .UsedRange.Rows.Count
    Max_Columns_Count_A = .UsedRange.Columns.Count
    add_Matrix_Range = .range(.Cells(1, 1), .Cells(Max_Rows_Count_A, Max_Columns_Count_A))
End With

Dim temp_str() As String

For Row_Index = 1 To Max_Rows_Count - 4
    For Column_Index = 0 To Max_Columns_Count
        If ADCLK_Matrix_Range(Row_Index, 7) = "Use-Limit" And ADCLK_Matrix_Range(Row_Index, 14) = "Hz" Then
            temp_str() = Split(ADCLK_Matrix_Range(Row_Index, 8), "_")
                    For Row_Index_A = 1 To Max_Rows_Count_A
                      'For Column_Index_A = 0 To Max_Columns_Count_A
                        If temp_str(14) = add_Matrix_Range(Row_Index_A, 1) Then
                                 ADCLK_Matrix_Range(Row_Index, 11) = add_Matrix_Range(Row_Index_A, 2)
                                 ADCLK_Matrix_Range(Row_Index, 12) = add_Matrix_Range(Row_Index_A, 3)
                         End If
                      'Next Column_Index_A
                  Next Row_Index_A
        End If
    Next Column_Index
Next Row_Index

With ADCLK_Matrix_Sheet
.range(.Cells(5, 1), .Cells(Max_Rows_Count, Max_Columns_Count)) = ADCLK_Matrix_Range
End With

End Function

