Attribute VB_Name = "LIB_HardIP_Calc"
Option Explicit
Public R_Path_PLD As New PinListData
Public CMRR_Values(10) As New SiteDouble
'ReDim CMRR_Value(10) As New SiteDouble
Public glb_ind As Long
Public gl_Save_deltavalue As New DSPWave
Public gl_Save_MeasNum As New DSPWave
Public gl_Save_DCK As New DSPWave

Type Type_MonoWithBlock
    Block As Long
    DSP_Bin As New DSPWave
    DSP_Dec As New DSPWave
End Type

Public Const Calc = "Calc"
Public Const CalcV = "CalcV" '"V"
Public Const CalcI = "CalcI" '"I"
Public Const CalcF = "CalcF" '"F"
Public Const CalcR = "CalcR" '"R"
Public Const CalcC = "CalcC" '"C"
Public Const CalcT = "CalcT" '"T"

'Public meas_val_before() As New SiteDouble
Public meas_val_before As New PinListData
Public meas_val_delay_instance_name As String
Public meas_val_first(10) As New SiteDouble
Public meas_val_delay_instance As String

Public Function Calc_delay(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    
    Dim meas_name As String
    Dim site As Variant
    Dim result As New SiteDouble
    Dim meas_val As New SiteDouble

    
    For i = 0 To argc - 1
        meas_name = argv(i)
        meas_val = GetStoredMeasurement(meas_name)
    
            If TheExec.TesterMode = testModeOffline Then
                meas_val = Rnd() * 1000000000000#
            End If
            
            If meas_val_delay_instance <> TheExec.DataManager.instanceName Then
                meas_val_first(i) = meas_val
            Else
                For Each site In TheExec.sites
                    If meas_val = 0 Then meas_val = 0.0000000001
                    If meas_val_first(i) = 0 Then meas_val_first(i) = 0.0000000001
                    
                    result = Format(meas_val.Invert.Subtract(meas_val_first(i).Invert).Multiply(0.5), "0.00000000000000000000000000")
                Next
                meas_val_first(i) = meas_val
                TestNameInput = "Time delay F" + CStr(i + 1)
                
                TheExec.Flow.TestLimit resultVal:=result, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
                
            End If
        
    Next i
    meas_val_delay_instance = TheExec.DataManager.instanceName
    
End Function
Public Function Calc_delay_Sicily(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim TestNameInput As String
    Dim Dict_Name As String
    Dim site As Variant
    Dim meas_val_now As New PinListData
    Dim meas_val As New PinListData

    Set meas_val = Nothing
    Dict_Name = argv(i)
    meas_val = GetStoredMeasurement(Dict_Name)

    
    If meas_val_delay_instance_name <> TheExec.DataManager.instanceName Then 'first time to enter this function
        Set meas_val_before = Nothing
        meas_val_before = meas_val
    Else
        Set meas_val_now = Nothing
        meas_val_now = meas_val
        '=================prevent divide 0==============
        For Each site In TheExec.sites
            For i = 0 To meas_val_now.Pins.Count - 1
                If meas_val_now.Pins(i).Value = 0 Then
                    meas_val_now.Pins(i).Value = 0.0000000001
                End If
            Next i
            For i = 0 To meas_val_before.Pins.Count - 1
                If meas_val_before.Pins(i).Value = 0 Then
                    meas_val_before.Pins(i).Value = 0.0000000001
                End If
            Next i
        Next site
        '=================prevent divide 0==============
        meas_val_now = meas_val_now.Math.Invert.Subtract(meas_val_before.Math.Invert).Multiply(0.5)
        Dim PLD_For_TestLimit As New PinListData
        For i = 0 To meas_val_now.Pins.Count - 1
            If UCase(meas_val_now.Pins(i)) Like "*DQS_P*" Then
                PLD_For_TestLimit.AddPin (meas_val_now.Pins(i))
                For Each site In TheExec.sites
                    PLD_For_TestLimit.Pins(meas_val_now.Pins(i)).Value = meas_val_now.Pins(i).Value
                Next site
            End If
        Next i
        TestNameInput = Report_TName_From_Instance(CalcC, "")
        TheExec.Flow.TestLimit resultVal:=PLD_For_TestLimit, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
        meas_val_before = meas_val
    End If

    meas_val_delay_instance_name = TheExec.DataManager.instanceName
    
End Function

Public Function Calc_SetFlag(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    
    Dim meas_name As String
    Dim site As Variant
    Dim meas_val As New SiteDouble


    For i = 0 To argc - 1
        meas_name = argv(i)
        meas_val = GetStoredMeasurement(meas_name)
        For Each site In TheExec.sites
            If meas_val(site) = 0 Then TheExec.sites(site).FlagState("F_" + meas_name) = logicTrue
        Next
    Next i
    
End Function
Public Function Calc_GrayCode(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String



    For i = 0 To argc - 1
        DSPWave_Dict = GetStoredCaptureData(argv(i))
        TestNameInput = TestNameInput & argv(i)
        Call rundsp.Transfer2GrayCode(DSPWave_Dict, DSPWave_GrayCode, DSPWave_GrayCodeDec)

        TestNameInput = Report_TName_From_Instance(CalcC, "X", "GrayCode", CInt(i))
        TheExec.Flow.TestLimit resultVal:=DSPWave_GrayCodeDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    

End Function

Public Function CMRR(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim site As Variant
    Dim CMRR_Value As New SiteDouble
    Dim Voltage_Value As Double
    
    'Voltage_Value = theexec.Specs.DC.item(argv(1)).CurrentValue
    TestNameInput = Report_TName_From_Instance(Calc, "X", "CMRR")
    
    For Each site In TheExec.sites
        'CMRR_Value = GetStoredCaptureData(argv(0))
        Voltage_Value = TheExec.specs.DC.Item(argv(1)).CurrentValue(site)
        CMRR_Value = GetStoredData(argv(0) + "_para")
        
        CMRR_Value = CMRR_Value * 1.25 / (2 ^ 17)
        
        OutputTname_format = Split(TestNameInput, "_")
        OutputTname_format(6) = "CMRR"
        OutputTname_format(7) = CStr(GetStoredData(argv(0) + "_para"))
        OutputTname_format(8) = Replace(CStr(TheExec.specs.DC.Item(argv(1)).CurrentValue(site)), ".", "p")
        TestNameInput = Merge_TName(OutputTname_format)
        CMRR_Value = CMRR_Value / Voltage_Value
        
    Next
    
    TheExec.Flow.TestLimit resultVal:=CMRR_Value, Tname:=TestNameInput, ForceResults:=tlForceFlow

''    For i = 0 To argc - 1
''        For Each Site In TheExec.sites
''            DSPWave_Dict = GetStoredCaptureData(argv(i))
''            TheExec.Flow.TestLimit resultVal:=DSPWave_Dict.ConvertStreamTo(tldspParallel, 21, 0, Bit0IsMsb).Multiply(50000).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
''        Next
''    Next i
    

End Function

Public Function PSRR(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_GrayCode As New DSPWave
    Dim DSPWave_GrayCodeDec As New DSPWave
    Dim TestNameInput As String
    Dim TestNameInput1 As String
    Dim OutputTname_format() As String
    Dim site As Variant
    Dim PSRR_Value As New SiteDouble
    Dim Voltage_Value As Double
    
    'Voltage_Value = theexec.Specs.DC.item(argv(1)).CurrentValue
    TestNameInput = Report_TName_From_Instance(Calc, "X", "PSRR")
    
    For Each site In TheExec.sites
        'CMRR_Value = GetStoredCaptureData(argv(0))
        Voltage_Value = TheExec.specs.DC.Item(argv(1)).CurrentValue(site)
        PSRR_Value = GetStoredData(argv(0) + "_para")
        
        PSRR_Value = PSRR_Value * 1.25 / (2 ^ 17)
        
        OutputTname_format = Split(TestNameInput, "_")
        OutputTname_format(6) = "PSRR"
        OutputTname_format(7) = CStr(GetStoredData(argv(0) + "_para"))
        OutputTname_format(8) = Replace(CStr(TheExec.specs.DC.Item(argv(1)).CurrentValue(site)), ".", "p")
        TestNameInput = Merge_TName(OutputTname_format)
        OutputTname_format(6) = "VDDIO12_MTR_GR"
        TestNameInput1 = Merge_TName(OutputTname_format)
        'PSRR_Value = PSRR_Value / Voltage_Value
        PSRR_Value = PSRR_Value.Power(-1).Multiply(0.2).Log10.Multiply(20)
    Next
    
    TheExec.Flow.TestLimit resultVal:=Voltage_Value, Tname:=TestNameInput1, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=PSRR_Value, Tname:=TestNameInput, ForceResults:=tlForceFlow

''    For i = 0 To argc - 1
''        For Each Site In TheExec.sites
''            DSPWave_Dict = GetStoredCaptureData(argv(i))
''            TheExec.Flow.TestLimit resultVal:=DSPWave_Dict.ConvertStreamTo(tldspParallel, 21, 0, Bit0IsMsb).Multiply(50000).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
''        Next
''    Next i
    

End Function
Public Function Calc_RXEye(argc As Integer, argv() As String) As Long
    '--- arg list ---
    ' 0:    StepSize
    ' 1:    mdll0_lsw,
    ' 2:    mdll0_msw,
    ' 3:    ddr0_dqs0_sw0,
    ' 4:    ddr0_dqs0_sw1,
    ' 5:    mdll1_lsw,
    ' 6:    mdll1_msw,
    ' 7:    ddr0_dqs1_sw0,
    ' 8:    ddr0_dqs1_sw1


    
    Dim InputKey As String
    Dim Step_Size As Integer
    
    Dim site As Variant
    Dim i As Integer
    
    Dim LSW_dspwave As New DSPWave
    Dim MSW_dspwave As New DSPWave
    Dim Combined_dspwave As New DSPWave
    Dim DecValueDspwave As New DSPWave
    
    Dim mdll_8x8 As New DSPWave
    
    
    Dim LSW_SampleSize As Integer
    Dim MSW_SampleSize As Integer
    Dim SampleSize As Integer
    
    
    '/* ------------------------------ */
    Dim mdll0 As New SiteDouble
    Dim mdll1 As New SiteDouble
    
    Dim dqs0rx_sweep As New SiteLong
    Dim dqs1rx_sweep As New SiteLong
    
    Dim ReportVal As New SiteDouble
    Dim LoVal As Double
    Dim TestNameInput As String
    Dim MaxContinuousOne As New SiteLong
    '/* ------------------------------ */
    
    
    Step_Size = Val(argv(0))
    
    
    DecValueDspwave.CreateConstant 0, 1, DspDouble
    

    '/*** --------------------------------------------- ***/
    '/*** ------------------- MDLL0 ------------------- ***/
    '/*** --------------------------------------------- ***/
    
    InputKey = LCase(argv(1))
    LSW_dspwave = GetStoredCaptureData(InputKey)
    InputKey = LCase(argv(2))
    MSW_dspwave = GetStoredCaptureData(InputKey)
    
    For Each site In TheExec.sites
        LSW_SampleSize = LSW_dspwave.SampleSize
        MSW_SampleSize = MSW_dspwave.SampleSize
        SampleSize = LSW_SampleSize + MSW_SampleSize
        Exit For
    Next site
    
'    Call rundsp.CombineDSPWave(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave)
'
'    '/* ------------------ update on 2017/09/20 ------------------ */
'
'    '/* --- separate 64 bits data to 8 x 8 bits --- */
'    Call rundsp.ConvertToLongAndSerialToParrel(Combined_dspwave, 8, mdll_8x8)
'
'
    '/* ----- update on 2018//04/17 make one rundsp of " CombineDSPWave and ConvertToLongAndSerialToParrel "--------*/
    Call rundsp.CombineDSPWave_and_ConvertToLongAndSerialToParrel(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave, 8, mdll_8x8)
    
    '/* --- Calculate average of  8 x 8 bits --- */
    For Each site In TheExec.sites
        mdll0 = mdll_8x8(site).CalcMean
    Next site
    
    '/* ------------------ update on 2017/09/20 ------------------ */
    
    
    
    InputKey = LCase(argv(3))
    LSW_dspwave = GetStoredCaptureData(InputKey)
    InputKey = LCase(argv(4))
    MSW_dspwave = GetStoredCaptureData(InputKey)
    ''SampleSize = LSW_SampleSize + MSW_SampleSize
    
    Call rundsp.CombineDSPWave(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave)
    
    '/*** --------------------------------------------- ***/
    dqs0rx_sweep = 0
    MaxContinuousOne = 0
    For Each site In TheExec.sites
        For i = 0 To SampleSize - 1
            If Combined_dspwave(site).Element(i) = 1 Then
                dqs0rx_sweep = dqs0rx_sweep + 1
            Else
                '/*** Count the number of the first continuous '1' ***/
                'If dqs0rx_sweep > 0 Then
                '    Exit For
                'End If
                
                '/*** Count the number of the Max continuous '1' ***/
                If dqs0rx_sweep > MaxContinuousOne Then
                    MaxContinuousOne = dqs0rx_sweep
                    dqs0rx_sweep = 0
                End If
            End If
        Next i
        '/*** if the Combined_dspwave.Element(END) = 1 ***/
        If dqs0rx_sweep < MaxContinuousOne Then
                dqs0rx_sweep = MaxContinuousOne
        End If
        
        
    Next site
    
    'TheExec.Flow.TestLimit resultVal:=dqs0rx_sweep, Tname:="Number_of_First_Continuous_One_DQS0RX", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=dqs0rx_sweep, Tname:="Number_of_Max_Continuous_One_DQS0RX", ForceResults:=tlForceNone
    
    'dqs0rx_sweep * step_size > mdll0 / 2
    
    ReportVal = dqs0rx_sweep.Multiply(Step_Size)
    
    TheExec.Datalog.WriteComment " *** DQS0RX_Sweep x Step_Size ( " & Step_Size & " ) ***"
    
    For Each site In TheExec.sites
        LoVal = mdll0
        
        If ReportVal = 0 Then ReportVal = -1        ' update by Kaino on 2017/09/20
        
        'Report_TestLimit_by_CZ_Format resultVal:=ReportVal, lowVal:=Str(LoVal), MeasType:="C", UserVar5:="EYEDQS0", scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "EYEDQS0", 0, , , , , tlForceNone)
        TheExec.Flow.TestLimit resultVal:=ReportVal, lowVal:=Str(LoVal), Tname:=TestNameInput, ForceResults:=tlForceNone
    Next site
    
    
    
    '/*** --------------------------------------------- ***/
    '/*** ------------------- MDLL1 ------------------- ***/
    '/*** --------------------------------------------- ***/
    
    InputKey = LCase(argv(5))
    LSW_dspwave = GetStoredCaptureData(InputKey)
    InputKey = LCase(argv(6))
    MSW_dspwave = GetStoredCaptureData(InputKey)
    
   'SampleSize = LSW_SampleSize + MSW_SampleSize
    
'    Call rundsp.CombineDSPWave(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave)
'    '/* ------------------ update on 2017/09/20 ------------------ */
'
'    '/* --- separate 64 bits data to 8 x 8 bits --- */
'    Call rundsp.ConvertToLongAndSerialToParrel(Combined_dspwave, 8, mdll_8x8)
    
    
    '/* ----- update on 2018//04/17 make one rundsp of " CombineDSPWave and ConvertToLongAndSerialToParrel "--------*/
    Call rundsp.CombineDSPWave_and_ConvertToLongAndSerialToParrel(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave, 8, mdll_8x8)
    
    
    '/* --- Calculate average of  8 x 8 bits --- */
    For Each site In TheExec.sites
        mdll1 = mdll_8x8(site).CalcMean
    Next site
    
    '/* ------------------ update on 2017/09/20 ------------------ */
    
    
    InputKey = LCase(argv(7))
    LSW_dspwave = GetStoredCaptureData(InputKey)
    InputKey = LCase(argv(8))
    MSW_dspwave = GetStoredCaptureData(InputKey)
    
    ''SampleSize = LSW_SampleSize + MSW_SampleSize
    
    Call rundsp.CombineDSPWave(LSW_dspwave, MSW_dspwave, LSW_SampleSize, MSW_SampleSize, Combined_dspwave)
    
    dqs1rx_sweep = 0
    MaxContinuousOne = 0
    For Each site In TheExec.sites
        For i = 0 To SampleSize - 1
            If Combined_dspwave(site).Element(i) = 1 Then
                dqs1rx_sweep = dqs1rx_sweep + 1
            Else
                '/*** Count the number of the first continuous '1' ***/
                'If dqs1rx_sweep > 0 Then
                '    Exit For
                'End If
                
                '/*** Count the number of the Max continuous '1' ***/
                If dqs1rx_sweep > MaxContinuousOne Then
                    MaxContinuousOne = dqs1rx_sweep
                    dqs1rx_sweep = 0
                End If
            End If
        Next i
        
        '/*** if the Combined_dspwave.Element(END) = 1 ***/
        If dqs1rx_sweep < MaxContinuousOne Then
                dqs1rx_sweep = MaxContinuousOne
        End If
        
    Next site
    
    'TheExec.Flow.TestLimit resultVal:=dqs1rx_sweep, Tname:="Number_of_First_Continuous_One_DQS1RX", ForceResults:=tlForceNone
    TheExec.Flow.TestLimit resultVal:=dqs1rx_sweep, Tname:="Number_of_Max_Continuous_One_DQS1RX", ForceResults:=tlForceNone
    
    
    
    'dqs1rx_sweep * step_size > mdll1 / 2
    
    ReportVal = dqs1rx_sweep.Multiply(Step_Size)
    
    TheExec.Datalog.WriteComment " *** DQS1RX_Sweep x Step_Size ( " & Step_Size & " ) ***"
    
    For Each site In TheExec.sites
        LoVal = mdll1
        
        If ReportVal = 0 Then ReportVal = -1        ' update by Kaino on 2017/09/20
        
        'Report_TestLimit_by_CZ_Format resultVal:=ReportVal, lowVal:=Str(LoVal), MeasType:="C", UserVar5:="EYEDQS1", scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "EYEDQS1", 0, , , , , tlForceNone)
        TheExec.Flow.TestLimit resultVal:=ReportVal, lowVal:=Str(LoVal), Tname:=TestNameInput, ForceResults:=tlForceNone
   
    Next site
    
End Function

Public Function Calc_R_Path_Cal(argc As Integer, argv() As String) As Long
    
    '' NOTE :
    '' argv(0) : I1 Dictionary KeyName
    '' argv(1) : I2 Dictionary KeyName
    '' argv(2) : I3 Dictionary KeyName
    '' argv(3) : Force Condition Equation Ex: VDDQL/2 => Evaluate(=_VDDQL_VAR_H/2)

    Dim site As Variant
    Dim Meas_I1_PLD As New PinListData: Meas_I1_PLD = GetStoredMeasurement(argv(0))
    Dim Meas_I2_PLD As New PinListData: Meas_I2_PLD = GetStoredMeasurement(argv(1))
    Dim Meas_I3_PLD As New PinListData: Meas_I3_PLD = GetStoredMeasurement(argv(2))
    Dim Force_Cond_Eq As String: Force_Cond_Eq = argv(3)
    Dim Cust_Str As String
    
    If UBound(argv) = 4 Then
        Cust_Str = argv(4)
    End If

    Dim R_Contact_PLD As New PinListData
    Dim RAK_Val() As Double
    Dim Total_RAK_Val As Double
    Dim Split_Name() As String
    Dim Force_Cond As Double
    Dim AddPin_Flag As Integer: AddPin_Flag = 1                                         'Flag for Pinlist global variable add pin
    Dim PinName As Variant
    Dim PinName_Glb_PLD As Variant
    Dim ForceCond_str As String
    Dim Ary_str(0) As String
    Dim DDR_R1 As New PinListData   ' DDR Test R1 Variable
    Dim DDR_R2 As New PinListData   ' DDR Test R2 Variable
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    On Error GoTo errHandler
    
    'Force_Cond_Eq = VDDQL_DDR0/2 ; Split_Name(0) is PowerName
    Split_Name = Split(Force_Cond_Eq, "/")
    ForceCond_str = "_" + Split_Name(0) + "_VAR" + "/" + Split_Name(1)
    Ary_str(0) = ForceCond_str
    Call HIP_Evaluate_ForceVal(Ary_str)
    Force_Cond = CDbl(Ary_str(0))
    
    For Each PinName In Meas_I1_PLD.Pins
        
        'If PinName is not exist then add new one to Global PinListData
        For Each PinName_Glb_PLD In R_Path_PLD.Pins
            If LCase(PinName_Glb_PLD) = LCase(PinName) Then
                'AddPin_Flag = 0 : no need to add pin since pin has already available
                AddPin_Flag = 0
                Exit For
            End If
        Next PinName_Glb_PLD
        
        If AddPin_Flag = 1 Then
            R_Path_PLD.AddPin (CStr(PinName))
        End If
        
        'Add PinName to Local PinListData
        R_Contact_PLD.AddPin (CStr(PinName))
        
        'Customize String for DDR Test only
        If Cust_Str = UCase("DDR_TEST") Then
            DDR_R1.AddPin (CStr(PinName))
            DDR_R2.AddPin (CStr(PinName))
        End If
        
        For Each site In TheExec.sites.Active
            
            Dim I1 As Double: I1 = Meas_I1_PLD.Pins(PinName).Value(site)
            Dim I2 As Double: I2 = Meas_I2_PLD.Pins(PinName).Value(site)
            Dim I3 As Double: I3 = Meas_I3_PLD.Pins(PinName).Value(site)
            Dim I3_I1 As Double: I3_I1 = Meas_I3_PLD.Pins(PinName).Value(site) - Meas_I1_PLD.Pins(PinName).Value(site)
            Dim I3_I2 As Double: I3_I2 = Meas_I3_PLD.Pins(PinName).Value(site) - Meas_I2_PLD.Pins(PinName).Value(site)
            
            'Initialize the Value on PinListData to prevent any cross usage between samples
            R_Path_PLD.Pins(PinName).Value(site) = 0
            R_Contact_PLD.Pins(PinName).Value(site) = 0
            
            'RAK_Val = TheHdw.PPMU.ReadRakValuesByPinnames(PinName, site)
       
''            If InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
''                Total_RAK_Val = RAK_Val(0) + FT_Card_RAK.Pins(PinName).Value(Site)
''            Else
''                Total_RAK_Val = RAK_Val(0) + CP_Card_RAK.Pins(PinName).Value(Site)
''            End If
            Total_RAK_Val = CurrentJob_Card_RAK.Pins(PinName).Value(site)
           
            If I1 <> 0 And I2 <> 0 And I3 <> 0 Then
                If I3_I1 > 0 And I3_I2 > 0 And (I1 * I2) > 0 Then
                    R_Path_PLD.Pins(PinName).Value(site) = (Force_Cond / I3) * (1 - ((I3_I1 * I3_I2) / (I1 * I2)) ^ 0.5)
                Else
                    R_Path_PLD.Pins(PinName).Value(site) = 999 ' report R= 999 when divide by 0
                    TheExec.Datalog.WriteComment (" Error : PinName " & CStr(PinName) & " , Site" & CStr(site) & " I3 should greater than I1,I2 And (I1*I2) should greater than 0!  ")
                End If
            Else
                R_Path_PLD.Pins(PinName).Value(site) = 999
                TheExec.Datalog.WriteComment (" Error : PinName " & CStr(PinName) & " , Site" & CStr(site) & " Division by Zero !   ")
            End If
            
            R_Contact_PLD.Pins(PinName).Value(site) = R_Path_PLD.Pins(PinName).Value(site) - Total_RAK_Val
                        
            'Customize String for DDR Test only
            If Cust_Str = UCase("DDR_TEST") Then
                If I1 <> 0 And I2 <> 0 Then
                    DDR_R1.Pins(PinName).Value(site) = (1 * Force_Cond / I1) - R_Path_PLD.Pins(PinName).Value(site)
                    DDR_R2.Pins(PinName).Value(site) = (1 * Force_Cond / I2) - R_Path_PLD.Pins(PinName).Value(site)
                Else
                    DDR_R1.Pins(PinName).Value(site) = 999
                    DDR_R2.Pins(PinName).Value(site) = 999
                End If
            End If
                        
        Next site
        
    Next PinName
    
    Dim Temp
    
    Temp = TheExec.Flow.TestLimitIndex
   
    For Each PinName In R_Contact_PLD.Pins
            TheExec.Flow.TestLimitIndex = Temp
            TestNameInput = Report_TName_From_Instance(CalcR, CStr(PinName), , 0)
            TheExec.Flow.TestLimit resultVal:=R_Contact_PLD.Pins(PinName), Unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
            'TheExec.Flow.TestLimit resultVal:=R_Contact_PLD.Pins(PinName), lowval:=0, hival:=5, Unit:=unitCustom, customUnit:="ohm", TName:=TestNameInput, ForceResults:=tlForceNone
    Next PinName
    
    If Cust_Str = UCase("DDR_TEST") Then
    
        Temp = TheExec.Flow.TestLimitIndex
        For Each PinName In DDR_R1.Pins
            TheExec.Flow.TestLimitIndex = Temp
            TestNameInput = Report_TName_From_Instance(CalcR, CStr(PinName))
            TheExec.Flow.TestLimit resultVal:=DDR_R1.Pins(PinName), Unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next PinName
        
        Temp = TheExec.Flow.TestLimitIndex
        For Each PinName In DDR_R2.Pins
            TheExec.Flow.TestLimitIndex = Temp
            TestNameInput = Report_TName_From_Instance(CalcR, CStr(PinName))
            TheExec.Flow.TestLimit resultVal:=DDR_R2.Pins(PinName), Unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next PinName
    End If
   
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_R_Path_Cal"
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function Calc_ConcatenateDSP(argc As Integer, argv() As String) As Long
    Dim site As Variant
    Dim i, j As Long
    Dim DSPWave_First As New DSPWave
    Dim DSPWave_Second As New DSPWave
    Dim DSPWave_Combine() As New DSPWave
    Dim TestNameInput As String
    Dim SplitByAt() As String
    Dim First_StartElement As Long
    Dim First_EndElement As Long
    Dim Second_StartElement As Long
    Dim Second_EndElement As Long
    
    Dim DictKey_DSPWave_Combine As String
    
    Dim DataString_First As String
    Dim DataString_Second As String
    Dim DataString_Combine As String
   ' Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    
    ReDim DSPWave_Combine(argc - 1) As New DSPWave
    Dim DSPWave_Combine_Dec As New DSPWave
    
    For i = 0 To argc - 1
        TestNameInput = "ConcatenateDSP_"
        SplitByAt = Split(argv(i), "@")
        DSPWave_First = GetStoredCaptureData(SplitByAt(0))
        First_StartElement = SplitByAt(1)
        First_EndElement = SplitByAt(2)
        DSPWave_Second = GetStoredCaptureData(SplitByAt(3))
        Second_StartElement = SplitByAt(4)
        Second_EndElement = SplitByAt(5)

        Call rundsp.ConcatenateDSP(DSPWave_First, First_StartElement, First_EndElement, DSPWave_Second, Second_StartElement, Second_EndElement, DSPWave_Combine(i))

        ''20170718 - Store Concatenate DSP to Dict.
        If UBound(SplitByAt) = 6 Then
            DictKey_DSPWave_Combine = SplitByAt(6)
            Call AddStoredCaptureData(DictKey_DSPWave_Combine, DSPWave_Combine(i))
        End If
        
        For Each site In TheExec.sites
            DataString_First = ""
            DataString_Second = ""
            DataString_Combine = ""
            For j = 0 To DSPWave_First.SampleSize - 1
                DataString_First = DataString_First & DSPWave_First(site).Element(j)
            Next j
            For j = 0 To DSPWave_Second.SampleSize - 1
                DataString_Second = DataString_Second & DSPWave_Second(site).Element(j)
            Next j
            For j = 0 To DSPWave_Combine(i).SampleSize - 1
                DataString_Combine = DataString_Combine & DSPWave_Combine(i)(site).Element(j)
            Next j
            
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Dictionary " & SplitByAt(0) & " Output Bits = " & DataString_First & " Extract Bits [" & First_StartElement & "-" & First_EndElement & "]" & _
                                                           " ,Dictionary " & SplitByAt(3) & " Output Bits = " & DataString_Second & " Extract Bits [" & Second_StartElement & "-" & Second_EndElement & "]" & _
                                                           " ,Dictionary " & DictKey_DSPWave_Combine & " Output Bits = " & DataString_Combine)
        Next site
        Call rundsp.BinToDec(DSPWave_Combine(i), DSPWave_Combine_Dec)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "ConcatenateDSP", 0)
        
        TheExec.Flow.TestLimit resultVal:=DSPWave_Combine_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
End Function

Public Function Calc_AverageDSP(argc As Integer, argv() As String) As Long

    Dim Val_SerialDSP_1 As New DSPWave
    Dim Val_SerialDSP_2 As New DSPWave
'''    Dim Val_ParallelDSP_1 As New DSPWave
'''    Dim Val_ParallelDSP_2 As New DSPWave
    Dim Temp As New SiteDouble
    Dim outWave As New DSPWave
    
'''    Dim SampleSize1 As Long
'''    Dim SampleSize2 As Long
'''    Dim Site As Variant

'    Val_SerialDSP_1.CreateConstant 0, 11, DspLong
'    Val_SerialDSP_2.CreateConstant 0, 11, DspLong
'    Val_ParallelDSP_1.CreateConstant 0, 1, DspLong
'    Val_ParallelDSP_2.CreateConstant 0, 1, DspLong
    
    Val_SerialDSP_1 = GetStoredCaptureData(argv(0))
    Val_SerialDSP_2 = GetStoredCaptureData(argv(1))
    
'''    For Each Site In TheExec.sites
'''        SampleSize1 = Val_SerialDSP_1(Site).SampleSize
'''        SampleSize2 = Val_SerialDSP_2(Site).SampleSize
'''        Exit For
'''    Next Site
'    For Each Site In TheExec.sites
'        SampleSize1 = Val_SerialDSP_1.SampleSize
'        SampleSize2 = Val_SerialDSP_2.SampleSize
'    Next Site

'''    Call rundsp.ConvertToLongAndSerialToParrel(Val_SerialDSP_1, SampleSize1, Val_ParallelDSP_1)
'''    Call rundsp.ConvertToLongAndSerialToParrel(Val_SerialDSP_2, SampleSize2, Val_ParallelDSP_2)
'''    Call rundsp.DSP_Add(Val_ParallelDSP_1, Val_ParallelDSP_2)
    Call rundsp.Calc_Average_DSP_Porcedure(Val_SerialDSP_1, Val_SerialDSP_2, outWave, Temp)
    
'    Temp = Val_ParallelDSP_1.Element(0)
'    Temp = Temp.Divide(2)
        
    'Report_TestLimit_by_CZ_Format resultVal:=Temp, ForceResults:=tlForceFlow, MeasType:="C"
    Dim TestNameInput As String
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0)
    TheExec.Flow.TestLimit resultVal:=Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow

End Function

Public Function Calc_BitwiseDSP(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim HEX_Str As String
    Dim bin_str As String
    Dim BitWidth As Long
    Dim DSP_Fixed_Bin As New DSPWave
    Dim DictKey As String
    Dim OperationKeyWord As String
    Dim DSP_DictKey As New DSPWave
    Dim DSP_ProcessOutput_BIN As New DSPWave
    Dim DSP_ProcessOutput_DEC As New DSPWave
    
    Dim Dict_Str As String
    Dim Fixed_Str As String
    Dim ProcessOutput_Str As String
    Dim testName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
''    Dim SV_BitWidth As New SiteLong
''    SV_BitWidth = 32
    For i = 0 To argc - 1
        SplitByAt = Split(argv(i), "@")
        HEX_Str = SplitByAt(2)
        BitWidth = Len(SplitByAt(2)) * 4
        bin_str = HEX_to_BIN(HEX_Str)
        
        Set DSP_Fixed_Bin = Nothing
        DSP_Fixed_Bin.CreateConstant 0, BitWidth, DspLong
        
        For Each site In TheExec.sites
            For j = 0 To DSP_Fixed_Bin.SampleSize - 1
                'DSP_Fixed_Bin(Site).Element(j) = Mid(bin_str, i + 1, 1)
                DSP_Fixed_Bin(site).Element(j) = Mid(bin_str, DSP_Fixed_Bin.SampleSize - j, 1)          'ZB correct  for Cyprus AMP dqpi, capi binary bits LSM-->MSB re-order  - 20170905
            Next j
        Next site
        
        DictKey = SplitByAt(0)
        
        OperationKeyWord = UCase(SplitByAt(1))
        
        DSP_DictKey = GetStoredCaptureData(DictKey)
        
        Select Case OperationKeyWord
            Case "OR"
                Call rundsp.DSP_BitWiseOr(DSP_DictKey, DSP_Fixed_Bin, BitWidth, DSP_ProcessOutput_BIN)
            Case "AND"
                Call rundsp.DSP_BitWiseAnd(DSP_DictKey, DSP_Fixed_Bin, BitWidth, DSP_ProcessOutput_BIN)
            Case "XOR"
                Call rundsp.DSP_BitWiseXOR(DSP_DictKey, DSP_Fixed_Bin, BitWidth, DSP_ProcessOutput_BIN)
            Case Else
        End Select
        
        For Each site In TheExec.sites
            Dict_Str = ""
            Fixed_Str = ""
            ProcessOutput_Str = ""
            For j = 0 To DSP_DictKey.SampleSize - 1
                Dict_Str = Dict_Str & DSP_DictKey(site).Element(j)
            Next j
            For j = 0 To DSP_Fixed_Bin.SampleSize - 1
                Fixed_Str = Fixed_Str & DSP_Fixed_Bin(site).Element(j)
            Next j
            For j = 0 To DSP_ProcessOutput_BIN.SampleSize - 1
                ProcessOutput_Str = ProcessOutput_Str & DSP_ProcessOutput_BIN(site).Element(j)
            Next j
        
           If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Dictionary Output Bits = " & Dict_Str & "[" & SplitByAt(0) & "]" & vbCrLf & _
                                                                                     "Hex Val                       = " & Fixed_Str & "[" & SplitByAt(1) & " " & SplitByAt(2) & "]" & vbCrLf & _
                                                                                     "Process Result                = " & ProcessOutput_Str)
        Next site

        Set DSP_ProcessOutput_DEC = Nothing
        DSP_ProcessOutput_DEC.CreateConstant 0, 1, DspDouble
        
        testName = OperationKeyWord & "_" & i
        
        Call rundsp.BinToDec(DSP_ProcessOutput_BIN, DSP_ProcessOutput_DEC)
        
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
                
        TheExec.Flow.TestLimit resultVal:=DSP_ProcessOutput_DEC.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i

End Function

'markchen

'ADC Calculate final efuse trim code after 85C trimming
'CDNS   => REFERENCE_CTRL_DIG = round(0.25*REFERENCE_CTRL_DIG_25 + 0.75*REFERENCE_CTRL_DIG_85)
'Sicily => ADC0_VREF_85C = round(0.25*ADC0_VREF_25C + 0.75*ADC0_VREF_85C_IM)

Public Function Calc_Dict_Store(argc As Integer, argv() As String) As Long

'Dim Dict_Store_DIG_25C As New DSPWave

 
'Dict_Store_DIG_25C = argv(0)
'Dim Dict_Store_DIG_85C As New DSPWave
'Dict_Store_DIG_85C = argv(1)


'Dict_Store_DIG_25C As String, Dict_Store_DIG_85C As String

Dim DSPWave_Dict_DIG_25C As New DSPWave
Dim DSPWave_Dict_DIG_85C As New DSPWave
Dim ADC_Trim_Code_DIG_25C As New DSPWave
Dim ADC_Trim_Code_DIG_85C As New DSPWave
Dim ADC_Trim_Code_DIG_sum As New DSPWave
Dim ADC_Trim_Code_DIG_final As New DSPWave
Dim eFuse_CTRL_DIG As New DSPWave

Dim Fuse_REFERENCE_CTRL_DIG_Name As String: Fuse_REFERENCE_CTRL_DIG_Name = argv(2)

Dim site As Variant


ADC_Trim_Code_DIG_25C.CreateConstant 0, 1, DspLong
ADC_Trim_Code_DIG_85C.CreateConstant 0, 1, DspLong
ADC_Trim_Code_DIG_sum.CreateConstant 0, 1, DspLong

DSPWave_Dict_DIG_25C = GetStoredCaptureData(argv(0))
DSPWave_Dict_DIG_85C = GetStoredCaptureData(argv(1))


Call HardIP_Bin2Dec(ADC_Trim_Code_DIG_25C, DSPWave_Dict_DIG_25C)
Call HardIP_Bin2Dec(ADC_Trim_Code_DIG_85C, DSPWave_Dict_DIG_85C)

For Each site In TheExec.sites.Active
    ADC_Trim_Code_DIG_sum(site).Element(0) = FormatNumber(ADC_Trim_Code_DIG_25C(site).Element(0) * 0.25 + ADC_Trim_Code_DIG_85C(site).Element(0) * 0.75, 0)
'        Call HardIP_Dec2Bin(ADC_Trim_Code_DIG_final, ADC_Trim_Code_DIG_sum, 8)
        
        If InStr(UCase(argv(0)), UCase("ADC0")) <> 0 Then
            TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_25C :" & ADC_Trim_Code_DIG_25C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_85C :" & ADC_Trim_Code_DIG_85C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_sum :" & ADC_Trim_Code_DIG_sum(site).Element(0)
            
         ElseIf InStr(UCase(argv(0)), UCase("ADC1")) <> 0 Then
            TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_25C :" & ADC_Trim_Code_DIG_25C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_85C :" & ADC_Trim_Code_DIG_85C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_sum :" & ADC_Trim_Code_DIG_sum(site).Element(0)
        
         ElseIf InStr(UCase(argv(0)), UCase("ADC2")) <> 0 Then
            TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_25C :" & ADC_Trim_Code_DIG_25C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_85C :" & ADC_Trim_Code_DIG_85C(site).Element(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_sum :" & ADC_Trim_Code_DIG_sum(site).Element(0)
        
        End If
    
Next site
Call HardIP_Dec2Bin(ADC_Trim_Code_DIG_final, ADC_Trim_Code_DIG_sum, 8)

' Dim Data_Temp As String
Dim final_Bin2_Str1(7) As String
Dim final_Bin2_Str As String
Dim efuse_REFERENCE_CTRL_DIG_Str1(7) As String
Dim efuse_REFERENCE_CTRL_DIG_Str As String
Dim i As Integer
For Each site In TheExec.sites.Active
        For i = 0 To 7
           ' Data_Temp = Data_Temp & (ADC_Trim_Code_DIG_final(site).Element(i))
             final_Bin2_Str1(i) = CStr(ADC_Trim_Code_DIG_final(site).Element(i))
                                             
        Next i
        final_Bin2_Str = Join(final_Bin2_Str1, "")
        
        If InStr(UCase(argv(0)), UCase("ADC0")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_final :" & final_Bin2_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC1")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_final :" & final_Bin2_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC2")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_final :" & final_Bin2_Str
        End If
          
        final_Bin2_Str = ""
       ' Data_Temp = ""
Next site

Call AddStoredCaptureData(Fuse_REFERENCE_CTRL_DIG_Name, ADC_Trim_Code_DIG_final)
TheExec.Datalog.WriteComment ("DigCap data store in dictionary " & "<<" & Fuse_REFERENCE_CTRL_DIG_Name & ">>")

eFuse_CTRL_DIG = GetStoredCaptureData(Fuse_REFERENCE_CTRL_DIG_Name)

For Each site In TheExec.sites.Active
        For i = 0 To 7
          efuse_REFERENCE_CTRL_DIG_Str1(i) = CStr(eFuse_CTRL_DIG(site).Element(i))
                                             
        Next i
        efuse_REFERENCE_CTRL_DIG_Str = Join(efuse_REFERENCE_CTRL_DIG_Str1, "")
        
        If InStr(UCase(argv(0)), UCase("ADC0")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " Fuse ADC0_VREF_85C :" & efuse_REFERENCE_CTRL_DIG_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC1")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " Fuse ADC1_VREF_85C :" & efuse_REFERENCE_CTRL_DIG_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC2")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " Fuse ADC2_VREF_85C :" & efuse_REFERENCE_CTRL_DIG_Str
        End If
          
        final_Bin2_Str = ""
       ' Data_Temp = ""
Next site


'
'For Each site In theexec.sites.Active
'    theexec.Datalog.WriteComment "site " & site & " ADC_Trim_Code_DIG_final :" & Data_Temp
''    theexec.Datalog.WriteComment "site " & site & " ADC_Trim_Code_DIG_final :" & ADC_Trim_Code_DIG_final(site).Element(0) & ADC_Trim_Code_DIG_final(site).Element(1) _
''    & ADC_Trim_Code_DIG_final(site).Element(2) & ADC_Trim_Code_DIG_final(site).Element(3) & ADC_Trim_Code_DIG_final(site).Element(4) _
''    & ADC_Trim_Code_DIG_final(site).Element(5) & ADC_Trim_Code_DIG_final(site).Element(6) & ADC_Trim_Code_DIG_final(site).Element(7)
'Next site


End Function

Public Function Calc_2S_Complement_DSP(argc As Integer, argv() As String) As Long
    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey As String
    Dim DictKey_2S_DEC As String
    
    Dim DSP_DictKey_BIN As New DSPWave
    Dim DSP_DictKey_DEC As New DSPWave
    
    Dim DSPWave_2S_Complement() As New DSPWave
    ReDim DSPWave_2S_Complement(argc - 1) As New DSPWave
    
    Dim testName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    Dim SL_BitWidth As New SiteLong
''    Call rundsp.WordWidthExample
    
    For i = 0 To argc - 1
        If InStr(argv(i), "@") <> 0 Then
            SplitByAt = Split(argv(i), "@")
            DictKey = SplitByAt(0)
            DictKey_2S_DEC = SplitByAt(1)
            testName = SplitByAt(2)
        Else
            DictKey = argv(i)
        End If
        
        DSP_DictKey_BIN = GetStoredCaptureData(DictKey)
        Set DSP_DictKey_DEC = Nothing
        DSP_DictKey_DEC.CreateConstant 0, 1, DspDouble
        'Call rundsp.BinToDec(DSP_DictKey_BIN, DSP_DictKey_DEC)
        
        For Each site In TheExec.sites
        
        
                    ''===================== BinToDec =====================
            DSP_DictKey_BIN(site) = DSP_DictKey_BIN(site).ConvertDataTypeTo(DspLong)
            DSP_DictKey_DEC(site) = DSP_DictKey_BIN(site).ConvertStreamTo(tldspParallel, DSP_DictKey_BIN(site).SampleSize, 0, Bit0IsMsb)
            ''===================== BinToDec (End) =====================
        
        
            SL_BitWidth(site) = DSP_DictKey_BIN(site).SampleSize
''            DSP_DictKey_DEC(0).Element(0) = 255
''            DSP_DictKey_DEC(1).Element(0) = 254
        Next site
        
        Set DSPWave_2S_Complement(i) = Nothing
        DSPWave_2S_Complement(i).CreateConstant 0, 1, DspDouble
        
        'Call rundsp.DSP_Convert_2S_Complement(DSP_DictKey_DEC, SL_BitWidth, DSPWave_2S_Complement(i))
        
        
        
                 ''===================== Convert_2S_Complement =====================
        For Each site In TheExec.sites
            DSPWave_2S_Complement(i)(site) = DSP_DictKey_DEC(site).ConvertDataTypeTo(DspLong)
            DSPWave_2S_Complement(i)(site).WordWidth = SL_BitWidth(site)
            DSPWave_2S_Complement(i)(site) = DSPWave_2S_Complement(i)(site).ConvertDataTypeTo(DspLong)
'            Debug.Print DSPWave_2S_Complement(i)(Site).Element(0)
        Next site
        ''===================== Convert_2S_Complement  (End) =====================
        
        
        
        
        If InStr(argv(i), "@") <> 0 Then
            Call AddStoredCaptureData(DictKey_2S_DEC, DSPWave_2S_Complement(i))
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", "DEC" & CStr(i), CInt(i))
            
            TheExec.Flow.TestLimit resultVal:=DSP_DictKey_DEC.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            
            Call Update_BC_PassFail_Flag
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
            
            TheExec.Flow.TestLimit resultVal:=DSPWave_2S_Complement(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            
            Call Update_BC_PassFail_Flag
        Else
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
            
            TheExec.Flow.TestLimit resultVal:=DSPWave_2S_Complement(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            
            Call Update_BC_PassFail_Flag
        End If
    Next i
End Function


Public Function MSBNEGATE(argc As Integer, argv() As String) As Long
    
    Dim temp_dsp As New DSPWave
    Dim i As Long
    Dim site As Variant
    
    For i = 0 To argc - 1
        temp_dsp = GetStoredCaptureData(argv(i))
        For Each site In TheExec.sites
            If temp_dsp.Element(temp_dsp.SampleSize - 1) = 0 Then
                temp_dsp.Element(temp_dsp.SampleSize - 1) = 1
            ElseIf temp_dsp.Element(temp_dsp.SampleSize - 1) = 1 Then
                temp_dsp.Element(temp_dsp.SampleSize - 1) = 0
            End If
            Call AddStoredCaptureData(argv(i), temp_dsp)
        Next site
    Next

End Function

'
Public Function Calc_TMPS_Coeff(argc As Integer, argv() As String) As Long

    Dim site As Variant
    
    Dim Coeff_A0_Sensor1 As New DSPWave, Coeff_A1_Sensor1 As New DSPWave, Coeff_A2_Sensor1 As New DSPWave, Coeff_A3_Sensor1 As New DSPWave, Coeff_A4_Sensor1 As New DSPWave
    Dim Coeff_A0_Sensor1_Dict As New DSPWave, Coeff_A1_Sensor1_Dict As New DSPWave, Coeff_A2_Sensor1_Dict As New DSPWave, Coeff_A3_Sensor1_Dict As New DSPWave, Coeff_A4_Sensor1_Dict As New DSPWave
    Dim DataOut_85C_Sensor1 As New DSPWave, DataOut_25C_Sensor1 As New DSPWave, DSPWave_Dict As New DSPWave
    
    Coeff_A0_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A1_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A2_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A3_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A4_Sensor1.CreateConstant 0, 1, DspLong
    DataOut_25C_Sensor1.CreateConstant 0, 1, DspLong
    DataOut_85C_Sensor1.CreateConstant 0, 1, DspLong
    
    On Error GoTo errHandler
    
    If TheExec.TesterMode = testModeOffline Then
        Set DataOut_25C_Sensor1 = Nothing
        DataOut_25C_Sensor1.CreateConstant 0, 4
    Else
        'DataOut_25C_Sensor1 = GetStoredCaptureData(argv(0))
        Call HardIP_Bin2Dec(DataOut_25C_Sensor1, GetStoredCaptureData(argv(0))) ' for Turks
    End If
    
    Call HardIP_Bin2Dec(DataOut_85C_Sensor1, GetStoredCaptureData(argv(1)))

    Call TMPS_Coeff_Calculation(Coeff_A0_Sensor1, Coeff_A1_Sensor1, Coeff_A2_Sensor1, Coeff_A3_Sensor1, Coeff_A4_Sensor1, DataOut_85C_Sensor1, DataOut_25C_Sensor1)

    Call HardIP_Dec2Bin(Coeff_A0_Sensor1_Dict, Coeff_A0_Sensor1, 15)
    Call HardIP_Dec2Bin(Coeff_A1_Sensor1_Dict, Coeff_A1_Sensor1, 14)
    Call HardIP_Dec2Bin(Coeff_A2_Sensor1_Dict, Coeff_A2_Sensor1, 12)
    Call HardIP_Dec2Bin(Coeff_A3_Sensor1_Dict, Coeff_A3_Sensor1, 10)
    Call HardIP_Dec2Bin(Coeff_A4_Sensor1_Dict, Coeff_A4_Sensor1, 11)

    Call AddStoredCaptureData(argv(2), Coeff_A0_Sensor1_Dict)
    Call AddStoredCaptureData(argv(3), Coeff_A1_Sensor1_Dict)
    Call AddStoredCaptureData(argv(4), Coeff_A2_Sensor1_Dict)
    Call AddStoredCaptureData(argv(5), Coeff_A3_Sensor1_Dict)
    Call AddStoredCaptureData(argv(6), Coeff_A4_Sensor1_Dict)
        
    Exit Function
errHandler:
        TheExec.Datalog.WriteComment "TMPS Calc Temp VBT function is error "
        TheExec.Datalog.WriteComment ("Error #: " & Str(err.number) & " " & err.Description)
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_TMPS_Coeff_1point(argc As Integer, argv() As String) As Long

    Dim site As Variant
    
    Dim Coeff_A0_Sensor1 As New DSPWave, Coeff_A1_Sensor1 As New DSPWave, Coeff_A2_Sensor1 As New DSPWave, Coeff_A3_Sensor1 As New DSPWave, Coeff_A4_Sensor1 As New DSPWave
    Dim Coeff_A0_Sensor1_Dict As New DSPWave, Coeff_A1_Sensor1_Dict As New DSPWave, Coeff_A2_Sensor1_Dict As New DSPWave, Coeff_A3_Sensor1_Dict As New DSPWave, Coeff_A4_Sensor1_Dict As New DSPWave
    Dim DataOut_85C_Sensor1 As New DSPWave, DataOut_25C_Sensor1 As New DSPWave, DSPWave_Dict As New DSPWave
    
    Coeff_A0_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A1_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A2_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A3_Sensor1.CreateConstant 0, 1, DspLong
    Coeff_A4_Sensor1.CreateConstant 0, 1, DspLong
    DataOut_25C_Sensor1.CreateConstant 0, 1, DspLong
    DataOut_85C_Sensor1.CreateConstant 0, 1, DspLong
    
    On Error GoTo errHandler
    
   '' DataOut_25C_Sensor1 = GetStoredCaptureData(argv(0))

    Call HardIP_Bin2Dec(DataOut_25C_Sensor1, GetStoredCaptureData(argv(0)))

    Call TMPS_Coeff_Calculation_1point(Coeff_A0_Sensor1, Coeff_A1_Sensor1, Coeff_A2_Sensor1, Coeff_A3_Sensor1, Coeff_A4_Sensor1, DataOut_25C_Sensor1)

    Call HardIP_Dec2Bin(Coeff_A0_Sensor1_Dict, Coeff_A0_Sensor1, 15)
    Call HardIP_Dec2Bin(Coeff_A1_Sensor1_Dict, Coeff_A1_Sensor1, 14)
    Call HardIP_Dec2Bin(Coeff_A2_Sensor1_Dict, Coeff_A2_Sensor1, 12)
    Call HardIP_Dec2Bin(Coeff_A3_Sensor1_Dict, Coeff_A3_Sensor1, 10)
    Call HardIP_Dec2Bin(Coeff_A4_Sensor1_Dict, Coeff_A4_Sensor1, 11)

    Call AddStoredCaptureData(argv(1), Coeff_A0_Sensor1_Dict)
    Call AddStoredCaptureData(argv(2), Coeff_A1_Sensor1_Dict)
    Call AddStoredCaptureData(argv(3), Coeff_A2_Sensor1_Dict)
    Call AddStoredCaptureData(argv(4), Coeff_A3_Sensor1_Dict)
    Call AddStoredCaptureData(argv(5), Coeff_A4_Sensor1_Dict)
        
    Exit Function
errHandler:
        TheExec.Datalog.WriteComment "TMPS Calc Temp VBT function is error "
        TheExec.Datalog.WriteComment ("Error #: " & Str(err.number) & " " & err.Description)
        If AbortTest Then Exit Function Else Resume Next
End Function



Public Function ADDRIO_TrimCodeAverage(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim DSPWave_Bin() As New DSPWave
    Dim DSPWave_Dec() As New DSPWave
    ReDim DSPWave_Bin(argc - 2) As New DSPWave
    ReDim DSPWave_Dec(argc - 2) As New DSPWave
    Dim DSPWave_AverageDec As New DSPWave
    DSPWave_AverageDec.CreateConstant 0, 1
    For i = 0 To argc - 2
        DSPWave_Bin(i) = GetStoredCaptureData(argv(i))
        Call rundsp.BinToDec(DSPWave_Bin(i), DSPWave_Dec(i))
        Call rundsp.DSP_Add(DSPWave_AverageDec, DSPWave_Dec(i))
    Next i
    Call rundsp.DSP_DivideConstant(DSPWave_AverageDec, argc - 1)
''    Call rundsp.DSP_ConvertDataTypeToLong(DSPWave_AverageDec)
    For Each site In TheExec.sites
        ''20170210-Rounding
        DSPWave_AverageDec(site).Element(0) = Int(DSPWave_AverageDec(site).Element(0) + 0.5)
    Next site

    Call AddStoredCaptureData(argv(argc - 1), DSPWave_AverageDec)
    TheExec.Flow.TestLimit resultVal:=DSPWave_AverageDec.Element(0), Tname:="ADDRIO_AverageTrimCode", ForceResults:=tlForceNone
End Function
Public Function Calc_MDLL_Monotonicity(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
''    Call CreateSimulateMDLL_Data(argc, argv)
    
    Dim DSPWaveBin() As New DSPWave
    ReDim DSPWaveBin(argc - 1) As New DSPWave
    Dim DSPWaveDec() As New DSPWave
    ReDim DSPWaveDec(argc - 1) As New DSPWave
    Dim testName As String
    testName = argv(0) & "_"
    For i = 1 To argc - 1
        DSPWaveBin(i) = GetStoredCaptureData(argv(i))
        Call rundsp.BinToDec(DSPWaveBin(i), DSPWaveDec(i))
    Next i
    Dim dataStr As String
    For Each site In TheExec.sites
        dataStr = ""
        For i = 1 To argc - 1
            If i = 1 Then
                dataStr = argv(i) & " = " & DSPWaveDec(i)(site).Element(0) & ", "
            Else
                dataStr = dataStr & argv(i) & " = " & DSPWaveDec(i)(site).Element(0) & ", "
            End If
        Next i
       If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " " & dataStr)
    Next site
    
    Dim MDLL_CurrentVal As New SiteLong
    Dim MDLL_PreviousVal  As New SiteLong
    Dim b_MDLL_DecreaseDirection As New SiteBoolean
    Dim b_MDLL_DecreaseAddIndex As New SiteBoolean
    Dim MDLL_DecreaseResultPass As New SiteLong
    Dim b_MDLL_TestResultFail As New SiteBoolean
    Dim MDLL_Index As New SiteLong
    b_MDLL_DecreaseDirection = False
    
    MDLL_DecreaseResultPass = 1
    b_MDLL_TestResultFail = False
    MDLL_Index = 1
    Dim StepSize As Long
    For Each site In TheExec.sites
        For i = 1 To argc - 1
            If i = 1 Then
                MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
            Else
                MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                b_MDLL_DecreaseDirection(site) = MDLL_CurrentVal.Subtract(MDLL_PreviousVal).compare(LessThanOrEqualTo, 0)
                
                If b_MDLL_DecreaseDirection(site) = False Then
                    MDLL_DecreaseResultPass(site) = 0
''                    b_MDLL_TestResultFail(Site) = True
                    Exit For
                End If
                
                b_MDLL_DecreaseAddIndex(site) = MDLL_CurrentVal.Subtract(MDLL_PreviousVal).compare(LessThan, 0)
                
                If b_MDLL_DecreaseAddIndex(site) = True Then
                    MDLL_Index(site) = MDLL_Index(site) + 1
                End If
''                If MDLL_Index(Site) > 1 Then
''''                    b_MDLL_TestResultFail(Site) = True
''                    Exit For
''                End If
                
                MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
            End If
        Next i
    Next site
    

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "MDLLDecrease", 0)
    
    TheExec.Flow.TestLimit resultVal:=MDLL_DecreaseResultPass, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceNone
    For Each site In TheExec.sites
        If MDLL_DecreaseResultPass.bitwiseand(1) Then
        Else
            MDLL_Index(site) = -99
        End If
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "MDLLUnique", 1)
    TheExec.Flow.TestLimit resultVal:=MDLL_Index, lowVal:=1, hiVal:=2, Tname:=TestNameInput, ForceResults:=tlForceNone
End Function
Public Function Calc_DDR_MDCC_Freq(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim SplitByAt() As String
    Dim F_Case As String
    Dim DictKey As String
    Dim Dict_DSP_DEC() As New DSPWave
    ReDim Dict_DSP_DEC(argc - 1) As New DSPWave
    Dim Dict_DSP_BINARY() As New DSPWave
    ReDim Dict_DSP_BINARY(argc - 1) As New DSPWave
    Dim site As Variant
    Dim Calc_DSP_DEC() As New DSPWave
    ReDim Calc_DSP_DEC(argc - 1) As New DSPWave
    Dim testName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    For i = 0 To argc - 1
        Calc_DSP_DEC(i).CreateConstant 0, 1, DspDouble
        SplitByAt = Split(argv(i), "@")
        F_Case = UCase(SplitByAt(0))
        DictKey = SplitByAt(1)
        Dict_DSP_BINARY(i) = GetStoredCaptureData(DictKey)
        testName = SplitByAt(2)
        
'        Call rundsp.BinToDec(Dict_DSP_BINARY(i), Dict_DSP_DEC(i))

        For Each site In TheExec.sites
            ''===================== BinToDec =====================
            Dict_DSP_BINARY(i)(site) = Dict_DSP_BINARY(i)(site).ConvertDataTypeTo(DspLong)
            Dict_DSP_DEC(i)(site) = Dict_DSP_BINARY(i)(site).ConvertStreamTo(tldspParallel, Dict_DSP_BINARY(i)(site).SampleSize, 0, Bit0IsMsb)
            ''===================== BinToDec (End) =====================
        Next site


        
        For Each site In TheExec.sites
            Select Case F_Case
                Case "F0"
                    Calc_DSP_DEC(i)(site).Element(0) = ((Dict_DSP_DEC(i)(site).Element(0)) / (114 * 2)) * 2.133 * 1000000000#
                Case "F1"
                    Calc_DSP_DEC(i)(site).Element(0) = (Dict_DSP_DEC(i)(site).Element(0)) / (100 * 2) * 1.466 * 1000000000#
                Case "F2"
                    Calc_DSP_DEC(i)(site).Element(0) = (Dict_DSP_DEC(i)(site).Element(0)) / (90 * 2) * 0.712 * 1000000000#
                Case "M9_F1"
                    Calc_DSP_DEC(i)(site).Element(0) = (Dict_DSP_DEC(i)(site).Element(0)) / (100 * 2) * 1.2 * 1000000000#
                Case Else
            
            End Select
        Next site
''        TheExec.Flow.TestLimit resultVal:=Calc_DSP_DEC(i).Element(0), Tname:="Calc_DDR_MDCC_Freq", unit:=unitHz, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance(CalcF, "X", , CInt(i))
        TheExec.Flow.TestLimit resultVal:=Calc_DSP_DEC(i).Element(0), Tname:=TestNameInput, Unit:=unitHz, ForceResults:=tlForceFlow
    Next i
End Function

Public Function Calc_MDLL_Monotonicity_DevideBlock_SEG(argc As Integer, argv() As String) As Long
   Dim i As Long
   Dim site As Variant
   Dim DSP_Captured() As New DSPWave
   ReDim DSP_Captured((argc - 2))
    
   Dim DSP_Arry_Bin() As New DSPWave
   ReDim DSP_Arry_Bin((argc - 2) * 4 - 1)
    
   Dim DSP_Arry_Dec() As New DSPWave
   ReDim DSP_Arry_Dec((argc - 2) * 4 - 1)
    
   Dim Uni_DLL_Indicator As New SiteLong
   Dim Max_Dec_Val As New SiteLong
    
   Dim TestNameInput As String
   Dim OutputTname_format() As String
    
   Dim testName As String
    
   For i = 0 To argc - 2
      DSP_Captured(i) = GetStoredCaptureData(argv(i + 1))
   Next i
    
''   '--- For Offline Simulation ------
''   If TheExec.TesterMode = testModeOffline Then
''      For Each Site In TheExec.sites.Active
''         'code1
''         DSP_Captured(0).Element(0) = 1
''         DSP_Captured(0).Element(1) = 1
''         DSP_Captured(0).Element(2) = 1
''         DSP_Captured(0).Element(3) = 1
''         DSP_Captured(0).Element(4) = 1
''         DSP_Captured(0).Element(5) = 1
''         DSP_Captured(0).Element(6) = 0
''         DSP_Captured(0).Element(7) = 0
''
''         'code6
''         DSP_Captured(0).Element(8) = 1
''         DSP_Captured(0).Element(9) = 1
''         DSP_Captured(0).Element(10) = 1
''         DSP_Captured(0).Element(11) = 1
''         DSP_Captured(0).Element(12) = 1
''         DSP_Captured(0).Element(13) = 1
''         DSP_Captured(0).Element(14) = 0
''         DSP_Captured(0).Element(15) = 0
''
''         'code0
''         DSP_Captured(0).Element(16) = 1
''         DSP_Captured(0).Element(17) = 1
''         DSP_Captured(0).Element(18) = 1
''         DSP_Captured(0).Element(19) = 1
''         DSP_Captured(0).Element(20) = 1
''         DSP_Captured(0).Element(21) = 1
''         DSP_Captured(0).Element(22) = 0
''         DSP_Captured(0).Element(23) = 0
''
''         'code4
''         DSP_Captured(0).Element(24) = 1
''         DSP_Captured(0).Element(25) = 1
''         DSP_Captured(0).Element(26) = 1
''         DSP_Captured(0).Element(27) = 1
''         DSP_Captured(0).Element(28) = 1
''         DSP_Captured(0).Element(29) = 1
''         DSP_Captured(0).Element(30) = 0
''         DSP_Captured(0).Element(31) = 0
''
''         'code5
''         DSP_Captured(1).Element(0) = 1
''         DSP_Captured(1).Element(1) = 1
''         DSP_Captured(1).Element(2) = 1
''         DSP_Captured(1).Element(3) = 1
''         DSP_Captured(1).Element(4) = 1
''         DSP_Captured(1).Element(5) = 1
''         DSP_Captured(1).Element(6) = 0
''         DSP_Captured(1).Element(7) = 0
''
''         'code2
''         DSP_Captured(1).Element(8) = 1
''         DSP_Captured(1).Element(9) = 1
''         DSP_Captured(1).Element(10) = 1
''         DSP_Captured(1).Element(11) = 1
''         DSP_Captured(1).Element(12) = 1
''         DSP_Captured(1).Element(13) = 1
''         DSP_Captured(1).Element(14) = 0
''         DSP_Captured(1).Element(15) = 0
''
''         'code7
''         DSP_Captured(1).Element(16) = 1
''         DSP_Captured(1).Element(17) = 1
''         DSP_Captured(1).Element(18) = 1
''         DSP_Captured(1).Element(19) = 1
''         DSP_Captured(1).Element(20) = 1
''         DSP_Captured(1).Element(21) = 1
''         DSP_Captured(1).Element(22) = 0
''         DSP_Captured(1).Element(23) = 0
''
''         'code3
''         DSP_Captured(1).Element(24) = 1
''         DSP_Captured(1).Element(25) = 1
''         DSP_Captured(1).Element(26) = 1
''         DSP_Captured(1).Element(27) = 1
''         DSP_Captured(1).Element(28) = 1
''         DSP_Captured(1).Element(29) = 1
''         DSP_Captured(1).Element(30) = 0
''         DSP_Captured(1).Element(31) = 0
''
''      Next Site
    
      '       Dim j, k As Long
      '       For j = 0 To 1
      '           For k = 0 To 3
      '               DSP_Captured(j).Element(k * 8) = 1
      '
      '           Next k
      '       Next j
       
'   End If
    
    
    
   For Each site In TheExec.sites.Active
        'cfgh_cadll_sts_mdll_code_grp1_w210,{3'b000, oct2[26:18], 1'b0, oct1[17:9],   1'b0, oct0[8:0]};
        'cfgh_cadll_sts_mdll_code_grp2_w543,{3'b000, oct5[53:45], 1'b0, oct4[44:36], 1'b0, oct3[35:27]};
        'cfgh_cadll_sts_mdll_code_grp3_w76,{13'b0,     oct7[71:63], 1'b0, oct6[62:54]};
        '
        'for the 32 bits of w210 DigCap, it maps to : {3'b0, oct2 values, 1'b0, oct1 values, 1'b0, oct0 values}.
        'for the 32 bits of w543 DigCap, it maps to : {3'b0, oct5 values, 1'b0, oct4 values, 1'b0, oct3 values}.
        'for the 32 bits of w76 DigCap, it maps to : {13'b0, oct7 values, 1'b0, oct6 values}.
        
        
'      DSP_Arry_Bin(4) = DSP_Captured(0).Select(0, 1, DSP_Captured(0).SampleSize / 4).Copy
'      DSP_Arry_Bin(0) = DSP_Captured(0).Select(DSP_Captured(0).SampleSize / 4, 1, DSP_Captured(0).SampleSize / 4).Copy
'      DSP_Arry_Bin(6) = DSP_Captured(0).Select(DSP_Captured(0).SampleSize / 2, 1, DSP_Captured(0).SampleSize / 4).Copy
'      DSP_Arry_Bin(1) = DSP_Captured(0).Select((DSP_Captured(0).SampleSize / 4) * 3, 1, DSP_Captured(0).SampleSize / 4).Copy
'      DSP_Arry_Bin(3) = DSP_Captured(1).Select(0, 1, DSP_Captured(1).SampleSize / 4).Copy
'      DSP_Arry_Bin(7) = DSP_Captured(1).Select(DSP_Captured(1).SampleSize / 4, 1, DSP_Captured(1).SampleSize / 4).Copy
'      DSP_Arry_Bin(2) = DSP_Captured(1).Select(DSP_Captured(1).SampleSize / 2, 1, DSP_Captured(1).SampleSize / 4).Copy
'      DSP_Arry_Bin(5) = DSP_Captured(1).Select((DSP_Captured(1).SampleSize / 4) * 3, 1, DSP_Captured(1).SampleSize / 4).Copy

      
      DSP_Arry_Bin(4) = DSP_Captured(0).Select(0, 1, 9).Copy
      DSP_Arry_Bin(0) = DSP_Captured(0).Select(10, 1, 9).Copy
      DSP_Arry_Bin(6) = DSP_Captured(0).Select(20, 1, 9).Copy
      
      DSP_Arry_Bin(1) = DSP_Captured(1).Select(0, 1, 9).Copy
      DSP_Arry_Bin(3) = DSP_Captured(1).Select(10, 1, 9).Copy
      DSP_Arry_Bin(7) = DSP_Captured(1).Select(20, 1, 9).Copy
      
      DSP_Arry_Bin(2) = DSP_Captured(2).Select(0, 1, 9).Copy
      DSP_Arry_Bin(5) = DSP_Captured(2).Select(10, 1, 9).Copy
      
      
   
      For i = 0 To UBound(DSP_Arry_Bin)
         DSP_Arry_Bin(i) = DSP_Arry_Bin(i).ConvertDataTypeTo(DspLong)
         DSP_Arry_Dec(i) = DSP_Arry_Bin(i).ConvertStreamTo(tldspParallel, DSP_Arry_Bin(i).SampleSize, 0, Bit0IsMsb)
         'nope, only for debugging purpose
         'Report_TestLimit_by_CZ_Format resultVal:=DSP_Arry_Dec(i).Element(0), ForceResults:=tlForceNone, UserVar6:="DSP_Arry_Dec" & i, UserVar5:=argv(0), MeasType:="C"
      Next i
      
      Uni_DLL_Indicator(site) = 1
      For i = 0 To UBound(DSP_Arry_Bin) - 1
         If Uni_DLL_Indicator(site) = 1 Then
            If DSP_Arry_Dec(i).Element(0) = DSP_Arry_Dec(i + 1).Element(0) Then
            ElseIf DSP_Arry_Dec(i).Element(0) = DSP_Arry_Dec(i + 1).Element(0) + 1 Then Uni_DLL_Indicator(site) = 2
            Else
               Uni_DLL_Indicator(site) = -2
               Exit For
            End If
         ElseIf Uni_DLL_Indicator(site) = 2 Then
            If DSP_Arry_Dec(i).Element(0) = DSP_Arry_Dec(i + 1).Element(0) Then
            Else
               Uni_DLL_Indicator(site) = -1
               Exit For
            End If
         End If
      Next i
      Max_Dec_Val(site) = DSP_Arry_Dec(0).Element(0)
   Next site
    
   Call GetFlowTName
    
   If gl_UseStandardTestName_Flag = True Then                     'Roger add
      Call Report_ALG_TName_From_Instance(OutputTname_format, "C", CStr(argv(0)) & "Max_Diff", gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex))
      TestNameInput = Merge_TName(OutputTname_format)
            
   Else
      TestNameInput = testName & "Max_Diff"
   End If
    
   TheExec.Flow.TestLimit resultVal:=Max_Dec_Val, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
   If gl_UseStandardTestName_Flag = True Then                     'Roger add
      Call Report_ALG_TName_From_Instance(OutputTname_format, "C", CStr(argv(0)) & "Decrease", gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex))
      TestNameInput = Merge_TName(OutputTname_format)
            
   Else
      TestNameInput = testName & "Decrease"
   End If
    
    
   TheExec.Flow.TestLimit resultVal:=Uni_DLL_Indicator, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    
    
    
   '''    Report_TestLimit_by_CZ_Format resultVal:=Max_Dec_Val, ForceResults:=tlForceFlow, MeasType:="C"
   '''
   '''    Report_TestLimit_by_CZ_Format resultVal:=Uni_DLL_Indicator, lowVal:=1, hiVal:=2, ForceResults:=tlForceFlow, MeasType:="C"
    
End Function
Public Function Calc_MDLL_Monotonicity_Analyze(argc As Integer, argv() As String) As Long
    Dim site As Variant
    Dim i, j, k As Integer
    Dim testName As String
    Dim TestNameInput As String
    Dim Max_Dec_Val As New SiteLong
    Dim DSP_Decimal() As New DSPWave
    Dim DSP_Captured() As New DSPWave
    Dim OutputTname_format() As String
    
    Dim MaxDiffRank As New SiteLong
    Dim DecreaseRank As New SiteLong
    Dim Uni_DLL_Indicator As New SiteLong
    
    ReDim DSP_Decimal((argc - 1))
    ReDim DSP_Captured((argc - 1))
    
    For i = 0 To argc - 1
        DSP_Captured(i) = GetStoredCaptureData(argv(i))
        For Each site In TheExec.sites.Active
            DSP_Decimal(i) = DSP_Captured(i).ConvertStreamTo(tldspParallel, DSP_Captured(i).SampleSize, 0, Bit0IsMsb)
        Next site
    Next i
    

    For Each site In TheExec.sites.Active
        MaxDiffRank(site) = 1
        DecreaseRank(site) = 1
        Uni_DLL_Indicator(site) = 1
        MaxDiffRank(site) = DSP_Decimal(0).Element(0)                                   ' Setting compare base
        For i = 0 To UBound(DSP_Decimal) - 1
            If i <> UBound(DSP_Decimal) - 1 Then
                If DecreaseRank(site) <> 0 Then
                    If DSP_Decimal(i).Element(0) < DSP_Decimal(i + 1).Element(0) Then   ' RuleCheck1:oct0>=oct1>=oct2>=oct3>=oct4>=oct5>=oct6>=oct7
                        DecreaseRank(site) = 0
                    End If
                End If
                If MaxDiffRank(site) > DSP_Decimal(i + 1).Element(0) Then               ' Record Minimum for RuleCheck3
                    MaxDiffRank(site) = DSP_Decimal(i + 1).Element(0)
                End If
            End If
            If Uni_DLL_Indicator(site) = 1 Then                                         ' RuleCheck2:The TypeNum must be less than two type
                If DSP_Decimal(i).Element(0) = DSP_Decimal(i + 1).Element(0) Then
                    Uni_DLL_Indicator(site) = 1
                ElseIf DSP_Decimal(i).Element(0) = DSP_Decimal(i + 1).Element(0) + 1 Then
                    Uni_DLL_Indicator(site) = 2                                         ' When OTC0 > OTC1 +1, then Uni_DLL_Indicator = 2
                Else
                    Uni_DLL_Indicator(site) = -2                                        ' delta(OTC(i) - OTC(i+1) )> 1 , Uni_DLL_Indicator = -2
                End If
            ElseIf Uni_DLL_Indicator(site) = 2 Then
                If DSP_Decimal(i).Element(0) = DSP_Decimal(i + 1).Element(0) Then
                    Uni_DLL_Indicator(site) = 2
                Else
                    Uni_DLL_Indicator(site) = -1                                         ' When Uni_DLL_Indicator = 2 means there have third kind of OTC value
                End If
            End If
        Next i
        MaxDiffRank(site) = DSP_Decimal(0).Element(0) - MaxDiffRank(site)               ' RuleCheck3:Maxmun & Minimum delta must be equal one
    Next site
    
  
    Call GetFlowTName
    If gl_UseStandardTestName_Flag = True Then
        gl_Tname_Alg_Index = CStr(TheExec.Flow.TestLimitIndex)
        TestNameInput = Report_TName_From_Instance("N", "x", Left(argv(0), InStr(1, argv(0), "_")) & "Decrease", CInt(gl_Tname_Alg_Index), , "qq")
        TheExec.Flow.TestLimit resultVal:=DecreaseRank, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
        gl_Tname_Alg_Index = CStr(TheExec.Flow.TestLimitIndex)
        TestNameInput = Report_TName_From_Instance("N", "x", Left(argv(0), InStr(1, argv(0), "_")) & "Unique", CInt(gl_Tname_Alg_Index))
        TheExec.Flow.TestLimit resultVal:=Uni_DLL_Indicator, lowVal:=1, hiVal:=2, Tname:=TestNameInput, ForceResults:=tlForceFlow
        gl_Tname_Alg_Index = CStr(TheExec.Flow.TestLimitIndex)
        TestNameInput = Report_TName_From_Instance("N", "x", Left(argv(0), InStr(1, argv(0), "_")) & "MaxDiff", CInt(gl_Tname_Alg_Index))
        TheExec.Flow.TestLimit resultVal:=MaxDiffRank, lowVal:=0, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    End If
  
   
End Function

Public Function Calc_LPDPTX_FXCode(argc As Integer, argv() As String) As Long
    Dim Dict_FXCode As String
    Dim Dict_Margin_5Bit As String
    Dim Dict_Margin_1Bit As String
    Dim DSP_FXCode_Bin As New DSPWave
    Dim DSP_FXCode_Dec As New DSPWave
    Dim DSP_Margin_5Bit_Dec As New DSPWave
    Dim DSP_Margin_5Bit_Bin As New DSPWave
    Dim DSP_Margin_1Bit_Dec As New DSPWave
    Dim DSP_Margin_1Bit_Bin As New DSPWave
    Dim site As Variant
    '' ----Added to truncate FXcode 20170426---
    Dim Dict_FXCode_5Bit As String
    Dim DSP_FXCode_5Bit_Bin As New DSPWave
    ''----------------------------------------
    
    ''----Added Post_Bin and Pre_Bin Procedure----
    Dim Dict_Post_Bin As String
    Dim Dict_Post_2R As String
    Dim Dict_Pre_Bin As String
    Dim Dict_Pre_2R As String
    Dim DSP_Post_Dec As New DSPWave
    Dim DSP_Post_Bin As New DSPWave
    Dim DSP_Pre_Dec As New DSPWave
    Dim DSP_Pre_Bin As New DSPWave
    Dim DSP_Post_2R_Dec As New DSPWave
    Dim DSP_Post_2R_Bin As New DSPWave
    Dim DSP_Pre_2R_Dec As New DSPWave
    Dim DSP_Pre_2R_Bin As New DSPWave
    ''-----------------------------------------------------------
    Dict_FXCode = argv(0)
    ''Dict_Margin_5Bit = argv(1)
    ''Dict_Margin_1Bit = argv(2)
    Dict_FXCode_5Bit = argv(1)
    Dict_Post_Bin = argv(2)
    Dict_Post_2R = argv(3)
    Dict_Pre_Bin = argv(4)
    Dict_Pre_2R = argv(5)
    
    
    DSP_FXCode_Bin = GetStoredCaptureData(Dict_FXCode)
    Call rundsp.BinToDec(DSP_FXCode_Bin, DSP_FXCode_Dec)
     
'     ''Simulation
'    DSP_FXCode_Dec(0).Element(0) = 12
'    DSP_FXCode_Dec(1).Element(0) = 15
    
    ''Truncate FXCode to 5 bit
    Call rundsp.DSPWaveDecToBinary(DSP_FXCode_Dec, 5, DSP_FXCode_5Bit_Bin)
    Call AddStoredCaptureData(Dict_FXCode_5Bit, DSP_FXCode_5Bit_Bin)
    
 
    DSP_Margin_5Bit_Dec.CreateConstant 0, 1, DspDouble
    DSP_Post_Dec.CreateConstant 0, 1, DspDouble
    DSP_Pre_Dec.CreateConstant 0, 1, DspDouble
    DSP_Post_2R_Dec.CreateConstant 0, 1, DspDouble
    DSP_Pre_2R_Dec.CreateConstant 0, 1, DspDouble
    
    
    DSP_Margin_5Bit_Bin.CreateConstant 0, 5, DspLong
    DSP_Margin_1Bit_Bin.CreateConstant 0, 1, DspLong
    DSP_Post_Bin.CreateConstant 0, 4, DspLong
    DSP_Pre_Bin.CreateConstant 0, 4, DspLong
    DSP_Post_2R_Bin.CreateConstant 0, 1, DspLong
    DSP_Pre_2R_Bin.CreateConstant 0, 1, DspLong
    
    For Each site In TheExec.sites
        DSP_Margin_5Bit_Dec(site).Element(0) = (DSP_FXCode_Dec(site).Element(0) + 18) / 2
        DSP_Margin_5Bit_Dec(site).Element(0) = DSP_Margin_5Bit_Dec(site).Element(0) - DSP_FXCode_Dec(site).Element(0) ''=> Rest of Margin
        
        
        If DSP_Margin_5Bit_Dec(site).Element(0) > 6 Then
           DSP_Post_Dec(site).Element(0) = Fix(DSP_Margin_5Bit_Dec.Element(0)) ''=>Integer of Rest of Margin
           DSP_Pre_Dec(site).Element(0) = 0
           DSP_Pre_2R_Dec(site).Element(0) = 0
           
            If DSP_Margin_5Bit_Dec(site).Element(0) - Int(DSP_Margin_5Bit_Dec(site).Element(0)) = 0 Then
                DSP_Post_2R_Dec.Element(0) = 0
            Else
                DSP_Post_2R_Dec.Element(0) = 1
            End If
        Else
           DSP_Pre_Dec(site).Element(0) = Fix(DSP_Margin_5Bit_Dec.Element(0))
           DSP_Post_Dec(site).Element(0) = 0
           DSP_Post_2R_Dec(site).Element(0) = 0
            
            If DSP_Margin_5Bit_Dec(site).Element(0) - Int(DSP_Margin_5Bit_Dec(site).Element(0)) = 0 Then
                DSP_Pre_2R_Dec.Element(0) = 0
            Else
                DSP_Pre_2R_Dec.Element(0) = 1
            End If
        End If
        
'        If DSP_Margin_5Bit_Dec(Site).Element(0) - Int(DSP_Margin_5Bit_Dec(Site).Element(0)) = 0 Then
'            DSP_Margin_1Bit_Bin.Element(0) = 0
'
'        Else
'            DSP_Margin_1Bit_Bin.Element(0) = 1
'            DSP_Margin_5Bit_Dec.Element(0) = Fix(DSP_Margin_5Bit_Dec.Element(0))
'        End If
    Next site
    
    ''Call AddStoredCaptureData(Dict_Margin_1Bit, DSP_Margin_1Bit_Bin)
    
    For Each site In TheExec.sites
       '' DSP_Margin_5Bit_Dec(Site) = DSP_Margin_5Bit_Dec(Site).ConvertDataTypeTo(DspLong)
        DSP_Post_Dec(site) = DSP_Post_Dec(site).ConvertDataTypeTo(DspLong)
        DSP_Pre_Dec(site) = DSP_Pre_Dec(site).ConvertDataTypeTo(DspLong)
        DSP_Post_2R_Dec(site) = DSP_Post_2R_Dec(site).ConvertDataTypeTo(DspLong)
        DSP_Pre_2R_Dec(site) = DSP_Pre_2R_Dec(site).ConvertDataTypeTo(DspLong)
    Next site
    
    ''Call rundsp.DSPWaveDecToBinary(DSP_Margin_5Bit_Dec, 5, DSP_Margin_5Bit_Bin)
    Call rundsp.DSPWaveDecToBinary(DSP_Post_Dec, 4, DSP_Post_Bin)
    Call rundsp.DSPWaveDecToBinary(DSP_Pre_Dec, 4, DSP_Pre_Bin)
    Call rundsp.DSPWaveDecToBinary(DSP_Post_2R_Dec, 1, DSP_Post_2R_Bin)
    Call rundsp.DSPWaveDecToBinary(DSP_Pre_2R_Dec, 1, DSP_Pre_2R_Bin)


    ''Call AddStoredCaptureData(Dict_Margin_5Bit, DSP_Margin_5Bit_Bin)
    Call AddStoredCaptureData(Dict_Post_Bin, DSP_Post_Bin)
    Call AddStoredCaptureData(Dict_Pre_Bin, DSP_Pre_Bin)
    Call AddStoredCaptureData(Dict_Post_2R, DSP_Post_2R_Bin)
    Call AddStoredCaptureData(Dict_Pre_2R, DSP_Pre_2R_Bin)
End Function

Public Function Calc_ADCPLL_fuse(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim DSPWave_Dict As New DSPWave
    Dim fuse_name As String
    Dim Data_Temp As String
    Dim Fuse_Value As New SiteLong
    Dim Dict_Name As String

''    For i = 0 To argc - 2 Step 2  'arg(0)=DSPWaveA, arg(1)=Fuse_nameA, arg(2)=DSPWaveB, arg(3)=Fuse_nameB......
    Dict_Name = argv(0)
    DSPWave_Dict = GetStoredCaptureData(Dict_Name)
    Data_Temp = ""
    
    For Each site In TheExec.sites
        For j = 0 To (DSPWave_Dict(site).SampleSize - 1)
            Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(j))
        Next j
        Fuse_Value(site) = Bin2Dec_rev(Data_Temp)
        Data_Temp = ""
    Next site

    fuse_name = UCase(argv(1))
    ''Call HIP_eFuse_Write("ECID", fuse_name, Fuse_Value)
    fuse_name = ""
''    Next i

End Function
Public Function Calc_GrayCodeToBin(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim Dict_DSP_Bin() As New DSPWave
    ReDim Dict_DSP_Bin(argc - 1) As New DSPWave
    Dim GrayCode_DSP_Bin() As New DSPWave
    ReDim GrayCode_DSP_Bin(argc - 1) As New DSPWave
    Dim GrayCode_DSP_Dec() As New DSPWave
    ReDim GrayCode_DSP_Dec(argc - 1) As New DSPWave
    Dim b_IsUnSigned As Boolean ''New SiteBoolean
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    b_IsUnSigned = argv(0)
    
    For i = 1 To argc - 1
        Dict_DSP_Bin(i) = GetStoredCaptureData(argv(i))
        'Call rundsp.DSP_GrayCode2Bin(b_IsUnSigned, Dict_DSP_Bin(i), GrayCode_DSP_Bin(i), GrayCode_DSP_Dec(i))
        Call GrayCode2Bin_TTR(b_IsUnSigned, Dict_DSP_Bin(i), GrayCode_DSP_Bin(i), GrayCode_DSP_Dec(i))
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "", CInt(i))
        
        TheExec.Flow.TestLimit resultVal:=GrayCode_DSP_Dec(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
End Function

Public Function CalcDutyDelay(argc As Integer, argv() As String) As Long

    Dim CalcDutyVal() As New PinListData
    ReDim CalcDutyVal(argc - 1) As New PinListData
    Dim DeltaDelayVal() As New PinListData
    ReDim DeltaDelayVal(argc - 1) As New PinListData
    
    Dim i As Long, j As Long, p As Long
    Dim site As Variant
    Dim PinName As String
    Dim b_FirstTime As Boolean
    b_FirstTime = True
    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    
    Dim TestNameInput As String
    Dim Freq_TestName_Input As String
    Dim Voltage_Name() As String
    Voltage_Name = Split(TheExec.DataManager.instanceName, "_")
    Freq_TestName_Input = argv(argc - 1)
    
    Dim MaxNumOfDuty As Long
    Dim StartNumOfDuty As Long
    StartNumOfDuty = 1
    MaxNumOfDuty = 113
    Dim OutputTname_format() As String
    
    For i = StartNumOfDuty To MaxNumOfDuty
        CalcDutyVal(i) = GetStoredMeasurement(argv(i))
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).Pins.Count - 1
                CalcDutyVal(i).Pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
        End If
        For j = 1 To CalcDutyVal(i).Pins.Count - 1
            If InStr(UCase(CalcDutyVal(i).Pins(j)), "_P") <> 0 Then
                PinName = CalcDutyVal(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                    If CalcDutyVal(i).Pins(j).Value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).Pins(j).Value = 1
                    End If
                Next site
            
                CalcDutyVal(i).Pins(j).Value = CalcDutyVal(i).Pins(j).Multiply(2).Invert
                
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        CalcDutyVal(i).Pins(j).Value = -999
    ''                TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & CStr(i - 1) & "_" & TestNameInput, ForceResults:=tlForceNone
                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).Pins(j), "", 0)
                TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone
            End If
        Next j
    Next i
    
    '' 20170228 - Add test method for octal
    Dim Freq_Dll_Str As String
    Freq_Dll_Str = argv(0)
    Dim TCycle_Val As Double
    Dim LSB_Val As Double
    Dim Oct_Ideal_Val As Double
    Select Case UCase(Freq_Dll_Str)
        Case "DDR_F0"
            TCycle_Val = 1 / (2133.3333 * MHz)
        Case "DDR_F1"
            TCycle_Val = 1 / (1466.6667 * MHz)
        Case "DDR_F2"
            TCycle_Val = 1 / (712 * MHz)
        Case "DDR_F1M9"
            TCycle_Val = 1 / (1200 * MHz)
        Case "DDR_F2M9"
            TCycle_Val = 1 / (600 * MHz)
    End Select
    
    LSB_Val = TCycle_Val / 128
    Oct_Ideal_Val = TCycle_Val / 8
    
    Dim OctantIndex As Long
    Dim OctantMaxNum As Long
    OctantIndex = 0
    OctantMaxNum = 7
    Dim Octant_Val() As New PinListData
    ReDim Octant_Val(OctantMaxNum) As New PinListData
    
    For i = StartNumOfDuty To MaxNumOfDuty Step 16
        If OctantIndex = 7 Then
            Octant_Val(OctantIndex) = CalcDutyVal(1).Math.Subtract(CalcDutyVal(i)).Add(TCycle_Val)
        Else
            Octant_Val(OctantIndex) = CalcDutyVal(i + 16).Math.Subtract(CalcDutyVal(i))
        End If
        For j = 1 To Octant_Val(OctantIndex).Pins.Count - 1
             If InStr(UCase(Octant_Val(OctantIndex).Pins(j)), "_P") <> 0 Then
                PinName = Octant_Val(OctantIndex).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        Octant_Val(OctantIndex).Pins(j).Value = -999
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Oct" & CStr(OctantIndex) & "_" & TestNameInput, ForceResults:=tlForceNone
                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, Octant_Val(OctantIndex).Pins(j), , 0)
                TheExec.Flow.TestLimit resultVal:=Octant_Val(OctantIndex).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone
            End If
        Next j
        OctantIndex = OctantIndex + 1
    Next i
    Dim OctPhaseError() As New PinListData
    ReDim OctPhaseError(OctantMaxNum) As New PinListData
    Dim OctPhaseError_Max As New PinListData
    Dim OctPhaseError_Min As New PinListData
    
    For i = 0 To OctantMaxNum
        OctPhaseError(i) = Octant_Val(i).Math.Subtract(Oct_Ideal_Val)
        If i = 0 Then
            OctPhaseError_Max = OctPhaseError(i)
            OctPhaseError_Min = OctPhaseError(i)
        End If
        For j = 1 To OctPhaseError(i).Pins.Count - 1
            If InStr(UCase(OctPhaseError(i).Pins(j)), "_P") <> 0 Then
                PinName = OctPhaseError(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                    If OctPhaseError(i).Pins(j).Value > OctPhaseError_Max.Pins(j).Value Then
                        OctPhaseError_Max.Pins(j).Value = OctPhaseError(i).Pins(j).Value
                    End If
                    If OctPhaseError(i).Pins(j).Value < OctPhaseError_Min.Pins(j).Value Then
                        OctPhaseError_Min.Pins(j).Value = OctPhaseError(i).Pins(j).Value
                    End If
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE" & CStr(i) & "_" & TestNameInput, ForceResults:=tlForceNone
                        OctPhaseError(i).Pins(j).Value = -999
                    End If
''                    Else
''                        TheExec.Flow.TestLimit resultVal:=OctPhaseError(i).Pins(j).Value, ScaleType:=scalePico, Tname:="PE" & CStr(i) & "_" & TestNameInput, ForceResults:=tlForceNone
''                    End If
''                    If i = OctantMaxNum Then
''                        If b_DivideZeroError(Site) = True Then
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceNone
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceNone
''                        Else
''                            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Max.Pins(j).Value, ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceNone
''                            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Min.Pins(j).Value, ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceNone
''                        End If
''                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, OctPhaseError(i).Pins(j), , 0)
                TheExec.Flow.TestLimit resultVal:=OctPhaseError(i).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone
''                If i = OctantMaxNum Then
''                    For Each Site In TheExec.sites
''                        If b_DivideZeroError(Site) = True Then
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceNone
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceNone
''                        End If
''                    Next Site
''                    TheExec.Flow.TestLimit resultVal:=OctPhaseError_Max.Pins(j), ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceNone
''                    TheExec.Flow.TestLimit resultVal:=OctPhaseError_Min.Pins(j), ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceNone
''
''                End If
            End If
        Next j
    Next i

    For j = 1 To OctPhaseError_Max.Pins.Count - 1
        If InStr(UCase(OctPhaseError_Max.Pins(j)), "_P") <> 0 Then
            PinName = OctPhaseError_Max.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            TestNameInput = TestNameInput & "_" & Freq_TestName_Input
            
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
                    OctPhaseError_Max.Pins(j).Value = -999
                    OctPhaseError_Min.Pins(j).Value = -999
                End If
            Next site
            
            TestNameInput = Report_TName_From_Instance(CalcF, OctPhaseError_Max.Pins(j), "", 0)
            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Max.Pins(j), scaletype:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput & "_" & Voltage_Name(UBound(Voltage_Name)), ForceResults:=tlForceNone
            
            TestNameInput = Report_TName_From_Instance(CalcF, OctPhaseError_Min.Pins(j), "", 0)
            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Min.Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone
        End If
    Next j

    
    For i = StartNumOfDuty To MaxNumOfDuty
        If i = 1 Then
        Else

            DeltaDelayVal(i) = CalcDutyVal(i).Math.Subtract(CalcDutyVal(i - 1))
            For j = 1 To DeltaDelayVal(i).Pins.Count - 1
                If InStr(UCase(DeltaDelayVal(i).Pins(j)), "_P") <> 0 Then
                    PinName = DeltaDelayVal(i).Pins(j)
                    TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                    TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                    TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                    
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delta_Delay_" & CStr(i - 2) & "_" & TestNameInput, ForceResults:=tlForceNone
                            DeltaDelayVal(i).Pins(j).Value = -999
                        End If
                    Next site
                    TestNameInput = Report_TName_From_Instance(CalcF, DeltaDelayVal(i).Pins(j), "", 0)
                    TheExec.Flow.TestLimit resultVal:=DeltaDelayVal(i).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone
                End If
            Next j
        End If
    Next i
    

    Dim DNL_Val() As New PinListData
    ReDim DNL_Val(argc - 1) As New PinListData
    Dim AryShiftNum As Long
    AryShiftNum = 2
    Dim b_Linearity_Fail As Boolean
    b_Linearity_Fail = False
    Dim DNL_Val_Max As New PinListData
    Dim DNL_Val_Min As New PinListData
    Dim No_Of_Valid_Delta_Delay As Long
    
    No_Of_Valid_Delta_Delay = 111
    
    ''20170818-Sum of DNL to be INL
    ''20170901
    Dim INL() As New PinListData
    ReDim INL(argc - 1) As New PinListData
    '' Assign pins to INL and initial value to 0
''    INL = DNL_Val(0)
''    INL = 0
    
    For i = 0 + AryShiftNum To No_Of_Valid_Delta_Delay + AryShiftNum
        DNL_Val(i) = DeltaDelayVal(i).Math.Divide(LSB_Val).Subtract(1)
        
        If i = 0 + AryShiftNum Then
            DNL_Val_Max = DNL_Val(i)
            DNL_Val_Min = DNL_Val(i)
            ''20170818 -  initial INL value to 0
        End If
            INL(i) = DNL_Val(i)
            INL(i) = 0
        
        For j = 1 To DeltaDelayVal(i).Pins.Count - 1
            
            If InStr(UCase(DeltaDelayVal(i).Pins(j)), "_P") <> 0 Then
                PinName = DeltaDelayVal(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
    
                   If DNL_Val(i).Pins(j).Value > DNL_Val_Max.Pins(j).Value Then
                       DNL_Val_Max.Pins(j).Value = DNL_Val(i).Pins(j).Value
                   End If
                   If DNL_Val(i).Pins(j).Value < DNL_Val_Min.Pins(j).Value Then
                       DNL_Val_Min.Pins(j).Value = DNL_Val(i).Pins(j).Value
                   End If
                   
                    If b_DivideZeroError(site) = True Then
                        DNL_Val(i).Pins(j).Value = -999
                    End If
                    
                       Select Case UCase(Freq_Dll_Str)
                           Case "DDR_F0"
                               If b_DivideZeroError(site) = True Then
                                    TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), "", 0)
                                    TheExec.Flow.TestLimit resultVal:=-999, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone
                               Else
                                    TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), "", 0)
                                    TheExec.Flow.TestLimit resultVal:=DNL_Val(i).Pins(j).Value, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone
                               End If
                               If DNL_Val(i).Pins(j).Value > 1 Or DNL_Val(i).Pins(j).Value < -1 Then
                                   b_Linearity_Fail = True
                               End If
                           Case "DDR_F1", "DDR_F1M9"
                               If b_DivideZeroError(site) = True Then
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), "", 0)
                                   TheExec.Flow.TestLimit resultVal:=-999, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone
                               Else
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), "", 0)
                                   TheExec.Flow.TestLimit resultVal:=DNL_Val(i).Pins(j).Value, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone
                               End If
                               If DNL_Val(i).Pins(j).Value > 1 Or DNL_Val(i).Pins(j).Value < -1 Then
                                   b_Linearity_Fail = True
                               End If
                           Case "DDR_F2", "DDR_F2M9"
                               If b_DivideZeroError(site) = True Then
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), "", 0)
                                   TheExec.Flow.TestLimit resultVal:=-999, lowVal:=-1, hiVal:=1.5, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone
                               Else
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).Pins(j), "", 0)
                                   TheExec.Flow.TestLimit resultVal:=DNL_Val(i).Pins(j).Value, lowVal:=-1, hiVal:=1.5, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone
                               End If
                               If DNL_Val(i).Pins(j).Value > 1.5 Or DNL_Val(i).Pins(j).Value < -1 Then
                                   b_Linearity_Fail = True
                               End If
                       End Select
                Next site
                
                ''20170818-Sum of DNL to be INL
                 If i = 0 + AryShiftNum Then
                    INL(i).Pins(j) = INL(i).Pins(j).Add(DNL_Val(i).Pins(j))
                Else
                    INL(i).Pins(j) = INL(i).Pins(j).Add(DNL_Val(i).Pins(j)).Add(INL(i - 1).Pins(j))
                End If

                ''20170830 - Bypass
'                TheExec.Flow.TestLimit resultVal:=DNL_Val(i).Pins(j), ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:="DNL" & CStr(i - 2) & "_" & TestNameInput, ForceResults:=tlForceNone
            End If
            
        Next j
    Next i
    
''    If i = No_Of_Valid_Delta_Delay + AryShiftNum Then
    For j = 1 To DNL_Val_Max.Pins.Count - 1
            
        If InStr(UCase(DNL_Val_Max.Pins(j)), "_P") <> 0 Then
            PinName = DNL_Val_Max.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            TestNameInput = TestNameInput & "_" & Freq_TestName_Input
            
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
                    DNL_Val_Max.Pins(j).Value = -999
                    DNL_Val_Min.Pins(j).Value = -999
                End If
            Next site
            

            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Max.Pins(j), "", 0)

            TheExec.Flow.TestLimit resultVal:=DNL_Val_Max.Pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone
            
            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Min.Pins(j), "", 0)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Min.Pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone
        End If
    Next j
''    End If

    Dim INL_Val_Max As New PinListData
    Dim INL_Val_Min As New PinListData
    For i = 0 + AryShiftNum To No_Of_Valid_Delta_Delay + AryShiftNum
        If i = 0 + AryShiftNum Then
            INL_Val_Max = INL(i)
            INL_Val_Min = INL(i)
        End If
       
        For j = 1 To DeltaDelayVal(i).Pins.Count - 1
            If InStr(UCase(DeltaDelayVal(i).Pins(j)), "_P") <> 0 Then
                PinName = DeltaDelayVal(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                   If INL(i).Pins(j).Value > INL_Val_Max.Pins(j).Value Then
                       INL_Val_Max.Pins(j).Value = INL(i).Pins(j).Value
                   End If
                   If INL(i).Pins(j).Value < INL_Val_Min.Pins(j).Value Then
                       INL_Val_Min.Pins(j).Value = INL(i).Pins(j).Value
                   End If
                                     
                    TestNameInput = Report_TName_From_Instance(CalcF, INL_Val_Min.Pins(j), "", 0)
                    TheExec.Flow.TestLimit resultVal:=INL(i).Pins(j).Value, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone
               Next site
             End If
             
        Next j
    Next i
    
    For j = 1 To INL_Val_Max.Pins.Count - 1
            
        If InStr(UCase(INL_Val_Max.Pins(j)), "_P") <> 0 Then
            PinName = INL_Val_Max.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            TestNameInput = TestNameInput & "_" & Freq_TestName_Input
            
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
                    DNL_Val_Max.Pins(j).Value = -999
                    DNL_Val_Min.Pins(j).Value = -999
                End If
            Next site
            
            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Max.Pins(j), "", 0)
                    
            TheExec.Flow.TestLimit resultVal:=INL_Val_Max.Pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone
            
            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Min.Pins(j), "", 0)
                    
            TheExec.Flow.TestLimit resultVal:=INL_Val_Min.Pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, Unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone
        End If
    Next j
    
    '' 20170818 - Test limit for INL
''    Dim INL_Val_Max As New SiteDouble
''    Dim INL_Val_Min As New SiteDouble
''    Dim Counter As Long
''
''    For i = 0 + AryShiftNum To No_Of_Valid_Delta_Delay + AryShiftNum
''        For j = 1 To INL.Pins.Count - 1
''            If InStr(UCase(INL.Pins(j)), "_P") <> 0 Then
''                PinName = INL.Pins(j)
''                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
''                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
''                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
''
''                If Counter = 0 Then
''                    INL_Val_Max = INL.Pins(j)
''                    INL_Val_Min = INL.Pins(j)
''                End If
''
''                For Each Site In TheExec.sites
''                    If INL.Pins(j).Value(Site) > INL_Val_Max(Site) Then
''                        INL_Val_Max(Site) = INL.Pins(j).Value(Site)
''                    End If
''                    If INL.Pins(j).Value(Site) < INL_Val_Min(Site) Then
''                        INL_Val_Min(Site) = INL.Pins(j).Value(Site)
''                    End If
''
''                    If b_DivideZeroError(Site) = True Then
''                        INL.Pins(j).Value = -999
''                        INL.Pins(j).Value = -999
''                    End If
''                Next Site
''
''                TheExec.Flow.TestLimit resultVal:=INL.Pins(j), Tname:="INL" & "_" & TestNameInput, ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone
''                Counter = Counter + 1
''            End If
''        Next j
''    Next i
''    TheExec.Flow.TestLimit resultVal:=INL_Val_Max, Tname:="INL_MAX" & "_" & Freq_TestName_Input, ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone
''    TheExec.Flow.TestLimit resultVal:=INL_Val_Min, Tname:="INL_MIN" & "_" & Freq_TestName_Input, ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone

'''''    For Each Site In TheExec.Sites
'''''         If b_DivideZeroError(Site) = True Then
'''''            TheExec.Flow.TestLimit resultVal:=-999, lowVal:=False, hiVal:=False, Tname:="Linearity_Pass" & "_" & TestNameInput, ForceResults:=tlForceNone
'''''         Else
'''''            TheExec.Flow.TestLimit resultVal:=b_Linearity_Fail, lowVal:=False, hiVal:=False, Tname:="Linearity_Pass" & "_" & TestNameInput, ForceResults:=tlForceNone
'''''        End If
'''''    Next Site
    
End Function

Public Function CalcDelayDelta_Sicily(argc As Integer, argv() As String) As Long

Dim DDR0_MC_DQS_DIFFx_F() As New PinListData
Dim SiteDouble_Frequency() As New SiteDouble
Dim SiteDouble_Delay() As New SiteDouble
Dim SiteDouble_Delta() As New SiteDouble

Dim NumberOfFreq As Long
Dim i, j As Long
Dim site As Variant
Dim TestNameInput As String

NumberOfFreq = CLng(argv(2)) - CLng(argv(1)) + 1

ReDim DDR0_MC_DQS_DIFFx_F(NumberOfFreq - 1) As New PinListData
ReDim SiteDouble_Frequency(NumberOfFreq - 1) As New SiteDouble
ReDim SiteDouble_Delay(NumberOfFreq - 1) As New SiteDouble
ReDim SiteDouble_Delta(NumberOfFreq - 2) As New SiteDouble


For i = 0 To NumberOfFreq - 1
    DDR0_MC_DQS_DIFFx_F(i) = GetStoredMeasurement(argv(0) & i)
    For j = 0 To DDR0_MC_DQS_DIFFx_F(i).Pins.Count - 1
        If InStr(UCase(DDR0_MC_DQS_DIFFx_F(i).Pins(j)), "DQS_P") <> 0 Then
            For Each site In TheExec.sites
                If DDR0_MC_DQS_DIFFx_F(i).Pins(j).Value = 0 Then
                    SiteDouble_Frequency(i) = 0.000000001
                    TheExec.Datalog.WriteComment "Site" & site & " : DDR F" & i & " frequency is 0"
                Else
                    SiteDouble_Frequency(i) = DDR0_MC_DQS_DIFFx_F(i).Pins(j).Value
                End If
            Next site
        End If
    Next j
Next i


For i = 0 To NumberOfFreq - 1
    SiteDouble_Delay(i) = SiteDouble_Frequency(i).Multiply(2).Invert
    TestNameInput = Report_TName_From_Instance(CalcF, "", , 0)
    TheExec.Flow.TestLimit resultVal:=SiteDouble_Delay(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
Next i

For i = 0 To NumberOfFreq - 2
    SiteDouble_Delta(i) = SiteDouble_Delay(i + 1).Subtract(SiteDouble_Delay(i))
    TestNameInput = Report_TName_From_Instance(CalcF, "", , 0)
    TheExec.Flow.TestLimit resultVal:=SiteDouble_Delta(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
Next i


End Function

Public Function CalcDutyDelay_Delta(argc As Integer, argv() As String) As Long

    Dim CalcDutyVal() As New PinListData
    ReDim CalcDutyVal(argc - 1) As New PinListData
    Dim DeltaDelayVal() As New PinListData
    ReDim DeltaDelayVal(argc - 1) As New PinListData
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    Dim i As Long, j As Long
    Dim site As Variant
    Dim PinName As String
    Dim b_FirstTime As Boolean
    b_FirstTime = True
    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    For i = 1 To argc - 1
        CalcDutyVal(i) = GetStoredMeasurement(argv(i))
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).Pins.Count - 1
                CalcDutyVal(i).Pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
        End If
        'For j = 1 To CalcDutyVal(i).Pins.Count - 1 Step 2
        For j = 0 To CalcDutyVal(i).Pins.Count - 1 Step 1           'Modify 20170908
            If j Mod 4 = 2 Or j Mod 4 = 3 Then                                  'Modify 20170908
                
                PinName = CalcDutyVal(i).Pins(j)
                For Each site In TheExec.sites
    
                    If CalcDutyVal(i).Pins(j).Value(site) = 0 Then
                        b_DivideZeroError(site) = True
                       If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).Pins(j).Value = 1
                    End If
                    
                    CalcDutyVal(i).Pins(j).Value = CalcDutyVal(i).Pins(j).Multiply(2).Invert
                    
                    If b_DivideZeroError(site) = True Then
                        TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).Pins(j), "", 0, i)
                        TheExec.Flow.TestLimit resultVal:=-999, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceNone
                    Else
                        TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).Pins(j), "", 0, i)
                        TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j).Value, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceNone
                    End If
                    
                Next site
            End If                                                          'Modify 20170908
        Next j
    Next i
    
    For i = 1 To argc - 1
        If i = 1 Then
        Else

            DeltaDelayVal(i) = CalcDutyVal(i).Math.Subtract(CalcDutyVal(i - 1))
            'For j = 1 To DeltaDelayVal(i).Pins.Count - 1 Step 2
            For j = 0 To DeltaDelayVal(i).Pins.Count - 1 Step 1     'Modify 20170908
                If j Mod 4 = 2 Or j Mod 4 = 3 Then                                  'Modify 20170908
                
                    PinName = DeltaDelayVal(i).Pins(j)
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            TestNameInput = Report_TName_From_Instance(CalcF, DeltaDelayVal(i).Pins(j), "", 0, i)
                            TheExec.Flow.TestLimit resultVal:=-999, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceNone
                        Else
                            TestNameInput = Report_TName_From_Instance(CalcF, DeltaDelayVal(i).Pins(j), "", 0, i)
                            TheExec.Flow.TestLimit resultVal:=DeltaDelayVal(i).Pins(j).Value, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceNone
                        End If
                    Next site
                    
                End If                      'Modify 20170908
            Next j
        End If
    Next i
    
End Function

Public Function CalcDelayJitter(argc As Integer, argv() As String) As Long
    
    Dim CalcDutyVal() As New PinListData
    ReDim CalcDutyVal(argc - 1) As New PinListData
    Dim DeltaDelayVal() As New PinListData
    ReDim DeltaDelayVal(argc - 1) As New PinListData
    
    Dim i As Long, j As Long
    Dim site As Variant

    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    
    Dim TestNameInput As String
    Dim TestNameFromPara As String
    Dim TestNameFreq As String
    Dim OutputTname_format() As String
    
    TestNameFromPara = argv(0)
    TestNameFromPara = LCase(Left(argv(0), 3))
    If InStr(argv(0), "712") Then
        TestNameFreq = LCase(Right(argv(0), 3))
    Else
        TestNameFreq = LCase(Right(argv(0), 4))
    End If
    
    Dim Voltage_Name() As String
    Voltage_Name = Split(TheExec.DataManager.instanceName, "_")
    
    Dim MaxNumOfDuty As Long
    Dim StartNumOfDuty As Long
    StartNumOfDuty = 1
    MaxNumOfDuty = 1
    Dim PinName As String
    For i = StartNumOfDuty To MaxNumOfDuty
        CalcDutyVal(i) = GetStoredMeasurement(argv(i))
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).Pins.Count - 1
                CalcDutyVal(i).Pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
        End If
        For j = 1 To CalcDutyVal(i).Pins.Count - 1
            If InStr(UCase(CalcDutyVal(i).Pins(j)), "_P") <> 0 Then
                PinName = CalcDutyVal(i).Pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                
                For Each site In TheExec.sites
                    If CalcDutyVal(i).Pins(j).Value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).Pins(j).Value = 1
                    End If
                Next site
                    
                CalcDutyVal(i).Pins(j).Value = CalcDutyVal(i).Pins(j).Multiply(2).Invert
                    
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & "_" & TestNameInput, ForceResults:=tlForceNone
                        CalcDutyVal(i).Pins(j).Value = -999
                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).Pins(j), "", 0)
                TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone
            End If
        Next j
    Next i
End Function

Public Function CalcJitter(argc As Integer, argv() As String) As Long
    
    Dim Dict_CalcDutyVal_1 As String
    Dim Dict_CalcDutyVal_2 As String
    Dim CalcDutyVal_1 As New PinListData
    Dim CalcDutyVal_2 As New PinListData
    Dim CalcDuty_Diff As New PinListData
    Dim i As Long, j As Long
    Dim site As Variant

    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    
    Dim TestNameInput As String
    Dim FreqTestName As String
    Dim TestNameFromPara As String
    Dim OutputTname_format() As String
    
    TestNameInput = argv(0)
    TestNameFromPara = LCase(Left(argv(0), 3))
    If InStr(TestNameInput, "712") Then
        FreqTestName = Right(TestNameInput, 3)
    Else
        FreqTestName = Right(TestNameInput, 4)
    End If
    
    Dim Voltage_Name() As String
    Voltage_Name = Split(TheExec.DataManager.instanceName, "_")
    
    Dict_CalcDutyVal_1 = argv(1)
    Dict_CalcDutyVal_2 = argv(2)
    
    CalcDutyVal_1 = GetStoredMeasurement(Dict_CalcDutyVal_1)
    CalcDutyVal_2 = GetStoredMeasurement(Dict_CalcDutyVal_2)
    Dim PinName As String
    
    For j = 1 To CalcDutyVal_1.Pins.Count - 1
        If InStr(UCase(CalcDutyVal_1.Pins(j)), "_P") <> 0 Then
            PinName = CalcDutyVal_1.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            
            For Each site In TheExec.sites
            
                If CalcDutyVal_1.Pins(j).Value(site) = 0 Then
                    b_DivideZeroError(site) = True
                    If gl_Disable_HIP_debug_log = False Then If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                    CalcDutyVal_1.Pins(j).Value = 1
                End If
                If CalcDutyVal_2.Pins(j).Value(site) = 0 Then
                    b_DivideZeroError(site) = True
                    TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                    CalcDutyVal_2.Pins(j).Value = 1
                End If
            Next site
            
            CalcDutyVal_1.Pins(j).Value = CalcDutyVal_1.Pins(j).Multiply(2).Invert
            CalcDutyVal_2.Pins(j).Value = CalcDutyVal_2.Pins(j).Multiply(2).Invert
        End If
    Next j
    
    CalcDuty_Diff = CalcDutyVal_1.Math.Subtract(CalcDutyVal_2)
    
    For j = 1 To CalcDuty_Diff.Pins.Count - 1
        If InStr(UCase(CalcDuty_Diff.Pins(j)), "_P") <> 0 Then
            PinName = CalcDuty_Diff.Pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
''                    TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Jitter" & "_" & TestNameInput, ForceResults:=tlForceNone
                    CalcDuty_Diff.Pins(j).Value = -999
                End If
            Next site

            TestNameInput = Report_TName_From_Instance(CalcF, CalcDuty_Diff.Pins(j), "", 0)
            TheExec.Flow.TestLimit resultVal:=CalcDuty_Diff.Pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone
        End If
    Next j

End Function

Public Function Calc_2S_Complement_To_SignDec(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey_2S_BIN As String
    Dim DictKey_SIGN_DEC As String
    
    Dim DSP_DictKey_2S_BIN As New DSPWave
    Dim DSP_DictKey_SIGN_DEC() As New DSPWave

    ReDim DSP_DictKey_SIGN_DEC(argc - 1) As New DSPWave
    
    Dim testName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    Dim SL_BitWidth As New SiteLong
    '' Format: Dict_2S_Com_A@Dict_SignDec_A@TestName_A,Dict_2S_Com_B@Dict_SignDec_B@TestName_B
    For i = 0 To argc - 1
        SplitByAt = Split(argv(i), "@")
        DictKey_2S_BIN = SplitByAt(0)
        DictKey_SIGN_DEC = SplitByAt(1)
        testName = SplitByAt(2)
        
        DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
''        Set DSP_DictKey_DEC = Nothing
''        DSP_DictKey_DEC.CreateConstant 0, 1, DspDouble
''        Call rundsp.BinToDec(DSP_DictKey_BIN, DSP_DictKey_DEC)
        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize
''            DSP_DictKey_DEC(0).Element(0) = 255
''            DSP_DictKey_DEC(1).Element(0) = 254
        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        Call AddStoredCaptureData(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))
        
''        TheExec.Flow.TestLimit resultVal:=DSP_DictKey_DEC.Element(0), Tname:="DEC_" & i, ForceResults:=tlForceFlow
        
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))

        TheExec.Flow.TestLimit resultVal:=DSP_DictKey_SIGN_DEC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
End Function

Public Function Calc_TMPS_Code2Temperature(argc As Integer, argv() As String) As Long
Dim site As Variant
Dim DataOut_TemperatureCode As New DSPWave
Dim DataOut_Temperature As New SiteDouble
Dim TestNameInput As String
Dim OutputTname_format() As String
Dim Temp_Pass As New SiteLong
'Dim code(165) As Long
'Dim Temperature(165) As Long
'Dim i As Integer
DataOut_TemperatureCode.CreateConstant 0, 1, DspLong

'For i = 0 To 165
'    code(i) = Worksheets("TMPS_Table").Cells(i + 2, 2).Value
'    Temperature(i) = Worksheets("TMPS_Table").Cells(i + 2, 1).Value
'Next i

Call HardIP_Bin2Dec(DataOut_TemperatureCode, GetStoredCaptureData(argv(0)))

For Each site In TheExec.sites
    DataOut_Temperature(site) = 53.2 - 0.08942 * (DataOut_TemperatureCode(site).Element(0) - 2400) - 0.0000142 * (DataOut_TemperatureCode(site).Element(0) - 2400) ^ 2 - 0.00000000231 * (DataOut_TemperatureCode(site).Element(0) - 2400) ^ 3 - 0.000000000000416 * (DataOut_TemperatureCode(site).Element(0) - 2400) ^ 4
Next site

'For Each Site In TheExec.sites
'    If DataOut_TemperatureCode(Site).Element(0) < code(165) Then
'            DataOut_Temperature(Site) = 999
'            GoTo Lable_NextSite
'    ElseIf DataOut_TemperatureCode(Site).Element(0) > code(0) Then
'            DataOut_Temperature(Site) = -999
'            GoTo Lable_NextSite
'    End If
'
'    For i = 0 To 165
'        If DataOut_TemperatureCode(Site).Element(0) < code(i) Then
'            If DataOut_TemperatureCode(Site).Element(0) > code(i + 1) Then
'                DataOut_Temperature(Site) = Temperature(i) + (Temperature(i + 1) - Temperature(i)) * (DataOut_TemperatureCode(Site).Element(0) - code(i)) / (code(i) - code(i + 1))
'                Exit For
'            End If
'        ElseIf DataOut_TemperatureCode(Site).Element(0) = code(i) Then
'                DataOut_Temperature(Site) = Temperature(i)
'                Exit For
'        End If
'    Next i
'Lable_NextSite:
'Next Site
If TheExec.DataManager.instanceName Like "*BV*" Then
    TheExec.Flow.TestLimit resultVal:=DataOut_Temperature, lowVal:=15, hiVal:=35, ForceResults:=tlForceNone
    Update_BC_PassFail_Flag
    
    If TheExec.CurrentJob = "CP1" Then
    Else: TheHdw.Wait 0.15
    End If
Else
    TestNameInput = Report_TName_From_Instance(CalcT, "X", , 0)
    TheExec.Flow.TestLimit resultVal:=DataOut_Temperature, ForceResults:=tlForceFlow, Tname:=TestNameInput
End If

Call TMPS_Temperature2iEDA(argv(0), DataOut_Temperature)


End Function

Public Function Calc_PCIE_ADC(argc As Integer, argv() As String) As Long
Dim site As Variant
Dim DataOut_ADC_Code_0 As New DSPWave
Dim DataOut_ADC_Code_1 As New DSPWave
Dim DataOut_ADC_Code_0_OffSet As New SiteLong
Dim DataOut_ADC_Code_1_OffSet As New SiteLong
Dim DataOut_ADC_Code_OffSet_Average As New SiteLong
Dim DataOut_ADC_Code_Average As New SiteLong
Dim DataOut_ADC_Code_Average_Dict As New DSPWave
Dim DataOut_ADC_Code_Final As New SiteLong
Dim DataOut_ADC_Voltage_0 As New SiteDouble
Dim DataOut_ADC_Voltage_1 As New SiteDouble
Dim DataOut_ADC_Voltage_Average As New SiteDouble
Dim DataOut_ADC_Voltage_Out As New SiteDouble
Dim Str_Split() As String
Dim i As Integer
Dim TestNameInput As String
Dim OutputTname_format() As String

DataOut_ADC_Code_0.CreateConstant 0, 1, DspLong
DataOut_ADC_Code_1.CreateConstant 0, 1, DspLong
DataOut_ADC_Code_Average_Dict.CreateConstant 0, 1, DspLong

If argv(0) Like "*adc_offset*" Then
Else
    DataOut_ADC_Code_Average_Dict = GetStoredCaptureData("ADC_OFFSET_AVERAGE_X")
End If

Call HardIP_Bin2Dec(DataOut_ADC_Code_0, GetStoredCaptureData(argv(0)))
Call HardIP_Bin2Dec(DataOut_ADC_Code_1, GetStoredCaptureData(argv(1)))

For Each site In TheExec.sites
    DataOut_ADC_Voltage_0(site) = TheHdw.DCVS.Pins("VDD12_PCIE").Voltage.Value * DataOut_ADC_Code_0(site).Element(0) / 255
    DataOut_ADC_Voltage_1(site) = TheHdw.DCVS.Pins("VDD12_PCIE").Voltage.Value * DataOut_ADC_Code_1(site).Element(0) / 255
    DataOut_ADC_Voltage_Average(site) = (DataOut_ADC_Voltage_0(site) + DataOut_ADC_Voltage_1(site)) / 2
    If argv(0) Like "*adc_offset_adc*" Then
        DataOut_ADC_Code_0_OffSet(site) = DataOut_ADC_Code_0(site).Element(0) - 128
        DataOut_ADC_Code_1_OffSet(site) = DataOut_ADC_Code_1(site).Element(0) - 128
        DataOut_ADC_Code_OffSet_Average(site) = (DataOut_ADC_Code_0_OffSet(site) + DataOut_ADC_Code_1_OffSet(site)) / 2
        DataOut_ADC_Code_Average_Dict(site).Element(0) = DataOut_ADC_Code_OffSet_Average(site)
    Else
    DataOut_ADC_Code_Average(site) = (DataOut_ADC_Code_0(site).Element(0) + DataOut_ADC_Code_1(site).Element(0)) / 2
    DataOut_ADC_Code_Final(site) = DataOut_ADC_Code_Average(site) - DataOut_ADC_Code_Average_Dict(site).Element(0)
    DataOut_ADC_Voltage_Out(site) = 0.25 * TheHdw.DCVS.Pins("VDD12_PCIE").Voltage.Value + DataOut_ADC_Code_Final(site) * TheHdw.DCVS.Pins("VDD12_PCIE").Voltage.Value * 0.5 / 256
    End If
Next site

If argv(0) Like "*adc_offset_adc*" Then
    Call AddStoredCaptureData("ADC_OFFSET_AVERAGE_X", DataOut_ADC_Code_Average_Dict)
End If

Str_Split = Split(argv(0), "_")

TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Voltage_0, Tname:="Voltage_" & argv(0), ForceResults:=tlForceFlow
TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Voltage_1, Tname:="Voltage_" & argv(1), ForceResults:=tlForceFlow
TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Voltage_Average, Tname:="Average_Voltage_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow

If argv(0) Like "*adc_offset_adc*" Then
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_0_OffSet, Tname:="OffSet_" & argv(0), ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_1_OffSet, Tname:="OffSet_" & argv(1), ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_OffSet_Average, Tname:="Average_OffSet_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
Else
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_Average, Tname:="Average_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Code_Final, Tname:="Final_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=DataOut_ADC_Voltage_Out, Tname:="Voltage_Out_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
End If

End Function

Public Function Calc_LPDPRX_Bin2Hex(argc As Integer, argv() As String) As Long
Dim i As Integer
Dim Data_Temp As String
Dim DSPWave_Dict As New DSPWave: DSPWave_Dict = GetStoredCaptureData(argv(0))
Dim hex_string As String
Dim site As Variant
    For Each site In TheExec.sites
    i = DSPWave_Dict(site).SampleSize - 1
        Do While (i >= 0)
            Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(i))
            i = i - 1
        Loop
        hex_string = Right(BinStr2HexStr(Data_Temp, DSPWave_Dict(site).SampleSize), 8)
        
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("<@Hexadecimal Code : " & UCase(argv(0)) & "|" & site & "|" & hex_string & ">")
        
        Data_Temp = ""
    Next site
End Function

Public Function TX_Level(argc As Integer, argv() As String) As Long
    
    ''20170711
    '' Calculate V_DP_K =  ((V_1 I_2-V_2 I_1 ) R_term)/(V_1-V_2+(R_term-R_(path) )(I_2-I_1 ) )
    '' where R_term = 45ohm; Rpath = trace + RAK
    ''TheExec.Datalog.WriteComment "Some Error in " & TheExec.DataManager.InstanceName
    
    Dim site As Variant
    Dim i As Long, j As Long
    Dim DictKey_V1 As String, DictKey_V2 As String
    Dim pld_V1 As New PinListData, pld_V2 As New PinListData
    Dim I1 As Double, I2 As Double
    Dim PinName As String
    Dim R_Term As Double
    Dim DictKey_Diff As String, DictKey_Diff_Calc As String
    Dim V_Diff As New SiteDouble
    
    DictKey_V1 = argv(0)
    DictKey_V2 = argv(1)
    I1 = CDbl(argv(2))
    I2 = CDbl(argv(3))
    PinName = argv(4)
    R_Term = CDbl(argv(5))
    
    If argc = 7 Then
        If argv(6) <> "" Then
            DictKey_Diff = argv(6)
        End If
    End If
    
    If argc = 8 Then
        If argv(7) <> "" Then
            'DictKey_Diff_Calc = argv(7)
            DictKey_Diff_Calc = argv(6) ' For Turks
        End If
    End If
    
    pld_V1 = GetStoredMeasurement(DictKey_V1)
    pld_V2 = GetStoredMeasurement(DictKey_V2)
    
    Dim R_Path As New SiteDouble
    'Dim R_Channel_RAK() As Double
    For Each site In TheExec.sites
        'R_Channel_RAK = TheHdw.PPMU.ReadRakValuesByPinnames(PinName, site)
        R_Path(site) = CurrentJob_Card_RAK.Pins(PinName).Value(site)
    Next site
    
    Dim V_DP_K As New SiteDouble
    
    For Each site In TheExec.sites
        V_DP_K(site) = ((pld_V1.Pins(PinName).Value(site) * I2 - pld_V2.Pins(PinName).Value(site) * I1) * R_Term) / (pld_V1.Pins(PinName).Value(site) - pld_V2.Pins(PinName).Value(site) + (R_Term - R_Path(site)) * (I2 - I1))
    Next site
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
''    TheExec.Flow.TestLimit resultVal:=V_DP_K, PinName:=PinName, Tname:="Volt_meas_TX_Level", ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcV, PinName, "", CInt(j))
    TheExec.Flow.TestLimit resultVal:=V_DP_K, PinName:=PinName, ForceResults:=tlForceFlow, Tname:=TestNameInput
    If argc = 7 Then
        If argv(6) <> "" Then
            Call AddStoredMeasurement(DictKey_Diff, V_DP_K)
        End If
    End If
    
    If argc = 8 Then
        If argv(7) <> "" Then
            V_Diff = GetStoredMeasurement(DictKey_Diff_Calc)
            'V_Diff = V_Diff.Subtract(V_DP_K)
            V_Diff = V_Diff.Subtract(V_DP_K).Abs 'For Tonga
            TestNameInput = Report_TName_From_Instance(CalcV, PinName, "", CInt(j))
            TheExec.Flow.TestLimit resultVal:=V_Diff, PinName:="P-N", ForceResults:=tlForceFlow, Tname:=TestNameInput
        End If
    End If
    
''    Dim TX_V1a As New PinListData
''    Dim TX_V1b As New PinListData
''    Dim TX_V1c As New PinListData
''    Dim TX_V1d As New PinListData
''
''    Dim TX_V2a As New PinListData
''    Dim TX_V2b As New PinListData
''    Dim TX_V2c As New PinListData
''    Dim TX_V2d As New PinListData
''
''    Dim TX0P As New SiteDouble
''    Dim TX1P As New SiteDouble
''
''    Dim TX0M As New SiteDouble
''    Dim TX1M As New SiteDouble
''
''    Dim R_path As Long
''    Dim R_term As Long
''    Dim Rdiff As New PinListData
''    Dim r_trace_TX0P() As Double
''    Dim r_trace_TX1P() As Double
''    Dim r_trace_TX0M() As Double
''    Dim r_trace_TX1M() As Double
''    Dim I1 As Double
''    Dim I2 As Double
''
''    R_path = 0
''    R_term = 50
''    I1 = 0.007
''    I2 = 0.009
''    Rdiff.AddPin ("TX0_P")
''    Rdiff.AddPin ("TX1_P")
''    Rdiff.AddPin ("TX0_M")
''    Rdiff.AddPin ("TX1_M")
''    For Each Site In TheExec.sites.Active
''        r_trace_TX0P = TheHdw.PPMU.ReadRakValuesByPinnames("TX0_P", Site)
''        r_trace_TX1P = TheHdw.PPMU.ReadRakValuesByPinnames("TX1_P", Site)
''        r_trace_TX0M = TheHdw.PPMU.ReadRakValuesByPinnames("TX0_M", Site)
''        r_trace_TX1M = TheHdw.PPMU.ReadRakValuesByPinnames("TX1_M", Site)
''
''        Rdiff.Pins("TX0_P") = R_term - CP_Card_RAK.Pins("TX0_P").Value - r_trace_TX0P(0)
''        Rdiff.Pins("TX1_P") = R_term - CP_Card_RAK.Pins("TX1_P").Value - r_trace_TX1P(0)
''        Rdiff.Pins("TX0_M") = R_term - CP_Card_RAK.Pins("TX0_M").Value - r_trace_TX0M(0)
''        Rdiff.Pins("TX1_M") = R_term - CP_Card_RAK.Pins("TX1_M").Value - r_trace_TX1M(0)
''    Next Site
''
''    TX_V1a = GetStoredMeasurement(argv(0))  'TX0_P,0.07,I1,TX0P
''    TX_V1b = GetStoredMeasurement(argv(1))  'TX1_P,0.07,I1,TX1P
''    TX_V1c = GetStoredMeasurement(argv(2))  'TX0_M,0.07,I1,TX0M
''    TX_V1d = GetStoredMeasurement(argv(3))  'TX1_M,0.07,I1,TX1M
''
''    TX_V2a = GetStoredMeasurement(argv(4))  'TX0_P,0.09,I2,TX0P
''    TX_V2b = GetStoredMeasurement(argv(5))  'TX1_P,0.09,I2,TX1P
''    TX_V2c = GetStoredMeasurement(argv(6))  'TX0_P,0.09,I2,TX0M
''    TX_V2d = GetStoredMeasurement(argv(7))  'TX1_P,0.09,I2,TX1M
''
''
'''(V1I2-V2I1)Rterm /  V1-V2+(Rterm-Rpath)(I2-I1)
''
''    For Each Site In TheExec.sites.Active
''        TX0P = ((TX_V1a.Pins("TX0_P").Value(Site) * I2 - TX_V2a.Pins("TX0_P").Value(Site) * I1) * R_term) / ((TX_V1a.Pins("TX0_P").Value(Site) - TX_V2a.Pins("TX0_P").Value(Site)) + Rdiff.Pins("TX0_P").Value(Site) * (I2 - I1))
''        TX1P = ((TX_V1b.Pins("TX1_P").Value(Site) * I2 - TX_V2b.Pins("TX1_P").Value(Site) * I1) * R_term) / ((TX_V1b.Pins("TX1_P").Value(Site) - TX_V2b.Pins("TX1_P").Value(Site)) + Rdiff.Pins("TX1_P").Value(Site) * (I2 - I1))
''        TX0M = ((TX_V1c.Pins("TX0_M").Value(Site) * I2 - TX_V2c.Pins("TX0_M").Value(Site) * I1) * R_term) / ((TX_V1c.Pins("TX0_M").Value(Site) - TX_V2c.Pins("TX0_M").Value(Site)) + Rdiff.Pins("TX0_M").Value(Site) * (I2 - I1))
''        TX1M = ((TX_V1d.Pins("TX1_M").Value(Site) * I2 - TX_V2d.Pins("TX1_M").Value(Site) * I1) * R_term) / ((TX_V1d.Pins("TX1_M").Value(Site) - TX_V2d.Pins("TX1_M").Value(Site)) + Rdiff.Pins("TX1_M").Value(Site) * (I2 - I1))
''
''    Next Site
''
''    TheExec.Flow.TestLimit resultVal:=TX0P, Tname:="TX0P_Level_H", ForceResults:=tlForceFlow
''    TheExec.Flow.TestLimit resultVal:=TX1P, Tname:="TX1P_Level_H", ForceResults:=tlForceFlow
''    TheExec.Flow.TestLimit resultVal:=TX0M, Tname:="TX0M_Level_H", ForceResults:=tlForceFlow
''    TheExec.Flow.TestLimit resultVal:=TX1M, Tname:="TX1M_Level_H", ForceResults:=tlForceFlow
        
    
End Function
Public Function TX_EQXXXXXXX(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim TX_Va_EQ As New PinListData
    Dim TX_Vb_EQ As New PinListData
    Dim TX_PM_EQ As New PinListData
    Dim OutputTname_format() As String
    Dim TestNameInput As String
            
    TX_Va_EQ = GetStoredMeasurement(argv(0))
    TX_Vb_EQ = GetStoredMeasurement(argv(1))
     TX_PM_EQ = TX_Va_EQ
'    TX_Va_EQ.AddPin ("Hello")
'    TX_Vb_EQ.AddPin ("Hi")
    If TheExec.TesterMode = testModeOffline Then
    
    For i = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active

            TX_Va_EQ.Pins(i).Value(site) = 10
            TX_Vb_EQ.Pins(i).Value(site) = 10
            
        Next site
'
    Next i
    
    End If
    
   
    
        
    For i = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active

            TX_PM_EQ.Pins(i).Value(site) = 20 * Log(TX_Va_EQ.Pins(i).Value(site) / TX_Vb_EQ.Pins(i).Value(site))

        Next site
'
    Next i

    For j = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active
                TestNameInput = Report_TName_From_Instance(CalcV, TX_PM_EQ.Pins(j), "", CInt(j))
                TheExec.Flow.TestLimit resultVal:=TX_PM_EQ.Pins(j).Value, Tname:=TestNameInput, ForceResults:=tlForceNone
        Next site
    Next j




End Function

Public Function TX_EQ(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim TX_Va_EQ As New PinListData
    Dim TX_Vb_EQ As New PinListData
    Dim TX_Vc_EQ As New PinListData
    
    Dim TX_PM_EQ As New PinListData
    Dim TX_PM_PRE As New PinListData
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    TX_Va_EQ = GetStoredMeasurement(argv(0))
    TX_Vb_EQ = GetStoredMeasurement(argv(1))
    TX_PM_EQ = TX_Va_EQ
    
    If (argc = 3) Then
        TX_Vc_EQ = GetStoredMeasurement(argv(2))
        TX_PM_PRE = TX_Vc_EQ
    End If
    
    
    
    If TheExec.TesterMode = testModeOffline Then
    For i = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active

            TX_Va_EQ.Pins(i).Value(site) = 10
            TX_Vb_EQ.Pins(i).Value(site) = 10
            If (argc = 3) Then
                TX_Vc_EQ.Pins(i).Value(site) = 10
            End If
        Next site
'
    Next i
    End If
    
    
    
        
    For i = 0 To TX_PM_EQ.Pins.Count - 1
        For Each site In TheExec.sites.Active

            If ((TX_Vb_EQ.Pins(i).Value(site) / TX_Va_EQ.Pins(i).Value(site)) > 0) Then
                TX_PM_EQ.Pins(i).Value(site) = 20 * Log10((TX_Vb_EQ.Pins(i).Value(site) / TX_Va_EQ.Pins(i).Value(site)))
            Else
                TX_PM_EQ.Pins(i).Value(site) = 0
            End If
            If (argc = 3) Then
                If (TX_Vc_EQ.Pins(i).Value(site) / TX_Vb_EQ.Pins(i).Value(site) > 0) Then
                    TX_PM_PRE.Pins(i).Value(site) = 20 * Log10(TX_Vc_EQ.Pins(i).Value(site) / TX_Vb_EQ.Pins(i).Value(site))
                Else
                    TX_PM_PRE.Pins(i).Value(site) = 0
                End If
            End If
        Next site
'
    Next i
    

    For j = 0 To TX_PM_EQ.Pins.Count - 1
        TestNameInput = Report_TName_From_Instance(CalcV, TX_PM_EQ.Pins(j), "", CInt(j))
        TheExec.Flow.TestLimit resultVal:=TX_PM_EQ.Pins(j), ForceResults:=tlForceFlow, Tname:=TestNameInput
    Next j

    If (argc = 3) Then
        For j = 0 To TX_PM_PRE.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance(CalcV, TX_PM_PRE.Pins(j), "", CInt(j))
            TheExec.Flow.TestLimit resultVal:=TX_PM_PRE.Pins(j), ForceResults:=tlForceFlow, Tname:=TestNameInput
        Next j
    End If


End Function
'CMRR and PSSR func modified for metrology 20170711
Public Function Calc_2S_Complement_To_SignDec_Modified(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey_2S_BIN As String
    Dim DictKey_SIGN_DEC As String
    
    Dim DSP_DictKey_2S_BIN As New DSPWave
    Dim DSP_DictKey_SIGN_DEC() As New DSPWave
Dim DSP_CMRR_CALC() As New DSPWave
Dim DSP_PSRR_CALC() As New DSPWave
    ReDim DSP_DictKey_SIGN_DEC(argc - 1) As New DSPWave
    ReDim DSP_CMRR_CALC(argc - 1) As New DSPWave
    ReDim DSP_PSRR_CALC(argc - 1) As New DSPWave
    Dim testName As String
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim StepIndex_Val As Long

    
    Dim SL_BitWidth As New SiteLong
    '' Format: Dict_2S_Com_A@Dict_SignDec_A@TestName_A,Dict_2S_Com_B@Dict_SignDec_B@TestName_B
    For i = 0 To argc - 1
        
        If InStr(TheExec.DataManager.instanceName, "T3") Then
        
            SplitByAt = Split(argv(i), "@")
            DictKey_2S_BIN = SplitByAt(0)
            
            DictKey_SIGN_DEC = SplitByAt(1)
            testName = SplitByAt(UBound(SplitByAt))
     
            DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
        Else
        
        
            DictKey_2S_BIN = argv(0)
            DictKey_SIGN_DEC = DictKey_2S_BIN
            testName = DictKey_2S_BIN
            DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
        End If

        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize

        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        
         Call AddStoredCaptureData(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))
        

        If Not ByPassTestLimit Then
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
            TheExec.Flow.TestLimit resultVal:=DSP_DictKey_SIGN_DEC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If

        If InStr(TheExec.DataManager.instanceName, "T2P6") <> 0 Then
            
                Set DSP_CMRR_CALC(i) = Nothing
                DSP_CMRR_CALC(i).CreateConstant 0, 1, DspDouble

                Dim CMRR_VIN_Calc As Double
                CMRR_VIN_Calc = CDbl(Replace(Split(DictKey_2S_BIN, "_")(2), "p", "."))
                For Each site In TheExec.sites
                    DSP_CMRR_CALC(i)(site).Element(0) = (DSP_DictKey_SIGN_DEC(i)(site).Element(0) / 131072) * 1.25
                    DSP_CMRR_CALC(i)(site).Element(0) = DSP_CMRR_CALC(i)(site).Element(0) / CMRR_VIN_Calc
                Next site
                Call AddStoredCaptureData(DictKey_2S_BIN, DSP_CMRR_CALC(i))
                If Not ByPassTestLimit Then
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CMRR", CInt(i))
                    TheExec.Flow.TestLimit resultVal:=DSP_CMRR_CALC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                End If
            
        End If
        ''=CMRR calculation END=
        ''=Osprey Metrology T2P7 PSRR calculation 20170605=
        If InStr(TheExec.DataManager.instanceName, "T2P7") <> 0 Then
            
                
                Set DSP_PSRR_CALC(i) = Nothing
                DSP_PSRR_CALC(i).CreateConstant 0, 1, DspDouble
                
                For Each site In TheExec.sites
                    If DSP_DictKey_SIGN_DEC(i)(site).Element(0) = 0 Then
                        DSP_DictKey_SIGN_DEC(i)(site).Element(0) = 1
                    End If
                    DSP_PSRR_CALC(i)(site).Element(0) = Abs((DSP_DictKey_SIGN_DEC(i)(site).Element(0) / 131072) * 1.25)
                    DSP_PSRR_CALC(i)(site).Element(0) = 20 * Log10(0.2 / DSP_PSRR_CALC(i)(site).Element(0))
                     ''Osprey Metrology T2P7 PSRR avergae store 20170606
                Next site
    
                Call AddStoredCaptureData(DictKey_2S_BIN, DSP_PSRR_CALC(i))
                If Not ByPassTestLimit Then
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "PSRR", CInt(i))
                 ' Call AddStoredCaptureData(SplitByAt(2), DSP_PSRR_CALC(i))
                    If Not ByPassTestLimit Then
                            TheExec.Flow.TestLimit resultVal:=DSP_PSRR_CALC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                    End If
                End If
        
        ''=PSRR calculation END=
        End If
    Next i

End Function

'CMRR and PSSR func modified for metrology 20170711
Public Function Calc_2S_Complement_To_SignDec_Modified_Nolimit(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey_2S_BIN As String
    Dim DictKey_SIGN_DEC As String
    
    Dim DSP_DictKey_2S_BIN As New DSPWave
    Dim DSP_DictKey_SIGN_DEC() As New DSPWave
Dim DSP_CMRR_CALC() As New DSPWave
Dim DSP_PSRR_CALC() As New DSPWave
    ReDim DSP_DictKey_SIGN_DEC(argc - 1) As New DSPWave
    ReDim DSP_CMRR_CALC(argc - 1) As New DSPWave
    ReDim DSP_PSRR_CALC(argc - 1) As New DSPWave
    Dim testName As String
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim StepIndex_Val As Long

    
    Dim SL_BitWidth As New SiteLong
    '' Format: Dict_2S_Com_A@Dict_SignDec_A@TestName_A,Dict_2S_Com_B@Dict_SignDec_B@TestName_B
    For i = 0 To argc - 1
        
        If InStr(TheExec.DataManager.instanceName, "T3") Then
        
            SplitByAt = Split(argv(i), "@")
            DictKey_2S_BIN = SplitByAt(0)
            
            DictKey_SIGN_DEC = SplitByAt(1)
            testName = SplitByAt(UBound(SplitByAt))
     
            DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
        Else
        
        
            DictKey_2S_BIN = argv(0)
    
            DictKey_SIGN_DEC = DictKey_2S_BIN
            
            testName = DictKey_2S_BIN
     
            DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
        End If

        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize

        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        
         Call AddStoredCaptureData(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))
        

    Next i

End Function


Public Function Calc_MDLL_Monotonicity_DevideBlock(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    
''    Call CreateSimulateMDLL_Data(argc, argv)
    
''    Dim DSPWaveBin() As New DSPWave
''    ReDim DSPWaveBin(argc - 1) As New DSPWave
    Dim DSPWaveDec() As New DSPWave
    ReDim DSPWaveDec((argc - 1) * 2 - 1) As New DSPWave
    Dim testName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    testName = argv(0) & "_"
    
    Dim DDR_MonoWithblock() As Type_MonoWithBlock
    ReDim DDR_MonoWithblock((argc - 1) * 2 - 1) As Type_MonoWithBlock
    Dim DSP_Input As New DSPWave
    Dim DSP_Input_UpperBIN As New DSPWave
    Dim DSP_Input_BelowBIN As New DSPWave
    Dim DSP_Input_UpperDEC As New DSPWave
    Dim DSP_Input_BelowDEC As New DSPWave
    Dim InputKey As String
    For i = 0 To argc - 2
        InputKey = LCase(argv(i + 1))
        DSP_Input = GetStoredCaptureData(InputKey)
        
        Call rundsp.SeprateDSP(DSP_Input, DSP_Input_UpperBIN, DSP_Input_BelowBIN)
        Call rundsp.BinToDec(DSP_Input_UpperBIN, DSP_Input_UpperDEC)
        Call rundsp.BinToDec(DSP_Input_BelowBIN, DSP_Input_BelowDEC)
        
        If InStr(InputKey, LCase("dll_l_1")) <> 0 Then
            DDR_MonoWithblock(i * 2).Block = 4
            DDR_MonoWithblock(i * 2).DSP_Bin = DSP_Input_UpperBIN
            DDR_MonoWithblock(i * 2).DSP_Dec = DSP_Input_UpperDEC
            DDR_MonoWithblock(i * 2 + 1).Block = 0
            DDR_MonoWithblock(i * 2 + 1).DSP_Bin = DSP_Input_BelowBIN
            DDR_MonoWithblock(i * 2 + 1).DSP_Dec = DSP_Input_BelowDEC
        ElseIf InStr(InputKey, LCase("dll_l_2")) <> 0 Then
            DDR_MonoWithblock(i * 2).Block = 6
            DDR_MonoWithblock(i * 2).DSP_Bin = DSP_Input_UpperBIN
            DDR_MonoWithblock(i * 2).DSP_Dec = DSP_Input_UpperDEC
            DDR_MonoWithblock(i * 2 + 1).Block = 1
            DDR_MonoWithblock(i * 2 + 1).DSP_Bin = DSP_Input_BelowBIN
            DDR_MonoWithblock(i * 2 + 1).DSP_Dec = DSP_Input_BelowDEC
        ElseIf InStr(InputKey, LCase("dll_m_1")) <> 0 Then
            DDR_MonoWithblock(i * 2).Block = 3
            DDR_MonoWithblock(i * 2).DSP_Bin = DSP_Input_UpperBIN
            DDR_MonoWithblock(i * 2).DSP_Dec = DSP_Input_UpperDEC
            DDR_MonoWithblock(i * 2 + 1).Block = 7
            DDR_MonoWithblock(i * 2 + 1).DSP_Bin = DSP_Input_BelowBIN
            DDR_MonoWithblock(i * 2 + 1).DSP_Dec = DSP_Input_BelowDEC
        ElseIf InStr(InputKey, LCase("dll_m_2")) <> 0 Then
            DDR_MonoWithblock(i * 2).Block = 2
            DDR_MonoWithblock(i * 2).DSP_Bin = DSP_Input_UpperBIN
            DDR_MonoWithblock(i * 2).DSP_Dec = DSP_Input_UpperDEC
            DDR_MonoWithblock(i * 2 + 1).Block = 5
            DDR_MonoWithblock(i * 2 + 1).DSP_Bin = DSP_Input_BelowBIN
            DDR_MonoWithblock(i * 2 + 1).DSP_Dec = DSP_Input_BelowDEC
        End If
    Next i
    
    Dim dataStr As String
    For Each site In TheExec.sites
        For i = 0 To UBound(DDR_MonoWithblock)
            dataStr = ""
            For j = 0 To DDR_MonoWithblock(i).DSP_Bin.SampleSize - 1
                If j = 0 Then
                    dataStr = DDR_MonoWithblock(i).DSP_Bin(site).Element(j)
                Else
                    dataStr = dataStr & DDR_MonoWithblock(i).DSP_Bin(site).Element(j)
                End If
            Next j
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " , Block = " & DDR_MonoWithblock(i).Block & " , Binary = " & dataStr & " , Decimal = " & DDR_MonoWithblock(i).DSP_Dec.Element(0))
        Next i
    Next site
    
    '' 20170713 - Sorting DDR_MonoWithblock by block
    Dim TempBlock As Long
    Dim sd_TempDSP_BIN As New DSPWave
    Dim sd_TempDSP_DEC As New DSPWave
    For i = 0 To UBound(DDR_MonoWithblock)
        For j = i To UBound(DDR_MonoWithblock)
            If DDR_MonoWithblock(i).Block > DDR_MonoWithblock(j).Block Then
                TempBlock = DDR_MonoWithblock(i).Block
                DDR_MonoWithblock(i).Block = DDR_MonoWithblock(j).Block
                DDR_MonoWithblock(j).Block = TempBlock
                
                sd_TempDSP_BIN = DDR_MonoWithblock(i).DSP_Bin
                DDR_MonoWithblock(i).DSP_Bin = DDR_MonoWithblock(j).DSP_Bin
                DDR_MonoWithblock(j).DSP_Bin = sd_TempDSP_BIN

                sd_TempDSP_DEC = DDR_MonoWithblock(i).DSP_Dec
                DDR_MonoWithblock(i).DSP_Dec = DDR_MonoWithblock(j).DSP_Dec
                DDR_MonoWithblock(j).DSP_Dec = sd_TempDSP_DEC
            End If
        Next j
    Next i
    
    '' Print info after sorting
    If gl_Disable_HIP_debug_log = False Then
        TheExec.Datalog.WriteComment ("Print info after sorting")
        For Each site In TheExec.sites
            For i = 0 To UBound(DDR_MonoWithblock)
                dataStr = ""
                For j = 0 To DDR_MonoWithblock(i).DSP_Bin.SampleSize - 1
                    If j = 0 Then
                        dataStr = DDR_MonoWithblock(i).DSP_Bin(site).Element(j)
                    Else
                        dataStr = dataStr & DDR_MonoWithblock(i).DSP_Bin(site).Element(j)
                    End If
                Next j
                TheExec.Datalog.WriteComment ("Site_" & site & " , Block = " & DDR_MonoWithblock(i).Block & " , Binary = " & dataStr & " , Decimal = " & DDR_MonoWithblock(i).DSP_Dec.Element(0))
            Next i
        Next site
    End If
    For i = 0 To UBound(DDR_MonoWithblock)
        DSPWaveDec(i) = DDR_MonoWithblock(i).DSP_Dec
    Next i
    
    For Each site In TheExec.sites
        For i = 0 To UBound(DDR_MonoWithblock)  'NEW 20170730
             'NEW 20170730
            'TestNameInput = Report_ALG_TName_From_Instance(OutputTname_format, "C", "X", "LockCodeRange", CInt(i))
            'TheExec.Flow.TestLimit resultVal:=DSPWaveDec(i)(Site).Element(0), lowVal:=0, hiVal:=119, Tname:=TestNameInput, ForceResults:=tlForceNone
            TestNameInput = Report_TName_From_Instance(CalcC, "X", "LockCodeRange", CInt(i))
            TheExec.Flow.TestLimit resultVal:=DSPWaveDec(i)(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
        Next i
    Next site
    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
    
    
    Dim MDLL_CurrentVal As New SiteLong
    Dim MDLL_PreviousVal  As New SiteLong
    Dim b_MDLL_DecreaseDirection As New SiteBoolean
    Dim b_MDLL_DecreaseAddIndex As New SiteBoolean
    Dim MDLL_DecreaseResultPass As New SiteLong
    Dim b_MDLL_TestResultFail As New SiteBoolean
    Dim MDLL_Index As New SiteLong
    b_MDLL_DecreaseDirection = False
    
    MDLL_DecreaseResultPass = 1
    b_MDLL_TestResultFail = False
    MDLL_Index = 1
    Dim StepSize As Long
    Dim StoreDecreaseVal As New SiteVariant
    Dim StoreDecreaseIndex As Long
    StoreDecreaseIndex = 0
    For Each site In TheExec.sites
'       For i = 1 To argc - 1
        For i = 0 To UBound(DDR_MonoWithblock)  'NEW 20170730
            If i = 0 Then
                MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
                
                StoreDecreaseVal(site) = CStr(MDLL_CurrentVal(site))
                StoreDecreaseIndex = StoreDecreaseIndex + 1
            Else
                MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                b_MDLL_DecreaseDirection(site) = MDLL_CurrentVal.Subtract(MDLL_PreviousVal).compare(LessThanOrEqualTo, 0)
                
                '' Fail  as below
                If b_MDLL_DecreaseDirection(site) = False Then
                    MDLL_DecreaseResultPass(site) = 0
''                    b_MDLL_TestResultFail(Site) = True
''                    Exit For
                End If
                
''                b_MDLL_DecreaseAddIndex(Site) = MDLL_CurrentVal.Subtract(MDLL_PreviousVal).compare(LessThan, 0)
''
''                If b_MDLL_DecreaseAddIndex(Site) = True Then
''                    MDLL_Index(Site) = MDLL_Index(Site) + 1
''
                StoreDecreaseVal(site) = StoreDecreaseVal(site) & "," & MDLL_CurrentVal(site)
                StoreDecreaseIndex = StoreDecreaseIndex + 1
''                End If
''                If MDLL_Index(Site) > 1 Then
''''                    b_MDLL_TestResultFail(Site) = True
''                    Exit For
''                End If
                
                MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
            End If
        Next i
    Next site
    
    Dim OriginalVal() As String
    Dim TempVal As Double
    Dim SortedVal() As Double
    
    Dim DiffVal_Num As New SiteLong
''    Dim DiffVal_Judge As New SiteBoolean
    Dim DiffVal_MaxSubMin As New SiteLong
    DiffVal_Num = 1
    
    For Each site In TheExec.sites
        OriginalVal = Split(StoreDecreaseVal(site), ",")
        ReDim SortedVal(UBound(OriginalVal)) As Double
        For i = 0 To UBound(OriginalVal)
            SortedVal(i) = CDbl(OriginalVal(i))
        Next i
''        SortedVal = CDbl(OriginalVal)
        For i = 0 To UBound(SortedVal)
            For j = i To UBound(SortedVal)
                If SortedVal(i) > SortedVal(j) Then
                    TempVal = SortedVal(i)
                    SortedVal(i) = SortedVal(j)
                    SortedVal(j) = TempVal
                End If
            Next j
        Next i
        For i = 0 To UBound(SortedVal) - 1
            If SortedVal(i + 1) - SortedVal(i) > 0 Then
                DiffVal_Num(site) = DiffVal_Num(site) + 1
            End If
        Next i
        DiffVal_MaxSubMin(site) = SortedVal(UBound(SortedVal)) - SortedVal(0)
    Next site

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Decrease", 0)

    TheExec.Flow.TestLimit resultVal:=MDLL_DecreaseResultPass, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceNone
''    For Each Site In TheExec.sites
''        If MDLL_DecreaseResultPass.BitwiseAnd(1) Then
''        Else
''            MDLL_Index(Site) = -99
''        End If
''    Next Site
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Unique", 0)
    
    TheExec.Flow.TestLimit resultVal:=DiffVal_Num, lowVal:=1, hiVal:=2, Tname:=testName & "Unique", ForceResults:=tlForceNone
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "MaxDiff", 0)

    TheExec.Flow.TestLimit resultVal:=DiffVal_MaxSubMin, lowVal:=0, hiVal:=1, Tname:=testName & "Max_Diff", ForceResults:=tlForceNone
End Function
Public Function Calc_Metrology_GainError(argc As Integer, argv() As String) As Long
    Dim Dict_ReturnKey As String
    Dim Dict_InputKey As String
    Dim InputVal As New PinListData
    Dim CalcVal As New PinListData
    
    Dict_ReturnKey = argv(0)
    Dict_InputKey = argv(1)
    InputVal = GetStoredMeasurement(Dict_InputKey)
    
    CalcVal.AddPin (InputVal.Pins(0))
    CalcVal = InputVal.Pins(0).Subtract(0.4).Divide(0.7975).Subtract(1)
    Call AddStoredMeasurement(Dict_ReturnKey, CalcVal)
End Function

Public Function Calc_MIPI_CodeTolerance(argc As Integer, argv() As String) As Long
        
    Dim i As Long, j As Long
    Dim X As Integer
    Dim site As Variant
    Dim InputDSPWave_BIN As New DSPWave
    Dim InputDSPWave_DEC As New DSPWave
    Dim MIPI_threshold_Code_value1(7) As New SiteDouble
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    For i = 0 To UBound(argv)
        InputDSPWave_BIN = GetStoredCaptureData(argv(i))
        Call rundsp.BinToDec(InputDSPWave_BIN, InputDSPWave_DEC)
        For Each site In TheExec.sites
            MIPI_threshold_Code_value1(i)(site) = InputDSPWave_DEC(site).Element(0)
        Next site
    Next i

    Dim MIPI_threshold_lower(0) As New SiteVariant
    Dim MIPI_threshold_high(0) As New SiteVariant
    Dim MIPI_threshold_found As New SiteBoolean
    Dim MIPI_trans_mapping As Variant
    
    Dim threshold_temp As Integer
    Dim threshold_flag1 As Boolean
    Dim p  As Long
    MIPI_trans_mapping = Array(-0.2, -0.15, -0.1, -0.05, 0.05, 0.1, 0.15, 0.2)

    X = 0

    For Each site In TheExec.sites
        threshold_temp = 0
        threshold_flag1 = False
        MIPI_threshold_found(site) = False
        
        For p = 0 To UBound(argv)

''            MIPI_threshold_Code_value1(p)(Site) = DigCapVal_DSSC_Out(0, p * 2)(Site) + 256 * DigCapVal_DSSC_Out(0, p * 2 + 1)(Site)

            If MIPI_threshold_Code_value1(p)(site) = 0 Then
                If threshold_flag1 = False Then
                    MIPI_threshold_lower(0)(site) = p
                    threshold_flag1 = True
                    MIPI_threshold_found(site) = True
                End If
                If threshold_flag1 = True Then
                    MIPI_threshold_high(0)(site) = p
                End If
            End If
            If MIPI_threshold_Code_value1(p)(site) > 0 Then
                threshold_temp = threshold_temp + 1
            End If
        Next p
        
        If threshold_temp = 0 Then
            MIPI_threshold_found = False
        End If
        
        If MIPI_threshold_lower(0)(site) <> "" Then
            MIPI_threshold_lower(0)(site) = MIPI_trans_mapping(MIPI_threshold_lower(0)(site))
        Else
            MIPI_threshold_lower(0)(site) = 999
        End If

         If MIPI_threshold_high(0)(site) <> "" Then
            MIPI_threshold_high(0)(site) = MIPI_trans_mapping(MIPI_threshold_high(0)(site))
        Else
            MIPI_threshold_high(0)(site) = 999
        End If

    Next site

    For p = 0 To 7
        TestNameInput = Report_TName_From_Instance(CalcC, "code1_" & p + 1, , CInt(X))
        
        TheExec.Flow.TestLimit MIPI_threshold_Code_value1(p), 0, 2 ^ 10 - 1, PinName:="code1_" & p + 1, ForceResults:=tlForceFlow
    Next p

    TestNameInput = Report_TName_From_Instance(CalcC, "MIPI_Tolerance1_1", , CInt(X))

    TheExec.Flow.TestLimit MIPI_threshold_lower(0), scaletype:=scaleNone, PinName:="MIPI_Tolerance1_1", ForceResults:=tlForceFlow
    'TheExec.Flow.TestLimit MIPI_threshold_lower(0), ScaleType:=None, PinName:="MIPI_Tolerance1_1", ForceResults:=tlForceFlow ''OscarLi_Compile,20190629
    TestNameInput = Report_TName_From_Instance(CalcC, "MIPI_Tolerance1_2", , CInt(X))
      
    TheExec.Flow.TestLimit MIPI_threshold_high(0), scaletype:=scaleNone, PinName:="MIPI_Tolerance1_2", ForceResults:=tlForceFlow
    'TheExec.Flow.TestLimit MIPI_threshold_high(0), ScaleType:=None, PinName:="MIPI_Tolerance1_1", ForceResults:=tlForceFlow ''OscarLi_Compile,20190629
    TestNameInput = Report_TName_From_Instance(CalcC, "MIPI_threshold_found", , CInt(X))
    
    TheExec.Flow.TestLimit MIPI_threshold_found, True, True, PinName:="MIPI_threshold_found", ForceResults:=tlForceFlow

End Function
Public Function Calc_Metrology_GainErrorOffset(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim Dict_tfe_vol_1 As String
    Dim Dict_tfe_vol_0 As String

    Dim CapturedCode1 As String
    Dim CapturedCode2 As String
    Dim CapturedCode3 As String
    Dim CapturedCode4 As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    Dim DSP_tfe_vol_1_in_decimal As New DSPWave
    Dim DSP_tfe_vol_1_in_binary As New DSPWave


    Dim SL_BitWidth As New SiteLong
    
    Dim X As Long

    Dict_tfe_vol_1 = argv(0)
    CapturedCode1 = argv(1)
    CapturedCode2 = argv(2)
    CapturedCode3 = argv(3)
    CapturedCode4 = argv(4)
    Dict_tfe_vol_0 = argv(5)


    Dim DSP_tfe_vol_0_in_2S_binary As New DSPWave
    Dim DSP_tfe_vol_0_in_decimal As New DSPWave

    Dim DSP_gainErrorOffset1 As New DSPWave
    Dim DSP_gainErrorOffset2 As New DSPWave
    Dim DSP_gainErrorOffset3 As New DSPWave
    Dim DSP_gainErrorOffset4 As New DSPWave

    Dim DSP_gainErrorOffset1_decimal As New DSPWave
    Dim DSP_gainErrorOffset2_decimal As New DSPWave
    Dim DSP_gainErrorOffset3_decimal As New DSPWave
    Dim DSP_gainErrorOffset4_decimal As New DSPWave

    X = 0

    DSP_gainErrorOffset1 = GetStoredCaptureData(CapturedCode1)
    DSP_gainErrorOffset2 = GetStoredCaptureData(CapturedCode2)
    DSP_gainErrorOffset3 = GetStoredCaptureData(CapturedCode3)
    DSP_gainErrorOffset4 = GetStoredCaptureData(CapturedCode4)

    DSP_tfe_vol_0_in_2S_binary = GetStoredCaptureData(Dict_tfe_vol_0)
    For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_tfe_vol_0_in_2S_binary(site).SampleSize
            
            'Test Run
            
'            '111111111111111000
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(0) = 0
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(1) = 0
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(2) = 0
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(3) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(4) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(5) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(6) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(7) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(8) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(9) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(10) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(11) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(12) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(13) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(14) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(15) = 1
'             DSP_tfe_vol_0_in_2S_binary(Site).Element(16) = 1
'            DSP_tfe_vol_0_in_2S_binary(Site).Element(17) = 1
'
'
'            '001000000000001110
'            DSP_gainErrorOffset1(Site).Element(0) = 0
'            DSP_gainErrorOffset1(Site).Element(1) = 1
'            DSP_gainErrorOffset1(Site).Element(2) = 1
'            DSP_gainErrorOffset1(Site).Element(3) = 1
'            DSP_gainErrorOffset1(Site).Element(4) = 0
'            DSP_gainErrorOffset1(Site).Element(5) = 0
'            DSP_gainErrorOffset1(Site).Element(6) = 0
'            DSP_gainErrorOffset1(Site).Element(7) = 0
'            DSP_gainErrorOffset1(Site).Element(8) = 0
'            DSP_gainErrorOffset1(Site).Element(9) = 0
'            DSP_gainErrorOffset1(Site).Element(10) = 0
'            DSP_gainErrorOffset1(Site).Element(11) = 0
'            DSP_gainErrorOffset1(Site).Element(12) = 0
'            DSP_gainErrorOffset1(Site).Element(13) = 0
'            DSP_gainErrorOffset1(Site).Element(14) = 0
'            DSP_gainErrorOffset1(Site).Element(15) = 1
'             DSP_gainErrorOffset1(Site).Element(16) = 0
'            DSP_gainErrorOffset1(Site).Element(17) = 0
'
'
'            '000111111111101011
'            DSP_gainErrorOffset2(Site).Element(0) = 1
'            DSP_gainErrorOffset2(Site).Element(1) = 1
'            DSP_gainErrorOffset2(Site).Element(2) = 0
'            DSP_gainErrorOffset2(Site).Element(3) = 1
'            DSP_gainErrorOffset2(Site).Element(4) = 0
'            DSP_gainErrorOffset2(Site).Element(5) = 1
'            DSP_gainErrorOffset2(Site).Element(6) = 1
'            DSP_gainErrorOffset2(Site).Element(7) = 1
'            DSP_gainErrorOffset2(Site).Element(8) = 1
'            DSP_gainErrorOffset2(Site).Element(9) = 1
'            DSP_gainErrorOffset2(Site).Element(10) = 1
'            DSP_gainErrorOffset2(Site).Element(11) = 1
'            DSP_gainErrorOffset2(Site).Element(12) = 1
'            DSP_gainErrorOffset2(Site).Element(13) = 1
'            DSP_gainErrorOffset2(Site).Element(14) = 1
'            DSP_gainErrorOffset2(Site).Element(15) = 0
'             DSP_gainErrorOffset2(Site).Element(16) = 0
'            DSP_gainErrorOffset2(Site).Element(17) = 0
'
'            '000111111111100110
'            DSP_gainErrorOffset3(Site).Element(0) = 0
'            DSP_gainErrorOffset3(Site).Element(1) = 1
'            DSP_gainErrorOffset3(Site).Element(2) = 1
'            DSP_gainErrorOffset3(Site).Element(3) = 0
'            DSP_gainErrorOffset3(Site).Element(4) = 0
'            DSP_gainErrorOffset3(Site).Element(5) = 1
'            DSP_gainErrorOffset3(Site).Element(6) = 1
'            DSP_gainErrorOffset3(Site).Element(7) = 1
'            DSP_gainErrorOffset3(Site).Element(8) = 1
'            DSP_gainErrorOffset3(Site).Element(9) = 1
'            DSP_gainErrorOffset3(Site).Element(10) = 1
'            DSP_gainErrorOffset3(Site).Element(11) = 1
'            DSP_gainErrorOffset3(Site).Element(12) = 1
'            DSP_gainErrorOffset3(Site).Element(13) = 1
'            DSP_gainErrorOffset3(Site).Element(14) = 1
'            DSP_gainErrorOffset3(Site).Element(15) = 0
'             DSP_gainErrorOffset3(Site).Element(16) = 0
'            DSP_gainErrorOffset3(Site).Element(17) = 0
'
'
'            '001000000000011010
'            DSP_gainErrorOffset4(Site).Element(0) = 0
'            DSP_gainErrorOffset4(Site).Element(1) = 1
'            DSP_gainErrorOffset4(Site).Element(2) = 0
'            DSP_gainErrorOffset4(Site).Element(3) = 1
'            DSP_gainErrorOffset4(Site).Element(4) = 1
'            DSP_gainErrorOffset4(Site).Element(5) = 0
'            DSP_gainErrorOffset4(Site).Element(6) = 0
'            DSP_gainErrorOffset4(Site).Element(7) = 0
'            DSP_gainErrorOffset4(Site).Element(8) = 0
'            DSP_gainErrorOffset4(Site).Element(9) = 0
'            DSP_gainErrorOffset4(Site).Element(10) = 0
'            DSP_gainErrorOffset4(Site).Element(11) = 0
'            DSP_gainErrorOffset4(Site).Element(12) = 0
'            DSP_gainErrorOffset4(Site).Element(13) = 0
'            DSP_gainErrorOffset4(Site).Element(14) = 0
'            DSP_gainErrorOffset4(Site).Element(15) = 1
'             DSP_gainErrorOffset4(Site).Element(16) = 0
'            DSP_gainErrorOffset4(Site).Element(17) = 0
            
            
            

    Next site
    DSP_tfe_vol_0_in_decimal.CreateConstant 0, 1, DspLong
    
    

    Call rundsp.DSP_2S_Complement_To_SignDec(DSP_tfe_vol_0_in_2S_binary, SL_BitWidth, DSP_tfe_vol_0_in_decimal)

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfe_vol_0", CInt(X))

    TheExec.Flow.TestLimit resultVal:=DSP_tfe_vol_0_in_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone

    Call rundsp.BinToDec(DSP_gainErrorOffset1, DSP_gainErrorOffset1_decimal)
    Call rundsp.BinToDec(DSP_gainErrorOffset2, DSP_gainErrorOffset2_decimal)
    Call rundsp.BinToDec(DSP_gainErrorOffset3, DSP_gainErrorOffset3_decimal)
    Call rundsp.BinToDec(DSP_gainErrorOffset4, DSP_gainErrorOffset4_decimal)

    DSP_tfe_vol_1_in_decimal.CreateConstant 0, 1, DspLong

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode1_Dec", CInt(X))

    TheExec.Flow.TestLimit resultVal:=DSP_gainErrorOffset1_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode2_Dec", CInt(X))
    
    TheExec.Flow.TestLimit resultVal:=DSP_gainErrorOffset2_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode3_Dec", CInt(X))
    
    TheExec.Flow.TestLimit resultVal:=DSP_gainErrorOffset3_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode4_Dec", CInt(X))
    
    TheExec.Flow.TestLimit resultVal:=DSP_gainErrorOffset4_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone

    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Gain_vol_1_LimitExceeded_Dec", CInt(X))
    
    For Each site In TheExec.sites


        DSP_tfe_vol_1_in_decimal(site).Element(0) = DSP_gainErrorOffset1_decimal(site).Element(0) + DSP_gainErrorOffset2_decimal(site).Element(0) + DSP_gainErrorOffset3_decimal(site).Element(0) + DSP_gainErrorOffset4_decimal(site).Element(0) - 4 * DSP_tfe_vol_0_in_decimal(site).Element(0)
        If (DSP_tfe_vol_1_in_decimal(site).Element(0) > 262143) Then
            TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  Gain_vol_1_LimitExceeded_Dec = " + CStr(DSP_tfe_vol_1_in_decimal(site).Element(0)) + ", Force tfe_vol_1_dec = 174762")
            DSP_tfe_vol_1_in_decimal(site).Element(0) = 174762
        End If

    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfe_vol_1_dec", CInt(X))
    
    TheExec.Flow.TestLimit resultVal:=DSP_tfe_vol_1_in_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
    

    Call rundsp.DSPWf_Dec2Binary(DSP_tfe_vol_1_in_decimal, 18, DSP_tfe_vol_1_in_binary)

    Call AddStoredCaptureData(Dict_tfe_vol_1, DSP_tfe_vol_1_in_binary)
    
    Dim tfe_vol_1_bin_str As String
    Dim i As Long
   
    For Each site In TheExec.sites

            tfe_vol_1_bin_str = ""
         For i = DSP_tfe_vol_1_in_binary(site).SampleSize - 1 To 0 Step -1
         
                tfe_vol_1_bin_str = tfe_vol_1_bin_str + CStr(DSP_tfe_vol_1_in_binary(site).Element(i))
            
         Next i
      
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  tfe_vol_1_binary_fuse_Code  " + tfe_vol_1_bin_str)
    Next site
     
   ' TheExec.Flow.TestLimit resultVal:=DSP_tfe_vol_1_in_binary.Element(0), Tname:="tfe_vol_1", ForceResults:=tlForceNone


End Function
Public Function Calc_Metrology_EncodeActualTemp(argc As Integer, argv() As String) As Long

Dim actual_temp As Double
Dim Dict_actual_temp As String
Dim site As Variant

actual_temp = CDbl(argv(0))
Dict_actual_temp = argv(1)


Dim actual_temp_cal1 As Double
actual_temp_cal1 = (actual_temp - 25) * 64

Dim conv_temp_rounded As Long

conv_temp_rounded = FormatNumber(actual_temp_cal1)


Dim DSP_conv_temp_rounded As New DSPWave
Dim DSP_conv_temp_rounded_binary As New DSPWave

DSP_conv_temp_rounded.CreateConstant 0, 1, DspLong
DSP_conv_temp_rounded_binary.CreateConstant 0, 1, DspLong


For Each site In TheExec.sites

DSP_conv_temp_rounded(site).Element(0) = conv_temp_rounded

Next site

Call rundsp.DSPWf_Dec2Binary(DSP_conv_temp_rounded, 10, DSP_conv_temp_rounded_binary)

Call AddStoredCaptureData(Dict_actual_temp, DSP_conv_temp_rounded_binary)


End Function

Public Function Calc_Metrology_DecodeActualTemp(argc As Integer, argv() As String) As Long

Dim Dict_decoded_temp As String
Dim Dict_encoded_temp As String
Dim site As Variant
Dim SL_BitWidth As New SiteLong

Dim Dict_encoded_temp_in_2S_binary As New DSPWave
Dim Dict_encoded_temp_in_Decimal As New DSPWave



Dim Dict_decoded_temp_in_Decimal As New DSPWave


Dict_encoded_temp = argv(0)
Dict_decoded_temp = argv(1)

Dict_encoded_temp_in_2S_binary = GetStoredCaptureData(Dict_encoded_temp)


''  Test Data for y0 25C

'For Each Site In TheExec.sites

  '  Dict_encoded_temp_in_2S_binary(Site).Element(0) = 1
   ' Dict_encoded_temp_in_2S_binary(Site).Element(1) = 0

   ' Dict_encoded_temp_in_2S_binary(Site).Element(2) = 1
   ' Dict_encoded_temp_in_2S_binary(Site).Element(3) = 1
   ' Dict_encoded_temp_in_2S_binary(Site).Element(4) = 0
   ' Dict_encoded_temp_in_2S_binary(Site).Element(5) = 1
    'Dict_encoded_temp_in_2S_binary(Site).Element(6) = 1
    'Dict_encoded_temp_in_2S_binary(Site).Element(7) = 0
    'Dict_encoded_temp_in_2S_binary(Site).Element(8) = 1
    'Dict_encoded_temp_in_2S_binary(Site).Element(9) = 1


'Next Site


''

For Each site In TheExec.sites
            SL_BitWidth(site) = Dict_encoded_temp_in_2S_binary(site).SampleSize
Next site

Dict_encoded_temp_in_Decimal.CreateConstant 0, 1, DspLong


Call rundsp.DSP_2S_Complement_To_SignDec(Dict_encoded_temp_in_2S_binary, SL_BitWidth, Dict_encoded_temp_in_Decimal)

Dict_decoded_temp_in_Decimal.CreateConstant 0, 1, DspDouble







For Each site In TheExec.sites

Dict_decoded_temp_in_Decimal(site).Element(0) = (CDbl(Dict_encoded_temp_in_Decimal(site).Element(0)) / 64) + 25

Next site



Call AddStoredCaptureData(Dict_decoded_temp, Dict_decoded_temp_in_Decimal)


End Function

Public Function Calc_Metrology_adc_tfe_temp_fuses(argc As Integer, argv() As String) As Long

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
    
    Dim X As Long
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    

    Dict_name_tfe_vol_x1 = argv(0)
    cal_tfe_vol_y1 = argv(1)
    fuse_read_tfe_vol_0 = argv(2)
    fuse_read_tfe_vol_1 = argv(3)
    fuse_read_tfe_x0 = argv(4)
    fuse_read_tfe_y0 = argv(5)

    fuse_write_tfe_temp_0 = argv(6)
    fuse_write_tfe_temp_1 = argv(7)




    'Get Cap data for t5p2 at 85C and Fuse Data for offset,gain and x0 at 25C
    Dim DSP_tfe_vol_x1_binary As New DSPWave
    Dim DSP_fuse_read_tfe_vol_0_2S_binary As New DSPWave
    Dim DSP_fuse_read_tfe_vol_1_binary As New DSPWave
    Dim DSP_fuse_read_tfe_x0_binary As New DSPWave





    DSP_tfe_vol_x1_binary = GetStoredCaptureData(Dict_name_tfe_vol_x1)
    DSP_fuse_read_tfe_vol_0_2S_binary = GetStoredCaptureData(fuse_read_tfe_vol_0)
    DSP_fuse_read_tfe_vol_1_binary = GetStoredCaptureData(fuse_read_tfe_vol_1)
    DSP_fuse_read_tfe_x0_binary = GetStoredCaptureData(fuse_read_tfe_x0)



    ' Test Inputs
    
'    For Each Site In TheExec.sites
'
'    'test Run
'
'            '000010000101001000
'            DSP_fuse_read_tfe_x0_binary(Site).Element(0) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(1) = 1
'            DSP_fuse_read_tfe_x0_binary(Site).Element(2) = 1
'            DSP_fuse_read_tfe_x0_binary(Site).Element(3) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(4) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(5) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(6) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(7) = 1
'            DSP_fuse_read_tfe_x0_binary(Site).Element(8) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(9) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(10) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(11) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(12) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(13) = 1
'            DSP_fuse_read_tfe_x0_binary(Site).Element(14) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(15) = 0
'             DSP_fuse_read_tfe_x0_binary(Site).Element(16) = 0
'            DSP_fuse_read_tfe_x0_binary(Site).Element(17) = 0
'
'
'            '000010011100101001
''            DSP_tfe_vol_x1_binary(Site).Element(0) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(1) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(2) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(3) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(4) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(5) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(6) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(7) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(8) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(9) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(10) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(11) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(12) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(13) = 1
''            DSP_tfe_vol_x1_binary(Site).Element(14) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(15) = 0
''             DSP_tfe_vol_x1_binary(Site).Element(16) = 0
''            DSP_tfe_vol_x1_binary(Site).Element(17) = 0
''
'
'            '111111111111111000
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(0) = 0
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(1) = 0
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(2) = 0
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(3) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(4) = 0
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(5) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(6) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(7) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(8) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(9) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(10) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(11) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(12) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(13) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(14) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(15) = 1
'             DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(16) = 1
'            DSP_fuse_read_tfe_vol_0_2S_binary(Site).Element(17) = 1
'
'
'
'
'            '100000000000011000
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(0) = 1
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(1) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(2) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(3) = 1
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(4) = 1
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(5) = 1
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(6) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(7) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(8) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(9) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(10) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(11) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(12) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(13) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(14) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(15) = 0
'             DSP_fuse_read_tfe_vol_1_binary(Site).Element(16) = 0
'            DSP_fuse_read_tfe_vol_1_binary(Site).Element(17) = 1
'
'
'
'    Next Site
'
    
    ''


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
    If cal_tfe_vol_y1 Like "CP2" Then

        actual_Temp_CP2 = 85

    End If

    Dim DSP_tfe_y1_in_double As New DSPWave

    DSP_tfe_y1_in_double.CreateConstant 0, 1, DspDouble

    For Each site In TheExec.sites

    DSP_tfe_y1_in_double(site).Element(0) = actual_Temp_CP2

    Next site

'    'Check for Encode Logic ..Can comment it
'
'    Dim encoded_tfe_y1_in_2S_binary As String
'    Dim DSP_tfe_y1_in_2S_binary As New DSPWave
'    encoded_tfe_y1_in_2S_binary = "encoded_tfe_y1_in_2S_binary"
'        Dim call_encode_argv(2) As String
'    call_encode_argv(0) = CStr(actual_Temp_CP2)
'    call_encode_argv(1) = encoded_tfe_y1_in_2S_binary
'
'    Dim call_encodeActualTemp As Long
'    call_encodeActualTemp = Calc_Metrology_EncodeActualTemp(1, call_encode_argv)
'
'    DSP_tfe_y1_in_2S_binary = GetStoredCaptureData(encoded_tfe_y1_in_2S_binary)
'
'    'Check End


    'Start the algo


    'Define Constants

    Dim C0 As Double
    Dim c1 As Double
    Dim C2 As Double
    Dim C3 As Double

    'Values for Constants

    C0 = CDbl("-21.5822184999726")
    c1 = CDbl("428.0092266096283") 'truncated one digit
    C2 = CDbl("-133.4543109228228") 'truncated one digit
    C3 = CDbl("19.0485545665615")
    



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




    Dim X0 As New SiteDouble
    Dim X1 As New SiteDouble

    Dim Y0 As New SiteDouble
    Dim Y1 As New SiteDouble

    Dim c1_cal As New SiteDouble
    Dim c0_cal As New SiteDouble

    Dim tfe_temp0_double As New SiteDouble
    Dim tfe_temp1_double As New SiteDouble

    Dim tfe_temp0_long As New SiteLong
    Dim tfe_temp1_long As New SiteLong

    Dim Dsp_tfe_temp0_in_decimal As New DSPWave
    Dim Dsp_tfe_temp1_in_decimal As New DSPWave
    
    Dsp_tfe_temp0_in_decimal.CreateConstant 0, 1, DspDouble
     Dsp_tfe_temp1_in_decimal.CreateConstant 0, 1, DspDouble
    
'    'Test Data
'    For Each Site In TheExec.sites
'
'    DSP_fuse_read_tfe_x0_in_decimal(Site).Element(0) = 8520
'    DSP_fuse_read_tfe_y0_in_double(Site).Element(0) = 22.7031
'
'    DSP_fuse_read_tfe_vol_0_in_decimal(Site).Element(0) = -8
'    DSP_fuse_read_tfe_vol_1_in_decimal(Site).Element(0) = 131097
'
'
'    DSP_tfe_vol_x1_in_decimal(Site).Element(0) = 10025
'    DSP_tfe_y1_in_double(Site).Element(0) = 86.1
'
'
'    Next Site
'
'    ''Test data end

    For Each site In TheExec.sites
                X = 0
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfex0", CInt(X))
                 
                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_x0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfey0", CInt(X))
                  
                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_y0_in_double(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfex1", CInt(X))
                    
                TheExec.Flow.TestLimit resultVal:=DSP_tfe_vol_x1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfey1", CInt(X))
                    
                TheExec.Flow.TestLimit resultVal:=DSP_tfe_y1_in_double(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_ALG_TName_From_Instance(OutputTname_format, "C", "X", "tfevol0", CInt(X))
                
                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfevol1", CInt(X))
                
                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                
               
                
                If (DSP_fuse_read_tfe_x0_in_decimal(site).Element(0) = DSP_tfe_vol_x1_in_decimal(site).Element(0)) Or (DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0) = 0) Then
                         
                         tfe_temp0_double(site) = 178956970

                        tfe_temp1_double(site) = 178956970
                        
                            Dsp_tfe_temp0_in_decimal(site).Element(0) = FormatNumber(tfe_temp0_double(site))

                        Dsp_tfe_temp1_in_decimal(site).Element(0) = FormatNumber(tfe_temp1_double(site))
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Error_code_temp_0", CInt(X))
                    
                    TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Error_code_temp_1", CInt(X))
                        
                    TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                Else
                

                    X0(site) = ((DSP_fuse_read_tfe_x0_in_decimal(site).Element(0) - DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0)) / CDbl(DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0))) * 16
            
        X1(site) = ((DSP_tfe_vol_x1_in_decimal(site).Element(0) - DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0)) / CDbl(DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0))) * 16


        Y0(site) = 273.15 + DSP_fuse_read_tfe_y0_in_double(site).Element(0) - C2 * X0(site) * X0(site) - C3 * X0(site) * X0(site) * X0(site)


        Y1(site) = 273.15 + DSP_tfe_y1_in_double(site).Element(0) - C2 * X1(site) * X1(site) - C3 * X1(site) * X1(site) * X1(site)


        c1_cal(site) = (Y1(site) - Y0(site)) / (X1(site) - X0(site))

        c0_cal(site) = (X1(site) * Y0(site) - X0(site) * Y1(site)) / (X1(site) - X0(site))

        tfe_temp0_double(site) = (c0_cal(site) - C0) * (2 ^ 13)

        tfe_temp1_double(site) = (c1_cal(site) - c1) * (2 ^ 13)

    'tfe_temp0_long(Site) = FormatNumber(tfe_temp0_double(Site))

    'tfe_temp1_long(Site) = FormatNumber(tfe_temp1_double(Site))
                
                
                    If (tfe_temp0_double(site) > 134217727) Or (tfe_temp0_double(site) < -134217728) Then
                    

                        TestNameInput = Report_TName_From_Instance(CalcC, "X", "UpperLimit_Reached_temp_0", CInt(X))
                            
                        TheExec.Flow.TestLimit resultVal:=tfe_temp0_double(site), Tname:=TestNameInput, ForceResults:=tlForceNone
 
                                               
                        tfe_temp0_double(site) = 178956970
                                           
                        
                    
                    End If
                    If (tfe_temp1_double(site) > 134217727) Or (tfe_temp1_double(site) < -134217728) Then
                        TestNameInput = Report_TName_From_Instance(CalcC, "X", "UpperLimit_Reached_temp_1", CInt(X))
                            
                        TheExec.Flow.TestLimit resultVal:=tfe_temp1_double(site), Tname:=TestNameInput, ForceResults:=tlForceNone
                    
                      tfe_temp1_double(site) = 178956970
                    End If
                
                
                    Dsp_tfe_temp0_in_decimal(site).Element(0) = FormatNumber(tfe_temp0_double(site))

                    Dsp_tfe_temp1_in_decimal(site).Element(0) = FormatNumber(tfe_temp1_double(site))
        
        
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "X0", CInt(X))
                    
                TheExec.Flow.TestLimit resultVal:=X0(site), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "X1", CInt(X))
                   
                TheExec.Flow.TestLimit resultVal:=X1(site), Tname:=TestNameInput, ForceResults:=tlForceNone

                TestNameInput = Report_TName_From_Instance(CalcC, "X", "Y0", CInt(X))
                    
                TheExec.Flow.TestLimit resultVal:=Y0(site), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "Y1", CInt(X))
                
                TheExec.Flow.TestLimit resultVal:=Y1(site), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "c1_calc", CInt(X))
                    
                TheExec.Flow.TestLimit resultVal:=c1_cal(site), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "c0_calc", CInt(X))
                    
                TheExec.Flow.TestLimit resultVal:=c0_cal(site), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "temp_0", CInt(X))
                
                TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "temp_1", CInt(X))
                    
                TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
                
            End If

    Next site



    Dim Dsp_tfe_temp0_in_binary As New DSPWave
    Dim Dsp_tfe_temp1_in_binary As New DSPWave


    Call rundsp.DSPWf_Dec2Binary(Dsp_tfe_temp0_in_decimal, 28, Dsp_tfe_temp0_in_binary)

    Call rundsp.DSPWf_Dec2Binary(Dsp_tfe_temp1_in_decimal, 28, Dsp_tfe_temp1_in_binary)

    'Algo end

    'Store Data
    
    
    ''test dspWave
    
'    Dim test_dspWave As New DSPWave
'    test_dspWave.CreateConstant 0, 5, DspLong
'
'    For Each Site In TheExec.sites
'    test_dspWave(Site).Element(0) = 2
'    test_dspWave(Site).Element(1) = -2
'    test_dspWave(Site).Element(2) = 3
'    test_dspWave(Site).Element(3) = 4
'    test_dspWave(Site).Element(4) = 14
'    Next Site
'
'
'    Dim test_dspWave_inBinary As New DSPWave
  '  Call rundsp.DSPWf_Dec2Binary(test_dspWave, 4, test_dspWave_inBinary)
    
    ''end test

    Call AddStoredCaptureData(fuse_write_tfe_temp_0, Dsp_tfe_temp0_in_binary)
    Call AddStoredCaptureData(fuse_write_tfe_temp_1, Dsp_tfe_temp1_in_binary)



End Function



Public Function Calc_Fmax_Divide_Fmin(argc As Integer, argv() As String) As Long
    Dim site As Variant
    Dim DSP_Freq As New PinListData
    Dim Dict_Freq_Value() As New PinListData
    Dim i As Integer
    Dim Max_Temp As New PinListData
    Dim Min_Temp As New PinListData
    Dim Divide_Temp As New PinListData
    Dim ArrayNum As Integer
    Dim j As Integer
    Dim increase_flag As New SiteBoolean
    
    ArrayNum = argc - 1
    ReDim Dict_Freq_Value(ArrayNum) As New PinListData
    site = 0
    For i = 0 To argc - 1
        Dict_Freq_Value(i) = GetStoredMeasurement(argv(i))
''            ''===========Verification===========
''            For Each Site In TheExec.sites
''                    Dict_Freq_Value(i).Pins(0).Value(Site) = 1
''            Next Site
''            If i = 5 Then
''                Dict_Freq_Value(i).Pins(0).Value(0) = 30
''                Dict_Freq_Value(i).Pins(0).Value(1) = 40
''                Dict_Freq_Value(i).Pins(0).Value(2) = 50
''                Dict_Freq_Value(i).Pins(0).Value(3) = 60
''                Dict_Freq_Value(i).Pins(0).Value(4) = 70
''                Dict_Freq_Value(i).Pins(0).Value(5) = 80
''            End If
''            ''==================================
        If i = 0 Then
            Max_Temp.AddPin (Dict_Freq_Value(i).Pins(0).Name)
            Min_Temp.AddPin (Dict_Freq_Value(i).Pins(0).Name)
            Divide_Temp.AddPin (Dict_Freq_Value(i).Pins(0).Name)
        End If
        For Each site In TheExec.sites
            increase_flag(site) = True
            
            If i = 0 Then
                Max_Temp.Pins(0).Value(site) = Dict_Freq_Value(i).Pins(0).Value(site)
                Min_Temp.Pins(0).Value(site) = Dict_Freq_Value(i).Pins(0).Value(site)
            Else
                If Dict_Freq_Value(i).Pins(0).Value(site) > Max_Temp.Pins(0).Value(site) Then
                    Max_Temp.Pins(0).Value(site) = Dict_Freq_Value(i).Pins(0).Value(site)
                Else
                    increase_flag = False
                End If
                If Dict_Freq_Value(i).Pins(0).Value(site) < Min_Temp.Pins(0).Value(site) Then
                    Min_Temp.Pins(0).Value(site) = Dict_Freq_Value(i).Pins(0).Value(site)
                End If
            End If
        
        Next site
        
    
    Next i
    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            Min_Temp.Pins(0).Value(site) = 1000
            TheExec.Datalog.WriteComment ("Offline Mode Site " & site & " Calc Min Freq Simulate Meas(Denominator)=1000 Hz ")
        Next site
    End If
    For Each site In TheExec.sites
        If Min_Temp.Pins(0).Value(site) = 0 Or increase_flag(site) = False Then
        
            Divide_Temp.Pins(0).Value(site) = 999
            If Min_Temp.Pins(0).Value(site) = 0 Then
                TheExec.Datalog.WriteComment ("Error! Site " & site & " Min Freq Meas(Denominator)=0 Hz ")
            Else
                TheExec.Datalog.WriteComment ("Error! Site " & site & " Not FRO0<FRO1<FRO2....<FRO23 ")
            End If
            
        Else
            Divide_Temp.Pins(0).Value(site) = Max_Temp.Pins(0).Value(site) / Min_Temp.Pins(0).Value(site)
        End If
    Next site
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    For i = 0 To Max_Temp.Pins.Count - 1

        TestNameInput = Report_TName_From_Instance(CalcF, Max_Temp.Pins(i), "Fmax", CInt(i))
        
        TheExec.Flow.TestLimit resultVal:=Max_Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        TestNameInput = Report_TName_From_Instance(CalcF, Max_Temp.Pins(i), "Fmin", CInt(i))
        
        TheExec.Flow.TestLimit resultVal:=Min_Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        TestNameInput = Report_TName_From_Instance(CalcF, Max_Temp.Pins(i), "FmaxDivideFmin", CInt(i))
        
        TheExec.Flow.TestLimit resultVal:=Divide_Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
End Function

Public Function Calc_MTR_REL_Freq_Diff_Percentage(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim freq_Dut As String
    Dim freq_ref As String

    Dim fdiff_percent As String


    Dim DSP_fdiff_percent As New DSPWave


    DSP_fdiff_percent.CreateConstant 0, 1, DspDouble


    Dim DSP_freq_Dut As New DSPWave
    Dim DSP_freq_ref As New DSPWave



    freq_Dut = argv(0)
    freq_ref = argv(1)
    fdiff_percent = argv(2)

    Dim testName As String
    testName = "f_diff_" + freq_Dut
    DSP_freq_Dut = GetStoredCaptureData(freq_Dut)
    DSP_freq_ref = GetStoredCaptureData(freq_ref)

   
    For Each site In TheExec.sites

    If DSP_freq_Dut(site).Element(0) <> 0 Then
        DSP_fdiff_percent(site).Element(0) = ((DSP_freq_Dut(site).Element(0) - DSP_freq_ref(site).Element(0)) / DSP_freq_Dut(site).Element(0)) * 100
     
    

    Else
        DSP_fdiff_percent(site).Element(0) = 99999
         If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  freq_of_Dut " + freq_Dut + " is 0")

    End If

                TheExec.Flow.TestLimit resultVal:=DSP_fdiff_percent(site).Element(0), Tname:=testName, ForceResults:=tlForceNone
    Next site


    Call AddStoredCaptureData(fdiff_percent, DSP_fdiff_percent)



End Function

Public Function Calc_MIPI_VCMTX(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    Dim DSPWave_Binary() As New DSPWave
    ReDim DSPWave_Binary(argc - 1) As New DSPWave
    
    Dim DSPWave_Combine As New DSPWave
    DSPWave_Combine.CreateConstant 0, 10, DspLong
    
'    Dim DSPWave_Combine_verify As New DSPWave
'    DSPWave_Combine_verify.CreateConstant 0, 10, DspLong

    
    Dim DSPWave_Combine_Dec As New DSPWave
    DSPWave_Combine_Dec.CreateConstant 0, 1, DspLong
    
    Dim testName As String
    Dim site As Variant
    
    For i = 0 To argc - 2
        DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
    Next i
    
    testName = argv(argc - 1)
    
    For j = 0 To DSPWave_Combine.SampleSize - 1
        For Each site In TheExec.sites
            If j < 8 Then
                DSPWave_Combine.Element(j) = DSPWave_Binary(0).Element(j)
            Else
                DSPWave_Combine.Element(j) = DSPWave_Binary(1).Element(j - 8)
            End If
        Next site
    Next j

    Call rundsp.ConvertToLongAndSerialToParrel(DSPWave_Combine, 10, DSPWave_Combine_Dec)
    
    Dim VCMTX As New DSPWave
    VCMTX.CreateConstant 0, 1, DspDouble
    Dim VDD12_MIPI_value As Double
    VDD12_MIPI_value = TheHdw.DCVS.Pins("VDD12_MIPI").Voltage.Main.Value
    If VDD12_MIPI_value = 0 Then
            VDD12_MIPI_value = 999
            TheExec.Datalog.WriteComment ("Error! Apply VDD12_MIPI=0 V  ")
    End If
    
    For Each site In TheExec.sites
        VCMTX(site).Element(0) = DSPWave_Combine_Dec(site).Element(0) / 1024 * VDD12_MIPI_value
    Next site
'    Call rundsp.DSPWaveDecToBinary(DSPWave_Combine_Dec, 10, DSPWave_Combine_verify)
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", , 0)

    TheExec.Flow.TestLimit resultVal:=VCMTX.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
End Function


Public Function Calc_MIPID_VCMTX(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    Dim DSPWave_Binary() As New DSPWave
    ReDim DSPWave_Binary(argc - 1) As New DSPWave
    
    Dim DSPWave_Combine As New DSPWave
    DSPWave_Combine.CreateConstant 0, 10, DspLong
    
'    Dim DSPWave_Combine_verify As New DSPWave
'    DSPWave_Combine_verify.CreateConstant 0, 10, DspLong

    
    Dim DSPWave_Combine_Dec As New DSPWave
    DSPWave_Combine_Dec.CreateConstant 0, 1, DspLong
    
    Dim testName As String
    Dim site As Variant
    
    For i = 0 To argc - 1 '20190523 CWCIOU
        DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
    Next i
    
    testName = argv(argc - 1)
    
    For j = 0 To DSPWave_Combine.SampleSize - 1
        For Each site In TheExec.sites
            If j < 8 Then
                DSPWave_Combine.Element(j) = DSPWave_Binary(0).Element(j)
            Else
                DSPWave_Combine.Element(j) = DSPWave_Binary(1).Element(j - 8)
            End If
        Next site
    Next j

    Call rundsp.ConvertToLongAndSerialToParrel(DSPWave_Combine, 10, DSPWave_Combine_Dec)
    
    Dim VCMTX As New DSPWave
    VCMTX.CreateConstant 0, 1, DspDouble
    Dim VDD18_MIPID_value As Double
    VDD18_MIPID_value = TheHdw.DCVS.Pins("VDD18_MIPID").Voltage.Main.Value '20190523 CWCIOU
    If VDD18_MIPID_value = 0 Then
            VDD18_MIPID_value = 999
            TheExec.Datalog.WriteComment ("Error! Apply VDD18_MIPID=0 V  ")
    End If
    
    For Each site In TheExec.sites
        VCMTX(site).Element(0) = DSPWave_Combine_Dec(site).Element(0) / 1024 * VDD18_MIPID_value
    Next site
'    Call rundsp.DSPWaveDecToBinary(DSPWave_Combine_Dec, 10, DSPWave_Combine_verify)
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0)
    'TestNameInput = Report_TName_From_Instance("V", "X", , 0)

    TheExec.Flow.TestLimit resultVal:=VCMTX.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
End Function

Public Function Calc_DigCapCombine(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long

    Dim DSPWave_Binary() As New DSPWave
    ReDim DSPWave_Binary(argc - 1) As New DSPWave
    
    Dim DSPWave_Combine As New DSPWave
    DSPWave_Combine.CreateConstant 0, 10, DspLong
    
'    Dim DSPWave_Combine_verify As New DSPWave
'    DSPWave_Combine_verify.CreateConstant 0, 10, DspLong

    
    Dim DSPWave_Combine_Dec As New DSPWave
    DSPWave_Combine_Dec.CreateConstant 0, 1, DspLong
    
    Dim TestNameInput As String
    Dim site As Variant
    
    For i = 0 To argc - 1
        DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
    Next i
    

    For j = 0 To DSPWave_Combine.SampleSize - 1
        For Each site In TheExec.sites
            If j < 8 Then
                DSPWave_Combine.Element(j) = DSPWave_Binary(0).Element(j)
            Else
                DSPWave_Combine.Element(j) = DSPWave_Binary(1).Element(j - 8)
            End If
        Next site
    Next j

    Call rundsp.ConvertToLongAndSerialToParrel(DSPWave_Combine, 10, DSPWave_Combine_Dec)
'    Call rundsp.DSPWaveDecToBinary(DSPWave_Combine_Dec, 10, DSPWave_Combine_verify)
    
    
    Dim OutputTname_format() As String

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "DEC" & i, CInt(i))
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_Combine_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
End Function
Public Function Calc_VDiff_t6p1_metrologyGR(argc As Integer, argv() As String) As Long
    Dim Dict_V2 As String
    Dim Dict_V1 As String
    Dim testName As String
    Dim Input_V1 As New PinListData
    Dim Input_V2 As New PinListData
    Dim result As New DSPWave
    Dim CalcVal As New PinListData
    Dim DummyPinListData As New PinListData
    Dim site As Variant
    Dim X As Integer
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    
    X = 0
   
    
    result.CreateConstant 0, 1, DspDouble

 
    
    Dict_V1 = argv(0)
    Dict_V2 = argv(1)
    testName = argv(2)
    Input_V1 = GetStoredMeasurement(Dict_V1)
      Input_V2 = GetStoredMeasurement(Dict_V2)
      
      
    DummyPinListData.AddPin (Input_V1.Pins(0))
      DummyPinListData = Input_V1.Pins(0).Subtract(Input_V2.Pins(0)).Abs
      
      
      
      
      For Each site In TheExec.sites
        result(site).Element(0) = DummyPinListData.Pins(0).Value
      Next site
      


'    CalcVal.AddPin (InputVal.Pins(0))
'    CalcVal = InputVal.Pins(0).Subtract(0.4).Divide(0.7975).Subtract(1)
    
         If Not ByPassTestLimit Then
            TestNameInput = Report_TName_From_Instance(CalcV, "X", , CInt(X))
            TheExec.Flow.TestLimit resultVal:=result.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If
'    Call AddStoredMeasurement(Dict_ReturnKey, CalcVal)
End Function
Public Function Calc_DigCapAvg(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long

    Dim DSPWave_Binary() As New DSPWave
    ReDim DSPWave_Binary(argc - 4) As New DSPWave
    
    Dim DSPWave_Dec() As New DSPWave
    ReDim DSPWave_Dec(argc - 4) As New DSPWave
    
    Dim DSPWave_Avg_Dec As New DSPWave
    DSPWave_Avg_Dec.CreateConstant 0, 1, DspLong
    'ReDim DSPWave_Avg(argc - 1) As New DSPWave
    
    Dim DSPWave_Avg_Bin As New DSPWave
    'ReDim DSPWave_Avg_Bin(argc - 3) As New DSPWave
    
    Dim testName As String
    Dim site As Variant
    Dim Dict As String
    Dim BitWidth As Long
    
    For i = 0 To 1
        DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
        Call rundsp.BinToDec(DSPWave_Binary(i), DSPWave_Dec(i))
    Next i
    
    testName = argv(argc - 1)
    BitWidth = argv(argc - 2)
    Dict = argv(argc - 3)
    
    For Each site In TheExec.sites
            DSPWave_Avg_Dec.Element(0) = Int(((DSPWave_Dec(0).Element(0) + DSPWave_Dec(1).Element(0)) / 2) + 0.5) ''Example 1). 78.4=>78  2). 78.5=79
    Next site
    Call rundsp.DSPWaveDecToBinary(DSPWave_Avg_Dec, BitWidth, DSPWave_Avg_Bin)
    Call AddStoredCaptureData(Dict, DSPWave_Avg_Bin)
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
    
    TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
    TheExec.Flow.TestLimit resultVal:=DSPWave_Avg_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
    
End Function
Public Function Calc_CalR_FVMI_IO(argc As Integer, argv() As String) As Long

    Dim StoredCurrent As New PinListData
    Dim CalR As New PinListData
    Dim ForceVoltVal As Double
    Dim PowerPinName As String
    
    Dim i, p As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim site As Variant
    Dim Pin  As Variant
    Dim Lowlimitval_temp As Double
    Dim Hilimitval_temp As Double
        
    PowerPinName = argv(1)
    StoredCurrent = GetStoredMeasurement(argv(0))
    ForceVoltVal = argv(2)
    
    For Each Pin In StoredCurrent.Pins
        For Each site In TheExec.sites
            If StoredCurrent.Pins(Pin).Value(site) = 0 Then
                StoredCurrent.Pins(Pin).Value(site) = 0.000000000001
            End If
        Next site
    Next Pin

    
    CalR = StoredCurrent.Math.Invert.Multiply(ForceVoltVal).Abs
          
    '===============RAK read
    Dim GetRakVal As New PinListData
    GetRakVal = CurrentJob_Card_RAK
       
            For Each site In TheExec.sites
                GetRakVal = CurrentJob_Card_RAK.Pins(PowerPinName).Value(site)
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment PowerPinName & " = " & CalR.Pins.Item(PowerPinName).Value(site) & ", RAK val = " & GetRakVal.Pins(PowerPinName).Value
                CalR.Pins.Item(PowerPinName).Value(site) = CalR.Pins.Item(PowerPinName).Value(site) - GetRakVal.Pins(PowerPinName).Value
            Next site
    
        For p = 0 To CalR.Pins.Count - 1
            If LCase(CalR.Pins.Item(p).Name) Like LCase((PowerPinName)) Then
                    TestNameInput = Report_TName_From_Instance("R", CalR.Pins(p), , CInt(p))
                    Hilimitval_temp = 96
                    Lowlimitval_temp = 64
                    TheExec.Flow.TestLimit CalR.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
            End If
        Next p
    
'
End Function
Public Function Calc_CalR_FVMI(argc As Integer, argv() As String) As Long

    Dim StoredCurrent As New PinListData
    Dim CalR As New PinListData
    Dim ForceVoltVal As Double
    Dim PowerPinName As String
    
    Dim i, p As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim site As Variant
    Dim Pin  As Variant
    Dim Lowlimitval_temp As Double
    Dim Hilimitval_temp As Double
        
    PowerPinName = argv(1)
    StoredCurrent = GetStoredMeasurement(argv(0))
    ForceVoltVal = TheHdw.DCVS.Pins(PowerPinName).Voltage.Value
    
    For Each Pin In StoredCurrent.Pins
        For Each site In TheExec.sites
            If StoredCurrent.Pins(Pin).Value(site) = 0 Then
                StoredCurrent.Pins(Pin).Value(site) = 0.000000000001
            End If
        Next site
    Next Pin
    
    CalR = StoredCurrent.Math.Invert.Multiply(ForceVoltVal).Abs
'
'    Dim RakV() As Double
'    Dim GetRakVal As Double
'    Dim RAK_Pin As String
'
'    Dim PinGetRakVal As New PinListData
'    Set PinGetRakVal = Nothing
'
'    For Each Pin In CalR.Pins
'        PinGetRakVal.AddPin CStr(Pin)
'        For Each site In TheExec.sites
'            RAK_Pin = CStr(Pin)
'            RakV = TheHdw.PPMU.ReadRakValuesByPinnames(RAK_Pin, site)
'
'            GetRakVal = RakV(0) + CurrentJob_Card_RAK.Pins(Pin).Value(site)
'            PinGetRakVal.Pins(Pin).Value = GetRakVal
'                If argc <= 2 And gl_Disable_HIP_debug_log = False Then
'                    TheExec.DataLog.WriteComment Pin & " = " & CalR.Pins.Item(Pin).Value(site) & ", RAK val = " & GetRakVal
'                ElseIf argv(2) <> "TTR" And gl_Disable_HIP_debug_log = False Then
'                    TheExec.DataLog.WriteComment Pin & " = " & CalR.Pins.Item(Pin).Value(site) & ", RAK val = " & GetRakVal
'                End If
'            CalR.Pins.Item(Pin).Value(site) = CalR.Pins.Item(Pin).Value(site) - GetRakVal
'        Next site
'    Next Pin
'
    If argc <= 2 Then
    Dim Temp_index As Long
    Temp_index = TheExec.Flow.TestLimitIndex
        For i = 0 To CalR.Pins.Count - 1
            TheExec.Flow.TestLimitIndex = Temp_index
            TestNameInput = Report_TName_From_Instance(CalcR, CalR.Pins(i), , CInt(i))
'            If i = 0 Then
                TheExec.Flow.TestLimit CalR.Pins(i), , , , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
'            Else
'                TheExec.Flow.TestLimit CalR.Pins(i), GetLowLimitFromFlow, GetHiLimitFromFlow, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceNone
'            End If
        Next i
    ElseIf argv(2) = "TTR" Then
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 2   'use second test limit spec for R,first test limit for current(don't need)

        Lowlimitval_temp = GetLowLimitFromFlow
        Hilimitval_temp = GetHiLimitFromFlow
         If TheExec.EnableWord("HIP_TTR_FailResultOnly") = True Then
        For Each site In TheExec.sites.Active
            For p = 0 To CalR.Pins.Count - 1
                If CalR.Pins(p).Value > Hilimitval_temp Or CalR.Pins(p).Value < Lowlimitval_temp Then
                    TestNameInput = Report_TName_From_Instance(CalcR, CalR.Pins(p), , CInt(p))
                    
                    'TheExec.Flow.TestLimit StoredCurrent.Pins(p), , , , , , unitAmp, , , ForceResults:=tlForceNone
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment CalR.Pins(p) & " = " & CalR.Pins(p).Value(site)
                    TheExec.Flow.TestLimit CalR.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceNone

                    
                End If
            Next p
        Next site
        Else
            For p = 0 To CalR.Pins.Count - 1

                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment CalR.Pins(p) & " = " & CalR.Pins(p).Value(site)
                    TestNameInput = Report_TName_From_Instance(CalcR, CalR.Pins(p), , CInt(p))
                    TheExec.Flow.TestLimit CalR.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceNone
            Next p
    End If
    End If
    
End Function

Public Function Calc_CalZ_FVMI(argc As Integer, argv() As String) As Long

    Dim StoredCurrent As New PinListData
    Dim StoredCurrent_I2 As New PinListData
    Dim StoredCurrent_I1 As New PinListData
    Dim CalR As New PinListData
    Dim ForceVoltVal As Double
    Dim PowerPinName As String
    
    Dim i, p As Long
    Dim TestNameInput As String
    Dim site As Variant
    Dim Pin  As Variant

        
        'argv() :V2,V1,I2,I1
        

    StoredCurrent_I1 = GetStoredMeasurement(argv(3))
    StoredCurrent_I2 = GetStoredMeasurement(argv(2))
  
    ForceVoltVal = argv(0) - argv(1)
    

    StoredCurrent = StoredCurrent_I2.Math.Subtract(StoredCurrent_I1)
    
        For Each Pin In StoredCurrent.Pins ' To prevent i=0
            For Each site In TheExec.sites
                If StoredCurrent.Pins(Pin).Value(site) = 0 Then
                    StoredCurrent.Pins(Pin).Value(site) = 0.000000000001
                End If
            Next site
        Next Pin
    
        CalR = StoredCurrent.Math.Invert.Multiply(ForceVoltVal).Abs

        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment CalR.Pins(p) & " = " & CalR.Pins(p).Value(site)
        TestNameInput = Report_TName_From_Instance(CalcR, "X", "", 0)
        TheExec.Flow.TestLimit CalR, , , , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow


End Function


Public Function Calc_MDLL_Monotonicity_DevideBlock_TTR(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long, X As Long
    Dim site As Variant
    
    Dim DSPWaveDec() As New DSPWave
    ReDim DSPWaveDec((argc - 1) * 2 - 1) As New DSPWave
    Dim testName As String
    Dim DDR_MonoWithblock() As Type_MonoWithBlock
    ReDim DDR_MonoWithblock((5 - 1) * 2 - 1) As Type_MonoWithBlock
    Dim DSP_Input As New DSPWave
    Dim DSP_Input_Update As New DSPWave
    Dim DSP_Input_Final As New DSPWave
    Dim DSP_Input_UpperBIN_1 As New DSPWave
    Dim DSP_Input_BelowBIN_1 As New DSPWave
    Dim DSP_Input_UpperDEC_1 As New DSPWave
    Dim DSP_Input_BelowDEC_1 As New DSPWave
    Dim DSP_Input_UpperBIN_2 As New DSPWave
    Dim DSP_Input_BelowBIN_2 As New DSPWave
    Dim DSP_Input_UpperDEC_2 As New DSPWave
    Dim DSP_Input_BelowDEC_2 As New DSPWave
    Dim InputKey As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    '' Merge all the DSP data to DSP_Input_Update and send it to rundsp SeprateDSP for only one time
    For i = 0 To argc - 1
        If i Mod 3 <> 0 Then
            InputKey = LCase(argv(i))
            Set DSP_Input = Nothing
            DSP_Input = GetStoredCaptureData(InputKey)
            
            If i = 1 Then
                DSP_Input_Update = DSP_Input
            Else
                For Each site In TheExec.sites
                    DSP_Input_Update = DSP_Input_Update.Concatenate(DSP_Input)
                Next site
            End If
        
        End If
     Next i
     
    Call rundsp.SeprateDSP_TTR(DSP_Input_Update, DSP_Input_UpperBIN_1, DSP_Input_UpperBIN_2, DSP_Input_BelowBIN_1, DSP_Input_BelowBIN_2, DSP_Input_UpperDEC_1, DSP_Input_UpperDEC_2, DSP_Input_BelowDEC_1, DSP_Input_BelowDEC_2)
    
    Dim Temp_DSP_Input_UpperBIN_1 As New DSPWave
    Dim Temp_DSP_Input_UpperBIN_2 As New DSPWave
    Dim Temp_DSP_Input_UpperDEC_1 As New DSPWave
    Dim Temp_DSP_Input_UpperDEC_2 As New DSPWave
    Dim Temp_DSP_Input_BelowBIN_1 As New DSPWave
    Dim Temp_DSP_Input_BelowBIN_2 As New DSPWave
    Dim Temp_DSP_Input_BelowDEC_1 As New DSPWave
    Dim Temp_DSP_Input_BelowDEC_2 As New DSPWave
    
    Temp_DSP_Input_UpperBIN_1.CreateConstant 0, 8, DspLong
    Temp_DSP_Input_UpperBIN_2.CreateConstant 0, 8, DspLong
    Temp_DSP_Input_UpperDEC_1.CreateConstant 0, 1, DspLong
    Temp_DSP_Input_UpperDEC_2.CreateConstant 0, 1, DspLong
    Temp_DSP_Input_BelowBIN_1.CreateConstant 0, 8, DspLong
    Temp_DSP_Input_BelowBIN_2.CreateConstant 0, 8, DspLong
    Temp_DSP_Input_BelowDEC_1.CreateConstant 0, 1, DspLong
    Temp_DSP_Input_BelowDEC_2.CreateConstant 0, 1, DspLong
    
    For X = 0 To 7 'wc 1220
        DDR_MonoWithblock(X).DSP_Bin.CreateConstant 0, 8
        DDR_MonoWithblock(X).DSP_Dec.CreateConstant 0, 1
    Next X
    
    Dim arg As Long: arg = 0
    Dim counter As Long: counter = 0
    
    For arg = 0 To argc - 1 Step 3
        
        For i = 0 To 3
           testName = argv(arg) & "_"
            InputKey = LCase(argv(i + 1))
    
            If InStr(InputKey, LCase("dll_code_l")) <> 0 Then
               
                Call Sub_MDLL(DSP_Input_UpperBIN_1, DSP_Input_UpperBIN_2, DSP_Input_BelowBIN_1, DSP_Input_BelowBIN_2, DSP_Input_UpperDEC_1, DSP_Input_UpperDEC_2, DSP_Input_BelowDEC_1, DSP_Input_BelowDEC_2, _
                             DDR_MonoWithblock(4).DSP_Bin, DDR_MonoWithblock(0).DSP_Bin, DDR_MonoWithblock(6).DSP_Bin, DDR_MonoWithblock(1).DSP_Bin, DDR_MonoWithblock(4).DSP_Dec, DDR_MonoWithblock(0).DSP_Dec, DDR_MonoWithblock(6).DSP_Dec, DDR_MonoWithblock(1).DSP_Dec, _
                             0 + 16 * counter, 7 + 16 * counter, 0 + 2 * counter) '1222 wc
                 DDR_MonoWithblock(4).Block = 4
                 DDR_MonoWithblock(0).Block = 0
                 DDR_MonoWithblock(6).Block = 6
                 DDR_MonoWithblock(1).Block = 1
                 
            ElseIf InStr(InputKey, LCase("dll_code_m")) <> 0 Then

                Call Sub_MDLL(DSP_Input_UpperBIN_1, DSP_Input_UpperBIN_2, DSP_Input_BelowBIN_1, DSP_Input_BelowBIN_2, DSP_Input_UpperDEC_1, DSP_Input_UpperDEC_2, DSP_Input_BelowDEC_1, DSP_Input_BelowDEC_2, _
                             DDR_MonoWithblock(3).DSP_Bin, DDR_MonoWithblock(7).DSP_Bin, DDR_MonoWithblock(2).DSP_Bin, DDR_MonoWithblock(5).DSP_Bin, DDR_MonoWithblock(3).DSP_Dec, DDR_MonoWithblock(7).DSP_Dec, DDR_MonoWithblock(2).DSP_Dec, DDR_MonoWithblock(5).DSP_Dec, _
                             8 + 16 * counter, 15 + 16 * counter, 1 + 2 * counter) '1222 wc

                 DDR_MonoWithblock(3).Block = 3
                 DDR_MonoWithblock(7).Block = 7
                 DDR_MonoWithblock(2).Block = 2
                 DDR_MonoWithblock(5).Block = 5

            End If
        Next i

'''''        Dim dataStr As String
'''''        For Each Site In TheExec.sites
'''''            For i = 0 To UBound(DDR_MonoWithblock)
'''''                dataStr = ""
'''''                For j = 0 To DDR_MonoWithblock(i).DSP_Bin.SampleSize - 1
'''''                    If j = 0 Then
'''''                        dataStr = DDR_MonoWithblock(i).DSP_Bin(Site).Element(j)
'''''                    Else
'''''                        dataStr = dataStr & DDR_MonoWithblock(i).DSP_Bin(Site).Element(j)
'''''                    End If
'''''                Next j
'''''                TheExec.Datalog.WriteComment ("Site_" & Site & " , Block = " & DDR_MonoWithblock(i).Block & " , Binary = " & dataStr & " , Decimal = " & DDR_MonoWithblock(i).DSP_Dec.Element(0))
'''''            Next i
'''''        Next Site
    
        '' 20170713 - Sorting DDR_MonoWithblock by block
        Dim TempBlock As Long
        Dim sd_TempDSP_BIN As New DSPWave
        Dim sd_TempDSP_DEC As New DSPWave
'        For i = 0 To UBound(DDR_MonoWithblock)
'            For j = i To UBound(DDR_MonoWithblock)
'                If DDR_MonoWithblock(i).Block > DDR_MonoWithblock(j).Block Then
'                    TempBlock = DDR_MonoWithblock(i).Block
'                    DDR_MonoWithblock(i).Block = DDR_MonoWithblock(j).Block
'                    DDR_MonoWithblock(j).Block = TempBlock
'
'                    sd_TempDSP_BIN = DDR_MonoWithblock(i).DSP_Bin
'                    DDR_MonoWithblock(i).DSP_Bin = DDR_MonoWithblock(j).DSP_Bin
'                    DDR_MonoWithblock(j).DSP_Bin = sd_TempDSP_BIN
'
'                    sd_TempDSP_DEC = DDR_MonoWithblock(i).DSP_Dec
'                    DDR_MonoWithblock(i).DSP_Dec = DDR_MonoWithblock(j).DSP_Dec
'                    DDR_MonoWithblock(j).DSP_Dec = sd_TempDSP_DEC
'                End If
'            Next j
'        Next i
    
'''''        '' Print info after sorting
'''''        TheExec.Datalog.WriteComment ("Print info after sorting")
'''''        For Each Site In TheExec.sites
'''''            For i = 0 To UBound(DDR_MonoWithblock)
'''''                dataStr = ""
'''''                For j = 0 To DDR_MonoWithblock(i).DSP_Bin.SampleSize - 1
'''''                    If j = 0 Then
'''''                        dataStr = DDR_MonoWithblock(i).DSP_Bin(Site).Element(j)
'''''                    Else
'''''                        dataStr = dataStr & DDR_MonoWithblock(i).DSP_Bin(Site).Element(j)
'''''                    End If
'''''                Next j
'''''                TheExec.Datalog.WriteComment ("Site_" & Site & " , Block = " & DDR_MonoWithblock(i).Block & " , Binary = " & dataStr & " , Decimal = " & DDR_MonoWithblock(i).DSP_Dec.Element(0))
'''''            Next i
'''''        Next Site
    
        For i = 0 To UBound(DDR_MonoWithblock)
            DSPWaveDec(i) = DDR_MonoWithblock(i).DSP_Dec
        Next i
    
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("")
        'For Each site In TheExec.sites
         For i = 0 To UBound(DDR_MonoWithblock)  'NEW 20170730
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(counter))
            'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            For Each site In TheExec.sites
                'TheExec.Flow.TestLimit resultVal:=DSPWaveDec(i)(Site).Element(0), LowVal:=0, HiVal:=119, Tname:=TestName & "Block_" & DDR_MonoWithblock(i).Block & "_Lock_Code_Range", ForceResults:=tlForceNone
                TheExec.Flow.TestLimit resultVal:=DSPWaveDec(i)(site).Element(0), lowVal:=0, hiVal:=119, Tname:=TestNameInput, ForceResults:=tlForceNone
            Next site
        Next i
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("")
        'Next site
    
        Dim MDLL_CurrentVal As New SiteLong
        Dim MDLL_PreviousVal  As New SiteLong
        Dim b_MDLL_DecreaseDirection As New SiteBoolean
        Dim b_MDLL_DecreaseAddIndex As New SiteBoolean
        Dim MDLL_DecreaseResultPass As New SiteLong
        Dim b_MDLL_TestResultFail As New SiteBoolean
        Dim MDLL_Index As New SiteLong
        b_MDLL_DecreaseDirection = False
    
        MDLL_DecreaseResultPass = 1
        b_MDLL_TestResultFail = False
        MDLL_Index = 1
        Dim StepSize As Long
        Dim StoreDecreaseVal As New SiteVariant
        Dim StoreDecreaseIndex As Long
        StoreDecreaseIndex = 0
        For Each site In TheExec.sites

            For i = 0 To UBound(DDR_MonoWithblock)  'NEW 20170730
                If i = 0 Then
                    MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                    MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
    
                    StoreDecreaseVal(site) = CStr(MDLL_CurrentVal(site))
                    StoreDecreaseIndex = StoreDecreaseIndex + 1
                Else
                    MDLL_CurrentVal(site) = DSPWaveDec(i)(site).Element(0)
                    b_MDLL_DecreaseDirection(site) = MDLL_CurrentVal.Subtract(MDLL_PreviousVal).compare(LessThanOrEqualTo, 0)
    
                    If b_MDLL_DecreaseDirection(site) = False Then
                        MDLL_DecreaseResultPass(site) = 0
                    End If

                    StoreDecreaseVal(site) = StoreDecreaseVal(site) & "," & MDLL_CurrentVal(site)
                    StoreDecreaseIndex = StoreDecreaseIndex + 1
   
                    MDLL_PreviousVal(site) = MDLL_CurrentVal(site)
                End If
            Next i
        Next site
    
        Dim OriginalVal() As String
        Dim TempVal As Double
        Dim SortedVal() As Double
    
        Dim DiffVal_Num As New SiteLong
        Dim DiffVal_MaxSubMin As New SiteLong
        DiffVal_Num = 1
    
        For Each site In TheExec.sites
            OriginalVal = Split(StoreDecreaseVal(site), ",")
            ReDim SortedVal(UBound(OriginalVal)) As Double
            For i = 0 To UBound(OriginalVal)
                SortedVal(i) = CDbl(OriginalVal(i))
            Next i
            For i = 0 To UBound(SortedVal)
                For j = i To UBound(SortedVal)
                    If SortedVal(i) > SortedVal(j) Then
                        TempVal = SortedVal(i)
                        SortedVal(i) = SortedVal(j)
                        SortedVal(j) = TempVal
                    End If
                Next j
            Next i
            For i = 0 To UBound(SortedVal) - 1
                If SortedVal(i + 1) - SortedVal(i) > 0 Then
                    DiffVal_Num(site) = DiffVal_Num(site) + 1
                End If
            Next i
            DiffVal_MaxSubMin(site) = SortedVal(UBound(SortedVal)) - SortedVal(0)
        Next site
        
        
'''''''''' Edited by Dylan 2019/12/04
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "Decrease", CInt(X))
            
        TheExec.Flow.TestLimit resultVal:=MDLL_DecreaseResultPass, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceNone
        
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "Unique", CInt(X))
        
        TheExec.Flow.TestLimit resultVal:=DiffVal_Num, lowVal:=1, hiVal:=2, Tname:=TestNameInput, ForceResults:=tlForceNone
        
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "MaxDiff", CInt(X))
            
        TheExec.Flow.TestLimit resultVal:=DiffVal_MaxSubMin, lowVal:=0, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceNone
''''''''''

        counter = counter + 1
    Next arg

End Function

Public Function Sub_MDLL(DSP_Input_UpperBIN_1 As DSPWave, DSP_Input_UpperBIN_2 As DSPWave, DSP_Input_BelowBIN_1 As DSPWave, DSP_Input_BelowBIN_2 As DSPWave, DSP_Input_UpperDEC_1 As DSPWave, DSP_Input_UpperDEC_2 As DSPWave, DSP_Input_BelowDEC_1 As DSPWave, DSP_Input_BelowDEC_2 As DSPWave, _
ByRef Temp_DSP_Input_UpperBIN_1 As DSPWave, ByRef Temp_DSP_Input_UpperBIN_2 As DSPWave, ByRef Temp_DSP_Input_BelowBIN_1 As DSPWave, ByRef Temp_DSP_Input_BelowBIN_2 As DSPWave, ByRef Temp_DSP_Input_UpperDEC_1 As DSPWave, ByRef Temp_DSP_Input_UpperDEC_2 As DSPWave, ByRef Temp_DSP_Input_BelowDEC_1 As DSPWave, ByRef Temp_DSP_Input_BelowDEC_2 As DSPWave, _
Binary_Start As Long, Binary_End As Long, Dec_data As Long) As Long

    Dim i As Long: i = 0
    Dim j  As Long: j = 0
    Dim site As Variant
    For Each site In TheExec.sites
        Temp_DSP_Input_UpperDEC_1(site).Element(0) = DSP_Input_UpperDEC_1(site).Element(Dec_data)
        Temp_DSP_Input_UpperDEC_2(site).Element(0) = DSP_Input_UpperDEC_2(site).Element(Dec_data)
        Temp_DSP_Input_BelowDEC_1(site).Element(0) = DSP_Input_BelowDEC_1(site).Element(Dec_data)
        Temp_DSP_Input_BelowDEC_2(site).Element(0) = DSP_Input_BelowDEC_2(site).Element(Dec_data)
        
        For i = Binary_Start To Binary_End
            Temp_DSP_Input_UpperBIN_1(site).Element(j) = DSP_Input_UpperBIN_1(site).Element(i)
            Temp_DSP_Input_UpperBIN_2(site).Element(j) = DSP_Input_UpperBIN_2(site).Element(i)
            Temp_DSP_Input_BelowBIN_1(site).Element(j) = DSP_Input_BelowBIN_1(site).Element(i)
            Temp_DSP_Input_BelowBIN_2(site).Element(j) = DSP_Input_BelowBIN_2(site).Element(i)
            j = j + 1
        Next i
        j = 0
    Next site
    
End Function

Public Function Calc_FromLoad_MTR_SE_CAL_Coeff(SensorTempName_rot As String, SensorTempName_rov As String, ByVal Temperature As Long, ByVal FuseSize_1 As Long, ByVal FuseSize_2 As Long, ByRef DSPWave_Coeff_1 As DSPWave, ByRef DSPWave_Coeff_2 As DSPWave, ByRef OutDspWaveToFuse_1 As DSPWave, ByRef OutDspWaveToFuse_2 As DSPWave, MTR_CAL_Sheet As Long) As Long

    Dim site As Variant
    Dim piU1(3, 7) As Double
    Dim piU2(2, 7) As Double
    Dim piU3(3, 7) As Double
    Dim piU4(2, 7) As Double
    Dim a1 As New DSPWave
    Dim a2 As New DSPWave
    Dim a3 As New DSPWave
    Dim a4 As New DSPWave
    a1.CreateConstant 0, 4, DspDouble
    a2.CreateConstant 0, 3, DspDouble
    a3.CreateConstant 0, 4, DspDouble
    a4.CreateConstant 0, 3, DspDouble
    Dim a1_max(3) As Double
    Dim a2_max(2) As Double
    Dim a3_max(3) As Double
    Dim a4_max(2) As Double
    Dim a1_min(3) As Double
    Dim a2_min(2) As Double
    Dim a3_min(3) As Double
    Dim a4_min(2) As Double
    
    Dim row As Long
    Dim col As Long


    
    Dim MTRMatricesSheet As Worksheet
    If MTR_CAL_Sheet = 0 Then
        Set MTRMatricesSheet = Sheets("MTR_CAL_matrices_Group1")
    For row = 2 To 5
            For col = 1 To 7
            piU1(row - 2, col - 1) = MTRMatricesSheet.Cells(row, col)
        Next col
    Next row
    
    For row = 7 To 9
            For col = 1 To 7
            piU2(row - 7, col - 1) = MTRMatricesSheet.Cells(row, col)
        Next col
    Next row
    
    For row = 11 To 14
            For col = 1 To 7
            piU3(row - 11, col - 1) = MTRMatricesSheet.Cells(row, col)
        Next col
    Next row
    
    For row = 16 To 18
            For col = 1 To 7
            piU4(row - 16, col - 1) = MTRMatricesSheet.Cells(row, col)
        Next col
    Next row
    Else
        Set MTRMatricesSheet = Sheets("MTR_CAL_matrices_Group2")
        For row = 2 To 5
            For col = 1 To 5
                piU1(row - 2, col - 1) = MTRMatricesSheet.Cells(row, col)
            Next col
        Next row
    
        For row = 7 To 9
            For col = 1 To 5
                piU2(row - 7, col - 1) = MTRMatricesSheet.Cells(row, col)
            Next col
        Next row
    
        For row = 11 To 14
            For col = 1 To 5
                piU3(row - 11, col - 1) = MTRMatricesSheet.Cells(row, col)
            Next col
        Next row
    
        For row = 16 To 18
            For col = 1 To 5
                piU4(row - 16, col - 1) = MTRMatricesSheet.Cells(row, col)
            Next col
        Next row
    End If
    
    For col = 2 To 5
        a1_max(col - 2) = MTRMatricesSheet.Cells(20, col)
    Next col
        For col = 2 To 5
        a1_min(col - 2) = MTRMatricesSheet.Cells(21, col)
    Next col
        For col = 2 To 4
        a2_max(col - 2) = MTRMatricesSheet.Cells(22, col)
    Next col
        For col = 2 To 4
        a2_min(col - 2) = MTRMatricesSheet.Cells(23, col)
    Next col
    For col = 2 To 5
        a3_max(col - 2) = MTRMatricesSheet.Cells(24, col)
    Next col
    For col = 2 To 5
        a3_min(col - 2) = MTRMatricesSheet.Cells(25, col)
    Next col
    For col = 2 To 4
        a4_max(col - 2) = MTRMatricesSheet.Cells(26, col)
    Next col
    For col = 2 To 4
        a4_min(col - 2) = MTRMatricesSheet.Cells(27, col)
    Next col
    
    
    
    
    
    
    
    Dim temp_rowVal_a1 As New SiteDouble
    Dim temp_rowVal_a2 As New SiteDouble
    Dim temp_rowVal_a3 As New SiteDouble
    Dim temp_rowVal_a4 As New SiteDouble
    Dim testName As String
    Dim currBinaryStr As String
    Dim totalBinaryStr As String
    Dim currElementDspWave As Long
'    Dim OutDspWaveToFuse As New DSPWave
'    OutDspWaveToFuse.CreateConstant 0, FuseSize, DspLong


    Dim decimalPlaces As Long
    decimalPlaces = 8
    
    Dim DSPWave_Matrix_rot As New DSPWave
    Dim DSPWave_Matrix_rov As New DSPWave
    DSPWave_Matrix_rot = GetStoredCaptureData(SensorTempName_rot)
    DSPWave_Matrix_rov = GetStoredCaptureData(SensorTempName_rov)
    
    If Temperature = 25 Then

        For Each site In TheExec.sites
            totalBinaryStr = ""
            For row = 0 To 3
                currBinaryStr = ""
                temp_rowVal_a1(site) = 0

                If MTR_CAL_Sheet = 0 Then
                    For col = 0 To 6
                    temp_rowVal_a1(site) = temp_rowVal_a1(site) + piU1(row, col) * DSPWave_Matrix_rot(site).Element(col)
                Next col
                Else
                    For col = 0 To 4
                        temp_rowVal_a1(site) = temp_rowVal_a1(site) + piU1(row, col) * DSPWave_Matrix_rot(site).Element(col)
                    Next col
                End If

                a1(site).Element(row) = (temp_rowVal_a1(site) - a1_min(row)) / (a1_max(row) - a1_min(row))
                temp_rowVal_a1 = 0
                If (row = 0) Then
                    Call MTR_Cal_DecimalToBinary(a1(site).Element(row), 15, decimalPlaces, currBinaryStr)
                Else
                    Call MTR_Cal_DecimalToBinary(a1(site).Element(row), 14, decimalPlaces, currBinaryStr)
                End If
                totalBinaryStr = totalBinaryStr + currBinaryStr

                'Added on 20180131 To Force Error
                If (a1(site).Element(row) = 0) Then
                    a1(site).Element(row) = -0.000001
                ElseIf (a1(site).Element(row) = 1) Then
                    a1(site).Element(row) = 1.000001
                End If
            Next row
            currElementDspWave = 0
            TheExec.Datalog.WriteComment ("Fuse Binary Str  a1 for Site:" + CStr(site) + " is " + totalBinaryStr)
            totalBinaryStr = StrReverse(totalBinaryStr)
            If Len(totalBinaryStr) = OutDspWaveToFuse_1.SampleSize Then
                Do While currElementDspWave < FuseSize_1
                    OutDspWaveToFuse_1(site).Element(currElementDspWave) = CInt(Mid(totalBinaryStr, currElementDspWave + 1, 1))
                    currElementDspWave = currElementDspWave + 1
                Loop
            End If
            
            
            totalBinaryStr = ""
            For row = 0 To 2
                currBinaryStr = ""
                temp_rowVal_a2(site) = 0
                    
                If MTR_CAL_Sheet = 0 Then
                    For col = 0 To 6
                    temp_rowVal_a2(site) = temp_rowVal_a2(site) + piU2(row, col) * DSPWave_Matrix_rov(site).Element(col)
                Next col
                Else
                    For col = 0 To 4
                        temp_rowVal_a2(site) = temp_rowVal_a2(site) + piU2(row, col) * DSPWave_Matrix_rov(site).Element(col)
                    Next col
                End If

                a2(site).Element(row) = (temp_rowVal_a2(site) - a2_min(row)) / (a2_max(row) - a2_min(row))
                temp_rowVal_a2 = 0
                If (row = 0) Then
                    Call MTR_Cal_DecimalToBinary(a2(site).Element(row), 15, decimalPlaces, currBinaryStr)
                Else
                    Call MTR_Cal_DecimalToBinary(a2(site).Element(row), 14, decimalPlaces, currBinaryStr)
                End If
                totalBinaryStr = totalBinaryStr + currBinaryStr

                'Added on 20180131 To Force Error
                If (a2(site).Element(row) = 0) Then
                    a2(site).Element(row) = -0.000001
                ElseIf (a2(site).Element(row) = 1) Then
                    a2(site).Element(row) = 1.000001
                End If
            Next row
            currElementDspWave = 0
            TheExec.Datalog.WriteComment ("Fuse Binary Str  a2 for Site:" + CStr(site) + " is " + totalBinaryStr)
            totalBinaryStr = StrReverse(totalBinaryStr)
            If Len(totalBinaryStr) = OutDspWaveToFuse_2.SampleSize Then

                Do While currElementDspWave < FuseSize_2
                    OutDspWaveToFuse_2(site).Element(currElementDspWave) = CInt(Mid(totalBinaryStr, currElementDspWave + 1, 1))
                    currElementDspWave = currElementDspWave + 1
                Loop
            End If
        Next site
            For row = 0 To 3
                testName = "a1_row_" + SensorTempName_rot + "_" + CStr(row + 1) + ":"
                TheExec.Flow.TestLimit resultVal:=a1.Element(row), Tname:=testName, ForceResults:=tlForceFlow
            Next row
            Set DSPWave_Coeff_1 = a1
            For row = 0 To 2
                testName = "a2_row_" + SensorTempName_rov + "_" + CStr(row + 1) + ":"
                TheExec.Flow.TestLimit resultVal:=a2.Element(row), Tname:=testName, ForceResults:=tlForceFlow
            Next row
            Set DSPWave_Coeff_2 = a2
            
    ElseIf Temperature = 85 Then

        For Each site In TheExec.sites
            totalBinaryStr = ""
            For row = 0 To 3
                temp_rowVal_a3(site) = 0
                currBinaryStr = ""


                If MTR_CAL_Sheet = 0 Then
                    For col = 0 To 6
                    temp_rowVal_a3(site) = temp_rowVal_a3(site) + piU3(row, col) * DSPWave_Matrix_rot(site).Element(col)
                Next col
                Else
                    For col = 0 To 4
                        temp_rowVal_a3(site) = temp_rowVal_a3(site) + piU3(row, col) * DSPWave_Matrix_rot(site).Element(col)
                    Next col
                End If

                a3(site).Element(row) = (temp_rowVal_a3(site) - a3_min(row)) / (a3_max(row) - a3_min(row))
                temp_rowVal_a3 = 0
                If (row = 0) Then
                    Call MTR_Cal_DecimalToBinary(a3(site).Element(row), 15, decimalPlaces, currBinaryStr)
                Else
                    Call MTR_Cal_DecimalToBinary(a3(site).Element(row), 14, decimalPlaces, currBinaryStr)
                End If
                totalBinaryStr = totalBinaryStr + currBinaryStr
                'Added on 20180131 To Force Error
                If (a3(site).Element(row) = 0) Then
                    a3(site).Element(row) = -0.000001
                ElseIf (a3(site).Element(row) = 1) Then
                    a3(site).Element(row) = 1.000001
                End If
            Next row
            currElementDspWave = 0
            TheExec.Datalog.WriteComment ("Fuse Binary Str  a3 for Site:" + CStr(site) + " is " + totalBinaryStr)
            totalBinaryStr = StrReverse(totalBinaryStr)
            If Len(totalBinaryStr) = OutDspWaveToFuse_1.SampleSize Then
                
                Do While currElementDspWave < FuseSize_1
                    OutDspWaveToFuse_1(site).Element(currElementDspWave) = CInt(Mid(totalBinaryStr, currElementDspWave + 1, 1))
                    currElementDspWave = currElementDspWave + 1
                Loop
            End If
            totalBinaryStr = ""
            For row = 0 To 2
                temp_rowVal_a4(site) = 0
                currBinaryStr = ""

                    
                If MTR_CAL_Sheet = 0 Then
                    For col = 0 To 6
                    temp_rowVal_a4(site) = temp_rowVal_a4(site) + piU4(row, col) * DSPWave_Matrix_rov(site).Element(col)
                Next col
                Else
                    For col = 0 To 4
                        temp_rowVal_a4(site) = temp_rowVal_a4(site) + piU4(row, col) * DSPWave_Matrix_rov(site).Element(col)
                    Next col
                End If
                    
                a4(site).Element(row) = (temp_rowVal_a4(site) - a4_min(row)) / (a4_max(row) - a4_min(row))
                temp_rowVal_a4 = 0
                If (row = 0) Then
                    Call MTR_Cal_DecimalToBinary(a4(site).Element(row), 15, decimalPlaces, currBinaryStr)
                Else
                    Call MTR_Cal_DecimalToBinary(a4(site).Element(row), 14, decimalPlaces, currBinaryStr)
                End If
                totalBinaryStr = totalBinaryStr + currBinaryStr

                'Added on 20180131 To Force Error
                If (a4(site).Element(row) = 0) Then
                    a4(site).Element(row) = -0.000001
                ElseIf (a4(site).Element(row) = 1) Then
                    a4(site).Element(row) = 1.000001
                End If
            Next row
            currElementDspWave = 0
            TheExec.Datalog.WriteComment ("Fuse Binary Str  a4 for Site:" + CStr(site) + " is " + totalBinaryStr)
            totalBinaryStr = StrReverse(totalBinaryStr)
            If Len(totalBinaryStr) = OutDspWaveToFuse_2.SampleSize Then

                Do While currElementDspWave < FuseSize_2
                    OutDspWaveToFuse_2(site).Element(currElementDspWave) = CInt(Mid(totalBinaryStr, currElementDspWave + 1, 1))
                    currElementDspWave = currElementDspWave + 1
                Loop
            End If
        Next site
        For row = 0 To 3
            testName = "a3_row_" + SensorTempName_rot + "_" + CStr(row + 1) + ":"
            TheExec.Flow.TestLimit resultVal:=a3.Element(row), Tname:=testName, ForceResults:=tlForceFlow
        Next row
        Set DSPWave_Coeff_1 = a3
        For row = 0 To 2
            testName = "a4_row_" + SensorTempName_rov + "_" + CStr(row + 1) + ":"
            TheExec.Flow.TestLimit resultVal:=a4.Element(row), Tname:=testName, ForceResults:=tlForceFlow
        Next row
        Set DSPWave_Coeff_2 = a4
    End If
End Function

Public Function MTR_Cal_DecimalToBinary(ByVal inputDecimal As Double, ByVal bitSize As Long, ByVal placesAfterDecimal As Long, ByRef outBinaryStr As String) As Long
    
    Dim i As Long
    Dim fractional As Double
    Dim integral  As Long
    Dim currIntegral As Long
    Dim decimalFract As Double
    Dim binaryStr As String
    Dim currDecimal As Double
    Dim theDecimal As Double
    Dim currCount As Long
 
    theDecimal = FormatNumber(inputDecimal, placesAfterDecimal)
       
     
    integral = Int(theDecimal)
    
    fractional = theDecimal - integral
    
    If (theDecimal > 0) And (theDecimal < 1) Then
    
        
        
        currCount = 0
        Do While currCount < bitSize
        
            currDecimal = fractional * 2
            currIntegral = Int(currDecimal)
            decimalFract = decimalFract + CStr(currIntegral) * (2 ^ (bitSize - currCount))
            binaryStr = binaryStr + CStr(currIntegral)
            fractional = currDecimal - currIntegral
        
            
            currCount = currCount + 1
            
        Loop
        outBinaryStr = binaryStr
        

    Else
        currCount = 0
        binaryStr = ""
        Do While currCount < bitSize
            binaryStr = binaryStr + "1"
            currCount = currCount + 1
        
        Loop
        outBinaryStr = binaryStr
    End If


End Function


Public Function TX_Low_Level(argc As Integer, argv() As String) As Long

    Dim DictKey_V1 As String, DictKey_V2 As String
    Dim pld_V1 As New PinListData, pld_V2 As New PinListData
    Dim pld_upd_V1 As New PinListData, pld_upd_V2 As New PinListData
    Dim Pin_Name_1 As String, Pin_Name_2 As String
    'Dim Rak_Pin_Name_1() As Double
    'Dim Rak_Pin_Name_2() As Double
    Dim GetRakVal As Double
    Dim OutputTname_format() As String
    Dim TestNameInput As String

    DictKey_V1 = argv(0)
    DictKey_V2 = argv(1)
    Pin_Name_1 = argv(2)
    Pin_Name_2 = argv(3)
    Dim site As Variant
    pld_V1 = GetStoredMeasurement(DictKey_V1)
    pld_V2 = GetStoredMeasurement(DictKey_V2)
    
    pld_upd_V1.AddPin (Pin_Name_1)
    pld_upd_V2.AddPin (Pin_Name_2)
    
    For Each site In TheExec.sites
        'Rak_Pin_Name_1 = TheHdw.PPMU.ReadRakValuesByPinnames(Pin_Name_1, site)
        'Rak_Pin_Name_2 = TheHdw.PPMU.ReadRakValuesByPinnames(Pin_Name_2, site)
        GetRakVal = (CurrentJob_Card_RAK.Pins(Pin_Name_1).Value(site) + CurrentJob_Card_RAK.Pins(Pin_Name_2).Value(site)) / 2
        pld_upd_V1.Pins(Pin_Name_1).Value(site) = pld_V1.Pins(Pin_Name_1).Multiply(45).Divide(45 + 45 + GetRakVal).Value(site)
        pld_upd_V2.Pins(Pin_Name_2).Value(site) = pld_V2.Pins(Pin_Name_2).Multiply(45).Divide(45 + 45 + GetRakVal).Value(site)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", "", 0)
    TheExec.Flow.TestLimit resultVal:=pld_upd_V1, ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", "", 0)
    TheExec.Flow.TestLimit resultVal:=pld_upd_V2, ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    
End Function

Public Function Calc_MIPI_Tolerance(argc As Integer, argv() As String) As Long



    Dim site As Variant
    Dim i, j As Long
    Dim DSPWave_First As New DSPWave
    Dim DSPWave_Second As New DSPWave
    Dim DSPWave_Combine() As New DSPWave
    Dim TestNameInput As String
    Dim SplitByAdd() As String
    Dim First_StartElement As Long
    Dim First_EndElement As Long
    Dim Second_StartElement As Long
    Dim Second_EndElement As Long
    
    Dim DictKey_DSPWave_Combine As String
    
    Dim DataString_First As String
    Dim DataString_Second As String
    Dim DataString_Combine As String
    
    ReDim DSPWave_Combine(argc - 1) As New DSPWave
    Dim DSPWave_Combine_Dec As New DSPWave
    Dim OutputTname_format() As String
'    Dim TestNameInput As String
    
    For i = 0 To argc - 1
        'TestNameInput = "ConcatenateDSP_"
        SplitByAdd = Split(argv(i), "+")
        DSPWave_First = GetStoredCaptureData(SplitByAdd(0))
        First_StartElement = 0
        First_EndElement = 7
        DSPWave_Second = GetStoredCaptureData(SplitByAdd(1))
        Second_StartElement = 0
        Second_EndElement = 1
        

        Call ConcatenateDSP_TTR(DSPWave_First, First_StartElement, First_EndElement, DSPWave_Second, Second_StartElement, Second_EndElement, DSPWave_Combine(i))
        
        ''20170718 - Store Concatenate DSP to Dict.
'        If UBound(SplitByAt) = 6 Then
'            DictKey_DSPWave_Combine = SplitByAt(6)
'            Call AddStoredCaptureData(DictKey_DSPWave_Combine, DSPWave_Combine(i))
'        End If
        
        For Each site In TheExec.sites
            DataString_First = ""
            DataString_Second = ""
            DataString_Combine = ""
            For j = 0 To DSPWave_First.SampleSize - 1
                DataString_First = DataString_First & DSPWave_First(site).Element(j)
            Next j
            For j = 0 To DSPWave_Second.SampleSize - 1
                DataString_Second = DataString_Second & DSPWave_Second(site).Element(j)
            Next j
            For j = 0 To DSPWave_Combine(i).SampleSize - 1
                DataString_Combine = DataString_Combine & DSPWave_Combine(i)(site).Element(j)
            Next j
            
           If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Dictionary " & SplitByAdd(0) & " Output Bits = " & DataString_First & " Extract Bits [" & First_StartElement & "-" & First_EndElement & "]" & _
                                                           " ,Dictionary " & SplitByAdd(1) & " Output Bits = " & DataString_Second & " Extract Bits [" & Second_StartElement & "-" & Second_EndElement & "]" & _
                                                           " ,Dictionary " & DictKey_DSPWave_Combine & " Output Bits = " & DataString_Combine)
        
        
        
        
        
        
        DSPWave_Combine(i)(site) = DSPWave_Combine(i)(site).ConvertDataTypeTo(DspLong)
        DSPWave_Combine_Dec(site) = DSPWave_Combine(i)(site).ConvertStreamTo(tldspParallel, DSPWave_Combine(i)(site).SampleSize, 0, Bit0IsMsb)
        
        
        Next site
        'Call rundsp.BinToDec(DSPWave_Combine(i), DSPWave_Combine_Dec)
                
        If gl_Disable_HIP_debug_log = False Then

            TestNameInput = Report_TName_From_Instance(CalcC, "X", "ConcatenateDSP", 0)
               
            TheExec.Flow.TestLimit resultVal:=DSPWave_Combine_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceNone
      
        End If
        
      
        Dim MIPI_threshold_Code_value_P(7) As New SiteDouble
        Dim MIPI_threshold_Code_value_N(7) As New SiteDouble
        
        For Each site In TheExec.sites
        
        
        'split code p and code n
        If i < 8 Then
            MIPI_threshold_Code_value_P(i)(site) = DSPWave_Combine_Dec(site).Element(0)
        Else
            i = i - 8
            MIPI_threshold_Code_value_N(i)(site) = DSPWave_Combine_Dec(site).Element(0)
            i = i + 8
        End If
        
        Next site

        
    Next i


    Dim MIPI_threshold_lower_p(0) As New SiteVariant
    Dim MIPI_threshold_high_p(0) As New SiteVariant
    'Dim MIPI_threshold_found_p As New SiteBoolean   'Change to SiteLong, due to SiteBoolean True = -1
    Dim MIPI_threshold_lower_n(0) As New SiteVariant
    Dim MIPI_threshold_high_n(0) As New SiteVariant
    'Dim MIPI_threshold_found_n As New SiteBoolean  'Change to SiteLong, due to SiteBoolean True = -1
    Dim MIPI_trans_mapping As Variant
    Dim MIPI_threshold_found_p_value As New SiteLong
    Dim MIPI_threshold_found_n_value As New SiteLong
    
    Dim threshold_temp_p As Integer
    Dim threshold_flag_p As Boolean
    Dim threshold_temp_n As Integer
    Dim threshold_flag_n As Boolean
    Dim p  As Long
    Dim n  As Long
    MIPI_trans_mapping = Array(-0.2, -0.15, -0.1, -0.05, 0.05, 0.1, 0.15, 0.2)

    For Each site In TheExec.sites
    
    'code p process
        threshold_temp_p = 0
        threshold_flag_p = False
        'MIPI_threshold_found_p(Site) = False
        MIPI_threshold_found_p_value(site) = -1  'Clear = -1
        
        
        For p = 0 To 7

            If MIPI_threshold_Code_value_P(p)(site) = 0 Then
                If threshold_flag_p = False Then
                    MIPI_threshold_lower_p(0)(site) = p
                    threshold_flag_p = True
                    MIPI_threshold_found_p_value(site) = 1 'True = 1
                End If
                If threshold_flag_p = True Then
                    MIPI_threshold_high_p(0)(site) = p
                End If
            End If
            If MIPI_threshold_Code_value_P(p)(site) > 0 Then
                threshold_temp_p = threshold_temp_p + 1
            End If
        Next p
        
        If threshold_temp_p = 0 Then
            MIPI_threshold_found_p_value(site) = 0 'Flase = 0
        End If
        
        If MIPI_threshold_lower_p(0)(site) <> "" Then
            MIPI_threshold_lower_p(0)(site) = MIPI_trans_mapping(MIPI_threshold_lower_p(0)(site))
        Else
            MIPI_threshold_lower_p(0)(site) = 999
        End If

         If MIPI_threshold_high_p(0)(site) <> "" Then
            MIPI_threshold_high_p(0)(site) = MIPI_trans_mapping(MIPI_threshold_high_p(0)(site))
        Else
            MIPI_threshold_high_p(0)(site) = 999
        End If

       'code n process
        threshold_temp_n = 0
        threshold_flag_n = False
        'MIPI_threshold_found_n(Site) = False
        MIPI_threshold_found_n_value(site) = -1 'Clear = -1
        
        
        For n = 0 To 7

            If MIPI_threshold_Code_value_N(n)(site) = 0 Then
                If threshold_flag_n = False Then
                    MIPI_threshold_lower_n(0)(site) = n
                    threshold_flag_n = True
                    MIPI_threshold_found_n_value(site) = 1 'True = 1
                End If
                If threshold_flag_n = True Then
                    MIPI_threshold_high_n(0)(site) = n
                End If
            End If
            If MIPI_threshold_Code_value_N(n)(site) > 0 Then
                threshold_temp_n = threshold_temp_n + 1
            End If
        Next n
        
        If threshold_temp_n = 0 Then
            MIPI_threshold_found_n_value(site) = 0 'Flase = 0
        End If
        
        If MIPI_threshold_lower_n(0)(site) <> "" Then
            MIPI_threshold_lower_n(0)(site) = MIPI_trans_mapping(MIPI_threshold_lower_n(0)(site))
        Else
            MIPI_threshold_lower_n(0)(site) = 999
        End If

         If MIPI_threshold_high_n(0)(site) <> "" Then
            MIPI_threshold_high_n(0)(site) = MIPI_trans_mapping(MIPI_threshold_high_n(0)(site))
        Else
            MIPI_threshold_high_n(0)(site) = 999
        End If


    Next site
    If gl_Disable_HIP_debug_log = False Then
  ' print datdlog
    For p = 0 To 7
        TestNameInput = Report_TName_From_Instance(CalcC, "code_P_" & p, "", CLng(p))
        TheExec.Flow.TestLimit MIPI_threshold_Code_value_P(p), 0, 2 ^ 10 - 1, PinName:="code_P_" & p, ForceResults:=tlForceNone, Tname:=TestNameInput
    Next p
    End If
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA0_Term_Tol1", "", 0)
    TheExec.Flow.TestLimit MIPI_threshold_lower_p(0), scaletype:=scaleNone, PinName:="DATA0_Term_Tol1", ForceResults:=tlForceFlow, Tname:=TestNameInput
        
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA0_Term_Tol2", "", 0)
    TheExec.Flow.TestLimit MIPI_threshold_high_p(0), scaletype:=scaleNone, PinName:="DATA0_Term_Tol2", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA0_Found_Thresh", "", 0)
    TheExec.Flow.TestLimit MIPI_threshold_found_p_value, 1, 1, PinName:="DATA0_Found_Thresh", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    If gl_Disable_HIP_debug_log = False Then
    
    For n = 0 To 7
        TestNameInput = Report_TName_From_Instance(CalcC, "code_N_" & n, "", CLng(n))
        TheExec.Flow.TestLimit MIPI_threshold_Code_value_N(n), 0, 2 ^ 10 - 1, PinName:="code_N_" & n, ForceResults:=tlForceNone, Tname:=TestNameInput
        
    Next n
    End If

    TestNameInput = Report_TName_From_Instance(CalcC, "DATA1_Term_Tol1", "", 0)
    TheExec.Flow.TestLimit MIPI_threshold_lower_n(0), scaletype:=scaleNone, PinName:="DATA1_Term_Tol1", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA1_Term_Tol2", "", 0)
    TheExec.Flow.TestLimit MIPI_threshold_high_n(0), scaletype:=scaleNone, PinName:="DATA1_Term_Tol2", ForceResults:=tlForceFlow, Tname:=TestNameInput
        
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA1_Found_Thresh", "", 0)
    TheExec.Flow.TestLimit MIPI_threshold_found_n_value, 1, 1, PinName:="DATA1_Found_Thresh", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
End Function


Public Function Calc_ADC_Error_code(argc As Integer, argv() As String) As Long

Dim site As Variant
Dim ADC_Trim_Code As New DSPWave: ADC_Trim_Code.CreateConstant 0, 1, DspLong
Dim Error_Code As New DSPWave: Error_Code.CreateConstant 0, 1, DspLong
Dim ERROR_CODE_Dict As New DSPWave
Dim ADC_Error_Code_Str As String
Dim ADC_Trim_Code_Str As String
Dim REFERENCE_CTRL As Long
Dim ADC_Error_Code_Str_25 As String
Dim ADC_Error_Code_Str_85 As String
Dim Error_Code_25C_Dec As New DSPWave: Error_Code_25C_Dec.CreateConstant 0, 1, DspLong
Dim Error_Code_85C_Dec As New DSPWave: Error_Code_85C_Dec.CreateConstant 0, 1, DspLong
Dim Error_Code_25C As New DSPWave
Dim Error_Code_85C As New DSPWave
Dim SL_BitWidth As New SiteLong
Dim ADC_Final_RefCtrl_Str As String
Dim ADC_Final_RefCtrl As New DSPWave: ADC_Final_RefCtrl.CreateConstant 0, 1, DspLong
Dim ADC_Final_RefCtrl_Dict As New DSPWave
Dim OutputTname_format() As String
Dim TestNameInput As String



    ADC_Trim_Code_Str = argv(0)
    ADC_Error_Code_Str = argv(1)
    REFERENCE_CTRL = argv(2)

    Call HardIP_Bin2Dec(ADC_Trim_Code, GetStoredCaptureData(ADC_Trim_Code_Str))
    For Each site In TheExec.sites
        Error_Code(site).Element(0) = ADC_Trim_Code(site).Element(0) - REFERENCE_CTRL
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, ADC_Error_Code_Str, "")
    TheExec.Flow.TestLimit resultVal:=Error_Code.Element(0), lowVal:=-127, hiVal:=127, ForceResults:=tlForceNone, Tname:=TestNameInput
        
    For Each site In TheExec.sites
        If Error_Code(site).Element(0) < -128 Then
            Error_Code(site).Element(0) = 128
        ElseIf Error_Code(site).Element(0) < 0 Then
            Error_Code(site).Element(0) = 2 ^ 8 + FormatNumber(Error_Code(site).Element(0))
        ElseIf Error_Code(site).Element(0) > 127 Then
            Error_Code(site).Element(0) = 127
        End If
    Next site
    Call HardIP_Dec2Bin(ERROR_CODE_Dict, Error_Code, 8)
    Call AddStoredCaptureData(ADC_Error_Code_Str, ERROR_CODE_Dict)
    
    
    If argc >= 4 Then
        ADC_Error_Code_Str_25 = argv(3)
        ADC_Error_Code_Str_85 = argv(1)
        ADC_Final_RefCtrl_Str = argv(4)
        
        Error_Code_25C = GetStoredCaptureData(ADC_Error_Code_Str_25)
        Error_Code_85C = GetStoredCaptureData(ADC_Error_Code_Str_85)
    
        For Each site In TheExec.sites
            SL_BitWidth(site) = Error_Code_25C(site).SampleSize
        Next site
        
        Call rundsp.DSP_2S_Complement_To_SignDec(Error_Code_25C, SL_BitWidth, Error_Code_25C_Dec)
        Call rundsp.DSP_2S_Complement_To_SignDec(Error_Code_85C, SL_BitWidth, Error_Code_85C_Dec)
    
        For Each site In TheExec.sites
            ADC_Final_RefCtrl(site).Element(0) = REFERENCE_CTRL + (Error_Code_25C_Dec(site).Element(0) + Error_Code_85C_Dec(site).Element(0)) / 2
        Next site
        TestNameInput = Report_TName_From_Instance(CalcC, "FinalReferenceControlCode", "")
        TheExec.Flow.TestLimit resultVal:=ADC_Final_RefCtrl.Element(0), ForceResults:=tlForceNone, Tname:=TestNameInput
        
        Call HardIP_Dec2Bin(ADC_Final_RefCtrl_Dict, ADC_Final_RefCtrl, 8)
        Call AddStoredCaptureData(ADC_Final_RefCtrl_Str, ADC_Final_RefCtrl_Dict)
    
    
    End If
    
    
End Function

Public Function ADC_code_toV(argc As Integer, argv() As String) As Long '----------------add by CSHO 20171227

Dim ADCcapcode As String
Dim USBvoltages As String
Dim USBvoltages2 As String
Dim devideV As Long
Dim ADC_voltages As String
Dim InputKey As String
Dim DSP_Input As New DSPWave
Dim LSB As Double
Dim bitprint As String
Dim i As Integer
Dim ADC_voltages_final As Double
Dim site As Variant

InputKey = argv(0)
USBvoltages = argv(1)
devideV = argv(2)

Set DSP_Input = Nothing
DSP_Input = GetStoredCaptureData(InputKey)
 For Each site In TheExec.sites
 For i = 0 To DSP_Input.SampleSize - 1
    If i = 0 Then
      bitprint = DSP_Input.Element(0)
    
    Else
      bitprint = bitprint & DSP_Input.Element(i)
    
    End If
  Next i


ADC_voltages = Bin2Dec(bitprint)

USBvoltages2 = TheExec.specs.DC.Item(USBvoltages).ContextValue

LSB = CDbl(USBvoltages2) / devideV

ADC_voltages_final = ADC_voltages * LSB

If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "site" & site & "Convert ADC codes to voltages" & " V: " & ADC_voltages_final
Next site


End Function



Public Function Calc_MTR_REL_Freq_Diff_AVG(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim freq_Dut As String
    Dim freq_ref As String
    
    Dim dut As String
    Dim ref As String

    Dim fdiff_percent As String
    Dim testName As String
    Dim Efuse_Dict_Name As String
'    Dim DSP_fdiff_percent As String

'    Dim DSP_freq_Dut As String
'    Dim DSP_freq_ref As String
    Dim index_name As String
    Dim Index_count As Long
    Dim i, k As Integer

    Dim freq_Dut_dsp As New DSPWave: freq_Dut_dsp = Nothing
    Dim freq_ref_dsp As New DSPWave: freq_ref_dsp = Nothing
    Dim fdiff_percent_dsp As New DSPWave: fdiff_percent_dsp = Nothing
    
    Dim freq_Dut_wav As New DSPWave: freq_Dut_wav = Nothing
    Dim freq_ref_wav As New DSPWave: freq_ref_wav = Nothing
    Dim fdiff_percent_wav As New DSPWave: fdiff_percent_wav = Nothing
    
    Dim freq_Dut_mean As New SiteDouble
    Dim freq_ref_mean As New SiteDouble
    Dim fdiff_percent_mean As New SiteDouble
    
    Dim freq_Dut_std As New SiteDouble
    Dim freq_ref_std As New SiteDouble
    Dim fdiff_percent_std As New SiteDouble
    
    Dim RSD_DUT As New SiteDouble
    Dim RSD_REF As New SiteDouble
    Dim R_Ref As New SiteDouble
    
    Dim DSP_fdiff_percent As New DSPWave: DSP_fdiff_percent.CreateConstant 0, 1, DspDouble
    Dim DSP_freq_Dut As New DSPWave: DSP_freq_Dut = Nothing
    Dim DSP_freq_ref As New DSPWave: DSP_freq_ref = Nothing
    
    Dim dut_array() As String
    Dim ref_array() As String
    Dim freq_Dut_array() As String
    Dim freq_ref_array() As String
    Dim fdiff_percent_array() As String
    Dim Check_Freq As New SiteBoolean
    Dim Check_STD As New SiteBoolean
    Dim Check_Ratio As New SiteBoolean
    
    Dim Freq_HiLimit As Double: Freq_HiLimit = 1150000000
    Dim Freq_LoLimit As Double: Freq_LoLimit = 650000000
    Dim STD_HiLimit As Double: STD_HiLimit = 0.2
    Dim STD_LoLimit As Double: STD_LoLimit = 0
    Dim F_Ratio_HiLimit As Double: F_Ratio_HiLimit = 106
    Dim F_Ratio_LoLimit As Double: F_Ratio_LoLimit = 94
    Dim Fuse_Code() As New DSPWave
    Dim Final_Fuse_Code As New DSPWave: Final_Fuse_Code = Nothing
    Dim Final_Fuse_Code_DEC As New DSPWave: Final_Fuse_Code_DEC = Nothing
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    

    Dim xxx As New DSPWave
    Dim yyy As New DSPWave
    Final_Fuse_Code_DEC.CreateConstant 0, 1, DspDouble
'    yyy.CreateConstant 0, 8, DspLong
'    xxx.CreateConstant 0, 8, DspLong
'    xxx(0).Element(0) = 1
'
'    xxx = xxx.ConvertStreamTo(tldspParallel, 8, 0, Bit0IsMsb)
'
'    xxx = xxx.ConvertDataTypeTo(DspLong)
'    xxx(0).Element(0) = 128
'    yyy = xxx.ConvertStreamTo(tldspSerial, 8, 0, Bit0IsMsb)




    dut = argv(0)
    ref = argv(1)
    freq_Dut = argv(2)
    freq_ref = argv(3)
    fdiff_percent = argv(4)
    
    index_name = argv(5)
    Index_count = argv(6)
    Efuse_Dict_Name = argv(7)
    
    
    dut_array = Split(dut, "@")
    ref_array = Split(ref, "@")
    freq_Dut_array = Split(freq_Dut, "@")
    freq_ref_array = Split(freq_ref, "@")
    fdiff_percent_array = Split(fdiff_percent, "@")
    
    
    For k = 0 To UBound(dut_array)
    
        testName = "f_diff_" + Replace(freq_Dut_array(k), index_name, TheExec.Flow.var(index_name).Value)
        DSP_freq_Dut = GetStoredCaptureData(dut_array(k))
        DSP_freq_ref = GetStoredCaptureData(ref_array(k))
        For Each site In TheExec.sites.Active
            DSP_freq_Dut = DSP_freq_Dut.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
            DSP_freq_ref = DSP_freq_ref.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
            DSP_freq_Dut = DSP_freq_Dut.Multiply(93750)
            DSP_freq_ref = DSP_freq_ref.Multiply(93750)
        Next site
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "======================================================================Start Calc Freq======================================================================"
        For Each site In TheExec.sites
            If DSP_freq_Dut.Element(0) <> 0 Then
                DSP_fdiff_percent.Element(0) = ((DSP_freq_Dut.Element(0) - DSP_freq_ref.Element(0)) / DSP_freq_Dut.Element(0)) * 100
            Else
                DSP_fdiff_percent.Element(0) = 99999
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  freq_of_Dut " + freq_Dut_array(k) + " is 0")
            End If
        Next site
        
        TestNameInput = Report_TName_From_Instance(CalcF, "X", Replace(freq_Dut_array(k), index_name, ""), CInt(TheExec.Flow.var(index_name).Value))
        
        TheExec.Flow.TestLimit resultVal:=DSP_freq_Dut.Element(0), Tname:=Replace(freq_Dut_array(k), index_name, TheExec.Flow.var(index_name).Value), ForceResults:=tlForceNone, scaletype:=scaleMega
        
        TestNameInput = Report_TName_From_Instance(CalcF, "X", Replace(freq_ref_array(k), index_name, ""), CInt(TheExec.Flow.var(index_name).Value))
        
        TheExec.Flow.TestLimit resultVal:=DSP_freq_ref.Element(0), Tname:=Replace(freq_ref_array(k), index_name, TheExec.Flow.var(index_name).Value), ForceResults:=tlForceNone, scaletype:=scaleMega
        
        TestNameInput = Report_TName_From_Instance(CalcF, "X", "Percent", CInt(TheExec.Flow.var(index_name).Value))
        
        TheExec.Flow.TestLimit resultVal:=DSP_fdiff_percent.Element(0), Tname:=testName, ForceResults:=tlForceNone
            
        
    
        
        
        Call AddStoredCaptureData(Replace(freq_Dut_array(k), index_name, TheExec.Flow.var(index_name).Value), DSP_freq_Dut)
        Call AddStoredCaptureData(Replace(freq_ref_array(k), index_name, TheExec.Flow.var(index_name).Value), DSP_freq_ref)
    
        Call AddStoredCaptureData(Replace(fdiff_percent_array(k), index_name, TheExec.Flow.var(index_name).Value), DSP_fdiff_percent)
        
        
        
        
        If (TheExec.Flow.var(index_name).Value + 1 = Index_count) Then
           If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "======================================================================Start Calc Mean,SD,Ratio======================================================================"
            Set freq_Dut_dsp = New DSPWave
            freq_Dut_dsp.CreateConstant 0, Index_count
            Set freq_ref_dsp = New DSPWave
            freq_ref_dsp.CreateConstant 0, Index_count
            Set fdiff_percent_dsp = New DSPWave
            fdiff_percent_dsp.CreateConstant 0, Index_count
            
            
            For Each site In TheExec.sites.Active
                Check_Freq = True
                Check_STD = True
                Check_Ratio = True
            Next site
            
            For i = 0 To Index_count - 1
                freq_Dut_wav = GetStoredCaptureData(Replace(freq_Dut_array(k), index_name, CStr(i)))
                freq_ref_wav = GetStoredCaptureData(Replace(freq_ref_array(k), index_name, CStr(i)))
                fdiff_percent_wav = GetStoredCaptureData(Replace(fdiff_percent_array(k), index_name, CStr(i)))
                
                For Each site In TheExec.sites.Active
                    freq_Dut_dsp.Element(i) = freq_Dut_wav.Element(0)
                    If (freq_Dut_wav.Element(0) < Freq_HiLimit And freq_Dut_wav.Element(0) > Freq_LoLimit) Then
                        Check_Freq = Check_Freq And True
                    Else
                        Check_Freq = False
                    End If
                    
                    freq_ref_dsp.Element(i) = freq_ref_wav.Element(0)
                        If (freq_ref_wav.Element(0) < Freq_HiLimit And freq_ref_wav.Element(0) > Freq_LoLimit) Then
                        Check_Freq = Check_Freq And True
                    Else
                        Check_Freq = False
                    End If
                    fdiff_percent_dsp.Element(i) = fdiff_percent_wav.Element(0)
                Next site
            Next i
            
            Dim freq_Dut_std_dbl As Double
            Dim freq_ref_std_dbl As Double
            
            
            
            For Each site In TheExec.sites.Active
                freq_Dut_mean = freq_Dut_dsp.CalcMeanWithStdDev(freq_Dut_std_dbl)
                freq_ref_mean = freq_ref_dsp.CalcMeanWithStdDev(freq_ref_std_dbl)
                fdiff_percent_mean = fdiff_percent_dsp.CalcMean
                
'                freq_Dut_dsp.CalcMeanWithStdDev (freq_Dut_std)
'                freq_ref_dsp.CalcMeanWithStdDev (freq_ref_std)
'                fdiff_percent_dsp.CalcMeanWithStdDev (fdiff_percent_std)
                If freq_Dut_mean = 0 Then
                    RSD_DUT = 0
                Else
                    RSD_DUT = 3 * freq_Dut_std_dbl / freq_Dut_mean * 100
                End If
                If (RSD_DUT < STD_HiLimit And RSD_DUT > STD_LoLimit) Then
                    Check_STD = Check_STD And True
                Else
                    Check_STD = False
                End If
                
                If freq_ref_mean = 0 Then
                    RSD_REF = 0
                Else
                    RSD_REF = 3 * freq_ref_std_dbl / freq_ref_mean * 100
                End If
                If (RSD_REF < STD_HiLimit And RSD_REF > STD_LoLimit) Then
                    Check_STD = Check_STD And True
                Else
                    Check_STD = False
                End If
                If freq_Dut_mean = 0 Then
                    R_Ref = 0
                Else
                    R_Ref = freq_ref_mean / freq_Dut_mean * 100
                End If
                If (R_Ref < F_Ratio_HiLimit And R_Ref > F_Ratio_LoLimit) Then
                    Check_Ratio = Check_Ratio And True
                Else
                    Check_Ratio = False
                End If
                
            Next site

            TheExec.Flow.TestLimit resultVal:=freq_Dut_mean, Tname:="freq_Dut_mean", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=freq_ref_mean, Tname:="freq_ref_mean", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=RSD_DUT, Tname:="RSD_DUT", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=RSD_REF, Tname:="RSD_REF", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=R_Ref, Tname:="R_Ref", ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=fdiff_percent_mean, Tname:="Avg_R0t0_E3", ForceResults:=tlForceFlow

            ReDim Fuse_Code(UBound(dut_array)) As New DSPWave
            Dim fuse_code_dec As New DSPWave
            
            Set Fuse_Code(k) = New DSPWave
            Fuse_Code(k).CreateConstant 0, 16, DspLong
            Set fuse_code_dec = New DSPWave
            fuse_code_dec.CreateConstant 0, 1, DspLong
            
            For Each site In TheExec.sites.Active
                If Check_Freq = False Then
                    fuse_code_dec.Element(0) = 65533    '0xFFFD
                ElseIf Check_STD = False Then
                    fuse_code_dec.Element(0) = 65534    '0xFFFE
                ElseIf Check_Ratio = False Then
                    fuse_code_dec.Element(0) = 65532    '0xFFFC
                Else
                    If (fdiff_percent_mean >= 0) Then
                        fuse_code_dec.Element(0) = Abs(fdiff_percent_mean) * 1000
                    Else
                        fuse_code_dec.Element(0) = Abs(fdiff_percent_mean) * 1000 + 32768
                    End If
                End If
                fuse_code_dec = fuse_code_dec.ConvertDataTypeTo(DspLong)
                Fuse_Code(k) = fuse_code_dec.ConvertStreamTo(tldspSerial, 16, 0, Bit0IsMsb)
               If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Efuse Write,Site:" + CStr(site) + " Value: " + CStr(fuse_code_dec.Element(0))
                Final_Fuse_Code = Final_Fuse_Code.Concatenate(Fuse_Code(k))
                Final_Fuse_Code_DEC.Element(0) = Final_Fuse_Code_DEC.Element(0) * (2 ^ 16) * k + fuse_code_dec.Element(0)
            Next site

        End If
    Next k
    If (TheExec.Flow.var(index_name).Value + 1 = Index_count) Then
    
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Final Efuse Write Value , Site:" + CStr(site) + " Value: " + CStr(Final_Fuse_Code_DEC.Element(0))
            Next site
        End If
        
        Call AddStoredCaptureData(Efuse_Dict_Name, Final_Fuse_Code_DEC)
    End If
End Function

Public Function Calc_MTR_AVG(argc As Integer, argv() As String) As Long

'    Dim index_name As String
'    Dim Sweep_Dictionary As String
    Dim Loop_count As Long
    Dim Loop_Index As Long
    
    Dim DSP_Capture As New DSPWave
    Dim i, j As Long

    Dim Sweep_index As Long
    Dim Sweep_Info() As Power_Sweep
    Dim Sweep_Count As Long
    Dim dict_key As String
    Dim DSP_Result As New DSPWave
    Dim site As Variant
    Dim Sweep_Mean As New SiteDouble
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    Loop_Index = argv(argc - 1)

    ReDim Sweep_Info(argc - 2) As Power_Sweep
    
    
    For i = 0 To argc - 2
        Sweep_Info(i).PinName = Split(argv(i), "@")(1)
        Sweep_Info(i).from = Split(argv(i), "@")(3)
        Sweep_Info(i).stop = Split(argv(i), "@")(4)
        Sweep_Info(i).step = Split(argv(i), "@")(5)
        If (CDbl(Sweep_Info(i).stop) < CDbl(Sweep_Info(i).from)) Then Sweep_Info(i).step = "-" & Sweep_Info(i).step
        Sweep_Info(i).Loop_Index_Name = Split(argv(i), "@")(6)
        Sweep_Info(i).Loop_count = Split(argv(i), "@")(7)
        Sweep_Info(i).Key = Split(argv(i), "@")(8)
    Next i
    
    Sweep_Count = CLng(Abs((Sweep_Info(0).stop - Sweep_Info(0).from) / Sweep_Info(0).step)) + 1
    Loop_count = CLng(Sweep_Info(0).Loop_count)
    
    If (TheExec.Flow.var(Sweep_Info(0).Loop_Index_Name).Value = Loop_count - 1 And Loop_Index = Sweep_Count - 1) Then
        
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "====================================Start Calc Mean===================================="

        
        For j = 0 To Sweep_Count - 1
            dict_key = ""
            For Sweep_index = 0 To UBound(Sweep_Info)
                If dict_key = "" Then
                    dict_key = Replace(CStr(CDbl(Sweep_Info(Sweep_index).from) + CDbl(Sweep_Info(Sweep_index).step) * j), ".", "p")
                Else
                    dict_key = dict_key & "_" & Replace(CStr(CDbl(Sweep_Info(Sweep_index).from) + CDbl(Sweep_Info(Sweep_index).step) * j), ".", "p")
                End If
            Next Sweep_index
            
            dict_key = Sweep_Info(0).Key & "_" & dict_key
            For Each site In TheExec.sites.Active
                DSP_Result.CreateConstant 0, Loop_count, DspDouble
            Next site
            For i = 0 To Loop_count - 1
                
                DSP_Capture = GetStoredCaptureData(dict_key & "_" & CStr(i))
                For Each site In TheExec.sites.Active
                    DSP_Capture = DSP_Capture.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, DSP_Capture.SampleSize, 0, Bit0IsMsb)
                'For Each site In TheExec.sites.Active
                    DSP_Result.Element(i) = DSP_Capture.Element(0)
                Next site
            Next i
            For Each site In TheExec.sites.Active
                Sweep_Mean = DSP_Result.CalcMean
            Next site
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", dict_key & "Mean", CInt(j))
                        
            TheExec.Flow.TestLimit resultVal:=Sweep_Mean, Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next j
    End If
    
'    index_name = argv(0)
'    Sweep_Dictionary = argv(1)
'    Loop_Count = argv(2)
'    Loop_Index = theexec.Flow.var(index_name).Value
'
'
'    If (Loop_Index = Loop_Count - 1) Then
'        DSP_Result.CreateConstant 0, Loop_Count, DspLong
'        Dictionary_Key = Split(Sweep_Dictionary, ":")
'
'        For Each key In Dictionary_Key
'            DSP_Capture = GetStoredCaptureData(CStr(key))
'
'        Next key
'
'
'    End If
    
    

End Function



Public Function Report_ALG_TName_From_Instance(ByRef TNameSeg() As String, MeasType As String, PinName As String, Tname As String, Optional TestSeqNum As Integer, Optional k As Long)

        'Modify from M9 module
        Dim instanceName As String
        Dim InstanceName_WO_Pset As String
        Dim InstNameSegs() As String
        Dim InTNameSegs() As String
        ReDim TNameSeg(9) As String

        instanceName = UCase(TheExec.DataManager.instanceName)
        InTNameSegs = Split(gl_Tname_Alg, ",")
        
        ''20190107 - Global name for saving Customize Subblock name
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
            If gl_Current_Instance_Tname_subblock <> "" Then            ''20190107 - Global name for saving Customize Subblock name
                TNameSeg(3) = gl_Current_Instance_Tname_subblock
            Else
                TNameSeg(3) = InstNameSegs(1)
            End If
        End If
        
        If TheExec.DataManager.instanceName Like "IDS_*IDS*" Then
            TNameSeg(9) = CStr(TestSeqNum)
        Else
            TNameSeg(9) = CStr(gl_Tname_Alg_Index)
        End If
        TNameSeg(5) = Replace(PinName, "_", "")
        
        If gl_Tname_Alg <> "" Then
            If UBound(InTNameSegs) < gl_Tname_Alg_Index Then
                TNameSeg(6) = "X"
                Else
                TNameSeg(6) = InTNameSegs(gl_Tname_Alg_Index)
            End If
        End If
        
        If MeasType = "I" Then
            TNameSeg(1) = TNameSeg(1) & MeasType
        Else
            TNameSeg(1) = "Calc"
        End If
          
        
        If gl_Sweep_Name <> "" Then
            If sweep_power_val_per_loop_count <> "" Then
                TNameSeg(8) = Replace(sweep_power_val_per_loop_count, ".", "p")
            Else
                TNameSeg(9) = TheExec.Flow.var(gl_Sweep_Name).Value
            End If
            
        Else
            If gl_Tname_Alg <> "" Then TNameSeg(9) = gl_Tname_Alg_Index
        End If
        
        If LCase(TNameSeg(4)) = "pp" Or LCase(TNameSeg(4)) = "dd" Or LCase(TNameSeg(4)) = "dp" Or LCase(TNameSeg(4)) = "cz" Or LCase(TNameSeg(4)) = "ht" Then TNameSeg(4) = "X"
        If LCase(TNameSeg(3)) = "pp" Or LCase(TNameSeg(3)) = "dd" Or LCase(TNameSeg(3)) = "dp" Or LCase(TNameSeg(3)) = "cz" Or LCase(TNameSeg(3)) = "ht" Then TNameSeg(3) = "X"
        
        If InStr(LCase(TheExec.DataManager.instanceName), "lapll") <> 0 Or InStr(LCase(TheExec.DataManager.instanceName), "usb2") <> 0 Or InStr(LCase(TheExec.DataManager.instanceName), "mipi") <> 0 Then
            TNameSeg(3) = UCase(TNameSeg(3))
            
            If InStr(LCase(TNameSeg(3)), "v") <> 0 Then
                TNameSeg(7) = Split(TNameSeg(3), "V")(0)
                TNameSeg(3) = "V" & Split(TNameSeg(3), "V")(1)
            ElseIf InStr(LCase(TNameSeg(3)), "t") <> 0 Then
                TNameSeg(7) = Split(TNameSeg(3), "T")(0)
                TNameSeg(3) = "T" & Split(TNameSeg(3), "T")(1)
            End If
        End If
        
        If InStr(LCase(TheExec.DataManager.instanceName), "lpdprx") <> 0 Then
            TNameSeg(3) = UCase(TNameSeg(3))
            
            If LCase(TNameSeg(3)) Like "rx2*" And InStr(TNameSeg(3), "L") <> 0 Then
                TNameSeg(7) = "L" & Split(TNameSeg(3), "L")(1)
                
                If UCase(TNameSeg(6)) Like "LN*" Then
                    TNameSeg(6) = Replace(UCase(TNameSeg(6)), "LN" & Split(TNameSeg(3), "L")(1), "")
                End If
                TNameSeg(3) = Split(TNameSeg(3), "L")(0)
            End If
        End If
        
        If InStr(LCase(TheExec.DataManager.instanceName), "pcie") <> 0 Then

                If UCase(TNameSeg(6)) Like "LN*" Then
                    TNameSeg(6) = UCase(Replace(TNameSeg(6), "_", ""))
                    TNameSeg(7) = UCase(Mid(TNameSeg(6), 1, 3))
                    TNameSeg(6) = UCase(Mid(TNameSeg(6), 4, Len(TNameSeg(6)) - 3))
                End If
                
            
        End If
        
        If InStr(LCase(TheExec.DataManager.instanceName), "amp") <> 0 Then
            TNameSeg(3) = UCase(TNameSeg(3))
            
            If LCase(TNameSeg(6)) Like "ddr*" Then
                TNameSeg(7) = UCase(Mid(TNameSeg(6), 1, 4))
                TNameSeg(6) = UCase(Mid(TNameSeg(6), 5, Len(TNameSeg(6)) - 4))
            End If
        End If
        
        '-------------------------------Pin Split--------------------------------------------------------
        If InStr(LCase(TheExec.DataManager.instanceName), "amp") <> 0 Then
            If LCase(TNameSeg(5)) Like "ddr*" Then
                TNameSeg(5) = Replace(TNameSeg(5), "_", "")
                TNameSeg(7) = UCase(Mid(TNameSeg(5), 1, 4))
                TNameSeg(5) = UCase(Mid(TNameSeg(5), 5, Len(TNameSeg(5)) - 4))
            End If
        End If
        '-------------------------------Pin Split--------------------------------------------------------
        

        '[X2]_[X3]_[X4]
'        If UBound(InstNameSegs) >= 5 Then
'                TNameSeg(6) = InstNameSegs(UBound(InstNameSegs) - 3)        '[X2]
'                TNameSeg(7) = InstNameSegs(UBound(InstNameSegs) - 2)        '[X3]
'                TNameSeg(8) = InstNameSegs(UBound(InstNameSegs) - 1)        '[X4]
'        ElseIf UBound(InstNameSegs) >= 4 Then
'                TNameSeg(6) = InstNameSegs(UBound(InstNameSegs) - 2)        '[X2]
'                TNameSeg(7) = InstNameSegs(UBound(InstNameSegs) - 1)        '[X3]
'        ElseIf UBound(InstNameSegs) >= 3 Then
'                TNameSeg(6) = InstNameSegs(UBound(InstNameSegs) - 1)        '[X2]
'        End If

'        Call SetupDatalogFormat(TestNameW:=90, PatternW:=100)
    gl_Tname_Alg_Index = gl_Tname_Alg_Index + 1
    
    Call SetupDatalogFormat(80, 100)
    
End Function

Public Function Calc_ADC_Convert_Average(argc As Integer, argv() As String) As Long

Dim Addvalue As Double
Dim Minuspin As String
Dim RefferenceCode_string As String
    Dim Transfer_Code_string As String
    Dim i As Long
    Dim j As Long
    Dim RefferenceCode As New DSPWave
    Dim Transfer_Code As New DSPWave
    Dim ADC_code As New DSPWave
    Dim ADC_code_DEC As New SiteDouble
    Dim ADC_code_average As New DSPWave
    Dim ADC_code_average_DEC As New SiteDouble
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    Dim averagecount As Long
    Dim temp_RefferanceCode As New DSPWave
    Dim temp_RefferanceCode_DEC As New SiteDouble
    Dim MinusValue As Double
    Dim Tname_String As String
    
    Dim RefferenceCode_DEC As New SiteDouble
    Dim Transfer_Code_DEC As New SiteDouble
    
    RefferenceCode.CreateConstant 0, 1, DspDouble
    Transfer_Code.CreateConstant 0, 1, DspDouble
    temp_RefferanceCode.CreateConstant 0, 1, DspDouble
    ADC_code_average.CreateConstant 0, 1, DspDouble
    
    ''''''''''''''''CalcArg:2,127,VDD12_PCIE,adc_offset_adc0_0,adc_offset_ adc1_1,ctlevos_in_p_adc0_2,ctlevos_in_p_adc1_3,ctlevos_in_n_adc0_4,ctlevos_in_n_adc1_5,vss_adc0_6,vss_adc1_7,ctlevos_in_cm_adc0_8'''''
    
    averagecount = argv(0)
    Addvalue = argv(1)
    Minuspin = argv(2)
    MinusValue = ProcessEvaluateDCSpec(Minuspin)
    
    '////////////////////////////////////// for DDR/SOC PLL referece code
    If LCase(argv(3)) = "dummyvref" Then
        Dim Dummytemp As New SiteDouble
            Dummytemp = 127
         Call AddStoredData("DummyVref" & "_para", Dummytemp)
    End If
    '/////////////////////////////////////////
    For i = 3 To 3 + averagecount - 1
        RefferenceCode_string = argv(i)
        RefferenceCode_DEC = GetStoredData(RefferenceCode_string & "_para")
        Set temp_RefferanceCode_DEC = temp_RefferanceCode_DEC.Add(RefferenceCode_DEC)
    Next i
    Set temp_RefferanceCode_DEC = temp_RefferanceCode_DEC.Divide(averagecount)
    If averagecount >= 2 Then
        TestNameInput = Report_TName_From_Instance(CalcC, Left(RefferenceCode_string, Len(RefferenceCode_string) - 3), Left(RefferenceCode_string, Len(RefferenceCode_string) - 3), CInt(i - 3))
        TheExec.Flow.TestLimit resultVal:=temp_RefferanceCode_DEC, ForceResults:=tlForceFlow, Tname:=TestNameInput
    End If
    
    For j = 3 + averagecount To UBound(argv) Step averagecount
        Set ADC_code_average_DEC = ADC_code_average_DEC.Multiply(0)
        For i = 0 To averagecount - 1
            Transfer_Code_string = argv(j + i)
            Transfer_Code_DEC = GetStoredData(Transfer_Code_string & "_para")
            Set ADC_code_DEC = Transfer_Code_DEC.Add(Addvalue).Add(temp_RefferanceCode_DEC.Multiply(-1)).Divide(256).Multiply(0.5).Multiply(MinusValue).Add(0.25 * (MinusValue))
            Set ADC_code_average_DEC = ADC_code_average_DEC.Add(ADC_code_DEC)
        Next i
        Set ADC_code_average_DEC = ADC_code_average_DEC.Divide(averagecount)
        Tname_String = argv(j)
        If averagecount >= 2 Then
            TestNameInput = Report_TName_From_Instance(CalcC, Left(Tname_String, Len(Tname_String) - 3), Left(Tname_String, Len(Tname_String) - 3))
            TheExec.Flow.TestLimit resultVal:=ADC_code_average_DEC, ForceResults:=tlForceFlow, Tname:=TestNameInput
        Else
            TestNameInput = Report_TName_From_Instance(CalcC, Tname_String, "_ADC", CInt(i - 3))
            TheExec.Flow.TestLimit resultVal:=ADC_code_average_DEC, ForceResults:=tlForceFlow, Tname:=TestNameInput
        End If
    Next j

End Function


Public Function compensate_Volt(argc As Integer, argv() As String) As Long

Dim RAK_Pin As String
Dim Meas_pins() As String
Dim inVoh As New PinListData
Dim MeasureValue As New PinListData
Dim Compensate_V As New PinListData
Dim voh_temp As New SiteDouble
Dim forceV_temp(8) As Double
Dim RakV() As Double
Dim GetRakVal As Double
Dim Pin  As Variant
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim p As Integer
Dim OutputTname_format() As String
Dim TestNameInput As String
Dim site As Variant

    inVoh = GetStoredMeasurement(argv(1))
    ReDim Meas_pins(inVoh.Pins.Count - 1)
    ''get pins
    For i = 0 To inVoh.Pins.Count - 1
        Meas_pins(i) = inVoh.Pins.Item(i).Name
    'add pins
        MeasureValue.AddPin (Meas_pins(i))
        Compensate_V.AddPin (Meas_pins(i))
    'disconect digital pins first
        TheHdw.Digital.Pins(Meas_pins(i)).Disconnect
    Next i
    'force V
        For Each Pin In Meas_pins
            For Each site In TheExec.sites
                voh_temp = inVoh.Pins(Pin).Value(site)
                 If voh_temp < -1 Or voh_temp > 6 Then
                        TheExec.Datalog.WriteComment "the force value " & voh_temp & "is out of PPMU range -1V ~ 6V, bypass force PPMU and set measurement result to 9999"
                        MeasureValue.Pins(Pin).Value(site) = 9999
                 Else
                         With TheHdw.PPMU.Pins(Pin)
                            '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
                             .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_InitialValue_FI_Range
                            .Connect
                            .Gate = tlOn
                            If TheExec.TesterMode = testModeOffline Then voh_temp = 0.1
                            .ForceV voh_temp, 0.05
                            '' 20160108 - Only keep 1 force value but current range can be different for force pin
                        End With
                        TheHdw.Wait (100 * us)
                        MeasureValue.Pins(Pin).Value(site) = TheHdw.PPMU.Pins(Pin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                End If
            Next site
        Next Pin
    ' calc VOH
      For Each Pin In Meas_pins
          For Each site In TheExec.sites
            RAK_Pin = CStr(Pin)
            'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(RAK_Pin, site)
                GetRakVal = CurrentJob_Card_RAK.Pins(Pin).Value(site)
                Compensate_V.Pins(Pin).Value(site) = inVoh.Pins(Pin).Value(site) + Abs(GetRakVal * MeasureValue.Pins(Pin).Value(site))
          Next site
      Next Pin
      

      
      For Each Pin In Meas_pins
            If TheExec.TesterMode = testModeOffline Then voh_temp = 0.1
            TestNameInput = Report_TName_From_Instance("I", inVoh.Pins(Pin), "OutputCurrent", CInt(i))
            For Each site In TheExec.sites
                voh_temp = inVoh.Pins(Pin).Value(site)
                TheExec.Flow.TestLimit MeasureValue.Pins(Pin).Value(site), -9.999, 9.999, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.4f", Tname:=TestNameInput, ForceVal:=voh_temp, ForceUnit:=unitVolt, ForceResults:=tlForceNone
            Next site
      Next Pin

      
'      For Each pin In Meas_pins
            TestNameInput = Report_TName_From_Instance(CalcV, inVoh.Pins(Pin), "", CInt(i))
'            For Each Site In TheExec.sites
                TheExec.Flow.TestLimit Compensate_V, , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.4f", Tname:=TestNameInput, ForceResults:=tlForceFlow
'            Next Site
'      Next pin
'
 
      'disconnect pins
      Dim TestSeq As Long
      
      For TestSeq = 0 To (inVoh.Pins.Count - 1)
            With TheHdw.PPMU.Pins(Meas_pins(TestSeq))
                .ForceI 0, 0.05
                .ForceV 0, 0.05
                .Gate = tlOff
                .Disconnect
            End With
            TheHdw.Digital.Pins(Meas_pins(TestSeq)).Connect
      
      Next TestSeq
 
End Function

Public Function USB3_ADC(argc As Integer, argv() As String) As Long '----------------add by CSHO 20171227

Dim ADCcapcode As String
Dim USBvoltages As String
Dim USBvoltages2 As String
Dim devideV As Long
Dim ADC_voltages As String
Dim InputKey As String
Dim DSP_Input As New DSPWave
Dim DSP_Input_2 As New SiteDouble
Dim ADC_Output As New SiteDouble
Dim LSB As Double
Dim bitprint As String
Dim i As Integer
Dim ADC_voltages_final As Double
Dim MinusValue As Double
Dim OutputTname_format() As String
Dim TestNameInput As String


USBvoltages = argv(0)
MinusValue = ProcessEvaluateDCSpec(USBvoltages)

devideV = argv(1)
DSP_Input.CreateConstant 0, 1, DspDouble

For i = 2 To argc - 1
    InputKey = argv(i)
    Set DSP_Input = Nothing
    DSP_Input_2 = GetStoredData(InputKey & "_para")
    'DSP_Input = GetStoredCaptureData(InputKey & "_para")
    ADC_Output = ADC_Output.Add(DSP_Input_2)
    ADC_Output = ADC_Output.Multiply(MinusValue).Divide(devideV)
    TestNameInput = Report_TName_From_Instance(CalcC, InputKey, "_ADC", CInt(i - 2))
    TheExec.Flow.TestLimit resultVal:=ADC_Output, ForceResults:=tlForceFlow, Tname:=TestNameInput
Next i

End Function

Public Function Calc_memcheck(argc As Integer, argv() As String) As Long
    
    Dim temp_dsp As New DSPWave
    Dim dataWave As New DSPWave
    Dim hexWave As New DSPWave
    Dim i As Long
    Dim CurSite As Variant
    Dim HexStr As String
    Dim DataFormat As String: DataFormat = "Hex"
    Dim cap_dec_data As New SiteLong
    Dim dc_read As New SiteLong: dc_read = 1
    Dim j As Integer
    Dim first_flag As New SiteBoolean
    Dim second_flag As New SiteBoolean
    Dim Dec_Str_All(3) As New DSPWave

        
        first_flag = False
        second_flag = False
        For i = 0 To argc - 1
            Dec_Str_All(i).CreateConstant 0, 4, DspLong
        Next i
    For i = 0 To argc - 1
        temp_dsp = GetStoredCaptureData(argv(i))
        For Each CurSite In TheExec.sites
            HexStr = ""
            ' convert bits to hex formatted stream
            Dim bin_str As String
            bin_str = ""
               For j = 0 To temp_dsp.SampleSize - 1
                    bin_str = bin_str & temp_dsp.Element(j)
            Next j
            bin_str = StrReverse(bin_str)
            TheExec.Datalog.WriteComment "(MSB -> LSB)"
            TheExec.Datalog.WriteComment bin_str

            hexWave = temp_dsp.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 4, 0, Bit0IsMsb)

            For j = (hexWave.SampleSize - 1) To 0 Step -1
                    HexStr = HexStr + Hex(hexWave.Element(j))
            Next j

               cap_dec_data(CurSite) = CLng("&H" & CStr(HexStr))
               dc_read(CurSite) = CLng("&H" & CStr(Hex(temp_dsp.Element(temp_dsp.SampleSize - 2)))) * dc_read(CurSite)  ' If two cycle HSC_READ are "1" then read=1
               TheExec.Datalog.WriteComment " Hex:  0x " & HexStr
               Dec_Str_All(i)(CurSite).Element(0) = cap_dec_data 'store data
        Next CurSite
        TheExec.Flow.TestLimit resultVal:=cap_dec_data, ForceResults:=tlForceFlow, Unit:=unitCustom, customUnit:=""   ', Tname:="FailBitCount", Unit:=unitNone, ScaleType:=scaleNone
    Next i


'/////////////// Judgement passing flag///////////////////
    
    Dim temp_HexStr As String
        For Each CurSite In TheExec.sites
            temp_HexStr = ""
            HexStr = ""
            For i = 0 To argc - 1
             temp_dsp = GetStoredCaptureData(argv(i))
                HexStr = CStr(Hex(Dec_Str_All(i)(CurSite).Element(0)))
                If first_flag = False Then
                    If HexStr = "E910" And temp_dsp(CurSite).Element(temp_dsp.SampleSize - 1) = 1 Then
                        first_flag = True
                        temp_HexStr = HexStr
                    ElseIf HexStr = "91E" And temp_dsp(CurSite).Element(temp_dsp.SampleSize - 1) = 0 Then
                        first_flag = True
                        temp_HexStr = HexStr
                    Else
                        first_flag = False
                    End If
                 Else
                    If HexStr <> temp_HexStr Then
                        Select Case HexStr
                            Case "E910":
                                If temp_dsp(CurSite).Element(temp_dsp.SampleSize - 1) = 1 Then second_flag = True
                            Case "91E":
                                If temp_dsp(CurSite).Element(temp_dsp.SampleSize - 1) = 0 Then second_flag = True
                        End Select
                    End If
                 End If
             Next i
        Next CurSite
'////////////////////////////////////////////////////////////
    TheExec.Flow.TestLimit resultVal:=dc_read, ForceResults:=tlForceFlow, Unit:=unitCustom, customUnit:=""
    TheExec.Flow.TestLimit resultVal:=second_flag, ForceResults:=tlForceFlow, Unit:=unitCustom, customUnit:=""

End Function

Public Function LP5_LB_PI(argc As Integer, argv() As String) As Long

   'New LP5 eye model 20190417
   
   Dim i As Long, j As Long, k As Long, L As Long
   Dim site As Variant
   Dim SplitByAt() As String
   Dim DSP_Captured() As New DSPWave
   Dim DSP_EYE() As New DSPWave
   Dim tmp_element As Long
   Dim tmp_name As String
   Dim EYE_arr() As Long
   Dim DSP_INV() As New DSPWave
   Dim DSP_CK() As New DSPWave
   Dim DSP_CKTemp() As New DSPWave
   Dim DSP_INVTemp() As New DSPWave
   ReDim DSP_INV(CStr(argc) - 1)
   ReDim DSP_CK(CStr(argc) - 1)
   ReDim DSP_INVTemp(CStr(argc) - 1)
   ReDim DSP_CKTemp(CStr(argc) - 1)
   
   Dim tmp_max_eye As Long
   Dim Eye_str As String
   Dim Eye_str_result() As New SiteVariant
   Dim Eye_str_long As New DSPWave
   Eye_str_long.CreateConstant 0, CLng(argc)
   ReDim Eye_str_result(CStr(argc))
   'ReDim Eye_str_long(CStr(argc)) As String
   Dim TestNameInput As String
   Dim DSP_Record() As New SiteVariant
   
   'argv(0) = "WCK0Sweep_2@WCK0Sweep_3@INVDQ0Sweep_0@INVDQ0Sweep_1"
   'argv(1) = "WCK1Sweep_2@WCK1Sweep_3@INVDQ1Sweep_0@INVDQ1Sweep_1"

   '' Split DSPWave captured to number of components of sweep
   
   
For Each site In TheExec.sites

   For i = 0 To argc - 1
      SplitByAt = Split(argv(i), "@") ' list of sweep names in order of concatination should be performed and INV if reverse is required
       ReDim Preserve DSP_Record((UBound(SplitByAt) + 1) * CStr(argc) - 1)
      ' Resize capture and final EYE DSPWaves to
      ReDim DSP_Captured(UBound(SplitByAt))
      ReDim DSP_EYE(UBound(SplitByAt))
      ReDim EYE_arr(UBound(SplitByAt))
      ReDim DSP_INV(UBound(SplitByAt))
      'ReDim Preserve DSP_INV(UBound(SplitByAt))

      Set DSP_EYE(i) = DSP_EYE(i).ConvertDataTypeTo(DspLong)
      Set DSP_INV(i) = DSP_INV(i).ConvertDataTypeTo(DspLong)
      Set DSP_CK(i) = DSP_CK(i).ConvertDataTypeTo(DspLong)
      ' ============== Prepare data capture for calculation ==============
        For j = 0 To UBound(SplitByAt)
        ' ======= INV Data order MSB -> LSB require inversion =======
            If SplitByAt(j) Like "INV*" Then
                tmp_name = Mid(SplitByAt(j), 4) ' remove INV from the beginning
                'tmp_name = SplitByAt(j)
                DSP_Captured(j) = GetStoredCaptureData(tmp_name)
                
                DSP_INVTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                
'                For k = DSP_Captured(j).SampleSize To 1 Step -1
                For k = 0 To DSP_Captured(j).SampleSize - 1
                                  
                    DSP_INVTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).Data) - k)
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = CStr(DSP_Record(i * (UBound(SplitByAt) + 1) + j)) & CStr(DSP_INVTemp(i).Element(k))
                   
                Next k
                 
 
                Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_INVTemp(i)) 'Merge all need flipped bit into one DSP
                
      
            Else
                
                DSP_Captured(j) = GetStoredCaptureData(SplitByAt(j))
                DSP_CKTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                
                For k = 0 To DSP_Captured(j).SampleSize - 1
                    DSP_CKTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).Data) - k)
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = DSP_Record(i * (UBound(SplitByAt) + 1) + j) & CStr(DSP_Captured(j).Element(UBound(DSP_Captured(j).Data) - k))
                Next k
               Set DSP_CK(i) = DSP_CK(i).Concatenate(DSP_CKTemp(i))
            
   
            End If
    
            If UCase(SplitByAt(j)) Like "*WCK*" Then
                DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "WCK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
            ElseIf UCase(SplitByAt(j)) Like "*CK*" Then
               DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "CK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
            Else
               DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "INV:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
            End If
            
        Next j
        
        '/////// INV_Cap reverse ////////////

        Dim iMidPt As Long
        Dim iUpper As Long
          iUpper = UBound(DSP_INV(i).Data)
          iMidPt = (UBound(DSP_INV(i).Data) - LBound(DSP_INV(i).Data)) \ 2 + LBound(DSP_INV(i).Data)
          For k = LBound(DSP_INV(i).Data) To iMidPt
              tmp_element = DSP_INV(i).Element(iUpper)
              DSP_INV(i).Element(iUpper) = DSP_INV(i).Element(k)
              DSP_INV(i).Element(k) = tmp_element
              iUpper = iUpper - 1
          Next k
        '////////////////////////////////////
        
        
      ' next sweep register sw0, sw1, ...

       ' ====== Concat EYE data ============

        Set DSP_EYE(i) = DSP_CK(i).Concatenate(DSP_INV(i))  ' Concatenate UnFlip code + INV Flip code
         'Set DSP_EYE(i) = DSP_INV(i).Concatenate(DSP_CK(i))
      '==========================================================================

      'theexec.Datalog.WriteComment "EYE " & i
      For k = 0 To DSP_EYE(i).SampleSize - 1
          'Debug.Print DSP_EYE(i).Element(k);
          If k = 0 Then
            Eye_str = DSP_EYE(i).Element(k)
            Else
            Eye_str = Eye_str & DSP_EYE(i).Element(k)
            End If

      Next k

      Eye_str_result(i) = Eye_str

      'Debug.Print
      'theexec.Datalog.WriteComment Eye_str
      '====== Calculate number of 'ones' in the EYE ===========
      tmp_max_eye = 0 ' reset tmp_max_eye
      For k = 0 To DSP_EYE(i).SampleSize - 1
         If DSP_EYE(i).Element(k) = 1 Then
            EYE_arr(i) = EYE_arr(i) + 1
         Else
            If tmp_max_eye < EYE_arr(i) Then
                tmp_max_eye = EYE_arr(i) ' update max eye width
            End If
            EYE_arr(i) = 0
         End If
      Next k

             If tmp_max_eye < EYE_arr(i) Then
                 tmp_max_eye = EYE_arr(i) ' update max eye width
             End If

      Eye_str_long(site).Element(i) = tmp_max_eye
      
      
'''      'T_Name Edit
'''      '**************************************
'''      TestNameInput = "EYEDDR" & CStr(i)
'''      TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(theexec.Flow.TestLimitIndex), 0)
'''      theexec.Flow.TestLimit resultVal:=Eye_str_long(i), FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow
'''      '**************************************

'      theexec.Datalog.WriteComment "EYE " & i & " width " & tmp_max_eye
     Next i ' next DDR bus : DQ0, DQ1, CA0, CA1 ...
    '=============================================================================
Next site

   'T_Name Edit
      '**************************************
    Dim TnumRecord As Long
    
    For i = 0 To argc - 1
        
            TnumRecord = TheExec.sites.Item(site).TestNumber
            TestNameInput = "EYEDDR" & CStr(i)
            TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(TheExec.Flow.TestLimitIndex), 0)
        
            For Each site In TheExec.sites
                TheExec.Flow.TestLimit resultVal:=Eye_str_long.Element(i), formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, scaletype:=scaleNoScaling
        '            TheExec.Flow.TestLimit lowVal:=mdll_low(i)(Site), resultVal:=Eye_str_long(Site).Element(i) * 4, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            'TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
    Next i
   '**************************************
    
    
For Each site In TheExec.sites
   TheExec.Datalog.WriteComment "/////////" & "Site: " & site & "/////////"
      Count = 0
    For L = 0 To argc - 1

        SplitByAt = Split(argv(L), "@")

        For i = 0 To UBound(SplitByAt)
           If SplitByAt(i) Like "INV*" Then
           TheExec.Datalog.WriteComment Mid(SplitByAt(i), 4)
           Else
           TheExec.Datalog.WriteComment SplitByAt(i)
           End If
           
           TheExec.Datalog.WriteComment DSP_Record(Count)
           Count = Count + 1
        Next i
        '****************************************
        TheExec.Datalog.WriteComment "EYE " & L
        TheExec.Datalog.WriteComment Eye_str_result(L)
        TheExec.Datalog.WriteComment "EYE " & L & " width " & Eye_str_long(site).Element(L)
    Next L

   
Next site

    TheExec.Datalog.WriteComment " ------------------ End ---"
    TheExec.Datalog.WriteComment "                           "




End Function
Public Function LP5_LB_DLL(argc As Integer, argv() As String) As Long

   'New LP5 eye model 20190417
   
    Dim i As Long, j As Long, k As Long, L As Long, z As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey As String
    Dim DSP_Captured() As New DSPWave
    Dim DSP_EYE() As New DSPWave
    Dim tmp_element As Long
    Dim tmp_name As String
    Dim EYE_arr() As Long
    Dim DSP_INV() As New DSPWave
    Dim DSP_CK() As New DSPWave
    Dim DSP_CKTemp() As New DSPWave
    Dim DSP_INVTemp() As New DSPWave
    ReDim DSP_INV(CStr(argc) - 1)
    ReDim DSP_CK(CStr(argc) - 1)
    ReDim DSP_INVTemp(CStr(argc) - 1)
    ReDim DSP_CKTemp(CStr(argc) - 1)
    Dim tmp_max_eye As Long
    Dim Eye_str As String
    'Dim Eye_str_result() As String
    Dim Eye_str_long As New DSPWave
    Eye_str_long.CreateConstant 0, CLng(argc - 1)
    Dim Eye_str_result() As New SiteVariant
    ReDim Eye_str_result(CStr(argc))
    'ReDim Eye_str_long(CStr(argc)) As String
    Dim TestNameInput As String
    Dim OutputTname_formatQQ() As String
    Dim Mdll_value() As String
    Dim Mdll_ChannelInfo() As String
    
    Dim mdll_12x8 As New DSPWave
    Dim mdll As New SiteDouble
    Dim mdll_low() As New SiteDouble
    ReDim mdll_low(CLng(argc - 2))
    Dim mdll_high() As New SiteDouble
    ReDim mdll_high(CLng(argc - 2))
    Dim Mdll_width As Long
   
   
    Dim DSP_Record() As New SiteVariant
   
   
   'argv(0) = "WCK0Sweep_2@WCK0Sweep_3@INVDQ0Sweep_0@INVDQ0Sweep_1"
   'argv(1) = "WCK1Sweep_2@WCK1Sweep_3@INVDQ1Sweep_0@INVDQ1Sweep_1"
   'argv(2) = "ch0_mdll_w210|ch0_mdll_w543|ch0_mdll_w76|ch1_mdll_w210|ch1_mdll_w543|ch1_mdll_w76"    'for mdll high low clac
   '' Split DSPWave captured to number of components of sweep
   
    For Each site In TheExec.sites

        For i = 0 To argc - 2

            SplitByAt = Split(argv(i), "@") ' list of sweep names in order of concatination should be performed and INV if reverse is required
            ReDim Preserve DSP_Record((UBound(SplitByAt) + 1) * CStr(argc) - 2)

          ' Resize capture and final EYE DSPWaves to
            ReDim DSP_Captured(UBound(SplitByAt))
            ReDim DSP_EYE(UBound(SplitByAt))
            ReDim EYE_arr(UBound(SplitByAt))
            ReDim DSP_INV(UBound(SplitByAt))
            Set DSP_EYE(i) = DSP_EYE(i).ConvertDataTypeTo(DspLong)
            Set DSP_INV(i) = DSP_INV(i).ConvertDataTypeTo(DspLong)
            Set DSP_CK(i) = DSP_CK(i).ConvertDataTypeTo(DspLong)
     
      ' ============== Prepare data capture for calculation ==============
            For j = 0 To UBound(SplitByAt)
        ' ======= INV Data order MSB -> LSB require inversion =======
                If SplitByAt(j) Like "INV*" Then
                    tmp_name = Mid(SplitByAt(j), 4) ' remove INV from the beginning
                    'tmp_name = SplitByAt(j)
                    DSP_Captured(j) = GetStoredCaptureData(tmp_name)
                    'DSP_Captured(j).CreateRandom 0, 1, 10, 1, DspLong '<- should be replaced by previous Line
                    'Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_Captured(j)) 'Merge all need flipped bit into one DSP
                    DSP_INVTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_INVTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).Data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = CStr(DSP_Record(i * (UBound(SplitByAt) + 1) + j)) & CStr(DSP_INVTemp(i).Element(k))
                    Next k
                    Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_INVTemp(i)) 'Merge all need flipped bit into one D
                Else
                    DSP_Captured(j) = GetStoredCaptureData(SplitByAt(j))
                    DSP_CKTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_CKTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).Data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = DSP_Record(i * (UBound(SplitByAt) + 1) + j) & CStr(DSP_Captured(j).Element(UBound(DSP_Captured(j).Data) - k))
                    Next k
                    Set DSP_CK(i) = DSP_CK(i).Concatenate(DSP_CKTemp(i))
                End If
            

                If UCase(SplitByAt(j)) Like "*WCK*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "WCK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                ElseIf UCase(SplitByAt(j)) Like "*CK*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "CK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                Else
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "INV:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                End If
                'ReDim Preserve DSP_Record(UBound(DSP_Record) + 1)
            Next j
     
        '/////// INV_Cap reverse ////////////

            Dim iMidPt As Long
            Dim iUpper As Long
            iUpper = UBound(DSP_INV(i).Data)
            iMidPt = (UBound(DSP_INV(i).Data) - LBound(DSP_INV(i).Data)) \ 2 + LBound(DSP_INV(i).Data)
            For k = LBound(DSP_INV(i).Data) To iMidPt
                tmp_element = DSP_INV(i).Element(iUpper)
                DSP_INV(i).Element(iUpper) = DSP_INV(i).Element(k)
                DSP_INV(i).Element(k) = tmp_element
                iUpper = iUpper - 1
            Next k
        '////////////////////////////////////

      ' next sweep register sw0, sw1, ...
       ' ====== Concat EYE data ============
            Set DSP_EYE(i) = DSP_CK(i).Concatenate(DSP_INV(i))  ' Concatenate UnFlip code + INV Flip code
            'Set DSP_EYE(i) = DSP_INV(i).Concatenate(DSP_CK(i))
      '==========================================================================
     
            'theexec.Datalog.WriteComment "EYE " & i
            For k = 0 To DSP_EYE(i).SampleSize - 1
                'Debug.Print DSP_EYE(i).Element(k);
                If k = 0 Then
                    Eye_str = DSP_EYE(i).Element(k)
                Else
                    Eye_str = Eye_str & DSP_EYE(i).Element(k)
                End If
            Next k
            Eye_str_result(i) = Eye_str

            'Debug.Print
            'theexec.Datalog.WriteComment Eye_str
            '====== Calculate number of 'ones' in the EYE ===========
            tmp_max_eye = 0 ' reset tmp_max_eye
            For k = 0 To DSP_EYE(i).SampleSize - 1
                If DSP_EYE(i).Element(k) = 1 Then
                    EYE_arr(i) = EYE_arr(i) + 1
                Else
                    If tmp_max_eye < EYE_arr(i) Then
                        tmp_max_eye = EYE_arr(i) ' update max eye width
                    End If
                    EYE_arr(i) = 0
                End If
                
                
                    If tmp_max_eye < EYE_arr(i) Then
                        tmp_max_eye = EYE_arr(i) ' update max eye width
                    End If
               
                
            Next k
            'Eye_str_long(i) = tmp_max_eye
            'Eye_str_long.CreateConstant 0, CLng(argc - 1)
            Eye_str_long(site).Element(i) = tmp_max_eye
        Next i ' next DDR bus : DQ0, DQ1, CA0, CA1 ...
    '=============================================================================
    Next site
    
    
    Dim DSP_Mdll_Temp As New DSPWave
    Dim DSP_Mdll_All() As New DSPWave
    Dim DSP_Mdll_Capture() As New DSPWave
    ReDim DSP_Mdll_All(argc - 1) As New DSPWave
    ReDim DSP_Mdll_Capture(argc - 1) As New DSPWave
    DSP_Mdll_Temp.CreateConstant 0, 1, DspLong
    For i = 0 To argc - 2
    '************Only for CACK read Mdll DSSCOUT and get Hi/Low limit****************
        DSP_Mdll_All(i).CreateConstant 0, 1, DspLong
        Mdll_ChannelInfo = Split(argv(UBound(argv)), "&")
        Mdll_value = Split(Mdll_ChannelInfo(i), "|")
        For z = 0 To UBound(Mdll_value)
            DSP_Mdll_Capture(i) = GetStoredCaptureData(Mdll_value(z))
'            rundsp.ConvertToLongAndSerialToParrel DSP_Mdll_Capture(i), DSP_Mdll_Capture(i).SampleSize, DSP_Mdll_Temp
            For Each site In TheExec.sites
                DSP_Mdll_Temp = DSP_Mdll_Capture(i).ConvertStreamTo(tldspParallel, DSP_Mdll_Capture(i).SampleSize, 0, Bit0IsMsb)
                DSP_Mdll_All(i).Element(0) = DSP_Mdll_All(i).Element(0) + DSP_Mdll_Temp.Element(0)
            Next site
        Next z
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "Site: " & site & "   ,octants code sum : " & DSP_Mdll_All(i).Element(0) & ",for Argc Number " & i + 1
            mdll_low(i)(site) = DSP_Mdll_All(i).Element(0) / 2     'fix 20190715
            mdll_high(i)(site) = DSP_Mdll_All(i).Element(0) * 2    ' fix 20190601
        Next site
    '*******************************************************************************
    Next i
    
    
    Dim TnumRecord As Long
    
    For i = 0 To argc - 2
        'For Each Site In TheExec.sites
            SplitByAt = Split(argv(i), "@")
            
            TnumRecord = TheExec.sites.Item(site).TestNumber
            
            TestNameInput = Left(SplitByAt(0), InStr(1, SplitByAt(0), "_")) & "EYEDDR" & CStr(i)
            
            TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(TheExec.Flow.TestLimitIndex), 0)
            
        'Next Site
           For Each site In TheExec.sites
''''''''''            TheExec.Flow.TestLimit LowVal:=mdll_low(i), HiVal:=mdll_high(i), resultVal:=Eye_str_long.Element(i) * 8, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, ScaleType:=scaleNoScaling
                 TheExec.Flow.TestLimit lowVal:=mdll_low(i), hiVal:=mdll_high(i), resultVal:=Eye_str_long.Element(i) * 8, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, scaletype:=scaleNoScaling
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
           ' TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
        'Next Site
    Next i
    
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "/////////" & "Site: " & site & "/////////"
        Count = 0
        For L = 0 To argc - 2
            SplitByAt = Split(argv(L), "@")
            For i = 0 To UBound(SplitByAt)
                If SplitByAt(i) Like "INV*" Then
                    TheExec.Datalog.WriteComment Mid(SplitByAt(i), 4)
                Else
                    TheExec.Datalog.WriteComment SplitByAt(i)
                End If
                TheExec.Datalog.WriteComment DSP_Record(Count)
                Count = Count + 1
            Next i
        '****************************************
        TheExec.Datalog.WriteComment "EYE " & L
        TheExec.Datalog.WriteComment Eye_str_result(L)
        TheExec.Datalog.WriteComment "EYE " & L & " width " & Eye_str_long(site).Element(L)
        Next L
    Next site
    TheExec.Datalog.WriteComment " ------------------ End ------------------"
    TheExec.Datalog.WriteComment "                           "

End Function


Public Function LP5_LB_RDDLL(argc As Integer, argv() As String) As Long

   'New LP5 eye model 20190417
   
    Dim i As Long, j As Long, k As Long, L As Long, z As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey As String
    Dim DSP_Captured() As New DSPWave
    Dim DSP_EYE() As New DSPWave
    Dim tmp_element As Long
    Dim tmp_name As String
    Dim EYE_arr() As Long
    Dim DSP_INV() As New DSPWave
    Dim DSP_CK() As New DSPWave
    Dim DSP_CKTemp() As New DSPWave
    Dim DSP_INVTemp() As New DSPWave
    ReDim DSP_INV(CStr(argc) - 1)
    ReDim DSP_CK(CStr(argc) - 1)
    ReDim DSP_INVTemp(CStr(argc) - 1)
    ReDim DSP_CKTemp(CStr(argc) - 1)
    Dim tmp_max_eye As Long
    Dim Eye_str As String
    'Dim Eye_str_result() As String
    Dim Eye_str_long As New DSPWave
    Eye_str_long.CreateConstant 0, CLng(argc - 1)
    Dim Eye_str_result() As New SiteVariant
    ReDim Eye_str_result(CStr(argc))
    'ReDim Eye_str_long(CStr(argc)) As String
    Dim TestNameInput As String
    Dim OutputTname_formatQQ() As String
    Dim Mdll_value() As String
    Dim Mdll_ChannelInfo() As String
    
    Dim mdll_12x8 As New DSPWave
    Dim mdll As New SiteDouble
    Dim mdll_low() As New SiteDouble
    ReDim mdll_low(CLng(argc - 2))
    Dim mdll_high() As New SiteDouble
    ReDim mdll_high(CLng(argc - 2))
    Dim Mdll_width As Long
   
   
    Dim DSP_Record() As New SiteVariant
   
   
   'argv(0) = "WCK0Sweep_2@WCK0Sweep_3@INVDQ0Sweep_0@INVDQ0Sweep_1"
   'argv(1) = "WCK1Sweep_2@WCK1Sweep_3@INVDQ1Sweep_0@INVDQ1Sweep_1"
   'argv(2) = "ch0_mdll_w210|ch0_mdll_w543|ch0_mdll_w76|ch1_mdll_w210|ch1_mdll_w543|ch1_mdll_w76"    'for mdll high low clac
   '' Split DSPWave captured to number of components of sweep
   
    For Each site In TheExec.sites

        For i = 0 To argc - 2

            SplitByAt = Split(argv(i), "@") ' list of sweep names in order of concatination should be performed and INV if reverse is required
            ReDim Preserve DSP_Record((UBound(SplitByAt) + 1) * CStr(argc) - 2)

          ' Resize capture and final EYE DSPWaves to
            ReDim DSP_Captured(UBound(SplitByAt))
            ReDim DSP_EYE(UBound(SplitByAt))
            ReDim EYE_arr(UBound(SplitByAt))
            ReDim DSP_INV(UBound(SplitByAt))
            Set DSP_EYE(i) = DSP_EYE(i).ConvertDataTypeTo(DspLong)
            Set DSP_INV(i) = DSP_INV(i).ConvertDataTypeTo(DspLong)
            Set DSP_CK(i) = DSP_CK(i).ConvertDataTypeTo(DspLong)
     
      ' ============== Prepare data capture for calculation ==============
            For j = 0 To UBound(SplitByAt)
        ' ======= INV Data order MSB -> LSB require inversion =======
                If SplitByAt(j) Like "INV*" Then
                    tmp_name = Mid(SplitByAt(j), 4) ' remove INV from the beginning
                    'tmp_name = SplitByAt(j)
                    DSP_Captured(j) = GetStoredCaptureData(tmp_name)
                    'DSP_Captured(j).CreateRandom 0, 1, 10, 1, DspLong '<- should be replaced by previous Line
                    'Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_Captured(j)) 'Merge all need flipped bit into one DSP
                    DSP_INVTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_INVTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).Data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = CStr(DSP_Record(i * (UBound(SplitByAt) + 1) + j)) & CStr(DSP_INVTemp(i).Element(k))
                    Next k
                    Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_INVTemp(i)) 'Merge all need flipped bit into one D
                Else
                    DSP_Captured(j) = GetStoredCaptureData(SplitByAt(j))
                    DSP_CKTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_CKTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).Data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = DSP_Record(i * (UBound(SplitByAt) + 1) + j) & CStr(DSP_Captured(j).Element(UBound(DSP_Captured(j).Data) - k))
                    Next k
                    Set DSP_CK(i) = DSP_CK(i).Concatenate(DSP_CKTemp(i))
                End If
            

                If UCase(SplitByAt(j)) Like "*WCK*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "WCK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                ElseIf UCase(SplitByAt(j)) Like "*CK*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "CK:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                ElseIf UCase(SplitByAt(j)) Like "*RDQS*" Then
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "RDQS:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                Else
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = "INV:" & DSP_Record(i * (UBound(SplitByAt) + 1) + j)
                End If
                'ReDim Preserve DSP_Record(UBound(DSP_Record) + 1)
            Next j
     
        '/////// INV_Cap reverse ////////////

            Dim iMidPt As Long
            Dim iUpper As Long
            iUpper = UBound(DSP_INV(i).Data)
            iMidPt = (UBound(DSP_INV(i).Data) - LBound(DSP_INV(i).Data)) \ 2 + LBound(DSP_INV(i).Data)
            For k = LBound(DSP_INV(i).Data) To iMidPt
                tmp_element = DSP_INV(i).Element(iUpper)
                DSP_INV(i).Element(iUpper) = DSP_INV(i).Element(k)
                DSP_INV(i).Element(k) = tmp_element
                iUpper = iUpper - 1
            Next k
        '////////////////////////////////////

      ' next sweep register sw0, sw1, ...
       ' ====== Concat EYE data ============
            Set DSP_EYE(i) = DSP_CK(i).Concatenate(DSP_INV(i))  ' Concatenate UnFlip code + INV Flip code
             'Set DSP_EYE(i) = DSP_INV(i).Concatenate(DSP_CK(i))
      '==========================================================================
     
            'theexec.Datalog.WriteComment "EYE " & i
            For k = 0 To DSP_EYE(i).SampleSize - 1
                'Debug.Print DSP_EYE(i).Element(k);
                If k = 0 Then
                    Eye_str = DSP_EYE(i).Element(k)
                Else
                    Eye_str = Eye_str & DSP_EYE(i).Element(k)
                End If
            Next k
            Eye_str_result(i) = Eye_str

            'Debug.Print
            'theexec.Datalog.WriteComment Eye_str
            '====== Calculate number of 'ones' in the EYE ===========
            tmp_max_eye = 0 ' reset tmp_max_eye
            For k = 0 To DSP_EYE(i).SampleSize - 1
                If DSP_EYE(i).Element(k) = 1 Then
                    EYE_arr(i) = EYE_arr(i) + 1
                Else
                    If tmp_max_eye < EYE_arr(i) Then
                        tmp_max_eye = EYE_arr(i) ' update max eye width
                    End If
                    EYE_arr(i) = 0
                End If
                
                
                    If tmp_max_eye < EYE_arr(i) Then
                        tmp_max_eye = EYE_arr(i) ' update max eye width
                    End If
              
            
            Next k
            'Eye_str_long(i) = tmp_max_eye
            'Eye_str_long.CreateConstant 0, CLng(argc - 1)
            Eye_str_long(site).Element(i) = tmp_max_eye
        Next i ' next DDR bus : DQ0, DQ1, CA0, CA1 ...
    '=============================================================================
    Next site
    
    
    Dim DSP_Mdll_Temp As New DSPWave
    Dim DSP_Mdll_All() As New DSPWave
    Dim DSP_Mdll_Capture() As New DSPWave
    ReDim DSP_Mdll_All(argc - 1) As New DSPWave
    ReDim DSP_Mdll_Capture(argc - 1) As New DSPWave
    DSP_Mdll_Temp.CreateConstant 0, 1, DspLong
    For i = 0 To argc - 2
    '************Only for CACK read Mdll DSSCOUT and get Hi/Low limit****************
        DSP_Mdll_All(i).CreateConstant 0, 1, DspLong
        Mdll_ChannelInfo = Split(argv(UBound(argv)), "&")
        Mdll_value = Split(Mdll_ChannelInfo(i), "|")
        For z = 0 To UBound(Mdll_value)
            DSP_Mdll_Capture(i) = GetStoredCaptureData(Mdll_value(z))
'            rundsp.ConvertToLongAndSerialToParrel DSP_Mdll_Capture(i), DSP_Mdll_Capture(i).SampleSize, DSP_Mdll_Temp
            For Each site In TheExec.sites
                DSP_Mdll_Temp = DSP_Mdll_Capture(i).ConvertStreamTo(tldspParallel, DSP_Mdll_Capture(i).SampleSize, 0, Bit0IsMsb)
                DSP_Mdll_All(i).Element(0) = DSP_Mdll_All(i).Element(0) + DSP_Mdll_Temp.Element(0)
            Next site
        Next z
        For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "Site: " & site & "   ,octants code sum : " & DSP_Mdll_All(i).Element(0) & ",for Argc Number " & i + 1
            mdll_low(i)(site) = DSP_Mdll_All(i).Element(0) / 8
            mdll_high(i)(site) = DSP_Mdll_All(i).Element(0)
        Next site
    '*******************************************************************************
    Next i
    
    
    Dim TnumRecord As Long
    
    For i = 0 To argc - 2
        
            SplitByAt = Split(argv(i), "@")
            
            TnumRecord = TheExec.sites.Item(site).TestNumber
            
            'TestNameInput = Left(SplitByAt(0), InStr(1, SplitByAt(0), "_")) & "EYEDDR" & CStr(i)
             TestNameInput = "EYE" & Left(SplitByAt(0), 7)
            
            TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(TheExec.Flow.TestLimitIndex), 0)
       
            For Each site In TheExec.sites
                'TheExec.Flow.TestLimit LowVal:=mdll_low(i)(Site), HiVal:=mdll_high(i)(Site), resultVal:=Eye_str_long(Site).Element(i) * 4, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, ScaleType:=scaleNoScaling
                TheExec.Flow.TestLimit lowVal:=mdll_low(i), hiVal:=mdll_high(i), resultVal:=Eye_str_long.Element(i), formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, scaletype:=scaleNoScaling
        '           TheExec.Flow.TestLimit lowVal:=mdll_low(i)(Site), resultVal:=Eye_str_long(Site).Element(i) * 4, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            'TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
    Next i
    
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "/////////" & "Site: " & site & "/////////"
        Count = 0
        For L = 0 To argc - 2
            SplitByAt = Split(argv(L), "@")
            For i = 0 To UBound(SplitByAt)
                If SplitByAt(i) Like "INV*" Then
                    TheExec.Datalog.WriteComment Mid(SplitByAt(i), 4)
                Else
                    TheExec.Datalog.WriteComment SplitByAt(i)
                End If
                TheExec.Datalog.WriteComment DSP_Record(Count)
                Count = Count + 1
            Next i
        '****************************************
        TheExec.Datalog.WriteComment "EYE " & L
        TheExec.Datalog.WriteComment Eye_str_result(L)
        TheExec.Datalog.WriteComment "EYE " & L & " width " & Eye_str_long(site).Element(L)
        Next L
    Next site
    TheExec.Datalog.WriteComment " ------------------ End ------------------"
    TheExec.Datalog.WriteComment "                           "

End Function
Public Function Calc_GPIO_DriverStrength(argc As Integer, argv() As String) As Long

    'GPIO_Pins=DS_Pins
    'Read Stored_PinListData=GPIO_IOL/IOH
    'Write Dict_name=gpio_iol_8
    
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim DS_Pins As New PinList: DS_Pins = argv(0)
    Dim Get_StoredData As New PinListData: Get_StoredData = GetStoredMeasurement(argv(1))
    Dim DS_Fuse_Name As String: DS_Fuse_Name = argv(2)
    Dim Fuse_Bit As Long: Fuse_Bit = argv(3)
    Dim DS_Data As New PinListData
    Dim Pin_Ary() As String
    Dim PinCnt As Long
    Dim DS_Data_DSPwave As New DSPWave
    Dim DS_Max As New SiteDouble, DS_Min As New SiteDouble, DS_Avg As New SiteDouble, DS_Result As New SiteDouble
    Dim DS_Max_Diff As New SiteDouble, DS_Min_Diff As New SiteDouble
    Dim Fuse_Bin As New DSPWave: Fuse_Bin.CreateConstant 0, Fuse_Bit
    Dim Fuse_Dec As New DSPWave: Fuse_Dec.CreateConstant 0, 1
    
    
    Call TheExec.DataManager.DecomposePinList(DS_Pins, Pin_Ary, PinCnt)
    DS_Data_DSPwave.CreateConstant 0, PinCnt
    
    For i = 0 To PinCnt - 1
        DS_Data.AddPin (Pin_Ary(i))
        DS_Data.Pins(Pin_Ary(i)) = Get_StoredData.Pins(Pin_Ary(i))
        For Each site In TheExec.sites
            DS_Data_DSPwave.Element(i) = DS_Data.Pins(Pin_Ary(i)).Value
        Next site
    Next i
    
    Dim DSP_Result As New DSPWave: DSP_Result.CreateConstant 0, UBound(Pin_Ary)
    
    For Each site In TheExec.sites
        DS_Data_DSPwave = DS_Data_DSPwave.Multiply(1000)
        DS_Avg = Format(DS_Data_DSPwave.CalcMean, "0.0")
        DSP_Result = DS_Data_DSPwave.Subtract(DS_Avg).Divide(1000)
        Fuse_Dec.Element(0) = DS_Avg.Abs.Multiply(10)
    Next site
 
    Call HardIP_Dec2Bin(Fuse_Bin, Fuse_Dec, Fuse_Bit)
    Call AddStoredCaptureData(DS_Fuse_Name, Fuse_Bin)
    
    Dim Pin As Variant
    Dim j As Long
    Dim testName() As String
        testName() = Split(CStr(argv(1)), "_")
    For j = 0 To UBound(Pin_Ary)
        For Each site In TheExec.sites
             If testName(1) = "ioh" Then
                TestNameInput = Report_TName_From_Instance(CalcI, CStr(Pin_Ary(j)), "CurrError" & Replace(DS_Fuse_Name, "_", ""), CInt(i))
                TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(testName(1)))
             Else
                TestNameInput = Report_TName_From_Instance(CalcI, CStr(Pin_Ary(j)), "CurrError" & Replace(DS_Fuse_Name, "_", ""), CInt(i))
                TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(testName(1)))
             End If
             If Fuse_Bit = 9 Then  'if is DS14
                TheExec.Flow.TestLimit DSP_Result.Data(j), hiVal:=0.005, lowVal:=-0.005, Tname:=TestNameInput, PinName:=Pin_Ary(j), Unit:=unitAmp, scaletype:=scaleMicro
            ElseIf Fuse_Bit = 8 Then 'if is DS8
                TheExec.Flow.TestLimit DSP_Result.Data(j), hiVal:=0.002, lowVal:=-0.002, Tname:=TestNameInput, PinName:=Pin_Ary(j), Unit:=unitAmp, scaletype:=scaleMicro
             Else
                TheExec.Flow.TestLimit DSP_Result.Data(j), hiVal:=0.001, lowVal:=-0.001, Tname:=TestNameInput, PinName:=Pin_Ary(j), Unit:=unitAmp, scaletype:=scaleMicro
            End If
        Next site
    Next j


    If testName(1) = "ioh" Then
        TestNameInput = Report_TName_From_Instance(CalcI, "", "CurrAvg", CInt(i))
        TestNameInput = Replace(TestNameInput, "IOL", UCase(testName(1)))
    Else
        TestNameInput = Report_TName_From_Instance(CalcI, "", "CurrAvg", CInt(i))
    End If
    
    
    If testName(1) = "ioh" Then
    
        If Fuse_Bit = 7 Then  'if is DS4
            TheExec.Flow.TestLimit DS_Avg.Divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", Unit:=unitAmp, lowVal:=-0.0127, hiVal:=0
        ElseIf Fuse_Bit = 8 Then   'if is DS8
            TheExec.Flow.TestLimit DS_Avg.Divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", Unit:=unitAmp, lowVal:=-0.0255, hiVal:=0
        Else  'if is DS14
            TheExec.Flow.TestLimit DS_Avg.Divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", Unit:=unitAmp, lowVal:=-0.0511, hiVal:=0
        End If
    Else
        If Fuse_Bit = 7 Then  'if is DS4
            TheExec.Flow.TestLimit DS_Avg.Divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", Unit:=unitAmp, hiVal:=0.0127, lowVal:=0
        ElseIf Fuse_Bit = 8 Then   'if is DS8
            TheExec.Flow.TestLimit DS_Avg.Divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", Unit:=unitAmp, hiVal:=0.0255, lowVal:=0
        Else  'if is DS14
            TheExec.Flow.TestLimit DS_Avg.Divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", Unit:=unitAmp, hiVal:=0.0511, lowVal:=0
        End If
      End If
End Function




Public Function Calc_MTR_BinStr2HexStr(ByVal BinStr As String, ByVal HexBit As Long) As String

    Dim i As Integer, j As Integer
    Dim BinStrLen As Long
    Dim HexMOD As Integer
    Dim HexStr As String
    Dim HexVal As String
    Dim HexLen As Long

    HexStr = ""
    
    BinStrLen = Len(BinStr)
    If (BinStrLen Mod (4)) > 0 Then
        HexLen = (BinStrLen \ 4) + 1
    Else
        HexLen = BinStrLen \ 4
    End If
    
    If HexBit > HexLen Then
        HexLen = HexBit
    End If

    HexMOD = HexLen * 4 - BinStrLen
    
    If HexMOD > 0 Then
        For i = 0 To HexMOD - 1
            BinStr = "0" & BinStr
        Next i
    End If

    For i = 0 To HexLen - 1
        If Mid(BinStr, i * 4 + 1, 4) = "0000" Then
            HexVal = "0"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0001" Then
            HexVal = "1"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0010" Then
            HexVal = "2"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0011" Then
            HexVal = "3"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0100" Then
            HexVal = "4"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0101" Then
            HexVal = "5"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0110" Then
            HexVal = "6"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0111" Then
            HexVal = "7"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1000" Then
            HexVal = "8"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1001" Then
            HexVal = "9"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1010" Then
            HexVal = "A"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1011" Then
            HexVal = "B"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1100" Then
            HexVal = "C"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1101" Then
            HexVal = "D"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1110" Then
            HexVal = "E"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1111" Then
            HexVal = "F"
        Else
            HexVal = "X"
        End If

        HexStr = HexStr & HexVal
    Next i

    Calc_MTR_BinStr2HexStr = HexStr

End Function

Public Function Calc_Voff_t6p2_MetrologyGR(argc As Integer, argv() As String) As Long
    Dim Dict_V0 As String
    Dim Dict_V1 As String
    Dim Dict_V2 As String
    Dim Fuse_BitCount As Double
    Dim Fuse_Voff_Round As String
    Dim Dict_Ratio_off_Per As String
    
    Dim Input_V0 As New PinListData
    Dim Input_V1 As New PinListData
    Dim Input_V2 As New PinListData
    
    Dim Voff_PinListData As New PinListData
    Dim Voff_PinListData_Round As New PinListData
    
    Dim UnSinged_Voff_Round As New DSPWave
    UnSinged_Voff_Round.CreateConstant 0, 1, DspDouble
    
    Dim Ratio_off_PinListData As New PinListData
    Dim Ratio_off_Per_DSP As New DSPWave
    Ratio_off_Per_DSP.CreateConstant 0, 1, DspDouble
    Dim site As Variant
    
    Dict_V0 = argv(0)
    Dict_V1 = argv(1)
    Dict_V2 = argv(2)
    Fuse_BitCount = argv(3)
    Fuse_Voff_Round = argv(4)
    Dict_Ratio_off_Per = argv(5)
    
    Input_V0 = GetStoredMeasurement(Dict_V0)
    Input_V1 = GetStoredMeasurement(Dict_V1)
    Input_V2 = GetStoredMeasurement(Dict_V2)
    
    Voff_PinListData.AddPin (Input_V1.Pins(0))
    Voff_PinListData = Input_V1.Pins(0).Subtract(Input_V0.Pins(0))
    Voff_PinListData = Voff_PinListData.Math.Divide(0.001)
    
    Voff_PinListData_Round.AddPin (Input_V1.Pins(0))
    Voff_PinListData_Round = Voff_PinListData.Pins(0).Divide(0.5)
    For Each site In TheExec.sites
        Voff_PinListData_Round.Pins(0).Value(site) = CDbl(FormatNumber(Voff_PinListData_Round.Pins(0).Value(site), 0))
    Next site
    
    Ratio_off_PinListData.AddPin (Input_V1.Pins(0))
    Ratio_off_PinListData = Input_V1.Pins(0).Divide(Input_V2.Pins(0)).Divide(2).Subtract(1)
    For Each site In TheExec.sites
            Ratio_off_Per_DSP(site).Element(0) = Ratio_off_PinListData.Pins(0).Value(site)
    Next site
    
    Call AddStoredCaptureData(Dict_Ratio_off_Per, Ratio_off_Per_DSP)
                                       
    TheExec.Flow.TestLimit resultVal:=Voff_PinListData.Pins(0), ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=Voff_PinListData_Round.Pins(0), ForceResults:=tlForceFlow
    For Each site In TheExec.sites
        If Voff_PinListData_Round.Pins(0).Value(site) < 0 Then
            UnSinged_Voff_Round(site).Element(0) = Voff_PinListData_Round.Pins(0).Value(site) + (2 ^ Fuse_BitCount)
        Else
            UnSinged_Voff_Round(site).Element(0) = Voff_PinListData_Round.Pins(0).Value(site)
        End If
    Next site
    
    Call AddStoredCaptureData(Fuse_Voff_Round, UnSinged_Voff_Round)
    
    TheExec.Flow.TestLimit resultVal:=Ratio_off_PinListData.Pins(0), ForceResults:=tlForceFlow
    
End Function

Public Function Calc_Ratio_off_average_t6p2_MetrologyGR(argc As Integer, argv() As String) As Long
    Dim i As Long
    Dim DSPWave_Ratio_off_per() As New DSPWave
    ReDim DSPWave_Ratio_off_per(argc - 3) As New DSPWave
    Dim DSPWave_Average As New DSPWave
    DSPWave_Average.CreateConstant 0, 1
    Dim DSPWave_Round_Average As New DSPWave
    DSPWave_Round_Average.CreateConstant 0, 1
    Dim DSPWave_Unsinged_Round_Average As New DSPWave
    DSPWave_Unsinged_Round_Average.CreateConstant 0, 1
    Dim site As Variant
    Dim Sweep_Count As Double
    Dim Fuse_BitCount As Double
    Sweep_Count = argc - 2
    Fuse_BitCount = argv(argc - 1)
    Dim Dict_Unsinged_Round_Avg As String
    Dict_Unsinged_Round_Avg = argv(argc - 2)
    
    For i = 0 To argc - 3
        DSPWave_Ratio_off_per(i) = GetStoredCaptureData(argv(i))
        Call rundsp.DSP_Add(DSPWave_Average, DSPWave_Ratio_off_per(i))
    Next i
    Call rundsp.DSP_DivideConstant(DSPWave_Average, Sweep_Count)
    
    If TheExec.TesterMode = testModeOffline Then            'for offline run
        For Each site In TheExec.sites
            DSPWave_Average(site).Element(0) = -0.00060171
        Next site
    End If
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_Average.Element(0), Tname:="Ratio_off_per_avg", ForceResults:=tlForceNone
    
    For Each site In TheExec.sites
        DSPWave_Round_Average(site).Element(0) = FormatNumber(DSPWave_Average(site).Element(0) * 1600, 0)
    Next site
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_Round_Average.Element(0), Tname:="Round_Ratio_off_per_avg", ForceResults:=tlForceNone
    
    For Each site In TheExec.sites
        If DSPWave_Round_Average(site).Element(0) < 0 Then
            DSPWave_Unsinged_Round_Average(site).Element(0) = DSPWave_Round_Average(site).Element(0) + (2 ^ Fuse_BitCount)
        Else
            DSPWave_Unsinged_Round_Average(site).Element(0) = DSPWave_Round_Average(site).Element(0)
        End If
    Next site
    
    Call AddStoredCaptureData(Dict_Unsinged_Round_Avg, DSPWave_Unsinged_Round_Average)
    
End Function

Public Function Calc_2S_Complement_To_SignDec_DivConst(argc As Integer, argv() As String) As Long

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey_2S_BIN As String
    Dim DictKey_SIGN_DEC As String
    
    Dim DSP_DictKey_2S_BIN As New DSPWave
    Dim DSP_DictKey_SIGN_DEC() As New DSPWave

    ReDim DSP_DictKey_SIGN_DEC(argc - 1) As New DSPWave
    
    Dim testName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    Dim DivConst As Double
    
    Dim SL_BitWidth As New SiteLong
    '' Format: Dict_2S_Com_A@Dict_SignDec_A@TestName_A,Dict_2S_Com_B@Dict_SignDec_B@TestName_B
    For i = 0 To argc - 1
        SplitByAt = Split(argv(i), "@")
        DictKey_2S_BIN = SplitByAt(0)
        DictKey_SIGN_DEC = SplitByAt(1)
        testName = SplitByAt(2)
        DivConst = SplitByAt(3)
        
        DSP_DictKey_2S_BIN = GetStoredCaptureData(DictKey_2S_BIN)
        
''        Set DSP_DictKey_DEC = Nothing
''        DSP_DictKey_DEC.CreateConstant 0, 1, DspDouble
''        Call rundsp.BinToDec(DSP_DictKey_BIN, DSP_DictKey_DEC)
        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize
''            DSP_DictKey_DEC(0).Element(0) = 255
''            DSP_DictKey_DEC(1).Element(0) = 254
        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        'DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        Call AddStoredCaptureData(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))

        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
        
        Call rundsp.DSP_DivideConstant(DSP_DictKey_SIGN_DEC(i), DivConst)
        
        TheExec.Flow.TestLimit resultVal:=DSP_DictKey_SIGN_DEC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
End Function

Public Function Calc_BinStr2HexStr(ByVal BinStr As String, ByVal HexBit As Long) As String

    Dim i As Integer, j As Integer
    Dim BinStrLen As Long
    Dim HexMOD As Integer
    Dim HexStr As String
    Dim HexVal As String
    Dim HexLen As Long

    HexStr = ""
    
    BinStrLen = Len(BinStr)
    If (BinStrLen Mod (4)) > 0 Then
        HexLen = (BinStrLen \ 4) + 1
    Else
        HexLen = BinStrLen \ 4
    End If
    
    If HexBit > HexLen Then
        HexLen = HexBit
    End If

    HexMOD = HexLen * 4 - BinStrLen
    
    If HexMOD > 0 Then
        For i = 0 To HexMOD - 1
            BinStr = "0" & BinStr
        Next i
    End If

    For i = 0 To HexLen - 1
        If Mid(BinStr, i * 4 + 1, 4) = "0000" Then
            HexVal = "0"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0001" Then
            HexVal = "1"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0010" Then
            HexVal = "2"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0011" Then
            HexVal = "3"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0100" Then
            HexVal = "4"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0101" Then
            HexVal = "5"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0110" Then
            HexVal = "6"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "0111" Then
            HexVal = "7"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1000" Then
            HexVal = "8"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1001" Then
            HexVal = "9"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1010" Then
            HexVal = "A"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1011" Then
            HexVal = "B"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1100" Then
            HexVal = "C"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1101" Then
            HexVal = "D"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1110" Then
            HexVal = "E"
        ElseIf Mid(BinStr, i * 4 + 1, 4) = "1111" Then
            HexVal = "F"
        Else
            HexVal = "X"
        End If

        HexStr = HexStr & HexVal
    Next i

    Calc_BinStr2HexStr = HexStr

End Function


Public Function Calc_Metrology_Trim_Vdiff(argc As Integer, argv() As String) As Long

    Dim Dict_V1 As String
    Dim Dict_V2 As String
    Dim Dict_Vdiff As String
    
    Dim PinList_V1 As New PinListData
    Dim PinList_V2 As New PinListData
    Dim PinList_Vdiff As New PinListData
    
    Dict_V1 = argv(0)
    Dict_V2 = argv(1)
    Dict_Vdiff = argv(2)
    
    PinList_V1 = GetStoredMeasurement(Dict_V1)
    PinList_V2 = GetStoredMeasurement(Dict_V2)
    
    PinList_Vdiff.AddPin (PinList_V1.Pins(0))
    PinList_Vdiff.Pins(0) = PinList_V1.Math.Subtract(PinList_V2)
    
    Call AddStoredMeasurement(Dict_Vdiff, PinList_Vdiff)
    
    TheExec.Datalog.WriteComment ("Voltage Difference Calculation")
    
End Function
Public Function Calc_DigCap_Avg_Store(argc As Integer, argv() As String) As Long
    Dim i As Long
    'Dim site As Variant
    Dim DSPWave_Bin() As New DSPWave
    Dim DSPWave_Dec() As New DSPWave
    ReDim DSPWave_Bin(argc - 3) As New DSPWave
    ReDim DSPWave_Dec(argc - 3) As New DSPWave
    'DSPWave_Dec(0).CreateConstant 0, 1, DspDouble
    Dim DSPWave_AverageDec As New DSPWave
    DSPWave_AverageDec.CreateConstant 0, 1, DspLong
    Dim DSPWave_AverageBin As New DSPWave
    'DSPWave_AverageBin.CreateConstant 0, 18
    Dim TestNameInput As String
    
    Dim SL_BitWidth As New SiteLong
    
    Dim Is_2sComplement As Boolean: Is_2sComplement = False
    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
    
    If UCase(argv(argc - 1)) = "2SCOMPLEMENT" Then Is_2sComplement = True
    
    For i = 0 To argc - 3
        DSPWave_Bin(i) = GetStoredCaptureData(argv(i))
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSPWave_Bin(i)(site).SampleSize
        Next site
        
        DSPWave_Dec(i).CreateConstant 0, 1, DspLong
        
        If Is_2sComplement = True Then
            Call rundsp.DSP_2S_Complement_To_SignDec(DSPWave_Bin(i), SL_BitWidth, DSPWave_Dec(i))
        Else
            Call rundsp.BinToDec(DSPWave_Bin(i), DSPWave_Dec(i))
        End If
        
        For Each site In TheExec.sites
            DSPWave_Dec(i)(site) = DSPWave_Dec(i)(site).ConvertDataTypeTo(DspLong)
        Next site
        Call rundsp.DSP_Add(DSPWave_AverageDec, DSPWave_Dec(i))
    Next i
    
    Call rundsp.DSP_DivideConstant(DSPWave_AverageDec, argc - 2)

    For Each site In TheExec.sites
        DSPWave_AverageDec(site).Element(0) = FormatNumber(DSPWave_AverageDec(site).Element(0), 0)
    Next site
    
    Call rundsp.DSPWf_Dec2Binary(DSPWave_AverageDec, SL_BitWidth, DSPWave_AverageBin)
    'Call AddStoredCaptureData(argv(argc - 1), DSPWave_AverageDec)

    TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))

    TheExec.Flow.TestLimit resultVal:=DSPWave_AverageDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Call AddStoredCaptureData(argv(argc - 2), DSPWave_AverageBin)
    
    TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
End Function

Public Function Calc_MetrologyGR_t5p5(argc As Integer, argv() As String) As Long
    Dim Vsrp As New SiteDouble
    Dim Vsrn As New SiteDouble
    Dim Vdiff As New SiteDouble
    Dim TestNameInput As String
    Dim site As Variant
    Dim OutputTname_format() As String
    
    'GetStoredMeasurement
    For Each site In TheExec.sites
        Vsrp = GetStoredMeasurement(argv(0))
        Vsrn = GetStoredMeasurement(argv(1))
        Vdiff = Vsrn.Subtract(Vsrp).Divide(0.000005)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", "Rpsr")
    OutputTname_format = Split(TestNameInput, "_")
    OutputTname_format(6) = "Rpsr"
    TestNameInput = Merge_TName(OutputTname_format)
    TheExec.Flow.TestLimit resultVal:=Vdiff, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
End Function

Public Function Calc_MetrologyGR_t1p0(argc As Integer, argv() As String) As Long
    Dim Vsrp As New SiteDouble
    Dim Vsrn As New SiteDouble
    Dim Vdiff As New SiteDouble
    Dim TestNameInput As String
    Dim site As Variant
    Dim OutputTname_format() As String
    
    'GetStoredMeasurement
    For Each site In TheExec.sites
        Vsrp = GetStoredMeasurement(argv(0))
        Vsrn = GetStoredMeasurement(argv(1))
        Vdiff = Vsrp.Subtract(Vsrn)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", "Rpsr")
    OutputTname_format = Split(TestNameInput, "_")
    OutputTname_format(6) = "Vdiff"
    OutputTname_format(7) = CStr(TheExec.Flow.var("SrcCodeIndx").Value)
    TestNameInput = Merge_TName(OutputTname_format)
    TheExec.Flow.TestLimit resultVal:=Vdiff, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
End Function



Public Function Calc_PCIE_RXTERM(argc As Integer, argv() As String) As Long
    Dim DSP_RCAL_TX_DIV4_CODE As New DSPWave: DSP_RCAL_TX_DIV4_CODE = GetStoredCaptureData(argv(0))
    Dim DSP_RXTERM_CODE_Binary As New DSPWave
    Dim DSP_RXTERM_CODE_Dec As New DSPWave
    Dim TestNameInput As String
    For Each site In TheExec.sites.Active
        DSP_RXTERM_CODE_Binary = DSP_RCAL_TX_DIV4_CODE.Select(0, 1, DSP_RCAL_TX_DIV4_CODE.SampleSize - 1).Copy
        DSP_RXTERM_CODE_Dec = DSP_RXTERM_CODE_Binary.ConvertStreamTo(tldspParallel, DSP_RXTERM_CODE_Binary.SampleSize, 0, Bit0IsMsb)
    Next site
    TestNameInput = Report_TName_From_Instance(CalcC, "")
    TheExec.Flow.TestLimit resultVal:=DSP_RXTERM_CODE_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Call AddStoredCaptureData(argv(1), DSP_RXTERM_CODE_Binary)
End Function
Public Function Calc_DigCap_MeanWithVariance(argc As Integer, argv() As String) As Long

'Arguments: SegmentSize_8, DictKey1, DictKey2, DictKey3.....


    Dim i As Long, j As Long
    Dim site As Variant
    Dim SegmentSize As Long
    Dim SegmentCount As Long
    Dim IndexOffset As Long
    Dim GroupCount As Long
    Dim str_Temp As String
    
    Dim mean As Double
    Dim STDEV As Double
    Dim Variance As Double
    
    Dim TestNameInput As String
    
'    Dim Output_Mean() As New DSPWave
'    Dim Output_STDEV() As New DSPWave
'    Dim Output_Variance() As New DSPWave
    Dim CalcResult() As New PinListData
    
'    Dim Mean() As New SiteDouble
'    Dim STDEV() As New SiteDouble
'    Dim Variance() As New SiteDouble
    
    Dim DSPwave_temp As New DSPWave
    Dim DSPWave_UnitSegment() As New DSPWave
    Dim DSPWave_MergedSegment() As New DSPWave

    
    SegmentSize = CLng(Split(argv(0), "_")(1))
    GroupCount = argc - 1 ''Argv(0) define segment size, the others are group1, group2 ....
    ReDim DSPWave_UnitSegment(GroupCount - 1)
    ReDim DSPWave_MergedSegment(GroupCount - 1)
    
    ReDim CalcResult(GroupCount - 1)
    Dim MSB_First_Flag As Boolean
    
'    ReDim Output_Mean(GroupCount - 1)
'    ReDim Output_STDEV(GroupCount - 1)
'    ReDim Output_Variance(GroupCount - 1)
    
    If UBound(Split(argv(0), "_")) = 2 Then
        If Split(argv(0), "_")(2) = "MSB" Then MSB_First_Flag = True
    End If
    
    
    For Each site In TheExec.sites.Active
        For i = 0 To GroupCount - 1
            DSPWave_UnitSegment(i) = GetStoredCaptureData(argv(i + 1))
            CalcResult(i).AddPin "Mean"
            CalcResult(i).AddPin "STDEV"
            CalcResult(i).AddPin "Variance"
'            Output_Mean(i).CreateConstant 0, 1, DspDouble
'            Output_STDEV(i).CreateConstant 0, 1, DspDouble
'            Output_Variance(i).CreateConstant 0, 1, DspDouble
        Next i
        
        SegmentCount = CLng(DSPWave_UnitSegment(0).SampleSize) \ SegmentSize
        
    
        For i = 0 To GroupCount - 1 '' For example, SegmentSize = 8; Binary 8 Bit -> Dec 1 number (DSPWave_MergedSegment)
            If MSB_First_Flag Then
                DSPWave_MergedSegment(i) = DSPWave_UnitSegment(i).ConvertStreamTo(tldspParallel, SegmentSize, 0, Bit0IsLsb)
            Else
                DSPWave_MergedSegment(i) = DSPWave_UnitSegment(i).ConvertStreamTo(tldspParallel, SegmentSize, 0, Bit0IsMsb)
            End If
            mean = DSPWave_MergedSegment(i).CalcMeanWithStdDev(STDEV)
            CalcResult(i).Pins("Mean").Value(site) = mean
            CalcResult(i).Pins("STDEV").Value(site) = STDEV
            CalcResult(i).Pins("Variance").Value(site) = STDEV * STDEV
        Next i
        
    Next site
   
    
    TheExec.Datalog.WriteComment ""
    
    For i = 0 To GroupCount - 1
        'TestNameInput = Report_TName_From_Instance("C", "X", , CInt(i))
        
        For Each site In TheExec.sites.Active
            For j = 0 To SegmentCount - 1
                TheExec.Flow.TestLimit resultVal:=DSPWave_MergedSegment(i).Data(j), _
                Tname:=Report_TName_From_Instance("C", "X", argv(i + 1), Instance_Data.TestSeqNum + CInt(i * (SegmentCount + 2) + j)), ForceResults:=tlForceNone
            Next j
                TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("Mean").Value(site), _
                Tname:=Report_TName_From_Instance("C", "X", argv(i + 1) & "Mean", Instance_Data.TestSeqNum + CInt(i * (SegmentCount + 2) + SegmentCount)), ForceResults:=tlForceNone
                TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("Variance").Value(site), _
                Tname:=Report_TName_From_Instance("C", "X", argv(i + 1) & "Variance", Instance_Data.TestSeqNum + CInt(i * (SegmentCount + 2) + SegmentCount + 1)), ForceResults:=tlForceNone
        Next site
        
    Next i
    
    'ReDim DSPWave_Avg_Bin(argc - 3) As New DSPWave
    
'    Dim TestName As String
'    Dim Site As Variant
'    Dim Dict As String
'    Dim BitWidth As Long
'
'    For i = 0 To 1
'        DSPWave_Binary(i) = GetStoredCaptureData(argv(i))
'        Call rundsp.BinToDec(DSPWave_Binary(i), DSPWave_Dec(i))
'    Next i
'
'    TestName = argv(argc - 1)
'    BitWidth = argv(argc - 2)
'    Dict = argv(argc - 3)
'
'    For Each Site In TheExec.sites
'            DSPWave_Avg_Dec.Element(0) = Int(((DSPWave_Dec(0).Element(0) + DSPWave_Dec(1).Element(0)) / 2) + 0.5) ''Example 1). 78.4=>78  2). 78.5=79
'    Next Site
'    Call rundsp.DSPWaveDecToBinary(DSPWave_Avg_Dec, BitWidth, DSPWave_Avg_Bin)
'    Call AddStoredCaptureData(Dict, DSPWave_Avg_Bin)
'    Dim TestNameInput As String
'    Dim OutputTname_format() As String
'
'    TestNameInput = Report_TName_From_Instance("C", "X", , CInt(i))
'
'    TheExec.Flow.TestLimit resultVal:=DSPWave_Avg_Dec.Element(0), TName:=TestNameInput, ForceResults:=tlForceNone
    
End Function

Public Function Trim_Pll_Freq(argc As Integer, argv() As String) As Long


    Dim UseLimitTname As String
    Dim TestNameInput As String
    Dim SplitTrimFreq() As String
    Dim SplitVro() As String
    Dim SplitDCO() As String
    Dim SplitCap() As String
    Dim SplitBias() As String
    Dim i, j, k As Long
    Dim F_TrimComplete() As New SiteBoolean
    Dim F_Vro() As New SiteBoolean
    Dim DSPWaveFromDict As New DSPWave
    Dim DSPWaveDecType As New DSPWave
    Dim BiasTargetIndex() As New SiteLong
    Dim FinalBiasIndex As New SiteLong
    Dim VroTarget As New SiteDouble
    Dim FinalVroIndex As New SiteLong
    Dim VroFromDict As New SiteLong: VroFromDict = 0
    Dim VroVoltage() As New SiteDouble
    Dim CapDSPWave As New DSPWave
    Dim BiasDSPWave As New DSPWave
    Dim DCODSPWave As New DSPWave
    Dim FinalDSPWave As New DSPWave
    Dim Cap_arry() As Long
    Dim Bias_arry() As Long
    Dim DCO_arry() As Long
    'Trim_Pll_Freq@1000,Vro@300,DCO@6,Fcount_Cap@Fcount-Cap0-@Fcount-Cap1-@Fcount-Cap3-,Fcount_Bias@Bias0@Bias1@Bias2@Bias3@Bias4@Bias5@Bias6@Bias7,Storename_src_bit
    'argv(0) : Trim_Pll_Freq@1000
    'argv(1) : Vro@300
    'argv(2) : DCO@6
    'argv(3) : Fcount_Cap@Fcount-Cap0-@Fcount-Cap1-@Fcount-Cap3-
    'argv(4) : Fcount_Bias@Bias0@Bias1@Bias2@Bias3@Bias4@Bias5@Bias6@Bias7
    'argv(5) : Storename_src_bit
    
    SplitTrimFreq = Split(argv(0), "@") 'Trim_Pll_Freq@1000
    SplitVro = Split(argv(1), "@") 'Vro@300
    SplitDCO = Split(argv(2), "@") 'DCO@5
    SplitCap = Split(argv(3), "@") 'Fcount_Cap@Fcount-Cap0-@Fcount-Cap1-@Fcount-Cap3-
    SplitBias = Split(argv(4), "@") 'Fcount_Bias@Bias0@Bias1@Bias2@Bias3@Bias4@Bias5@Bias6@Bias7
    ReDim F_TrimComplete(UBound(SplitCap))
    ReDim F_Vro(UBound(SplitCap))
    ReDim VroVoltage(UBound(SplitCap))
    ReDim BiasTargetIndex(UBound(SplitCap))
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                        'Follow Customer Instruction to create the bit space
    ReDim Bias_arry(2)  'Bit0-2
    ReDim Cap_arry(1)   'Bit3-4
    ReDim DCO_arry(2)   'Bit5-7
    CapDSPWave.CreateConstant 0, (UBound(Cap_arry) + 1)
    BiasDSPWave.CreateConstant 0, (UBound(Bias_arry) + 1)
    DCODSPWave.CreateConstant 0, (UBound(DCO_arry) + 1)
    FinalDSPWave.CreateConstant 0, ((CapDSPWave.SampleSize) + (BiasDSPWave.SampleSize) + (DCODSPWave.SampleSize))
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    For i = 0 To UBound(SplitCap)
        VroVoltage(i) = 0
        F_TrimComplete(i) = False
        F_Vro(i) = False
    Next i
        
    For i = 1 To UBound(SplitBias)
        For j = 1 To UBound(SplitCap)
            For Each site In TheExec.sites.Active
                If Not F_TrimComplete(j) Then
                    DSPWaveFromDict = GetStoredCaptureData(LCase(SplitCap(j) & SplitBias(i)))
                    ''DSPWaveFromDict.Element(i + 4) = 1 '' Debug
                    DSPWaveDecType = DSPWaveFromDict.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
                    If DSPWaveDecType.Element(0) >= CInt(SplitTrimFreq(1)) Then
                        F_TrimComplete(j) = True
                        VroVoltage(j) = GetStoredMeasurement(CStr(LCase(SplitCap(j) & SplitBias(i) & SplitVro(0))))
                        ''VroVoltage(j) = VroVoltage(j).Add(j * 0.15) '' Debug
                        If VroVoltage(j) * 1000 >= SplitVro(1) Then
                            F_Vro(j) = True
                            VroTarget = VroVoltage(j)
                            FinalVroIndex = j
                            BiasTargetIndex(j) = i
                        Else
                            F_Vro(j) = False
                        End If
                    End If
                End If
            Next site
        Next j
    Next i
    
    For Each site In TheExec.sites.Active
        For k = 1 To UBound(SplitCap)
            If F_Vro(k) Then
                If VroVoltage(k) <= VroTarget Then
                    FinalVroIndex = k
                    VroTarget = VroVoltage(k)
                    FinalBiasIndex = BiasTargetIndex(k)
                End If
            End If
        Next k
        
        If Right(SplitBias(FinalBiasIndex), 1) = "s" Then
        
            TheExec.Datalog.WriteComment ("ERROR: No parameter can reach the target")
        Else
            FinalBiasIndex = CLng(Right(SplitBias(FinalBiasIndex), 1))
            FinalVroIndex = CLng(Mid(SplitCap(FinalVroIndex), (Len(SplitCap(FinalVroIndex)) - 1), 1))
'            CapDSPWave.CreateConstant 0, (UBound(Cap_arry) + 1)
'            BiasDSPWave.CreateConstant 0, (UBound(Bias_arry) + 1)
'            DCODSPWave.CreateConstant 0, (UBound(DCO_arry) + 1)
'            FinalDSPWave.CreateConstant 0, ((CapDSPWave.SampleSize) + (BiasDSPWave.SampleSize) + (DCODSPWave.SampleSize))
            Call Dec2Bin(FinalBiasIndex, Bias_arry())
            For i = 0 To UBound(Bias_arry)
                FinalDSPWave.Element(i) = Bias_arry((UBound(Bias_arry)) - i)
            Next i
            Call Dec2Bin(FinalVroIndex, Cap_arry())
            For j = 0 To UBound(Cap_arry)
                FinalDSPWave.Element(j + UBound(Bias_arry) + 1) = Cap_arry((UBound(Cap_arry)) - j) 'Bit start from 3
            Next j
            
           ' If Bias_arry(0) = 1 And Bias_arry(1) = 1 And Bias_arry(2) = 1 And FinalVroIndex = 0 Then
           
            UseLimitTname = CStr(Instance_Data.Tname(TheExec.Flow.TestLimitIndex))          ' Dylan Edit by 20190616
            TestNameInput = Report_TName_From_Instance("Calc", "X", UseLimitTname, 0, 0)
            If FinalBiasIndex = 7 And FinalVroIndex = 0 Then
                TheExec.Flow.TestLimit resultVal:=1, hiVal:=0, lowVal:=0, Tname:=TestNameInput, ForceResults:=tlForceFlowFail
            Else
                TheExec.Flow.TestLimit resultVal:=0, hiVal:=0, lowVal:=0, Tname:=TestNameInput, ForceResults:=tlForceFlowPass
            End If
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1       ' Edited for avoid useLimit index messy
            
            Call Dec2Bin(CLng(SplitDCO(1)), DCO_arry())
            For k = 0 To UBound(DCO_arry)
                FinalDSPWave.Element(k + UBound(Bias_arry) + UBound(Cap_arry) + 1 + 1) = DCO_arry((UBound(DCO_arry)) - k) 'Bit start from 5
            Next k
        End If
        TheExec.Datalog.WriteComment ("Final Bias Index : " & FinalBiasIndex)
        TheExec.Datalog.WriteComment ("Final Cap Index : " & FinalVroIndex)
        TheExec.Datalog.WriteComment ("DCO : " & CLng(SplitDCO(1)))
    Next site
    Call AddStoredCaptureData(LCase(argv(5)), FinalDSPWave)
End Function



Public Function Calc_DCC_Skew_Range_DSP(argc As Integer, argv() As String) As Long
 
    ''''Demo String : CH@CH0@CH1,DQ@DQ0@DQ1,CountIN@0x1F@0x0@1Fx0,Count100@0x1F@0x0@1Fx0,SkewFactor@0.5,InputFactor@1.5, PatternBit@13
    Dim i, j, k, y As Long
    Dim SplitCH() As String
    Dim SplitDQ() As String
    Dim SplitCountIN() As String
    Dim SplitCount100() As String
    Dim DC_Skew_Input_Array() As New SiteDouble
    Dim DC_Input_CLK_UP As New SiteDouble
    Dim DC_Input_CLK_NO_DCC As New SiteDouble
    Dim DC_Input_CLK_DOWN As New SiteDouble
    Dim DC_Skew_Input_CLK_UP As New SiteDouble
    Dim DC_Skew_Input_CLK_NO_DCC As New SiteDouble
    Dim DC_Skew_Input_CLK_DOWN As New SiteDouble
    Dim DCC_RANGE_UP As New SiteDouble
    Dim DCC_RANGE_DOWN As New SiteDouble
    Dim SkewFactor As Double
    Dim InputFactor As Double
    Dim DSPWaveTemp As New DSPWave
    Dim DSPWaveTemp1 As New DSPWave
    Dim DSPWaveDec As New DSPWave
    Dim DSPWaveDec1 As New DSPWave
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    SplitCH = Split(argv(0), "@")   'CH@CH0@CH1
    SplitDQ = Split(argv(1), "@")   'DQ@DQ0@DQ1
    SplitCountIN = Split(argv(2), "@")   'CountIN@0x1F@0x0@1Fx0
    SplitCount100 = Split(argv(3), "@")   'Count100@0x1F@0x0@1Fx0
    SkewFactor = Split(argv(4), "@")(1) 'Factor1@0.5
    InputFactor = Split(argv(5), "@")(1) 'Factor2@1.5
    DSPWaveTemp.CreateConstant 0, Split(argv(6), "@")(1), DspLong 'PatternBit@13
    DSPWaveTemp1.CreateConstant 0, Split(argv(6), "@")(1), DspLong 'PatternBit@13
    DSPWaveDec.CreateConstant 0, 1
    DSPWaveDec1.CreateConstant 0, 1
    ReDim DC_Skew_Input_Array(UBound(SplitCountIN) - 1)
    
    For i = 0 To (UBound(SplitCH) - 1)
        For j = 0 To (UBound(SplitDQ) - 1)
            For Each site In TheExec.sites.Active
                For k = 0 To (UBound(SplitCountIN) - 1)
''''''''''                    DSPWaveTemp = GetStoredCaptureData(SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCountIN(k + 1) & SplitCountIN(0))
''''''''''                    DSPWaveTemp1 = GetStoredCaptureData(SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCount100(k + 1) & SplitCount100(0))
                    DSPWaveDec = GetStoredCaptureData("2SDEC_" & SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCountIN(k + 1) & SplitCountIN(0))
                    DSPWaveDec1 = GetStoredCaptureData("2SDEC_" & SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCount100(k + 1) & SplitCount100(0))
                    
''''''''''                    DSPWaveDec = DSPWaveTemp.ConvertStreamTo(tldspParallel, Split(argv(6), "@")(1), 0, Bit0IsMsb)
''''''''''                    DSPWaveDec1 = DSPWaveTemp1.ConvertStreamTo(tldspParallel, Split(argv(6), "@")(1), 0, Bit0IsMsb)
                    If DSPWaveDec1.Element(0) = 0 Then
                        TheExec.Datalog.WriteComment ("Can't divide by 0")
                    Else
                        DC_Skew_Input_Array(k) = (DSPWaveDec.Element(0) / DSPWaveDec1.Element(0)) * SkewFactor
                    End If
                Next k
                DC_Input_CLK_UP = DC_Skew_Input_Array(0) + 0.5
                DC_Input_CLK_NO_DCC = DC_Skew_Input_Array(1) + 0.5
                DC_Input_CLK_DOWN = DC_Skew_Input_Array(2) + 0.5
                
                DC_Skew_Input_CLK_UP = DC_Skew_Input_Array(0)
                DC_Skew_Input_CLK_NO_DCC = DC_Skew_Input_Array(1)
                DC_Skew_Input_CLK_DOWN = DC_Skew_Input_Array(2)
                DCC_RANGE_UP = DC_Skew_Input_CLK_UP - DC_Skew_Input_CLK_NO_DCC
                DCC_RANGE_DOWN = DC_Skew_Input_CLK_DOWN - DC_Skew_Input_CLK_NO_DCC
''''''''''                TheExec.Datalog.WriteComment ("Site " & Site & " : " & SplitCH(i + 1) & "_" & SplitDQ(j + 1) & "_" & "DCC_RANGE_UP" & " = " & DCC_RANGE_UP)
''''''''''                TheExec.Datalog.WriteComment ("Site " & Site & " : " & SplitCH(i + 1) & "_" & SplitDQ(j + 1) & "_" & "DCC_RANGE_DOWN" & " = " & DCC_RANGE_DOWN)
            Next site
            
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Input_CLK", CInt(i), , "replace;7=UP", , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DC_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", Unit:=unitCustom, customUnit:="%"
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Input_CLK", CInt(i), , "replace;7=NODCC", , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DC_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", Unit:=unitCustom, customUnit:="%"
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Input_CLK", CInt(i), , "replace;7=DOWN", , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DC_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", Unit:=unitCustom, customUnit:="%"
            
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Skew_Input_CLK", CInt(i), , "replace;7=UP", , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DC_Skew_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", Unit:=unitCustom, customUnit:="%"
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Skew_Input_CLK", CInt(i), , "replace;7=NODCC", , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DC_Skew_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", Unit:=unitCustom, customUnit:="%"
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Skew_Input_CLK", CInt(i), , "replace;7=DOWN", , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DC_Skew_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", Unit:=unitCustom, customUnit:="%"
            
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DCC_RANGE_UP", CInt(i), , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DCC_RANGE_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", Unit:=unitCustom, customUnit:="%"
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DCC_RANGE_DOWN", CInt(i), , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DCC_RANGE_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling, formatStr:="%1.2f", Unit:=unitCustom, customUnit:="%"
        
        Next j
    Next i
End Function
Public Function Calc_DiCap_ParallelMode_For_IPPM(argc As Integer, argv() As String) As Long
' Edided by 20190613
'**********************************************************
' Format : DictionaryName & dc_meas @ t_meas
' Calculate Function : duty_cycle = (dc_meas_int*2) / (t_meas_int+2^17)
'**********************************************************

    Dim site As Variant
    Dim i, j, k As Integer
    Dim BitPosition As Long
    Dim RegSplit() As String
    Dim AssembleStr() As String
    Dim FormatSplit() As String
    Dim BinaryWave As New DSPWave
    Dim CalDSPWave As New DSPWave
    Dim duty_cycle() As New DSPWave
    Dim SplitDSPWave() As New DSPWave
    ReDim AssembleStr(argc - 1)
    ReDim duty_cycle(argc - 1)
    
    For i = 0 To argc - 1
        'BitPosition = 0
        FormatSplit = Split(argv(i), "&")
        RegSplit = Split(FormatSplit(1), "@")
        
        ReDim SplitDSPWave(UBound(RegSplit))
        duty_cycle(i).CreateConstant 0, 1, DspDouble
        BinaryWave = GetStoredCaptureData(FormatSplit(0))
        
        For Each site In TheExec.sites
            For j = 0 To BinaryWave.SampleSize - 1
                AssembleStr(i) = CStr(BinaryWave(site).Element(j)) & AssembleStr(i)
            Next j
            BitPosition = 0
            For j = 0 To UBound(RegSplit)
               
                SplitDSPWave(j).CreateConstant 0, 1, DspLong
                BinaryWave = BinaryWave(site).ConvertDataTypeTo(DspLong)
                CalDSPWave = BinaryWave(site).Select(CLng(BitPosition), 1, CLng(RegSplit(j)))
                SplitDSPWave(j) = CalDSPWave.ConvertStreamTo(tldspParallel, CLng(RegSplit(j)), 0, Bit0IsMsb)
                BitPosition = BitPosition + CLng(RegSplit(j))
            Next j
            duty_cycle(i).Element(0) = (SplitDSPWave(0).Element(0) * 2) / (SplitDSPWave(1).Element(0) + 2 ^ 17)
            duty_cycle(i).Element(0) = duty_cycle(i).Element(0) * 100
        Next site
        TheExec.Datalog.WriteComment FormatSplit(0) & " Binary Value : " & AssembleStr(i)
    Next i
    For i = 0 To argc - 1
        TheExec.Flow.TestLimit resultVal:=duty_cycle(i).Element(0), Tname:=FormatSplit(0), ForceResults:=tlForceFlow
    Next
End Function

Public Function prasing_ADC(RAW_DSP As DSPWave, ADC_bits As Long, sgmt_size) As DSPWave

    Dim new_DSP As New DSPWave
    Dim i As Long, j As Long
    Dim sgmt_cnt As Long
    sgmt_cnt = RAW_DSP.SampleSize / ADC_bits / sgmt_size
    
    For i = 0 To sgmt_cnt - 1
        For j = 0 To sgmt_size - 1
            If new_DSP.SampleSize = 0 Then
                new_DSP = RAW_DSP.Select(0, sgmt_size, ADC_bits).Copy
            Else
                new_DSP = new_DSP.Concatenate(RAW_DSP.Select(i * ADC_bits * sgmt_size + j, sgmt_size, ADC_bits).Copy)
            End If
        Next j
    Next i
    
    Set prasing_ADC = new_DSP.ConvertStreamTo(tldspParallel, ADC_bits, 0, Bit0IsMsb)

End Function

Public Function Calc_12bADC_MeanWithVariance(argc As Integer, argv() As String) As Long
'Alg::Calc_DigCap_MeanWithVariance(SegmentSize_32,CFG_FIFO_SDM5|2047,CFG_FIFO_SDM6|1024,CFG_FIFO_SDM7|3072)

    Dim i As Long, j As Long, k As Long
    Dim site As Variant
    Dim SegmentSize As Long: SegmentSize = CLng(Split(argv(0), "_")(1))
    Dim GroupCount As Long: GroupCount = argc - 1
    Dim ADC_bits As Long: ADC_bits = 12
    
    Dim mean As Double
    Dim STDEV As Double
    Dim TestNameInput As String
    Dim Limit_Dev As Long: Limit_Dev = 60
    
    Dim CalcResult() As New PinListData
    Dim Limit_Val() As Long
    Dim Key() As String
    ReDim CalcResult(GroupCount - 1)
    ReDim Limit_Val(GroupCount - 1)
    ReDim Key(GroupCount - 1)
    
    Dim DSPWave_Ori() As New DSPWave
    ReDim DSPWave_Ori(GroupCount - 1)
    
    Dim ADC_result() As New DSPWave
    ReDim ADC_result(GroupCount - 1)

    For i = 0 To GroupCount - 1
        Key(i) = Split(argv(i + 1), "|")(0)
        DSPWave_Ori(i) = GetStoredCaptureData(Key(i))
        CalcResult(i).AddPin "Mean"
        CalcResult(i).AddPin "Variance"
        CalcResult(i).AddPin "MaxErr"
        Limit_Val(i) = Split(argv(i + 1), "|")(1)
        Set ADC_result(i) = Nothing
        ADC_result(i) = prasing_ADC(DSPWave_Ori(i), ADC_bits, SegmentSize)
        'Call DebugPrintRawDigCap(DSPWave_Ori(i), SegmentSize)
    Next i

    For Each site In TheExec.sites.Active
        For i = 0 To GroupCount - 1
            mean = ADC_result(i).CalcMeanWithStdDev(STDEV)
            CalcResult(i).Pins("Mean").Value(site) = mean
            CalcResult(i).Pins("Variance").Value(site) = STDEV * STDEV
            CalcResult(i).Pins("MaxErr").Value(site) = ADC_result(i).CalcMaximumValue - ADC_result(i).CalcMinimumValue
        Next i
    Next site
    
    k = 0
    For i = 0 To GroupCount - 1
        For Each site In TheExec.sites.Active
            If True Then
                For j = 0 To ADC_result(i).SampleSize - 1
                    TheExec.Flow.TestLimit ADC_result(i).Data(j), Limit_Val(i) - Limit_Dev, Limit_Val(i) + Limit_Dev, _
                    Tname:=Report_TName_From_Instance("C", "X", Key(i), Instance_Data.TestSeqNum + CInt(k)), ForceResults:=tlForceNone
                    k = k + 1
                Next j
            End If
            TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("Mean").Value(site), _
            Tname:=Report_TName_From_Instance("C", "X", Key(i), Instance_Data.TestSeqNum + CInt(k + 1), , "replace;7=Mean"), ForceResults:=tlForceNone
            TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("Variance").Value(site), _
            Tname:=Report_TName_From_Instance("C", "X", Key(i), Instance_Data.TestSeqNum + CInt(k + 2), , "replace;7=Variance"), ForceResults:=tlForceNone
            TheExec.Flow.TestLimit resultVal:=CalcResult(i).Pins("MaxErr").Value(site), _
            Tname:=Report_TName_From_Instance("C", "X", Key(i), Instance_Data.TestSeqNum + CInt(k + 3), , "replace;7=MaxErr"), ForceResults:=tlForceNone
            k = k + 3
        Next site
    Next i
    
End Function

Public Function DebugPrintRawDigCap(InWave As DSPWave, sgmt_size As Long)
    
    Dim i As Long, j As Long
    Dim PrtStr As String
    Dim PartialWave As New DSPWave
    
    For i = 0 To (InWave.SampleSize / sgmt_size) - 1
        PrtStr = ""
        Set PartialWave = Nothing
        
        PartialWave = InWave.Select(i * sgmt_size, 1, sgmt_size).Copy
        For j = 0 To sgmt_size - 1
            PrtStr = PrtStr & PartialWave.Element(j)
        Next j
        TheExec.Datalog.WriteComment "Line" & Space(3 - Len(CStr(i))) & i & ":" & PrtStr
        
    Next i
End Function

Public Function Calc_ADC_average(argc As Integer, argv() As String) As Long

    Dim InWave As New DSPWave
    Dim i As Long, j As Long, k As Long
    Dim ADC_wave(2) As New DSPWave
    Dim Chk_wave(2) As New DSPWave
    Dim mean As Double
    Dim STDEV As Double
    Dim MaxErr As Double
    
    InWave = GetStoredCaptureData(argv(0))
    InWave = InWave.ConvertDataTypeTo(DspLong)
    
    For i = 0 To 2
        k = 0
        'Chk_wave(i) = InWave.Select(13 * (i + 1) - 1, 39, 256).Copy
        Chk_wave(i) = InWave.Select(3328 * i - 1 + 13, 13, 256).Copy
        Set ADC_wave(i) = Nothing
        ADC_wave(i) = ADC_wave(i).ConvertDataTypeTo(DspLong)
        For j = 0 To 255
'            ADC_wave(i) = ADC_wave(i).Concatenate(InWave.Select(39 * j + 13 * i, 1, 12).Copy)
             ADC_wave(i) = ADC_wave(i).Concatenate(InWave.Select(13 * j + 3328 * i, 1, 12).Copy)
        Next j
        
        ADC_wave(i) = ADC_wave(i).ConvertStreamTo(tldspParallel, 12, 0, Bit0IsMsb)
        
        For j = 0 To 255
            TheExec.Flow.TestLimit ADC_wave(i).Data(j), _
            Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k)), ForceResults:=tlForceNone
            TheExec.Flow.TestLimit Chk_wave(i).Data(j), _
            Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k), , "replace;7=Chk"), ForceResults:=tlForceNone
            k = k + 1
        Next j
        
        mean = ADC_wave(i).CalcMeanWithStdDev(STDEV)
        STDEV = STDEV * STDEV
        MaxErr = ADC_wave(i).CalcMaximumValue - ADC_wave(i).CalcMinimumValue
                    
        TheExec.Flow.TestLimit resultVal:=mean, Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k), , "replace;7=Mean"), ForceResults:=tlForceNone
        TheExec.Flow.TestLimit resultVal:=STDEV, Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k + 1), , "replace;7=Variance"), ForceResults:=tlForceNone
        TheExec.Flow.TestLimit resultVal:=MaxErr, Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k + 2), , "replace;7=MaxErr"), ForceResults:=tlForceNone
    Next i
    

End Function
Public Function Calc_D2D_MAX_MIN_V2(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim i, j, k As Integer
    Dim FirstLoop As String
    Dim SecondLoop As String
    Dim SplitCalStr() As String
    Dim TestNameInput As String
    Dim ValueMax As New SiteLong
    Dim ValueMin As New SiteLong
    Dim MaxCKDNL As New SiteLong
    Dim MinCKDNL As New SiteLong
    
    Dim DCKMAX As New SiteLong
    Dim DCKMin As New SiteLong
    Dim Idsvalue As New SiteDouble
    Dim DCKIdsvalue As New SiteDouble
    
    Dim IndexOfMinimumValue As Long
    Dim IndexOfMaximumValue As Long
    
    Dim SaveDCK_DSPWave As New DSPWave
    Dim DNLValue_DSPWave As New DSPWave
    Dim PreValue_DSPWave As New DSPWave
    Dim DeltaDCK_DSPWave As New DSPWave
    Dim DictValue_DSPWave As New DSPWave
    Dim DNLDCKValue_DSPWave As New DSPWave
    Dim SaveMeasNum_DSPWave As New DSPWave
    Dim SaveDeltaValue_DSPWave As New DSPWave
    
    
    
    FirstLoop = CStr(TheExec.Flow.var("SrcCodeIndx").Value)
    SecondLoop = CStr(TheExec.Flow.var("SrcCodeIndx1").Value)
''''''''''    Debug.Print "SrcCodeIndx Value : " & FirstLoop
''''''''''    Debug.Print "SrcCodeIndx1 Value : " & SecondLoop
    
    
''''''''''    LoopNum = SecondLoop = SrcCodeIndx1
''''''''''    LoopNum1 = FirstLoop = SrcCodeIndx
    
    For i = 0 To argc - 1
        SplitCalStr = Split(argv(i), "@")
        DictValue_DSPWave = GetStoredCaptureData(SplitCalStr(0))
        For Each site In TheExec.sites.Active
            DictValue_DSPWave = DictValue_DSPWave.ConvertDataTypeTo(DspLong)
            DictValue_DSPWave = DictValue_DSPWave.ConvertStreamTo(tldspParallel, DictValue_DSPWave.SampleSize, 0, Bit0IsMsb)
        Next site
        Call AddStoredCaptureData(SplitCalStr(0) & "_" & SecondLoop, DictValue_DSPWave)
               
        If SecondLoop = 0 Then
            SaveDeltaValue_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspLong
            SaveMeasNum_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)) + 1, DspLong
            
            If FirstLoop = 0 Then
                SaveDCK_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)) + 1, DspLong
            Else
                SaveDCK_DSPWave = GetStoredCaptureData("SaveDCK_DSPWaveData")
            End If
            For Each site In TheExec.sites.Active
                SaveDCK_DSPWave.Element(FirstLoop) = DictValue_DSPWave.Element(0)
                SaveMeasNum_DSPWave.Element(SecondLoop) = DictValue_DSPWave.Element(0)
            Next site
            Call AddStoredCaptureData("SaveDCK_DSPWaveData", SaveDCK_DSPWave)
            Call AddStoredCaptureData("SaveMeasNum_DSPWaveData", SaveMeasNum_DSPWave)
            Call AddStoredCaptureData("SaveDeltaValue_DSPWaveData", SaveDeltaValue_DSPWave)
            
        ElseIf SecondLoop <> 64 Then
            SaveMeasNum_DSPWave = GetStoredCaptureData("SaveMeasNum_DSPWaveData")
            SaveDeltaValue_DSPWave = GetStoredCaptureData("SaveDeltaValue_DSPWaveData")
            PreValue_DSPWave = GetStoredCaptureData(SplitCalStr(0) & "_" & SecondLoop - 1)
            
            For Each site In TheExec.sites.Active
                SaveDeltaValue_DSPWave.Element(SecondLoop - 1) = Abs(DictValue_DSPWave.Element(0) - PreValue_DSPWave.Element(0))
                SaveMeasNum_DSPWave.Element(SecondLoop) = DictValue_DSPWave.Element(0)
            Next site
            Call AddStoredCaptureData("SaveMeasNum_DSPWaveData", SaveMeasNum_DSPWave)
            Call AddStoredCaptureData("SaveDeltaValue_DSPWaveData", SaveDeltaValue_DSPWave)
            
        ElseIf SecondLoop = 64 Then
            SaveMeasNum_DSPWave = GetStoredCaptureData("SaveMeasNum_DSPWaveData")
            SaveDeltaValue_DSPWave = GetStoredCaptureData("SaveDeltaValue_DSPWaveData")
            
            For Each site In TheExec.sites.Active
                ValueMax = SaveMeasNum_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
                ValueMin = SaveMeasNum_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
                Idsvalue = (ValueMax - ValueMin) / (CLng(SplitCalStr(1)) + 1)
''''''''''                TheExec.Datalog.WriteComment "MAX_Value:  " & SaveMeasNum_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
''''''''''                TheExec.Datalog.WriteComment "MIN_Value:  " & SaveMeasNum_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
''''''''''                TheExec.Datalog.WriteComment "FinalIds value : " & Idsvalue
            Next site
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-START", CInt(i))
            TheExec.Flow.TestLimit resultVal:=ValueMax, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-END", CInt(i))
            TheExec.Flow.TestLimit resultVal:=ValueMin, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-IDEAL", CInt(i))
            TheExec.Flow.TestLimit resultVal:=Idsvalue, Tname:=TestNameInput, ForceResults:=tlForceFlow

            DNLValue_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspDouble
            For j = 0 To CLng(SplitCalStr(1)) - 1
                For Each site In TheExec.sites.Active
                    DNLValue_DSPWave.Element(j) = (SaveDeltaValue_DSPWave.Element(j) - Idsvalue) / Idsvalue
''''''''''                    TheExec.Datalog.WriteComment "CK Delta value" & j & " = " & CStr(SaveDeltaValue_DSPWave.Element(j))
''''''''''                    TheExec.Datalog.WriteComment "DNL value on point" & j & " = " & CStr(DNLValue_DSPWave.Element(j))
                Next site
                TestNameInput = Report_TName_From_Instance("C", "X", "CKDeltavalue" & Format(j, "00"), CInt(i))
                TheExec.Flow.TestLimit resultVal:=SaveDeltaValue_DSPWave.Element(j), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TestNameInput = Report_TName_From_Instance("C", "X", "CKDNLvalue" & Format(j, "00"), CInt(i))
                TheExec.Flow.TestLimit resultVal:=DNLValue_DSPWave.Element(j), Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next j
            
            For Each site In TheExec.sites.Active
                MaxCKDNL = DNLValue_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
                MinCKDNL = DNLValue_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
''''''''''                TheExec.Datalog.WriteComment "MAX DNL:" & MaxCKDNL
''''''''''                TheExec.Datalog.WriteComment "MIN DNL:" & MinCKDNL
            Next site
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MAX-DNL", CInt(i))
            TheExec.Flow.TestLimit resultVal:=MaxCKDNL, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MIN-DNL", CInt(i))
            TheExec.Flow.TestLimit resultVal:=MinCKDNL, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MAXStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=CStr(SaveDeltaValue_DSPWave.CalcMaximumValue(IndexOfMaximumValue)), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MINStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=CStr(SaveDeltaValue_DSPWave.CalcMinimumValue(IndexOfMinimumValue)), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        End If
            
        If SecondLoop = 64 And FirstLoop = 64 Then
            SaveDCK_DSPWave = GetStoredCaptureData("SaveDCK_DSPWaveData")
            
            For Each site In TheExec.sites.Active
                DCKMAX = SaveDCK_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
                DCKMin = SaveDCK_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
                DCKIdsvalue = (DCKMAX - DCKMin) / (CLng(SplitCalStr(1)) + 1)
            Next site
            TheExec.Datalog.WriteComment "------------------------------------------------- SDLL_DCK summary--------------------------------------------------------------"
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-START", CInt(i))
            TheExec.Flow.TestLimit resultVal:=DCKMAX, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-END", CInt(i))
            TheExec.Flow.TestLimit resultVal:=DCKMin, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-IDEAL", CInt(i))
            TheExec.Flow.TestLimit resultVal:=DCKIdsvalue, Tname:=TestNameInput, ForceResults:=tlForceFlow

            DeltaDCK_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspDouble
            DNLDCKValue_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspDouble

            For j = 1 To CLng(SplitCalStr(1))
                For Each site In TheExec.sites.Active
                    DeltaDCK_DSPWave.Element(j - 1) = SaveDCK_DSPWave.Element(j - 1) - SaveDCK_DSPWave.Element(j)
                    DNLDCKValue_DSPWave.Element(j - 1) = (DeltaDCK_DSPWave.Element(j - 1) - DCKIdsvalue) / DCKIdsvalue
                Next site
                TestNameInput = Report_TName_From_Instance("C", "X", "DCKDeltavalue" & Format(j - 1, "00"), CInt(i))
                TheExec.Flow.TestLimit resultVal:=SaveDCK_DSPWave.Element(j - 1), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TestNameInput = Report_TName_From_Instance("C", "X", "DCKDNLvalue" & Format(j - 1, "00"), CInt(i))
                TheExec.Flow.TestLimit resultVal:=DNLDCKValue_DSPWave.Element(j - 1), Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next j
            
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MAX-DNL", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=DNLDCKValue_DSPWave.CalcMaximumValue(IndexOfMaximumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MIN-DNL", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=DNLDCKValue_DSPWave.CalcMinimumValue(IndexOfMinimumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
                
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MAXStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=DeltaDCK_DSPWave.CalcMaximumValue(IndexOfMaximumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MINStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.Flow.TestLimit resultVal:=DeltaDCK_DSPWave.CalcMinimumValue(IndexOfMinimumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
            Next site
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
            
        End If
    Next i
End Function



Public Function P2PBundle_eye(argc As Integer, argv() As String) As Long

' Format:P2PBundle_eye([SweepLoopName,StartValue,TargetValue,DivideForEye,Site0@Site1&Site2@Site3])
' SweepLoopName : This StringName should be same with Split_Loop_DigSrc_Str(6)
' Site0@Site1&Site2@Site : Exchange data site0 & site1 , site2 & site3

    Dim site As Variant
    Dim EyeWidth As Integer
    Dim EyeWidthTemp As Integer
    Dim i, j, k, z, X As Integer
    Dim strTemp() As String
    Dim EyeDivide As String
    Dim TempString As String
    Dim SiteBundle() As String
    Dim SweepConterStr As String
    Dim DestinationSite As String
    Dim CounterByStart As String
    Dim CounterByTarget As String
    Dim CounterByWidth As String
    Dim DictCounterName As String
    Dim SiteBundleIndex() As String
    Dim Eye_information() As String
    Dim EyeStep() As String
    Dim DataSite() As New SiteVariant
    Dim Mdll_lock As New DSPWave
    Dim Mdll_lockvalue As New DSPWave
    Dim mdll_low() As New SiteDouble
    Dim mdll_high() As New SiteDouble
    ReDim mdll_low(0)
    ReDim mdll_high(0)
    Dim Eye_precent As Double
    
    Dim DSPWaveTemp As New DSPWave
    Dim New_DSPWave As New DSPWave
    Dim PrintEye() As New SiteLong
    Dim UnitCellrecord() As New SiteVariant
    Dim TestNameInput As String
    Dim TestNameInputeye As String
    Dim TnumRecord As Long
    
    

    For i = 0 To argc - 5
        Eye_information = Split(argv(0), "=")
        EyeStep = Split(Eye_information(1), "@")
        
        CounterByStart = EyeStep(0)
        CounterByTarget = EyeStep(1)
        CounterByWidth = EyeStep(2)
        EyeDivide = argv(1)
        argv(2) = Replace(argv(2), "]", "")
        strTemp = Split(argv(3), "+")
        DictCounterName = Replace(Eye_information(0), "[", "")
        SiteBundleIndex = Split(argv(2), "&")
        ReDim DataSite((UBound(strTemp)))
        Public_GetStoredString (DictCounterName)
        SweepConterStr = gl_SpecialString
        New_DSPWave.CreateConstant 0, 2 * (UBound(strTemp) + 1), DspLong
        Mdll_lockvalue = GetStoredCaptureData(argv(4))
        'Mdll_lockvalue.ConvertDataTypeTo (DspLong)
        Mdll_lock.CreateConstant 0, 1, DspLong
        For Each site In TheExec.sites
            Mdll_lock(site) = Mdll_lockvalue(site).ConvertStreamTo(tldspParallel, Mdll_lockvalue.SampleSize, 0, Bit0IsMsb)
            mdll_low(0)(site) = Mdll_lock(site).Element(0) / 4
            mdll_high(0)(site) = Mdll_lock(site).Element(0) / 2
        Next site
        
        
        
        
         ReDim PrintEye((UBound(strTemp) + 1) * CounterByWidth) As New SiteLong
         ReDim UnitCellrecord((UBound(strTemp) + 1) * CounterByWidth) As New SiteVariant
        For j = 0 To UBound(strTemp)
            DSPWaveTemp = GetStoredCaptureData(strTemp(j))
            If CLng(SweepConterStr) <> CLng(CounterByStart) Then
                DataSite(j) = GetStoredMeasurement(strTemp(j) & "_" & "AssemblyStr")
            End If
            For k = 0 To UBound(SiteBundleIndex)
                SiteBundle = Split(SiteBundleIndex(k), "@")
                For X = 0 To UBound(SiteBundle)
                    TempString = ""
                    DestinationSite = SiteBundle(UBound(SiteBundle) - X)
                    
                    For z = 0 To DSPWaveTemp(DestinationSite).SampleSize - 1
                        If z = 0 Then
                            TempString = CStr(DSPWaveTemp(DestinationSite).Element(0))
                        Else
                            TempString = CStr(DSPWaveTemp(DestinationSite).Element(z)) & TempString
                        End If
                    Next z
                    DataSite(j)(SiteBundle(X)) = TempString & DataSite(j)(SiteBundle(X))
                Next X
            Next k
            Call AddStoredMeasurement(strTemp(j) & "_" & "AssemblyStr", DataSite(j))
            If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            
                Dim UnitCellString As String
                Dim SweepStep As Long
                SweepStep = CLng(CounterByTarget / EyeDivide)
                
                
                'ReDim PrintEye((UBound(StrTemp) + 1) * CounterByWidth) As New SiteLong
                
                For Each site In TheExec.sites
                    For z = 1 To CounterByWidth     '6= Unitcell Number
                        UnitCellString = ""
                        For k = 0 To SweepStep
                        
                          If k = 0 Then
                            UnitCellString = Mid(DataSite(j)(site), z + k * 6, 1)
                          Else
                            UnitCellString = UnitCellString & Mid(DataSite(j)(site), z + k * 6, 1)
                    
                          End If
                        Next k
                    
                        EyeWidth = 0
                        EyeWidthTemp = 0
                    
                        For k = 0 To Len(UnitCellString)
                            If Mid(UnitCellString, k + 1, 1) = "1" Then
                                EyeWidthTemp = EyeWidthTemp + 1
                            ElseIf k = Len(UnitCellString) And EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                            Else
                                If EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                                Else
                                    EyeWidthTemp = 0
                                End If
                            End If
                        Next k
                      
'                       Dim PrintEye() As SiteLong
'                       ReDim PrintEye(SweepStep * CounterByWidth)
                      
                      PrintEye((z - 1) + j * CounterByWidth)(site) = EyeWidth
                      UnitCellrecord((z - 1) + j * CounterByWidth)(site) = UnitCellString
                      
'                      If EyeDivide <> "" Then
'                            EyeWidthTemp = CStr(FormatNumber((EyeWidth * EyeDivide), 0))
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidthTemp
'                      Else
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidth
'                      End If
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & CStr(StrTemp(j)) & " : " & UnitCellString
                    Next z
                
                Next site
           End If
        Next j
        
           
         If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            TheExec.Datalog.WriteComment "=========================== Count EYE=========================== "
              For Each site In TheExec.sites
              
                   For k = 0 To UBound(strTemp)
                   'SplitByAt = Split(argv(i), "@")
                   
                   'TnumRecord = TheExec.sites.item(Site).TestNumber
                  
                       For z = 0 To CounterByWidth - 1
                        TestNameInput = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE"
                        TestNameInputeye = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE-percent"
                        
                        TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, X, 0)
                        
                   
                       
                        TheExec.Flow.TestLimit lowVal:=mdll_low(0), hiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                        
                        
                        TestNameInputeye = Report_TName_From_Instance("calc", "x", TestNameInputeye, X, 0)
                        
                        Eye_precent = (PrintEye(z + k * CounterByWidth) * EyeDivide) / Mdll_lock(site).Element(0)
                        
                        TheExec.Flow.TestLimit lowVal:=25, hiVal:=50, resultVal:=Format(Eye_precent * 100, 0), formatStr:="%i", Tname:=TestNameInputeye, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                        
                        
                        
                        
                        'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
                          'theexec.sites.item(Site).TestNumber = theexec.sites.item(Site).TestNumber + 1
                       Next z
                   'TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
                   Next k
           
               Next site
               
               
               For Each site In TheExec.sites
                 For k = 0 To UBound(strTemp)
                    For z = 0 To CounterByWidth - 1
                      If EyeDivide <> "" Then
                            
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth) * EyeDivide
                      Else
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth)
                      End If
                      TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z - 1 & "_" & CStr(strTemp(k)) & " : " & CStr(UnitCellrecord(z + k * CounterByWidth))
                    Next z
                 Next k
              Next site
               
               
               
        End If
    Next i
End Function
        

Public Function P2PBundle_Unflipeye(argc As Integer, argv() As String) As Long

' Format:P2PBundle_eye([SweepLoopName,StartValue,TargetValue,DivideForEye,Site0@Site1&Site2@Site3])
' SweepLoopName : This StringName should be same with Split_Loop_DigSrc_Str(6)
' Site0@Site1&Site2@Site : Exchange data site0 & site1 , site2 & site3

    Dim site As Variant
    Dim EyeWidth As Integer
    Dim EyeWidthTemp As Integer
    Dim i, j, k, z, X As Integer
    Dim strTemp() As String
    Dim EyeDivide As String
    Dim TempString As String
    Dim SiteBundle() As String
    Dim SweepConterStr As String
    Dim DestinationSite As String
    Dim CounterByStart As String
    Dim CounterByTarget As String
    Dim CounterByWidth As String
    Dim DictCounterName As String
    Dim SiteBundleIndex() As String
    Dim Eye_information() As String
    Dim EyeStep() As String
    Dim DataSite() As New SiteVariant
    Dim Mdll_lock As New DSPWave
    Dim Mdll_lockvalue As New DSPWave
    Dim mdll_low() As New SiteDouble
    Dim mdll_high() As New SiteDouble
    ReDim mdll_low(0)
    ReDim mdll_high(0)
    Dim Eye_precent As Double
    
    Dim DSPWaveTemp As New DSPWave
    Dim New_DSPWave As New DSPWave
    Dim PrintEye() As New SiteLong
    Dim UnitCellrecord() As New SiteVariant
    Dim TestNameInput As String
    Dim TestNameInputeye As String
    Dim TnumRecord As Long
    
    'Alg::P2PBundle_unflipeye([Loop_DigSrc=64@804@10,1,0@1],Bundle+++++,D2D_CMN__MDLL_LOCK_CODE_mdll_dcode_lock_NV)

    For i = 0 To argc - 5
        Eye_information = Split(argv(0), "=")
        EyeStep = Split(Eye_information(1), "@")
        
        CounterByStart = EyeStep(0)
        CounterByTarget = EyeStep(1)
        CounterByWidth = EyeStep(2)
        EyeDivide = argv(1)
        argv(2) = Replace(argv(2), "]", "")
        strTemp = Split(argv(3), "+")
        DictCounterName = Replace(Eye_information(0), "[", "")
        SiteBundleIndex = Split(argv(2), "&")
        ReDim DataSite((UBound(strTemp)))
        Public_GetStoredString (DictCounterName)
        SweepConterStr = gl_SpecialString
        New_DSPWave.CreateConstant 0, 2 * (UBound(strTemp) + 1), DspLong
        Mdll_lockvalue = GetStoredCaptureData(argv(4))
        'Mdll_lockvalue.ConvertDataTypeTo (DspLong)
        Mdll_lock.CreateConstant 0, 1, DspLong
        For Each site In TheExec.sites
            Mdll_lock(site) = Mdll_lockvalue(site).ConvertStreamTo(tldspParallel, Mdll_lockvalue.SampleSize, 0, Bit0IsMsb)
            mdll_low(0)(site) = Mdll_lock(site).Element(0) / 4
            mdll_high(0)(site) = Mdll_lock(site).Element(0) / 2
        Next site
        
        
        
        
         ReDim PrintEye((UBound(strTemp) + 1) * CounterByWidth) As New SiteLong
         ReDim UnitCellrecord((UBound(strTemp) + 1) * CounterByWidth) As New SiteVariant
        For j = 0 To UBound(strTemp)
            DSPWaveTemp = GetStoredCaptureData(strTemp(j))
            If CLng(SweepConterStr) <> CLng(CounterByStart) Then
                DataSite(j) = GetStoredMeasurement(strTemp(j) & "_" & "AssemblyStr")
            End If
            For k = 0 To UBound(SiteBundleIndex)
                SiteBundle = Split(SiteBundleIndex(k), "@")
               ''''' For x = 0 To UBound(SiteBundle)
                 For Each site In TheExec.sites
                    TempString = ""
                    '''''DestinationSite = SiteBundle(x)  'SiteBundle(UBound(SiteBundle) - x)  change for no flip site
                    
                    
                    For z = 0 To DSPWaveTemp(site).SampleSize - 1
                        If z = 0 Then
                            TempString = CStr(DSPWaveTemp(site).Element(0))
                        Else
                            TempString = TempString & CStr(DSPWaveTemp(site).Element(z)) 'CStr(DSPWaveTemp(DestinationSite).Element(z)) & TempString  change for no unit cell flip
                        End If
                    Next z
                    DataSite(j)(site) = TempString & DataSite(j)(site)
                Next site
            Next k
            Call AddStoredMeasurement(strTemp(j) & "_" & "AssemblyStr", DataSite(j))
            If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            
                Dim UnitCellString As String
                Dim SweepStep As Long
                SweepStep = CLng(CounterByTarget / EyeDivide)
                
                
                'ReDim PrintEye((UBound(StrTemp) + 1) * CounterByWidth) As New SiteLong
                
                For Each site In TheExec.sites
                    For z = 1 To CounterByWidth     '6= Unitcell Number
                        UnitCellString = ""
                        For k = 0 To SweepStep
                        
                          If k = 0 Then
                            UnitCellString = Mid(DataSite(j)(site), z + k * 6, 1)
                          Else
                            UnitCellString = UnitCellString & Mid(DataSite(j)(site), z + k * 6, 1)
                    
                          End If
                        Next k
                    
                        EyeWidth = 0
                        EyeWidthTemp = 0
                    
                        For k = 0 To Len(UnitCellString)
                            If Mid(UnitCellString, k + 1, 1) = "1" Then
                                EyeWidthTemp = EyeWidthTemp + 1
                            ElseIf k = Len(UnitCellString) And EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                            Else
                                If EyeWidthTemp > EyeWidth Then
                                    EyeWidth = EyeWidthTemp
                                    EyeWidthTemp = 0
                                Else
                                    EyeWidthTemp = 0
                                End If
                            End If
                        Next k
                      
'                       Dim PrintEye() As SiteLong
'                       ReDim PrintEye(SweepStep * CounterByWidth)
                      
                      PrintEye((z - 1) + j * CounterByWidth)(site) = EyeWidth
                      UnitCellrecord((z - 1) + j * CounterByWidth)(site) = UnitCellString
                      
'                      If EyeDivide <> "" Then
'                            EyeWidthTemp = CStr(FormatNumber((EyeWidth * EyeDivide), 0))
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidthTemp
'                      Else
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & "EyeWidth : " & EyeWidth
'                      End If
'                            TheExec.Datalog.WriteComment "Site" & CStr(Site) & " , " & "UnitCell" & z - 1 & "_" & CStr(StrTemp(j)) & " : " & UnitCellString
                    Next z
                
                Next site
           End If
        Next j
        
           
         If CLng(SweepConterStr) = CLng(CounterByTarget) Then
            TheExec.Datalog.WriteComment "=========================== Count EYE=========================== "
              For Each site In TheExec.sites
              
                   For k = 0 To UBound(strTemp)
                   'SplitByAt = Split(argv(i), "@")
                   
                   'TnumRecord = TheExec.sites.item(Site).TestNumber
                  
                       For z = 0 To CounterByWidth - 1
                        TestNameInput = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE"
                        TestNameInputeye = "UnitCell" & z & "_" & CStr(strTemp(k)) & "EYE-percent"
                        
                        TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, X, 0)
                        
                   
                       
                        'TheExec.Flow.TestLimit LowVal:=mdll_low(0), HiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceNone, ScaleType:=scaleNoScaling
                        TheExec.Flow.TestLimit lowVal:=mdll_low(0), hiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                        
                        TestNameInputeye = Report_TName_From_Instance("calc", "x", TestNameInputeye, X, 0)
                        
                        Eye_precent = (PrintEye(z + k * CounterByWidth) * EyeDivide) / Mdll_lock(site).Element(0)
                        
                        TheExec.Flow.TestLimit lowVal:=25, hiVal:=50, resultVal:=Format(Eye_precent * 100, 0), formatStr:="%i", Tname:=TestNameInputeye, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                        
                        
                        
                        
                        'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
                          'theexec.sites.item(Site).TestNumber = theexec.sites.item(Site).TestNumber + 1
                       Next z
                   'TheExec.sites.item(Site).TestNumber = TheExec.sites.item(Site).TestNumber + 1
                   Next k
           
               Next site
               
               
               For Each site In TheExec.sites
                 For k = 0 To UBound(strTemp)
                    For z = 0 To CounterByWidth - 1
                      If EyeDivide <> "" Then
                            
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth) * EyeDivide
                      Else
                            TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z & "_" & "EyeWidth : " & PrintEye(z + k * CounterByWidth)
                      End If
                      TheExec.Datalog.WriteComment "Site" & CStr(site) & " , " & "UnitCell" & z & "_" & CStr(strTemp(k)) & " : " & CStr(UnitCellrecord(z + k * CounterByWidth))
                    Next z
                 Next k
              Next site
               
               
               
        End If
    Next i
End Function
