Attribute VB_Name = "VBT_LIB_HardIP"
' dummy comment
Option Explicit

Public gl_TName_Pat As String           'Roger add, For TName
Public gl_SweepNum As String

'Public DACInitialFlag As Boolean
'==========================================================Roger Add,for power sweep
Type Power_Sweep
    Loop_count As Long
    Loop_Index_Name As String
    PinName As String
    from As String
    stop As String
    step As String
    Count As Long
    Key As String
End Type
'==========================================================
'' 20151229 - hard ip dssc code, add rd-rd9
Public Src_DSPWave As New DSPWave
Public Src1_DSPWave As New DSPWave
Public Src2_DSPWave As New DSPWave
Public Src3_DSPWave As New DSPWave
Public Src4_DSPWave As New DSPWave
Public Src5_DSPWave As New DSPWave
Public Src6_DSPWave As New DSPWave
Public Src7_DSPWave As New DSPWave
Public Src8_DSPWave As New DSPWave
Public Src9_DSPWave As New DSPWave

'' 20160121 - hard ip dssc code, add rd10-rd12 for Starling
Public Src10_DSPWave As New DSPWave
Public Src11_DSPWave As New DSPWave
Public Src12_DSPWave As New DSPWave

Public CP_Card_RAK As New PinListData
Public FT_Card_RAK As New PinListData
Public WLFT1_Card_RAK As New PinListData
Public CurrentJob_Card_RAK As New PinListData
Public FourceV As Double

Public ADDRIO_Norm_Y_T1() As Double
Public ADDRIO_Norm_Y_T2() As Double
Public ADDRIO_Norm_Y_T1_T2_ReadOnce_Flag As Boolean

Enum CalculateMethodSetup_PPMU
    PPMU_DEFAULT = 0
    PPMU_STORE_I = 1
    VIR_DIFF_PN_ABS = 4
    VIR_DIFF_PN = 5
    VIR_VOD_VOCM_XI0_Off = 8
    VIR_VOD_VOCM_PN = 9
    VIR_DDIO = 10
End Enum

Enum CalculateMethodSetup_DSPWave
     DigCap_DEFAULT_SETUP = 0
     DigCap_MultiPinsOperation = 1
End Enum

'Enum InstrumentSpecialSetup_PPMU
'     DEFAULT_SETUP = 0
'     PPMU_SerialMeasurement = 1
'     DigitalConnectPPMU2 = 2 ' 20160204
'End Enum

Enum CalculateMethodSetup
    DEFAULT_SINGLE = 0
    DIFF_1ST = 1
    DIFF_2ND = 2
    DIFF_PT12 = 3
    DIFF_PN = 4
    DIFF_DCO = 5
    DIFF_DAC = 6
    Force_VDD12_RX = 7
    RATIO_FREQ = 8
    PPMU_TestLimit_TTR = 15
    Average_voltage = 10
    PPMU_STORE_I = 11
    VIR_DIFF_PN_ABS = 12
    VIR_DIFF_PN = 13
    VIR_VOD_VOCM_XI0_Off = 14
    VIR_VOD_VOCM_PN = 9
    VIR_DDIO = 16
''    DigCap_MultiPinsOperation = 4
    DigCap_DEFAULT_SETUP = 17
    DigCap_MultiPinsOperation = 18
End Enum


Type DSSC_CodeSearchCond
    MeasureValue As New SiteDouble
    SearchCode As New SiteLong
    TargetCodeFind As New SiteBoolean
    TransitionPoint As New SiteVariant
    patternPass As New SiteBoolean
End Type

Enum InstrumentSpecialSetup
     DEFAULT_SETUP = 0
     DigitalConnectPPMU = 1
     PPMU_SerialMeasurement = 3 'For turks HIP_USB
     DigitalConnectPPMU2 = 2 'For turks HIP_USB
     PPMU_AccurateMeasurement = 4
     PPMU_2mA_Force_I_Range = 5
     PPMU_200uA_Force_I_Range = 6
     PPMU_20uA_Force_I_Range = 7
     EUSB_T10T11_Split_Force_I_Range = 8
End Enum

Public Stored_MeasI_PPMU As New PinListData

'' 20151117 - Event source combine HiZ/VT mode
Enum EventSourceWithTerminationMode 'not use, for reference.
     BOTH_VT = 1
     VOH_VT = 2
     VOL_VT = 3
     BOTH_HIZ = 4
     VOH_HIZ = 5
     VOL_HIZ = 6
End Enum

Enum Enum_RAK
     Default = 0
     R_TraceOnly = 1
     R_PathWithContact = 2
End Enum

'' 20151224 - Print measured frequency during shmoo if need
Public G_MeasFreqForCZ As New PinListData

Public SweepWhileFirstTimeFlag As New SiteBoolean ''VBT_LIB_HardIP
Public SourceCode_First6Bit As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_Last6Bit_MSB As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_Last6Bit As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_First6Bit_MSB As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_First6Bit_Code0To4 As New SiteVariant ''VBT_LIB_HardIP
Public SourceCode_Last6Bit_Code0To4 As New SiteVariant ''VBT_LIB_HardIP
Public Dec_First6Bit_Code0To4 As New SiteDouble ''VBT_LIB_HardIP
Public Dec_Last6Bit_Code0To4 As New SiteDouble ''VBT_LIB_HardIP
Public Final_Dec_First6Bit_Code0To4 As New SiteDouble ''VBT_LIB_HardIP
Public Final_Dec_Last6Bit_Code0To4 As New SiteDouble ''VBT_LIB_HardIP
Public PostiveIndex As New SiteLong ''VBT_LIB_HardIP
Public Source12Bits As New SiteVariant ''VBT_LIB_HardIP
Public Final_point_flag As New SiteBoolean ''VBT_LIB_HardIP
Public Final_point As New SiteLong ''VBT_LIB_HardIP
Public NegativeIndex As New SiteLong ''VBT_LIB_HardIP
Public Imped_LowLimit As Double ''VBT_LIB_HardIP, ImpedanceMeasurement_2Point
Public Imped_HighLimit As Double ''VBT_LIB_HardIP, ImpedanceMeasurement_2Point

Public G_pld_DigCapInfo() As New PinListData ''VBT_LIB_HardIP

Enum DDR_Eye_setup
    DDR_EYE_False = 0
    DDR_EYE_1ST = 1
    DDR_EYE_2ND = 2
End Enum

''20160729 - Use global value to denfine default setting
Public Const pc_Def_VFI_FreqInterval = 0.001
Public Const pc_Def_VFI_FreqThresholdPercentage = 0.5
Public Const pc_Def_VFI_MeasCurrRange = 0.02

Public Const pc_Def_VIR_MeasCurrRange = 0.05

Public Const pc_Def_VFI_UVI80_VoltCalmp = 6
Public Const pc_Def_VFI_UVI80_InitialVal_FI = 0
Public Const pc_Def_VFI_UVI80_ReadPoint = 2
Public Const pc_Def_VFI_UVI80_VoltageRange = 7
Public Const pc_Def_VFI_UVI80__InitialVal_FI = 0

Public Const pc_Def_Default_Range_By_Instrument = 0
Public Const pc_Def_UVI80_Init_MeasCurrRange = 0.2

Public Const pc_Def_PPMU_InitialValue_FI = 0
Public Const pc_Def_PPMU_InitialValue_FI_Range = 0.05
Public Const pc_Def_PPMU_Max_InitialValue_FI_Range = 0.05
Public Const pc_Def_PPMU_InitialValue_FV = 0
Public Const pc_Def_PPMU_InitialValue_FV_Range = 0
Public Const pc_Def_PPMU_FI_Range_200uA = 0.0002
Public Const pc_Def_PPMU_ReadPoint_FreqDC = 20
Public Const pc_Def_PPMU_ReadPoint = 10
Public Const pc_Def_PPMU_ClampVHi = 2
Public Const pc_Def_PPMU_Digital_MaxCurrRange = 0.05

Public Const pc_Def_DSSC_Amplitude = 1

Public Const pc_Def_HexVS_ReadPoint = 10000

Public Const pc_Def_UVS256_ReadPoint = 1
Public Const pc_Def_UVS256_CurrentRangeRatio = 0.09

Public Const pc_Def_Power1p2 = 1.2

Public Const pc_Def_DCTIME_InitialCurrent = 0
Public Const pc_Def_DCTIME_SampleSize = 10
  
Public Const pc_Def_DiffMeter_HWAverageSize = 64
Public Const pc_Def_DiffMeter_VoltRange = 1.4
Public Const pc_Def_DiffMeter_ReadPoint = 100000
Public Const pc_Def_HardIP_PatGenTimeout = 10

Public Const pc_Str_InitOff_Pins = "All_Digital"

Public Const pc_Def_VFI_MI_WaitTime = 1 * ms
Public Const pc_Def_VFI_MI_WaitTime_PPMU = 100 * us
Public Const pc_Def_VFI_MI_WaitTime_UVI80 = 1 * ms

Public TPModeAsCharz_GLB As Boolean
'Public FlowShmooString_GLB As String               '20170717 already defined in Lib_Digital_Shmoo.bas

Type MeasTrimImpedInfo
    Pat As String
    MeasPinsAry() As String
    IsDifferential As Boolean
End Type

''20170405-Global string to record all functional result if flow for loop specified (gray code)
Public gs_RecordGrayCodeTestResult As String

'' 20170523
Public Const pc_Def_VFI_ForceI_Val = 0

''20170518-Global string to Customize DigCap
Public MTR_CusDigCap As String
Public MTR_VIN As String
''20170524-Global string to T4P3
Public MTR_T4P3_MeasVolt_P(1023) As New PinListData
Public MTR_T4P3_MeasVolt_N(1023) As New PinListData
Public MTR_T4P3_DigSrc(255) As String
Public StoreIndex_MTR_P As Integer
Public StoreIndex_MTR_N As Integer
''20170605-Globastring forMetrology T2P6 CMRR
Public CMRR_VIN As Double
''20170606-Globastring forMetrology T2P6 CMRR average
Public CMRR_Average(71) As New DSPWave
Public StoreIndex_CMRR_Average As Integer
Public PSRR_Average(35) As New DSPWave
Public StoreIndex_PSRR_Average As Integer
'Store for Avg Dic for metrology 20170711
Public MetroAvgDic As New Dictionary

Public gl_CZ_FlowTestNameIndex As Long
Public gl_DSSC_OUT_STR As String
Public gl_DSSC_CALC_STR As String
Public gl_Sweep_Glb_TName As String '' 20190529 - Add for sweep force V

Public gl_Flag_HardIP_Characterization_1stRun As Boolean
Public gl_Flag_HardIP_Trim_Set_PrePoint As Boolean
Public gl_Flag_HardIP_Trim_Set_PostPoint As Boolean
Public gl_Flag_HardIP_Disable_Functional_Result As Boolean
'Sweep name for Sweepsrc
Public Sweepnameforsweep() As String
Public temp_CUS_String As String
Public srcnameindex As Integer
Type EyeDiagram
    Value() As New SiteVariant
End Type
Public Function Meas_Vdiff_func(Optional patset As Pattern, Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional PpmuMeasureP_Pin As String, Optional PpmuMeasureN_Pin As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
    Optional ForceV1p As String, Optional ForceV2p As String, Optional ForceV1n As String, Optional ForceV2n As String, _
    Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String = "", _
    Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, _
    Optional Meas_StoreName As String, Optional Calc_Eqn As String, Optional TestLimitPerPin_VFI As String = "FFF", _
    Optional Validating_ As Boolean) As Long
    
    Dim PatCount As Long
    Dim PattArray() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen

    On Error GoTo errHandler
    
    Dim pat_count As Long
    Dim i As Long, k As Long, j As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim TestNum As Long
    Dim TestNumber As Long
    Dim patt_ary() As String
    Dim InDSPwave As New DSPWave
    Dim OutDspWave As New DSPWave
    Dim show_Dec As String
    Dim show_out As String
    Dim site As Variant

    Dim patt As Variant
    Dim Pat As String
    Dim HighLimitVal() As Double
    Dim LowLimitVal() As Double
    Dim Idiff As New SiteDouble, Vdiff As New SiteDouble, Vocm As New SiteDouble
    Dim TxVDD As New SiteDouble, TxVss As New SiteDouble
    Dim MeasIp1 As New PinListData
    Dim MeasIp2 As New PinListData
    Dim MeasIn1 As New PinListData
    Dim MeasIn2 As New PinListData
    Dim MeasVdiff As New PinListData
    Dim MeasVocm As New PinListData
    Dim ForceV1pAry() As String
    Dim ForceV2pAry() As String
    Dim ForceV1nAry() As String
    Dim ForceV2nAry() As String
                                                                                                                                                                                                                                                               
    Dim Zp As New SiteDouble
    Dim Zn As New SiteDouble
    
    Dim TestNameInput As String
    Dim OutputTname_format() As String
 
    Dim TestPinArrayP() As String, TestPinArrayN() As String
    Dim Pin As Variant

    ''20160923 - Analyze Interpose_PreMeas to force setting with different sequence.
    Dim Interpose_PreMeas_Ary() As String
    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
    
    ''Defined for TTR
    Dim DiffPins As New PinList
    
    Dim TName_Ary() As String
    Dim Tname As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    '----180524----------------------------------------------------------------
    Call GetFlowTName
    
    DiffPins = PpmuMeasureP_Pin & "," & PpmuMeasureN_Pin
                                                                                                                                                                                                                                                           
    TestSequenceArray = Split(TestSequence, ",")

    TestPinArrayP = Split(PpmuMeasureP_Pin, "+")
    TestPinArrayN = Split(PpmuMeasureN_Pin, "+")
    
    ForceV1pAry = Split(ForceV1p, "+")
    ForceV2pAry = Split(ForceV2p, "+")
    ForceV1nAry = Split(ForceV1n, "+")
    ForceV2nAry = Split(ForceV2n, "+")
    
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)  ''20141219


    ''20161130-Get test name from flow table
    Dim FlowTestNme() As String
    If TPModeAsCharz_GLB Then
        gl_CZ_FlowTestName_Counter = 0
        Call GetFlowTestName(FlowTestNme)
    End If
                                                                                                                                                                                                                                                               
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    '' 20160923 - Add Interpose_PrePat entry point
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    'TName_Ary = Split(gl_Tname_Meas, "+")
    
'    If (UBound(TestSequenceArray) > UBound(TName_Ary)) Then
'        ReDim Preserve TName_Ary(UBound(TestSequenceArray)) As String
'
'    End If
    
    TheHdw.Patterns(patset).Load
    
    Call PATT_GetPatListFromPatternSet(patset.Value, patt_ary, pat_count)
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    
    ''20161107-Return sweep test name
    Dim Rtn_SweepTestName As String
    Rtn_SweepTestName = ""
    
    ''========================================================================================
    ''20170203 - Analyze Meas_StoreName and store the measurement for futher use.
    Dim Rtn_MeasVolt As New PinListData
    Dim MeasStoreName_Ary() As String
    MeasStoreName_Ary = Split(Meas_StoreName, "+")
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
    ''========================================================================================
                                                                                                                                                                                                                                                               
    For Each patt In patt_ary
        
        Pat = CStr(patt)
        
        TheHdw.Patterns(Pat).Load
        
        'If theexec.DataManager.InstanceName = "LPDPTX_1D_PP_CEBA0_S_FULP_AN_TX00_DCT_JTG_VMX_ALLFV_SI_DPTX_1D_NV" Then Stop
        
        Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, _
                                               DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPwave, Rtn_SweepTestName)

        If TPModeAsCharz_GLB = True Then
            If Rtn_SweepTestName <> "" Then
''                Rtn_SweepTestName = "_" & Rtn_SweepTestName
                For i = 0 To UBound(FlowTestNme)
                    FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepcode", Rtn_SweepTestName)
                Next i
            End If
        End If
        
        Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
        
        Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
        
        '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Patterns(Pat).start
        Else
            Call TheHdw.Patterns(Pat).Test(pfAlways, 0)
        End If
        
        TestSeqNum = 0
        Dim Force_idx As Integer
        Force_idx = 0
        For Each Ts In TestSequenceArray
'            If (UBound(TName_Ary) < TestSeqNum) Then
'                TName = ""
'            Else
'                TName = TName_Ary(TestSeqNum)
'            End If


            ''20150907 - Only need CPUA_Flag_In_Pat to do control
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            TestOptLen = Len(Ts)
            
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                End If
            End If
                                                                                                                                                                                                                                                               
            For k = 1 To TestOptLen
                TestOption = Mid(Ts, k, 1)
                For Each site In TheExec.sites.Active: TestNum = TheExec.sites.Item(site).TestNumber: Exit For: Next site
                TheHdw.Digital.Pins(Replace(DiffPins, "+", ",")).Disconnect
                
                With TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum))
                    .ForceV ForceV1pAry(Force_idx)
''                   .ClampVHi = 2
                    .ClampVHi = pc_Def_PPMU_ClampVHi
                    .Connect
                    .Gate = tlOn
                End With
               
                With TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum))
                    .ForceV ForceV1nAry(Force_idx)
''                   .ClampVHi = 2
                    .ClampVHi = pc_Def_PPMU_ClampVHi
                    .Connect
                    .Gate = tlOn
                End With
                 
                TheHdw.Wait 0.001
                DebugPrintFunc_PPMU CStr(TestPinArrayP(TestSeqNum))
''                MeasIp1 = TheHdw.PPMU.Pins(TestPinArrayP(k - 1)).Read(tlPPMUReadMeasurements, 10)
                MeasIp1 = TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                DebugPrintFunc_PPMU CStr(TestPinArrayN(TestSeqNum))
''                MeasIn1 = TheHdw.PPMU.Pins(TestPinArrayN(k - 1)).Read(tlPPMUReadMeasurements, 10)
                MeasIn1 = TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                 
                TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum)).ForceV ForceV2pAry(Force_idx)
                TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum)).ForceV ForceV2nAry(Force_idx)
                TheHdw.Wait 0.001
                 
                DebugPrintFunc_PPMU CStr(TestPinArrayP(TestSeqNum))
''                MeasIp2 = TheHdw.PPMU.Pins(TestPinArrayP(k - 1)).Read(tlPPMUReadMeasurements, 10)
                MeasIp2 = TheHdw.PPMU.Pins(TestPinArrayP(TestSeqNum)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                DebugPrintFunc_PPMU CStr(TestPinArrayN(TestSeqNum))
''                MeasIn2 = TheHdw.PPMU.Pins(TestPinArrayN(k - 1)).Read(tlPPMUReadMeasurements, 10)
                MeasIn2 = TheHdw.PPMU.Pins(TestPinArrayN(TestSeqNum)).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)

                MeasVdiff = MeasIp1
                MeasVocm = MeasIp1
                
                'Dim RakCh_p() As Double
                'Dim RakCh_n() As Double
                Dim RAKVal_p As Double
                Dim RAKVal_n As Double
                 
                For Each site In TheExec.sites
                    For i = 0 To MeasIp1.Pins.Count - 1
                        'RakCh_p = TheHdw.PPMU.ReadRakValuesByPinnames(MeasIp1.Pins(i), site)
                        'RakCh_n = TheHdw.PPMU.ReadRakValuesByPinnames(MeasIn1.Pins(i), site)
                        
                        If (MeasIp1.Pins(i).Value(site) - MeasIp2.Pins(i).Value(site)) = 0 Then MeasIp1.Pins(i).Value(site) = MeasIp1.Pins(i).Value(site) + 0.000000001
                        If (MeasIn1.Pins(i).Value(site) - MeasIn2.Pins(i).Value(site)) = 0 Then MeasIn1.Pins(i).Value(site) = MeasIn1.Pins(i).Value(site) + 0.000000001
                       
                        Zp(site) = (CDbl(ForceV1pAry(Force_idx)) - CDbl(ForceV2pAry(Force_idx))) / (MeasIp1.Pins(i).Value(site) - MeasIp2.Pins(i).Value(site))
                        Zn(site) = (CDbl(ForceV1nAry(Force_idx)) - CDbl(ForceV2nAry(Force_idx))) / (MeasIn1.Pins(i).Value(site) - MeasIn2.Pins(i).Value(site))
                        TxVDD(site) = CDbl(ForceV2pAry(Force_idx)) + Zp(site) * MeasIp2.Pins(i).Value(site) * (-1)
                        TxVss(site) = ForceV2nAry(Force_idx) - Zn(site) * MeasIn2.Pins(i).Value(site)
                        
                        RAKVal_p = (CurrentJob_Card_RAK.Pins(MeasIp1.Pins(i)).Value(site))
                        RAKVal_n = (CurrentJob_Card_RAK.Pins(MeasIn1.Pins(i)).Value(site))
                        
                        Idiff(site) = (TxVDD(site) - TxVss(site)) / (Zp(site) - (RAKVal_p) + 100 + Zn(site) - (RAKVal_n))
                        Vdiff(site) = Abs((TxVDD(site) - (Zp(site) - RAKVal_p) * Idiff(site)) - (TxVss(site) + (Zn(site) - RAKVal_n) * Idiff(site)))
                        Vocm(site) = 0.5 * ((TxVDD(site) - (Zp(site) - RAKVal_p) * Idiff(site)) + (TxVss(site) + (Zn(site) - RAKVal_n) * Idiff(site)))
                        
                        MeasVdiff.Pins(i).Value = Vdiff(site)
                        MeasVocm.Pins(i).Value = Vocm(site)
 
                    Next i
                Next site
                   
                For Each site In TheExec.sites.Active: TestNum = TheExec.sites.Item(site).TestNumber: Exit For: Next site
                
                ''20160906 - Check store measurement or not
                If Meas_StoreName <> "" Then
                    If MeasStoreName_Ary(TestSeqNum) <> "" Then
                        Store_Rtn_Meas(StoreIndex) = MeasVdiff
                        Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                        StoreIndex = StoreIndex + 1
                    End If
                End If
                

                
                '' Added TestLimitPerPin 20170814
                Dim p As Long

                For p = 0 To MeasVdiff.Pins.Count - 1
                    TestNameInput = Report_TName_From_Instance("Vdiff", MeasVdiff.Pins(p), , TestSeqNum, p)
                    TheExec.Flow.TestLimit MeasVdiff.Pins(p), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Next p

                For p = 0 To MeasVocm.Pins.Count - 1
                    TestNameInput = Report_TName_From_Instance("Vocm", MeasVocm.Pins(p), , TestSeqNum, p)
                    If UCase(CUS_Str_MainProgram) = "VOCM" Then: TheExec.Flow.TestLimit MeasVocm.Pins(p), , , scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Next p
                If TheExec.sites.Active.Count = 0 Then Exit Function

            Next k
            ''20161206-Restore force condiction after measurement

            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                End If
            End If
            
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)                'Jump out CPUA loop
            
            Force_idx = Force_idx + 1
        
        Next Ts
                                                                                                                                                                                                                                                               
        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & pat_count & "): " & Pat & ""
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        pat_count = pat_count + 1
    Next patt
    
    '' 20160923 - Add Interpose_PostTest entry point
    Call SetForceCondition(Interpose_PostTest)
    
    ''Comment by Martin for TTR
    TheHdw.PPMU.Pins(Replace(DiffPins, "+", ",")).Disconnect
    TheHdw.Digital.Pins(Replace(DiffPins, "+", ",")).Connect
        
    '' 20160211 - Process DigCapData by using DSP
    If DigCap_Sample_Size <> 0 Then
        Dim DigCapPinAry() As String, NumberPins As Long
        Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
        
        If NumberPins > 1 Then
            Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
            Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins, , DigCap_Pin.Value)
        ElseIf NumberPins = 1 Then
            Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave)
            Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, , , DigCap_Pin.Value)
        End If
    End If
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Connect
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
                                                                                                                                                                                                                                                               
    '' 20160907 - Process calculate equation by dictionary.
    If Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Calc_Eqn)
    End If
                                                                                                                                                                                                                                                               
    '' 20160713 - Call write functional result if cpu flag in pattern
    If (CPUA_Flag_In_Pat) Then
        Call HardIP_WriteFuncResult(, , Inst_Name_Str)
    End If
                                                                                                                                                                                                                                                               
    DebugPrintFunc patset.Value  ' print all debug information
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If

    Exit Function
                                                                                                                                                                                                                                                               
errHandler:
    TheExec.Datalog.WriteComment "error in Meas_Vdiff_func"
    If AbortTest Then Exit Function Else Resume Next
                                                                                                                                                                                                                                                               
End Function
                                                        
Public Function Meas_VIR_IO_Universal_func(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
Optional Measure_Pin_PPMU As String, Optional ForceV As String, Optional ForceI As String, Optional MeasureI_Range As String = "0.05", _
Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
Optional InstSpecialSetting As InstrumentSpecialSetup = 0, Optional SpecialCalcValSetting As CalculateMethodSetup = 0, Optional RAK_Flag As Enum_RAK = 0, _
Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String, _
Optional Flag_SingleLimit As Boolean = False, Optional TestLimitPerPin_VIR As String = "FFF", _
Optional ForceFunctional_Flag As Boolean = False, _
Optional Meas_StoreName As String, Optional Calc_Eqn As String, _
Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, Optional WaitTime_VIRZ As String, Optional Validating_ As Boolean) As Long

''Optional b_ProcessDigCapByDSP As Boolean = False, _
''==================================================================================
'' 20150621 - Check with CCWu: FRCPortName As String, Optional DisableFRC As Boolean = False not use in this function
'' 20150717 - Impedance measurement by using 2 point measure method, Define "Z" for TestSequence - On going
''                - EX: Pin1, Pin2 + Pin3, Pin4     V1, V2 + V3, V4
''                - V1 and V2 use for Pin1 of impedence measurement
''                - V1 and V2 use for Pin2 of impedence measurement
'' 20150717 - Get I from previous item and apply the current value to next item, use enum for the feature
''                - EX: TestSequence: "V,V,V"
''                  If second V want to apply calcuated I value that Force I value argument should be "0,keyword,0"
'' 20150727 - MeasureI_Range is use for test sequence "I", "R" and "Z"
''==================================================================================
    Dim site As Variant
    Dim PattArray() As String, PatCount As Long
    
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen
    
    Dim i As Long, j As Long, k As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, ForceISequenceArray() As String, ForceVSequenceArray() As String
    Dim TestOption As Variant, Ts As Variant, TestSeqNum As Integer
    Dim TestPinArrayIV() As String, TestIrange() As String
    Dim TestSeqNumIdx As Long
    Dim InDSPwave As New DSPWave, OutDspWave As New DSPWave
    Dim ShowDec As String, ShowOut As String
    Dim Pat As String, patt As Variant
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Dim Rtn_MeasVolt As New PinListData, Rtn_MeasVolt_CUS_R As New PinListData, Rtn_MeasCurr As New PinListData
    Dim FlowForLoopName() As String   ' Sequences : Code , Voltage , Loop Index
    Dim MeasStoreName_Ary() As String
    Dim Interpose_PreMeas_Ary() As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName  '20170728 Added for HardIP_WriteFuncResult Output
    Dim WaitTime_VIRZ_Ary() As String
    
''    Dim RTN_InterposeString As String
    Dim OutputTname() As String
    On Error GoTo errHandler
    
    
    '----180524----------------------------------------------------------------
    Call GetFlowTName
    
    If ForceI Like "*@*" Then
    ForceI = Replace(ForceI, "@", "")
    End If
    '''''========================================================================================
    If WaitTime_VIRZ <> "" Then   ' update wait parameter for V,I,R,Z
        WaitTime_VIRZ_Ary = Split(WaitTime_VIRZ, ",")
        If UBound(WaitTime_VIRZ_Ary) = 0 Then
            ReDim Preserve WaitTime_VIRZ_Ary(3) As String
            WaitTime_VIRZ_Ary(1) = WaitTime_VIRZ_Ary(0)
            WaitTime_VIRZ_Ary(2) = WaitTime_VIRZ_Ary(0)
            WaitTime_VIRZ_Ary(3) = WaitTime_VIRZ_Ary(0)
        ElseIf UBound(WaitTime_VIRZ_Ary) < 3 Then
            ReDim Preserve WaitTime_VIRZ_Ary(3) As String
        End If
    Else
        ReDim WaitTime_VIRZ_Ary(3) As String
    End If
    '''''========================================================================================
    If Measure_Pin_PPMU Like "*@*" Then Measure_Pin_PPMU = Replace(Measure_Pin_PPMU, "@", "")
    Shmoo_Pattern = patset.Value

    Call tl_PinListDataSort(True)
    
    If (InStr(MeasureI_Range, ":") <> 0) Then MeasureI_Range = Select_MeasIRange(MeasureI_Range, CurrentJobName_U)  ' support different Meter_Range in different Job, add by Roger 20170628
    
    Call VIR_AnalyzedInputStringByAt(Measure_Pin_PPMU, ForceV, ForceI, MeasureI_Range)
   
    If ForceI = "" Then ForceI = 0
    If ForceV = "" Then ForceV = 0
    If MeasureI_Range = "" Then MeasureI_Range = pc_Def_VIR_MeasCurrRange
    
    Call VIR_CheckForceVal(ForceI, ForceV)

    Call VIR_ProcessInputString(TestSequence, ForceI, ForceV, Measure_Pin_PPMU, MeasureI_Range, Meas_StoreName, Interpose_PreMeas, _
                                              TestSequenceArray(), ForceISequenceArray(), ForceVSequenceArray(), TestPinArrayIV(), _
                                              TestIrange(), MeasStoreName_Ary(), Interpose_PreMeas_Ary())
 
'    Call HIP_Evaluate_ForceVal_New(ForceVSequenceArray())
'    Call HIP_Evaluate_ForceVal_New(ForceISequenceArray())
    ''========================================================================================
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
    ''========================================================================================
     ''20170807 - CZ test name index
    gl_CZ_FlowTestNameIndex = 0
    ''========================================================================================
    
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    
    ''20161130-Get test name from flow table
    Dim FlowTestNme() As String
    If TPModeAsCharz_GLB Then
        gl_CZ_FlowTestName_Counter = 0
        Call GetFlowTestName(FlowTestNme)
    End If
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    '' 20160923 - Add Interpose_PrePat entry point
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    

    ''20161205 - Force_Flow_Shmoo_Condition
    If TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then Force_Flow_Shmoo_Condition
    'Do Flow Shmoo
    
    If patset.Value <> "" Then
         gl_TName_Pat = patset.Value
        Shmoo_Pattern = patset.Value '' 20170808 add for shmoo pattern name print
        TheHdw.Patterns(patset).Load
        Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
    Else
        ReDim PattArray(0)
        PattArray(0) = ""
    End If
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True

    Dim Rtn_SweepTestName As String
    Rtn_SweepTestName = ""
    
    For Each patt In PattArray
        If patt <> "" Then

        Pat = CStr(patt)
        
        Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, _
                                               DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPwave, Rtn_SweepTestName)

        If TPModeAsCharz_GLB = True Then
            If Rtn_SweepTestName <> "" Then
''                Rtn_SweepTestName = "_" & Rtn_SweepTestName
                For i = 0 To UBound(FlowTestNme)
                    FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepcode", Rtn_SweepTestName)
                Next i
            Else
                Call SimulateFlowForSweep(FlowShmooString_GLB)
                If FlowShmooString_GLB <> "" Then
                    For i = 0 To UBound(FlowTestNme)
                        FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepvoltage", FlowShmooString_GLB)
                    Next i
                End If
            End If
        End If

        Set OutDspWave = Nothing
        Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
        
        Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
        
        ''20160306-Sweep Volt
        If UCase(DigSrc_FlowForLoopIntegerName) = "SWEEP_V" Then
            Call Cust_Sweep_V
        End If
        
        '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Patterns(Pat).start
        Else
            Call TheHdw.Patterns(Pat).Test(pfAlways, 0)
        End If
        End If
        
        TestSeqNum = 0
        
        Call ProcessTestNameInputString(OutputTname, UBound(TestSequenceArray))
        For Each Ts In TestSequenceArray
            
            ''20150907 - Only need CPUA_Flag_In_Pat to do control
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            ''20160923 - Add Interpose_PreMeas entry point by each sequence
            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
                End If
            End If
            
            TestOptLen = Len(Ts)
            
            TestSeqNumIdx = TestSeqNum
            
            For k = 1 To TestOptLen
                
                TestOption = Mid(Ts, k, 1)
                
                '' 20160106 - If "ForceFunctional_Flag" = True to let TestOption = "N" to make the test instance only run functional test
                If ForceFunctional_Flag = True Then
                    TestOption = "N"
                End If
                
                '' 20160705 - If second case is N that will cause error
                If (Measure_Pin_PPMU <> "") Then
                    Call Meas_VIR_IO_PreSetupBeforeMeasurement(TestPinArrayIV, TestSeqNumIdx)
                    
                    Select Case UCase(TestOption)
                    
                        Case "V"
                        
                            Call IO_HardIP_PPMU_Measure_V(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceISequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, Rtn_MeasVolt, FlowTestNme, _
                                    SpecialCalcValSetting, InstSpecialSetting, RAK_Flag, CUS_Str_MainProgram, Rtn_SweepTestName, OutputTname(TestSeqNum), WaitTime_VIRZ_Ary(0))
 
                             ''20160906 - Check store measurement or not
                            If Meas_StoreName <> "" Then
                                If MeasStoreName_Ary(TestSeqNum) <> "" Then
                                    Store_Rtn_Meas(StoreIndex) = Rtn_MeasVolt
                                    Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                                    StoreIndex = StoreIndex + 1
                                End If
                            End If
 
''                             ''20151028  CUS_MeasV_And_CalR -- TYCHENGG
''                            ''========================================================================================
''                            If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
''                                Call CUS_VIR_MainProgram_MeasV_CalR(TestPinArrayIV, TestSeqNum, CUS_CalR_Seq_Ary, ForceISequenceArray, Rtn_MeasVolt_CUS_R, CUS_CalR_VDD)
''                            End If
''                            ''========================================================================================
                            
                        Case "I"
                            
                            If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
                            
                            Call IO_HardIP_PPMU_Measure_I(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, TestIrange, FlowTestNme, CUS_Str_MainProgram, SpecialCalcValSetting, Rtn_MeasCurr, Rtn_SweepTestName, InstSpecialSetting, OutputTname(TestSeqNum), WaitTime_VIRZ_Ary(1))
                            
                            ''20160906 - Check store measurement or not
                            If Meas_StoreName <> "" Then
                                If MeasStoreName_Ary(TestSeqNum) <> "" Then
                                    Store_Rtn_Meas(StoreIndex) = Rtn_MeasCurr
                                    Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
                                    StoreIndex = StoreIndex + 1
                                End If
                            End If
                            
                        Case "R"
                            
                            Call IO_HardIP_PPMU_Measure_R(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, TestIrange, FlowTestNme, RAK_Flag, Rtn_SweepTestName, CUS_Str_MainProgram, OutputTname(TestSeqNum), WaitTime_VIRZ_Ary(2), SpecialCalcValSetting)
                        
                        Case "Z"
                            
                            If (Len(Ts) <> 1) Then ForceVSequenceArray(TestSeqNum) = ForceVSequenceArray(TestSeqNum) & ";sweep"

                            Call IO_HardIP_PPMU_Measure_Z(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), TestLimitPerPin_VIR, TestIrange, FlowTestNme, RAK_Flag, Rtn_SweepTestName, OutputTname(TestSeqNum), WaitTime_VIRZ_Ary(3))
                                    
                        Case "N"
                        
                        Case Else
                            TheExec.Datalog.WriteComment "Error Test Option, please select V, I or R"
                    
                    End Select
                    
                    Call Meas_VIR_IO_PostSetupAfterMeasurement(TestPinArrayIV, TestSeqNumIdx)
                End If
            Next k
            
            ''20161206-Restore force condiction after measurement

            If Interpose_PreMeas <> "" Then
                If UBound(Interpose_PreMeas_Ary) = 0 Then
                    Call SetForceCondition("RESTOREPREMEAS")
                ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
                    Call SetForceCondition("RESTOREPREMEAS")
                End If
            End If
            
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
            
        Next Ts
        
        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & ""
                
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        
        PatCount = PatCount + 1
        
        '' 20160923 - Add Interpose_PostTest entry point
        Call SetForceCondition(Interpose_PostTest)
    
'        If gl_FlowForLoop_DigSrc_SweepCode <> "" Then   '20180509 TER add
'            gl_FlowForLoop_DigSrc_SweepCode = ""
'        End If
    
        '' 20160211 - Process DigCapData by using DSP
''        If b_ProcessDigCapByDSP = True Then
            If DigCap_Sample_Size <> 0 Then
                Dim DigCapPinAry() As String, NumberPins As Long
                Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
                
                If NumberPins > 1 Then
                    Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
                    Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins, , DigCap_Pin.Value)
                ElseIf NumberPins = 1 Then
                    Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
                    Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave)
                    Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, , , DigCap_Pin.Value)
                End If
            End If
            
        If gl_FlowForLoop_DigSrc_SweepCode <> "" Then   '20180814 TER add
            gl_FlowForLoop_DigSrc_SweepCode = ""
            gl_FlowForLoop_DigSrc_SweepCode_Dec = "" '20190613 CT add for Decimal value printing
        End If

''        End If
    Next patt
    
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Connect
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    
    If DisableFRC = True Then
        Call ReStart_FRC(FRCPortName)
    End If
    '' 20160907 - Process calculate equation by dictionary.
    If Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Calc_Eqn)
    End If

    '' 20160713 - Call write functional result if cpu flag in pattern
    If (CPUA_Flag_In_Pat) Then
        Call HardIP_WriteFuncResult(, , Inst_Name_Str)
    End If
    
    DebugPrintFunc patset.Value  ' print all debug information
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Meas_VIR_IO_Universal_func"
    If AbortTest Then Exit Function Else Resume Next
  
End Function




Public Function Meas_FreqVoltCurr_Universal_func(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
Optional MeasV_Pins As String, _
Optional MeasF_PinS_SingleEnd As String, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, Optional MeasF_Flag_MeasureThreshold As Boolean = False, Optional MeasF_ThresholdPercentage As Double = 0.5, Optional MeasF_WaitTime As String, _
Optional MeasI_pinS As String, Optional MeasI_Range As String, Optional MeasI_WaitTime As String, _
Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
Optional SpecialCalcValSetting As CalculateMethodSetup = 0, _
Optional InstSpecialSetting As InstrumentSpecialSetup = 0, _
Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String = "", _
Optional Flag_SingleLimit As Boolean = False, Optional TestLimitPerPin_VFI As String = "FFF", _
Optional MeasF_PinS_Differential As String, Optional ForceFunctional_Flag As Boolean = False, _
Optional MeasF_WalkingStrobe_Flag As Boolean, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, Optional MeasF_WalkingStrobe_StepVoltage As Double, Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double, _
Optional Meas_StoreName As String, Optional Calc_Eqn As String, _
Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, Optional CharSetName As String, _
Optional ForceV_Val As String, Optional ForceI_Val As String, Optional UVI80_MeasV_WaitTime As String = "", _
Optional RAK_Flag As Enum_RAK, Optional WaitTime_VIRZ As String, Optional MSB_First_Flag As Boolean = False, Optional BV_Enable As Boolean, Optional Validating_ As Boolean) As Long

    
    '=======Turks add to bypass HIP_Universal shmoo post body=======
    Dim DevChar_Setup As String
    
    If TheExec.DevChar.Setups.IsRunning = True Then
        DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
        If TheExec.DevChar.Results(DevChar_Setup).startTime Like "1/1/0001*" Or TheExec.DevChar.Results(DevChar_Setup).startTime Like "0001/1/1*" Then ' initial run of shmoo, not the first point
            Shmoo_End = False
        End If
    End If
    
    If TheExec.DevChar.Setups.IsRunning = True And Shmoo_End = True Then
        Exit Function
   End If
    Dim PatCount As Long
    Dim PattArray() As String
    Dim Loopendnumber() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If
    
    If InStr(CUS_Str_MainProgram, "vt_sweep") > 0 And AMP_EYE_VT_CZ_Flag = True Then
        Call SWEEP_VT(CUS_Str_MainProgram, Interpose_PrePat)
    Else
        gl_Sweep_vt = ""
    End If
    
    If InStr(UCase(CUS_Str_MainProgram), UCase("V_Sweep")) > 0 Then
        Call SWEEP_V(CUS_Str_MainProgram, Interpose_PrePat)
    Else
        gl_Sweep_vt = ""
    End If

    
    Call HardIP_InitialSetupForPatgen
    If MSB_First_Flag = True Then
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Note: DSSC MSB First. Please check pattern content"
    End If
    
    
    Call ShmooEndFunction  ' For Shmoo DIgsrc
    Call HardIP_InitialSetupForPatgen
    Dim m_InstanceName As String
    m_InstanceName = LCase(TheExec.DataManager.instanceName)
    
    Dim i As Long, j As Long, k As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, MeasPinAry_V() As String, MeasPinAry_F() As String, MeasPinAry_I() As String, MeasPinAry_IRange() As String
    Dim MeasPinAry_F_Differential() As String
    Dim MeasureF_Pin_Differential As New PinList
    Dim Ts As Variant, TestOption As Variant, site As Variant
    Dim TestSeqNum As Integer
    Dim MeasureV_pin As New PinList, MeasureF_Pin_SingleEnd As New PinList, MeasureI_pin As New PinList
    Dim MeasureI_Pin_CurrentRange As String
    Dim TestNum As Long
    Dim InDSPwave As New DSPWave, OutDspWave As New DSPWave
    Dim ShowDec As String, ShowOut As String
    Dim patt As Variant
    Dim Pat As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Dim MeasureV_Pin_PPMU As String, MeasureV_Pin_UVI80 As String
    Dim d_MeasF_Interval As Double
    Dim FreqPinsCheckType() As String
    Dim ThisPinType As String
    Dim MeasF_EventSource As FreqCtrEventSrcSel
    Dim MeasF_EnableVtMode As Boolean
    Dim Split_F_Str() As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName  '20170728 Added for HardIP_WriteFuncResult Output
    Dim restore_Flag As Boolean
    
    
    ''20160906 - Return measurement to directionary if needed
    Dim Rtn_MeasVolt As New PinListData, Rtn_MeasCurr As New PinListData, Rtn_MeasFreq As New PinListData
    Dim MeasStoreName_Ary() As String
    Dim Interpose_PreMeas_Ary() As String
    
''    Dim RTN_InterposeString As String
    
    On Error GoTo errHandler
    Dim CheckDSPWave As New DSPWave
    Dim Sweep_Enable As Boolean: Sweep_Enable = False
    Dim Sweep_Loop_Calc_Eqn As String: Sweep_Loop_Calc_Eqn = ""
'    Dim Sweep_Calc_Eqn As Boolean: Sweep_Calc_Eqn = False
    Dim Sweep_Calc_Eqn_index As String: Sweep_Calc_Eqn_index = ""
    Dim Sweep_Dictionary As String: Sweep_Dictionary = ""
    Dim Sweep_Calc_Eqn As String: Sweep_Calc_Eqn = ""
    
    Dim OutputTname() As String

    Call tl_PinListDataSort(True)
    Dim instance_name As String

    instance_name = TheExec.DataManager.instanceName

'**************************************************
'SeaHawk Edited by 20190606
    Dim SpecialUsePatName As String
    SpecialUsePatName = CStr(patset)
    If CUS_Str_DigCapData Like "*Special_DigCapData_Setting*" Then
        gl_SpecialString = ""
        SpecialUsePatName = SpecialUsePatName & "_SpecialDigCap"
        Public_GetStoredString (SpecialUsePatName)
        CUS_Str_DigCapData = gl_SpecialString
    End If
    
    If CUS_Str_DigSrcData Like "*Special_DigSrcData_Setting*" Then
        gl_SpecialString = ""
        SpecialUsePatName = SpecialUsePatName & "_SpecialDigSrc"
        Public_GetStoredString (SpecialUsePatName)
        CUS_Str_DigSrcData = gl_SpecialString
    End If
'**************************************************
    
       If gl_Flag_HardIP_Characterization_1stRun = False Then 'Then: Exit Function
        If TheExec.DevChar.Setups.IsRunning = True And CStr(TheExec.DevChar.Setups.ActiveSetupName) Like "*SweepDigSrc*" Then
            Call ReDefineDigSrcForCharacterization(DigSrc_Assignment)
        End If
      End If
    
    '================================================================ Roger
    If InStr(1, LCase(Interpose_PrePat), "sweep:") <> 0 Then
        Dim Sweep_Info() As Power_Sweep
        Dim Sweep_CUS_Str_DigCapData As String
        Call SortSweepInfo(Sweep_Info, Interpose_PrePat)
        Sweep_Enable = True
        Sweep_CUS_Str_DigCapData = CUS_Str_DigCapData
        Sweep_Calc_Eqn = Calc_Eqn
    End If
    '================================================================
    ''20170322-Store MeasF mid value for VT
    Dim SplitFreqVtValue() As String
    Dim DictKey_StoreVT As String
    Dim Dict_VT_Value As New SiteDouble
    
    
    
    'If (UCase(MeasI_Range) Like "*CP*:*" Or UCase(MeasI_Range) Like "*FT*:*") Then MeasI_Range = Select_MeasIRange(MeasI_Range, CurrentJobName_U)   ' support different Meter_Range in different Job, add by Roger 20170628
    
    '' 20160201 - Check input argumenets whether have "@" in the first character. Add it If no "@" in the beginning. Then remove it to process fomat.
    Call VFI_AnalyzedInputStringByAt(MeasV_Pins, MeasF_PinS_SingleEnd, MeasI_pinS, MeasI_Range, MeasF_PinS_Differential, ForceV_Val, ForceI_Val)
    
    Dim ForceV_Val_Ary() As String
    Dim ForceI_Val_Ary() As String
    Dim MeasurePin_ForceV_Val As String
    Dim MeasurePin_ForceI_Val As String
    Dim MeasI_WaitTime_Ary() As String
    Dim MeasF_WaitTime_Ary() As String
    Dim UVI80_MeasV_WaitTime_Ary() As String
    
    
    
    If TestSequence = "" Then                       '20170714
        ReDim TestSequenceArray(0) As String
        TestSequenceArray(0) = TestSequence
    Else
        TestSequenceArray = Split(TestSequence, ",")
    End If
    MeasStoreName_Ary = Split(Meas_StoreName, ",")
    Interpose_PreMeas_Ary = SplitInputCondition(Interpose_PreMeas, "|") ''Carter, 20190616
    Dim PreMeas_Ary() As String
    If Interpose_PreMeas <> "" Then
        PreMeas_Ary = ParseData_InterPose(Interpose_PreMeas_Ary, TestSequenceArray)
    End If
    '----------------------------20180523
    
    'Roger New,20180510 TName
    '--------------------------------------------------------------------
    'Call GetFlowTName
    
    '----------------------------20180523
        
    'Call VFI_ProcessInputString(TestSequence, MeasV_PinS, MeasI_pinS, MeasF_PinS_SingleEnd, MeasF_PinS_Differential, MeasI_Range, Meas_StoreName, Interpose_PreMeas, _
                                            ForceV_Val, ForceI_Val, _
                                            TestSequenceArray(), MeasPinAry_V(), MeasPinAry_I(), MeasPinAry_F(), _
                                            MeasPinAry_F_Differential(), MeasPinAry_IRange(), MeasStoreName_Ary(), Interpose_PreMeas_Ary(), ForceV_Val_Ary(), ForceI_Val_Ary())

    'Call VFI_ProcessWaitTimeString(MeasI_WaitTime, MeasF_WaitTime, UVI80_MeasV_WaitTime, MeasI_WaitTime_Ary(), MeasF_WaitTime_Ary(), UVI80_MeasV_WaitTime_Ary(), TestSequenceArray())
                                            
    
'    Call HIP_Evaluate_ForceVal(ForceV_Val_Ary())
'
'    Call HIP_Evaluate_ForceVal(ForceI_Val_Ary())
    
    ''20170807 - CZ test name index
'    gl_CZ_FlowTestNameIndex = 0

'    Call Freq_ProcessEventSourceTerminationMode(MeasF_EventSourceWithTerminationMode, MeasF_EventSource, MeasF_EnableVtMode)
    
    ''20141219 Get use-limit from flow table
'    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    
    ''20161130-Get test name from flow table
    Dim FlowTestNme() As String
    ''========================================================================================
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
    ''========================================================================================
    If TheExec.DevChar.Setups.IsRunning Then
        If CharSetName <> "" And InStr(UCase(Interpose_PrePat), ":TERM:") <> 0 Then
        'HIO:can not applylevelsTiming for the first point  of run_shmoo
        ElseIf TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Contains(tlDevCharShmooAxis_Y) Then
            If gl_Flag_HardIP_Trim_Set_PrePoint And Not (gl_Flag_HardIP_Characterization_1stRun) Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value)
            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", 24000000#)
            Else
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            End If
        Else
                        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                End If
    ElseIf BV_Enable Then
        Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    Dim Loop_Idx As Long
    Dim Loop_count As Long
    Dim Loop_Init As Long
    Dim Loop_Max As Long
    Dim Loop_Step As Long
    Dim Loop_BitNum As Long
    Dim Loop_RegName As String
    Dim SplitLoop_RegName() As String
    Dim Split_Loop_DigSrc_Str() As String
    Dim BinStr As String
    Dim Loop_SplitByComma() As String
    Dim Loop_SplitByEqual() As String
    Dim Loop_Digsrc_name As String
    Dim Split_DigSrc_Equation() As String
    
    Loop_Idx = 0
    Loop_Init = 0
    Loop_Max = 0
    Loop_Step = 1
    
    If (Sweep_Enable = True) Then
        Loop_Max = Sweep_Info(0).Count - 1
    
    End If
    Dim timer_ As Double
    
    'timer_ = theexec.Timer()
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1164 As Long: VBT_LIB_HardIP_ProfileMark_1164 = ProfileMarkEnter(2, instance_name & "_" & "ProcessInputToGLB&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1158")   ' Profile Mark
    
''If TheExec.DataManager.instanceName Like "*D2DEXLBK*" Then   'CSHOX for D2D ELB too long
If DigSrc_Equation Like "*Duplicate*" Then    '' Determine if it for for too long string case
'' 190814 re-edit set a format to loop the DigSrc assignment : Duplicate:[512](Loop count):[D2D_PHY__ZCPU_ZCPD__zcpu+D2D_PHY__ZCPU_ZCPD__zcpd](LoopName)
    Split_DigSrc_Equation = Split(DigSrc_Equation, ":")
    Loop_Digsrc_name = Split_DigSrc_Equation(2) ''LoopName
    For i = 0 To CInt(Split_DigSrc_Equation(1)) - 1   ''LoopCount
        If i = 0 Then
            DigSrc_Equation = Loop_Digsrc_name
        Else
            DigSrc_Equation = DigSrc_Equation & "+" & Loop_Digsrc_name
        End If
    Next i
End If
   
    
    Call ProcessInputToGLB(patset, TestSequence, CPUA_Flag_In_Pat, DisableComparePins, DisableConnectPins, DisableFRC, FRCPortName, MeasV_Pins, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, MeasF_Flag_MeasureThreshold, _
                            MeasF_ThresholdPercentage, MeasF_WaitTime, MeasI_pinS, MeasI_Range, MeasI_WaitTime, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, _
                            DigSrc_FlowForLoopIntegerName, SpecialCalcValSetting, InstSpecialSetting, CUS_Str_MainProgram, CUS_Str_DigCapData, CUS_Str_DigSrcData, Flag_SingleLimit, TestLimitPerPin_VFI, MeasF_PinS_Differential, ForceFunctional_Flag, _
                            MeasF_WalkingStrobe_Flag, MeasF_WalkingStrobe_StartV, MeasF_WalkingStrobe_EndV, MeasF_WalkingStrobe_StepVoltage, MeasF_WalkingStrobe_BothVohVolDiffV, MeasF_WalkingStrobe_interval, MeasF_WalkingStrobe_miniFreq, Meas_StoreName, _
                            Calc_Eqn, Interpose_PrePat, Interpose_PreMeas, Interpose_PostTest, CharSetName, ForceV_Val, ForceI_Val, UVI80_MeasV_WaitTime, RAK_Flag, WaitTime_VIRZ)
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1164    ' Profile Mark
            
    'theexec.Datalog.WriteComment "ProcessInputToGLB Time : " & FormatNumber(theexec.Timer(timer_), 6) & ":" & theexec.DataManager.instanceName & ":" & TestSequence & ":" & CStr(DigSrc_Sample_Size) & ":" & DigSrc_Equation & ":" & DigSrc_Assignment
    
    If InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
        meas_val_delay_instance_name = ""
        'Ex: CUS_Str_MainProgram ==> Loop_DigSrc;1;119;10;32;ddr0_mdll0_lsw:ddr0_mdll1_lsw:ddr1_mdll0_lsw:ddr1_mdll1_lsw
        Split_Loop_DigSrc_Str = Split(CUS_Str_MainProgram, ";")
        Loop_Init = Split_Loop_DigSrc_Str(1)
        Loop_Max = Split_Loop_DigSrc_Str(2)
        Loop_Step = Split_Loop_DigSrc_Str(3)
        Loop_BitNum = Split_Loop_DigSrc_Str(4)
        Loop_RegName = Split_Loop_DigSrc_Str(5)
        SplitLoop_RegName = Split(Loop_RegName, ":")
        Loopendnumber = Split(Split_Loop_DigSrc_Str(6), "$")
        
    End If
    
    Dim loop_i As Long, Loop_j As Long
    Dim Temp_Equal_Str As String
    Dim Final_Comma_Str As String
    Temp_Equal_Str = ""
    Final_Comma_Str = ""
    
    For Loop_count = Loop_Init To Loop_Max
       If InStr(UCase(CUS_Str_MainProgram), UCase("loopendnum")) <> 0 And Loop_count = Loop_Max Then
           
                Loop_count = Loopendnumber(1)
     
       End If
      
      If InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
            Call Public_AddStoredString(Split_Loop_DigSrc_Str(0), CStr(Loop_count))
      End If
      
        
        If InStr(UCase(CUS_Str_MainProgram), UCase("Calc_Freq_SDLL_SWP")) <> 0 Then gl_Tname_Alg_Index = Loop_count
        
        'TypeName (Loop_count)
        
        If (Sweep_Enable = True) Then
            CUS_Str_DigCapData = Sweep_CUS_Str_DigCapData
            
            Call SetForceSweepVoltAndTName(Sweep_Info, CUS_Str_DigCapData, Loop_count)
            
            If InStr(UCase(TheExec.DataManager.instanceName), "MTRGR_T2P6") <> 0 Or InStr(UCase(TheExec.DataManager.instanceName), "MTRGR_T2P7") <> 0 Then
                Calc_Eqn = Replace(Calc_Eqn, Replace(Split(Split(Calc_Eqn, ":")(2), "(")(1), ")", ""), Split(CUS_Str_DigCapData, ":")(2))
                CUS_Str_DigCapData = Replace(CUS_Str_DigCapData, Split(CUS_Str_DigCapData, ":")(1), Split(CUS_Str_DigCapData, ":")(1) & CStr(Loop_count))
            Else
                Calc_Eqn = Sweep_Calc_Eqn & "," & CStr(Loop_count)
            End If
        End If
        
        
        If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
            BinStr = Dec2BinStr32Bit_Rev(Loop_BitNum, Loop_count)
            Loop_SplitByComma = Split(DigSrc_Assignment, ";")
            
            For loop_i = 0 To UBound(Loop_SplitByComma)
                Loop_SplitByEqual = Split(Loop_SplitByComma(loop_i), "=")
                For Loop_j = 0 To UBound(SplitLoop_RegName)
                    If UCase(Loop_SplitByEqual(0)) = UCase(SplitLoop_RegName(Loop_j)) Then
                        Loop_SplitByEqual(1) = BinStr
                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
                        Exit For
                    Else
                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
                    End If
                Next Loop_j
                If loop_i = 0 Then
                    Final_Comma_Str = Temp_Equal_Str
                Else
                    Final_Comma_Str = Final_Comma_Str & ";" & Temp_Equal_Str
                End If
            Next loop_i
            DigSrc_Assignment = Final_Comma_Str
        End If
        
        
        '' 20190529 - Add for sweep force V
        If InStr(Interpose_PrePat, "x_sweep") <> 0 Then
            gl_Sweep_Glb_TName = CDbl(Val(TheExec.Flow.var("x_sweep").Value)) / 1000
            Interpose_PrePat = Replace(Interpose_PrePat, "x_sweep", gl_Sweep_Glb_TName)
            
            'USB_DP:V:x_sweep;Sweep_Name:
            If InStr(Interpose_PrePat, "Sweep_Name") <> 0 Then
                Interpose_PrePat = Replace(Interpose_PrePat, "Sweep_Name:", "")
            End If
        End If
        
        If InStr(Interpose_PrePat, "x_power_sweep") <> 0 Then
            gl_Sweep_Glb_TName = CDbl(Val(TheExec.Flow.var("x_power_sweep").Value)) / 1000
            Interpose_PrePat = Replace(Interpose_PrePat, "x_power_sweep", gl_Sweep_Glb_TName)
            
            If InStr(Interpose_PrePat, "Sweep_Name") <> 0 Then
                Interpose_PrePat = Replace(Interpose_PrePat, "Sweep_Name:", "")
            End If
        End If
        
          
        '' 20190531 - Add for sweep Volt by shmoo
        If InStr(Interpose_PrePat, "Volt_sweep_GLB") <> 0 Then
        For Each site In TheExec.sites
            gl_Sweep_Glb_TName = CDbl(TheExec.specs.Globals("Volt_sweep_GLB").CurrentValue)
            Exit For
        Next site
            Interpose_PrePat = Replace(Interpose_PrePat, "Volt_sweep_GLB", gl_Sweep_Glb_TName)
            

            If InStr(Interpose_PrePat, "Sweep_Name") <> 0 Then
                Interpose_PrePat = Replace(Interpose_PrePat, "Sweep_Name:", "")
            End If
        End If
         
        
        '' 20160923 - Add Interpose_PrePat entry point
        If Interpose_PrePat <> "" Then
            Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
        End If
        
        If gl_Flag_HardIP_Characterization_1stRun Then: Exit Function
            
        ''20161205 - Force_Flow_Shmoo_Condition
        If TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then Force_Flow_Shmoo_Condition
        'Do Flow Shmoo
        
        If patset.Value <> "" Then
            Shmoo_Pattern = patset.Value '' 20170808 add for shmoo pattern name print
            TheHdw.Patterns(patset).Load
            Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
        Else
            ReDim PattArray(0)
            PattArray(0) = ""
        End If
            
        If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
        If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
        
        ''20161107-Return sweep test name
        Dim Rtn_SweepTestName As String
        Rtn_SweepTestName = ""
        gl_TName_Pat = patset.Value
        
        Dim current_pat_index As Integer
        current_pat_index = 0
        
        
        '20191003 add for CPM with Multi_Init(DigSrc)_PL
        Dim DigSrc_Equation_temp_array() As String
        Dim DigSrc_Assignment_temp_array() As String
        Dim DigSec_Multi_Init_PL__Seq_index As Long: DigSec_Multi_Init_PL__Seq_index = 0
       
        If (InStr(UCase(CUS_Str_MainProgram), "CPM_MULTI_INIT_DIGSRC") > 0) Then
            DigSrc_Equation_temp_array = Split(DigSrc_Equation, "|")
            DigSrc_Assignment_temp_array = Split(DigSrc_Assignment, "|")
        End If
       
        
        
        For Each patt In PattArray
            If patt <> "" Then
                TheExec.Flow.TestLimitIndex = 0
                Pat = CStr(patt)
                TheHdw.Patterns(Pat).Load
'                                                                                                                                                            Dim VBT_LIB_HardIP_ProfileMark_1267 As Long: VBT_LIB_HardIP_ProfileMark_1267 = ProfileMarkEnter(2, instance_name & "_" & "GenDigSrc&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1264")    ' Profile Mark
                If (InStr(UCase(CUS_Str_MainProgram), "CPM_INIT_DIGSRC") > 0) And (UCase(patt) Like "*_PL*") Then

                Else
                    '20191003 add for CPM with Multi_Init(DigSrc)_PL
                    If (InStr(UCase(CUS_Str_MainProgram), "CPM_MULTI_INIT_DIGSRC") > 0) Then
                        DigSrc_Equation = DigSrc_Equation_temp_array(DigSec_Multi_Init_PL__Seq_index)
                        DigSrc_Assignment = DigSrc_Assignment_temp_array(DigSec_Multi_Init_PL__Seq_index)
                        DigSec_Multi_Init_PL__Seq_index = DigSec_Multi_Init_PL__Seq_index + 1
                    End If
                    
                    
                
                Set InDSPwave = Nothing
                
'*******************************New Feature for trimcode table*******************************
'Added by  20190509
                If LCase(DigSrc_Assignment) Like "*table*" Then DigSrc_Assignment = DigSrc_Assignment & "_" & CStr(TheExec.Flow.var("SrcCodeIndx").Value)
'********************************************************************************************
                
                Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, _
                                                       DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPwave, Rtn_SweepTestName, MSB_First_Flag)
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1267    ' Profile Mark
                                End If
                Set OutDspWave = Nothing
                Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                 
                Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
                

                If InStr(UCase(CUS_Str_MainProgram), "MTR_UVI80_SETUP") <> 0 Then
                    Call MTR_UVI80_Setup
                End If
            
            
            
                Dim SplitByCommaStr() As String
                Dim ForcePin_X As String
                Dim ForcePin_Y As String
                Dim SweepIndexStr_X As String
                Dim ForceVal_X As Double
                If LCase(CUS_Str_MainProgram) Like "*x_sweep*" Then
                    SplitByCommaStr = Split(CUS_Str_MainProgram, ",")
                    SweepIndexStr_X = SplitByCommaStr(0)
                         ForcePin_X = SplitByCommaStr(1)
                         
                          ForceVal_X = CDbl(Val(TheExec.Flow.var(SweepIndexStr_X).Value)) / 1000
                          TheExec.Datalog.WriteComment "ForcePin = " & ForcePin_X & "; ForceVal_X = " & ForceVal_X & "V"
                          'TheExec.Datalog.WriteComment "ForcePin = " & SplitByCommaStr(2) & ";  ForceVal_X  = " & ForceVal_X & "V"
                          'TheHdw.DCVS.Pins(ForcePin_X).Voltage.Value = ForceVal_X
                          'TheHdw.DCVS.Pins(SplitByCommaStr(2)).Voltage.Value = ForceVal_X
                          
                        TheHdw.Digital.Pins(ForcePin_X).Disconnect
                        
                            With TheHdw.PPMU.Pins(ForcePin_X)
                                .Gate = tlOff
                                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
                                .ForceV CDbl(ForceVal_X), 0.02
                                .Connect
                                .Gate = tlOn
                            End With
                          
                          
                        FourceV = ForceVal_X
                End If
            
            
            
                '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
                If (CPUA_Flag_In_Pat) Then
                    Call TheHdw.Patterns(Pat).start
                Else
                    Call TheHdw.Patterns(Pat).Test(pfNever, 0)
                End If
            End If
            
            'TestSeqNum = 0
            
            'Call ProcessTestNameInputString(OutputTname, UBound(TestSequenceArray))    Remove
            
'            If PatCount > 1 Then
'                Dim ot_cnt As Long
'                For ot_cnt = 0 To UBound(OutputTname)
'                    OutputTname(ot_cnt) = OutputTname(ot_cnt) & Split(Split(Split(Pat, "\")(UBound(Split(Pat, "\"))), ":")(0), "_")(12)
'                Next ot_cnt
'            End If
            

            TestSeqNum = 0


            For Each Ts In TestSequenceArray
                Instance_Data.TestSeqNum = TestSeqNum
                ''20150907 - Only need CPUA_Flag_In_Pat to do control
                If (CPUA_Flag_In_Pat) Then
                    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
                Else
                    Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
                End If
                
                If InStr(MeasF_PinS_SingleEnd, "$") Then
                    Dim MeasF_Set() As String
                    MeasF_Set = Split(MeasF_PinS_SingleEnd, ",")
                End If
                
                ''20160923 - Add Interpose_PreMeas entry point by each sequence
                    
            '''------- Carter, 20190616
'                If Interpose_PreMeas <> "" Then
'                    If UBound(Interpose_PreMeas_Ary) = 0 Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
'                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
'                    End If
'                End If
            '''------- Carter, 20190616
                
                TestOptLen = Len(Ts)
                
                
                
                For k = 1 To TestOptLen
                    Instance_Data.TestSeqSweepNum = k - 1
                    TestOption = Mid(Ts, k, 1)
                    
               '''-------Start - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                    If Interpose_PreMeas <> "" Then
                        If PreMeas_Ary(TestSeqNum, k - 1) <> "" Then
                            Call SetForceCondition(PreMeas_Ary(TestSeqNum, k - 1) & ";STOREPREMEAS")
                        End If
                    End If
               '''-------End - Add per sweep feature for interpose_premeas - Carter, 20190614-------
               
                    For Each site In TheExec.sites.Active
                        TestNum = TheExec.sites.Item(site).TestNumber
                    Next site
                    
                    '----------------0427 begin-------------------------------------
                    If InStr(MeasF_PinS_SingleEnd, "$") Then
                        MeasureF_Pin_SingleEnd = Replace(MeasF_Set(current_pat_index), "$", "")
                    End If
                    '----------------0427 end---------------------------------------
                    Select Case UCase(TestOption)
                        Case "V"
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1347 As Long: VBT_LIB_HardIP_ProfileMark_1347 = ProfileMarkEnter(2, instance_name & "_" & "MeasV&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1345")    ' Profile Mark
                            
                            Call HardIP_MeasureVolt
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1347    ' Profile Mark
                        Case "F"
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1352 As Long: VBT_LIB_HardIP_ProfileMark_1352 = ProfileMarkEnter(2, instance_name & "_" & "MeasF&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1350")    ' Profile Mark
                            
                            Call HardIP_MeasureFreq
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1352    ' Profile Mark
                        Case "I"
                            If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1358 As Long: VBT_LIB_HardIP_ProfileMark_1358 = ProfileMarkEnter(2, instance_name & "_" & "MeasI&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1356")    ' Profile Mark
                            
                            Call HardIP_MeasureCurrent
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1358    ' Profile Mark
                        Case "R"
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1363 As Long: VBT_LIB_HardIP_ProfileMark_1363 = ProfileMarkEnter(2, instance_name & "_" & "MeasR&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1361")    ' Profile Mark
                            
                            Call HardIP_SetupAndMeasureR
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1363    ' Profile Mark
                        Case "Z"
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1368 As Long: VBT_LIB_HardIP_ProfileMark_1368 = ProfileMarkEnter(2, instance_name & "_" & "MeasZ&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1366")    ' Profile Mark
                            
                            Call HardIP_SetupAndMeasureZ
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1368    ' Profile Mark
                        Case "P"
                            HardIP_BySeqCurrentProfile
                        Case "N"
                            restore_Flag = True
                        Case Else
                            TheExec.Datalog.WriteComment "Error: Test Option " & UCase(TestOption) & " cannot be recognized!!!"
                    End Select
                    If TheExec.sites.Active.Count = 0 Then Exit Function
                    
                '''-------Start - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                    If Interpose_PreMeas <> "" Then
                        If PreMeas_Ary(TestSeqNum, k - 1) <> "" And UCase(TestOption) <> "N" Then
                            Call SetForceCondition("RESTOREPREMEAS")
                        End If
                    End If
                '''-------End - Add per sweep feature for interpose_premeas - Carter, 20190614-------
                Next k
                
                ''20161206-Restore force condiction after measurement
    ''            Call SetForceCondition("RESTORE")
    
    '''------- Carter, 20190616
'                If Interpose_PreMeas <> "" And Ts <> "N" Then
'                    If UBound(Interpose_PreMeas_Ary) = 0 Then
'                        Call SetForceCondition("RESTOREPREMEAS")
'                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                        Call SetForceCondition("RESTOREPREMEAS")
'                    End If
'                End If
   '''------- Carter, 20190616
                
                TestSeqNum = TestSeqNum + 1
                
                If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
                Instance_Data.TestSeqNum = TestSeqNum
                
            Next Ts
            
            If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & ""
            
            TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end
            
            PatCount = PatCount + 1
            
            '' 20160923 - Add Interpose_PostTest entry point
            Call SetForceCondition(Interpose_PostTest)
            
'            If gl_FlowForLoop_DigSrc_SweepCode <> "" Then         '20180509
'                gl_FlowForLoop_DigSrc_SweepCode = ""
'            End If
            
            '' 20160211 - Process DigCapData by using DSP
    ''        If b_ProcessDigCapByDSP = True Then
                If DigCap_Sample_Size <> 0 Then
                    Dim DigCapPinAry() As String, NumberPins As Long
                    Dim CUS_Str_DigCapData_temp As String
                    Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
                    
                    
                    If NumberPins > 1 Then
                        Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
                        Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
                        Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins, , DigCap_Pin.Value)
    
                    ElseIf NumberPins = 1 Then
                        Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
                        Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
'                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1429 As Long: VBT_LIB_HardIP_ProfileMark_1429 = ProfileMarkEnter(2, instance_name & "_" & "MeasC&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1427")    ' Profile Mark
                        
                        Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, CUS_Str_MainProgram, , DigCap_Pin.Value, , MSB_First_Flag)
'                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1429    ' Profile Mark
                    End If
                End If
                
                '' 20160907 - Process calculate equation by dictionary.
                If Calc_Eqn <> "" And InStr(LCase(TestSequence), "p") = 0 Then
                    Call ProcessCalcEquation(Calc_Eqn)
                End If
                
                '' 20160713 - Call write functional result if cpu flag in pattern
                'If (CPUA_Flag_In_Pat) Then
                    Call HardIP_WriteFuncResult(, , Inst_Name_Str)
                'End If
                
                If gl_FlowForLoop_DigSrc_SweepCode <> "" Then        '20180814
                    gl_FlowForLoop_DigSrc_SweepCode = ""
                    gl_FlowForLoop_DigSrc_SweepCode_Dec = "" '20190613 CT add for Decimal value printing
                End If
                
    ''        End If
    
            current_pat_index = current_pat_index + 1
            
                If Interpose_PreMeas <> "" And restore_Flag = True Then
                    Call SetForceCondition("RESTOREPREMEAS")

                End If
            
            
            
               If LCase(CUS_Str_MainProgram) Like "*x_sweep*" Then

                    With TheHdw.PPMU.Pins(ForcePin_X)
                            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range ''FVMI - Carter, 20190503
                            .Disconnect
                            .Gate = tlOff
                    End With
                    TheHdw.Digital.Pins(ForcePin_X).Connect ''Connect Digital pins after measurement - Carter, 20190503

                End If
                    
                  
            gl_Sweep_Glb_TName = "" '' 20190529 - Add for sweep force V
                  
        Next patt
        
''        ''20170405-Record all functional test result from flow for loop opcode, use global string to store them
        If CUS_Str_DigSrcData <> "" And UCase(CUS_Str_DigSrcData) = UCase("BinToGray") Then
            If CPUA_Flag_In_Pat = False Then
                Call DisplayForLoopFuncResult_EndOfTest(CUS_Str_DigSrcData, Rtn_SweepTestName, CPUA_Flag_In_Pat, DigSrc_FlowForLoopIntegerName)
            End If
        End If
     If MeasureV_pin <> "" Then
         Call EndSetupForMeasureVoltPins(MeasureV_Pin_PPMU, MeasureV_Pin_UVI80)
     End If
     
     If DisableConnectPins <> "" Then TheHdw.Digital.Pins(DisableConnectPins).Connect
     If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
         
     If DisableFRC = True Then
         Call ReStart_FRC(FRCPortName)
     End If
     
     DebugPrintFunc patset.Value   ' print all debug information
     
     If TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then
     'Do Flow Shmoo
         If Flow_Shmoo_Port_Name <> "" Then Restart_All_Freerun_Clk
     End If
     
     If Interpose_PrePat <> "" Then
         Call SetForceCondition("RESTOREPREPAT")
     End If
     
     ''=============================== CharSetName ====================================
     Dim p As Variant
     If TheExec.DevChar.Setups.IsRunning = False And CharSetName <> "" Then
         Dim ApplyPins As String, Setup_mode As String, p_ary() As String, p_cnt As Long
         'If TheExec.DevChar.Setups(CharSetName).TestMethod.Value = tlDevCharTestMethod_Reburst Then TheExec.Datalog.WriteComment "[PrintCharCondition:" & PrintCharSetup(Interpose_PrePat_GLB) & ",Test]"
         Setup_mode = TheExec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Name
         If (LCase(Setup_mode) <> "vid" And LCase(Setup_mode) <> "vicm") Then
             ApplyPins = TheExec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins
             TheExec.DataManager.DecomposePinList ApplyPins, p_ary, p_cnt
             For Each p In p_ary
                 TheExec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = p
                 run_shmoo CharSetName
             Next p
             TheExec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = ApplyPins
         Else
             ApplyPins = TheExec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins
             p_ary = Split(ApplyPins, ",")
             For Each p In p_ary
                 TheExec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = p
                 run_shmoo CharSetName
             Next p
             TheExec.DevChar.Setups(CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins = ApplyPins
             'run_shmoo CharSetName
         End If
     End If
    
    If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 And Loop_Step <> 1 Then
        Loop_count = Loop_count + Loop_Step - 1
    End If
 
    If InStr(UCase(CUS_Str_MainProgram), UCase("V_Sweep")) > 0 Then
        sweep_power_val_per_loop_count = ""
    End If
 
 
    Next Loop_count
    ''================================================================================

    If InStr(DigSrc_Assignment, "digsrctable") <> 0 Then
                  'Table_Decvalue = ""
                  gl_SweepNum = ""
    End If

    ReDim TestConditionSeqData(0)
    Dim Instance_Data_temp() As Instance_Type
    ReDim Instance_Data_temp(0)
    Instance_Data = Instance_Data_temp(0)
    
    Meas_StoreName_Flag = False ''Carter, 20190521
    temp_CUS_String = ""
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Meas_FreqVoltCurr_Universal_func"
'    Resume Next
    If AbortTest Then Exit Function Else Resume Next
  
End Function


Public Function Opt_DdrLpBkFunc2(DqsSwpPat As Pattern, DqSwpPat As Pattern, _
                            DisableComparePins As PinList, DisableConnectPins As PinList, _
                            DigCap_Pin As PinList, NoOfBists As Long, _
                            DqSwpNoOfBits As String, DqsSwpNoOfBits As String, _
                            Optional DispCaptStrm As Boolean = True, _
                            Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As String, _
                            Optional DqsDigSrc_Equation As String, Optional DqDigSrc_Equation As String, _
                            Optional DigSrc_Assignment As String, _
                            Optional CUS_Str_DigSrcData As String, _
                            Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
                            Optional EyeTestRegName As String, _
                            Optional DigCap_Sample_Size_Dqs As Long, _
                            Optional CUS_Str_DigCapData_Dqs As String, _
                            Optional DigCap_Sample_Size_Dq As Integer, _
                            Optional CUS_Str_DigCapData_Dq As String, _
                            Optional Interpose_PrePat As String, _
                            Optional SweepVtStr As String, _
                            Optional Calc_Eqn As String, _
                            Optional DigSrc_FlowForLoopIntegerName As String = "", _
                            Optional Validating_ As Boolean) As Long

    ''''--------------------------------------------------------------------------------------------------
    ''''    Based on TMA V04C function "Meas_FreqVoltCurr_Univeral_func" in VBT_LIB_HardIP_New Module
    ''''    rev 0, by Zheng Xiao, Apple Inc, 1/1/2016
    ''''
    ''''    Adapted for Starling DDR external loopback test.
    ''''        - MDLL code no longer captured in this test
    ''''        - DQ ELB lower limit now fixed as 1/4 of PI eye openings (DQ 64, CA 128)
    ''''        - Impedance cal settings needed to be sourced in
    ''''    rev 1, by Zheng Xiao, Apple Inc, 1/26/2016
    ''''--------------------------------------------------------------------------------------------------
    ''''    Test fucntion for DDR (AMP) loopback test, with data eye sweeping. Applicable to all buses, both
    '''         internal and external loopback tests.
    ''''    There are two sweeps, involving 2 patterns, sweeping left and right from the center point, respectively.
    ''''        The eye width will be the combination of the 2 sweeps, and being tested
    ''''
    ''''    In case of multiple eyes, the maximum eye will be tested.
    ''''
    ''''    DQ sweep : Moving DQ strobe the captured bits stream are consecutive results from the first lane to the last
    ''''    DQS sweep: moving DQS strobe
    ''''
    ''''--------------------------------------------------------------------------------------------------
    ''''    Usage
    ''''        Opt_DdrLpBkFunc2 is to be used to construct test instance directly.
    ''''--------------------------------------------------------------------------------------------------
    ''''    Function calls
    ''''        - PATT_GetPatListFromPatternSet (original)
    ''''        - DigCapSetup (original): inline codes should be used here.
    ''''        - SetupDigSrcDspWave (original): setup dssc dig source. inline be better
    ''''        - FindMaxEyeWidth: DSP fucntional call stitch 2 sweeps to a single eye diagram of each BIST,
    ''''            reporting the eyewidths
    ''''        - DebugPrintFunc (original)
    ''''--------------------------------------------------------------------------------------------------
    ''''    Modifications:
    ''''        - Completely re-written for speciallized function for DDR eye-sweep based loopback tests
    ''''        - Eliminated the need to pass the first sweep results via global variable, by including
    ''''            both sweeps in a single function
    ''''        - Instead of using VBT for waveform conversion, processing, and eye width finding, using
    ''''            DSP based function calls for efficiencies and multi-site handling
    ''''--------------------------------------------------------------------------------------------------
    ''''    Argument List
    ''''        DqSwpPat:           Pattern set for DQ sweep
    ''''        DqsSwpPat:          Pattern set for DQS sweep
    ''''        DisableComparePins: Retained from the original function. Pins to be masked
    ''''        DisableConnectPins: Retained. Pins to be disconnected. (DisconnectPins is a more suitable name)
    ''''        DigCap_Pin:         Retained. Pin group on which the digital data to be sourced
    ''''        NoOfBists:          DDR LB test consists of individual blocks, suchas lanes, byte.
    ''''        DqSwpNoOfBits:      Data points in SWQ sweep
    ''''        DqsSwpNoOfBits:     Data points in SWK sweep
    ''''        DispCaptStream:     If true the captured data would be displayed as bit stream
    ''''--------------------------------------------------------------------------------------------------
    ''''    NOTE: The impedance settings are to be sourced.
    ''''        They include zcpu, zcpd, dspu, and dspd for each instance. There are 2 instances for Starling
    ''''            Among them the first 3 are to be calibrated based on the test condition and performance mode,
    ''''            and dspd fixed.
    ''''        At the time this function is being developed, it's not clear how those calibration results as well
    ''''            dspd would be passed to this function. For the moment the an assumption is made that those
    ''''            settings are available and assigned. Will use locally defined variables with hard coded settings
    ''''            for them.
    ''''        This will be updated before get in Starling DDR flow based on the actual
    ''''

    Dim pat_count As Long
    Dim i As Long, k As Long, j As Long
    Dim patt_ary() As String
    Dim site As Variant
    Dim patt As Variant
    Dim Pat As String

    Dim EyeStrobes As Long
    Dim EyeStrobes_DQ As Long     '-- 2018_0920 HH
    Dim EyeStrobes_DQS As Long    '-- 2018_0920 HH

    Dim DqSwpWf As New DSPWave, DqsSwpWf As New DSPWave         ' captured sweeep results as well MDLL cal code if applicable
    Dim DqEyeWf As New DSPWave, DqsEyeWf As New DSPWave         ' for starting, the captured wf are Eye wf, no conversion necessary
    Dim EyeWidthWf As New DSPWave

    Dim Ddr0ImpWf As New DSPWave, Ddr1ImpWf As New DSPWave      ' impedance settings to be sourced in
    Dim DdrImpRegWidthWf As New DSPWave                         ' register bit width
    Dim Ddr0ImpDigSrcWf As New DSPWave
    Dim Ddr1ImpDigSrcWf As New DSPWave
    Dim NoOfSrcBits As Long
    Dim repeats As Long
    Dim isIndDataRepeat As Boolean
    Dim isAllDataRepeat As Boolean
    Dim DigSrc_Sample_Size_DQ As Long
    Dim DigSrc_Sample_Size_DQS As Long
    Dim SplitSize() As String
    Dim Testname_CZ_Vt As String: Testname_CZ_Vt = ""
    Dim Instname_split() As String
    Dim TempStr() As String
    Dim p As Long
    Dim BistIdx As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim DigSrc_Sample_Size_Long As Long

    'speed up the first run test time
    If Validating_ Then
        Call PrLoadPattern(DqsSwpPat.Value)
        Call PrLoadPattern(DqSwpPat.Value)
        Exit Function    ' Exit after validation
    End If

    Dim reg_ddr0_zcpu As New SiteLong, reg_ddr0_zcpd As New SiteLong, reg_ddr0_dspu As New SiteLong, reg_ddr0_dspd As New SiteLong
    Dim reg_ddr1_zcpu As New SiteLong, reg_ddr1_zcpd As New SiteLong, reg_ddr1_dspu As New SiteLong, reg_ddr1_dspd As New SiteLong

    On Error GoTo errHandler

    'Roger New,20180510 TName
    '--------------------------------------------------------------------
    Call GetFlowTName

'    theexec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
'    theexec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = True
'    theexec.Datalog.ApplySetup
    '----------------------------20180523
'''------------------------------------------------------------------------------------------------------------------------
''' DigSrc setup
'''------------------------------------------------------------------------------------------------------------------------
    Dim DqsInDspWave As New DSPWave '''temp
    Dim DqInDspWave As New DSPWave '''temp
    Dim EyeStrobes_bywidth() As String
    Dim EyeStrobes_DQSbywidth() As String
    Dim Cus_bywidth As Boolean



'/////////////////////////////// Customize Bits width//////////////////////////  add 20180925

    If NoOfBists <> 0 Then
    
        DqSwpNoOfBits = CLng(DqSwpNoOfBits)             ' change type to long for same with original
        DqsSwpNoOfBits = CLng(DqsSwpNoOfBits)
        EyeStrobes = DqSwpNoOfBits / NoOfBists
        EyeStrobes_DQ = DqSwpNoOfBits / NoOfBists     '-- 2018_0920 HH
        EyeStrobes_DQS = DqsSwpNoOfBits / NoOfBists '-- 2018_0920 HH
        Cus_bywidth = False

    Else
        EyeStrobes_bywidth = Split(DqSwpNoOfBits, "+")
        EyeStrobes_DQSbywidth = Split(DqsSwpNoOfBits, "+")
        
        Cus_bywidth = True
        
        
    End If
    
 '///////////////////////////////////////////////////////////////////////////////


    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableComparePins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True

    '' 20170222 - Sweep Vt from SweepVtStr
    Dim SplitByColon() As String
    Dim SourceIndexStr As String, SourceIndex As Long
    Dim StartVal As Double, StepVal As Double, FinalVal As Double
    Dim ReplaceStr() As String

    If AMP_EYE_VT_CZ_Flag = True Then
        If SweepVtStr <> "" Then
        SplitByColon = Split(SweepVtStr, ":")
        SourceIndexStr = SplitByColon(0)
        SourceIndex = TheExec.Flow.var(SourceIndexStr).Value
        StartVal = SplitByColon(1)
        StepVal = SplitByColon(2)
        FinalVal = StartVal + SourceIndex * StepVal

            If InStr(UCase(Interpose_PrePat), ":VT:") <> 0 Then
                'NEW 20170731 Purpose to only update VT value and keep the other interpose setting the same
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

                'NEW 20170731 For Char TestName
                FinalVal = Format(FinalVal, "0.000")
                Instname_split = Split(TheExec.DataManager.instanceName, "_")
                If FinalVal < 0 Then
                    Testname_CZ_Vt = Replace(CStr(FinalVal), "-", "m")
                Else
                    Testname_CZ_Vt = CStr(FinalVal)
                End If
                Testname_CZ_Vt = Replace(Testname_CZ_Vt, ".", "p")
                Testname_CZ_Vt = "_" & Instname_split(9) & "_" & Instname_split(10) & "_" & "VT" & "_" & Testname_CZ_Vt & "_" & Instname_split(UBound(Instname_split))

            End If
        End If
    End If

    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If


'    TName_Ary = Split(gl_Tname_Meas, "+")

    TheHdw.Patterns(DqsSwpPat).Load
    gl_TName_Pat = DqsSwpPat.Value
    Call PATT_GetPatListFromPatternSet(DqsSwpPat.Value, patt_ary, pat_count)
    ''''add src for ddr ''''''''''''SP 20180221
    Dim Rtn_SweepTestName As String
    Rtn_SweepTestName = ""
    For Each patt In patt_ary
        If DigSrc_Sample_Size <> "" Then
            Dim DqsSwpPat_Str As String
            DqsSwpPat_Str = CStr(patt)
            DigSrc_Sample_Size_Long = CLng(DigSrc_Sample_Size)
            Call GeneralDigSrcSetting(DqsSwpPat_Str, DigSrc_pin, DigSrc_Sample_Size_Long, DigSrc_DataWidth, DqsDigSrc_Equation, DigSrc_Assignment, _
                                                       DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, DqsInDspWave, Rtn_SweepTestName)
        End If
    Next patt

    '' 1. To get all DigCap bits (Not only bits for eye test) -- TYCHENGG
    If CUS_Str_DigCapData_Dqs <> "" Then
        Dim ShowDec_Dqs As String
        Dim ShowOut_Dqs As String
        Dim DigCapIndex_Dqs As Integer
        Dim DqsDataWf As New DSPWave, DqsTempWf As New DSPWave
        Dim Dqs_DSSC_OUT_Wf(0) As New DSPWave
        Dim Dqs_DSSC_OUT_Full(0) As New DSPWave
        Dim CUS_Sub_Str_DigCapData_Dqs As String
        Dim DqsNewWf As New DSPWave
    End If
    ''----------------------------------------------------

    For Each patt In patt_ary
        Pat = CStr(patt)

        Dim pat_name() As String
        Dim pat_name_module() As String
        Dim Pat_name1() As String

        pat_name_module = Split(Pat, ":")
        pat_name = Split(pat_name_module(0), "\")

        pat_name(0) = pat_name(UBound(pat_name))
        pat_name(0) = Replace(pat_name(0), ".", "_")
        Pat_name1 = Split(TheExec.DataManager.instanceName, "_")

        Call DigCapSetup(Pat, DigCap_Pin, pat_name(0) & "_" & Pat_name1(UBound(Pat_name1)), CLng(DigCap_Sample_Size_Dqs), DqsSwpWf)      'DqsSwpWf = 288

'        Call TheHdw.Patterns(Pat).test(pfAlways, 0)
        If gl_flag_CZ_Nominal_Measured_1st_Point Then: 'Call CZ_TNum_Increment      '20180713 TER add for increaseing FTR TNum @ CZ
        Call TheHdw.Patterns(Pat).Test(pfAlways, 0)
        If gl_flag_CZ_Nominal_Measured_1st_Point Then: 'Call CZ_TNum_Decrement        '20180717 TER add for increaseing FTR TNum @ CZ

        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & pat_count & "): " & Pat & ""
    Next patt

    If TheExec.TesterMode = testModeOffline Then
        For Each site In TheExec.sites
            For i = 0 To DigCap_Sample_Size_Dqs - 1
                DqsSwpWf.Element(i) = FormatNumber(Rnd())
            Next i
        Next site
    End If

    '' 2. To Split DigCapData portion and DqsEyeWf portion -- TYCHENGG
    If CUS_Str_DigCapData_Dqs <> "" Then
        Call DSSC_Special_Str_Filter(CUS_Str_DigCapData_Dqs, EyeTestRegName, DqsSwpWf, _
                                        CUS_Sub_Str_DigCapData_Dqs, DqsTempWf, DqsDataWf)

        ' DqsSwpWf = 288
        ' DqsTempWf = 256
        ' DqsDataWf = 32
        
        
        If Cus_bywidth = False Then                                ' add 20180925
         
            Else
          
                DqsSwpNoOfBits = UBound(EyeStrobes_bywidth) + 1
        
        End If
        
        
        
        For Each site In TheExec.sites

            Dqs_DSSC_OUT_Full(0) = DqsSwpWf
            DqsSwpWf.CreateConstant 0, DqsSwpNoOfBits
            DqsSwpWf = DqsTempWf.Copy          ''
            Dqs_DSSC_OUT_Wf(0) = DqsDataWf

        Next site
        
        ' DqsSwpWf = DqsTempWf = 256
        ' DqsDataWf = Dqs_DSSC_OUT_Wf(0) = 32

        ' Print Out total 288 bits
''        Call HardIP_Digcap_Print_New(CUS_Str_DigCapData_Dqs, Dqs_DSSC_OUT_Full, CLng(DigCap_Sample_Size_Dqs), 0, ShowDec_Dqs, ShowOut_Dqs, , DigCap_DSPWaveSetting)
        Call DigCapDataProcessByDSP(CUS_Str_DigCapData_Dqs, Dqs_DSSC_OUT_Full(0), CLng(DigCap_Sample_Size_Dqs), 0)

    End If
    ''----------------------------------------------------

    For Each site In TheExec.sites
        DqsEyeWf = DqsSwpWf.Copy          '''' the original captured waveform would become stile after DSP functional call
    Next site

    ''''
    '''' DQ sweep
    ''''
    If DqSwpPat <> "" Then

        TheHdw.Patterns(DqSwpPat).Load
        Call PATT_GetPatListFromPatternSet(DqSwpPat.Value, patt_ary, pat_count)
        ''''add src for ddr ''''''''''''SP 20180221
        Rtn_SweepTestName = ""
        For Each patt In patt_ary
            If DigSrc_Sample_Size <> "" Then
                Dim DqSwpPat_Str As String
                DqSwpPat_Str = CStr(patt)
                DigSrc_Sample_Size_Long = CLng(DigSrc_Sample_Size)
                Call GeneralDigSrcSetting(DqSwpPat_Str, DigSrc_pin, DigSrc_Sample_Size_Long, DigSrc_DataWidth, DqDigSrc_Equation, DigSrc_Assignment, _
                                                           DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, DqInDspWave, Rtn_SweepTestName)
            End If
        Next patt

        '' 3. To get all DigCap bits (Not only bits for eye test) -- TYCHENGG
        If CUS_Str_DigCapData_Dq <> "" Then
            Dim ShowDec_Dq As String
            Dim ShowOut_Dq As String
            Dim DigCapIndex_Dq As Integer
            Dim DqDataWf As New DSPWave, DqTempWf As New DSPWave
            Dim Dq_DSSC_OUT_Wf(0) As New DSPWave
            Dim Dq_DSSC_OUT_Full_Wf(0) As New DSPWave
            Dim CUS_Sub_Str_DigCapData_Dq As String
        End If
        ''----------------------------------------------------

        For Each patt In patt_ary
            Pat = CStr(patt)

            Call DigCapSetup(Pat, DigCap_Pin, "Meas_cap", CLng(DigCap_Sample_Size_Dq), DqSwpWf)  ' DqSwpWf = 288

'            Call TheHdw.Patterns(Pat).test(pfAlways, 0)
            If gl_flag_CZ_Nominal_Measured_1st_Point Then: 'Call CZ_TNum_Increment      '20180713 TER add for increaseing FTR TNum @ CZ
            Call TheHdw.Patterns(Pat).Test(pfAlways, 0)
            If gl_flag_CZ_Nominal_Measured_1st_Point Then: 'Call CZ_TNum_Decrement        '20180717 TER add for increaseing FTR TNum @ CZ

            If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & pat_count & "): " & Pat & ""
        Next patt


        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                For i = 0 To DigCap_Sample_Size_Dq - 1
                    DqSwpWf.Element(i) = FormatNumber(Rnd())
                Next i
            Next site
        End If

        '' 4. To Split DigCapData portion and DqsEyeWf portion -- TYCHENGG
        If CUS_Str_DigCapData_Dq <> "" Then
            Call DSSC_Special_Str_Filter(CUS_Str_DigCapData_Dq, EyeTestRegName, DqSwpWf, _
                                         CUS_Sub_Str_DigCapData_Dq, DqTempWf, DqDataWf)

           ' DqSwpWf = 288
           ' CUS_Sub_Str_DigCapData_Dq : Remaining DSSC_OUT String
           ' DqTempWf = 256
           ' DqDataWf = 32
           
        If Cus_bywidth = False Then
         
          Else
          
          DqSwpNoOfBits = UBound(EyeStrobes_bywidth) + 1
        
        End If

            For Each site In TheExec.sites
                Dq_DSSC_OUT_Full_Wf(0) = DqSwpWf
                DqSwpWf.CreateConstant 0, DqSwpNoOfBits
                DqSwpWf = DqTempWf.Copy          ''
                Dq_DSSC_OUT_Wf(0) = DqDataWf
            Next site

'''''            Call HardIP_Digcap_Print_New(CUS_Str_DigCapData_Dq, Dq_DSSC_OUT_Full_Wf, CLng(DigCap_Sample_Size_Dq), 0, ShowDec_Dq, ShowOut_Dq, , DigCap_DSPWaveSetting)   ''
        Call DigCapDataProcessByDSP(CUS_Str_DigCapData_Dq, Dq_DSSC_OUT_Full_Wf(0), CLng(DigCap_Sample_Size_Dq), 0)

        End If

    Else

        'If DqSwpPat no pattern
        For Each site In TheExec.sites
            DqSwpWf.CreateConstant 0, DqsSwpNoOfBits
            Dq_DSSC_OUT_Wf(0) = DqDataWf
        Next site

    End If
    ''----------------------------------------------------

    For Each site In TheExec.sites.Active
        DqEyeWf = DqSwpWf.Copy          '''' the original captured waveform would become stile after DSP functional call
    Next site


    '''' stitch the eye diagrams from 2 sweeps, find and report the eye widths.
    EyeWidthWf.CreateConstant 0, 1
    
    If Cus_bywidth = False Then
    
    Call rundsp.FindMaxEyeWidth_reverse(DqsEyeWf, DqEyeWf, NoOfBists, EyeWidthWf)
    
    Else
    
    Dim Cont_width As New DSPWave  ' add 20180925 for difference width bits
    Dim W As Integer
    
    Cont_width.CreateConstant 0, CLng(DqsSwpNoOfBits), DspLong
    
    For W = 0 To DqsSwpNoOfBits - 1
        
        Cont_width.Element(W) = EyeStrobes_bywidth(W)
        
    Next W
     
     Call rundsp.FindMaxEyeWidth_reverse_bywidth(DqsEyeWf, DqEyeWf, Cont_width, EyeWidthWf) ' add 20180925

    End If
    
   
    
    '                                        256 , 256

    ''''
    '''' Test the eye opening, per lane
    '''' eyewidth: limits from the flow table
    ''''
    
    If NoOfBists = 0 Then
    
       NoOfBists = DqSwpNoOfBits
       
     End If
     

    For BistIdx = 0 To NoOfBists - 1

        If LCase(TheExec.DataManager.instanceName) Like "*cacs*_ck*" Then
            TestNameInput = Report_TName_From_Instance(CalcC, "", "DDR" & CStr(BistIdx) & "_EYE_CACS_CK" & Testname_CZ_Vt, 0, BistIdx)
            TheExec.Flow.TestLimit resultVal:=EyeWidthWf.Element(BistIdx), Tname:=TestNameInput, ForceResults:=tlForceFlow
        Else
            TestNameInput = Report_TName_From_Instance(CalcC, "", "DDR" & CStr(BistIdx \ 2) & "_EYE_DQ_DQS" & CStr(BistIdx Mod 2) & Testname_CZ_Vt, 0, BistIdx)
            TheExec.Flow.TestLimit resultVal:=EyeWidthWf.Element(BistIdx), Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If

    Next BistIdx

    If DisableConnectPins <> "" Then TheHdw.Digital.Pins(DisableComparePins).Connect
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False

    ''''
    '''' display the captured information, as well eye diagrams in the captured order
    ''''
  
    If Cus_bywidth = True Then
       
       DqsSwpNoOfBits = CStr(DigCap_Sample_Size_Dqs)
       DqSwpNoOfBits = CStr(DigCap_Sample_Size_Dq)
       DqsSwpNoOfBits = CLng(DqsSwpNoOfBits)
       DqSwpNoOfBits = CLng(DqSwpNoOfBits)
    End If
    
    
    
    If DispCaptStrm Then

       Dim BitStrM As String
       For Each site In TheExec.sites.Active
         
            '''' Dqs sweep
            BitStrM = CStr(DqsSwpWf(site).Element(0))
            For i = 1 To DqsSwpNoOfBits - 1                          ' DqsSwpNoOfBits = 256
                BitStrM = BitStrM & CStr(DqsSwpWf(site).Element(i))  ''DqSwpWf0=>DqSwpWf
            Next i

            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ": 1st Sweep " & DqsSwpNoOfBits & " bits(LSB->MSB) = " & BitStrM  'cw

            '''' Dq sweep
            BitStrM = CStr(DqSwpWf(site).Element(0))               ''DqsSwpWf0=>DqsSwpWf
            For i = 1 To DqSwpNoOfBits - 1
                BitStrM = BitStrM & CStr(DqSwpWf(site).Element(i)) ''DqsSwpWf0=>DqsSwpWf
            Next i
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "        2nd Sweep " & DqSwpNoOfBits & " bits(LSB->MSB) = " & BitStrM                     'cw

            '''' Swq eyediagram, per eye sweep

            For BistIdx = 0 To NoOfBists - 1
            
                If Cus_bywidth = True Then
                 
                EyeStrobes = EyeStrobes_bywidth(BistIdx)
                
                End If
                
                
'                EyeStrobes_DQ = EyeStrobes_bywidth(BistIdx)
'                EyeStrobes_DQS = EyeStrobes_bywidth(BistIdx)
                
                Dim EyeSt As Integer
                Dim BistByte As Integer, Ddr As Integer

                BistByte = BistIdx Mod 2        '''' BistByte 0 or 1
                Ddr = BistIdx \ 2                       '''' 0, 1: 0; 2, 3: 1
                
                If Cus_bywidth = True Then
                
                If BistIdx = 0 Then
                   EyeSt = 0
                Else
                    EyeSt = EyeSt + EyeStrobes_bywidth(BistIdx - 1)
                End If
'                EyeSt = BistIdx * EyeStrobes
                End If

                BitStrM = CStr(DqsEyeWf(site).Element(EyeSt))
                For i = 1 To EyeStrobes - 1
                    BitStrM = BitStrM & CStr(DqsEyeWf(site).Element(EyeSt + i))
                Next i

                If gl_Disable_HIP_debug_log = False Then

                    If LCase(TheExec.DataManager.instanceName) Like "*cacs*_ck*" Then
                    ''cw add for Skye
                        TheExec.Datalog.WriteComment "         cacs Eye, DDR" & BistIdx & "eye0" & ": " & BitStrM
                    Else
                        TheExec.Datalog.WriteComment "         dq    Eye, DDR" & CInt(BistIdx \ 2) & "eye" & (BistIdx Mod 2) & ": " & BitStrM
                    End If
                End If
                    
                BitStrM = CStr(DqEyeWf(site).Element(EyeSt))
                For i = 1 To EyeStrobes - 1
                    BitStrM = BitStrM & CStr(DqEyeWf(site).Element(EyeSt + i))
                Next i

                If gl_Disable_HIP_debug_log = False Then
                    If LCase(TheExec.DataManager.instanceName) Like "*cacs*_ck*" Then
                    ''cw add for Skye
                        TheExec.Datalog.WriteComment "         ck   Eye, DDR" & BistIdx & "eye0" & ": " & BitStrM
                    Else
                        TheExec.Datalog.WriteComment "         dqs   Eye, DDR" & CInt(BistIdx \ 2) & "eye" & (BistIdx Mod 2) & ": " & BitStrM
                    End If
                End If

            Next BistIdx
        Next site

    End If


    For Each site In TheExec.sites
        DqsSwpWf.CreateConstant 0, DqsSwpNoOfBits
    Next site

    For Each site In TheExec.sites
        DqsSwpWf.CreateConstant 0, DqsSwpNoOfBits
    Next site

    Pat = DqsSwpPat.Value & "," & DqSwpPat.Value
    Shmoo_Pattern = DqsSwpPat.Value & "," & DqSwpPat.Value
    DebugPrintFunc Pat

     '' 20170712 - Process calculate equation by dictionary.
     If Calc_Eqn <> "" Then
         Call ProcessCalcEquation(Calc_Eqn)
     End If

    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Opt_DdrLpBkFunc2"
    If AbortTest Then Exit Function Else Resume Next
  
End Function


Public Function Opt_DdrLpBkFunc3(DqSwpPat As Pattern, DqsSwpPat As Pattern, _
                            DisableComparePins As PinList, DisableConnectPins As PinList, _
                            DigCap_Pin As PinList, NoOfBists As Integer, _
                            DqSwpNoOfBits As Long, DqsSwpNoOfBits As Long, _
                            Optional DispCaptStrm As Boolean = False, _
                            Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As String, _
                            Optional DqDigSrc_Equation As String, Optional DqsDigSrc_Equation As String, _
                            Optional DigSrc_Assignment As String, _
                            Optional CUS_Str_DigSrcData As String, _
                            Optional DigCap_DSPWaveSetting As CalculateMethodSetup_DSPWave = 0, _
                            Optional EyeTestRegName As String, _
                            Optional DigCap_Sample_Size_Dq As Long, _
                            Optional CUS_Str_DigCapData_Dq As String, _
                            Optional DigCap_Sample_Size_Dqs As Long, _
                            Optional CUS_Str_DigCapData_Dqs As String, _
                            Optional Interpose_PrePat As String, _
                            Optional SweepVtStr As String, _
                            Optional Calc_Eqn As String, _
                                                        Optional BV_Enable As Boolean, _
                            Optional Validating_ As Boolean) As Long

    Dim i As Long
    Dim site As Variant
    Dim Pat As String
    Dim EyeStrobes As Long
    Dim DqSwpWf As New DSPWave, DqsSwpWf As New DSPWave
    Dim Testname_CZ_Vt As String: Testname_CZ_Vt = ""
    Dim Instname_split() As String
    Dim TempStr() As String
    Dim p As Long
    Dim BistIdx As Long
    Dim PatCnt As Long
    Dim PatNames() As String
    Dim DSP_Eye_StartBit_DQ As New DSPWave
    Dim DSP_Eye_BitLength_DQ As New DSPWave
    Dim DSP_Eye_StartBit_DQS As New DSPWave
    Dim DSP_Eye_BitLength_DQS As New DSPWave
    Dim DSP_Eye_Width As New DSPWave
    Dim DQ_EYE_Data As New DSPWave
    Dim DQS_EYE_Data As New DSPWave
    Dim testName As String

    ''''' Sweep Vt from SweepVtStr
    Dim SplitByColon() As String
    Dim SourceIndexStr As String, SourceIndex As Long
    Dim StartVal As Double, StepVal As Double, FinalVal As Double
    Dim ReplaceStr() As String
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    
    ''''' Speed up the first run test time
    If Validating_ Then
        Call PrLoadPattern(DqsSwpPat.Value)
        Call PrLoadPattern(DqSwpPat.Value)
        Exit Function    ''''' Exit after validation
    End If
    
    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If
    
    On Error GoTo errHandler
       
    Call GetFlowTName
       
    If BV_Enable Then
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    EyeStrobes = DqSwpNoOfBits / NoOfBists
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableComparePins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    
    ''''' Input Parsing Function to get every EYE start point bit and EYE bit length
    Call Opt_Input_Parsing(CUS_Str_DigCapData_Dq, EyeTestRegName, DSP_Eye_StartBit_DQ, DSP_Eye_BitLength_DQ)
    Call Opt_Input_Parsing(CUS_Str_DigCapData_Dqs, EyeTestRegName, DSP_Eye_StartBit_DQS, DSP_Eye_BitLength_DQS)
    
    If AMP_EYE_VT_CZ_Flag = True Then
        If SweepVtStr <> "" Then
        SplitByColon = Split(SweepVtStr, ":")
        SourceIndexStr = SplitByColon(0)
        SourceIndex = TheExec.Flow.var(SourceIndexStr).Value
        StartVal = SplitByColon(1)
        StepVal = SplitByColon(2)
        FinalVal = StartVal + SourceIndex * StepVal
        
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
               
                ''''' For Char TestName
                FinalVal = Format(FinalVal, "0.000")
                Instname_split = Split(TheExec.DataManager.instanceName, "_")
                If FinalVal < 0 Then
                    Testname_CZ_Vt = Replace(CStr(FinalVal), "-", "m")
                Else
                    Testname_CZ_Vt = CStr(FinalVal)
                End If
                Testname_CZ_Vt = Replace(Testname_CZ_Vt, ".", "p")
                'Testname_CZ_Vt = "_" & Instname_split(10) & "_" & Instname_split(1) & "_" & Instname_split(11) & "_" & "VT" & "_" & Testname_CZ_Vt & "_" & Instname_split(UBound(Instname_split))
                Testname_CZ_Vt = "_" & Instname_split(1) & "_" & Instname_split(9) & "_" & Instname_split(10) & "_" & "VT" & "_" & Testname_CZ_Vt & "_" & Instname_split(UBound(Instname_split)) ' update190925 for eye plot
            End If
        End If
    End If

    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
     
    ''''' Offline Simulation Data
    If TheExec.TesterMode = testModeOffline Then
        DqSwpWf.CreateConstant 0, DigCap_Sample_Size_Dq          ''''' New added to create space for DqSwpWf . Placed TheExec before pattern run for DSP optimization
        DqsSwpWf.CreateConstant 0, DigCap_Sample_Size_Dqs      ''''' New added to create space for DqsSwpWf . Placed TheExec before pattern run for DSP optimization
        
        For Each site In TheExec.sites
            For i = 0 To DigCap_Sample_Size_Dq - 1
               DqSwpWf.Element(i) = Round(Rnd())
            Next i
            
            For i = 0 To DigCap_Sample_Size_Dqs - 1
                DqsSwpWf.Element(i) = Round(Rnd())
            Next i
        Next site
    End If
      
    ''''' Capture Setup and Pattern Run
    PatNames() = TheExec.DataManager.Raw.GetPatternsInSet(DqSwpPat.Value, PatCnt)
    Call DigCapSetup(PatNames(0), DigCap_Pin, "Capture_Code_0", DigCap_Sample_Size_Dq, DqSwpWf)
    Call TheHdw.Patterns(PatNames(0)).Test(pfAlways, 0)
    Call Update_BC_PassFail_Flag

    If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCnt & "): " & PatNames(0) & ""
 
    PatNames() = TheExec.DataManager.Raw.GetPatternsInSet(DqsSwpPat.Value, PatCnt)
    Call DigCapSetup(PatNames(0), DigCap_Pin, "Capture_Code_1", DigCap_Sample_Size_Dqs, DqsSwpWf)
    Call TheHdw.Patterns(PatNames(0)).Test(pfAlways, 0)
    Call Update_BC_PassFail_Flag

    If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCnt & "): " & PatNames(0) & ""
    ''''' End of Capture Setup and Pattern Run

    ''''' DSP DQ/DQS Filter and EYE Width Calculation
    Call rundsp.DSP_Opt_EYE(DqSwpWf, DSP_Eye_StartBit_DQ, DSP_Eye_BitLength_DQ, DqsSwpWf, DSP_Eye_StartBit_DQS, DSP_Eye_BitLength_DQS, NoOfBists, DQ_EYE_Data, DQS_EYE_Data, DSP_Eye_Width)
    
    ''''' DQ/DQS all Capture Code Limits
    Call DigCapDataProcessByDSP(CUS_Str_DigCapData_Dq, DqSwpWf, DigCap_Sample_Size_Dq, 0, , , DigCap_Pin.Value)
    Call DigCapDataProcessByDSP(CUS_Str_DigCapData_Dqs, DqsSwpWf, DigCap_Sample_Size_Dqs, 0, , , DigCap_Pin.Value)

    ''''' EYE Width Limits
    For BistIdx = 0 To NoOfBists - 1
        If LCase(Inst_Name_Str) Like "*cacs*_ck*" Then
        
            If AMP_EYE_VT_CZ_Flag = True Then
                'TestName = Report_TName_From_Instance("calc", "", "EYE_CACS_CK_" & Testname_CZ_Vt & "DDR" & CStr(BistIdx), 0)
                TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), Tname:="DDR" & CStr(BistIdx) & "_EYE_CACS_CK" & Testname_CZ_Vt, ForceResults:=tlForceFlow
            Else
                testName = Report_TName_From_Instance("calc", "", "EYE_CACS_CK_" & "DDR" & CStr(BistIdx), 0)
                TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), Tname:=testName, ForceResults:=tlForceFlow
            End If
            
            'TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), TName:="DDR" & CStr(BistIdx) & "_EYE_CACS_CK" & Testname_CZ_Vt, ForceResults:=tlForceFlow
        Else
            If AMP_EYE_VT_CZ_Flag = True Then
                TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), Tname:="DDR" & CStr(BistIdx \ 2) & "_EYE_DQ_DQS" & CStr(BistIdx Mod 2) & Testname_CZ_Vt, ForceResults:=tlForceFlow
                'TestName = Report_TName_From_Instance("calc", "", "EYE_DQ_DQS_" & CStr(BistIdx Mod 2) & Testname_CZ_Vt & "DDR" & CStr(BistIdx \ 2), 0)
            Else
                testName = Report_TName_From_Instance("calc", "", "EYE_DQ_DQS_" & CStr(BistIdx Mod 2) & "DDR" & CStr(BistIdx \ 2), 0)
                TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), Tname:=testName, ForceResults:=tlForceFlow
            End If
            
            'TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), TName:="DDR" & CStr(BistIdx \ 2) & "_EYE_DQ_DQS" & CStr(BistIdx Mod 2) & Testname_CZ_Vt, ForceResults:=tlForceFlow
        End If
                Call Update_BC_PassFail_Flag
    Next BistIdx
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableComparePins).Connect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
        
    ''''' EYE Printing : Display the captured information, as well eye diagrams in the captured order
    If DispCaptStrm Then

        Dim BitStrM As String
        Dim EyeSt As Integer
        
        For Each site In TheExec.sites.Active

            ''''' DQ_EYE_Data Sweep
            BitStrM = CStr(DQ_EYE_Data(site).Element(0))
            For i = 1 To DqSwpNoOfBits - 1
                BitStrM = BitStrM & CStr(DQ_EYE_Data(site).Element(i))
            Next i
            TheExec.Datalog.WriteComment "Site " & site & ": 1st Sweep " & DqSwpNoOfBits & " bits(LSB->MSB) = " & BitStrM

            ''''' DQS_EYE_Data Sweep
            BitStrM = CStr(DQS_EYE_Data(site).Element(0))
            For i = 1 To DqsSwpNoOfBits - 1
                BitStrM = BitStrM & CStr(DQS_EYE_Data(site).Element(i))
            Next i
            TheExec.Datalog.WriteComment "        2nd Sweep " & DqsSwpNoOfBits & " bits(LSB->MSB) = " & BitStrM

            For BistIdx = 0 To NoOfBists - 1
                EyeSt = BistIdx * EyeStrobes

                BitStrM = CStr(DQ_EYE_Data(site).Element(EyeSt))
                For i = 1 To EyeStrobes - 1
                    BitStrM = BitStrM & CStr(DQ_EYE_Data(site).Element(EyeSt + i))
                Next i

                If LCase(Inst_Name_Str) Like "*cacs*_ck*" Then
                    TheExec.Datalog.WriteComment "         CACS Eye, DDR" & BistIdx & "eye0" & ": " & BitStrM
                Else
                    TheExec.Datalog.WriteComment "         DQ    Eye, DDR" & CInt(BistIdx \ 2) & "eye" & (BistIdx Mod 2) & ": " & BitStrM
                End If

                BitStrM = CStr(DQS_EYE_Data(site).Element(EyeSt))
                For i = 1 To EyeStrobes - 1
                    BitStrM = BitStrM & CStr(DQS_EYE_Data(site).Element(EyeSt + i))
                Next i

                If LCase(Inst_Name_Str) Like "*cacs*_ck*" Then
                    TheExec.Datalog.WriteComment "         CK   Eye, DDR" & BistIdx & "eye0" & ": " & BitStrM
                Else
                    TheExec.Datalog.WriteComment "         DQS   Eye, DDR" & CInt(BistIdx \ 2) & "eye" & (BistIdx Mod 2) & ": " & BitStrM
                End If

            Next BistIdx
        Next site

    End If
    ''''' End of EYE Printing
 
    Pat = DqSwpPat.Value & "," & DqsSwpPat.Value
    Shmoo_Pattern = DqSwpPat.Value & "," & DqsSwpPat.Value
    DebugPrintFunc Pat
    
     ''''' Process calculate equation by dictionary
     If Calc_Eqn <> "" Then
         Call ProcessCalcEquation(Calc_Eqn)
     End If
    
''     Added to print out performance mode power level
     
     
     If Inst_Name_Str Like "DDR_*" Then '''20190509
            Call CUS_DDR_DCS_PrintOut
     End If
     
   
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in Opt_DdrLpBkFunc3"
    If AbortTest Then Exit Function Else Resume Next
  
End Function

Public Function DigSrc_DigCap_Universal_func(Optional patset As Pattern, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
    Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String = "", _
    Optional Interpose_PrePat As String, Optional Interpose_PostTest As String, Optional Validating_ As Boolean) As Long
    
    Dim PatCount As Long, PattArray() As String
   
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen

    Dim InDSPwave As New DSPWave
    Dim OutDspWave() As New DSPWave
    Dim ShowDec As String, ShowOut As String
    Dim site As Variant
    Dim patt As Variant
    Dim Pat As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Dim i As Long, j As Long, k As Long

''    Dim RTN_InterposeString As String
    On Error GoTo errHandler
    
    ''20141219 Get use-limit from flow table
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    '' 20160923 - Add Interpose_PrePat entry point
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    Call HardIP_InitialSetupForPatgen
    gl_TName_Pat = patset.Value
    TheHdw.Patterns(patset).Load
    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
    
    ReDim OutDspWave(PatCount - 1) As New DSPWave
    
   For i = 0 To PatCount - 1
        Pat = CStr(PattArray(i))
        
        TheHdw.Patterns(Pat).Load

        Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, _
                                               DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPwave)
        
        
        Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size / PatCount, OutDspWave(i))
        
        'Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    
    
        Call TheHdw.Patterns(Pat).Test(pfAlways, 0)

        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & ""
        
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
    
        Call SetForceCondition(Interpose_PostTest)
    Next i
        
        Dim OutDspWave_final As New DSPWave
        OutDspWave_final.CreateConstant 0, DigCap_Sample_Size, DspLong
        For Each site In TheExec.sites.Active
            For i = 0 To PatCount - 1
                For j = 0 To 11
                    OutDspWave_final.Element(i * 12 + j) = OutDspWave(i).Element(j)
                Next j
            Next i
        Next site
        '' 20160211 - Process DigCapData by using DSP
        If DigCap_Sample_Size <> 0 Then
            Dim DigCapPinAry() As String, NumberPins As Long
            Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
            
            If NumberPins > 1 Then
                Call CreateSimulateDataDSPWave_Parallel(OutDspWave_final, DigCap_Sample_Size)
                Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave_final, NumberPins)
                Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave_final, DigCap_Sample_Size, NumberPins)
            ElseIf NumberPins = 1 Then
                Call CreateSimulateDataDSPWave(OutDspWave_final, DigCap_Sample_Size, DigCap_DataWidth)
                Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave_final, NumberPins)
                Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave_final, DigCap_Sample_Size, DigCap_DataWidth)
            End If
        End If

    
    DebugPrintFunc patset.Value  ' print all debug information
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in DigSrc_DigCap_Universal_func"
    If AbortTest Then Exit Function Else Resume Next
  
End Function

Public Function TMPS_Voltage_Print(PowerName As String) As Long
If gl_Disable_HIP_debug_log = False Then
    TheExec.Datalog.WriteComment "*******************************"
    TheExec.Datalog.WriteComment "Set " & PowerName & " : " & TheHdw.DCVS.Pins(PowerName).Voltage.Value
    TheExec.Datalog.WriteComment "*******************************"
End If

End Function



Public Function ReMeasImpedByAveTrimCode(Optional patset As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasR_Pins_SingleEnd As String, Optional MeasR_Pins_Differential As String, Optional StrForceVolt As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimStoreName As String, _
    Optional Fixed_DigSrc_DataWidth As Long, Optional Fixed_DigSrc_Sample_Size As Long, Optional Fixed_DigSrc_Equation As String, Optional Fixed_DigSrc_Assignment As String, _
    Optional b_PD_Mode As Boolean = True) As Long
    
    Dim PatCount As Long, PattArray() As String
    Dim InDSPwave As New DSPWave
    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long
    Dim MeasureImped As New PinListData

    Dim PatWithPinsInfo(3) As MeasTrimImpedInfo
    Dim Pin_Ary() As String, Pin_Cnt As Long, Pin As Variant, TempPins As String
    Dim b_IsDifferential As Boolean
    Dim SingleEndSplitByAdd() As String, DifferentialSplitByAdd() As String, PatternSplitByAdd() As String
    Dim InfoCounter As Long
    
    On Error GoTo ErrorHandler
    
    Call HardIP_InitialSetupForPatgen
    
    SingleEndSplitByAdd = Split(MeasR_Pins_SingleEnd, "+")
    DifferentialSplitByAdd = Split(MeasR_Pins_Differential, "+")
    PatternSplitByAdd = Split(patset, "+")
    
    For InfoCounter = 0 To UBound(PatWithPinsInfo)
        If MeasR_Pins_SingleEnd <> "" Then
            TheExec.DataManager.DecomposePinList SingleEndSplitByAdd(InfoCounter), Pin_Ary, Pin_Cnt
            b_IsDifferential = False
            
        ElseIf MeasR_Pins_Differential <> "" Then
            TheExec.DataManager.DecomposePinList DifferentialSplitByAdd(InfoCounter), Pin_Ary, Pin_Cnt
            
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
        
        TheHdw.Patterns(PatternSplitByAdd(InfoCounter)).Load
        Call PATT_GetPatListFromPatternSet(PatternSplitByAdd(InfoCounter), PattArray, PatCount)
        PatWithPinsInfo(InfoCounter).Pat = PattArray(0)
        PatWithPinsInfo(InfoCounter).MeasPinsAry = Pin_Ary
        PatWithPinsInfo(InfoCounter).IsDifferential = b_IsDifferential
    Next InfoCounter

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
    
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    
    Dim InitialFixedDSPWave As New DSPWave
    Dim FinalTrimDSPWave_BIN As New DSPWave
    Dim FinalTrimDSPWave_DEC As New DSPWave
    Dim Trim_DigSrc_Sample_Size As Long

    Trim_DigSrc_Sample_Size = DigSrc_Sample_Size - Fixed_DigSrc_Sample_Size
    
    FinalTrimDSPWave_DEC.CreateConstant 0, 1, DspLong
    FinalTrimDSPWave_BIN.CreateConstant 0, Trim_DigSrc_Sample_Size, DspLong
    
    If Fixed_DigSrc_Equation <> "" Then
        For Each site In TheExec.sites.Active
            Call Create_DigSrc_Data(DigSrc_pin, Fixed_DigSrc_DataWidth, Fixed_DigSrc_Sample_Size, Fixed_DigSrc_Equation, Fixed_DigSrc_Assignment, InitialFixedDSPWave, site)
        Next site
        
        If (TheExec.TesterMode = testModeOffline) Then
            FinalTrimDSPWave_DEC.CreateConstant 18, 1, DspLong
        Else
            FinalTrimDSPWave_DEC = GetStoredCaptureData(TrimStoreName)
        End If
''        FinalTrimDSPWave_DEC = FinalTrimDSPWave_DEC.ConvertDataTypeTo(DspLong)
''        FinalTrimDSPWave_BIN = FinalTrimDSPWave_BIN.ConvertDataTypeTo(DspLong)
        Call rundsp.DSPWaveDecToBinary(FinalTrimDSPWave_DEC, Trim_DigSrc_Sample_Size, FinalTrimDSPWave_BIN)
       
        Call rundsp.CombineDSPWave(InitialFixedDSPWave, FinalTrimDSPWave_BIN, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, InDSPwave)
    End If
    Dim OutputTrimCode As String
    For Each site In TheExec.sites
        OutputTrimCode = ""
        For k = 0 To InDSPwave(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(InDSPwave(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
    Next site

    For InfoCounter = 0 To UBound(PatWithPinsInfo)
        For Each Pin In PatWithPinsInfo(InfoCounter).MeasPinsAry
            Call SetupDigSrcDspWave(PatWithPinsInfo(InfoCounter).Pat, DigSrc_pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPwave)
    
            Call TheHdw.Patterns(PatWithPinsInfo(InfoCounter).Pat).start
            Call SubMeasR(CPUA_Flag_In_Pat, CStr(Pin), ForceVolt, MeasureImped, PatWithPinsInfo(InfoCounter).IsDifferential, b_PD_Mode)
            TheExec.Flow.TestLimit resultVal:=MeasureImped, Unit:=unitCustom, customUnit:="ohm", Tname:="SourceAverCode" & "_Pin_" & Pin, ForceResults:=tlForceFlow
        Next Pin
    Next InfoCounter
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in ReMeasImpedByAveTrimCode function"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TMPS_Bin2Dec(ByRef DataOut_85C As DSPWave, Optional DSPWave_Dict As DSPWave) As Long

Dim i As Integer
Dim Data_Temp As String
Dim site As Variant
    For Each site In TheExec.sites
        For i = 0 To (DSPWave_Dict(site).SampleSize - 1)
            Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(i))
        Next i
            DataOut_85C(site).Element(0) = Bin2Dec_rev(Data_Temp)
            Data_Temp = ""
    Next site

End Function

Public Function TMPS_Dec2Bin(ByRef Read_Code As DSPWave, Optional DSPWave_Dict As DSPWave, Optional dspwavesize) As Long
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
Public Function Eye_Diagram(LaneNumber As Long) As Long
Dim i, j, k As Long
Dim Eye_Diagram_Binary_Lane_Temp(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane() As EyeDiagram
ReDim Eye_Diagram_Binary_Lane(LaneNumber - 1) As EyeDiagram
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim First_src_code_Temp As New SiteVariant
Dim End_src_code_Temp As New SiteVariant
Dim WithinEye As Boolean
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
Dim X As Integer
Dim iLen As Integer
Dim Temp_Counter_Act As Long
Dim Total_Zero_Count As Long

For k = 0 To LaneNumber - 1
    For Each site In TheExec.sites
        For i = -31 To 31
        Eye_Diagram_Binary_Lane_Temp(i + 31)(site) = ""
            For j = 1 To 32
                Eye_Diagram_Binary_Lane_Temp(i + 31)(site) = Eye_Diagram_Binary_Lane_Temp(i + 31)(site) & Mid(Eye_Diagram_Binary(i + 31)(site), LaneNumber * j - (LaneNumber - (k + 1)), 1)
            Next j
'        If k = LaneNumber - 1 Then: Eye_Diagram_Binary_Lane_Temp(i + 31)(Site) = ""
        Next i
    Next site
    Eye_Diagram_Binary_Lane(k).Value = Eye_Diagram_Binary_Lane_Temp
Next k
For k = 0 To LaneNumber - 1
    vertical_width = 0
    First_src_code = 0
    End_src_code = 0
    horizontal_width = 0
    timing_res_start = 0
    timing_res_end = 0
    Zero_counter = 0
    
    For Each site In TheExec.sites
        iLen = Len(Eye_Diagram_Binary_Lane(k).Value(0)(site)) - 1
        'process   the  Max Zero horizontal
        For i = -31 To 31
            Temp_counter = 0
            Total_Zero_Count = 0
            Temp_Counter_Act = 0
            timing_res_start_temp = 0
            timing_res_end_temp = 0
            For X = 0 To iLen
                If Mid(Eye_Diagram_Binary_Lane(k).Value(i + 31)(site), iLen - X + 1, 1) = 0 Then
                    Temp_counter = Temp_counter + 1
                    Total_Zero_Count = Total_Zero_Count + 1
                    If X = iLen And Temp_counter > Temp_Counter_Act Then: Temp_Counter_Act = Temp_counter
                ElseIf Mid(Eye_Diagram_Binary_Lane(k).Value(i + 31)(site), iLen - X + 1, 1) = 1 Then
                    If Temp_counter > Temp_Counter_Act Then
                        Temp_Counter_Act = Temp_counter
                        timing_res_end_temp = 32 - (X - Temp_Counter_Act + 1)
                        timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                    End If
                    Temp_counter = 0
                End If
            Next X
                If horizontal_width < Temp_Counter_Act Then
                   horizontal_width = Temp_Counter_Act
                   timing_res_end = timing_res_end_temp
                   timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then: Zero_counter = Total_Zero_Count
        Next i
'============================= Vertical Width Process Start (In Site Loop) ================================' ''20180918 -- TYCHENGG
        For X = 0 To iLen
            vertical_width = 0
            WithinEye = False
            For i = -31 To 31 '' -31,31 can be defined as constant
                If WithinEye = False Then
                    If Mid(Eye_Diagram_Binary_Lane(k).Value(i + 31)(site), X + 1, 1) = 0 Then
                        WithinEye = True
                        First_src_code_Temp = i
                    End If
                Else
                    If Mid(Eye_Diagram_Binary_Lane(k).Value(i + 31)(site), X + 1, 1) = 1 Then
                        WithinEye = False
                        End_src_code_Temp = i - 1
                        vertical_width = Abs(End_src_code_Temp - First_src_code_Temp)
                        If vertical_width > Abs(End_src_code - First_src_code) Then
                            First_src_code = First_src_code_Temp
                            End_src_code = End_src_code_Temp
                        End If
                    End If
                End If
            Next i
            vertical_width = Abs(End_src_code_Temp - First_src_code_Temp)
            If vertical_width > Abs(End_src_code - First_src_code) Then
                First_src_code = First_src_code_Temp
                End_src_code = End_src_code_Temp
            End If
        Next X
        vertical_width = Abs(End_src_code - First_src_code) + 1
'============================= Vertical Width Process End   (In Site Loop) ================================' ''20180918 -- TYCHENGG
'//////////////////////// for all 1 eye by csho/////////////////
        If vertical_width = "" Then: vertical_width = 0
        If First_src_code = "" Then: First_src_code = 0
        If End_src_code = "" Then: End_src_code = 0
        If horizontal_width = "" Then: horizontal_width = 0
        If timing_res_end = "" Then: timing_res_end = 0
        If timing_res_start = "" Then: timing_res_start = 0
        If Zero_counter = "" Then: Zero_counter = 0
'/////////////////////////////////////////////////////////////////////////
    Next site
    If LaneNumber = 1 Then
        TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane"
        TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane"
        TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane"
        TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane"
        TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane"
        TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane"
        TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane"
    Else
        TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane" & k
        TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane" & k
        TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane" & k
        TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane" & k
        TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane" & k
        TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane" & k
        TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane" & k
    End If
Next k
For k = 0 To LaneNumber - 1
    For Each site In TheExec.sites
        For i = -31 To 31
            If LaneNumber = 1 Then
                If i <= -10 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).Value(i + 31)(site) & ", h0dac_off:" & i & ",, " & TheExec.DataManager.instanceName & "]")
                ElseIf -10 < i And i < 0 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).Value(i + 31)(site) & ", h0dac_off: " & i & ",, " & TheExec.DataManager.instanceName & "]")
                ElseIf 0 <= i And i < 10 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).Value(i + 31)(site) & ", h0dac_off:  " & i & ",, " & TheExec.DataManager.instanceName & "]")
                ElseIf i >= 10 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).Value(i + 31)(site) & ", h0dac_off: " & i & ",, " & TheExec.DataManager.instanceName & "]")
                End If
            Else
                If i <= -10 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).Value(i + 31)(site) & ", h0dac_off:" & i & ", Lane" & k & ", " & TheExec.DataManager.instanceName & "]")
                ElseIf -10 < i And i < 0 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).Value(i + 31)(site) & ", h0dac_off: " & i & ", Lane" & k & ", " & TheExec.DataManager.instanceName & "]")
                ElseIf 0 <= i And i < 10 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).Value(i + 31)(site) & ", h0dac_off:  " & i & ", Lane" & k & ", " & TheExec.DataManager.instanceName & "]")
                ElseIf i >= 10 Then
                    Call TheExec.Datalog.WriteComment("[Eye Diagram, " & HramLotId(site) & "-" & CStr(HramWaferId(site)) & ", " & CStr(XCoord(site)) & ", " & CStr(YCoord(site)) & ", " & site & ", " & Eye_Diagram_Binary_Lane(k).Value(i + 31)(site) & ", h0dac_off: " & i & ", Lane" & k & ", " & TheExec.DataManager.instanceName & "]")
                End If
            End If
        Next i
    Next site
Next k

End Function

Public Function PPMU_Impedance_Function(PowerPins As String, Power_Force_V As Double, Sink_Groups As String, Meas_Groups As String, Sink1_Current As Double, Sink2_Current As Double, LowLimit As Double, HiLimit As Double) As Long
  
    Dim Sink_Groups_Array() As String
    Dim Sink_Groups_Num As Double
    Dim Meas_Groups_Array() As String
    Dim Meas_Groups_Num As Double
    Dim ResultPower1 As New PinListData
    Dim ResultPower2 As New PinListData

    Dim site As Variant
    Dim Calculate_Contact_R As New PinListData

    On Error GoTo errHandler
    
    
    Sink_Groups_Array = Split(Sink_Groups, ",")
    Sink_Groups_Num = UBound(Sink_Groups_Array)
    
    Meas_Groups_Array = Split(Meas_Groups, ",")
    Meas_Groups_Num = UBound(Meas_Groups_Array)
    
    If Sink_Groups_Num <> Meas_Groups_Num Then
        TheExec.Datalog.WriteComment "None Match Pin Num"
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True
        'DisconnectVDDCA 'SEC DRAM
        TheHdw.Wait 0.001
        
        TheHdw.DCVS.Pins(PowerPins).Voltage.Main = Power_Force_V 'force_v pclinzg
        TheHdw.Wait 0.001
        TheHdw.Digital.Pins(Sink_Groups).Disconnect
        TheHdw.Digital.Pins(Meas_Groups).Disconnect
        TheHdw.PPMU.Pins(Sink_Groups).Gate = tlOn
        TheHdw.PPMU.Pins(Meas_Groups).Gate = tlOn
        TheHdw.PPMU.Pins(Sink_Groups).Connect
        TheHdw.PPMU.Pins(Meas_Groups).Connect
        TheHdw.PPMU.Pins(Sink_Groups).ClampVHi = 1
        TheHdw.PPMU.Pins(Sink_Groups).ClampVLo = -1
        TheHdw.PPMU.Pins(Meas_Groups).ForceI 0, 0.00002
    '========================Sink1======================
        TheHdw.PPMU.Pins(Sink_Groups).ForceV 0, 0.00002
        TheHdw.PPMU.Pins(Meas_Groups).ForceV 0, 0.00002
        TheHdw.Wait 0.0005

        TheHdw.PPMU.Pins(Sink_Groups).ForceI Sink1_Current, Sink1_Current
        TheHdw.PPMU.Pins(Meas_Groups).ForceI 0, 0.00002
        TheHdw.Wait 0.0001
        ResultPower1 = TheHdw.PPMU(Meas_Groups).Read(tlPPMUReadMeasurements)
        TheExec.Flow.TestLimit resultVal:=ResultPower1, hiVal:=0.5, Unit:=unitVolt, ForceVal:=Sink1_Current, ForceUnit:=unitAmp, Tname:="ForceI1"
        
    '========================Sink2======================
        TheHdw.PPMU.Pins(Sink_Groups).ForceI Sink2_Current, Sink2_Current
        TheHdw.Wait 0.0005
        ResultPower2 = TheHdw.PPMU(Meas_Groups).Read(tlPPMUReadMeasurements)
        TheExec.Flow.TestLimit resultVal:=ResultPower2, hiVal:=0.5, Unit:=unitVolt, ForceVal:=Sink2_Current, ForceUnit:=unitAmp, Tname:="ForceI2"
        TheHdw.PPMU.Pins(Meas_Groups).Disconnect
        TheHdw.PPMU.Pins(Sink_Groups).Disconnect
'
        TheHdw.PPMU.Pins(Sink_Groups).Gate = tlOff
        TheHdw.PPMU.Pins(Meas_Groups).Gate = tlOff
        TheHdw.DCVS.Pins(PowerPins).Voltage.Main = 0
        TheHdw.Wait 0.001
    
          
        Calculate_Contact_R = ResultPower1.Math.Subtract(ResultPower2)
        Calculate_Contact_R = Calculate_Contact_R.Math.Divide(Sink1_Current - Sink2_Current).Abs
        
        TheExec.Flow.TestLimit resultVal:=Calculate_Contact_R, lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitCustom, customUnit:="Ohm", Tname:="Resistance"
        DebugPrintFunc ""
        
    End If
    Exit Function
errHandler:
        TheExec.AddOutput "Error in PPMU_Impedance"
        If AbortTest Then Exit Function Else Resume Next
End Function
Public Function PPMU_Impedance(PowerPins As String, Power_Force_V As Double, Sink_Groups As String, Meas_Groups As String, Sink1_Current As Double, Sink2_Current As Double, LowLimit As Double, HiLimit As Double) As Long
    
  
    Dim Sink_Groups_Array() As String
    Dim Sink_Groups_Num As Double
    Dim Meas_Groups_Array() As String
    Dim Meas_Groups_Num As Double
    Dim ResultPower1 As New PinListData
    Dim ResultPower2 As New PinListData
    Dim ResultPower3 As New PinListData
    
    Dim site As Variant
    Dim Calculate_Contact_R As New PinListData
    Dim Calculate_path_R As New PinListData
    Dim Calculate_trace_R As New PinListData
    Dim Calculate_path_V As New PinListData
    Dim Table_trace_R As New PinListData
    Dim R_AK_sink() As Double
    Dim RAK_PPMU As New PinListData
  


    On Error GoTo errHandler
    
    
    Sink_Groups_Array = Split(Sink_Groups, ",")
    Sink_Groups_Num = UBound(Sink_Groups_Array)
    
    Meas_Groups_Array = Split(Meas_Groups, ",")
    Meas_Groups_Num = UBound(Meas_Groups_Array)
    
    If Sink_Groups_Num <> Meas_Groups_Num Then
        TheExec.Datalog.WriteComment "None Match Pin Num"
    Else
        'thehdw.Digital.ApplyLevelsTiming True, True, True
        'DisconnectVDDCA 'SEC DRAM
        'thehdw.Wait 0.001
        
        ' no power request 170630 WC
        'thehdw.DCVS.Pins(PowerPins).Voltage.Main = Power_Force_V 'force_v pclinzg
        'thehdw.Wait 0.005
        TheHdw.Digital.Pins(Sink_Groups).Disconnect
        TheHdw.Digital.Pins(Meas_Groups).Disconnect
        TheHdw.PPMU.Pins(Sink_Groups).Connect
        TheHdw.PPMU.Pins(Meas_Groups).Connect
        
    '========================Sink1======================
        TheHdw.PPMU.Pins(Sink_Groups).ForceI Sink1_Current, Sink1_Current
        TheHdw.PPMU.Pins(Meas_Groups).ForceI 0, 0.00002
        
        TheHdw.Wait 0.005
        ResultPower1 = TheHdw.PPMU(Meas_Groups).Read(tlPPMUReadMeasurements)
        ResultPower3 = TheHdw.PPMU(Sink_Groups).Read(tlPPMUReadMeasurements)
      
        
    '========================Sink2======================
        TheHdw.PPMU.Pins(Sink_Groups).ForceI Sink2_Current, Sink2_Current
        TheHdw.PPMU.Pins(Meas_Groups).ForceI 0, 0.00002
        
        TheHdw.Wait 0.005
        ResultPower2 = TheHdw.PPMU(Meas_Groups).Read(tlPPMUReadMeasurements)
        
        Dim pins_sink() As String
        Dim sink_num As Long
        Dim Meas_sink() As String
        Dim Meas_num As Long
        
        TheExec.DataManager.DecomposePinList Sink_Groups, pins_sink, sink_num
        TheExec.DataManager.DecomposePinList Meas_Groups, Meas_sink, Meas_num
        
        For Each site In TheExec.sites.Active
        
        'R_AK_sink = TheHdw.PPMU.ReadRakValuesByPinnames(Sink_Groups, site)

        
        ' no power request 170630 WC
        'thehdw.DCVS.Pins(PowerPins).Voltage.Main = 0
        'thehdw.Wait 0.001
    
          
        Calculate_Contact_R = ResultPower1.Math.Subtract(ResultPower2)
        Calculate_path_V = ResultPower3.Math.Subtract(Calculate_Contact_R)
        Calculate_Contact_R = Calculate_Contact_R.Math.Divide(Sink1_Current - Sink2_Current).Abs
        Calculate_path_R = Calculate_path_V.Math.Divide(Sink1_Current).Abs
        Calculate_trace_R = Calculate_path_R
        Table_trace_R = Calculate_trace_R
        RAK_PPMU = Calculate_trace_R
        
        Dim i As Integer
        For i = 0 To sink_num - 1
        Calculate_trace_R.Pins(pins_sink(i)).Value = Calculate_path_R.Pins(pins_sink(i)).Value - Calculate_Contact_R.Pins(Meas_sink(i)).Value '- R_AK_sink(i) Merge into Trace R
        'RAK_PPMU.Pins(pins_sink(i)).Value(site) = R_AK_sink(i)
        Table_trace_R.Pins(pins_sink(i)).Value = CurrentJob_Card_RAK.Pins(pins_sink(i)).Value
        Next i
        
        Next site
        
        TheExec.Flow.TestLimit resultVal:=Calculate_Contact_R, lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitCustom, Tname:="Contact_R", customUnit:="Ohm", ForceResults:=tlForceNone
        TheExec.Flow.TestLimit resultVal:=Calculate_path_R, lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitCustom, Tname:="Path_R", customUnit:="Ohm", ForceResults:=tlForceNone
        TheExec.Flow.TestLimit resultVal:=Calculate_trace_R, lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitCustom, Tname:="Trace_R_Calculate", customUnit:="Ohm", ForceResults:=tlForceNone
        TheExec.Flow.TestLimit resultVal:=Table_trace_R, lowVal:=LowLimit, hiVal:=HiLimit, Unit:=unitCustom, Tname:="Trace_R_Table", customUnit:="Ohm", ForceResults:=tlForceNone
        'TheExec.Flow.TestLimit resultVal:=RAK_PPMU, lowval:=LowLimit, hival:=HiLimit, Unit:=unitCustom, Tname:="RAK", customUnit:="Ohm", ForceResults:=tlForceNone
        
        TheHdw.PPMU.Pins(Sink_Groups).Disconnect
        TheHdw.PPMU.Pins(Meas_Groups).Disconnect

         
        DebugPrintFunc ""
        
    End If
    Exit Function
errHandler:
        TheExec.AddOutput "Error in PPMU_Impedance"
        If AbortTest Then Exit Function Else Resume Next
End Function
Public Function TrimCodeDig(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, Optional CUS_Str_DigCapData As String, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PrePat As String, Optional DigCap_Pin As PinList, Optional DigCap_Sample_Size As Long, _
    Optional Validating_ As Boolean, Optional Interpose_PostTest As String) As Long
    
    Dim PatCount As Long, PattArray() As String
    Dim OutDspWave As New DSPWave
               TrimTarget = 0.3 ''Just for Flag    2017/Dec
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If
    Call HardIP_InitialSetupForPatgen
    Dim Ts As Variant, TestSequenceArray() As String
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPwave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    
    Dim CapValue As New PinListData, CapValue_V1 As New PinListData, CapValue_V2 As New PinListData
    CapValue.AddPin ("CapValueString")
    On Error GoTo ErrorHandler

    Call GetFlowTName
 
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Interpose_PrePat <> "" Then ''''180109 update
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TheHdw.Patterns(patset).Load

    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)

    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    Dim b_HigherThanTarget As New SiteBoolean
    b_HigherThanTarget = False
    
    Dim LastSectionV1V2_Index As Long
    LastSectionV1V2_Index = 0
    Dim OutputTrimCode As String, TestNameInput As String
    Dim SourceTrimCode As String
    Dim TestLimitIndex As Long '
    
'    Dim TrimStart_1st() As String
    Dim Dec_TrimStart_1st As Long
    
    Dim StoredTargetTrimCode As New DSPWave
    Dim b_MatchTagetCap As New SiteBoolean
    Dim b_DisplayCap As New SiteBoolean
    b_MatchTagetCap = False
    b_DisplayCap = False
    
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    Dim StoreEachTrimCode() As New DSPWave
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    
    Dim DigCapData() As New PinListData
    ReDim DigCapData(DigSrc_Sample_Size + 1) As New PinListData
    
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i

    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", "")
    End If

'' ===============================   TrimStart_1st = TrimStart===============================================
    Dec_TrimStart_1st = Bin2Dec(TrimStart)
    
    InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong

    Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, DigSrc_Sample_Size, InDSPwave)

    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, InDSPwave)
    TheExec.Datalog.WriteComment ("========First Time Setup========")
    
    For Each site In TheExec.sites
        SourceTrimCode = ""
        For k = 0 To InDSPwave(site).SampleSize - 1
            SourceTrimCode = SourceTrimCode & CStr(InDSPwave(site).Element(k))
        Next k
        TheExec.Datalog.WriteComment ("Site_" & site & " Initial Source Trim Code = " & SourceTrimCode)
    Next site
    
    For Each site In TheExec.sites
        StoreEachTrimCode(0)(site).Data = InDSPwave(site).Data
    Next site
    
'=========Set Up DigCap parameter=============================
    Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Call TheHdw.Patterns(PattArray(0)).start
     'For Each Site In TheExec.sites
        CapValue.Value = OutDspWave.Element(0)
     'Next Site
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    ''Update Interpose_PreMeas 20170801
    Dim TestSeqNum As Integer
    TestSeqNum = 0
           
    TheHdw.Digital.Patgen.HaltWait
'    For Each Site In TheExec.sites
        DigCapData(0) = CapValue
        b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
        PastDSPWave = InDSPwave
        TestNameInput = "Cap_Value_"
        TestLimitIndex = 0
'    Next Site

    Dim CUS_Str_MainProgram As String: CUS_Str_MainProgram = ""

    
    For Each site In TheExec.sites
        If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
            
            TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
            
        End If
    Next site
    
    b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
    b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
    
    For Each site In TheExec.sites
        If b_DisplayCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
            StoredTargetTrimCode(site).Data = InDSPwave(site).Data
            b_StopTrimCodeProcess(site) = True
        End If
    Next site
    
    TheExec.Datalog.WriteComment ("======================================================================================")
    
    Dim b_KeepGoing As New SiteBoolean
    
    Dim PreviousCap As New PinListData
   
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    StoreEachIndex = 1
    
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
        
        For j = SrcStartBit To (SrcEndBit + 1) Step -1
        ''===============up for src bit step

'            For Each Site In TheExec.sites
'                CapValue = CStr(OutDspWave(Site).Element(0))
'            Next Site
            
            If b_FirstExecution = True Then
                b_ControlNextBit = True
                If j = SrcEndBit Then ''=============Trim from Format untill same======1124
                    b_ControlNextBit = False
                End If
            Else
            
            ''20160716-Control next bit to 1 no matter first or last progress
                b_ControlNextBit = True
                If j = SrcEndBit Then
                    b_ControlNextBit = False
                End If
            End If

            If b_FirstExecution = True And j = SrcEndBit Then
                Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPwave)
            Else
                Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HigherThanTarget, j, b_ControlNextBit, InDSPwave)
            End If
            
            Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, InDSPwave)
            
                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site).Data = InDSPwave(site).Data
                Next site
            
            '' Debug use
            '' ==============================================================================================
            '' 20160716 - Modify trim code rule
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
                
                For Each site In TheExec.sites
                    If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
                    If b_KeepGoing(site) = True Then
                        SourceTrimCode = ""
                        For k = 0 To InDSPwave(site).SampleSize - 1
                            SourceTrimCode = SourceTrimCode & CStr(InDSPwave(site).Element(k))
                        Next k
                        Dim OutputDec As String
                        TheExec.Datalog.WriteComment ("Site_" & site & " Source Trim Code = " & SourceTrimCode)
                    End If
                    End If
                Next site
            Set OutDspWave = Nothing
            Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
            Call TheHdw.Patterns(PattArray(0)).start
            
            ''Update Interpose_PreMeas 20170801
            TestSeqNum = 0
                            
            
            TheHdw.Digital.Patgen.HaltWait
            
                        CapValue.Value = OutDspWave.Element(0)
            b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                PastDSPWave = InDSPwave

            TestLimitIndex = TestLimitIndex + 1
            
            '' 20160712 - Modify to use WriteComment to display output frequency.
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
'                    TheExec.Datalog.WriteComment ("Site " & Site & " Output CapValue = " & CapValue.Pins(Site).Value(Site))
                    TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
                End If
            Next site
            
            b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
            b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_MatchTagetCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
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

            TheExec.Datalog.WriteComment ("======================================================================================")
        Next j
        For Each site In TheExec.sites
            If CapValue = 1 Then
                SourceTrimCode = ""
                    SourceTrimCode = SourceTrimCode & "0"
                For k = 1 To InDSPwave(site).SampleSize - 1
                    SourceTrimCode = SourceTrimCode & CStr(InDSPwave(site).Element(k))
                Next k
                InDSPwave(site).Element(0) = 0
            Else
                SourceTrimCode = ""
                    
                For k = 0 To InDSPwave(site).SampleSize - 1
                    SourceTrimCode = SourceTrimCode & CStr(InDSPwave(site).Element(k))
                Next k
            
            End If
           TheExec.Datalog.WriteComment ("Site_" & site & " Final Source Trim Code = " & SourceTrimCode)
        Next site
    
    Next i
        
    End If
    
    If TrimStoreName <> "" Then
       Call Checker_StoreDigCapAllToDictionary(TrimStoreName, InDSPwave)
    End If
    
    Call HardIP_WriteFuncResult(, , Inst_Name_Str)

    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel InDSPwave, DigSrc_Sample_Size, ConvertedDataWf
    
    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, StoredTargetTrimCode)
    Call TheHdw.Patterns(PattArray(0)).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
   
    TheHdw.Digital.Patgen.HaltWait
    
    If Interpose_PrePat <> "" Then '''180109 update
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    Call SetForceCondition(Interpose_PostTest)
    
    Dim sl_FUSE_Val As New SiteLong

    DebugPrintFunc patset.Value
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeCap function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function

Public Function TrimCodeDig_SeaHawk(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, Optional CUS_Str_DigCapData As String, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PrePat As String, Optional DigCap_Pin As PinList, Optional DigCap_Sample_Size As Long, _
    Optional Validating_ As Boolean, Optional Interpose_PostTest As String, Optional TrimOffset As String, Optional TrimBase As String, Optional DigSrc_Sample_Size_Real As String, Optional DigCap_DataWidth As Long, Optional Digcap_key As String) As Long

    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
    TheHdw.DSSC.MoveMode = tlDSSCMoveModeDatabus
    
    
    Dim SourceTrimCode_SiteVariant As New SiteVariant
    
    Dim SourceTrimCode_temp As Long
    Dim SourceTrimCode_final As Long
    Dim TrimBase_temp() As String
    Dim TrimBase_Num As Long
    Dim TrimBase_Item() As String
    Dim TrimBase_DSPWave As New DSPWave
    Dim R As Integer
    Dim TrimBase_DSPWave_Final() As New DSPWave
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    TrimBase_temp = Split(TrimBase, ";")
    TrimBase_Num = UBound(TrimBase_temp)
    
    Dim DigSrc_Sample_Size_Real_Temp() As String    ' Fix 20190812
    Dim Divide_Result As Long
    DigSrc_Sample_Size_Real_Temp = Split(DigSrc_Sample_Size_Real, "@")
    Divide_Result = DigSrc_Sample_Size_Real_Temp(1) / DigSrc_Sample_Size_Real_Temp(0)
          
    Dim Bin_arry() As Long
    ReDim Bin_arry(DigSrc_Sample_Size_Real_Temp(0) - 1)
    Dim InDspWave_New As New DSPWave
    Dim Dec_Trim_Temp As Long
    Dim t As Integer
    
    Dim PatCount As Long, PattArray() As String
    Dim OutDspWave As New DSPWave
               TrimTarget = 0.3 ''Just for Flag    2017/Dec
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If
    Call HardIP_InitialSetupForPatgen
    Dim Ts As Variant, TestSequenceArray() As String
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPwave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long, z As Long
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    
    Dim CapValue As New PinListData, CapValue_V1 As New PinListData, CapValue_V2 As New PinListData
    CapValue.AddPin ("CapValueString")
    On Error GoTo ErrorHandler

    Call GetFlowTName
 
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Interpose_PrePat <> "" Then ''''180109 update
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TheHdw.Patterns(patset).Load

    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)

    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    Dim b_HigherThanTarget As New SiteBoolean
    b_HigherThanTarget = False
    
    Dim LastSectionV1V2_Index As Long
    LastSectionV1V2_Index = 0
    Dim OutputTrimCode As String
    Dim SourceTrimCode As String
    Dim TestLimitIndex As Long '
    
'    Dim TrimStart_1st() As String
    Dim Dec_TrimStart_1st As Long
    
    Dim StoredTargetTrimCode As New DSPWave
    Dim b_MatchTagetCap As New SiteBoolean
    Dim b_DisplayCap As New SiteBoolean
    b_MatchTagetCap = False
    b_DisplayCap = False
    
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    Dim StoreEachTrimCode() As New DSPWave
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    
    Dim DigCapData() As New PinListData
    ReDim DigCapData(DigSrc_Sample_Size + 1) As New PinListData
    
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i

    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", "")
    End If

'' ===============================   TrimStart_1st = TrimStart===============================================
    Dec_TrimStart_1st = Bin2Dec(TrimStart)
    
    InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong

    Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, DigSrc_Sample_Size, InDSPwave)

    'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, InDspWave)
    TheExec.Datalog.WriteComment ("========First Time Setup========")
    
    For Each site In TheExec.sites
        SourceTrimCode = ""
        For k = 0 To InDSPwave(site).SampleSize - 1
            SourceTrimCode = SourceTrimCode & CStr(InDSPwave(site).Element(k))
        Next k
        TheExec.Datalog.WriteComment ("Site_" & site & " Initial Source Trim Code = " & SourceTrimCode)
        SourceTrimCode = StrReverse(SourceTrimCode)
        Dec_Trim_Temp = Bin2Dec(SourceTrimCode)
        TrimOffset = CInt(TrimOffset)
        Dec_Trim_Temp = Dec_Trim_Temp + TrimOffset
        
        'TheExec.Datalog.WriteComment ("Site_" & Site & " Initial Source Trim Code = " & SourceTrimCode)
    Next site
    
    InDspWave_New.CreateConstant 0, DigSrc_Sample_Size_Real_Temp(1)
                
    Call Dec2Bin(Dec_Trim_Temp, Bin_arry())
    
    For z = 0 To UBound(Bin_arry)
        InDspWave_New.Element(z) = Bin_arry(UBound(Bin_arry) - z)
        For t = 1 To Divide_Result - 1
            InDspWave_New.Element(z + ((UBound(Bin_arry) + 1) * t)) = InDspWave_New.Element(z)
        Next t
    Next z
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)
    
    For Each site In TheExec.sites
        StoreEachTrimCode(0)(site).Data = InDSPwave(site).Data
    Next site
    
'=========Set Up DigCap parameter=============================
    Dim Decompose_DigCapData() As String
    Dim OutDspWave_elementnum As Long
    Dim Digcap_width() As String
    Dim Digcap_sum As Long
    
    
    If Digcap_key <> "" Then
        Decompose_DigCapData() = Split(CUS_Str_DigCapData, ",")
        For i = 1 To UBound(Decompose_DigCapData)
            Digcap_width() = Split(Decompose_DigCapData(i), ":")
            Digcap_sum = Digcap_sum + Digcap_width(0)
            If InStr(Decompose_DigCapData(i), Digcap_key) <> 0 Then
                OutDspWave_elementnum = Digcap_sum - 1
                TheExec.Datalog.WriteComment "capture digcap name and bits  " & Decompose_DigCapData(i)
                Exit For
            End If
        Next i
    Else
        OutDspWave_elementnum = 0
    End If
     
    
    Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Call TheHdw.Patterns(PattArray(0)).start
'     For Each Site In TheExec.sites
        CapValue.Value = OutDspWave.Element(OutDspWave_elementnum)
'     Next Site
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    ''Update Interpose_PreMeas 20170801
    Dim TestSeqNum As Integer
    TestSeqNum = 0
           
    TheHdw.Digital.Patgen.HaltWait
'    For Each Site In TheExec.sites
    DigCapData(0) = CapValue
    b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
    PastDSPWave = InDSPwave
    TestNameInput = "Cap_Value_"
    TestLimitIndex = 0
'    Next Site

    Dim CUS_Str_MainProgram As String: CUS_Str_MainProgram = ""
    
    For Each site In TheExec.sites
        If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
            TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
        End If
    Next site
    
    b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
    b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
    
    For Each site In TheExec.sites
        If b_DisplayCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
            StoredTargetTrimCode(site).Data = InDSPwave(site).Data
            b_StopTrimCodeProcess(site) = True
        End If
    Next site
    
    TheExec.Datalog.WriteComment ("======================================================================================")
    
    Dim b_KeepGoing As New SiteBoolean
    
    Dim PreviousCap As New PinListData
   
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    StoreEachIndex = 1
    
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
                
            For j = SrcStartBit To (SrcEndBit + 1) Step -1
                ''===============up for src bit step
        
        '            For Each Site In TheExec.sites
        '                CapValue = CStr(OutDspWave(Site).Element(0))
        '            Next Site
                If b_FirstExecution = True Then
                    b_ControlNextBit = True
                    If j = SrcEndBit Then ''=============Trim from Format untill same======1124
                        b_ControlNextBit = False
                    End If
                Else
                    ''20160716-Control next bit to 1 no matter first or last progress
                    b_ControlNextBit = True
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                End If
        
                If b_FirstExecution = True And j = SrcEndBit Then
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPwave)
                Else
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HigherThanTarget, j, b_ControlNextBit, InDSPwave)
                End If
        
        '            Dim Dec_Trim_Temp As Long
                    
                For Each site In TheExec.sites
                    SourceTrimCode = ""
                    For k = 0 To InDSPwave(site).SampleSize - 1
                        SourceTrimCode = SourceTrimCode & CStr(InDSPwave(site).Element(k))
                    Next k
                    SourceTrimCode = StrReverse(SourceTrimCode)
                    Dec_Trim_Temp = Bin2Dec(SourceTrimCode)
                    TrimOffset = CInt(TrimOffset)
                    Dec_Trim_Temp = Dec_Trim_Temp + TrimOffset
                    
        '                Dim DigSrc_Sample_Size_Real_Temp() As String
        '                Dim Divide_Result As Long
        '                DigSrc_Sample_Size_Real_Temp = Split(DigSrc_Sample_Size_Real, "@")
        '                Divide_Result = DigSrc_Sample_Size_Real_Temp(1) / DigSrc_Sample_Size_Real_Temp(0)
        '
        '
        '                Dim Bin_arry() As Long
        '                ReDim Bin_arry(DigSrc_Sample_Size_Real_Temp(0) - 1)
        '                Dim InDspWave_New As New DSPWave
                    InDspWave_New.CreateConstant 0, DigSrc_Sample_Size_Real_Temp(1)
                    
                    Call Dec2Bin(Dec_Trim_Temp, Bin_arry())
                    
                    Dim Array_code As String
                    'Dim t As Integer
                    
                    Array_code = ""
                    
                    For z = 0 To UBound(Bin_arry)
                        InDspWave_New.Element(z) = Bin_arry(UBound(Bin_arry) - z)
                        For t = 1 To Divide_Result - 1
                            InDspWave_New.Element(z + ((UBound(Bin_arry) + 1) * t)) = InDspWave_New.Element(z)
                        Next t
                        Array_code = Array_code & InDspWave_New(site).Element(z)
                    Next z
                    TheExec.Datalog.WriteComment "Site " & site & "  ,Digcap Decimal value:  " & Dec_Trim_Temp & "  ,Array Code:" & Array_code
                Next site
                    
         
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)
                   
                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site).Data = InDSPwave(site).Data
                Next site
                    
                '' Debug use
                '' ==============================================================================================
                '' 20160716 - Modify trim code rule
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
                    
                For Each site In TheExec.sites
                    If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
                        If b_KeepGoing(site) = True Then

                            SourceTrimCode_SiteVariant(site) = ""
                            
                            For k = 0 To InDSPwave(site).SampleSize - 1

                                SourceTrimCode_SiteVariant = SourceTrimCode_SiteVariant & CStr(InDSPwave(site).Element(k))
                            Next k
                            Dim OutputDec As String

                            TheExec.Datalog.WriteComment ("Site_" & site & " Source Trim Code = " & SourceTrimCode_SiteVariant(site))
                        End If
                    End If
                Next site
                
                Set OutDspWave = Nothing
                Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                Call TheHdw.Patterns(PattArray(0)).start
                    
                    ''Update Interpose_PreMeas 20170801
                TestSeqNum = 0
                                
                
                TheHdw.Digital.Patgen.HaltWait
                
                CapValue.Value = OutDspWave.Element(OutDspWave_elementnum)
                b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                PastDSPWave = InDSPwave
        
                TestLimitIndex = TestLimitIndex + 1
                    
                    '' 20160712 - Modify to use WriteComment to display output frequency.
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
        '                    TheExec.Datalog.WriteComment ("Site " & Site & " Output CapValue = " & CapValue.Pins(Site).Value(Site))
                        TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
                    End If
                Next site
               
                b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
                        If b_MatchTagetCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
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
                TheExec.Datalog.WriteComment ("======================================================================================")
            Next j
        
        
                    
                    
                    
            If TrimOffset <> "" Then
                    
            'Dim SourceTrimCodeArray(8) As Long
                TrimOffset = CInt(TrimOffset)
                If TrimBase <> "" Then
                        ''Dim TrimBase_SiteNum() As String
                        
        ''                    Dim gDictDSPWaves As Scriptin.Dictionary
        ''                    Set gDictDSPWaves = New Scripting.Dictionary
                        
                    TrimBase_temp = Split(TrimBase, ";")
                    TrimBase_Num = UBound(TrimBase_temp)
        
                    ReDim TrimBase_DSPWave_Final(TrimBase_Num)
                    TrimBase_DSPWave.CreateConstant 0, DigSrc_Sample_Size_Real_Temp(0)
                    TrimBase_DSPWave_Final(0).CreateConstant 0, DigSrc_Sample_Size_Real_Temp(0)
                    

                    For R = 0 To TrimBase_Num
                        TrimBase_Item = Split(TrimBase_temp(R), ":")
                        For Each site In TheExec.sites

                            SourceTrimCode_SiteVariant(site) = StrReverse(SourceTrimCode_SiteVariant(site))   ' for Bin to Dec need to reverse
                            SourceTrimCode_temp = Bin2Dec(SourceTrimCode_SiteVariant(site))
                            SourceTrimCode_temp = SourceTrimCode_temp + TrimOffset - 256
                            SourceTrimCode_final = TrimBase_Item(1) - SourceTrimCode_temp
                            
                            
                            Dim printDec As String
                            printDec = SourceTrimCode_final
                            'TrimBase_DSPWave(Site).Element(0) = SourceTrimCode_final
                            'SourceTrimCode = Dec2Bin(SourceTrimCode_temp, SourceTrimCodeArray())
                            '///////////////////////////////////Add to src 9 bit binary 20190528/////////////////////////////////
                            If SourceTrimCode_final < 0 Then
                                If SourceTrimCode_final <= 0 - 2 ^ (DigSrc_Sample_Size_Real_Temp(0) - 1) Then
                                    TheExec.Datalog.WriteComment ("Your number is too small and current bits are not enough to save")
                                    GoTo ErrorHandler
                                Else
                                    SourceTrimCode_final = SourceTrimCode_final + 2 ^ (DigSrc_Sample_Size_Real_Temp(0))
                                End If
                            End If
                       
                            Dim Bin_arry_Base() As Long
                            ReDim Bin_arry_Base(DigSrc_Sample_Size_Real_Temp(0) - 1)
                            Dim Array_code_Base As String
                            
                            Array_code_Base = ""
                            Call Dec2Bin(SourceTrimCode_final, Bin_arry_Base())
                    
                            For z = 0 To UBound(Bin_arry_Base)
                                TrimBase_DSPWave.Element(z) = Bin_arry_Base(UBound(Bin_arry) - z)
                                Array_code_Base = Array_code_Base & TrimBase_DSPWave(site).Element(z)
                            Next
                            '////////////////////////////////////////////////////////////////////////////////////////////////////
                            TheExec.Datalog.WriteComment ("Site_" & site & " Final Source Trim Code = " & SourceTrimCode_SiteVariant(site) & " & " & TrimBase_Item(0) & " = " & Array_code_Base & ",Decimal value: " & printDec)

                            'theexec.Datalog.WriteComment ("Site_" & Site & " Final Source Trim Code = " & SourceTrimCode & " & " & TrimBase_Item(0) & " = " & SourceTrimCode_final)
                        Next site
                        
                        TrimBase_DSPWave_Final(R) = TrimBase_DSPWave
                        Call AddStoredCaptureData(TrimBase_Item(0), TrimBase_DSPWave_Final(R))
                       '' TrimBase_DSPWave(R).CreateConstant , SourceTrimCode_final
                       '' AddStoredCaptureData TrimBase_Item(0) + CStr(site), TrimBase_DSPWave(R)
                       '' TheExec.Datalog.WriteComment ("Site_" & site & " : " & TrimBase_Item(0) & " = " & SourceTrimCode_final)
                    Next R
                End If
            Else
                For Each site In TheExec.sites
                    If CapValue = 1 Then

                        SourceTrimCode_SiteVariant(site) = ""
                        SourceTrimCode_SiteVariant(site) = SourceTrimCode_SiteVariant(site) & "0"
                        For k = 1 To InDSPwave(site).SampleSize - 1
                            SourceTrimCode_SiteVariant(site) = SourceTrimCode_SiteVariant(site) & CStr(InDSPwave(site).Element(k))
                        Next k
                        InDSPwave(site).Element(0) = 0
                    Else

                        
                        SourceTrimCode_SiteVariant(site) = ""
                        For k = 0 To InDSPwave(site).SampleSize - 1
                            SourceTrimCode_SiteVariant(site) = SourceTrimCode_SiteVariant(site) & CStr(InDSPwave(site).Element(k))
                        Next k
                    End If

                    TheExec.Datalog.WriteComment ("Site_" & site & " Final Source Trim Code = " & SourceTrimCode_SiteVariant(site))
                Next site
            End If
        Next i
    End If
    
    Call HardIP_WriteFuncResult(, , Inst_Name_Str)
    
    '//////////////////////////////////////////    add for capture bit more than one and print
    
    If DigCap_DataWidth <> 0 Then
        Dim DigCapPinAry() As String, NumberPins As Long
        Dim CUS_Str_DigCapData_temp As String
        Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
    
        If NumberPins > 1 Then
            Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
            Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins, , DigCap_Pin.Value)
    
        ElseIf NumberPins = 1 Then
            Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
    '                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1429 As Long: VBT_LIB_HardIP_ProfileMark_1429 = ProfileMarkEnter(2, instance_name & "_" & "MeasC&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1427")    ' Profile Mark
            Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, CUS_Str_MainProgram, , DigCap_Pin.Value)
    '                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1429    ' Profile Mark
        End If
    End If
    '//////////////////////////////////////////
    
    If TrimStoreName <> "" Then
       Call Checker_StoreDigCapAllToDictionary(TrimStoreName, InDSPwave)
    End If
    
    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel InDSPwave, DigSrc_Sample_Size, ConvertedDataWf
    
    Call GetFlowTName
    
    If gl_UseStandardTestName_Flag = True Then
        Call Report_ALG_TName_From_Instance(OutputTname_format, "C", StrSeparatebyEqual(0), gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex))
        TestNameInput = Merge_TName(OutputTname_format)
    Else
        TestNameInput = TheExec.DataManager.instanceName & "DDR_Sweep"
    End If
        
        
    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TestNameInput, PinName:="Voffset_Trim_Dec", ForceResults:=tlForceFlow
        
    
'    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), TName:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, StoredTargetTrimCode)
    Call TheHdw.Patterns(PattArray(0)).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
   
    TheHdw.Digital.Patgen.HaltWait
    
    If Interpose_PrePat <> "" Then '''180109 update
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    Call SetForceCondition(Interpose_PostTest)
    
    Dim sl_FUSE_Val As New SiteLong

    DebugPrintFunc patset.Value
    
    TheHdw.DSSC.MoveMode = tlDSSCMoveModeIIM
    
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeDig_SeaHawk function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function





Public Function TrimCodeBasicDig(Optional patset As Pattern, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
Optional TrimTarget As Double, Optional TrimStart As String, Optional TrimFormat As String, Optional TrimStoreName As String, _
Optional IncreaseFlag As Boolean = True, Optional BinarySearchFlag As Boolean = True, Optional TrimPrcocessAll As Boolean = True, _
Optional Interpose_PrePat As String, Optional DigCap_Pin As PinList, Optional DigCap_Sample_Size As Long, Optional Validating_ As Boolean, _
Optional Interpose_PostTest As String, Optional TrimOffset As String, Optional TrimBase As String, Optional DigSrc_Sample_Size_Real As String, Optional DigSrc_Assignment As String) As Long
'Dylan Edited 20190615

    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If
    On Error GoTo ErrorHandler
    
    Dim site As Variant
    Dim PatCount As Long
    
    Dim TrimBaseStr() As String
    Dim TrimBaseNum() As String
    Dim OffsetDelta As Integer
    Dim Inst_Name_Str As String: Inst_Name_Str = TheExec.DataManager.instanceName
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    Dim LimitCodeStart As Integer
    Dim LimitCodeEnd As Integer
    Dim i, j, k, X As Integer
    Dim SrcEndBit As Integer
    Dim RegSeparate As Integer
    Dim SrcStartBit As Integer
    Dim ExecutionMax As Integer
    Dim DecTrimStart As Integer
    Dim InDSPwave As New DSPWave
    Dim OutDspWave As New DSPWave
    Dim InDspWave_New As New DSPWave
    Dim InitialDSPWave As New DSPWave
    Dim CaptureDSPWave() As New DSPWave
    ReDim CaptureDSPWave(0)
    Dim ProcessDoneDSPWave() As New DSPWave
    ReDim ProcessDoneDSPWave(0)
    Dim ConvertedDataWf As New DSPWave
    
    Dim PattArray() As String
    Dim SourceTrimCode As String
    Dim CapValue As New PinListData
    Dim StrSeparatebyEqual() As String
    Dim StrSeparatebyColon() As String
    Dim StrSeparatebyComma() As String
    Dim RegSeparatebyComma() As String
    Dim EachRegSize As New SiteLong
    Dim InitStateCode As New SiteLong
    Dim TrimOriginalSize As New SiteLong
    Dim TrimScanPoint As New SiteLong
    Dim TrimOffsetPoint As New SiteLong
    Dim CaptureDSPFlag As New SiteBoolean
    Dim b_HigherThanTarget As New SiteBoolean
    Dim b_StopTrimCodeProcess As New SiteBoolean
    Dim DigSrc_Sample_Size_Real_Temp() As String
    Dim assignment As New DSPWave
    Dim DigSrc_Assignment_Temp() As String
    Dim AssignmentDSPWave As New DSPWave
    Dim Src_dig As New SiteBoolean
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    
                
    Call HardIP_InitialSetupForPatgen
    Call GetFlowTName
    TheHdw.Patterns(patset).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    CapValue.AddPin ("CapValueString")
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    RegSeparatebyComma = Split(DigSrc_Sample_Size_Real, ";")
    Src_dig = False
    assignment.CreateConstant 0, 1, DspLong
    AssignmentDSPWave.CreateConstant 0, 1, DspLong
    
    If DigSrc_Assignment <> "" Then
    
    DigSrc_Assignment_Temp = Split(DigSrc_Assignment, "+")
    AssignmentDSPWave.CreateConstant 0, UBound(DigSrc_Assignment_Temp) + 1, DspLong
    For i = 0 To UBound(DigSrc_Assignment_Temp)
        If LCase(DigSrc_Assignment_Temp(i)) <> "sweep" Then
          assignment = GetStoredCaptureData(DigSrc_Assignment_Temp(i))
          AssignmentDSPWave.Element(i) = 1
          Src_dig = True
        End If
    Next
    
    End If
    
    For i = 0 To ExecutionMax
        StrSeparatebyEqual = Split(StrSeparatebyComma(0), "=")
        StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
        SrcEndBit = StrSeparatebyColon(1)
        SrcStartBit = StrSeparatebyColon(0)
        DigSrc_Sample_Size_Real_Temp = Split(RegSeparatebyComma(i), "@")
        CaptureDSPWave(0).CreateConstant 0, 1, DspLong                  ' Avoid sweep fail which any site
        RegSeparate = DigSrc_Sample_Size_Real_Temp(1) / DigSrc_Sample_Size_Real_Temp(0)
        InDSPwave.CreateConstant 0, CLng(DigSrc_Sample_Size_Real_Temp(1)), DspLong
        InDspWave_New.CreateConstant 0, CLng(DigSrc_Sample_Size_Real_Temp(1)), DspLong
        
        If BinarySearchFlag = False Then                                ' This initial method Only support linear search mode
            If TrimStart = "" Then
                If IncreaseFlag = True Then                             ' Based on IncreaseFlag state to determine start point
                    TrimStart = 0
                Else
                    TrimStart = CStr(2 ^ (SrcStartBit + 1) - 1)
                End If
            End If
        End If
        
'''        For Each site In TheExec.sites.Active
            CaptureDSPFlag = False
            If UBound(StrSeparatebyColon) > 1 Then
                InitStateCode = StrSeparatebyColon(2)
            End If
            EachRegSize = CLng(DigSrc_Sample_Size_Real_Temp(0))
            TrimScanPoint = CLng(TrimStart)
            TrimOffsetPoint = CLng(TrimOffset)
            TrimOriginalSize = CLng(SrcStartBit) + 1
'''        Next site
        
        
        
        TrimStart = CStr(CInt(TrimStart) + CInt(TrimOffset))
        InitialDSPWave.CreateConstant TrimStart, 1, DspLong                           ' Define first trim code from TrimStart
        rundsp.CreateFlexibleDSPWave InitialDSPWave, CLng(DigSrc_Sample_Size_Real_Temp(0)), InDSPwave
        rundsp.ElementTransformer InDSPwave, CLng(DigSrc_Sample_Size_Real_Temp(0)), CLng(DigSrc_Sample_Size_Real_Temp(1))
        
               
        If BinarySearchFlag = True Then                                               ' Binary Search
            For j = SrcStartBit + 1 To SrcEndBit Step -1
            
               If j = SrcEndBit Then
                   b_ControlNextBit = False
            
               Else
                    b_ControlNextBit = True
            
               End If
               
            
                rundsp.ReAssignmentDSPWave InDSPwave, RegSeparate, InDspWave_New, Src_dig, assignment, AssignmentDSPWave              ' ReAssignment DSPWave element to each register
                
                
                
                '*****************************************************For HardIP_D2D debug*****************************************************
'''''''''''                Dim Constant As Integer
'''''''''''                Dim ForConstantCode As String
'''''''''''                Dim ForConstantSplit() As String
'''''''''''                ForConstantCode = "1,0,1,0,1,0,1,0"
'''''''''''                ForConstantCode = StrReverse(ForConstantCode)
'''''''''''
'''''''''''                ForConstantSplit = Split(ForConstantCode, ",")
'''''''''''                If theexec.DataManager.instanceName = "D2D_ZCPDD2DIMPCL_PP_SHKA0_C_FULP_AN_AMXX_DLL_JTG_PRG_ALLFV_SI_D2DIMPCL_ZCPD_NV" Then
'''''''''''                    Constant = 0
'''''''''''                    For k = 0 To 7                                    '|TrimCode|Constanct|TrimCode|Constanct|Constanct|
'''''''''''                        If CStr(ForConstantSplit(Constant)) = "0" Then
'''''''''''                            InDspWave_New(0).Element(k) = 0
'''''''''''                        Else
'''''''''''                            InDspWave_New(0).Element(k) = 1
'''''''''''                        End If
'''''''''''                        Constant = Constant + 1
'''''''''''                    Next k
'''''''''''                    Constant = 0
'''''''''''                    For k = 8 To 15
'''''''''''                        If CStr(ForConstantSplit(Constant)) = "0" Then
'''''''''''                            InDspWave_New(0).Element(k) = 0
'''''''''''                        Else
'''''''''''                            InDspWave_New(0).Element(k) = 1
'''''''''''                        End If
'''''''''''                        Constant = Constant + 1
'''''''''''                    Next k
'''''''''''                    Constant = 0
'''''''''''                    For k = 24 To 31
'''''''''''                        If CStr(ForConstantSplit(Constant)) = "0" Then
'''''''''''                            InDspWave_New(0).Element(k) = 0
'''''''''''                        Else
'''''''''''                            InDspWave_New(0).Element(k) = 1
'''''''''''                        End If
'''''''''''                        Constant = Constant + 1
'''''''''''                    Next k
'''''''''''                End If
                '*****************************************************For HardIP_D2D debug*****************************************************
                
                
                
                
                
                For Each site In TheExec.sites.Active
                    SourceTrimCode = ""
                    For k = InDspWave_New.SampleSize - 1 To 0 Step -1
                        SourceTrimCode = SourceTrimCode & CStr(InDspWave_New.Element(k))
                    Next k
                    
                    If j = SrcStartBit + 1 Then
                        TheExec.Datalog.WriteComment ("InitialBit , Trim Code Bit " & SourceTrimCode)
                    Else
                        TheExec.Datalog.WriteComment ("Setup Bit " & (j) & ", Trim Code Bit " & SourceTrimCode)
                    End If
                Next site
                
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)   '''' not LSB to MSB ???
                Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                
                Dim DSPwave_temp As New DSPWave
                Dim W As Long
                
                
                For Each site In TheExec.sites.Active
                
                DSPwave_temp.CreateConstant 0, 1, DspLong
                
                
                
                DSPwave_temp = InDspWave_New.ConvertStreamTo(tldspParallel, CLng(DigSrc_Sample_Size_Real_Temp(0)), 0, Bit0IsMsb)
                
                For W = 0 To DSPwave_temp.SampleSize - 1
                
                TheExec.Datalog.WriteComment CStr(DSPwave_temp.Element(W))
                
                Next
                Next site
                            
                Call TheHdw.Patterns(PattArray(0)).start
                TheHdw.Digital.Patgen.HaltWait
                
                For Each site In TheExec.sites.Active
                    CapValue.Value = OutDspWave.Element(0)
                    ' TrimTarget is your expected transfer-point
                    b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(EqualTo, 0)
                    
                    '///////////////// fail stop ////////////////////
                    If TrimPrcocessAll = False Then
                        If b_HigherThanTarget = True Then
                            If CaptureDSPFlag = False Then
                                CaptureDSPWave(0) = InDspWave_New.Copy                  ' For each site arry(0) is uncalculate value
                                CaptureDSPFlag = True
                            End If
                            b_StopTrimCodeProcess(site) = True
                        End If
                    Else
                    '/////////////////  do all  ////////////////////
                        If TrimPrcocessAll = True Then
                           CaptureDSPWave(0) = InDspWave_New.Copy
                        End If
                    End If
                    '///////////////////////////////////////////////
                    TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & CapValue.Value)
                Next site
    
                TheExec.Datalog.WriteComment ("======================================================================================")

                If j <> SrcEndBit Then '20190530 Need to include initial bit so endbit don't need to do transform
                    If j - 1 = SrcEndBit Then
                      b_ControlNextBit = False
                      
                     End If
                      
                     
                    rundsp.SetupBinaryTrimCodeBit InDspWave_New, b_HigherThanTarget, j - 1, InitStateCode, TrimOffsetPoint, TrimOriginalSize, InDSPwave, b_ControlNextBit, AssignmentDSPWave, EachRegSize
                    rundsp.ReAssignmentDSPWave InDSPwave, RegSeparate, InDspWave_New, Src_dig, assignment, AssignmentDSPWave      ' ReAssignment DSPWave element to each register
                    
         
                    
                End If
                
                If TrimPrcocessAll = False Then                                     ' Immediately stop if all site capture done
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                
            Next j
           
            
        Else                                                                        ' Linear Search
            If IncreaseFlag = True Then
                LimitCodeEnd = CInt((2 ^ (SrcStartBit + 1) - 1))
                LimitCodeStart = TrimStart
            Else                                                                    ' Decrease sweep
                LimitCodeEnd = TrimStart
                LimitCodeStart = 0
            End If
            
            For j = LimitCodeStart To LimitCodeEnd
            
                rundsp.ReAssignmentDSPWave InDSPwave, RegSeparate, InDspWave_New, Src_dig, assignment, AssignmentDSPWave   ' ReAssignment DSPWave element to each register
                
                For Each site In TheExec.sites.Active
                    SourceTrimCode = ""
                    For k = InDspWave_New.SampleSize - 1 To 0 Step -1
                        SourceTrimCode = SourceTrimCode & CStr(InDspWave_New.Element(k))
                    Next k
                    If j = LimitCodeStart Then
                        b_StopTrimCodeProcess = False                               ' Inititial b_StopTrimCodeProcess flag
                        TheExec.Datalog.WriteComment ("InitialCode, Trim Code Bit " & SourceTrimCode)
                    Else
                        TheExec.Datalog.WriteComment ("LinearSweep " & TrimScanPoint(0) & "," & " Trim Code Bit " & SourceTrimCode)
                    End If
                Next site
                
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)
                Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                
                Call TheHdw.Patterns(PattArray(0)).start
                TheHdw.Digital.Patgen.HaltWait
                
                For Each site In TheExec.sites.Active
                    CapValue.Value = OutDspWave.Element(0)
                    ' TrimTarget is your expected transfer-point
                    b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(EqualTo, 0)
                    If b_HigherThanTarget = True Then
                    
                        If CaptureDSPFlag = False Then
                            CaptureDSPWave(0) = InDspWave_New.Copy                  ' For each site arry(0) is uncalculate value
                            CaptureDSPFlag = True
                        End If
                        
                        If TrimPrcocessAll = False Then
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                    TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
                Next site
    
                TheExec.Datalog.WriteComment ("======================================================================================")
                
                If j <> LimitCodeEnd Then
'                    Judgment addition "1" or "0" based on b_HigherThanTarget value
                    rundsp.SetupLinearTrimCodeBit IncreaseFlag, TrimScanPoint, b_HigherThanTarget, EachRegSize, InDSPwave, TrimPrcocessAll
                End If
                
                If TrimPrcocessAll = False Then                                     ' Immediately stop if all site capture done
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
            Next j
 
        End If
        
        Call HardIP_WriteFuncResult(, , Inst_Name_Str)
        rundsp.ConvertToLongAndSerialToParrel InDSPwave, EachRegSize, ConvertedDataWf
        
        Call GetFlowTName
        
        If gl_UseStandardTestName_Flag = True Then
            Call Report_ALG_TName_From_Instance(OutputTname_format, "C", TrimStoreName, gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex))
            TestNameInput = Merge_TName(OutputTname_format)
                  
        Else
            TestNameInput = TheExec.DataManager.instanceName & "DDR_Sweep"
        End If
        
        
        TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TestNameInput, PinName:="D2D_Trim_Dec", ForceResults:=tlForceFlow
        
        'TheExec.Flow.TestLimit ConvertedDataWf.Element(0), TName:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow
            
           
        If TrimBase <> "" Then
            TrimBaseStr = Split(TrimBase, ";")
            OffsetDelta = CInt(TrimOffset) + 2 ^ CInt(StrSeparatebyColon(0))
            For j = 0 To UBound(TrimBaseStr)
                TrimBaseNum = Split(TrimBaseStr(j), ":")
                ReDim Preserve CaptureDSPWave(UBound(CaptureDSPWave) + 1)
                ReDim Preserve ProcessDoneDSPWave(UBound(ProcessDoneDSPWave) + 1)                   'This dspwave will save as to Dictionary
                
                CaptureDSPWave(UBound(CaptureDSPWave)).CreateConstant 0, 1, DspLong
                ProcessDoneDSPWave(UBound(ProcessDoneDSPWave)).CreateConstant 0, 1, DspLong
                                
                If CaptureDSPFlag.Any(True) Then
                    rundsp.CalculateDSPWaveforTrimCode CaptureDSPWave(0), EachRegSize, CaptureDSPWave(UBound(CaptureDSPWave)), _
                                                       CInt(OffsetDelta), CInt(TrimBaseNum(1)), ProcessDoneDSPWave(UBound(ProcessDoneDSPWave))
                                                       
                    
                    
                    
'''''                    TheExec.Flow.TestLimit CaptureDSPWave(UBound(CaptureDSPWave)).Element(0), TName:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow
                                                       
                    AddStoredCaptureData TrimBaseNum(0), ProcessDoneDSPWave(UBound(ProcessDoneDSPWave))
                End If
            Next j
        End If
    Next i
    
    
    If TrimStoreName <> "" Then
       Call Checker_StoreDigCapAllToDictionary(TrimStoreName, InDSPwave)
    End If
    
        
    If Interpose_PrePat <> "" Then '''180109 update
        Call SetForceCondition("RESTOREPREPAT")
    End If

    Call SetForceCondition(Interpose_PostTest)
        
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeBasicDig function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function


Public Function PCIE_Eye_Diagram_0() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane5(62) As New SiteVariant

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
            Eye_Diagram_Binary_Lane0(i + 31)(site) = Eye_Diagram_Binary_Lane0(i + 31)(site) & Mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 4, 1)
            'Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
           ' Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
            'Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
           ' Eye_Diagram_Binary_Lane4(i + 31)(Site) = Eye_Diagram_Binary_Lane4(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane0(i + 31)(site))
      '  Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane0(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim X As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane0(i + 31)(site)) - 1
                For X = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + Mid(Eye_Diagram_Binary_Lane0(i + 31)(site), iLen - X + 1, 1) * 2 ^ X
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For X = 0 To iLen
                        If Mid(Eye_Diagram_Binary_Lane0(i + 31)(site), iLen - X + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If X = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf Mid(Eye_Diagram_Binary(i + 31)(site), iLen - X + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (X - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next X
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
            '  Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
        
         '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
        
Next site


    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane0"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane0"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane0"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane0"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane0"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane0"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane0"
    

    Call PCIE_Eye_Diagram_1
    Call PCIE_Eye_Diagram_2
    Call PCIE_Eye_Diagram_3
    Call PCIE_Eye_Diagram_4
    'Call PCIE_Eye_Diagram_5
End Function

'errHandler:
'    TheExec.DataLog.WriteComment "error in Meas_FreqVoltCurr_Universal_func"
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function

Public Function PCIE_Eye_Diagram_1() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
            'Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
            Eye_Diagram_Binary_lane1(i + 31)(site) = Eye_Diagram_Binary_lane1(i + 31)(site) & Mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 3, 1)
           ' Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
            'Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
           ' Eye_Diagram_Binary_Lane4(i + 31)(Site) = Eye_Diagram_Binary_Lane4(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(site))
      '  Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim X As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_lane1(i + 31)(site)) - 1
                For X = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + Mid(Eye_Diagram_Binary_lane1(i + 31)(site), iLen - X + 1, 1) * 2 ^ X
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For X = 0 To iLen
                        If Mid(Eye_Diagram_Binary_lane1(i + 31)(site), iLen - X + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If X = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf Mid(Eye_Diagram_Binary(i + 31)(site), iLen - X + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (X - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next X
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
            ' Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
        
        
         '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
Next site



    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane1"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane1"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane1"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane1"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane1"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane1"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane1"
End Function

Public Function PCIE_Eye_Diagram_2() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
          '  Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
           ' Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
         Eye_Diagram_Binary_Lane2(i + 31)(site) = Eye_Diagram_Binary_Lane2(i + 31)(site) & Mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 2, 1)
          ' Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
          '  Eye_Diagram_Binary_Lane4(i + 31)(Site) = Eye_Diagram_Binary_Lane4(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(site))
'        Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim X As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane2(i + 31)(site)) - 1
                For X = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + Mid(Eye_Diagram_Binary_Lane2(i + 31)(site), iLen - X + 1, 1) * 2 ^ X
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For X = 0 To iLen
                        If Mid(Eye_Diagram_Binary_Lane2(i + 31)(site), iLen - X + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If X = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf Mid(Eye_Diagram_Binary(i + 31)(site), iLen - X + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (X - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next X
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
            '  Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
        
         '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
        
Next site


    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane2"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane2"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane2"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane2"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane2"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane2"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane2"
End Function

Public Function PCIE_Eye_Diagram_3() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
'            Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
'            Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
'            Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
            Eye_Diagram_Binary_Lane3(i + 31)(site) = Eye_Diagram_Binary_Lane3(i + 31)(site) & Mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 1, 1)
'            Eye_Diagram_Binary_Lane4(i + 31)(Site) = Eye_Diagram_Binary_Lane4(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(site))
'        Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim X As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane3(i + 31)(site)) - 1
                For X = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + Mid(Eye_Diagram_Binary_Lane3(i + 31)(site), iLen - X + 1, 1) * 2 ^ X
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                 
                        
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For X = 0 To iLen
                        If Mid(Eye_Diagram_Binary_Lane3(i + 31)(site), iLen - X + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If X = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                
                                
                                End If
                            End If
                        ElseIf Mid(Eye_Diagram_Binary(i + 31)(site), iLen - X + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (X - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next X
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
             ' Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
        
   '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
Next site

 



    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane3"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane3"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane3"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane3"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane3"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane3"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane3"
End Function

Public Function PCIE_Eye_Diagram_4() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
'            Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
'            Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
'            Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
'            Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
            Eye_Diagram_Binary_Lane4(i + 31)(site) = Eye_Diagram_Binary_Lane4(i + 31)(site) & Mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 1, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(site))
'       Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim X As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane4(i + 31)(site)) - 1
                For X = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + Mid(Eye_Diagram_Binary_Lane4(i + 31)(site), iLen - X + 1, 1) * 2 ^ X
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For X = 0 To iLen
                        If Mid(Eye_Diagram_Binary_Lane4(i + 31)(site), iLen - X + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If X = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf Mid(Eye_Diagram_Binary(i + 31)(site), iLen - X + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (X - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next X
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
              'Eye_Diagram_Binary(i + 31)(site) = ""
        Next i
        
    '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
Next site


    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane4"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane4"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane4"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane4"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane4"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane4"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane4"
End Function







Public Function PCIE_Eye_Diagram_5() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane5(62) As New SiteVariant

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
'            Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
'            Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
'            Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
'            Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
'            Eye_Diagram_Binary_Lane4(i + 31)(site) = Eye_Diagram_Binary_Lane4(i + 31)(site) & Mid(Eye_Diagram_Binary(i + 31)(site), 6 * j - 1, 1)
            Eye_Diagram_Binary_Lane5(i + 31)(site) = Eye_Diagram_Binary_Lane5(i + 31)(site) & Mid(Eye_Diagram_Binary(i + 31)(site), 6 * j - 1, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 5, Binary string = " & Eye_Diagram_Binary_Lane5(i + 31)(site))
'       Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim X As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane5(i + 31)(site)) - 1
                For X = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + Mid(Eye_Diagram_Binary_Lane5(i + 31)(site), iLen - X + 1, 1) * 2 ^ X
                Next
                'process vertical width
                If TempDec(i + 31) < 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For X = 0 To iLen
                        If Mid(Eye_Diagram_Binary_Lane5(i + 31)(site), iLen - X + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If X = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf Mid(Eye_Diagram_Binary(i + 31)(site), iLen - X + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (X - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next X
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
              Eye_Diagram_Binary(i + 31)(site) = ""
        Next i
Next site


    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane5"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane5"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane5"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane5"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane5"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane5"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane5"
End Function



Public Function LPDPRX_Eye_Diagram() As Long
Dim i As Integer
Dim site As Variant
Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"
For Each site In TheExec.sites
' horizontal_width = 0
'    Zero_counter = 0
    For i = -31 To 31
        Call TheExec.Datalog.WriteComment("Site(" & site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim X As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary(i + 31)(site)) - 1
                For X = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + Mid(Eye_Diagram_Binary(i + 31)(site), iLen - X + 1, 1) * 2 ^ X
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For X = 0 To iLen
                        If Mid(Eye_Diagram_Binary(i + 31)(site), iLen - X + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If X = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf Mid(Eye_Diagram_Binary(i + 31)(site), iLen - X + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (X - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next X
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
              Eye_Diagram_Binary(i + 31)(site) = ""
        Next i
         '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
Next site

For Each site In TheExec.sites
    If vertical_width <> Empty Or horizontal_width <> Empty Then
    
        TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width"
        TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start"
        TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End"
        TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width"
        TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start"
        TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End"
        TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero"
    Else
        TheExec.Datalog.WriteComment "*****NO SIGNAL OPENING*****"
    End If
   Next site


End Function
Public Function Enable_HIP_Datalog_Format()

                With TheExec.Datalog
            .Setup.DatalogSetup.DisableInstanceNameInPTR = False
            .Setup.DatalogSetup.DisablePinNameInPTR = True
            .Setup.DatalogSetup.DisableChannelNumberInPTR = True
            .Setup.DatalogSetup.PTR_InstanceNameIsTINameOnly = True
    
    .Setup.Shared.Ascii.Columns.EnableCustomWidths = True
    .Setup.Shared.Ascii.Columns.Parametric.testName.Width = 150
    .Setup.Shared.Ascii.Columns.Parametric.Measured.Width = 16
    .Setup.Shared.Ascii.Columns.Functional.testName.Width = 150
    .Setup.Shared.Ascii.Columns.Functional.Pattern.Width = 80
            .ApplySetup
                End With
                
End Function

Public Function Disable_HIP_Datalog_Format()

                With TheExec.Datalog
            .Setup.DatalogSetup.DisableInstanceNameInPTR = False
            .Setup.DatalogSetup.DisablePinNameInPTR = False
            .Setup.DatalogSetup.DisableChannelNumberInPTR = True
            .Setup.DatalogSetup.PTR_InstanceNameIsTINameOnly = False
            .ApplySetup
                End With

End Function

Public Function Set_SEPVM_Ref_Level_Div()

    
    With TheHdw.DCVI.Pins("HSC_SEPVM_TEST_N")
        .Gate = False
        .mode = tlDCVIModeVoltage
        .Voltage = 0
        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
        .SetCurrentAndRange 0.002, 0.02
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    
    With TheHdw.DCVI.Pins("HSC_SEPVM_TEST_P_SRC")
'        .Gate = False
'        .mode = tlDCVIModeVoltage
'        .Voltage = 6
'        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
'        .SetCurrentAndRange 0.02, 0.2
'        .Connect tlDCVIConnectDefault
'        .Gate = True
        .mode = tlDCVIModeVoltage
        ''20170509 - Comment this
''            .Voltage = ForceV
        .Voltage = 6
        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
        ''20161018 - Swap current and current range sequence to avoid mode alarm
''            .Current = MI_TestCond_UVI80(i).CurrentRange
''            .CurrentRange.Value = MI_TestCond_UVI80(i).CurrentRange
        .SetCurrentAndRange 0.02, 0.02
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    'thehdw.Wait 100 * ms
End Function
Public Function SEPVM_Ref_measurement(UVI_P_SRC As PinList, UVI_P As PinList, UVI_N As PinList, Vdiff_High As PinList, Vdiff_Low As PinList, Source_P_Voltage As Double, Source_N_Voltage As Double, _
TargetVoltage As Double, TheHdwWait As Double)
'UVI_P_SRC=HSC_SEPVM_TEST_P_SRC  @  UVI_P=HSC_SEPVM_TEST_P_Meas  @  UVI_N=HSC_SEPVM_TEST_N  @ Vdiff_High=HSC_SEPVM_P_High  @  Vdiff_Low=HSC_SEPVM_N_Low
'@ Source_P_Voltage=6  @  Source_N_Voltage=0  @  TargetVoltage=0.75
    On Error GoTo errHandler
    Dim Vdiff_Pin As String: Vdiff_Pin = UVI_P + "," + UVI_N
    Dim Disconnect_Pin As String: Disconnect_Pin = UVI_P_SRC + "," + UVI_N + "," + UVI_P
    Dim MeasureVolt_P As New PinListData
    Dim MeasureVolt_N As New PinListData
'===============================================================================================================
'for 3. Force UVI80-1 6V==================================================================================
'===============================================================================================================
      With TheHdw.DCVI.Pins(UVI_P_SRC)
          .Gate = False
          .mode = tlDCVIModeVoltage
          .Voltage = Source_P_Voltage
          .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
          .SetCurrentAndRange 0.2, 0.2
          .Connect tlDCVIConnectHighForce
          .Gate = True
      End With
'===============================================================================================================
'for 3. Force UVI80-2 0V==================================================================================
'===============================================================================================================
      With TheHdw.DCVI.Pins(UVI_P)
          .Gate = False
          .mode = tlDCVIModeVoltage
'          .Voltage = 0 '-----------------------------------------------------------------------------------sense line only
          .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
          .SetCurrentAndRange 0.002, 0.02
          .Connect tlDCVIConnectDefault
          .Gate = True
      End With
'===============================================================================================================
'for 3. Force UVI80-N 0V========================================================================================
'===============================================================================================================
      With TheHdw.DCVI.Pins(UVI_N)
          .Gate = False
          .mode = tlDCVIModeVoltage
          .Voltage = Source_N_Voltage
          .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
          .SetCurrentAndRange 0.002, 0.02
          .Connect tlDCVIConnectDefault
          .Gate = True
      End With
'===============================================================================================================
'for 4. Use UVI80-2/UVI80-N VDIFF mode to measure the voltage across HSC_SEPVM_TEST_P/N=========================
'===============================================================================================================
'     TheHdw.DCVI.Pins(Vdiff_Pin).Connect '----------------------------------- Gate on all DCVIs
'
'     TheHdw.DCDiffMeter.Pins(Vdiff_High).LowSide.Pins = (Vdiff_Low) '----------------------- Specify the low side of the DCDiffMeter
'
'     With TheHdw.DCDiffMeter.Pins(Vdiff_High) ' ---------------------------------------------------- Set up the DCDiffMeter
'         .Connect tlDCDiffMeterConnectDefault
'         .VoltageRange = TargetVoltage
'     End With
'
'     TheHdw.Wait (TheHdwWait) ' --------------------------------------------------------------------------------- Program a wait time
'===============================================================================================================
'===============================================================================================================
'for 4.1. Use SINGLE END mode to measure the voltage across HSC_SEPVM_TEST_P/N=========================
'===============================================================================================================
    TheHdw.Wait (TheHdwWait) ' --------------------------------------------------------------------------------- Program a wait time
    MeasureVolt_P = TheHdw.DCVI.Pins(UVI_P).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
    MeasureVolt_N = TheHdw.DCVI.Pins(UVI_N).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)

'===============================================================================================================


'===============================================================================================================
'Program the DCDiffMeter to make a measurement==================================================================
'===============================================================================================================
     Dim Results As New SiteDouble
'     Results = TheHdw.DCDiffMeter.Pins(Vdiff_High).Read(tlStrobe, 100, -1, tlDCDiffMeterReadingFormatAverage)
     Results = MeasureVolt_P.Math.Subtract(MeasureVolt_N)
     TheExec.Flow.TestLimit resultVal:=Results, PinName:="HSC_SEPVM_TEST_Vdiff", Unit:=unitVolt, ForceResults:=tlForceFlow
     
     Dim ErrorValue As New PinListData: ErrorValue.AddPin ("ErrorValue")
'     ErrorValue = Results.Pins(Vdiff_High).Subtract(TargetVoltage)
     ErrorValue = Results.Subtract(TargetVoltage)
     TheExec.Flow.TestLimit resultVal:=ErrorValue.Pins, Unit:=unitVolt, ForceResults:=tlForceFlow
'===============================================================================================================
'for Gate off and disconnect the DCVI===========================================================================
'===============================================================================================================
     With TheHdw.DCVI.Pins(Disconnect_Pin)
        .mode = tlDCVIModeVoltage
        .Gate = False
        .Disconnect
     End With
     
'     With TheHdw.DCDiffMeter.Pins(Vdiff_High) ' ---------------------------------------------------- Set up the DCDiffMeter
'         .Disconnect tlDCDiffMeterConnectDefault
'         .VoltageRange = TargetVoltage
'     End With
     
     Exit Function

errHandler:
    TheExec.AddOutput "Error in SEPVM_Ref_measurement"
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function SEPVM_Ref2_Calibration(UVI_P As PinList, UVI_N As PinList, Vdiff_High As PinList, Vdiff_Low As PinList, Source_P_Voltage As Double, Source_N_Voltage As Double, _
TargetVoltage As Double, TheHdwWait As Double)
'UVI_P=HSC_SEPVM_TEST_P_Meas  @  UVI_N=HSC_SEPVM_TEST_N  @ Vdiff_High=HSC_SEPVM_P_High  @  Vdiff_Low=HSC_SEPVM_N_Low
'@ Source_P_Voltage=0.75  @ Source_N_Voltage=0  @  TargetVoltage=0.75  @ ErrorValueTarget=0.00022
    Dim ErrorValueTarget As Long
    
    ErrorValueTarget = 0.00022
    Dim SEPDSP As New DSPWave
    
    
    On Error GoTo errHandler
    Dim Vdiff_Pin As String: Vdiff_Pin = UVI_P + "," + UVI_N
        Dim Results As New SiteDouble
        Dim ErrorValue As New PinListData: ErrorValue.AddPin ("ErrorValue")
        Dim site As Variant
        Dim ForceCalibration As New SiteDouble: ForceCalibration = Source_P_Voltage ''--------let original value =0.75mV
        Dim MeasureVolt_P As New PinListData
        Dim MeasureVolt_N As New PinListData
    On Error GoTo errHandler
'===============================================================================================================
'for step3 Force UVI80-2 0.75V==================================================================================
'===============================================================================================================
    With TheHdw.DCVI.Pins(UVI_P)
        .Gate = False
        .mode = tlDCVIModeVoltage
        .Voltage = Source_P_Voltage
        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
        .SetCurrentAndRange 0.002, 0.02
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    'thehdw.Wait 100 * ms
    
'===============================================================================================================
'for step3 Force UVI80-N 0V=====================================================================================
'===============================================================================================================
    With TheHdw.DCVI.Pins(UVI_N)
        .Gate = False
        .mode = tlDCVIModeVoltage
        .Voltage = Source_N_Voltage
        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
        .SetCurrentAndRange 0.002, 0.02
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    'thehdw.Wait 100 * ms
'===============================================================================================================
'for step4. Use UVI80-2/UVI80-N VDIFF mode to measure the voltage across HSC_SEPVM_TEST_P/N=====================
'===============================================================================================================
                       
'    TheHdw.DCVI.Pins(Vdiff_Pin).Gate = False
'    TheHdw.DCVI.Pins(Vdiff_Pin).Connect ' ---------------Gate on all DCVIs
'    TheHdw.DCDiffMeter.Pins(Vdiff_High).LowSide.Pins = (Vdiff_Low) ' ---Specify the low side of the DCDiffMeter
'    TheHdw.DCVI.Pins(Vdiff_Pin).Gate = True
'
'
'    With TheHdw.DCDiffMeter.Pins(Vdiff_High) ' ---------------------------------Set up the DCDiffMeter
'        .Connect tlDCDiffMeterConnectDefault
'        .VoltageRange = Source_P_Voltage
'    End With
'    TheHdw.Wait 10 * ms
    TheHdw.Wait (TheHdwWait) '--------------------------------------------------------------- Program a wait time
    
    MeasureVolt_P = TheHdw.DCVI.Pins(UVI_P).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
    MeasureVolt_N = TheHdw.DCVI.Pins(UVI_N).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
    
'===============================================================================================================
'Program the DCDiffMeter to make a measurement==================================================================
'===============================================================================================================
'    Results = TheHdw.DCDiffMeter.Pins(Vdiff_High).Read(tlStrobe, 100, -1, tlDCDiffMeterReadingFormatAverage)
    Results = MeasureVolt_P.Math.Subtract(MeasureVolt_N)
    ErrorValue = Results.Subtract(Source_P_Voltage) '-------------------------------------------- the value of differ from 0.75mV
    
      For Each site In TheExec.sites
      
         Dim LoopCount As New SiteDouble: LoopCount = 0
         Dim step As Integer: step = 1
         
          Do While Abs(ErrorValue) > ErrorValueTarget And LoopCount < 11 And ForceCalibration < 7 '-------------------------------------------- the value of error target 220E-06
                            
                LoopCount = LoopCount + step
            
                ForceCalibration = ForceCalibration - ErrorValue '----------------------------- for loop error value add the last result
                With TheHdw.DCVI.Pins(UVI_P)
                    .Voltage = ForceCalibration '---------------------------------------------- for force last result add error value
                End With
                With TheHdw.DCVI.Pins(UVI_N)
                    .Voltage = 0
                End With
                  TheHdw.Wait (TheHdwWait) ' -------------------------------------------------------Program a wait time
                MeasureVolt_P = TheHdw.DCVI.Pins(UVI_P).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
                MeasureVolt_N = TheHdw.DCVI.Pins(UVI_N).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
                
'                Results = TheHdw.DCDiffMeter.Pins(Vdiff_High).Read(tlStrobe, 100, -1, tlDCDiffMeterReadingFormatAverage)
                Results = MeasureVolt_P.Math.Subtract(MeasureVolt_N)
                ErrorValue = Results.Subtract(TargetVoltage)
                
                If TheExec.TesterMode = testModeOffline Then '---------------------------------for avoid offline in to infinite loop
                  ErrorValue = 0.0001
                End If
                
                TheExec.Datalog.WriteComment "site " & site & " LoopCount " & LoopCount & "  ForceVoltage" & ForceCalibration & "  Results" & Results & "  ErrorValue" & ErrorValue & " "
                    
          Loop
          
      Next site
'Alarm *ForceVoltage out of range* or *Loop out of range*
'===============================================================================================================
'Print measured and error value=================================================================================
'===============================================================================================================
    
    TheExec.Flow.TestLimit resultVal:=Results, PinName:="HSC_SEPVM_TEST_Vdiff", Unit:=unitVolt, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=ErrorValue.Pins, Unit:=unitVolt, ForceResults:=tlForceFlow
    For Each site In TheExec.sites
    If LoopCount > 9 Or ForceCalibration > 2 Then
    
        TheExec.Datalog.WriteComment "site " & site & " LoopCount " & LoopCount & "  ForceVoltage" & ForceCalibration & " Alarm *ForceVoltage out of range* or *Loop out of range*"
    Else
    TheExec.Datalog.WriteComment "site " & site & " LoopCount " & LoopCount & "  ForceVoltage" & ForceCalibration & "  "
    End If
    
    Next site
'===============================================================================================================
'Gate off and disconnect the DCVI===============================================================================
'===============================================================================================================
                    
    With TheHdw.DCVI.Pins(Vdiff_Pin)
       .mode = tlDCVIModeVoltage
       .Gate = False
       .Disconnect
    End With


    Exit Function

errHandler:
    TheExec.AddOutput "Error in SEPVM_Ref2_measurement"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function ReSet_SEPVM_Ref_Level_Div()

    With TheHdw.DCVI.Pins("HSC_SEPVM_TEST_P_SRC,HSC_SEPVM_TEST_N,HSC_SEPVM_TEST_P_MEA")
        .Gate(tlDCVIGateHiZ) = False
        TheHdw.Wait 0.001
        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
        .SetCurrentAndRange 0.02, 0.02
        .Disconnect
        .mode = tlDCVIModeCurrent
    End With
    TheHdw.Wait 10 * ms
    
'===============================================================================================================
'for step3 Force UVI80-N 0V=====================================================================================
'===============================================================================================================
'    With thehdw.DCVI.Pins("HSC_SEPVM_TEST_N")
'        .Gate = False
''        .mode = tlDCVIModeVoltage
'        .Voltage = 0
'        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
''        .SetCurrentAndRange 0.002, 0.02
'        .Disconnect tlDCVIConnectDefault
''        .Gate = True
'    End With
'
'    With thehdw.DCVI.Pins("HSC_SEPVM_TEST_P_MEAS")
'        .Gate = False
''        .mode = tlDCVIModeVoltage
'        .Voltage = 0
'        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
''        .SetCurrentAndRange 0.002, 0.02
'        .Disconnect tlDCVIConnectDefault
''        .Gate = True
'    End With
    
End Function

Public Function LDO_Calibration(Optional Pat As Pattern, Optional TestSequence As String, Optional MeasV_Pins As String, Optional MeaV_WaitTime As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional TrimMethod As String, Optional TrimStepSize As Double, Optional Validating_ As Boolean)
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim pats() As String
    Dim code() As New SiteLong: ReDim code(UBound(Split(TestSequence, ",")))
    Dim BestCode() As New SiteLong: ReDim BestCode(UBound(Split(TestSequence, ",")))
    Dim vout() As New SiteDouble: ReDim vout(UBound(Split(TestSequence, ",")))
        For i = 0 To UBound(Split(TestSequence, ","))
            For Each site In TheExec.sites.Active
                code(i) = TrimStart
                vout(i) = 0
            Next site
        Next i
    Dim NumberOfMeasV As Integer: NumberOfMeasV = UBound(Split(TestSequence, ",")) + 1
    Dim PreviousNegative() As New SiteBoolean: ReDim PreviousNegative(UBound(Split(TestSequence, ",")))
    Dim PreviousPositive() As New SiteBoolean: ReDim PreviousPositive(UBound(Split(TestSequence, ",")))
    Dim DecideTrim() As New SiteBoolean: ReDim DecideTrim(UBound(Split(TestSequence, ",")))
    Dim Trim_Flag As Boolean
        For i = 0 To UBound(Split(TestSequence, ","))
            For Each site In TheExec.sites.Active
                PreviousNegative(i) = False
                PreviousPositive(i) = False
                DecideTrim(i) = False
            Next site
        Next i
    Dim blockName() As String: blockName = Split(TheExec.DataManager.instanceName, "_")
    Dim MeasValue() As New PinListData: ReDim MeasValue(NumberOfMeasV - 1)
    Dim PreviousMeasValue() As New PinListData: ReDim PreviousMeasValue(NumberOfMeasV - 1)
    Dim BestVal() As New PinListData: ReDim BestVal(NumberOfMeasV - 1)
    Dim StepCount As Long: StepCount = 0
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim PreviousTargetCompare() As New SiteDouble: ReDim PreviousTargetCompare(UBound(Split(TestSequence, ",")))
    Dim TrimStoreName_Array() As String: TrimStoreName_Array = Split(TrimStoreName, ",")
    Dim PinName As String
    Dim TempVal As Integer
    Dim FinalTrimCode() As New DSPWave: ReDim FinalTrimCode(UBound(TrimStoreName_Array))
    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
    
    Call ProcessInputToGLB(Pat, TestSequence, True, , , , , MeasV_Pins, , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , MeaV_WaitTime)
    
    For i = 0 To NumberOfMeasV - 1
        For j = 0 To UBound(Split(MeasV_Pins, ","))
            PreviousMeasValue(i).AddPin (Split(MeasV_Pins, ",")(j))
            PreviousMeasValue(i).Pins(Split(MeasV_Pins, ",")(j)).Value = 0
            MeasValue(i).AddPin (Split(MeasV_Pins, ",")(j))
            MeasValue(i).Pins(Split(MeasV_Pins, ",")(j)).Value = 0
            BestVal(i).AddPin (Split(MeasV_Pins, ",")(j))
            BestVal(i).Pins(Split(MeasV_Pins, ",")(j)).Value = 0
        Next j
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
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value)
            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", 24000000#)
            Else
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            End If
        End If
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet Pat.Value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    TrimCodeValue_Min = 0
    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
  

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
    Call LDO_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array(), MeaV_WaitTime)

    BestCode = code
    For Each site In TheExec.sites.Active
        For i = 0 To NumberOfMeasV - 1
            BestVal(i) = MeasValue(i)
        Next i
    Next site
    For i = 0 To UBound(Split(TestSequence, ","))
        If LCase(TrimMethod) = "linearsearch" Then
            For Each site In TheExec.sites.Active
                If vout(i).compare(LessThan, TrimTarget) Then
                    code(i) = code(i) + 1
                ElseIf vout(i).compare(GreaterThan, TrimTarget) Then
                    code(i) = code(i) - 1
                End If
                DecideTrim(i) = True
            Next site
        Else
            For Each site In TheExec.sites.Active
                code(i) = code(i) + Fix((TrimTarget - vout(i)) / TrimStepSize)
'                If Fix((TrimTarget - vout(i)) / TrimStepSize) <> 0 Then: DecideTrim(i) = True
                DecideTrim(i) = True
            Next site
        End If
    Next i
StartTrim:
    For i = 0 To UBound(Split(TestSequence, ","))
        If i = 0 Then
            Trim_Flag = DecideTrim(i).Any(True)
        Else
            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
        End If
    Next i
        If Trim_Flag Then
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

            Call LDO_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array(), MeaV_WaitTime)
        End If
    For k = 0 To UBound(Split(TestSequence, ","))
        For Each site In TheExec.sites.Active
            If DecideTrim(k) Then
                If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                    BestCode = code
'                    For i = 0 To NumberOfMeasV - 1
                        For j = 0 To UBound(Split(MeasV_Pins, ","))
                            BestVal(k).Pins(j).Value = MeasValue(k).Pins(j).Value
                        Next j
'                    Next i
                    DecideTrim(k) = False
                ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
                    code(k) = TrimCodeValue_Max
                    DecideTrim(k) = True
                ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
                    code(k) = TrimCodeValue_Min
                    DecideTrim(k) = True
                ElseIf code(k).compare(EqualTo, TrimCodeValue_Max) Or code(k).compare(EqualTo, TrimCodeValue_Min) Then
                    BestCode(k) = code(k)
'                    For i = 0 To NumberOfMeasV - 1
                        For j = 0 To UBound(Split(MeasV_Pins, ","))
                            BestVal(k).Pins(j).Value = MeasValue(k).Pins(j).Value
                        Next j
'                    Next i
                    DecideTrim(k) = False
                ElseIf vout(k).compare(LessThan, TrimTarget) And PreviousPositive(k) Then
    '                If vout(k).Subtract(TrimTarget).Abs > PreviousTargetCompare(k) Then
                        BestCode(k) = code(k) + 1
'                        For i = 0 To NumberOfMeasV - 1
                            For j = 0 To UBound(Split(MeasV_Pins, ","))
                                BestVal(k).Pins(j).Value = PreviousMeasValue(k).Pins(j).Value
                            Next j
'                        Next i
    '                Else
    '                    BestCode(k) = code(k)
    '                    For i = 0 To NumberOfMeasV - 1
    '                        For j = 0 To UBound(Split(MeasV_PinS, ","))
    '                            BestVal(i).Pins(j).Value = MeasValue(i).Pins(j).Value
    '                        Next j
    '                    Next i
    '                End If
                    DecideTrim(k) = False
                ElseIf vout(k).compare(LessThan, TrimTarget) And Not (PreviousPositive(k)) Then
                    code(k) = code(k) + 1
                    PreviousNegative(k) = True
                    PreviousTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
'                    For i = 0 To NumberOfMeasV - 1
                        For j = 0 To UBound(Split(MeasV_Pins, ","))
                            PreviousMeasValue(k).Pins(j).Value = MeasValue(k).Pins(j).Value
                        Next j
'                    Next i
                    DecideTrim(k) = True
                ElseIf vout(k).compare(GreaterThan, TrimTarget) And PreviousNegative(k) Then
    '                If vout(k).Subtract(TrimTarget).Abs > PreviousTargetCompare(k) Then
    '                    BestCode(k) = code(k) - 1
    '                    For i = 0 To NumberOfMeasV - 1
    '                        For j = 0 To UBound(Split(MeasV_PinS, ","))
    '                            BestVal(i).Pins(j).Value = PreviousMeasValue(i).Pins(j).Value
    '                        Next j
    '                    Next i
    '                Else
    '                    If DecideTrim(k) Then BestCode(k) = code(k)
                        BestCode(k) = code(k)
'                        For i = 0 To NumberOfMeasV - 1
                            For j = 0 To UBound(Split(MeasV_Pins, ","))
    '                            If DecideTrim(k) = True Then BestVal(i).Pins(j).Value = MeasValue(i).Pins(j).Value
                                BestVal(k).Pins(j).Value = MeasValue(k).Pins(j).Value
                            Next j
'                        Next i
    '                End If
                    DecideTrim(k) = False
                ElseIf vout(k).compare(GreaterThan, TrimTarget) And Not (PreviousNegative(k)) Then
                    code(k) = code(k) - 1
                    PreviousPositive(k) = True
                    PreviousTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
'                    For i = 0 To NumberOfMeasV - 1
                        For j = 0 To UBound(Split(MeasV_Pins, ","))
                            PreviousMeasValue(k).Pins(j).Value = MeasValue(k).Pins(j).Value
                        Next j
'                    Next i
                    DecideTrim(k) = True
                End If
            End If
        Next site
    Next k
    
    For i = 0 To UBound(Split(TestSequence, ","))
        If i = 0 Then
            Trim_Flag = DecideTrim(i).Any(True)
        Else
            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
        End If
    Next i
    
    If Trim_Flag Then GoTo StartTrim

        
    For i = 0 To NumberOfMeasV - 1
        For j = 0 To UBound(Split(MeasV_Pins, ","))
            PinName = Split(MeasV_Pins, ",")(j)
            TestNameInput = Report_TName_From_Instance("V", PinName, "", i, 0)
            TheExec.Flow.TestLimit resultVal:=BestVal(i), Tname:=TestNameInput, PinName:=Split(MeasV_Pins, ",")(j), ForceResults:=tlForceFlow
        Next j
    Next i
    For i = 0 To UBound(Split(TestSequence, ","))
        TestNameInput = Report_TName_From_Instance("C", "", blockName(0) & "Trim", i, 0)
        TheExec.Flow.TestLimit resultVal:=BestCode(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    For i = 0 To UBound(Split(TestSequence, ","))
        For Each site In TheExec.sites
            TempVal = BestCode(i)
            For j = 0 To TrimCodeSize - 1
                FinalTrimCode_Array(j) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next j
            FinalTrimCode(i).Data = FinalTrimCode_Array
        Next site
        Call AddStoredCaptureData(TrimStoreName_Array(i), FinalTrimCode(i))
    Next i
    DebugPrintFunc Pat.Value
    
    Call HardIP_WriteFuncResult(, , Inst_Name_Str)
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in LDO_Calibration"
    If AbortTest Then Exit Function Else Resume Next
End Function


