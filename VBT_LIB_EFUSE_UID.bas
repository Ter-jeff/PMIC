Attribute VB_Name = "VBT_LIB_EFUSE_UID"
Option Explicit

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_UIDBlankChk(ReadPatSet As Pattern, PinRead As PinList, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)
    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UIDBlankChk"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(ReadPatSet, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim m_stage As String
    Dim i As Long
    Dim ResultFlag As New SiteLong
    Dim testName As String, PinName As String
    Dim PrintSiteVarResult As String
    Dim SiteVarValue As Long
    Dim Block1Sum As Double, Block2Sum As Double
    Dim allBlank As New SiteBoolean

    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, , InitPinsLo
    
    '****** Initialize Site Varaible ******
    'Assume Max site number
    Dim m_siteVar As String
    m_siteVar = "UIDChk_Var"
    For Each site In TheExec.sites.Existing
        TheExec.sites(site).SiteVariableValue(m_siteVar) = -1
    Next site

    '================================================
    '=  Setup HRAM/DSSC capture cycles              =
    '================================================
    Dim SignalCap As String, CapWave As New DSPWave

    SignalCap = "Margin_DigCap"
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, UIDReadCycle, CapWave
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)
    ''TheHdw.Digital.Patgen.HaltWait
    
    'auto_eFuse_DSSC_ReadDigCap_32bits UIDReadCycle, PinRead.Value, gS_SingleStrArray, CapWave, allblank

    ''''20151229 update, Using 1st category as the m_stage (uid)
    m_stage = LCase(UIDFuse.Category(0).Stage)
    
    Call UpdateDLogColumns(gI_UID_catename_maxLen)
    
    ''''201811XX update
If (gB_eFuse_newMethod = True) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''----------------------------------------------------
    ''''201812XX New Method by DSPWave
    ''''----------------------------------------------------
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong
    Dim blank_stage As New SiteBoolean

    m_Fusetype = eFuse_UID
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_stage = True

    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>

    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    
    ''''----------------------------------------------------
'    If (TheExec.TesterMode = testModeOffline) Then
'        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
'        ''''Here MarginRead it only read the programmed Site.
'        'Dim Temp_USO As New DSPWave
'        'CapWave.CreateConstant 0, UIDReadCycle, DspLong
'
'        'gL_eFuse_Sim_Blank = 0
'
'        Call RunDSP.eFuse_SingleBitWave2CapWave32Bits(eFuse_UID, CapWave)
'    End If

    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, CapWave, m_FBC, blank_stage, allBlank)

    If (blank_stage.Any(False) = True) Then
        ''''''''if there is any site which is non-blank, then decode to gDW_CFG_Read_Decimal_Cate [check later]
        ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
        ''''it will be present by Hex and Binary compare with the limit later on.
        ''''<NOTICE> gDW_CFG_Read_Decimal_Cate has been decided in non-blank and/or MarginRead
        ''Call auto_eFuse_setReadData(eFuse_CFG, gDW_CFG_Read_DoubleBitWave, gDW_CFG_Read_Decimal_Cate)
        'Call auto_eFuse_setReadData(eFuse_CFG)
        
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(eFuse_UID)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_UID)

        Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_UID, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_UID, False, gB_eFuse_printReadCate)
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
    m_SiteVarValue = blank_stage
    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
    'If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only

    For Each site In TheExec.sites
        If (auto_eFuse_GetAllPatTestPass_Flag("UID") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    gL_UID_FBC = m_FBC
    
    Call UpdateDLogColumns(gI_UID_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "UID_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "UID_Blank_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value

Exit Function

End If
    
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151229 New
''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_UIDBlankChk_byStage(ReadPatSet As Pattern, PinRead As PinList, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)
    
'''On Error GoTo errHandler
'''    Dim funcName As String:: funcName = "auto_UIDBlankChk_byStage"
'''
'''    ''''----------------------------------------------------------------------------------------------------
'''    ''''<Important>
'''    ''''Must be put before all implicit array variables, otherwise the validation will be error.
'''    '==================================
'''    '=  Validate/Load Read patterns   =
'''    '==================================
'''    ''''20161114 update to Validate/load pattern
'''    Dim ReadPatt As String
'''    If (auto_eFuse_PatSetToPat_Validation(ReadPatSet, ReadPatt, Validating_) = True) Then Exit Function
'''    ''''----------------------------------------------------------------------------------------------------
'''
'''    Dim Site As Variant
'''    Dim m_stage As String
'''    Dim m_jobinStage_flag As Boolean
'''    Dim i As Long
'''    Dim PatMargin As String
'''    Dim ResultFlag As New SiteLong
'''    Dim testName As String, PinName As String
'''    Dim PrintSiteVarResult As String
'''    Dim SiteVarValue As Long
'''    Dim SingleBitArray() As Long
'''    Dim DoubleBitArray() As Long
'''    ''ReDim SingleBitArray(UIDTotalBits - 1)
'''    ''Dim SingleBitSum As Long
'''    Dim SingleDoubleFBC As Long
'''    Dim blank_stage As New SiteBoolean
'''
'''    Dim Block1Sum As Double, Block2Sum As Double
'''    Dim allBlank As New SiteBoolean
'''
'''    '=============================================================================
'''    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
'''    '=============================================================================
'''    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, , InitPinsLo
'''
'''    '****** Initialize Site Varaible ******
'''    'Assume Max site number
'''    Dim m_siteVar As String
'''    m_siteVar = "UIDChk_Var"
'''    For Each Site In TheExec.Sites.Existing
'''        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = -1
'''    Next Site
'''
'''    '================================================
'''    '=  Setup HRAM/DSSC capture cycles              =
'''    '================================================
'''    Dim SignalCap As String, CapWave As New DSPWave
'''
'''    SignalCap = "SignalCapture"
'''    auto_eFuse_DSSC_DigCapSetup PatMargin, PinRead, SignalCap, UIDReadCycle, CapWave
'''    Call TheHdw.Patterns(PatMargin).Test(pfAlways, 0, tlResultModeDomain)
'''
'''    TheHdw.Digital.Patgen.HaltWait
'''
'''    'auto_eFuse_DSSC_ReadDigCap_32bits UIDReadCycle, PinRead.Value, gS_SingleStrArray(), CapWave, allblank
'''
'''    ''''20151229 update, Using 1st category as the m_stage (uid)
'''    m_stage = LCase(UIDFuse.Category(0).Stage)
'''
'''    ''''it's used to identify if the Job Name is existed in the UID portion of the eFuse BitDef table.
'''    m_jobinStage_flag = auto_eFuse_JobExistInStage("UID", True) ''''<MUST>
'''
'''    Call UpdateDLogColumns(gI_UID_catename_maxLen)
'''
'''    For Each Site In TheExec.Sites
'''        ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'''        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'''            For i = 0 To UIDReadCycle - 1
'''                gS_SingleStrArray(i, Site) = StrReverse(gS_SingleStrArray(i, Site))
'''            Next i
'''        End If
'''
'''        ''''depends on the specific Job to judge if it's blank on the specific stage
'''        testName = "UID_BlankChk_" + UCase(gS_JobName)
'''
'''        SingleDoubleFBC = 0 ''''init
'''        ''''Decompose Read Cycle String to get SingleBitArray and SingleBitSum
'''        ''Call auto_Decompose_StrArray_to_BitArray("UID", gS_SingleStrArray, SingleBitArray, SingleBitSum)
'''        Call auto_OR_2Blocks("UID", gS_SingleStrArray, SingleBitArray, DoubleBitArray) ''''calc gL_UID_FBC
'''        Call auto_eFuse_BlankChk_FBC_byStage("UID", SingleBitArray, blank_stage, SingleDoubleFBC)
'''
'''        ''''20151229 add
'''        ''''=============== Start of Simulated Data ===============
'''        If (TheExec.TesterMode = testModeOffline) Then
'''            If (gS_JobName = m_stage) Then
'''                blank_stage(Site) = True
'''            Else
'''                blank_stage(Site) = False
'''            End If
'''        End If
'''        ''''===============   End of Simulated Data ===============
'''
'''       If blank_stage(Site) = False Then  'If not blank
'''
'''            '*** In here it means the eFuse is not blank ****
'''            If (gS_JobName = m_stage) Then
'''                '*** Check the summation for UID re-test flow     ****
'''                Call Cal_UIDChkSum(gS_SingleStrArray, Block1Sum, Block2Sum)
'''                Block1Sum = Block1Sum / gL_UIDCodeBitWidth
'''                Block2Sum = Block2Sum / gL_UIDCodeBitWidth
'''
'''                If (gS_EFuse_Orientation = "UP2DOWN") Then
'''                    TheExec.Flow.TestLimit resultVal:=Block1Sum, lowval:=UID_ChkSum_LoLimit, hival:=UID_ChkSum_HiLimit, Tname:="UID_ChksumU"
'''                    TheExec.Flow.TestLimit resultVal:=Block2Sum, lowval:=UID_ChkSum_LoLimit, hival:=UID_ChkSum_HiLimit, Tname:="UID_ChksumD"
'''
'''                ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
'''                    ''Block1Sum -> Right Side CheckSum
'''                    ''Block2Sum -> Left  Side CheckSum
'''                    TheExec.Flow.TestLimit resultVal:=Block1Sum, lowval:=UID_ChkSum_LoLimit, hival:=UID_ChkSum_HiLimit, Tname:="UID_ChksumR"
'''                    TheExec.Flow.TestLimit resultVal:=Block2Sum, lowval:=UID_ChkSum_LoLimit, hival:=UID_ChkSum_HiLimit, Tname:="UID_ChksumL"
'''
'''                ElseIf (gS_EFuse_Orientation = "SingleUp") Then ''''only one block
'''                    TheExec.Flow.TestLimit resultVal:=Block1Sum, lowval:=UID_ChkSum_LoLimit, hival:=UID_ChkSum_HiLimit, Tname:="UID_Chksum"
'''
'''                ''''The below is reserved.
'''                ElseIf (gS_EFuse_Orientation = "SingleDown") Then
'''                ElseIf (gS_EFuse_Orientation = "SingleRight") Then
'''                ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
'''                End If
'''
'''                ''''testName = "UID_Blank"
'''                If ((gS_EFuse_Orientation = "SingleUp")) Then
'''                    If (Block1Sum >= UID_ChkSum_LoLimit) And (Block1Sum <= UID_ChkSum_HiLimit) Then
'''                        ResultFlag(Site) = 0 'Pass
'''                        PinName = "Pass"
'''                        SiteVarValue = 2
'''                    Else
'''                        ResultFlag(Site) = 1 'Fail
'''                        PinName = "Fail"
'''                        SiteVarValue = 0
'''                    End If
'''                Else
'''                    'This stand for 20% (anh ask to modify it on 2013/11/26)
'''                    If (Block1Sum >= UID_ChkSum_LoLimit) And (Block2Sum >= UID_ChkSum_LoLimit) And _
'''                       (Block1Sum <= UID_ChkSum_HiLimit) And (Block2Sum <= UID_ChkSum_HiLimit) Then
'''                        ResultFlag(Site) = 0 'Pass
'''                        PinName = "Pass"
'''                        SiteVarValue = 2
'''                    Else
'''                        ResultFlag(Site) = 1 'Fail
'''                        PinName = "Fail"
'''                        SiteVarValue = 0
'''                    End If
'''                End If
'''            Else
'''                If (SingleDoubleFBC = 0) Then
'''                    ResultFlag(Site) = 0    ''Pass Blank check
'''                    PinName = "Pass"
'''                    SiteVarValue = 2
'''                Else
'''                    ResultFlag(Site) = 1    ''Fail Blank check
'''                    PinName = "Fail"
'''                    SiteVarValue = 0
'''                End If
'''            End If
'''
'''        Else ' True means this Efuse is balnk in this JobStage.
'''            If (gS_JobName <> m_stage And allBlank(Site) = True) Then
'''                ''''Becuase it should have been programmed in m_stage (uid),
'''                ''''so the allBlank() must NOT be True
'''                ResultFlag(Site) = 1   ''Fail then NO fuse
'''                PinName = "Fail"
'''                SiteVarValue = 0
'''            Else
'''                ''''possible cases::
'''                ''''(1) gS_JobName==m_stage And allBlank(Site)=True  (CP1 fresh_Die case)
'''                ''''(2) gS_JobName==m_stage And allBlank(Site)=False =>blank_stage(Site)=False (CP1 retestDie case), it has been processed as above
'''                ''''(3) gS_JobName<>m_stage And allBlank(Site)=False
'''                If (auto_eFuse_GetAllPatTestPass_Flag("UID") = False) Then
'''                    ResultFlag(Site) = 1   ''Fail then NO fuse
'''                    PinName = "Fail"
'''                    SiteVarValue = 0
'''                Else
'''                    ResultFlag(Site) = 0   ''Pass Blank check
'''                    PinName = "Pass"
'''                    SiteVarValue = 1
'''
'''                    ''''<MUST>
'''                    ''''it's used to identify if the Job Name is existed in the UID portion of the eFuse BitDef table.
'''                    If (m_jobinStage_flag = False) Then
'''                        ''''<Important> Then it will NOT go WritebyStage to let the user confusion.
'''                        SiteVarValue = 2
'''                    End If
'''                End If
'''            End If
'''        End If
'''
'''        If (SiteVarValue = 1) Then
'''            ''''20160531
'''            ''''SiteVarValue = 1, do the pre-decode for CRC Pgm if needed
'''            Call auto_Decode_UIDBinary_Data(DoubleBitArray, False)
'''        End If
'''
'''        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'''
'''        If (False) Then
'''            PrintSiteVarResult = "Site (" + CStr(Site) + ") " + m_siteVar + " = " + CStr(SiteVarValue)
'''            TheExec.Datalog.WriteComment PrintSiteVarResult & vbCrLf
'''        End If
'''
'''        ''Binning out
'''        TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowval:=1, hival:=2, Tname:=m_siteVar, PinName:="Value"
'''        TheExec.Flow.TestLimit resultVal:=ResultFlag, lowval:=0, hival:=0, Tname:=testName, PinName:=PinName
'''
'''    Next Site
'''
'''    Call UpdateDLogColumns__False
'''
'''    DebugPrintFunc ReadPatSet.Value
'''
'''Exit Function
'''
'''errHandler:
'''    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'''    If AbortTest Then Exit Function Else Resume Next

End Function



''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_UID_Read_by_OR_2Blocks(ReadPatSet As Pattern, PinRead As PinList, _
                    condstr As String, _
                    Optional catename_grp As String = "", _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UID_Read_by_OR_2Blocks"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(ReadPatSet, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim eFuse_Pgm_Bit() As Long
    Dim i As Long, j As Long, k As Long
    Dim m_tmpStr As String

    Dim FailCnt As New SiteLong
    Dim testName As String
    Dim blank As New SiteBoolean
    Dim cycleNum As Long, BitPerCycle As Long

    Dim SingleStrArray() As String
    Dim DoubleBitArray() As Long
    Dim SingleBitArray() As Long
    
    cycleNum = UIDReadCycle: BitPerCycle = UIDBitsPerCycle
    
    ''ReDim SingleStrArray(CycleNum - 1)
    ''ReDim DoubleBitArray(UIDBitsPerBlockUsed - 1)
    ''ReDim SingleBitArray(UIDTotalBits - 1)

    condstr = LCase(condstr) ''''<MUST>

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
  
    '====================================================
    '=  1. Define the Bit length of Programming         =
    '=  2. Define the cycle number of Compare Pattern   =
    '====================================================
    ReDim eFuse_Pgm_Bit(UIDTotalBits - 1)

    '================================================
    '=  In Fiji, C651 ask we to OR block 1          =
    '=  and block2 before compare with programmed   =
    '=  data (confirmed on 6/28 morning meeting     =
    '================================================
    Dim SignalCap As String, CapWave As New DSPWave
    SignalCap = "SignalCapture"

    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, UIDReadCycle, CapWave
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0)
    testName = "UID_ORMarginRead_" + UCase(condstr)
    
''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    Dim m_Fusetype As eFuseBlockType
    Dim m_FBC As New SiteLong       ''''it means Read Bits (Single vs Double)=>Sigle-Double-Bits Check
    Dim m_cmpResult As New SiteLong ''''it means the comparison of the Read and Pgm Bits
    Dim m_bitFlag_mode As Long

    m_Fusetype = eFuse_UID
    m_FBC = -1       ''''init to failure
    m_cmpResult = -1 ''''init to failure

    ''''--------------------------------------------------------------------------
    '''' Offline Simulation Start                                                |
    ''''--------------------------------------------------------------------------
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        Call rundsp.eFuse_SingleBitWave2CapWave32Bits(eFuse_UID, CapWave)
        Call auto_eFuse_print_capWave32Bits(eFuse_UID, CapWave, False) ''''True to print out
    End If
    ''''--------------------------------------------------------------------------
    '''' Offline Simulation End                                                  |
    ''''--------------------------------------------------------------------------

    If (condstr = "cp1_early") Then
        m_bitFlag_mode = 0
    ElseIf (condstr = "stage") Then
        m_bitFlag_mode = 1
    ElseIf (condstr = "all") Then
        m_bitFlag_mode = 1 ''''update later, was 2
    Else
        ''''default, here it prevents any typo issue
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
        m_FBC = -1
        m_cmpResult = -1
    End If

    Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, m_cmpResult)

    gL_UID_FBC = m_FBC

    ''''''[NOTICE] Decode and Print have moved to SingleDoubleBit()

    Call UpdateDLogColumns(gI_UID_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=m_cmpResult, lowVal:=0, hiVal:=0
    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value

Exit Function

End If
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_UIDSingleDoubleBit(Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UIDSingleDoubleBit"

    Dim Block1Sum As Double
    Dim Block2Sum As Double
    Dim site As Variant
    Dim SingleBitArray() As Long, singleBitSum As Long
    Dim DoubleBitArray() As Long, doubleBitSum As Long
    Dim m_siteVar As String
    m_siteVar = "UIDChk_Var"
    
    
 ''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''gDW_CFG_Read_Decimal_Cate
    ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
    ''''it will be present by Hex and Binary compare with the limit later on.
    ''''<NOTICE> gDW_CFG_Read_Decimal_Cate has been decided in non-blank and/or MarginRead
    ''Call auto_eFuse_setReadData(eFuse_CFG, gDW_CFG_Read_DoubleBitWave, gDW_CFG_Read_Decimal_Cate) ''''was
    'Call auto_eFuse_setReadData(eFuse_CFG)
    
    ''''201901XX New for TTR/PTE improvement
    Call auto_eFuse_setReadData_forSyntax(eFuse_UID)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_UID)
    
    ''''All the read action has been down in blank and/or MarginRead
    ''''gDW_CFG_Read_cmpsgWavePerCyc used to display the cmpare result (2-bit mode)
    Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_UID, gB_eFuse_printBitMap)
    If (gS_JobName = "cp1_early") Then
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_UID, True, gB_eFuse_printReadCate)
    Else
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_UID, False, gB_eFuse_printReadCate)
    End If
    
    
    ''''Print CRC calcBits information
    Dim m_crcBitWave As New DSPWave
    Dim mS_hexStr As New SiteVariant
    Dim mS_bitStrM As New SiteVariant
    Dim m_debugCRC As Boolean
    Dim m_cnt As Long
    'Dim m_siteVar As String
    m_siteVar = "UIDChk_Var"
    m_debugCRC = False

    ''''<MUST> Initialize
    gS_UID_Read_calcCRC_hexStr = "0x00000000"
    gS_UID_Read_calcCRC_bitStrM = ""
    CRC_Shift_Out_String = ""
    If (auto_eFuse_check_Job_cmpare_Stage(gS_UID_CRC_Stage) >= 0) Then
        Call rundsp.eFuse_Read_to_calc_CRCWave(eFuse_UID, gL_UID_CRC_BitWidth, m_crcBitWave)
        TheHdw.Wait 1# * ms ''''check if it needs

        If (m_debugCRC = False) Then
            ''''Here get gS_CFG_Read_calcCRC_hexStr for the syntax check
            Call auto_eFuse_bitWave_to_binStr_HexStr(m_crcBitWave, gS_UID_Read_calcCRC_bitStrM, gS_UID_Read_calcCRC_hexStr, True, m_debugCRC)
        Else
            ''''m_debugCRC=True => Debug purpose for the print
            TheExec.Datalog.WriteComment "------Read CRC Category Result------"
            Call auto_eFuse_bitWave_to_binStr_HexStr(m_crcBitWave, gS_UID_Read_calcCRC_bitStrM, gS_UID_Read_calcCRC_hexStr, True, m_debugCRC)
            TheExec.Datalog.WriteComment ""

            ''''[Pgm CRC calcBits] only gS_CFG_CRC_Stage=Job and CFGChk_Var=1
            If (gS_UID_CRC_Stage = gS_JobName) Then
                m_cnt = 0
                For Each site In TheExec.sites
                    If (TheExec.sites(site).SiteVariableValue(m_siteVar) = 1) Then
                        If (m_cnt = 0) Then TheExec.Datalog.WriteComment "------Pgm CRC calcBits------"
                        Call auto_eFuse_bitWave_to_binStr_HexStr(gDW_Pgm_BitWaveForCRCCalc, mS_bitStrM, mS_hexStr, False, m_debugCRC)
                        m_cnt = m_cnt + 1
                    End If
                Next site
                If (m_cnt > 0) Then TheExec.Datalog.WriteComment ""
            End If

            TheExec.Datalog.WriteComment "------Read CRC calcBits------"
            Call auto_eFuse_bitWave_to_binStr_HexStr(gDW_Read_BitWaveForCRCCalc, mS_bitStrM, mS_hexStr, False, m_debugCRC)
            TheExec.Datalog.WriteComment ""

            ''''----------------------------------------------------------------------------------------------------------------
            ''''The Below is the original format from Orange team
            ''''----------------------------------------------------------------------------------------------------------------
            ''' 20170623 Per discussion with Jack, update as the below
            ''''means that the CRC codes have been blown into the Fuse,
            ''''we MUST display the below bit string for the tool calculation.
''            CRC_Shift_Out_String = mS_bitStrM
''            For Each Site In TheExec.Sites
''                TheExec.Datalog.WriteComment "CFG bit string for CRC calculation on Site " & CStr(Site) & " (MSB to LSB)= " & CStr(CRC_Shift_Out_String(Site))   ''20170623
''                TheExec.Datalog.WriteComment Chr(13) & "Totally " & CStr(Len(CRC_Shift_Out_String(Site))) & " bits for CRC calculation"
''            Next Site
''            TheExec.Datalog.WriteComment ""
            ''''----------------------------------------------------------------------------------------------------------------
        End If

    End If
    
    ''''----------------------------------------------------------------------------------------------
'    If (gS_JobName <> "cp1_early") Then
'        For Each Site In TheExec.Sites
'            DoubleBitArray = gDW_UID_Read_DoubleBitWave.Data
'
'            gS_UID_Direct_Access_Str(Site) = "" ''''is a String [(bitLast)......(bit0)]
'
'            For i = 0 To UBound(DoubleBitArray)
'                gS_UID_Direct_Access_Str(Site) = CStr(DoubleBitArray(i)) + gS_UID_Direct_Access_Str(Site)
'            Next i
'            ''TheExec.Datalog.WriteComment "gS_CFG_Direct_Access_Str=" + CStr(gS_CFG_Direct_Access_Str(Site))
'
'            ''''20161114 update for print all bits (DTR) in STDF
'            Call auto_eFuse_to_STDF_allBits("UID", gS_UID_Direct_Access_Str(Site))
'        Next Site
'    End If
    ''''----------------------------------------------------------------------------------------------

    ''''gL_CFG_FBC has been check in Blank/MarginRead
    Call UpdateDLogColumns(gI_UID_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=gL_UID_FBC, lowVal:=0, hiVal:=0, Tname:="MON_FBCount_" + UCase(gS_JobName) '2d-s=0
    Call UpdateDLogColumns__False

Exit Function

End If
    
    
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_UIDWrite_byCondition(WritePattSet As Pattern, PinWrite As PinList, _
                    PwrPin As String, vpwr As Double, _
                    condstr As String, _
                    Optional catename_grp As String, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UIDWrite_byCondition"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Write patterns  =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim WritePatt As String
    If (auto_eFuse_PatSetToPat_Validation(WritePattSet, WritePatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim i As Long
    Dim site As Variant
    Dim Expand_eFuse_Pgm_Bit() As Long, eFuse_Pgm_Bit() As Long, eFusePatCompare() As String
    Dim SegmentSize As Long
    Dim DigSrcSignalName As String
    Dim Expand_Size As Long

    ReDim gL_UIDFuse_Pgm_Bit(TheExec.sites.Existing.Count - 1, UIDTotalBits - 1)

    '=======================================
    '=  Step1. Read eFuse default values   =
    '=======================================

    condstr = LCase(condstr) ''''<MUST>

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

    TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName

    Call TurnOnEfusePwrPins(PwrPin, vpwr)

    '========================================================
    '=  Step3. Wrap up data for Efuse programming by DSSC   =
    '========================================================
    Expand_Size = CLng(UIDTotalBits) * CLng(UIDWriteBitExpandWidth) 'Because there are repeat cycle in C651 pattern, we have to create multiple DSSC data.
    'However,UIDWriteBitExpandWidth depends on how many repeat cycle in pin of STROBE.
    ReDim eFuse_Pgm_Bit(UIDTotalBits - 1)
    ReDim Expand_eFuse_Pgm_Bit(Expand_Size - 1)
    ReDim eFusePatCompare(UIDReadCycle - 1)

    'This subroutine is designed to (1). pack programming bit for DSSC (Expand_eFuse_Pgm_Bit())
    '                               (2). produce a output comparing data (eFusePatCompare())

    DigSrcSignalName = "DigSrcSignal"
    
    
''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''201808XX update
'    If (TheExec.TesterMode = testModeOffline) Then
'        If (condStr <> "cp1_early") Then
'            For Each Site In TheExec.Sites
'                Call eFuseENGFakeValue_Sim
'            Next Site
'        End If
'    End If

    Dim m_stage As String
    Dim m_catename As String
    Dim m_catenameVbin As String
    Dim m_crc_idx As Long
    Dim m_calcCRC As New SiteLong
    
    Dim m_cmpStage As String
    Dim m_pgmRes As New SiteLong
    Dim m_defreal As String
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_vbinResult As New SiteDouble
    Dim m_pgmWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_pgmDigSrcWave As New DSPWave
    
    m_pgmWave.CreateConstant 0, gDL_BitsPerBlock, DspLong ''''gDL_BitsPerBlock = EConfigBitPerBlockUsed
    ''''<MUST and Importance>
    For Each site In TheExec.sites
        gDW_Pgm_RawBitWave = m_pgmWave.Copy
    Next site

'    If (condStr = "cp1_early") Then
'        m_cmpStage = "cp1_early"
'    Else
'        ''''condStr = "stage"
'        m_cmpStage = gS_JobName
'    End If
    
    ''''Only composite case "real or bincut" PgmBits Wave per Stage requirement
    For i = 0 To UBound(UIDFuse.Category)
        With UIDFuse.Category(i)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_defreal = LCase(.Default_Real)
        End With
        
        'If (m_stage = gS_JobName) Then ''''was If (m_stage = m_cmpStage) Then
            If (m_algorithm = "crc") Then
                m_crc_idx = i
                ''''special handle on the next process
                ''''skip it here
            ElseIf (m_defreal = "real" Or m_defreal = "bincut") Then
'                If (m_algorithm = "vddbin") Then
'                    Call auto_eFuse_Vddbin_bincut_setWriteData(eFuse_UID, i)
'                End If
                ''''---------------------------------------------------------------------------
                With UIDFuse.Category(i)
                    ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
                    Call rundsp.eFuse_updatePgmWave_byCategory(.BitIndexWave, .Write.BitArrWave)
                End With
                ''''---------------------------------------------------------------------------
            End If
        'Else
            ''''doNothing
        'End If
    Next i
    
    ''''process CRC bits calculation
    If (gS_UID_CRC_Stage = gS_JobName) Then
        Dim mSL_bitwidth As New SiteLong
        mSL_bitwidth = gL_UID_CRC_BitWidth
        ''''CRC case
        With UIDFuse.Category(m_crc_idx)
            ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
            Call rundsp.eFuse_updatePgmWave_CRCbits(eFuse_UID, mSL_bitwidth, .BitIndexWave)
        End With
    End If
    
    ''''composite effective PgmBits per Stage requirement
    m_pgmRes = 0
    If (m_cmpStage = "cp1_early") Then
        ''''condStr = "cp1_early"
        Call auto_eFuse_Gen_PgmBitSrcWave_onlyDeaultBits(eFuse_UID, m_pgmDigSrcWave, m_pgmRes)
    Else
        ''''condStr = "stage"
        ''''Here Parameter bitFlag_mode=1 (Stage Bits)
        m_calcCRC = 0
        Call rundsp.eFuse_Gen_PgmBitSrcWave(eFuse_UID, 1, m_pgmDigSrcWave, m_pgmRes)
    End If
    TheHdw.Wait 1# * ms
    
    Call UpdateDLogColumns(gI_UID_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=m_pgmRes, lowVal:=1, hiVal:=1, Tname:="UID_PGM_" + UCase(gS_JobName)
    Call UpdateDLogColumns__False

    ''''use global DSP variable
    Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_UID_Pgm_SingleBitWave, gDL_TotalBits, gDL_ReadCycles, gB_eFuse_printBitMap)
    'Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_UID_Pgm_SingleBitWave, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_UID, gDW_UID_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)
    
    If (m_cmpStage = "cp1_early") Then
        ''''if it's same values on all Sites to save TT and improve PTE
        Call eFuse_DSSC_SetupDigSrcWave_allSites(WritePatt, PinWrite, DigSrcSignalName, m_pgmDigSrcWave)
    Else
        Call eFuse_DSSC_SetupDigSrcWave(WritePatt, PinWrite, DigSrcSignalName, m_pgmDigSrcWave)
    End If

    ''''In the MarginRead process, it will use gDW_XXX_Pgm_SingleBitWave / gDW_XXX_Pgm_DoubleBitWave to do the comparison with Read

    ''''Write Pattern for programming eFuse
    Call TheHdw.Patterns(WritePatt).Test(pfAlways, 0)   'Write ECID

    Call TurnOffEfusePwrPins(PwrPin, vpwr)
    DebugPrintFunc WritePattSet.Value
    
Exit Function
    
End If
    
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20150604 New for the new UID eFuse ChkList table
''''Should replace the function "VBT_encoding_AES_128bit_NEW()"
Public Function auto_UID_Encoding_128bits() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UID_Encoding_128bits"

    Dim site As Variant
    Dim rnOutput() As String
    Dim i As Long, idx As Long
    Dim iter As Long
    Dim arrayPos As Long
    Dim tsName As String
    Dim UID_Check() As Long
    ReDim UID_Check(gL_UIDCode_Block - 1)
    
    ReDim UID_Code_BitStr(TheExec.sites.Existing.Count - 1, gL_UIDCode_Block - 1)
    ReDim rnOutput(UIDBitsPerCode - 1)

    '################################################################################################
    '# START !! #####################################################################################
    '################################################################################################
    Call UpdateDLogColumns(gI_UID_catename_maxLen)


''''201811XX update
If (gB_eFuse_newMethod) Then
'
'    If (gB_eFuse_DSPMode = True) Then
'        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
'    Else
'        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
'    End If

    Dim j As Long
    Dim m_catename As String
    Dim m_alogrithm As String
    Dim m_bitwidth As Long
    Dim m_UID_Code_BitStr As String
    Dim m_UID_Code_HexStr As String
    Dim F_PrintOneTime As Boolean:: F_PrintOneTime = True

    Dim m_pgmWave As New DSPWave
    Dim m_RndData As New DSPWave
    Dim m_PrintDataArr() As Long
    ReDim m_PrintDataArr(gL_UIDCodeBitWidth)

    'm_RndData.CreateConstant 0, gL_UIDCodeBitWidth / RNG_UTIL_BYTE_SIZE, DspLong
    'iter = gL_UIDCodeBitWidth / RNG_UTIL_BYTE_SIZE

    For i = 0 To UBound(UIDFuse.Category)

        With UIDFuse.Category(i)
            m_catename = .Name
            m_bitwidth = .BitWidth
            m_alogrithm = LCase(.algorithm)
        End With

        If (m_alogrithm = "uid") Then
            m_RndData.CreateConstant 0, m_bitwidth / RNG_UTIL_BYTE_SIZE, DspLong
            iter = m_bitwidth / RNG_UTIL_BYTE_SIZE
            For Each site In TheExec.sites
                For j = 0 To iter - 1
                    m_RndData(site).Select(j, 1, 1).Replace RNG_cryptoRandomByte()
                Next j
                m_pgmWave(site) = m_RndData(site).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, RNG_UTIL_BYTE_SIZE, 0, Bit0IsLsb).ConvertDataTypeTo(DspLong)
            

            With UIDFuse.Category(i).Write
                .BitArrWave = m_pgmWave.Copy
                .BitSummation = m_pgmWave.CalcSum
                .Decimal = (.BitSummation / m_bitwidth)
                .Value = .Decimal
                .Value = CStr(.Value)
            End With
            
'            If (DisplayUID = True) Then
'                If (F_PrintOneTime) Then
'                    TheExec.Datalog.WriteComment ""
'                    TheExec.Datalog.WriteComment FormatNumeric("Site(" + CStr(Site) + ") LSB", 32) + Space(122) + "MSB"
'                    m_PrintDataArr = m_pgmWave(Site).Data
'                    F_PrintOneTime = False
'                End If
'                For j = 0 To (m_BitWidth - 1)
'                    m_UID_Code_BitStr = m_UID_Code_BitStr + CStr(m_PrintDataArr(j))
'                Next j
'                m_UID_Code_HexStr = auto_BinStr2HexStr(m_UID_Code_BitStr, m_BitWidth / 4)
'
'                With UIDFuse.Category(i).Write
'                    .BitstrM(Site) = m_UID_Code_BitStr
'                    .BitStrL(Site) = StrReverse(m_UID_Code_BitStr)
'                    .HexStr(Site) = m_UID_Code_HexStr
'                End With
'
'                TheExec.Datalog.WriteComment vbTab & "CompareValue UID" & i & " : " + m_UID_Code_BitStr
'                i = -1
'                For j = 0 To UBound(m_PrintDataArr) + 1
'                    If (j Mod UIDBitsPerCode = 0) Then
'                        i = i + 1
'                        If (i <> 0) Then theexec.Datalog.WriteComment vbTab & "CompareValue UID" & (i - 1) & " : " + UID_Code_BitStr(Site, i - 1)
'                        If (i = gL_UIDCode_Block) Then Exit For
'                    End If
'                    UID_Code_BitStr(Site, i) = UID_Code_BitStr(Site, i) + CStr(m_PrintDataArr(j))
'                Next j
'                m_UID_Code_HexStr = ""
'                m_UID_Code_BitStr = ""
'            End If
            
            Next site
        End If

    Next i

    If (DisplayUID = True) Then
    Dim m_cnt As Integer
    Dim m_tmpStr As String
        For Each site In TheExec.sites
            
            
            TheExec.Datalog.WriteComment ""
            TheExec.Datalog.WriteComment FormatNumeric("Site(" + CStr(site) + ") MSB", 32) + Space(122) + "LSB"
            m_cnt = 0
            For i = 0 To UBound(UIDFuse.Category)
            m_UID_Code_BitStr = ""
                m_alogrithm = LCase(UIDFuse.Category(i).algorithm)
            
                If (m_alogrithm = "uid") Then
                    With UIDFuse.Category(i)
                        m_PrintDataArr = .Write.BitArrWave(site).Data
                        m_bitwidth = .BitWidth
                    End With
                    
                    For j = 0 To (m_bitwidth - 1)
                        m_UID_Code_BitStr = m_UID_Code_BitStr + CStr(m_PrintDataArr(j))
                    Next j
                    
                    With UIDFuse.Category(i).Write
                        .BitStrM(site) = m_UID_Code_BitStr
                        .BitStrL(site) = StrReverse(m_UID_Code_BitStr)
                        .HexStr(site) = m_UID_Code_HexStr
                    End With
                    
                    arrayPos = 1
                    If (m_bitwidth > UIDBitsPerCode) Then arrayPos = Round(m_bitwidth / UIDBitsPerCode)
                    
                    Do
                        m_tmpStr = Mid(StrReverse(m_UID_Code_BitStr), arrayPos, UIDBitsPerCode)
                        TheExec.Datalog.WriteComment vbTab & "CompareValue UID" & m_cnt & " : " + m_tmpStr
                        UID_Check(m_cnt) = InStr(1, m_tmpStr, "1", vbTextCompare)
                        m_cnt = m_cnt + 1
                        arrayPos = arrayPos - 1
                    Loop Until (arrayPos = 0)
                End If
            Next i
            For i = 0 To gL_UIDCode_Block - 1
                tsName = "UID" + CStr(i) + " Encrytion"
                TheExec.Flow.TestLimit resultVal:=UID_Check(i), lowVal:=1, Tname:=tsName, PinName:="Check"
            Next i
        Next site
        
    End If




Exit Function
End If
    
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


