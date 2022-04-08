Attribute VB_Name = "VBT_LIB_EFUSE_MONITOR"
Option Explicit

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_MONITORBlankChk(ReadPatSet As Pattern, PinRead As PinList, _
                    Optional InterfaceType As String = "", _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MONITORBlankChk"

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
    Dim i As Long, j As Long, k As Long

    Dim FBCChkPass As Boolean
    Dim ResultFlag As New SiteLong
    Dim testName As String, PinName As String
    Dim PrintSiteVarResult As String
    Dim SiteVarValue As Long

    Dim SingleBitArray() As Long
    ''ReDim SingleBitArray(MONITORTotalBitCount - 1)

    ''Dim DoubleBitSum As Long
    Dim DoubleBitArray() As Long
    ''ReDim DoubleBitArray(MONITORBitPerBlockUsed - 1)

    Dim SingleDoubleBitMismatch As New SiteLong
    Dim SignalCap As String, CapWave As New DSPWave
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    Dim m_jobinStage_flag As Boolean

    ReDim gL_Sim_FuseBits(TheExec.sites.Existing.Count - 1, MONITORTotalBitCount - 1) ''''it's for the simulation

    Dim leftStr As String
    Dim rightStr As String
    
     m_jobinStage_flag = auto_eFuse_JobExistInStage("MON", True) ''''<MUST>

    SignalCap = "SignalCapture" 'define capture signal name

    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ   'Karl chnage 0416

    '****** Initialize Site Varaible ******
    Dim m_siteVar As String
    m_siteVar = "MONChk_Var"
    For Each site In TheExec.sites.Existing
        TheExec.sites(site).SiteVariableValue(m_siteVar) = -1
    Next site

    '================================================
    '=  Setup HRAM/DSSC capture cycles           =
    '================================================
    '*** Setup HARM/DSSC Trigger and Capture parameter ***
    gDL_eFuse_Orientation = gE_eFuse_Orientation
    ''''----------------------------------------------------
    ''''201812XX New Method by DSPWave
    ''''----------------------------------------------------
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong

    m_Fusetype = eFuse_MON
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_stage = True

    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    Dim m_SampleSize As Long
    Dim m_SerialType As Boolean
    'InterfaceType = "APB"
    If (InterfaceType = "APB") Then
        'm_SampleSize = (gL_MON_CRC_MSBbit + 1) * 2
        gDB_SerialType = True
        m_SerialType = True
        m_SampleSize = MONITORTotalBitCount
    Else
        gDB_SerialType = False
        m_SerialType = False
        m_SampleSize = MONITORReadCycle
    End If
    
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, m_SampleSize, CapWave  'setup
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)    'run read pattern and capture
     
''''201811XX update
If (gB_eFuse_newMethod = True) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

'    ''''----------------------------------------------------
'    ''''201812XX New Method by DSPWave
'    ''''----------------------------------------------------
'    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
'    Dim m_Fusetype As eFuseBlockType
'    Dim m_SiteVarValue As New SiteLong
'    Dim m_ResultFlag As New SiteLong
'
'    m_Fusetype = eFuse_MON
'    m_FBC = -1               ''''initialize
'    m_ResultFlag = -1        ''''initialize
'    m_SiteVarValue = -1      ''''initialize
'    allblank = True
'    blank_stage = True
'
'    'gDL_eFuse_Orientation = eFuse_2_Bit
'    gDL_eFuse_Orientation = gE_eFuse_Orientation
'
'    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>

    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    gL_eFuse_Sim_Blank = 0
     Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, CapWave, m_FBC, blank_stage, allBlank, m_SerialType)
    ''''----------------------------------------------------

    If (blank_stage.Any(False) = True) Then
        ''''''''if there is any site which is non-blank, then decode to gDW_CFG_Read_Decimal_Cate [check later]
        ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
        ''''it will be present by Hex and Binary compare with the limit later on.
        ''''<NOTICE> gDW_CFG_Read_Decimal_Cate has been decided in non-blank and/or MarginRead
        ''Call auto_eFuse_setReadData(eFuse_CFG, gDW_CFG_Read_DoubleBitWave, gDW_CFG_Read_Decimal_Cate)
        'Call auto_eFuse_setReadData(eFuse_CFG)
        
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(eFuse_MON)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_MON)

        Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_MON, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_MON, False, gB_eFuse_printReadCate)
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
    m_SiteVarValue = blank_stage
    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
    If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only

    For Each site In TheExec.sites
        If (auto_eFuse_GetAllPatTestPass_Flag("MON") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    gL_MON_FBC = m_FBC
    
    Call UpdateDLogColumns(gI_MON_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "MON_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "MON_Blank_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value

Exit Function

End If
    
    
    
    
    
    
    
'
'    ''''While being in the "Retest" stage, the DSSC Read result will be used in the instance "MONITORSingleDoubleBit"
'    ''''So it's needed to set SingleStrArray() to global gS_SingleStrArray()
'
'    auto_eFuse_DSSC_ReadDigCap_32bits MONITORReadCycle, PinRead.Value, gS_SingleStrArray, capWave, allblank 'read back in singlestrarray
'
'    ''''it's used to identify if the Job Name is existed in the MON portion of the eFuse BitDef table.
'    m_jobinStage_flag = auto_eFuse_JobExistInStage("MON", True) ''''<MUST>
'
'    Call UpdateDLogColumns(gI_MON_catename_maxLen)
'
'    For Each Site In TheExec.Sites
'        ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To MONITORReadCycle - 1
'                gS_SingleStrArray(i, Site) = StrReverse(gS_SingleStrArray(i, Site))
'            Next i
'        End If
'
'        ''''20160329 Add, 20161108 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            allblank(Site) = False
'            blank_stage(Site) = True ''''True or False(sim for re-test)
'            If (blank_stage(Site) = True And gS_JobName = "cp1") Then allblank(Site) = True
'
'            If (blank_stage = False) Then ''''20160202, simulation for retest mode
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                TheExec.Datalog.WriteComment vbTab & "[ blank_stage = False, Simulation for the ReTest Mode on Job[" + UCase(gS_JobName) + "] ]"
'                Call eFuseENGFakeValue_Sim
'
'                Dim m_tmpStr As String
'                Dim Expand_eFuse_Pgm_Bit() As Long
'                Dim eFusePatCompare() As String
'                ReDim Expand_eFuse_Pgm_Bit(MONITORTotalBitCount * MONITOR_Repeat_Cyc_for_Pgm - 1)
'                ReDim eFusePatCompare(MONITORReadCycle - 1)
'                Call auto_Make_MONITOR_Pgm_for_Sim(SingleBitArray, Expand_eFuse_Pgm_Bit, eFusePatCompare, blank_stage(Site), False) ''''showPrint if True
'
'                ''''to composite gS_SingleStrArray(i, Site) for the simulation in SingleDoubleBit_Check/Syntax
'                For i = 0 To MONITORReadCycle - 1
'                    m_tmpStr = ""
'                    For j = 0 To MONITORReadBitWidth - 1
'                        k = j + i * MONITORReadBitWidth ''''MUST
'                        m_tmpStr = CStr(SingleBitArray(k)) + m_tmpStr
'                        gL_Sim_FuseBits(Site, k) = SingleBitArray(k)
'                    Next j
'                    gS_SingleStrArray(i, Site) = m_tmpStr
'                Next i
'            End If
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        '====================================================
'        '=  Print the all eFuse Bit data from digCap        =
'        '====================================================
'        Call auto_MONITOR_Memory_Read(gS_SingleStrArray, SingleBitArray, DoubleBitArray, FBCChkPass) 'Read DSSC data
'
'        ''''----------------------------------------------------------------------------------
'        ''''<NOTICE>
'        ''''Because we use 'V' instead of 'L' in the DSSC pattern,
'        ''''so that "Blank-Fuse" can not be judged by the API 'Patgen.PatternBurstPassed'.
'        ''''But it was decided in the routine auto_eFuse_DSSC_ReadDigCap_32bits()
'        ''''If theHdw.digital.Patgen.PatternBurstPassed(Site) = False Then  'If not blank
'        ''''----------------------------------------------------------------------------------
'
'        testName = "MON_BlankChk_" + UCase(gS_JobName)
'
'        If allblank(Site) = False Then ''False means this Efuse is NOT balnk.
'
'            '*** In here it means the eFuse is not blank ****
'            If gS_JobName Like "cp*" Then
'                If FBCChkPass = True Then    'If doubleBit*2-singleBit=0 (i.e. FBCChkPass=true)
'                    ResultFlag(Site) = 0     'Pass Blank check criterion
'                    SiteVarValue = 2
'                    PinName = "Pass"
'                Else  'Fail FBCChkPass       'If doubleBit*2-singleBit<>0 (i.e. FBCChkPass=false)
'                    ResultFlag(Site) = 1     'Fail Blank check criterion
'                    SiteVarValue = 0
'                    PinName = "Fail"
'                End If  'if BlankChkPass = True Then
'
'            ElseIf (gS_JobName Like "*ft*") Then ''''WLFT/FT1/FT2/FT3
'                If FBCChkPass = True Then
'                    ResultFlag(Site) = 0     'Pass Blank check criterion
'                    SiteVarValue = 2
'                    PinName = "Pass"
'                Else
'                    ResultFlag(Site) = 1       'Fail Blank check criterion
'                    SiteVarValue = 0
'                    PinName = "Fail"
'                End If
'            End If
'
'        Else ''True means this Efuse is balnk.
'
'            If gS_JobName = "cp1" Then
'                ''''it's used to identify if the Job Name is existed in the MON portion of the eFuse BitDef table.
'                If (m_jobinStage_flag = False) Then
'                    ResultFlag(Site) = 0        ''Pass Blank check criterion
'                    PinName = "Pass"
'                    SiteVarValue = 2 ''No write in CP1
'                Else
'                    ''''case: m_jobinStage_flag = True
'                    ''''Here it's used to check if HardIP pattern test pass or not.
'                    If (auto_eFuse_GetAllPatTestPass_Flag("MON") = False) Then
'                        ResultFlag(Site) = 1   'Fail as NO fuse
'                        PinName = "Fail"
'                        SiteVarValue = 0
'                    Else
'                        ResultFlag(Site) = 0   'Pass Blank check criterion
'                        PinName = "Pass"
'                        SiteVarValue = 1
'                    End If
'                End If
'
'            ElseIf (gS_JobName = "cp2" And m_jobinStage_flag = True) Then
'                ''''Here it's used to check if HardIP pattern test pass or not.
'                If (auto_eFuse_GetAllPatTestPass_Flag("MON") = False) Then
'                    ResultFlag(Site) = 1   'Fail as NO fuse
'                    PinName = "Fail"
'                    SiteVarValue = 0
'                Else
'                    ResultFlag(Site) = 0   'Pass Blank check criterion
'                    PinName = "Pass"
'                    SiteVarValue = 1
'                End If
'
'            ''''WLFT/FT1/FT2/FT3 should not see a blank eFuse
'            ''''ElseIf (gS_JobName Like "*ft*") Then
'            Else
'                ResultFlag(Site) = 1    'Fail Blank check criterion
'                PinName = "Fail"
'                SiteVarValue = 0
'            End If
'
'        End If
'
'        '====================================================================
'        '=      Check if right and left half block data are consistent      =
'        '====================================================================
'        'This is redundant for CP but necessary for FT in case eFuse bit flip
'        SingleDoubleBitMismatch(Site) = 0   'Deault is Data match if blank=True
'        If allblank(Site) = False Then
'            If (gS_EFuse_Orientation = "UP2DOWN") Then
'                For i = 0 To MONITORReadCycle / 2 - 1
'                    If gS_SingleStrArray(i, Site) <> gS_SingleStrArray(i + MONITORReadCycle / 2, Site) Then
'                        SingleDoubleBitMismatch(Site) = 1   'Data mismatch
'                        Exit For
'                    End If
'                Next i
'            ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
'                For i = 0 To MONITORReadCycle - 1
'                    rightStr = Mid(gS_SingleStrArray(i, Site), MONITORBitsPerRow + 1, MONITORBitsPerRow)
'                    leftStr = Mid(gS_SingleStrArray(i, Site), 1, MONITORBitsPerRow)
'                    If (rightStr <> leftStr) Then
'                        SingleDoubleBitMismatch(Site) = 1   'Data mismatch
'                        Exit For
'                    End If
'                Next i
'            ElseIf (gS_EFuse_Orientation = "SingleUp") Then
'                ''''because there is only 1 block
'                SingleDoubleBitMismatch(Site) = 0
'            ElseIf (gS_EFuse_Orientation = "SingleDown") Then
'            ElseIf (gS_EFuse_Orientation = "SingleRight") Then
'            ElseIf (gS_EFuse_Orientation = "SingleLeft") Then
'            End If
'        End If
'
'        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'
'        If (SiteVarValue <> 1) Then
'            TheExec.Datalog.WriteComment ""
'            TheExec.Datalog.WriteComment "Read All MONITOR eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'            Call auto_PrintAllBitbyDSSC(SingleBitArray, MONITORReadCycle, MONITORTotalBitCount, MONITORReadBitWidth) ''must use MONITORReadBitWidth in Right2Left mode
'        Else
'            ''''20160531
'            ''''SiteVarValue = 1, do the pre-decode for CRC Pgm if needed
'            Call auto_Decode_MONBinary_Data(DoubleBitArray, False)
'        End If
'
'        If (False) Then
'            PrintSiteVarResult = "Site (" + CStr(Site) + ") " + m_siteVar + " = " + CStr(SiteVarValue)
'            TheExec.Datalog.WriteComment PrintSiteVarResult & vbCrLf
'        End If
'
'        gB_MON_decode_flag(Site) = False ''''20160531, <MUST> reset and init
'
'        ''Binning out
'        TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'        TheExec.Flow.TestLimit resultVal:=SingleDoubleBitMismatch, lowVal:=0, hiVal:=0, Tname:="Chk Mismatch" '''', PinName:=PinName
'        TheExec.Flow.TestLimit resultVal:=ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName, PinName:=PinName
'
'        If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
'
'    Next Site
'
'    Call UpdateDLogColumns__False
'    DebugPrintFunc ReadPatSet.Value
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_MONITORWrite_byCondition(WritePattSet As Pattern, PinWrite As PinList, _
                    PwrPin As String, vpwr As Double, _
                    condstr As String, _
                    Optional catename_grp As String, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MONITORWrite_byCondition"

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

    Dim site As Variant
    Dim eFuse_Pgm_Bit() As Long
    Dim Expand_eFuse_Pgm_Bit() As Long, eFusePatCompare() As String
    Dim i As Long, j As Long
    Dim SegmentSize As Long

    Dim DigSrcSignalName As String
    Dim Expand_Size As Long
    Dim Even_Count As Long
    Dim Odd_Count As Long

    Dim Count As Long
    Dim FuseBitSize As Long

    condstr = LCase(condstr) ''''<MUST>

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

    TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName
    

    '====================================================
    '=  1. Define the Bit length of Programming         =
    '=  2. Define the cycle number of Compare Pattern   =
    '====================================================
    FuseBitSize = MONITORBlock * MONITORBitPerBlockUsed
    Expand_Size = (FuseBitSize * MONITOR_Repeat_Cyc_for_Pgm) 'Because there are repeat cycle in C651 pattern, we have to create multiple DSSC
    ReDim Expand_eFuse_Pgm_Bit(Expand_Size - 1)
    ReDim eFusePatCompare(MONITORReadCycle - 1)
    ReDim gL_MONFuse_Pgm_Bit(TheExec.sites.Existing.Count - 1, FuseBitSize - 1)
    
    '========================================================
    '=  1. Make DigSrc Pattern for eFuse programming        =
    '=  2. Make Read Pattern for eFuse Read tests           =
    '========================================================
    DigSrcSignalName = "MONITOR_DigSrcSignal"

    Call UpdateDLogColumns(gI_MON_catename_maxLen)

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''201808XX update
    If (TheExec.TesterMode = testModeOffline) Then
        If (condstr <> "cp1") Then
            For Each site In TheExec.sites
                Call eFuseENGFakeValue_Sim
            Next site
        End If
    End If

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

    If (condstr = "cp1_early") Then
        m_cmpStage = "cp1_early"
    Else
        ''''condStr = "stage"
        m_cmpStage = gS_JobName
    End If
    
    ''''Only composite case "real or bincut" PgmBits Wave per Stage requirement
    For i = 0 To UBound(MONFuse.Category)
        With MONFuse.Category(i)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_defreal = LCase(.Default_Real)
        End With
        
        If (m_stage = gS_JobName) Then ''''was If (m_stage = m_cmpStage) Then
            If (m_algorithm = "crc") Then
                m_crc_idx = i
                ''''special handle on the next process
                ''''skip it here
            ElseIf (m_defreal = "real" Or m_defreal = "bincut") Then
                If (m_algorithm = "vddbin") Then
                    Call auto_eFuse_Vddbin_bincut_setWriteData(eFuse_MON, i)
                End If
            'If (m_defreal = "real") Then
                ''''---------------------------------------------------------------------------
                With MONFuse.Category(i)
                    ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
                    Call rundsp.eFuse_updatePgmWave_byCategory(.BitIndexWave, .Write.BitArrWave)
                End With
                ''''---------------------------------------------------------------------------
            End If
        Else
            ''''doNothing
        End If
    Next i
    
    '''process CRC bits calculation
    If (gS_MON_CRC_Stage = gS_JobName) Then
        Dim mSL_bitwidth As New SiteLong
        mSL_bitwidth = gL_MON_CRC_BitWidth
        ''''CRC case
        With MONFuse.Category(m_crc_idx)
            ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
            Call rundsp.eFuse_updatePgmWave_CRCbits(eFuse_MON, mSL_bitwidth, .BitIndexWave)
        End With
    End If
    
    ''''composite effective PgmBits per Stage requirement
    m_pgmRes = 0
    If (m_cmpStage = "cp1_early") Then
        ''''condStr = "cp1_early"
        Call auto_eFuse_Gen_PgmBitSrcWave_onlyDeaultBits(eFuse_MON, m_pgmDigSrcWave, m_pgmRes)
    Else
        '''condStr = "stage"
        '''Here Parameter bitFlag_mode=1 (Stage Bits)
        m_calcCRC = 0
        Call rundsp.eFuse_Gen_PgmBitSrcWave(eFuse_MON, 1, m_pgmDigSrcWave, m_pgmRes)
    End If
    TheHdw.Wait 1# * ms
    
    Call UpdateDLogColumns(gI_MON_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=m_pgmRes, lowVal:=1, hiVal:=1, Tname:="MON_PGM_" + UCase(gS_JobName)
    Call UpdateDLogColumns__False

    ''''use global DSP variable
    Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_MON_Pgm_SingleBitWave, gDL_TotalBits, gDL_ReadCycles, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_MON, gDW_MON_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)
    
    Call TurnOnEfusePwrPins(PwrPin, vpwr)
    
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







'
'    For Each Site In TheExec.Sites
'
'        If (TheExec.TesterMode = testModeOffline) Then ''''20160526 update
'            Call eFuseENGFakeValue_Sim
'        End If
'
'        ''''20160104 New
'        If (condstr = "stage") Then
'            SegmentSize = auto_Make_MONITOR_Pgm_and_Read_Array(eFuse_Pgm_Bit(), Expand_eFuse_Pgm_Bit(), eFusePatCompare())
'        ElseIf (condstr = "category") Then
'            SegmentSize = auto_Make_MONITOR_Pgm_and_Read_Array_byCategory(catename_grp, Expand_eFuse_Pgm_Bit(), eFuse_Pgm_Bit(), eFusePatCompare())
'        Else
'            ''''default=all bits are 0, here it prevents any typo issue
'            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (Stage,Category)"
'            For i = 0 To FuseBitSize - 1
'                eFuse_Pgm_Bit(i) = 0
'            Next i
'            For i = 0 To Expand_Size - 1
'                Expand_eFuse_Pgm_Bit(i) = 0
'            Next i
'            SegmentSize = Expand_Size
'        End If
'
'        '============================================================================
'        '= << This subroutine is for creating DSSC Signal fro 'STROBE' pin >>       =
'        '=                                                                          =
'        '=  1. This subroutine is for passing Expand_eFuse_Pgm_Bit() to a Dspwave.  =
'        '=  2. Store this Dspwave into DigSrc memory for STROBE pin                 =
'        '============================================================================
'
'        DSSC_SetupDigSrcWave WritePatt, PinWrite, DigSrcSignalName, SegmentSize, Expand_eFuse_Pgm_Bit
'
'        'Print out programming bits
'        Call auto_PrintAllPgmBits(eFuse_Pgm_Bit(), MONITORReadCycle, MONITORTotalBitCount, MONITORReadBitWidth) ''must use MONITORReadBitWidth
'
'        For i = 0 To MONITORTotalBitCount - 1
'            gL_MONFuse_Pgm_Bit(Site, i) = eFuse_Pgm_Bit(i)
'        Next i
'
'    Next Site
'
'    Call UpdateDLogColumns__False
'
'    TheHdw.DSSC.Pins(PinWrite).Pattern(WritePatt).Source.Signals.DefaultSignal = DigSrcSignalName
'
'    'Step2. Write Pattern for programming eFuse
'    TheHdw.Wait 0.0001
'
'    Call TheHdw.Patterns(WritePatt).test(pfAlways, 0)   'Write MONITOR
'
'    DebugPrintFunc WritePattSet.Value
'
'    Call TurnOffEfusePwrPins(PwrPin, vpwr)

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_MONITORRead_by_OR_2Blocks(ReadPatSet As Pattern, PinRead As PinList, _
                    condstr As String, _
                    Optional InterfaceType As String = "", _
                    Optional catename_grp As String = "", _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MONITORRead_by_OR_2Blocks"

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
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m_tmpStr As String
    Dim FailCnt As New SiteLong
    Dim testName As String

    Dim SingleStrArray() As String, DoubleBitArray() As Long
    Dim SingleBitArray() As Long

    Dim CapWave As New DSPWave
    Dim blank As New SiteBoolean
    Dim SignalCap As String
    
    ''''-------------------------------------------------------------------------------------
    ''ReDim SingleStrArray(MONITORReadCycle - 1, TheExec.Sites.Existing.Count - 1)
    ''ReDim DoubleBitArray(MONITORBitPerBlockUsed - 1)
    ''ReDim SingleBitArray(MONITORTotalBitCount - 1)
    
    condstr = LCase(condstr) ''''<MUST>

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
  
    '====================================================
    '=  1. Define the Bit length of Programming         =
    '=  2. Define the cycle number of Compare Pattern   =
    '====================================================
    ReDim eFuse_Pgm_Bit(MONITORTotalBitCount - 1)
    
    Dim m_SampleSize As Long
    Dim m_SerialType As Boolean
    'InterfaceType = "APB"
    If (InterfaceType = "APB") Then
        'm_SampleSize = (gL_MON_CRC_MSBbit + 1) * 2
        gDB_SerialType = True
        m_SerialType = True
        m_SampleSize = MONITORTotalBitCount
    Else
        gDB_SerialType = False
        m_SerialType = False
        m_SampleSize = MONITORReadCycle
    End If

    SignalCap = "MarginCapture"
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, m_SampleSize, CapWave
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)
    testName = "MON_ORMarginRead" + "_" + UCase(gS_JobName)
     
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

    m_Fusetype = eFuse_MON
    m_FBC = -1       ''''init to failure
    m_cmpResult = -1 ''''init to failure

    ''''--------------------------------------------------------------------------
    '''' Offline Simulation Start                                                |
    ''''--------------------------------------------------------------------------
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        Call rundsp.eFuse_SingleBitWave2CapWave32Bits(eFuse_MON, CapWave)
        Call auto_eFuse_print_capWave32Bits(eFuse_MON, CapWave, False) ''''True to print out
    End If
    ''''--------------------------------------------------------------------------
    '''' Offline Simulation End                                                  |
    ''''--------------------------------------------------------------------------

    If (condstr = "cp1_early") Then
        m_bitFlag_mode = 0
    ElseIf (condstr = "stage") Then
        m_bitFlag_mode = 1
    ElseIf (condstr = "all") Then
        m_bitFlag_mode = 2 ''''update later, was 2
    Else
        ''''default, here it prevents any typo issue
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
        m_FBC = -1
        m_cmpResult = -1
    End If

    Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, m_cmpResult, , , m_SerialType)

    gL_MON_FBC = m_FBC

    ''''''[NOTICE] Decode and Print have moved to SingleDoubleBit()

    Call UpdateDLogColumns(gI_MON_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=m_cmpResult, lowVal:=0, hiVal:=0
    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value

Exit Function

End If
     
     
     
     
'    auto_eFuse_DSSC_ReadDigCap_32bits MONITORReadCycle, PinRead.Value, SingleStrArray, capWave, blank ''''Here Must use local variable
'
'    Call UpdateDLogColumns(gI_MON_catename_maxLen)
'
'    '================================================
'    '=  1. Make Program bit array                   =
'    '=  2. Make Read Compare bit array              =
'    '================================================
'    For Each Site In TheExec.Sites
'        ''''20151026 add, 20161108 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            ''''------------------------------------------------------
'            ''''20161108 update
'            Call auto_Decompose_StrArray_to_BitArray("MON", gS_SingleStrArray, SingleBitArray, 0)
'            ''''------------------------------------------------------
'
'            ''''to composite gS_SingleStrArray(i, Site) for the simulation in SingleDoubleBit_Check/Syntax
'            For i = 0 To MONITORReadCycle - 1
'                m_tmpStr = ""
'                For j = 0 To MONITORReadBitWidth - 1
'                    k = j + i * MONITORReadBitWidth ''''MUST
'                    If (SingleBitArray(k) = 0 And gL_MONFuse_Pgm_Bit(Site, k) = 1) Then
'                        SingleBitArray(k) = gL_MONFuse_Pgm_Bit(Site, k)
'                        gL_Sim_FuseBits(Site, k) = SingleBitArray(k)
'                    End If
'                    m_tmpStr = CStr(SingleBitArray(k)) + m_tmpStr
'                Next j
'                gS_SingleStrArray(i, Site) = m_tmpStr
'                SingleStrArray(i, Site) = m_tmpStr
'            Next i
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        ''''gS_SingleStrArray() can be used later in auto_MONITORSingleDoubleBit()
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To MONITORReadCycle - 1
'                ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'                SingleStrArray(i, Site) = StrReverse(SingleStrArray(i, Site))
'                gS_SingleStrArray(i, Site) = SingleStrArray(i, Site)
'            Next i
'        Else
'            For i = 0 To MONITORReadCycle - 1
'                gS_SingleStrArray(i, Site) = SingleStrArray(i, Site)
'            Next i
'        End If
'
'        '================================================
'        '=  Compare EFuse Read with EFuse Program       =
'        '================================================
'        Call auto_OR_2Blocks("MON", SingleStrArray, SingleBitArray, DoubleBitArray) ''''calc gL_MON_FBC
'
'        For i = 0 To MONITORTotalBitCount - 1
'            eFuse_Pgm_Bit(i) = gL_MONFuse_Pgm_Bit(Site, i)
'        Next i
'
'        FailCnt(Site) = 0
'
'        ''''20160104 Update
'        If (condstr = "all") Then
'            Call auto_MONCompare_DoubleBit_PgmBit_byAll(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = "stage") Then
'            Call auto_MONCompare_DoubleBit_PgmBit_byStage(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = "category") Then
'            Call auto_eFuse_Compare_DoubleBit_PgmBit_byCategory("MON", catename_grp, DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        Else
'            ''''default, here it prevents any typo issue
'            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (All,Stage,Category)"
'            FailCnt(Site) = -1
'        End If
'
'        '====================================================
'        '=  Print the all eFuse Bit data from DSSC          =
'        '====================================================
'        TheExec.Datalog.WriteComment ""
'        TheExec.Datalog.WriteComment "Read All MONITOR eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'        Call auto_PrintAllBitbyDSSC(SingleBitArray, MONITORReadCycle, MONITORTotalBitCount, MONITORReadBitWidth) ''must use MONITORReadBitWidth in Right2Left mode
'
'        '========================================================
'        '=  1. Decode binary data to meaningful decimal data    =
'        '=  2. Wrap Up the string for writing to HKEY           =
'        '========================================================
'
'        ''''<Important> User Need to check the content inside
'        Call auto_Decode_MONBinary_Data(DoubleBitArray)
'
'        ''''' 20161019 update
'        If (checkJob_less_Stage_Sequence(gS_MON_CRC_Stage) = True) Then
'            gS_MON_CRC_HexStr(Site) = "00000000"
'        Else
'            gS_MON_CRC_HexStr(Site) = auto_MONITOR_CRC2HexStr(DoubleBitArray)
'        End If
'
'    Next Site 'For Each Site In TheExec.Sites
'
'    TheExec.Flow.TestLimit resultVal:=FailCnt, lowVal:=0, hiVal:=0, Tname:=testName
'
'    Call UpdateDLogColumns__False
'    DebugPrintFunc ReadPatSet.Value
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_MONITORSingleDoubleBit(Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MONITORSingleDoubleBit"
    
    Dim site As Variant

    Dim i As Long, j As Long, k As Long

    ''Dim SingleStrArray() As String
    Dim SingleBitArray() As Long
    Dim DoubleBitArray() As Long

    Dim tmpStr As String
    
    ''''--------------------------------------------------------------------------------
    ''ReDim SingleStrArray(MONITORReadCycle - 1, TheExec.Sites.Existing.Count - 1)
    ''ReDim DoubleBitArray(MONITORBitPerBlockUsed - 1)
    ''ReDim SingleBitArray(MONITORTotalBitCount - 1)
    
    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

    ''''------------------------------------------------------------------------------------------------------------------
    ''''<Important Notice>
    ''''------------------------------------------------------------------------------------------------------------------
    ''''gS_SingleStrArray() was extracted in the module auto_MONITORRead_by_OR_2Blocks() then used in auto_MONITORSingleDoubleBit()
    ''''gS_SingleStrArray() is the result of the NormRead or MarginRead
    ''''
    ''''So it doesn't need to run the pattern and DSSC to get the SignalStrArray, and save test time
    ''''------------------------------------------------------------------------------------------------------------------
    
    Call UpdateDLogColumns(gI_MON_catename_maxLen)
    
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
    Call auto_eFuse_setReadData_forSyntax(eFuse_MON)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_MON)
    
    ''''All the read action has been down in blank and/or MarginRead
    ''''gDW_CFG_Read_cmpsgWavePerCyc used to display the cmpare result (2-bit mode)
    Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_MON, gB_eFuse_printBitMap)
    If (gS_JobName = "cp1_early") Then
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_MON, True, gB_eFuse_printReadCate)
    Else
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_MON, False, gB_eFuse_printReadCate)
    End If
    
    
    ''''Print CRC calcBits information
    Dim m_crcBitWave As New DSPWave
    Dim mS_hexStr As New SiteVariant
    Dim mS_bitStrM As New SiteVariant
    Dim m_debugCRC As Boolean
    Dim m_cnt As Long
    Dim m_siteVar As String
    m_siteVar = "MONChk_Var"
    m_debugCRC = True

    ''''<MUST> Initialize
    gS_MON_Read_calcCRC_hexStr = "0x00000000"
    gS_MON_Read_calcCRC_bitStrM = ""
    CRC_Shift_Out_String = ""
    If (auto_eFuse_check_Job_cmpare_Stage(gS_MON_CRC_Stage) >= 0) Then
        Call rundsp.eFuse_Read_to_calc_CRCWave(eFuse_MON, gL_MON_CRC_BitWidth, m_crcBitWave)
        TheHdw.Wait 1# * ms ''''check if it needs

        If (m_debugCRC = False) Then
            ''''Here get gS_CFG_Read_calcCRC_hexStr for the syntax check
            Call auto_eFuse_bitWave_to_binStr_HexStr(m_crcBitWave, gS_MON_Read_calcCRC_bitStrM, gS_MON_Read_calcCRC_hexStr, True, m_debugCRC)
        Else
            ''''m_debugCRC=True => Debug purpose for the print
            TheExec.Datalog.WriteComment "------Read CRC Category Result------"
            Call auto_eFuse_bitWave_to_binStr_HexStr(m_crcBitWave, gS_MON_Read_calcCRC_bitStrM, gS_MON_Read_calcCRC_hexStr, True, m_debugCRC)
            TheExec.Datalog.WriteComment ""

            ''''[Pgm CRC calcBits] only gS_CFG_CRC_Stage=Job and CFGChk_Var=1
            If (gS_MON_CRC_Stage = gS_JobName) Then
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
    If (gS_JobName <> "cp1_early") Then
        For Each site In TheExec.sites
            DoubleBitArray = gDW_MON_Read_DoubleBitWave.Data
            
            gS_MON_Direct_Access_Str(site) = "" ''''is a String [(bitLast)......(bit0)]
            
            For i = 0 To UBound(DoubleBitArray)
                gS_MON_Direct_Access_Str(site) = CStr(DoubleBitArray(i)) + gS_MON_Direct_Access_Str(site)
            Next i
            ''TheExec.Datalog.WriteComment "gS_CFG_Direct_Access_Str=" + CStr(gS_CFG_Direct_Access_Str(Site))
    
            ''''20161114 update for print all bits (DTR) in STDF
            Call auto_eFuse_to_STDF_allBits("MON", gS_MON_Direct_Access_Str(site))
        Next site
    End If
    ''''----------------------------------------------------------------------------------------------

    ''''gL_CFG_FBC has been check in Blank/MarginRead
    Call UpdateDLogColumns(gI_MON_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=gL_MON_FBC, lowVal:=0, hiVal:=0, Tname:="MON_FBCount_" + UCase(gS_JobName) '2d-s=0
    Call UpdateDLogColumns__False

Exit Function

End If
 
    
    
    
    
    

'    For Each Site In TheExec.Sites
'        ''''20161108 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            If (False) Then ''''True for Debug, 20161031
'                TheExec.Datalog.WriteComment "---Offline--- Read All MONITOR eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'                Call auto_PrintAllBitbyDSSC(SingleBitArray(), MONITORReadCycle, MONITORTotalBitCount, MONITORReadBitWidth)
'            End If
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        '==============================================================
'        '=  Extract DigCap data for exhibit MONITOR eFuse              =                                                                                                                  =
'        '==============================================================
'        'Before comparing with program bit, you have to OR 2 block bit by bit
'        Call auto_OR_2Blocks("MON", gS_SingleStrArray, SingleBitArray, DoubleBitArray) ''''get gL_MON_FBC(site)
'
'        ''''20170220 update
'        If (gL_MON_FBC(Site) > 0) Then
'            TmpStr = "The Fail Bit Count of MONITOR eFuse at Site(" + CStr(Site) + ") is " + CStr(gL_MON_FBC(Site))
'            TmpStr = TmpStr + " (Max FBC =0)"
'            TheExec.Datalog.WriteComment TmpStr
'        End If
'
'        If (TheExec.Sites(Site).SiteVariableValue("MONChk_Var") <> 1 Or gB_MON_decode_flag(Site) = False) Then
'            ''''ReTest Stage
'            ''''<Important> User Need to check the content inside
'            Call auto_Decode_MONBinary_Data(DoubleBitArray)
'
'            ''''' 20161019 update
'            If (checkJob_less_Stage_Sequence(gS_MON_CRC_Stage) = True) Then
'                gS_MON_CRC_HexStr(Site) = "00000000"
'            Else
'                gS_MON_CRC_HexStr(Site) = auto_MONITOR_CRC2HexStr(DoubleBitArray)
'            End If
'        End If
'
'        ''''----------------------------------------------------------------------------------------------
'        gS_MON_Direct_Access_Str(Site) = "" ''''is a String [(bitLast)......(bit0)]
'        For i = 0 To UBound(DoubleBitArray)
'            gS_MON_Direct_Access_Str(Site) = CStr(DoubleBitArray(i)) + gS_MON_Direct_Access_Str(Site)
'        Next i
'        ''theExec.Datalog.WriteComment "gS_MON_Direct_Access_Str=" + CStr(gS_MON_Direct_Access_Str(Site))
'
'        ''''20161114 update for print all bits (DTR) in STDF
'        Call auto_eFuse_to_STDF_allBits("MON", gS_MON_Direct_Access_Str(Site))
'        ''''----------------------------------------------------------------------------------------------
'
'    Next Site
'
'    TheExec.Flow.TestLimit resultVal:=gL_MON_FBC, lowVal:=0, hiVal:=0, Tname:="FailBitCount"
'
'    Call UpdateDLogColumns__False

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_ChkAllMONITOREfuseData() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ChkAllMONITOREfuseData"

    Dim i As Long
    Dim site As Variant

    Dim m_tsname As String
    Dim m_catename As String
    Dim m_bitStrM As String
    Dim m_bitwidth As Long
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_decimal As Variant ''20160506 update, was Long
    Dim m_bitsum As Long
    Dim m_value As Variant
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim vbinIdx As Long
    Dim vbinflag As Long
    Dim tmpStr As String
    Dim m_crchexStr As String
    Dim m_testValue As Variant
    Dim m_stage As String
    Dim m_defreal As String
    Dim m_HexStr As String

    Call UpdateDLogColumns(gI_MON_catename_maxLen)
    
 
''''201811XX update
If (gB_eFuse_newMethod) Then
    
    Dim m_Fusetype As eFuseBlockType
    Dim condstr As String
    
    m_Fusetype = eFuse_MON
    
    ''''20170630 update
    If (Trim(condstr) = "") Then
        condstr = "all"
    End If
    condstr = LCase(condstr)

    Call auto_eFuse_SyntaxCheck(m_Fusetype, condstr)
    
Exit Function
End If
    
    
    
    
'    For Each Site In TheExec.Sites
'
'        ''''Check the result and its test limit
'        For i = 0 To UBound(MONFuse.Category)
'            m_stage = LCase(MONFuse.Category(i).Stage)
'            m_catename = MONFuse.Category(i).Name
'            m_algorithm = LCase(MONFuse.Category(i).Algorithm)
'            m_defreal = LCase(MONFuse.Category(i).Default_Real)
'            m_lolmt = MONFuse.Category(i).LoLMT
'            m_hilmt = MONFuse.Category(i).HiLMT
'            m_bitwidth = MONFuse.Category(i).Bitwidth
'            m_decimal = MONFuse.Category(i).Read.Decimal(Site)
'            m_bitStrM = MONFuse.Category(i).Read.BitstrM(Site)
'            m_value = MONFuse.Category(i).Read.Value(Site)
'            m_bitSum = MONFuse.Category(i).Read.BitSummation(Site)
'            m_hexStr = MONFuse.Category(i).Read.HexStr(Site)
'            m_tsName = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
'
'            If (m_algorithm = "freq") Then
'                m_testValue = m_decimal
'                ''''20160808 add for the special case.
'                ''''ECID::AFREQ_CTRL_TRIM_CP1 == MON::AFREQ_CTRL_TRIM
'''''                If (m_defreal = "real") Then
'''''                    If (UCase(m_catename) = "AFREQ_CTRL_TRIM") Then
'''''                         m_lolmt = ECIDFuse.Category(ECIDIndex("AFREQ_CTRL_TRIM_CP1")).Read.Decimal(Site)
'''''                         m_hilmt = m_lolmt
'''''                    End If
'''''                End If
'
'            ElseIf (m_algorithm = "trim") Then
'                m_testValue = m_decimal
'                ''''20160531 update
'                ''''<NOTICE> 20160803 update, user maintain
'                ''''CFG::Thermal_Sensor_0_Data_TRIMG ==== MON::THERMAL_PARAM_TRIMG
'                ''''CFG::Thermal_Sensor_0_Data_TRIMO ==== MON::THERMAL_PARAM_TRIMO
'                ''''------------------------------------------------------------------
'                ''''CFG::TrimG_SOC_0 ==== MON::MON_SOC_TrimG_0
'                ''''CFG::TrimO_SOC_0 ==== MON::MON_SOC_TrimO_0
'                If (m_defreal = "real") Then
'                    If (UCase(m_catename) = UCase("MON_SOC_TrimG_0")) Then
'                        m_lolmt = CFGFuse.Category(CFGIndex("TrimG_SOC_0")).Read.Decimal(Site)
'                        m_hilmt = m_lolmt
'                    ElseIf (UCase(m_catename) = UCase("MON_SOC_TrimO_0")) Then
'                        m_lolmt = CFGFuse.Category(CFGIndex("TrimO_SOC_0")).Read.Decimal(Site)
'                        m_hilmt = m_lolmt
'                    End If
'                End If
'
'            ''''20161019 update
'            ElseIf (m_algorithm = "crc") Then
'                m_value = auto_BinStr2HexStr(m_bitStrM, 4) ''''<MUST>
'                ''''20170309, per Jack's recommendation to use Read_Data to do the CRC calculation result (=gS_MON_CRC_HexStr)
'                ''''          and compare with the CRC readCode.
'                m_crchexStr = UCase(gS_MON_CRC_HexStr(Site))
'                m_tsName = m_tsName + "_" + m_crchexStr
'
'                If (gB_eFuse_Disable_ChkLMT_Flag = True) Then
'                    ''''set Pass
'                    TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                Else
'                    If (UCase(CStr(m_value)) = m_crchexStr) Then
'                        ''''Pass
'                        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                    Else
'                        ''''Fail
'                        TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                    End If
'                End If
'
'            ElseIf (m_algorithm Like "*reserve*") Then
'                m_testValue = m_decimal ''''20160927 update, was m_bitsum
'            Else
'                ''''other cases, 20160927 update
'                m_testValue = m_decimal
'                ''''TheExec.Datalog.WriteComment "undefined Algorithm: " + m_algorithm
'            End If
'
'            ''''20160108 New
'            If (m_algorithm <> "crc") Then
'                Call auto_eFuse_chkLoLimit("MON", i, m_stage, m_lolmt)
'                Call auto_eFuse_chkHiLimit("MON", i, m_stage, m_hilmt)
'
'                ''''20170811 update
'                If (m_bitwidth >= 32) Then
'                    ''m_tsName = m_tsName + "_" + m_hexStr
'                    ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
'                    m_lolmt = auto_Value2HexStr(m_lolmt, m_bitwidth)
'                    m_hilmt = auto_Value2HexStr(m_hilmt, m_bitwidth)
'
'                    ''''------------------------------------------
'                    ''''compare with lolmt, hilmt
'                    ''''m_testValue 0 means fail
'                    ''''m_testValue 1 means pass
'                    ''''------------------------------------------
'                    m_testValue = auto_TestStringLimit(m_hexStr, CStr(m_lolmt), CStr(m_hilmt))
'                    m_lolmt = 1
'                    m_hilmt = 1
'                Else
'                    ''''20160620 update
'                    ''''20160927 update the new logical methodology for the unexpected binary decode.
'                    If (auto_isHexString(CStr(m_lolmt)) = True) Then
'                        ''''translate to double value
'                        m_lolmt = auto_HexStr2Value(m_lolmt)
'                    Else
'                        ''''doNothing, m_lolmt = m_lolmt
'                    End If
'
'                    If (auto_isHexString(CStr(m_hilmt)) = True) Then
'                        ''''translate to double value
'                        m_hilmt = auto_HexStr2Value(m_hilmt)
'                    Else
'                        ''''doNothing, m_hilmt = m_hilmt
'                    End If
'                End If
'                TheExec.Flow.TestLimit resultVal:=m_testValue, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsName
'            End If
'        Next i
'        TheExec.Datalog.WriteComment ""
'    Next Site
'
'    Call UpdateDLogColumns__False
'
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_MONITORRead_Decode(ReadPatSet As Pattern, PinRead As PinList, Optional Validating_ As Boolean, _
                                        Optional InterfaceType As String = "")

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MONITORRead_Decode"

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
    Dim i As Long, j As Long

    Dim SingleStrArray() As String
    Dim SingleBitArray() As Long
    Dim DoubleBitArray() As Long
    
    ''ReDim SingleStrArray(MONITORReadCycle - 1, TheExec.Sites.Existing.Count - 1)
    ''ReDim SingleBitArray(MONITORTotalBitCount - 1)
    ''ReDim DoubleBitArray(MONITORBitPerBlockUsed - 1)

    Dim SignalCap As String, CapWave As New DSPWave
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean

    ReDim gL_Sim_FuseBits(TheExec.sites.Existing.Count - 1, MONITORTotalBitCount - 1) ''''it's for the simulation

    SignalCap = "SignalCapture" 'define capture signal name

    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    '================================================
    '=  Setup HRAM/DSSC capture cycles           =
    '================================================
    '*** Setup HARM/DSSC Trigger and Capture parameter ***
    
    Dim m_SampleSize As Long
    Dim m_SerialType As Boolean
    'InterfaceType = "APB"
    If (InterfaceType = "APB") Then
        'm_SampleSize = (gL_MON_CRC_MSBbit + 1) * 2
        gDB_SerialType = True
        m_SerialType = True
        m_SampleSize = MONITORTotalBitCount
    Else
        gDB_SerialType = False
        m_SerialType = False
        m_SampleSize = MONITORReadCycle
    End If
    
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, m_SampleSize, CapWave  'setup
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)    'run read pattern and capture

    ''''201811XX update
If (gB_eFuse_newMethod = True) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If
    
    'Dim m_SampleSize As Long
    'Dim m_SerialType As Boolean
    
    gDL_eFuse_Orientation = gE_eFuse_Orientation
    
    ''''----------------------------------------------------
    ''''201812XX New Method by DSPWave
    ''''----------------------------------------------------
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong
    Dim testName As String, PinName As String

    m_Fusetype = eFuse_MON
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_stage = True

    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
'    'InterfaceType = "APB"
'    If (InterfaceType = "APB") Then
'        'm_SampleSize = (gL_MON_CRC_MSBbit + 1) * 2
'        gDB_SerialType = True
'        m_SerialType = True
'        m_SampleSize = MONITORTotalBitCount
'    Else
'        gDB_SerialType = False
'        m_SerialType = False
'        m_SampleSize = MONITORReadCycle
'    End If
'
'    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, m_SampleSize, capWave  'setup
'    'auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, MONITORReadCycle, capWave  'setup
'    Call TheHdw.Patterns(ReadPatt).test(pfAlways, 0, tlResultModeDomain)    'run read pattern and capture

    'gDL_eFuse_Orientation = eFuse_2_Bit


    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    gL_eFuse_Sim_Blank = 0
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, CapWave, m_FBC, blank_stage, allBlank, m_SerialType)
    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, capWave, m_FBC, blank_stage, allblank)
    ''''----------------------------------------------------

    If (True) Then
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(eFuse_MON)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_MON)

        Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_MON, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_MON, False, gB_eFuse_printReadCate)
    End If
    gL_MON_FBC = m_FBC
    
    testName = "MON_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value
    
    ''''20170111 Add
    Call auto_eFuse_ReadAllData_to_DictDSPWave("MON", False, False)

Exit Function

End If





'    auto_eFuse_DSSC_ReadDigCap_32bits MONITORReadCycle, PinRead.Value, SingleStrArray, capWave, allblank 'read back in singlestrarray
'
'    Call UpdateDLogColumns(gI_SEN_catename_maxLen)
'
'    For Each Site In TheExec.Sites
'        ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To MONITORReadCycle - 1
'                SingleStrArray(i, Site) = StrReverse(SingleStrArray(i, Site))
'            Next i
'        End If
'
'        ''''20160202 Add, 20161107 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            allblank(Site) = False
'            blank_stage(Site) = True ''''True or False(sim for re-test)
'            If (blank_stage(Site) = True And gS_JobName = "cp1") Then allblank(Site) = True
'
'            TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'            If (blank_stage(Site) = True) Then
'                TheExec.Datalog.WriteComment vbTab & "[ blank_stage = True, Simulate Category (m_stage < Job[" + UCase(gS_JobName) + "]) ]"
'            Else
'                TheExec.Datalog.WriteComment vbTab & "[ blank_stage = False, Simulate Category (m_stage <= Job[" + UCase(gS_JobName) + "]) ]"
'            End If
'            Call eFuseENGFakeValue_Sim
'
'            Dim Expand_eFuse_Pgm_Bit() As Long
'            Dim eFusePatCompare() As String
'            ReDim Expand_eFuse_Pgm_Bit(MONITORTotalBitCount * MONITOR_Repeat_Cyc_for_Pgm - 1)
'            ReDim eFusePatCompare(MONITORReadCycle - 1)
'            Call auto_Make_MONITOR_Pgm_for_Sim(SingleBitArray, Expand_eFuse_Pgm_Bit, eFusePatCompare, blank_stage(Site), False) ''''showPrint if True
'
'            Dim k As Long
'            Dim m_tmpStr As String
'            For i = 0 To MONITORReadCycle - 1
'                m_tmpStr = ""
'                For j = 0 To MONITORReadBitWidth - 1
'                    k = j + i * MONITORReadBitWidth ''''MUST
'                    gL_Sim_FuseBits(Site, k) = SingleBitArray(k) ''''it's used for the Read/Syntax simulation
'                    m_tmpStr = CStr(SingleBitArray(k)) + m_tmpStr
'                Next j
'                SingleStrArray(i, Site) = m_tmpStr
'            Next i
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        Call auto_OR_2Blocks("MON", SingleStrArray, SingleBitArray, DoubleBitArray)  ''''calc gL_MON_FBC
'
'        '====================================================
'        '=  Print the all eFuse Bit data from DSSC          =
'        '====================================================
'        TheExec.Datalog.WriteComment ""
'        TheExec.Datalog.WriteComment "Read All MONITOR eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'        Call auto_PrintAllBitbyDSSC(SingleBitArray, MONITORReadCycle, MONITORTotalBitCount, MONITORReadBitWidth)
'
'        Call auto_Decode_MONBinary_Data(DoubleBitArray, Not allblank(Site)) ''''true for debug, only NOT allBlank to show the decode result.
'
'    Next Site ''For Each Site In TheExec.Sites
'
'    TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0
'
'    Call UpdateDLogColumns__False
'    DebugPrintFunc ReadPatSet.Value
'
'    ''''20170111 Add
'    Call auto_eFuse_ReadAllData_to_DictDSPWave("MON", False, False)

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function
