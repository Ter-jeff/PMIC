Attribute VB_Name = "VBT_LIB_EFUSE_Config"
Option Explicit

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_ConfigBlankChk(ReadPatSet As Pattern, PinRead As PinList, _
                    Optional condstr As String = "", _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ConfigBlankChk"

    ''''-------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(ReadPatSet, ReadPatt, Validating_) = True) Then Exit Function
    ''''-------------------------------------------------------------------------------------------------
    
    Dim site As Variant
    Dim i As Long, j As Long, k As Long
    Dim BlankChkPass As Boolean
    Dim blank_firstbits As Boolean
    Dim blank_no57bit As Boolean
    Dim blank_stage_no64bits As Boolean
    Dim ResultFlag As New SiteLong
    Dim testName As String, PinName As String
    Dim m_jobinStage_flag As Boolean
    Dim m_stage_SingleDoubleFBC As Long
    Dim PrintSiteVarResult As String
    Dim SiteVarValue As Long
    Dim SingleBitArray() As Long
    Dim DoubleBitArray() As Long
    Dim SingleDoubleBitMismatch As New SiteLong
    Dim Count As Long
    Dim SignalCap As String, CapWave As New DSPWave
    Dim blank_stage As New SiteBoolean
    Dim allBlank As New SiteBoolean
    Dim leftStr As String
    Dim rightStr As String
    Dim bypass_flag As Boolean
    
    ''''20170630 update
    Dim blank_Cond As Boolean
    Dim blank_SCAN As Boolean
    Dim blank_stage_noCond_SCAN As Boolean

    ReDim gL_CFG_Sim_FuseBits(TheExec.sites.Existing.Count - 1, EConfigTotalBitCount - 1) ''''it's for the simulation
''    ''''''''<Important> Can NOT do the below declaration, otherwise the simulation bits (Early Bits) will be clear in this test.
''    ''''''''<MUST> using Preserve to reserve the previous simulation data here.
''    ReDim Preserve gL_CFG_Sim_FuseBits(TheExec.Sites.Existing.Count - 1, EConfigTotalBitCount - 1) ''''20170630 update, it's for the simulation

    SignalCap = "SignalCapture" 'define capture signal name

    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

    '****** Initialize Site Varaible ******
    Dim m_siteVar As String
    m_siteVar = "CFGChk_Var"
    For Each site In TheExec.sites.Existing
        TheExec.sites(site).SiteVariableValue(m_siteVar) = -1
    Next site

    '=============================================
    '=  Setup HRAM/DSSC capture cycles           =
    '=============================================
    '*** Setup HARM/DSSC Trigger and Capture parameter ***
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, EConfigReadCycle, CapWave  'setup
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)    'run read pattern and capture
    
    ''''20151229 New
    ''''It's used to check if 'gS_JobName' is existed in all CFGFuse programming stages
    ''''it's used to identify if the Job Name is existed in the CFG portion of the eFuse BitDef table.
    m_jobinStage_flag = auto_eFuse_JobExistInStage("CFG", True) ''''<MUST>

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
    Dim m_bitFlag_mode As Long

    'gDL_eFuse_Orientation = eFuse_2_Bit ''20190513
     gDL_eFuse_Orientation = gE_eFuse_Orientation
    gL_eFuse_Sim_Blank = 0
    
    m_Fusetype = eFuse_CFG
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_stage = True
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    If (LCase(condstr) = "stage") Then
        m_bitFlag_mode = 1
    ElseIf (LCase(condstr) = "real") Then
        m_bitFlag_mode = 3
    Else
        ''''default, here it prevents any typo issue
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
        m_FBC = -1
        'm_cmpResult = -1
    End If

    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, blank_stage, allBlank)
    ''''----------------------------------------------------

    If (blank_stage.Any(False) = True) Then
        ''''''''if there is any site which is non-blank, then decode to gDW_CFG_Read_Decimal_Cate [check later]
        ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
        ''''it will be present by Hex and Binary compare with the limit later on.
        ''''<NOTICE> gDW_CFG_Read_Decimal_Cate has been decided in non-blank and/or MarginRead
        ''Call auto_eFuse_setReadData(eFuse_CFG, gDW_CFG_Read_DoubleBitWave, gDW_CFG_Read_Decimal_Cate)
        'Call auto_eFuse_setReadData(eFuse_CFG)
        
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(eFuse_CFG)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_CFG)

        Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_CFG, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_CFG, False, gB_eFuse_printReadCate)
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
    m_SiteVarValue = blank_stage
    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
    If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only

    For Each site In TheExec.sites
        If (auto_eFuse_GetAllPatTestPass_Flag("CFG") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    gL_CFG_FBC = m_FBC
    
    Call UpdateDLogColumns(gI_CFG_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "CFG_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "CFG_Blank_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value

Exit Function

End If

'    ''''While being in the "Retest" stage, the DSSC Read result will be used in the instance "ConfigSingleDoubleBit"
'    ''''So it's needed to set SingleStrArray() to global gS_SingleStrArray()
'    auto_eFuse_DSSC_ReadDigCap_32bits EConfigReadCycle, PinRead.Value, gS_SingleStrArray, capWave, allblank 'read back in singlestrarray
'
'    Call UpdateDLogColumns(gI_CFG_catename_maxLen)
'    Count = 0 'Initialization
'    Call auto_GetSiteFlagName(Count, gS_cfgFlagname, False)
'    If (gB_findCFGCondTable_flag) Then
'        If (Count = 0) Then
'            gS_cfgFlagname = "ALL_0" ''''was "A00"
'        ElseIf (Count <> 1) Then
'            TheExec.Datalog.WriteComment vbCrLf & "<WARNING> There are more one CFG condition Flag selected. Please check it!! " ''''20160927 add
'            gS_cfgFlagname = "ALL_0"
'            SiteVarValue = 0
'            For Each Site In TheExec.Sites
'                TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'                ''''20161110 update per Jack's comment, "F_config_flag_missing" (BinTable Flag)
'                TheExec.Sites.Item(Site).FlagState("F_config_flag_missing") = logicTrue
'            Next Site
'            TheExec.Flow.TestLimit resultVal:=Count, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Error" ''''set fail
'            TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'            GoTo CFGChk_End
'        End If
'    Else
'        ''''20160902 update for the case CP1 fuse CFG_Condition already, then CP2/CPx needs to get the Flag name.
'        If (Count <> 1) Then
'            gS_cfgFlagname = "ALL_0"
'        End If
'    End If
'
'    ''''20170911 update
'    ''Call auto_display_CFG_Cond_Table_by_PKGName(gS_cfgFlagname)
'
'    For Each Site In TheExec.Sites
'        ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To (EConfigReadCycle - 1)
'                gS_SingleStrArray(i, Site) = StrReverse(gS_SingleStrArray(i, Site))
'            Next i
'        End If
'
'        ''''20160202 update, 20161108 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            allblank(Site) = False
'            blank_stage(Site) = True ''''True or False(sim for re-test)
'            If (blank_stage(Site) = True And gS_JobName = "cp1") Then allblank(Site) = True
'
'            Dim m_tmpStr As String
'            Dim Expand_eFuse_Pgm_Bit() As Long
'            Dim eFusePatCompare() As String
'            ReDim Expand_eFuse_Pgm_Bit(EConfigTotalBitCount * EConfig_Repeat_Cyc_for_Pgm - 1)
'            ReDim eFusePatCompare(EConfigReadCycle - 1)
'            ReDim SingleBitArray(EConfigTotalBitCount - 1)
'
'            ''''20170630 update
'            If (gB_findCFGCondTable_flag And gS_JobName = "cp1") Then ''''<MUST>
'                blank_Cond = False
'                blank_SCAN = False
'                blank_stage_noCond_SCAN = True
'                blank_stage(Site) = True
'                allblank(Site) = False
'
'                If (gB_eFuse_CFG_Cond_FTF_done_Flag = True) Then ''''20170923 add
'                    blank_stage(Site) = False
'                    allblank(Site) = False
'                End If
'
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                TheExec.Datalog.WriteComment vbTab & "[ blank_Cond = False, Simulation for the ReTest Mode on Job[" + UCase(gS_JobName) + "] ]"
'                Call eFuseENGFakeValue_Sim
'                ''''20170923 add
'                If (blank_stage = True) Then
'                    Call auto_make_CFG_Pgm_for_Simulation_Early(SingleBitArray, Expand_eFuse_Pgm_Bit, eFusePatCompare, blank_Cond, blank_stage_noCond_SCAN, False) ''''showPrint if True
'                Else
'                    Call auto_make_CFG_Pgm_for_Simulation(SingleBitArray, Expand_eFuse_Pgm_Bit, eFusePatCompare, blank_stage(Site), False) ''''showPrint if True
'                End If
'            ElseIf (blank_stage = False) Then ''''20160202, simulation for retest mode
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                TheExec.Datalog.WriteComment vbTab & "[ blank_stage = False, Simulation for the ReTest Mode on Job[" + UCase(gS_JobName) + "] ]"
'                Call eFuseENGFakeValue_Sim
'                Call auto_make_CFG_Pgm_for_Simulation(SingleBitArray, Expand_eFuse_Pgm_Bit, eFusePatCompare, blank_stage(Site), False) ''''showPrint if True
'            End If
'
'            ''''to composite gS_SingleStrArray(i, Site) for the simulation in SingleDoubleBit_Check/Syntax
'            For i = 0 To EConfigReadCycle - 1
'                m_tmpStr = ""
'                For j = 0 To EConfigReadBitWidth - 1
'                    k = j + i * EConfigReadBitWidth ''''MUST
'                    m_tmpStr = CStr(SingleBitArray(k)) + m_tmpStr
'                    gL_CFG_Sim_FuseBits(Site, k) = SingleBitArray(k)
'                Next j
'                gS_SingleStrArray(i, Site) = m_tmpStr
'            Next i
'
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        '====================================================
'        '=  Print the all eFuse Bit data from digCap        =
'        '====================================================
'        Call auto_EConfig_Memory_Read(gS_SingleStrArray, SingleBitArray, DoubleBitArray, BlankChkPass) ''''for allbits
'
'        ''''20151230 update
'        blank_stage = True ''''initial
'        m_stage_SingleDoubleFBC = 0
'        Call auto_eFuse_BlankChk_FBC_byStage("CFG", SingleBitArray, blank_stage, m_stage_SingleDoubleFBC) ''''for bits in (gS_jobName) stage
'
'        ''''<MUST>
'        ''''20170630 update
'        If (gB_findCFGCondTable_flag) Then
'            gB_CFG_blank_Cond(Site) = True ''''<MUST>
'            gB_CFG_blank_SCAN(Site) = True ''''<MUST>
'            Call auto_CFG_blank_check_Cond_SCAN(SingleBitArray, blank_Cond, blank_SCAN, blank_stage_noCond_SCAN)
'            gB_CFG_blank_Cond(Site) = blank_Cond
'            gB_CFG_blank_SCAN(Site) = blank_SCAN
'        End If
'
'        ''''----------------------------------------------------------------------------------
'        ''''<NOTICE>
'        ''''Because we use 'V' instead of 'L' in the DSSC pattern,
'        ''''so that "Blank-Fuse" can not be judged by the API 'Patgen.PatternBurstPassed'.
'        ''''But it was decided in the routine auto_eFuse_DSSC_ReadDigCap_32bits()
'        ''''If theHdw.digital.Patgen.PatternBurstPassed(Site) = False Then  'If not blank
'        ''''----------------------------------------------------------------------------------
'        bypass_flag = False
'
'        If (allblank(Site) = False) Then ''False means this Efuse is NOT balnk.
'
'            If (gS_JobName = gS_CFG_firstbits_stage And gB_findCFGCondTable_flag = True) Then ''''20170630 update
'                ''''Should NOT have this case
'                If (blank_Cond = True And blank_SCAN = False And blank_stage_noCond_SCAN = True) Then
'                    If (gS_JobName = gS_CFG_SCAN_stage) Then
'                        ResultFlag(Site) = 1  'Fail Blank check criterion
'                        PinName = "Fail"
'                        SiteVarValue = 0
'                    Else
'                        ResultFlag(Site) = 0  'Pass Blank check criterion
'                        PinName = "Pass"
'                        SiteVarValue = 1
'                    End If
'                ElseIf (blank_Cond = False And blank_SCAN = True And blank_stage_noCond_SCAN = True) Then
'                    ResultFlag(Site) = 1  'Fail Blank check criterion
'                    PinName = "Fail"
'                    SiteVarValue = 0
'                ElseIf (blank_Cond = True And blank_SCAN = True And blank_stage_noCond_SCAN = False) Then
'                    ResultFlag(Site) = 1  'Fail Blank check criterion
'                    PinName = "Fail"
'                    SiteVarValue = 0
'                ''ElseIf (blank_Cond = False And blank_SCAN = False And blank_stage_noCond_SCAN = True) Then
'                ElseIf (blank_stage_noCond_SCAN = True) Then
'                    ''''20170804 update, if gS_CFG_firstbits_stage='cp1', here is a case
'                    ResultFlag(Site) = 0  'Pass Blank check criterion
'                    PinName = "Pass"
'                    SiteVarValue = 1
'                End If
'
'            ElseIf (gS_JobName = gS_CFG_firstbits_stage And gB_findCFGCondTable_flag = False) Then ''''20150720, 20170630 update
'                ''''<NOTICE> In this Stage, it means that CFG(firstbits) will be blown and checked in detail.
'                ''''20151230 update
'                Call auto_CFG_blank_check_firstbits(SingleBitArray, blank_firstbits, blank_no57bit, True, blank_stage_no64bits)
'
'                ''''Get Config security code from OI
'                Count = 0 'Initialization
'                Call auto_GetSiteFlagName(Count, gS_cfgFlagname, True)
'
'                If (Count = 0) Then ''''it means that NO any assign flag CFG_XXX, so using ALL_0 as default
'                    gS_cfgFlagname = "ALL_0"  ''''20151229 update
'                    ResultFlag(Site) = 1    'Fail Blank check criterion
'                    SiteVarValue = 0
'                    testName = "CFG64bits_Blank_Chk"
'                    PinName = "Fail"
'                    blank_firstbits = False ''''force NOT to burn
'                    blank_no57bit = False   ''''force NOT to burn
'                    ''''20161110 update per Jack's comment, "F_config_flag_missing" (BinTable Flag)
'                    TheExec.Sites.Item(Site).FlagState("F_config_flag_missing") = logicTrue
'                    TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_No_Assign"
'                ElseIf (Count > 1) Then
'                    ResultFlag(Site) = 1   'Fail Blank check criterion
'                    SiteVarValue = 0
'                    testName = "CFG64bits_Blank_Chk"
'                    PinName = "Fail"
'                    blank_firstbits = False ''''force NOT to burn
'                    blank_no57bit = False   ''''force NOT to burn
'                    ''''20161110 update per Jack's comment, "F_config_flag_missing" (BinTable Flag)
'                    TheExec.Sites.Item(Site).FlagState("F_config_flag_missing") = logicTrue
'                    TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Error"
'                Else
'                    ''''csae Count==1
'                    TheExec.Datalog.WriteComment funcName + ":: The Selected CFG Condition is " + gS_cfgFlagname + " (from OI or Eng mode)"
'                    TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Set_" + gS_cfgFlagname
'
'                    ResultFlag(Site) = 0       'Pass Blank check criterion
'                    testName = "CFG64bits_Blank_Chk"
'                    PinName = "Pass"
'                    SiteVarValue = 2
'
'                    If ((blank_stage_no64bits = True) And (blank_firstbits = True)) Then
'                        SiteVarValue = 1
'                    ElseIf ((blank_stage_no64bits = True) And (blank_firstbits = False)) Then
'                        ''''case blank_firstbits==False (first 64 bits <> 0)
'                        ''''SiteVarValue = 2
'                        ''''<NOTICE> 20151231, Special Case
'                        If (gB_CFG_SVM = True) Then
'                            If ((blank_firstbits = False) And (blank_no57bit = True) And (UCase(gS_cfgFlagname) <> "A00")) Then
'                                ''''Case:: CFG_SVM is Enable, and already burn 'CFG_A00'
'                                ''''But would like to blow other conditions except 'A00'
'                                SiteVarValue = 1
'                            End If
'                        End If
'                    End If
'
'                    ''''Need to check overall FBC's result again
'                    If BlankChkPass = False Then  'If FBC<>0 (i.e. BlankChkPass=false) then
'                        ResultFlag(Site) = 1      'Fail Blank check criterion
'                        SiteVarValue = 0
'                        testName = "CFG_Blank_Chk"
'                        PinName = "Fail"
'                    End If
'                End If
'
'            ElseIf (gS_JobName Like "cp*" Or gS_JobName Like "ft*" Or gS_JobName = "wlft") Then ''''ex: cp1,cp2,ft1,ft2,wlft
'                ''''---------------------------------------
'                ''''1st Step:: get the CFG condition
'                ''''---------------------------------------
'                Count = 0 'Initialization
'                Call auto_GetSiteFlagName(Count, gS_cfgFlagname, True)
'
'                If (gS_JobName Like "ft*" Or gS_JobName = "wlft") Then
'                    ''''<Important> In FT* and WLFT stage,MUST have ONE CFG flag for the syntax check only
'                    If (Count = 1) Then
'                        TheExec.Datalog.WriteComment funcName + ":: The Selected CFG Condition is " + gS_cfgFlagname + " (from OI or Eng mode)"
'                        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Set_" + gS_cfgFlagname
'                    Else
'                        gS_cfgFlagname = "ALL_0"
'                        BlankChkPass = False ''''Force to Fail
'                        ''''20161110 update per Jack's comment, "F_config_flag_missing" (BinTable Flag)
'                        TheExec.Sites.Item(Site).FlagState("F_config_flag_missing") = logicTrue
'                        TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Error" ''''set fail here
'                    End If
'                ElseIf (gS_JobName Like "cp*") Then
'                    If (Count = 1) Then
'                        TheExec.Datalog.WriteComment funcName + ":: The Selected CFG Condition is " + gS_cfgFlagname + " (from OI or Eng mode)"
'                        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Set_" + gS_cfgFlagname
'                    ElseIf (Count > 1) Then
'                        ''''20161110 update per Jack's comment, "F_config_flag_missing" (BinTable Flag)
'                        TheExec.Sites.Item(Site).FlagState("F_config_flag_missing") = logicTrue
'                        TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Error" ''''set fail here
'                        BlankChkPass = False ''''Force to Fail
'                    ElseIf (Count = 0) Then
'                        ''''<Very Important> 20151230
'                        ''''In CP* stage, we set default all zero 'ALL_0' if no selection of CFG Flag.
'                        gS_cfgFlagname = "ALL_0"  ''''20151229 update
'                        If (gB_findCFGCondTable_flag) Then gS_cfgFlagname = "A00" ''''<MUST>
'                        ''If (gB_CFG_SVM = True) Then gS_cfgFlagname = "CFG_A00"
'                        TheExec.Datalog.WriteComment "Count=0, Set Default CFG Condition = " + gS_cfgFlagname
'                        ''''TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Set_all0" ''''set pass here
'                    End If
'                End If
'
'                ''''---------------------------------------
'                ''''2nd Step:: decide the SiteVarValue
'                ''''---------------------------------------
'                testName = "CFG_BlankChk_" + UCase(gS_JobName)
'
'                If BlankChkPass = True Then  'If FBC=0 (i.e. BlankChkPass=true) then
'                    ''''Here it's used to check if HardIP pattern test pass or not.
'                    If (auto_eFuse_GetAllPatTestPass_Flag("CFG") = False) Then
'                        ResultFlag(Site) = 1  'Fail Blank check criterion
'                        PinName = "Fail"
'                        SiteVarValue = 0
'                    Else
'                        ResultFlag(Site) = 0   'Pass Blank check criterion
'                        PinName = "Pass"
'                        ''''<MUST>
'                        ''''it's used to identify if the Job Name is existed in the CFG portion of the eFuse BitDef table.
'                        If (m_jobinStage_flag = False) Then
'                            ''''<Important> Then it will NOT go WritebyStage to let the user confusion.
'                            SiteVarValue = 2
'                        Else
'                            ''''m_jobinStage_flag is True
'                            ''''20170630 update
'                            If (gB_findCFGCondTable_flag) Then
'                                If (blank_stage_noCond_SCAN) Then
'                                    If (blank_Cond = False And blank_SCAN = False) Then
'                                        SiteVarValue = 1 ''''<1st> Here is most case
'                                    Else
'                                        ResultFlag(Site) = 1  'Fail Blank check criterion
'                                        PinName = "Fail"
'                                        SiteVarValue = 0
'                                        TheExec.Datalog.WriteComment "[WARNING] Site(" & Site & ") blank_Cond=" + CStr(blank_Cond) + ", blank_SCAN=" + CStr(blank_SCAN)
'                                    End If
'                                Else
'                                    SiteVarValue = 2
'                                End If
'                            ElseIf (blank_stage(Site) = True And m_stage_SingleDoubleFBC = 0) Then
'                                SiteVarValue = 1
'                            ElseIf (blank_stage(Site) = False And m_stage_SingleDoubleFBC = 0) Then
'                                SiteVarValue = 2
'                            Else
'                                ResultFlag(Site) = 1  'Fail Blank check criterion
'                                PinName = "Fail"
'                                SiteVarValue = 0
'                            End If
'                        End If
'                    End If
'                Else  'Fail BlankChk, BlankChkPass==False
'                    ResultFlag(Site) = 1   'Fail Blank check criterion
'                    SiteVarValue = 0
'                    PinName = "Fail"
'                End If
'
'            Else 'Char / HTOL / ...etc
'                bypass_flag = True
'                ResultFlag(Site) = 0   'Pass Blank check criterion
'                SiteVarValue = 2
'                testName = "CFG_BlankChk_" + UCase(gS_JobName)
'                PinName = "Pass"
'            End If
'
'        Else ''True means this Efuse is blank of allbits.
'
'            gB_CFGSVM_BIT_Read_ValueisONE(Site) = False ''''<MUST>
'            testName = "CFG_BlankChk_" + UCase(gS_JobName)
'            If gS_JobName = "cp1" Then
'                ''''Here it's used to check if HardIP pattern test pass or not.
'                If (auto_eFuse_GetAllPatTestPass_Flag("CFG") = False) Then
'                    ResultFlag(Site) = 1  'Fail Blank check criterion
'                    PinName = "Fail"
'                    SiteVarValue = 0
'                Else
'                    ResultFlag(Site) = 0   'Pass Blank check criterion
'                    PinName = "Pass"
'                    SiteVarValue = 1
'
'                    ''''<Important> 20151230
'                    If (gS_CFG_firstbits_stage = "cp1") Then
'                        'Get Config security code from OI/Eng
'                        Count = 0 'Initialization
'                        Call auto_GetSiteFlagName(Count, gS_cfgFlagname, True)
'                        testName = "CFG64bits_Blank_Chk"
'                        If (Count = 1) Then
'                            TheExec.Datalog.WriteComment funcName + ":: The Selected CFG Condition is " + gS_cfgFlagname + " (from OI or Eng mode)"
'                            TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Set_" + gS_cfgFlagname
'                        Else
'                            ''''case Count<>1
'                            gS_cfgFlagname = "ALL_0"
'                            ResultFlag(Site) = 1 'Set it to Fail
'                            SiteVarValue = 0
'                            PinName = "Fail"
'                            ''''20161110 update per Jack's comment, "F_config_flag_missing" (BinTable Flag)
'                            TheExec.Sites.Item(Site).FlagState("F_config_flag_missing") = logicTrue
'                            TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Error"
'                        End If
'                    End If
'                End If
'            Else 'Should NOT be All Blank at non-CP1 flow
'                ResultFlag(Site) = 1  'Fail Blank check criterion
'                PinName = "Fail"
'                SiteVarValue = 0
'            End If
'        End If
'
'        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'
'        If (SiteVarValue <> 1) Then
'            TheExec.Datalog.WriteComment ""
'            TheExec.Datalog.WriteComment "Read All Config eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'            Call auto_PrintAllBitbyDSSC(SingleBitArray, EConfigReadCycle, EConfigTotalBitCount, EConfigReadBitWidth) ''must use EConfigReadBitWidth in Right2Left mode
'        Else
'            ''''20161114 update
'            Call auto_Decode_CfgBinary_Data(DoubleBitArray, False)
'        End If
'
'        '====================================================================
'        '=      Check if right and left half block data are consistent     =
'        '====================================================================
'        'This is redundant for CP but necessary for FT in case eFuse bit flip
'        SingleDoubleBitMismatch(Site) = 0   'Deault is Data match if allBlank=True or bypass_flag=True
'        If ((allblank(Site) = False) And (bypass_flag = False)) Then
'            If (gS_EFuse_Orientation = "UP2DOWN") Then
'                For i = 0 To EConfigReadCycle / 2 - 1
'                    If gS_SingleStrArray(i, Site) <> gS_SingleStrArray(i + EConfigReadCycle / 2, Site) Then
'                        SingleDoubleBitMismatch(Site) = 1   'Data mismatch
'                        Exit For
'                    End If
'                Next i
'            ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
'                For i = 0 To EConfigReadCycle - 1
'                    rightStr = Mid(gS_SingleStrArray(i, Site), EConfigBitsPerRow + 1, EConfigBitsPerRow)
'                    leftStr = Mid(gS_SingleStrArray(i, Site), 1, EConfigBitsPerRow)
'                    If (rightStr <> leftStr) Then
'                        SingleDoubleBitMismatch(Site) = 1   'Data mismatch
'                        Exit For
'                    End If
'                Next i
'            ElseIf (gS_EFuse_Orientation = "SingleUp") Then
'                ''''doNothing, becaause there is only one block
'                SingleDoubleBitMismatch(Site) = 0 ''''keep default
'            End If
'        End If
'
'        If (False) Then
'            PrintSiteVarResult = "Site (" + CStr(Site) + ") " + m_siteVar + " = " + CStr(SiteVarValue)
'            TheExec.Datalog.WriteComment PrintSiteVarResult & vbCrLf
'        End If
'
'        gB_CFG_decode_flag(Site) = False ''''20160531, <MUST> reset and init
'
'        ''Binning out
'        TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'        TheExec.Flow.TestLimit resultVal:=SingleDoubleBitMismatch, lowVal:=0, hiVal:=0, Tname:="Chk_Mismatch"
'        TheExec.Flow.TestLimit resultVal:=ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName, PinName:=PinName
'
'        If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
'
'    Next Site
'
'CFGChk_End:
'    Call UpdateDLogColumns__False
'    DebugPrintFunc ReadPatSet.Value
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_ConfigWrite_byCondition(WritePattSet As Pattern, PinWrite As PinList, _
                    PwrPin As String, vpwr As Double, _
                    condstr As String, _
                    Optional catename_grp As String, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ConfigWrite_byCondition"

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
    Dim Expand_eFuse_Pgm_Bit() As Long
    Dim eFusePatCompare() As String
    Dim i As Long
    Dim SegmentSize As Long

    'Dim DigSrcSignalName As String
    Dim DigSrcSignalName() As String
    Dim Expand_Size As Long

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

    TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName
    
  
    '====================================================
    '=  1. Define the Bit length of Programming         =
    '=  2. Define the cycle number of Compare Pattern   =
    '====================================================
    Expand_Size = (EConfigTotalBitCount * EConfig_Repeat_Cyc_for_Pgm) 'Because there are repeat cycle in C651 pattern, we have to create multiple DSSC
    ReDim Expand_eFuse_Pgm_Bit(Expand_Size - 1)
    ReDim eFusePatCompare(EConfigReadCycle - 1)
    ReDim gL_CFGFuse_Pgm_Bit(TheExec.sites.Existing.Count - 1, EConfigTotalBitCount - 1)
    
    '========================================================
    '=  1. Make DigSrc Pattern for eFuse programming        =
    '=  2. Make Read Pattern for eFuse Read tests           =
    '========================================================
  
    'DigSrcSignalName = "CFG_DigSrcSignal"
    
    condstr = LCase(condstr) ''''<MUST>
    
    Dim PatAry() As String
    Dim m_PatAry() As String
    Dim PatCnt As Long
    Dim k As Long
    Dim m_patset As New Pattern
    Dim m_patValidateCnt As Long
    PatAry = TheExec.DataManager.Raw.GetPatternsInSet(WritePattSet, PatCnt)
    ReDim DigSrcSignalName(PatCnt - 1)
    
    If (PatCnt <> 0) Then    ''' 20171218 [JH] Due to split original CFG programming pattern into 2 patterns (4k (HF) & 8k (ALL))
        ReDim m_PatAry(PatCnt - 1)
        For k = 0 To UBound(PatAry)
            ''''----------------------------------------------------------------------------------------------------
            ''''<Important>
            ''''Must be put before all implicit array variables, otherwise the validation will be error.
            '==================================
            '=  Validate/Load Write patterns  =
            '==================================
            ''''20161114 update to Validate/load pattern
            m_patset.Value = PatAry(k)
            If (auto_eFuse_PatSetToPat_Validation(m_patset, WritePatt, Validating_) = True) Then m_patValidateCnt = m_patValidateCnt + 1
            ''''----------------------------------------------------------------------------------------------------
            m_PatAry(k) = WritePatt
            DigSrcSignalName(k) = "CFG_DigSrcSignal_" & CStr(k)
        Next k
        WritePatt = m_PatAry(0) ''''set as default
    End If
    If (Validating_) Then Exit Function
    

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''201808XX update
    If (TheExec.TesterMode = testModeOffline) Then
        If (condstr <> "cp1_early") Then
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
    For i = 0 To UBound(CFGFuse.Category)
        With CFGFuse.Category(i)
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
                    Call auto_eFuse_Vddbin_bincut_setWriteData(eFuse_CFG, i)
                End If
                ''''---------------------------------------------------------------------------
                With CFGFuse.Category(i)
                    ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
                    Call rundsp.eFuse_updatePgmWave_byCategory(.BitIndexWave, .Write.BitArrWave)
                End With
                ''''---------------------------------------------------------------------------
            End If
        Else
            ''''doNothing
        End If
    Next i
    
    ''''process CRC bits calculation
    If (gS_CFG_CRC_Stage = gS_JobName) Then
        Dim mSL_bitwidth As New SiteLong
        mSL_bitwidth = gL_CFG_CRC_BitWidth
        ''''CRC case
        With CFGFuse.Category(m_crc_idx)
            ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
            Call rundsp.eFuse_updatePgmWave_CRCbits(eFuse_CFG, mSL_bitwidth, .BitIndexWave)
        End With
    End If
    
    ''''composite effective PgmBits per Stage requirement
    m_pgmRes = 0
    
    Dim m_SampleSize As Long:: m_SampleSize = gDL_TotalBits
    Dim m_ReadCycle As Long:: m_ReadCycle = gDL_ReadCycles

    If (m_cmpStage = "cp1_early") Then
        ''''condStr = "cp1_early"
        Call auto_eFuse_Gen_PgmBitSrcWave_onlyDeaultBits(eFuse_CFG, m_pgmDigSrcWave, m_pgmRes)
    ElseIf (gB_EFUSE_DVRV_ENABLE And m_cmpStage <> "cp1_early") Then
        Call rundsp.eFuse_Gen_PgmBitSrcWave_OnlyRV(eFuse_CFG, 1, gL_CFG_SegCNT, m_pgmDigSrcWave, m_pgmRes)
        For Each site In TheExec.sites.Active
            m_SampleSize = m_pgmDigSrcWave(site).SampleSize
            Exit For
        Next
        m_ReadCycle = gL_CFG_SegCNT
    Else
        ''''condStr = "stage"
        ''''Here Parameter bitFlag_mode=1 (Stage Bits)
        m_calcCRC = 0
        Call rundsp.eFuse_Gen_PgmBitSrcWave(eFuse_CFG, 1, m_pgmDigSrcWave, m_pgmRes)
    End If
    TheHdw.Wait 1# * ms
    
    Call UpdateDLogColumns(gI_CFG_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=m_pgmRes, lowVal:=1, hiVal:=1, Tname:="CFG_PGM_" + UCase(gS_JobName)
    Call UpdateDLogColumns__False

    ''''use global DSP variable
    Call auto_eFuse_print_PgmBitsWave_BitMap(m_pgmDigSrcWave, m_SampleSize, m_ReadCycle, gB_eFuse_printBitMap)
    'Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_CFG_Pgm_SingleBitWave, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_CFG, gDW_CFG_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)
    
    Call TurnOnEfusePwrPins(PwrPin, vpwr)
    
    Dim m_SplitDigSrcWave As New DSPWave
    Dim m_SplitSampleSize As Long
    
    For k = 0 To UBound(PatAry)
        For Each site In TheExec.sites
            m_SplitSampleSize = m_pgmDigSrcWave(site).SampleSize / PatCnt
            m_SplitDigSrcWave.CreateConstant 0, m_SplitSampleSize, DspLong
            m_SplitDigSrcWave = m_pgmDigSrcWave(site).Select(k * m_SplitSampleSize, 1, m_SplitSampleSize).Copy
            If (m_SplitDigSrcWave(site).CalcSum <> 0) Then
                If (m_cmpStage = "cp1_early") Then
                    ''''if it's same values on all Sites to save TT and improve PTE
                    Call eFuse_DSSC_SetupDigSrcWave_allSites(m_PatAry(k), PinWrite, DigSrcSignalName(k), m_SplitDigSrcWave(site))
                Else
                    Call eFuse_DSSC_SetupDigSrcWave(m_PatAry(k), PinWrite, DigSrcSignalName(k), m_SplitDigSrcWave(site))
                End If
                Call TheHdw.Patterns(m_PatAry(k)).Test(pfAlways, 0)   'Write ECID
            End If
        Next site
    Next


    ''''In the MarginRead process, it will use gDW_XXX_Pgm_SingleBitWave / gDW_XXX_Pgm_DoubleBitWave to do the comparison with Read

    ''''Write Pattern for programming eFuse
    'Call TheHdw.Patterns(WritePatt).test(pfAlways, 0)   'Write ECID

    Call TurnOffEfusePwrPins(PwrPin, vpwr)
    DebugPrintFunc WritePattSet.Value
    
Exit Function
    
End If

'
'    Call UpdateDLogColumns(gI_CFG_catename_maxLen)
'
'    For Each Site In TheExec.Sites
'
'        If (TheExec.TesterMode = testModeOffline) Then ''''20160526 update
'            ''If (condStr <> "cp1_early") Then Call eFuseENGFakeValue_Sim ''''20170630 update
'            Call eFuseENGFakeValue_Sim
'        End If
'
'        ''''20151229 New
'        If (condstr = "stage") Then
'            SegmentSize = auto_Make_EConfig_Pgm_and_Read_Array(eFuse_Pgm_Bit(), Expand_eFuse_Pgm_Bit(), eFusePatCompare())
'
'        ElseIf (condstr = "cp1_early") Then ''''20170630, only Cond/scan
'            SegmentSize = auto_Make_EConfig_Pgm_CP1_Early(eFuse_Pgm_Bit(), Expand_eFuse_Pgm_Bit(), eFusePatCompare())
'            ''''TTR, because it's same value for all sites
'            Exit For
'
'        ElseIf (condstr = "category") Then
'            SegmentSize = auto_Make_EConfig_Pgm_and_Read_Array_byCategory(catename_grp, Expand_eFuse_Pgm_Bit(), eFuse_Pgm_Bit(), eFusePatCompare())
'        Else
'            ''''default=all bits are 0, here it prevents any typo issue
'            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (CP1_Early,Stage,Category)"
'            For i = 0 To EConfigTotalBitCount - 1
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
'        If SegmentSize <> -1 Then
'            DSSC_SetupDigSrcWave WritePatt, PinWrite, DigSrcSignalName, SegmentSize, Expand_eFuse_Pgm_Bit
'
'            'Print out programming bits
'            'ECfgPrintPgm eFuse_Pgm_Bit() ' Expand_eFuse_Pgm_Bit(0) is Row 0 and bit 0. Expand_eFuse_Pgm_Bit(1) is Row 0 and bit 1
'            Call auto_PrintAllPgmBits(eFuse_Pgm_Bit(), EConfigReadCycle, EConfigTotalBitCount, EConfigReadBitWidth) ''must use EConfigReadBitWidth
'
'            For i = 0 To EConfigTotalBitCount - 1
'                gL_CFGFuse_Pgm_Bit(Site, i) = eFuse_Pgm_Bit(i)
'            Next i
'        End If ''If SegmentSize <> -1 Then
'    Next Site
'    If (condstr <> "cp1_early") Then TheHdw.DSSC.Pins(PinWrite).Pattern(WritePatt).Source.Signals.DefaultSignal = DigSrcSignalName
'    Call UpdateDLogColumns__False
'
'    ''''--------------------------------------------------------------------------------------------------
'    ''''20180522 reserve for the future
'    ''''The statement "DSSC_SetupDigSrcWave" in previous SiteLoop should be skip in using the below method.
'    ''''---------------------------------------------------------------------------------------------------
'    If (True) Then
'        If (condstr = "cp1_early") Then ''''For TTR Purpose, if the pgm values are same on all Sites.
'            ''''For Marginal Read purpose
'            For Each Site In TheExec.Sites
'                For i = 0 To EConfigTotalBitCount - 1
'                    gL_CFGFuse_Pgm_Bit(Site, i) = eFuse_Pgm_Bit(i)
'                Next i
'                ''''could be masked
'                If (False) Then Call auto_PrintAllPgmBits(eFuse_Pgm_Bit(), EConfigReadCycle, EConfigTotalBitCount, EConfigReadBitWidth) ''must use EConfigReadBitWidth
'            Next Site
'            eFuse_DSSC_SetupDigSrcArr_allSites WritePatt, PinWrite, DigSrcSignalName, SegmentSize, Expand_eFuse_Pgm_Bit
'        End If
'    End If
'    ''''---------------------------------------------------------------------------------------------------
'
'    'Step2. Write Pattern for programming eFuse
'    TheHdw.Wait 0.0001
'    Call TheHdw.Patterns(WritePatt).test(pfAlways, 0)   'Write EConfig
'    DebugPrintFunc WritePattSet.Value
'
'    Call TurnOffEfusePwrPins(PwrPin, vpwr)

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_ConfigRead_by_OR_2Blocks(ReadPatSet As Pattern, PinRead As PinList, _
                    condstr As String, _
                    Optional catename_grp As String = "", _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ConfigRead_by_OR_2Blocks"

    ''''-------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(ReadPatSet, ReadPatt, Validating_) = True) Then Exit Function
    ''''-------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim eFuse_Pgm_Bit() As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim m_tmpStr As String
    Dim Count As Long, FailCnt As New SiteLong
    Dim testName As String

    Dim SingleStrArray() As String
    Dim DoubleBitArray() As Long
    Dim SingleBitArray() As Long

    Dim CapWave As New DSPWave
    Dim blank As New SiteBoolean
    Dim SignalCap As String
    Dim crcBinStr As String     ''''' 20161003 ADD CRC

    ''ReDim SingleStrArray(EConfigReadCycle - 1, TheExec.Sites.Existing.Count - 1)
    ''ReDim SingleBitArray(EConfigTotalBitCount - 1)
    ''ReDim DoubleBitArray(EConfigBitPerBlockUsed - 1)

    condstr = LCase(condstr) ''''<MUST>

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
  
    '====================================================
    '=  1. Define the Bit length of Programming         =
    '=  2. Define the cycle number of Compare Pattern   =
    '====================================================
    ReDim eFuse_Pgm_Bit(EConfigTotalBitCount - 1)
  
    '================================================
    '=  In Fiji, C651 ask we to OR block 1          =
    '=  and block2 before compare with programmed   =
    '=  data (confirmed on 6/28 morning meeting     =
    '================================================
    SignalCap = "SignalCapture"
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, EConfigReadCycle, CapWave
    
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)
    ''''testName = "CFG_ORMarginRead" + "_" + UCase(gS_JobName)

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

    m_Fusetype = eFuse_CFG
    m_FBC = -1       ''''init to failure
    m_cmpResult = -1 ''''init to failure

    ''''--------------------------------------------------------------------------
    '''' Offline Simulation Start                                                |
    ''''--------------------------------------------------------------------------
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        Call rundsp.eFuse_SingleBitWave2CapWave32Bits(eFuse_CFG, CapWave)
        Call auto_eFuse_print_capWave32Bits(eFuse_CFG, CapWave, False) ''''True to print out
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
    ElseIf (gB_EFUSE_DVRV_ENABLE = True Or condstr = "real") Then
        m_bitFlag_mode = 3
    Else
        ''''default, here it prevents any typo issue
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
        m_FBC = -1
        m_cmpResult = -1
    End If

    Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, m_cmpResult)

    gL_CFG_FBC = m_FBC

    ''''''[NOTICE] Decode and Print have moved to SingleDoubleBit()

    Call UpdateDLogColumns(gI_CFG_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=m_cmpResult, lowVal:=0, hiVal:=0
    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value

Exit Function

End If



    
     
'    auto_eFuse_DSSC_ReadDigCap_32bits EConfigReadCycle, PinRead.Value, SingleStrArray, capWave, blank ''''Here Must use local variable
'
'    Call UpdateDLogColumns(gI_CFG_catename_maxLen)
'
'    '================================================
'    '=  1. Make Program bit array                   =
'    '=  2. Make Read Compare bit array              =
'    '================================================
'    For Each Site In TheExec.Sites
'        ''''20151026 add, 20161107 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            ''''------------------------------------------------------
'            ''''20161107 update
'            Call auto_Decompose_StrArray_to_BitArray("CFG", gS_SingleStrArray, SingleBitArray, 0)
'            ''''------------------------------------------------------
'
'            ''''to composite gS_SingleStrArray(i, Site) for the simulation in SingleDoubleBit_Check/Syntax
'            For i = 0 To EConfigReadCycle - 1
'                m_tmpStr = ""
'                For j = 0 To EConfigReadBitWidth - 1
'                    k = j + i * EConfigReadBitWidth ''''MUST
'                    If (SingleBitArray(k) = 0 And gL_CFGFuse_Pgm_Bit(Site, k) = 1) Then
'                        SingleBitArray(k) = gL_CFGFuse_Pgm_Bit(Site, k)
'                        gL_CFG_Sim_FuseBits(Site, k) = SingleBitArray(k)
'                    End If
'                    m_tmpStr = CStr(SingleBitArray(k)) + m_tmpStr
'                Next j
'                gS_SingleStrArray(i, Site) = m_tmpStr
'                SingleStrArray(i, Site) = m_tmpStr
'            Next i
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        ''''gS_SingleStrArray() can be used later in auto_ConfigSingleDoubleBit()
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To EConfigReadCycle - 1
'                ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'                SingleStrArray(i, Site) = StrReverse(SingleStrArray(i, Site))
'                gS_SingleStrArray(i, Site) = SingleStrArray(i, Site)
'            Next i
'        Else
'            For i = 0 To EConfigReadCycle - 1
'                gS_SingleStrArray(i, Site) = SingleStrArray(i, Site)
'            Next i
'        End If
'
'        If (gB_findCFGCondTable_flag) Then
'            ''''20170630, At present, doNothing
'        Else
'            ''''---------------------------------------------------------------------------------------------
'            ''''20160905 update for the case which SVM 57bit is already '1' in the previous Stage
'            ''''but Pgm set to '0' in current Job.
'            ''''Define gC_CFGSVM_BIT = 57 in the module LIB_EFUSE_Custom
'            If (gS_JobName = gS_CFG_firstbits_stage And gB_CFG_SVM = True) Then
'                Dim m_pgmIndex As Long
'                Dim m_pgmIndex_sym As Long ''''symmetrical position
'
'                If (gB_CFGSVM_BIT_Read_ValueisONE(Site) = True) Then
'                    If (gS_EFuse_Orientation = "RIGHT2LEFT") Then
'                        m_pgmIndex = (EConfigReadBitWidth * Fix(gC_CFGSVM_BIT / EConfigBitsPerRow)) + (gC_CFGSVM_BIT Mod EConfigBitsPerRow) ''106=(32*Fix(57/16))+(57 mod 16) ''''(57 mod 16)=9
'                        m_pgmIndex_sym = m_pgmIndex + EConfigBitsPerRow ''105+16=121
'                        gL_CFGFuse_Pgm_Bit(Site, m_pgmIndex) = 1
'                        gL_CFGFuse_Pgm_Bit(Site, m_pgmIndex_sym) = 1
'                    ElseIf (gS_EFuse_Orientation = "SingleUp") Then
'                        m_pgmIndex = gC_CFGSVM_BIT
'                        gL_CFGFuse_Pgm_Bit(Site, m_pgmIndex) = 1
'                    ElseIf (gS_EFuse_Orientation = "UP2DOWN") Then
'                        m_pgmIndex = gC_CFGSVM_BIT
'                        m_pgmIndex_sym = gC_CFGSVM_BIT + EConfigBitPerBlockUsed
'                        gL_CFGFuse_Pgm_Bit(Site, m_pgmIndex) = 1
'                        gL_CFGFuse_Pgm_Bit(Site, m_pgmIndex_sym) = 1
'                    End If
'                End If
'            End If
'        End If
'        ''''---------------------------------------------------------------------------------------------
'
'        '================================================
'        '=  Compare EFuse Read with EFuse Program       =
'        '================================================
'
'        Call auto_OR_2Blocks("CFG", SingleStrArray, SingleBitArray(), DoubleBitArray())  ''''calc gL_CFG_FBC
'
'        For i = 0 To EConfigTotalBitCount - 1
'            eFuse_Pgm_Bit(i) = gL_CFGFuse_Pgm_Bit(Site, i)
'        Next i
'
'        FailCnt(Site) = 0
'
'        ''''20151222 Update
'        If (condstr = "all") Then
'            Call auto_CFGCompare_DoubleBit_PgmBit_byAll(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = "cp1_early") Then
'            ''''20170630 only for Cond/SCAN
'            Call auto_CFGCompare_DoubleBit_PgmBit_byStage_Early(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = "stage") Then
'            Call auto_CFGCompare_DoubleBit_PgmBit_byStage(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = "category") Then
'            Call auto_eFuse_Compare_DoubleBit_PgmBit_byCategory("CFG", catename_grp, DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        Else
'            ''''default, here it prevents any typo issue
'            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (All,CP1_Early,Stage,Category)"
'            FailCnt(Site) = -1
'        End If
'
'        '====================================================
'        '=  Print the all eFuse Bit data from DSSC          =
'        '====================================================
'        TheExec.Datalog.WriteComment ""
'        TheExec.Datalog.WriteComment "Read All Config eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'        Call auto_PrintAllBitbyDSSC(SingleBitArray(), EConfigReadCycle, EConfigTotalBitCount, EConfigReadBitWidth) ''must use EConfigReadBitWidth in Right2Left mode
'
'        '========================================================
'        '=  1. Decode binary data to meaningful decimal data    =
'        '=  2. Wrap Up the string for writing to HKEY           =
'        '========================================================
'        ''''' 20161003 ADD CRC
'        If (Trim(gS_CFG_CRC_Stage) <> "") Then
'            If (checkJob_less_Stage_Sequence(gS_CFG_CRC_Stage) = True) Then
'                gS_CFG_CRC_HexStr(Site) = "0000"
'            Else
'                gS_CFG_CRC_HexStr(Site) = auto_CFG_CRC2HexStr(DoubleBitArray, crcBinStr)
'            End If
'        End If
'
'        ''''<Important> User Need to check the content inside
'        Call auto_Decode_CfgBinary_Data(DoubleBitArray)
'
'    Next Site ''For Each Site In TheExec.Sites
'
'    TheExec.Flow.TestLimit resultVal:=FailCnt, lowVal:=0, hiVal:=0 '''', Tname:=testName
'
'    Call UpdateDLogColumns__False
'    DebugPrintFunc ReadPatSet.Value
'
    ''''----------------------------------------------------------------------------------------------------------
    ''''<Notice> User Maintain
''''    TheHdw.Wait 0.0001
''''    If (False) Then
''''        Call checkIEDAString("eFuseSOCTRIM1", SOCTrim1Str, SOCTrim1Str)
''''        Call checkIEDAString("eFuseSOCTRIM2", SOCTrim2Str, SOCTrim2Str)
''''        Call checkIEDAString("eFuseIDSCPU", IDS_CPU_Str, IDS_CPU_Str)
''''        Call checkIEDAString("eFuseIDSSOC", IDS_SOC_Str, IDS_SOC_Str)
''''        Call checkIEDAString("eFuseIDSGPU", IDS_GPU_Str, IDS_GPU_Str)
''''        Call checkIEDAString("eFuseIDSSRAM", IDS_SRAM_Str, IDS_SRAM_Str)
''''        Call checkIEDAString("eFuseBINSRAMMD1", BIN_SRAM_MODE1_Str, BIN_SRAM_MODE1_Str)
''''        Call checkIEDAString("eFuseBINSRAMMD2", BIN_SRAM_MODE2_Str, BIN_SRAM_MODE2_Str)
''''        Call checkIEDAString("eFuseBINGPUMD1", BIN_GPU_MODE1_Str, BIN_GPU_MODE1_Str)
''''        Call checkIEDAString("eFuseBINGPUMD2", BIN_GPU_MODE2_Str, BIN_GPU_MODE2_Str)
''''        Call checkIEDAString("eFuseBINGPUMD3", BIN_GPU_MODE3_Str, BIN_GPU_MODE3_Str)
''''        Call checkIEDAString("eFuseBINGPUMD4", BIN_GPU_MODE4_Str, BIN_GPU_MODE4_Str)
''''        Call checkIEDAString("eFuseBINSOCMD1", BIN_SOC_MODE1_Str, BIN_SOC_MODE1_Str)
''''
''''        theExec.Datalog.WriteComment vbCrLf & "Test Instance   :: " + InstanceName
''''        theExec.Datalog.WriteComment "eFuseSOCTRIM1   :: " + SOCTrim1Str
''''        theExec.Datalog.WriteComment "eFuseSOCTRIM2   :: " + SOCTrim2Str
''''        theExec.Datalog.WriteComment "eFuseIDSCPU     :: " + IDS_CPU_Str
''''        theExec.Datalog.WriteComment "eFuseIDSSOC     :: " + IDS_SOC_Str
''''        theExec.Datalog.WriteComment "eFuseIDSGPU     :: " + IDS_GPU_Str
''''        theExec.Datalog.WriteComment "eFuseIDSSRAM    :: " + IDS_SRAM_Str
''''        theExec.Datalog.WriteComment "eFuseBINSRAMMD1 :: " + BIN_SRAM_MODE1_Str
''''        theExec.Datalog.WriteComment "eFuseBINSRAMMD2 :: " + BIN_SRAM_MODE2_Str
''''        theExec.Datalog.WriteComment "eFuseBINGPUMD1  :: " + BIN_GPU_MODE1_Str
''''        theExec.Datalog.WriteComment "eFuseBINGPUMD2  :: " + BIN_GPU_MODE2_Str
''''        theExec.Datalog.WriteComment "eFuseBINGPUMD3  :: " + BIN_GPU_MODE3_Str
''''        theExec.Datalog.WriteComment "eFuseBINGPUMD4  :: " + BIN_GPU_MODE4_Str
''''        theExec.Datalog.WriteComment "eFuseBINSOCMD1  :: " + BIN_SOC_MODE1_Str
''''        ''TheExec.Datalog.WriteComment vbCrLf
''''
''''        '===================================================
''''        '=  Write Config Data to Register Edit (HKEY)      =
''''        '===================================================
''''        Call RegKeySave("eFuseSOCTRIM1", SOCTrim1Str)
''''        Call RegKeySave("eFuseSOCTRIM2", SOCTrim2Str)
''''        Call RegKeySave("eFuseIDSCPU", IDS_CPU_Str)
''''        Call RegKeySave("eFuseIDSSOC", IDS_SOC_Str)
''''        Call RegKeySave("eFuseIDSGPU", IDS_GPU_Str)
''''        Call RegKeySave("eFuseIDSSRAM", IDS_SRAM_Str)
''''        Call RegKeySave("eFuseBINSRAMMD1", BIN_SRAM_MODE1_Str)
''''        Call RegKeySave("eFuseBINSRAMMD2", BIN_SRAM_MODE2_Str)
''''        Call RegKeySave("eFuseBINGPUMD1", BIN_GPU_MODE1_Str)
''''        Call RegKeySave("eFuseBINGPUMD2", BIN_GPU_MODE2_Str)
''''        Call RegKeySave("eFuseBINGPUMD3", BIN_GPU_MODE3_Str)
''''        Call RegKeySave("eFuseBINGPUMD4", BIN_GPU_MODE4_Str)
''''        Call RegKeySave("eFuseBINSOCMD1", BIN_SOC_MODE1_Str)
''''    End If
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_ConfigSingleDoubleBit(Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ConfigSingleDoubleBit"
    
    Dim site As Variant
    Dim i As Long, j As Long, k As Long
    Dim DoubleBitArray() As Long
    Dim SingleBitArray() As Long
    Dim tmpStr As String
    Dim crcBinStr As String
    Dim m_siteVar As String
    m_siteVar = "CFGChk_Var"
    ''''--------------------------------------------------------------------------
    
    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

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
    Call auto_eFuse_setReadData_forSyntax(eFuse_CFG)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_CFG)
    
    ''''All the read action has been down in blank and/or MarginRead
    ''''gDW_CFG_Read_cmpsgWavePerCyc used to display the cmpare result (2-bit mode)
    Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_CFG, gB_eFuse_printBitMap)
    If (gS_JobName = "cp1_early") Then
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_CFG, True, gB_eFuse_printReadCate)
    Else
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_CFG, False, gB_eFuse_printReadCate)
    End If
    
    
    ''''Print CRC calcBits information
    Dim m_crcBitWave As New DSPWave
    Dim mS_hexStr As New SiteVariant
    Dim mS_bitStrM As New SiteVariant
    Dim m_debugCRC As Boolean
    Dim m_cnt As Long
    m_debugCRC = False
    
    ''''<MUST> Initialize
    gS_CFG_Read_calcCRC_hexStr = "0x0000"
    gS_CFG_Read_calcCRC_bitStrM = ""
    CRC_Shift_Out_String = ""
    If (auto_eFuse_check_Job_cmpare_Stage(gS_CFG_CRC_Stage) >= 0) Then
        Call rundsp.eFuse_Read_to_calc_CRCWave(eFuse_CFG, gL_CFG_CRC_BitWidth, m_crcBitWave)
        TheHdw.Wait 1# * ms ''''check if it needs
        
        If (m_debugCRC = False) Then
            ''''Here get gS_CFG_Read_calcCRC_hexStr for the syntax check
            Call auto_eFuse_bitWave_to_binStr_HexStr(m_crcBitWave, gS_CFG_Read_calcCRC_bitStrM, gS_CFG_Read_calcCRC_hexStr, True, m_debugCRC)
        Else
            ''''m_debugCRC=True => Debug purpose for the print
            TheExec.Datalog.WriteComment "------Read CRC Category Result------"
            Call auto_eFuse_bitWave_to_binStr_HexStr(m_crcBitWave, gS_CFG_Read_calcCRC_bitStrM, gS_CFG_Read_calcCRC_hexStr, True, m_debugCRC)
            TheExec.Datalog.WriteComment ""

            ''''[Pgm CRC calcBits] only gS_CFG_CRC_Stage=Job and CFGChk_Var=1
            If (gS_CFG_CRC_Stage = gS_JobName) Then
                m_cnt = 0
                TheExec.Datalog.WriteComment "------Pgm CRC calcBits------"
                For Each site In TheExec.sites
                    'If (TheExec.Sites(Site).SiteVariableValue(m_siteVar) = 1) Then
                        If (m_cnt = 0) Then TheExec.Datalog.WriteComment "------Pgm CRC calcBits------"
                        Call auto_eFuse_bitWave_to_binStr_HexStr(gDW_Pgm_BitWaveForCRCCalc, mS_bitStrM, mS_hexStr, False, m_debugCRC)
                        m_cnt = m_cnt + 1
                    'End If
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
'            DoubleBitArray = gDW_CFG_Read_DoubleBitWave.Data
'
'            gS_CFG_Direct_Access_Str(Site) = "" ''''is a String [(bitLast)......(bit0)]
'
'            For i = 0 To UBound(DoubleBitArray)
'                gS_CFG_Direct_Access_Str(Site) = CStr(DoubleBitArray(i)) + gS_CFG_Direct_Access_Str(Site)
'            Next i
'            ''TheExec.Datalog.WriteComment "gS_CFG_Direct_Access_Str=" + CStr(gS_CFG_Direct_Access_Str(Site))
'
'            ''''20161114 update for print all bits (DTR) in STDF
'            Call auto_eFuse_to_STDF_allBits("Config", gS_CFG_Direct_Access_Str(Site))
'        Next Site
'    End If
    
        For Each site In TheExec.sites
            DoubleBitArray = gDW_CFG_Read_DoubleBitWave.Data
            
            gS_CFG_Direct_Access_Str(site) = "" ''''is a String [(bitLast)......(bit0)]
            
            For i = 0 To UBound(DoubleBitArray)
                gS_CFG_Direct_Access_Str(site) = CStr(DoubleBitArray(i)) + gS_CFG_Direct_Access_Str(site)
            Next i
            ''TheExec.Datalog.WriteComment "gS_CFG_Direct_Access_Str=" + CStr(gS_CFG_Direct_Access_Str(Site))
    
            ''''20161114 update for print all bits (DTR) in STDF
            Call auto_eFuse_to_STDF_allBits("Config", gS_CFG_Direct_Access_Str(site))
        Next site
    
    ''''----------------------------------------------------------------------------------------------

    ''''gL_CFG_FBC has been check in Blank/MarginRead
    Call UpdateDLogColumns(gI_CFG_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=gL_CFG_FBC, lowVal:=0, hiVal:=0, Tname:="CFG_FBCount_" + UCase(gS_JobName) '2d-s=0
    Call UpdateDLogColumns__False

        ''''20171024 enable here
    ''''----------------------------------------------------------------------------------------
    ''''IEDA, user definition here for the register Key
    ''''Path:: HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\ (check by regEdit)
    If (True) Then
        Dim allCFG_First_bitStr As String
        Dim tmpCFG_First_bitStr As String
        Dim m_bitStrM As String
        Dim m_algorithm As String
        
        Dim m_name As String:: m_name = ""
        Dim tempIEDA_BKM As String
        Dim IEDA_BKM As String
        
        ''''20171024 add, 20171113 update
        Dim m_regName_CFG_Cond As String
        If (gB_findCFGCondTable_flag = True) Then
            m_regName_CFG_Cond = "SVM_CFuse_288Bits"
        ElseIf (gB_findCFGTable_flag = True) Then
            m_regName_CFG_Cond = "CFG_First_64Bits"
        Else
            Exit Function
        End If

        
        ''''initialize
        allCFG_First_bitStr = ""
        IEDA_BKM = ""

        For Each site In TheExec.sites.Existing
            ''''initialize
            tmpCFG_First_bitStr = ""
            tempIEDA_BKM = ""

            ''''1st: get the tmpXXXstr per site
            For i = 0 To UBound(CFGFuse.Category)
                m_bitStrM = CFGFuse.Category(i).Read.BitStrM(site)
                m_algorithm = LCase(CFGFuse.Category(i).algorithm)
                If (m_algorithm = "cond") Then
                    tmpCFG_First_bitStr = tmpCFG_First_bitStr + m_bitStrM
                End If
                If (m_name Like "*BKM*") Then
                    tempIEDA_BKM = tempIEDA_BKM + m_bitStrM
                End If
            Next i

            ''''2nd: integrate tmpXXXstr to allXXXstr for iEDA register key
            If (site = TheExec.sites.Existing.Count - 1) Then
                allCFG_First_bitStr = allCFG_First_bitStr + tmpCFG_First_bitStr
                IEDA_BKM = IEDA_BKM + tempIEDA_BKM
            Else
                allCFG_First_bitStr = allCFG_First_bitStr + tmpCFG_First_bitStr + ","
                IEDA_BKM = IEDA_BKM + tempIEDA_BKM + ","
            End If
        Next site

        allCFG_First_bitStr = auto_checkIEDAString(allCFG_First_bitStr)
        IEDA_BKM = auto_checkIEDAString(IEDA_BKM)

        TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName
        TheExec.Datalog.WriteComment " ConfigRead (all sites iEDA format)::"
        TheExec.Datalog.WriteComment " " + m_regName_CFG_Cond + " = " & allCFG_First_bitStr
        If Not (UCase(TheExec.DataManager.instanceName) Like "*EARLY*") Then
            TheExec.Datalog.WriteComment " " + "BKM" + " = " & IEDA_BKM
        End If
        TheExec.Datalog.WriteComment ""

        Call RegKeySave(m_regName_CFG_Cond, allCFG_First_bitStr)
        
        If Not (UCase(TheExec.DataManager.instanceName) Like "*EARLY*") Then
            Call RegKeySave("BKM", IEDA_BKM)
        End If
    End If
    ''''----------------------------------------------------------------------------------------
    
Exit Function

End If

'    ''''------------------------------------------------------------------------------------------------------------------
'    ''''<Important Notice>
'    ''''------------------------------------------------------------------------------------------------------------------
'    ''''gS_SingleStrArray() was extracted in the module auto_ConfigRead_by_OR_2Blocks() then used in auto_ConfigSingleDoubleBit()
'    ''''gS_SingleStrArray() is the result of the NormRead or MarginRead
'    ''''
'    ''''So it doesn't need to run the pattern and DSSC to get the SignalStrArray, and save test time
'    ''''------------------------------------------------------------------------------------------------------------------
'
'    Call UpdateDLogColumns(gI_CFG_catename_maxLen)
'
'    TheExec.Datalog.WriteComment ""
'
'    For Each Site In TheExec.Sites
'
'        ''''20160202 add to simulate for Retest Stage
'        ''''20161107 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            If (False) Then ''''True for Debug, 20161031
'                TheExec.Datalog.WriteComment "---Offline--- Read All Config eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'                Call auto_PrintAllBitbyDSSC(SingleBitArray(), EConfigReadCycle, EConfigTotalBitCount, EConfigReadBitWidth)
'            End If
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        '==============================================================
'        '=  Extract DigCap data for exhibit Configuration eFuse       =                                                                                                                  =
'        '==============================================================
'        'Before comparing with program bit, you have to OR 2 block bit by bit
'        Call auto_OR_2Blocks("CFG", gS_SingleStrArray, SingleBitArray, DoubleBitArray)
'
'        ''''20170220 update
'        If (gL_CFG_FBC(Site) > 0) Then
'            TmpStr = "The Fail Bit Count of Config eFuse at Site(" + CStr(Site) + ") is " + CStr(gL_CFG_FBC(Site))
'            TmpStr = TmpStr + " (Max FBC=0)"
'            TheExec.Datalog.WriteComment TmpStr
'        End If
'
'        If (TheExec.Sites(Site).SiteVariableValue(m_siteVar) <> 1 Or gB_CFG_decode_flag(Site) = False) Then
'            ''''ReTest Stage
'            ''''to get CFGFuse Result for the IEDA data,Decode
'            ''''<Important> User Need to check the content inside
'            Call auto_Decode_CfgBinary_Data(DoubleBitArray)
'
'            ''''' 20161012 ADD CRC
'            If (Trim(gS_CFG_CRC_Stage) <> "") Then
'                If (checkJob_less_Stage_Sequence(gS_CFG_CRC_Stage) = True) Then
'                    gS_CFG_CRC_HexStr(Site) = "0000"
'                Else
'                    gS_CFG_CRC_HexStr(Site) = auto_CFG_CRC2HexStr(DoubleBitArray, crcBinStr)
'                End If
'            End If
'
'            If (False) Then ''''True for Debug, 20161031
'                TheExec.Datalog.WriteComment ""
'                TheExec.Datalog.WriteComment "Read All Config eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'                Call auto_PrintAllBitbyDSSC(SingleBitArray(), EConfigReadCycle, EConfigTotalBitCount, EConfigReadBitWidth) ''must use EConfigReadBitWidth in Right2Left mode
'            End If
'        End If
'
'        '=====================================================================================
'        '=  Get EConfig data from DSSC for writing to HKEY                                   =
'        '=====================================================================================
'        gS_CFG_Direct_Access_Str(Site) = "" ''''is a String [(bitLast)......(bit0)]
'
'        For i = 0 To UBound(DoubleBitArray)
'            gS_CFG_Direct_Access_Str(Site) = CStr(DoubleBitArray(i)) + gS_CFG_Direct_Access_Str(Site)
'        Next i
'        ''TheExec.Datalog.WriteComment "gS_CFG_Direct_Access_Str=" + CStr(gS_CFG_Direct_Access_Str(Site))
'
'        ''''20161114 update for print all bits (DTR) in STDF
'        Call auto_eFuse_to_STDF_allBits("Config", gS_CFG_Direct_Access_Str(Site))
'        ''''----------------------------------------------------------------------------------------------
'        ''''20160630 Add
'        gS_CFG_SingleBit_Str(Site) = ""     ''''is a String [(bitLast)......(bit0)]
'        For i = 0 To UBound(SingleBitArray)
'            gS_CFG_SingleBit_Str(Site) = CStr(SingleBitArray(i)) + gS_CFG_SingleBit_Str(Site)
'        Next i
'        ''TheExec.Datalog.WriteComment "gS_CFG_SingleBit_Str=" + CStr(gS_CFG_SingleBit_Str(Site))
'        ''''----------------------------------------------------------------------------------------------
'    Next Site
'
'    ''''' Added by the request of Chris Vu
'    For Each Site In TheExec.Sites
'        ''' This prerequesite of "CFGFuse.Category(CFGIndex("PCIE_REFPLL_FCAL_VCO_DIGCTRL")).Read.Decimal > 1" needs to be modified by different project owner.
'        ''' The purpose is to make sure the pcie trim code has been fused. Otherwise this CRC print out is meaningless.
'        ''' 20170623 Per discussion with Jack, update as the below
'        If (checkJob_less_Stage_Sequence(gS_CFG_CRC_Stage) = True) Then
'            ''''Here it means the CRC codes are NOT blown into the Fuse.
'            ''''So we do NOT need to check the CRC result.
'            ''''doNothing
'        Else
'            ''''means that the CRC codes have been blown into the Fuse, we MUST display the below bit string for the tool calculation.
'            TheExec.Datalog.WriteComment "CFG bit string for CRC calculation on Site " & CStr(Site) & " (MSB to LSB)= " & CStr(CRC_Shift_Out_String(Site))   ''20170623
'            TheExec.Datalog.WriteComment Chr(13) & "Totally " & CStr(Len(CRC_Shift_Out_String(Site))) & " bits for CRC calculation"
'        End If
'    Next Site
'    TheExec.Datalog.WriteComment ""
'    TheExec.Flow.TestLimit resultVal:=gL_CFG_FBC, lowVal:=0, hiVal:=0, Tname:="FailBitCount"
'    Call UpdateDLogColumns__False
'
'    ''''20171024 enable here
'    ''''----------------------------------------------------------------------------------------
'    ''''IEDA, user definition here for the register Key
'    ''''Path:: HKEY_CURRENT_USER\Software\VB and VBA Program Settings\IEDA\ (check by regEdit)
'    If (True) Then
''        Dim allCFG_First_bitStr As String
''        ''Dim allCFG_IDS_bitStr As String
''        ''Dim allCFG_VddBin_bitStr As String
''        Dim tmpCFG_First_bitStr As String
''        ''Dim tmpCFG_IDS_bitStr As String
''        ''Dim tmpCFG_VddBin_bitStr As String
''        Dim m_bitStrM As String
''        Dim m_algorithm As String
''
''        ''''20171024 add, 20171113 update
''        Dim m_regName_CFG_Cond As String
''        If (gB_findCFGCondTable_flag = True) Then
''            m_regName_CFG_Cond = "SVM_CFuse_288Bits"
''        ElseIf (gB_findCFGTable_flag = True) Then
''            m_regName_CFG_Cond = "CFG_First_64Bits"
''        Else
''            Exit Function
''        End If
''
''
''        ''''initialize
''        allCFG_First_bitStr = ""
''        ''allCFG_IDS_bitStr = ""
''        ''allCFG_VddBin_bitStr = ""
''
''        For Each Site In TheExec.Sites.Existing
''            ''''initialize
''            tmpCFG_First_bitStr = ""
''            ''tmpCFG_IDS_bitStr = ""
''            ''tmpCFG_VddBin_bitStr = ""
''
''            ''''1st: get the tmpXXXstr per site
''            For i = 0 To UBound(CFGFuse.Category)
''                m_bitStrM = CFGFuse.Category(i).Read.BitstrM(Site)
''                m_algorithm = LCase(CFGFuse.Category(i).Algorithm)
''                ''If (m_algorithm = "firstbits") Then
''                If (m_algorithm = "cond") Then
''                    tmpCFG_First_bitStr = tmpCFG_First_bitStr + m_bitStrM
''
''''                ElseIf (m_algorithm = "ids") Then
''''                    tmpCFG_IDS_bitStr = tmpCFG_IDS_bitStr + m_bitstrM
''''
''''                ElseIf (m_algorithm = "vddbin") Then
''''                    tmpCFG_VddBin_bitStr = tmpCFG_VddBin_bitStr + m_bitstrM
''
''                End If
''            Next i
''
''            ''''2nd: integrate tmpXXXstr to allXXXstr for iEDA register key
''            If (Site = TheExec.Sites.Existing.Count - 1) Then
''                allCFG_First_bitStr = allCFG_First_bitStr + tmpCFG_First_bitStr
''                ''allCFG_IDS_bitStr = allCFG_IDS_bitStr + tmpCFG_IDS_bitStr
''                ''allCFG_VddBin_bitStr = allCFG_VddBin_bitStr + tmpCFG_VddBin_bitStr
''            Else
''                allCFG_First_bitStr = allCFG_First_bitStr + tmpCFG_First_bitStr + ","
''                ''allCFG_IDS_bitStr = allCFG_IDS_bitStr + tmpCFG_IDS_bitStr + ","
''                ''allCFG_VddBin_bitStr = allCFG_VddBin_bitStr + tmpCFG_VddBin_bitStr + ","
''            End If
''        Next Site
''
''        allCFG_First_bitStr = auto_checkIEDAString(allCFG_First_bitStr)
''        ''allCFG_IDS_bitStr = auto_checkIEDAString(allCFG_IDS_bitStr)
''        ''allCFG_VddBin_bitStr = auto_checkIEDAString(allCFG_VddBin_bitStr)
''
''        TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.InstanceName
''        TheExec.Datalog.WriteComment " ConfigRead (all sites iEDA format)::"
''        TheExec.Datalog.WriteComment " " + m_regName_CFG_Cond + " = " & allCFG_First_bitStr
''        ''TheExec.Datalog.WriteComment " CFG_IDS_bits   = " & allCFG_IDS_bitStr
''        ''TheExec.Datalog.WriteComment " CFG_VddBin_bits= " & allCFG_VddBin_bitStr
''        TheExec.Datalog.WriteComment ""
''
''        Call RegKeySave(m_regName_CFG_Cond, allCFG_First_bitStr)
''        ''Call RegKeySave("Hram_IDS_37bit", allCFG_IDS_bitStr)
''        ''Call RegKeySave("Hram_BinData_46bit", allCFG_VddBin_bitStr)
''
''        ''''to Verify
''        ''TmpStr = RegKeyRead("SVM_CFuse_288Bits")
''        ''TheExec.Datalog.WriteComment m_regName_CFG_Cond + " = " + TmpStr & vbCrLf
'    End If
    ''''----------------------------------------------------------------------------------------
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_ChkAllConfigEfuseData(Optional condstr As String = "ALL") As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ChkAllConfigEfuseData"

    Dim i As Long, j As Long, k As Long
    Dim Config64bitResult As New SiteLong
    Dim CFGCondbitResult As New SiteLong ''''20170630
    Dim site As Variant
    Dim m_pkgname As String
    Dim m_tsname As String
    Dim m_tsName0 As String
    Dim m_stage As String
    Dim m_catename As String
    Dim m_catenameVbin As String
    Dim m_bitStrM As String
    Dim m_bitwidth As Long
    Dim m_algorithm As String
    Dim m_resolution As Double
    Dim m_decimal As Variant ''20160506 update, was Long
    Dim m_decimal_ids As Double
    Dim m_bitsum As Long
    Dim m_value As Variant
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim TmpVal As Variant
    Dim vbinIdx As Long
    Dim vbinflag As Long
    Dim tmpStr As String
    Dim mem_Config_code As String
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim MaxLevelIndex As Long
    Dim m_testValue As Variant
    Dim m_Pmode As Long
    Dim m_unitType As UnitType
    Dim m_scale As tlScaleType
    Dim m_crchexStr As String
    
    ''''20170630 add
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_HexStr As String
    Dim m_bitStrM_Tab As String
    Dim m_hexStr_Tab As String
    Dim m_bitStr32M_Read As String

    ''''201811XX
    Dim m_customUnit As String
    Dim mSL_bitSum As New SiteLong
    Dim mSV_value As New SiteVariant
    Dim mSV_decimal As New SiteVariant
    Dim mSV_bitStrM As New SiteVariant
    Dim mSV_hexStr As New SiteVariant
    Dim mSL_valueSum As New SiteLong
    Dim m_findTMPS_flag As Boolean
    
'    Dim mSL_valueSum As New SiteLong
'    Dim m_findTMPS_flag As Boolean

    'Get Hi/Low limits from Check List table (this subroutine is only suitable for default vdd-binning values)
    ''''By using CFGFuse type structure here
    Call UpdateDLogColumns(gI_CFG_catename_maxLen)

    ''''20170630 update
    If (Trim(condstr) = "") Then
        condstr = "all"
    End If
    condstr = LCase(condstr)

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If
    ''''gS_cfgFlagname is gotten from OI
    ''''gS_CFG_Cond_Read_bitStrM(site) has been decided in auto_eFuse_setReadData_XXXX()
    ''''gS_CFGCondTable_bitsStr, it has been decised in auto_CFGConstant_Initialize()
''auto_StartWatchTimer

    If (condstr = "cp1_early") Then

        ''''------------------------------------------------------------------
        ''''Check the result and its test limit
        For i = 0 To UBound(CFGFuse.Category)
            With CFGFuse.Category(i)
                m_stage = LCase(.Stage)
                m_catename = .Name
                m_algorithm = LCase(.algorithm)
                m_value = .Read.Value(site)
                m_lolmt = .LoLMT
                m_hilmt = .HiLMT
                m_bitwidth = .BitWidth
                m_defval = .DefaultValue
                m_resolution = .Resoultion
                m_defreal = LCase(.Default_Real)
                ''''-----------------------------------
                mSL_bitSum = .Read.BitSummation
                mSV_decimal = .Read.Decimal
                mSV_bitStrM = .Read.BitStrM
                mSV_hexStr = .Read.HexStr
                mSV_value = .Read.Value
                ''''-----------------------------------
            End With
            
            ''''Here it presents the CFG_Cond_[Flag Selected] syntax check
            If (i = 0) Then
                m_tsname = "CFG_Cond_" + gS_cfgFlagname '''' + "_Read_" + gS_CFG_Cond_Read_pkgname(Site)) then need siteloop for the tsname per site
                'mSV_value = gL_CFG_Cond_compResult
                TheExec.Flow.TestLimit resultVal:=gL_CFG_Cond_compResult, lowVal:=0, hiVal:=0, Tname:=m_tsname
            End If
            
            If (m_stage = gS_JobName Or m_algorithm = "cond") Then
                m_tsname = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
                m_unitType = unitNone
                m_scale = scaleNoScaling ''= scaleNone ''''default
                
                ''''<NOTICE> 20160108
                Call auto_eFuse_chkLoLimit("CFG", i, m_stage, m_lolmt)
                Call auto_eFuse_chkHiLimit("CFG", i, m_stage, m_hilmt)
    
                If (m_bitwidth >= 32) Then
                    ''''keep Limit as Hex String
                    ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
                    m_lolmt = auto_Value2HexStr(m_lolmt, m_bitwidth)
                    m_hilmt = auto_Value2HexStr(m_hilmt, m_bitwidth)
                    
                    ''''----------------------------------------------
                    ''''auto_TestStringLimit compare with lolmt, hilmt
                    ''''return  0 means Fail
                    ''''return  1 means Pass
                    ''''----------------------------------------------
                    For Each site In TheExec.sites
                        m_HexStr = mSV_hexStr(site)
                        mSV_value(site) = auto_TestStringLimit(m_HexStr, CStr(m_lolmt), CStr(m_hilmt)) - 1
                    Next site
                    ''''mSV_value=0: Pass, = -1 Fail
                    m_lolmt = 0
                    m_hilmt = 0
                Else
                    ''''translate to double value
                    If (auto_isHexString(CStr(m_lolmt)) = True) Then m_lolmt = auto_HexStr2Value(m_lolmt)
                    If (auto_isHexString(CStr(m_hilmt)) = True) Then m_hilmt = auto_HexStr2Value(m_hilmt)
                End If

                If (auto_eFuse_check_Job_cmpare_Stage(m_stage) < 1 Or gB_eFuse_Disable_ChkLMT_Flag = True) Then ''''Job<=Stage
                    TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsname, Unit:=m_unitType, scaletype:=m_scale, customForceunit:=m_customUnit
                Else
                    TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=0, hiVal:=0, Tname:=m_tsname, Unit:=m_unitType, scaletype:=m_scale, customForceunit:=m_customUnit
                End If

            End If ''''end of If (m_stage = gS_JobName)
        Next i
    Else
        ''''other case
        ''''------------------------------------------------------------------
        ''''Check the result and its test limit
        m_findTMPS_flag = False ''''<MUST> initial
        For i = 0 To UBound(CFGFuse.Category)
            With CFGFuse.Category(i)
                m_stage = LCase(.Stage)
                m_catename = .Name
                m_algorithm = LCase(.algorithm)
                m_value = .Read.Value(site)
                m_lolmt = .LoLMT
                m_hilmt = .HiLMT
                m_bitwidth = .BitWidth
                m_defval = .DefaultValue
                m_resolution = .Resoultion
                m_defreal = LCase(.Default_Real)
                ''''-----------------------------------
                mSL_bitSum = .Read.BitSummation
                mSV_decimal = .Read.Decimal
                mSV_bitStrM = .Read.BitStrM
                mSV_hexStr = .Read.HexStr
                mSV_value = .Read.Value
                ''''-----------------------------------
            End With
            
            m_tsname = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
            m_unitType = unitNone
            m_customUnit = ""
            m_scale = scaleNoScaling ''=scaleNone ''''default

            ''''Here it presents the CFG_Cond_[Flag Selected] syntax check
            If (i = 0) Then
                m_tsName0 = "CFG_Cond_" + gS_cfgFlagname '''' + "_Read_" + gS_CFG_Cond_Read_pkgname(Site)) then need siteloop for the tsname per site
                'mSV_value = gL_CFG_Cond_compResult
                TheExec.Flow.TestLimit resultVal:=gL_CFG_Cond_compResult, lowVal:=0, hiVal:=0, Tname:=m_tsName0
            End If
            
            ''If (m_stage = gS_JobName) Then ''by Job
            If (True) Then
                
                ''''<NOTICE> 20160108
                Call auto_eFuse_chkLoLimit("CFG", i, m_stage, m_lolmt)
                Call auto_eFuse_chkHiLimit("CFG", i, m_stage, m_hilmt)
                
                ''''<TRY>dummy Test
                If (m_algorithm = "ids") Then
                    m_unitType = unitCustom
                    m_customUnit = "mA"
                    If (CDbl(m_lolmt) = 0# And CDbl(m_hilmt) = 0#) Then ''''Need to check
                        m_lolmt = 1# * m_resolution
                        m_hilmt = ((2 ^ m_bitwidth) - 1) * m_resolution
                    Else
                        If (CDbl(m_lolmt) = 0#) Then m_lolmt = 1# * m_resolution  '0 means nothing, can not be acceptable
                    End If
                
                ElseIf (m_algorithm = "base" And m_defreal Like "safe*voltage") Then
                    m_unitType = unitCustom
                    m_customUnit = "mV"

                ElseIf (m_algorithm = "vddbin") Then
                    m_unitType = unitCustom
                    m_customUnit = "mV"
                    If (m_resolution = 0#) Then m_customUnit = "" ''''decimal case
                    
                    ''If (m_defreal Like "safe*voltage") Then
                    ''''Has been done in auto_eFuse_setReadData_XXX()
                    ''''mSV_value = mSV_decimal.Multiply(m_resolution).Add(gD_BaseVoltage)
                    
                    If (m_defreal = "bincut") Then
                        'm_isBinCut_flag = True
                        ''''<Notice> User maintain
                        ''''-------------------------------------
                        ''''m_catenameVBin = "MS001" ''''<NOTICE> M8 uses MS001 on both power VDD_SOC and VDD_SOC_AON
                        m_catenameVbin = m_catename
                        For Each site In TheExec.sites
                            ''''' Every testing stage needs to check this formula with bincut fused value (2017/5/17, Jack)
                            vbinflag = auto_CheckVddBinInRangeNew(m_catenameVbin, m_resolution)
                            
                            ''''20160329 Add for the offline simulation
                            If (vbinflag = 0 And TheExec.TesterMode = testModeOffline) Then
                                vbinflag = 1
                            End If
                            
                            ''''was m_vddbinEnum, its equal to m_Pmode
                            m_Pmode = VddBinStr2Enum(m_catenameVbin) ''''20160329 add
                            ''tmpVal = auto_VfuseStr_to_Vdd_New(m_bitStrM, m_bitwidth, m_resolution)
                            
                            MaxLevelIndex = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step
                            m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(MaxLevelIndex) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(MaxLevelIndex)
                            m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
                            ''''judge the result
                            If (vbinflag <> 1) Then
                                mSV_value = -999
                                tmpStr = m_catenameVbin + "(Site " + CStr(site) + ") = " + CStr(mSV_value) + " is not in range"
                                TheExec.Datalog.WriteComment tmpStr
                            End If
                            
                            If (auto_eFuse_check_Job_cmpare_Stage(m_stage) <> -1 Or gB_eFuse_Disable_ChkLMT_Flag = True) Then ''''Job<=Stage
                                TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsname, scaletype:=m_scale, Unit:=m_unitType, customUnit:=m_customUnit
                            Else
                                TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=0, hiVal:=0, Tname:=m_tsname, scaletype:=m_scale, Unit:=m_unitType, customUnit:=m_customUnit
                            End If
                        Next site
                    End If
                
                ElseIf (m_algorithm = "crc") Then
                    m_lolmt = 0
                    m_hilmt = 0
                    m_tsName0 = m_tsname
                    ''''20170309, per Jack's recommendation to use Read_Data to do the CRC calculation result (=gS_CFG_CRC_HexStr)
                    ''''          and compare with the CRC readCode.
                    If (gB_eFuse_Disable_ChkLMT_Flag = True) Then
                        ''''set Pass
                        For Each site In TheExec.sites
                            m_tsname = m_tsName0 + "_" + UCase(gS_CFG_Read_calcCRC_hexStr(site))
                            TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:=m_tsname
                        Next site
                    Else
                        For Each site In TheExec.sites
                            m_crchexStr = UCase(gS_CFG_Read_calcCRC_hexStr(site))
                            m_tsname = m_tsName0 + "_" + m_crchexStr
                            ''''<NOTICE>
                            ''''mSV_hexStr  is the CRC HexStr of Read eFuse Category
                            ''''m_crchexStr is the CRC HexStr by the calculation of read bits.
                            If (UCase(mSV_hexStr) = m_crchexStr) Then
                                ''''Pass
                                TheExec.Flow.TestLimit resultVal:=0, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsname
                            Else
                                ''''Fail
                                TheExec.Flow.TestLimit resultVal:=1, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsname
                            End If
                        Next site
                    End If
                End If

                If (m_bitwidth >= 32) Then
                    ''''keep Limit as Hex String
                    ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
                    m_lolmt = auto_Value2HexStr(m_lolmt, m_bitwidth)
                    m_hilmt = auto_Value2HexStr(m_hilmt, m_bitwidth)
                    
                    ''''----------------------------------------------
                    ''''auto_TestStringLimit compare with lolmt, hilmt
                    ''''return  0 means Fail
                    ''''return  1 means Pass
                    ''''----------------------------------------------
                    For Each site In TheExec.sites
                        m_HexStr = mSV_hexStr(site)
                        mSV_value(site) = auto_TestStringLimit(m_HexStr, CStr(m_lolmt), CStr(m_hilmt)) - 1
                    Next site
                    ''''mSV_value=0: Pass, = -1 Fail
                    m_lolmt = 0
                    m_hilmt = 0
                Else
                    ''''translate to double value
                    If (auto_isHexString(CStr(m_lolmt)) = True) Then m_lolmt = auto_HexStr2Value(m_lolmt)
                    If (auto_isHexString(CStr(m_hilmt)) = True) Then m_hilmt = auto_HexStr2Value(m_hilmt)
                End If
    
                ''''201812XX update
                If (m_defreal <> "bincut" And m_algorithm <> "crc") Then
                    If (auto_eFuse_check_Job_cmpare_Stage(m_stage) >= 0 Or gB_eFuse_Disable_ChkLMT_Flag = True) Then
                        ''''Job >= Stage or disable limit check
                        TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsname, scaletype:=m_scale, Unit:=m_unitType, customUnit:=m_customUnit
                    Else
                        ''''Job < Stage
                        TheExec.Flow.TestLimit resultVal:=mSV_value, lowVal:=0, hiVal:=0, Tname:=m_tsname, scaletype:=m_scale, Unit:=m_unitType, customUnit:=m_customUnit
                    End If
                End If

            End If ''''end of If()
        Next i
    End If

    Call UpdateDLogColumns__False
     
    ''''201811XX reset
    If (gS_JobName = "cp1_early") Then gS_JobName = "cp1" ''''<MUST>

''auto_StopWatchTimer ("0099")

Exit Function

End If




'    If (gB_findCFGCondTable_flag = True) Then
'        ''''gS_cfgFlagname is gotten from OI
'        ''''Get String: gS_CFGCondTable_bitsStr
'        gS_CFGCondTable_bitsStr = ""
'
'        ''''20170915 update
'        ''''all zero case, default
'        For j = 0 To CFGFuse.Category(gI_CFG_firstbits_index).Bitwidth - 1
'            gS_CFGCondTable_bitsStr = "0" + gS_CFGCondTable_bitsStr
'        Next j
'        For i = 0 To UBound(CFGTable.Category)
'            m_pkgname = UCase(CFGTable.Category(i).pkgName)
'            If (UCase(gS_cfgFlagname) = "ALL_0") Then ''''20160614 update
'''                ''''all zero case
'''                For j = 0 To CFGFuse.Category(gI_CFG_firstbits_index).BitWidth - 1
'''                    gS_CFGCondTable_bitsStr = "0" + gS_CFGCondTable_bitsStr
'''                Next j
'                Exit For
'            ElseIf (m_pkgname = UCase(gS_cfgFlagname)) Then
'                gS_CFGCondTable_bitsStr = CFGTable.Category(i).BitstrM
'                Exit For
'            End If
'        Next i
'    ElseIf (gB_findCFGTable_flag = True) Then
'        ''''gS_cfgFlagname is gotten from OI
'        ''''Get String: gS_cfgTable_First64bitsStr
'        gS_cfgTable_First64bitsStr = ""
'        For i = 0 To UBound(CFGTable.Category)
'            m_pkgname = UCase(CFGTable.Category(i).pkgName)
'            If (UCase(gS_cfgFlagname) = "ALL_0") Then ''''20160614 update
'                ''''all zero case
'                For j = 0 To CFGFuse.Category(gI_CFG_firstbits_index).Bitwidth - 1
'                    gS_cfgTable_First64bitsStr = "0" + gS_cfgTable_First64bitsStr
'                Next j
'                Exit For
'            ElseIf (m_pkgname = UCase(gS_cfgFlagname)) Then
'                gS_cfgTable_First64bitsStr = CFGTable.Category(i).BitstrM
'                Exit For
'            End If
'        Next i
'    End If
'
'    If (condstr = "cp1_early") Then
'        For Each Site In TheExec.Sites
'            ''''gS_CFGCondTable_bitsStr
'            ''''gI_CFG_firstbits_index
'            m_bitStrM = CFGFuse.Category(gI_CFG_firstbits_index).Read.BitstrM(Site)
'
'            ''''the purpose is to let FT3 parts retestable with FT1 or FT2 T/P (per C651 request)
'            If (gS_JobName Like "cp*" Or gS_JobName Like "ft*" Or gS_JobName = "wlft") Then ''''CP* or WLFT or FT*
'                If (m_bitStrM <> gS_CFGCondTable_bitsStr) Then
'                    mem_Config_code = CFGFuse.Category(gI_CFG_firstbits_index).Read.ValStr(Site)
'                    TheExec.Datalog.WriteComment "CFG Condition from DSSC is " & mem_Config_code & " (but OI select " & gS_cfgFlagname & " !!)"
'                    CFGCondbitResult(Site) = 0 'Fail
'                Else
'                    CFGCondbitResult(Site) = 1 'Pass
'                End If
'            Else    'Char/HTOL/.... don't need to care about CFG Condition Bits
'                CFGCondbitResult(Site) = 1 'Pass
'            End If
'
'            ''''<Notice>
'            ''''Only Check Cond/SCAN
'            ''''------------------------------------------------------------------
'            ''''Check the result and its test limit
'            For i = 0 To UBound(CFGFuse.Category)
'                m_stage = LCase(CFGFuse.Category(i).Stage)
'                m_catename = CFGFuse.Category(i).Name
'                m_algorithm = LCase(CFGFuse.Category(i).Algorithm)
'                m_decimal = CFGFuse.Category(i).Read.Decimal(Site)
'                m_bitStrM = CFGFuse.Category(i).Read.BitstrM(Site)
'                m_value = CFGFuse.Category(i).Read.Value(Site)
'                m_lolmt = CFGFuse.Category(i).LoLMT
'                m_hilmt = CFGFuse.Category(i).HiLMT
'                m_bitwidth = CFGFuse.Category(i).Bitwidth
'                m_bitSum = CFGFuse.Category(i).Read.BitSummation(Site)
'                m_defval = CFGFuse.Category(i).DefaultValue
'                m_resolution = CFGFuse.Category(i).Resoultion
'                m_defreal = LCase(CFGFuse.Category(i).Default_Real)
'                m_hexStr = CFGFuse.Category(i).Read.HexStr(Site) ''''20170811 add
'                m_tsName = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
'                m_unitType = unitNone
'                m_scale = scaleNone ''''default
'
'                If (m_algorithm = "cond") Then
'                    ''''Only Compare here
'                    k = CFGTabIndex(gS_cfgFlagname) ''''<Important> 20170717 update
'                    ''''per 32bits to display and judgement
'                    For j = 0 To UBound(CFGTable.Category(k).Cate32bit)
'                        m_catename = CFGTable.Category(k).Cate32bit(j).Name
'                        m_MSBBit = CFGTable.Category(k).Cate32bit(j).MSBbit
'                        m_LSBbit = CFGTable.Category(k).Cate32bit(j).LSBbit
'                        m_bitStrM_Tab = CFGTable.Category(k).Cate32bit(j).BitstrM
'                        m_hexStr_Tab = CFGTable.Category(k).Cate32bit(j).HexStr
'
'                        m_tsName = "CFG_Cond["
'                        m_tsName = m_tsName + Format(m_MSBBit, "000") + ":" + Format(m_LSBbit, "000") + "]"
'                        m_tsName = m_tsName + "_" + gS_cfgFlagname + "_" + m_hexStr_Tab
'
'                        ''''--------------------------------------
'                        ''''m_bitStrM_Tab as the below
'                        ''''--------------------------------------
'                        ''''if j=0, CFG_CONDITION_7_0   [031:000]
'                        ''''if j=1, CFG_CONDITION_15_8  [063:032]
'                        ''''if j=2, CFG_CONDITION_23_16 [095:064]
'                        ''''if j=3, CFG_CONDITION_31_24 [127:096]
'                        ''''if j=4, CFG_CONDITION_62_32 [159:128]
'                        ''''if j=5, CFG_CONDITION_70_63 [191:160]
'                        ''''if j=6, CFG_CONDITION_71_71 [223:192]
'                        ''''if j=7, CFG_CONDITION_72_72 [255:224]
'                        ''''if j=8, CFG_CONDITION_80_73 [287:256]
'                        ''''--------------------------------------
'                        ''''Here m_bitStrM is Read decode [287:0]
'                        ''''--------------------------------------
'                        ''m_bitStr32M_Read = Mid(m_bitStrM, (j * 32) + 1, 32)
'                        m_bitStr32M_Read = Mid(m_bitStrM, ((UBound(CFGTable.Category(k).Cate32bit) - j) * 32) + 1, 32)
'                        If (m_bitStr32M_Read = m_bitStrM_Tab) Then
'                            TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                        Else
'                            ''''Fail
'                            TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                        End If
'                    Next j
'
'                ElseIf (m_algorithm = "scan") Then
'                    ''''only cond/scan in Early stage
'                    m_testValue = m_decimal
'
'                Else
'                    ''''other cases skip
'                    ''''doNothing
'                End If
'
'                ''''20160108 New
'                If (m_algorithm = "scan") Then
'                    ''''<NOTICE> 20160108
'                    Call auto_eFuse_chkLoLimit("CFG", i, m_stage, m_lolmt)
'                    Call auto_eFuse_chkHiLimit("CFG", i, m_stage, m_hilmt)
'
'                    ''''20170811 update
'                    If (m_bitwidth >= 32) Then
'                        ''m_tsName = m_tsName + "_" + m_hexStr
'                        ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
'                        m_lolmt = auto_Value2HexStr(m_lolmt, m_bitwidth)
'                        m_hilmt = auto_Value2HexStr(m_hilmt, m_bitwidth)
'
'                        ''''------------------------------------------
'                        ''''compare with lolmt, hilmt
'                        ''''return -1 means less
'                        ''''return  0 means equal
'                        ''''return  1 means large
'                        ''''------------------------------------------
'                        m_testValue = auto_TestStringLimit(m_hexStr, CStr(m_lolmt), CStr(m_hilmt))
'                        m_lolmt = 1
'                        m_hilmt = 1
'                    Else
'                        ''''20160927 update the new logical methodology for the unexpected binary decode.
'                        If (auto_isHexString(CStr(m_lolmt)) = True) Then
'                            ''''translate to double value
'                            m_lolmt = auto_HexStr2Value(m_lolmt)
'                        Else
'                            ''''doNothing, m_lolmt = m_lolmt
'                        End If
'
'                        If (auto_isHexString(CStr(m_hilmt)) = True) Then
'                            ''''translate to double value
'                            m_hilmt = auto_HexStr2Value(m_hilmt)
'                        Else
'                            ''''doNothing, m_hilmt = m_hilmt
'                        End If
'                    End If
'                    TheExec.Flow.TestLimit resultVal:=m_testValue, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsName, unit:=m_unitType, ScaleType:=m_scale
'                End If
'            Next i
'            TheExec.Datalog.WriteComment ""
'        Next Site
'
'    ElseIf (condstr = "all") Then
'
'        For Each Site In TheExec.Sites
'            m_bitStrM = CFGFuse.Category(gI_CFG_firstbits_index).Read.BitstrM(Site)
'
'            If (gB_findCFGCondTable_flag = True) Then
'                ''''gS_CFGCondTable_bitsStr
'                ''''gI_CFG_firstbits_index
'                m_bitStrM = CFGFuse.Category(gI_CFG_firstbits_index).Read.BitstrM(Site)
'
'                ''''the purpose is to let FT3 parts retestable with FT1 or FT2 T/P (per C651 request)
'                If (gS_JobName Like "cp*" Or gS_JobName Like "ft*" Or gS_JobName = "wlft") Then ''''CP* or WLFT or FT*
'                    If (m_bitStrM <> gS_CFGCondTable_bitsStr) Then
'                        mem_Config_code = CFGFuse.Category(gI_CFG_firstbits_index).Read.ValStr(Site)
'                        TheExec.Datalog.WriteComment "CFG Condition from DSSC is " & mem_Config_code & " (but OI select " & gS_cfgFlagname & " !!)"
'                        CFGCondbitResult(Site) = 0 'Fail
'                    Else
'                        CFGCondbitResult(Site) = 1 'Pass
'                    End If
'                Else    'Char/HTOL/.... don't need to care about CFG Condition Bits
'                    CFGCondbitResult(Site) = 1 'Pass
'                End If
'            ElseIf (gB_findCFGTable_flag = True) Then
'                ''''gS_cfgTable_First64bitsStr
'                ''''gI_CFG_firstbits_index
'                ''''the purpose is to let FT3 parts retestable with FT1 or FT2 T/P (per C651 request)
'                If (gS_JobName Like "cp*" Or gS_JobName Like "ft*" Or gS_JobName = "wlft") Then ''''CP* or WLFT or FT*
'                    If (m_bitStrM <> gS_cfgTable_First64bitsStr) Then
'                        mem_Config_code = CFGFuse.Category(gI_CFG_firstbits_index).Read.ValStr(Site)
'                        TheExec.Datalog.WriteComment "First 64bit of Config from DSSC is " & mem_Config_code & " (but OI select " & gS_cfgFlagname & " !!)"
'                        Config64bitResult(Site) = 0 'Fail
'                    Else
'                        Config64bitResult(Site) = 1 'Pass
'                    End If
'                Else    'Char/HTOL/.... don't need to care about first 64 bits of Config eFuse
'                    Config64bitResult(Site) = 1 'Pass
'                End If
'            End If
'
'            ''''<Notice> User maintain this function auto_CheckVddBinInRange()
'            ''''Has been replaced by auto_CheckVddBinInRangeNew()
'            ''''------------------------------------------------------------------
'
'            ''''Check the result and its test limit
'            For i = 0 To UBound(CFGFuse.Category)
'                m_stage = LCase(CFGFuse.Category(i).Stage)
'                m_catename = CFGFuse.Category(i).Name
'                m_algorithm = LCase(CFGFuse.Category(i).Algorithm)
'                m_decimal = CFGFuse.Category(i).Read.Decimal(Site)
'                m_bitStrM = CFGFuse.Category(i).Read.BitstrM(Site)
'                m_value = CFGFuse.Category(i).Read.Value(Site)
'                m_lolmt = CFGFuse.Category(i).LoLMT
'                m_hilmt = CFGFuse.Category(i).HiLMT
'                m_bitwidth = CFGFuse.Category(i).Bitwidth
'                m_bitSum = CFGFuse.Category(i).Read.BitSummation(Site)
'                m_defval = CFGFuse.Category(i).DefaultValue
'                m_resolution = CFGFuse.Category(i).Resoultion
'                m_defreal = LCase(CFGFuse.Category(i).Default_Real)
'                m_hexStr = CFGFuse.Category(i).Read.HexStr(Site) ''''20170811 add
'                m_tsName = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
'                m_unitType = unitNone
'                m_scale = scaleNone ''''default
'
'                If (m_algorithm = "cond") Then
'                    k = CFGTabIndex(gS_cfgFlagname) ''''<Important> 20170717 update
'                    ''''per 32bits to display and judgement
'                    For j = 0 To UBound(CFGTable.Category(k).Cate32bit)
'                        m_catename = CFGTable.Category(k).Cate32bit(j).Name
'                        m_MSBBit = CFGTable.Category(k).Cate32bit(j).MSBbit
'                        m_LSBbit = CFGTable.Category(k).Cate32bit(j).LSBbit
'                        m_bitStrM_Tab = CFGTable.Category(k).Cate32bit(j).BitstrM
'                        m_hexStr_Tab = CFGTable.Category(k).Cate32bit(j).HexStr
'
'                        m_tsName = "CFG_Cond["
'                        m_tsName = m_tsName + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "]"
'                        m_tsName = m_tsName + "_" + gS_cfgFlagname + "_" + m_hexStr_Tab
'
'                        ''''--------------------------------------
'                        ''''m_bitStrM_Tab as the below
'                        ''''--------------------------------------
'                        ''''if j=0, CFG_CONDITION_7_0   [031:000]
'                        ''''if j=1, CFG_CONDITION_15_8  [063:032]
'                        ''''if j=2, CFG_CONDITION_23_16 [095:064]
'                        ''''if j=3, CFG_CONDITION_31_24 [127:096]
'                        ''''if j=4, CFG_CONDITION_62_32 [159:128]
'                        ''''if j=5, CFG_CONDITION_70_63 [191:160]
'                        ''''if j=6, CFG_CONDITION_71_71 [223:192]
'                        ''''if j=7, CFG_CONDITION_72_72 [255:224]
'                        ''''if j=8, CFG_CONDITION_80_73 [287:256]
'                        ''''--------------------------------------
'                        ''''Here m_bitStrM is Read decode [287:0]
'                        ''''--------------------------------------
'                        ''m_bitStr32M_Read = Mid(m_bitStrM, (j * 32) + 1, 32)
'                        m_bitStr32M_Read = Mid(m_bitStrM, ((UBound(CFGTable.Category(k).Cate32bit) - j) * 32) + 1, 32)
'                        If (m_bitStr32M_Read = m_bitStrM_Tab) Then
'                            TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                        Else
'                            ''''Fail
'                            TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                        End If
'                    Next j
'                ElseIf (m_algorithm = "firstbits") Then
'                    ''''Only Compare here
'                    TheExec.Flow.TestLimit resultVal:=Config64bitResult, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                ElseIf (m_algorithm = "ids") Then
'                    ''''TMA's experience:
'                    ''''User needs to get the hi/lo limit in IDS Measurement Function
'
'                    ''''<Notice> User maintain
'                    ''''-------------------------------------
'                    ''''<TMA> Call Get_IDS_UseLimit() FIRST in the IDS current item
'                    ''''<Prefer> Or put a fixed number in the efuse bitdef table
'                    ''''Then the following limits can be ignored.
'
'                    ''''because 'm_resolution' its unit is mA (New)
'                    ''''m_resolution = m_resolution * 0.001 ''''in unit A.
'
'                    ''''------------------------------------------------------------------
'                    ''''<Important> 20160617 update, per Jack's commet
'                    ''''------------------------------------------------------------------
'                    ''''IDS limit should get from IDS_testInstance in the specific m_stage
'                    ''''But if the job <> m_stage, then will use the table as the limite (no bincut).
'                    If (gS_JobName <> m_stage) Then
'                        m_lolmt = CFGFuse.Category(i).LoLMT_R
'                        m_hilmt = CFGFuse.Category(i).HiLMT_R
'                    End If
'                    ''''------------------------------------------------------------------
'                    ''''Per Jack's comment::
'                    ''''IDS limit should get from bincut table if bincut is there.
'                    ''''------------------------------------------------------------------
'''''                    If (m_defreal = "bincut") Then ''''need to check it later
'''''                        If gS_JobName Like "*cp*" Then m_hilmt = BinCut(P_mode, PassBinCutNum).IDS_CP_LIMIT(VBIN_RESULT(P_mode).Step)
'''''                        If gS_JobName Like "*ft1*" Then m_hilmt = BinCut(P_mode, PassBinCutNum).IDS_FT_LIMIT(VBIN_RESULT(P_mode).Step)
'''''                        If gS_JobName Like "*ft2*" Then m_hilmt = BinCut(P_mode, PassBinCutNum).IDS_FT2_LIMIT(VBIN_RESULT(P_mode).Step)
'''''                    End If
'                    ''''------------------------------------------------------------------
'
'                    ''''hi/lo limit unit is also 'mA' (New)
'                    m_unitType = unitAmp
'                    m_scale = scaleMilli
'                    If (m_lolmt = 0# And m_hilmt = 0#) Then ''''Need to check
'                        m_lolmt = 1# * m_resolution   '0 means nothing, can not be acceptable
'                        m_hilmt = ((2 ^ m_bitwidth) - 1) * m_resolution
'                    End If
'                    If (m_lolmt = 0#) Then m_lolmt = 1# * m_resolution '0 means nothing, can not be acceptable
'                    m_decimal_ids = CDbl(m_decimal) * m_resolution ''unit: mA
'                    m_testValue = m_decimal_ids ''''unit is mA
'
'                    ''''------------------------------------------------
'                    ''''20160710 using unit:A as the limit test
'                    m_lolmt = m_lolmt * 0.001
'                    m_hilmt = m_hilmt * 0.001
'                    m_testValue = m_testValue * 0.001
'                    ''''------------------------------------------------
'
'                ElseIf (m_algorithm = "base") Then ''''M8 Case
'                    ''''m_value has been translated in SingleDoubleBit
'                    If (m_defreal = "decimal") Then ''''20160624 update
'                        m_testValue = m_decimal
'                    Else
'                        m_testValue = (m_decimal + 1) * gD_BaseStepVoltage
'                        m_unitType = unitVolt
'                        m_scale = scaleMilli
'                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
'                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
'                        m_testValue = m_value * 0.001 ''''to unit:V
'                    End If
'
'                ElseIf (m_algorithm = "vddbin") Then
'                    ''''<Notice> User maintain
'                    ''''-------------------------------------
'                    ''''m_catenameVBin = "MS001" ''''<NOTICE> M8 uses MS001 on both power VDD_SOC and VDD_SOC_AON
'                    m_catenameVbin = m_catename
'                    If (m_defreal = "decimal") Then
'                        ''''No unit
'                        m_testValue = m_decimal
'                    ElseIf (m_defreal Like "safe*voltage" Or m_defreal = "default") Then
'                        m_unitType = unitVolt
'                        m_scale = scaleMilli
'                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
'                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
'                        m_testValue = m_value * 0.001 ''''to unit:V
'                    ElseIf (m_defreal = "bincut") Then     'use realistic vdd-binning values
'                        ''''' Every testing stage needs to check this formula with bincut fused value (2017/5/17, Jack)
'                        vbinflag = auto_CheckVddBinInRangeNew(m_catenameVbin, m_resolution)
'
'                        ''''20160329 Add for the offline simulation
'                        If (vbinflag = 0 And TheExec.TesterMode = testModeOffline) Then
'                            vbinflag = 1
'                        End If
'
'                        ''''was m_vddbinEnum, its equal to m_Pmode
'                        m_Pmode = VddBinStr2Enum(m_catenameVbin) ''''20160329 add
'                        tmpVal = auto_VfuseStr_to_Vdd_New(m_bitStrM, m_bitwidth, m_resolution)
'                        MaxLevelIndex = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step
'                        m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(MaxLevelIndex) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(MaxLevelIndex)
'                        m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
'                        ''''judge the result
'                        If (vbinflag = 1) Then
'                            m_value = tmpVal
'                        Else
'                            m_value = -999
'                            TmpStr = m_catename + "(Site " + CStr(Site) + ") = " + CStr(tmpVal) + " is not in range"
'                            TheExec.Datalog.WriteComment TmpStr
'                        End If
'                        m_unitType = unitVolt
'                        m_scale = scaleMilli
'                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
'                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
'                        m_testValue = m_value * 0.001 ''''to unit:V
'                    End If
'
'                ''''20161019 update
'                ElseIf (m_algorithm = "crc") Then
'                    m_value = auto_BinStr2HexStr(m_bitStrM, 4) ''''<MUST> it's the Read Code of CRC category
'                    ''''20170309, per Jack's recommendation to use Read_Data to do the CRC calculation result (=gS_CFG_CRC_HexStr)
'                    ''''          and compare with the CRC readCode.
'                    m_crchexStr = UCase(gS_CFG_CRC_HexStr(Site))
'
'                    m_tsName = m_tsName + "_" + m_crchexStr
'
'                    If (gB_eFuse_Disable_ChkLMT_Flag = True) Then
'                        ''''set Pass
'                        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                    Else
'                        If (UCase(CStr(m_value)) = m_crchexStr) Then
'                            ''''Pass
'                            TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                        Else
'                            ''''Fail
'                            TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=m_tsName
'                        End If
'                    End If
'                Else
'                    ''''other cases, 20160927 update
'                    m_testValue = m_decimal
'                End If
'
'                ''''20160108 New
'                If (m_algorithm <> "firstbits" And m_algorithm <> "crc" And m_algorithm <> "cond") Then
'                    ''''<NOTICE> 20160108
'                    Call auto_eFuse_chkLoLimit("CFG", i, m_stage, m_lolmt)
'                    Call auto_eFuse_chkHiLimit("CFG", i, m_stage, m_hilmt)
'
'                    ''''20170811 update
'                    If (m_bitwidth >= 32) Then
'                        ''m_tsName = m_tsName + "_" + m_hexStr
'                        ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
'                        m_lolmt = auto_Value2HexStr(m_lolmt, m_bitwidth)
'                        m_hilmt = auto_Value2HexStr(m_hilmt, m_bitwidth)
'
'                        ''''------------------------------------------
'                        ''''compare with lolmt, hilmt
'                        ''''m_testValue 0 means fail
'                        ''''m_testValue 1 means pass
'                        ''''------------------------------------------
'                        m_testValue = auto_TestStringLimit(m_hexStr, CStr(m_lolmt), CStr(m_hilmt))
'                        m_lolmt = 1
'                        m_hilmt = 1
'                    Else
'                        ''''20160927 update the new logical methodology for the unexpected binary decode.
'                        If (auto_isHexString(CStr(m_lolmt)) = True) Then
'                            ''''translate to double value
'                            m_lolmt = auto_HexStr2Value(m_lolmt)
'                        Else
'                            ''''doNothing, m_lolmt = m_lolmt
'                        End If
'
'                        If (auto_isHexString(CStr(m_hilmt)) = True) Then
'                            ''''translate to double value
'                            m_hilmt = auto_HexStr2Value(m_hilmt)
'                        Else
'                            ''''doNothing, m_hilmt = m_hilmt
'                        End If
'                    End If
'                    TheExec.Flow.TestLimit resultVal:=m_testValue, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsName, unit:=m_unitType, ScaleType:=m_scale
'                End If
'            Next i
'
'            ''''--------------------------------------------------------------------------------------------
'            ''''20160503 update for the 2nd way to prevent for the temp sensor trim will all zero values
'            ''''20160907 update
'            Dim m_valueSum As Long
'            Dim m_matchTMPS_flag As Boolean
'            m_valueSum = 0 ''''initialize
'            m_matchTMPS_flag = False
'            m_stage = "" ''''<MUST> 20160617 update, if the "trim" is existed then m_stage has its correct value.
'            For i = 0 To UBound(CFGFuse.Category)
'                m_catename = UCase(CFGFuse.Category(i).Name)
'                m_algorithm = LCase(CFGFuse.Category(i).Algorithm)
'                If (m_catename Like "TEMP_SENSOR*" Or m_algorithm = "tmps") Then ''''was m_algorithm = "trim", 20171103 update
'                    m_stage = LCase(CFGFuse.Category(i).Stage)
'                    m_decimal = CFGFuse.Category(i).Read.Decimal(Site)
'                    m_valueSum = m_valueSum + m_decimal
'                    m_matchTMPS_flag = True
'                End If
'            Next i
'            If (m_matchTMPS_flag = True) Then
'                ''''if Job >= m_stage then m_valueSum >= 1
'                If (checkJob_less_Stage_Sequence(m_stage) = False) Then
'                    TheExec.Flow.TestLimit resultVal:=m_valueSum, lowVal:=1, Tname:="CFG_TMPS_SUM"
'                Else
'                    ''''if Job < m_stage then m_valueSum = 0
'                    TheExec.Flow.TestLimit resultVal:=m_valueSum, lowVal:=0, hiVal:=0, Tname:="CFG_TMPS_SUM"
'                End If
'            End If
'            ''''--------------------------------------------------------------------------------------------
'
'            TheExec.Datalog.WriteComment ""
'        Next Site
'    Else
'        ''''do nothing
'    End If
'
'    Call UpdateDLogColumns__False
'
'    ''''201811XX reset
'    If (gS_JobName = "cp1_early") Then gS_JobName = "cp1" ''''<MUST>

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_ConfigRead_Decode(ReadPatSet As Pattern, PinRead As PinList, Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ConfigRead_Decode"

    ''''-------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '=====================================================
    '=  Validate/Load Read patterns (save 1st Run time)  =
    '=====================================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(ReadPatSet, ReadPatt, Validating_) = True) Then Exit Function
    ''''-------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim eFuse_Pgm_Bit() As Long
    Dim i As Long
    Dim j As Long

    Dim SingleStrArray() As String
    Dim DoubleBitArray() As Long
    Dim SingleBitArray() As Long

    Dim CapWave As New DSPWave
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    Dim SignalCap As String
    Dim crcBinStr As String     ''''' 20161003 ADD CRC

    ''''-------------------------------------------------------------------------------------
    ReDim SingleStrArray(EConfigReadCycle - 1, TheExec.sites.Existing.Count - 1)
    ''ReDim SingleBitArray(EConfigTotalBitCount - 1)
    ''ReDim DoubleBitArray(EConfigBitPerBlockUsed - 1)
    ReDim gL_CFG_Sim_FuseBits(TheExec.sites.Existing.Count - 1, EConfigTotalBitCount - 1) ''''it's for the simulation

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
  
    '====================================================
    '=  1. Define the Bit length of Programming         =
    '=  2. Define the cycle number of Compare Pattern   =
    '====================================================
    ReDim eFuse_Pgm_Bit(EConfigTotalBitCount - 1)
  
    SignalCap = "SignalCapture"
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, EConfigReadCycle, CapWave
    
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)


''''201812XX update
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
    Dim testName As String

    m_Fusetype = eFuse_CFG
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_stage = True

    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>

    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, CapWave, m_FBC, blank_stage, allBlank)
    ''''----------------------------------------------------

    ''''Always decode here
    If (True) Then
        ''''''''if there is any site which is non-blank, then decode to gDW_CFG_Read_Decimal_Cate [check later]
        ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
        ''''it will be present by Hex and Binary compare with the limit later on.
        ''''<NOTICE> gDW_CFG_Read_Decimal_Cate has been decided in non-blank and/or MarginRead
        ''Call auto_eFuse_setReadData(eFuse_CFG, gDW_CFG_Read_DoubleBitWave, gDW_CFG_Read_Decimal_Cate)
        'Call auto_eFuse_setReadData(eFuse_CFG)

        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(eFuse_CFG)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_CFG)
    
        Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_CFG, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_CFG, False, gB_eFuse_printReadCate)
    End If

    gL_CFG_FBC = m_FBC
    
    ''Call UpdateDLogColumns(gI_CFG_catename_maxLen)
    testName = "CFG_Read_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName
    ''Call UpdateDLogColumns__False

    DebugPrintFunc ReadPatSet.Value
    
    ''''20170111 Add
    Call auto_eFuse_ReadAllData_to_DictDSPWave("CFG", False, False)

Exit Function

End If


'
'
'    auto_eFuse_DSSC_ReadDigCap_32bits EConfigReadCycle, PinRead.Value, SingleStrArray, capWave, allblank ''''Here Must use local variable
'
'    Call UpdateDLogColumns(gI_CFG_catename_maxLen)
'
'    Dim Count As Long
'    Count = 0 'Initialization
'    Call auto_GetSiteFlagName(Count, gS_cfgFlagname, False)
'    If (gB_findCFGCondTable_flag) Then
'        If (Count = 0) Then
'            gS_cfgFlagname = "A00"
'        End If
'    Else
'        ''''20160902 update for the case CP1 fuse CFG_Condition already, then CP2/CPx needs to get the Flag name.
'        If (Count <> 1) Then
'            gS_cfgFlagname = "ALL_0"
'        End If
'    End If
'
'    '================================================
'    '=  1. Make Program bit array                   =
'    '=  2. Make Read Compare bit array              =
'    '================================================
'    For Each Site In TheExec.Sites
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To EConfigReadCycle - 1
'                ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'                SingleStrArray(i, Site) = StrReverse(SingleStrArray(i, Site))
'            Next i
'        End If
'
'        ''''20160202 update, 20161107 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            allblank(Site) = False
'            blank_stage(Site) = True ''''True or False(sim for re-test)
'            If (blank_stage(Site) = True And gS_JobName = "cp1") Then allblank(Site) = True
'
'            If (gB_CFGSVM_A00_CP1 = True And gS_JobName <> "cp1") Then
'                If (gB_CFG_SVM = True) Then
'                    gB_CFGSVM_BIT_Read_ValueisONE(Site) = True
'                End If
'            End If
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
'            ReDim Expand_eFuse_Pgm_Bit(EConfigTotalBitCount * EConfig_Repeat_Cyc_for_Pgm - 1)
'            ReDim eFusePatCompare(EConfigReadCycle - 1)
'            Call auto_make_CFG_Pgm_for_Simulation(SingleBitArray, Expand_eFuse_Pgm_Bit, eFusePatCompare, blank_stage(Site), False) ''''debug/showPrint if True
'
'            Dim k As Long
'            Dim m_tmpStr As String
'            For i = 0 To EConfigReadCycle - 1
'                m_tmpStr = ""
'                For j = 0 To EConfigReadBitWidth - 1
'                    k = j + i * EConfigReadBitWidth ''''MUST
'                    gL_CFG_Sim_FuseBits(Site, k) = SingleBitArray(k) ''''it's used for the Read/Syntax simulation
'                    m_tmpStr = CStr(SingleBitArray(k)) + m_tmpStr
'                Next j
'                SingleStrArray(i, Site) = m_tmpStr
'            Next i
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        Call auto_OR_2Blocks("CFG", SingleStrArray, SingleBitArray, DoubleBitArray)  ''''calc gL_CFG_FBC
'
'        '====================================================
'        '=  Print the all eFuse Bit data from DSSC          =
'        '====================================================
'        TheExec.Datalog.WriteComment ""
'        TheExec.Datalog.WriteComment "Read All Config eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'        Call auto_PrintAllBitbyDSSC(SingleBitArray, EConfigReadCycle, EConfigTotalBitCount, EConfigReadBitWidth) ''must use EConfigReadBitWidth in Right2Left mode
'
'        ''''' 20161003 ADD CRC
'        If (Trim(gS_CFG_CRC_Stage) <> "") Then
'            If (checkJob_less_Stage_Sequence(gS_CFG_CRC_Stage) = True) Then
'                gS_CFG_CRC_HexStr(Site) = "0000"
'            Else
'                gS_CFG_CRC_HexStr(Site) = auto_CFG_CRC2HexStr(DoubleBitArray, crcBinStr)
'            End If
'        End If
'
'        ''''<Important> User Need to check the content inside
'        Call auto_Decode_CfgBinary_Data(DoubleBitArray, Not allblank(Site)) ''''true for debug, only NOT allBlank to show the decode result.
'
'    Next Site ''For Each Site In TheExec.Sites
'
'    TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0
'
'    Call UpdateDLogColumns__False
'    DebugPrintFunc ReadPatSet.Value
'    ''''----------------------------------------------------------------------------------------------------------
'
'    ''''20170111 Add
'    Call auto_eFuse_ReadAllData_to_DictDSPWave("CFG", False, False)

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

''''20161104 New
''''This function is used to do the pre-Check if
Public Function auto_eFuse_IDS_BinCut_PreCheck()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_IDS_BinCut_PreCheck"
    
    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim m_tmpStr As String
    Dim m_decimal As Long
    Dim m_algorithm As String
    Dim m_defreal As String
    Dim m_exist_bincut_CFG As Boolean
    'Dim m_exist_bincut_UDR As Boolean
    Dim m_exist_bincut_UDR_E As Boolean
    Dim m_exist_bincut_UDR_P As Boolean
    Dim m_CFG_idsSum As Long ''''bit summation of IDS    in CFG.
    Dim m_CFG_bctSum As Long ''''bit summation of Bincut in CFG.
    'Dim m_UDR_bctSum As Long ''''bit summation of Bincut in UDR.
    Dim m_UDR_E_bctSum As Long ''''bit summation of Bincut in UDR.
    Dim m_UDR_P_bctSum As Long ''''bit summation of Bincut in UDR.
    Dim m_readValueSum As New SiteLong
    Dim m_IDS_isZero_Falg As Boolean        ''''True if anyone of IDS catrgories is zero.
    Dim m_CFG_bincut_isZero_Flag As Boolean ''''True if anyone of CFG bincut catrgories is zero.
    'Dim m_UDR_bincut_isZero_Flag As Boolean ''''True if anyone of UDR bincut catrgories is zero.
    Dim m_UDR_E_bincut_isZero_Flag As Boolean ''''True if anyone of UDR bincut catrgories is zero.
    Dim m_UDR_P_bincut_isZero_Flag As Boolean ''''True if anyone of UDR bincut catrgories is zero.
    Dim m_stage As String

    m_exist_bincut_CFG = False
    m_exist_bincut_UDR_E = False
    m_exist_bincut_UDR_P = False

    ''''First, check if there is the "BinCut" in the field "Default or Real" on the BitDefTable.
    For Each site In TheExec.sites
        For i = 0 To UBound(CFGFuse.Category)
            m_algorithm = LCase(CFGFuse.Category(i).algorithm)
            m_defreal = LCase(CFGFuse.Category(i).Default_Real)
            If (m_algorithm = "vddbin" And m_defreal = "bincut") Then
                m_exist_bincut_CFG = True
                Exit For
            End If
        Next i

        ''''20171103 update
        'If (gB_findUDR_flag = True) Then
        '    For i = 0 To UBound(UDRFuse.Category)
        '        m_algorithm = LCase(UDRFuse.Category(i).algorithm)
        '        m_defreal = LCase(UDRFuse.Category(i).Default_Real)
        '        If (m_algorithm = "vddbin" And m_defreal = "bincut") Then
        '            m_exist_bincut_UDR = True
        '            Exit For
        '        End If
        '    Next i
        'ElseIf (gB_findUDRE_flag = True) Then
            For i = 0 To UBound(UDRE_Fuse.Category)
                m_algorithm = LCase(UDRE_Fuse.Category(i).algorithm)
                m_defreal = LCase(UDRE_Fuse.Category(i).Default_Real)
                If (m_algorithm = "vddbin" And m_defreal = "bincut") Then
                m_exist_bincut_UDR_E = True
                    Exit For
                End If
            Next i
        'ElseIf (gB_findUDRP_flag = True) Then
            For i = 0 To UBound(UDRP_Fuse.Category)
                m_algorithm = LCase(UDRP_Fuse.Category(i).algorithm)
                m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
                If (m_algorithm = "vddbin" And m_defreal = "bincut") Then
                m_exist_bincut_UDR_P = True
                    Exit For
                End If
            Next i
        'End If
    Next site ''For Each Site In TheExec.Sites

    ''''Only both CFG and UDR have the bincut then go the below check process (2nd Step).
    If (m_exist_bincut_CFG = False Or m_exist_bincut_UDR_E = False Or m_exist_bincut_UDR_P = False) Then
        TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:="eFuse_noBinCut"
        Exit Function
    End If

    ''''2nd Step, check if (all IDS sum is zero) and (all Vddbin sum is zero)
    ''''          check if (all IDS is non-zero) and (all Vddbin is non-zero)
    Call UpdateDLogColumns(gI_CFG_catename_maxLen)

    For Each site In TheExec.sites
        ''''------------------------------------
        ''''initialize per Site
        ''''------------------------------------
        m_CFG_idsSum = 0
        m_CFG_bctSum = 0
        'm_UDR_bctSum = 0
        m_UDR_E_bctSum = 0
        m_UDR_P_bctSum = 0
        m_IDS_isZero_Falg = False
        m_CFG_bincut_isZero_Flag = False
        'm_UDR_bincut_isZero_Flag = False
        m_UDR_E_bincut_isZero_Flag = False
        m_UDR_P_bincut_isZero_Flag = False
        ''''------------------------------------
        For i = 0 To UBound(CFGFuse.Category)
            m_algorithm = LCase(CFGFuse.Category(i).algorithm)
            m_defreal = LCase(CFGFuse.Category(i).Default_Real)
            m_stage = UCase(CFGFuse.Category(i).Stage)
            If (m_algorithm = "ids" And m_defreal = "real" And UCase(m_stage) = "CP1") Then
                m_decimal = CFGFuse.Category(i).Read.Decimal(site)
                m_CFG_idsSum = m_CFG_idsSum + m_decimal
                If (m_decimal = 0) Then m_IDS_isZero_Falg = True
            ElseIf (m_algorithm = "vddbin" And m_defreal = "bincut") Then
                m_decimal = CFGFuse.Category(i).Read.Decimal(site)
                m_CFG_bctSum = m_CFG_bctSum + m_decimal
                If (m_decimal = 0) Then m_CFG_bincut_isZero_Flag = True
            End If
        Next i
        
        ''''20171103 update
        'If (gB_findUDR_flag = True) Then
        '    For i = 0 To UBound(UDRFuse.Category)
        '        m_algorithm = LCase(UDRFuse.Category(i).algorithm)
        '        m_defreal = LCase(UDRFuse.Category(i).Default_Real)
        '        If (m_algorithm = "vddbin" And m_defreal = "bincut") Then
        '            m_decimal = UDRFuse.Category(i).Read.Decimal(Site)
        '            m_UDR_bctSum = m_UDR_bctSum + m_decimal
        '            If (m_decimal = 0) Then m_UDR_bincut_isZero_Flag = True
        '        End If
        '    Next i
        'ElseIf (gB_findUDRE_flag = True) Then
            For i = 0 To UBound(UDRE_Fuse.Category)
                m_algorithm = LCase(UDRE_Fuse.Category(i).algorithm)
                m_defreal = LCase(UDRE_Fuse.Category(i).Default_Real)
                If (m_algorithm = "vddbin" And m_defreal = "bincut") Then
                    m_decimal = UDRE_Fuse.Category(i).Read.Decimal(site)
                m_UDR_E_bctSum = m_UDR_E_bctSum + m_decimal
                If (m_decimal = 0) Then m_UDR_E_bincut_isZero_Flag = True
                End If
            Next i
        'ElseIf (gB_findUDRP_flag = True) Then
            For i = 0 To UBound(UDRP_Fuse.Category)
                m_algorithm = LCase(UDRP_Fuse.Category(i).algorithm)
                m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
                If (m_algorithm = "vddbin" And m_defreal = "bincut") Then
                    m_decimal = UDRP_Fuse.Category(i).Read.Decimal(site)
                m_UDR_P_bctSum = m_UDR_P_bctSum + m_decimal
                If (m_decimal = 0) Then m_UDR_P_bincut_isZero_Flag = True
                End If
            Next i
        'End If

        ''''<Important>
        ''''Check if IDS/BinCut are burned at the same RUN or all Empty.
        If (m_CFG_idsSum = 0 And m_CFG_bctSum = 0 And m_UDR_E_bctSum = 0 And m_UDR_P_bctSum = 0) Then
            m_readValueSum(site) = 0 ''''set it Pass
        ElseIf (m_CFG_idsSum <> 0 And m_CFG_bctSum <> 0 And m_UDR_E_bctSum <> 0 And m_UDR_P_bctSum <> 0) Then
            If (m_IDS_isZero_Falg Or m_CFG_bincut_isZero_Flag Or m_UDR_E_bincut_isZero_Flag Or m_UDR_P_bincut_isZero_Flag) Then
                ''''means that there is "zero" in one of these categories
                ''''Fail case
                If (m_IDS_isZero_Falg) Then TheExec.Datalog.WriteComment vbTab & "<WARNING> CFG IDS:: There is the Empty (Zero) case."
                If (m_CFG_bincut_isZero_Flag) Then TheExec.Datalog.WriteComment vbTab & "<WARNING> CFG BinCut:: There is the Empty (Zero) case."
                ''''If (m_UDR_bincut_isZero_Flag) Then TheExec.Datalog.WriteComment vbTab & "<WARNING> UDR BinCut:: There is the Empty (Zero) case."
                If (m_UDR_E_bincut_isZero_Flag) Then TheExec.Datalog.WriteComment vbTab & "<WARNING> UDR_E BinCut:: There is the Empty (Zero) case."
                If (m_UDR_P_bincut_isZero_Flag) Then TheExec.Datalog.WriteComment vbTab & "<WARNING> UDR_P BinCut:: There is the Empty (Zero) case."
                m_readValueSum(site) = m_CFG_idsSum + m_CFG_bctSum + m_UDR_E_bctSum + m_UDR_P_bctSum
            Else
                m_readValueSum(site) = 0 ''''set it Pass
            End If
        Else
            ''''Fail case, IDS, BinCut does NOT burn at the same RUN.
            '''m_readValueSum(Site) = m_CFG_idsSum + m_CFG_bctSum + m_UDR_bctSum
            m_readValueSum(site) = m_CFG_idsSum + m_CFG_bctSum + m_UDR_E_bctSum + m_UDR_P_bctSum
        End If
    Next site ''For Each Site In TheExec.Sites
    
    TheExec.Flow.TestLimit resultVal:=m_readValueSum, lowVal:=0, hiVal:=0

    Call UpdateDLogColumns__False
    ''''----------------------------------------------------------------------------------------------------------

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

''''20161115, its from Tcay (SWLin generated), and update here
''''It's used to do the post check eFuse IDS/BinCut if meet the BinCut Table (GradeValue)
''''was auto_Compare_GradeValue_IDS_from_eFuse, rename auto_eFuse_IDS_BinCut_PostCheck
Public Function auto_eFuse_IDS_BinCut_PostCheck() As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_IDS_BinCut_PostCheck"
    
    Dim site As Variant
    Dim i As Long, j As Long
    Dim tmpStr As String
    Dim IDS_FuseNameTmp As String
    Dim tmpIdsvalue As Double            ''''eFuse IDS value in (mA)
    Dim IDS_FuseValue() As New SiteDouble
    Dim IDS_FuseName() As String
    Dim PmodeNum As Integer
    Dim BinCut_Domain_Index As New Dictionary
    Dim BinCut_Pmode_Nmae As String
    Dim Bincut_Domain_Name As String    ''''BinCut Domain name
    Dim Pmode_match_flag As Boolean
    Dim Pmode_FuseCateName As String      ''''After serach to get the related eFuse category name of this Pmode
    Dim tmpAlgorithm As String
    Dim tmpDefreal As String
    Dim tmpResolution As Double
    Dim Bincut_Domain_CateName As String        ''''eFuse Vddbin category name
    Dim tmpBinCutNum As Long
    Dim tmpEQNum As Long
    Dim Bincut_totalEQStepNum As Long
    Dim tmpGRADEVDD As Double
    Dim tmpIdsCurrent As Double
   
    ''-----------------------------------------------------------
    Call UpdateDLogColumns(30)
    ReDim IDS_FuseValue(UBound(pinGroup_CorePower)) As New SiteDouble
    ReDim IDS_FuseName(UBound(pinGroup_CorePower)) As String
    
    For Each site In TheExec.sites
        'pinGroup_CorePower(bincutDomain) to Fuse BDF ids name
        For i = 0 To UBound(pinGroup_CorePower)
            'Special handle for bincut domain -> BDF ids name, Need to maintain with each pjt
            '-----------------------------------------------------------
'            If (LCase(pinGroup_CorePower(i)) Like ("*gpu0")) Then
'                IDS_FuseNameTmp = "ids_vdd_gpu0_bincheck"
'            ElseIf (LCase(pinGroup_CorePower(i)) Like ("*gpu1")) Then
'                IDS_FuseNameTmp = "ids_vdd_gpu1_bincheck"
'            ElseIf (LCase(pinGroup_CorePower(i)) Like ("*pcpu0")) Then
'                IDS_FuseNameTmp = "ids_vdd_pcpu0_25c_4"
'            ElseIf (LCase(pinGroup_CorePower(i)) Like ("*pcpu1")) Then
'                IDS_FuseNameTmp = "ids_vdd_pcpu1_25c_4"
'            Else
            IDS_FuseNameTmp = LCase("ids_" + pinGroup_CorePower(i))
'            End If
            '-----------------------------------------------------------
            tmpIdsvalue = CFGFuse.Category(CFGIndex(IDS_FuseNameTmp)).Read.Value
            tmpStr = Format(tmpIdsvalue, "0.000")
            IDS_FuseValue(i) = tmpIdsvalue * 0.001
            IDS_FuseName(i) = IDS_FuseNameTmp
            TheExec.Datalog.WriteComment "Site(" & site & "), " & IDS_FuseNameTmp & FormatNumeric(tmpStr, 8) & " mA"
            
            If Not (BinCut_Domain_Index.Exists(pinGroup_CorePower(i))) Then
                BinCut_Domain_Index.Add pinGroup_CorePower(i), i
            End If
        Next i
        TheExec.Datalog.WriteComment "" '20191021
    Next site
    '----------------------------------------------------------------------------------------------
   
    For PmodeNum = 0 To MaxPerformanceModeCount - 1
        If AllBinCut(PmodeNum).Used = True Then
            BinCut_Pmode_Nmae = UCase(VddBinName(PmodeNum)) 'EX:VDD_SOC_MS001
            Bincut_Domain_Name = UCase(AllBinCut(PmodeNum).powerPin)    'EX:VDD_SOC
            
            ''''-----------------------------------------------------------------
            '''' <MUST> search for the matched eFuse Vddbin category name
            ''''-----------------------------------------------------------------
            Pmode_match_flag = False
            Pmode_FuseCateName = ""
            
            'Special handle for bincut domain -> fuse category, Need to maintain with each pjt
            '-----------------------------------------------------------
            If Bincut_Domain_Name = "VDD_PCPU" Then
                ''''search in UDRFuse Category
                For i = 0 To UBound(UDRP_Fuse.Category)
                    tmpAlgorithm = LCase(UDRP_Fuse.Category(i).algorithm)
                    tmpDefreal = LCase(UDRP_Fuse.Category(i).Default_Real)
                    If (tmpAlgorithm = "vddbin" And tmpDefreal = "bincut") Then
                        Bincut_Domain_CateName = UCase(UDRP_Fuse.Category(i).Name)
                        If (VddBinStr2Enum(Bincut_Domain_CateName) = PmodeNum) Then
                            Pmode_match_flag = True
                            Pmode_FuseCateName = Bincut_Domain_CateName
                            tmpResolution = UDRP_Fuse.Category(i).Resoultion
                            Exit For
                        End If
                    End If
                Next i
'                ElseIf Bincut_Domain_Name = "VDD_PCPU1" Then
'                ''''search in UDRFuse Category
'                For i = 0 To UBound(UDRP01_Fuse.Category)
'                    tmpAlgorithm = LCase(UDRP01_Fuse.Category(i).algorithm)
'                    tmpDefreal = LCase(UDRP01_Fuse.Category(i).Default_Real)
'                    If (tmpAlgorithm = "vddbin" And tmpDefreal = "bincut") Then
'                        Bincut_Domain_CateName = UCase(UDRP01_Fuse.Category(i).Name)
'                        If (VddBinStr2Enum(Bincut_Domain_CateName) = PmodeNum) Then
'                            Pmode_match_flag = True
'                            Pmode_FuseCateName = Bincut_Domain_CateName
'                            tmpResolution = UDRP01_Fuse.Category(i).Resoultion
'                            Exit For
'                        End If
'                    End If
'                Next i
                ElseIf Bincut_Domain_Name = "VDD_ECPU" Then
                ''''search in UDRE_Fuse Category
                For i = 0 To UBound(UDRE_Fuse.Category)
                    tmpAlgorithm = LCase(UDRE_Fuse.Category(i).algorithm)
                    tmpDefreal = LCase(UDRE_Fuse.Category(i).Default_Real)
                    If (tmpAlgorithm = "vddbin" And tmpDefreal = "bincut") Then
                        Bincut_Domain_CateName = UCase(UDRE_Fuse.Category(i).Name)
                        If (VddBinStr2Enum(Bincut_Domain_CateName) = PmodeNum) Then
                            Pmode_match_flag = True
                            Pmode_FuseCateName = Bincut_Domain_CateName
                            tmpResolution = UDRE_Fuse.Category(i).Resoultion
                            Exit For
                        End If
                    End If
                Next i
                '-----------------------------------------------------------
                Else 'Bincut_Domain_Name = "VDD_SOC" Or Bincut_Domain_Name = "VDD_GPU" Or Bincut_Domain_Name = "VDD_DCS_DDR" Or ...(other domain) Then
                    For i = 0 To UBound(CFGFuse.Category)
                        tmpAlgorithm = LCase(CFGFuse.Category(i).algorithm)
                        tmpDefreal = LCase(CFGFuse.Category(i).Default_Real)
                        If (tmpAlgorithm = "vddbin" And tmpDefreal = "bincut") Then
                            Bincut_Domain_CateName = UCase(CFGFuse.Category(i).Name)
                            If (VddBinStr2Enum(Bincut_Domain_CateName) = PmodeNum) Then
                                Pmode_match_flag = True
                                Pmode_FuseCateName = Bincut_Domain_CateName
                                tmpResolution = CFGFuse.Category(i).Resoultion
                                Exit For
                            End If
                        End If
                    Next i
                End If
              
            ''''<Important>  WARNING if there is NO the Pmode existed in the eFuse category.
            If (Pmode_match_flag = False) Then
                TheExec.Datalog.WriteComment "<WARNING> This perfomance mode (" + BinCut_Pmode_Nmae + ") is not exisitng in eFuse Bit-Def table Or is the safe voltage."
                TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=BinCut_Pmode_Nmae + "_NotRealBinCut"
                TheExec.Datalog.WriteComment ""
                GoTo Next_Pmode
            End If

            ''''-----------------------------------------------------------------
            ''''Here Bincut_Domain_Name is one of "VDD_CPU", "VDD_GPU", "VDD_SOC"
            ''''-----------------------------------------------------------------
            For Each site In TheExec.sites
                ''''<MUST User Maintain>, select which BinCut table to be used
                If (TheExec.TesterMode = testModeOffline) Then
                    tmpBinCutNum = 1
                Else
                    tmpBinCutNum = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                    If tmpBinCutNum = 4 Then
                        TheExec.Datalog.WriteComment "BinCutNum is mess up"
                        TheExec.ErrorLogMessage ("BinCutNum is mess up")
                        TheExec.Flow.TestLimit resultVal:=tmpBinCutNum, lowVal:=1, hiVal:=3, Tname:="BincutNum Erro!!"
                    End If
                End If
                ''''-------------------------------------------------------------------------

                tmpEQNum = 999 ''''<MUST> Reset and Initial
                Bincut_totalEQStepNum = BinCut(PmodeNum, tmpBinCutNum).Mode_Step + 1
                
                tmpIdsCurrent = IDS_FuseValue(BinCut_Domain_Index.Item(UCase(Bincut_Domain_Name)))
                'Special handle for bincut domain -> BDF ids name, Need to maintain with each pjt
                '-----------------------------------------------------------
                If (LCase(Bincut_Domain_Name) Like ("*pcpu")) Then
                    tmpGRADEVDD = auto_eFuse_GetReadValue("UDRP", Pmode_FuseCateName, False)
'                ElseIf (LCase(Bincut_Domain_Name) Like ("*pcpu1")) Then
'                    tmpGRADEVDD = auto_eFuse_GetReadValue("UDRP", Pmode_FuseCateName, False)
                ElseIf (LCase(Bincut_Domain_Name) Like ("*ecpu")) Then
                    tmpGRADEVDD = auto_eFuse_GetReadValue("UDRE", Pmode_FuseCateName, False)
                '-----------------------------------------------------------
                Else
                    tmpGRADEVDD = auto_eFuse_GetReadValue("CFG", Pmode_FuseCateName, False)
                End If
                
                TheExec.Flow.TestLimit resultVal:=tmpGRADEVDD * 0.001, scaletype:=scaleMilli, Unit:=unitVolt, Tname:=BinCut_Pmode_Nmae & "_BCT"
                
                '********************************************************************************
                '                  To check fused IDS_VDD_CPU is not empty
                '********************************************************************************
                If tmpIdsCurrent = 0# Then
                    TheExec.AddOutput BinCut_Pmode_Nmae + ", its IDS current is zero."
                    TheExec.Datalog.WriteComment BinCut_Pmode_Nmae + ", its IDS current is zero."
                    TheExec.Flow.TestLimit resultVal:=tmpIdsCurrent, lowVal:=1, hiVal:=1, Tname:=BinCut_Pmode_Nmae & "_IDS"
                Else
                    ''TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=BinCut_Pmode_Nmae & "_IDS"
                    TheExec.Flow.TestLimit resultVal:=tmpIdsCurrent, scaletype:=scaleMilli, Unit:=unitAmp, Tname:=BinCut_Pmode_Nmae & "_IDS"
                    Call auto_Compare_EQN_Voltage_Per_Site(tmpIdsCurrent, PmodeNum, tmpGRADEVDD, tmpResolution, tmpBinCutNum, tmpEQNum)
                    If tmpEQNum > Bincut_totalEQStepNum Then
                        TheExec.Datalog.WriteComment "The EQN Number Can Not Be Found!!"
                        TheExec.ErrorLogMessage "The EQN Number Can Not Be Found!!"
                    End If
                End If
                TheExec.Flow.TestLimit resultVal:=tmpEQNum, lowVal:=1, hiVal:=Bincut_totalEQStepNum, Tname:=BinCut_Pmode_Nmae & "_EQ"
                TheExec.Datalog.WriteComment ""
            Next site
        Else
            ''''Here is .Used = False
            ''''Debug
            If (False) Then
                TheExec.Datalog.WriteComment "unUsed >>> PmodeNum = " & PmodeNum & ", VddBinName = " + UCase(VddBinName(PmodeNum))
            End If
        End If ''''If AllBinCut(PmodeNum).Used = True Then
Next_Pmode:
    Next PmodeNum

    Call UpdateDLogColumns__False

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

''''20170711 update
''''20170630, add for the eaarly Fuse CFG_Condition and SEP SCAN bits
Public Function auto_ConfigBlankChk_Early_byStage(ReadPatSet As Pattern, PinRead As PinList, _
                    Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ConfigBlankChk_Early_byStage"

    ''''-------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(ReadPatSet, ReadPatt, Validating_) = True) Then Exit Function
    ''''-------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim i As Long, j As Long
    Dim FBCChkPass As Boolean
    Dim ResultFlag As New SiteLong
    Dim testName As String, PinName As String
    Dim PrintSiteVarResult As String
    Dim SiteVarValue As Long
    Dim SingleBitArray() As Long
    Dim DoubleBitArray() As Long
    Dim SingleDoubleFBC As Long
    Dim SignalCap As String, CapWave As New DSPWave
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    Dim Count As Long
    Dim m_jobinStage_flag As Boolean
    Dim blank_Cond As Boolean
    Dim blank_SCAN As Boolean
    Dim blank_stage_noCond_SCAN As Boolean

    ReDim gL_CFG_Sim_FuseBits(TheExec.sites.Existing.Count - 1, EConfigTotalBitCount - 1) ''''it's for the simulation

    SignalCap = "SignalCapture" 'define capture signal name

    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    '****** Initialize Site Varaible ******
    Dim m_siteVar As String
    m_siteVar = "CFGChk_Var"
    For Each site In TheExec.sites.Existing
        TheExec.sites(site).SiteVariableValue(m_siteVar) = -1
    Next site

    '================================================
    '=  Setup HRAM/DSSC capture cycles              =
    '================================================
    '*** Setup HARM/DSSC Trigger and Capture parameter ***
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, EConfigReadCycle, CapWave  'setup
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)    'run read pattern and capture
    

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    Count = 0 'Initialization
    Call auto_GetSiteFlagName(Count, gS_cfgFlagname, True)
    If Count <> 1 Then
        If Count = 0 Then TheExec.Datalog.WriteComment vbCrLf & "<WARNING> There is NO any CFG eFuse conditions Flag selected. Please check it!! " ''''20160927 add
        SiteVarValue = 0
        For Each site In TheExec.sites
            TheExec.sites(site).SiteVariableValue(m_siteVar) = SiteVarValue
            ''''20161110 update per Jack's comment, "F_config_flag_missing" (BinTable Flag)
            TheExec.sites.Item(site).FlagState("F_config_flag_missing") = logicTrue
        Next site
        TheExec.Flow.TestLimit resultVal:=Count, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Error" ''''set fail
        TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
        DebugPrintFunc ReadPatSet.Value
        Exit Function
    Else
        ''''Count==1
        TheExec.Datalog.WriteComment funcName + ":: The Selected CFG Condition is " + gS_cfgFlagname + " (from OI or Eng mode)"
        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_" + gS_cfgFlagname ''''set pass
    End If
    
    ''''20170911 update, debug purpose
    If (False) Then Call auto_display_CFG_Cond_Table_by_PKGName(gS_cfgFlagname)
    
    ''''----------------------------------------------------
    ''''201808XX New Method by DSPWave
    ''''----------------------------------------------------
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong
    
    m_jobinStage_flag = auto_eFuse_JobExistInStage("CFG", True) ''''<MUST>
    
    ''''201811XX
    ''''Set Job as cp1_early to fuse early bits
    ''''Will be reset back @ Syntax Check
    If (gS_JobName = "cp1") Then gS_JobName = "cp1_early" ''''<MUST>
    
    m_Fusetype = eFuse_CFG
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_stage = True

    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=0 (Stage Early Bits)
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 0, CapWave, m_FBC, blank_stage, allBlank)
    ''''----------------------------------------------------

    If (blank_stage.Any(False) = True) Then
        ''''''''if there is any site which is non-blank, then decode to gDW_CFG_Read_Decimal_Cate [check later]
        ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
        ''''it will be present by Hex and Binary compare with the limit later on.
        ''''<NOTICE> gDW_CFG_Read_Decimal_Cate has been decided in non-blank and/or MarginRead
        ''Call auto_eFuse_setReadData(eFuse_CFG, gDW_CFG_Read_DoubleBitWave, gDW_CFG_Read_Decimal_Cate)
        'Call auto_eFuse_setReadData(eFuse_CFG)
        
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(eFuse_CFG)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_CFG)
    
        Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_CFG, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_CFG, False, gB_eFuse_printReadCate)
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
    m_SiteVarValue = blank_stage
    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
    If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only

    For Each site In TheExec.sites
        If (auto_eFuse_GetAllPatTestPass_Flag("CFG") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    gL_CFG_FBC = m_FBC
    
    Call UpdateDLogColumns(gI_CFG_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "CFG_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "CFG_Blank_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName
    
    If (gB_EFUSE_DVRV_ENABLE = True And gS_JobName = "cp1_early") Then gS_JobName = "cp1"

    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value

Exit Function

End If


'    ''''While being in the "Retest" stage, the DSSC Read result will be used in the instance "XXXX_SingleDoubleBit"
'    ''''So it's needed to set SingleStrArray() to global gS_SingleStrArray()
'    ''''201811XX, move to afterward
'    auto_eFuse_DSSC_ReadDigCap_32bits EConfigReadCycle, PinRead.Value, gS_SingleStrArray, capWave, allblank 'read back in singlestrarray
'
'    ''''it's used to identify if the Job Name is existed in the CFG portion of the eFuse BitDef table.
'    ''''20170630 update, check if the Job Name is existed in the CFG_Condition_Table
'    m_jobinStage_flag = auto_eFuse_JobExistInStage("CFG", True) ''''<MUST>
'
'    Call UpdateDLogColumns(gI_CFG_catename_maxLen)
'
'    ''''20160526 update
'    ''''---------------------------------------------------------------
'    If (gS_JobName = gS_CFG_firstbits_stage Or gS_JobName = "wlft" Or gS_JobName Like "ft*") Then ''''20151214 update
'        ''''Get Config security code from OI
'        Count = 0 'Initialization
'        Call auto_GetSiteFlagName(Count, gS_cfgFlagname, True)
'        If Count <> 1 Then
'            If Count = 0 Then TheExec.Datalog.WriteComment vbCrLf & "<WARNING> There is NO any CFG eFuse conditions Flag selected. Please check it!! " ''''20160927 add
'            SiteVarValue = 0
'            For Each Site In TheExec.Sites
'                TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'                ''''20161110 update per Jack's comment, "F_config_flag_missing" (BinTable Flag)
'                TheExec.Sites.Item(Site).FlagState("F_config_flag_missing") = logicTrue
'            Next Site
'            TheExec.Flow.TestLimit resultVal:=Count, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Error" ''''set fail
'            TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'            GoTo CFGChk_byStage_Early
'        Else
'            ''''Count==1
'            TheExec.Datalog.WriteComment funcName + ":: The Selected CFG Condition is " + gS_cfgFlagname + " (from OI or Eng mode)"
'            TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Set_" + gS_cfgFlagname ''''set pass
'        End If
'    Else
'        Count = 0 'Initialization
'        Call auto_GetSiteFlagName(Count, gS_cfgFlagname, False)
'        If (gB_findCFGCondTable_flag) Then
'            If (Count = 0) Then
'                gS_cfgFlagname = "A00"
'            ElseIf (Count <> 1) Then
'                TheExec.Datalog.WriteComment vbCrLf & "<WARNING> There are more one CFG condition Flag selected. Please check it!! " ''''20160927 add
'                gS_cfgFlagname = "ALL_0"
'                SiteVarValue = 0
'                For Each Site In TheExec.Sites
'                    TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'                    ''''20161110 update per Jack's comment, "F_config_flag_missing" (BinTable Flag)
'                    TheExec.Sites.Item(Site).FlagState("F_config_flag_missing") = logicTrue
'                Next Site
'                TheExec.Flow.TestLimit resultVal:=Count, lowVal:=1, hiVal:=1, Tname:="CFG_Flag_Error" ''''set fail
'                TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'                GoTo CFGChk_byStage_Early
'            End If
'        Else
'            ''''20160902 update for the case CP1 fuse CFG_Condition already, then CP2/CPx needs to get the Flag name.
'            If (Count <> 1) Then
'                gS_cfgFlagname = "ALL_0"
'            End If
'        End If
'    End If
'    ''''---------------------------------------------------------------
'
'    ''''20170911 update
'    Call auto_display_CFG_Cond_Table_by_PKGName(gS_cfgFlagname)
'
'    For Each Site In TheExec.Sites
'        ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To EConfigReadCycle - 1
'                gS_SingleStrArray(i, Site) = StrReverse(gS_SingleStrArray(i, Site))
'            Next i
'        End If
'
'        ''''20160202 Add, 20161108 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            allblank(Site) = False
'            blank_stage(Site) = True ''''True or False(sim for re-test)
'            If (blank_stage(Site) = True And gS_JobName = "cp1") Then allblank(Site) = True
'
'            If (gB_eFuse_CFG_Cond_FTF_done_Flag = True) Then ''''20170923 add
'                blank_stage(Site) = False
'                allblank(Site) = False
'            End If
'
'            If (gB_findCFGTable_flag) Then
'                If (gB_CFGSVM_A00_CP1 = True And gS_JobName <> "cp1") Then
'                    If (gB_CFG_SVM = True) Then
'                        gB_CFGSVM_BIT_Read_ValueisONE(Site) = True
'                    End If
'                End If
'            End If
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
'            ReDim Expand_eFuse_Pgm_Bit(EConfigTotalBitCount * EConfig_Repeat_Cyc_for_Pgm - 1)
'            ReDim eFusePatCompare(EConfigReadCycle - 1)
'
'            If (blank_stage = True) Then
'                blank_Cond = True
'                blank_stage_noCond_SCAN = True
'                Call auto_make_CFG_Pgm_for_Simulation_Early(SingleBitArray, Expand_eFuse_Pgm_Bit, eFusePatCompare, blank_Cond, blank_stage_noCond_SCAN, False) ''''showPrint if True
'            Else
'                Call auto_make_CFG_Pgm_for_Simulation(SingleBitArray, Expand_eFuse_Pgm_Bit, eFusePatCompare, blank_stage(Site), False) ''''debug/showPrint if True
'            End If
'            ''''20161031 update
'            ''''to composite gS_SingleStrArray(i, Site) for the simulation in SingleDoubleBit_Check/Syntax
'            Dim k As Long
'            Dim m_tmpStr As String
'            For i = 0 To EConfigReadCycle - 1
'                m_tmpStr = ""
'                For j = 0 To EConfigReadBitWidth - 1
'                    k = j + i * EConfigReadBitWidth ''''MUST
'                    gL_CFG_Sim_FuseBits(Site, k) = SingleBitArray(k) ''''it's used for the Read/Syntax simulation
'                    m_tmpStr = CStr(SingleBitArray(k)) + m_tmpStr
'                Next j
'                gS_SingleStrArray(i, Site) = m_tmpStr
'            Next i
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        ''''----------------------------------------------------------------------------------
'        ''''<NOTICE>
'        ''''Because we use 'V' instead of 'L' in the DSSC pattern,
'        ''''so that "Blank-Fuse" can not be judged by the API 'Patgen.PatternBurstPassed'.
'        ''''But it was decided in the routine auto_eFuse_DSSC_ReadDigCap_32bits()
'        ''''If theHdw.digital.Patgen.PatternBurstPassed(Site) = False Then  'If not blank
'        ''''----------------------------------------------------------------------------------
'
'        ''''depends on the specific Job to judge if it's blank on the specific stage
'        testName = "CFG_BlankChk_Early_" + UCase(gS_JobName)
'
'        SingleDoubleFBC = 0 ''''init
'        Call auto_OR_2Blocks("CFG", gS_SingleStrArray, SingleBitArray, DoubleBitArray) ''''calc gL_CFG_FBC
'        ''''Call auto_eFuse_BlankChk_FBC_byStage("CFG", SingleBitArray, blank_stage, SingleDoubleFBC)
'
'        Call auto_CFG_blank_check_Cond_SCAN(SingleBitArray, blank_Cond, blank_SCAN, blank_stage_noCond_SCAN)
'
'        If (blank_Cond = True And blank_SCAN = True) Then ''If both are blank
'            ResultFlag(Site) = 0    ''Pass Blank check
'            PinName = "Pass"
'            SiteVarValue = 1
'            If (m_jobinStage_flag = False) Then SiteVarValue = 2 ''''Read Only
'        ElseIf (blank_Cond = False And blank_SCAN = False) Then ''If both are not blank
'            ResultFlag(Site) = 0    ''Pass Blank check
'            PinName = "Pass"
'            SiteVarValue = 2
'        Else ''only one is not blank (ex:blank_Cond=True,blank_SCAN=False)
'            ResultFlag(Site) = 1    ''Fail Blank check
'            PinName = "Fail"
'            SiteVarValue = 0
'        End If
'
'        '====================================================
'        '=  Print the all eFuse Bit data from DSSC          =
'        '====================================================
'        TheExec.Datalog.WriteComment ""
'        TheExec.Datalog.WriteComment "Read All Config eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'        Call auto_PrintAllBitbyDSSC(SingleBitArray(), EConfigReadCycle, EConfigTotalBitCount, EConfigReadBitWidth)
'
''        If (SiteVarValue = 1) Then
''            ''''20161108 update
''            ''''SiteVarValue = 1, do the pre-decode for CRC Pgm if needed
''            Call auto_Decode_CfgBinary_Data(DoubleBitArray, False)
''        End If
'
'        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'
'        If (False) Then
'            PrintSiteVarResult = "Site (" + CStr(Site) + ") " + m_siteVar + " = " + CStr(SiteVarValue)
'            TheExec.Datalog.WriteComment PrintSiteVarResult & vbCrLf
'        End If
'
'        gB_CFG_decode_flag(Site) = False ''''20160531, <MUST> reset and init
'
'        If (SiteVarValue = 0) Then
'            TheExec.Datalog.WriteComment "Site(" & Site & ") blank_Cond=" + CStr(blank_Cond) + ", blank_SCAN=" + CStr(blank_SCAN)
'        End If
'        ''Binning out
'        TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'        TheExec.Flow.TestLimit resultVal:=ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName, PinName:=PinName
'
'        If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
'    Next Site

'CFGChk_byStage_Early:
    'Call UpdateDLogColumns__False
    'DebugPrintFunc ReadPatSet.Value
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function



Public Function auto_mapping_fusing_BKM(FuseType As String, m_catename As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_mapping_fusing_BKM"
    
    Dim BKM_LotID_Temp As String
    Dim BKM_WaferID_Temp As String
    Dim BKM_Lot_Wafer_ID_Temp As String
    Dim BKM_Path As String
    Dim BKM_File_Name As String
    Dim BKM_ver As String
    Dim BKM_Par_Temp() As String
    Dim BkM_ver_temp As String
    Dim BKM_Decode As Long

    
BKM_LotID_Temp = TheExec.Datalog.Setup.LotSetup.LotID
        
        BKM_WaferID_Temp = TheExec.Datalog.Setup.WaferSetup.ID
        BKM_WaferID_Temp = Format(BKM_WaferID_Temp, "00")
        BKM_Lot_Wafer_ID_Temp = BKM_LotID_Temp & "-" & BKM_WaferID_Temp
        'BKM_Lot_Wafer_ID = "N800P0-02"
        'If UCase(BKM_Lot_Wafer_ID_Temp) <> UCase(gS_BKM_Lot_Wafer_ID) Then
             BKM_Path = "X:\BKM\" & BKM_LotID_Temp & "\" & BKM_Lot_Wafer_ID_Temp & "*" & ".txt"
             BKM_File_Name = Dir(BKM_Path)
             
            ''20191221
            If (Dir() <> "") Then   '''if more than one matched files
                TheExec.Datalog.WriteComment ""
                TheExec.Datalog.WriteComment "!!!!   More than one matched files    !!!!"
                TheExec.Datalog.WriteComment ""
                TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="More than one matched files"
                GoTo the_end
                
             ElseIf BKM_File_Name = "" Then
                TheExec.Datalog.WriteComment ""
                TheExec.Datalog.WriteComment "!!!!   Can not Find BKM File in the path " & BKM_Path & "    !!!!"
                TheExec.Datalog.WriteComment ""
                TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="File_Not_Found"
''                For Each site In TheExec.sites
''                    TheExec.sites.Item(site).FlagState("F_BKM_Fail") = logicTrue
''                Next site
                
                gS_BKM_Number = gS_BKM_Unknown
                gS_efuse_BKM_Ver = Dic_BKM(gS_BKM_Unknown)
                gS_BKM_Lot_Wafer_ID = BKM_Lot_Wafer_ID_Temp
                GoTo the_end
             Else
                 
                BKM_Par_Temp = Split(BKM_File_Name, "_")
                BkM_ver_temp = Replace(BKM_Par_Temp(1), ".txt", "")
                
                If (Dic_BKM.Exists(BkM_ver_temp) = False) Then
                    gS_BKM_Number = gS_BKM_Unknown
                    gS_efuse_BKM_Ver = Dic_BKM(gS_BKM_Unknown)
                    gS_BKM_Lot_Wafer_ID = BKM_Lot_Wafer_ID_Temp
                    TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="BKM_Number_Not_Match"
                    TheExec.Datalog.WriteComment ""
                    TheExec.Datalog.WriteComment "The related BKM file " + BKM_File_Name + " does not match with " + gS_BKM_Name
                    TheExec.Datalog.WriteComment ""
                    GoTo the_end
                Else
                    BKM_ver = BkM_ver_temp
                    gS_BKM_Number = BKM_ver
                    gS_efuse_BKM_Ver = auto_BKM2Fuse_Mapping(CStr(BKM_ver))
                    gS_BKM_Lot_Wafer_ID = BKM_Lot_Wafer_ID_Temp
                End If
             End If
         
       'End If
       
       BKM_Decode = CLng("&H" + Replace(auto_BinStr2HexStr(gS_efuse_BKM_Ver, 1), "X", ""))
       
       TheExec.Datalog.WriteComment ""
       TheExec.Datalog.WriteComment " BKM Number is    " & gS_BKM_Number & "    Mapping to Fuse is    " & CStr(gS_efuse_BKM_Ver)
       TheExec.Datalog.WriteComment ""
       TheExec.Flow.TestLimit resultVal:=BKM_Decode, lowVal:=0, hiVal:=CFGFuse.Category(CFGIndex("bkm_package")).HiLMT, Tname:="BKM_Value", ForceResults:=tlForceNone
       TheExec.Datalog.WriteComment ""
       
       For Each site In TheExec.sites
         'Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, BKM_Decode, True)
         gS_BKM_IEDA(site) = CStr(BKM_Decode)
         gS_BKM_Fuse_IEDA(site) = CStr(BKM_Decode)
       Next site
       
    'Call auto_eFuse_SetPatTestPass_Flag_SiteAware(FuseType, m_catename, Pass_Fail_Flag, True)
        Dim m_value As New SiteLong
        m_value = BKM_Decode
        Call auto_eFuse_SetWriteVariable_SiteAware(FuseType, m_catename, m_value, True)
    
the_end:
    
Exit Function
    
    
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_getting_fusing_BKM(FuseType As String, m_catename As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_getting_fusing_BKM"
    Dim fuse_value_All As New SiteDouble
    Dim Fuse_Value As Double
    
    For Each site In TheExec.sites
        Fuse_Value = CFGFuse.Category(CFGIndex(m_catename)).Read.Decimal(site)
        fuse_value_All(site) = Fuse_Value
        gS_BKM_IEDA(site) = Fuse_Value
        gS_BKM_Fuse_IEDA(site) = Fuse_Value
    Next site
    TheExec.Flow.TestLimit fuse_value_All, 0, 15, Tname:="BKM_Group_Index"
    
Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_ConfigWrite_CFG_DV(CFG_DV_pat As Pattern, PwrPin As String, vpwr As Double, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CFG_DV"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim patt As String
    If (auto_eFuse_PatSetToPat_Validation(CFG_DV_pat, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    ''TheHdw.Patterns(CFG_DV_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    Call TurnOnEfusePwrPins(PwrPin, vpwr)

    Call TheHdw.Patterns(patt).Test(pfAlways, 0)
    DebugPrintFunc CFG_DV_pat.Value

    Call TurnOffEfusePwrPins(PwrPin, vpwr)

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Parsing_BKM_Info(BKM_Name As String) As Long

On Error GoTo errHandler

Dim funcName As String:: funcName = "Parsing_BKM_Info"
Dim wb As Workbook
Dim ws As Worksheet
Dim i As Long
Dim j As Long
Dim BKM_Nmae As String
Dim BKM_Decode As String

Set wb = Application.ActiveWorkbook
Set ws = wb.Sheets(BKM_Name)
ws.Activate
 
For i = 1 To ws.UsedRange.Rows.Count
    If (UCase(ws.Cells(i, "A")) = UCase("Decimal")) Then Exit For
Next


For j = i + 1 To ws.UsedRange.Rows.Count
    If (ws.Cells(j, "C") <> "") Then
    
       If (UCase(ws.Cells(j, "C")) Like UCase("*unknown*")) Then
           gS_BKM_Unknown = UCase(ws.Cells(j, "C"))
           BKM_Decode = Trim(ws.Cells(j, "B"))
           If Dic_BKM.Exists(UCase(gS_BKM_Unknown)) = False Then Call Dic_BKM.Add(gS_BKM_Unknown, BKM_Decode)
       Else
           BKM_Nmae = UCase(ws.Cells(j, "C"))
           BKM_Decode = Trim(ws.Cells(j, "B"))
           If Dic_BKM.Exists(BKM_Nmae) = False Then Call Dic_BKM.Add(BKM_Nmae, BKM_Decode)
       End If
    
    End If
Next

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IDS_LIMIT(ids_category_name As String) As Long

On Error GoTo errHandler

Dim funcName As String:: funcName = "IDS_Limit"
Dim Temp As New SiteDouble
Dim IDS As New SiteDouble
Dim instance_name As String:: instance_name = TheExec.DataManager.instanceName + "_" + ids_category_name

Temp = CFGFuse.Category(CFGIndex(ids_category_name)).Read.Decimal
IDS = Temp.Multiply(CFGFuse.Category(CFGIndex(ids_category_name)).Resoultion)

TheExec.Flow.TestLimit resultVal:=IDS, lowVal:=0, hiVal:=1000, Unit:=unitCustom, customUnit:="mA", PinName:=ids_category_name, formatStr:="%f", scaletype:=scaleNoScaling, ForceResults:=tlForceNone

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
