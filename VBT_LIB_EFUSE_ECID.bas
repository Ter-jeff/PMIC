Attribute VB_Name = "VBT_LIB_EFUSE_ECID"

Option Explicit

Public Function auto_eFuse_Initialize()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Initialize"

    Dim m_Site As Variant
    Dim m_len As Long
    Dim m_deflen As Long
    Dim m_tmpStr As String

    gS_JobName = LCase(TheExec.CurrentJob)
    gB_ReadWaferData_flag = False

    ''''--------------------------------------------------------------------------
    ''''201812XX update
    ''''--------------------------------------------------------------------------
    gB_eFuse_newMethod = True      ''''Using DSP_Method:: True, False
    gB_eFuse_printBitMap = True   ''''print Bit Map
    gB_eFuse_printPgmCate = True   ''''print Pgm  Category
    gB_eFuse_printReadCate = True  ''''print Read Category
    
    gB_EFUSE_DVRV_ENABLE = False
    If (TheExec.EnableWord("CFG_Partial") = True) Then gB_EFUSE_DVRV_ENABLE = True
    ''''--------------------------------------------------------------------------
    ''''DSP Mode Possibility
    ''''TRUE::  tlDSPModeForceAutomatic or tlDSPModeAutomatic
    ''''FALSE:: tlDSPModeHostThread     or tlDSPModeHostDebug
    gB_eFuse_DSPMode = True
    ''''--------------------------------------------------------------------------
    '''' set the simulation to blank = True or False
    '''' 0: all sites blank = True
    '''' 1: all sites blank = False
    '''' 2: blank = True or False by site
    gL_eFuse_Sim_Blank = 0 ''''0,1,2
    ''''--------------------------------------------------------------------------
    
    ''''201812XX update
    If (gL_1st_FuseSheetRead = 1) Then
        Call auto_GetSiteFlagName(0, gS_cfgFlagname, False)
        If (gS_cfgFlagname_pre <> gS_cfgFlagname) Then gL_1st_FuseSheetRead = 0
    End If

    gB_CFG_SVM = TheExec.Flow.EnableWord("CFG_SVM")
    gB_CFGSVM_A00_CP1 = TheExec.Flow.EnableWord("CFG_SVM_A00_CP1") ''''20161108 add
    ''''20161108 update
    If (gB_CFGSVM_A00_CP1 = True) Then
        ''''In function auto_CFGConstant_Initialize(), it will trun on the "A00" Flag.
        ''''[NOTICE] In the "Flow_Table_Main_Init_EnableWd", it also do the same thing as the following.
        TheExec.Flow.EnableWord("CFG_SVM") = True
        TheExec.Flow.EnableWord("CFG_A00") = True
    End If

    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment funcName + ":: EnableWord CFG_SVM         = " + CStr(gB_CFG_SVM)
    TheExec.Datalog.WriteComment funcName + ":: EnableWord CFG_SVM_A00_CP1 = " + CStr(gB_CFGSVM_A00_CP1)
  
    ''''<Important>
    ''''20160217 New for CP2 samples back to CP1 retest, user MUST enable the word "eFuse_Disable_ChkLMT"
    gB_eFuse_Disable_ChkLMT_Flag = TheExec.Flow.EnableWord("eFuse_Disable_ChkLMT")
    TheExec.Datalog.WriteComment funcName + ":: EnableWord eFuse_Disable_ChkLMT = " + CStr(gB_eFuse_Disable_ChkLMT_Flag)
    
    gB_eFuse_CFG_Cond_FTF_done_Flag = TheExec.Flow.EnableWord("eFuse_CFG_Cond_FTF_done")
    TheExec.Datalog.WriteComment funcName + ":: EnableWord eFuse_CFG_Cond_FTF_done = " + CStr(gB_eFuse_CFG_Cond_FTF_done_Flag)
    TheExec.Datalog.WriteComment ""

    If (gL_1st_FuseSheetRead = 0) Then
        gS_BKM_Lot_Wafer_ID = ""
        gS_BKM_Number = "-1"
        Call Parsing_BKM_Info(gS_BKM_Name)
        
        gB_findCFGTable_flag = False
        gB_findCFGCondTable_flag = False


        ''''-----------------------------------------------------------------------
        '''' 20151217 New, eFuse Revision CTRL sheet (Miner case)
        ''''-----------------------------------------------------------------------
        ''If (False) Then
        ''    Call auto_eFuseRevCtrl_initialize(True) ''''True is to show the datalog
        ''End If
        ''''-----------------------------------------------------------------------
        
        Call parse_eFuse_ChkList_New(gS_eFuse_sheetName) ''''New eFuse Bit Definition format
        gS_pre_eFuse_sheetName = gS_eFuse_sheetName
        
        ''''20170220, to prevent the case without Config Table Sheet
        If (gB_findCFG_flag) Then
            If (gB_CFG_SVM = True) Then
               Call parse_CFG_Condition_Table_Sheet(gS_cfgTable_SVM_sheetName)
                gS_pre_cfgtable_sheetName = gS_cfgTable_SVM_sheetName
            Else
               Call parse_CFG_Condition_Table_Sheet(gS_cfgTable_sheetName)
                gS_pre_cfgtable_sheetName = gS_cfgTable_sheetName
            End If
        End If

        Call auto_eFuseCategoryResult_Initialize
    
        If (gB_findECID_flag) Then Call auto_ECIDConstant_Initialize
        If (gB_findCFG_flag) Then Call auto_CFGConstant_Initialize
        If (gB_findUID_flag) Then Call auto_UIDConstant_Initialize
        If (gB_findUDR_flag) Then Call auto_UDRConstant_Initialize
        If (gB_findSEN_flag) Then Call auto_SENConstant_Initialize
        If (gB_findMON_flag) Then Call auto_MONConstant_Initialize
        If (gB_findCMP_flag) Then Call auto_CMPConstant_Initialize
        If (gB_findUDRE_flag) Then Call auto_UDRE_Constant_Initialize
        If (gB_findUDRP_flag) Then Call auto_UDRP_Constant_Initialize
        If (gB_findCMPE_flag) Then Call auto_CMPE_Constant_Initialize
        If (gB_findCMPP_flag) Then Call auto_CMPP_Constant_Initialize
    
        ''''--------------------------------------------------------------------------------
        m_len = 0
        m_deflen = 35
        ''''20150710 update
        If (gI_ECID_catename_maxLen < m_deflen) Then gI_ECID_catename_maxLen = m_deflen
        If (gI_CFG_catename_maxLen < m_deflen) Then gI_CFG_catename_maxLen = m_deflen
        If (gI_UID_catename_maxLen < m_deflen) Then gI_UID_catename_maxLen = m_deflen
        If (gI_UDR_catename_maxLen < m_deflen) Then gI_UDR_catename_maxLen = m_deflen
        If (gI_SEN_catename_maxLen < m_deflen) Then gI_SEN_catename_maxLen = m_deflen
        If (gI_MON_catename_maxLen < m_deflen) Then gI_MON_catename_maxLen = m_deflen
        If (gI_CMP_catename_maxLen < m_deflen) Then gI_CMP_catename_maxLen = m_deflen
        If (gI_UDRE_catename_maxLen < m_deflen) Then gI_UDRE_catename_maxLen = m_deflen
        If (gI_UDRP_catename_maxLen < m_deflen) Then gI_UDRP_catename_maxLen = m_deflen
        If (gI_CMPE_catename_maxLen < m_deflen) Then gI_CMPE_catename_maxLen = m_deflen
        If (gI_CMPP_catename_maxLen < m_deflen) Then gI_CMPP_catename_maxLen = m_deflen

        If (gI_ECID_catename_maxLen > m_len) Then m_len = gI_ECID_catename_maxLen
        If (gI_CFG_catename_maxLen > m_len) Then m_len = gI_CFG_catename_maxLen
        If (gI_UID_catename_maxLen > m_len) Then m_len = gI_UID_catename_maxLen
        If (gI_UDR_catename_maxLen > m_len) Then m_len = gI_UDR_catename_maxLen
        If (gI_SEN_catename_maxLen > m_len) Then m_len = gI_SEN_catename_maxLen
        If (gI_MON_catename_maxLen > m_len) Then m_len = gI_MON_catename_maxLen
        If (gI_CMP_catename_maxLen > m_len) Then m_len = gI_CMP_catename_maxLen
        If (gI_UDRE_catename_maxLen > m_len) Then m_len = gI_UDRE_catename_maxLen
        If (gI_UDRP_catename_maxLen > m_len) Then m_len = gI_UDRP_catename_maxLen
        If (gI_CMPE_catename_maxLen > m_len) Then m_len = gI_CMPE_catename_maxLen
        If (gI_CMPP_catename_maxLen > m_len) Then m_len = gI_CMPP_catename_maxLen
        If (m_len < m_deflen) Then m_len = m_deflen
        gL_eFuse_catename_maxLen = m_len
        ''''--------------------------------------------------------------------------------
    
        ''''20171115 update, to check if UDRE/UDRP have the same base voltage
        If (gB_findUDRE_flag And gB_findUDRP_flag) Then
            If (gD_UDRE_BaseVoltage <> gD_UDRP_BaseVoltage) Then
                m_tmpStr = "gD_UDRE_BaseVoltage(" + CStr(gD_UDRE_BaseVoltage) + ") <> gD_UDRP_BaseVoltage(" + CStr(gD_UDRP_BaseVoltage) + ")"
                ''MsgBox (m_tmpStr)
                m_tmpStr = "<WARNING> " + funcName + ":: " + m_tmpStr
                TheExec.AddOutput m_tmpStr
                TheExec.Datalog.WriteComment m_tmpStr
            Else
                gD_BaseVoltage = gD_UDRE_BaseVoltage
            End If
            
            If (gD_UDRE_BaseStepVoltage <> gD_UDRP_BaseStepVoltage) Then
                m_tmpStr = "gD_UDRE_BaseStepVoltage(" + CStr(gD_UDRE_BaseStepVoltage) + ") <> gD_UDRP_BaseStepVoltage(" + CStr(gD_UDRP_BaseStepVoltage) + ")"
                ''MsgBox (m_tmpStr)
                m_tmpStr = "<WARNING> " + funcName + ":: " + m_tmpStr
                TheExec.AddOutput m_tmpStr
                TheExec.Datalog.WriteComment m_tmpStr
            Else
                gD_BaseStepVoltage = gD_UDRE_BaseStepVoltage
            End If
            
            If (gD_UDRE_VBaseFuse <> gD_UDRP_VBaseFuse) Then
                m_tmpStr = "gD_UDRE_VBaseFuse(" + CStr(gD_UDRE_VBaseFuse) + ") <> gD_UDRP_VBaseFuse(" + CStr(gD_UDRP_VBaseFuse) + ")"
                ''MsgBox (m_tmpStr)
                m_tmpStr = "<WARNING> " + funcName + ":: " + m_tmpStr
                TheExec.AddOutput m_tmpStr
                TheExec.Datalog.WriteComment m_tmpStr
            Else
                gD_VBaseFuse = gD_UDRE_VBaseFuse
            End If
        End If
        ''''--------------------------------------------------------------------------------
        ''''<MUST be here after the above> because gD_BaseVoltage should be confirmed first
        Call auto_precheck_SafeVoltage_Base_VddBin ''''20160608 Add, 20180725 update
        Call auto_eFuse_copyLMTtoLMT_R ''''20160617 update
        ''''--------------------------------------------------------------------------------

        ''''20180711 add, 201811XX,
        ''''<Importance and MUST> after auto_precheck_SafeVoltage_Base_VddBin()
        Call auto_eFuse_glbVar_Init
    
        ''''-------------------------------------------------------------------
        ''''20170220 New, it can be the optional by the customer.
        ''''only do once, just try to instead the original method, save TTR
        ''m_deflen = 128
        ''Call auto_UpdateDLogColumns_New(m_deflen)
        ''''-------------------------------------------------------------------
        Call Protect_eFuse_Sheet ''''MUST be shere
        
    End If ''''end of If (gL_1st_FuseSheetRead = 0)

    ''''-------------------------------------------------------------------------
    TheExec.Datalog.WriteComment funcName + "::" + FormatNumeric("eFuse_BitDef_sheetName", 25) + " = " + gS_pre_eFuse_sheetName
    If (gB_findCFG_flag And (gB_findCFGTable_flag Or gB_findCFGCondTable_flag)) Then
        TheExec.Datalog.WriteComment funcName + "::" + FormatNumeric("CFG_Table_sheetName", 25) + " = " + gS_pre_cfgtable_sheetName
    End If
    If (UCase(gS_EFuse_Orientation) = UCase("SingleUp")) Then
        TheExec.Datalog.WriteComment funcName + "::" + FormatNumeric("eFuse_Orientation", 25) + " = " + gS_EFuse_Orientation + " (1-Bit)"
    ElseIf (UCase(gS_EFuse_Orientation) = "RIGHT2LEFT") Then
        TheExec.Datalog.WriteComment funcName + "::" + FormatNumeric("eFuse_Orientation", 25) + " = " + gS_EFuse_Orientation + " (2-Bit)"
    Else
        TheExec.Datalog.WriteComment funcName + "::" + FormatNumeric("eFuse_Orientation", 25) + " = " + gS_EFuse_Orientation
    End If
    TheExec.Datalog.WriteComment funcName + "::" + FormatNumeric("eFuse_DigCap_BitOrder", 25) + " = " + gC_eFuse_DigCap_BitOrder
    ''''-------------------------------------------------------------------------
    
    ''''20180711 add, run everytime
    Call auto_eFuse_onProgramStarted
    gStr_PatName = ""

    If (TheExec.EnableWord("Pgm2File")) Then Call auto_CreateConstant

    TheExec.Flow.TestLimit gL_1st_FuseSheetRead, 0, 1, Tname:="eFuseInit_" + UCase(gS_JobName)
    
    ''''debug purpose if True
    'If (False) Then Call auto_eFuse_ReadCategory
    
    gL_1st_FuseSheetRead = 1 ''''201808XX, <Just put on the last line>
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'''''20160125, it's used to check if needs to run EVS flow
'''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
'Public Function auto_ECID_ALLBlankChk(ReadPatSet As Pattern, PinRead As PinList, _
'                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
'                    Optional Validating_ As Boolean)
'    If FunctionList.Exists("auto_ECID_ALLBlankChk") = False Then FunctionList.Add "auto_ECID_ALLBlankChk", ""
'
'On Error GoTo errHandler
'    Dim funcName As String:: funcName = "auto_ECID_ALLBlankChk"
'
'    ''''----------------------------------------------------------------------------------------------------
'    ''''<Important>
'    ''''Must be put before all implicit array variables, otherwise the validation will be error.
'    '==================================
'    '=  Validate/Load Read patterns   =
'    '==================================
'    ''''20161114 update to Validate/load pattern
'    Dim ReadPatt As String
'    If (auto_eFuse_PatSetToPat_Validation(ReadPatSet, ReadPatt, Validating_) = True) Then Exit Function
'    ''''----------------------------------------------------------------------------------------------------
'
'    Dim Site As Variant
'    Dim SiteVarValue As Long, PrintSiteVarResult As String
'    Dim SignalCap As String, CapWave As New DSPWave
'    Dim SingleStrArray() As String
'    Dim allblank As New SiteBoolean
'
'    SignalCap = "SignalCapture" 'define capture signal name
'
'    '=============================================================================
'    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
'    '=============================================================================
'    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
'
'    '****** Initialize Site Varaible ******
'    Dim m_siteVar As String
'    m_siteVar = "ECIDChk_Var"
'    For Each Site In TheExec.Sites.Existing
'        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = -1
'    Next Site
'
'    '================================================
'    '=  Setup HRAM/DSSC capture cycles              =
'    '================================================
'    '*** Setup HARM/DSSC Trigger and Capture parameter ***
'    If (True) Then
'        ''''''''For the Pattern with Parallel 32bits Capture Case
'        auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, EcidReadCycle, CapWave  'setup
'        Call TheHdw.Patterns(ReadPatt).test(pfAlways, 0, tlResultModeDomain)   'run read pattern and capture
'        auto_eFuse_DSSC_ReadDigCap_32bits EcidReadCycle, PinRead.Value, SingleStrArray, CapWave, allblank 'read back to singlestrarray
'    Else
'        ''''20170126 update (20180122 update from Tosp as Example)
'        ''''''''For the Pattern with Serial 1bit Capture Case
'        auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, (EcidReadCycle * 32), CapWave 'setup, 16x32=512 bits
'        Call TheHdw.Patterns(ReadPatt).test(pfAlways, 0, tlResultModeDomain)   'run read pattern and capture
'        auto_eFuse_DSSC_ReadDigCap_32bits EcidReadCycle, PinRead.Value, SingleStrArray, CapWave, allblank, True, "bit0_bitLast" 'read back to singlestrarray
'    End If
'
'    For Each Site In TheExec.Sites
'
'        ''''20151027 add
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            allblank(Site) = True
'            If (gS_JobName <> "cp1") Then allblank(Site) = False
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        If allblank(Site) = False Then  'If not blank
'            SiteVarValue = 2
'        Else ' True means this ECIDFuse is all balnk.
'            SiteVarValue = 1
'        End If
'
'        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'
'        If (False) Then
'            PrintSiteVarResult = "Site (" + CStr(Site) + ") " + m_siteVar + " = " + CStr(SiteVarValue)
'            TheExec.Datalog.WriteComment PrintSiteVarResult
'        End If
'
'        ''Binning out
'        TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'        If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
'
'    Next Site
'    TheExec.Datalog.WriteComment ""
'
'    Call UpdateDLogColumns__False
'    DebugPrintFunc ReadPatSet.Value
'
'Exit Function
'
'errHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function
'
''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
''''This Function is used on ECID DEID blank check only
Public Function auto_ECIDBlankChk_DEID(ReadPatSet As Pattern, PinRead As PinList, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)
  
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ECIDBlankChk_DEID"

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
    Dim BlankChkPass As Boolean, BlownFuseChkPass As Boolean, FuseBlankChkPass As Boolean
    Dim DblFBCount As Long
    Dim ResultFlag As New SiteLong
    Dim testName As String, PinName As String
    Dim SiteVarValue As Long, PrintSiteVarResult As String
    Dim CurrInstance As String
    Dim SignalCap As String, CapWave As New DSPWave
    Dim SingleStrArray() As String
    Dim m_stage As String
    Dim allBlank As New SiteBoolean
    Dim blank_early As New SiteBoolean
    Dim blank_stage As New SiteBoolean

    Dim singleBitSum As Long
    Dim SingleBitArray() As Long
    ''ReDim SingleBitArray(ECIDTotalBits - 1)
    Dim SingleBitArrayStr() As String
    ReDim gL_ECID_Sim_FuseBits(TheExec.sites.Existing.Count - 1, ECIDTotalBits - 1) ''''20161107 update, it's for the simulation

    SignalCap = "SignalCapture" 'define capture signal name

    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ   'Karl chnage 0416

    '****** Initialize Site Varaible ******
    Dim m_siteVar As String
    m_siteVar = "ECIDChk_Var"
    For Each site In TheExec.sites.Existing
        TheExec.sites(site).SiteVariableValue(m_siteVar) = -1
    Next site

    Dim m_jobinStage_flag As Boolean
    ''''it's used to identify if the Job Name is existed in the UDR portion of the eFuse BitDef table.
    m_jobinStage_flag = auto_eFuse_JobExistInStage("ECID", True) ''''<MUST>

    '================================================
    '=  Setup HRAM/DSSC capture cycles           =
    '================================================
    '*** Setup HARM/DSSC Trigger and Capture parameter ***
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, EcidReadCycle, CapWave  'setup
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)   'run read pattern and capture
    

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''----------------------------------------------------
    ''''201808XX New Method by DSPWave
    ''''----------------------------------------------------
    'Dim mW_SingleBitWave As New DSPWave
    'Dim mW_DoubleBitWave As New DSPWave
    
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong
    
    
    m_Fusetype = eFuse_ECID
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_early = True

''    ''''<MUST> Clear Memory
''    Set gW_ECID_Read_singleBitWave = Nothing
''    Set gW_ECID_Read_doubleBitWave = Nothing

    'gDL_eFuse_Orientation = eFuse_2_Bit
    gDL_eFuse_Orientation = gE_eFuse_Orientation
    If (gS_JobName = "cp1") Then gS_JobName = "cp1_early" ''''<MUST>
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
   'condstr = LCase(condstr)

    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=0 (Stage Early Bits)
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 0, CapWave, m_FBC, blank_early, allBlank)
    ''''----------------------------------------------------

    If (blank_early.Any(False) = True) Then
        ''''if there is any site which is non-blank, then decode to gDW_ECID_Read_Decimal_Cate
        ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
        ''''it will be present by Hex and Binary compare with the limit later on.
        ''Call auto_eFuse_setReadData(eFuse_ECID, gDW_ECID_Read_DoubleBitWave, gDW_ECID_Read_Decimal_Cate)
        'Call auto_eFuse_setReadData(eFuse_ECID)
        
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(eFuse_ECID)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_ECID)
    
''        If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
''        End If
        
        Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_ECID, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_ECID, False, gB_eFuse_printReadCate)
        
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
    m_SiteVarValue = blank_early
    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
    If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only
    For Each site In TheExec.sites
        If (auto_eFuse_GetAllPatTestPass_Flag("ECID") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    gL_ECID_FBC = m_FBC
    
    ''''if blank_early(Site)=False, check if read DEID bits are same as Prober (CP1 only)
    ''''When blank=False, We do NOT need to do that because it will be checked in Syntax check again.

    Call UpdateDLogColumns(gI_ECID_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "FailBitCount_CP1_Early" ''+ UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "ECID_Blank_CP1_Early" '' + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value
    
Exit Function

End If

'    auto_eFuse_DSSC_ReadDigCap_32bits EcidReadCycle, PinRead.Value, gS_SingleStrArray, CapWave, allblank 'read back to singlestrarray
'    ''''set to global gS_SingleStrArray for SingleDoubleBit usage while in Re-probing stage.
'
'    Call UpdateDLogColumns(gI_ECID_catename_maxLen)
'    'blank check
'    SiteVarValue = 0  'initial
'
'    ''''Here using the programming stage of 'Lot_ID' stands for the Stage for all ECID bits.(It Should BE!!!)
'    m_stage = LCase(ECIDFuse.Category(ECIDIndex("Lot_ID")).Stage)
'
'    ''''Need to be checked
'    ''If (gB_eFuse_newMethod = False) Then
'        If (m_stage = "cp1_early") Then m_stage = "cp1"
'        For i = 0 To UBound(ECIDFuse.Category)
'            If (LCase(ECIDFuse.Category(i).Stage) = "cp1_early") Then
'                ECIDFuse.Category(i).Stage = "CP1"
'            End If
'        Next i
'    ''End If
'
'    For Each Site In TheExec.Sites
'        ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To EcidReadCycle - 1
'                gS_SingleStrArray(i, Site) = StrReverse(gS_SingleStrArray(i, Site))
'            Next i
'        End If
'
'        ''''20151027 add, 20161107 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            allblank(Site) = True ''''False for Retest mode
'            If (gS_JobName <> "cp1") Then allblank(Site) = False
'
'            If (allblank(Site) = False) Then ''''sim for re-test
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                TheExec.Datalog.WriteComment vbTab & "[ Simulation for the ReTest Mode (DEID and nonDEID)]"
'                Call eFuseENGFakeValue_Sim
'                ''''-----------------------------------------------------------------------------------------
'                Dim m_tmpStr As String
'                Dim Expand_eFuse_Pgm_Bit() As Long, eFuse_Pgm_Bit() As Long
'                ReDim eFuse_Pgm_Bit(ECIDTotalBits - 1)
'                ReDim Expand_eFuse_Pgm_Bit(ECIDTotalBits * EcidWriteBitExpandWidth - 1)
'                Call auto_EcidPgmBit_DEID_forCheck(Expand_eFuse_Pgm_Bit, eFuse_Pgm_Bit, False) ''''True for Debug Display
'                ''''to composite gS_SingleStrArray(i, Site) for the simulation in SingleDoubleBit_Check/Syntax
'                For i = 0 To EcidReadCycle - 1
'                    m_tmpStr = ""
'                    For j = 0 To EcidReadBitWidth - 1
'                        k = j + i * EcidReadBitWidth ''''MUST
'                        m_tmpStr = CStr(eFuse_Pgm_Bit(k)) + m_tmpStr
'                        gL_ECID_Sim_FuseBits(Site, k) = eFuse_Pgm_Bit(k)
'                    Next j
'                    gS_SingleStrArray(i, Site) = m_tmpStr
'                Next i
'                ''''-----------------------------------------------------------------------------------------
'            Else
'                ''''allBlank = True
'                For i = 0 To ECIDTotalBits - 1
'                    gL_ECID_Sim_FuseBits(Site, i) = 0
'                Next i
'            End If
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        DblFBCount = 0 ''''Must have
'        If allblank(Site) = False Then  'If not blank
'            '*** In here it means the eFuse is not blank ****
'            If gS_JobName Like "cp*" Then
'
'                If (gB_ReadWaferData_flag = True) Then
'                    ''''Read DSSC data and check only DUT ID (LotID,WaferID,XY) and Reserved bits
'                    Call auto_ECID_Memory_Read_DEID(gS_SingleStrArray, BlownFuseChkPass, DblFBCount)
'                Else
'                    ''''In CP stage, it needs to check if the Prober's information is same as the eFuse Read.
'                    BlownFuseChkPass = False ''''set it to fail
'                    TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please run TestInstance <ReadWaferData> first."
'                End If
'
'                testName = "ECID_GPIB"
'                If BlownFuseChkPass = True Then
'                    ResultFlag(Site) = 0 'stand for this test item pass
'                    PinName = "Same"
'                    SiteVarValue = 2
'                Else 'Namely, BlownFuseChkPass = False
'                    ResultFlag(Site) = 1    'Fail Blank check criterion
'                    PinName = "Diff"
'                    SiteVarValue = 0
'                    'Retest Branch Fail
'                End If
'
'            Else  ''''non CP tests, set it to default pass, and will be checked in SingleDoubleBit/Syntax Check item
'                ''''<NOTICE>
'                ''''INFO FT Stage (WLFT) could use 'ReadWaferData', but its XY could be NOT equal to DUT's ECID
'                DblFBCount = 0
'                BlownFuseChkPass = True
'                testName = "ECID_Blank"
'                PinName = "Pass"
'                ResultFlag(Site) = 0 'stand for this test item pass
'                SiteVarValue = 2
'            End If
'
'        Else ' True means this ECIDFuse is all balnk.
'            testName = "ECID_Blank"
'            ''''Here it's used to check if HardIP pattern test pass or not.
'            If (auto_eFuse_GetAllPatTestPass_Flag("ECID") = False) Then
'                ResultFlag(Site) = 1  'Fail Blank check criterion
'                PinName = "Fail"
'                SiteVarValue = 0
'            Else
'                ''If CP, it implies this is a first-time test, eFuse certainly is blank
'                If gS_JobName = m_stage Then
'                    ResultFlag(Site) = 0   'Pass Blank check criterion
'                    testName = "ECID_Blank"
'                    PinName = "Pass"
'                    SiteVarValue = 1
'                Else 'Should not be Blank at non-CP1(m_stage) flow
'                    ResultFlag(Site) = 1    'Fail Blank check criterion
'                    PinName = "Fail"
'                    SiteVarValue = 0
'                End If
'            End If
'        End If
'
'        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'
'        singleBitSum = 0
'        Call auto_Decompose_StrArray_to_BitArray("ECID", gS_SingleStrArray, SingleBitArray, singleBitSum)
'        If (gS_ECID_CRC_PgmFlow <> "") Then
'            ReDim SingleBitArrayStr(UBound(SingleBitArray)) ''''<MUST> 20161114 update
'            For i = 0 To UBound(SingleBitArray)
'                SingleBitArrayStr(i) = CStr(SingleBitArray(i))
'            Next i
'            ''''20170815 update for CRC calc
'            Call auto_EcidPrintData(1, SingleBitArrayStr, False, False)
'        End If
'
'        ''''20151228 update
'        If (True) Then ''''set True for the Engineer Debug
'            ''''Decompose Read Cycle String to get SingleBitArray and SingleBitSum
'            ''SingleBitSum = 0
'            ''Call auto_Decompose_StrArray_to_BitArray("ECID", gS_SingleStrArray, SingleBitArray, SingleBitSum)
'            TheExec.Datalog.WriteComment ""
'            TheExec.Datalog.WriteComment "Read All ECID eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'            Call auto_PrintAllBitbyDSSC(SingleBitArray(), EcidReadCycle, ECIDTotalBits, EcidReadBitWidth)
'        End If
'
'        If (False) Then
'            PrintSiteVarResult = "Site (" + CStr(Site) + ") " + m_siteVar + " = " + CStr(SiteVarValue)
'            TheExec.Datalog.WriteComment PrintSiteVarResult & vbCrLf
'        End If
'
'        gB_ECID_decode_flag(Site) = False ''''20160531, <MUST> reset and init
'
'        ''Binning out
'        TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'        TheExec.Flow.TestLimit resultVal:=DblFBCount, lowVal:=0, hiVal:=0, Tname:="FailBitCount", PinName:="Value"
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

'''' nonDEID=non DeviceID
'''' This function is used for the blank check for the nonDEID part.
''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_ECIDBlankChk(ReadPatSet As Pattern, PinRead As PinList, _
                                  Optional condstr As String = "stage", _
                                  Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                                  Optional Validating_ As Boolean)
  
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ECIDBlankChk"

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
    Dim DblFBCount As Long
    Dim ResultFlag As New SiteLong
    Dim testName As String, PinName As String
    Dim SiteVarValue As Long, PrintSiteVarResult As String
    Dim SignalCap As String, CapWave As New DSPWave
    Dim allBlank As New SiteBoolean
    Dim blank_nonDEID As New SiteBoolean
    Dim m_stage As String
    Dim m_jobinStage_flag As Boolean
    Dim singleBitSum As Long
    Dim SingleBitArray() As Long
    ''ReDim SingleBitArray(ECIDTotalBits - 1)
    Dim m_tmpStr As String
    Dim SingleBitArrayStr() As String

    ''''20161107 update
    ''''''''<Important> Can NOT do the below declaration, otherwise the simulation bits (DEID) will be clear in this nonDEID test.
    ''''''''<MUST> using Preserve to reserve the previous simulation data here.
    ReDim Preserve gL_ECID_Sim_FuseBits(TheExec.sites.Existing.Count - 1, ECIDTotalBits - 1) ''''20161107 update, it's for the simulation

    SignalCap = "SignalCapture" 'define capture signal name

    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ   'Karl chnage 0416

    '****** Initialize Site Varaible ******
    Dim m_siteVar As String
    m_siteVar = "ECIDChk_Var"
    For Each site In TheExec.sites.Existing
        TheExec.sites(site).SiteVariableValue(m_siteVar) = -1
    Next site

    ''''it's used to identify if the Job Name is existed in the UDR portion of the eFuse BitDef table.
    m_jobinStage_flag = auto_eFuse_JobExistInStage("ECID", True) ''''<MUST>
    
    '================================================
    '=  Setup HRAM/DSSC capture cycles              =
    '================================================
    '*** Setup HARM/DSSC Trigger and Capture parameter ***
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, EcidReadCycle, CapWave  'setup
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)    'run read pattern and capture

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''----------------------------------------------------
    ''''201808XX New Method by DSPWave
    ''''----------------------------------------------------
    'Dim mW_SingleBitWave As New DSPWave
    'Dim mW_DoubleBitWave As New DSPWave
    
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong
    Dim blank_early As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    Dim m_bitFlag_mode As Long
    
    m_Fusetype = eFuse_ECID
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_early = True
    
    If (condstr = "cp1_early") Then
        m_bitFlag_mode = 0
    ElseIf (condstr = "stage") Then
        m_bitFlag_mode = 1
    ElseIf (condstr = "all") Then
        m_bitFlag_mode = 2
    Else
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
        m_FBC = -1
    End If

''    ''''<MUST> Clear Memory
''    Set gW_ECID_Read_singleBitWave = Nothing
''    Set gW_ECID_Read_doubleBitWave = Nothing

    'If (gS_JobName = "cp1") Then gS_JobName = "cp1_early" ''''<MUST>
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=0 (Stage Early Bits)
    'If (TheExec.TesterMode = testModeOffline) Then gL_eFuse_Sim_Blank = 1
    
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, blank_early, allBlank)
    ''''----------------------------------------------------

    If (blank_early.Any(False) = True) Then
        ''''if there is any site which is non-blank, then decode to gDW_ECID_Read_Decimal_Cate
        ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
        ''''it will be present by Hex and Binary compare with the limit later on.
        ''Call auto_eFuse_setReadData(eFuse_ECID, gDW_ECID_Read_DoubleBitWave, gDW_ECID_Read_Decimal_Cate)
        'Call auto_eFuse_setReadData(eFuse_ECID)
        
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(eFuse_ECID)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_ECID)
    
''        If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
''        End If
        
        Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_ECID, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_ECID, False, gB_eFuse_printReadCate)
        
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
    m_SiteVarValue = blank_early
    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
    If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only
    For Each site In TheExec.sites
        If (auto_eFuse_GetAllPatTestPass_Flag("ECID") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    gL_ECID_FBC = m_FBC
    
    ''''if blank_early(Site)=False, check if read DEID bits are same as Prober (CP1 only)
    ''''When blank=False, We do NOT need to do that because it will be checked in Syntax check again.

    Call UpdateDLogColumns(gI_ECID_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "FailBitCount_CP1" ''+ UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "ECID_Blank_CP1" '' + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value
    
Exit Function

End If


errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function



''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_EcidWrite_byCondition(WritePattSet As Pattern, PinWrite As PinList, _
                    PwrPin As String, vpwr As Double, _
                    condstr As String, _
                    Optional catename_grp As String, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)
  
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidWrite_byCondition"

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
    Dim Expand_eFuse_Pgm_Bit() As Long
    Dim eFuse_Pgm_Bit() As Long
    Dim eFusePatCompare() As String
    Dim i As Long, j As Long
    Dim SegmentSize As Long
    Dim DigSrcSignalName As String
    Dim Expand_Size As Long
    
    '====================================
    '=  eFuse R/W Flow Description      = (2011/10 by Jack)
    '====================================
    'Step1. Fetch the Lot ID, Wafer ID, X,Y coordinates from prober
    'Step2. Wrap up data for Efuse programming through DSSC approach
    'Step3. Program eFuse
    'Step4. Read back eFuse content to make sure programming success
    
    '===================================================
    '=  Step1.Setup Power Supply and IO pin voltages   =
    '====================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ  'SEC DRAM

    TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName
   
    condstr = LCase(condstr) ''''<MUST>

    '========================================================
    '=  Step2. Define Array sizes for DSSC source   =
    '========================================================
    Expand_Size = ECIDTotalBits * EcidWriteBitExpandWidth  'Because there are repeat cycle in C651 pattern, we have to create multiple DSSC
    'However,EcidWriteBitExpandWidth depends on how many repeat cycle in pin of STROBE.
    'This subroutine is designed to (1). pack programming bit for DSSC (Expand_eFuse_Pgm_Bit())
    '                               (2). produce a output comparing data (eFusePatCompare())
    
    '================================================
    '=  Modifying Pattern for Read Pattern Compare  =
    '================================================
       
    ReDim gL_ECIDFuse_Pgm_Bit(TheExec.sites.Existing.Count - 1, ECIDTotalBits - 1)
    
    DigSrcSignalName = "DigSrcSignal"

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''201808XX update
    If (TheExec.TesterMode = testModeOffline) Then
        If (condstr <> LCase("DEID")) Then
            For Each site In TheExec.sites
                Call eFuseENGFakeValue_Sim
            Next site
        End If
    End If

    Dim m_stage As String
    Dim m_cmpStage As String
    Dim m_pgmRes As New SiteLong
    Dim m_defreal As String
    Dim m_algorithm As String
    Dim m_pgmWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_crc_idx As Long
    Dim m_CRCconvertFlag As Boolean:: m_CRCconvertFlag = False

    m_pgmWave.CreateConstant 0, gDL_BitsPerBlock, DspLong ''''gDL_BitsPerBlock = EcidBitPerBlockUsed
    ''''<MUST and Importance>
    For Each site In TheExec.sites
        gDW_Pgm_RawBitWave = m_pgmWave.Copy
    Next site

    If (condstr = LCase("DEID") Or condstr = "cp1_early") Then
        m_cmpStage = "cp1_early"
    Else
        ''''condStr = "nonDEID" or "stage"
        m_cmpStage = gS_JobName
    End If
    
    ''''Only composite case "real or bincut" PgmBits Wave per Stage requirement
    For i = 0 To UBound(ECIDFuse.Category) - 1
        With ECIDFuse.Category(i)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_defreal = LCase(.Default_Real)
        End With
        
        If (m_stage = m_cmpStage) Then
            If (m_algorithm = "crc") Then
                m_crc_idx = i
                ''''special handle on the next process
                ''''skip it here
            ElseIf (m_defreal = "real" Or m_defreal = "bincut") Then
                With ECIDFuse.Category(i)
                    ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
                    Call rundsp.eFuse_updatePgmWave_byCategory(.BitIndexWave, .Write.BitArrWave)
                End With
            End If
        Else
            ''''doNothing
        End If
    Next i
    
    ''''process CRC bits calculation
    If (gS_ECID_CRC_Stage = gS_JobName) Then
        Dim mSL_bitwidth As New SiteLong
        mSL_bitwidth = gL_ECID_CRC_BitWidth
        ''''CRC case
        With ECIDFuse.Category(m_crc_idx)
            ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
            
            'gDL_CRC_EndBit = gL_ECID_CRC_EndBit
            
            Call rundsp.eFuse_updatePgmWave_CRCbits(eFuse_ECID, mSL_bitwidth, .BitIndexWave)
        End With
    End If
    
'    ''''process CRC bits calculation
'    For i = 0 To UBound(ECIDFuse.Category) - 1
'        m_defreal = LCase(ECIDFuse.Category(i).Default_Real)
'        m_stage = LCase(ECIDFuse.Category(i).Stage)
'        If (m_stage = m_cmpStage) Then
'            If (m_defreal = "crc") Then
'                ''''special handle on the this process
'                Exit For ''''<MUST>
'            End If
'        Else
'            ''''doNothing
'        End If
'    Next i
    
    ''''composite effective PgmBits per Stage requirement
    Dim m_pgmDoubleBitWave As New DSPWave
    Dim m_pgmSingleBitWave As New DSPWave
    Dim m_pgmDigSrcWave As New DSPWave
    
    If (m_cmpStage = "cp1_early") Then
        ''''condStr = "DEID" or "cp1_early"
        ''''Here Parameter bitFlag_mode=0 (Stage Early Bits)
        Call rundsp.eFuse_Gen_PgmBitSrcWave(eFuse_ECID, 0, m_pgmDigSrcWave, m_pgmRes)
    Else
        ''''condStr = "nonDEID" or "stage"
        ''''Here Parameter bitFlag_mode=1 (Stage Bits)
        Call rundsp.eFuse_Gen_PgmBitSrcWave(eFuse_ECID, 1, m_pgmDigSrcWave, m_pgmRes)
    End If
    TheHdw.Wait 1# * ms
    TheExec.Flow.TestLimit resultVal:=m_pgmRes, lowVal:=1, hiVal:=1, Tname:="ECID_PGM_" + UCase(gS_JobName)

    ''''use global DSP variable
    Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_ECID_Pgm_SingleBitWave, gDL_TotalBits, gDL_ReadCycles, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_ECID, gDW_ECID_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)
    
    Call TurnOnEfusePwrPins(PwrPin, vpwr)
    
    Call eFuse_DSSC_SetupDigSrcWave(WritePatt, PinWrite, DigSrcSignalName, m_pgmDigSrcWave)
    
    ''''if it's same values on all Sites to save TT and improve PTE
    ''Call eFuse_DSSC_SetupDigSrcWave_allSites(WritePatt, PinWrite, DigSrcSignalName, m_pgmDigSrcWave)

''''    ''''It will be used @ MarginRead process
''''    Call RunDSP.eFuse_DspWave_Copy(m_pgmSingleBitWave, gW_ECID_Pgm_singleBitWave)
''''    Call RunDSP.eFuse_DspWave_Copy(m_pgmDoubleBitWave, gW_ECID_Pgm_doubleBitWave)

    ''''Write Pattern for programming eFuse
    Call TheHdw.Patterns(WritePatt).Test(pfAlways, 0)   'Write ECID
    TheHdw.Wait 100# * us

    Call TurnOffEfusePwrPins(PwrPin, vpwr)
    DebugPrintFunc WritePattSet.Value
    
Exit Function
    
End If

'    For Each Site In TheExec.Sites
'        ''''20161107 update
'        If (TheExec.TesterMode = testModeOffline) Then
'            ''If (condStr <> LCase("DEID")) Then Call eFuseENGFakeValue_Sim
'            Call eFuseENGFakeValue_Sim
'        End If
'
'        ReDim eFuse_Pgm_Bit(ECIDTotalBits - 1)
'        ReDim Expand_eFuse_Pgm_Bit(Expand_Size - 1)
'        ReDim eFusePatCompare(EcidReadCycle - 1)
'
'        ''''20151221 New
'        If (condstr = LCase("DEID")) Then
'            SegmentSize = auto_EcidMakePgmBit_DEID(Expand_eFuse_Pgm_Bit(), eFuse_Pgm_Bit())
'        ElseIf (condstr = LCase("nonDEID")) Then
'            SegmentSize = auto_EcidMakePgmBit_nonDEID(Expand_eFuse_Pgm_Bit(), eFuse_Pgm_Bit())
'        ElseIf (condstr = LCase("DEID_WAT")) Then
'            SegmentSize = auto_EcidMakePgmBit_DEID_DEV_WAT(Expand_eFuse_Pgm_Bit(), eFuse_Pgm_Bit())
'        ElseIf (condstr = "stage") Then
'            SegmentSize = auto_EcidMakePgmBit_byStage(Expand_eFuse_Pgm_Bit(), eFuse_Pgm_Bit())
'        ElseIf (condstr = "category") Then
'            SegmentSize = auto_EcidMakePgmBit_byCategory(catename_grp, Expand_eFuse_Pgm_Bit(), eFuse_Pgm_Bit())
'        Else
'            ''''default=all bits are 0, here it prevents any typo issue
'            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (DEID,DEID_WAT,nonDEID,Stage,Category)"
'            For i = 0 To ECIDTotalBits - 1
'                eFuse_Pgm_Bit(i) = 0
'            Next i
'            For i = 0 To Expand_Size - 1
'                Expand_eFuse_Pgm_Bit(i) = 0
'            Next i
'            SegmentSize = Expand_Size
'        End If
'
'        For i = 0 To ECIDTotalBits - 1
'            gL_ECIDFuse_Pgm_Bit(Site, i) = eFuse_Pgm_Bit(i)
'        Next i
'
'        DSSC_SetupDigSrcWave WritePatt, PinWrite, DigSrcSignalName, SegmentSize, Expand_eFuse_Pgm_Bit
'
'        'Print out All program bits
'        Call auto_PrintAllPgmBits(eFuse_Pgm_Bit, EcidReadCycle, ECIDTotalBits, EcidReadBitWidth)
'
'    Next Site
'    TheHdw.DSSC.Pins(PinWrite).Pattern(WritePatt).Source.Signals.DefaultSignal = DigSrcSignalName
'
'    'Step2. Write Pattern for programming eFuse
'    TheHdw.Wait 0.0001
'    Call TheHdw.Patterns(WritePatt).test(pfAlways, 0)   'Write ECID
'    TheHdw.Wait 0.0001
'
'    Call TurnOffEfusePwrPins(PwrPin, vpwr)
'    DebugPrintFunc WritePattSet.Value
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_ECID_Read_by_OR_2Blocks(ReadPatSet As Pattern, PinRead As PinList, _
                    condstr As String, _
                    Optional catename_grp As String = "", _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)
  
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ECID_Read_by_OR_2Blocks"

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
    Dim SegmentSize As Long
    Dim Expand_Size As Long
    Dim FailCnt As New SiteLong
    Dim testName As String, TestNum() As Long, instanceName As String, Testcnt As Long
    Dim blank As New SiteBoolean

    Dim SingleStrArray() As String
    Dim SingleBitArray() As Long
    Dim SingleBitArrayStr() As String
    Dim DoubleBitArray() As Long

    Dim CapWave As New DSPWave

    condstr = LCase(condstr) ''''<MUST>

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
  
    '====================================================
    '=  1. Define the Bit length of Programming         =
    '=  2. Define the cycle number of Compare Pattern   =
    '====================================================
    Expand_Size = (ECIDTotalBits * EcidWriteBitExpandWidth)  'Because there are repeat cycle in C651 pattern, we have to create multiple DSSC

    '*** Burst read pattern for capturing Q[31:0] ouput data ***
    Call PatSetToPat_EFuse(ReadPatSet, ReadPatt)

    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, "SignalCapture", EcidReadCycle, CapWave

    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)

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

    m_Fusetype = eFuse_ECID
    m_FBC = -1       ''''init to failure
    m_cmpResult = -1 ''''init to failure

    ''''--------------------------------------------------------------------------
    '''' Offline Simulation Start                                                |
    ''''--------------------------------------------------------------------------
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave by using Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        ''''Thus its capWave should be the result (_Pgm_singleBitWave or _read_SingleBitWave of the blank check)
        Call rundsp.eFuse_SingleBitWave2CapWave32Bits(eFuse_ECID, CapWave)
        Call auto_eFuse_print_capWave32Bits(eFuse_ECID, CapWave, False) ''''True to print out
    End If
    ''''--------------------------------------------------------------------------
    '''' Offline Simulation End                                                  |
    ''''--------------------------------------------------------------------------

'    If (condStr = "cp1_early") Then
'        m_bitFlag_mode = 0
'    ElseIf (condStr = "stage") Then
'        m_bitFlag_mode = 1
'    ElseIf (condStr = "all") Then
'        m_bitFlag_mode = 2
'    Else
'        ''''default, here it prevents any typo issue
'        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
'        m_FBC = -1
'        m_cmpResult = -1
'    End If

    If (condstr = "deid" Or condstr = "cp1_early") Then
        m_bitFlag_mode = 0
    ElseIf (condstr = "nondeid" Or condstr = "stage") Then
        m_bitFlag_mode = 1
    ElseIf (condstr = "all") Then
        m_bitFlag_mode = 2
    Else
        ''''default, here it prevents any typo issue
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (DEID,nonDEID,All,Stage)"
        m_FBC = -1
        m_cmpResult = -1
    End If
    
    Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, m_cmpResult)
    
    gL_ECID_FBC = m_FBC

    ''''''[NOTICE] Decode and Print have moved to SingleDoubleBit()

    TheExec.Flow.TestLimit resultVal:=m_cmpResult, lowVal:=0, hiVal:=0

Exit Function

End If

'    auto_eFuse_DSSC_ReadDigCap_32bits EcidReadCycle, PinRead.Value, SingleStrArray, CapWave, blank ''''here must use local SingleStrArray
'
'    Call UpdateDLogColumns(gI_ECID_catename_maxLen)
'
'    For Each Site In TheExec.Sites
'        ''''20151014 add, 20161107 update
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            ''''------------------------------------------------------
'            ''''20161107 update
'            Call auto_Decompose_StrArray_to_BitArray("ECID", gS_SingleStrArray, SingleBitArray, 0)
'            ''''------------------------------------------------------
'
'            ''''to composite gS_SingleStrArray(i, Site) for the simulation in SingleDoubleBit_Check/Syntax
'            For i = 0 To EcidReadCycle - 1
'                m_tmpStr = ""
'                For j = 0 To EcidReadBitWidth - 1
'                    k = j + i * EcidReadBitWidth ''''MUST
'                    If (SingleBitArray(k) = 0 And gL_ECIDFuse_Pgm_Bit(Site, k) = 1) Then
'                        SingleBitArray(k) = gL_ECIDFuse_Pgm_Bit(Site, k)
'                        gL_ECID_Sim_FuseBits(Site, k) = SingleBitArray(k) ''''20161107 update
'                    End If
'                    m_tmpStr = CStr(SingleBitArray(k)) + m_tmpStr
'                Next j
'                gS_SingleStrArray(i, Site) = m_tmpStr
'                SingleStrArray(i, Site) = m_tmpStr
'            Next i
'
'            If (False) Then
'                ''''20161018 try for debug
'                TheExec.Datalog.WriteComment "---testModeOffline---" + funcName
'                TheExec.Datalog.WriteComment "Read All ECID eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'                Call auto_PrintAllBitbyDSSC(SingleBitArray, EcidReadCycle, ECIDTotalBits, EcidReadBitWidth)
'            End If
'        End If
'        ''''===============   End of Simulated Data ===============
'
'
'        '============================================================
'        '=  Reorder Bit string arrray to Bit array                  =
'        '============================================================
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To EcidReadCycle - 1
'                ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'                SingleStrArray(i, Site) = StrReverse(SingleStrArray(i, Site))
'                gS_SingleStrArray(i, Site) = SingleStrArray(i, Site)
'            Next i
'        Else
'            For i = 0 To EcidReadCycle - 1
'                gS_SingleStrArray(i, Site) = SingleStrArray(i, Site)
'            Next i
'        End If
'
'        FailCnt(Site) = 0
'        ''''Decompose_StrArray_to_BitArray SingleStrArray, SingleBitArray
'        Call auto_OR_2Blocks("ECID", SingleStrArray, SingleBitArray, DoubleBitArray)  ''''calc gL_ECID_FBC
'
'        '====================================================
'        '=  Print the all eFuse Bit data from DSSC          =
'        '====================================================
'        'Print out the translation of Efuse bit data
'        ReDim eFuse_Pgm_Bit(UBound(SingleBitArray)) ''''Must be here
'        ReDim SingleBitArrayStr(UBound(SingleBitArray)) ''''<MUST> 20161114 update
'        For i = 0 To UBound(SingleBitArray)
'            SingleBitArrayStr(i) = CStr(SingleBitArray(i))
'            eFuse_Pgm_Bit(i) = gL_ECIDFuse_Pgm_Bit(Site, i)
'        Next i
'
'        TheExec.Datalog.WriteComment ""
'        TheExec.Datalog.WriteComment "Read All ECID eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'        Call auto_PrintAllBitbyDSSC(SingleBitArray, EcidReadCycle, ECIDTotalBits, EcidReadBitWidth)
'
'        If (gS_EFuse_Orientation = "SingleUp") Then
'            ''''Only 1 block
'            If (condstr = LCase("nonDEID")) Then
'                ''''20161018 update, False: NOT to store DEID to STDF, because it was done in DEID item.
'                ''''For nonDEID, it should re-decode all ECID contents
'                Call auto_EcidPrintData(1, SingleBitArrayStr, False)
'            Else
'                Call auto_EcidPrintData(1, SingleBitArrayStr) 'readout eFuse and set and print out HramLotId ... for block 1
'            End If
'        Else
'            If (condstr = LCase("nonDEID")) Then
'                ''''20161018 update, False: NOT to store DEID to STDF, because it was done in DEID item.
'                ''''For nonDEID, it should re-decode all ECID contents
'                Call auto_EcidPrintData(1, SingleBitArrayStr, False)
'                Call auto_EcidPrintData(2, SingleBitArrayStr, False)
'            Else
'                Call auto_EcidPrintData(1, SingleBitArrayStr)  'readout eFuse and set and print out HramLotId ... for block 1
'                Call auto_EcidPrintData(2, SingleBitArrayStr)  'readout eFuse and set and print out HramLotId ... for block 2
'            End If
'        End If
'
'        ''''20151209 New
'        If (condstr = LCase("DEID")) Then
'            Call auto_EcidCompare_DoubleBit_PgmBit_DEID(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = LCase("DEID_WAT")) Then
'            Call auto_EcidCompare_DoubleBit_PgmBit_DEID_DEV_WAT(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = "all") Then
'            Call auto_EcidCompare_DoubleBit_PgmBit(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = LCase("nonDEID")) Then
'            Call auto_EcidCompare_DoubleBit_PgmBit_nonDEID(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = "stage") Then
'            Call auto_EcidCompare_DoubleBit_PgmBit_byStage(DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        ElseIf (condstr = "category") Then
'            Call auto_eFuse_Compare_DoubleBit_PgmBit_byCategory("ECID", catename_grp, DoubleBitArray, eFuse_Pgm_Bit, FailCnt)
'        Else
'            ''''default, here it prevents any typo issue
'            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (DEID,DEID_WAT,nonDEID,All,Stage,Category)"
'            FailCnt(Site) = -1
'        End If
'
'        TheExec.Flow.TestLimit resultVal:=FailCnt, lowVal:=0, hiVal:=0
'
'    Next Site 'For Each Site In TheExec.Sites
'
'    Call UpdateDLogColumns__False
'
'    DebugPrintFunc ReadPatSet.Value
  
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_EcidSingleDoubleBit(Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidSingleDoubleBit"
    
    Dim site As Variant
    Dim m_stage As String
    Dim i As Long, j As Long, k As Long

    Dim ChkResult As New SiteLong
    Dim mstr_ECID_DEID As String
    Dim mstr_ECID_effbit As String
    Dim mstr_ECID_bits_for_CRC As String         'added  20170623

    Dim bcnt As Long
    Dim m_DEID_startbit As Long
    Dim m_DEID_endbit As Long
    m_DEID_startbit = ECIDFuse.Category(ECIDIndex("Lot_ID")).MSBbit
    m_DEID_endbit = ECIDFuse.Category(ECIDIndex("Y_Coordinate")).LSBbit

    Dim SingleBitArray() As Long, DoubleBitArray() As Long
    ReDim SingleBitArray(ECIDTotalBits - 1)
    ReDim DoubleBitArray(EcidBitPerBlockUsed - 1)

    Dim SingleBitArrayStr() As String
    ''ReDim SingleBitArrayStr(ECIDTotalBits - 1)

    TheExec.Datalog.WriteComment vbCrLf & "Test Instance :: " + TheExec.DataManager.instanceName

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''if there is any site which is non-blank, then decode to gDW_ECID_Read_Decimal_Cate
    ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
    ''''it will be present by Hex and Binary compare with the limit later on.
    ''Call auto_eFuse_setReadData(eFuse_ECID, gDW_ECID_Read_DoubleBitWave, gDW_ECID_Read_Decimal_Cate)
    'Call auto_eFuse_setReadData(eFuse_ECID)

    ''''201901XX New for TTR/PTE improvement
    Call auto_eFuse_setReadData_forSyntax(eFuse_ECID)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_ECID)
    
    If Not (TheExec.DataManager.instanceName Like "*NONDEID*") Then auto_eFuse_Print_Device_code
    
    ''''All the read action has been down in blank and/or MarginRead
    ''''gDW_ECID_Read_cmpsgWavePerCyc used to display the cmpare result (2-bit mode)
    Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_ECID, gB_eFuse_printBitMap)
    If (gS_JobName = "cp1_early") Then
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_ECID, True, gB_eFuse_printReadCate)
    Else
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_ECID, False, gB_eFuse_printReadCate)
    End If
'    Call auto_eFuse_print_DSSCReadWave_Category(eFuse_ECID, False, gB_eFuse_printReadCate)
    
    ''''----------------------------------------------------------------------------------------------
    
    'gDW_CFG_Read_DoubleBitWave = gDW_CFG_Read_DoubleBitWave.ConvertDataTypeTo(DspLong)
    For Each site In TheExec.sites
        DoubleBitArray = gDW_ECID_Read_DoubleBitWave(site).Data
        ''SingleBitArray = gDW_CFG_Read_SingleBitWave.Data
        gS_ECID_Direct_Access_Str(site) = "" ''''is a String [(bitLast)......(bit0)]
        
        For i = 0 To UBound(DoubleBitArray)
            gS_ECID_Direct_Access_Str(site) = CStr(DoubleBitArray(i)) + gS_ECID_Direct_Access_Str(site)
        Next i
        ''TheExec.Datalog.WriteComment "gS_ECID_Direct_Access_Str=" + CStr(gS_ECID_Direct_Access_Str(Site))

        ''''20161114 update for print all bits (DTR) in STDF
        Call auto_eFuse_to_STDF_allBits("ECID", gS_ECID_Direct_Access_Str(site))
        
''        gS_ECID_SingleBit_Str(Site) = "" ''''is a String [(bitLast)......(bit0)]
''        For i = 0 To UBound(SingleBitArray)
''            gS_ECID_SingleBit_Str(Site) = CStr(SingleBitArray(i)) + gS_ECID_SingleBit_Str(Site)
''        Next i
''        ''TheExec.Datalog.WriteComment "gS_ECID_SingleBit_Str=" + CStr(gS_ECID_SingleBit_Str(Site))

        
        ''''[NOTICE] it will be updated later on
''        ''''' 20161003 ADD CRC
''        If (Trim(gS_ECID_CRC_Stage) <> "") Then
''            gS_ECID_CRC_HexStr(Site) = auto_ECID_CRC2HexStr(DoubleBitArray(), gL_ECID_CRC_EndBit) '2016.09.24. Add
''        End If

        ''''20151222 update for all cases
        If (gB_ReadWaferData_flag = False) Then
            ''''FT case, XY is from ECID Read
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(site, HramXCoord(site))
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(site, HramYCoord(site))
            XCoord(site) = HramXCoord(site)
            YCoord(site) = HramYCoord(site)
        Else
            ''''all CP cases and WLFT from the prober
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(site, XCoord(site))
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(site, YCoord(site))
        End If
    
    Next site
    ''''----------------------------------------------------------------------------------------------
    
    Call UpdateDLogColumns(gI_ECID_catename_maxLen)
    ''for ECID Syntax check------------------------------------------------------------------------------
    Call auto_ECID_SyntaxCheck_DEID
    ''for ECID Syntax check------------------------------------------------------------------------------

    ''''gL_ECID_FBC has been check in Blank/MarginRead
    TheExec.Flow.TestLimit resultVal:=gL_ECID_FBC, lowVal:=0, hiVal:=0, Tname:="ECID_FBCount_" + UCase(gS_JobName) '2d-s=0
    Call UpdateDLogColumns__False
    
    Call auto_eFuse_ReadAllData_to_DictDSPWave("ECID", False, False)
        ''''register and print out the IEDA data----------------------------------------------------------------------------
    If (True) Then
        Dim LotStr As String
        Dim Waferstr As String
        Dim X_Coor_Str As String
        Dim Y_Coor_Str As String
        Dim ECID_DEID_Str As String
        
        Dim mSV_LotID As New SiteVariant
        Dim mSV_WaferID As New SiteVariant
        Dim mSV_XCoord As New SiteVariant
        Dim mSV_YCoord As New SiteVariant
        
        With ECIDFuse
            mSV_LotID = .Category(ECIDIndex("Lot_ID")).Read.ValStr
            mSV_WaferID = .Category(ECIDIndex("Wafer_ID")).Read.ValStr
            mSV_XCoord = .Category(ECIDIndex("X_Coordinate")).Read.ValStr
            mSV_YCoord = .Category(ECIDIndex("Y_Coordinate")).Read.ValStr
        End With

        LotStr = ""
        Waferstr = ""
        X_Coor_Str = ""
        Y_Coor_Str = ""
        ECID_DEID_Str = ""
        
        For Each site In TheExec.sites.Existing
            LotStr = LotStr + mSV_LotID(site)
            Waferstr = Waferstr + CStr(mSV_WaferID(site))
            X_Coor_Str = X_Coor_Str + CStr(mSV_XCoord(site))
            Y_Coor_Str = Y_Coor_Str + CStr(mSV_YCoord(site))
            ECID_DEID_Str = ECID_DEID_Str + ECIDFuse.Category(gI_Index_DEID).Read.BitStrL(site)
            If (site = TheExec.sites.Existing.Count - 1) Then
            Else
                LotStr = LotStr + ","
                Waferstr = Waferstr + ","
                X_Coor_Str = X_Coor_Str + ","
                Y_Coor_Str = Y_Coor_Str + ","
                ECID_DEID_Str = ECID_DEID_Str + ","
            End If
'            If (Site = TheExec.sites.Existing.Count - 1) Then
'                LotStr = LotStr + ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(Site)
'                Waferstr = Waferstr + ECIDFuse.Category(ECIDIndex("Wafer_ID")).Read.ValStr(Site)
'                X_Coor_Str = X_Coor_Str + ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.ValStr(Site)
'                Y_Coor_Str = Y_Coor_Str + ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.ValStr(Site)
'                ECID_DEID_Str = ECID_DEID_Str + ECIDFuse.Category(gI_Index_DEID).Read.BitStrL(Site)
'            Else
'                LotStr = LotStr + ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(Site) + ","
'                Waferstr = Waferstr + ECIDFuse.Category(ECIDIndex("Wafer_ID")).Read.ValStr(Site) + ","
'                X_Coor_Str = X_Coor_Str + ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.ValStr(Site) + ","
'                Y_Coor_Str = Y_Coor_Str + ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.ValStr(Site) + ","
'                ECID_DEID_Str = ECID_DEID_Str + ECIDFuse.Category(gI_Index_DEID).Read.BitStrL(Site) + ","
'            End If

        Next site

        LotStr = auto_checkIEDAString(LotStr)
        Waferstr = auto_checkIEDAString(Waferstr)
        X_Coor_Str = auto_checkIEDAString(X_Coor_Str)
        Y_Coor_Str = auto_checkIEDAString(Y_Coor_Str)
        ECID_DEID_Str = auto_checkIEDAString(ECID_DEID_Str)
        
        TheExec.Datalog.WriteComment vbCrLf & "Test Instance :: " + TheExec.DataManager.instanceName
        TheExec.Datalog.WriteComment " ECID (all sites iEDA format)::"
        TheExec.Datalog.WriteComment " Lot ID    = " + LotStr
        TheExec.Datalog.WriteComment " Wafer ID  = " + Waferstr
        TheExec.Datalog.WriteComment " X_Coor    = " + X_Coor_Str
        TheExec.Datalog.WriteComment " Y_Coor    = " + Y_Coor_Str
        TheExec.Datalog.WriteComment " ECID_DEID = " + ECID_DEID_Str & vbCrLf

        '============================================
        '=  Write Data to Register Edit (HKEY)      =
        '============================================
        Call RegKeySave("eFuseLotNumber", LotStr)
        Call RegKeySave("eFuseWaferID", Waferstr)
        Call RegKeySave("eFuseDieX", X_Coor_Str)
        Call RegKeySave("eFuseDieY", Y_Coor_Str)
        Call RegKeySave("Hram_ECID_53bit", ECID_DEID_Str)

    End If
    ''''End of register and print out the IEDA data----------------------------------------------------------------------------

    
        ''''201811XX reset
    If (gS_JobName = "cp1_early") Then gS_JobName = "cp1" ''''<MUST>
Exit Function

End If

'
'    ''''------------------------------------------------------------------------------------------------------------------
'    ''''<Important Notice>
'    ''''gS_SingleStrArray() was extracted in the module auto_ECID_Read_by_OR_2Blocks() then used in auto_EcidSingleDoubleBit()
'    ''''gS_SingleStrArray() is the result of the NormRead or MarginRead
'    ''''
'    ''''So it doesn't need to run the pattern and DSSC to get the SignalStrArray, and save test time
'    ''''------------------------------------------------------------------------------------------------------------------
'    ''TheExec.Datalog.WriteComment vbCrLf & "Test Instance :: " + TheExec.DataManager.InstanceName
'
'    Call UpdateDLogColumns(gI_ECID_catename_maxLen)
'
'    ''''Here using the programming stage of 'Lot_ID' stands for the Stage for all ECID bits.(It Should BE!!!)
'    m_stage = LCase(ECIDFuse.Category(ECIDIndex("Lot_ID")).Stage)
'
'    For Each Site In TheExec.Sites
'        ' OR 2 block bit by bit
'        ' calc gL_ECID_FBC
'        Call auto_OR_2Blocks("ECID", gS_SingleStrArray(), SingleBitArray(), DoubleBitArray())     'calc gL_ECID_FBC
'
'        'ECID_512bit
'        '==================================================================================
'        '=  Get ECID LotID,WaferID and X/Y Coordinates from DSSC for writing to HKEY      =
'        '==================================================================================
'        gS_ECID_Direct_Access_Str(Site) = "" ''''is a String [(bitLast)......(bit0)]
'
'        ''''<User Maintain by the DAP pattern>
'        ''''For i = 0 To ((UBound(DoubleBitArray) + 1) / 2) - 1 'TMA case, 'Because there are only 128 bits in ECID DAP pattern
'        For i = 0 To UBound(DoubleBitArray) 'there are only 256 bits in ELBA ECID DAP pattern
'            gS_ECID_Direct_Access_Str(Site) = CStr(DoubleBitArray(i)) + gS_ECID_Direct_Access_Str(Site)
'        Next i
'        ''''20161114 update for print all bits (DTR) in STDF
'        If (gS_JobName <> "cp1") Then
'            ''''In CP1, it will be display in auto_EcidSingleDoubleBit_nonDEID()
'            Call auto_eFuse_to_STDF_allBits("ECID", gS_ECID_Direct_Access_Str(Site))
'        End If
'        ''''----------------------------------------------------------------------------------------------
'        ''''20160630 Add
'        ReDim SingleBitArrayStr(UBound(SingleBitArray)) ''''<MUST> 20161114 update
'        gS_ECID_SingleBit_Str(Site) = ""     ''''is a String [(bitLast)......(bit0)]
'        For i = 0 To UBound(SingleBitArray)
'            gS_ECID_SingleBit_Str(Site) = CStr(SingleBitArray(i)) + gS_ECID_SingleBit_Str(Site)
'            SingleBitArrayStr(i) = CStr(SingleBitArray(i))
'        Next i
'        ''''----------------------------------------------------------------------------------------------
'
'        mstr_ECID_DEID = "" ''''is a String [LSB......MSB]
'        mstr_ECID_effbit = ""
'        mstr_ECID_bits_for_CRC = "" ''''<MUST>be Clear per Site, added on 20170623
'        bcnt = 0
'        For k = 0 To EcidRowPerBlock - 1    ''0...15(R2L), 0...15(SUP)
'            For j = 0 To EcidBitsPerRow - 1 ''0...15(R2L), 0...31(SUP)
'                ''''if only count to Ycoord, it should be only DEID.
'                If (bcnt >= m_DEID_startbit And bcnt <= m_DEID_endbit) Then
'                    mstr_ECID_DEID = CStr(DoubleBitArray(bcnt)) + mstr_ECID_DEID
'                End If
'                ''''20150721 update
'                mstr_ECID_effbit = CStr(DoubleBitArray(bcnt)) + mstr_ECID_effbit
'                ''' added on 20170623
'                If bcnt <= gL_ECID_CRC_EndBit Then
'                    mstr_ECID_bits_for_CRC = CStr(DoubleBitArray(bcnt)) + mstr_ECID_bits_for_CRC    '''[MSB ....... LSB]
'                End If
'                bcnt = bcnt + 1 ''''<MUST>
'            Next j
'        Next k
'
'        ECIDFuse.Category(gI_Index_DEID).Read.BitStrL(Site) = mstr_ECID_DEID
'        ECIDFuse.Category(gI_Index_DEID).Read.BitstrM(Site) = StrReverse(mstr_ECID_DEID)
'
'        '====================================================
'        '=  Print the all eFuse Bit data from HRAM          =
'        '====================================================
'        TheExec.Datalog.WriteComment ""
'
'        If (TheExec.Sites(Site).SiteVariableValue("ECIDChk_Var") <> 1 Or gB_ECID_decode_flag(Site) = False) Then
'            If (gS_EFuse_Orientation = "SingleUp") Then
'                ''''Only 1 block
'                Call auto_EcidPrintData(1, SingleBitArrayStr)
'            Else
'                Call auto_EcidPrintData(1, SingleBitArrayStr)
'                Call auto_EcidPrintData(2, SingleBitArrayStr)
'            End If
'            TheExec.Datalog.WriteComment ""
'            Call auto_PrintAllBitbyDSSC(SingleBitArray(), EcidReadCycle, ECIDTotalBits, EcidReadBitWidth)
'        End If
'
'        ''''' 20161003 ADD CRC
'        If (Trim(gS_ECID_CRC_Stage) <> "") Then
'            gS_ECID_CRC_HexStr(Site) = auto_ECID_CRC2HexStr(DoubleBitArray(), gL_ECID_CRC_EndBit) '2016.09.24. Add
'        End If
'
'        ''''20151222 update for all cases
'        If (gB_ReadWaferData_flag = False) Then
'            ''''FT case, XY is from ECID Read
'            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(Site, HramXCoord(Site))
'            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(Site, HramYCoord(Site))
'        Else
'            ''''all CP cases and WLFT from the prober
'            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(Site, XCoord(Site))
'            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(Site, YCoord(Site))
'        End If
'
'        '========================================================================
'        '=  Add a gating rule for eFuse consolidation                           =
'        '=  The first letter has to be [A-Z], the 2nd to 6th letters have to    =
'        '=  be [0-9] or [A-Z]. Besides, the summation for bit6~53th has to be   =
'        '=  greater than 3. And the reserved area has to be all 0s (blank)      =
'        '========================================================================
'        Call auto_printECIDContent_AllBits
'
'        If (gS_JobName = m_stage) Then
'            If (TheExec.Flow.EnableWord("WAT_Enable") = True) Then
'                ChkResult(Site) = auto_Chk_ECID_Content_DEID_DEV_WAT(mstr_ECID_effbit)
'            Else
'                ''''only for the first DEID and reserved bits,
'                ''''<NOTICE> but others' limits should be zero in the very First time.
'                ''''20160907 update, bypass others except DEID categories.
'                ChkResult(Site) = auto_Chk_ECID_Content_DEID(mstr_ECID_DEID)
'            End If
'        ElseIf (gS_JobName <> m_stage) Then ''''non CP1 (m_stage) Job
'            ChkResult(Site) = auto_Chk_ECID_Content_AllBits(mstr_ECID_effbit)
'        End If
'
'        TheExec.Flow.TestLimit resultVal:=ChkResult, lowVal:=1, hiVal:=1, Tname:="ECID_Syntax_Chk"
'        TheExec.Flow.TestLimit resultVal:=gL_ECID_FBC, lowVal:=0, hiVal:=EcidHiLimitSingleDoubleBitCheck, Tname:="FailBitCount"  '2d-s=0
'
'        If (gS_JobName <> "cp1") Then
'            ''' 20170623 add
'            TheExec.Datalog.WriteComment ""
'            TheExec.Datalog.WriteComment "ECID bit string for CRC calculation on site " & CStr(Site) & " (MSB to LSB) = " & mstr_ECID_bits_for_CRC
'            TheExec.Datalog.WriteComment ""
'        End If
'    Next Site  'For Each Site In TheExec.Sites
'
'    Call UpdateDLogColumns__False
'
'    ''''20170111 Add
'    Call auto_eFuse_ReadAllData_to_DictDSPWave("ECID", False, False)
'
'    ''''register and print out the IEDA data----------------------------------------------------------------------------
'    If (True) Then
''        Dim LotStr As String
''        Dim Waferstr As String
''        Dim X_Coor_Str As String
''        Dim Y_Coor_Str As String
''        Dim ECID_DEID_Str As String
''
''        LotStr = ""
''        Waferstr = ""
''        X_Coor_Str = ""
''        Y_Coor_Str = ""
''        ECID_DEID_Str = ""
''
''        For Each Site In TheExec.Sites.Existing
''
''            If (Site = TheExec.Sites.Existing.Count - 1) Then
''                LotStr = LotStr + ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(Site)
''                Waferstr = Waferstr + ECIDFuse.Category(ECIDIndex("Wafer_ID")).Read.ValStr(Site)
''                X_Coor_Str = X_Coor_Str + ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.ValStr(Site)
''                Y_Coor_Str = Y_Coor_Str + ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.ValStr(Site)
''                ECID_DEID_Str = ECID_DEID_Str + ECIDFuse.Category(gI_Index_DEID).Read.BitStrL(Site)
''            Else
''                LotStr = LotStr + ECIDFuse.Category(ECIDIndex("Lot_ID")).Read.ValStr(Site) + ","
''                Waferstr = Waferstr + ECIDFuse.Category(ECIDIndex("Wafer_ID")).Read.ValStr(Site) + ","
''                X_Coor_Str = X_Coor_Str + ECIDFuse.Category(ECIDIndex("X_Coordinate")).Read.ValStr(Site) + ","
''                Y_Coor_Str = Y_Coor_Str + ECIDFuse.Category(ECIDIndex("Y_Coordinate")).Read.ValStr(Site) + ","
''                ECID_DEID_Str = ECID_DEID_Str + ECIDFuse.Category(gI_Index_DEID).Read.BitStrL(Site) + ","
''            End If
''
''        Next Site
''
''        LotStr = auto_checkIEDAString(LotStr)
''        Waferstr = auto_checkIEDAString(Waferstr)
''        X_Coor_Str = auto_checkIEDAString(X_Coor_Str)
''        Y_Coor_Str = auto_checkIEDAString(Y_Coor_Str)
''        ECID_DEID_Str = auto_checkIEDAString(ECID_DEID_Str)
''
''        TheExec.Datalog.WriteComment vbCrLf & "Test Instance :: " + TheExec.DataManager.InstanceName
''        TheExec.Datalog.WriteComment " ECID (all sites iEDA format)::"
''        TheExec.Datalog.WriteComment " Lot ID    = " + LotStr
''        TheExec.Datalog.WriteComment " Wafer ID  = " + Waferstr
''        TheExec.Datalog.WriteComment " X_Coor    = " + X_Coor_Str
''        TheExec.Datalog.WriteComment " Y_Coor    = " + Y_Coor_Str
''        TheExec.Datalog.WriteComment " ECID_DEID = " + ECID_DEID_Str & vbCrLf
''
''        '============================================
''        '=  Write Data to Register Edit (HKEY)      =
''        '============================================
''        Call RegKeySave("eFuseLotNumber", LotStr)
''        Call RegKeySave("eFuseWaferID", Waferstr)
''        Call RegKeySave("eFuseDieX", X_Coor_Str)
''        Call RegKeySave("eFuseDieY", Y_Coor_Str)
''        Call RegKeySave("Hram_ECID_53bit", ECID_DEID_Str)
'
'    End If
    ''''End of register and print out the IEDA data----------------------------------------------------------------------------
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_ReadWaferData()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ReadWaferData"

    Dim LotTmp As String
    Dim Loc_dash As Long
    Dim site As Variant
    
    Dim i As Long
    Dim ch1st As String
    Dim ch2_6 As String
    Dim ascVal As Long
    Dim tmpStr As String
    Dim chkVal As Long
    Dim m_chkStage As New SiteBoolean
    Dim m_tmpwfid As String
    Dim m_len As Long

    ''''20170217 update
    LotTmp = Trim(UCase(TheExec.Datalog.Setup.LotSetup.LotID))
    m_tmpwfid = Trim(CStr(TheExec.Datalog.Setup.WaferSetup.ID))
    m_len = Len(LotTmp)

    '=== Simulated Data ===
    If (TheExec.TesterMode = testModeOffline) Then
        If (LotTmp = "" And m_tmpwfid = "") Then
            TheExec.Datalog.Setup.LotSetup.LotID = "N98G17-02C1" ''''"DUMMY0-25C1"
            TheExec.Datalog.Setup.WaferSetup.ID = "2"
            LotTmp = "N98G17-02C1"
            m_tmpwfid = "2"
        End If
    Else
        ''''20170217 update
        If (LotTmp = "") Then
            LotTmp = "000000" ''''to avoid runtime VBT error stop
            TheExec.Datalog.WriteComment "[WARNING] Input LotID is Empty, Set it to (000000). "
        End If
    End If
    
    Loc_dash = InStr(1, LotTmp, "-")
    
    If Loc_dash <> 0 Then
        LotID = Mid(LotTmp, 1, Loc_dash - 1)
    Else
        LotID = LotTmp
    End If
   
    Call UpdateDLogColumns(30)

    ''''Syntax Check LotID of the prober
    ch1st = Mid(LotID, 1, 1)
    ascVal = Asc(LCase(ch1st))
    If (ascVal < 97 Or ascVal > 122) Then ''''a=97 and z=122 in ANSI character
        chkVal = 0 'Fail
        tmpStr = "First Character of Prober LotID (" + UCase(ch1st) + ") is not [A-Z]."
        TheExec.Datalog.WriteComment tmpStr
    Else
        chkVal = 1 'Pass

        If (Len(LotID) <> EcidCharPerLotId) Then ''''EcidCharPerLotId=6
            tmpStr = "Character Numbers of Prober LotID (" + UCase(LotID) + ") is NOT Six Characters."
            TheExec.Datalog.WriteComment tmpStr
            chkVal = 0 'Fail
        Else
            For i = 2 To EcidCharPerLotId  ''''EcidCharPerLotId=6
                ch2_6 = Mid(LotID, i, 1)
                ascVal = Asc(LCase(ch2_6))
                If ascVal < 97 Or ascVal > 122 Then    'a=97 and z=122 in ANSI character
                    If ascVal < 48 Or ascVal > 57 Then ''0'=48 and '9'=57 in ANSI character
                        chkVal = 0  'Fail
                        tmpStr = "Second-to-Sixth Characters of Prober LotID (" + UCase(LotID) + ") are not [A-Z] or [0-9]."
                        TheExec.Datalog.WriteComment tmpStr
                        Exit For
                    Else
                        chkVal = 1 'Pass
                    End If
                Else
                End If
            Next i
        End If
    End If
    If (chkVal = 0) Then
        m_chkStage = False
        Call auto_eFuse_SetPatTestPass_Flag_SiteAware("ECID", "Lot_ID", m_chkStage, True)
    End If

    TheExec.Flow.TestLimit chkVal, 1, 1, Tname:="Prober_LotID", PinName:=LotID
    ''''-----------------------------------------------------------------------------
    
    Call auto_eFuse_LotID_to_setWriteVariable(LotID)

    ''''20170217 update
    If (TheExec.TesterMode = testModeOffline) Then
        If m_tmpwfid <> "" Then
            If (IsNumeric(m_tmpwfid) = True) Then
                WaferID = CLng(m_tmpwfid)
            Else
                tmpStr = "<Offline> Prober WaferID (" + m_tmpwfid + ") is NOT numeric, set it to 25 (psudo wafer id)."
                TheExec.Datalog.WriteComment tmpStr
                WaferID = 25
            End If
        Else
            WaferID = 25
            ''TheExec.Datalog.WriteComment vbTab & "<Offline> Set WaferID to 25 (pseudo wafer id)"
        End If
    Else
        ''''Here is the Online Mode
        If m_tmpwfid <> "" Then
            If (IsNumeric(m_tmpwfid) = True) Then
                WaferID = CLng(m_tmpwfid)
            Else
                tmpStr = "Prober WaferID (" + m_tmpwfid + ") is NOT numeric, set it to 0."
                TheExec.Datalog.WriteComment tmpStr
                WaferID = 0
            End If
        Else
            WaferID = 0
        End If
    End If
    
    ''''Syntax Check WaferID of the prober
    If (WaferID < 1 Or WaferID > 25) Then ''''range 1...25
        tmpStr = "Prober WaferID (" + CStr(WaferID) + ") is out of the range [1...25]."
        TheExec.Datalog.WriteComment tmpStr
        m_chkStage = False
        Call auto_eFuse_SetPatTestPass_Flag_SiteAware("ECID", "Wafer_ID", m_chkStage, True)
    End If
    TheExec.Flow.TestLimit WaferID, 1, 25, Tname:="Prober_WaferID", PinName:=CStr(WaferID)
    ''''-----------------------------------------------------------------------------
    
    ''''---------- For CharZ Datalog --------
    HramLotId = LotID
    HramWaferId = WaferID
    ''''-------------------------------------
    
    Call auto_eFuse_SetWriteVariable_SiteAware("ECID", "Wafer_ID", HramWaferId, False)
    
    For Each site In TheExec.sites
        XCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetXCoord(site)
        YCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetYCoord(site)
        If (TheExec.TesterMode = testModeOffline) Then
            ''Call setXY(11, 18) ''''Engineer Trial
            If (XCoord(site) = -32768 Or YCoord(site) = -32768) Then
                Call setXY(5, 6) ''''set a pseudo XY coordinate
                ''TheExec.Datalog.WriteComment vbTab & "<Offline> Call setXY(5, 6) (pseudo XY_Coordinate)"
                XCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetXCoord(site)
                YCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetYCoord(site)
            End If
        End If

        ''''Syntax Check XY of the prober
        If (XCoord(site) < XCOORD_LoLMT Or XCoord(site) > XCOORD_HiLMT) Then
            tmpStr = "Prober X_Coordinate (" + CStr(XCoord(site)) + ") is out of the range [" + CStr(XCOORD_LoLMT) + "..." + CStr(XCOORD_HiLMT) + "]."
            TheExec.Datalog.WriteComment tmpStr
            Call auto_eFuse_SetPatTestPass_Flag("ECID", "X_Coordinate", False, True)
        End If
        If (YCoord(site) < YCOORD_LoLMT Or YCoord(site) > YCOORD_HiLMT) Then
            tmpStr = "Prober Y_Coordinate (" + CStr(YCoord(site)) + ") is out of the range [" + CStr(YCOORD_LoLMT) + "..." + CStr(YCOORD_HiLMT) + "]."
            TheExec.Datalog.WriteComment tmpStr
            Call auto_eFuse_SetPatTestPass_Flag("ECID", "Y_Coordinate", False, True)
        End If

        ''''-----------------------------------------------------------------------------
        Call auto_eFuse_SetWriteDecimal("ECID", "X_Coordinate", XCoord(site), False, False)
        Call auto_eFuse_SetWriteDecimal("ECID", "Y_Coordinate", YCoord(site), False, False)
        
        ''''20171007 update to simulate WFLT case
        If (TheExec.TesterMode = testModeOffline And gS_JobName = "wlft") Then
            TheExec.Datalog.WriteComment vbTab & "<Offline> Simulate WLFT Prober XY <> DUT ECID XY, Below is DUT Simulate ECID XY"
            Call auto_eFuse_SetWriteDecimal("ECID", "X_Coordinate", XCoord(site) + 1, True, False)
            Call auto_eFuse_SetWriteDecimal("ECID", "Y_Coordinate", YCoord(site) + 1, True, False)
        End If
        ''''-----------------------------------------------------------------------------
    Next site

    TheExec.Flow.TestLimit XCoord, XCOORD_LoLMT, XCOORD_HiLMT, Tname:="Prober_X"
    TheExec.Flow.TestLimit YCoord, YCOORD_LoLMT, YCOORD_HiLMT, Tname:="Prober_Y"

    gB_ReadWaferData_flag = True


    ''''Print Out the Prober's Information
    Dim m_user_prober As String
    TheExec.Datalog.WriteComment vbCrLf & funcName + "::"
    TheExec.Datalog.WriteComment "---------------------------"
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "  Lot ID = " + LotID
        TheExec.Datalog.WriteComment "Wafer ID = " + CStr(WaferID)
        TheExec.Datalog.WriteComment "X coor (site " + CStr(site) + ")= " + CStr(XCoord(site))
        TheExec.Datalog.WriteComment "Y coor (site " + CStr(site) + ")= " + CStr(YCoord(site))
        
        ''''20161021 update per Laba and C651 PE request
        ''''20161118 Per Jack's comment, bypass WLFT case.
        ''''-----------------------------------------------------------------------------------
        m_user_prober = auto_WaferData_to_HexECID(LotID, WaferID, XCoord(site), YCoord(site))
        TheExec.Datalog.WriteComment "Prober Hex code = " + m_user_prober
        If (True And Not (gS_JobName Like "*wlft*")) Then
            ''''Write to PRR-Part_TEXT in STDF
            Call TheExec.Datalog.Setup.SetPRRPartInfo(tl_SelectSite, site, , m_user_prober)
        End If
        ''''-----------------------------------------------------------------------------------
        TheExec.Datalog.WriteComment "---------------------------"
    Next site


    Call UpdateDLogColumns__False

Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_ReadHandlerData()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ReadHandlerData"

    ''TheExec.Datalog.WriteComment ""
    'update by Jason's request to fixed Galaxy multi-site format, 140411
    Dim site As Variant
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment ("<@Chuck_ID=" & site & "|" & RegKeyRead("Cover_ID") & ">")
        TheExec.Datalog.WriteComment ("<@Dut_Temperature=" & site & "|" & RegKeyRead("Dut_Temperature") & ">")
        TheExec.Datalog.WriteComment ("<@Handler_Arm_ID=" & site & "|" & RegKeyRead("Handler_Arm_ID") & ">")
        TheExec.Datalog.WriteComment ("<@Rework_Flag=" & site & "|" & RegKeyRead("FT_ReTest") & ">")
        TheExec.Datalog.WriteComment ("<@Socket_ID=" & site & "|" & RegKeyRead("Socket_ID") & ">")
        TheExec.Datalog.WriteComment ("<@Sort_Stage=" & site & "|" & RegKeyRead("Sort_Code") & ">")
    Next site
    TheExec.Datalog.WriteComment ""

Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_ShowECIDData()

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ShowECIDData"

    TheExec.Datalog.WriteComment ""
    'update by Jason's request to fixed Galaxy multi-site format, 140411
    Dim site As Variant
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "<@efuse_lot_ID=" & site & "|" & HramLotId(site) & ">"
        TheExec.Datalog.WriteComment "<@efuse_wafer_ID=" & site & "|" & HramLotId(site) & "." & Format(CStr(HramWaferId(site)), "00") & ">"
    Next site
    TheExec.Datalog.WriteComment ""

Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''nonDEID=non DeviceID
Public Function auto_EcidSingleDoubleBit_nonDEID(Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidSingleDoubleBit_nonDEID"
    
    Dim site As Variant
    Dim i As Long, j As Long, k As Long

    Dim ChkResult As New SiteLong
    Dim mstr_ECID_effbit As String
    Dim mstr_ECID_bits_for_CRC As String

    Dim SingleBitArray() As Long, DoubleBitArray() As Long
    ReDim SingleBitArray(ECIDTotalBits - 1)
    ReDim DoubleBitArray(EcidBitPerBlockUsed - 1)

    Dim bcnt As Long
    Dim blank As New SiteBoolean
    Dim cycleNum As Long, BitPerCycle As Long

    Dim SingleBitArrayStr() As String
    ''ReDim SingleBitArrayStr(ECIDTotalBits - 1)
    Dim m_tmpStr As String
    Dim m_siteVar As String
    m_siteVar = "ECIDChk_Var"

    ''''------------------------------------------------------------------------------------------------------------------
    ''''<Important Notice>
    ''''gS_SingleStrArray() was extracted in the module auto_ECID_Read_by_OR_2Blocks_TMPS_ADC() then used in auto_EcidSingleDoubleBit_TMPS_ADC()
    ''''gS_SingleStrArray() is the result of the NormRead or MarginRead
    ''''
    ''''So it doesn't need to run the pattern and DSSC to get the SignalStrArray, and save test time
    ''''------------------------------------------------------------------------------------------------------------------
    TheExec.Datalog.WriteComment vbCrLf & "Test Instance   :: " + TheExec.DataManager.instanceName
    
    cycleNum = EcidReadCycle: BitPerCycle = ECIDBitPerCycle
    
    
''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    Call auto_eFuse_setReadData_forSyntax(eFuse_ECID)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_ECID)
    
    ''''All the read action has been down in blank and/or MarginRead
    ''''gDW_ECID_Read_cmpsgWavePerCyc used to display the cmpare result (2-bit mode)
    Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_ECID, gB_eFuse_printBitMap)
    If (gS_JobName = "cp1_early") Then
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_ECID, True, gB_eFuse_printReadCate)
    Else
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_ECID, False, gB_eFuse_printReadCate)
    End If
    
    
    ''''Print CRC calcBits information
    Dim m_crcBitWave As New DSPWave
    Dim mS_hexStr As New SiteVariant
    Dim mS_bitStrM As New SiteVariant
    Dim m_debugCRC As Boolean
    Dim m_cnt As Long
    m_debugCRC = False
    
    ''''<MUST> Initialize
    gS_ECID_Read_calcCRC_hexStr = "0x0000"
    gS_ECID_Read_calcCRC_bitStrM = ""
    CRC_Shift_Out_String = ""
    If (auto_eFuse_check_Job_cmpare_Stage(gS_ECID_CRC_Stage) >= 0) Then
        Call rundsp.eFuse_Read_to_calc_CRCWave(eFuse_ECID, gL_ECID_CRC_BitWidth, m_crcBitWave)
        TheHdw.Wait 1# * ms ''''check if it needs
        
        If (m_debugCRC = False) Then
            ''''Here get gS_CFG_Read_calcCRC_hexStr for the syntax check
            Call auto_eFuse_bitWave_to_binStr_HexStr(m_crcBitWave, gS_ECID_Read_calcCRC_bitStrM, gS_ECID_Read_calcCRC_hexStr, True, m_debugCRC)
        Else
            ''''m_debugCRC=True => Debug purpose for the print
            TheExec.Datalog.WriteComment "------Read CRC Category Result------"
            Call auto_eFuse_bitWave_to_binStr_HexStr(m_crcBitWave, gS_ECID_Read_calcCRC_bitStrM, gS_ECID_Read_calcCRC_hexStr, True, m_debugCRC)
            TheExec.Datalog.WriteComment ""

            ''''[Pgm CRC calcBits] only gS_CFG_CRC_Stage=Job and CFGChk_Var=1
            If (gS_ECID_CRC_Stage = gS_JobName) Then
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
    'gDW_ECID_Read_DoubleBitWave = gDW_ECID_Read_DoubleBitWave.ConvertDataTypeTo(DspLong)
    For Each site In TheExec.sites
        DoubleBitArray = gDW_ECID_Read_DoubleBitWave(site).Data
        ''SingleBitArray = gDW_CFG_Read_SingleBitWave.Data
        gS_ECID_Direct_Access_Str(site) = "" ''''is a String [(bitLast)......(bit0)]
        
        For i = 0 To UBound(DoubleBitArray)
            gS_ECID_Direct_Access_Str(site) = CStr(DoubleBitArray(i)) + gS_ECID_Direct_Access_Str(site)
        Next i
        ''TheExec.Datalog.WriteComment "gS_ECID_Direct_Access_Str=" + CStr(gS_ECID_Direct_Access_Str(Site))

        ''''20161114 update for print all bits (DTR) in STDF
        Call auto_eFuse_to_STDF_allBits("ECID", gS_ECID_Direct_Access_Str(site))

        ''''20151222 update for all cases
        If (gB_ReadWaferData_flag = False) Then
            ''''FT case, XY is from ECID Read
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(site, HramXCoord(site))
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(site, HramYCoord(site))
            XCoord(site) = HramXCoord(site)
            YCoord(site) = HramYCoord(site)
        Else
            ''''all CP cases and WLFT from the prober
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(site, XCoord(site))
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(site, YCoord(site))
        End If
    
    Next site
    ''''----------------------------------------------------------------------------------------------
    
    Call UpdateDLogColumns(gI_ECID_catename_maxLen)
    ''for ECID Syntax check------------------------------------------------------------------------------
    Call auto_ECID_SyntaxCheck_All
    ''for ECID Syntax check------------------------------------------------------------------------------

    ''''gL_ECID_FBC has been check in Blank/MarginRead
    TheExec.Flow.TestLimit resultVal:=gL_ECID_FBC, lowVal:=0, hiVal:=0, Tname:="ECID_FBCount_" + UCase(gS_JobName) '2d-s=0
    
    If (gS_JobName = "cp1") Then
        ''' 20170623 add
        Dim m_DoubleBitArray() As Long

        Dim m_DoubleBitWave As New DSPWave
        For Each site In TheExec.sites
'            m_DoubleBitWave = gDW_ECID_Read_DoubleBitWave(Site).Copy
'
'            m_DoubleBitArray = m_DoubleBitWave.Data
            m_DoubleBitArray = gDW_ECID_Read_DoubleBitWave(site).Data

            mstr_ECID_bits_for_CRC = "" ''''<MUST>be Clear per Site, added on 20170623
            
            bcnt = 0
            For i = 0 To UBound(DoubleBitArray)
                ''' added on 20170623
                If bcnt <= gL_ECID_CRC_EndBit Then
                    mstr_ECID_bits_for_CRC = CStr(m_DoubleBitArray(bcnt)) + mstr_ECID_bits_for_CRC    '''[MSB ....... LSB]
                End If
                bcnt = bcnt + 1 ''''<MUST>
            Next i
            TheExec.Datalog.WriteComment ""
            TheExec.Datalog.WriteComment "ECID bit string for CRC calculation on site " & CStr(site) & " (MSB to LSB) = " & mstr_ECID_bits_for_CRC
            TheExec.Datalog.WriteComment ""
        Next
    End If
    
    Call UpdateDLogColumns__False
    
    ''''20170111 Add
    Call auto_eFuse_ReadAllData_to_DictDSPWave("ECID", False, False)

Exit Function

End If

'    Call UpdateDLogColumns(gI_ECID_catename_maxLen)
'
'    For Each Site In TheExec.Sites
'        ' OR 2 block bit by bit
'        ' calc gL_ECID_FBC
'
'        Call auto_OR_2Blocks("ECID", gS_SingleStrArray(), SingleBitArray(), DoubleBitArray())     'calc gL_ECID_FBC
'
'        ''''----------------------------------------------------------------------------------------------
'        gS_ECID_Direct_Access_Str(Site) = "" ''''is a String [(bitLast)......(bit0)]
'
'        ''''<User Maintain by the DAP pattern>
'        mstr_ECID_effbit = ""
'        mstr_ECID_bits_for_CRC = "" ''''<MUST>be Clear per Site, added on 20170623
'        bcnt = 0
'        For i = 0 To UBound(DoubleBitArray)
'            gS_ECID_Direct_Access_Str(Site) = CStr(DoubleBitArray(i)) + gS_ECID_Direct_Access_Str(Site)
'            mstr_ECID_effbit = CStr(DoubleBitArray(i)) + mstr_ECID_effbit ''''20161108 move here from the below.
'
'            ''' added on 20170623
'            If bcnt <= gL_ECID_CRC_EndBit Then
'                mstr_ECID_bits_for_CRC = CStr(DoubleBitArray(bcnt)) + mstr_ECID_bits_for_CRC    '''[MSB ....... LSB]
'            End If
'            bcnt = bcnt + 1 ''''<MUST>
'        Next i
'
'        ''''20161114 update for print all bits (DTR) in STDF
'        Call auto_eFuse_to_STDF_allBits("ECID", gS_ECID_Direct_Access_Str(Site))
'        ''''----------------------------------------------------------------------------------------------
'        ''''----------------------------------------------------------------------------------------------
'        ''''20160630 Add
'        ReDim SingleBitArrayStr(UBound(SingleBitArray)) ''''<MUST> 20161114 update
'        gS_ECID_SingleBit_Str(Site) = ""     ''''is a String [(bitLast)......(bit0)]
'        For i = 0 To UBound(SingleBitArray)
'            gS_ECID_SingleBit_Str(Site) = CStr(SingleBitArray(i)) + gS_ECID_SingleBit_Str(Site)
'            SingleBitArrayStr(i) = CStr(SingleBitArray(i)) ''''20161108 move here from the below.
'        Next i
'        ''''----------------------------------------------------------------------------------------------
'
'        '====================================================
'        '=  Print the all eFuse Bit data from HRAM          =
'        '====================================================
'        TheExec.Datalog.WriteComment ""
'
'        If (TheExec.Sites(Site).SiteVariableValue("ECIDChk_Var") <> 1) Then
'            ''''only extract nonDEID data to ECIDFuse.Category.Read...
'            If (gS_EFuse_Orientation = "SingleUp") Then
'                ''''Only 1 block
'                ''''20161018 update, False: NOT to store DEID to STDF, because it was done in DEID item.
'                ''''For nonDEID, it should re-decode all ECID contents
'                Call auto_EcidPrintData(1, SingleBitArrayStr, False)
'            Else
'                ''''20161018 update, False: NOT to store DEID to STDF, because it was done in DEID item.
'                ''''For nonDEID, it should re-decode all ECID contents
'                Call auto_EcidPrintData(1, SingleBitArrayStr, False)
'                Call auto_EcidPrintData(2, SingleBitArrayStr, False)
'            End If
'
'            ''''print the Bitmap
'            Call auto_PrintAllBitbyDSSC(SingleBitArray(), EcidReadCycle, ECIDTotalBits, ECIDBitPerCycle)
'        End If
'
'        ''''' 20161003 ADD CRC
'        If (Trim(gS_ECID_CRC_Stage) <> "") Then
'            gS_ECID_CRC_HexStr(Site) = auto_ECID_CRC2HexStr(DoubleBitArray(), gL_ECID_CRC_EndBit) '2016.09.24. Add
'        End If
'        '========================================================================
'        '=  Add a gating rule for eFuse consolidation                           =
'        '=  The first letter has to be [A-Z], the 2nd to 6th letters have to    =
'        '=  be [0-9] or [A-Z]. Besides, the summation for bit6~53th has to be   =
'        '=  greater than 3. And the reserved area has to be all 0s (blank)      =
'        '========================================================================
'        Call auto_printECIDContent_AllBits
'
'        ChkResult(Site) = auto_Chk_ECID_Content_AllBits(mstr_ECID_effbit)
'
'        TheExec.Flow.TestLimit resultVal:=ChkResult, lowVal:=1, hiVal:=1, Tname:="ECID_Syntax_Chk_All"
'        TheExec.Flow.TestLimit resultVal:=gL_ECID_FBC, lowVal:=0, hiVal:=EcidHiLimitSingleDoubleBitCheck, Tname:="FailBitCount"
'
'        If (gS_JobName = "cp1") Then
'            ''' 20170623 add
'            TheExec.Datalog.WriteComment ""
'            TheExec.Datalog.WriteComment "ECID bit string for CRC calculation on site " & CStr(Site) & " (MSB to LSB) = " & mstr_ECID_bits_for_CRC
'            TheExec.Datalog.WriteComment ""
'        End If
'    Next Site  'For Each Site In TheExec.Sites
'
'    Call UpdateDLogColumns__False
'
'    ''''20170111 Add
'    Call auto_eFuse_ReadAllData_to_DictDSPWave("ECID", False, False)

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function ECID_Info()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "ECID_Info"

    Dim site As Variant
    Dim eFuseLotID() As String
    Dim eFuseWaferID() As String
    Dim eFuseDieX() As String
    Dim eFuseDieY() As String
    Dim tmpStr As String
    Dim compare As New SiteLong
''''    Dim max_x As Long
''''    Dim max_y As Long
''''    max_x = 50
''''    max_y = 50
    compare = compare.bitwiseand(0)
    
    ''''Call RegKeySave("eFuseLotNumber", "N12345,N12345,N12345,N12345,N12345,N12345,N12345,N12345")
    ''''Call RegKeySave("eFuseWaferID", "12,12,12,12,12,12,12,12")
    ''''Call RegKeySave("eFuseDieX", "1,2,3,4,5,6,7,8")
    ''''Call RegKeySave("eFuseDieY", "1,2,3,4,5,6,7,8")
    tmpStr = RegKeyRead("eFuseWaferID")
    eFuseWaferID = Split(tmpStr, ",")
    tmpStr = RegKeyRead("eFuseDieX")
    eFuseDieX = Split(tmpStr, ",")
    tmpStr = RegKeyRead("eFuseDieY")
    eFuseDieY = Split(tmpStr, ",")
    tmpStr = RegKeyRead("eFuseLotNumber")
    eFuseLotID = Split(tmpStr, ",")
    LotID = RegKeyRead("LotID")
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment ("Site " + CStr(site) + " " + eFuseLotID(site) + "-" + eFuseWaferID(site) + " X" + eFuseDieX(site) + "Y" + eFuseDieY(site))
        If eFuseLotID(site) = LotID Then
            compare = 1
        Else
            compare = 0
        End If
        If (TheExec.TesterMode = testModeOffline) Then compare = 1 ''''offline, set it to pass
    Next site
    TheExec.Datalog.WriteComment ""

    TheExec.Flow.TestLimit compare, 1, 1

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20160106 Update
Public Function auto_CleanRegData_New()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CleanRegData_New"

    Dim k As Long
    Dim siteNCnt As Long
    Dim tmpStr As String
    Dim site As Variant
    
    siteNCnt = TheExec.sites.Existing.Count
    
    tmpStr = ""
    If siteNCnt > 1 Then
        For k = 1 To siteNCnt - 1
            tmpStr = tmpStr + ","
        Next k
    End If
    
    Call RegKeySave("eFuseLotNumber", tmpStr)
    Call RegKeySave("eFuseWaferID", tmpStr)
    Call RegKeySave("eFuseDieX", tmpStr)
    Call RegKeySave("eFuseDieY", tmpStr)
    Call RegKeySave("eFuseIDSSOC", tmpStr)
    Call RegKeySave("eFuseIDSCPU", tmpStr)
    Call RegKeySave("eFuseIDSFIXED", tmpStr)
    Call RegKeySave("eFuseIDSGPU", tmpStr)
    Call RegKeySave("IDSSRAMSOC", tmpStr)
    Call RegKeySave("IDSSRAMCPU1", tmpStr)
    Call RegKeySave("IDSSRAMCPU2", tmpStr)
    Call RegKeySave("eFuseSpareParameter1", tmpStr)
    Call RegKeySave("eFuseSpareParameter2", tmpStr)
    Call RegKeySave("eFuseSpareParameter3", tmpStr)
    Call RegKeySave("eFuseSpareParameter4", tmpStr)
    Call RegKeySave("Hram_ECID_53bit", tmpStr)
    Call RegKeySave("Hram_DVFM_64bit", tmpStr)
    Call RegKeySave("Hram_BinData_46bit", tmpStr)
    Call RegKeySave("Hram_IDS_37bit", tmpStr)
    Call RegKeySave("eFuseSOCTRIM1", tmpStr)
    Call RegKeySave("eFuseSOCTRIM2", tmpStr)
    Call RegKeySave("eFuseBINFIXEDMD1", tmpStr)
    Call RegKeySave("eFuseBINGPUMD1", tmpStr)
    Call RegKeySave("eFuseBINGPUMD2", tmpStr)
    Call RegKeySave("eFuseBINGPUMD3", tmpStr)
    Call RegKeySave("eFuseBINGPUMD4", tmpStr)
    Call RegKeySave("eFuseBINSOCMD1", tmpStr)
    Call RegKeySave("eFuseBINSOCMD2", tmpStr)
    Call RegKeySave("eFuseIDSSRAMSOC", tmpStr)
    Call RegKeySave("eFuseIDSSRAMCPU1", tmpStr)
    Call RegKeySave("eFuseIDSSRAMCPU2", tmpStr)
    Call RegKeySave("HardBinName", "") ''tmpStr
    Call RegKeySave("SoftBinName", "") ''tmpStr
    
    Call RegKeySave("tmps0", tmpStr)
    Call RegKeySave("tmps1", tmpStr)
    Call RegKeySave("tmps2", tmpStr)
    Call RegKeySave("tmps3", tmpStr)
    Call RegKeySave("tmps4", tmpStr)

    If (gB_findCFGCondTable_flag = True) Then
        Call RegKeySave("SVM_CFuse_288Bits", tmpStr) ''''20171103 add SVM_CFuse_288Bits
    ElseIf (gB_findCFGTable_flag = True) Then
        Call RegKeySave("CFG_First_64Bits", tmpStr)
    End If
        
    Call UpdateDLogColumns(30)

    ''''if the TP does NOT run the TestInstance "eFuse_Initialize"
    If (Trim(gS_JobName) = "") Then
        gS_JobName = LCase(TheExec.CurrentJob)
    End If

    ''''For all NOT CP/WLFT Jobs, reset datalog X,Y Coordinates to N/A (-32768).
    If ((gS_JobName Like "*cp*") Or (gS_JobName Like "*wlft*")) Then
        TheExec.Flow.TestLimit 1, 1, 1, Tname:="CleanRegData", PinName:=UCase(gS_JobName)
    Else
        If TheExec.Flow.EnableWord("FT_SIM") = True Then
            ''''it's Allen's functioality, doing the FT simulation on CP environment (engineer).
            'Do not Reset X,Y to -32768(N/A)
            TheExec.Flow.TestLimit 1, 1, 1, Tname:="CleanRegData", PinName:=UCase(gS_JobName)
        Else
            ''multisite for FT, initialize XY to N/A(-32768) for all sites
            ''Write X and Y coordinates to FT STDF file
            ''After running ECID, it will have the correct XY information.

            If (TheExec.TesterMode = testModeOffline) Then ''''20160202 for the simulation
                For Each site In TheExec.sites
                    XCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetXCoord(site)
                    YCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetYCoord(site)
                    If (XCoord(site) = -32768 Or YCoord(site) = -32768) Then
                        Call setXY(5, 6) ''''set a pseudo XY coordinate
                        TheExec.Datalog.WriteComment vbTab & "Call setXY(5, 6) (pseudo XY_Coordinate)"
                        XCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetXCoord(site)
                        YCoord(site) = TheExec.Datalog.Setup.WaferSetup.GetYCoord(site)
                    End If
                Next site
            Else
                ''''<MUST> Very Important
                For Each site In TheExec.sites.Existing
                    Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(site, "-32768")
                    Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(site, "-32768")
                Next site
            End If

            TheExec.Flow.TestLimit 1, 1, 1, Tname:="CleanRegData_XY", PinName:=UCase(gS_JobName)
        End If
    End If

    Call UpdateDLogColumns__False

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20161121 Add for gereral Function Test with the (.pat/.pat.gz) in datalog
Public Function auto_Function_Test(patset As Pattern, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Function_Test"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim patt As String
    If (auto_eFuse_PatSetToPat_Validation(patset, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    Call TheHdw.Patterns(patt).Test(pfAlways, 0)
    DebugPrintFunc patset.Value

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20170106 New Function
''''<NOTICE> Default is MSBFirst=False, DSPWave element(0) is LSB bit.
Public Function auto_eFuse_ReadAllData_to_DictDSPWave(ByVal FuseType As String, Optional MSBFirst As Boolean = False, Optional showPrint As Boolean = False) As Long
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_ReadAllData_to_DictDSPWave"

    Dim site As Variant
    Dim i As Long, j As Long
    Dim m_catename As String
    Dim m_dlogstr As String
    Dim m_Fusetype As String
    Dim m_keyname As String
    Dim m_bitstr As String
    Dim m_tailStr As String
    Dim m_bitwidth As Long

    Dim m_debugPrint As Boolean
    Dim m_dspWave() As New DSPWave

    If (MSBFirst) Then
        m_tailStr = " [MSB...LSB]"
    Else
        m_tailStr = " [LSB...MSB]"
    End If

    m_Fusetype = UCase(Trim(FuseType))

    m_debugPrint = False ''''debug print with function GetStoredCaptureData()
    If (showPrint) Then TheExec.Datalog.WriteComment ""
    
    If (FuseType = "ECID") Then
        ReDim m_dspWave(UBound(ECIDFuse.Category))
        For i = 0 To UBound(ECIDFuse.Category)
            m_catename = ECIDFuse.Category(i).Name
            m_bitwidth = ECIDFuse.Category(i).BitWidth
            
            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = ECIDFuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = ECIDFuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -36)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -10) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i

    ElseIf (FuseType = "CFG") Then
        ReDim m_dspWave(UBound(CFGFuse.Category))
        For i = 0 To UBound(CFGFuse.Category)
            m_catename = CFGFuse.Category(i).Name
            m_bitwidth = CFGFuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = CFGFuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = CFGFuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -36)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -10) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i
        
    ElseIf (FuseType = "UID") Then
        ReDim m_dspWave(UBound(UIDFuse.Category))
        For i = 0 To UBound(UIDFuse.Category)
            m_catename = UIDFuse.Category(i).Name
            m_bitwidth = UIDFuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = UIDFuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = UIDFuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -36)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -10) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i

    ElseIf (FuseType = "UDR") Then
        ReDim m_dspWave(UBound(UDRFuse.Category))
        For i = 0 To UBound(UDRFuse.Category)
            m_catename = UDRFuse.Category(i).Name
            m_bitwidth = UDRFuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = UDRFuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = UDRFuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -36)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -10) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i

    ElseIf (FuseType = "SEN") Then
        ReDim m_dspWave(UBound(SENFuse.Category))
        For i = 0 To UBound(SENFuse.Category)
            m_catename = SENFuse.Category(i).Name
            m_bitwidth = SENFuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = SENFuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = SENFuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -36)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -10) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i
        
    ElseIf (FuseType = "MON") Then
        ReDim m_dspWave(UBound(MONFuse.Category))
        For i = 0 To UBound(MONFuse.Category)
            m_catename = MONFuse.Category(i).Name
            m_bitwidth = MONFuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = MONFuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = MONFuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -36)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -10) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i

    ElseIf (FuseType = "CMP") Then
        ReDim m_dspWave(UBound(CMPFuse.Category))
        For i = 0 To UBound(CMPFuse.Category)
            m_catename = CMPFuse.Category(i).Name
            m_bitwidth = CMPFuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = CMPFuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = CMPFuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -35)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -1) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i

    ElseIf (FuseType = "UDRE") Then
        ReDim m_dspWave(UBound(UDRE_Fuse.Category))
        For i = 0 To UBound(UDRE_Fuse.Category)
            m_catename = UDRE_Fuse.Category(i).Name
            m_bitwidth = UDRE_Fuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = UDRE_Fuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = UDRE_Fuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -36)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -10) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i

    ElseIf (FuseType = "UDRP") Then
        ReDim m_dspWave(UBound(UDRP_Fuse.Category))
        For i = 0 To UBound(UDRP_Fuse.Category)
            m_catename = UDRP_Fuse.Category(i).Name
            m_bitwidth = UDRP_Fuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = UDRP_Fuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = UDRP_Fuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -36)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -10) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i

    ElseIf (FuseType = "CMPE") Then
        ReDim m_dspWave(UBound(CMPE_Fuse.Category))
        For i = 0 To UBound(CMPE_Fuse.Category)
            m_catename = CMPE_Fuse.Category(i).Name
            m_bitwidth = CMPE_Fuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = CMPE_Fuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = CMPE_Fuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -35)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -1) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i

    ElseIf (FuseType = "CMPP") Then
        ReDim m_dspWave(UBound(CMPP_Fuse.Category))
        For i = 0 To UBound(CMPP_Fuse.Category)
            m_catename = CMPP_Fuse.Category(i).Name
            m_bitwidth = CMPP_Fuse.Category(i).BitWidth

            Call m_dspWave(i).CreateConstant(0, m_bitwidth, DspLong)
            For Each site In TheExec.sites
                If (MSBFirst) Then
                    m_bitstr = CMPP_Fuse.Category(i).Read.BitStrM(site)
                Else
                    m_bitstr = CMPP_Fuse.Category(i).Read.BitStrL(site)
                End If
                For j = 1 To Len(m_bitstr)
                    m_dspWave(i).Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
                Next j
                If (showPrint) Then
                    m_Fusetype = FormatNumeric(FuseType, -5)
                    m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadAllData_to_DictDSPWave", -35)
                    m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 25) + " = " + FormatNumeric(m_bitstr, -1) + m_tailStr
                    TheExec.Datalog.WriteComment m_dlogstr
                End If
            Next site
            m_keyname = FuseType + "_" + m_catename
            Call AddStoredCaptureData(m_keyname, m_dspWave(i))

            ''''Debug
            Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)
        Next i

    Else
        TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,MON,SEN,CMP,UDRE,UDRP,CMPE,CMPP)"
        GoTo errHandler
        ''''nothing
    End If

    If (False) Then
        Call UpdateDLogColumns(30)
        TheExec.Flow.TestLimit 1, 1, 1
        Call UpdateDLogColumns__False
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20170110 New Function
''''<NOTICE> Default is MSBFirst=False, DSPWave element(0) is LSB bit.
Public Function auto_eFuse_ReadData_to_DictDSPWave_byCategory(ByVal FuseType As String, m_catename As String, Optional MSBFirst As Boolean = False, Optional showPrint As Boolean = True) As Long
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_ReadData_to_DictDSPWave_byCategory"

    Dim i As Long
    Dim j As Long
    Dim m_len As Long
    Dim m_decimal As Long
    Dim m_dlogstr As String
    Dim m_Fusetype As String
    Dim m_keyname As String
    Dim m_bitstr As String
    Dim m_tailStr As String
    Dim site As Variant

    Dim m_dspWave As New DSPWave

    FuseType = UCase(Trim(FuseType))
    m_len = Len(m_catename) + 2

    If (MSBFirst) Then
        m_tailStr = " [MSB...LSB]"
    Else
        m_tailStr = " [LSB...MSB]"
    End If

    If (showPrint) Then TheExec.Datalog.WriteComment ""
    Call m_dspWave.Clear

    For Each site In TheExec.sites

        If (FuseType = "ECID") Then
            i = ECIDIndex(m_catename)
            If (MSBFirst) Then
                m_bitstr = ECIDFuse.Category(i).Read.BitStrM(site)
            Else
                m_bitstr = ECIDFuse.Category(i).Read.BitStrL(site)
            End If

        ElseIf (FuseType = "CFG") Then
            i = CFGIndex(m_catename)
            If (MSBFirst) Then
                m_bitstr = CFGFuse.Category(i).Read.BitStrM(site)
            Else
                m_bitstr = CFGFuse.Category(i).Read.BitStrL(site)
            End If

        ElseIf (FuseType = "UID") Then
            i = UIDIndex(m_catename)
            If (MSBFirst) Then
                m_bitstr = UIDFuse.Category(i).Read.BitStrM(site)
            Else
                m_bitstr = UIDFuse.Category(i).Read.BitStrL(site)
            End If

        ElseIf (FuseType = "UDR") Then
            i = UDRIndex(m_catename)
            If (MSBFirst) Then
                m_bitstr = UDRFuse.Category(i).Read.BitStrM(site)
            Else
                m_bitstr = UDRFuse.Category(i).Read.BitStrL(site)
            End If

        ElseIf (FuseType = "SEN") Then
            i = SENIndex(m_catename)
            If (MSBFirst) Then
                m_bitstr = SENFuse.Category(i).Read.BitStrM(site)
            Else
                m_bitstr = SENFuse.Category(i).Read.BitStrL(site)
            End If

        ElseIf (FuseType = "MON") Then
            i = MONIndex(m_catename)
            If (MSBFirst) Then
                m_bitstr = MONFuse.Category(i).Read.BitStrM(site)
            Else
                m_bitstr = MONFuse.Category(i).Read.BitStrL(site)
            End If

        ElseIf (FuseType = "CMP") Then
            i = CMPIndex(m_catename)
            If (MSBFirst) Then
                m_bitstr = CMPFuse.Category(i).Read.BitStrM(site)
            Else
                m_bitstr = CMPFuse.Category(i).Read.BitStrL(site)
            End If

        Else
            TheExec.Datalog.WriteComment funcName + ":: Please have a correct Fuse type (ECID,CFG,UID,UDR,MON,SEN,CMP)"
            GoTo errHandler
            ''''nothing
        End If
        
        Call m_dspWave.CreateConstant(0, Len(m_bitstr), DspLong)
        For j = 1 To Len(m_bitstr)
            m_dspWave.Element(j - 1) = CLng(Mid(m_bitstr, j, 1))
        Next j

        ''showPrint = True
        If (showPrint) Then
            m_Fusetype = FormatNumeric(FuseType, 3)
            m_Fusetype = m_Fusetype + FormatNumeric("Fuse ReadData_to_DictDSPWave_byCategory", -40)
            m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, 20) + " = " + FormatNumeric(m_bitstr, -1) + m_tailStr
            TheExec.Datalog.WriteComment m_dlogstr
        End If

    Next site

    m_keyname = FuseType + "_" + m_catename
    Call AddStoredCaptureData(m_keyname, m_dspWave)

    ''''Debug
    Dim m_debugPrint As Boolean
    m_debugPrint = True
    Call auto_eFuse_Print_GetStoredCaptureData(m_keyname, MSBFirst, m_debugPrint)

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20170106 used for the debug
''''Default MSBFirst is False
Public Function auto_eFuse_Print_GetStoredCaptureData(m_keyname As String, Optional MSBFirst As Boolean = False, Optional showPrint As Boolean = True) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_Print_GetStoredCaptureData"
    
    Dim site As Variant
    Dim i As Long, j As Long
    Dim m_debugPrint As Boolean
    Dim m_dspWaveRTN As New DSPWave
    Dim m_dlogstr As String

    If (showPrint) Then
        m_dspWaveRTN.Clear ''''<BeCarefully> it can not be used here.
        m_dspWaveRTN = GetStoredCaptureData(m_keyname)
    
        For Each site In TheExec.sites
            m_dlogstr = ""
            For j = 0 To m_dspWaveRTN.SampleSize - 1
                If (j = 0) Then
                    m_dlogstr = CStr(m_dspWaveRTN.Element(j))
                Else
                    m_dlogstr = m_dlogstr + "," + CStr(m_dspWaveRTN.Element(j))
                End If
            Next j
            If (MSBFirst) Then
                m_dlogstr = vbTab & "Site(" & site & ") GetStoredCaptureData:: " + UCase(m_keyname) + " = [MSB] " + m_dlogstr + " [LSB]"
            Else
                m_dlogstr = vbTab & "Site(" & site & ") GetStoredCaptureData:: " + UCase(m_keyname) + " = [LSB] " + m_dlogstr + " [MSB]"
            End If
            TheExec.Datalog.WriteComment m_dlogstr
        Next site
        TheExec.Datalog.WriteComment ""
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

''''20151222 New
''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
''''20170220 New, ECID Read Decode
Public Function auto_ECIDRead_Decode(ReadPatSet As Pattern, PinRead As PinList, Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ECIDRead_Decode"

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
    Dim ResultFlag As New SiteLong
    Dim testName As String
    Dim SingleStrArray() As String
    Dim SingleBitArray() As Long
    Dim DoubleBitArray() As Long
    
    Dim SingleBitArrayStr() As String

    Dim SignalCap As String, CapWave As New DSPWave
    Dim allBlank As New SiteBoolean
    'Dim blank_stage As New SiteBoolean
    ReDim gL_ECID_Sim_FuseBits(TheExec.sites.Existing.Count - 1, ECIDTotalBits - 1) ''''20161107 update, it's for the simulation

    SignalCap = "SignalCapture" 'define capture signal name

    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    '****** Initialize Site Varaible ******
    Dim m_siteVar As String
    m_siteVar = "ECIDChk_Var"
''    For Each Site In TheExec.Sites.Existing
''        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = -1
''    Next Site

    '================================================
    '=  Setup HRAM/DSSC capture cycles              =
    '================================================
    '*** Setup HARM/DSSC Trigger and Capture parameter ***
    auto_eFuse_DSSC_DigCapSetup ReadPatt, PinRead, SignalCap, EcidReadCycle, CapWave  'setup
    Call TheHdw.Patterns(ReadPatt).Test(pfAlways, 0, tlResultModeDomain)   'run read pattern and capture
    'auto_eFuse_DSSC_ReadDigCap_32bits EcidReadCycle, PinRead.Value, SingleStrArray, CapWave, allblank 'read back to singlestrarray

    Call UpdateDLogColumns(30)
    
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    ''''----------------------------------------------------
    ''''201808XX New Method by DSPWave
    ''''----------------------------------------------------
    'Dim mW_SingleBitWave As New DSPWave
    'Dim mW_DoubleBitWave As New DSPWave
    
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong
    Dim blank_early As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    
    m_Fusetype = eFuse_ECID
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_early = True

''    ''''<MUST> Clear Memory
''    Set gW_ECID_Read_singleBitWave = Nothing
''    Set gW_ECID_Read_doubleBitWave = Nothing

    'If (gS_JobName = "cp1") Then gS_JobName = "cp1_early" ''''<MUST>
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=0 (Stage Early Bits)
    'If (TheExec.TesterMode = testModeOffline) Then gL_eFuse_Sim_Blank = 1
    
    If (gS_JobName <> "cp1") Then gL_eFuse_Sim_Blank = 1
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, CapWave, m_FBC, blank_early, allBlank)
    ''''----------------------------------------------------

    If (blank_early.Any(False) = True) Then
        ''''if there is any site which is non-blank, then decode to gDW_ECID_Read_Decimal_Cate
        ''''if the element = -9999 means that its bits exceed 32 bits and bitSum<>0.
        ''''it will be present by Hex and Binary compare with the limit later on.
        ''Call auto_eFuse_setReadData(eFuse_ECID, gDW_ECID_Read_DoubleBitWave, gDW_ECID_Read_Decimal_Cate)
        'Call auto_eFuse_setReadData(eFuse_ECID)
        
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(eFuse_ECID)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(eFuse_ECID)
    
''        If (gS_JobName Like "cp*" And gB_ReadWaferData_flag = True) Then
''        End If
        
        Call auto_eFuse_print_DSSCReadWave_BitMap(eFuse_ECID, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(eFuse_ECID, False, gB_eFuse_printReadCate)
        
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
'    m_SiteVarValue = blank_early
'    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
'    For Each Site In TheExec.Sites
'        If (auto_eFuse_GetAllPatTestPass_Flag("ECID") = True And m_FBC = 0) Then
'            m_ResultFlag = 0  'Pass Blank check criterion
'        Else
'            m_ResultFlag = 1  'Fail Blank check criterion
'            m_SiteVarValue = 0
'        End If
'
'        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = m_SiteVarValue(Site)
'    Next Site
    gL_ECID_FBC = m_FBC
    
    ''''if blank_early(Site)=False, check if read DEID bits are same as Prober (CP1 only)
    ''''When blank=False, We do NOT need to do that because it will be checked in Syntax check again.

    Call UpdateDLogColumns(gI_ECID_catename_maxLen)

'    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "FailBitCount_CP1_Early" ''+ UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

'    testName = "ECID_Blank_CP1_Early" '' + UCase(gS_JobName)
'    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    DebugPrintFunc ReadPatSet.Value
    
Exit Function

End If

    
    
    
    

'    testName = "ECIDRead_DeCode_" + UCase(gS_JobName)
'
'    For Each Site In TheExec.Sites
'        ''''if the BitOrder is LSB in the PinMap sheet, so we do the reverse here.
'        If (gC_eFuse_DigCap_BitOrder = "LSB") Then
'            For i = 0 To EcidReadCycle - 1
'                SingleStrArray(i, Site) = StrReverse(SingleStrArray(i, Site))
'            Next i
'        End If
'
'        ''''=============== Start of Simulated Data ===============
'        If (TheExec.TesterMode = testModeOffline) Then
'            allblank(Site) = False
'            blank_stage(Site) = True ''''True or False (re-test)
'
'            'If (gS_JobName <> "cp1") Then
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                TheExec.Datalog.WriteComment vbTab & "[ Simulate Category (m_stage <= Job) ]"
'                Call eFuseENGFakeValue_Sim
'
'                Dim m_tmpStr As String
'                Dim Expand_eFuse_Pgm_Bit() As Long, eFuse_Pgm_Bit() As Long
'                ReDim eFuse_Pgm_Bit(ECIDTotalBits - 1)
'                ReDim Expand_eFuse_Pgm_Bit(ECIDTotalBits * EcidWriteBitExpandWidth - 1)
'
'                If (True) Then
'                    ''''<MUST>
'                    ''''update here because they will be clear after non-DEID
'                    Call auto_eFuse_SetWriteDecimal("ECID", "Wafer_ID", WaferID, False, False)
'                    Call auto_eFuse_SetWriteDecimal("ECID", "X_Coordinate", XCoord(Site), False, False)
'                    Call auto_eFuse_SetWriteDecimal("ECID", "Y_Coordinate", YCoord(Site), False, False)
'                End If
'
'                Call auto_EcidPgmBit_DEID_forCheck(Expand_eFuse_Pgm_Bit(), eFuse_Pgm_Bit(), False)
'                ''''to composite gS_SingleStrArray(i, Site) for the simulation in SingleDoubleBit_Check/Syntax
'                For i = 0 To EcidReadCycle - 1
'                    m_tmpStr = ""
'                    For j = 0 To EcidReadBitWidth - 1
'                        k = j + i * EcidReadBitWidth ''''MUST
'                        m_tmpStr = CStr(eFuse_Pgm_Bit(k)) + m_tmpStr
'                        gL_ECID_Sim_FuseBits(Site, k) = eFuse_Pgm_Bit(k) ''''20161107 update
'                    Next j
'                    ''gS_SingleStrArray(i, Site) = m_tmpStr
'                    SingleStrArray(i, Site) = m_tmpStr ''''use local variable only
'                Next i
'            'End If
'        End If
'        ''''===============   End of Simulated Data ===============
'
'        Call auto_OR_2Blocks("ECID", SingleStrArray, SingleBitArray, DoubleBitArray)
'
'        ReDim SingleBitArrayStr(UBound(SingleBitArray)) ''''<MUST>
'        For i = 0 To UBound(SingleBitArray)
'            SingleBitArrayStr(i) = CStr(SingleBitArray(i))
'        Next i
'
'        '====================================================
'        '=  Print the all eFuse Bit data from DSSC          =
'        '====================================================
'        TheExec.Datalog.WriteComment ""
'        TheExec.Datalog.WriteComment "Read All ECID eFuse bits from DSSC at Site (" & CStr(Site) & ")"
'        Call auto_PrintAllBitbyDSSC(SingleBitArray, EcidReadCycle, ECIDTotalBits, EcidReadBitWidth)
'
'        Call auto_EcidPrintData(1, SingleBitArrayStr, False, False)
'        Call auto_printECIDContent_AllBits
'
'        gB_ECID_decode_flag(Site) = False ''''reset
'
'        ''Binning out
'        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=testName
'
'        If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
'
'    Next Site
'
'    ''''20170111 Add
'    Call auto_eFuse_ReadAllData_to_DictDSPWave("ECID", False, False)
'
'    Call UpdateDLogColumns__False
'    DebugPrintFunc ReadPatSet.Value
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function


Public Function EFUSE_Resistance(patset As Pattern, PwrPin As String, vpwr As Double, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler

Dim result As New SiteDouble
Dim funcName As String:: funcName = "auto_EFUSE_Resistance"
Dim current_data As New PinListData
Dim site As Variant
Dim power_pin_arr() As String
Dim power_pin_number As Long
Dim i As Long
Dim temp_power_pin As String
        
    If Validating_ Then
        Call PrLoadPattern(patset.Value)
        Exit Function    ' Exit after validation
    End If
    
    Call TheExec.DataManager.DecomposePinList(PwrPin, power_pin_arr, power_pin_number)
    
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    TheExec.Datalog.WriteComment "Setting: " + CStr(vpwr) + " V"
 
    Call TurnOnEfusePwrPins(PwrPin, vpwr)
    TheHdw.Wait 0.001

    Call HardIP_InitialSetupForPatgen
    Call TheHdw.Patterns(patset).start
    TheHdw.DCVS.Pins(PwrPin).SetCurrentRanges 0.02, 0.02
    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
    
    TheHdw.Wait 0.01
    
    current_data = TheHdw.DCVS.Pins(PwrPin).Meter.Read.Math.Multiply(1000)
    
    For i = 0 To power_pin_number - 1
        temp_power_pin = power_pin_arr(i)
        For Each site In TheExec.sites
        
            If (current_data.Pins(temp_power_pin).Value < 0.0000001) Then
                current_data.Pins(temp_power_pin).Value = 0.0000001
            End If
            
            result = vpwr / current_data.Pins(temp_power_pin).Value
            '''theexec.Datalog.WriteComment "Site" + CStr(site) + ": " + temp_power_pin + " " + FormatNumber(CStr(current_data.Pins(temp_power_pin).Value), 3) + " mA" + " " + FormatNumber(CStr(result), 3) + " Kohm"
        Next
        
        TheExec.Flow.TestLimit resultVal:=result.Multiply(1000), lowVal:=140, hiVal:=330, Tname:=temp_power_pin
    Next
    
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end
    Call TurnOffEfusePwrPins(PwrPin, vpwr)
    
    TheExec.Datalog.WriteComment ""

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
    
    
End Function



