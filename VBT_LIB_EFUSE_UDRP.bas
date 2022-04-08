Attribute VB_Name = "VBT_LIB_EFUSE_UDRP"
Option Explicit

''''20160805 update for UDRP_Ver1 UDRP CMPP Fuse
Public Function auto_CMPP_Syntax_Chk(CMPP_pat As Pattern, OutPin As PinList, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CMPP_Syntax_Chk"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(CMPP_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_CMPP As New DSPWave
    Dim CMPP_BitStr As New SiteVariant
    Dim i As Long, j As Long, k As Long
    Dim PatCMPPArray() As String
    Dim pat_count As Long, Status As Boolean
    Dim fail_flag As Boolean
    ''Dim Str1 As String
    ''Dim TstNumArray() As Long, TstNumCount As Long
    Dim DigCapArray() As Long
    ''Dim ResultFlag As New SiteLong
    Dim testName As String
    Dim PinName As String
    
    Dim CMPP_PrintRow As Long
    Dim CMPP_BitPerRow As Long
    Dim CMPP_CapBits As Long
    Dim CMPP_TotalBit As Long
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    ''Dim SiteVarValue As Long
    
    Dim m_catenameUDRP As String ''''UDRP_Fuse Category name
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_startbit As Long
    Dim m_endbit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant ''''20160901 update
    Dim m_value As Variant
    Dim m_bitsum As Long
    Dim m_bitStrM As String
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim m_lolmtV As Variant
    Dim m_hilmtV As Variant
    Dim m_binarr() As Long
    Dim tmpVbin As Double
    Dim tmpVfuse As String
    Dim tmpdlgStr As String
    Dim tmpStr1 As String
    Dim tmpStr As String
    Dim PrintSiteVarResult As String
    Dim step_vdd As Long
    Dim m_defvalUDRP As Variant ''''20160905
    Dim m_HexStr As String

    Dim Flag_CMPPCategoryMatch As Boolean
    
    Dim MaxLevelIndex As Long
    Dim m_catenameVbin As String
    Dim m_defreal As String
    Dim m_resolution As Double
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long

    Dim tmpStrL As String
    Dim TmpVal As Variant
    Dim m_testValue As Variant

    Dim vbinflag As Long
    Dim m_stage As String
    Dim m_tsname As String
    Dim m_siteVar As String
    'Dim m_hexStr As String
    ''Dim m_vddbinEnum As Long
    Dim m_Pmode As Long
    Dim m_unitType As UnitType
    Dim m_scale As tlScaleType

    m_siteVar = "UDR_PChk_Var"
    
    CMPP_CapBits = gL_CMPP_DigCapBits_Num
    CMPP_TotalBit = gL_CMPP_DigCapBits_Num

    ''''display bit number per row
    If (gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp") Then
        CMPP_BitPerRow = 32
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        CMPP_BitPerRow = 16
    End If

    ''''20150731 update
    CMPP_PrintRow = IIf((CMPP_TotalBit Mod CMPP_BitPerRow) > 0, Floor(CMPP_TotalBit / CMPP_BitPerRow) + 1, Floor(CMPP_TotalBit / CMPP_BitPerRow))
    ReDim DigCapArray((CMPP_PrintRow * CMPP_BitPerRow) - 1)

    TheHdw.Patterns(CMPP_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Status = GetPatListFromPatternSet(CMPP_pat.Value, PatCMPPArray, pat_count)
    
    
If (gB_eFuse_newMethod) Then
    
    TheExec.Datalog.WriteComment ""

    Call auto_eFuse_DSSC_DigCapSetup(PatCMPPArray(0), OutPin, "CMPP_cap", CMPP_CapBits, Trim_code_CMPP)
    Call TheHdw.Patterns(PatCMPPArray(0)).Test(pfAlways, 0)
    gStr_PatName = PatCMPPArray(0)
    Call UpdateDLogColumns(gI_CMPP_catename_maxLen)
        
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        gL_eFuse_Sim_Blank = 1
        Call auto_CMP_Sim_New(eFuse_CMPP, True)
'        Call RunDSP.eFuse_SingleBitWave2CapWave32Bits(eFuse_CMP, Trim_code_CMP)
'        Call auto_eFuse_print_capWave32Bits(eFuse_CMP, Trim_code_CMP, False) ''''True to print out
        For Each site In TheExec.sites
            Trim_code_CMPP = gDW_CMPP_Pgm_SingleBitWave.Copy
        Next
    End If
    
    ''''----------------------------------------------------
    ''''201812XX New Method by DSPWave
    ''''----------------------------------------------------
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong
    Dim m_cmpResult As New SiteLong

    m_Fusetype = eFuse_CMPP
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = False
    blank_stage = False
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
    Dim m_SerialType As Boolean:: m_SerialType = False
    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
        m_SerialType = True
        gDB_SerialType = True
    End If
    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_CMPP, m_FBC, blank_stage, allBlank, True)

    'Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, 1, Trim_code_USO, m_FBC, m_cmpResult)


    Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

    Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
    Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate, gStr_PatName)
    
    'Call auto_eFuse_CMPP_Parsing_HLlimit(m_Fusetype)
    
    Dim condstr As String:: condstr = ""
    Call auto_eFuse_SyntaxCheck(m_Fusetype, condstr)

Exit Function

End If
    
    
    
    

'    For j = 0 To pat_count - 1 'Actually, we have only one pattern
'        Call auto_eFuse_DSSC_DigCapSetup(PatCMPPArray(j), OutPin, "CMPP_cap", CMPP_CapBits, Trim_code_CMPP)
'        Call TheHdw.Patterns(PatCMPPArray(j)).test(pfAlways, 0)
'
'        Call UpdateDLogColumns(gI_CMPP_catename_maxLen)
'
'        For Each Site In TheExec.Sites
'
'            CMPP_BitStr(Site) = ""
'            gB_CMPP_decode_flag(Site) = False
'
'            For i = 0 To UBound(DigCapArray)
'               DigCapArray(i) = 0  ''''MUST be [LSB(0)......MSB(lastbit)]
'            Next i
'
'            ''''20150717 update
'            ''''composite to the CMPP_BitStr() from the DSSC Capture
'            If (UCase(gS_CMPP_PatBitOrder) = "MSB") Then
'                ''''<Notice>
'                ''''the pattern DSSC capture from MSB to LSB (bitLast to bit0)
'                ''''so Trim_code_CMPP.Element(0) is MSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To CMPP_CapBits - 1
'                    DigCapArray(i) = Trim_code_CMPP.Element(CMPP_CapBits - 1 - i)    ''''Reverse Bit String
'                    CMPP_BitStr(Site) = CMPP_BitStr(Site) & Trim_code_CMPP.Element(i) ''''[MSB ... LSB]
'                Next i
'            Else
'                ''''<Notice>
'                ''''the pattern DSSC capture from LSB to MSB (bit0 to bitLast)
'                ''''so Trim_code_CMPP.Element(0) is LSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To CMPP_CapBits - 1
'                    DigCapArray(i) = Trim_code_CMPP.Element(i)
'                    CMPP_BitStr(Site) = DigCapArray(i) & CMPP_BitStr(Site) ''''[MSB ... LSB]
'                Next i
'            End If
'
'            ''''20160324 update
'            ''''=============== Start of Simulated Data ===============
'            If (TheExec.TesterMode = testModeOffline) Then
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                CMPP_BitStr(Site) = auto_CMP_Sim_New(True) ''''<Notice> CMPP_BitStr MUST be [MSB ... LSB] [bitLast...bit0]
'                For i = 0 To CMPP_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(CMPP_BitStr(Site)), i + 1, 1))
'                Next i
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            '=======================================================
'            '= Print out the caputured bit data from DigCap        =
'            '=======================================================
'            TheExec.Datalog.WriteComment ""
'            Call auto_PrintAllBitbyDSSC(DigCapArray, CMPP_PrintRow, UBound(DigCapArray) + 1, CMPP_BitPerRow)
'
'            TheExec.Datalog.WriteComment "Site(" & Site & "), CMPP pat:" & PatCMPPArray(j) & ", Shift out code [" + CStr(CMPP_CapBits - 1) + ":0]=" + CMPP_BitStr(Site)
'            TheExec.Datalog.WriteComment ""
'
'            Dim Count_1s As Long
'            Count_1s = 0
'            For i = 1 To Len(CMPP_BitStr(Site))
'                If (CStr(Mid(CMPP_BitStr(Site), i, 1)) = "1") Then
'                    Count_1s = Count_1s + 1
'                End If
'            Next i
'
'            TheExec.Datalog.WriteComment "Site(" & Site & "), CMPP pat:" & PatCMPPArray(j) & " UDRPVer1 total 1s amount = " & CStr(Count_1s)
'
'            ''Str1 = TheExec.DataManager.InstanceName
'            ''Call TheExec.DataManager.GetTestNumbers(Str1, TstNumArray, TstNumCount) ''waste time in 9.0, 20180320 update
'
'            ''''20160901 update, here is decoding.
'            Call auto_Decode_CMPP_Binary_Data(DigCapArray)
'
'            ''''judge limit
'            For i = 0 To UBound(CMPP_Fuse.Category)
'                tmpdlgStr = ""
'                m_catename = CMPP_Fuse.Category(i).Name
'                m_algorithm = LCase(CMPP_Fuse.Category(i).Algorithm)
'                m_startbit = CMPP_Fuse.Category(i).SeqStart
'                m_endbit = CMPP_Fuse.Category(i).SeqEnd
'                m_bitwidth = CMPP_Fuse.Category(i).Bitwidth
'                m_decimal = CMPP_Fuse.Category(i).Read.Decimal(Site)
'                m_value = CMPP_Fuse.Category(i).Read.Value(Site)
'                m_hexStr = CMPP_Fuse.Category(i).Read.HexStr(Site)
'
'                ''''initialize everytime
'                Flag_CMPPCategoryMatch = False
'                m_lolmtV = -1
'                m_hilmtV = m_lolmtV
'
'                For k = 0 To UBound(UDRP_Fuse.Category)
'                    m_catenameUDRP = UCase(UDRP_Fuse.Category(k).Name)
'                    If (m_catenameUDRP = UCase(m_catename)) Then
'                        ''m_defvalUDRP = UDRP_Fuse.Category(k).DefaultValue ''''20160905
'                        m_lolmtV = UDRP_Fuse.Category(k).Read.Decimal(Site)
'                        m_hilmtV = m_lolmtV
'
'''''                        ''''Cayman speciacl case example
'''''                        If UCase(m_catename) = UCase("PLL_KVCO_Trim") Then 'V01J 160526 only
'''''                            m_lolmtV = 15
'''''                            m_hilmtV = m_lolmtV
'''''                        ElseIf (UCase(m_catename) = UCase("ADCLK_SCR2_vsns_cal_fuse")) Then
'''''                            m_lolmtV = 0
'''''                            m_hilmtV = m_lolmtV
'''''                        End If
'
'                        Flag_CMPPCategoryMatch = True
'                        Exit For
'                    End If
'                Next k
'
'                If Flag_CMPPCategoryMatch = False Then
'                    TheExec.Datalog.WriteComment ("====================================================================")
'                    TheExec.Datalog.WriteComment ("    Can't CMPP Category : " & m_catename & " in UDRP Category!!!")
'                    TheExec.Datalog.WriteComment ("====================================================================")
'                End If
'
'                ''''20170811 update
'                If (m_bitwidth >= 32) Then
'                    ''m_catename = m_catename + "_" + m_hexStr
'                    ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
'                    m_lolmtV = auto_Value2HexStr(m_lolmtV, m_bitwidth)
'                    m_hilmtV = auto_Value2HexStr(m_hilmtV, m_bitwidth)
'
'                    ''''------------------------------------------
'                    ''''compare with lolmt, hilmt
'                    ''''m_decimal 0 means fail
'                    ''''m_decimal 1 means pass
'                    ''''------------------------------------------
'                    m_decimal = auto_TestStringLimit(m_hexStr, CStr(m_lolmtV), CStr(m_hilmtV))
'                    m_lolmtV = 1
'                    m_hilmtV = 1
'                Else
'                    ''''20160620 update
'                    ''''20160927 update the new logical methodology for the unexpected binary decode.
'                    If (auto_isHexString(CStr(m_lolmtV)) = True) Then
'                        ''''translate to double value
'                        m_lolmtV = auto_HexStr2Value(m_lolmtV)
'                    Else
'                        ''''doNothing, m_lolmtV = m_lolmtV
'                    End If
'
'                    If (auto_isHexString(CStr(m_hilmtV)) = True) Then
'                        ''''translate to double value
'                        m_hilmtV = auto_HexStr2Value(m_hilmtV)
'                    Else
'                        ''''doNothing, m_hilmtV = m_hilmtV
'                    End If
'                End If
'
'                TheExec.Flow.TestLimit resultVal:=m_decimal, lowVal:=m_lolmtV, hiVal:=m_hilmtV, Tname:=m_catename
'            Next i
'            ''''end-------------------
'
'            If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
'
'        Next Site
'        Call UpdateDLogColumns__False
'    Next j
'
'    DebugPrintFunc CMPP_pat.Value
 
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDRP_USI(USI_pat As Pattern, InPin As PinList, Optional condstr As String = "stage", Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_USI"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim patt As String
    If (auto_eFuse_PatSetToPat_Validation(USI_pat, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_USI As New DSPWave
    Dim i As Long, j As Long, k As Long
    Dim pat_count As Long
    Dim Status As Boolean
    Dim PatUSIArray() As String
    Dim CheckEfuseVer As New SiteVariant
    
    Dim usiarrSize As Long
    usiarrSize = gL_UDRP_USI_DigSrcBits_Num * gC_UDRP_USI_DSSCRepeatCyclePerBit

    Dim PgmBitArr() As Long
    ReDim PgmBitArr(gL_UDRP_USI_DigSrcBits_Num - 1)

    Dim USI_Array() As Long
    ReDim USI_Array(TheExec.sites.Existing.Count - 1, usiarrSize - 1)
    
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant ''''was long, 20160608 update
    Dim m_bitStrM As String
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim tmpVbin As Variant ''''was long, 20160608 update
    Dim tmpVfuse As String
    Dim tmpdlgStr As String
    Dim tmpStr As String
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_bitsum As Long
    Dim m_stage As String
    Dim TmpVal As Variant
    Dim m_resolution As Double
    Dim m_USI_BitStr As New SiteVariant
    ''''------------------------------------------------------------

    TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName

    ''''20171016 update
    ''''--------------------------------
    Dim m_CP1_Early_Flag As Boolean
    condstr = LCase(condstr)

    
    If (gS_JobName = "cp1_early") Then
    'If (gS_JobName = "cp1" And condStr = "cp1_early") Then
        m_CP1_Early_Flag = True
        'gS_JobName = "cp1_early" ''''used to program the category with stage = "cp1_early"
    Else
        m_CP1_Early_Flag = False
    End If
    ''''--------------------------------
 
If (gB_eFuse_newMethod = True) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If
    
    ''''201808XX update
    If (TheExec.TesterMode = testModeOffline) Then
        If (condstr <> "cp1_early") Then
            For Each site In TheExec.sites
                Call auto_UDRP_USI_Sim(False, False) ''''True for print debug
                Call eFuseENGFakeValue_Sim
            Next site
        End If
    End If
    
    'Dim m_stage As String
    'Dim m_catename As String
    Dim m_catenameVbin As String
    Dim m_crc_idx As Long
    Dim m_calcCRC As New SiteLong
    
    Dim m_cmpStage As String
    Dim m_pgmRes As New SiteLong
    'Dim m_defreal As String
    'Dim m_algorithm As String
    'Dim m_resolution As Double
    Dim m_vbinResult As New SiteDouble
    Dim m_pgmWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_pgmDigSrcWave As New DSPWave
    
    m_pgmWave.CreateConstant 0, gDL_BitsPerBlock, DspLong ''''gDL_BitsPerBlock = EConfigBitPerBlockUsed
    ''''<MUST and Importance>
    For Each site In TheExec.sites
        gDW_Pgm_RawBitWave = m_pgmWave.Copy
    Next site
    
        ''''Only composite case "real or bincut" PgmBits Wave per Stage requirement
    For i = 0 To UBound(UDRP_Fuse.Category)
        With UDRP_Fuse.Category(i)
            m_stage = LCase(.Stage)
            m_algorithm = LCase(.algorithm)
            m_defreal = LCase(.Default_Real)
        End With
        
        If (m_stage = gS_JobName) Then ''''was If (m_stage = m_cmpStage) Then
'            If (m_algorithm = "crc") Then
'                m_crc_idx = i
'                ''''special handle on the next process
'                ''''skip it here
'            ElseIf (m_defreal = "real" Or m_defreal = "bincut") Then
            If (m_defreal = "real" Or m_defreal = "bincut") Then
                If (m_algorithm = "vddbin") Then
                    Call auto_eFuse_Vddbin_bincut_setWriteData(eFuse_UDRP, i)
                End If
                ''''---------------------------------------------------------------------------
                With UDRP_Fuse.Category(i)
                    ''''Here it will update the global DSP DSPWave "gDW_Pgm_RawBitWave" inside
                    Call rundsp.eFuse_updatePgmWave_byCategory(.BitIndexWave, .Write.BitArrWave)
                End With
                ''''---------------------------------------------------------------------------
            End If
        Else
            ''''doNothing
        End If
    Next i
    
    ''''composite effective PgmBits per Stage requirement
    m_pgmRes = 0
    If (condstr = "cp1_early") Then
        ''''condStr = "cp1_early"
        Call auto_eFuse_Gen_PgmBitSrcWave_onlyDeaultBits(eFuse_UDRP, m_pgmDigSrcWave, m_pgmRes)
    Else
        ''''condStr = "stage"
        ''''Here Parameter bitFlag_mode=1 (Stage Bits)
        m_calcCRC = 0
        Call rundsp.eFuse_Gen_PgmBitSrcWave(eFuse_UDRP, 1, m_pgmDigSrcWave, m_pgmRes)
    End If
    TheHdw.Wait 1# * ms
    
    Call UpdateDLogColumns(gI_UDRP_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=m_pgmRes, lowVal:=1, hiVal:=1, Tname:="UDRP_PGM_" + UCase(gS_JobName)
    Call UpdateDLogColumns__False

    ''''use global DSP variable
    Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_UDRP_Pgm_SingleBitWave, gDL_TotalBits, gDL_ReadCycles, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_UDRP, gDW_UDRP_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)
    
        Dim m_size As Long
        Dim m_tmpArr() As Long
        Dim m_outArr() As Long
        Dim m_tmpWave1 As New DSPWave
        Dim outWave As New DSPWave
        
        'outWave.CreateConstant 0, m_size, DspLong
        For Each site In TheExec.sites
            m_size = m_pgmDigSrcWave(site).SampleSize
        m_USI_BitStr(site) = ""
        Next
        
    If (gL_UDRP_USI_PatBitOrder = "LSB") Then
        For Each site In TheExec.sites
            m_tmpWave1(site) = m_pgmDigSrcWave(site).Copy.ConvertDataTypeTo(DspLong)
            m_tmpArr = m_tmpWave1(site).Data
            For i = 0 To m_size - 1
                m_USI_BitStr(site) = CStr(m_tmpArr(i)) + m_USI_BitStr(site)
            Next i
        Next
    Else
        outWave.CreateConstant 0, m_size, DspLong
        
        For Each site In TheExec.sites
            m_tmpWave1(site) = m_pgmDigSrcWave(site).Copy.ConvertDataTypeTo(DspLong)
            m_tmpArr = m_tmpWave1(site).Data
            m_outArr = outWave(site).Data
                For i = 0 To m_size - 1
            ''outWave.Element(i) = m_tmpWave.Element(m_size - i - 1) ''''waste TT
                m_outArr(i) = m_tmpArr(m_size - i - 1) ''''save TT
                 m_USI_BitStr(site) = m_USI_BitStr(site) + CStr(m_tmpArr(m_size - i - 1))
            Next i
        
        outWave(site).Data = m_outArr ''''save TT
        Next
    End If
    
    ''''--------------------------------------------------------------------------------------------
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    ''TheHdw.Patterns(USI_pat).Load

    Status = GetPatListFromPatternSet(USI_pat.Value, PatUSIArray, pat_count)

    For j = 0 To pat_count - 1
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "USI Pattern: " + PatUSIArray(j)
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "Site(" + CStr(site) + "), Shift In Code [" + CStr(UBound(PgmBitArr)) + ":0]=" + m_USI_BitStr(site)
        Next site

        Call eFuse_DSSC_SetupDigSrcWave(PatUSIArray(j), InPin, "USI_Src", outWave)
        'UDR_SetupDigSrcArray PatUSIArray(j), InPin, "USI_Src", usiarrSize, USI_Array
        Call TheHdw.Patterns(PatUSIArray(j)).Test(pfAlways, 0)
    Next j

    TheHdw.Wait 100# * us
    DebugPrintFunc USI_pat.Value

Exit Function
End If
    
    
    
    
    
    


    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDRP_USO_Syntax_Chk(Optional condstr As String = "all") As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_USO_Syntax_Chk"

'    ''''----------------------------------------------------------------------------------------------------
'    ''''<Important>
'    ''''Must be put before all implicit array variables, otherwise the validation will be error.
'    '==================================
'    '=  Validate/Load Read patterns   =
'    '==================================
'    ''''20161114 update to Validate/load pattern
'    Dim ReadPatt As String
'    If (auto_eFuse_PatSetToPat_Validation(USO_pat, ReadPatt, Validating_) = True) Then Exit Function
'    ''''----------------------------------------------------------------------------------------------------
'
'    Dim Site As Variant
'    Dim Trim_code_USO As New DSPWave
'    Dim USO_BitStr As New SiteVariant
'    Dim i As Long, j As Long, k As Integer
'    Dim PatUSOArray() As String
'    Dim pat_count As Long, Status As Boolean
'
'    Dim DigCapArray() As Long
'    Dim USO_PrintRow As Long
'    Dim USO_BitPerRow As Long
'    Dim USO_CapBits As Long
'    Dim USO_TotalBit As Long
'
'    USO_CapBits = gL_UDRP_USO_DigCapBits_Num
'    USO_TotalBit = gL_UDRP_USO_DigCapBits_Num
'
'    ''''display bit number per row
'    If (gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp") Then
'        USO_BitPerRow = 32
'    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
'        USO_BitPerRow = 16
'    End If
'
'    ''''20150731 update
'    USO_PrintRow = IIf((USO_TotalBit Mod USO_BitPerRow) > 0, Floor(USO_TotalBit / USO_BitPerRow) + 1, Floor(USO_TotalBit / USO_BitPerRow))
'    ReDim DigCapArray((USO_PrintRow * USO_BitPerRow) - 1)
'
'    Dim MaxLevelIndex As Long
'    Dim m_catenameVbin As String
'    Dim m_catename As String
'    Dim m_algorithm As String
'    Dim m_defreal As String
'    Dim m_resolution As Double
'    Dim m_LSBbit As Long
'    Dim m_MSBBit As Long
'    Dim m_bitwidth As Long
'    Dim m_decimal As Variant ''20160506 update, was Long
'    Dim m_value As Variant
'    Dim m_bitsum As Long
'    Dim m_bitStrM As String
'    Dim m_lolmt As Variant
'    Dim m_hilmt As Variant
'    Dim tmpdlgStr As String
'    Dim tmpStrL As String
'    Dim TmpStr As String
'    Dim tmpVal As Variant
'    Dim m_testValue As Variant
'    Dim step_vdd As Long
'    Dim vbinflag As Long
'    Dim m_stage As String
'    Dim m_tsName As String
'    Dim m_siteVar As String
'    Dim m_hexStr As String
'    ''Dim m_vddbinEnum As Long
'    Dim m_Pmode As Long
'    Dim m_unitType As UnitType
'    Dim m_scale As tlScaleType
'
'    Dim allblank As New SiteBoolean
'    Dim blank_stage As New SiteBoolean
'
'    Dim m_bitFlag_mode As Long
'
'    m_siteVar = "UDR_PChk_Var"
'
'    ''TheHdw.Patterns(USO_pat).Load
'    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'
'    Status = GetPatListFromPatternSet(USO_pat.Value, PatUSOArray, pat_count)
'    TheExec.Datalog.WriteComment ""
'
'    ''''20171016 update
'    ''''--------------------------------
'    Dim m_testFlag As Boolean
'    Dim m_CP1_Early_Flag As Boolean
    condstr = LCase(condstr)
'
'    'If (gS_JobName = "cp1" And condStr = "cp1_early") Then
'    If (gS_JobName = "cp1_early") Then
'        'm_CP1_Early_Flag = True
'        gS_JobName = "cp1_early" ''''used to syntax check the category with stage = "cp1_early"
'    Else
'        m_CP1_Early_Flag = False
'    End If
'    ''''--------------------------------
'
'If (gB_eFuse_newMethod) Then
'
'    ''''----------------------------------------------------
'    ''''201812XX New Method by DSPWave
'    ''''----------------------------------------------------
'    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType




    Dim site As Variant

    Dim i As Long, j As Long, k As Long

    ''Dim SingleStrArray() As String
    Dim SingleBitArray() As Long
    Dim DoubleBitArray() As Long

    Dim tmpStr As String

    m_Fusetype = eFuse_UDRP
'    m_FBC = -1               ''''initialize
'    m_ResultFlag = -1        ''''initialize
'    m_SiteVarValue = -1      ''''initialize
'    allblank = False
'    blank_stage = True
'
'    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
'
'    gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
'    Dim m_SerialType As Boolean:: m_SerialType = False
'    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
'        m_SerialType = True
'        gDB_SerialType = True
'    End If
'
'    Dim m_CompareFlag As Boolean:: m_CompareFlag = False
'
'    For Each Site In TheExec.sites.Active
'        If (TheExec.sites(Site).SiteVariableValue(m_siteVar) = 1) Then
'            m_CompareFlag = True
'            Exit For
'        End If
'    Next
'
'    If (m_CompareFlag = True) Then
'        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(0), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
'        Call TheHdw.Patterns(PatUSOArray(0)).test(pfAlways, 0)
'    End If
'
'    If (TheExec.TesterMode = testModeOffline) Then
'        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
'        ''''Here MarginRead it only read the programmed Site.
'        If (gS_JobName = "cp1_early") Then
'            gL_eFuse_Sim_Blank = 0
'        Else
'            gL_eFuse_Sim_Blank = 1
'        End If
'        Call RunDSP.eFuse_SingleBitWave2CapWave32Bits(eFuse_UDRP, Trim_code_USO)
'        Call auto_eFuse_print_capWave32Bits(eFuse_UDRP, Trim_code_USO, False) ''''True to print out
'    End If
'
'    If (condstr = "cp1_early") Then
'        m_bitFlag_mode = 0
'    ElseIf (condstr = "stage") Then
'        m_bitFlag_mode = 1
'    ElseIf (condstr = "all") Then
'        m_bitFlag_mode = 2 ''''update later, was 2
'    Else
'        ''''default, here it prevents any typo issue
'        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
'        m_FBC = -1
'        m_cmpResult = -1
'    End If
'
'    ''''20160506 update
'    ''''due to the additional characters "_USI_USO_compare", so plus 18.
'    Call UpdateDLogColumns(gI_UDR_catename_maxLen + 18)
'
'    If (gL_UDRP_USO_PatBitOrder = "LSB") Then
'        m_PatBitOrder = "bit0_bitLast"
'    Else
'        m_PatBitOrder = "bitLast_bit0"
'    End If
'
'    ''''Offline simulation inside
'    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
'    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, m_bitFlag_mode, Trim_code_USO, m_FBC, blank_stage, allblank, True, m_PatBitOrder)
'
'    'If (TheExec.sites(Site).SiteVariableValue(m_siteVar) = 1) Then
'
'
'    If (m_CompareFlag = True) Then
'        Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, 1, Trim_code_USO, m_FBC, m_cmpResult, , , True, m_PatBitOrder)
'    End If

    Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

    Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
    Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate, gStr_PatName)
    
    Call auto_eFuse_SyntaxCheck(m_Fusetype, condstr)
    
        ''''----------------------------------------------------------------------------------------------
    If (gS_JobName <> "cp1_early") Then
        Dim m_UDRP_DATA_STR As New SiteVariant
        For Each site In TheExec.sites
            DoubleBitArray = gDW_UDRP_Read_DoubleBitWave.Data
            
            m_UDRP_DATA_STR(site) = "" ''''is a String [(bitLast)......(bit0)]
            
            For i = 0 To UBound(DoubleBitArray)
                m_UDRP_DATA_STR(site) = CStr(DoubleBitArray(i)) + m_UDRP_DATA_STR(site)
            Next i
            ''TheExec.Datalog.WriteComment "gS_CFG_Direct_Access_Str=" + CStr(gS_CFG_Direct_Access_Str(Site))
    
            ''''20161114 update for print all bits (DTR) in STDF
            Call auto_eFuse_to_STDF_allBits("UDRP", m_UDRP_DATA_STR(site))
        Next site
    End If
    ''''----------------------------------------------------------------------------------------------

    
    

    ''''201811XX reset
    If (gS_JobName = "cp1_early") Then gS_JobName = "cp1" ''''<MUST>

Exit Function

'End If
    
    
    
    
    

'    For j = 0 To pat_count - 1
'        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
'        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
'
'        ''''20160506 update
'        ''''due to the additional characters "_USI_USO_compare", so plus 18.
'        Call UpdateDLogColumns(gI_UDRP_catename_maxLen + 18)
'
'        For Each Site In TheExec.Sites
'            ''''initialize
'            USO_BitStr(Site) = ""  ''''MUST, and it's [bitLast ... bit0]
'            For i = 0 To UBound(DigCapArray)
'               DigCapArray(i) = 0  ''''MUST be [bit0 ... bitLast]
'            Next i
'
'            ''--------------------------------------------------------------------------------------
'            ''''20150717 update
'            ''''composite to the USO_BitStr() from the DSSC Capture
'            If (UCase(gL_UDRP_USO_PatBitOrder) = "MSB") Then
'                ''''<Notice>
'                ''''the pattern DSSC capture from MSB to LSB (bitLast to bit0)
'                ''''so Trim_code_USO.Element(0) is MSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(USO_CapBits - 1 - i)    ''''Reverse Bit String
'                    USO_BitStr(Site) = USO_BitStr(Site) & Trim_code_USO.Element(i) ''''[MSB ... LSB]
'                Next i
'            Else
'                ''''<Notice>
'                ''''the pattern DSSC capture from LSB to MSB (bit0 to bitLast)
'                ''''so Trim_code_USO.Element(0) is LSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(i)
'                    USO_BitStr(Site) = Trim_code_USO.Element(i) & USO_BitStr(Site) ''''[MSB ... LSB]
'                Next i
'            End If
'            ''--------------------------------------------------------------------------------------
'
'            ''''20160324 updae
'            ''''=============== Start of Simulated Data ===============
'            If (TheExec.TesterMode = testModeOffline) Then
'                ''''20160906 trial for the ugly codes
'                ''''<Issued codes> Shift out code [383:0]=111110000110100000000000000000001111100001101000000000000000000011111110111010000000000000000000111110100110100000000000000000001111101001101000000000000000000011111000011010000000000000000000111110000100000000000000000000001111100001000000000000000000000011111000010000000001000000010000111110000110000000000000000000001111100001000000000000000000000011111000010000000000000000000000
'                ''gS_UDRP_USI_BitStr(Site) = "111110000110100000000000000000001111100001101000000000000000000011111110111010000000000000000000111110100110100000000000000000001111101001101000000000000000000011111000011010000000000000000000111110000100000000000000000000001111100001000000000000000000000011111000010000000001000000010000111110000110000000000000000000001111100001000000000000000000000011111000010000000000000000000000"
'
'                USO_BitStr(Site) = gS_UDRP_USI_BitStr(Site)
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(gS_UDRP_USI_BitStr(Site)), i + 1, 1))
'                Next i
'
'                If (gS_JobName <> "cp1" Or TheExec.Sites.Item(Site).FlagState("F_UDRP_Early_Enable") = logicTrue) Then ''''was "cp1"
'                    TmpStr = ""
'                    For i = 0 To USO_CapBits - 1
'                        If (DigCapArray(i) = 0) Then
'                            DigCapArray(i) = gL_Sim_FuseBits(Site, i)
'                        Else
'                            gL_Sim_FuseBits(Site, i) = DigCapArray(i)
'                        End If
'                        TmpStr = CStr(DigCapArray(i)) + TmpStr
'                    Next i
'                    USO_BitStr(Site) = TmpStr ''''<MUST>
'                End If
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            If (True) Then
'                '=======================================================
'                '= Print out the caputured bit data from DigCap        =
'                '=======================================================
'                TheExec.Datalog.WriteComment ""
'                Call auto_PrintAllBitbyDSSC(DigCapArray, USO_PrintRow, UBound(DigCapArray) + 1, USO_BitPerRow)
'            End If
'            ''--------------------------------------------------------------------------------------
'
'            TheExec.Datalog.WriteComment "Site(" & Site & "), UDRP USO pat:" & PatUSOArray(j) & ", Shift out code [" + CStr(USO_CapBits - 1) + ":0]=" + USO_BitStr(Site)
'            ''TheExec.Datalog.WriteComment ""
'
'            ''''----------------------------------------------------------------------------------------------
'            ''''20161114 update for print all bits (DTR) in STDF
'            ''''20171016 update to excluding "cp1_early"
'            If (m_CP1_Early_Flag = False) Then Call auto_eFuse_to_STDF_allBits("UDRP", USO_BitStr(Site))
'            ''''----------------------------------------------------------------------------------------------
'
'            ''''20150717 New
'            Call auto_Decode_UDRP_Binary_Data(DigCapArray)
'
'            ''''----------------------------------------------------------------------------------
'            ''''judge pass/fail for the specific test limit
'            tmpStrL = StrReverse(USO_BitStr(Site)) ''''translate to [LSB......MSB]
'            For i = 0 To UBound(UDRP_Fuse.Category)
'                tmpdlgStr = ""
'                m_stage = LCase(UDRP_Fuse.Category(i).Stage)
'                m_catename = UDRP_Fuse.Category(i).Name
'                m_algorithm = LCase(UDRP_Fuse.Category(i).Algorithm)
'                m_LSBbit = UDRP_Fuse.Category(i).LSBbit
'                m_MSBBit = UDRP_Fuse.Category(i).MSBbit
'                m_bitwidth = UDRP_Fuse.Category(i).Bitwidth
'                m_lolmt = UDRP_Fuse.Category(i).LoLMT
'                m_hilmt = UDRP_Fuse.Category(i).HiLMT
'                m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
'                m_resolution = UDRP_Fuse.Category(i).Resoultion
'
'                m_decimal = UDRP_Fuse.Category(i).Read.Decimal(Site)
'                m_value = UDRP_Fuse.Category(i).Read.Value(Site)
'                m_bitSum = UDRP_Fuse.Category(i).Read.BitSummation(Site)
'                m_hexStr = UDRP_Fuse.Category(i).Read.HexStr(Site)
'                m_unitType = unitNone
'                m_scale = scaleNone ''''default
'                m_tsName = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
'
'                m_bitStrM = StrReverse(Mid(tmpStrL, m_LSBbit + 1, m_bitwidth))
'
'                m_testFlag = True ''''20171016 update
'                If (m_CP1_Early_Flag = True) Then ''''20171016 update
'                    ''''only compare these category with stage="cp1_early"
'                    If (m_stage = condstr) Then
'                        ''''other cases
'                        m_testValue = m_decimal
'                    Else
'                        ''''Here it's an excluding case
'                        ''''<MUST>
'                        m_testFlag = False
'                        m_testValue = 0
'                        m_lolmt = 0
'                        m_hilmt = 0
'                    End If
'
'                ElseIf (m_algorithm = "base") Then
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
'                    ''''<Notice> User Maintain
'                    ''''Ex:: step_vdd_cpu_p1 = VDD_BIN(vdd_cpu_p1).MODE_STEP
'                    m_testValue = 0 ''''default to fail
'                    If (m_defreal = "decimal") Then ''''20160624 update
'                        m_testValue = m_decimal
'                    ElseIf (m_defreal Like "safe*voltage" Or m_defreal = "default") Then
'                        m_unitType = unitVolt
'                        m_scale = scaleMilli
'                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
'                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
'                        m_testValue = m_value * 0.001 ''''to unit:V
'                    ElseIf (m_defreal = "bincut") Then
'                        m_catenameVbin = m_catename '150127
'                        ''''' Every testing stage needs to check this formula with bincut fused value (2017/5/17, Jack)
'                        vbinflag = auto_CheckVddBinInRangeNew(m_catenameVbin, m_resolution)
'
'                        ''''20160329 Add for the offline simulation
'                        If (vbinflag = 0 And TheExec.TesterMode = testModeOffline) Then
'                            vbinflag = 1
'                        End If
'
'                        m_Pmode = VddBinStr2Enum(m_catenameVbin)
'                        tmpVal = auto_VfuseStr_to_Vdd_New(m_bitStrM, m_bitwidth, m_resolution)
'                        MaxLevelIndex = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step '150127
'                        m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(MaxLevelIndex) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(MaxLevelIndex)
'                        m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
'                        ''''judge the result
'                        If (vbinflag = 1) Then
'                            m_value = tmpVal
'                        Else
'                            m_value = -999
'                            TmpStr = m_catename + "(Site " + CStr(Site) + ") = " + CStr(tmpVal) + " is not in range" '150127
'                            TheExec.Datalog.WriteComment TmpStr
'                        End If
'                        m_unitType = unitVolt
'                        m_scale = scaleMilli
'                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
'                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
'                        m_testValue = m_value * 0.001 ''''to unit:V
'                    End If
'                Else
'                    ''''other cases, 20160927 update
'                    m_testValue = m_decimal
'                End If
'
'                ''''20160108 New
'                m_tsName = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
'                Call auto_eFuse_chkLoLimit("UDRP", i, m_stage, m_lolmt)
'                Call auto_eFuse_chkHiLimit("UDRP", i, m_stage, m_hilmt)
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
'
'                ''''20171016 update
'                If (m_testFlag) Then
'                    TheExec.Flow.TestLimit resultVal:=m_testValue, lowVal:=m_lolmt, hiVal:=m_hilmt, Tname:=m_tsName, unit:=m_unitType, ScaleType:=m_scale
'                End If
'            Next i
'            ''''----------------------------------------------------------------------------------
'
'            ''''20160503 update for the 2nd way to prevent for the temp sensor trim will all zero values
'            ''''20160907 update
'            Dim m_valueSum As Long
'            Dim m_matchTMPS_flag As Boolean
'            m_valueSum = 0 ''''initialize
'            m_matchTMPS_flag = False
'            m_stage = "" ''''<MUST> 20160617 update, if the "trim/tmps" is existed then m_stage has its correct value.
'            For i = 0 To UBound(UDRP_Fuse.Category)
'                m_catename = UCase(UDRP_Fuse.Category(i).Name)
'                m_algorithm = LCase(UDRP_Fuse.Category(i).Algorithm)
'                If (m_catename Like "TEMP_SENSOR*" Or m_algorithm = "tmps") Then ''''was m_algorithm = "trim", 20171103 update
'                    m_stage = LCase(UDRP_Fuse.Category(i).Stage)
'                    m_decimal = UDRP_Fuse.Category(i).Read.Decimal(Site)
'                    m_valueSum = m_valueSum + m_decimal
'                    m_matchTMPS_flag = True
'                End If
'            Next i
'            If (m_matchTMPS_flag = True) Then
'                ''''if Job >= m_stage then m_valueSim >= 1
'                If (checkJob_less_Stage_Sequence(m_stage) = False) Then
'                    TheExec.Flow.TestLimit resultVal:=m_valueSum, lowVal:=1, Tname:="UDRP_TMPS_SUM"
'                    ''TheExec.Datalog.WriteComment ""
'                Else
'                    ''''if Job < m_stage then m_valueSim = 0
'                    ''''20180105 update
'                    If (gB_eFuse_Disable_ChkLMT_Flag = True) Then ''''case CP2 back to CP1 retest
'                        TheExec.Flow.TestLimit resultVal:=m_valueSum, lowVal:=0, Tname:="UDRP_TMPS_SUM"
'                    Else
'                        TheExec.Flow.TestLimit resultVal:=m_valueSum, lowVal:=0, hiVal:=0, Tname:="UDRP_TMPS_SUM"
'                    End If
'                End If
'            End If
'            ''''--------------------------------------------------------------------------------------------
'
'            ''''20160503 update
'            ''''compare both USI and USO for the specific stage, it's only when siteVar is '1'.
'            ''''Must be after the decode then you have the Read buffer value
'            ''''20180105 update
'            ''''The below is used to compare both USI and USO contents.
'            If (TheExec.Sites(Site).SiteVariableValue(m_siteVar) = 1) Then
'                Dim m_writeBitStrM As String
'                Dim m_readBitStrM As String
'                Dim m_usiusoCmp As Long
'                m_usiusoCmp = 0 ''''<MUST> default Compare Pass:0, Fail:1
'                For i = 0 To UBound(UDRP_Fuse.Category)
'                    m_stage = LCase(UDRP_Fuse.Category(i).Stage)
'                    If (m_stage = gS_JobName) Then
'                        m_catename = UDRP_Fuse.Category(i).Name
'                        m_writeBitStrM = UCase(UDRP_Fuse.Category(i).Write.BitstrM(Site))
'                        m_readBitStrM = UCase(UDRP_Fuse.Category(i).Read.BitstrM(Site))
'                        m_tsName = Replace(m_catename, " ", "_")
'                        m_tsName = m_tsName + "_USI_USO_compare"
'                        If (m_writeBitStrM <> m_readBitStrM) Then
'                            ''''20180105 update
'                            m_usiusoCmp = 1 ''''Fail Comparison
'                            TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(Site) + "), " + m_tsName + " Failure."
'                        Else
'                            ''''case: m_writeBitStrM = m_readBitStrM
'                            ''''reserve
'                        End If
'                    End If
'                Next i
'                TheExec.Flow.TestLimit resultVal:=m_usiusoCmp, lowVal:=0, hiVal:=0, Tname:="UDRP_USO_USI_Cmp"
'            End If
'            ''''--------------------------------------------------------------------------------------------
'
'            TheExec.Datalog.WriteComment ""
'            If (TheExec.Sites.ActiveCount = 0) Then Exit Function 'chihome
'        Next Site
'        Call UpdateDLogColumns__False
'    Next j
'
'     ''''20171016 update
'    If (m_CP1_Early_Flag = True) Then
'        gS_JobName = "cp1" ''''reset
'    End If
'    DebugPrintFunc USO_pat.Value

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20150731 New
Public Function auto_UDRP_USO_BlankChk_byStage(USO_pat As Pattern, OutPin As PinList, _
                    Optional m_decode_flag As Boolean = True, _
                    Optional condstr As String = "stage", _
                    Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_USO_BlankChk_byStage"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(USO_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_USO As New DSPWave
    Dim USO_BitStr As New SiteVariant
    Dim i As Long, j As Long, k As Long
    Dim PatUSOArray() As String
    Dim pat_count As Long, Status As Boolean
    Dim DigCapArray() As Long
    Dim ResultFlag As New SiteLong
    Dim testName As String
    Dim PinName As String

    Dim USO_PrintRow As Long
    Dim USO_BitPerRow As Long
    Dim USO_CapBits As Long
    Dim USO_TotalBit As Long
    Dim blank_stage As New SiteBoolean
    Dim allBlank As New SiteBoolean
    Dim SiteVarValue As Long
    Dim PrintSiteVarResult As String
    Dim SingleDoubleFBC As Long
    Dim m_jobinStage_flag As Boolean
    ReDim gL_Sim_FuseBits(TheExec.sites.Existing.Count - 1, gL_UDRP_USO_DigCapBits_Num - 1) ''''it's for the simulation
    
    USO_CapBits = gL_UDRP_USO_DigCapBits_Num
    USO_TotalBit = gL_UDRP_USO_DigCapBits_Num

    ''''it's used to identify if the Job Name is existed in the UDRP portion of the eFuse BitDef table.
    m_jobinStage_flag = auto_eFuse_JobExistInStage("UDRP", True) ''''<MUST>

    ''''display bit number per row
    If (gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp") Then
        USO_BitPerRow = 32
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        USO_BitPerRow = 16
    End If

    ''''20150731 update
    USO_PrintRow = IIf((USO_TotalBit Mod USO_BitPerRow) > 0, Floor(USO_TotalBit / USO_BitPerRow) + 1, Floor(USO_TotalBit / USO_BitPerRow))
    ReDim DigCapArray((USO_PrintRow * USO_BitPerRow) - 1)

    '****** Initialize Site Varaible ******
    Dim m_siteVar As String
    m_siteVar = "UDR_PChk_Var"
    For Each site In TheExec.sites.Existing
        blank_stage(site) = True
        TheExec.sites(site).SiteVariableValue(m_siteVar) = -1
    Next site

    ''TheHdw.Patterns(USO_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Status = GetPatListFromPatternSet(USO_pat.Value, PatUSOArray, pat_count)
    
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
    Dim m_PatBitOrder As String

    m_Fusetype = eFuse_UDRP
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_stage = True
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    'gDL_eFuse_Orientation = eFuse_1_Bit
    gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
    Dim m_SerialType As Boolean:: m_SerialType = False
    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
        m_SerialType = True
        gDB_SerialType = True
    End If
    
    'Actually, we have only one pattern
    Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(0), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
    Call TheHdw.Patterns(PatUSOArray(0)).Test(pfAlways, 0)
    
    gStr_PatName = PatUSOArray(0)

    Call UpdateDLogColumns(gI_UDRP_catename_maxLen) ''''<MUST>must after execute pattern
    
    If (gL_UDRP_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        Dim Temp_USO As New DSPWave
        Temp_USO.CreateConstant 0, USO_CapBits, DspLong
        
'        If (gS_JobName = "cp1_early") Then
'            gL_eFuse_Sim_Blank = 0
'        Else
'            gL_eFuse_Sim_Blank = 1
'        End If
        If (UCase(gS_JobName) Like "*CP*") Then
        gL_eFuse_Sim_Blank = 0
        End If
        gDW_UDRP_Pgm_SingleBitWave = Temp_USO.Copy
        Call rundsp.eFuse_SingleBitWave2CapWave32Bits(eFuse_UDRP, Temp_USO)
        'Call auto_eFuse_print_capWave32Bits(eFuse_UDR, Temp_USO, False) ''''True to print out
        If (m_PatBitOrder = "bit0_bitLast") Then
            For Each site In TheExec.sites
                Trim_code_USO(site) = Temp_USO(site).Copy
            Next
        Else
            Call ReverseWave(Temp_USO, Trim_code_USO, m_PatBitOrder, USO_CapBits)
        End If
    End If
    
    'Dim condstr As String:: condstr = "stage"
    Dim m_bitFlag_mode As Long
    
    If (condstr = "cp1_early") Then
        m_bitFlag_mode = 0
    ElseIf (condstr = "stage") Then
        m_bitFlag_mode = 1
    ElseIf (condstr = "all") Then
        m_bitFlag_mode = 2 ''''update later, was 2
    Else
        ''''default, here it prevents any typo issue
'        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
'        m_FBC = -1
'        m_cmpResult = -1
    End If
    
    If (gL_UDRP_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    'gL_eFuse_Sim_Blank = 1
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_USO, m_FBC, blank_stage, allBlank, m_SerialType, m_PatBitOrder)

    If (blank_stage.Any(False) = True) Then
   
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

        Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate)
   
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
    m_SiteVarValue = blank_stage
    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
    If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only
        
    For Each site In TheExec.sites
        If (auto_eFuse_GetAllPatTestPass_Flag("UDRP") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    'gL_UDR_FBC = m_FBC
    
    Call UpdateDLogColumns(gI_UDRP_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "UDRP_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "UDRP_Blank_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    'DebugPrintFunc ReadPatSet.Value

Exit Function
End If
    
    
    
    
    
    

'    For j = 0 To pat_count - 1 'Actually, we have only one pattern
'        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
'        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
'
'        Call UpdateDLogColumns(gI_UDRP_catename_maxLen)
'
'        For Each Site In TheExec.Sites
'            ''''initialize
'            allblank(Site) = True
'            USO_BitStr(Site) = ""  ''''MUST, and it's [MSB ... LSB]
'            For i = 0 To UBound(DigCapArray)
'               DigCapArray(i) = 0  ''''MUST be [LSB(0) ... MSB(lastbit)]
'            Next i
'
'            ''''20150717 update
'            ''''composite to the USO_BitStr() from the DSSC Capture
'            If (UCase(gL_UDRP_USO_PatBitOrder) = "MSB") Then
'                ''''<Notice>
'                ''''the pattern DSSC capture from MSB to LSB (bitLast to bit0)
'                ''''so Trim_code_USO.Element(0) is MSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(USO_CapBits - 1 - i)    ''''Reverse Bit String
'                    USO_BitStr(Site) = USO_BitStr(Site) & Trim_code_USO.Element(i) ''''[MSB ... LSB]
'                    If (DigCapArray(i) <> 0) Then allblank(Site) = False
'                Next i
'            Else
'                ''''<Notice>
'                ''''the pattern DSSC capture from LSB to MSB (bit0 to bitLast)
'                ''''so Trim_code_USO.Element(0) is LSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(i)
'                    USO_BitStr(Site) = Trim_code_USO.Element(i) & USO_BitStr(Site) ''''[MSB ... LSB]
'                    If (DigCapArray(i) <> 0) Then allblank(Site) = False
'                Next i
'            End If
'
'            ''''20160324 update, 20161108 update
'            ''''=============== Start of Simulated Data ===============
'            If (TheExec.TesterMode = testModeOffline) Then
'                allblank(Site) = False
'                blank_stage(Site) = True ''''True or False(sim for re-test)
'                If (blank_stage(Site) = True And gS_JobName = "cp1") Then allblank(Site) = True
'
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                If (blank_stage(Site) = True) Then
'                    TheExec.Datalog.WriteComment vbTab & "[ blank_stage = True, Simulate Category (m_stage < Job[" + UCase(gS_JobName) + "]) ]"
'                Else
'                    TheExec.Datalog.WriteComment vbTab & "[ blank_stage = False, Simulate Category (m_stage <= Job[" + UCase(gS_JobName) + "]) ]"
'                End If
'                Call eFuseENGFakeValue_Sim
'                Call auto_UDRP_USI_Sim(blank_stage(Site), False) ''''True for print debug
'                USO_BitStr(Site) = gS_UDRP_USI_BitStr(Site)
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(gS_UDRP_USI_BitStr(Site)), i + 1, 1))
'                    ''''it's used for the Read/Syntax simulation
'                    gL_Sim_FuseBits(Site, i) = DigCapArray(i)
'                Next i
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            ''''20150731 update
'            ''''depends on the specific Job to judge if it's blank_stage on the specific stage bits
'            testName = "UDRP_BlankChk_" + UCase(gS_JobName)
'            SingleDoubleFBC = 0 ''''init
'            Call auto_eFuse_BlankChk_FBC_byStage("UDRP", DigCapArray, blank_stage, SingleDoubleFBC)
'
'            '=======================================================
'            '= Print out the caputured bit data from DigCap        =
'            '=======================================================
'            TheExec.Datalog.WriteComment ""
'            Call auto_PrintAllBitbyDSSC(DigCapArray, USO_PrintRow, UBound(DigCapArray) + 1, USO_BitPerRow)
'
'            If blank_stage(Site) = False Then ''If not blank
'                If (SingleDoubleFBC = 0) Then
'                    ResultFlag(Site) = 0    ''Pass Blank check
'                    PinName = "Pass"
'                    SiteVarValue = 2
'                Else
'                    ResultFlag(Site) = 1    ''Fail Blank check
'                    PinName = "Fail"
'                    SiteVarValue = 0
'                End If
'            Else '' True means this Efuse is balnk.
'                If (gS_JobName <> "cp1" And allblank(Site) = True) Then
'                    ''''It's NOT allowed. So set it Failed.
'                    ResultFlag(Site) = 1   ''Fail then NO fuse
'                    PinName = "Fail"
'                    SiteVarValue = 0
'                Else
'                    ''''Here it's used to check if HardIP pattern test pass or not.
'                    If (auto_eFuse_GetAllPatTestPass_Flag("UDRP") = False) Then
'                        ResultFlag(Site) = 1   ''Fail then NO fuse
'                        PinName = "Fail"
'                        SiteVarValue = 0
'                    Else
'                        ResultFlag(Site) = 0   ''Pass Blank check
'                        PinName = "Pass"
'                        SiteVarValue = 1
'                        ''''<MUST>
'                        ''''it's used to identify if the Job Name is existed in the UDRP portion of the eFuse BitDef table.
'                        If (m_jobinStage_flag = False) Then
'                            ''''<Important> Then it will NOT go WritebyStage to let the user confusion.
'                            SiteVarValue = 2
'                        Else
'                            ''''check it later if any other case
'                        End If
'                    End If
'                End If
'            End If
'
'            TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'
'            If (False) Then
'                PrintSiteVarResult = "Site (" + CStr(Site) + ") " + m_siteVar + " = " + CStr(SiteVarValue)
'                TheExec.Datalog.WriteComment PrintSiteVarResult & vbCrLf
'            End If
'
'            TheExec.Datalog.WriteComment "Site(" & Site & "), UDRP USO pat:" & PatUSOArray(j) & ", Shift out code [" + CStr(USO_CapBits - 1) + ":0]=" + USO_BitStr(Site)
'            TheExec.Datalog.WriteComment ""
'
'            ''''20150717 New, 20160321 update
'            If (SiteVarValue <> 1 And m_decode_flag) Then Call auto_Decode_UDRP_Binary_Data(DigCapArray)
'
'            gB_UDRP_decode_flag(Site) = False ''''20160531, <MUST> reset and init
'
'            ''''------------------------------------------------------------------------
'            ''''<Very Important>
'            ''''20171211 update to prevent Re-Fuse while UFR pattern failure
'            '''' It will casue the all Blank result to mis-leading the blank judgement
'            If (TheExec.Sites.Item(Site).FlagState("F_udrp_blank") = logicTrue) Then
'                ''''Here it means that UFR read mode pattern is failure
'                TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: Site(" + CStr(Site) + ") UDRP_UFR_NV Function Failure"
'                TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: Site(" + CStr(Site) + ") Set " + m_siteVar + " = 0"
'                SiteVarValue = 0
'                TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'            End If
'            ''''------------------------------------------------------------------------
'
'            ''Binning out
'            TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'            TheExec.Flow.TestLimit resultVal:=ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName, PinName:=PinName
'            If (TheExec.Sites.ActiveCount = 0) Then Exit Function 'chihome
'        Next Site
'
'        Call UpdateDLogColumns__False
'    Next j
'
'    DebugPrintFunc USO_pat.Value
 
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDRP_USO_Read_Decode(USO_pat As Pattern, OutPin As PinList, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_USO_Read_Decode"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(USO_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_USO As New DSPWave
    Dim USO_BitStr As New SiteVariant
    Dim i As Long, j As Long
    Dim PatUSOArray() As String
    Dim pat_count As Long, Status As Boolean

    Dim DigCapArray() As Long
    Dim ResultFlag As New SiteLong
    Dim testName As String
    Dim PinName As String
    
    Dim USO_PrintRow As Long
    Dim USO_BitPerRow As Long
    Dim USO_CapBits As Long
    Dim USO_TotalBit As Long
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    ''Dim SiteVarValue As Long
    ReDim gL_Sim_FuseBits(TheExec.sites.Existing.Count - 1, gL_UDRP_USO_DigCapBits_Num - 1) ''''it's for the simulation

    USO_CapBits = gL_UDRP_USO_DigCapBits_Num
    USO_TotalBit = gL_UDRP_USO_DigCapBits_Num

    ''''display bit number per row
    If (gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp") Then
        USO_BitPerRow = 32
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        USO_BitPerRow = 16
    End If

    ''''20150731 update
    USO_PrintRow = IIf((USO_TotalBit Mod USO_BitPerRow) > 0, Floor(USO_TotalBit / USO_BitPerRow) + 1, Floor(USO_TotalBit / USO_BitPerRow))
    ReDim DigCapArray((USO_PrintRow * USO_BitPerRow) - 1)

    ''TheHdw.Patterns(USO_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Status = GetPatListFromPatternSet(USO_pat.Value, PatUSOArray, pat_count)
    
    
If (gB_eFuse_newMethod = True) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If
    
    'gDL_eFuse_Orientation = eFuse_1_Bit
    
    'Actually, we have only one pattern
    Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(0), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
    Call TheHdw.Patterns(PatUSOArray(0)).Test(pfAlways, 0)
    
    gStr_PatName = PatUSOArray(0)

    Call UpdateDLogColumns(gI_UDRP_catename_maxLen) ''''<MUST>must after execute pattern
    
    ''''----------------------------------------------------
    ''''201812XX New Method by DSPWave
    ''''----------------------------------------------------
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong

    m_Fusetype = eFuse_UDRP
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_stage = True
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
    Dim m_SerialType As Boolean:: m_SerialType = False
    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
        m_SerialType = True
        gDB_SerialType = True
    End If
    
'    If (TheExec.TesterMode = testModeOffline) Then
'        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
'        ''''Here MarginRead it only read the programmed Site.
'        If (gS_JobName = "cp1_early") Then
'            gL_eFuse_Sim_Blank = 0
'        Else
'            gL_eFuse_Sim_Blank = 1
'        End If
'        Call RunDSP.eFuse_SingleBitWave2CapWave32Bits(eFuse_UDR, Trim_code_USO)
'        Call auto_eFuse_print_capWave32Bits(eFuse_UDR, Trim_code_USO, False) ''''True to print out
'    End If
    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    Dim m_PatBitOrder As String
    If (gL_UDRP_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_USO, m_FBC, blank_stage, allBlank, True, m_PatBitOrder)

    If (True) Then
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

        Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate, gStr_PatName)
    End If
      
    testName = "UDRP_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    Call UpdateDLogColumns__False
    'DebugPrintFunc ReadPatSet.Value
    
    ''''20170111 Add
    Call auto_eFuse_ReadAllData_to_DictDSPWave("UDRP", False, False)
    
Exit Function
End If
    
    
    
    
    
'    For j = 0 To pat_count - 1 'Actually, we have only one pattern
'        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
'        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
'
'        Call UpdateDLogColumns(gI_UDRP_catename_maxLen) ''''<MUST>must after execute pattern
'
'        For Each Site In TheExec.Sites
'            ''''initialize
'            USO_BitStr(Site) = ""  ''''MUST, and it's [MSB ... LSB]
'            For i = 0 To UBound(DigCapArray)
'               DigCapArray(i) = 0  ''''MUST be [LSB(0)......MSB(lastbit)]
'            Next i
'
'            ''''20150717 update
'            ''''composite to the USO_BitStr() from the DSSC Capture
'            If (UCase(gL_UDRP_USO_PatBitOrder) = "MSB") Then
'                ''''<Notice>
'                ''''the pattern DSSC capture from MSB to LSB (bitLast to bit0)
'                ''''so Trim_code_USO.Element(0) is MSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(USO_CapBits - 1 - i)    ''''Reverse Bit String
'                    USO_BitStr(Site) = USO_BitStr(Site) & Trim_code_USO.Element(i) ''''[MSB ... LSB]
'                Next i
'            Else
'                ''''<Notice>
'                ''''the pattern DSSC capture from LSB to MSB (bit0 to bitLast)
'                ''''so Trim_code_USO.Element(0) is LSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(i)
'                    USO_BitStr(Site) = Trim_code_USO.Element(i) & USO_BitStr(Site) ''''[MSB ... LSB]
'                Next i
'            End If
'
'            ''''=============== Start of Simulated Data ===============
'            ''''20161108 update
'            If (TheExec.TesterMode = testModeOffline) Then
'                allblank(Site) = False
'                blank_stage(Site) = True ''''True or False(sim for re-test)
'                If (blank_stage(Site) = True And gS_JobName = "cp1") Then allblank(Site) = True
'
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                If (blank_stage(Site) = True) Then
'                    TheExec.Datalog.WriteComment vbTab & "[ blank_stage = True, Simulate Category (m_stage < Job[" + UCase(gS_JobName) + "]) ]"
'                Else
'                    TheExec.Datalog.WriteComment vbTab & "[ blank_stage = False, Simulate Category (m_stage <= Job[" + UCase(gS_JobName) + "]) ]"
'                End If
'                Call eFuseENGFakeValue_Sim
'                Call auto_UDRP_USI_Sim(blank_stage(Site), False) ''''True for print debug
'                USO_BitStr(Site) = gS_UDRP_USI_BitStr(Site)
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(gS_UDRP_USI_BitStr(Site)), i + 1, 1))
'                    ''''it's used for the Read/Syntax simulation
'                    gL_Sim_FuseBits(Site, i) = DigCapArray(i)
'                Next i
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            '=======================================================
'            '= Print out the caputured bit data from DigCap        =
'            '=======================================================
'            TheExec.Datalog.WriteComment ""
'            Call auto_PrintAllBitbyDSSC(DigCapArray, USO_PrintRow, UBound(DigCapArray) + 1, USO_BitPerRow)
'
'            Call auto_Decode_UDRP_Binary_Data(DigCapArray, Not allblank(Site)) ''''True for Debug, only Not allBlank to show the decode result
'
'            ''Binning out
'            TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0
'        Next Site
'
'        Call UpdateDLogColumns__False
'
'    Next j
'
'    DebugPrintFunc USO_pat.Value
'
'    ''''20170111 Add
'    Call auto_eFuse_ReadAllData_to_DictDSPWave("UDRP", False, False)

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20171016 used for the stage "CP1_Early"
Public Function auto_UDRP_USO_BlankChk_Early_byStage(USO_pat As Pattern, OutPin As PinList, Optional m_decode_flag As Boolean = True, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_USO_BlankChk_Early_byStage"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(USO_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_USO As New DSPWave
    Dim USO_BitStr As New SiteVariant
    Dim i As Long, j As Long
    Dim PatUSOArray() As String
    Dim pat_count As Long, Status As Boolean

    Dim DigCapArray() As Long
    Dim ResultFlag As New SiteLong
    Dim testName As String
    Dim PinName As String
    
    Dim USO_PrintRow As Long
    Dim USO_BitPerRow As Long
    Dim USO_CapBits As Long
    Dim USO_TotalBit As Long
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    Dim blank_early As Boolean
    Dim blank_stage_noEarly As Boolean
    Dim SiteVarValue As Long
    Dim PrintSiteVarResult As String
    ReDim gL_Sim_FuseBits(TheExec.sites.Existing.Count - 1, gL_UDRP_USO_DigCapBits_Num - 1) ''''it's for the simulation

    USO_CapBits = gL_UDRP_USO_DigCapBits_Num
    USO_TotalBit = gL_UDRP_USO_DigCapBits_Num

    ''''display bit number per row
    If (gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp") Then
        USO_BitPerRow = 32
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        USO_BitPerRow = 16
    End If

    ''''20150731 update
    USO_PrintRow = IIf((USO_TotalBit Mod USO_BitPerRow) > 0, Floor(USO_TotalBit / USO_BitPerRow) + 1, Floor(USO_TotalBit / USO_BitPerRow))
    ReDim DigCapArray((USO_PrintRow * USO_BitPerRow) - 1)

    '****** Initialize Site Varaible ******
    Dim m_siteVar As String
    m_siteVar = "UDR_PChk_Var"
    For Each site In TheExec.sites.Existing
        allBlank(site) = True
        TheExec.sites(site).SiteVariableValue(m_siteVar) = -1
    Next site

    ''TheHdw.Patterns(USO_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Status = GetPatListFromPatternSet(USO_pat.Value, PatUSOArray, pat_count)
    
    ''''20171016 update
    If (gS_JobName = "cp1") Then
        gS_JobName = "cp1_early"
    End If

    ''''20180105 New
    ''''It's used to check if gS_JobName is existed in all UDRP_Fuse programming stages
    ''''it's used to identify if the Job Name is existed in the UDRP portion of the eFuse BitDef table.
    Dim m_jobinStage_flag As Boolean
    m_jobinStage_flag = auto_eFuse_JobExistInStage("UDRP", True) ''''<MUST>
    
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
    Dim m_PatBitOrder As String

    m_Fusetype = eFuse_UDRP
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = True
    blank_stage = True
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    'gDL_eFuse_Orientation = eFuse_1_Bit
     gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
    Dim m_SerialType As Boolean:: m_SerialType = False
    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
        m_SerialType = True
        gDB_SerialType = True
    End If
    
    'Actually, we have only one pattern
    Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(0), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
    Call TheHdw.Patterns(PatUSOArray(0)).Test(pfAlways, 0)

    Call UpdateDLogColumns(gI_UDRP_catename_maxLen) ''''<MUST>must after execute pattern

'    ''''----------------------------------------------------
'    ''''201812XX New Method by DSPWave
'    ''''----------------------------------------------------
'    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
'    Dim m_Fusetype As eFuseBlockType
'    Dim m_SiteVarValue As New SiteLong
'    Dim m_ResultFlag As New SiteLong
'    Dim m_PatBitOrder As String
'
'    m_Fusetype = eFuse_UDRP
'    m_FBC = -1               ''''initialize
'    m_ResultFlag = -1        ''''initialize
'    m_SiteVarValue = -1      ''''initialize
'    allblank = True
'    blank_stage = True
'
'    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    If (gL_UDRP_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 0, Trim_code_USO, m_FBC, blank_stage, allBlank, m_SerialType, m_PatBitOrder)

    If (blank_stage.Any(False) = True) Then
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

        Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate)
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
    m_SiteVarValue = blank_stage
    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
    If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only
        
    For Each site In TheExec.sites
        If (auto_eFuse_GetAllPatTestPass_Flag("UDRP") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    'gL_UDR_FBC = m_FBC
    
    Call UpdateDLogColumns(gI_UDRP_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "UDRP_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "UDRP_Blank_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    'DebugPrintFunc ReadPatSet.Value
    
Exit Function

End If
    
    
    
    
    
'    For j = 0 To pat_count - 1 'Actually, we have only one pattern
'        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
'        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
'
'        Call UpdateDLogColumns(gI_UDRP_catename_maxLen) ''''<MUST>must after execute pattern
'
'        For Each Site In TheExec.Sites
'            ''''initialize
'            USO_BitStr(Site) = ""  ''''MUST, and it's [MSB ... LSB]
'            For i = 0 To UBound(DigCapArray)
'               DigCapArray(i) = 0  ''''MUST be [LSB(0)......MSB(lastbit)]
'            Next i
'
'            ''''20150717 update
'            ''''composite to the USO_BitStr() from the DSSC Capture
'            If (UCase(gL_UDRP_USO_PatBitOrder) = "MSB") Then
'                ''''<Notice>
'                ''''the pattern DSSC capture from MSB to LSB (bitLast to bit0)
'                ''''so Trim_code_USO.Element(0) is MSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(USO_CapBits - 1 - i)    ''''Reverse Bit String
'                    USO_BitStr(Site) = USO_BitStr(Site) & Trim_code_USO.Element(i) ''''[MSB ... LSB]
'                Next i
'            Else
'                ''''<Notice>
'                ''''the pattern DSSC capture from LSB to MSB (bit0 to bitLast)
'                ''''so Trim_code_USO.Element(0) is LSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(i)
'                    USO_BitStr(Site) = Trim_code_USO.Element(i) & USO_BitStr(Site) ''''[MSB ... LSB]
'                Next i
'            End If
'
'
'            ''''20160324 update, 20161108 update
'            ''''=============== Start of Simulated Data ===============
'            If (TheExec.TesterMode = testModeOffline) Then
'                allblank(Site) = False
'                blank_stage(Site) = True ''''True or False(sim for re-test)
'                'If (blank_stage(Site) = True And gS_JobName = "cp1") Then allBlank(Site) = True
'                If (blank_stage(Site) = True And gS_JobName = "cp1_early") Then allblank(Site) = True
'
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                If (blank_stage(Site) = True) Then
'                    TheExec.Datalog.WriteComment vbTab & "[ blank_stage = True, Simulate Category (m_stage < Job[" + UCase(gS_JobName) + "]) ]"
'                Else
'                    TheExec.Datalog.WriteComment vbTab & "[ blank_stage = False, Simulate Category (m_stage <= Job[" + UCase(gS_JobName) + "]) ]"
'                End If
'                Call eFuseENGFakeValue_Sim
'                Call auto_UDRP_USI_Sim(blank_stage(Site), False) ''''True for print debug
'                USO_BitStr(Site) = gS_UDRP_USI_BitStr(Site)
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(gS_UDRP_USI_BitStr(Site)), i + 1, 1))
'                    ''''it's used for the Read/Syntax simulation
'                    gL_Sim_FuseBits(Site, i) = DigCapArray(i)
'                Next i
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            '=======================================================
'            '= Print out the caputured bit data from DigCap        =
'            '=======================================================
'            TheExec.Datalog.WriteComment ""
'            Call auto_PrintAllBitbyDSSC(DigCapArray, USO_PrintRow, UBound(DigCapArray) + 1, USO_BitPerRow)
'
'            For i = 0 To USO_TotalBit - 1
'                If DigCapArray(i) <> 0 Then
'                    allblank(Site) = False
'                    Exit For
'                End If
'            Next i
'
'            testName = "UDRP_BlankChk_CP1_Early" ''''"UDRP_BlankChk_" + UCase(gS_JobName)
'
'            Call auto_UDRP_blank_check_Early(DigCapArray, blank_early, blank_stage_noEarly)
'            ''TheExec.Datalog.WriteComment "...blank_early = " & blank_Early
'            ''TheExec.Datalog.WriteComment "...blank_stage_noEarly = " & blank_stage_noEarly
'
'            ResultFlag(Site) = 0 'stand for this test item pass
'            PinName = "Pass"
'            SiteVarValue = 2 ''default read only
'
'            ''''20180105 update
'            If (m_jobinStage_flag = False) Then
'                ''''Read Only
'                SiteVarValue = 2
'
'            ElseIf allblank(Site) = True Then  'If all blank
'                SiteVarValue = 1
'
'            Else  'If not allblank
'                If gS_JobName = "cp1_early" Then
'                    If (blank_early = True) Then
'                        SiteVarValue = 1
'                    End If
'                End If
'            End If
'
'            TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'
'            If (False) Then
'                PrintSiteVarResult = "Site (" + CStr(Site) + ") " + m_siteVar + " = " + CStr(SiteVarValue)
'                TheExec.Datalog.WriteComment PrintSiteVarResult & vbCrLf
'            End If
'
'            TheExec.Datalog.WriteComment "Site(" & Site & "), UDRP USO pat:" & PatUSOArray(j) & ", Shift out code [" + CStr(USO_CapBits - 1) + ":0]=" + USO_BitStr(Site)
'            TheExec.Datalog.WriteComment ""
'
'            ''''20150717 New, 20160321 update
'            If (SiteVarValue <> 1 And m_decode_flag) Then Call auto_Decode_UDRP_Binary_Data(DigCapArray)
'
'            gB_UDRP_decode_flag(Site) = False ''''20160531, <MUST> reset and init
'
'            ''''------------------------------------------------------------------------
'            ''''<Very Important>
'            ''''20171211 update to prevent Re-Fuse while UFR pattern failure
'            '''' It will casue the all Blank result to mis-leading the blank judgement
'            If (TheExec.Sites.Item(Site).FlagState("F_udrp_blank") = logicTrue) Then
'                ''''Here it means that UFR read mode pattern is failure
'                TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: Site(" + CStr(Site) + ") UDRP_UFR_NV Function Failure"
'                TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: Site(" + CStr(Site) + ") Set " + m_siteVar + " = 0"
'                SiteVarValue = 0
'                TheExec.Sites(Site).SiteVariableValue(m_siteVar) = SiteVarValue
'            End If
'            ''''------------------------------------------------------------------------
'
'            ''Binning out
'            TheExec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:="Value"
'            TheExec.Flow.TestLimit resultVal:=ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName, PinName:=PinName
'            If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
'        Next Site
'
'        Call UpdateDLogColumns__False
'
'    Next j
'
'    DebugPrintFunc USO_pat.Value
'
'    ''''20171016 update
'    If (gS_JobName = "cp1_early") Then
'        gS_JobName = "cp1" ''''Reset
'    End If
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20160324 used for the offline simulation
Public Function auto_UDRP_USI_Sim(m_blank As Boolean, Optional showPrint As Boolean = False) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_USI_Sim"

    Dim site As Variant
    Dim i As Long, j As Long, k As Long
    
    Dim usiarrSize As Long
    usiarrSize = gL_UDRP_USI_DigSrcBits_Num * gC_UDRP_USI_DSSCRepeatCyclePerBit

    Dim PgmBitArr() As Long
    ReDim PgmBitArr(gL_UDRP_USI_DigSrcBits_Num - 1)

    Dim USI_Array() As Long
    ReDim USI_Array(TheExec.sites.Existing.Count - 1, usiarrSize - 1)
    
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant
    Dim m_bitStrM As String
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim tmpVbin As Variant
    Dim tmpVfuse As String
    Dim tmpdlgStr As String
    Dim tmpStr As String
    Dim m_defval As Variant
    Dim m_defreal As String
    Dim m_bitsum As Long
    Dim m_stage As String
    Dim TmpVal As Variant
    Dim m_resolution As Double
    ''''------------------------------------------------------------

    For Each site In TheExec.sites
        '''' initialize
        gS_UDRP_USI_BitStr(site) = ""
        For i = 0 To UBound(PgmBitArr)
            PgmBitArr(i) = 0
        Next i

        '''' 1st Step: get the PgmBitArr() per Site
        For i = 0 To UBound(UDRP_Fuse.Category)
            tmpdlgStr = ""
            m_catename = UDRP_Fuse.Category(i).Name
            m_algorithm = LCase(UDRP_Fuse.Category(i).algorithm)
            m_LSBbit = UDRP_Fuse.Category(i).LSBbit
            m_MSBBit = UDRP_Fuse.Category(i).MSBbit
            m_bitwidth = UDRP_Fuse.Category(i).BitWidth
            m_lolmt = UDRP_Fuse.Category(i).LoLMT
            m_hilmt = UDRP_Fuse.Category(i).HiLMT
            m_defval = UDRP_Fuse.Category(i).DefaultValue
            m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
            m_stage = LCase(UDRP_Fuse.Category(i).Stage)
            m_resolution = UDRP_Fuse.Category(i).Resoultion

            ''''20150710 new datalog format
            tmpdlgStr = "Site(" + CStr(site) + ") Simulation : " + FormatNumeric(m_catename, gI_UDRP_catename_maxLen)
            tmpdlgStr = tmpdlgStr + " [" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "] = "
            
            If (m_algorithm = "base") Then
                If (checkJob_less_Stage_Sequence(m_stage) = True) Then
                    TmpVal = 0
                    m_decimal = 0
                Else
                    If (m_defreal = "decimal") Then ''''20160624 update
                        m_decimal = m_defval
                        TmpVal = m_decimal
                    Else
                        TmpVal = gD_BaseVoltage
                        m_decimal = gD_VBaseFuse
                    End If
                End If

            ElseIf (m_algorithm = "vddbin") Then
                ''''<Notice>
                ''''Here m_catename MUST be same as the content of Enum EcidVddBinningFlow
                ''''Ex: VDD_SRAM_P1 in (Enum EcidVddBinningFlow)
                ''''Ex: m_decimal = VBIN_RESULT(VddBinStr2Enum("VDD_CPU_P1")).GRADEVDD(Site)
                If (checkJob_less_Stage_Sequence(m_stage) = True) Then
                    tmpVbin = 0
                Else
                    If (m_defreal = "bincut") Then
                        tmpVbin = VBIN_RESULT(VddBinStr2Enum(m_catename)).GRADEVDD(site)
                        ''''20160329 add for the offline simulation, 20160714 update
                        If ((tmpVbin = 0 Or tmpVbin = -1) And TheExec.TesterMode = testModeOffline) Then
                            tmpVbin = gD_UDRP_BaseVoltage + m_resolution * auto_eFuse_GetWriteDecimal("UDRP", m_catename, False)
                        End If
                    Else
                        tmpVbin = m_defval
                    End If
                End If

                If (m_defreal = "decimal") Then ''''20160624 update
                    m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRP", m_catename, m_defreal, m_defval, False)
                Else
                    m_decimal = auto_Vbin_to_VfuseStr_New(tmpVbin, m_bitwidth, tmpVfuse, m_resolution)
                End If
            
            ElseIf (m_algorithm = "app") Then
                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRP", m_catename, m_defreal, m_defval, False)

            Else ''other cases
                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRP", m_catename, m_defreal, m_defval, False)

            End If

            ''''-------------------------------------------------------------------------------------------------------
            ''''20150825 update
            Call auto_eFuse_Dec2PgmArr_Write_byStage("UDRP", m_stage, i, m_decimal, m_LSBbit, m_MSBBit, PgmBitArr, False) ''''False for the Simulation
            m_decimal = UDRP_Fuse.Category(i).Write.Decimal(site)
            m_bitStrM = UDRP_Fuse.Category(i).Write.BitStrM(site)
            tmpStr = " [" + m_bitStrM + "]"
            If (m_algorithm = "vddbin") Then
                If (m_defreal = "decimal") Then
                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + tmpStr
                Else
                    ''''<Notice> Here using .Value to store VDDBIN value
                    UDRP_Fuse.Category(i).Write.Value(site) = tmpVbin
                    UDRP_Fuse.Category(i).Write.ValStr(site) = CStr(tmpVbin)
                    tmpdlgStr = tmpdlgStr + FormatNumeric(CStr(tmpVbin) + "mV", 10) + tmpStr + " = " + FormatNumeric(m_decimal, -5)
                End If
            ElseIf (m_algorithm = "base") Then ''''20160624 update
                If (m_defreal = "decimal") Then ''''20160624 update
                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + tmpStr
                Else
                    ''''<Notice> Here using .Value to store VDDBIN value
                    UDRP_Fuse.Category(i).Write.Value(site) = TmpVal
                    UDRP_Fuse.Category(i).Write.ValStr(site) = CStr(TmpVal)
                    tmpdlgStr = tmpdlgStr + FormatNumeric(CStr(TmpVal) + "mV", 10) + tmpStr + " = " + FormatNumeric(m_decimal, -5)
                End If
            Else
                tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + tmpStr
            End If
            ''''-------------------------------------------------------------------------------------------------------

            ''''20171016 update, black=True, set "Job<=Stage" category bits == Zero
            If (m_blank = True And (checkJob_less_Stage_Sequence(m_stage) = True Or m_stage = gS_JobName)) Then
                For j = m_LSBbit To m_MSBBit
                    PgmBitArr(j) = 0
                Next j
            End If

            If (showPrint) Then TheExec.Datalog.WriteComment tmpdlgStr
        Next i

        ''''20150717 update
        '''' 2nd Step: composite to the UDRP_USI_Array() for the DSSC Source
        k = 0
        tmpdlgStr = ""
        If (UCase(gL_UDRP_USI_PatBitOrder) = "MSB") Then
            ''''case: gL_UDRP_USI_PatBitOrder is MSB
            ''''<Notice> USI_Array(0) is MSB, so it should be PgmBitArr(lastbit)
            For i = 0 To UBound(PgmBitArr)
                tmpdlgStr = CStr(PgmBitArr(i)) + tmpdlgStr ''''[MSB...LSB]
                For j = 1 To gC_UDRP_USI_DSSCRepeatCyclePerBit        ''''here j start from 1
                    USI_Array(site, k) = PgmBitArr(UBound(PgmBitArr) - i)
                    k = k + 1
                Next j
            Next i
        Else
            ''''case: gL_UDRP_USI_PatBitOrder is LSB
            ''''<Notice> USI_Array(0) is LSB, so it should be PgmBitArr(0)
            For i = 0 To UBound(PgmBitArr)
                tmpdlgStr = CStr(PgmBitArr(i)) + tmpdlgStr ''''[MSB...LSB]
                For j = 1 To gC_UDRP_USI_DSSCRepeatCyclePerBit  ''''here j start from 1
                    USI_Array(site, k) = PgmBitArr(i)
                    k = k + 1
                Next j
            Next i
        End If
        ''''<NOTICE> Here gS_UDRP_USI_BitStr is Always [MSB(lastbit)...LSB(bit0)]
        gS_UDRP_USI_BitStr(site) = tmpdlgStr ''''[MSB(lastbit)...LSB(bit0)]
        If (showPrint) Then TheExec.Datalog.WriteComment ""
    Next site

    If (False) Then
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment funcName + ", Site(" + CStr(site) + "), Shift In Code [" + CStr(UBound(PgmBitArr)) + ":0]=" + gS_UDRP_USI_BitStr(site)
        Next site
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDRP_USO_COMPARE(USO_pat As Pattern, OutPin As PinList, Optional condstr As String = "all", Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_USO_COMPARE"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(USO_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_USO As New DSPWave
    Dim USO_BitStr As New SiteVariant
    Dim i As Long, j As Long, k As Integer
    Dim PatUSOArray() As String
    Dim pat_count As Long, Status As Boolean

    Dim DigCapArray() As Long
    Dim USO_PrintRow As Long
    Dim USO_BitPerRow As Long
    Dim USO_CapBits As Long
    Dim USO_TotalBit As Long

    USO_CapBits = gL_UDRP_USO_DigCapBits_Num
    USO_TotalBit = gL_UDRP_USO_DigCapBits_Num

    ''''display bit number per row
    If (gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp") Then
        USO_BitPerRow = 32
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        USO_BitPerRow = 16
    End If

    ''''20150731 update
    USO_PrintRow = IIf((USO_TotalBit Mod USO_BitPerRow) > 0, Floor(USO_TotalBit / USO_BitPerRow) + 1, Floor(USO_TotalBit / USO_BitPerRow))
    ReDim DigCapArray((USO_PrintRow * USO_BitPerRow) - 1)
 
    Dim MaxLevelIndex As Long
    Dim m_catenameVbin As String
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_defreal As String
    Dim m_resolution As Double
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_bitwidth As Long
    Dim m_decimal As Variant ''20160506 update, was Long
    Dim m_value As Variant
    Dim m_bitsum As Long
    Dim m_bitStrM As String
    Dim m_lolmt As Variant
    Dim m_hilmt As Variant
    Dim tmpdlgStr As String
    Dim tmpStrL As String
    Dim tmpStr As String
    Dim TmpVal As Variant
    Dim m_testValue As Variant
    Dim step_vdd As Long
    Dim vbinflag As Long
    Dim m_stage As String
    Dim m_tsname As String
    Dim m_siteVar As String
    Dim m_HexStr As String
    ''Dim m_vddbinEnum As Long
    Dim m_Pmode As Long
    Dim m_unitType As UnitType
    Dim m_scale As tlScaleType
    
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    
    Dim m_bitFlag_mode As Long

    m_siteVar = "UDR_PChk_Var"
    
    ''TheHdw.Patterns(USO_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    Status = GetPatListFromPatternSet(USO_pat.Value, PatUSOArray, pat_count)
    TheExec.Datalog.WriteComment ""
    
    ''''20171016 update
    ''''--------------------------------
    Dim m_testFlag As Boolean
    Dim m_CP1_Early_Flag As Boolean
    condstr = LCase(condstr)

    'If (gS_JobName = "cp1" And condStr = "cp1_early") Then
    If (gS_JobName = "cp1_early") Then
        m_CP1_Early_Flag = True
        gS_JobName = "cp1_early" ''''used to syntax check the category with stage = "cp1_early"
    Else
        m_CP1_Early_Flag = False
    End If
    ''''--------------------------------
    
If (gB_eFuse_newMethod) Then
    
    ''''----------------------------------------------------
    ''''201812XX New Method by DSPWave
    ''''----------------------------------------------------
    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
    Dim m_Fusetype As eFuseBlockType
    Dim m_SiteVarValue As New SiteLong
    Dim m_ResultFlag As New SiteLong
    Dim m_cmpResult As New SiteLong
    Dim m_PatBitOrder As String

    m_Fusetype = eFuse_UDRP
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = False
    blank_stage = True
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
    Dim m_SerialType As Boolean:: m_SerialType = False
    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
        m_SerialType = True
        gDB_SerialType = True
    End If
    
'    Dim m_CompareFlag As Boolean:: m_CompareFlag = False
'
'    For Each Site In TheExec.sites.Active
'        If (TheExec.sites(Site).SiteVariableValue(m_siteVar) = 1) Then
'            m_CompareFlag = True
'            Exit For
'        End If
'    Next
    
    'If (m_CompareFlag = True) Then
        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(0), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
        Call TheHdw.Patterns(PatUSOArray(0)).Test(pfAlways, 0)
    'End If
    

    
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
    
    ''''20160506 update
    ''''due to the additional characters "_USI_USO_compare", so plus 18.
    Call UpdateDLogColumns(gI_UDR_catename_maxLen + 18)
    
    If (gL_UDRP_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        Dim Temp_USO As New DSPWave
        Temp_USO.CreateConstant 0, USO_CapBits, DspLong
        If (gS_JobName = "cp1_early") Then
            gL_eFuse_Sim_Blank = 0
        Else
            gL_eFuse_Sim_Blank = 1
        End If
        Call rundsp.eFuse_SingleBitWave2CapWave32Bits(eFuse_UDRP, Temp_USO)
        Call auto_eFuse_print_capWave32Bits(eFuse_UDRP, Temp_USO, False) ''''True to print out
        If (m_PatBitOrder = "bit0_bitLast") Then
            For Each site In TheExec.sites
                Trim_code_USO(site) = Temp_USO(site).Copy
            Next
        Else
            Call ReverseWave(Temp_USO, Trim_code_USO, m_PatBitOrder, USO_CapBits)
        End If
    End If
    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, m_bitFlag_mode, Trim_code_USO, m_FBC, blank_stage, allblank, True, m_PatBitOrder)

    'If (TheExec.sites(Site).SiteVariableValue(m_siteVar) = 1) Then

    Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, m_bitFlag_mode, Trim_code_USO, m_FBC, m_cmpResult, , , True, m_PatBitOrder)

    TheExec.Flow.TestLimit resultVal:=m_cmpResult, lowVal:=0, hiVal:=0, Tname:=m_tsname, Unit:=m_unitType, scaletype:=m_scale
    
Exit Function

End If

    
    
    

'    For j = 0 To pat_count - 1
'        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
'        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
'
'        ''''20160506 update
'        ''''due to the additional characters "_USI_USO_compare", so plus 18.
'        Call UpdateDLogColumns(gI_UDRP_catename_maxLen + 18)
'
'        For Each Site In TheExec.sites
'            ''''initialize
'            USO_BitStr(Site) = ""  ''''MUST, and it's [bitLast ... bit0]
'            For i = 0 To UBound(DigCapArray)
'               DigCapArray(i) = 0  ''''MUST be [bit0 ... bitLast]
'            Next i
'
'            ''--------------------------------------------------------------------------------------
'            ''''20150717 update
'            ''''composite to the USO_BitStr() from the DSSC Capture
'            If (UCase(gL_UDRP_USO_PatBitOrder) = "MSB") Then
'                ''''<Notice>
'                ''''the pattern DSSC capture from MSB to LSB (bitLast to bit0)
'                ''''so Trim_code_USO.Element(0) is MSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(USO_CapBits - 1 - i)    ''''Reverse Bit String
'                    USO_BitStr(Site) = USO_BitStr(Site) & Trim_code_USO.Element(i) ''''[MSB ... LSB]
'                Next i
'            Else
'                ''''<Notice>
'                ''''the pattern DSSC capture from LSB to MSB (bit0 to bitLast)
'                ''''so Trim_code_USO.Element(0) is LSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = Trim_code_USO.Element(i)
'                    USO_BitStr(Site) = Trim_code_USO.Element(i) & USO_BitStr(Site) ''''[MSB ... LSB]
'                Next i
'            End If
'            ''--------------------------------------------------------------------------------------
'
'            ''''20160324 updae
'            ''''=============== Start of Simulated Data ===============
'            If (TheExec.TesterMode = testModeOffline) Then
'                ''''20160906 trial for the ugly codes
'                ''''<Issued codes> Shift out code [383:0]=111110000110100000000000000000001111100001101000000000000000000011111110111010000000000000000000111110100110100000000000000000001111101001101000000000000000000011111000011010000000000000000000111110000100000000000000000000001111100001000000000000000000000011111000010000000001000000010000111110000110000000000000000000001111100001000000000000000000000011111000010000000000000000000000
'                ''gS_UDRP_USI_BitStr(Site) = "111110000110100000000000000000001111100001101000000000000000000011111110111010000000000000000000111110100110100000000000000000001111101001101000000000000000000011111000011010000000000000000000111110000100000000000000000000001111100001000000000000000000000011111000010000000001000000010000111110000110000000000000000000001111100001000000000000000000000011111000010000000000000000000000"
'
'                USO_BitStr(Site) = gS_UDRP_USI_BitStr(Site)
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(gS_UDRP_USI_BitStr(Site)), i + 1, 1))
'                Next i
'
'                If (gS_JobName <> "cp1" Or TheExec.sites.Item(Site).FlagState("F_UDRP_Early_Enable") = logicTrue) Then ''''was "cp1"
'                    TmpStr = ""
'                    For i = 0 To USO_CapBits - 1
'                        If (DigCapArray(i) = 0) Then
'                            DigCapArray(i) = gL_Sim_FuseBits(Site, i)
'                        Else
'                            gL_Sim_FuseBits(Site, i) = DigCapArray(i)
'                        End If
'                        TmpStr = CStr(DigCapArray(i)) + TmpStr
'                    Next i
'                    USO_BitStr(Site) = TmpStr ''''<MUST>
'                End If
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            If (True) Then
'                '=======================================================
'                '= Print out the caputured bit data from DigCap        =
'                '=======================================================
'                TheExec.Datalog.WriteComment ""
'                Call auto_PrintAllBitbyDSSC(DigCapArray, USO_PrintRow, UBound(DigCapArray) + 1, USO_BitPerRow)
'            End If
'            ''--------------------------------------------------------------------------------------
'
'            TheExec.Datalog.WriteComment "Site(" & Site & "), UDRP USO pat:" & PatUSOArray(j) & ", Shift out code [" + CStr(USO_CapBits - 1) + ":0]=" + USO_BitStr(Site)
'            ''TheExec.Datalog.WriteComment ""
'
'            ''''----------------------------------------------------------------------------------------------
'            ''''20161114 update for print all bits (DTR) in STDF
'            ''''20171016 update to excluding "cp1_early"
'            If (m_CP1_Early_Flag = False) Then Call auto_eFuse_to_STDF_allBits("UDRP", USO_BitStr(Site))
'            ''''----------------------------------------------------------------------------------------------
'
'            ''''20150717 New
'            Call auto_Decode_UDRP_Binary_Data(DigCapArray)
'
'            ''''----------------------------------------------------------------------------------
'            ''''judge pass/fail for the specific test limit
'            tmpStrL = StrReverse(USO_BitStr(Site)) ''''translate to [LSB......MSB]
'            For i = 0 To UBound(UDRP_Fuse.Category)
'                tmpdlgStr = ""
'                m_stage = LCase(UDRP_Fuse.Category(i).Stage)
'                m_catename = UDRP_Fuse.Category(i).Name
'                m_algorithm = LCase(UDRP_Fuse.Category(i).Algorithm)
'                m_LSBbit = UDRP_Fuse.Category(i).LSBbit
'                m_MSBBit = UDRP_Fuse.Category(i).MSBbit
'                m_bitwidth = UDRP_Fuse.Category(i).Bitwidth
'                m_lolmt = UDRP_Fuse.Category(i).LoLMT
'                m_hilmt = UDRP_Fuse.Category(i).HiLMT
'                m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
'                m_resolution = UDRP_Fuse.Category(i).Resoultion
'
'                m_decimal = UDRP_Fuse.Category(i).Read.Decimal(Site)
'                m_value = UDRP_Fuse.Category(i).Read.Value(Site)
'                m_bitsum = UDRP_Fuse.Category(i).Read.BitSummation(Site)
'                m_hexStr = UDRP_Fuse.Category(i).Read.HexStr(Site)
'                m_unitType = unitNone
'                m_scale = scaleNone ''''default
'                m_tsName = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
'
'                m_bitStrM = StrReverse(Mid(tmpStrL, m_LSBbit + 1, m_bitwidth))
'
'                m_testFlag = True ''''20171016 update
'                If (m_CP1_Early_Flag = True) Then ''''20171016 update
'                    ''''only compare these category with stage="cp1_early"
'                    If (m_stage = condstr) Then
'                        ''''other cases
'                        m_testValue = m_decimal
'                    Else
'                        ''''Here it's an excluding case
'                        ''''<MUST>
'                        m_testFlag = False
'                        m_testValue = 0
'                        m_lolmt = 0
'                        m_hilmt = 0
'                    End If
'
'                ElseIf (m_algorithm = "base") Then
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
'                    ''''<Notice> User Maintain
'                    ''''Ex:: step_vdd_cpu_p1 = VDD_BIN(vdd_cpu_p1).MODE_STEP
'                    m_testValue = 0 ''''default to fail
'                    If (m_defreal = "decimal") Then ''''20160624 update
'                        m_testValue = m_decimal
'                    ElseIf (m_defreal Like "safe*voltage" Or m_defreal = "default") Then
'                        m_unitType = unitVolt
'                        m_scale = scaleMilli
'                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
'                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
'                        m_testValue = m_value * 0.001 ''''to unit:V
'                    ElseIf (m_defreal = "bincut") Then
'                        m_catenameVbin = m_catename '150127
'                        ''''' Every testing stage needs to check this formula with bincut fused value (2017/5/17, Jack)
'                        vbinflag = auto_CheckVddBinInRangeNew(m_catenameVbin, m_resolution)
'
'                        ''''20160329 Add for the offline simulation
'                        If (vbinflag = 0 And TheExec.TesterMode = testModeOffline) Then
'                            vbinflag = 1
'                        End If
'
'                        m_Pmode = VddBinStr2Enum(m_catenameVbin)
'                        tmpVal = auto_VfuseStr_to_Vdd_New(m_bitStrM, m_bitwidth, m_resolution)
'                        MaxLevelIndex = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step '150127
'                        m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(MaxLevelIndex) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(MaxLevelIndex)
'                        m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
'                        ''''judge the result
'                        If (vbinflag = 1) Then
'                            m_value = tmpVal
'                        Else
'                            m_value = -999
'                            TmpStr = m_catename + "(Site " + CStr(Site) + ") = " + CStr(tmpVal) + " is not in range" '150127
'                            TheExec.Datalog.WriteComment TmpStr
'                        End If
'                        m_unitType = unitVolt
'                        m_scale = scaleMilli
'                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
'                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
'                        m_testValue = m_value * 0.001 ''''to unit:V
'                    End If
'                Else
'                    ''''other cases, 20160927 update
'                    m_testValue = m_decimal
'                End If
'
'                ''''20160108 New
'                m_tsName = Replace(m_catename, " ", "_") ''''20151028, benefit for the script
'                Call auto_eFuse_chkLoLimit("UDRP", i, m_stage, m_lolmt)
'                Call auto_eFuse_chkHiLimit("UDRP", i, m_stage, m_hilmt)
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
'
'                ''''20171016 update
'                If (m_testFlag) Then
'                    TheExec.Flow.TestLimit resultVal:=m_testValue, lowVal:=m_lolmt, hiVal:=m_hilmt, TName:=m_tsName, Unit:=m_unitType, scaletype:=m_scale
'                End If
'            Next i
'            ''''----------------------------------------------------------------------------------
'
'            ''''20160503 update for the 2nd way to prevent for the temp sensor trim will all zero values
'            ''''20160907 update
'            Dim m_valueSum As Long
'            Dim m_matchTMPS_flag As Boolean
'            m_valueSum = 0 ''''initialize
'            m_matchTMPS_flag = False
'            m_stage = "" ''''<MUST> 20160617 update, if the "trim/tmps" is existed then m_stage has its correct value.
'            For i = 0 To UBound(UDRP_Fuse.Category)
'                m_catename = UCase(UDRP_Fuse.Category(i).Name)
'                m_algorithm = LCase(UDRP_Fuse.Category(i).Algorithm)
'                If (m_catename Like "TEMP_SENSOR*" Or m_algorithm = "tmps") Then ''''was m_algorithm = "trim", 20171103 update
'                    m_stage = LCase(UDRP_Fuse.Category(i).Stage)
'                    m_decimal = UDRP_Fuse.Category(i).Read.Decimal(Site)
'                    m_valueSum = m_valueSum + m_decimal
'                    m_matchTMPS_flag = True
'                End If
'            Next i
'            If (m_matchTMPS_flag = True) Then
'                ''''if Job >= m_stage then m_valueSim >= 1
'                If (checkJob_less_Stage_Sequence(m_stage) = False) Then
'                    TheExec.Flow.TestLimit resultVal:=m_valueSum, lowVal:=1, TName:="UDRP_TMPS_SUM"
'                    ''TheExec.Datalog.WriteComment ""
'                Else
'                    ''''if Job < m_stage then m_valueSim = 0
'                    ''''20180105 update
'                    If (gB_eFuse_Disable_ChkLMT_Flag = True) Then ''''case CP2 back to CP1 retest
'                        TheExec.Flow.TestLimit resultVal:=m_valueSum, lowVal:=0, TName:="UDRP_TMPS_SUM"
'                    Else
'                        TheExec.Flow.TestLimit resultVal:=m_valueSum, lowVal:=0, hiVal:=0, TName:="UDRP_TMPS_SUM"
'                    End If
'                End If
'            End If
'            ''''--------------------------------------------------------------------------------------------
'
'            ''''20160503 update
'            ''''compare both USI and USO for the specific stage, it's only when siteVar is '1'.
'            ''''Must be after the decode then you have the Read buffer value
'            ''''20180105 update
'            ''''The below is used to compare both USI and USO contents.
'            If (TheExec.sites(Site).SiteVariableValue(m_siteVar) = 1) Then
'                Dim m_writeBitStrM As String
'                Dim m_readBitStrM As String
'                Dim m_usiusoCmp As Long
'                m_usiusoCmp = 0 ''''<MUST> default Compare Pass:0, Fail:1
'                For i = 0 To UBound(UDRP_Fuse.Category)
'                    m_stage = LCase(UDRP_Fuse.Category(i).Stage)
'                    If (m_stage = gS_JobName) Then
'                        m_catename = UDRP_Fuse.Category(i).Name
'                        m_writeBitStrM = UCase(UDRP_Fuse.Category(i).Write.BitstrM(Site))
'                        m_readBitStrM = UCase(UDRP_Fuse.Category(i).Read.BitstrM(Site))
'                        m_tsName = Replace(m_catename, " ", "_")
'                        m_tsName = m_tsName + "_USI_USO_compare"
'                        If (m_writeBitStrM <> m_readBitStrM) Then
'                            ''''20180105 update
'                            m_usiusoCmp = 1 ''''Fail Comparison
'                            TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(Site) + "), " + m_tsName + " Failure."
'                        Else
'                            ''''case: m_writeBitStrM = m_readBitStrM
'                            ''''reserve
'                        End If
'                    End If
'                Next i
'                TheExec.Flow.TestLimit resultVal:=m_usiusoCmp, lowVal:=0, hiVal:=0, TName:="UDRP_USO_USI_Cmp"
'            End If
'            ''''--------------------------------------------------------------------------------------------
'
'            TheExec.Datalog.WriteComment ""
'            If (TheExec.sites.ActiveCount = 0) Then Exit Function 'chihome
'        Next Site
'        Call UpdateDLogColumns__False
'    Next j
'
'     ''''20171016 update
'    If (m_CP1_Early_Flag = True) Then
'        gS_JobName = "cp1" ''''reset
'    End If
'    DebugPrintFunc USO_pat.Value

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function
