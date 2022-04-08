Attribute VB_Name = "VBT_LIB_EFUSE_UDR"
Option Explicit

''''20160805 update for UDR_Ver1 UDR CMPFuse
Public Function auto_CMP_Syntax_Chk(CMP_pat As Pattern, OutPin As PinList, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CMP_Syntax_Chk"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(CMP_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_CMP As New DSPWave
    Dim CMP_BitStr As New SiteVariant
    Dim i As Long, j As Long, k As Long
    Dim PatCMPArray() As String
    Dim pat_count As Long, Status As Boolean
    Dim fail_flag As Boolean
    ''Dim Str1 As String
    ''Dim TstNumArray() As Long, TstNumCount As Long
    Dim DigCapArray() As Long
    ''Dim ResultFlag As New SiteLong
    Dim testName As String
    Dim PinName As String
    
    Dim CMP_PrintRow As Long
    Dim CMP_BitPerRow As Long
    Dim CMP_CapBits As Long
    Dim CMP_TotalBit As Long
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    ''Dim SiteVarValue As Long
    
    Dim m_catenameUDR As String ''''UDRFuse Category name
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
    Dim m_defvalUDR As Variant ''''20160905
    Dim m_HexStr As String

    Dim Flag_CMPCategoryMatch As Boolean
    
    CMP_CapBits = gL_CMP_DigCapBits_Num
    CMP_TotalBit = gL_CMP_DigCapBits_Num

    ''''display bit number per row
    If (gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp") Then
        CMP_BitPerRow = 32
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        CMP_BitPerRow = 16
    End If

    ''''20150731 update
    CMP_PrintRow = IIf((CMP_TotalBit Mod CMP_BitPerRow) > 0, Floor(CMP_TotalBit / CMP_BitPerRow) + 1, Floor(CMP_TotalBit / CMP_BitPerRow))
    ReDim DigCapArray((CMP_PrintRow * CMP_BitPerRow) - 1)



    Dim MaxLevelIndex As Long
    Dim m_catenameVbin As String
    'Dim m_catename As String
    'Dim m_algorithm As String
    Dim m_defreal As String
    Dim m_resolution As Double
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    'Dim m_BitWidth As Long
    'Dim m_decimal As Variant ''20160506 update, was Long
    'Dim m_value As Variant
    'Dim m_bitSum As Long
    'Dim m_bitstrM As String
    'Dim m_lolmt As Variant
    'Dim m_hilmt As Variant
    'Dim tmpdlgStr As String
    Dim tmpStrL As String
    'Dim TmpStr As String
    Dim TmpVal As Variant
    Dim m_testValue As Variant
    'Dim step_vdd As Long
    Dim vbinflag As Long
    Dim m_stage As String
    Dim m_tsname As String
    Dim m_siteVar As String
    'Dim m_hexStr As String
    ''Dim m_vddbinEnum As Long
    Dim m_Pmode As Long
    Dim m_unitType As UnitType
    Dim m_scale As tlScaleType
    
    'Dim allblank As New SiteBoolean
    'Dim blank_stage As New SiteBoolean

    m_siteVar = "CMPChk_Var"
    
    
If (gB_eFuse_newMethod) Then
    TheHdw.Patterns(CMP_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    Status = GetPatListFromPatternSet(CMP_pat.Value, PatCMPArray, pat_count)
    TheExec.Datalog.WriteComment ""

    Call auto_eFuse_DSSC_DigCapSetup(PatCMPArray(0), OutPin, "CMP_cap", CMP_CapBits, Trim_code_CMP)
    Call TheHdw.Patterns(PatCMPArray(0)).Test(pfAlways, 0)
    
    Call UpdateDLogColumns(gI_CMP_catename_maxLen)
        
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        gL_eFuse_Sim_Blank = 1
        Call auto_CMP_Sim_New(eFuse_CMP, True)
'        Call RunDSP.eFuse_SingleBitWave2CapWave32Bits(eFuse_CMP, Trim_code_CMP)
'        Call auto_eFuse_print_capWave32Bits(eFuse_CMP, Trim_code_CMP, False) ''''True to print out
        For Each site In TheExec.sites
            Trim_code_CMP = gDW_CMP_Pgm_SingleBitWave.Copy
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

    m_Fusetype = eFuse_CMP
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
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_CMP, m_FBC, blank_stage, allBlank, True)

    'Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, 1, Trim_code_USO, m_FBC, m_cmpResult)


    Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

    Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
    'Call auto_eFuse_print_DSSCReadWave_Category(m_FuseType, False, gB_eFuse_printReadCate)
    Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate, gStr_PatName)

    'Call auto_eFuse_CMP_Parsing_HLlimit(m_FuseType)
    
    Dim condstr As String:: condstr = ""
    Call auto_eFuse_SyntaxCheck(m_Fusetype, condstr)

    

Exit Function

End If







'    TheHdw.Patterns(CMP_pat).Load
'    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'    Status = GetPatListFromPatternSet(CMP_pat.Value, PatCMPArray, pat_count)
'
'    For j = 0 To pat_count - 1 'Actually, we have only one pattern
'        Call auto_eFuse_DSSC_DigCapSetup(PatCMPArray(j), OutPin, "CMP_cap", CMP_CapBits, Trim_code_CMP)
'        Call TheHdw.Patterns(PatCMPArray(j)).test(pfAlways, 0)
'
'        Call UpdateDLogColumns(gI_CMP_catename_maxLen)
'
'        For Each Site In TheExec.Sites
'
'            CMP_BitStr(Site) = ""
'            gB_CMP_decode_flag(Site) = False
'
'            For i = 0 To UBound(DigCapArray)
'               DigCapArray(i) = 0  ''''MUST be [LSB(0)......MSB(lastbit)]
'            Next i
'
'            ''''20150717 update
'            ''''composite to the CMP_BitStr() from the DSSC Capture
'            If (UCase(gS_CMP_PatBitOrder) = "MSB") Then
'                ''''<Notice>
'                ''''the pattern DSSC capture from MSB to LSB (bitLast to bit0)
'                ''''so Trim_code_CMP.Element(0) is MSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To CMP_CapBits - 1
'                    DigCapArray(i) = Trim_code_CMP.Element(CMP_CapBits - 1 - i)    ''''Reverse Bit String
'                    CMP_BitStr(Site) = CMP_BitStr(Site) & Trim_code_CMP.Element(i) ''''[MSB ... LSB]
'                Next i
'            Else
'                ''''<Notice>
'                ''''the pattern DSSC capture from LSB to MSB (bit0 to bitLast)
'                ''''so Trim_code_CMP.Element(0) is LSB
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For i = 0 To CMP_CapBits - 1
'                    DigCapArray(i) = Trim_code_CMP.Element(i)
'                    CMP_BitStr(Site) = DigCapArray(i) & CMP_BitStr(Site) ''''[MSB ... LSB]
'                Next i
'            End If
'
'            ''''20160324 update
'            ''''=============== Start of Simulated Data ===============
'            If (TheExec.TesterMode = testModeOffline) Then
'                TheExec.Datalog.WriteComment vbCrLf & vbTab & "[ In Offline Mode:: Simulation Engineering Data (" + TheExec.DataManager.InstanceName + ") ]"
'                CMP_BitStr(Site) = auto_CMP_Sim_New(True) ''''<Notice> CMP_BitStr MUST be [MSB ... LSB] [bitLast...bit0]
'                For i = 0 To CMP_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(CMP_BitStr(Site)), i + 1, 1))
'                Next i
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            '=======================================================
'            '= Print out the caputured bit data from DigCap        =
'            '=======================================================
'            TheExec.Datalog.WriteComment ""
'            Call auto_PrintAllBitbyDSSC(DigCapArray, CMP_PrintRow, UBound(DigCapArray) + 1, CMP_BitPerRow)
'
'            TheExec.Datalog.WriteComment "Site(" & Site & "), CMP pat:" & PatCMPArray(j) & ", Shift out code [" + CStr(CMP_CapBits - 1) + ":0]=" + CMP_BitStr(Site)
'            TheExec.Datalog.WriteComment ""
'
'            Dim Count_1s As Long
'            Count_1s = 0
'            For i = 1 To Len(CMP_BitStr(Site))
'                If (CStr(Mid(CMP_BitStr(Site), i, 1)) = "1") Then
'                    Count_1s = Count_1s + 1
'                End If
'            Next i
'
'            TheExec.Datalog.WriteComment "Site(" & Site & "), CMP pat:" & PatCMPArray(j) & " UDRVer1 total 1s amount = " & CStr(Count_1s)
'
'            ''Str1 = TheExec.DataManager.InstanceName
'            ''Call TheExec.DataManager.GetTestNumbers(Str1, TstNumArray, TstNumCount) ''waste time in 9.0, 20180320 update
'
'            ''''20160901 update, here is decoding.
'            Call auto_Decode_CMPBinary_Data(DigCapArray)
'
'            ''''judge limit
'            For i = 0 To UBound(CMPFuse.Category)
'                tmpdlgStr = ""
'                m_catename = CMPFuse.Category(i).Name
'                m_algorithm = LCase(CMPFuse.Category(i).Algorithm)
'                m_startbit = CMPFuse.Category(i).SeqStart
'                m_endbit = CMPFuse.Category(i).SeqEnd
'                m_bitwidth = CMPFuse.Category(i).Bitwidth
'                m_decimal = CMPFuse.Category(i).Read.Decimal(Site)
'                m_value = CMPFuse.Category(i).Read.Value(Site)
'                m_hexStr = CMPFuse.Category(i).Read.HexStr(Site)
'
'                ''''initialize everytime
'                Flag_CMPCategoryMatch = False
'                m_lolmtV = -1
'                m_hilmtV = m_lolmtV
'
'                For k = 0 To UBound(UDRFuse.Category)
'                    m_catenameUDR = UCase(UDRFuse.Category(k).Name)
'                    If (m_catenameUDR = UCase(m_catename)) Then
'                        ''m_defvalUDR = UDRFuse.Category(k).DefaultValue ''''20160905
'                        m_lolmtV = UDRFuse.Category(k).Read.Decimal(Site)
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
'                        Flag_CMPCategoryMatch = True
'                        Exit For
'                    End If
'                Next k
'
'                If Flag_CMPCategoryMatch = False Then
'                    TheExec.Datalog.WriteComment ("====================================================================")
'                    TheExec.Datalog.WriteComment ("    Can't CMP Category : " & m_catename & " in UDR Category!!!")
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
'    DebugPrintFunc CMP_pat.Value
 
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDR_USI(USI_pat As Pattern, InPin As PinList, Optional condstr As String = "stage", Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDR_USI"

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
    usiarrSize = gL_USI_DigSrcBits_Num * gC_USI_DSSCRepeatCyclePerBit

    Dim PgmBitArr() As Long
    ReDim PgmBitArr(gL_USI_DigSrcBits_Num - 1)

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

    'If (gS_JobName = "cp1" And condStr = "cp1_early") Then
    If (gS_JobName = "cp1_early") Then
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
                Call auto_UDR_USI_Sim(False, False) ''''True for print debug
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
    For i = 0 To UBound(UDRFuse.Category)
        With UDRFuse.Category(i)
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
                    Call auto_eFuse_Vddbin_bincut_setWriteData(eFuse_UDR, i)
                End If
                ''''---------------------------------------------------------------------------
                With UDRFuse.Category(i)
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
        Call auto_eFuse_Gen_PgmBitSrcWave_onlyDeaultBits(eFuse_UDR, m_pgmDigSrcWave, m_pgmRes)
    Else
        ''''condStr = "stage"
        ''''Here Parameter bitFlag_mode=1 (Stage Bits)
        m_calcCRC = 0
        Call rundsp.eFuse_Gen_PgmBitSrcWave(eFuse_UDR, 1, m_pgmDigSrcWave, m_pgmRes)
    End If
    TheHdw.Wait 1# * ms
    
    Call UpdateDLogColumns(gI_UDR_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=m_pgmRes, lowVal:=1, hiVal:=1, Tname:="UDR_PGM_" + UCase(gS_JobName)
    Call UpdateDLogColumns__False

    ''''use global DSP variable
    Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_UDR_Pgm_SingleBitWave, gDL_TotalBits, gDL_ReadCycles, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_UDR, gDW_UDR_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)

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
    
    
    If (gL_USI_PatBitOrder = "LSB") Then
        For Each site In TheExec.sites
            m_tmpWave1(site) = m_pgmDigSrcWave(site).Copy.ConvertDataTypeTo(DspLong)
            m_tmpArr = m_tmpWave1(site).Data
            For i = 0 To m_size - 1
                m_USI_BitStr(site) = CStr(m_tmpArr(i)) + m_USI_BitStr(site)
            Next i
    Next
    Else
            'm_pgmDigSrcWave
    '        Dim m_tmp As New DSPWave
    '        For Each Site In TheExec.sites
    '            m_tmp(Site) = m_pgmDigSrcWave(Site).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
    '            m_pgmDigSrcWave(Site) = m_tmp(Site).ConvertDataTypeTo(DspLong)
    
        'Dim i As Long
        
    
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
    
    


errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDR_USO_Syntax_Chk(Optional condstr As String = "all") As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDR_USO_Syntax_Chk"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern

    Dim m_Fusetype As eFuseBlockType

    Dim site As Variant

    Dim i As Long, j As Long, k As Long

    ''Dim SingleStrArray() As String
    Dim SingleBitArray() As Long
    Dim DoubleBitArray() As Long

    Dim tmpStr As String


    m_Fusetype = eFuse_UDR

    Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

    Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
    Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate, gStr_PatName)
    
    Call auto_eFuse_SyntaxCheck(m_Fusetype, condstr)
    
            ''''----------------------------------------------------------------------------------------------
    If (gS_JobName <> "cp1_early") Then
        Dim m_UDR_DATA_STR As New SiteVariant
        For Each site In TheExec.sites
            DoubleBitArray = gDW_UDR_Read_DoubleBitWave.Data
            
            m_UDR_DATA_STR(site) = "" ''''is a String [(bitLast)......(bit0)]
            
            For i = 0 To UBound(DoubleBitArray)
                m_UDR_DATA_STR(site) = CStr(DoubleBitArray(i)) + m_UDR_DATA_STR(site)
            Next i
            ''TheExec.Datalog.WriteComment "gS_CFG_Direct_Access_Str=" + CStr(gS_CFG_Direct_Access_Str(Site))
    
            ''''20161114 update for print all bits (DTR) in STDF
            Call auto_eFuse_to_STDF_allBits("UDR", m_UDR_DATA_STR(site))
        Next site
    End If
    ''''----------------------------------------------------------------------------------------------

        ''''201811XX reset
    If (gS_JobName = "cp1_early") Then gS_JobName = "cp1" ''''<MUST>
Exit Function


errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function
Public Function auto_UDR_UFP(UFP_pat As Pattern, PwrPin As String, vpwr As Double, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDR_UFP"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim patt As String
    If (auto_eFuse_PatSetToPat_Validation(UFP_pat, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    ''TheHdw.Patterns(UFP_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    Call TurnOnEfusePwrPins(PwrPin, vpwr)

    Call TheHdw.Patterns(patt).Test(pfAlways, 0) ''''was UFP_pat
    DebugPrintFunc UFP_pat.Value

    Call TurnOffEfusePwrPins(PwrPin, vpwr)

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDR_UFR(UFR_pat As Pattern, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDR_UFR"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim patt As String
    If (auto_eFuse_PatSetToPat_Validation(UFR_pat, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    ''TheHdw.Patterns(UFR_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    Call TheHdw.Patterns(patt).Test(pfAlways, 0) '''' was UFR_pat
    DebugPrintFunc UFR_pat.Value

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20150731 New
Public Function auto_UDR_USO_BlankChk_byStage(USO_pat As Pattern, OutPin As PinList, _
                    Optional condstr As String = "stage", _
                    Optional m_decode_flag As Boolean = True, _
                    Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDR_USO_BlankChk_byStage"

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
    ReDim gL_Sim_FuseBits(TheExec.sites.Existing.Count - 1, gL_USO_DigCapBits_Num - 1) ''''it's for the simulation
    
    USO_CapBits = gL_USO_DigCapBits_Num
    USO_TotalBit = gL_USO_DigCapBits_Num

    ''''it's used to identify if the Job Name is existed in the UDR portion of the eFuse BitDef table.
    m_jobinStage_flag = auto_eFuse_JobExistInStage("UDR", True) ''''<MUST>

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
    m_siteVar = "UDRChk_Var"
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

    m_Fusetype = eFuse_UDR
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
     
    Call UpdateDLogColumns(gI_UDR_catename_maxLen) ''''<MUST>must after execute pattern
    
'    ''''----------------------------------------------------
'    ''''201812XX New Method by DSPWave
'    ''''----------------------------------------------------
'    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
'    Dim m_Fusetype As eFuseBlockType
'    Dim m_SiteVarValue As New SiteLong
'    Dim m_ResultFlag As New SiteLong
'    Dim m_PatBitOrder As String
'
'    m_Fusetype = eFuse_UDR
'    m_FBC = -1               ''''initialize
'    m_ResultFlag = -1        ''''initialize
'    m_SiteVarValue = -1      ''''initialize
'    allblank = True
'    blank_stage = True
'
'    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>

    If (gL_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        Dim Temp_USO As New DSPWave
        Temp_USO.CreateConstant 0, USO_CapBits, DspLong
        
        gL_eFuse_Sim_Blank = 0
'        If (gS_JobName = "cp1_early") Then
'            gL_eFuse_Sim_Blank = 0
'        Else
'            gL_eFuse_Sim_Blank = 1
'        End If
        gDW_UDR_Pgm_SingleBitWave = Temp_USO.Copy
        Call rundsp.eFuse_SingleBitWave2CapWave32Bits(eFuse_UDR, Temp_USO)
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
    condstr = LCase(condstr)
    
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
    

    
    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
        'gL_eFuse_Sim_Blank = 1
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_USO, m_FBC, blank_stage, allBlank, m_SerialType, m_PatBitOrder)

    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_USO, m_FBC, blank_stage, allblank, True, m_PatBitOrder)
    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_USO, m_FBC, blank_stage, allblank, True)

    If (blank_stage.Any(False) = True) Then
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

        Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate, gStr_PatName)

        'Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate)
    Else
        ''''all active sites (blank_early) are true
        m_ResultFlag = 0
    End If
    
    ''''Boolean True=-1, False=0
    m_SiteVarValue = blank_stage
    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
    If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only
        
    For Each site In TheExec.sites
        If (auto_eFuse_GetAllPatTestPass_Flag("UDR") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    'gL_UDR_FBC = m_FBC
    
    Call UpdateDLogColumns(gI_UDR_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "UDR_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "UDR_Blank_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    'DebugPrintFunc ReadPatSet.Value

Exit Function
End If




'    For j = 0 To pat_count - 1 'Actually, we have only one pattern
'        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
'        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
'
'        Call UpdateDLogColumns(gI_UDR_catename_maxLen)
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
'            If (UCase(gL_USO_PatBitOrder) = "MSB") Then
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
'                Call auto_UDR_USI_Sim(blank_stage(Site), False) ''''True for print debug
'                USO_BitStr(Site) = gS_USI_BitStr(Site)
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(gS_USI_BitStr(Site)), i + 1, 1))
'                    ''''it's used for the Read/Syntax simulation
'                    gL_Sim_FuseBits(Site, i) = DigCapArray(i)
'                Next i
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            ''''20150731 update
'            ''''depends on the specific Job to judge if it's blank_stage on the specific stage bits
'            testName = "UDR_BlankChk_" + UCase(gS_JobName)
'            SingleDoubleFBC = 0 ''''init
'            Call auto_eFuse_BlankChk_FBC_byStage("UDR", DigCapArray, blank_stage, SingleDoubleFBC)
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
'                    If (auto_eFuse_GetAllPatTestPass_Flag("UDR") = False) Then
'                        ResultFlag(Site) = 1   ''Fail then NO fuse
'                        PinName = "Fail"
'                        SiteVarValue = 0
'                    Else
'                        ResultFlag(Site) = 0   ''Pass Blank check
'                        PinName = "Pass"
'                        SiteVarValue = 1
'                        ''''<MUST>
'                        ''''it's used to identify if the Job Name is existed in the UDR portion of the eFuse BitDef table.
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
'            TheExec.Datalog.WriteComment "Site(" & Site & "), UDR USO pat:" & PatUSOArray(j) & ", Shift out code [" + CStr(USO_CapBits - 1) + ":0]=" + USO_BitStr(Site)
'            TheExec.Datalog.WriteComment ""
'
'            ''''20150717 New, 20160321 update
'            If (SiteVarValue <> 1 And m_decode_flag) Then Call auto_Decode_UDRBinary_Data(DigCapArray)
'
'            gB_UDR_decode_flag(Site) = False ''''20160531, <MUST> reset and init
'
'            ''''------------------------------------------------------------------------
'            ''''<Very Important>
'            ''''20171211 update to prevent Re-Fuse while UFR pattern failure
'            '''' It will casue the all Blank result to mis-leading the blank judgement
'            If (TheExec.Sites.Item(Site).FlagState("F_udr_blank") = logicTrue) Then
'                ''''Here it means that UFR read mode pattern is failure
'                TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: Site(" + CStr(Site) + ") UDR_UFR_NV Function Failure"
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
'
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDR_USO_Read_Decode(USO_pat As Pattern, OutPin As PinList, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDR_USO_Read_Decode"

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
    ReDim gL_Sim_FuseBits(TheExec.sites.Existing.Count - 1, gL_USO_DigCapBits_Num - 1) ''''it's for the simulation
    
    USO_CapBits = gL_USO_DigCapBits_Num
    USO_TotalBit = gL_USO_DigCapBits_Num

    ''''it's used to identify if the Job Name is existed in the UDR portion of the eFuse BitDef table.
    m_jobinStage_flag = auto_eFuse_JobExistInStage("UDR", True) ''''<MUST>

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
    m_siteVar = "UDRChk_Var"
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

    m_Fusetype = eFuse_UDR
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

    Call UpdateDLogColumns(gI_UDR_catename_maxLen) ''''<MUST>must after execute pattern
    
'    ''''----------------------------------------------------
'    ''''201812XX New Method by DSPWave
'    ''''----------------------------------------------------
'    Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
'    Dim m_Fusetype As eFuseBlockType
'    Dim m_SiteVarValue As New SiteLong
'    Dim m_ResultFlag As New SiteLong
'    Dim m_PatBitOrder As String
'
'    m_Fusetype = eFuse_UDR
'    m_FBC = -1               ''''initialize
'    m_ResultFlag = -1        ''''initialize
'    m_SiteVarValue = -1      ''''initialize
'    allblank = True
'    blank_stage = True
'
'    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>

    If (gL_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
'    If (TheExec.TesterMode = testModeOffline) Then
'        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
'        ''''Here MarginRead it only read the programmed Site.
'        Dim Temp_USO As New DSPWave
'        If (gS_JobName = "cp1_early") Then
'            gL_eFuse_Sim_Blank = 0
'        Else
'            gL_eFuse_Sim_Blank = 1
'        End If
'        For Each Site In TheExec.Sites
'            Call auto_UDR_USI_Sim(False, False) ''''True for print debug
'            Call eFuseENGFakeValue_Sim
'        Next Site
'        Call RunDSP.eFuse_SingleBitWave2CapWave32Bits(eFuse_UDR, Trim_code_USO)
'        Call auto_eFuse_print_capWave32Bits(eFuse_UDR, Trim_code_USO, False) ''''True to print out
'        If (True) Then
'            For Each Site In TheExec.Sites
'                Trim_code_USO(Site) = Temp_USO(Site).Copy
'            Next
'        Else
'            Call ReverseWave(Temp_USO, Trim_code_USO, m_PatBitOrder, USO_CapBits)
'        End If
'    End If
    
    'Dim condstr As String:: condstr = "stage"
    Dim m_bitFlag_mode As Long
    
'    If (condstr = "cp1_early") Then
'        m_bitFlag_mode = 0
'    ElseIf (condstr = "stage") Then
'        m_bitFlag_mode = 1
'    ElseIf (condstr = "all") Then
'        m_bitFlag_mode = 2 ''''update later, was 2
'    Else
'        ''''default, here it prevents any typo issue
''        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
''        m_FBC = -1
''        m_cmpResult = -1
'    End If
    

    
    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
        'gL_eFuse_Sim_Blank = 1
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_USO, m_FBC, blank_stage, allBlank, True, m_PatBitOrder)
    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_USO, m_FBC, blank_stage, allblank, True)

    'If (blank_stage.Any(False) = True) Then
        ''''201901XX New for TTR/PTE improvement
        Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

        Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate)
    'Else
        ''''all active sites (blank_early) are true
        'm_ResultFlag = 0
    'End If
    
    ''''Boolean True=-1, False=0
'    m_SiteVarValue = blank_stage
'    m_SiteVarValue = m_SiteVarValue.Add(2) ''''if is True +2 = 1 (Blank), is False +2 =2 (nonBlank)
'    If (m_jobinStage_flag = False) Then m_SiteVarValue = 2 ''''<MUST> just Read only
'
'    For Each Site In TheExec.Sites
'        If (auto_eFuse_GetAllPatTestPass_Flag("UDR") = True And m_FBC = 0) Then
'            m_ResultFlag = 0  'Pass Blank check criterion
'        Else
'            m_ResultFlag = 1  'Fail Blank check criterion
'            m_SiteVarValue = 0
'        End If
'        TheExec.Sites(Site).SiteVariableValue(m_siteVar) = m_SiteVarValue(Site)
'    Next Site
    'gL_UDR_FBC = m_FBC
    
    Call UpdateDLogColumns(gI_UDR_catename_maxLen)

'    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
'
'    testName = "UDR_FailBitCount_" + UCase(gS_JobName)
'    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

'    testName = "UDR_Blank_" + UCase(gS_JobName)
'    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    'DebugPrintFunc ReadPatSet.Value

Exit Function
End If
    
    
    
    
    
    
    

'    For j = 0 To pat_count - 1 'Actually, we have only one pattern
'        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
'        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
'
'        Call UpdateDLogColumns(gI_UDR_catename_maxLen) ''''<MUST>must after execute pattern
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
'            If (UCase(gL_USO_PatBitOrder) = "MSB") Then
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
'                Call auto_UDR_USI_Sim(blank_stage(Site), False) ''''True for print debug
'                USO_BitStr(Site) = gS_USI_BitStr(Site)
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(gS_USI_BitStr(Site)), i + 1, 1))
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
'            Call auto_Decode_UDRBinary_Data(DigCapArray, Not allblank(Site)) ''''True for Debug, only Not allBlank to show the decode result
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
'    Call auto_eFuse_ReadAllData_to_DictDSPWave("UDR", False, False)

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20171016 used for the stage "CP1_Early"
Public Function auto_UDR_USO_BlankChk_Early_byStage(USO_pat As Pattern, OutPin As PinList, Optional m_decode_flag As Boolean = True, Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDR_USO_BlankChk_Early_byStage"

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
    ReDim gL_Sim_FuseBits(TheExec.sites.Existing.Count - 1, gL_USO_DigCapBits_Num - 1) ''''it's for the simulation

    USO_CapBits = gL_USO_DigCapBits_Num
    USO_TotalBit = gL_USO_DigCapBits_Num

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
    m_siteVar = "UDRChk_Var"
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
    ''''It's used to check if gS_JobName is existed in all UDRFuse programming stages
    ''''it's used to identify if the Job Name is existed in the UDR portion of the eFuse BitDef table.
    Dim m_jobinStage_flag As Boolean
    m_jobinStage_flag = auto_eFuse_JobExistInStage("UDR", True) ''''<MUST>

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

    m_Fusetype = eFuse_UDR
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

    Call UpdateDLogColumns(gI_UDR_catename_maxLen) ''''<MUST>must after execute pattern


    
    If (gL_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 0, Trim_code_USO, m_FBC, blank_stage, allBlank, True, m_PatBitOrder)
    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, 1, Trim_code_USO, m_FBC, blank_stage, allblank, True)

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
        If (auto_eFuse_GetAllPatTestPass_Flag("UDR") = True And m_FBC = 0) Then
            m_ResultFlag = 0  'Pass Blank check criterion
        Else
            m_ResultFlag = 1  'Fail Blank check criterion
            m_SiteVarValue = 0
        End If
        TheExec.sites(site).SiteVariableValue(m_siteVar) = m_SiteVarValue(site)
    Next site
    'gL_UDR_FBC = m_FBC
    
    Call UpdateDLogColumns(gI_UDR_catename_maxLen)

    TheExec.Flow.TestLimit resultVal:=m_SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar '', PinName:="Value"
    
    testName = "UDR_FailBitCount_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_FBC, lowVal:=0, hiVal:=0, Tname:=testName '', PinName:="Value"

    testName = "UDR_Blank_" + UCase(gS_JobName)
    TheExec.Flow.TestLimit resultVal:=m_ResultFlag, lowVal:=0, hiVal:=0, Tname:=testName ', PinName:=PinName

    Call UpdateDLogColumns__False
    'DebugPrintFunc ReadPatSet.Value
    
Exit Function

End If





'    For j = 0 To pat_count - 1 'Actually, we have only one pattern
'        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
'        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
'
'        Call UpdateDLogColumns(gI_UDR_catename_maxLen) ''''<MUST>must after execute pattern
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
'            If (UCase(gL_USO_PatBitOrder) = "MSB") Then
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
'                Call auto_UDR_USI_Sim(blank_stage(Site), False) ''''True for print debug
'                USO_BitStr(Site) = gS_USI_BitStr(Site)
'                For i = 0 To USO_CapBits - 1
'                    DigCapArray(i) = CLng(Mid(StrReverse(gS_USI_BitStr(Site)), i + 1, 1))
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
'            testName = "UDR_BlankChk_CP1_Early" ''''"UDR_BlankChk_" + UCase(gS_JobName)
'
'            Call auto_UDR_blank_check_Early(DigCapArray, blank_early, blank_stage_noEarly)
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
'            TheExec.Datalog.WriteComment "Site(" & Site & "), UDR USO pat:" & PatUSOArray(j) & ", Shift out code [" + CStr(USO_CapBits - 1) + ":0]=" + USO_BitStr(Site)
'            TheExec.Datalog.WriteComment ""
'
'            ''''20150717 New, 20160321 update
'            If (SiteVarValue <> 1 And m_decode_flag) Then Call auto_Decode_UDRBinary_Data(DigCapArray)
'
'            gB_UDR_decode_flag(Site) = False ''''20160531, <MUST> reset and init
'
'            ''''------------------------------------------------------------------------
'            ''''<Very Important>
'            ''''20171211 update to prevent Re-Fuse while UFR pattern failure
'            '''' It will casue the all Blank result to mis-leading the blank judgement
'            If (TheExec.Sites.Item(Site).FlagState("F_udr_blank") = logicTrue) Then
'                ''''Here it means that UFR read mode pattern is failure
'                TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: Site(" + CStr(Site) + ") UDR_UFR_NV Function Failure"
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
Public Function auto_UDR_USI_Sim(m_blank As Boolean, Optional showPrint As Boolean = False) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDR_USI_Sim"

    Dim site As Variant
    Dim i As Long, j As Long, k As Long
    
    Dim usiarrSize As Long
    usiarrSize = gL_USI_DigSrcBits_Num * gC_USI_DSSCRepeatCyclePerBit

    Dim PgmBitArr() As Long
    ReDim PgmBitArr(gL_USI_DigSrcBits_Num - 1)

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
        gS_USI_BitStr(site) = ""
        For i = 0 To UBound(PgmBitArr)
            PgmBitArr(i) = 0
        Next i

        '''' 1st Step: get the PgmBitArr() per Site
        For i = 0 To UBound(UDRFuse.Category)
            tmpdlgStr = ""
            m_catename = UDRFuse.Category(i).Name
            m_algorithm = LCase(UDRFuse.Category(i).algorithm)
            m_LSBbit = UDRFuse.Category(i).LSBbit
            m_MSBBit = UDRFuse.Category(i).MSBbit
            m_bitwidth = UDRFuse.Category(i).BitWidth
            m_lolmt = UDRFuse.Category(i).LoLMT
            m_hilmt = UDRFuse.Category(i).HiLMT
            m_defval = UDRFuse.Category(i).DefaultValue
            m_defreal = LCase(UDRFuse.Category(i).Default_Real)
            m_stage = LCase(UDRFuse.Category(i).Stage)
            m_resolution = UDRFuse.Category(i).Resoultion

            ''''20150710 new datalog format
            tmpdlgStr = "Site(" + CStr(site) + ") Simulation : " + FormatNumeric(m_catename, gI_UDR_catename_maxLen)
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
                            tmpVbin = gD_BaseVoltage + m_resolution * auto_eFuse_GetWriteDecimal("UDR", m_catename, False)
                        End If
                    Else
                        tmpVbin = m_defval
                    End If
                End If

                If (m_defreal = "decimal") Then ''''20160624 update
                    m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDR", m_catename, m_defreal, m_defval, False)
                Else
                    m_decimal = auto_Vbin_to_VfuseStr_New(tmpVbin, m_bitwidth, tmpVfuse, m_resolution)
                End If

            ElseIf (m_algorithm = "app") Then
                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDR", m_catename, m_defreal, m_defval, False)

            Else ''other cases
                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDR", m_catename, m_defreal, m_defval, False)

            End If

            ''''-------------------------------------------------------------------------------------------------------
            ''''20150825 update
            Call auto_eFuse_Dec2PgmArr_Write_byStage("UDR", m_stage, i, m_decimal, m_LSBbit, m_MSBBit, PgmBitArr, False) ''''False for the Simulation
            m_decimal = UDRFuse.Category(i).Write.Decimal(site)
            m_bitStrM = UDRFuse.Category(i).Write.BitStrM(site)
            tmpStr = " [" + m_bitStrM + "]"
            If (m_algorithm = "vddbin") Then
                If (m_defreal = "decimal") Then
                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + tmpStr
                Else
                    ''''<Notice> Here using .Value to store VDDBIN value
                    UDRFuse.Category(i).Write.Value(site) = tmpVbin
                    UDRFuse.Category(i).Write.ValStr(site) = CStr(tmpVbin)
                    tmpdlgStr = tmpdlgStr + FormatNumeric(CStr(tmpVbin) + "mV", 10) + tmpStr + " = " + FormatNumeric(m_decimal, -5)
                End If
            ElseIf (m_algorithm = "base") Then ''''20160624 update
                If (m_defreal = "decimal") Then ''''20160624 update
                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + tmpStr
                Else
                    ''''<Notice> Here using .Value to store VDDBIN value
                    UDRFuse.Category(i).Write.Value(site) = TmpVal
                    UDRFuse.Category(i).Write.ValStr(site) = CStr(TmpVal)
                    tmpdlgStr = tmpdlgStr + FormatNumeric(CStr(TmpVal) + "mV", 10) + tmpStr + " = " + FormatNumeric(m_decimal, -5)
                End If
            Else
                tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + tmpStr
            End If
            ''''-------------------------------------------------------------------------------------------------------
            
            ''''20160324 update
            ''''Here it's used to simulate Blank for the specific Stage/Category
            ''If (m_blank = True And m_stage = gS_JobName) Then
            
            ''''20171016 update, black=True, set "Job<=Stage" category bits == Zero
            If (m_blank = True And (checkJob_less_Stage_Sequence(m_stage) = True Or m_stage = gS_JobName)) Then
                For j = m_LSBbit To m_MSBBit
                    PgmBitArr(j) = 0
                Next j
            End If

            If (showPrint) Then TheExec.Datalog.WriteComment tmpdlgStr
        Next i

        ''''20150717 update
        '''' 2nd Step: composite to the USI_Array() for the DSSC Source
        k = 0
        tmpdlgStr = ""
        If (UCase(gL_USI_PatBitOrder) = "MSB") Then
            ''''case: gL_USI_PatBitOrder is MSB
            ''''<Notice> USI_Array(0) is MSB, so it should be PgmBitArr(lastbit)
            For i = 0 To UBound(PgmBitArr)
                tmpdlgStr = CStr(PgmBitArr(i)) + tmpdlgStr ''''[MSB...LSB]
                For j = 1 To gC_USI_DSSCRepeatCyclePerBit        ''''here j start from 1
                    USI_Array(site, k) = PgmBitArr(UBound(PgmBitArr) - i)
                    k = k + 1
                Next j
            Next i
        Else
            ''''case: gL_USI_PatBitOrder is LSB
            ''''<Notice> USI_Array(0) is LSB, so it should be PgmBitArr(0)
            For i = 0 To UBound(PgmBitArr)
                tmpdlgStr = CStr(PgmBitArr(i)) + tmpdlgStr ''''[MSB...LSB]
                For j = 1 To gC_USI_DSSCRepeatCyclePerBit  ''''here j start from 1
                    USI_Array(site, k) = PgmBitArr(i)
                    k = k + 1
                Next j
            Next i
        End If
        ''''<NOTICE> Here gS_USI_BitStr is Always [MSB(lastbit)...LSB(bit0)]
        gS_USI_BitStr(site) = tmpdlgStr ''''[MSB(lastbit)...LSB(bit0)]
        If (showPrint) Then TheExec.Datalog.WriteComment ""
    Next site

    If (False) Then
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment funcName + ", Site(" + CStr(site) + "), Shift In Code [" + CStr(UBound(PgmBitArr)) + ":0]=" + gS_USI_BitStr(site)
        Next site
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDR_USO_COMPARE(USO_pat As Pattern, OutPin As PinList, Optional condstr As String = "all", Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDR_USO_COMPARE"

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

    USO_CapBits = gL_USO_DigCapBits_Num
    USO_TotalBit = gL_USO_DigCapBits_Num

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

    m_siteVar = "UDRChk_Var"
    
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
        'm_CP1_Early_Flag = True
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

    m_Fusetype = eFuse_UDR
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
    
    'If (m_CompareFlag = True) Then
        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(0), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
        Call TheHdw.Patterns(PatUSOArray(0)).Test(pfAlways, 0)
    'End If
    
    If (TheExec.TesterMode = testModeOffline) Then
        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
        ''''Here MarginRead it only read the programmed Site.
        If (gS_JobName = "cp1_early") Then
            gL_eFuse_Sim_Blank = 0
        Else
            gL_eFuse_Sim_Blank = 1
        End If
        Call rundsp.eFuse_SingleBitWave2CapWave32Bits(eFuse_UDR, Trim_code_USO)
        Call auto_eFuse_print_capWave32Bits(eFuse_UDR, Trim_code_USO, False) ''''True to print out
    End If
    
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
    
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, m_bitFlag_mode, Trim_code_USO, m_FBC, blank_stage, allblank, True, m_PatBitOrder)

    'If (TheExec.sites(Site).SiteVariableValue(m_siteVar) = 1) Then

    Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, 1, Trim_code_USO, m_FBC, m_cmpResult, , , True, m_PatBitOrder)

    TheExec.Flow.TestLimit resultVal:=m_cmpResult, lowVal:=0, hiVal:=0, Tname:=m_tsname, Unit:=m_unitType, scaletype:=m_scale
    
Exit Function

End If


Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

