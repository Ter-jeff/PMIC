Attribute VB_Name = "VBT_LIB_EFUSE_DAP"
Option Explicit

''''20150817 Update
''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_Efuse_DAP(DAP_pat As Pattern, OutPin As PinList, CpatureSize As Long, _
                               ByVal FuseType As String, PatBitOrder As String, _
                               Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Efuse_DAP"

    ''''-------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(DAP_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''-------------------------------------------------------------------------------------------------

    Dim pat_count As Long, Status As Boolean
    Dim i As Long, j As Long
    Dim dapWave As New DSPWave
    Dim PatDAPArray() As String
    Dim site As Variant
    Dim DAP_BitStr As New SiteVariant
    Dim DigCapArray() As Long
    ReDim DigCapArray(CpatureSize - 1)

    Dim DAP_PrintRow As Long
    Dim DAP_BitPerRow As Long
    Dim DAP_CapBits As Long
    Dim DAP_TotalBit As Long

    Dim m_cmpstr As String
    Dim m_tmpStr As String

    ''''-------------------------------------------------------------
    ''''20170630 update for efuse demo code/pattern only
    ''''-------------------------------------------------------------
    If (TheExec.TesterMode = testModeOffline) Then
        If UCase(FuseType) = "ECID" Then
            CpatureSize = EcidBitPerBlockUsed ''''20170630 update
        ElseIf (UCase(FuseType) = "CFG") Then
            CpatureSize = EConfigBitPerBlockUsed ''''20170630 update
        End If
        ReDim DigCapArray(CpatureSize - 1) ''''<MUST>
    End If
    ''''-------------------------------------------------------------

    ''''display bit number per row
    If (gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp") Then
        DAP_BitPerRow = 32
    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        DAP_BitPerRow = 16
    End If

    DAP_CapBits = CpatureSize
    DAP_TotalBit = CpatureSize

    Dim DAPcompareFlag As Long

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
   
    'In order to enter Jtag access mode, we have to set XO0 Vih =0.9v
    'TheHdw.Digital.Pins("XO0").Levels.Value(chVih) = 0.9
    TheHdw.Wait 0.0001
    Status = GetPatListFromPatternSet(DAP_pat.Value, PatDAPArray, pat_count)

'''201812XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If
    
    Dim m_Fusetype As eFuseBlockType
    Dim m_reverse As New SiteLong
    Dim m_cmpRes As New SiteLong
    Dim mDW_Read_DoubleBitWave As New DSPWave

    If UCase(FuseType) = "ECID" Then
        m_Fusetype = eFuse_ECID
        mDW_Read_DoubleBitWave = gDW_ECID_Read_DoubleBitWave
    ElseIf (UCase(FuseType) = "CFG") Then
        m_Fusetype = eFuse_CFG
        mDW_Read_DoubleBitWave = gDW_CFG_Read_DoubleBitWave
    End If
    
    If (UCase(PatBitOrder) = UCase("bit0_bitLast")) Then
        ''''the pattern DSSC capture from (bit0 to bitLast)="bit0_bitLast"
        ''''<Notice> so dapWave.Element(0) is 'bit0'
        ''''set reverse = 0
        m_reverse = 0
    Else
        ''''the pattern DSSC capture from (bitLast to bit0)="bitLast_bit0"
        ''''<Notice> so dapWave.Element(0) is 'bitLast'
        ''''set reverse = 1
        m_reverse = 1
    End If

    For i = 0 To pat_count - 1 'Actually, we have only one pattern
        Call auto_eFuse_DSSC_DigCapSetup(PatDAPArray(i), OutPin, "DAP_cap", CpatureSize, dapWave)
        Call TheHdw.Patterns(PatDAPArray(i)).Test(pfAlways, 0)
        TheExec.Datalog.WriteComment "DAP pat:" & PatDAPArray(i)

        ''''--------------------------------------------------------------------------
        If (TheExec.TesterMode = testModeOffline) Then
'            If UCase(FuseType) = "ECID" Then
'                Call RunDSP.eFuse_DspWave_Copy(gDW_ECID_Read_DoubleBitWave, dapWave)
'            ElseIf (UCase(FuseType) = "CFG") Then
'                Call RunDSP.eFuse_DspWave_Copy(gDW_CFG_Read_DoubleBitWave, dapWave)
'            End If
            Call rundsp.eFuse_DspWave_Copy(mDW_Read_DoubleBitWave, dapWave)
            PatBitOrder = UCase("bit0_bitLast")
        End If
        ''''--------------------------------------------------------------------------
        
        m_cmpRes = -1 ''''initial <MUST>
        Call rundsp.eFuse_compare_MarginRead_DoubleBitWave(m_Fusetype, dapWave, m_reverse, m_cmpRes)
        TheHdw.Wait 1# * ms
        
        Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
        If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

        Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
        Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate)

        Call UpdateDLogColumns(30)
        TheExec.Flow.TestLimit resultVal:=m_cmpRes, lowVal:=0, hiVal:=0, Tname:=UCase(FuseType) + "_DAP_" + UCase(gS_JobName)
        Call UpdateDLogColumns__False
    Next i

    ''''20160324 update
    DebugPrintFunc DAP_pat.Value

Exit Function

End If


'    For i = 0 To pat_count - 1 'Actually, we have only one pattern
'
'        Call auto_eFuse_DSSC_DigCapSetup(PatDAPArray(i), OutPin, "DAP_cap", CpatureSize, dapWave)
'        Call TheHdw.Patterns(PatDAPArray(i)).test(pfAlways, 0)
'        TheExec.Datalog.WriteComment "DAP pat:" & PatDAPArray(i)
'
'        Call UpdateDLogColumns(30)
'
'        For Each Site In TheExec.Sites
'            ''''initialize
'            For j = 0 To CpatureSize - 1
'               DigCapArray(j) = 0
'            Next j
'
'            DAP_BitStr(Site) = "" ''''always be [bitLast......bit0]
'
'            ''''20160129 add
'            ''''=============== Start of Simulated Data ===============
'            If (TheExec.TesterMode = testModeOffline) Then
'                Dim SingleBitArray() As Long
'                Dim DoubleBitArray() As Long
'                Dim m_size As Long
'                Dim m_bcnt As Long
'                Dim m_effbits As Long
'                ''''For Single-Up, DoubleBitArray == SingleBitArray
'                If UCase(FuseType) = "ECID" Then
'                    ''ReDim SingleBitArray(ECIDTotalBits - 1)
'                    ''ReDim DoubleBitArray(EcidBitPerBlockUsed - 1)
'                    Call auto_OR_2Blocks("ECID", gS_SingleStrArray, SingleBitArray, DoubleBitArray)
'                    m_effbits = EcidBitPerBlockUsed
'                    CpatureSize = EcidBitPerBlockUsed ''''20170630 update
'                ElseIf (UCase(FuseType) = "CFG") Then
'                    ''ReDim SingleBitArray(EConfigTotalBitCount - 1)
'                    ''ReDim DoubleBitArray(EConfigBitPerBlockUsed - 1)
'                    Call auto_OR_2Blocks("CFG", gS_SingleStrArray, SingleBitArray, DoubleBitArray)
'                    m_effbits = EConfigBitPerBlockUsed
'                    CpatureSize = EConfigBitPerBlockUsed ''''20170630 update
'                End If
'                DAP_CapBits = CpatureSize  ''''20170630 update
'                DAP_TotalBit = CpatureSize ''''20170630 update
'
'                ''''build up array DAPWave.Element() for the following simulation
'                If (UCase(PatBitOrder) = UCase("bitLast_bit0")) Then
'                    For j = 0 To m_effbits - 1
'                        dapWave.Element(j) = DoubleBitArray(m_effbits - j - 1)
'                    Next j
'                Else
'                    ''''"bit0_bitLast"
'                    For j = 0 To m_effbits - 1
'                        dapWave.Element(j) = DoubleBitArray(j)
'                    Next j
'                End If
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            ''''--------------------------------------------------------------------------
'            ''''20150817 update
'            ''''composite to the DAP_BitStr() from the DSSC Capture
'            If (UCase(PatBitOrder) = UCase("bitLast_bit0")) Then
'                ''''<Notice>
'                ''''the pattern DSSC capture from (bitLast to bit0)="bitLast_bit0"
'                ''''so DAPWave.Element(0) is 'bitLast'
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For j = 0 To DAP_CapBits - 1
'                    DigCapArray(j) = dapWave.Element(DAP_CapBits - 1 - j)    ''''Reverse Bit String
'                    DAP_BitStr(Site) = DAP_BitStr(Site) & dapWave.Element(j) ''''always be [bitLast...bit0]
'                Next j
'            Else
'                ''''<Notice>
'                ''''the pattern DSSC capture from (bit0 to bitLast)="bit0_bitLast"
'                ''''so DAPWave.Element(0) is 'bit0'
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For j = 0 To DAP_CapBits - 1
'                    DigCapArray(j) = dapWave.Element(j)
'                    DAP_BitStr(Site) = dapWave.Element(j) & DAP_BitStr(Site) ''''always be [bitLast...bit0]
'                Next j
'            End If
'            ''''--------------------------------------------------------------------------
'
'            '=======================================================
'            '= Print out the caputured bit data from DigCap        =
'            '=======================================================
'            'TheExec.Flow.TestLimit resultval:=0, lowval:=0, hival:=0
'            DAP_PrintRow = IIf((DAP_TotalBit Mod DAP_BitPerRow) > 0, Floor(DAP_TotalBit / DAP_BitPerRow) + 1, Floor(DAP_TotalBit / DAP_BitPerRow))
'            TheExec.Datalog.WriteComment ""
'            Call auto_PrintAllBitbyDSSC(DigCapArray, DAP_PrintRow, DAP_TotalBit, DAP_BitPerRow)
'
'            ''''TheExec.Datalog.WriteComment ""
'            TheExec.Datalog.WriteComment "Site(" & Site & "), Shift out code [" + CStr(DAP_CapBits - 1) + ":0]=" + DAP_BitStr(Site)
'
'            '============================================================================
'            '=  Judge if Direct-Access data is same as DAP(JTAG) Read-out data          =
'            '============================================================================
'            m_cmpstr = ""
'
'            If UCase(FuseType) = "ECID" Then
'                m_cmpstr = gS_ECID_Direct_Access_Str(Site)
'
'            ElseIf UCase(FuseType) = "CFG" Then
'                m_cmpstr = gS_CFG_Direct_Access_Str(Site)
'
'            Else
'                TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: Current it only supports the fusetype (ECID,CFG) !!!"
'                DAPcompareFlag = 1  'Fail
'            End If
'
'            If DAP_BitStr(Site) = m_cmpstr Then
'                DAPcompareFlag = 0  'Pass
'            Else
'                DAPcompareFlag = 1  'Fail
'            End If
'            TheExec.Flow.TestLimit resultVal:=DAPcompareFlag, lowVal:=0, hiVal:=0
'       Next Site
'
'       Call UpdateDLogColumns__False
'
'    Next i
'
'    ''''20160324 update
'    DebugPrintFunc DAP_pat.Value

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20160630 New for the whole bits by JTAG read
''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
''''20170217 update, <NOTICE> APB read for all bits not only double bits, different with DAP read.
Public Function auto_Efuse_JTAG(JTAG_pat As Pattern, OutPin As PinList, CpatureSize As Long, _
                               ByVal FuseType As String, PatBitOrder As String, _
                               Optional Validating_ As Boolean) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Efuse_JTAG"

    ''''-------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
    Dim ReadPatt As String
    If (auto_eFuse_PatSetToPat_Validation(JTAG_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''-------------------------------------------------------------------------------------------------

    Dim pat_count As Long, Status As Boolean
    Dim i As Long, j As Long
    Dim JTAGWave As New DSPWave
    Dim PatJTAGArray() As String
    Dim site As Variant
    Dim JTAG_BitStr As New SiteVariant
    Dim DigCapArray() As Long
    ReDim DigCapArray(CpatureSize - 1)

    Dim JTAG_PrintRow As Long
    Dim JTAG_BitPerRow As Long
    Dim JTAG_CapBits As Long
    Dim JTAG_TotalBit As Long

    Dim m_cmpstr As String
    Dim m_tmpStr As String

    ''''display bit number per row
    JTAG_BitPerRow = 32
''''    If (gS_EFuse_Orientation = "UP2DOWN") Or (gS_EFuse_Orientation = "SingleUp") Then
''''        JTAG_BitPerRow = 32
''''    ElseIf (gS_EFuse_Orientation = "RIGHT2LEFT") Then
''''        JTAG_BitPerRow = 16
''''    End If

    JTAG_CapBits = CpatureSize
    JTAG_TotalBit = CpatureSize

    Dim JTAGcompareFlag As Long

    'Load pattern and level/timing into tester head
    TheHdw.Patterns(JTAG_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.0001
    Status = GetPatListFromPatternSet(JTAG_pat.Value, PatJTAGArray, pat_count)


'''201812XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If
    
    Dim m_Fusetype As eFuseBlockType
    Dim m_reverse As New SiteLong
    Dim m_cmpRes As New SiteLong

    If UCase(FuseType) = "ECID" Then
        m_Fusetype = eFuse_ECID
    ElseIf (UCase(FuseType) = "CFG") Then
        m_Fusetype = eFuse_CFG
    End If
    
    If (UCase(PatBitOrder) = UCase("bit0_bitLast")) Then
        ''''the pattern DSSC capture from (bit0 to bitLast)="bit0_bitLast"
        ''''<Notice> so dapWave.Element(0) is 'bit0'
        ''''set reverse = 0
        m_reverse = 0
    Else
        ''''the pattern DSSC capture from (bitLast to bit0)="bitLast_bit0"
        ''''<Notice> so dapWave.Element(0) is 'bitLast'
        ''''set reverse = 1
        m_reverse = 1
    End If

    For i = 0 To pat_count - 1 'Actually, we have only one pattern
        Call auto_eFuse_DSSC_DigCapSetup(PatJTAGArray(i), OutPin, "JTAG_cap", CpatureSize, JTAGWave)
        Call TheHdw.Patterns(PatJTAGArray(i)).Test(pfAlways, 0)
        TheExec.Datalog.WriteComment "JTAG pat:" & PatJTAGArray(i)

        ''''--------------------------------------------------------------------------
        If (TheExec.TesterMode = testModeOffline) Then
            If UCase(FuseType) = "ECID" Then
                Call rundsp.eFuse_DspWave_Copy(gDW_ECID_Read_SingleBitWave, JTAGWave)
            ElseIf (UCase(FuseType) = "CFG") Then
                Call rundsp.eFuse_DspWave_Copy(gDW_CFG_Read_SingleBitWave, JTAGWave)
            End If
            PatBitOrder = UCase("bit0_bitLast")
        End If
        ''''--------------------------------------------------------------------------
        
        m_cmpRes = -1 ''''initial <MUST>
        Call rundsp.eFuse_compare_MarginRead_SingleBitWave(m_Fusetype, JTAGWave, m_reverse, m_cmpRes)
        TheHdw.Wait 1# * ms

        Call UpdateDLogColumns(30)
        TheExec.Flow.TestLimit resultVal:=m_cmpRes, lowVal:=0, hiVal:=0, Tname:=UCase(FuseType) + "_JTAG_" + UCase(gS_JobName)
        Call UpdateDLogColumns__False
    Next i

    ''''20160324 update
    DebugPrintFunc JTAG_pat.Value

Exit Function

End If

'    For i = 0 To pat_count - 1 'Actually, we have only one pattern
'
'        Call auto_eFuse_DSSC_DigCapSetup(PatJTAGArray(i), OutPin, "JTAG_cap", CpatureSize, JTAGWave)
'        Call TheHdw.Patterns(PatJTAGArray(i)).test(pfAlways, 0)
'        TheExec.Datalog.WriteComment "JTAG pat:" & PatJTAGArray(i)
'
'        Call UpdateDLogColumns(30)
'
'        For Each Site In TheExec.Sites
'            ''''initialize
'            For j = 0 To CpatureSize - 1
'               DigCapArray(j) = 0
'            Next j
'
'            JTAG_BitStr(Site) = "" ''''always be [bitLast......bit0]
'
'            ''''20160129 add
'            ''''=============== Start of Simulated Data ===============
'            If (TheExec.TesterMode = testModeOffline) Then
'                Dim SingleBitArray() As Long
'                Dim DoubleBitArray() As Long
'                Dim m_size As Long
'                Dim m_bcnt As Long
'                Dim m_effbits As Long
'                ''''For Single-Up, DoubleBitArray == SingleBitArray
'                If UCase(FuseType) = "ECID" Then
'                    ''ReDim SingleBitArray(ECIDTotalBits - 1)
'                    ''ReDim DoubleBitArray(EcidBitPerBlockUsed - 1)
'                    Call auto_OR_2Blocks("ECID", gS_SingleStrArray, SingleBitArray, DoubleBitArray)
'                    m_effbits = ECIDTotalBits ''''EcidBitPerBlockUsed
'                ElseIf (UCase(FuseType) = "CFG") Then
'                    ''ReDim SingleBitArray(EConfigTotalBitCount - 1)
'                    ''ReDim DoubleBitArray(EConfigBitPerBlockUsed - 1)
'                    Call auto_OR_2Blocks("CFG", gS_SingleStrArray, SingleBitArray, DoubleBitArray)
'                    m_effbits = EConfigTotalBitCount ''''EConfigBitPerBlockUsed
'                End If
'
'                ''''build up array JTAGWave.Element() for the following simulation
'                If (UCase(PatBitOrder) = UCase("bitLast_bit0")) Then
'                    For j = 0 To m_effbits - 1
'                        JTAGWave.Element(j) = SingleBitArray(m_effbits - j - 1) ''''DoubleBitArray(m_effbits - j - 1)
'                    Next j
'                Else
'                    ''''"bit0_bitLast"
'                    For j = 0 To m_effbits - 1
'                        JTAGWave.Element(j) = SingleBitArray(j) ''''DoubleBitArray(j)
'                    Next j
'                End If
'            End If
'            ''''===============   End of Simulated Data ===============
'
'            ''''--------------------------------------------------------------------------
'            ''''20150817 update
'            ''''composite to the JTAG_BitStr() from the DSSC Capture
'            If (UCase(PatBitOrder) = UCase("bitLast_bit0")) Then
'                ''''<Notice>
'                ''''the pattern DSSC capture from (bitLast to bit0)="bitLast_bit0"
'                ''''so JTAGWave.Element(0) is 'bitLast'
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For j = 0 To JTAG_CapBits - 1
'                    DigCapArray(j) = JTAGWave.Element(JTAG_CapBits - 1 - j)    ''''Reverse Bit String
'                    JTAG_BitStr(Site) = JTAG_BitStr(Site) & JTAGWave.Element(j) ''''always be [bitLast...bit0]
'                Next j
'            Else
'                ''''<Notice>
'                ''''the pattern DSSC capture from (bit0 to bitLast)="bit0_bitLast"
'                ''''so JTAGWave.Element(0) is 'bit0'
'                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
'                For j = 0 To JTAG_CapBits - 1
'                    DigCapArray(j) = JTAGWave.Element(j)
'                    JTAG_BitStr(Site) = JTAGWave.Element(j) & JTAG_BitStr(Site) ''''always be [bitLast...bit0]
'                Next j
'            End If
'            ''''--------------------------------------------------------------------------
'
'            '=======================================================
'            '= Print out the caputured bit data from DigCap        =
'            '=======================================================
'            'TheExec.Flow.TestLimit resultval:=0, lowval:=0, hival:=0
'            JTAG_PrintRow = IIf((JTAG_TotalBit Mod JTAG_BitPerRow) > 0, Floor(JTAG_TotalBit / JTAG_BitPerRow) + 1, Floor(JTAG_TotalBit / JTAG_BitPerRow))
'            TheExec.Datalog.WriteComment ""
'            Call auto_PrintAllBitbyDSSC(DigCapArray, JTAG_PrintRow, JTAG_TotalBit, JTAG_BitPerRow)
'
'            ''''TheExec.Datalog.WriteComment ""
'            TheExec.Datalog.WriteComment "Site(" & Site & "), Shift out code [" + CStr(JTAG_CapBits - 1) + ":0]=" + JTAG_BitStr(Site)
'
'            '=======================================================================
'            '=  Judge if Direct-Access data is same as JTAG Read-out data          =
'            '=======================================================================
'            m_cmpstr = ""
'
'            If UCase(FuseType) = "ECID" Then
'                m_cmpstr = gS_ECID_SingleBit_Str(Site)
'
'            ElseIf UCase(FuseType) = "CFG" Then
'                m_cmpstr = gS_CFG_SingleBit_Str(Site)
'
'            Else
'                TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: Current it only supports the fusetype (ECID,CFG) !!!"
'                JTAGcompareFlag = 1  'Fail
'            End If
'
'            If JTAG_BitStr(Site) = m_cmpstr Then
'                JTAGcompareFlag = 0  'Pass
'            Else
'                JTAGcompareFlag = 1  'Fail
'            End If
'            TheExec.Flow.TestLimit resultVal:=JTAGcompareFlag, lowVal:=0, hiVal:=0
'       Next Site
'
'       Call UpdateDLogColumns__False
'
'    Next i
'
'    ''''20160324 update
'    DebugPrintFunc JTAG_pat.Value

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

