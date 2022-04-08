Attribute VB_Name = "LIB_EFUSE_UDR"
Option Explicit

''''20160804 used for the offline simulation
Public Function auto_CMP_Sim_New(ByVal FuseType As eFuseBlockType, Optional showPrint As Boolean = False) As String
    'If FunctionList.Exists("auto_CMP_Sim_New") = False Then FunctionList.Add "auto_CMP_Sim_New", ""

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CMP_Sim_New"

    Dim site As Variant
    Dim i As Long, j As Long, k As Long
    
    Dim cmpBitStr As String
    Dim m_catename As String
    Dim m_algorithm As String
    Dim m_LSBbit As Long
    Dim m_MSBBit As Long
    Dim m_bitwidth As Long
    Dim tmpdlgStr As String
    
    Dim m_decimal As Variant
    Dim m_bitStrM As String
    
    Dim m_tmpDSPWave As New DSPWave
    ''''------------------------------------------------------------

'    Dim tmp_CapBitArr() As Long
'    ReDim tmp_CapBitArr(gL_CMP_DigCapBits_Num - 1)
    'Site = TheExec.Sites.SiteNumber
    
If (gB_eFuse_newMethod) Then
    Dim m_Fuseblock As EFuseCategorySyntax
    Dim m_Fuseblock2GetData As EFuseCategorySyntax
    Dim m_pgmSgWave As New DSPWave
    Dim m_BlockIndex As Long
    Dim match_Flag As Boolean:: match_Flag = False
    Dim m_Site As Variant
    Dim mL_DigCapBits_Num As Long
    'Dim Site As Variant
    
    If (FuseType = 9) Then
        m_Fuseblock = CMPFuse
        m_Fuseblock2GetData = UDRFuse
        mL_DigCapBits_Num = gL_CMP_DigCapBits_Num
    ElseIf (FuseType = 10) Then
        m_Fuseblock = CMPE_Fuse
        m_Fuseblock2GetData = UDRE_Fuse
        mL_DigCapBits_Num = gL_CMPE_DigCapBits_Num
    ElseIf (FuseType = 11) Then
        m_Fuseblock = CMPP_Fuse
        m_Fuseblock2GetData = UDRP_Fuse
        mL_DigCapBits_Num = gL_CMPP_DigCapBits_Num
    Else
    End If
    
    m_pgmSgWave.CreateConstant 0, 1, DspLong
    
    For i = 0 To UBound(m_Fuseblock.Category)
        With m_Fuseblock.Category(i)
            m_catename = .Name
            m_algorithm = LCase(.algorithm)
            m_MSBBit = .MSBbit
            m_LSBbit = .LSBbit
            m_bitwidth = .BitWidth
        End With
        
        For j = 0 To UBound(m_Fuseblock2GetData.Category)
            If (UCase(m_catename) = UCase(m_Fuseblock2GetData.Category(j).Name)) Then
                m_BlockIndex = j
                match_Flag = True
                Exit For
            End If
        Next j
            
        If (match_Flag = False) Then
            PrintDataLog "m_Fuseblock .Category(i).Name:: <" + m_catename + ">, it's NOT existed in the Category."
            GoTo errHandler
        End If
        
        m_tmpDSPWave.CreateConstant 0, m_bitwidth, DspLong
        'm_tmpDSPWave = m_Fuseblock2GetData.Category(m_BlockIndex).Read.BitArrWave
        For Each site In TheExec.sites
            m_tmpDSPWave(site) = m_Fuseblock2GetData.Category(m_BlockIndex).Read.BitArrWave(site)
            'm_pgmSgWave(Site).Select(m_LSBbit, 1, m_bitwidth).Replace m_tmpDSPWave(Site)
            m_pgmSgWave = m_pgmSgWave.Concatenate(m_tmpDSPWave.Select(0, 1, m_bitwidth).ConvertDataTypeTo(DspLong).Copy)
            'm_pgmSgWave(site).Select(m_LSBbit, 1, m_bitwidth).Replace m_tmpDSPWave(site).Select(0, 1, m_bitwidth).Copy
        Next site
        'm_pgmSgWave.Select(m_LSBbit, 1, m_bitwidth).Replace m_tmpDSPWave
        Set m_tmpDSPWave = Nothing
        
'        For Each Site In TheExec.Sites
'        'm_pgmSgWave(Site).Select(m_LSBbit, 1, m_bitwidth).Replace m_Fuseblock2GetData.Category(m_BlockIndex).Read.BitArrWave
'        Next Site
    Next i
 
    Dim m_tmp As New DSPWave
    
    For Each site In TheExec.sites
        m_tmp(site) = m_pgmSgWave.Select(1, 1, mL_DigCapBits_Num).Copy
        If (FuseType = 9) Then
            CMPFuse = m_Fuseblock
            gDW_CMP_Pgm_SingleBitWave(site) = m_tmp(site).Copy
        ElseIf (FuseType = 10) Then
            CMPE_Fuse = m_Fuseblock
            gDW_CMPE_Pgm_SingleBitWave(site) = m_tmp(site).Copy
        ElseIf (FuseType = 11) Then
            CMPP_Fuse = m_Fuseblock
            gDW_CMPP_Pgm_SingleBitWave(site) = m_tmp(site).Copy
        Else
        End If
    Next
    ''======


    
Exit Function

End If
 
    
    
    
'    For i = 0 To UBound(CMPFuse.Category)
'        m_catename = CMPFuse.Category(i).Name
'        m_algorithm = LCase(CMPFuse.Category(i).Algorithm)
'        m_LSBbit = CMPFuse.Category(i).LSBbit
'        m_MSBBit = CMPFuse.Category(i).MSBbit
'        m_bitwidth = CMPFuse.Category(i).BitWidth
'
'        If (m_LSBbit <= m_MSBBit) Then
'            For j = m_LSBbit To m_MSBBit
'                tmp_CapBitArr(j) = CLng(Mid(UDRFuse.Category(UDRIndex(m_catename)).Read.BitStrL(site), j - m_LSBbit + 1, 1))
'            Next j
'        Else
'            ''''case m_LSBbit > m_MSBbit <need to check>
'            For j = m_MSBBit To m_LSBbit
'                tmp_CapBitArr(j) = CLng(Mid(UDRFuse.Category(UDRIndex(m_catename)).Read.BitStrM(site), j - m_MSBBit + 1, 1))
'            Next j
'        End If
'
'    Next i
'
'    ''''' ken 20160804 add error bit for double confirm
'    'tmp_CapBitArr(122) = 1  '''ken for debug
'    'tmp_CapBitArr(128) = 1  '''ken for debug
'
'    '''''Composite the effective Bit String
'    cmpBitStr = ""
'    For i = 0 To gL_CMP_DigCapBits_Num - 1
'        cmpBitStr = CStr(tmp_CapBitArr(i)) + cmpBitStr ''''[bitLast...bit0] [MSB...LSB]
'    Next i
'
'    auto_CMP_Sim_New = cmpBitStr
'
'    If (showPrint) Then TheExec.Datalog.WriteComment vbTab & funcName + ":: cmpBitStr[" + CStr(gL_CMP_DigCapBits_Num - 1) + ":0] = " + cmpBitStr

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function
Public Function auto_eFuse_CMP_Parsing_HLlimit(ByVal FuseType As eFuseBlockType)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_CMP_Parsing_HLlimit"
    
    Dim i, j As Long
    Dim m_Fuseblock As EFuseCategorySyntax
    Dim m_Fuseblock2GetData As EFuseCategorySyntax
    Dim m_pgmSgWave As New DSPWave
    Dim m_BlockIndex As Long
    Dim match_Flag As Boolean:: match_Flag = False
    Dim m_Site As Variant
    Dim m_catename As String
       
       
    If (FuseType = 9) Then
        m_Fuseblock = CMPFuse
        m_Fuseblock2GetData = UDRFuse
    ElseIf (FuseType = 10) Then
        m_Fuseblock = CMPE_Fuse
        m_Fuseblock2GetData = UDRE_Fuse
    ElseIf (FuseType = 11) Then
        m_Fuseblock = CMPP_Fuse
        m_Fuseblock2GetData = UDRP_Fuse
    Else
    End If
    
    For i = 0 To UBound(m_Fuseblock.Category)
        With m_Fuseblock.Category(i)
            m_catename = .Name
            'm_stage = LCase(.Stage)
'            m_algorithm = LCase(.Algorithm)
'            m_msbbit = .MSBbit
'            m_lsbbit = .LSBbit
'            m_BitWidth = .Bitwidth
            'm_defval = .DefaultValue
            'm_defreal = LCase(.Default_Real)
            'm_resolution = .Resoultion
        End With
        
        For j = 0 To UBound(m_Fuseblock2GetData.Category)
            If (UCase(m_catename) = UCase(m_Fuseblock2GetData.Category(j).Name)) Then
                m_BlockIndex = j
                match_Flag = True
                Exit For
            End If
        Next j
            
        If (match_Flag = False) Then
            PrintDataLog "m_Fuseblock.Category(i).Name:: <" + m_catename + ">, it's NOT existed in the Category."
            GoTo errHandler
        End If
        
        With m_Fuseblock.Category(i)
'            .HiLMT = m_Fuseblock2GetData.Category(m_BlockIndex).HiLMT
            '.HiLMT_R = m_Fuseblock2GetData.Category(m_BlockIndex).HiLMT_R
            '.LoLMT = m_Fuseblock2GetData.Category(m_BlockIndex).LoLMT
            '.LoLMT_R = m_Fuseblock2GetData.Category(m_BlockIndex).LoLMT_R
            For Each m_Site In TheExec.sites
                With m_Fuseblock2GetData.Category(m_BlockIndex).Read
                m_Fuseblock.Category(i).Read.Decimal(m_Site) = .Decimal(m_Site)
                m_Fuseblock.Category(i).HiLMT = .Decimal(m_Site)
                m_Fuseblock.Category(i).LoLMT = .Decimal(m_Site)
                End With
            Next m_Site
        End With
 
    Next i
    
    
    If (FuseType = 9) Then
        CMPFuse = m_Fuseblock
    ElseIf (FuseType = 10) Then
        CMPE_Fuse = m_Fuseblock
    ElseIf (FuseType = 11) Then
        CMPP_Fuse = m_Fuseblock
    Else
    End If
    
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
