Attribute VB_Name = "LIB_EFUSE_UDRE"
Option Explicit

Public Function auto_eFuse_CMPE_Parsing_HLlimit(ByVal FuseType As eFuseBlockType)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_CMPE_Parsing_HLlimit"
    
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
