Attribute VB_Name = "VBT_LIB_EFUSE_Pgm2File"
Public gL_ECIDFuse_PgmBits As New DSPWave
Public gL_CFGFuse_PgmBits As New DSPWave
Public gL_UIDFuse_PgmBits As New DSPWave
Public gL_SENFuse_PgmBits As New DSPWave
Public gL_MONFuse_PgmBits As New DSPWave
Public gL_UDRFuse_PgmBits As New DSPWave
Public gL_UDREFuse_PgmBits As New DSPWave
Public gL_UDRPFuse_PgmBits As New DSPWave

Public gDW_CFGFuse_TotalBits As New DSPWave

Public gL_Fuse_export_value As New SiteVariant
Public gL_Fuse_ECID_value As New SiteVariant
Public gL_Fuse_CFG_value As New SiteVariant
Public gL_Fuse_UID_value As New SiteVariant
Public gL_Fuse_MON_value As New SiteVariant
Public gL_Fuse_UDR_value As New SiteVariant
Public gL_Fuse_UDRE_value As New SiteVariant
Public gL_Fuse_UDRP_value As New SiteVariant

Public gL_Fuse_export_value_hex As New SiteVariant
Public gL_Fuse_ECID_value_hex As New SiteVariant
Public gL_Fuse_CFG_value_hex As New SiteVariant
Public gL_Fuse_UID_value_hex As New SiteVariant
Public gL_Fuse_MON_value_hex As New SiteVariant
Public gL_Fuse_UDR_value_hex As New SiteVariant
Public gL_Fuse_UDRE_value_hex As New SiteVariant
Public gL_Fuse_UDRP_value_hex As New SiteVariant

Private gL_DictCP1Fuse As New Dictionary
Public gL_ImportDataArray() As Long
Public gB_ParsePseudoFuseFile As Boolean



Public Function AddStoredFuseData(KeyName As String, ByRef obj As Variant)
On Error GoTo errHandler
    
    Dim funcName As String:: funcName = "AddStoredFuseData"
    
    KeyName = LCase(KeyName)
    If gL_DictCP1Fuse.Exists(KeyName) Then
        gL_DictCP1Fuse.Remove (KeyName)
    End If
    gL_DictCP1Fuse.Add KeyName, obj

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetStoredFuseData(KeyName As String, Optional showPrint As Boolean = False) As String
On Error GoTo errHandler
    
    Dim funcName As String:: funcName = "GetStoredFuseData"
    
    KeyName = LCase(KeyName)
    If Not gL_DictCP1Fuse.Exists(KeyName) Then
        TheExec.ErrorLogMessage "Stored capture data " & KeyName & " not found."
    Else
        GetStoredFuseData = gL_DictCP1Fuse(KeyName)
    End If

    If (showPrint) Then
        Debug.Print GetStoredFuseData
    End If
        
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_CreateConstant()
On Error GoTo errHandler

    Dim funcName As String:: funcName = "auto_CreateConstant"

    If TheExec.Flow.EnableWord("Pgm2File") Then
    
        ''20181218 UseFindHeader
        ''''ECID CreateConstant
        If (gB_findECID_flag) Then
            Dim m_tmpWave_ECID As New DSPWave
            m_tmpWave_ECID.CreateConstant 0, ECIDTotalBits, DspLong
            gL_ECIDFuse_PgmBits = m_tmpWave_ECID.Copy
        End If
        ''''CFG CreateConstant
        If (gB_findCFG_flag) Then
            Dim m_tmpWave_CFG As New DSPWave
            Dim m_tmpWave_CFG_2 As New DSPWave
            m_tmpWave_CFG.CreateConstant 0, EConfigTotalBitCount, DspLong
            gL_CFGFuse_PgmBits = m_tmpWave_CFG.Copy
        End If
        ''''UID CreateConstant
        'If (gB_findUID_flag) Then
    ''        Dim m_tmpWave_UID As New DSPWave
    ''        m_tmpWave_UID.CreateConstant 0, UIDTotalBits, DspLong
    ''        gL_UIDFuse_PgmBits = m_tmpWave_UID.Copy
        'End If
        ''''UDR CreateConstant
        'If (gB_findUDR_flag) Then
    ''        Dim m_tmpWave_UDR As New DSPWave
    ''        m_tmpWave_UDR.CreateConstant 0, gL_USI_DigSrcBits_Num, DspLong
    ''        gL_UDRFuse_PgmBits = m_tmpWave_UDR.Copy
        'End If
        'If (gB_findSEN_flag) Then
        'End If
        ''''MON CreateConstant
        If (gB_findMON_flag) Then
            Dim m_tmpWave_MON As New DSPWave
            m_tmpWave_MON.CreateConstant 0, MONITORTotalBitCount, DspLong
            gL_MONFuse_PgmBits = m_tmpWave_MON.Copy
        End If
        ''''UDRE CreateConstant
        If (gB_findUDRE_flag) Then
            Dim m_tmpWave_UDRE As New DSPWave
            m_tmpWave_UDRE.CreateConstant 0, gL_UDRE_USI_DigSrcBits_Num, DspLong
            gL_UDREFuse_PgmBits = m_tmpWave_UDRE.Copy
        End If
        ''''UDRP CreateConstant
        If (gB_findUDRP_flag) Then
            Dim m_tmpWave_UDRP As New DSPWave
            m_tmpWave_UDRP.CreateConstant 0, gL_UDRP_USI_DigSrcBits_Num, DspLong
            gL_UDRPFuse_PgmBits = m_tmpWave_UDRP.Copy
        End If
        

        gL_Fuse_ECID_value = ""
        gL_Fuse_CFG_value = ""
        gL_Fuse_UID_value = ""
        gL_Fuse_MON_value = ""
        gL_Fuse_UDR_value = ""
        gL_Fuse_UDRE_value = ""
        gL_Fuse_UDRP_value = ""

    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_fusetofile(FilePath As String) As Long
    On Error GoTo errHandler
    
    Dim funcName As String:: funcName = "auto_fusetofile"
    
    Dim Ouputdata As String
    Dim OutputFilePath As String
    Dim ProgramName As String

    ''Need to create the path on your computer
    OutputFilePath = FilePath
    
    ''20181219
    ''check folder exist and create folder
    Dim Obj_GetFuseFile As New FileSystemObject
    Dim array_tmp() As String
    Dim OutputFold As String
    
    Dim ObjShell
    Set ObjShell = VBA.CreateObject("Wscript.shell")
    
    array_tmp = Split(OutputFilePath, "\", 1)
    
    If (Obj_GetFuseFile.FolderExists(OutputFilePath)) Then
    Else
        Call Shell("cmd /k md " & OutputFilePath, 0)  ''201903 multi folder
    End If
    
    TheHdw.Wait 10 * ms
    
    ''Get current date
    Dim day As String
    day = Date
    day = Replace(day, "/", "_")

    ''Get LotID & WaferID
    Dim lot_id_tmp As String:: lot_id_tmp = TheExec.Datalog.Setup.LotSetup.LotID
    Dim wafer_id_tmp As String:: wafer_id_tmp = TheExec.Datalog.Setup.WaferSetup.ID
    
    ''Create export file name -> LotID + WaferID + ProgramName + Date + .CSV
    Dim OutputFile As String
    ProgramName = Replace(TheExec.TestProgram.Name, ".igxl", "", 1)
    OutputFile = OutputFilePath & "\" & UCase(gS_JobName) & "_" & lot_id_tmp & "_" & wafer_id_tmp & "_" & ProgramName & "_" & "TrialRun" & "_" & day & ".CSV" ''201903 OutputFileName modify
    
    If Dir(OutputFile) = Empty Then ' (TD1)check whether csv file exist or not

        Open OutputFile For Append As #45

        Dim catname As String:: catname = "(X" & "_" & "Y)"
        Dim catECID As String:: catECID = "ECID_START"
        Dim catCFG As String:: catCFG = "CFG_START"
        Dim catUID As String:: catUID = "UID_START"
        Dim catUDR As String:: catUDR = "UDR_START"
        Dim catMON As String:: catMON = "MON_START"
        Dim catUDRP As String:: catUDRP = "UDRP_START"
        Dim catUDRE As String:: catUDRE = "UDRE_START"
        Dim i As Long

        ''20181218 UseFindHeader
        If (gB_findECID_flag) Then
            For i = 0 To UBound(ECIDFuse.Category) - 1
                catECID = catECID & "," & ECIDFuse.Category(i).Name  'combine category ECID
            Next i
        End If
        
        If (gB_findCFG_flag) Then
            For i = 0 To UBound(CFGFuse.Category)
                catCFG = catCFG & "," & CFGFuse.Category(i).Name     'combine category CFG
            Next i
        End If
        
        If (gB_findUID_flag) Then
            For i = 0 To UBound(UIDFuse.Category)
                catUID = catUID & "," & UIDFuse.Category(i).Name     'combine category UID
            Next i
        End If
        
        If (gB_findMON_flag) Then
            For i = 0 To UBound(MONFuse.Category)
                catMON = catMON & "," & MONFuse.Category(i).Name    'combine category MON
            Next i
        End If
        
        If (gB_findUDRE_flag) Then
            For i = 0 To UBound(UDRE_Fuse.Category)
                catUDRE = catUDRE & "," & UDRE_Fuse.Category(i).Name    'combine category UDRE
            Next i
        End If
        
        If (gB_findUDRP_flag) Then
            For i = 0 To UBound(UDRP_Fuse.Category)
                catUDRP = catUDRP & "," & UDRP_Fuse.Category(i).Name    'combine category UDRP
            Next i
        End If
        
''        If (gB_findUDR_flag) Then End If
        
        ''Print Out; bit; def; Category
        ''Print #45, catname & "," & catECID & "," & catUID & "," & catCFG & "," & catUDR & "," & catMON & "," & "EFUSE_END"
        Print #45, catname & "," & catECID & "," & catCFG & "," & catUDRP & "," & catUDRE & "," & catMON & "," & "EFUSE_END"
        Close #45
    End If
        
    Open OutputFile For Append As #45
        For Each site In TheExec.sites
            ''dump hex value to .csv file
            Print #45, gL_Fuse_export_value & _
                        ",ECID_START" & gL_Fuse_ECID_value & _
                        ",CFG_START" & gL_Fuse_CFG_value & _
                        ",UDRP_START" & gL_Fuse_UDRP_value & _
                        ",UDRE_START" & gL_Fuse_UDRE_value & _
                        ",MON_START" & gL_Fuse_MON_value & _
                        ",EFUSE_END"
''                        ",UID_START" & gL_Fuse_UID_value
        Next site
    Close #45

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_dump_fuse_data(FuseType As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_dump_fuse_data"
    Dim reftable As String:: reftable = "EFUSE_BitDef_Table"
          
    Dim X_Tmp As String
    Dim Y_Tmp As String
    Dim site As Variant
    Dim m_algorithm As String
    Dim m_bitwidth As Long
    Dim m_name As String
    Dim m_default_real As String
    
    For Each site In TheExec.sites
        ''Get Coordinate (X,Y) value
        X_Tmp = TheExec.Datalog.Setup.WaferSetup.GetXCoord(site)
        Y_Tmp = TheExec.Datalog.Setup.WaferSetup.GetYCoord(site)
        gL_Fuse_export_value = X_Tmp & "_" & Y_Tmp
        gL_Fuse_export_value_hex = X_Tmp & "_" & Y_Tmp & "_hex"

        If FuseType = "ECID" And gB_findECID_flag Then
            ''Reset the global value
            If (Trim(gL_Fuse_ECID_value) <> "") Then
                gL_Fuse_ECID_value = ""
                gL_Fuse_ECID_value_hex = ""
            End If
            
            For i = 0 To UBound(ECIDFuse.Category) - 1
                m_algorithm = LCase(ECIDFuse.Category(i).algorithm)
                m_bitwidth = ECIDFuse.Category(i).BitWidth
                m_name = ECIDFuse.Category(i).Name
                If (m_algorithm = "lotid") Then
                    gL_Fuse_ECID_value = gL_Fuse_ECID_value & "," & ECIDFuse.Category(i).Read.Value(site)
                ElseIf (m_name = "Wafer_ID") Then
                    gL_Fuse_ECID_value = gL_Fuse_ECID_value & "," & ECIDFuse.Category(i).Read.Value(site)
                ElseIf (m_name = "X_Coordinate") Then
                    gL_Fuse_ECID_value = gL_Fuse_ECID_value & "," & ECIDFuse.Category(i).Read.Value(site)
                ElseIf (m_name = "Y_Coordinate") Then
                    gL_Fuse_ECID_value = gL_Fuse_ECID_value & "," & ECIDFuse.Category(i).Read.Value(site)
                ElseIf (m_bitwidth > 31 Or m_algorithm = "crc") Then
                    gL_Fuse_ECID_value = gL_Fuse_ECID_value & "," & ECIDFuse.Category(i).Read.HexStr(site)
                Else
                    gL_Fuse_ECID_value = gL_Fuse_ECID_value & "," & ECIDFuse.Category(i).Read.Value(site)
                End If
                ''Create hex value
                gL_Fuse_ECID_value_hex = gL_Fuse_ECID_value_hex & "," & ECIDFuse.Category(i).Read.HexStr(site)
            Next i
        ElseIf (FuseType = "CFG" And gB_findCFG_flag) Then
            ''Reset the global value
            If (Trim(gL_Fuse_CFG_value) <> "") Then
                gL_Fuse_CFG_value = ""
                gL_Fuse_CFG_value_hex = ""
            End If
            
            For i = 0 To UBound(CFGFuse.Category)
                m_algorithm = LCase(CFGFuse.Category(i).algorithm)
                m_bitwidth = CFGFuse.Category(i).BitWidth
                m_default_real = LCase(CFGFuse.Category(i).Default_Real)
                If (m_bitwidth > 31 Or m_algorithm = "crc") Then
                    gL_Fuse_CFG_value = gL_Fuse_CFG_value & "," & CFGFuse.Category(i).Read.HexStr(site)
                Else
                    gL_Fuse_CFG_value = gL_Fuse_CFG_value & "," & CFGFuse.Category(i).Read.Value(site)
                End If
                ''create hex value
                gL_Fuse_CFG_value_hex = gL_Fuse_CFG_value_hex & "," & CFGFuse.Category(i).Read.HexStr(site)
            Next i
        ElseIf (FuseType = "UID" And gB_findUID_flag) Then
            ''Reset the global value
            If (Trim(gL_Fuse_UID_value) <> "") Then gL_Fuse_UID_value = ""

            For i = 0 To UBound(UIDFuse.Category)
                m_algorithm = LCase(UIDFuse.Category(i).algorithm)
                m_bitwidth = UIDFuse.Category(i).BitWidth
                If (m_bitwidth > 31 Or m_algorithm = "crc") Then
                    gL_Fuse_UID_value = gL_Fuse_UID_value & "," & UIDFuse.Category(i).Read.HexStr(site)
                Else
                    gL_Fuse_UID_value = gL_Fuse_UID_value & "," & UIDFuse.Category(i).Read.Value(site)
                End If
            Next i
    
''        ElseIf (fusetype = "SEN") Then
''            ''''20180518, reset the global value
''            If (Trim(gL_Fuse_SEN_value) <> "") Then gL_Fuse_SEN_value = ""
''
''            For i = 0 To UBound(SENFuse.Category)
''                'gL_Fuse_SEN_value = gL_Fuse_SEN_value & "," & SENFuse.Category(i).Read.Decimal(ss)
''                m_algorithm = LCase(SENFuse.Category(i).algorithm)
''                m_bitwidth = SENFuse.Category(i).BitWidth
''                If (m_bitwidth > 31 Or m_algorithm = "crc") Then
''                    gL_Fuse_SEN_value = gL_Fuse_SEN_value & "," & SENFuse.Category(i).Read.HexStr(Site)
''                Else
''                    gL_Fuse_SEN_value = gL_Fuse_SEN_value & "," & SENFuse.Category(i).Read.Value(Site)
''                End If
''            Next i
    
        ElseIf (FuseType = "MON" And gB_findMON_flag) Then
            ''Reset the global value
            If (Trim(gL_Fuse_MON_value) <> "") Then gL_Fuse_MON_value = ""

            For i = 0 To UBound(MONFuse.Category)
                m_algorithm = LCase(MONFuse.Category(i).algorithm)
                m_bitwidth = MONFuse.Category(i).BitWidth
                If (m_bitwidth > 31 Or m_algorithm = "crc") Then
                    gL_Fuse_MON_value = gL_Fuse_MON_value & "," & MONFuse.Category(i).Read.HexStr(site)
                Else
                    gL_Fuse_MON_value = gL_Fuse_MON_value & "," & MONFuse.Category(i).Read.Value(site)
                End If
            Next i
        ElseIf (FuseType = "UDR" And gB_findUDR_flag) Then
            ''Reset the global value
            If (Trim(gL_Fuse_UDR_value) <> "") Then
                gL_Fuse_UDR_value = ""
                gL_Fuse_UDR_value_hex = ""
            End If

            For i = 0 To UBound(UDRFuse.Category)
                m_algorithm = LCase(UDRFuse.Category(i).algorithm)
                m_bitwidth = UDRFuse.Category(i).BitWidth
                m_default_real = LCase(UDRFuse.Category(i).Default_Real)
                If (m_bitwidth > 31 Or m_algorithm = "crc") Then
                    gL_Fuse_UDR_value = gL_Fuse_UDR_value & "," & UDRFuse.Category(i).Read.HexStr(site)
                Else
                    gL_Fuse_UDR_value = gL_Fuse_UDR_value & "," & UDRFuse.Category(i).Read.Value(site)
                End If
                ''create hex value
                gL_Fuse_UDR_value_hex = gL_Fuse_UDR_value_hex & "," & UDRFuse.Category(i).Read.HexStr(site)
            Next i
            
            ElseIf (FuseType = "UDRP" And gB_findUDRP_flag) Then
            ''Reset the global value
            If (Trim(gL_Fuse_UDRP_value) <> "") Then
                gL_Fuse_UDRP_value = ""
                gL_Fuse_UDRP_value_hex = ""
            End If

            For i = 0 To UBound(UDRP_Fuse.Category)
                m_algorithm = LCase(UDRP_Fuse.Category(i).algorithm)
                m_bitwidth = UDRP_Fuse.Category(i).BitWidth
                m_default_real = LCase(UDRP_Fuse.Category(i).Default_Real)
                If (m_bitwidth > 31 Or m_algorithm = "crc") Then
                    gL_Fuse_UDRP_value = gL_Fuse_UDRP_value & "," & UDRP_Fuse.Category(i).Read.HexStr(site)
                Else
                    gL_Fuse_UDRP_value = gL_Fuse_UDRP_value & "," & UDRP_Fuse.Category(i).Read.Value(site)
                End If
                
                ''Create hex value
                gL_Fuse_UDRP_value_hex = gL_Fuse_UDRP_value_hex & "," & UDRP_Fuse.Category(i).Read.HexStr(site)
            Next i
    
        ElseIf (FuseType = "UDRE" And gB_findUDRE_flag) Then
            ''Reset the global value
            If (Trim(gL_Fuse_UDRE_value) <> "") Then
                gL_Fuse_UDRE_value = ""
                gL_Fuse_UDRE_value_hex = ""
            End If

            For i = 0 To UBound(UDRE_Fuse.Category)
                m_algorithm = LCase(UDRE_Fuse.Category(i).algorithm)
                m_bitwidth = UDRE_Fuse.Category(i).BitWidth
                m_default_real = LCase(UDRE_Fuse.Category(i).Default_Real)
                If (m_bitwidth > 31 Or m_algorithm = "crc") Then
                    gL_Fuse_UDRE_value = gL_Fuse_UDRE_value & "," & UDRE_Fuse.Category(i).Read.HexStr(site)
                Else
                    gL_Fuse_UDRE_value = gL_Fuse_UDRE_value & "," & UDRE_Fuse.Category(i).Read.Value(site)
                End If
                ''Create hex value
                gL_Fuse_UDRE_value_hex = gL_Fuse_UDRE_value_hex & "," & UDRE_Fuse.Category(i).Read.HexStr(site)
            Next i
    
        Else
            TheExec.AddOutput "<Error> " + funcName + ":: No this fusetype: " + FuseType
            TheExec.Datalog.WriteComment "<Error> " + funcName + ":: No this fusetype: " + FuseType
            GoTo errHandler
        End If
    Next site

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_eFuse_pgm2file(FuseType As String, pgmWave As DSPWave, Optional b_toFile As Boolean = True) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_pgm2file"
    
    FuseType = UCase(Trim(FuseType))
    
    Dim i As Long
    Dim SingleBitArray() As Long
    Dim DoubleBitArray() As Long
    Dim singleBitSum As Long
    Dim m_dummySum As Long
    Dim SingleBitArrayStr() As String
    
    TheExec.Datalog.WriteComment vbCrLf & funcName + ":: is " + FuseType
    
    Call auto_dump_fuse_data(FuseType)
    
'    If (FuseType = "ECID") Then
'        If (b_toFile) Then Call auto_dump_fuse_data(FuseType)
'
'    ElseIf (FuseType = "CFG") Then
'        Call auto_dump_fuse_data(FuseType)
'
'    ElseIf (FuseType = "UID") Then
'        For Each Site In TheExec.Sites
'            SingleBitArray = pgmWave.Data
'            Call auto_Gen_DoubleBitArray(FuseType, SingleBitArray, DoubleBitArray, m_dummySum)
'            Call auto_Decode_UIDBinary_Data(DoubleBitArray, True)
'        Next Site
'        Call auto_dump_fuse_data(FuseType)
'
'    ElseIf (FuseType = "SEN") Then
'        For Each Site In TheExec.Sites
'            SingleBitArray = pgmWave.Data
'            Call auto_Gen_DoubleBitArray(FuseType, SingleBitArray, DoubleBitArray, m_dummySum)
'            'Call auto_Decode_SENBinary_Data(DoubleBitArray, True)
'        Next Site
'        Call auto_dump_fuse_data(FuseType)
'
'    ElseIf (FuseType = "MON") Then
'        For Each Site In TheExec.Sites
'            SingleBitArray = pgmWave.Data
'            Call auto_Gen_DoubleBitArray(FuseType, SingleBitArray, DoubleBitArray, m_dummySum)
'            Call auto_Decode_MONBinary_Data(DoubleBitArray, True)
'        Next Site
'        Call auto_dump_fuse_data(FuseType)
'
'    ElseIf (FuseType = "UDR") Then
'        For Each Site In TheExec.Sites
'            DoubleBitArray = pgmWave.Data
'            Call auto_Decode_UDRBinary_Data(DoubleBitArray, True)
'        Next Site
'        Call auto_dump_fuse_data(FuseType)
'
'    ElseIf (FuseType = "UDRE") Then
'        For Each Site In TheExec.Sites
'            DoubleBitArray = pgmWave.Data
'            Call auto_Decode_UDRE_Binary_Data(DoubleBitArray, True)
'        Next Site
'        Call auto_dump_fuse_data(FuseType)
'
'    ElseIf (FuseType = "UDRP") Then
'        For Each Site In TheExec.Sites
'            DoubleBitArray = pgmWave.Data
'            Call auto_Decode_UDRP_Binary_Data(DoubleBitArray, True)
'        Next Site
'        Call auto_dump_fuse_data(FuseType)
'
'    Else
'    End If
        
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function


Public Function auto_eFuse_SetReadValue(ByVal FuseType As String, _
                                        NameIndex As Long, _
                                        m_value As Variant, _
                                        m_algorithm As String, _
                                        m_DefaultReal As String, _
                                        Optional showPrint As Boolean = True) As Variant
                    
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_eFuse_SetReadValue"

    Dim m_len As Integer
    Dim m_dlogstr As String
    Dim ss As Variant
    Dim m_resolution As Double
    Dim m_DataStr As String
    Dim m_catename As String
    Dim i, j As Long
    Dim m_MSBBit As Long
    Dim m_LSBbit As Long
    Dim m_bitwidth As Long
    Dim m_binarr() As Long
    Dim m_HighBit As Long
    Dim m_LowBit As Long
    Dim m_FuseValue As Variant
    Dim m_MSBFirst As Boolean:: m_MSBFirst = False
    Dim m_HexStr As String
    Dim m_BinStr As String

    If (True) Then
        
        Dim m_FuseCate As EFuseCategoryParamSyntax
        
        m_DataStr = ""
        ss = TheExec.sites.SiteNumber
        m_FuseValue = m_value
    
        If (FuseType = "ECID") Then
            m_FuseCate = ECIDFuse.Category(NameIndex)
        ElseIf (FuseType = "CFG") Then
            m_FuseCate = CFGFuse.Category(NameIndex)
        ElseIf (FuseType = "UID") Then
        ElseIf (FuseType = "UDRP") Then
            m_FuseCate = UDRP_Fuse.Category(NameIndex)
        ElseIf (FuseType = "UDRE") Then
            m_FuseCate = UDRE_Fuse.Category(NameIndex)
        ElseIf (FuseType = "MON") Then
            m_FuseCate = MONFuse.Category(NameIndex)
        End If
    
        With m_FuseCate
            m_catename = .Name
            m_LSBbit = .LSBbit
            m_MSBBit = .MSBbit
            m_bitwidth = .BitWidth
        End With
        
        m_LowBit = m_LSBbit
        m_HighBit = m_MSBBit
        
        If (m_MSBBit < m_LSBbit) Then
            m_LowBit = m_MSBBit
            m_HighBit = m_LSBbit
            m_MSBFirst = True
        End If
        
        m_FuseCate.Read.ValStr(ss) = CStr(m_value)
    
        If (m_algorithm = "LOTID") Then
            For i = 1 To EcidCharPerLotId
                m_DataStr = m_DataStr + auto_MappingCharToBinStr(Mid(CStr(m_value), i, 1))
            Next i
            m_FuseCate.Read.BitStrM(ss) = StrReverse(CStr(m_DataStr))
            m_FuseCate.Read.BitStrL(ss) = CStr(m_DataStr)
        
        ElseIf (m_algorithm = "IDS") Then
            m_FuseCate.Read.ValStr(ss) = Format(m_value, "0.0000") + "mA"
            m_resolution = m_FuseCate.Resoultion
            If m_resolution <= 0 Then m_resolution = 1
            m_FuseValue = CDbl(m_value / m_resolution)
        
        ElseIf (m_algorithm = "VDDBIN") Then
            m_FuseCate.Read.ValStr(ss) = Format(m_value, "0.000#") + "mV"
            m_resolution = m_FuseCate.Resoultion
            If m_resolution <= 0 Then m_resolution = 1
            If (UCase(m_DefaultReal) = "DECIMAL") Then
                m_FuseValue = m_value
                m_value = m_value * m_resolution + gD_BaseVoltage
                m_FuseCate.Read.ValStr(ss) = m_value
            Else
                If m_value < gD_BaseVoltage Then
                    m_FuseValue = 0
                Else
                    m_FuseValue = CDbl((m_value - gD_BaseVoltage) / m_resolution)
                End If
            End If
            
        ElseIf (m_algorithm = "BASE") Then
            m_resolution = m_FuseCate.Resoultion
            If m_resolution = 0 Then m_resolution = gD_UDRE_BaseStepVoltage
            If (UCase(m_FuseCate.Default_Real) Like "*SAFE VOLTAGE") Then
                'm_value = gD_UDRP_VBaseFuse
                m_FuseValue = CDbl((m_value) / m_resolution) - 1
            Else
                m_FuseValue = m_value
                m_value = m_value * m_resolution + 1
            End If
            m_FuseCate.Read.ValStr(ss) = Format(m_value, "0.000#") + "mV"
            'm_value = gD_UDRP_VBaseFuse
        ElseIf (m_algorithm = "CRC" Or (m_bitwidth >= 32 And m_MSBFirst = False)) Then
            Call auto_HexStr2BinStr_EFUSE(CStr(m_value), m_bitwidth, m_binarr)
            m_FuseCate.Read.HexStr(ss) = CStr(m_value)
        ElseIf (m_bitwidth >= 32 And m_MSBFirst = True) Then
            Call auto_HexStr2BinStr_EFUSE(CStr(m_value), m_bitwidth, m_binarr)
                Dim m_TmpBinarr() As Long
                ReDim m_TmpBinarr(m_bitwidth)
                For i = 0 To m_bitwidth - 1
                    m_TmpBinarr(m_bitwidth - 1 - i) = m_binarr(i)
                Next
                Call auto_eFuse_bitArr_to_binStr_HexStr(m_TmpBinarr, m_BinStr, m_HexStr)
                m_binarr = m_TmpBinarr
                m_FuseCate.Read.HexStr(ss) = m_HexStr
                m_FuseCate.Read.ValStr(ss) = m_HexStr
                m_FuseValue = m_HexStr
                m_value = m_HexStr
        Else
            
        End If
        
        If (m_bitwidth < 32 And m_algorithm <> "CRC") Then
            Call auto_Dec2Bin_EFuse(m_FuseValue, m_bitwidth, m_binarr)
        End If
        
        
        If (m_MSBFirst = False Or m_algorithm = "CRC") Then
            j = 0
            For i = m_LowBit To m_HighBit
                gL_ImportDataArray(i) = m_binarr(j)
                j = j + 1
            Next i
        Else
            If (m_algorithm <> "LOTID") Then
                j = m_bitwidth - 1
                For i = m_LowBit To m_HighBit
                    gL_ImportDataArray(i) = m_binarr(j)
                    j = j - 1
                Next i
            Else
                For i = m_LowBit To m_HighBit
                    gL_ImportDataArray(i) = CLng(Mid(m_DataStr, i + 1, 1))
                Next i
            End If
        End If
        
        m_FuseCate.Read.Decimal(ss) = m_FuseValue
        m_FuseCate.Read.Value(ss) = m_value
        
        m_DataStr = CStr(m_FuseValue)
    Else

'    ss = TheExec.Sites.SiteNumber
'    FuseType = UCase(Trim(FuseType))
'    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)
'    m_DataStr = ""
'
'    If (FuseType = "ECID") Then
'        m_catename = ECIDFuse.Category(NameIndex).Name
'        m_LSBbit = ECIDFuse.Category(NameIndex).LSBbit
'        m_MSBBit = ECIDFuse.Category(NameIndex).MSBbit
'        m_bitwidth = ECIDFuse.Category(NameIndex).Bitwidth
'
'        ''20181221 modify
'        m_LowBit = m_LSBbit
'        m_HighBit = m_MSBBit
'
'        If (m_MSBBit < m_LSBbit) Then
'            m_LowBit = m_MSBBit
'            m_HighBit = m_LSBbit
'        End If
'
'        If (m_algorithm = "LOTID") Then
'            For i = 1 To EcidCharPerLotId
'                m_DataStr = m_DataStr + auto_MappingCharToBinStr(Mid(CStr(m_value), i, 1))
'            Next i
'            ECIDFuse.Category(NameIndex).Read.BitstrM(ss) = StrReverse(CStr(m_DataStr))
'            ECIDFuse.Category(NameIndex).Read.BitStrL(ss) = CStr(m_DataStr)
'
'        ElseIf (m_algorithm = "CRC" Or m_bitwidth >= 32) Then
'            ECIDFuse.Category(NameIndex).Read.HexStr(ss) = CStr(m_value)
'            Call auto_HexStr2BinStr_EFUSE(CStr(m_value), m_bitwidth, m_binarr)
'
'        Else
'            Call auto_Dec2Bin_EFuse(m_value, m_bitwidth, m_binarr)
'        End If
'
'        'If (m_Algorithm <> "LOTID") Then
'        If (UCase(ECIDFuse.Category(NameIndex).MSBFirst) <> "Y" Or m_algorithm = "CRC") Then
'            j = 0
'            For i = m_LowBit To m_HighBit
'                gL_ImportDataArray(i) = m_binarr(j)
'                j = j + 1
'            Next i
'        Else
'            If (m_algorithm <> "LOTID") Then
'                j = m_bitwidth - 1
'            For i = m_LowBit To m_HighBit
'                gL_ImportDataArray(i) = m_binarr(j)
'                j = j - 1
'            Next i
'            Else
'                For i = m_LowBit To m_HighBit
'                    gL_ImportDataArray(i) = CLng(Mid(m_DataStr, i + 1, 1))
'                Next i
'            End If
'        End If
'
'        ECIDFuse.Category(NameIndex).Read.Value(ss) = m_value
'        m_DataStr = CStr(m_value)
'
'    ElseIf (FuseType = "CFG") Then
'        m_catename = CFGFuse.Category(NameIndex).Name
'        m_LSBbit = CFGFuse.Category(NameIndex).LSBbit
'        m_MSBBit = CFGFuse.Category(NameIndex).MSBbit
'        m_bitwidth = CFGFuse.Category(NameIndex).Bitwidth
'
'        If (m_algorithm = "IDS") Then
'            CFGFuse.Category(NameIndex).Write.Decimal(ss) = m_value
'            CFGFuse.Category(NameIndex).Read.ValStr(ss) = Format(m_value, "0.0000") + "mA"
'            m_resolution = CFGFuse.Category(NameIndex).Resoultion
'            If m_resolution <= 0 Then m_resolution = 1
'            m_value = CDbl(m_value / m_resolution)
'
'        ElseIf (m_algorithm = "VDDBIN") Then
'            CFGFuse.Category(NameIndex).Read.ValStr(ss) = Format(m_value, "0.000#") + "mV"
'            m_resolution = CFGFuse.Category(NameIndex).Resoultion
'            If m_resolution <= 0 Then m_resolution = 1
'            m_value = CDbl((m_value - gD_BaseVoltage) / m_resolution)
'            CFGFuse.Category(NameIndex).Write.Decimal(ss) = m_value
'        ElseIf (m_algorithm = "CRC" Or m_algorithm = "CONDNA" Or m_bitwidth >= 32) Then
'            CFGFuse.Category(NameIndex).Read.HexStr(ss) = CStr(m_value)
'            Call auto_HexStr2BinStr_EFUSE(CStr(m_value), m_bitwidth, m_binarr)
'        End If
'
'        If (m_bitwidth < 32 And m_algorithm <> "CRC" And m_algorithm <> "CONDNA") Then
'            Call auto_Dec2Bin_EFuse(m_value, m_bitwidth, m_binarr)
'        End If
'
'        j = 0
'        For i = m_LSBbit To m_MSBBit
'            gL_ImportDataArray(i) = m_binarr(j)
'            j = j + 1
'        Next i
'
'        CFGFuse.Category(NameIndex).Read.Value(ss) = m_value
'        m_DataStr = CStr(m_value)
'
'    ElseIf (FuseType = "UID") Then
'    ElseIf (FuseType = "UDRP") Then
'        m_catename = UDRP_Fuse.Category(NameIndex).Name
'        m_LSBbit = UDRP_Fuse.Category(NameIndex).LSBbit
'        m_MSBBit = UDRP_Fuse.Category(NameIndex).MSBbit
'        m_bitwidth = UDRP_Fuse.Category(NameIndex).Bitwidth
'
'        If (m_algorithm = "IDS") Then
'            UDRP_Fuse.Category(NameIndex).Write.Decimal(ss) = m_value
'            UDRP_Fuse.Category(NameIndex).Read.ValStr(ss) = Format(m_value, "0.0000") + "mA"
'            m_resolution = UDRP_Fuse.Category(NameIndex).Resoultion
'            If m_resolution <= 0 Then m_resolution = 1
'            m_value = CDbl(m_value / m_resolution)
'
'        ElseIf (m_algorithm = "VDDBIN") Then
'            UDRP_Fuse.Category(NameIndex).Read.ValStr(ss) = Format(m_value, "0.000#") + "mV"
'            m_resolution = UDRP_Fuse.Category(NameIndex).Resoultion
'            If m_resolution <= 0 Then m_resolution = 1
'            m_value = CDbl((m_value - gD_UDRP_BaseVoltage) / m_resolution)
'            UDRP_Fuse.Category(NameIndex).Write.Decimal(ss) = m_value
'        ElseIf (m_algorithm = "BASE") Then
'            UDRP_Fuse.Category(NameIndex).Write.Decimal(ss) = gD_UDRP_BaseVoltage
'            UDRP_Fuse.Category(NameIndex).Read.ValStr(ss) = Format(m_value, "0.000#") + "mV"
'''            m_resolution = UDRP_Fuse.Category(NameIndex).Resoultion
'''            If m_resolution <= 0 Then m_resolution = 1
'            m_value = gD_UDRP_VBaseFuse
'        ElseIf (m_bitwidth >= 32) Then
'            UDRP_Fuse.Category(NameIndex).Read.HexStr(ss) = CStr(m_value)
'            Call auto_HexStr2BinStr_EFUSE(CStr(m_value), m_bitwidth, m_binarr)
'        Else
'            UDRP_Fuse.Category(NameIndex).Read.ValStr(ss) = CStr(m_value)
'        End If
'
'        If (m_bitwidth < 32) Then
'            Call auto_Dec2Bin_EFuse(m_value, m_bitwidth, m_binarr)
'        End If
'
'        j = 0
'        For i = m_LSBbit To m_MSBBit
'            gL_ImportDataArray(i) = m_binarr(j)
'            j = j + 1
'        Next i
'
'        UDRP_Fuse.Category(NameIndex).Read.Value(ss) = m_value
'        m_DataStr = CStr(m_value)
'
'    ElseIf (FuseType = "UDRE") Then
'        m_catename = UDRE_Fuse.Category(NameIndex).Name
'        m_LSBbit = UDRE_Fuse.Category(NameIndex).LSBbit
'        m_MSBBit = UDRE_Fuse.Category(NameIndex).MSBbit
'        m_bitwidth = UDRE_Fuse.Category(NameIndex).Bitwidth
'
'        If (m_algorithm = "IDS") Then
'            UDRE_Fuse.Category(NameIndex).Write.Decimal(ss) = m_value
'            UDRE_Fuse.Category(NameIndex).Read.ValStr(ss) = Format(m_value, "0.0000") + "mA"
'            m_resolution = UDRE_Fuse.Category(NameIndex).Resoultion
'            If m_resolution <= 0 Then m_resolution = 1
'            m_value = CDbl(m_value / m_resolution)
'
'        ElseIf (m_algorithm = "VDDBIN") Then
'            UDRE_Fuse.Category(NameIndex).Read.ValStr(ss) = Format(m_value, "0.000#") + "mV"
'            m_resolution = UDRE_Fuse.Category(NameIndex).Resoultion
'            If m_resolution <= 0 Then m_resolution = 1
'            m_value = CDbl((m_value - gD_UDRE_BaseVoltage) / m_resolution)
'            UDRE_Fuse.Category(NameIndex).Write.Decimal(ss) = m_value
'        ElseIf (m_algorithm = "BASE") Then
'            UDRE_Fuse.Category(NameIndex).Write.Decimal(ss) = gD_UDRE_BaseVoltage
'            UDRE_Fuse.Category(NameIndex).Read.ValStr(ss) = Format(m_value, "0.000#") + "mV"
'            m_value = gD_UDRE_VBaseFuse
'        ElseIf (m_bitwidth >= 32) Then
'            UDRE_Fuse.Category(NameIndex).Read.HexStr(ss) = CStr(m_value)
'            Call auto_HexStr2BinStr_EFUSE(CStr(m_value), m_bitwidth, m_binarr)
'        Else
'            UDRE_Fuse.Category(NameIndex).Read.ValStr(ss) = CStr(m_value)
'        End If
'
'        If (m_bitwidth < 32) Then
'            Call auto_Dec2Bin_EFuse(m_value, m_bitwidth, m_binarr)
'        End If
'
'        j = 0
'        For i = m_LSBbit To m_MSBBit
'            gL_ImportDataArray(i) = m_binarr(j)
'            j = j + 1
'        Next i
'
'        UDRE_Fuse.Category(NameIndex).Read.Value(ss) = m_value
'        m_DataStr = CStr(m_value)
'
'    ElseIf (FuseType = "MON") Then
'        m_catename = MONFuse.Category(NameIndex).Name
'        m_LSBbit = MONFuse.Category(NameIndex).LSBbit
'        m_MSBBit = MONFuse.Category(NameIndex).MSBbit
'        m_bitwidth = MONFuse.Category(NameIndex).Bitwidth
'
'        If (m_algorithm = "CRC" Or m_bitwidth >= 32) Then
'            MONFuse.Category(NameIndex).Read.HexStr(ss) = CStr(m_value)
'            Call auto_HexStr2BinStr_EFUSE(CStr(m_value), m_bitwidth, m_binarr)
'        Else
'            Call auto_Dec2Bin_EFuse(m_value, m_bitwidth, m_binarr)
'        End If
'
'        j = 0
'        For i = m_LSBbit To m_MSBBit
'            gL_ImportDataArray(i) = m_binarr(j)
'            j = j + 1
'        Next i
'
'        MONFuse.Category(NameIndex).Read.Value(ss) = m_value
'        m_DataStr = CStr(m_value)
'    End If
'
'    auto_eFuse_SetReadValue = m_value
    
    End If
    
    If (FuseType = "ECID") Then
        ECIDFuse.Category(NameIndex) = m_FuseCate
    ElseIf (FuseType = "CFG") Then
        CFGFuse.Category(NameIndex) = m_FuseCate
    ElseIf (FuseType = "UID") Then
    ElseIf (FuseType = "UDRP") Then
        UDRP_Fuse.Category(NameIndex) = m_FuseCate
    ElseIf (FuseType = "UDRE") Then
        UDRE_Fuse.Category(NameIndex) = m_FuseCate
    ElseIf (FuseType = "MON") Then
        MONFuse.Category(NameIndex) = m_FuseCate
    End If

    If (showPrint) Then
        FuseType = FormatNumeric(FuseType, 4)
        FuseType = FuseType + FormatNumeric("Fuse SetReadDecimal", -25)
        m_dlogstr = vbTab & "Site(" + CStr(ss) + ") " + FuseType + m_catename + " = "
        If (m_DataStr = "") Then
            m_dlogstr = m_dlogstr + FormatNumeric(m_value, -10)
        Else
            m_dlogstr = m_dlogstr + m_DataStr
        End If
        TheExec.Datalog.WriteComment m_dlogstr
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function auto_filetofuse(FilePath As String, FileNameKeyword As String) As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_filetofuse"
    
    Dim m_file As String    ':: m_file = "D:\TMP_TRY\temp.csv"
    Dim m_lineStr As String
    Dim Key As String
    Dim addr As Variant
    Dim Data As Variant
    Dim bitdef As Variant
    Dim debug_mode As Boolean 'debug_mode = true print datalog
    Dim FileName As String
    debug_mode = True
    
    Dim lot_id_tmp As String:: lot_id_tmp = TheExec.Datalog.Setup.LotSetup.LotID
    Dim wafer_id_tmp As String:: wafer_id_tmp = TheExec.Datalog.Setup.WaferSetup.ID

    Dim Obj_GetFuseFile As New FileSystemObject
    Dim Obj_FuseFolder As Folder
    Dim Obj_FuseFile As File
    Dim FileName_tmp As String
    Dim counter As Integer:: counter = 0
    
    Dim array_tmp() As String
    Dim OutputFold As String
    
    If (gB_ParsePseudoFuseFile = False) Then
    
        array_tmp = Split(FilePath, "\", 1)
        
        ''20181219
        ''check folder exist and create folder
        If (Obj_GetFuseFile.FolderExists(FilePath)) Then
        Else
             Call Shell("cmd /k md " & FilePath, 0) ''201903 multi folder
        End If
        
        ''check file exist
        FileName_tmp = LCase("*" & lot_id_tmp & "_" & wafer_id_tmp & "_" & "*" & FileNameKeyword & ".csv")
        
        Set Obj_FuseFolder = Obj_GetFuseFile.GetFolder(FilePath)
    
        For Each Obj_FuseFile In Obj_FuseFolder.Files
            If (LCase(Obj_FuseFile.Name) Like FileName_tmp) Then
                FileName = Obj_FuseFile.Name
                counter = counter + 1
                m_file = Obj_FuseFile.Path
                If (counter >= 2) Then
                    TheExec.Datalog.WriteComment "<Error> File:: " + FileName_tmp + " , please check it out!!"
                    GoTo errHandler
                End If
            End If
        Next
        
        If (counter = 0) Then
            TheExec.Datalog.WriteComment "<Error> File:: " + FileName_tmp + " doesn't exist, please check it!"
            GoTo errHandler
        End If
    
        Open m_file For Input As #1
            Do Until EOF(1)
                ''Create Efuse Dictionary
                Line Input #1, m_lineStr
                m_lineStr = Trim(m_lineStr)
                addr = InStr(1, m_lineStr, ",")
                Key = Mid(m_lineStr, 1, addr - 1)
                Data = Mid(m_lineStr, addr + 1)
                Call AddStoredFuseData(Key, Data)
            Loop
        Close #1
        gB_ParsePseudoFuseFile = True
    End If

    ''Get data
    ''20181218 UseFindHeader
    If (gB_findECID_flag) Then
        Dim ECID_BitDef() As String
        ReDim ECID_BitDef(UBound(ECIDFuse.Category) - 1)
        Dim ECID_file_value() As Variant
        ReDim ECID_file_value(UBound(ECIDFuse.Category) - 1)
    End If
    If (gB_findCFG_flag) Then
        Dim CFG_BitDef() As String
        ReDim CFG_BitDef(UBound(CFGFuse.Category))
        Dim CFG_file_value() As Variant
        ReDim CFG_file_value(UBound(CFGFuse.Category))
    End If
    If (gB_findUID_flag) Then
        Dim UID_BitDef() As String
        'ReDim UID_BitDef(UBound(UIDFuse.Category))
        ReDim UID_BitDef(200)
        Dim UID_file_value() As Variant
        'ReDim UID_file_value(UBound(UIDFuse.Category))
        ReDim UID_file_value(200)
    End If
    If (gB_findUDR_flag) Then
        Dim UDR_BitDef() As String
        'ReDim UDR_BitDef(UBound(UDRFuse.Category))
        'Dim UDR_file_value() As Variant
        'ReDim UDR_file_value(UBound(UDRFuse.Category))
    End If
    If (gB_findMON_flag) Then
        Dim MON_BitDef() As String
        ReDim MON_BitDef(UBound(MONFuse.Category))
        Dim Mon_file_value() As Variant
        ReDim Mon_file_value(UBound(MONFuse.Category))
    End If
    If (gB_findUDRE_flag) Then
        Dim UDRE_BitDef() As String
        ReDim UDRE_BitDef(UBound(UDRE_Fuse.Category))
        Dim UDRE_file_value() As Variant
        ReDim UDRE_file_value(UBound(UDRE_Fuse.Category))
    End If
    If (gB_findUDRP_flag) Then
        Dim UDRP_BitDef() As String
        ReDim UDRP_BitDef(UBound(UDRP_Fuse.Category))
        Dim UDRP_file_value() As Variant
        ReDim UDRP_file_value(UBound(UDRP_Fuse.Category))
    End If
    
    Dim idx_ECID As Integer
    Dim idx_UID As Integer
    Dim idx_CFG As Integer
    'Dim idx_UDR As Integer
    Dim idx_UDRP As Integer
    Dim idx_UDRE As Integer
    Dim idx_MON As Integer
    Dim idx_END As Integer

    Dim m_lineArr_header() As String
    Dim m_lineArr_data() As String
    Dim m_lineArr_elem As String

    Dim X_Tmp As String
    Dim Y_Tmp As String
    Dim site As Variant
    Dim i As Long
    Dim m_NameIndex As Long
    Dim m_algorithm As String
    Dim m_DefaultReal As String
    Dim m_tmpData As New DSPWave
    
       
    bitdef = GetStoredFuseData("(X_Y)")
    m_lineArr_header = Split(bitdef, ",")
    For i = 0 To UBound(m_lineArr_header)
        m_lineArr_elem = UCase(Trim(m_lineArr_header(i)))
        If (m_lineArr_elem = UCase("ECID_START")) Then
            idx_ECID = i
''        ElseIf (m_lineArr_elem = UCase("UID_START")) Then
''            idx_UID = i
        ElseIf (m_lineArr_elem = UCase("CFG_START")) Then
            idx_CFG = i
''        ElseIf (m_lineArr_elem = UCase("UDR_START")) Then
''            idx_UDR = i
        ElseIf (m_lineArr_elem = UCase("UDRP_START")) Then
            idx_UDRP = i
        ElseIf (m_lineArr_elem = UCase("UDRE_START")) Then
            idx_UDRE = i
        ElseIf (m_lineArr_elem = UCase("MON_START")) Then
            idx_MON = i
        ElseIf (m_lineArr_elem = UCase("EFUSE_END")) Then
            idx_END = i
        End If
    Next i
    
    For Each site In TheExec.sites
        ''Get data from Dictionary
        X_Tmp = TheExec.Datalog.Setup.WaferSetup.GetXCoord(site)
        Y_Tmp = TheExec.Datalog.Setup.WaferSetup.GetYCoord(site)
        Key = X_Tmp & "_" & Y_Tmp
        
        Data = GetStoredFuseData(Key)
        m_lineArr_data = Split(Data, ",")
        bitdef = GetStoredFuseData("(X_Y)")
        m_lineArr_header = Split(bitdef, ",")
        
        ''20181218 UseFindHeader
        '-------------read ECID data write to ECID_file_value()
        ''''Total ECID Cate#  = (idx_UID) - (idx_ECID)
        'For i = 0 To (idx_UID - idx_ECID) - 2
        If (gB_findECID_flag) Then
            ReDim gL_ImportDataArray(ECIDTotalBits - 1)
            m_tmpData.CreateConstant 0, ECIDTotalBits, DspLong
            
            For i = 0 To (idx_CFG - idx_ECID) - 2
                ECID_BitDef(i) = UCase(Trim(m_lineArr_header(idx_ECID + 1 + i)))
                ECID_file_value(i) = UCase(Trim(m_lineArr_data(idx_ECID + 1 + i)))
                m_NameIndex = ECIDIndex(ECID_BitDef(i))
                m_algorithm = UCase(ECIDFuse.Category(m_NameIndex).algorithm)
                m_DefaultReal = ECIDFuse.Category(m_NameIndex).Default_Real
    
                Call auto_eFuse_SetReadValue("ECID", m_NameIndex, ECID_file_value(i), m_algorithm, m_DefaultReal, debug_mode)
                
'                If IsNumeric(ECID_file_value(i)) Then
'                    Call auto_eFuse_SetReadDecimal("ECID", ECID_BitDef(i), ECID_file_value(i), False)
'                End If
            Next i
            
            Call Gen_ProgramBitArray(ECIDTotalBits, EcidReadCycle, EcidBitsPerRow, EcidReadBitWidth)
            
            
            m_tmpData.Data = gL_ImportDataArray
            gL_ECIDFuse_PgmBits = gL_ECIDFuse_PgmBits.BitwiseOr(m_tmpData)
            
            If (debug_mode) Then TheExec.Datalog.WriteComment ""
        End If
        
''        '-------------read UID data write to UID_file_value()
''        If (gB_findUID_flag)  Then
''        For i = 0 To (idx_CFG - idx_UID) - 2
''            UID_BitDef(i) = UCase(Trim(m_lineArr_header(idx_UID + 1 + i)))
''            UID_file_value(i) = UCase(Trim(m_lineArr_data(idx_UID + 1 + i)))
''            Call auto_eFuse_SetReadValue("UID", UID_BitDef(i), UID_file_value(i), debug_mode)
''        Next i
''        If (debug_mode) Then TheExec.DataLog.WriteComment ""
''        endif
        
        ''Read CFG data write to CFG_file_value()
        If (gB_findCFG_flag) Then
            ReDim gL_ImportDataArray(EConfigTotalBitCount)
            m_tmpData.CreateConstant 0, EConfigTotalBitCount, DspLong
            For i = 0 To (idx_UDRP - 1 - idx_CFG) - 2
                CFG_BitDef(i) = UCase(Trim(m_lineArr_header(idx_CFG + 1 + i)))
                CFG_file_value(i) = UCase(Trim(m_lineArr_data(idx_CFG + 1 + i)))
                m_NameIndex = CFGIndex(CFG_BitDef(i))
                m_algorithm = UCase(CFGFuse.Category(m_NameIndex).algorithm)
                m_DefaultReal = CFGFuse.Category(m_NameIndex).Default_Real
                
                Call auto_eFuse_SetReadValue("CFG", m_NameIndex, CFG_file_value(i), m_algorithm, m_DefaultReal, debug_mode)
                
'                If IsNumeric(CFG_file_value(i)) Then
'                    Call auto_eFuse_SetReadDecimal("CFG", CFG_BitDef(i), CFG_file_value(i), False)
'                End If
            Next i
            
            Call Gen_ProgramBitArray(EConfigTotalBitCount, EConfigReadCycle, EConfigBitsPerRow, EConfigReadBitWidth)
            
            m_tmpData.Data = gL_ImportDataArray
            gL_CFGFuse_PgmBits = gL_CFGFuse_PgmBits.BitwiseOr(m_tmpData)
            
            If (debug_mode) Then TheExec.Datalog.WriteComment ""
        End If
        
        'read UDR data write to UDR_file_value()
''        If (gB_findUDR_flag)  Then
''        For i = 0 To (idx_END - idx_UDR) - 2
''            UDR_BitDef(i) = UCase(Trim(m_lineArr_header(idx_UDR + 1 + i)))
''            UDR_file_value(i) = UCase(Trim(m_lineArr_data(idx_UDR + 1 + i)))
''            Call auto_eFuse_SetReadValue("UDR", UDR_BitDef(i), UDR_file_value(i), debug_mode)
''
''        Next i
''        If (debug_mode) Then TheExec.DataLog.WriteComment ""
''        endif
        
        If (gB_findUDRP_flag) Then
            ReDim gL_ImportDataArray(gL_UDRP_USO_DigCapBits_Num - 1)
            m_tmpData.CreateConstant 0, gL_UDRP_USO_DigCapBits_Num, DspLong
            For i = 0 To (idx_UDRE - idx_UDRP) - 2
                UDRP_BitDef(i) = UCase(Trim(m_lineArr_header(idx_UDRP + 1 + i)))
                UDRP_file_value(i) = UCase(Trim(m_lineArr_data(idx_UDRP + 1 + i)))
                m_NameIndex = UDRP_Index(UDRP_BitDef(i))
                m_algorithm = UCase(UDRP_Fuse.Category(m_NameIndex).algorithm)
                m_DefaultReal = UDRP_Fuse.Category(m_NameIndex).Default_Real
                
                Call auto_eFuse_SetReadValue("UDRP", m_NameIndex, UDRP_file_value(i), m_algorithm, m_DefaultReal, debug_mode)
    
'                If IsNumeric(UDRP_file_value(i)) Then
'                    Call auto_eFuse_SetReadDecimal("UDRP", UDRP_BitDef(i), UDRP_file_value(i), False)
'                End If
            Next i
            
            m_tmpData.Data = gL_ImportDataArray
            
            If (gL_UDRE_USI_PatBitOrder = "LSB") Then
            Else
                    'm_pgmDigSrcWave
            '        Dim m_tmp As New DSPWave
            '        For Each Site In TheExec.sites
            '            m_tmp(Site) = m_pgmDigSrcWave(Site).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
            '            m_pgmDigSrcWave(Site) = m_tmp(Site).ConvertDataTypeTo(DspLong)
            
                'Dim i As Long
                Dim m_size As Long
                Dim m_tmpArr() As Long
                Dim m_outArr() As Long
                Dim m_tmpWave1 As New DSPWave
                Dim outWave As New DSPWave
                
                'outWave.CreateConstant 0, m_size, DspLong
                'For Each Site In TheExec.Sites
                    m_size = m_tmpData(site).SampleSize
                'Next
                
                outWave.CreateConstant 0, m_size, DspLong
                
                'For Each Site In TheExec.Sites
                    m_tmpWave1(site) = m_tmpData(site).Copy.ConvertDataTypeTo(DspLong)
                    m_tmpArr = m_tmpWave1(site).Data
                    m_outArr = outWave(site).Data
                        For i = 0 To m_size - 1
                    ''outWave.Element(i) = m_tmpWave.Element(m_size - i - 1) ''''waste TT
                        m_outArr(i) = m_tmpArr(m_size - i - 1) ''''save TT
                    Next i
                
                outWave(site).Data = m_outArr ''''save TT
                'Next
            End If
    
            
            gL_UDRPFuse_PgmBits = gL_UDRPFuse_PgmBits.BitwiseOr(outWave)
            
            If (debug_mode) Then TheExec.Datalog.WriteComment ""
        End If
        
        If (gB_findUDRE_flag) Then
            ReDim gL_ImportDataArray(gL_UDRE_USO_DigCapBits_Num - 1)
            m_tmpData.CreateConstant 0, gL_UDRE_USO_DigCapBits_Num, DspLong
            
            For i = 0 To (idx_MON - idx_UDRE) - 2
                UDRE_BitDef(i) = UCase(Trim(m_lineArr_header(idx_UDRE + 1 + i)))
                UDRE_file_value(i) = UCase(Trim(m_lineArr_data(idx_UDRE + 1 + i)))
                m_NameIndex = UDRE_Index(UDRE_BitDef(i))
                m_algorithm = UCase(UDRE_Fuse.Category(m_NameIndex).algorithm)
                m_DefaultReal = UDRE_Fuse.Category(m_NameIndex).Default_Real
                Call auto_eFuse_SetReadValue("UDRE", m_NameIndex, UDRE_file_value(i), m_algorithm, m_DefaultReal, debug_mode)
                
'                If IsNumeric(UDRE_file_value(i)) Then
'                    Call auto_eFuse_SetReadDecimal("UDRE", UDRE_BitDef(i), UDRE_file_value(i), False)
'                End If
            Next i
            
            m_tmpData.Data = gL_ImportDataArray
            
            If (gL_UDRP_USI_PatBitOrder = "LSB") Then
            Else
                    'm_pgmDigSrcWave
            '        Dim m_tmp As New DSPWave
            '        For Each Site In TheExec.sites
            '            m_tmp(Site) = m_pgmDigSrcWave(Site).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
            '            m_pgmDigSrcWave(Site) = m_tmp(Site).ConvertDataTypeTo(DspLong)
            
'                Dim i As Long
'                Dim m_size As Long
'                Dim m_tmpArr() As Long
'                Dim m_outArr() As Long
'                Dim m_tmpWave1 As New DSPWave
'                Dim outWave As New DSPWave
                
                'outWave.CreateConstant 0, m_size, DspLong
                'For Each Site In TheExec.Sites
                    m_size = m_tmpData(site).SampleSize
                'Next
                
                outWave.CreateConstant 0, m_size, DspLong
                
                'For Each Site In TheExec.Sites
                    m_tmpWave1(site) = m_tmpData(site).Copy.ConvertDataTypeTo(DspLong)
                    m_tmpArr = m_tmpWave1(site).Data
                    m_outArr = outWave(site).Data
                        For i = 0 To m_size - 1
                    ''outWave.Element(i) = m_tmpWave.Element(m_size - i - 1) ''''waste TT
                        m_outArr(i) = m_tmpArr(m_size - i - 1) ''''save TT
                    Next i
                
                outWave(site).Data = m_outArr ''''save TT
                'Next
            End If
            
            gL_UDREFuse_PgmBits = gL_UDREFuse_PgmBits.BitwiseOr(outWave)

            If (debug_mode) Then TheExec.Datalog.WriteComment ""
        End If
        
        ''Read MON data write to Mon_file_value()
        If (gB_findMON_flag) Then
            ReDim gL_ImportDataArray(MONITORTotalBitCount)
            m_tmpData.CreateConstant 0, MONITORTotalBitCount, DspLong
            For i = 0 To (idx_END - idx_MON) - 2
                MON_BitDef(i) = UCase(Trim(m_lineArr_header(idx_MON + 1 + i)))
                Mon_file_value(i) = UCase(Trim(m_lineArr_data(idx_MON + 1 + i)))
                m_NameIndex = MONIndex(MON_BitDef(i))
                m_algorithm = UCase(MONFuse.Category(m_NameIndex).algorithm)
                m_DefaultReal = MONFuse.Category(m_NameIndex).Default_Real
                Call auto_eFuse_SetReadValue("MON", m_NameIndex, Mon_file_value(i), m_algorithm, m_DefaultReal, debug_mode)
                
'                 If IsNumeric(Mon_file_value(i)) Then
'                     Call auto_eFuse_SetReadDecimal("MON", MON_BitDef(i), Mon_file_value(i), False)
'                 End If
            Next i
            
            Call Gen_ProgramBitArray(MONITORTotalBitCount, MONITORReadCycle, MONITORBitsPerRow, MONITORReadBitWidth)
            
            m_tmpData.Data = gL_ImportDataArray
            gL_MONFuse_PgmBits = gL_MONFuse_PgmBits.BitwiseOr(m_tmpData)
            
            If (debug_mode) Then TheExec.Datalog.WriteComment ""
        
        End If

    Next site
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Gen_ProgramBitArray(m_TotalSize As Long, _
                                    m_ReadCycle As Long, _
                                    m_BitsPerRow As Long, _
                                    m_ReadBitWidth As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Gen_ProgramBitArray"

    Dim i, j, k As Long
    Dim k1, k2 As Long
    
    If (gS_EFuse_Orientation = "RIGHT2LEFT") Then
        Dim TempeFuseArray() As Long
        ReDim TempeFuseArray(m_TotalSize)
        
        For i = 0 To (m_TotalSize - 1)
            TempeFuseArray(i) = gL_ImportDataArray(i)
        Next i
        
        k = 0 ''''must be here
        For i = 0 To m_ReadCycle - 1
            For j = 0 To m_BitsPerRow - 1
                ''k1: Right block
                ''k2:  Left block
                k1 = (i * m_ReadBitWidth) + j
                k2 = (i * m_ReadBitWidth) + m_BitsPerRow + j
                gL_ImportDataArray(k1) = TempeFuseArray(k)
                gL_ImportDataArray(k2) = TempeFuseArray(k)
                k = k + 1
            Next j
        Next i
    End If

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_ConfigWrite_byCondition_Pgm2File(condstr As String)
'Public Function auto_ConfigWrite_byCondition_Pgm2File(WritePattSet As Pattern, _
'                                                      condstr As String, _
'                                                      Optional catename_grp As String, _
'                                                      Optional Validating_ As Boolean)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ConfigWrite_byCondition"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Write patterns  =
    '==================================
    ''''20161114 update to Validate/load pattern
    'Dim WritePatt As String
    'If (auto_eFuse_PatSetToPat_Validation(WritePattSet, WritePatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim eFuse_Pgm_Bit() As Long
    Dim Expand_eFuse_Pgm_Bit() As Long
    Dim eFusePatCompare() As String
    Dim i As Long
    Dim SegmentSize As Long
    Dim Expand_Size As Long

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
    
    condstr = LCase(condstr) ''''<MUST>
    
    ''20181120, add for pgm2file
    Dim m_tmpData As New DSPWave
    m_tmpData.CreateConstant 0, EConfigTotalBitCount, DspLong

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
    Dim m_Fusetype As eFuseBlockType
    m_Fusetype = eFuse_CFG
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    gDL_eFuse_Orientation = gE_eFuse_Orientation
    
    m_pgmWave.CreateConstant 0, gDL_BitsPerBlock, DspLong ''''gDL_BitsPerBlock = EConfigBitPerBlockUsed
    ''''<MUST and Importance>
    For Each site In TheExec.sites
        gDW_Pgm_RawBitWave = m_pgmWave.Copy
    Next site

    If (condstr = "cp1_early") Then
        m_cmpStage = "cp1_early"
        gS_JobName = "cp1_early"
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
    
    For Each site In TheExec.sites
        m_tmpData(site) = gDW_CFG_Pgm_SingleBitWave(site).Copy 'was eFuse_Pgm_Bit
        gL_CFGFuse_PgmBits(site) = gL_CFGFuse_PgmBits(site).BitwiseOr(m_tmpData(site))
    Next
    
    If (m_cmpStage = "cp1_early") Then
        gS_JobName = "cp1"
    End If
    
Exit Function
    
End If

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_ConfigSingleDoubleBit_Pgm2File(Optional condstr As String = "")

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_ConfigSingleDoubleBit"
    
    Dim site As Variant
    Dim i As Long, j As Long, k As Long
    Dim DoubleBitArray() As Long
    Dim SingleBitArray() As Long
    Dim tmpStr As String
    Dim crcBinStr As String
    Dim m_siteVar As String
    Dim FuseStr As String:: FuseStr = "CFG"
    m_siteVar = "CFGChk_Var"
    ''''--------------------------------------------------------------------------
    
    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    If (TheExec.EnableWord("Pgm2File")) Then ''''for Pgm2File (Pgm2Read)
        gL_CFG_FBC = 0 ''''set dummy
        
        Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
        Dim m_Fusetype As eFuseBlockType
        Dim m_SiteVarValue As New SiteLong
        Dim m_ResultFlag As New SiteLong
        Dim m_bitFlag_mode As Long
        Dim blank_stage As New SiteBoolean
        Dim allBlank As New SiteBoolean
        Dim CapWave As New DSPWave
    
        gDL_eFuse_Orientation = gE_eFuse_Orientation
        gL_eFuse_Sim_Blank = 0
        
        m_Fusetype = eFuse_CFG
        m_FBC = -1               ''''initialize
        m_ResultFlag = -1        ''''initialize
        m_SiteVarValue = -1      ''''initialize
        allBlank = True
        blank_stage = True
        
        Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
        
        condstr = "all"
        If (LCase(condstr = "cp1_early")) Then
            m_bitFlag_mode = 0
        ElseIf (LCase(condstr) = "stage") Then
            m_bitFlag_mode = 1
        ElseIf (LCase(condstr) = "all") Then
            m_bitFlag_mode = 2
        ElseIf (LCase(condstr) = "real") Then
            m_bitFlag_mode = 3
        Else
            ''''default, here it prevents any typo issue
            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
            m_FBC = -1
            'm_cmpResult = -1
        End If
        
        For Each site In TheExec.sites
            CapWave(site) = gL_CFGFuse_PgmBits(site).ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb)
        Next site

        Call rundsp.eFuse_Wave32bits_to_SingleDoubleBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, blank_stage, allBlank)
    End If
    
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
        End If
        
    End If
    
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
    Call auto_eFuse_pgm2file(FuseStr, gL_CFGFuse_PgmBits)

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

        For Each site In TheExec.sites.Existing
            ''''initialize
            tmpCFG_First_bitStr = ""

            ''''1st: get the tmpXXXstr per site
            For i = 0 To UBound(CFGFuse.Category)
                m_bitStrM = CFGFuse.Category(i).Read.BitStrM(site)
                m_algorithm = LCase(CFGFuse.Category(i).algorithm)
                If (m_algorithm = "cond") Then
                    tmpCFG_First_bitStr = tmpCFG_First_bitStr + m_bitStrM
                End If
            Next i

            ''''2nd: integrate tmpXXXstr to allXXXstr for iEDA register key
            If (site = TheExec.sites.Existing.Count - 1) Then
                allCFG_First_bitStr = allCFG_First_bitStr + tmpCFG_First_bitStr
            Else
                allCFG_First_bitStr = allCFG_First_bitStr + tmpCFG_First_bitStr + ","
            End If
        Next site

        allCFG_First_bitStr = auto_checkIEDAString(allCFG_First_bitStr)

        TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName
        TheExec.Datalog.WriteComment " ConfigRead (all sites iEDA format)::"
        TheExec.Datalog.WriteComment " " + m_regName_CFG_Cond + " = " & allCFG_First_bitStr
        TheExec.Datalog.WriteComment ""

        Call RegKeySave(m_regName_CFG_Cond, allCFG_First_bitStr)
    End If
    ''''----------------------------------------------------------------------------------------
    
    'Call auto_eFuse_pgm2file("CFG")
    
Exit Function

End If
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function auto_UDRP_USI_Pgm2File(Optional condstr As String = "stage") As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_USI_Pgm2File"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    'Dim patt As String
    'If (auto_eFuse_PatSetToPat_Validation(USI_pat, patt, Validating_) = True) Then Exit Function
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
    ''''------------------------------------------------------------

    TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName

    ''''20171016 update
    ''''--------------------------------
    Dim m_CP1_Early_Flag As Boolean
    condstr = LCase(condstr)
    

    
    If (condstr = "cp1_early") Then
        m_CP1_Early_Flag = True
        gS_JobName = "cp1_early" ''''used to program the category with stage = "cp1_early"
    Else
        m_CP1_Early_Flag = False
    End If
 
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
    
    Dim m_catenameVbin As String
    Dim m_crc_idx As Long
    Dim m_calcCRC As New SiteLong
    
    Dim m_cmpStage As String
    Dim m_pgmRes As New SiteLong
    Dim m_vbinResult As New SiteDouble
    Dim m_pgmWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_pgmDigSrcWave As New DSPWave
    Dim m_Fusetype As eFuseBlockType
    m_Fusetype = eFuse_UDRP
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    ''20181120, add for pgm2file
    Dim m_tmpData As New DSPWave
    m_tmpData.CreateConstant 0, gDL_BitsPerBlock, DspLong
    
    gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
    Dim m_SerialType As Boolean:: m_SerialType = False
    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
        m_SerialType = True
        gDB_SerialType = True
    End If
    
        ''20181120, add for pgm2file
    'Dim m_tmpData As New DSPWave
    'm_tmpData.CreateConstant 0, gDL_BitsPerBlock, DspLong
    
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
    'Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_UDRP_Pgm_SingleBitWave, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_UDRP, gDW_UDRP_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)
    
    If (gL_UDRP_USI_PatBitOrder = "LSB") Then
    Else
            'm_pgmDigSrcWave
    '        Dim m_tmp As New DSPWave
    '        For Each Site In TheExec.sites
    '            m_tmp(Site) = m_pgmDigSrcWave(Site).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
    '            m_pgmDigSrcWave(Site) = m_tmp(Site).ConvertDataTypeTo(DspLong)
    
        'Dim i As Long
        Dim m_size As Long
        Dim m_tmpArr() As Long
        Dim m_outArr() As Long
        Dim m_tmpWave1 As New DSPWave
        Dim outWave As New DSPWave
        
        'outWave.CreateConstant 0, m_size, DspLong
        For Each site In TheExec.sites
            m_size = m_pgmDigSrcWave(site).SampleSize
        Next
        
        outWave.CreateConstant 0, m_size, DspLong
        
        For Each site In TheExec.sites
            m_tmpWave1(site) = m_pgmDigSrcWave(site).Copy.ConvertDataTypeTo(DspLong)
            m_tmpArr = m_tmpWave1(site).Data
            m_outArr = outWave(site).Data
                For i = 0 To m_size - 1
            ''outWave.Element(i) = m_tmpWave.Element(m_size - i - 1) ''''waste TT
                m_outArr(i) = m_tmpArr(m_size - i - 1) ''''save TT
            Next i
        
        outWave(site).Data = m_outArr ''''save TT
        Next
    End If
    
    For Each site In TheExec.sites
        m_tmpData(site) = outWave(site).Copy 'was eFuse_Pgm_Bit
        gL_UDRPFuse_PgmBits(site) = gL_UDRPFuse_PgmBits(site).BitwiseOr(m_tmpData(site))
    Next
    
    If (gS_JobName = "cp1_early") Then
        gS_JobName = "cp1"
    End If
    
    ''''--------------------------------------------------------------------------------------------
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    ''TheHdw.Patterns(USI_pat).Load

'    Status = GetPatListFromPatternSet(USI_pat.Value, PatUSIArray, pat_count)
'
'    For j = 0 To pat_count - 1
'        TheExec.Datalog.WriteComment ""
'        TheExec.Datalog.WriteComment "USI Pattern: " + PatUSIArray(j)
'        For Each Site In TheExec.Sites
'            TheExec.Datalog.WriteComment "Site(" + CStr(Site) + "), Shift In Code [" + CStr(UBound(PgmBitArr)) + ":0]=" + gS_USI_BitStr(Site)
'        Next Site
'
'        Call eFuse_DSSC_SetupDigSrcWave(PatUSIArray(j), InPin, "USI_Src", outWave)
'        'UDR_SetupDigSrcArray PatUSIArray(j), InPin, "USI_Src", usiarrSize, USI_Array
'        Call TheHdw.Patterns(PatUSIArray(j)).test(pfAlways, 0)
'    Next j

'    TheHdw.Wait 100# * us
'    DebugPrintFunc USI_pat.Value

    
Exit Function
End If




'
'
'
'
'
'
'
'
'
'    ''20181120, add for pgm2file
'    Dim m_tmpData As New DSPWave
'    m_tmpData.CreateConstant 0, gL_UDRP_USI_DigSrcBits_Num, DspLong
'
'    For Each Site In TheExec.Sites
'
'        If (TheExec.TesterMode = testModeOffline) Then ''''20160526 update
'            Call eFuseENGFakeValue_Sim
'        End If
'
'        '''' initialize
'        gS_UDRP_USI_BitStr(Site) = ""
'        For i = 0 To UBound(PgmBitArr)
'            PgmBitArr(i) = 0
'        Next i
'
'        '''' 1st Step: get the PgmBitArr() per Site
'        For i = 0 To UBound(UDRP_Fuse.Category)
'            tmpdlgStr = ""
'            m_catename = UDRP_Fuse.Category(i).Name
'            m_algorithm = LCase(UDRP_Fuse.Category(i).Algorithm)
'            m_LSBbit = UDRP_Fuse.Category(i).LSBbit
'            m_MSBBit = UDRP_Fuse.Category(i).MSBbit
'            m_bitwidth = UDRP_Fuse.Category(i).Bitwidth
'            m_lolmt = UDRP_Fuse.Category(i).LoLMT
'            m_hilmt = UDRP_Fuse.Category(i).HiLMT
'            m_defval = UDRP_Fuse.Category(i).DefaultValue
'            m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
'            m_stage = LCase(UDRP_Fuse.Category(i).Stage)
'            m_resolution = UDRP_Fuse.Category(i).Resoultion
'
'            ''''20150710 new datalog format
'            tmpdlgStr = "Site(" + CStr(Site) + ") Programming : " + FormatNumeric(m_catename, gI_UDRP_catename_maxLen)
'            tmpdlgStr = tmpdlgStr + " [" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "] = "
'
'            If (m_algorithm = "base") Then
'                If (gS_JobName = m_stage) Then
'                    If (m_defreal = "decimal") Then ''''20160624 update
'                        m_decimal = m_defval
'                    Else
'                        ''''put in auto_UDRP_Constant_Initialize()
'                        ''''gD_UDRP_VBaseFuse = gD_UDRP_BaseVoltage / gD_UDRP_BaseStepVoltage - 1
'                        tmpVal = gD_UDRP_BaseVoltage
'                        m_decimal = gD_UDRP_VBaseFuse  ''''21=(550/25)-1, code=001010
'                    End If
'                Else
'                    tmpVal = 0
'                    m_decimal = 0
'                End If
'
'            ElseIf (m_algorithm = "fuse") Then
'                ''''get UDRP Fuse version and its binary code
'                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRP", m_catename, m_defreal, m_defval)
'                If (m_decimal < m_lolmt Or m_decimal > m_hilmt) Then
'                    CheckEfuseVer(Site) = m_decimal ''Chk Fail
'                Else
'                    CheckEfuseVer(Site) = 0 ''Chk Pass
'                End If
'
'            ElseIf (m_algorithm = "vddbin") Then
'                ''''<Notice>
'                ''''Here m_catename MUST be same as the content of Enum EcidVddBinningFlow
'                ''''Ex: VDD_SRAM_P1 in (Enum EcidVddBinningFlow)
'                ''''Ex: m_decimal = VBIN_RESULT(VddBinStr2Enum("VDD_CPU_P1")).GRADEVDD(Site)
'
'                If (gS_JobName = m_stage) Then
'                    If (m_defreal = "bincut") Then
'                        tmpVbin = VBIN_RESULT(VddBinStr2Enum(m_catename)).GRADEVDD(Site)
'
'                        ''''20160329 add for the offline simulation, 20160714 update
'                        If ((tmpVbin = 0 Or tmpVbin = -1) And TheExec.TesterMode = testModeOffline) Then
'                            tmpVbin = gD_UDRP_BaseVoltage + m_resolution * auto_eFuse_GetWriteDecimal("UDRP", m_catename, False)
'                        End If
'                    Else
'                        tmpVbin = m_defval
'                    End If
'                Else
'                    ''''Set tmpVbin to 0, cause stage of category is not match current job
'                    tmpVbin = 0
'                End If
'
'                If (m_defreal = "decimal") Then ''''20160624 update
'                    m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRP", m_catename, m_defreal, m_defval)
'                Else
'                    m_decimal = auto_Vbin_to_VfuseStr_New(tmpVbin, m_bitwidth, tmpVfuse, m_resolution)
'                End If
'
'            ElseIf (m_algorithm = "app") Then
'                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRP", m_catename, m_defreal, m_defval)
'
'            Else ''other cases
'                ''''20150720 update
'                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRP", m_catename, m_defreal, m_defval)
'
'            End If
'
'            ''''-------------------------------------------------------------------------------------------------------
'            ''''20150825 update
'            Call auto_eFuse_Dec2PgmArr_Write_byStage("UDRP", m_stage, i, m_decimal, m_LSBbit, m_MSBBit, PgmBitArr)
'            m_decimal = UDRP_Fuse.Category(i).Write.Decimal(Site)
'            m_bitStrM = UDRP_Fuse.Category(i).Write.BitstrM(Site)
'            TmpStr = " [" + m_bitStrM + "]"
'            If (m_algorithm = "vddbin") Then
'                If (m_defreal = "decimal") Then ''''20160624 update
'                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + TmpStr
'                Else
'                    ''''<Notice> Here using .Value to store VDDBIN value
'                    UDRP_Fuse.Category(i).Write.Value(Site) = tmpVbin
'                    UDRP_Fuse.Category(i).Write.ValStr(Site) = CStr(tmpVbin)
'                    tmpdlgStr = tmpdlgStr + FormatNumeric(CStr(tmpVbin) + "mV", 10) + TmpStr + " = " + FormatNumeric(m_decimal, -5)
'                End If
'            ElseIf (m_algorithm = "base") Then ''''20160624 update
'                If (m_defreal = "decimal") Then ''''20160624 update
'                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + TmpStr
'                Else
'                    ''''<Notice> Here using .Value to store VDDBIN value
'                    UDRP_Fuse.Category(i).Write.Value(Site) = tmpVal
'                    UDRP_Fuse.Category(i).Write.ValStr(Site) = CStr(tmpVal)
'                    tmpdlgStr = tmpdlgStr + FormatNumeric(CStr(tmpVal) + "mV", 10) + TmpStr + " = " + FormatNumeric(m_decimal, -5)
'                End If
'            Else
'                tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + TmpStr
'            End If
'            ''''-------------------------------------------------------------------------------------------------------
'            TheExec.Datalog.WriteComment tmpdlgStr
'        Next i
'
'        ''20181120, add for pgm2file
'        m_tmpData.Data = PgmBitArr
'        gL_UDRPFuse_PgmBits = gL_UDRPFuse_PgmBits.BitwiseOr(m_tmpData)
'
'        ''''20150717 update
'        '''' 2nd Step: composite to the USI_Array() for the DSSC Source
'        k = 0
'        tmpdlgStr = ""
'        If (UCase(gL_UDRP_USI_PatBitOrder) = "MSB") Then
'            ''''case: gL_UDRP_USI_PatBitOrder is MSB
'            ''''<Notice> USI_Array(0) is MSB, so it should be PgmBitArr(lastbit)
'            For i = 0 To UBound(PgmBitArr)
'                tmpdlgStr = CStr(PgmBitArr(i)) + tmpdlgStr ''''[MSB...LSB]
'                For j = 1 To gC_UDRP_USI_DSSCRepeatCyclePerBit        ''''here j start from 1
'                    USI_Array(Site, k) = PgmBitArr(UBound(PgmBitArr) - i)
'                    k = k + 1
'                Next j
'            Next i
'        Else
'            ''''case: gL_UDRP_USI_PatBitOrder is LSB
'            ''''<Notice> USI_Array(0) is LSB, so it should be PgmBitArr(0)
'            For i = 0 To UBound(PgmBitArr)
'                tmpdlgStr = CStr(PgmBitArr(i)) + tmpdlgStr ''''[MSB...LSB]
'                For j = 1 To gC_UDRP_USI_DSSCRepeatCyclePerBit  ''''here j start from 1
'                    USI_Array(Site, k) = PgmBitArr(i)
'                    k = k + 1
'                Next j
'            Next i
'        End If
'        ''''<NOTICE> Here gS_UDRP_USI_BitStr is Always [MSB(lastbit)...LSB(bit0)]
'        gS_UDRP_USI_BitStr(Site) = tmpdlgStr ''''[MSB(lastbit)...LSB(bit0)]
'        TheExec.Datalog.WriteComment ""
'    Next Site
'
'    Call UpdateDLogColumns(gI_UDRP_catename_maxLen)
'
'    ''''20171016 update
'    If (m_CP1_Early_Flag) Then
'        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="UDRP_PGM_Stage_" + UCase(condstr)
'        gS_JobName = "cp1" ''''<MUST> Reset back to cp1
'    Else
'        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="UDRP_PGM_Stage_" + UCase(gS_JobName)
'    End If
'    Call UpdateDLogColumns__False
'
'    If (TheExec.Sites.ActiveCount = 0) Then Exit Function 'chihome
'
'    ''''--------------------------------------------------------------------------------------------
''    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'    ''TheHdw.Patterns(USI_pat).Load
'
''    Status = GetPatListFromPatternSet(USI_pat.Value, PatUSIArray, pat_count)
'
''    For j = 0 To pat_count - 1
''        TheExec.Datalog.WriteComment ""
''        TheExec.Datalog.WriteComment "UDRP USI Pattern: " + PatUSIArray(j)
''        For Each Site In TheExec.sites
''            TheExec.Datalog.WriteComment "Site(" + CStr(Site) + "), Shift In Code [" + CStr(UBound(PgmBitArr)) + ":0]=" + gS_UDRP_USI_BitStr(Site)
''        Next Site
''
''        UDR_SetupDigSrcArray PatUSIArray(j), InPin, "UDRP_USI_Src", usiarrSize, USI_Array
''        Call TheHdw.Patterns(PatUSIArray(j)).test(pfAlways, 0)
''    Next j
'
'    DebugPrintFunc USI_pat.Value
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDRP_USO_Syntax_Chk_Pgm2File(Optional condstr As String = "all") As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRP_USO_Syntax_Chk_Pgm2File"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
'    Dim ReadPatt As String
'    If (auto_eFuse_PatSetToPat_Validation(USO_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_USO As New DSPWave
    Dim USO_BitStr As New SiteVariant
    Dim i As Long, j As Long, k As Long
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
    Dim FuseStr As String:: FuseStr = "UDRP"

    m_siteVar = "UDR_PChk_Var"
    
    ''TheHdw.Patterns(USO_pat).Load
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    'Status = GetPatListFromPatternSet(USO_pat.Value, PatUSOArray, pat_count)
    TheExec.Datalog.WriteComment ""
    
    ''''20171016 update
    ''''--------------------------------
    Dim m_testFlag As Boolean
    Dim m_CP1_Early_Flag As Boolean
    condstr = LCase(condstr)
    
    If (condstr = "cp1_early") Then
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
    
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    Dim m_bitFlag_mode As Long
    Dim CapWave As New DSPWave

    m_Fusetype = eFuse_UDRP
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = False
    blank_stage = True
    
    'Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(0), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
    'Call TheHdw.Patterns(PatUSOArray(0)).test(pfAlways, 0)
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
    Dim m_SerialType As Boolean:: m_SerialType = False
    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
        m_SerialType = True
        gDB_SerialType = True
    End If
        Dim m_PatBitOrder As String
    
    If (gL_UDRP_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
'    If (TheExec.TesterMode = testModeOffline) Then
'        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
'        ''''Here MarginRead it only read the programmed Site.
'        Dim Temp_USO As New DSPWave
'        Temp_USO.CreateConstant 0, USO_CapBits, DspLong
'
'        gL_eFuse_Sim_Blank = 0
'
'        Call RunDSP.eFuse_SingleBitWave2CapWave32Bits(eFuse_UDRP, Temp_USO)
'        'Call auto_eFuse_print_capWave32Bits(eFuse_UDR, Temp_USO, False) ''''True to print out
'        If (m_PatBitOrder = "bit0_bitLast") Then
'            For Each Site In TheExec.Sites
'                Trim_code_USO(Site) = Temp_USO(Site).Copy
'            Next
'        Else
'            Call ReverseWave(Temp_USO, Trim_code_USO, m_PatBitOrder, USO_CapBits)
'        End If
'    End If
    
    
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
    Call UpdateDLogColumns(gI_UDRP_catename_maxLen + 18)

    
    If (TheExec.EnableWord("Pgm2File")) Then
        'CapWave = gL_UDRPFuse_PgmBits.Copy
        For Each site In TheExec.sites
            CapWave(site) = gL_UDRPFuse_PgmBits(site).Copy
        Next site
        
        Call rundsp.eFuse_Wave1bit_to_SingleDoubleBitWave(m_Fusetype, m_bitFlag_mode, m_SerialType, CapWave, m_FBC, blank_stage, allBlank)

        'Call RunDSP.eFuse_Wave32bits_to_SingleDoubleBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, blank_stage, allblank)
    End If
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    'If (TheExec.Sites(Site).SiteVariableValue(m_siteVar) = 1) Then
        'Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, m_bitFlag_mode, Trim_code_USO, m_FBC, m_cmpResult, , , True, m_PatBitOrder)
    'End If
    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, m_bitFlag_mode, Trim_code_USO, m_FBC, blank_stage, allblank, True)

    'Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, 1, Trim_code_USO, m_FBC, m_cmpResult)


    Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

    Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
    Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate)
    
    Call auto_eFuse_pgm2file(FuseStr, gL_UDRPFuse_PgmBits)
    
    'Call auto_eFuse_SyntaxCheck(m_Fusetype, condstr)

        ''''201811XX reset
    If (gS_JobName = "cp1_early") Then gS_JobName = "cp1" ''''<MUST>
Exit Function

End If
    
    
    
    
    
    
    
    

    ''20181120, add for pgm2file
'    If (condstr = "all" Or condstr = "") Then
'        Call auto_eFuse_pgm2file("UDRP", gL_UDRPFuse_PgmBits)
'    End If

''    For j = 0 To pat_count - 1
''        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
''        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
''
''        ''''20160506 update
''        ''''due to the additional characters "_USI_USO_compare", so plus 18.
''        Call UpdateDLogColumns(gI_UDRP_catename_maxLen + 18)
''
''        For Each Site In TheExec.sites
''            ''''initialize
''            USO_BitStr(Site) = ""  ''''MUST, and it's [bitLast ... bit0]
''            For i = 0 To UBound(DigCapArray)
''               DigCapArray(i) = 0  ''''MUST be [bit0 ... bitLast]
''            Next i
''
''            ''--------------------------------------------------------------------------------------
''            ''''20150717 update
''            ''''composite to the USO_BitStr() from the DSSC Capture
''            If (UCase(gL_UDRP_USO_PatBitOrder) = "MSB") Then
''                ''''<Notice>
''                ''''the pattern DSSC capture from MSB to LSB (bitLast to bit0)
''                ''''so Trim_code_USO.Element(0) is MSB
''                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
''                For i = 0 To USO_CapBits - 1
''                    DigCapArray(i) = Trim_code_USO.Element(USO_CapBits - 1 - i)    ''''Reverse Bit String
''                    USO_BitStr(Site) = USO_BitStr(Site) & Trim_code_USO.Element(i) ''''[MSB ... LSB]
''                Next i
''            Else
''                ''''<Notice>
''                ''''the pattern DSSC capture from LSB to MSB (bit0 to bitLast)
''                ''''so Trim_code_USO.Element(0) is LSB
''                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
''                For i = 0 To USO_CapBits - 1
''                    DigCapArray(i) = Trim_code_USO.Element(i)
''                    USO_BitStr(Site) = Trim_code_USO.Element(i) & USO_BitStr(Site) ''''[MSB ... LSB]
''                Next i
''            End If
''            ''--------------------------------------------------------------------------------------
''
''            ''''20160324 updae
''            ''''=============== Start of Simulated Data ===============
''            If (TheExec.TesterMode = testModeOffline) Then
''                ''''20160906 trial for the ugly codes
''                ''''<Issued codes> Shift out code [383:0]=111110000110100000000000000000001111100001101000000000000000000011111110111010000000000000000000111110100110100000000000000000001111101001101000000000000000000011111000011010000000000000000000111110000100000000000000000000001111100001000000000000000000000011111000010000000001000000010000111110000110000000000000000000001111100001000000000000000000000011111000010000000000000000000000
''                ''gS_UDRP_USI_BitStr(Site) = "111110000110100000000000000000001111100001101000000000000000000011111110111010000000000000000000111110100110100000000000000000001111101001101000000000000000000011111000011010000000000000000000111110000100000000000000000000001111100001000000000000000000000011111000010000000001000000010000111110000110000000000000000000001111100001000000000000000000000011111000010000000000000000000000"
''
''                USO_BitStr(Site) = gS_UDRP_USI_BitStr(Site)
''                For i = 0 To USO_CapBits - 1
''                    DigCapArray(i) = CLng(Mid(StrReverse(gS_UDRP_USI_BitStr(Site)), i + 1, 1))
''                Next i
''
''                If (gS_JobName <> "cp1" Or TheExec.sites.Item(Site).FlagState("F_UDRP_Early_Enable") = logicTrue) Then ''''was "cp1"
''                    TmpStr = ""
''                    For i = 0 To USO_CapBits - 1
''                        If (DigCapArray(i) = 0) Then
''                            DigCapArray(i) = gL_Sim_FuseBits(Site, i)
''                        Else
''                            gL_Sim_FuseBits(Site, i) = DigCapArray(i)
''                        End If
''                        TmpStr = CStr(DigCapArray(i)) + TmpStr
''                    Next i
''                    USO_BitStr(Site) = TmpStr ''''<MUST>
''                End If
''            End If
''            ''''===============   End of Simulated Data ===============
''
''            If (True) Then
''                '=======================================================
''                '= Print out the caputured bit data from DigCap        =
''                '=======================================================
''                TheExec.Datalog.WriteComment ""
''                Call auto_PrintAllBitbyDSSC(DigCapArray, USO_PrintRow, UBound(DigCapArray) + 1, USO_BitPerRow)
''            End If
''            ''--------------------------------------------------------------------------------------
''
''            TheExec.Datalog.WriteComment "Site(" & Site & "), UDRP USO pat:" & PatUSOArray(j) & ", Shift out code [" + CStr(USO_CapBits - 1) + ":0]=" + USO_BitStr(Site)
''            ''TheExec.Datalog.WriteComment ""
''
''            ''''----------------------------------------------------------------------------------------------
''            ''''20161114 update for print all bits (DTR) in STDF
''            ''''20171016 update to excluding "cp1_early"
''            If (m_CP1_Early_Flag = False) Then Call auto_eFuse_to_STDF_allBits("UDRP", USO_BitStr(Site))
''            ''''----------------------------------------------------------------------------------------------
''
''            ''''20150717 New
''            Call auto_Decode_UDRP_Binary_Data(DigCapArray)
''
''            ''''----------------------------------------------------------------------------------
''            ''''judge pass/fail for the specific test limit
''            tmpStrL = StrReverse(USO_BitStr(Site)) ''''translate to [LSB......MSB]
''            For i = 0 To UBound(UDRP_Fuse.Category)
''                tmpdlgStr = ""
''                m_stage = LCase(UDRP_Fuse.Category(i).Stage)
''                m_CateName = UDRP_Fuse.Category(i).Name
''                m_Algorithm = LCase(UDRP_Fuse.Category(i).Algorithm)
''                m_LSBbit = UDRP_Fuse.Category(i).LSBbit
''                m_MSBbit = UDRP_Fuse.Category(i).MSBbit
''                m_BitWidth = UDRP_Fuse.Category(i).BitWidth
''                m_lolmt = UDRP_Fuse.Category(i).LoLMT
''                m_hilmt = UDRP_Fuse.Category(i).HiLMT
''                m_defreal = LCase(UDRP_Fuse.Category(i).Default_Real)
''                m_resolution = UDRP_Fuse.Category(i).Resoultion
''
''                m_decimal = UDRP_Fuse.Category(i).Read.Decimal(Site)
''                m_value = UDRP_Fuse.Category(i).Read.Value(Site)
''                m_bitsum = UDRP_Fuse.Category(i).Read.BitSummation(Site)
''                m_hexStr = UDRP_Fuse.Category(i).Read.HexStr(Site)
''                m_unitType = unitNone
''                m_scale = scaleNone ''''default
''                m_tsName = Replace(m_CateName, " ", "_") ''''20151028, benefit for the script
''
''                m_bitStrM = StrReverse(Mid(tmpStrL, m_LSBbit + 1, m_BitWidth))
''
''                m_testFlag = True ''''20171016 update
''                If (m_CP1_Early_Flag = True) Then ''''20171016 update
''                    ''''only compare these category with stage="cp1_early"
''                    If (m_stage = condStr) Then
''                        ''''other cases
''                        m_testValue = m_decimal
''                    Else
''                        ''''Here it's an excluding case
''                        ''''<MUST>
''                        m_testFlag = False
''                        m_testValue = 0
''                        m_lolmt = 0
''                        m_hilmt = 0
''                    End If
''
''                ElseIf (m_Algorithm = "base") Then
''                    If (m_defreal = "decimal") Then ''''20160624 update
''                        m_testValue = m_decimal
''                    Else
''                        m_testValue = (m_decimal + 1) * gD_BaseStepVoltage
''                        m_unitType = unitVolt
''                        m_scale = scaleMilli
''                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
''                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
''                        m_testValue = m_value * 0.001 ''''to unit:V
''                    End If
''
''                ElseIf (m_Algorithm = "vddbin") Then
''                    ''''<Notice> User Maintain
''                    ''''Ex:: step_vdd_cpu_p1 = VDD_BIN(vdd_cpu_p1).MODE_STEP
''                    m_testValue = 0 ''''default to fail
''                    If (m_defreal = "decimal") Then ''''20160624 update
''                        m_testValue = m_decimal
''                    ElseIf (m_defreal Like "safe*voltage" Or m_defreal = "default") Then
''                        m_unitType = unitVolt
''                        m_scale = scaleMilli
''                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
''                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
''                        m_testValue = m_value * 0.001 ''''to unit:V
''                    ElseIf (m_defreal = "bincut") Then
''                        m_catenameVbin = m_CateName '150127
''                        ''''' Every testing stage needs to check this formula with bincut fused value (2017/5/17, Jack)
''                        vbinflag = auto_CheckVddBinInRangeNew(m_catenameVbin, m_resolution)
''
''                        ''''20160329 Add for the offline simulation
''                        If (vbinflag = 0 And TheExec.TesterMode = testModeOffline) Then
''                            vbinflag = 1
''                        End If
''
''                        m_Pmode = VddBinStr2Enum(m_catenameVbin)
''                        tmpVal = auto_VfuseStr_to_Vdd_New(m_bitStrM, m_BitWidth, m_resolution)
''                        MaxLevelIndex = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step '150127
''                        m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(MaxLevelIndex) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(MaxLevelIndex)
''                        m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
''                        ''''judge the result
''                        If (vbinflag = 1) Then
''                            m_value = tmpVal
''                        Else
''                            m_value = -999
''                            TmpStr = m_CateName + "(Site " + CStr(Site) + ") = " + CStr(tmpVal) + " is not in range" '150127
''                            TheExec.Datalog.WriteComment TmpStr
''                        End If
''                        m_unitType = unitVolt
''                        m_scale = scaleMilli
''                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
''                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
''                        m_testValue = m_value * 0.001 ''''to unit:V
''                    End If
''                Else
''                    ''''other cases, 20160927 update
''                    m_testValue = m_decimal
''                End If
''
''                ''''20160108 New
''                m_tsName = Replace(m_CateName, " ", "_") ''''20151028, benefit for the script
''                Call auto_eFuse_chkLoLimit("UDRP", i, m_stage, m_lolmt)
''                Call auto_eFuse_chkHiLimit("UDRP", i, m_stage, m_hilmt)
''
''                ''''20170811 update
''                If (m_BitWidth >= 32) Then
''                    ''m_tsName = m_tsName + "_" + m_hexStr
''                    ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
''                    m_lolmt = auto_Value2HexStr(m_lolmt, m_BitWidth)
''                    m_hilmt = auto_Value2HexStr(m_hilmt, m_BitWidth)
''
''                    ''''------------------------------------------
''                    ''''compare with lolmt, hilmt
''                    ''''m_testValue 0 means fail
''                    ''''m_testValue 1 means pass
''                    ''''------------------------------------------
''                    m_testValue = auto_TestStringLimit(m_hexStr, CStr(m_lolmt), CStr(m_hilmt))
''                    m_lolmt = 1
''                    m_hilmt = 1
''                Else
''                    ''''20160620 update
''                    ''''20160927 update the new logical methodology for the unexpected binary decode.
''                    If (auto_isHexString(CStr(m_lolmt)) = True) Then
''                        ''''translate to double value
''                        m_lolmt = auto_HexStr2Value(m_lolmt)
''                    Else
''                        ''''doNothing, m_lolmt = m_lolmt
''                    End If
''
''                    If (auto_isHexString(CStr(m_hilmt)) = True) Then
''                        ''''translate to double value
''                        m_hilmt = auto_HexStr2Value(m_hilmt)
''                    Else
''                        ''''doNothing, m_hilmt = m_hilmt
''                    End If
''                End If
''
''                ''''20171016 update
''                If (m_testFlag) Then
''                    TheExec.Flow.TestLimit resultVal:=m_testValue, lowval:=m_lolmt, hival:=m_hilmt, Tname:=m_tsName, unit:=m_unitType, ScaleType:=m_scale
''                End If
''            Next i
''            ''''----------------------------------------------------------------------------------
''
''            ''''20160503 update for the 2nd way to prevent for the temp sensor trim will all zero values
''            ''''20160907 update
''            Dim m_valueSum As Long
''            Dim m_matchTMPS_flag As Boolean
''            m_valueSum = 0 ''''initialize
''            m_matchTMPS_flag = False
''            m_stage = "" ''''<MUST> 20160617 update, if the "trim/tmps" is existed then m_stage has its correct value.
''            For i = 0 To UBound(UDRP_Fuse.Category)
''                m_CateName = UCase(UDRP_Fuse.Category(i).Name)
''                m_Algorithm = LCase(UDRP_Fuse.Category(i).Algorithm)
''                If (m_CateName Like "TEMP_SENSOR*" Or m_Algorithm = "tmps") Then ''''was m_algorithm = "trim", 20171103 update
''                    m_stage = LCase(UDRP_Fuse.Category(i).Stage)
''                    m_decimal = UDRP_Fuse.Category(i).Read.Decimal(Site)
''                    m_valueSum = m_valueSum + m_decimal
''                    m_matchTMPS_flag = True
''                End If
''            Next i
''            If (m_matchTMPS_flag = True) Then
''                ''''if Job >= m_stage then m_valueSim >= 1
''                If (checkJob_less_Stage_Sequence(m_stage) = False) Then
''                    TheExec.Flow.TestLimit resultVal:=m_valueSum, lowval:=1, Tname:="UDRP_TMPS_SUM"
''                    ''TheExec.Datalog.WriteComment ""
''                Else
''                    ''''if Job < m_stage then m_valueSim = 0
''                    ''''20180105 update
''                    If (gB_eFuse_Disable_ChkLMT_Flag = True) Then ''''case CP2 back to CP1 retest
''                        TheExec.Flow.TestLimit resultVal:=m_valueSum, lowval:=0, Tname:="UDRP_TMPS_SUM"
''                    Else
''                        TheExec.Flow.TestLimit resultVal:=m_valueSum, lowval:=0, hival:=0, Tname:="UDRP_TMPS_SUM"
''                    End If
''                End If
''            End If
''            ''''--------------------------------------------------------------------------------------------
''
''            ''''20160503 update
''            ''''compare both USI and USO for the specific stage, it's only when siteVar is '1'.
''            ''''Must be after the decode then you have the Read buffer value
''            ''''20180105 update
''            ''''The below is used to compare both USI and USO contents.
''            If (TheExec.sites(Site).SiteVariableValue(m_siteVar) = 1) Then
''                Dim m_writeBitStrM As String
''                Dim m_readBitStrM As String
''                Dim m_usiusoCmp As Long
''                m_usiusoCmp = 0 ''''<MUST> default Compare Pass:0, Fail:1
''                For i = 0 To UBound(UDRP_Fuse.Category)
''                    m_stage = LCase(UDRP_Fuse.Category(i).Stage)
''                    If (m_stage = gS_JobName) Then
''                        m_CateName = UDRP_Fuse.Category(i).Name
''                        m_writeBitStrM = UCase(UDRP_Fuse.Category(i).Write.BitStrM(Site))
''                        m_readBitStrM = UCase(UDRP_Fuse.Category(i).Read.BitStrM(Site))
''                        m_tsName = Replace(m_CateName, " ", "_")
''                        m_tsName = m_tsName + "_USI_USO_compare"
''                        If (m_writeBitStrM <> m_readBitStrM) Then
''                            ''''20180105 update
''                            m_usiusoCmp = 1 ''''Fail Comparison
''                            TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(Site) + "), " + m_tsName + " Failure."
''                        Else
''                            ''''case: m_writeBitStrM = m_readBitStrM
''                            ''''reserve
''                        End If
''                    End If
''                Next i
''                TheExec.Flow.TestLimit resultVal:=m_usiusoCmp, lowval:=0, hival:=0, Tname:="UDRP_USO_USI_Cmp"
''            End If
''            ''''--------------------------------------------------------------------------------------------
''
''            TheExec.Datalog.WriteComment ""
''            If (TheExec.sites.ActiveCount = 0) Then Exit Function 'chihome
''        Next Site
''        Call UpdateDLogColumns__False
''    Next j

     ''''20171016 update
    If (m_CP1_Early_Flag = True) Then
        gS_JobName = "cp1" ''''reset
    End If
    DebugPrintFunc USO_pat.Value

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDRE_USI_Pgm2File(Optional condstr As String = "stage") As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRE_USI_Pgm2File"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    'Dim patt As String
    'If (auto_eFuse_PatSetToPat_Validation(USI_pat, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_USI As New DSPWave
    Dim i As Long, j As Long, k As Long
    Dim pat_count As Long
    Dim Status As Boolean
    Dim PatUSIArray() As String
    Dim CheckEfuseVer As New SiteVariant
    
    Dim usiarrSize As Long
    usiarrSize = gL_UDRE_USI_DigSrcBits_Num * gC_UDRE_USI_DSSCRepeatCyclePerBit

    Dim PgmBitArr() As Long
    ReDim PgmBitArr(gL_UDRE_USI_DigSrcBits_Num - 1)

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
    ''''------------------------------------------------------------

    TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName

    ''''20171016 update
    ''''--------------------------------
    Dim m_CP1_Early_Flag As Boolean
    condstr = LCase(condstr)
    
    ''20181120, add for pgm2file
'    Dim m_tmpData As New DSPWave
'    m_tmpData.CreateConstant 0, EConfigTotalBitCount, DspLong
'
    If (condstr = "cp1_early") Then
        m_CP1_Early_Flag = True
        gS_JobName = "cp1_early" ''''used to program the category with stage = "cp1_early"
    Else
        m_CP1_Early_Flag = False
    End If
 
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
                Call auto_UDRE_USI_Sim(False, False) ''''True for print debug
                Call eFuseENGFakeValue_Sim
            Next site
        End If
    End If
    
    Dim m_catenameVbin As String
    Dim m_crc_idx As Long
    Dim m_calcCRC As New SiteLong
    
    Dim m_cmpStage As String
    Dim m_pgmRes As New SiteLong
    Dim m_vbinResult As New SiteDouble
    Dim m_pgmWave As New DSPWave
    Dim m_tmpWave As New DSPWave
    Dim m_pgmDigSrcWave As New DSPWave
    Dim m_Fusetype As eFuseBlockType
    m_Fusetype = eFuse_UDRE
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
    Dim m_SerialType As Boolean:: m_SerialType = False
    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
        m_SerialType = True
        gDB_SerialType = True
    End If
    
    ''20181120, add for pgm2file
    Dim m_tmpData As New DSPWave
    m_tmpData.CreateConstant 0, gDL_BitsPerBlock, DspLong
    
    m_pgmWave.CreateConstant 0, gDL_BitsPerBlock, DspLong ''''gDL_BitsPerBlock = EConfigBitPerBlockUsed
    ''''<MUST and Importance>
    For Each site In TheExec.sites
        gDW_Pgm_RawBitWave = m_pgmWave.Copy
    Next site
    
        ''''Only composite case "real or bincut" PgmBits Wave per Stage requirement
    For i = 0 To UBound(UDRE_Fuse.Category)
        With UDRE_Fuse.Category(i)
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
                    Call auto_eFuse_Vddbin_bincut_setWriteData(eFuse_UDRE, i)
                End If
                ''''---------------------------------------------------------------------------
                With UDRE_Fuse.Category(i)
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
        Call auto_eFuse_Gen_PgmBitSrcWave_onlyDeaultBits(eFuse_UDRE, m_pgmDigSrcWave, m_pgmRes)
    Else
        ''''condStr = "stage"
        ''''Here Parameter bitFlag_mode=1 (Stage Bits)
        m_calcCRC = 0
        Call rundsp.eFuse_Gen_PgmBitSrcWave(eFuse_UDRE, 1, m_pgmDigSrcWave, m_pgmRes)
    End If
    TheHdw.Wait 1# * ms
    
    Call UpdateDLogColumns(gI_UDRE_catename_maxLen)
    TheExec.Flow.TestLimit resultVal:=m_pgmRes, lowVal:=1, hiVal:=1, Tname:="UDRE_PGM_" + UCase(gS_JobName)
    Call UpdateDLogColumns__False

    ''''use global DSP variable
    Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_UDRE_Pgm_SingleBitWave, gDL_TotalBits, gDL_ReadCycles, gB_eFuse_printBitMap)
    'Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_UDRE_Pgm_SingleBitWave, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_UDRE, gDW_UDRE_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)
    
    If (gL_UDRE_USI_PatBitOrder = "LSB") Then
    Else
            'm_pgmDigSrcWave
    '        Dim m_tmp As New DSPWave
    '        For Each Site In TheExec.sites
    '            m_tmp(Site) = m_pgmDigSrcWave(Site).ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspSerial, 32, 0, Bit0IsMsb)
    '            m_pgmDigSrcWave(Site) = m_tmp(Site).ConvertDataTypeTo(DspLong)
    
        'Dim i As Long
        Dim m_size As Long
        Dim m_tmpArr() As Long
        Dim m_outArr() As Long
        Dim m_tmpWave1 As New DSPWave
        Dim outWave As New DSPWave
        
        'outWave.CreateConstant 0, m_size, DspLong
        For Each site In TheExec.sites
            m_size = m_pgmDigSrcWave(site).SampleSize
        Next
        
        outWave.CreateConstant 0, m_size, DspLong
        
        For Each site In TheExec.sites
            m_tmpWave1(site) = m_pgmDigSrcWave(site).Copy.ConvertDataTypeTo(DspLong)
            m_tmpArr = m_tmpWave1(site).Data
            m_outArr = outWave(site).Data
                For i = 0 To m_size - 1
            ''outWave.Element(i) = m_tmpWave.Element(m_size - i - 1) ''''waste TT
                m_outArr(i) = m_tmpArr(m_size - i - 1) ''''save TT
            Next i
        
        outWave(site).Data = m_outArr ''''save TT
        Next
    End If
    
    For Each site In TheExec.sites
        m_tmpData(site) = outWave(site).Copy 'was eFuse_Pgm_Bit
        gL_UDREFuse_PgmBits(site) = gL_UDREFuse_PgmBits(site).BitwiseOr(m_tmpData(site))
    Next
    
    If (gS_JobName = "cp1_early") Then
        gS_JobName = "cp1"
    End If
    
    ''''--------------------------------------------------------------------------------------------
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    ''TheHdw.Patterns(USI_pat).Load

'    Status = GetPatListFromPatternSet(USI_pat.Value, PatUSIArray, pat_count)
'
'    For j = 0 To pat_count - 1
'        TheExec.Datalog.WriteComment ""
'        TheExec.Datalog.WriteComment "USI Pattern: " + PatUSIArray(j)
'        For Each Site In TheExec.Sites
'            TheExec.Datalog.WriteComment "Site(" + CStr(Site) + "), Shift In Code [" + CStr(UBound(PgmBitArr)) + ":0]=" + gS_USI_BitStr(Site)
'        Next Site
'
'        Call eFuse_DSSC_SetupDigSrcWave(PatUSIArray(j), InPin, "USI_Src", outWave)
'        'UDR_SetupDigSrcArray PatUSIArray(j), InPin, "USI_Src", usiarrSize, USI_Array
'        Call TheHdw.Patterns(PatUSIArray(j)).test(pfAlways, 0)
'    Next j

'    TheHdw.Wait 100# * us
'    DebugPrintFunc USI_pat.Value

    
Exit Function
End If
    
    
    
    
'
'
'
'    ''20181120, add for pgm2file
'    Dim m_tmpData As New DSPWave
'    m_tmpData.CreateConstant 0, gL_UDRE_USI_DigSrcBits_Num, DspLong
'
'    For Each Site In TheExec.Sites
'
'        If (TheExec.TesterMode = testModeOffline) Then ''''20160526 update
'            Call eFuseENGFakeValue_Sim
'        End If
'
'        '''' initialize
'        gS_UDRE_USI_BitStr(Site) = ""
'        For i = 0 To UBound(PgmBitArr)
'            PgmBitArr(i) = 0
'        Next i
'
'        '''' 1st Step: get the PgmBitArr() per Site
'        For i = 0 To UBound(UDRE_Fuse.Category)
'            tmpdlgStr = ""
'            m_catename = UDRE_Fuse.Category(i).Name
'            m_algorithm = LCase(UDRE_Fuse.Category(i).Algorithm)
'            m_LSBbit = UDRE_Fuse.Category(i).LSBbit
'            m_MSBBit = UDRE_Fuse.Category(i).MSBbit
'            m_bitwidth = UDRE_Fuse.Category(i).Bitwidth
'            m_lolmt = UDRE_Fuse.Category(i).LoLMT
'            m_hilmt = UDRE_Fuse.Category(i).HiLMT
'            m_defval = UDRE_Fuse.Category(i).DefaultValue
'            m_defreal = LCase(UDRE_Fuse.Category(i).Default_Real)
'            m_stage = LCase(UDRE_Fuse.Category(i).Stage)
'            m_resolution = UDRE_Fuse.Category(i).Resoultion
'
'            ''''20150710 new datalog format
'            tmpdlgStr = "Site(" + CStr(Site) + ") Programming : " + FormatNumeric(m_catename, gI_UDRE_catename_maxLen)
'            tmpdlgStr = tmpdlgStr + " [" + Format(m_MSBBit, "0000") + ":" + Format(m_LSBbit, "0000") + "] = "
'
'            If (m_algorithm = "base") Then
'                If (gS_JobName = m_stage) Then
'                    If (m_defreal = "decimal") Then ''''20160624 update
'                        m_decimal = m_defval
'                    Else
'                        ''''put in auto_UDRE_Constant_Initialize()
'                        ''''gD_UDRE_VBaseFuse = gD_UDRE_BaseVoltage / gD_UDRE_BaseStepVoltage - 1
'                        tmpVal = gD_UDRE_BaseVoltage
'                        m_decimal = gD_UDRE_VBaseFuse  ''''21=(550/25)-1, code=001010
'                    End If
'                Else
'                    tmpVal = 0
'                    m_decimal = 0
'                End If
'
'            ElseIf (m_algorithm = "fuse") Then
'                ''''get UDRE Fuse version and its binary code
'                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRE", m_catename, m_defreal, m_defval)
'                If (m_decimal < m_lolmt Or m_decimal > m_hilmt) Then
'                    CheckEfuseVer(Site) = m_decimal ''Chk Fail
'                Else
'                    CheckEfuseVer(Site) = 0 ''Chk Pass
'                End If
'
'            ElseIf (m_algorithm = "vddbin") Then
'                ''''<Notice>
'                ''''Here m_catename MUST be same as the content of Enum EcidVddBinningFlow
'                ''''Ex: VDD_SRAM_P1 in (Enum EcidVddBinningFlow)
'                ''''Ex: m_decimal = VBIN_RESULT(VddBinStr2Enum("VDD_CPU_P1")).GRADEVDD(Site)
'
'                If (gS_JobName = m_stage) Then
'                    If (m_defreal = "bincut") Then
'                        tmpVbin = VBIN_RESULT(VddBinStr2Enum(m_catename)).GRADEVDD(Site)
'
'                        ''''20160329 add for the offline simulation, 20160714 update
'                        If ((tmpVbin = 0 Or tmpVbin = -1) And TheExec.TesterMode = testModeOffline) Then
'                            tmpVbin = gD_UDRE_BaseVoltage + m_resolution * auto_eFuse_GetWriteDecimal("UDRE", m_catename, False)
'                        End If
'                    Else
'                        tmpVbin = m_defval
'                    End If
'                Else
'                    ''''Set tmpVbin to 0, cause stage of category is not match current job
'                    tmpVbin = 0
'                End If
'
'                If (m_defreal = "decimal") Then ''''20160624 update
'                    m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRE", m_catename, m_defreal, m_defval)
'                Else
'                    m_decimal = auto_Vbin_to_VfuseStr_New(tmpVbin, m_bitwidth, tmpVfuse, m_resolution)
'                End If
'
'            ElseIf (m_algorithm = "app") Then
'                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRE", m_catename, m_defreal, m_defval)
'
'            Else ''other cases
'                ''''20150720 update
'                m_decimal = auto_eFuse_Get_DefaultRealDecimal("UDRE", m_catename, m_defreal, m_defval)
'
'            End If
'
'            ''''-------------------------------------------------------------------------------------------------------
'            ''''20150825 update
'            Call auto_eFuse_Dec2PgmArr_Write_byStage("UDRE", m_stage, i, m_decimal, m_LSBbit, m_MSBBit, PgmBitArr)
'            m_decimal = UDRE_Fuse.Category(i).Write.Decimal(Site)
'            m_bitStrM = UDRE_Fuse.Category(i).Write.BitstrM(Site)
'            TmpStr = " [" + m_bitStrM + "]"
'            If (m_algorithm = "vddbin") Then
'                If (m_defreal = "decimal") Then ''''20160624 update
'                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + TmpStr
'                Else
'                    ''''<Notice> Here using .Value to store VDDBIN value
'                    UDRE_Fuse.Category(i).Write.Value(Site) = tmpVbin
'                    UDRE_Fuse.Category(i).Write.ValStr(Site) = CStr(tmpVbin)
'                    tmpdlgStr = tmpdlgStr + FormatNumeric(CStr(tmpVbin) + "mV", 10) + TmpStr + " = " + FormatNumeric(m_decimal, -5)
'                End If
'            ElseIf (m_algorithm = "base") Then ''''20160624 update
'                If (m_defreal = "decimal") Then ''''20160624 update
'                    tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + TmpStr
'                Else
'                    ''''<Notice> Here using .Value to store VDDBIN value
'                    UDRE_Fuse.Category(i).Write.Value(Site) = tmpVal
'                    UDRE_Fuse.Category(i).Write.ValStr(Site) = CStr(tmpVal)
'                    tmpdlgStr = tmpdlgStr + FormatNumeric(CStr(tmpVal) + "mV", 10) + TmpStr + " = " + FormatNumeric(m_decimal, -5)
'                End If
'            Else
'                tmpdlgStr = tmpdlgStr + FormatNumeric(m_decimal, 10) + TmpStr
'            End If
'            ''''-------------------------------------------------------------------------------------------------------
'            TheExec.Datalog.WriteComment tmpdlgStr
'        Next i
'
'        ''20181120, add for pgm2file
'        m_tmpData.Data = PgmBitArr
'        gL_UDREFuse_PgmBits = gL_UDREFuse_PgmBits.BitwiseOr(m_tmpData)
'
'        ''''20150717 update
'        '''' 2nd Step: composite to the USI_Array() for the DSSC Source
'        k = 0
'        tmpdlgStr = ""
'        If (UCase(gL_UDRE_USI_PatBitOrder) = "MSB") Then
'            ''''case: gL_UDRE_USI_PatBitOrder is MSB
'            ''''<Notice> USI_Array(0) is MSB, so it should be PgmBitArr(lastbit)
'            For i = 0 To UBound(PgmBitArr)
'                tmpdlgStr = CStr(PgmBitArr(i)) + tmpdlgStr ''''[MSB...LSB]
'                For j = 1 To gC_UDRE_USI_DSSCRepeatCyclePerBit        ''''here j start from 1
'                    USI_Array(Site, k) = PgmBitArr(UBound(PgmBitArr) - i)
'                    k = k + 1
'                Next j
'            Next i
'        Else
'            ''''case: gL_UDRE_USI_PatBitOrder is LSB
'            ''''<Notice> USI_Array(0) is LSB, so it should be PgmBitArr(0)
'            For i = 0 To UBound(PgmBitArr)
'                tmpdlgStr = CStr(PgmBitArr(i)) + tmpdlgStr ''''[MSB...LSB]
'                For j = 1 To gC_UDRE_USI_DSSCRepeatCyclePerBit  ''''here j start from 1
'                    USI_Array(Site, k) = PgmBitArr(i)
'                    k = k + 1
'                Next j
'            Next i
'        End If
'        ''''<NOTICE> Here gS_UDRE_USI_BitStr is Always [MSB(lastbit)...LSB(bit0)]
'        gS_UDRE_USI_BitStr(Site) = tmpdlgStr ''''[MSB(lastbit)...LSB(bit0)]
'        TheExec.Datalog.WriteComment ""
'    Next Site
'
'    Call UpdateDLogColumns(gI_UDRE_catename_maxLen)
'
'    ''''20171016 update
'    If (m_CP1_Early_Flag) Then
'        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="UDRE_PGM_Stage_" + UCase(condstr)
'        gS_JobName = "cp1" ''''<MUST> Reset back to cp1
'    Else
'        TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="UDRE_PGM_Stage_" + UCase(gS_JobName)
'    End If
'    Call UpdateDLogColumns__False
'
'    If (TheExec.Sites.ActiveCount = 0) Then Exit Function 'chihome
'
'    ''''--------------------------------------------------------------------------------------------
''    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
'    ''TheHdw.Patterns(USI_pat).Load
'
''    Status = GetPatListFromPatternSet(USI_pat.Value, PatUSIArray, pat_count)
''
''    For j = 0 To pat_count - 1
''        TheExec.Datalog.WriteComment ""
''        TheExec.Datalog.WriteComment "UDRE USI Pattern: " + PatUSIArray(j)
''        For Each Site In TheExec.sites
''            TheExec.Datalog.WriteComment "Site(" + CStr(Site) + "), Shift In Code [" + CStr(UBound(PgmBitArr)) + ":0]=" + gS_UDRE_USI_BitStr(Site)
''        Next Site
''
''        UDR_SetupDigSrcArray PatUSIArray(j), InPin, "UDRE_USI_Src", usiarrSize, USI_Array
''        Call TheHdw.Patterns(PatUSIArray(j)).test(pfAlways, 0)
''    Next j
'
'    DebugPrintFunc USI_pat.Value
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_UDRE_USO_Syntax_Chk_Pgm2File(Optional condstr As String = "all") As Long

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_UDRE_USO_Syntax_Chk_Pgm2File"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Read patterns   =
    '==================================
    ''''20161114 update to Validate/load pattern
'    Dim ReadPatt As String
'    If (auto_eFuse_PatSetToPat_Validation(USO_pat, ReadPatt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Dim site As Variant
    Dim Trim_code_USO As New DSPWave
    Dim USO_BitStr As New SiteVariant
    Dim i As Long, j As Long, k As Long
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
    Dim FuseStr As String:: FuseStr = "UDRE"

    m_siteVar = "UDR_EChk_Var"
    
    ''TheHdw.Patterns(USO_pat).Load
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    'Status = GetPatListFromPatternSet(USO_pat.Value, PatUSOArray, pat_count)
    TheExec.Datalog.WriteComment ""
    
    ''''20171016 update
    ''''--------------------------------
    Dim m_testFlag As Boolean
    Dim m_CP1_Early_Flag As Boolean
    condstr = LCase(condstr)
    
    If (condstr = "cp1_early") Then
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
    
    Dim allBlank As New SiteBoolean
    Dim blank_stage As New SiteBoolean
    Dim m_bitFlag_mode As Long
    Dim CapWave As New DSPWave

    m_Fusetype = eFuse_UDRE
    m_FBC = -1               ''''initialize
    m_ResultFlag = -1        ''''initialize
    m_SiteVarValue = -1      ''''initialize
    allBlank = False
    blank_stage = True
    
    'Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(0), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
    'Call TheHdw.Patterns(PatUSOArray(0)).test(pfAlways, 0)
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    
    gDL_eFuse_Orientation = gE_eFuse_Orientation_1Bits
    Dim m_SerialType As Boolean:: m_SerialType = False
    If (gDL_eFuse_Orientation = eFuse_1_Bit) Then
        m_SerialType = True
        gDB_SerialType = True
    End If
        Dim m_PatBitOrder As String
    
    If (gL_UDRP_USO_PatBitOrder = "LSB") Then
        m_PatBitOrder = "bit0_bitLast"
    Else
        m_PatBitOrder = "bitLast_bit0"
    End If
    
'    If (TheExec.TesterMode = testModeOffline) Then
'        ''''Simulation  (capWave = Pgm_singleBitWave OR Read_singleBitWave)
'        ''''Here MarginRead it only read the programmed Site.
'        Dim Temp_USO As New DSPWave
'        Temp_USO.CreateConstant 0, USO_CapBits, DspLong
'
'        gL_eFuse_Sim_Blank = 0
'
'        Call RunDSP.eFuse_SingleBitWave2CapWave32Bits(eFuse_UDRP, Temp_USO)
'        'Call auto_eFuse_print_capWave32Bits(eFuse_UDR, Temp_USO, False) ''''True to print out
'        If (m_PatBitOrder = "bit0_bitLast") Then
'            For Each Site In TheExec.Sites
'                Trim_code_USO(Site) = Temp_USO(Site).Copy
'            Next
'        Else
'            Call ReverseWave(Temp_USO, Trim_code_USO, m_PatBitOrder, USO_CapBits)
'        End If
'    End If
    
    
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
    Call UpdateDLogColumns(gI_UDRE_catename_maxLen + 18)

    
    If (TheExec.EnableWord("Pgm2File")) Then
        'CapWave = gL_UDRPFuse_PgmBits.Copy
        For Each site In TheExec.sites
            CapWave(site) = gL_UDREFuse_PgmBits(site).Copy
        Next site
        
        Call rundsp.eFuse_Wave1bit_to_SingleDoubleBitWave(m_Fusetype, m_bitFlag_mode, m_SerialType, CapWave, m_FBC, blank_stage, allBlank)

        'Call RunDSP.eFuse_Wave32bits_to_SingleDoubleBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, blank_stage, allblank)
    End If
    ''''Offline simulation inside
    ''''Here Parameter bitFlag_mode=1 (Stage Bits)
    'If (TheExec.Sites(Site).SiteVariableValue(m_siteVar) = 1) Then
        'Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, m_bitFlag_mode, Trim_code_USO, m_FBC, m_cmpResult, , , True, m_PatBitOrder)
    'End If
    'Call auto_eFuse_DSSC_ReadDigCap_32bits_NEW(m_Fusetype, m_bitFlag_mode, Trim_code_USO, m_FBC, blank_stage, allblank, True)

    'Call auto_eFuse_compare_Read_PgmBitWave(m_Fusetype, 1, Trim_code_USO, m_FBC, m_cmpResult)


    Call auto_eFuse_setReadData_forSyntax(m_Fusetype)
    If (gB_eFuse_printReadCate) Then Call auto_eFuse_setReadData_forDatalog(m_Fusetype)

    Call auto_eFuse_print_DSSCReadWave_BitMap(m_Fusetype, gB_eFuse_printBitMap)
    Call auto_eFuse_print_DSSCReadWave_Category(m_Fusetype, False, gB_eFuse_printReadCate)
    
    Call auto_eFuse_pgm2file(FuseStr, gL_UDREFuse_PgmBits)
    
    'Call auto_eFuse_SyntaxCheck(m_Fusetype, condstr)

        ''''201811XX reset
    If (gS_JobName = "cp1_early") Then gS_JobName = "cp1" ''''<MUST>
Exit Function

End If


    
    
    
    ''20181120, add for pgm2file
'    If (condStr = "all" Or condStr = "") Then
'        Call auto_eFuse_pgm2file("UDRE", gL_UDREFuse_PgmBits)
'    End If

''    For j = 0 To pat_count - 1
''        Call auto_eFuse_DSSC_DigCapSetup(PatUSOArray(j), OutPin, "USO_cap", USO_CapBits, Trim_code_USO)
''        Call TheHdw.Patterns(PatUSOArray(j)).test(pfAlways, 0)
''
''        ''''20160506 update
''        ''''due to the additional characters "_USI_USO_compare", so plus 18.
''        Call UpdateDLogColumns(gI_UDRE_catename_maxLen + 18)
''
''        For Each Site In TheExec.sites
''            ''''initialize
''            USO_BitStr(Site) = ""  ''''MUST, and it's [bitLast ... bit0]
''            For i = 0 To UBound(DigCapArray)
''               DigCapArray(i) = 0  ''''MUST be [bit0 ... bitLast]
''            Next i
''
''            ''--------------------------------------------------------------------------------------
''            ''''20150717 update
''            ''''composite to the USO_BitStr() from the DSSC Capture
''            If (UCase(gL_UDRE_USO_PatBitOrder) = "MSB") Then
''                ''''<Notice>
''                ''''the pattern DSSC capture from MSB to LSB (bitLast to bit0)
''                ''''so Trim_code_USO.Element(0) is MSB
''                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
''                For i = 0 To USO_CapBits - 1
''                    DigCapArray(i) = Trim_code_USO.Element(USO_CapBits - 1 - i)    ''''Reverse Bit String
''                    USO_BitStr(Site) = USO_BitStr(Site) & Trim_code_USO.Element(i) ''''[MSB ... LSB]
''                Next i
''            Else
''                ''''<Notice>
''                ''''the pattern DSSC capture from LSB to MSB (bit0 to bitLast)
''                ''''so Trim_code_USO.Element(0) is LSB
''                ''''DigCapArray(0) Must be 'bit0' for the function auto_PrintAllBitbyDSSC()
''                For i = 0 To USO_CapBits - 1
''                    DigCapArray(i) = Trim_code_USO.Element(i)
''                    USO_BitStr(Site) = Trim_code_USO.Element(i) & USO_BitStr(Site) ''''[MSB ... LSB]
''                Next i
''            End If
''            ''--------------------------------------------------------------------------------------
''
''            ''''20160324 updae
''            ''''=============== Start of Simulated Data ===============
''            If (TheExec.TesterMode = testModeOffline) Then
''                ''''20160906 trial for the ugly codes
''                ''''<Issued codes> Shift out code [383:0]=111110000110100000000000000000001111100001101000000000000000000011111110111010000000000000000000111110100110100000000000000000001111101001101000000000000000000011111000011010000000000000000000111110000100000000000000000000001111100001000000000000000000000011111000010000000001000000010000111110000110000000000000000000001111100001000000000000000000000011111000010000000000000000000000
''                ''gS_UDRE_USI_BitStr(Site) = "111110000110100000000000000000001111100001101000000000000000000011111110111010000000000000000000111110100110100000000000000000001111101001101000000000000000000011111000011010000000000000000000111110000100000000000000000000001111100001000000000000000000000011111000010000000001000000010000111110000110000000000000000000001111100001000000000000000000000011111000010000000000000000000000"
''
''                USO_BitStr(Site) = gS_UDRE_USI_BitStr(Site)
''                For i = 0 To USO_CapBits - 1
''                    DigCapArray(i) = CLng(Mid(StrReverse(gS_UDRE_USI_BitStr(Site)), i + 1, 1))
''                Next i
''
''                If (gS_JobName <> "cp1" Or TheExec.sites.Item(Site).FlagState("F_UDRE_Early_Enable") = logicTrue) Then ''''was "cp1"
''                    TmpStr = ""
''                    For i = 0 To USO_CapBits - 1
''                        If (DigCapArray(i) = 0) Then
''                            DigCapArray(i) = gL_Sim_FuseBits(Site, i)
''                        Else
''                            gL_Sim_FuseBits(Site, i) = DigCapArray(i)
''                        End If
''                        TmpStr = CStr(DigCapArray(i)) + TmpStr
''                    Next i
''                    USO_BitStr(Site) = TmpStr ''''<MUST>
''                End If
''            End If
''            ''''===============   End of Simulated Data ===============
''
''            If (True) Then
''                '=======================================================
''                '= Print out the caputured bit data from DigCap        =
''                '=======================================================
''                TheExec.Datalog.WriteComment ""
''                Call auto_PrintAllBitbyDSSC(DigCapArray, USO_PrintRow, UBound(DigCapArray) + 1, USO_BitPerRow)
''            End If
''            ''--------------------------------------------------------------------------------------
''
''            TheExec.Datalog.WriteComment "Site(" & Site & "), UDRE USO pat:" & PatUSOArray(j) & ", Shift out code [" + CStr(USO_CapBits - 1) + ":0]=" + USO_BitStr(Site)
''            ''TheExec.Datalog.WriteComment ""
''
''            ''''----------------------------------------------------------------------------------------------
''            ''''20161114 update for print all bits (DTR) in STDF
''            ''''20171016 update to excluding "cp1_early"
''            If (m_CP1_Early_Flag = False) Then Call auto_eFuse_to_STDF_allBits("UDRE", USO_BitStr(Site))
''            ''''----------------------------------------------------------------------------------------------
''
''            ''''20150717 New
''            Call auto_Decode_UDRE_Binary_Data(DigCapArray)
''
''            ''''----------------------------------------------------------------------------------
''            ''''judge pass/fail for the specific test limit
''            tmpStrL = StrReverse(USO_BitStr(Site)) ''''translate to [LSB......MSB]
''            For i = 0 To UBound(UDRE_Fuse.Category)
''                tmpdlgStr = ""
''                m_stage = LCase(UDRE_Fuse.Category(i).Stage)
''                m_CateName = UDRE_Fuse.Category(i).Name
''                m_Algorithm = LCase(UDRE_Fuse.Category(i).Algorithm)
''                m_LSBbit = UDRE_Fuse.Category(i).LSBbit
''                m_MSBbit = UDRE_Fuse.Category(i).MSBbit
''                m_BitWidth = UDRE_Fuse.Category(i).BitWidth
''                m_lolmt = UDRE_Fuse.Category(i).LoLMT
''                m_hilmt = UDRE_Fuse.Category(i).HiLMT
''                m_defreal = LCase(UDRE_Fuse.Category(i).Default_Real)
''                m_resolution = UDRE_Fuse.Category(i).Resoultion
''
''                m_decimal = UDRE_Fuse.Category(i).Read.Decimal(Site)
''                m_value = UDRE_Fuse.Category(i).Read.Value(Site)
''                m_bitsum = UDRE_Fuse.Category(i).Read.BitSummation(Site)
''                m_hexStr = UDRE_Fuse.Category(i).Read.HexStr(Site)
''                m_unitType = unitNone
''                m_scale = scaleNone ''''default
''                m_tsName = Replace(m_CateName, " ", "_") ''''20151028, benefit for the script
''
''                m_bitStrM = StrReverse(Mid(tmpStrL, m_LSBbit + 1, m_BitWidth))
''
''                m_testFlag = True ''''20171016 update
''                If (m_CP1_Early_Flag = True) Then ''''20171016 update
''                    ''''only compare these category with stage="cp1_early"
''                    If (m_stage = condStr) Then
''                        ''''other cases
''                        m_testValue = m_decimal
''                    Else
''                        ''''Here it's an excluding case
''                        ''''<MUST>
''                        m_testFlag = False
''                        m_testValue = 0
''                        m_lolmt = 0
''                        m_hilmt = 0
''                    End If
''
''                ElseIf (m_Algorithm = "base") Then
''                    If (m_defreal = "decimal") Then ''''20160624 update
''                        m_testValue = m_decimal
''                    Else
''                        m_testValue = (m_decimal + 1) * gD_BaseStepVoltage
''                        m_unitType = unitVolt
''                        m_scale = scaleMilli
''                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
''                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
''                        m_testValue = m_value * 0.001 ''''to unit:V
''                    End If
''
''                ElseIf (m_Algorithm = "vddbin") Then
''                    ''''<Notice> User Maintain
''                    ''''Ex:: step_vdd_cpu_p1 = VDD_BIN(vdd_cpu_p1).MODE_STEP
''                    m_testValue = 0 ''''default to fail
''                    If (m_defreal = "decimal") Then ''''20160624 update
''                        m_testValue = m_decimal
''                    ElseIf (m_defreal Like "safe*voltage" Or m_defreal = "default") Then
''                        m_unitType = unitVolt
''                        m_scale = scaleMilli
''                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
''                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
''                        m_testValue = m_value * 0.001 ''''to unit:V
''                    ElseIf (m_defreal = "bincut") Then
''                        m_catenameVbin = m_CateName '150127
''                        ''''' Every testing stage needs to check this formula with bincut fused value (2017/5/17, Jack)
''                        vbinflag = auto_CheckVddBinInRangeNew(m_catenameVbin, m_resolution)
''
''                        ''''20160329 Add for the offline simulation
''                        If (vbinflag = 0 And TheExec.TesterMode = testModeOffline) Then
''                            vbinflag = 1
''                        End If
''
''                        m_Pmode = VddBinStr2Enum(m_catenameVbin)
''                        tmpVal = auto_VfuseStr_to_Vdd_New(m_bitStrM, m_BitWidth, m_resolution)
''                        MaxLevelIndex = BinCut(m_Pmode, CurrentPassBinCutNum).Mode_Step '150127
''                        m_lolmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmin(MaxLevelIndex) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(MaxLevelIndex)
''                        m_hilmt = BinCut(m_Pmode, CurrentPassBinCutNum).CP_Vmax(0) + BinCut(m_Pmode, CurrentPassBinCutNum).CP_GB(0)
''                        ''''judge the result
''                        If (vbinflag = 1) Then
''                            m_value = tmpVal
''                        Else
''                            m_value = -999
''                            TmpStr = m_CateName + "(Site " + CStr(Site) + ") = " + CStr(tmpVal) + " is not in range" '150127
''                            TheExec.Datalog.WriteComment TmpStr
''                        End If
''                        m_unitType = unitVolt
''                        m_scale = scaleMilli
''                        m_lolmt = m_lolmt * 0.001 ''''to unit:V
''                        m_hilmt = m_hilmt * 0.001 ''''to unit:V
''                        m_testValue = m_value * 0.001 ''''to unit:V
''                    End If
''                Else
''                    ''TheExec.Datalog.WriteComment "undefined Algorithm: " + m_algorithm
''                    ''''other cases, 20160927 update
''                    m_testValue = m_decimal
''                End If
''
''                ''''20160108 New
''                m_tsName = Replace(m_CateName, " ", "_") ''''20151028, benefit for the script
''                Call auto_eFuse_chkLoLimit("UDRE", i, m_stage, m_lolmt)
''                Call auto_eFuse_chkHiLimit("UDRE", i, m_stage, m_hilmt)
''
''                ''''20170811 update
''                If (m_BitWidth >= 32) Then
''                    ''m_tsName = m_tsName + "_" + m_hexStr
''                    ''''process m_lolmt / m_hilmt to hex string with prefix '0x'
''                    m_lolmt = auto_Value2HexStr(m_lolmt, m_BitWidth)
''                    m_hilmt = auto_Value2HexStr(m_hilmt, m_BitWidth)
''
''                    ''''------------------------------------------
''                    ''''compare with lolmt, hilmt
''                    ''''m_testValue 0 means fail
''                    ''''m_testValue 1 means pass
''                    ''''------------------------------------------
''                    m_testValue = auto_TestStringLimit(m_hexStr, CStr(m_lolmt), CStr(m_hilmt))
''                    m_lolmt = 1
''                    m_hilmt = 1
''                Else
''                    ''''20160620 update
''                    ''''20160927 update the new logical methodology for the unexpected binary decode.
''                    If (auto_isHexString(CStr(m_lolmt)) = True) Then
''                        ''''translate to double value
''                        m_lolmt = auto_HexStr2Value(m_lolmt)
''                    Else
''                        ''''doNothing, m_lolmt = m_lolmt
''                    End If
''
''                    If (auto_isHexString(CStr(m_hilmt)) = True) Then
''                        ''''translate to double value
''                        m_hilmt = auto_HexStr2Value(m_hilmt)
''                    Else
''                        ''''doNothing, m_hilmt = m_hilmt
''                    End If
''                End If
''
''                ''''20171016 update
''                If (m_testFlag) Then
''                    TheExec.Flow.TestLimit resultVal:=m_testValue, lowval:=m_lolmt, hival:=m_hilmt, Tname:=m_tsName, unit:=m_unitType, ScaleType:=m_scale
''                End If
''            Next i
''            ''''----------------------------------------------------------------------------------
''
''            ''''20160503 update for the 2nd way to prevent for the temp sensor trim will all zero values
''            ''''20160907 update
''            Dim m_valueSum As Long
''            Dim m_matchTMPS_flag As Boolean
''            m_valueSum = 0 ''''initialize
''            m_matchTMPS_flag = False
''            m_stage = "" ''''<MUST> 20160617 update, if the "trim/tmps" is existed then m_stage has its correct value.
''            For i = 0 To UBound(UDRE_Fuse.Category)
''                m_CateName = UCase(UDRE_Fuse.Category(i).Name)
''                m_Algorithm = LCase(UDRE_Fuse.Category(i).Algorithm)
''                If (m_CateName Like "TEMP_SENSOR*" Or m_Algorithm = "tmps") Then ''''was m_algorithm = "trim", 20171103 update
''                    m_stage = LCase(UDRE_Fuse.Category(i).Stage)
''                    m_decimal = UDRE_Fuse.Category(i).Read.Decimal(Site)
''                    m_valueSum = m_valueSum + m_decimal
''                    m_matchTMPS_flag = True
''                End If
''            Next i
''            If (m_matchTMPS_flag = True) Then
''                ''''if Job >= m_stage then m_valueSim >= 1
''                If (checkJob_less_Stage_Sequence(m_stage) = False) Then
''                    TheExec.Flow.TestLimit resultVal:=m_valueSum, lowval:=1, Tname:="UDRE_TMPS_SUM"
''                    ''TheExec.Datalog.WriteComment ""
''                Else
''                    ''''if Job < m_stage then m_valueSim = 0
''                    ''''20180105 update
''                    If (gB_eFuse_Disable_ChkLMT_Flag = True) Then ''''case CP2 back to CP1 retest
''                        TheExec.Flow.TestLimit resultVal:=m_valueSum, lowval:=0, Tname:="UDRE_TMPS_SUM"
''                    Else
''                        TheExec.Flow.TestLimit resultVal:=m_valueSum, lowval:=0, hival:=0, Tname:="UDRE_TMPS_SUM"
''                    End If
''                End If
''            End If
''            ''''--------------------------------------------------------------------------------------------
''
''            ''''20160503 update
''            ''''compare both USI and USO for the specific stage, it's only when siteVar is '1'.
''            ''''Must be after the decode then you have the Read buffer value
''            ''''20180105 update
''            ''''The below is used to compare both USI and USO contents.
''            If (TheExec.sites(Site).SiteVariableValue(m_siteVar) = 1) Then
''                Dim m_writeBitStrM As String
''                Dim m_readBitStrM As String
''                Dim m_usiusoCmp As Long
''                m_usiusoCmp = 0 ''''<MUST> default Compare Pass:0, Fail:1
''                For i = 0 To UBound(UDRE_Fuse.Category)
''                    m_stage = LCase(UDRE_Fuse.Category(i).Stage)
''                    If (m_stage = gS_JobName) Then
''                        m_CateName = UDRE_Fuse.Category(i).Name
''                        m_writeBitStrM = UCase(UDRE_Fuse.Category(i).Write.BitStrM(Site))
''                        m_readBitStrM = UCase(UDRE_Fuse.Category(i).Read.BitStrM(Site))
''                        m_tsName = Replace(m_CateName, " ", "_")
''                        m_tsName = m_tsName + "_USI_USO_compare"
''                        If (m_writeBitStrM <> m_readBitStrM) Then
''                            ''''20180105 update
''                            m_usiusoCmp = 1 ''''Fail Comparison
''                            TheExec.Datalog.WriteComment vbTab & "Site(" + CStr(Site) + "), " + m_tsName + " Failure."
''                        Else
''                            ''''case: m_writeBitStrM = m_readBitStrM
''                            ''''reserve
''                        End If
''                    End If
''                Next i
''                TheExec.Flow.TestLimit resultVal:=m_usiusoCmp, lowval:=0, hival:=0, Tname:="UDRE_USO_USI_Cmp"
''            End If
''            ''''--------------------------------------------------------------------------------------------
''
''            TheExec.Datalog.WriteComment ""
''            If (TheExec.sites.ActiveCount = 0) Then Exit Function 'chihome
''        Next Site
''        Call UpdateDLogColumns__False
''    Next j

     ''''20171016 update
    If (m_CP1_Early_Flag = True) Then
        gS_JobName = "cp1" ''''reset
    End If
    DebugPrintFunc USO_pat.Value

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''20161114, update with Validating_ for the validate/load the pattern in OnProgramValidation
Public Function auto_MONITORWrite_byCondition_Pgm2File(condstr As String, _
                                                       Optional InterfaceType As String = "")
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MONITORWrite_byCondition_Pgm2File"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Write patterns  =
    '==================================
    ''''20161114 update to Validate/load pattern
'    Dim WritePatt As String
'    If (auto_eFuse_PatSetToPat_Validation(WritePattSet, WritePatt, Validating_) = True) Then Exit Function
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
    
'    ''20181120, add for pgm2file
'    Dim m_tmpData As New DSPWave
'    m_tmpData.CreateConstant 0, gDL_BitsPerBlock, DspLong

    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

    TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName
    'Call TurnOnEfusePwrPins(PwrPin, vpwr)

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
    'DigSrcSignalName = "MONITOR_DigSrcSignal"

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
    Dim m_Fusetype As eFuseBlockType:: m_Fusetype = eFuse_MON
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    gDL_eFuse_Orientation = gE_eFuse_Orientation
    
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
    
        ''20181120, add for pgm2file
    Dim m_tmpData As New DSPWave
    m_tmpData.CreateConstant 0, gDL_BitsPerBlock, DspLong
    
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
    'Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_MON_Pgm_SingleBitWave, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_MON, gDW_MON_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)
    
    For Each site In TheExec.sites
        m_tmpData(site) = gDW_MON_Pgm_SingleBitWave(site).Copy 'was eFuse_Pgm_Bit
        gL_MONFuse_PgmBits(site) = gL_MONFuse_PgmBits(site).BitwiseOr(m_tmpData(site))
    Next
    
'    If (m_cmpStage = "cp1_early") Then
'        ''''if it's same values on all Sites to save TT and improve PTE
'        Call eFuse_DSSC_SetupDigSrcWave_allSites(WritePatt, PinWrite, DigSrcSignalName, m_pgmDigSrcWave)
'    Else
'        Call eFuse_DSSC_SetupDigSrcWave(WritePatt, PinWrite, DigSrcSignalName, m_pgmDigSrcWave)
'    End If
'
'    ''''In the MarginRead process, it will use gDW_XXX_Pgm_SingleBitWave / gDW_XXX_Pgm_DoubleBitWave to do the comparison with Read
'
'    ''''Write Pattern for programming eFuse
'    Call TheHdw.Patterns(WritePatt).test(pfAlways, 0)   'Write ECID
'
'    Call TurnOffEfusePwrPins(PwrPin, vpwr)
'    DebugPrintFunc WritePattSet.Value
    
Exit Function
    
End If

    
    
    
    
'    Dim m_tmpData As New DSPWave
'    m_tmpData.CreateConstant 0, MONITORTotalBitCount, DspLong
'
'    For Each Site In TheExec.sites
'
'        If (TheExec.TesterMode = testModeOffline) Then ''''20160526 update
'            Call eFuseENGFakeValue_Sim
'        End If
'
'        ''''20160104 New
'        If (condStr = "stage") Then
'            SegmentSize = auto_Make_MONITOR_Pgm_and_Read_Array(eFuse_Pgm_Bit(), Expand_eFuse_Pgm_Bit(), eFusePatCompare())
'        ElseIf (condStr = "category") Then
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
'        ''20181120, add for pgm2file
'        m_tmpData.Data = eFuse_Pgm_Bit
'        gL_MONFuse_PgmBits = gL_MONFuse_PgmBits.BitwiseOr(m_tmpData)
'
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

'    TheHdw.DSSC.Pins(PinWrite).pattern(WritePatt).Source.Signals.DefaultSignal = DigSrcSignalName
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

Public Function auto_MONITORSingleDoubleBit_Pgm2File(condstr As String, _
                                                    Optional InterfaceType As String = "")
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_MONITORSingleDoubleBit_Pgm2File"
    
    Dim site As Variant

    Dim i As Long, j As Long, k As Long

    ''Dim SingleStrArray() As String
    Dim SingleBitArray() As Long
    Dim DoubleBitArray() As Long

    Dim tmpStr As String
    Dim FuseStr As String:: FuseStr = "MON"
    
    ''''--------------------------------------------------------------------------------
    ''ReDim SingleStrArray(MONITORReadCycle - 1, TheExec.Sites.Existing.Count - 1)
    ''ReDim DoubleBitArray(MONITORBitPerBlockUsed - 1)
    ''ReDim SingleBitArray(MONITORTotalBitCount - 1)
    
    '=============================================================================
    '=  Setup Voltage and Timing then ramp-ip power on VDD18_EFUSE0/1            =
    '=============================================================================
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

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
    
    If (TheExec.EnableWord("Pgm2File")) Then ''''for Pgm2File (Pgm2Read)
        gL_MON_FBC = 0 ''''set dummy
        
            gDL_eFuse_Orientation = gE_eFuse_Orientation
        ''''----------------------------------------------------
        ''''201812XX New Method by DSPWave
        ''''----------------------------------------------------
        Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
        Dim m_Fusetype As eFuseBlockType
        Dim m_SiteVarValue As New SiteLong
        Dim m_ResultFlag As New SiteLong
        Dim CapWave As New DSPWave
        Dim m_bitFlag_mode As Long
        Dim blank_stage As New SiteBoolean
        Dim allBlank As New SiteBoolean
    
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
        
        If (LCase(condstr = "cp1_early")) Then
            m_bitFlag_mode = 0
        ElseIf (LCase(condstr) = "stage") Then
            m_bitFlag_mode = 1
        ElseIf (LCase(condstr) = "all") Then
            m_bitFlag_mode = 2
        ElseIf (LCase(condstr) = "real") Then
            m_bitFlag_mode = 3
        Else
            ''''default, here it prevents any typo issue
            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
            m_FBC = -1
            'm_cmpResult = -1
        End If
        
        If (InterfaceType = "APB") Then
            For Each site In TheExec.sites
                CapWave(site) = gL_MONFuse_PgmBits(site).Copy
            Next site
             Call rundsp.eFuse_Wave1bit_to_SingleDoubleBitWave(m_Fusetype, m_bitFlag_mode, m_SerialType, CapWave, m_FBC, blank_stage, allBlank)
            
        Else
            For Each site In TheExec.sites
                CapWave(site) = gL_MONFuse_PgmBits(site).ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb)
            Next site
            Call rundsp.eFuse_Wave32bits_to_SingleDoubleBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, blank_stage, allBlank)
        End If

        'Call RunDSP.eFuse_Wave32bits_to_SingleDoubleBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, blank_stage, allblank)
    End If
    
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
    
    Call auto_eFuse_pgm2file(FuseStr, gL_MONFuse_PgmBits)
    
    
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
    
    
    
    
    
'
'
'
'
'
'    ''20181120, add for pgm2file
'    Call auto_eFuse_pgm2file("MON", gL_MONFuse_PgmBits)
'
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
'''        Call auto_OR_2Blocks("MON", gS_SingleStrArray, SingleBitArray, DoubleBitArray) ''''get gL_MON_FBC(site)
'
'        ''20181120, add for pgm2file
'        If (TheExec.EnableWord("Pgm2File")) Then ''''for Pgm2File (Pgm2Read)
'            SingleBitArray = gL_MONFuse_PgmBits.Data
'            Call auto_Gen_DoubleBitArray("MON", SingleBitArray, DoubleBitArray, 0)
'            gL_MON_FBC(Site) = 0 ''''set dummy
'        Else
'            Call auto_OR_2Blocks("MON", gS_SingleStrArray(), SingleBitArray(), DoubleBitArray())
'        End If
'
'        ''''20170220 update
'        If (gL_MON_FBC(Site) > 0) Then
'            TmpStr = "The Fail Bit Count of MONITOR eFuse at Site(" + CStr(Site) + ") is " + CStr(gL_MON_FBC(Site))
'            TmpStr = TmpStr + " (Max FBC =0)"
'            TheExec.Datalog.WriteComment TmpStr
'        End If
'
'        'If (TheExec.sites(Site).SiteVariableValue("MONChk_Var") <> 1 Or gB_MON_decode_flag(Site) = False) Then
'            ''''ReTest Stage
'            ''''<Important> User Need to check the content inside
'            'Call auto_Decode_MONBinary_Data(DoubleBitArray)
'
'            ''''' 20161019 update
'            If (checkJob_less_Stage_Sequence(gS_MON_CRC_Stage) = True) Then
'                gS_MON_CRC_HexStr(Site) = "00000000"
'            Else
'                gS_MON_CRC_HexStr(Site) = auto_MONITOR_CRC2HexStr(DoubleBitArray)
'            End If
'        'End If
'
'        ''''----------------------------------------------------------------------------------------------
'        gS_MON_Direct_Access_Str(Site) = "" ''''is a String [(bitLast)......(bit0)]
'        For i = 0 To UBound(DoubleBitArray)
'            gS_MON_Direct_Access_Str(Site) = CStr(DoubleBitArray(i)) + gS_MON_Direct_Access_Str(Site)
'        Next i
'        ''TheExec.Datalog.WriteComment "gS_MON_Direct_Access_Str=" + CStr(gS_MON_Direct_Access_Str(Site))
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

''''nonDEID=non DeviceID
Public Function auto_EcidSingleDoubleBit_nonDEID_Pgm2File(Optional InitPinsHi As PinList, _
                                                          Optional InitPinsLo As PinList, _
                                                          Optional InitPinsHiZ As PinList)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidSingleDoubleBit_nonDEID_Pgm2File"
    
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
    Dim FuseStr As String:: FuseStr = "ECID"

    ''''------------------------------------------------------------------------------------------------------------------
    ''''<Important Notice>
    ''''gS_SingleStrArray() was extracted in the module auto_ECID_Read_by_OR_2Blocks_TMPS_ADC() then used in auto_EcidSingleDoubleBit_TMPS_ADC()
    ''''gS_SingleStrArray() is the result of the NormRead or MarginRead
    ''''
    ''''So it doesn't need to run the pattern and DSSC to get the SignalStrArray, and save test time
    ''''------------------------------------------------------------------------------------------------------------------
    TheExec.Datalog.WriteComment vbCrLf & "Test Instance   :: " + TheExec.DataManager.instanceName

    cycleNum = EcidReadCycle: BitPerCycle = ECIDBitPerCycle
    
    
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If
    
    If (TheExec.EnableWord("Pgm2File") = True) Then
        Dim tmpWave As New DSPWave
        For Each site In TheExec.sites
            gDW_ECID_Read_DoubleBitWave(site) = gL_ECIDFuse_PgmBits(site).Copy
        Next
        Call rundsp.eFuse_decode_DSSCReadWave(eFuse_ECID, gDW_ECID_Read_DoubleBitWave, tmpWave)
    End If

    TheExec.Datalog.WriteComment "<" & TheExec.DataManager.instanceName & ">"
    
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
    
    Call auto_eFuse_pgm2file(FuseStr, gL_ECIDFuse_PgmBits)

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
    'theexec.Flow.TestLimit resultVal:=gL_ECID_FBC, lowVal:=0, hiVal:=0, TName:="ECID_FBCount_" + UCase(gS_JobName) '2d-s=0
    
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

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function auto_EcidWrite_byCondition_Pgm2File(condstr As String)
'Public Function auto_EcidWrite_byCondition_Pgm2File(WritePattSet As Pattern, _
'                                                    condstr As String, _
'                                                    Optional catename_grp As String, _
'                                                    Optional Validating_ As Boolean)
'
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_EcidWrite_byCondition"

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load Write patterns  =
    '==================================
    ''''20161114 update to Validate/load pattern
    'Dim WritePatt As String
    'If (auto_eFuse_PatSetToPat_Validation(WritePattSet, WritePatt, Validating_) = True) Then Exit Function
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
    'TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ  'SEC DRAM

    TheExec.Datalog.WriteComment vbCrLf & "TestInstance:: " + TheExec.DataManager.instanceName

    'Call TurnOnEfusePwrPins(PwrPin, vpwr)
   
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
    
    'DigSrcSignalName = "DigSrcSignal"

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
    Dim m_Fusetype As eFuseBlockType
    m_Fusetype = eFuse_ECID
    
    Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
    gDL_eFuse_Orientation = gE_eFuse_Orientation
    
    ''20181120, add for pgm2file
    Dim m_tmpData As New DSPWave
    m_tmpData.CreateConstant 0, gDL_BitsPerBlock, DspLong

    m_pgmWave.CreateConstant 0, gDL_BitsPerBlock, DspLong ''''gDL_BitsPerBlock = EcidBitPerBlockUsed
    ''''<MUST and Importance>
    For Each site In TheExec.sites
        gDW_Pgm_RawBitWave = m_pgmWave.Copy
    Next site

    If (condstr = LCase("DEID") Or condstr = "cp1_early") Then
        m_cmpStage = "cp1_early"
        gS_JobName = "cp1_early"
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
        If (i = 16) Then
            Debug.Print "X"
        End If
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
    'Call auto_eFuse_print_PgmBitsWave_BitMap(gDW_ECID_Pgm_SingleBitWave, gB_eFuse_printBitMap)
    Call auto_eFuse_print_PgmBitsWave_Category(eFuse_ECID, gDW_ECID_Pgm_DoubleBitWave, gB_eFuse_printPgmCate)
    
    For Each site In TheExec.sites
        m_tmpData(site) = gDW_ECID_Pgm_SingleBitWave(site).Copy 'was eFuse_Pgm_Bit
        gL_ECIDFuse_PgmBits(site) = gL_ECIDFuse_PgmBits(site).BitwiseOr(m_tmpData(site))
    Next
    
    If (m_cmpStage = "cp1_early") Then
        gS_JobName = "cp1"
    End If
    
    'Call eFuse_DSSC_SetupDigSrcWave(WritePatt, PinWrite, DigSrcSignalName, m_pgmDigSrcWave)
    
    ''''if it's same values on all Sites to save TT and improve PTE
    ''Call eFuse_DSSC_SetupDigSrcWave_allSites(WritePatt, PinWrite, DigSrcSignalName, m_pgmDigSrcWave)

''''    ''''It will be used @ MarginRead process
''''    Call RunDSP.eFuse_DspWave_Copy(m_pgmSingleBitWave, gW_ECID_Pgm_singleBitWave)
''''    Call RunDSP.eFuse_DspWave_Copy(m_pgmDoubleBitWave, gW_ECID_Pgm_doubleBitWave)

    ''''Write Pattern for programming eFuse
   ' Call TheHdw.Patterns(WritePatt).test(pfAlways, 0)   'Write ECID
    'TheHdw.Wait 100# * us

    'Call TurnOffEfusePwrPins(PwrPin, vpwr)
    'DebugPrintFunc WritePattSet.Value
    
Exit Function
    
End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function auto_EcidSingleDoubleBit_Pgm2File(Optional condstr As String = "")

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
    Dim FuseStr As String:: FuseStr = "ECID"
    ''ReDim SingleBitArrayStr(ECIDTotalBits - 1)

    TheExec.Datalog.WriteComment vbCrLf & "Test Instance :: " + TheExec.DataManager.instanceName

''''201811XX update
If (gB_eFuse_newMethod) Then
    If (gB_eFuse_DSPMode = True) Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic       ''''tlDSPModeForceAutomatic, tlDSPModeAutomatic
    Else
        TheHdw.DSP.ExecutionMode = tlDSPModeHostThread      ''''tlDSPModeHostThread, tlDSPModeHostDebug
    End If

    If (TheExec.EnableWord("Pgm2File")) Then ''''for Pgm2File (Pgm2Read)
        gL_ECID_FBC = 0 ''''set dummy
        
        Dim m_FBC As New SiteLong ''''Fail bit count of the singleBit and doobleBit
        Dim m_Fusetype As eFuseBlockType
        Dim m_SiteVarValue As New SiteLong
        Dim m_ResultFlag As New SiteLong
        Dim m_bitFlag_mode As Long
        Dim blank_stage As New SiteBoolean
        Dim allBlank As New SiteBoolean
        Dim CapWave As New DSPWave
    
        gDL_eFuse_Orientation = gE_eFuse_Orientation
        gL_eFuse_Sim_Blank = 0
        
        m_Fusetype = eFuse_ECID
        m_FBC = -1               ''''initialize
        m_ResultFlag = -1        ''''initialize
        m_SiteVarValue = -1      ''''initialize
        allBlank = True
        blank_stage = True
        
        Call auto_eFuse_param2globalDSPVar(m_Fusetype) ''''<MUST be here First>
        
        condstr = "all"
        If (LCase(condstr = "cp1_early")) Then
            m_bitFlag_mode = 0
        ElseIf (LCase(condstr) = "stage") Then
            m_bitFlag_mode = 1
        ElseIf (LCase(condstr) = "all") Then
            m_bitFlag_mode = 2
        ElseIf (LCase(condstr) = "real") Then
            m_bitFlag_mode = 3
        Else
            ''''default, here it prevents any typo issue
            TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: please have a correct condStr (cp1_early,all,stage)"
            m_FBC = -1
            'm_cmpResult = -1
        End If
        
        For Each site In TheExec.sites
            CapWave(site) = gL_ECIDFuse_PgmBits(site).ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb)
        Next site

        Call rundsp.eFuse_Wave32bits_to_SingleDoubleBitWave(m_Fusetype, m_bitFlag_mode, CapWave, m_FBC, blank_stage, allBlank)
    End If

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
        Else
            ''''all CP cases and WLFT from the prober
            Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(site, XCoord(site))
            Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(site, YCoord(site))
        End If
    
    Next site
    ''''----------------------------------------------------------------------------------------------
    
    Call auto_eFuse_pgm2file(FuseStr, gL_ECIDFuse_PgmBits)
    
    Call UpdateDLogColumns(gI_ECID_catename_maxLen)
    ''for ECID Syntax check------------------------------------------------------------------------------
    'Call auto_ECID_SyntaxCheck_DEID
    'Call auto_ECID_SyntaxCheck_All
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
    
    
    
    
    
    

'    ''''Here using the programming stage of 'Lot_ID' stands for the Stage for all ECID bits.(It Should BE!!!)
'    m_stage = LCase(ECIDFuse.Category(ECIDIndex("Lot_ID")).Stage)
'
'    For Each Site In TheExec.sites
'        ' OR 2 block bit by bit
'        ' calc gL_ECID_FBC
'''        Call auto_OR_2Blocks("ECID", gS_SingleStrArray(), SingleBitArray(), DoubleBitArray())     'calc gL_ECID_FBC
'        ''20181120, add for pgm2file
'        If (TheExec.EnableWord("Pgm2File")) Then
'            SingleBitArray = gL_ECIDFuse_PgmBits.Data
'            Call auto_Gen_DoubleBitArray("ECID", SingleBitArray, DoubleBitArray, 0)
'            gL_ECID_FBC(Site) = 0
'        Else
'            Call auto_OR_2Blocks("ECID", gS_SingleStrArray(), SingleBitArray(), DoubleBitArray())
'        End If
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
'''        If (TheExec.sites(Site).SiteVariableValue("ECIDChk_Var") <> 1 Or gB_ECID_decode_flag(Site) = False) Then
'            If (gS_EFuse_Orientation = "SingleUp") Then
'                ''''Only 1 block
'                Call auto_EcidPrintData(1, SingleBitArrayStr)
'            Else
'                Call auto_EcidPrintData(1, SingleBitArrayStr)
'                Call auto_EcidPrintData(2, SingleBitArrayStr)
'            End If
'            TheExec.Datalog.WriteComment ""
'            Call auto_PrintAllBitbyDSSC(SingleBitArray(), EcidReadCycle, ECIDTotalBits, EcidReadBitWidth)
'''        End If
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
'        TheExec.Flow.TestLimit resultVal:=ChkResult, lowVal:=1, hiVal:=1, TName:="ECID_Syntax_Chk"
'        TheExec.Flow.TestLimit resultVal:=gL_ECID_FBC, lowVal:=0, hiVal:=EcidHiLimitSingleDoubleBitCheck, TName:="FailBitCount"  '2d-s=0
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
'        Dim LotStr As String
'        Dim Waferstr As String
'        Dim X_Coor_Str As String
'        Dim Y_Coor_Str As String
'        Dim ECID_DEID_Str As String
'
'        LotStr = ""
'        Waferstr = ""
'        X_Coor_Str = ""
'        Y_Coor_Str = ""
'        ECID_DEID_Str = ""
'
'        For Each Site In TheExec.sites.Existing
'
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
'
'        Next Site
'
'        LotStr = auto_checkIEDAString(LotStr)
'        Waferstr = auto_checkIEDAString(Waferstr)
'        X_Coor_Str = auto_checkIEDAString(X_Coor_Str)
'        Y_Coor_Str = auto_checkIEDAString(Y_Coor_Str)
'        ECID_DEID_Str = auto_checkIEDAString(ECID_DEID_Str)
'
'        TheExec.Datalog.WriteComment vbCrLf & "Test Instance :: " + TheExec.DataManager.InstanceName
'        TheExec.Datalog.WriteComment " ECID (all sites iEDA format)::"
'        TheExec.Datalog.WriteComment " Lot ID    = " + LotStr
'        TheExec.Datalog.WriteComment " Wafer ID  = " + Waferstr
'        TheExec.Datalog.WriteComment " X_Coor    = " + X_Coor_Str
'        TheExec.Datalog.WriteComment " Y_Coor    = " + Y_Coor_Str
'        TheExec.Datalog.WriteComment " ECID_DEID = " + ECID_DEID_Str & vbCrLf
'
'        '============================================
'        '=  Write Data to Register Edit (HKEY)      =
'        '============================================
'        Call RegKeySave("eFuseLotNumber", LotStr)
'        Call RegKeySave("eFuseWaferID", Waferstr)
'        Call RegKeySave("eFuseDieX", X_Coor_Str)
'        Call RegKeySave("eFuseDieY", Y_Coor_Str)
'        Call RegKeySave("Hram_ECID_53bit", ECID_DEID_Str)
'
'    End If
'    ''''End of register and print out the IEDA data----------------------------------------------------------------------------
    
Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next

End Function


