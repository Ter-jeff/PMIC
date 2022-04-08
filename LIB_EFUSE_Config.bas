Attribute VB_Name = "LIB_EFUSE_Config"

Option Explicit

Public Function auto_GetSiteFlagName(ByRef cnt As Long, ByRef flagName As String, Optional showPrint As Boolean = False)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_GetSiteFlagName"
        
    Dim ss As Variant
    Dim SiteFlagBoolean As New SiteLong
    Dim i As Long
    Dim m_pkgname As String

    If (showPrint) Then TheExec.Datalog.WriteComment ""

    For Each ss In TheExec.sites
        cnt = 0
        For i = 0 To UBound(CFGTable.Category)
            m_pkgname = CFGTable.Category(i).pkgName
            SiteFlagBoolean = TheExec.Flow.SiteFlag(ss, m_pkgname)
            If SiteFlagBoolean > 0 Then
                cnt = cnt + 1
                ''''<Important> 20170630 check if it is already existed '''' add
                If (UCase(m_pkgname) = UCase(flagName)) Then
                    Exit For
                Else
                    flagName = m_pkgname
                End If
            End If
        Next i
        'This Cnt >0 implies the enabled siteFlag has been detected
        If cnt > 0 Then Exit For
    Next ss
    
    'We just allow one siteFlag selected. So, if no or more than one siteFlag selected, we have to stop and
    'exit the testing.
    If cnt > 1 Then
        TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: There are more than one CFG eFuse conditions Flag selected. Please check it!! "
    ElseIf cnt = 0 Then
        ''If (showPrint) Then TheExec.Datalog.WriteComment "<WARNING> " + funcName + ":: There is NO any CFG eFuse conditions Flag selected. Please check it!! "
    Else
        If (showPrint) Then TheExec.Datalog.WriteComment funcName + ":: The CFG eFuse Condition Flag is " + flagName + "."
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function auto_CheckVddBinInRangeNew(binflowStr As String, m_stepVoltage As Double) As Long
                                
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_CheckVddBinInRangeNew"
    
    Dim i As Long
    Dim vbinIdx As Long
    Dim step_vdd As Long
    Dim remainder As Double
    Dim divisor As Double
    Dim cal_voltage As Double
    Dim ids_current As New SiteDouble
    Dim m_catename As String
    Dim m_decimal As Long
    Dim m_resolution As Double
    Dim m_match_flag As Boolean
    Dim m_dlgstr As String
    Dim ss As Variant

    ss = TheExec.sites.SiteNumber
    
    '---------------------------------------------'
    divisor = m_stepVoltage
    '---------------------------------------------'

    binflowStr = UCase(binflowStr) ''''to make sure it's UCase
    vbinIdx = VddBinStr2Enum(binflowStr)
    step_vdd = BinCut(vbinIdx, CurrentPassBinCutNum).Mode_Step
   
    ''get IDS value
    Dim m_pwrpin As String
    Dim m_pwrpinIDSname As String
    'm_pwrpin = UCase(BinCut(vbinIdx, CurrentPassBinCutNum).PowerPin)
    m_pwrpin = UCase(AllBinCut(vbinIdx).powerPin)
    m_pwrpinIDSname = "IDS_" + m_pwrpin
    
    m_match_flag = False
    For i = 0 To UBound(CFGFuse.Category)
        m_catename = UCase(CFGFuse.Category(i).Name)
        If (m_pwrpinIDSname = m_catename) Then ''''20160729 update
            m_decimal = CFGFuse.Category(CFGIndex(m_catename)).Read.Decimal(ss)
            ''''<Be carefully> because (.Resolution)'s unit is mA, so multiple by 0.001 then unit changes from mA to A.
            m_resolution = CFGFuse.Category(CFGIndex(m_catename)).Resoultion * 0.001
            ids_current(ss) = m_decimal * m_resolution
            m_match_flag = True
            Exit For
        End If
    Next i

    ''''initialization
    auto_CheckVddBinInRangeNew = 0
    
    If (m_match_flag = False) Then
        m_dlgstr = "<WARNING> " + funcName + ":: There is NO matching IDS Name (" + m_pwrpinIDSname + ")."
        TheExec.AddOutput m_dlgstr
        TheExec.Datalog.WriteComment m_dlgstr
        auto_CheckVddBinInRangeNew = 0
        Exit Function
    End If
    
    If (ids_current(ss) <= 0) Then
        m_dlgstr = "<WARNING> " + funcName + ":: There is zero/negative IDS Current Value (" + CStr(ids_current(ss)) + ")."
        TheExec.AddOutput m_dlgstr
        TheExec.Datalog.WriteComment m_dlgstr
        auto_CheckVddBinInRangeNew = 0
        Exit Function
    End If
    
    For i = 0 To step_vdd
        cal_voltage = BinCut(vbinIdx, CurrentPassBinCutNum).c(i) - BinCut(vbinIdx, CurrentPassBinCutNum).M(i) * (Log(ids_current(ss) * 1000) / Log(10))
        remainder = cal_voltage / divisor
        remainder = Floor(remainder)
        cal_voltage = remainder * divisor
        If cal_voltage > BinCut(vbinIdx, CurrentPassBinCutNum).CP_Vmax(i) Then
            cal_voltage = BinCut(vbinIdx, CurrentPassBinCutNum).CP_Vmax(i)
        ElseIf cal_voltage < BinCut(vbinIdx, CurrentPassBinCutNum).CP_Vmin(i) Then
            cal_voltage = BinCut(vbinIdx, CurrentPassBinCutNum).CP_Vmin(i)
        End If
        If VBIN_RESULT(vbinIdx).GRADEVDD(ss) = cal_voltage + BinCut(vbinIdx, CurrentPassBinCutNum).CP_GB(i) Then
            auto_CheckVddBinInRangeNew = 1 'pass
            Exit For
        End If
    Next i

Exit Function

errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
     
End Function

''''20160330 Update for the offline simulation
Public Function auto_Copy_CFGTable_Data_to_Array_forSim(p_stage As String, Flag1 As String, StartBit As Long, lastBit As Long, _
                                                        ByRef Array1() As Long, Optional chkStage As Boolean = True) As Double

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Copy_CFGTable_Data_to_Array_forSim"

    Dim i As Long, j As Long
    Dim m_bitwidth As Long
    Dim m_size As Long
    Dim m_jobMatch As Boolean  'can not assume A00 as all blank, it comes from A00 with SVM case

    Dim ss As Variant
    ss = TheExec.sites.SiteNumber

    m_bitwidth = lastBit - StartBit + 1
    Flag1 = UCase(Trim(Flag1))
    If (Flag1 = "") Then Flag1 = "ALL_0"

    If (chkStage = True) Then
        If (LCase(p_stage) = gS_JobName) Then
            m_jobMatch = True
        Else
            m_jobMatch = False
        End If
    Else
        ''''bypass stage check and set m_jobMatch is True
        m_jobMatch = True
    End If

    ''''<Important> default: all Config Security Code bits = 0
    For j = 0 To (m_bitwidth - 1)
        Array1(StartBit + j) = 0
    Next j
        
    If (m_jobMatch = True And Flag1 <> "ALL_0") Then
        For i = 0 To UBound(CFGTable.Category)
            If (Flag1 = UCase(CFGTable.Category(i).pkgName)) Then
                m_size = UBound(CFGTable.Category(i).BitVal) + 1
                If (m_bitwidth = m_size) Then
                    For j = 0 To UBound(CFGTable.Category(i).BitVal)
                        Array1(StartBit + j) = CFGTable.Category(i).BitVal(j)
                    Next j
                Else
                    TheExec.Datalog.WriteComment "<WARNING> UnMatch Length between BitWidth (" + CStr(m_bitwidth) + " and CFGTable_BitValue size (" + CStr(m_size) + ") !!!"
                End If
                Exit For
            End If
        Next i
        
        ''''20151230
        ''''Here is a special case to process bit57 (gC_CFGSVM_BIT)
        ''''If bit57 is already blown, ex CFG_A00 at CFG_SVM mode
        If (gB_CFG_SVM = True And gB_CFGSVM_BIT_Read_ValueisONE(ss) = True) Then
            If (StartBit = 0) Then
                TheExec.Datalog.WriteComment "<NOTICE> " + funcName + ":: Site(" + CStr(ss) + ") CFG Bit(" + CStr(gC_CFGSVM_BIT) + ") is ONE already, Set PgmBit=0 to avoid Re-Blown."
                Array1(gC_CFGSVM_BIT) = 0 ''''NO blown again
            End If
        End If
    Else
        ''''default: all bits = 0 (if NOT match job)
        ''''doNothing
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

''''Update for the offline simulation
''''20170630 Copy_CFG_CondTable_Data_to_Array_forSim
Public Function auto_Copy_CFG_CondTable_Data_to_Array_forSim(Flag1 As String, StartBit As Long, lastBit As Long, _
                                                        ByRef Array1() As Long, Optional chkStage As Boolean = True) As Double

On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_Copy_CFG_CondTable_Data_to_Array_forSim"

    Dim i As Long, j As Long
    Dim m_bitwidth As Long
    Dim m_size As Long
    Dim m_jobMatch As Boolean  'can not assume A00 as all blank, it comes from A00 with SVM case
    Dim p_stage As String
    Dim ss As Variant
    ss = TheExec.sites.SiteNumber

    m_bitwidth = lastBit - StartBit + 1
    Flag1 = UCase(Trim(Flag1))
    If (Flag1 = "") Then Flag1 = "ALL_0"

    If (chkStage = True) Then
        ''''20170630 update
        m_jobMatch = False
        For i = 0 To UBound(CFGTable.Category(0).condition)
            p_stage = LCase(CFGTable.Category(0).condition(i).Stage)
            If (p_stage = gS_JobName) Then
                m_jobMatch = True
                Exit For
            End If
        Next i
    Else
        ''''bypass stage check and set m_jobMatch is True
        m_jobMatch = True
    End If

    ''''<Important> default: all Config Security Code bits = 0
    For j = 0 To (m_bitwidth - 1)
        Array1(StartBit + j) = 0
    Next j
        
    If (m_jobMatch = True And Flag1 <> "ALL_0") Then
        ''''20170717 update
        For i = 0 To UBound(CFGTable.Category)
            If (Flag1 = UCase(CFGTable.Category(i).pkgName)) Then
                m_size = UBound(CFGTable.Category(i).BitVal_byStage) + 1
                If (m_bitwidth = m_size) Then
                    For j = 0 To UBound(CFGTable.Category(i).BitVal_byStage)
                        Array1(StartBit + j) = CFGTable.Category(i).BitVal_byStage(j)
                    Next j
                Else
                    TheExec.Datalog.WriteComment "<WARNING> UnMatch Length between BitWidth (" + CStr(m_bitwidth) + " and CFGTable_BitValue size (" + CStr(m_size) + ") !!!"
                End If
                Exit For
            End If
        Next i
    Else
        ''''default: all bits = 0 (if NOT match job)
        ''''doNothing
    End If

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function auto_BKM2Fuse_Mapping(BKM As String) As String
On Error GoTo errHandler
    Dim funcName As String:: funcName = "auto_BKM2Fuse_Mapping"
    
    
    If (Dic_BKM.Exists(BKM)) Then
        auto_BKM2Fuse_Mapping = Dic_BKM(BKM)
    End If
    
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



