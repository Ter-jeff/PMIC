Attribute VB_Name = "VBT_LIB_DC_AP"
Option Explicit

Public Function IDS_eFuse_Write(FuseType As String, Fuse_StoreName As String, Flag_Name As String) As Long

    Dim i As Long
    
    Dim k As Integer
    Dim m_len As Long

    Dim m_resolution As Double
    
    Dim showPrint As Boolean: showPrint = True
    
    Dim site As Variant
    
    Dim Data_Temp As String
    Dim m_valStr As String
    Dim m_dlogstr As String
    Dim m_Fusetype As String
    Dim m_catename As String
    Dim EFUSE_Field_Ary() As String
    Dim Fuse_StoreName_Ary() As String
    
    Dim m_dbl As New SiteDouble
    Dim m_value As New SiteDouble
    Dim m_decimal As New SiteDouble
    Dim m_dbl_round As New SiteDouble
    
    Dim Pass_Fail_Flag As New SiteBoolean
    
    Dim m_catename_pinlistdata As New PinListData
    
    Dim EFUSE_IDS_Dic As New Scripting.Dictionary
    
    On Error GoTo errHandler
    
    m_len = auto_eFuse_GetCatenameMaxLen(FuseType)
    
    If Fuse_StoreName <> "" Then
        Fuse_StoreName_Ary = Split(Fuse_StoreName, "+")
        For k = 0 To UBound(Fuse_StoreName_Ary)
           If Not EFUSE_IDS_Dic.Exists(Fuse_StoreName_Ary(k)) Then EFUSE_IDS_Dic.Item(Fuse_StoreName_Ary(k)) = gl_IDS_INFO_Dic(Fuse_StoreName_Ary(k))
        Next k
    End If
    ''Example ==>
    ''EFUSE_IDS_Dic.Item("ids_vdd_ave")(0) => "VDD_AVE"
    ''EFUSE_IDS_Dic.Item("ids_vdd_ave")(1) => "pp_scya0_s_fulp_io_popx_nan_daa_dio_allfv_si_ids"
    ''EFUSE_IDS_Dic.Item("ids_vdd_ave")(2) => All_Power_data_IDS_GB.Pins(i)
    ''EFUSE_IDS_Dic.Item("ids_vdd_ave")(3) => "0"
    ''EFUSE_IDS_Dic.Item("ids_vdd_ave")(4) => "271.557"
    If Fuse_StoreName <> "" Then
'        EFUSE_Field_Ary = Split(m_catename, "+")
        For k = 0 To UBound(Fuse_StoreName_Ary)
            If Fuse_StoreName_Ary(k) <> "" Then
                m_catename = Fuse_StoreName_Ary(k)
                m_decimal = EFUSE_IDS_Dic.Item(Fuse_StoreName_Ary(k))(2)
                Call ids_cal_resolution(FuseType, m_catename, m_decimal, m_value, m_resolution)
                
                For Each site In TheExec.sites
                    If TheExec.Flow.SiteFlag(site, Flag_Name) = 1 Then
                        Pass_Fail_Flag(site) = False
                    ElseIf TheExec.Flow.SiteFlag(site, Flag_Name) = 0 Then
                        Pass_Fail_Flag(site) = True
                    Else
                        Pass_Fail_Flag(site) = False
                        TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
                    End If
'                    Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catename, Pass_Fail_Flag(site), False)
'                    Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, m_value(site), False)
                    
                    If (showPrint) Then
                       m_Fusetype = FuseType
                       m_valStr = Format(m_decimal * 1000, "0.000000")
                       m_Fusetype = FormatNumeric(m_Fusetype, 4)
                       m_Fusetype = m_Fusetype + FormatNumeric("Fuse IDS_SetWriteDecimal_SetPatTestPass_Flag ", -1)
                       m_dlogstr = vbTab & "Site(" + CStr(site) + ") " + m_Fusetype + FormatNumeric(m_catename, m_len) + " = " + FormatNumeric(m_value, -10) + _
                                    " (" + FormatNumeric(m_valStr + " mA", 12) + _
                                    " / " + Format(m_resolution * 1000, "0.000000") + "mA)"
                       TheExec.Datalog.WriteComment m_dlogstr
                    End If
                Next site
                Call auto_eFuse_SetPatTestPass_Flag_SiteAware(FuseType, m_catename, Pass_Fail_Flag, False)
                Call auto_eFuse_SetWriteVariable_SiteAware(FuseType, m_catename, m_value, False)
                
            End If
        Next k
    End If
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in IDS_eFuse_Write_"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function DCVS_IDS_main_current_Delta(Delta_Pin As PinList, FuseType As String)
    
    Dim i As Long, j As Long

    Dim IDS_from_Efuse As New SiteDouble
    Dim IDS_from_DCVS As New SiteDouble
    Dim IDS_Delta As New SiteDouble
    Dim IDS_PwrName As String
    Dim HiLimit_IDS_Delta As Double
    Dim LoLimit_IDS_Delta As Double

    Set gS_delta_IDS_pcpu = New SiteDouble
    Set gS_delta_IDS_ecpu = New SiteDouble
    Set gS_delta_IDS_gpu = New SiteDouble
    Set gS_delta_IDS_dcs_ddr = New SiteDouble
    Set gS_delta_IDS_cpu_sram = New SiteDouble
    Set gS_delta_IDS_ave = New SiteDouble
    
    On Error GoTo errHandler
    ''Seperate the Delta_Pin list into Array, Carter - 20190115, Start
    Dim Pins() As String, Pin_Cnt As Long
    Dim CFG_Pins() As String
    TheExec.DataManager.DecomposePinList Delta_Pin, Pins(), Pin_Cnt

    ReDim CFG_Pins(UBound(Pins))
    For i = 0 To UBound(Pins())
        If LCase(Pins(i)) Like "*sram_cpu" Then
            CFG_Pins(i) = "ids_" & LCase(Replace(Pins(i), "SRAM_CPU", "CPU_SRAM"))
        Else
            CFG_Pins(i) = "ids_" & LCase(Pins(i))
        End If
    Next i
    
    If TheExec.TesterMode = testModeOffline Then
        For i = 0 To UBound(Pins())
            IDS_from_DCVS = All_Power_data_IDS_GB.Pins(Pins(i))
            IDS_from_Efuse = IDS_Delta.Add(0.01 + Rnd() * 0.0001)
            IDS_Delta = IDS_from_DCVS.Subtract(IDS_from_Efuse)
            
            TheExec.Flow.TestLimit resultVal:=IDS_from_Efuse, lowVal:=0.0001, Tname:=Pins(i) & "_CP1", PinName:=Pins(i) & "_CP1"
            TheExec.Flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=IDS_PwrName & "_WLFT", PinName:=Pins(i) & "_WLFT"
            TheExec.Flow.TestLimit IDS_Delta, Tname:=Pins(i) & "_Delta", PinName:=Pins(i) & "_Delta", ForceResults:=tlForceNone
        Next i

    Else
        If FuseType = "CFG" Then
            For i = 0 To UBound(CFGFuse.Category())
                If LCase(CFGFuse.Category(i).algorithm) Like "*ids*" Then
                    IDS_PwrName = LCase(CFGFuse.Category(i).Name)
                    ''Do the IDS_Delta if the IDS_PwrName exists in Delta_Pin and Delta_IDS_Dic exists, Carter - 20190116, Start
                    For j = 0 To UBound(Pins())
                        If IDS_PwrName = CFG_Pins(j) Then
                            IDS_from_DCVS = All_Power_data_IDS_GB.Pins(Pins(j))
                            IDS_from_Efuse = CFGFuse.Category(i).Read.Decimal.Multiply(CFGFuse.Category(i).Resoultion * 0.001)
                            IDS_Delta = IDS_from_DCVS.Subtract(IDS_from_Efuse)
                            If IDS_PwrName Like "*pcpu*" Then
                                gS_delta_IDS_pcpu = IDS_Delta
                            ElseIf IDS_PwrName Like "*ecpu*" Then
                                gS_delta_IDS_ecpu = IDS_Delta
                            ElseIf IDS_PwrName Like "*gpu*" Then
                                gS_delta_IDS_gpu = IDS_Delta
                            ElseIf IDS_PwrName Like "*dcs_ddr*" Then
                                gS_delta_IDS_dcs_ddr = IDS_Delta
                            ElseIf IDS_PwrName Like "*cpu_sram*" Then
                                 gS_delta_IDS_cpu_sram = IDS_Delta
                            ElseIf IDS_PwrName Like "*ave*" Then
                                gS_delta_IDS_ave = IDS_Delta
                            End If
                            TheExec.Flow.TestLimit resultVal:=IDS_from_Efuse, lowVal:=0.0001, Tname:=IDS_PwrName & "_CP1", PinName:=IDS_PwrName & "_CP1"
                            TheExec.Flow.TestLimit resultVal:=IDS_from_DCVS, Tname:=IDS_PwrName & "_WLFT", PinName:=IDS_PwrName & "_WLFT"
                            TheExec.Flow.TestLimit IDS_Delta, Tname:=IDS_PwrName & "_Delta", PinName:=IDS_PwrName & "_Delta", ForceResults:=tlForceFlow
                            Exit For
                        End If
                    Next j
                    ''Do the IDS_Delta if the IDS_PwrName exists in Delta_Pin and Delta_IDS_Dic exists, Carter - 20190116, End
                End If
            Next i
        End If
    End If

'============================================================================================
'=  Record Delta IDS to HardKeyReg (added on 2017/7/10)                                     =
'============================================================================================
    VBT_IEDA_Registry "WLFT_Delta_IDS_PCPU", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_ECPU", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_GPU", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_DCS_DDR", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_CPU_SRAM", True
    VBT_IEDA_Registry "WLFT_Delta_IDS_AVE", True

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in DCVS_IDS_main_current_Delta"
    If AbortTest Then Exit Function Else Resume Next
End Function

