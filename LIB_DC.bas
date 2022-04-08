Attribute VB_Name = "LIB_DC"
Option Explicit

Public Function IDS_Store2Dic(Fuse_StoreName As String, Core_Power_Pin As String, All_Power_data As PinListData, patt As Pattern)
    
    On Error GoTo errHandler
    
    Dim i As Integer
    Dim Pin_Cnt As Long
    
    Dim Val_Hi() As String
    Dim Val_Lo() As String
    Dim m_dlogstr As String
    Dim Core_PinAry() As String
    Dim Fuse_StoreName_Ary() As String
    
    Dim Core_Power_data As New PinListData
    
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    ' if no Use-Limits on this test, FlowLimitsInfo is nothing
    If FlowLimitsInfo Is Nothing Then
        TheExec.AddOutput "Could not get the limits info", vbRed, True
        Exit Function
    End If
    
    FlowLimitsInfo.GetLowLimits Val_Lo
    FlowLimitsInfo.GetHighLimits Val_Hi
    
    TheExec.DataManager.DecomposePinList Core_Power_Pin, Core_PinAry, Pin_Cnt
    For i = 0 To Pin_Cnt - 1
        Core_Power_data.AddPin (UCase(Core_PinAry(i)))
        Core_Power_data.Pins(UCase(Core_PinAry(i))) = All_Power_data.Pins(UCase(Core_PinAry(i)))
    Next
    
    Fuse_StoreName_Ary = Split(Fuse_StoreName, "+")
    ReDim ids_info_ary(UBound(Fuse_StoreName_Ary))
'    TheExec.Datalog.WriteComment "-----Start to save IDS value-----"
    For i = 0 To UBound(Fuse_StoreName_Ary)
        If Fuse_StoreName_Ary(i) <> "" Then
            ids_info_ary(i).Pat = patt
            ids_info_ary(i).LoLimit = Val_Lo(i)
            ids_info_ary(i).HiLimit = Val_Hi(i)
            ids_info_ary(i).Pin = Core_Power_data.Pins(i).Name
            ids_info_ary(i).MeasureValue = Core_Power_data.Pins(i)
            
            ''Example ==> gl_IDS_INFO_Dic.Add "ids_vdd_ave", Array(Pin, PAT, MeasureValue, LoLimit, HiLimit)
            ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(0) => "VDD_AVE"
            ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(1) => "pp_scya0_s_fulp_io_popx_nan_daa_dio_allfv_si_ids"
            ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(2) => All_Power_data_IDS_GB.Pins(i)
            ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(3) => "0"
            ''gl_IDS_INFO_Dic.Item("ids_vdd_ave")(4) => "271.557"
            If Not gl_IDS_INFO_Dic.Exists(Fuse_StoreName_Ary(i)) Then
                gl_IDS_INFO_Dic.Add Fuse_StoreName_Ary(i), _
                    Array(ids_info_ary(i).Pin, ids_info_ary(i).Pat, ids_info_ary(i).MeasureValue, ids_info_ary(i).LoLimit, ids_info_ary(i).HiLimit)
                    
'                m_dlogstr = vbTab & "Saving IDS and Fuse Field Name: " + ids_info_ary(i).Pin + "(" + Fuse_StoreName_Ary(i) + ")"
'                TheExec.Datalog.WriteComment m_dlogstr
            End If
        End If
    Next i

'    TheExec.Datalog.WriteComment "-----End to save IDS value-----"
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "Error in IDS_Store2Dic"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ids_cal_resolution(FuseType As String, m_catename As String, ids_val As SiteDouble, ids_output As SiteDouble, resolution_output As Double, Optional showPrint = False)
    
    Dim funcName As String:: funcName = "ids_cal_resolution" ''"auto_calc_IDS_Decimal"
    
    Dim m_str As String
    
    Dim ids_bool As New SiteBoolean
    
    Dim fuse_ids_resolution As Double
  
    Dim ids_max As New SiteDouble
    Dim IDS_resolution As New SiteDouble
    Dim ids_resolution_round As New SiteDouble
    
    On Error GoTo errHandler
       
    ids_max = ids_val.Maximum(0)
    fuse_ids_resolution = auto_eFuse_GetIDSResolution(FuseType, m_catename)
    fuse_ids_resolution = fuse_ids_resolution * 0.001 ''''<NOTICE> update unit from mA to A
    IDS_resolution = ids_max.Divide(fuse_ids_resolution)
    ids_resolution_round = IDS_resolution.Add(0.5).Truncate
    
    ids_bool = ids_resolution_round.Subtract(IDS_resolution).compare(GreaterThan, 0)
''-------Substitute the code in the belowing--------
    If ids_bool.Any(True) Then
        TheExec.sites.Selected = ids_bool
        ids_output = ids_resolution_round
        TheExec.sites.Selected = True
    End If
    
    If ids_bool.Any(False) Then
        TheExec.sites.Selected = ids_bool.LogicalNot
        ids_output = ids_resolution_round.Add(1)
        TheExec.sites.Selected = True
    End If

    ids_output = ids_output.Maximum(0)
       
    If (showPrint) Then
        For Each site In TheExec.sites.Active
            m_str = CStr(ids_val(site)) + "/" + CStr(fuse_ids_resolution)
            Debug.Print funcName + "......" + m_str + " = ids_resolution = " + CStr(IDS_resolution(site))
            Debug.Print funcName + "......" + m_str + " = ids_resolution_round = " + CStr(ids_resolution_round(site))
            Debug.Print funcName + "......" + m_str + " = " + CStr(ids_output(site))
            Debug.Print "---------------------------------------------------------------------" + vbCrLf
        Next site
    End If
    
    resolution_output = fuse_ids_resolution
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error in ids_cal_resolution"
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function Find_Range_WaitTime(Range_List() As Double, Instrucment_Type As String) As Variant

    Dim funcName As String:: funcName = "Find_Range_WaitTime"
    
    Dim var As Variant
    Dim SattleTime() As Double
    
    ReDim SattleTime(UBound(Range_List))
    
    If LCase(Instrucment_Type) = "hexvs" Then
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.01 Then
                SattleTime(var) = 100 * ms
            
            ElseIf Range_List(var) = 0.1 Then
                SattleTime(var) = 10 * ms
                
            ElseIf Range_List(var) = 1 Then
                SattleTime(var) = 1 * ms
            
            ElseIf Range_List(var) >= 15 Then
                SattleTime(var) = 100 * us
                
            End If
        Next var
        
    ElseIf LCase(Instrucment_Type) = "vhdvs" Then
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.000004 Then
                SattleTime(var) = 18 * ms
            
            ElseIf Range_List(var) = 0.00002 Then
                SattleTime(var) = 4 * ms
                
            ElseIf Range_List(var) = 0.0002 Then
                SattleTime(var) = 4 * ms
            
            ElseIf Range_List(var) = 0.002 Then
                SattleTime(var) = 3.5 * ms
                
            ElseIf Range_List(var) = 0.02 Then
                SattleTime(var) = 540 * us
                
            ElseIf Range_List(var) = 0.04 Then
                SattleTime(var) = 260 * us
                
            ElseIf Range_List(var) = 0.2 Then
                SattleTime(var) = 210 * us
            
            ElseIf Range_List(var) = 0.4 Then
                SattleTime(var) = 90 * us
            
            ElseIf Range_List(var) = 0.7 Then
                SattleTime(var) = 100 * us
                
            ElseIf Range_List(var) = 0.8 Then
                SattleTime(var) = 100 * us
                
            ElseIf Range_List(var) = 1.4 Then
                SattleTime(var) = 50 * us
                
            ElseIf Range_List(var) = 2.8 Then
                SattleTime(var) = 45 * us
                
            ElseIf Range_List(var) = 5.6 Then
                SattleTime(var) = 30 * us
                
            End If
        Next var
        
        
    ElseIf LCase(Instrucment_Type) = "dc-07" Then
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.000002 Or Range_List(var) = 0.000004 Then
                SattleTime(var) = 6 * ms
            
            ElseIf Range_List(var) = 0.00002 Or Range_List(var) = 0.00004 Then
                SattleTime(var) = 1.5 * ms
                
            ElseIf Range_List(var) = 0.0002 Or Range_List(var) = 0.0004 Then
                SattleTime(var) = 1.4 * ms
            
            ElseIf Range_List(var) = 0.002 Or Range_List(var) = 0.004 Then
                SattleTime(var) = 11 * ms
                
            ElseIf Range_List(var) = 0.02 Or Range_List(var) = 0.04 Then
                SattleTime(var) = 1.5 * ms
            
            ElseIf Range_List(var) = 0.2 Or Range_List(var) = 0.4 Then
                SattleTime(var) = 260 * us
            
            ElseIf Range_List(var) >= 1 Then
                SattleTime(var) = 1.6 * ms
                
            End If
        Next var
    End If
    
    Find_Range_WaitTime = SattleTime
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Find_Range_Accuracy(Range_List() As Double, Instrucment_Type As String) As Variant

    Dim funcName As String:: funcName = "Find_Range_WaitTime"
    
    Dim var As Variant
    Dim Accuracy() As Double
    
    ReDim Accuracy(UBound(Range_List))
    
    If LCase(Instrucment_Type) = "hexvs" Then
        
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.01 Then
                
                Accuracy(var) = Range_List(var) * 0.01 + 0.00005
            ElseIf Range_List(var) = 0.1 Then
                
                Accuracy(var) = Range_List(var) * 0.01 + 0.0005
            ElseIf Range_List(var) = 1 Then
                
                Accuracy(var) = Range_List(var) * 0.01 + 0.005
            ElseIf Range_List(var) >= 15 Then
                
                Accuracy(var) = Range_List(var) * 0.01 + (Range_List(var) \ 15) * 0.05
            End If
        Next var
    ElseIf LCase(Instrucment_Type) = "vhdvs" Then
        
        For var = 0 To UBound(Range_List)
            
            If Range_List(var) = 0.000004 Then
                
                Accuracy(var) = Range_List(var) * 0.005 + 0.000000026
            ElseIf Range_List(var) = 0.00002 Then
                
                Accuracy(var) = Range_List(var) * 0.005 + 0.00000012
            ElseIf Range_List(var) = 0.0002 Then
                
                Accuracy(var) = Range_List(var) * 0.005 + 0.0000012
            ElseIf Range_List(var) = 0.002 Then
                
                Accuracy(var) = Range_List(var) * 0.005 + 0.000012
            ElseIf Range_List(var) = 0.02 Then
                
                Accuracy(var) = Range_List(var) * 0.005 + 0.00012
            ElseIf Range_List(var) = 0.04 Then
                
                Accuracy(var) = Range_List(var) * 0.007 + 0.00024
            ElseIf Range_List(var) = 0.2 Then
                
                Accuracy(var) = Range_List(var) * 0.005 + 0.0012
            ElseIf Range_List(var) = 0.4 Then
                
                Accuracy(var) = Range_List(var) * 0.007 + 0.0024
            ElseIf Range_List(var) = 0.7 Then
                
                Accuracy(var) = Range_List(var) * 0.005 + 0.006
            ElseIf Range_List(var) = 0.8 Then
                
                Accuracy(var) = Range_List(var) * 0.005 + 0.006
            ElseIf Range_List(var) = 1.4 Then
                
                Accuracy(var) = Range_List(var) * 0.007 + 0.01
            ElseIf Range_List(var) = 2.8 Then
                
                Accuracy(var) = Range_List(var) * 0.007 + 0.02
            ElseIf Range_List(var) = 5.6 Then
                
                Accuracy(var) = Range_List(var) * 0.007 + 0.04
            End If
        Next var
    ElseIf LCase(Instrucment_Type) = "dc-07" Then
        
        For var = 0 To UBound(Range_List)
            If Range_List(var) = 0.00002 Or Range_List(var) = 0.00004 Then
                
                Accuracy(var) = Range_List(var) * 0.002 + 0.000000075
            ElseIf Range_List(var) = 0.0002 Or Range_List(var) = 0.0004 Then
                
                Accuracy(var) = Range_List(var) * 0.002 + 0.0000004
            ElseIf Range_List(var) = 0.002 Or Range_List(var) = 0.004 Then
                
                Accuracy(var) = Range_List(var) * 0.002 + 0.000004
            ElseIf Range_List(var) = 0.02 Or Range_List(var) = 0.04 Then
                
                Accuracy(var) = Range_List(var) * 0.002 + 0.00004
            ElseIf Range_List(var) = 0.2 Or Range_List(var) = 0.4 Then
                
                Accuracy(var) = Range_List(var) * 0.002 + 0.0004
            ElseIf Range_List(var) >= 1 Then
                
                Accuracy(var) = Range_List(var) * 0.002 + 0.008
            End If
        Next var
    End If
    
    Find_Range_Accuracy = Accuracy
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
