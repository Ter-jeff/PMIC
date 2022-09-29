Attribute VB_Name = "LIB_Digital_Shmoo_Setup"
'T-AutoGen-Version:OTC Automation/Validation - Version: 2.23.66.70/ with Build Version - 2.23.66.70
'Test Plan:E:\Raze\SERA\Sera_A0_TestPlan_190314_70544_3.xlsx, MD5=e7575c200756d6a60c3b3f0138121179
'SCGH:Skip SCGH file
'Pattern List:Skip Pattern List Csv
'SettingFolder:E:\ADC3\02.Development\02.Coding\Source\Automation\Automation\bin\Debug
'VBT is not using Central-[Warning] :Z:\Teradyne\ADC\L\Log\LibBas_PMIC:User specified a personal VBT library folder, should not use for Production T/P!
Option Explicit

Public CHAR_USL_HVCC As Double
Public CHAR_USL_LVCC As Double
Public CHAR_LSL_HVCC As Double
Public CHAR_LSL_LVCC As Double
Public Charz_Power_condition As String
Public Charz_Force_Power_condition As String
Public CharSetName_GLB As String
Public Re_store   As Double
Public g_PPMU_Connected As String
Public PrePat_Restore_String As String
Public PreMeas_Restore_String As String
Public PrePatStore As Boolean
Public PreMeasStore As Boolean

'================================================================================
'180425 update for trace compensation
Public gTerm_cond_All As String
Public gTerm_Restore_cond As String
Public gTerm_cond_Flag As Boolean
Public gTerm_cond_shm_Flag As Boolean
'================================================================================

Public Function Force_Condition_V(force_pin As String, force_val As Double)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Force_Condition_V"

    Dim p_ary() As String, p_cnt As Long
    p_ary = Split(force_pin, ",")    ' added to allow pinlist with comma such as "VDD_CPU,VDD_SRAM"
    If (LCase(TheExec.DataManager.pintype(LCase(p_ary(0)))) Like "power") Then
        SetPowerValue force_pin, force_val
    Else
        TheHdw.Digital.Pins(force_pin).Disconnect
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = force_pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Voltage, "0.000000"))
            Else
                PrePat_Restore_String = PrePat_Restore_String & ";" & force_pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Voltage, "0.000000"))
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = force_pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Voltage, "0.000000"))
            Else
                PreMeas_Restore_String = PreMeas_Restore_String & ";" & force_pin & ":V:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Voltage, "0.000000"))
            End If
        End If

        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = force_pin + ":DisConnectPPMU;" + force_pin + ":ConnectDigital"
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + force_pin + ":DisConnectPPMU;" + force_pin + ":ConnectDigital"
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = force_pin + ":DisConnectPPMU;" + force_pin + ":ConnectDigital"
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + force_pin + ":DisConnectPPMU;" + force_pin + ":ConnectDigital"
            End If
        End If

        With TheHdw.PPMU.Pins(force_pin)
            '.mode = tlPPMUForceVMeasureI
            .ForceV (force_val), 0.02
            .Connect
            .Gate = tlOn
        End With

        If g_PPMU_Connected <> "" Then
            g_PPMU_Connected = g_PPMU_Connected & "," & force_pin
        Else
            g_PPMU_Connected = force_pin
        End If
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'20170419 add for case I
Public Function Force_Condition_I(force_pin As String, force_val As Double)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Force_Condition_I"

    '20190416 top
    If (LCase(TheExec.DataManager.pintype(LCase(force_pin))) Like "power") Then
        SetPowerValue_I force_pin, force_val
    ElseIf (LCase(TheExec.DataManager.pintype(LCase(force_pin))) Like "analog") Then    'Update 0609
        SetPowerValue_I force_pin, force_val
        '20190416 end
    Else
        TheHdw.Digital.Pins(force_pin).Disconnect
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = force_pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Current, "0.000000"))
            Else
                PrePat_Restore_String = PrePat_Restore_String & ";" & force_pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Current, "0.000000"))
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = force_pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Current, "0.000000"))
            Else
                PreMeas_Restore_String = PreMeas_Restore_String & ";" & force_pin & ":I:" & CStr(Format(TheHdw.PPMU.Pins(force_pin).Current, "0.000000"))
            End If
        End If

        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = force_pin + ":DisConnectPPMU"
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + force_pin + ":DisConnectPPMU"
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = force_pin + ":DisConnectPPMU"
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + force_pin + ":DisConnectPPMU"
            End If
        End If

        With TheHdw.PPMU.Pins(force_pin)
            '.mode = tlPPMUForceIMeasureV
            .ForceI (force_val)
            .Connect
            .Gate = tlOn
        End With

        If g_PPMU_Connected <> "" Then
            g_PPMU_Connected = g_PPMU_Connected & "," & force_pin
        Else
            g_PPMU_Connected = force_pin
        End If
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function





'Public Function Get_Shmoo_ApplyPin(Pin_Ary() As String, Pin_Cnt As Long)
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "Get_Shmoo_ApplyPin"
'
'    Dim Shmoo_Pin_Str As String
'    Dim Shmoo_Tracking_Item As Variant, shmoo_axis As Variant
'    Dim DevChar_Setup As String
'    Shmoo_Pin_Str = ""
'    DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
'    If TheExec.DevChar.Setups.IsRunning = True Then
'        With TheExec.DevChar.Setups(DevChar_Setup)
'            For Each shmoo_axis In .Shmoo.Axes.List
'                If Shmoo_Pin_Str <> "" Then
'                    Shmoo_Pin_Str = Shmoo_Pin_Str & "," & TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).ApplyTo.Pins
'                Else
'                    Shmoo_Pin_Str = TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).ApplyTo.Pins
'                End If
'                With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).TrackingParameters
'                    For Each Shmoo_Tracking_Item In .List
'                        Shmoo_Pin_Str = Shmoo_Pin_Str & "," & .Item(Shmoo_Tracking_Item).ApplyTo.Pins
'                    Next Shmoo_Tracking_Item
'                End With
'            Next shmoo_axis
'        End With
'    End If
'    TheExec.DataManager.DecomposePinList Shmoo_Pin_Str, Pin_Ary, Pin_Cnt
'
'    Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

Public Function SetForceCondition(Setup_string As String) As Long

    Dim Cat_temp() As String
    Dim cat_split_temp As Variant

    Dim Pin_info_temp() As String

    Dim Edge_and_TimeSet() As String    '20180702 add

    Dim PinName As String, Pin_Type As String, pin_value As String, pin_restore As Boolean
    Dim SetSite, Split_cnt As Integer
    Dim ins_name  As String
    Dim Site      As Variant

    g_Retention_ForceV = ""
    g_Retention_VDD = ""
    'CharSetName_GLB = "" ''avoid VBT error if HIP universal(Meas_FreqVoltCurr_Universal_func) wants to do the shmoo by opcode "test"
    '============================customer===================================================================
    g_PPMU_Connected = ""

    '===============================================================================================
    On Error GoTo err1
    Dim funcName  As String:: funcName = "Edge_and_TimeSet"
    
    'get instance name
    ins_name = TheExec.DataManager.InstanceName
    '===============================================================================================
    '180425 update for trace compensation
    Setup_string = LCase(Replace(Setup_string, " ", ""))
    '    Analyze_shmoo_setup ' for trace compensation
    Call Add_Term_Restore(Setup_string)  ' for trace compensation
    If (Setup_string = "restorepremeas") Then
        Setup_string = LCase(PreMeas_Restore_String)
        PreMeas_Restore_String = ""
        If (Setup_string <> "") Then TheExec.Datalog.WriteComment "restore premeas Force Condtion:" & Setup_string
    ElseIf (Setup_string = "restoreprepat_term") Then
        Setup_string = LCase(PrePat_Restore_String) & ":term"    'remain keyword "term" in order to impact Add_Term_Resotore function
        PrePat_Restore_String = ""
        If (Setup_string <> "") Then TheExec.Datalog.WriteComment "restore prepat Force Condtion:" & Setup_string
    ElseIf (Setup_string = "restoreprepat") Then
        Setup_string = LCase(PrePat_Restore_String)
        PrePat_Restore_String = ""
        If (Setup_string <> "") Then TheExec.Datalog.WriteComment "restore prepat Force Condtion:" & Setup_string

    End If


    If (UCase(Setup_string) Like "*STOREPREMEAS") Then
        PreMeasStore = True
        PreMeas_Restore_String = ""
        Setup_string = Replace(UCase(Setup_string), ";STOREPREMEAS", "")
        If TheExec.DevChar.Setups.IsRunning = True Then    ' 20180702 add
            Charz_Force_Power_condition = Setup_string
        End If
    ElseIf (UCase(Setup_string) Like "*STOREPREPAT") Then
        PrePatStore = True
        PrePat_Restore_String = ""
        Setup_string = Replace(UCase(Setup_string), ";STOREPREPAT", "")
        If TheExec.DevChar.Setups.IsRunning = True Then    ' 20180702 add
            Charz_Force_Power_condition = Setup_string
        End If
    End If
    '===============================================================================================

    If (Setup_string = "") Then Exit Function
    'get site Number
    For Each Site In TheExec.Sites.Active
        SetSite = Site
    Next Site

    Cat_temp = Split(Setup_string, ";")    'compatible with Autogen and used in central

    TheExec.Datalog.WriteComment "Force Condtion:" & Setup_string


    ' Setup force condition to global
    ' Charz_Force_Power_condition = Setup_string  ' 20180702 mask

    '    Get_Shmoo_ApplyPin Shmoo_Pin_ary, Shmoo_Pin_Cnt
    Dim flag_shmoo_set_current_point As Boolean
    flag_shmoo_set_current_point = True
    For Each cat_split_temp In Cat_temp

        If cat_split_temp = "" Then GoTo continue1

        pin_restore = False

        If (InStr(LCase(cat_split_temp), "restore") > 0) Then
            pin_restore = True
        End If

        Pin_info_temp = Split(cat_split_temp, ":")
        Split_cnt = UBound(Pin_info_temp) + 1


        '\\\\\\\\\\\\\\\SAVE H/L Limit\\\\\\\\\\\\\
        If ((LCase(cat_split_temp) Like "usl*") Or LCase(cat_split_temp) Like "lsl*") Then
            If (Split_cnt = 1) Then
                GoTo continue1
            ElseIf (Split_cnt = 2) Then
                '\\\\\\Save HVCC Limit\\\\\\
                If (Pin_info_temp(0) = "USL") Then
                    If (Pin_info_temp(1) = "") Then
                        CHAR_USL_HVCC = 9999
                    Else
                        CHAR_USL_HVCC = CDbl(Pin_info_temp(1))
                        CHAR_USL_LVCC = CDbl(Pin_info_temp(1))
                    End If
                    GoTo continue1
                End If
                '\\\\\\Save LVCC Limit\\\\\\
                If (Pin_info_temp(0) = "LSL") Then
                    If (Pin_info_temp(1) = "") Then
                        CHAR_LSL_LVCC = 9999
                    Else
                        CHAR_LSL_HVCC = CDbl(Pin_info_temp(1))
                        CHAR_LSL_LVCC = CDbl(Pin_info_temp(1))
                    End If
                End If
            End If
            GoTo continue1
        End If

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '\\\\\\Read CharSetName_GLB\\\\\\
        If (InStr(LCase(cat_split_temp), "charsetname") > 0) Then
            CharSetName_GLB = Pin_info_temp(1)
            GoTo continue1
        End If

        Dim InStrTmp As String
        '''        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\Original\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        '''        If (UBound(Pin_info_temp) >= 2) Then
        '''            PinName = Pin_info_temp(0)
        '''            pin_type = Pin_info_temp(1)
        '''            'pin_value = CStr(Pin_info_temp(2))
        '''            InStrTmp = Pin_info_temp(2)
        '''
        '''            ''0421-Roger & Roy
        '''            If InStr(InStrTmp, "_") > 0 Or InStr(InStrTmp, "+") > 0 Or InStr(InStrTmp, "-") > 0 Or InStr(InStrTmp, "*") > 0 Or InStr(InStrTmp, "/") > 0 Then ' can not evaluate if only with  single number
        '''                For Each Site In TheExec.Sites.Active
        '''                    pin_value = CStr(Evaluate(InStrTmp))
        '''                Next Site
        '''            Else
        '''                pin_value = CStr(InStrTmp)
        '''            End If
        '''        ElseIf (UBound(Pin_info_temp) = 1) Then
        '''            PinName = Pin_info_temp(0)
        '''            If InStr(InStrTmp, "_") > 0 Or InStr(InStrTmp, "+") > 0 Or InStr(InStrTmp, "-") > 0 Or InStr(InStrTmp, "*") > 0 Or InStr(InStrTmp, "/") > 0 Then ' can not evaluate if only with  single number
        '''                For Each Site In TheExec.Sites.Active
        '''                    pin_value = CStr(Evaluate(Pin_info_temp(1)))
        '''                Next Site
        '''            Else
        '''                pin_value = CStr((Pin_info_temp(1)))
        '''            End If
        '''            pin_type = ""
        '''        End If
        '''        '////////////////////////////////////////////////////////////////////////////////////////////////

        '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\20161229 Roy Modified for  Evaluate\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
        If (UBound(Pin_info_temp) >= 2) Then
            PinName = Pin_info_temp(0)
            Pin_Type = Pin_info_temp(1)
            pin_value = CStr(Spec_Evaluate_DC_for_flow_loop(Pin_info_temp(0), Pin_info_temp(1), Pin_info_temp(2)))
        ElseIf (UBound(Pin_info_temp) = 1) Then
            PinName = Pin_info_temp(0)
            Pin_Type = ""
            pin_value = CStr(Pin_info_temp(1))
        End If
        '////////////////////////////////////////////////////////////////////////////////////////////////


        If (LCase(Pin_Type) = "setupfv") Then
            'PAD_MTR_ANALOG_TEST_P:SetupFV:Vprog,Irange,CustomizeWaitTime
            Call SetupDCVI_ForceV(Pin_info_temp(0), Pin_info_temp(2))
            GoTo continue1
        End If

        If (LCase(Pin_Type) = "setupfi") Then
            'PAD_MTR_ANALOG_TEST_P:SetupFI:Vprog,Iprog
            ' Mode Alarm: voltage above Vprog
            ' Voltage Clamp Alarm: voltage above Vprog+960mV
            Call SetupDCVI_ForceI(Pin_info_temp(0), Pin_info_temp(2))
            GoTo continue1
        End If

        If (LCase(Pin_Type) = "restoredcvi") Then
            With TheHdw.DCVI.Pins(Pin_info_temp(0))
                .Gate = False
                .Disconnect
            End With
            GoTo continue1
        End If
        ''===============================================================================================
        ''170425 update for trace compensation
        If (LCase(Pin_Type) = "term") Then
            '            Call Trace_res_Compensation(CStr(cat_split_temp), flag_shmoo_set_current_point)
            GoTo continue1
        End If
        ''        If (pin_type = "term") Then Print
        ''            'process later"?
        '''            Call Trace_Compensation(CStr(cat_split_temp), flag_shmoo_set_current_point)?
        ''            GoTo continue1
        ''        End If
        ''===============================================================================================
        '///////////////////////////////// Case I ///////////////////////////////////////////////////////
        If (UCase(Pin_Type) = "I") And pin_value <> "" Then
            Force_Condition_I PinName, CDbl(pin_value)
            GoTo continue1
        End If
        '///////////////////////////////// Case V* ///////////////////////////////////////////////////////
        If (UCase(Pin_Type) = "V") And pin_value <> "" Then
            Force_Condition_V PinName, CDbl(pin_value)
            GoTo continue1
        End If
        Dim i     As Long
        'Modify for force condition "VRET" 20171213
        If (UCase(Pin_Type) = "VRET") And pin_value <> "" Then
            Dim rp_ary() As String, rp_cnt As Long, pn As String, p_val As String
            TheExec.DataManager.DecomposePinList PinName, rp_ary, rp_cnt
            pn = Join(rp_ary, ",")
            p_val = pin_value
            If UBound(rp_ary) >= 1 Then For i = 1 To UBound(rp_ary): p_val = p_val & "," & pin_value: Next i
        If g_Retention_VDD = "" Then
            g_Retention_VDD = pn
            g_Retention_ForceV = p_val
        Else
            g_Retention_VDD = g_Retention_VDD & "," & pn
            g_Retention_ForceV = g_Retention_ForceV & "," & p_val
        End If
        GoTo continue1
    End If
    If (UCase(Pin_Type) = "VID") And pin_value <> "" Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":VID:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVid)), "0.000")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VID:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVid)), "0.000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":VID:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVid)), "0.000")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VID:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVid)), "0.000")
            End If
        End If
        TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVid) = CDbl(pin_value)
        GoTo continue1
    End If
    '////////
    If (UCase(Pin_Type) = "VOD") And pin_value <> "" Then
        If (LCase(TheExec.DataManager.pintype(PinName)) <> "differential") Then
            TheExec.AddOutput "[Alarm] Type: VOD ,Pin: " & PinName & " is not Differential Pin"
            GoTo continue1
        End If
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":VOD:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVod)), "0.000")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VOD:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVod)), "0.000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":VOD:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVod)), "0.000")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VOD:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVod)), "0.000")
            End If
        End If
        TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVod) = CDbl(pin_value)
        GoTo continue1
    End If
    '/////////
    If (UCase(Pin_Type) = "VICM") And pin_value <> "" Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":VICM:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVicm)), "0.000")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VICM:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVicm)), "0.000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":VICM:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVicm)), "0.000")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VICM:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVicm)), "0.000")
            End If
        End If
        TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chVicm) = CDbl(pin_value)
        GoTo continue1
    End If
    '/////////
    If (UCase(Pin_Type) = "VIH") And pin_value <> "" Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":VIH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVih)), "0.000")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VIH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVih)), "0.000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":VIH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVih)), "0.000")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VIH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVih)), "0.000")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Levels.Value(chVih) = CDbl(pin_value)
        GoTo continue1
    End If
    '/////////
    If (UCase(Pin_Type) = "VIL") And pin_value <> "" Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":VIL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVil)), "0.000")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VIL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVil)), "0.000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":VIL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVil)), "0.000")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VIL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVil)), "0.000")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Levels.Value(chVil) = CDbl(pin_value)
        GoTo continue1
    End If
    '/////////
    If (UCase(Pin_Type) = "VOH") And pin_value <> "" Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":VOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVoh)), "0.000")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVoh)), "0.000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":VOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVoh)), "0.000")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VOH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVoh)), "0.000")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Levels.Value(chVoh) = CDbl(pin_value)
        GoTo continue1
    End If
    '/////////
    If (UCase(Pin_Type) = "VOL") And pin_value <> "" Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":VOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVol)), "0.000")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVol)), "0.000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":VOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVol)), "0.000")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VOL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVol)), "0.000")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Levels.Value(chVol) = CDbl(pin_value)
        GoTo continue1
    End If
    '/////////
    If (UCase(Pin_Type) = "VT") And pin_value <> "" And (Not (UCase(TheExec.CurrentChanMap) Like "*FT*" And UCase(PinName) = "DDRIOPINS")) Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                    PrePat_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                Else
                    PrePat_Restore_String = PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                End If
            Else
                If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                End If
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                    PreMeas_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                Else
                    PreMeas_Restore_String = PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                End If
            Else
                If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                End If
            End If
        End If
        TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeVt
        TheHdw.Digital.Pins(PinName).Levels.Value(chVt) = CDbl(pin_value)
        GoTo continue1
    ElseIf (UCase(TheExec.CurrentChanMap) Like "*FT*" And UCase(PinName) = "DDRIOPINS") Then
        PrePatStore = False
        GoTo continue1
    End If
    '/////////
    If (UCase(Pin_Type) = "HIZ") And pin_value <> "" Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                    PrePat_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                Else
                    PrePat_Restore_String = PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                End If
            Else
                If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                End If
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                    PreMeas_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                Else
                    PreMeas_Restore_String = PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                End If
            Else
                If (TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeVt) Then
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":HIZ:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
                End If
            End If
        End If
        TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeLargeHiZ
        TheHdw.Digital.Pins(PinName).Levels.Value(chVt) = CDbl(pin_value)
        GoTo continue1
    End If
    '        If (UCase(pin_type) = "VT") And pin_value <> "" Then
    '           If (PrePatStore = True) Then
    '                If (PrePat_Restore_String = "") Then
    '                    If (LCase(TheExec.DataManager.pintype(PinName)) = "i/o") Then
    '                        PrePat_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
    '                    Else
    '                        PrePat_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chDiff_Vt)), "0.000")
    '                    End If
    '                Else
    '                    If (LCase(TheExec.DataManager.pintype(PinName)) = "i/o") Then
    '                        PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
    '                    Else
    '                        PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chDiff_Vt)), "0.000")
    '                   End If
    '                End If
    '            ElseIf (PreMeasStore = True) Then
    '                If (PreMeas_Restore_String = "") Then
    '                    If (LCase(TheExec.DataManager.pintype(PinName)) = "i/o") Then
    '                        PreMeas_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
    '                    Else
    '                        PreMeas_Restore_String = PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chDiff_Vt)), "0.000")
    '                    End If
    '                Else
    '                    If (LCase(TheExec.DataManager.pintype(PinName)) = "i/o") Then
    '                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVt)), "0.000")
    '                    Else
    '                        PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VT:" + Format(CStr(TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chDiff_Vt)), "0.000")
    '                    End If
    '                End If
    '            End If
    '            If (LCase(TheExec.DataManager.pintype(PinName)) = "i/o") Then
    '                TheHdw.Digital.Pins(PinName).Levels.DriverMode = tlDriverModeVt
    '                TheHdw.Digital.Pins(PinName).Levels.Value(chVt) = CDbl(pin_value)
    '            Else
    '                TheHdw.Digital.Pins(PinName).DifferentialLevels.Value(chDiff_Vt) = CDbl(pin_value)
    '            End If
    '            GoTo continue1
    '        End If
    '/////////
    If (UCase(Pin_Type) = "VCH") And pin_value <> "" Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":VCH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVch)), "0.000")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VCH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVch)), "0.000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":VCH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVch)), "0.000")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VCH:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVch)), "0.000")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Levels.Value(chVch) = CDbl(pin_value)
        GoTo continue1
    End If
    '/////////
    If (UCase(Pin_Type) = "VCL") And pin_value <> "" Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":VCL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVcl)), "0.000")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":VCL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVcl)), "0.000")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":VCL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVcl)), "0.000")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":VCL:" + Format(CStr(TheHdw.Digital.Pins(PinName).Levels.Value(chVcl)), "0.000")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Levels.Value(chVcl) = CDbl(pin_value)
        GoTo continue1
    End If
    '///////// 20180702 add for change D0, D1, D2, D3, R0, R1
    If (UCase(Pin_Type) Like "*D0*") And pin_value <> "" Then
        Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0)), "0.000#########")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0)), "0.000#########")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0)), "0.000#########")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0)), "0.000#########")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD0) = CDbl(pin_value)
        GoTo continue1
    End If

    If (UCase(Pin_Type) Like "*D1*") And pin_value <> "" Then
        Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1)), "0.000#########")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1)), "0.000#########")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1)), "0.000#########")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1)), "0.000#########")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD1) = CDbl(pin_value)
        GoTo continue1
    End If

    If (UCase(Pin_Type) Like "*D2*") And pin_value <> "" Then
        Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2)), "0.000#########")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2)), "0.000#########")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2)), "0.000#########")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2)), "0.000#########")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD2) = CDbl(pin_value)
        GoTo continue1
    End If

    If (UCase(Pin_Type) Like "*D3*") And pin_value <> "" Then
        Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3)), "0.000#########")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3)), "0.000#########")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3)), "0.000#########")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3)), "0.000#########")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD3) = CDbl(pin_value)
        GoTo continue1
    End If

    'SY mask timing without D4 setting
    ' ============================================
    '''        If (UCase(Pin_Type) Like "*D4*") And pin_value <> "" Then
    '''            Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
    '''            If (PrePatStore = True) Then
    '''                If (PrePat_Restore_String = "") Then
    '''                    PrePat_Restore_String = pinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(pinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD4)), "0.000#########")
    '''                Else
    '''                    PrePat_Restore_String = PrePat_Restore_String + ";" + pinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(pinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD4)), "0.000#########")
    '''                End If
    '''            ElseIf (PreMeasStore = True) Then
    '''                If (PreMeas_Restore_String = "") Then
    '''                    PreMeas_Restore_String = pinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(pinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD4)), "0.000#########")
    '''                Else
    '''                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + pinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(pinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD4)), "0.000#########")
    '''                End If
    '''            End If
    '''                   TheHdw.Digital.Pins(pinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeD4) = CDbl(pin_value)
    '''            GoTo continue1
    '''        End If
    '===============================================
    If (UCase(Pin_Type) Like "*R0*") And pin_value <> "" Then
        Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0)), "0.000#########")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0)), "0.000#########")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0)), "0.000#########")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0)), "0.000#########")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR0) = CDbl(pin_value)
        GoTo continue1
    End If

    If (UCase(Pin_Type) Like "*R1*") And pin_value <> "" Then
        Edge_and_TimeSet = Split(UCase(Pin_Type), ",")
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1)), "0.000#########")
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1)), "0.000#########")
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1)), "0.000#########")
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":" + UCase(Pin_Type) + ":" + Format(CStr(TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1)), "0.000#########")
            End If
        End If
        TheHdw.Digital.Pins(PinName).Timing.EdgeTime(Edge_and_TimeSet(1), chEdgeR1) = CDbl(pin_value)
        GoTo continue1
    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ PPMU Connect Control\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If (LCase(Pin_info_temp(1)) = "connectppmu") Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":DisConnectPPMU"
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":DisConnectPPMU"
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":DisConnectPPMU"
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":DisConnectPPMU"
            End If
        End If
        '20170619 evans.lo
        TheHdw.Digital.Pins(Pin_info_temp(0)).Disconnect
        TheHdw.PPMU.Pins(Pin_info_temp(0)).Connect
        GoTo continue1
    End If
    If (LCase(Pin_info_temp(1)) = "disconnectppmu") Then
        TheHdw.PPMU.Pins(Pin_info_temp(0)).Disconnect
        GoTo continue1
    End If
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Relay Control\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If (LCase(Pin_info_temp(1)) = "relay_on") Then
        TheHdw.Utility.Pins(PinName).State = tlUtilBitOn
        GoTo continue1
    End If
    If (LCase(Pin_info_temp(1)) = "relay_off") Then
        TheHdw.Utility.Pins(PinName).State = tlUtilBitOff
        GoTo continue1
    End If
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Digital Connect Control\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If (LCase(Pin_info_temp(1)) = "connectdigital") Then
        TheHdw.Digital.Pins(Pin_info_temp(0)).Connect

        GoTo continue1
    End If
    If (LCase(Pin_info_temp(1)) = "disconnectdigital") Then
        TheHdw.Digital.Pins(Pin_info_temp(0)).Disconnect


        If (PrePatStore = True) Then  '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ This part is for PLL I measurement 20170714 Kim
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":ConnectDigital"
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":ConnectDigital"
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":ConnectDigital"
            Else
                PreMeas_Restore_String = PreMeas_Restore_String + ";" + PinName + ":ConnectDigital"
            End If
        End If

        GoTo continue1
    End If
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Digital compare\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If (LCase(Pin_info_temp(1)) = "disablecompare") Then
        TheHdw.Digital.Pins(Pin_info_temp(0)).DisableCompare = True
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":EnableCompare"
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":EnableCompare"
            End If
        End If
        GoTo continue1
    End If
    If (LCase(Pin_info_temp(1)) = "enablecompare") Then
        TheHdw.Digital.Pins(Pin_info_temp(0)).DisableCompare = False


        GoTo continue1
    End If
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If (InStr(LCase(Pin_info_temp(1)), "init") > 0) Then
        If (LCase(Pin_info_temp(1)) = "inithi") Then
            TheHdw.Digital.Pins(Pin_info_temp(0)).InitState = chInitHi
        End If
        If (LCase(Pin_info_temp(1)) = "initlo") Then
            TheHdw.Digital.Pins(Pin_info_temp(0)).InitState = chInitLo
        End If

        If (LCase(Pin_info_temp(1)) = "inithiz") Then
            TheHdw.Digital.Pins(Pin_info_temp(0)).InitState = chInitoff
        End If
        GoTo continue1
    End If


    '        \\\\\\Setup Timeing for run pattern\\\\\\
    If (LCase(Pin_info_temp(1)) = "ac spec") Then
        TheExec.Overlays.ApplyUniformSpecToHW Pin_info_temp(0), CDbl(Spec_Evaluate_AC(Pin_info_temp(2)))
        GoTo continue1
    End If
    If (LCase(Pin_info_temp(1)) = "tck") Then
        TheExec.Overlays.ApplyUniformSpecToHW Pin_info_temp(0), CDbl(Spec_Evaluate_AC(Pin_info_temp(2)))
        GoTo continue1
    End If
    If (LCase(Pin_info_temp(1)) = "shiftin") Then
        TheExec.Overlays.ApplyUniformSpecToHW Pin_info_temp(0), CDbl(Spec_Evaluate_AC(Pin_info_temp(2)))
        GoTo continue1
    End If


    '\\\\\\Setup Timeing for nwire\\\\\\
    '\\\\\\Disable FRC\\\\\\
    If (LCase(Pin_info_temp(1)) = "disable_frc") Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":enable_frc"
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":enable_frc"
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":enable_frc"
            Else
                PreMeas_Restore_String = PrePat_Restore_String + ";" + PinName + ":enable_frc"
            End If
        End If
        Disable_FRC Pin_info_temp(0)
        GoTo continue1
    End If

    '\\\\\\Enable FRC\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    If (LCase(Pin_info_temp(1)) = "enable_frc") Then
        Enable_FRC Pin_info_temp(0)
        GoTo continue1
    End If

    '\\\\\\Disable FRC\\\\\\ 20170817 disable FRC by switch relay
    If (LCase(Pin_info_temp(1)) = "disable_frc_relay") Then
        If (PrePatStore = True) Then
            If (PrePat_Restore_String = "") Then
                PrePat_Restore_String = PinName + ":enable_frc_relay"
            Else
                PrePat_Restore_String = PrePat_Restore_String + ";" + PinName + ":enable_frc_relay"
            End If
        ElseIf (PreMeasStore = True) Then
            If (PreMeas_Restore_String = "") Then
                PreMeas_Restore_String = PinName + ":enable_frc_relay"
            Else
                PreMeas_Restore_String = PrePat_Restore_String + ";" + PinName + ":enable_frc_relay"
            End If
        End If

        TheExec.Datalog.WriteComment "=======Disable XO0======="
        TheHdw.Digital.Pins(Pin_info_temp(0)).Disconnect
        With TheHdw.PPMU.Pins(Pin_info_temp(0))
            .Disconnect
            .ForceV 0, 0.002
            .Connect
            .Gate = tlOn
        End With
        'TheHdw.Utility.Pins("K1").State = tlUtilBitOn
        GoTo continue1
    End If

    '\\\\\\Enable FRC\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\20170817 enable FRC by switch relay
    If (LCase(Pin_info_temp(1)) = "enable_frc_relay") Then
        TheExec.Datalog.WriteComment "=======Enable XO0======="
        TheHdw.PPMU.Pins(Pin_info_temp(0)).Gate = tlOff
        TheHdw.PPMU.Pins(Pin_info_temp(0)).Disconnect
        TheHdw.Digital.Pins(Pin_info_temp(0)).Connect
        'TheHdw.Utility.Pins("K1").State = tlUtilBitOff
        GoTo continue1
    End If


    '\\\\\\Enable FRC\\\\\\\obsolete\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    '        If (LCase(Pin_info_temp(1)) = "nwire") Then
    '            Re_store = CStr(TheExec.Specs.AC(Pin_info_temp(0)).CurrentValue(SetSite))
    '            If (LCase(Pin_info_temp(0)) Like "x*") Then
    '                If XI0_GP <> "" Then
    '                    Call VaryFreq("XI0_Port", CDbl(Pin_info_temp(2)), "XI0_Freq_VAR")
    '                ElseIf XI0_Diff_GP <> "" Then
    '                    Call VaryFreq("XI0_Diff_Port", CDbl(Pin_info_temp(2)), "XI0_Diff_Freq_VAR")
    '                End If
    '            End If
    '            If (LCase(Pin_info_temp(0)) Like "rt*") Then
    '                If RTCLK_GP <> "" Then
    '                    Call VaryFreq("RT_CLK32768_Port", CDbl(Pin_info_temp(2)), "RT_CLK32768_Freq_VAR")
    '                ElseIf RTCLK_Diff_GP <> "" Then
    '                    Call VaryFreq("RT_CLK32768_Diff_Port", CDbl(Pin_info_temp(2)), "RT_CLK32768_Diff_Freq_VAR")
    '                End If
    '            End If
    '            GoTo continue1
    '        End If

    TheExec.AddOutput "[Warning] Setup string" & cat_split_temp & " not support"


continue1:
    ''================================================================================================================================
    ''180425 update for trace compensation
    '    Next cat_split_temp
    '    If flag_shmoo_set_current_point = True Then Shmoo_Set_Current_Point 'restore shmoo conidtion overrided in SetForceCondition
    '    TheHdw.Wait 0.001 '20160302  add settling time
    '    If (PrePatStore = True) Then
    '        TheExec.DataLog.WriteComment "Save PrePat Force Condtion:" & PrePat_Restore_String
    '    ElseIf (PreMeasStore = True) Then
    '        TheExec.DataLog.WriteComment "Save PreMeasForce Condtion:" & PreMeas_Restore_String
    '    End If

Next cat_split_temp
If Not g_Vbump_function = True Then
    Shmoo_Set_Current_Point    'restore shmoo conidtion overrided in SetForceCondition
End If

If (gTerm_cond_All <> "") Then
    Call Trace_Compensation(flag_shmoo_set_current_point)
End If
TheHdw.Wait 0.001    '20160302  add settling time
If (PrePatStore = True) Then
    TheExec.Datalog.WriteComment "Save PrePat Force Condtion:" & PrePat_Restore_String
ElseIf (PreMeasStore = True) Then
    TheExec.Datalog.WriteComment "Save PreMeasForce Condtion:" & PreMeas_Restore_String
End If
''================================================================================================================================

PrePatStore = False
PreMeasStore = False

'/////////////////// 20180703 add for clean content of  Charz_Force_Power_condition ////////////////////////
If TheExec.DevChar.Setups.IsRunning = True Then
    Dim SetupName As String
    Dim X_RangeFrom As Double
    Dim Y_RangeFrom As Double

    SetupName = TheExec.DevChar.Setups.ActiveSetupName
    If Not ((TheExec.DevChar.Results(SetupName).StartTime Like "1/1/0001*" Or TheExec.DevChar.Results(SetupName).StartTime Like "0001/1/1*")) Then
        With TheExec.DevChar.Setups(SetupName)
            If .Shmoo.Axes.Count > 1 Then
                X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.from
                Y_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Range.from
                For Each Site In TheExec.Sites    ''20181101 current point need site value
                    XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                    YVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).Value
                Next Site
                If XVal = X_RangeFrom And YVal = Y_RangeFrom Then
                    gl_flag_end_shmoo = False
                End If
                If gl_flag_end_shmoo = True Then
                    Charz_Force_Power_condition = ""
                End If
            Else
                X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.from
                For Each Site In TheExec.Sites    ''20181101 current point need site value
                    XVal = TheExec.DevChar.Results(SetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_X).Value
                Next Site
                If XVal = X_RangeFrom Then
                    gl_flag_end_shmoo = False
                End If
                If gl_flag_end_shmoo = True Then
                    Charz_Force_Power_condition = ""
                End If
            End If
        End With
    End If
End If
'///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Exit Function
err1:
Debug.Print "Error String : " & Setup_string
TheExec.AddOutput "Error ForeceCondition : " & Setup_string
If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Add_Term_Restore(Setup_string As String)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Add_Term_Restore"

    ' extract "term" out of Setu_string and process later?
    ' pgrp1: pin1A,pin1B?
    ' case 1:?
    '   pin1A:vih:1.8;pin1A:term:S,0,50,    => add_fc pin1A:vil:0?
    '   shmoo pin1A:1,2,0.1?
    ' case 1:?
    '   pin1A:vih:1.8;pin1A:vil:0;pin1A:term:S,0,50,    => add_fc?
    '   shmoo pin1A:1,2,0.1?
    ' case 2:?
    '   pin1A:vih:1.8;pgrp1:term:S,0,50,    => add_fc pgrp1:vih:1.8;pgrp1:vil:0 (if pin1A:vil<>pin1B.vil or pin1A:vih<>pin1B:vil => add_fc pin1A:vih:1.8, pin1B:vih:1.9;pin1A:vil:0; pin1B:vil:0.1)?
    '   shmoo pgrp1:vih:1,2,0.1?
    ' case 3:?
    '   pin1A:vih:1.8;pin1B:vih:1.9;pgrp1:term:S,0,50,    => add_fc pgrp1:vil:0 (if pin1A:vil<>pin1B.vil => add_fc pin1A:vil:0, pin1B:vil:0.1)?
    '   shmoo pgrp1:vih:1,2,0.1?
    ' case 2:?
    '   pin1A:vih:1.8;pin1A:vil:0;pgrp1:term:S,0,50,    => add_fc pgrp1:vil:0 (if pin1A:vil<>pin1B.vil => add_fc pin1B:vil:0.1)?
    '   shmoo pgrp1:vih:1,2,0.1?
    ' case 2:?
    '   pin1A:vih:1.8;pin1B:vil:0.1;pgrp1:term:S,0,50,    => add_fc pgrp1:vil:0.1 (if pin1A:vil<>pin1B.vil => add_fc pin1A:vil:0)?
    '   shmoo pgrp1:vih:1,2,0.1?
    ' case 2:?
    '   pin1A:vih:1.8;pin1B:vil:0.1;pgrp1:term:S,0,50,    => add_fc pgrp1:vih:1.8;pgrp1:vil:0.1 (if pin1A:vil<>pin1B.vil or pin1A:vih<>pin1B:vil => add_fc pin1A:vih:1.8, pin1B:vih:1.9;pin1A:vil:0; pin1B:vil:0.1)?
    '   shmoo pin1A:vih:1,2,0.1;pin1B:vih:1.1,2.1,0.1?
    Dim condition_ary() As String
    Dim cond_element_ary() As String
    Dim cond      As Variant
    Dim Pin_Cnt As Long, Pin_Ary() As String
    Dim p         As String
    Dim term_cond_ary() As String
    Dim cond_add  As String
    Dim term_setting_ary() As String, term_type As String
    Dim DevChar_Setup As String
    ' Setup_string = p1:term:S,0,50, => gTerm_Restore_cond = p1:vih:1;p1:vil:0
    If InStr(Setup_string, "term") <= 0 Then
        Exit Function
    Else:
        '        gTerm_cond_Flag = True
    End If
    If TheExec.DevChar.Setups.IsRunning = True Then
        DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
        If gTerm_Restore_cond <> "" Then Setup_string = gTerm_Restore_cond & ";" & Setup_string
        gTerm_cond_shm_Flag = True
        Exit Function
    End If
    gTerm_cond_All = ""             'p1:term:S,0,50,
    gTerm_Restore_cond = ""         'p1:vih:1;p1:vil:0
    condition_ary = Split(Setup_string, ";")

    For Each cond In condition_ary
        If cond <> "" Then
            cond_element_ary = Split(cond, ":")
            p = cond_element_ary(0)
            If UBound(cond_element_ary) > 0 Then
                If cond_element_ary(1) = "term" Then
                    term_setting_ary = Split(cond_element_ary(2), ",")
                    term_type = term_setting_ary(0)
                    If gTerm_cond_All = "" Then
                        gTerm_cond_All = cond
                    Else:
                        gTerm_cond_All = gTerm_cond_All & ";" & cond
                    End If
                    If term_type = "s" Then
                        cond_add = p & ":vih:" & Format(TheHdw.Digital.Pins(p).Levels.Value(chVih), "0.0000") & ";" & _
                                   p & ":vil:" & Format(TheHdw.Digital.Pins(p).Levels.Value(chVil), "0.0000")
                    ElseIf term_type = "d" Then
                        cond_add = p & ":vid:" & Format(TheHdw.Digital.Pins(p).DifferentialLevels.Value(chVid), "0.0000") & ";" & _
                                   p & ":vicm:" & Format(TheHdw.Digital.Pins(p).DifferentialLevels.Value(chVicm), "0.0000")
                    End If
                    If gTerm_Restore_cond = "" Then
                        gTerm_Restore_cond = cond_add
                    Else:
                        gTerm_Restore_cond = gTerm_Restore_cond & ";" & cond_add
                    End If
                End If
            End If
        End If
    Next cond


    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function





Public Function Trace_Compensation(flag_shmoo_set_current_point As Boolean) As Double
    Dim Pin_info() As String
    Dim TypeName  As String
    Dim Vi_dut, Rs, Rt_dut_d, Rt_dut_u, Vt_dut As Double
    Dim PinName   As String
    Dim PinData   As String
    Dim Delta_Rs  As New PinListData
    Dim Site      As Variant
    Dim Job       As String
    Dim vi_h, vi_l, vi_d, vi_cm As New PinListData
    Dim pin_grp() As String, pin_grp1() As String
    Dim Pin_Cnt As Long, pin_cnt1 As Long, x As Long
    Dim pin_temp As Variant, p1 As Variant
    Dim RAK_Val() As Double
    Dim Rt_dut    As Double
    Dim Term_cond As Variant, term_cond_ary() As String
    Dim PinData_ary() As String
    Dim DevChar_Setup As String
    Dim Shmoo_Tracking_Item As Variant
    Dim shmoo_axis As Variant
    Dim Shmoo_Param_Type As String, Shmoo_Param_Name As String, shmoo_pin As String, Shmoo_value As Double, Port_name As String
    Dim Shmoo_Step_Name As String, Shmoo_TimeSets As String

    On Error GoTo err1
    Dim funcName  As String:: funcName = "Trace_Compensation"

    term_cond_ary = Split(gTerm_cond_All, ";")
    For Each Term_cond In term_cond_ary
        Pin_info = Split(Term_cond, ":")
        PinName = Pin_info(0)
        PinData = Pin_info(2)
        Vt_dut = -999
        Rt_dut_u = 999999999
        Rt_dut_d = 999999999

        PinData_ary = Split(PinData, ",")
        TypeName = PinData_ary(0)

        If (PinData_ary(1) <> "") Then
            Vt_dut = CDbl(PinData_ary(1))
        End If
        If (PinData_ary(2) <> "") Then
            Rt_dut_u = CDbl(PinData_ary(2))
            Rt_dut = Rt_dut_u
        End If
        If (PinData_ary(3) <> "") Then
            Rt_dut_d = CDbl(PinData_ary(3))
            Rt_dut = Rt_dut_d
        End If

        Rs = 50
        If (TheExec.DataManager.pintype(PinName) = "Differential") Then
            '180430 avoid PinName is PinGroupPin
            Dim GroupDiff_pin_grp() As String
            Dim GruopDiff_Pin_Cnt As Long, II As Long
            Dim GroupDiff_pin_temp As Variant
            Call TheExec.DataManager.DecomposePinList(PinName, GroupDiff_pin_grp, GruopDiff_Pin_Cnt)
            pin_grp = Split(PinName, ",")
            Pin_Cnt = UBound(pin_grp)
        Else:
            Call TheExec.DataManager.DecomposePinList(PinName, pin_grp, Pin_Cnt)
        End If

        For Each pin_temp In pin_grp
            '            pin_temp = LCase(pin_temp)
            '            Delta_Rs.AddPin (pin_temp)
            '' 20170711 - Use CurrentJob_Card_RAK to replace RAK of each job
            '            For Each site In TheExec.sites
            '                RAK_Val = TheHdw.PPMU.ReadRakValuesByPinnames(pin_temp, site)
            '                Delta_Rs.pin(pin_temp).Value = CurrentJob_Card_RAK.Pins(pin_temp).Value + RAK_Val(0)
            '            Next site


            If (TypeName = "s") Then
                pin_temp = LCase(pin_temp)
                Delta_Rs.AddPin (pin_temp)

                For Each Site In TheExec.Sites
                    RAK_Val = TheHdw.PPMU.ReadRakValuesByPinnames(pin_temp, Site)
                    Delta_Rs.Pin(pin_temp).Value = CurrentJob_Card_RAK.Pins(pin_temp).Value + RAK_Val(0)
                Next Site

                If (TheExec.DataManager.pintype(pin_temp) <> "I/O") Then
                    TheExec.AddOutput "Pin : " + pin_temp + " Is Not Single-End Pin!"
                    Exit Function
                End If
                If (Vt_dut = -999) Then Vt_dut = 0
                For Each Site In TheExec.Sites
                    TheHdw.Digital.Pins(pin_temp).Levels.Value(chVih) = TheHdw.Digital.Pins(pin_temp).Levels.Value(chVih) * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) - Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut_u
                    TheHdw.Digital.Pins(pin_temp).Levels.Value(chVil) = TheHdw.Digital.Pins(pin_temp).Levels.Value(chVil) * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) - Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut_u
                    TheExec.Datalog.WriteComment pin_temp & ": VIH(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.Value(chVih))
                    TheExec.Datalog.WriteComment pin_temp & ": VIL(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.Value(chVil))
                Next Site
            ElseIf (TypeName = "d") Then
                ' 180504 Assume the RAK value of Differential pin P/N is the same, just take P or N value
                '        if does not assumethe same, the formulas below needs to be modified.
                For II = 0 To GruopDiff_Pin_Cnt - 2
                    Delta_Rs.AddPin LCase(GroupDiff_pin_grp(II))
                    For Each Site In TheExec.Sites
                        '                        For Each GroupDiff_pin_temp In GroupDiff_pin_grp
                        RAK_Val = TheHdw.PPMU.ReadRakValuesByPinnames(GroupDiff_pin_grp(II), Site)
                        Delta_Rs.Pin(GroupDiff_pin_grp(II)).Value = CurrentJob_Card_RAK.Pins(GroupDiff_pin_grp(II)).Value + RAK_Val(0)
                        '                        Next GroupDiff_pin_temp
                        If (Vt_dut <> -999) Then
                            '                            For Each site In TheExec.sites
                            TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) * (1 + (Rs + Delta_Rs.Pin(GroupDiff_pin_grp(II)).Value(Site)) * (1 / Rt_dut)) - Vt_dut * (Rs + Delta_Rs.Pin(GroupDiff_pin_grp(II)).Value(Site)) / Rt_dut
                            TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm) * (1 + (Rs + Delta_Rs.Pin(GroupDiff_pin_grp(II)).Value(Site)) * (1 / Rt_dut)) - Vt_dut * (Rs + Delta_Rs.Pin(GroupDiff_pin_grp(II)).Value(Site)) / Rt_dut
                            TheExec.Datalog.WriteComment pin_temp & ": VID(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid))
                            TheExec.Datalog.WriteComment pin_temp & ": VICM(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm))
                            '                            Next site
                        Else:
                            '                            For Each site In TheExec.sites
                            TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) * (Rs + Delta_Rs.Pin(GroupDiff_pin_grp(II)).Value + Rt_dut) / Rt_dut
                            TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm)
                            TheExec.Datalog.WriteComment pin_temp & ": VID(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid))
                            TheExec.Datalog.WriteComment pin_temp & ": VICM(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm))
                            '                            Next site
                        End If
                    Next Site
                Next II
            End If
        Next pin_temp
    Next Term_cond

    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
Public Function SetPowerValue(ByVal Pin As String, ByVal pin_value As String)
    Dim PinList() As String
    Dim PinNum    As Integer
    Dim get_type  As String
    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim var       As Variant
    Dim PinName   As String

    On Error GoTo err1
    Dim funcName  As String:: funcName = "SetPowerValue"

    Call TheExec.DataManager.DecomposePinList(Pin, PinList, typesCount)


    For Each var In PinList
        '        PinName = TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes(tlDevCharShmooAxis_X).ApplyTo.Pins
        '        If (var = PinName) Then GoTo cont
        Call TheExec.DataManager.GetChannelTypes(var, typesCount, stringTypes)
        If (stringTypes(0) Like "DCVS*") Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.Value), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.Value), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.Value), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.Value), "0.000")
                End If
            End If
            If g_Vbump_function = True Then    'add for SelSram function
                If Not g_CharInputString_Voltage_Dict.Exists(UCase(var)) = True Then
                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(pin_value)
                Else
                    g_CharInputString_Voltage_Dict.Remove (UCase(var))
                    g_CharInputString_Voltage_Dict.Add UCase(var), CDbl(pin_value)
                End If
            Else
                TheHdw.DCVS.Pins(var).Voltage.Main.Value = CDbl(pin_value)
            End If

        End If
        If (stringTypes(0) Like "DCVI*") Then
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = var + ":" + "V:" + Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "V:" + Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.000")
                End If
            End If

            'thehdw.DCVI.Pins(var).mode = tlDCVIModeVoltage
            'thehdw.DCVI.Pins(var).Voltage = CDbl(pin_value)
            With TheHdw.DCVI.Pins(var)
                If .Mode <> tlDCVIModeVoltage Then
                    .Gate = False
                    .Mode = tlDCVIModeVoltage
                End If
                .Voltage = CDbl(pin_value)    ' MI_TestCond_UVI80(i).FV_Val
                .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
                .SetCurrentAndRange 0.2, 0.2    'MI_TestCond_UVI80(i).CurrentRange, MI_TestCond_UVI80(i).CurrentRange
                .Connect tlDCVIConnectDefault
                .Gate = True
            End With
        End If
cont:
    Next var

    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function
'20170419 add for case I
Public Function SetPowerValue_I(ByVal Pin As String, ByVal pin_value As String)
    Dim PinList() As String
    Dim PinNum    As Integer
    Dim get_type  As String
    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim var       As Variant
    Dim PinName   As String

    On Error GoTo err1
    Dim funcName  As String:: funcName = "SetPowerValue_I"
    Call TheExec.DataManager.DecomposePinList(Pin, PinList, typesCount)


    For Each var In PinList    '
        Call TheExec.DataManager.GetChannelTypes(var, typesCount, stringTypes)

        If (stringTypes(0) Like "DCVI*") Then
            TheHdw.DCVI.Pins(var).Mode = tlDCVIModeCurrent
            If (PrePatStore = True) Then
                If (PrePat_Restore_String = "") Then
                    PrePat_Restore_String = var + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(var).Current), "0.000")
                Else
                    PrePat_Restore_String = PrePat_Restore_String + ";" + var + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(var).Current), "0.000")
                End If
            ElseIf (PreMeasStore = True) Then
                If (PreMeas_Restore_String = "") Then
                    PreMeas_Restore_String = var + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(var).Current), "0.000")
                Else
                    PreMeas_Restore_String = PreMeas_Restore_String + ";" + var + ":" + "I:" + Format(CStr(TheHdw.DCVI.Pins(var).Current), "0.000")
                End If
            End If
            '20190416 top
            ' check later: evans lo
            If CDbl(pin_value) < TheHdw.DCVI.Pins(var).Current Then
                TheHdw.DCVI.Pins(var).CurrentRange = Abs(CDbl(pin_value))
            End If
            TheHdw.DCVI.Pins(var).Current = CDbl(pin_value)
            '20190416 end
        ElseIf (stringTypes(0) Like "DCVS*") Then
            TheExec.AddOutput "[Warning]  No force I mode for DCVS"
        Else
            TheExec.AddOutput "[Warning] " & var & "not support force I"
        End If
cont:
    Next var

    On Error GoTo 0
    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'Public Function SetPinValue(ByVal Pin As String, ByVal Pin_Type, ByVal pin_value As String)
'
'    Dim PinList() As String
'    Dim PinNum    As Integer
'    Dim get_type  As String
'    Dim typesCount As Long
'    Dim numericTypes() As Long
'    Dim stringTypes() As String
'    Dim var       As Variant
'
'    On Error GoTo err1
'
'    Call TheExec.DataManager.DecomposePinList(Pin, PinList, typesCount)
'
'
'    For Each var In PinList
'        Call TheExec.DataManager.GetChannelTypes(var, typesCount, stringTypes)
'        ''        If (stringTypes(0) Like "VIH*") Then thehdw.Digital.Pins(var).Levels.Value(chVih) = CDbl(pin_value)
'        ''        If (stringTypes(0) Like "VIL*") Then thehdw.Digital.Pins(var).Levels.Value(chVil) = CDbl(pin_value)
'        Pin_Type = UCase(Pin_Type)
'        If (stringTypes(0) = "I/O") Then
'            If (Pin_Type Like "VIH*") Then TheHdw.Digital.Pins(var).Levels.Value(chVih) = CDbl(pin_value)
'            If (Pin_Type Like "VIL*") Then TheHdw.Digital.Pins(var).Levels.Value(chVil) = CDbl(pin_value)
'        End If
'
'    Next var
'
'    Exit Function
'err1:
'    If AbortTest Then Exit Function Else Resume Next
'End Function

Public Function PrintCharSetup(ByVal Setup_string As String) As String


    Dim Cat_temp() As String
    Dim cat_split_temp As Variant

    Dim Pin_info_temp() As String

    Dim PinName, Pin_Type, pin_value As String
    Dim OutputString As String
    Dim Temp      As String
    Dim Site      As Variant
    Dim SetSite   As Integer
    On Error GoTo err1
    Dim funcName  As String:: funcName = "PrintCharSetup"

    For Each Site In TheExec.Sites.Active
        SetSite = Site
    Next Site



    OutputString = ""

    Cat_temp = Split(Setup_string, ";")
    For Each cat_split_temp In Cat_temp

        Pin_info_temp = Split(cat_split_temp, ":")
        If UBound(Pin_info_temp) < 2 Then GoTo continue2
        If (Pin_info_temp(2) = "") Then GoTo continue2

        PinName = Pin_info_temp(0)
        Pin_Type = Pin_info_temp(1)
        pin_value = Pin_info_temp(2)

        PinName = Replace(PinName, "+", ",")



        If (Pin_Type = "V") Then
            Temp = PrintCharValue(PinName)
            If (OutputString = "") Then
                OutputString = Temp
            Else
                OutputString = OutputString & "," & Temp
            End If
            GoTo continue2
        End If
        '        If (UCase(pin_type) = "VDIFF" Or UCase(pin_type) = "VCM") Then
        '            Temp = PrintCharValue(PinName)
        '            If (OutputString = "") Then
        '                OutputString = PinName & ":" & Temp
        '            Else
        '                OutputString = OutputString & "," & PinName & ":" & Temp
        '            End If
        '            GoTo continue2
        '        End If




continue2:
    Next cat_split_temp

    PrintCharSetup = OutputString
    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function PrintCharValue(ByVal Pin As String) As String
    Dim PinList() As String
    Dim PinNum    As Integer
    Dim get_type  As String
    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim var       As Variant
    Dim OutputString As String

    On Error GoTo err1
    Dim funcName  As String:: funcName = "PrintCharValue"
    
    OutputString = ""
    Call TheExec.DataManager.DecomposePinList(Pin, PinList, typesCount)


    For Each var In PinList
        Call TheExec.DataManager.GetChannelTypes(var, typesCount, stringTypes)
        If (stringTypes(0) Like "DCVS*") Then
            If (OutputString = "") Then
                '20170120 change print format
                OutputString = var & ":V:" & Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.Value), "0.0000")
            Else
                OutputString = OutputString & "," & var & ":V:" & Format(CStr(TheHdw.DCVS.Pins(var).Voltage.Main.Value), "0.0000")
            End If
        End If

        If (stringTypes(0) Like "DCVI*") Then
            If (OutputString = "") Then
                '20170120 change print format
                OutputString = var & ":V:" & Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.0000")
            Else
                OutputString = OutputString & "," & var & ":V:" & Format(CStr(TheHdw.DCVI.Pins(var).Voltage), "0.0000")
            End If

        End If

    Next var


    PrintCharValue = OutputString
    Exit Function
err1:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

'Public Function Set_TestName_PTR(ByVal TName As String, ByVal Pin As String) As String
'    '    Dim Tname As String
'    On Error GoTo err1
'    Set_TestName_PTR = TName & " " & Pin & " <> " & TName
'    Exit Function
'err1:
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'Public Function DO_TestLimit_PTR(ByVal Result As SiteDouble, ByVal USL As Double, ByVal LSL As Double, ByVal TName As String)
'    On err GoTo err1
'    TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = True
'    TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = True
'    TheExec.Datalog.ApplySetup
'
'    TheExec.Flow.TestLimit ResultVal:=Result, hiVal:=USL, lowVal:=LSL, TName:=TName, ForceResults:=tlForceNone
'
'
'    TheExec.Datalog.Setup.DatalogSetup.DisableInstanceNameInPTR = False
'    TheExec.Datalog.Setup.DatalogSetup.DisablePinNameInPTR = False
'    TheExec.Datalog.ApplySetup
'    Exit Function
'err1:
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'Public Function Trace_res_Compensation_old(inString As String, flag_shmoo_set_current_point As Boolean) As Double
'
'    Dim Pin_info() As String
'    Dim TypeName  As String
'    Dim Vi_dut, Rs, Rt_dut_d, Rt_dut_u, Vt_dut As Double
'    Dim PinName   As String
'    Dim PinData   As String
'    Dim Delta_Rs  As New PinListData
'    Dim Site      As Variant
'    Dim Job       As String
'    Dim vi_h, vi_l, vi_d, vi_cm As New PinListData
'    Dim pin_grp() As String, pin_grp1() As String
'    Dim Pin_Cnt As Long, pin_cnt1 As Long, x As Long
'    Dim pin_temp As Variant, p1 As Variant
'    Dim RAK_Val() As Double
'    Dim Rt_dut    As Double
'
'    On Error GoTo err1
'    If (LCase(TheExec.CurrentChanMap) Like "*cp*") Then Job = "CP"
'    If (LCase(TheExec.CurrentChanMap) Like "*ft*") Then Job = "FT"
'
'    Pin_info = Split(inString, ":")
'    PinName = Split(inString, ":")(0)
'    PinData = Split(inString, ":")(2)
'
'    TypeName = Split(PinData, ",")(0)
'
'    If (Split(PinData, ",")(1) <> "") Then
'        Vt_dut = CDbl(Split(PinData, ",")(1))
'    Else
'        Vt_dut = -999
'    End If
'    If (Split(PinData, ",")(2) <> "") Then
'        Rt_dut_u = CDbl(Split(PinData, ",")(2))
'        Rt_dut = Rt_dut_u
'    Else
'        Rt_dut_u = 999999999
'    End If
'    If (Split(PinData, ",")(3) <> "") Then
'        Rt_dut_d = CDbl(Split(PinData, ",")(3))
'        Rt_dut = Rt_dut_d
'    Else
'        Rt_dut_d = 999999999
'    End If
'
'    Rs = 50
'    If (TheExec.DataManager.pintype(PinName) = "Differential") Then
'        pin_grp = Split(PinName, ",")
'        Pin_Cnt = UBound(pin_grp)
'    Else
'        Call TheExec.DataManager.DecomposePinList(PinName, pin_grp, Pin_Cnt)
'    End If
'    '    =============================================================================================
'    '    for term case 20170710
'    Dim DevChar_Setup As String
'    Dim Shmoo_Tracking_Item As Variant
'    Dim shmoo_axis As Variant
'    Dim Shmoo_Param_Type As String, Shmoo_Param_Name As String, shmoo_pin As String, Shmoo_value As Double, Port_name As String
'    Dim Shmoo_Step_Name As String, Shmoo_TimeSets As String
'    '    =============================================================================================
'
'    For Each pin_temp In pin_grp
'        Delta_Rs.AddPin (pin_temp)
'        ''        If (Job = "CP") Then
'        ''            For Each Site In theexec.sites
'        ''                RAK_Val = thehdw.PPMU.ReadRakValuesByPinnames(pin_temp, Site)
'        ''                Delta_Rs.Pin(pin_temp).Value = CP_Card_RAK.Pins(0).Value + RAK_Val(0)
'        ''            Next Site
'        ''        ElseIf (Job = "FT") Then
'        ''            For Each Site In theexec.sites
'        ''                '''Delta_Rs.pin(pin_temp).Value = FT_Card_RAK.Pins(pin_temp).Value
'        ''                RAK_Val = thehdw.PPMU.ReadRakValuesByPinnames(pin_temp)
'        ''                Delta_Rs.Pin(pin_temp).Value = FT_Card_RAK.Pins(pin_temp).Value + RAK_Val(0)
'        ''            Next Site
'        ''        End If
'        '' 20170711 - Use CurrentJob_Card_RAK to replace RAK of each job
'        For Each Site In TheExec.Sites
'            RAK_Val = TheHdw.PPMU.ReadRakValuesByPinnames(pin_temp, Site)
'            Delta_Rs.Pin(pin_temp).Value = CurrentJob_Card_RAK.Pins(0).Value + RAK_Val(0)
'        Next Site
'
'
'        If (TypeName = "S") Then
'            If (TheExec.DataManager.pintype(pin_temp) <> "I/O") Then
'                TheExec.AddOutput "Pin : " + pin_temp + " Is Not Single-End Pin!"
'                Exit Function
'            End If
'
'            If (Vt_dut = -999) Then Vt_dut = 0
'            For Each Site In TheExec.Sites
'                ''===============================================================================================================
'                ''              20170608 VIH/VIL shmoo
'                DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
'
'                If TheExec.DevChar.Setups.IsRunning = False Then
'                    TheHdw.Digital.Pins(pin_temp).Levels.Value(chVih) = TheHdw.Digital.Pins(pin_temp).Levels.Value(chVih) * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut_u
'                    TheHdw.Digital.Pins(pin_temp).Levels.Value(chVil) = TheHdw.Digital.Pins(pin_temp).Levels.Value(chVil) * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut_u
'                    TheExec.Datalog.WriteComment pin_temp & ": VIH(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.Value(chVih))
'                    TheExec.Datalog.WriteComment pin_temp & ": VIL(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).Levels.Value(chVil))
'                ElseIf TheExec.DevChar.Results(DevChar_Setup).StartTime Like "1/1/0001*" Or TheExec.DevChar.Results(DevChar_Setup).StartTime Like "0001/1/1*" Then
'                Else
'                    Shmoo_Set_Current_Point
'                    With TheExec.DevChar.Setups(DevChar_Setup).Shmoo
'                        For Each shmoo_axis In .Axes.List
'                            Shmoo_Param_Type = .Axes.Item(shmoo_axis).Parameter.Type
'                            Shmoo_Param_Name = .Axes.Item(shmoo_axis).Parameter.Name
'                            '                            shmoo_pin = .Axes.Item(shmoo_axis).ApplyTo.Pins
'                            '                            Shmoo_TimeSets = .Axes.Item(shmoo_axis).ApplyTo.Timesets
'                            If LCase(Shmoo_Param_Type) = "level" Then
'                                Shmoo_value = TheExec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).Value
'                                If LCase(Shmoo_Param_Name) = "vih" Then
'                                    TheHdw.Digital.Pins(pin_temp).Levels.Value(chVih) = Shmoo_value * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut_u
'                                    flag_shmoo_set_current_point = False
'                                ElseIf LCase(Shmoo_Param_Name) = "vil" Then
'                                    TheHdw.Digital.Pins(pin_temp).Levels.Value(chVil) = Shmoo_value * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut_u
'                                    flag_shmoo_set_current_point = False
'                                End If
'                            End If
'                            With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).TrackingParameters
'                                For Each Shmoo_Tracking_Item In .List
'                                    Shmoo_Param_Type = .Item(Shmoo_Tracking_Item).Type
'                                    Shmoo_Param_Name = .Item(Shmoo_Tracking_Item).Name
'                                    '                                    shmoo_pin = .Item(Shmoo_Tracking_Item).ApplyTo.Pins
'                                    '                                    Shmoo_TimeSets = .Item(Shmoo_Tracking_Item).ApplyTo.Timesets
'                                    If LCase(Shmoo_Param_Type) = "level" Then
'                                        Shmoo_value = TheExec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).TrackingParameters(Shmoo_Tracking_Item).Value
'                                        If LCase(Shmoo_Param_Name) = "vih" Then
'                                            TheHdw.Digital.Pins(pin_temp).Levels.Value(chVil) = Shmoo_value * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut_u
'                                            flag_shmoo_set_current_point = False
'                                        ElseIf LCase(Shmoo_Param_Name) = "vil" Then
'                                            TheHdw.Digital.Pins(pin_temp).Levels.Value(chVih) = Shmoo_value * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut_d + 1 / Rt_dut_u)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut_u
'                                            flag_shmoo_set_current_point = False
'                                        End If
'                                    End If
'                                Next Shmoo_Tracking_Item
'                            End With
'                        Next shmoo_axis
'                    End With
'                End If
'                ''===============================================================================================================
'            Next Site
'
'        ElseIf (TypeName = "D") Then
'            If (Vt_dut <> -999) Then
'                For Each Site In TheExec.Sites
'                    '                        vi_d = thehdw.Digital.Pins(PinName).DifferentialLevels.Value(chVid) * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value) * (1 / Rt_dut)) + Vt_dut * (Delta_Rs.Pin(pin_temp).Value + Delta_Rs.Pin(pin_temp).Value) / Rt_dut
'                    '                        vi_cm = thehdw.Digital.Pins(PinName).DifferentialLevels.Value(chVicm) * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value) * (1 / Rt_dut)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value) / Rt_dut
'                    DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
'                    If TheExec.DevChar.Setups.IsRunning = False Then
'                        TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut
'                        TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm) * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut
'                        TheExec.Datalog.WriteComment pin_temp & ": VID(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid))
'                        TheExec.Datalog.WriteComment pin_temp & ": VICM(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm))
'                    ElseIf TheExec.DevChar.Results(DevChar_Setup).StartTime Like "1/1/0001*" Or TheExec.DevChar.Results(DevChar_Setup).StartTime Like "0001/1/1*" Then
'                    Else
'                        Shmoo_Set_Current_Point
'                        With TheExec.DevChar.Setups(DevChar_Setup).Shmoo
'                            For Each shmoo_axis In .Axes.List
'                                Shmoo_Param_Type = .Axes.Item(shmoo_axis).Parameter.Type
'                                Shmoo_Param_Name = .Axes.Item(shmoo_axis).Parameter.Name
'                                '                            shmoo_pin = .Axes.Item(shmoo_axis).ApplyTo.Pins
'                                '                            Shmoo_TimeSets = .Axes.Item(shmoo_axis).ApplyTo.Timesets
'                                If LCase(Shmoo_Param_Type) = "level" Then
'                                    Shmoo_value = TheExec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).Value
'                                    If LCase(Shmoo_Param_Name) = "vid" Then
'                                        TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) = Shmoo_value * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut
'                                        flag_shmoo_set_current_point = False
'                                    ElseIf LCase(Shmoo_Param_Name) = "vicm" Then
'                                        TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm) = Shmoo_value * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut
'                                        flag_shmoo_set_current_point = False
'                                    End If
'                                End If
'                                With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).TrackingParameters
'                                    For Each Shmoo_Tracking_Item In .List
'                                        Shmoo_Param_Type = .Item(Shmoo_Tracking_Item).Type
'                                        Shmoo_Param_Name = .Item(Shmoo_Tracking_Item).Name
'                                        '                                    shmoo_pin = .Item(Shmoo_Tracking_Item).ApplyTo.Pins
'                                        '                                    Shmoo_TimeSets = .Item(Shmoo_Tracking_Item).ApplyTo.Timesets
'                                        If LCase(Shmoo_Param_Type) = "level" Then
'                                            Shmoo_value = TheExec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).TrackingParameters(Shmoo_Tracking_Item).Value
'                                            If LCase(Shmoo_Param_Name) = "vid" Then
'                                                TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) = Shmoo_value * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut
'                                                flag_shmoo_set_current_point = False
'                                            ElseIf LCase(Shmoo_Param_Name) = "vicm" Then
'                                                TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm) = Shmoo_value * (1 + (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) * (1 / Rt_dut)) + Vt_dut * (Rs + Delta_Rs.Pin(pin_temp).Value(Site)) / Rt_dut
'                                                flag_shmoo_set_current_point = False
'                                            End If
'                                        End If
'                                    Next Shmoo_Tracking_Item
'                                End With
'                            Next shmoo_axis
'                        End With
'                    End If
'                Next Site
'            Else
'                For Each Site In TheExec.Sites
'                    TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid) * (Rs + Delta_Rs.Pin(pin_temp).Value + Rt_dut) / Rt_dut
'                    TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm) = TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm)
'                    TheExec.Datalog.WriteComment pin_temp & ": VID(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVid))
'                    TheExec.Datalog.WriteComment pin_temp & ": VICM(" & Site & ")&= " & CStr(TheHdw.Digital.Pins(pin_temp).DifferentialLevels.Value(chVicm))
'                Next Site
'            End If
'        End If
'
'    Next pin_temp
'
'
'    Exit Function
'err1:
'    If AbortTest Then Exit Function Else Resume Next
'End Function


Public Function Spec_Evaluate_DC_for_flow_loop(ByVal pininfo As String, ByVal condition_info As String, ByVal temp_pin_info As String) As String
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Spec_Evaluate_DC_for_flow_loop"

    Dim temp_pin_name, temp_flowint_name, calc_info As String
    Dim temp_pininfo_arr() As String
    Dim i         As Integer
    Dim fake_calc_info, outstring_ori, outstring_evaluated As String
    Dim flow_int_value As Double
    Dim Site      As Variant

    For Each Site In TheExec.Sites
        Exit For
    Next Site


    outstring_ori = pininfo & ":" & condition_info & ":" & temp_pin_info

    If InStr(temp_pin_info, "+") > 0 Or InStr(temp_pin_info, "-") > 0 Or InStr(temp_pin_info, "*") > 0 Or InStr(temp_pin_info, "/") > 0 Then
        calc_info = temp_pin_info

        temp_pin_info = Replace(temp_pin_info, "(", "")
        temp_pin_info = Replace(temp_pin_info, ")", "")
        temp_pin_info = Replace(temp_pin_info, "+", "~")
        temp_pin_info = Replace(temp_pin_info, "-", "~")
        temp_pin_info = Replace(temp_pin_info, "*", "~")
        temp_pin_info = Replace(temp_pin_info, "/", "~")

        fake_calc_info = Replace(calc_info, "(", " ( ")
        fake_calc_info = Replace(fake_calc_info, ")", " ) ")
        fake_calc_info = Replace(fake_calc_info, "+", " + ")
        fake_calc_info = Replace(fake_calc_info, "-", " - ")
        fake_calc_info = Replace(fake_calc_info, "*", " * ")
        fake_calc_info = Replace(fake_calc_info, "/", " / ")
        fake_calc_info = " " & fake_calc_info & " "

        temp_pininfo_arr = Split(temp_pin_info, "~")

        For i = 0 To UBound(temp_pininfo_arr)
            If InStr(temp_pininfo_arr(i), "_") <> 0 Then

                If InStr(LCase(temp_pininfo_arr(i)), "flow_int") <> 0 Then
                    flow_int_value = TheExec.Flow.var(temp_pininfo_arr(i)).Value
                    calc_info = Replace(fake_calc_info, " " & temp_pininfo_arr(i) & " ", flow_int_value, , 1)
                    fake_calc_info = calc_info
                Else
                    temp_pin_name = temp_pininfo_arr(i)
                    temp_pininfo_arr(i) = CStr(TheExec.Specs.DC.Item(Mid(temp_pininfo_arr(i), 2)).CurrentValue(Site))
                    calc_info = Replace(fake_calc_info, " " & temp_pin_name & " ", temp_pininfo_arr(i), , 1)
                    fake_calc_info = calc_info
                End If


            End If
        Next i
        Spec_Evaluate_DC_for_flow_loop = CStr(Evaluate(calc_info))
        outstring_evaluated = pininfo & ":" & condition_info & ":" & calc_info & "=" & CStr(Spec_Evaluate_DC_for_flow_loop)
        TheExec.Datalog.WriteComment "Calculate_Result:" & Trim(outstring_evaluated)

    Else
        If (InStr(temp_pin_info, "_") = 1) Then
            If TheExec.Specs.DC.Contains(Mid(temp_pin_info, 2)) Then
                Spec_Evaluate_DC_for_flow_loop = CStr(TheExec.Specs.DC.Item(Mid(temp_pin_info, 2)).CurrentValue(Site))
                outstring_evaluated = pininfo & ":" & condition_info & "=" & CStr(Spec_Evaluate_DC_for_flow_loop)
                TheExec.Datalog.WriteComment "Calculate_Result:" & Trim(outstring_evaluated)
            Else
                Spec_Evaluate_DC_for_flow_loop = CStr(temp_pin_info)
            End If
        Else
            If InStr(LCase(temp_pin_info), "flow_int") <> 0 Then
                Spec_Evaluate_DC_for_flow_loop = TheExec.Flow.var(temp_pin_info).Value
                outstring_evaluated = pininfo & ":" & condition_info & "=" & CStr(Spec_Evaluate_DC_for_flow_loop)
                TheExec.Datalog.WriteComment "Calculate_Result:" & Trim(outstring_evaluated)
            Else
                Spec_Evaluate_DC_for_flow_loop = CStr(temp_pin_info)
            End If
        End If
    End If

    ''Debug.Print calc_info & "=" & Spec_Evaluate_DC_for_flow_loop


    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Spec_Evaluate_AC(ByVal temp_pin_info As String) As String
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Spec_Evaluate_AC"

    Dim temp_pin_name, calc_info As String
    Dim temp_pininfo_arr() As String
    Dim i         As Integer
    Dim fake_calc_info As String

    If InStr(temp_pin_info, "+") > 0 Or InStr(temp_pin_info, "-") > 0 Or InStr(temp_pin_info, "*") > 0 Or InStr(temp_pin_info, "/") > 0 Then
        calc_info = temp_pin_info

        temp_pin_info = Replace(temp_pin_info, "(", "")
        temp_pin_info = Replace(temp_pin_info, ")", "")
        temp_pin_info = Replace(temp_pin_info, "+", "~")
        temp_pin_info = Replace(temp_pin_info, "-", "~")
        temp_pin_info = Replace(temp_pin_info, "*", "~")
        temp_pin_info = Replace(temp_pin_info, "/", "~")

        fake_calc_info = Replace(calc_info, "(", " ( ")
        fake_calc_info = Replace(fake_calc_info, ")", " ) ")
        fake_calc_info = Replace(fake_calc_info, "+", " + ")
        fake_calc_info = Replace(fake_calc_info, "-", " - ")
        fake_calc_info = Replace(fake_calc_info, "*", " * ")
        fake_calc_info = Replace(fake_calc_info, "/", " / ")
        fake_calc_info = " " & fake_calc_info & " "

        temp_pininfo_arr = Split(temp_pin_info, "~")

        For i = 0 To UBound(temp_pininfo_arr)
            If InStr(temp_pininfo_arr(i), "_") <> 0 Then
                temp_pin_name = temp_pininfo_arr(i)
                temp_pininfo_arr(i) = CStr(TheExec.Specs.ac.Item(Mid(temp_pininfo_arr(i), 2)).ContextValue)
                calc_info = Replace(fake_calc_info, " " & temp_pin_name & " ", temp_pininfo_arr(i), , 1)
            End If
        Next i
        Spec_Evaluate_AC = CStr(Evaluate(calc_info))
    Else
        If (InStr(temp_pin_info, "_") = 1) Then
            If TheExec.Specs.ac.Contains(Mid(temp_pin_info, 2)) Then
                Spec_Evaluate_AC = CStr(TheExec.Specs.ac.Item(Mid(temp_pin_info, 2)).ContextValue)
            Else
                Spec_Evaluate_AC = CStr(temp_pin_info)
            End If
        Else
            Spec_Evaluate_AC = CStr(temp_pin_info)
        End If
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190416 top
'Public Function Shmoo_Set_Current_Point()
'' Set up shmoo condition for current shmoo point (including tracking)
'' Use Set_Level_Timing_Specto set hardware
'    Dim Shmoo_Pin_Str As String
'    Dim Shmoo_Tracking_Item As Variant, shmoo_axis As Variant
'    Dim DevChar_Setup As String
'    Dim Shmoo_Param_Type As String, Shmoo_Param_Name As String, shmoo_pin As String, Shmoo_value As Double, Port_name As String
'    Dim Shmoo_Step_Name As String, Shmoo_TimeSets As String
'    Dim arg_ary() As String
'    Dim Site As Variant
'    If TheExec.DevChar.Setups.IsRunning = False Then
'        Shmoo_End = False
'    Else
'        DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
'        If Shmoo_End = True Then Exit Function  ' Prevent from setting  to last shmoo point; set Shmoo_End at the end of   PrintShmooInfo
'        If TheExec.DevChar.Results(DevChar_Setup).StartTime Like "1/1/0001 12:00:00 AM" Then Exit Function  ' initial run of shmoo, not the first point
'        With TheExec.DevChar.Setups(DevChar_Setup).Shmoo
'            For Each shmoo_axis In .Axes.List
'                If LCase(.Axes(shmoo_axis).InterposeFunctions.PrePoint.Name) Like "freerunclk_set_xy" Then
'                    arg_ary = Split(.Axes(shmoo_axis).InterposeFunctions.PrePoint.Arguments, ",")
'                    Port_name = arg_ary(1)
'                End If
'                Shmoo_Param_Type = .Axes.Item(shmoo_axis).Parameter.Type
'                Shmoo_Param_Name = .Axes.Item(shmoo_axis).Parameter.Name
'                shmoo_pin = .Axes.Item(shmoo_axis).ApplyTo.Pins
'                Shmoo_TimeSets = .Axes.Item(shmoo_axis).ApplyTo.Timesets
'                For Each Site In TheExec.sites
'                    Shmoo_value = TheExec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).Value
'                    'Debug.Print Shmoo_value
'                    Set_Level_Timing_Spec Shmoo_Param_Type, Shmoo_Param_Name, shmoo_pin, Shmoo_TimeSets, Shmoo_value, Port_name
'                Next Site
'                With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.Axes(shmoo_axis).TrackingParameters
'                    For Each Shmoo_Tracking_Item In .List
'                            Shmoo_Param_Type = .Item(Shmoo_Tracking_Item).Type
'                            Shmoo_Param_Name = .Item(Shmoo_Tracking_Item).Name
'                            shmoo_pin = .Item(Shmoo_Tracking_Item).ApplyTo.Pins
'                            Shmoo_TimeSets = .Item(Shmoo_Tracking_Item).ApplyTo.Timesets
'                            For Each Site In TheExec.sites
'                                Shmoo_value = TheExec.DevChar.Results(DevChar_Setup).Shmoo.CurrentPoint.Axes(shmoo_axis).TrackingParameters(Shmoo_Tracking_Item).Value
'                                Set_Level_Timing_Spec Shmoo_Param_Type, Shmoo_Param_Name, shmoo_pin, Shmoo_TimeSets, Shmoo_value, Port_name
'                            Next Site
'                    Next Shmoo_Tracking_Item
'                End With
'            Next shmoo_axis
'        End With
'    End If
'End Function

'Public Function Set_Level_Timing_Spec(Shmoo_Param_Type As String, Shmoo_Param_Name As String, shmoo_pin As String, Shmoo_TimeSets As String, Shmoo_value As Double, Port_name As String)
''Set instrument hardware
'    Dim InstName As String
'    Dim FRC_pin_name As String, Shmoo_Spec As String
'    If Shmoo_TimeSets <> "" Then
'        TheExec.ErrorLogMessage "Set up Timing set is not supported"
'        Exit Function
'    End If
'    Select Case Shmoo_Param_Type
'        Case "AC Spec", "DC Spec":
'            TheExec.Overlays.ApplyUniformSpecToHW Shmoo_Param_Name, Shmoo_value
'            Shmoo_Spec = Shmoo_Param_Name
'        Case "Level":
'        '20160925 Force to Ucase
'            Select Case UCase(Shmoo_Param_Name)
'                Case "VMAIN":
'                    InstName = GetInstrument(shmoo_pin, 0)
'                    Select Case InstName
'                       Case "DC-07"
'                            TheHdw.DCVI.Pins(shmoo_pin).Voltage = Shmoo_value
'                       Case "VHDVS"
'                            TheHdw.DCVS.Pins(shmoo_pin).Voltage.Main.Value = Shmoo_value
'                       Case "HexVS"
'                            TheHdw.DCVS.Pins(shmoo_pin).Voltage.Main.Value = Shmoo_value
'                       Case Else
'                    End Select
'                Case "VT":
'                   TheHdw.Digital.Pins(shmoo_pin).Levels.DriverMode = tlDriverModeVt
'                   TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVt) = Shmoo_value
'                Case "VIH": TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVih) = Shmoo_value
'                Case "VIL": TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVil) = Shmoo_value
'                Case "VOH": TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVoh) = Shmoo_value
'                Case "VOL": TheHdw.Digital.Pins(shmoo_pin).Levels.Value(chVol) = Shmoo_value
'                Case "VID": TheHdw.Digital.Pins(shmoo_pin).DifferentialLevels.Value(chVid) = Shmoo_value
'                Case "VOD": TheHdw.Digital.Pins(shmoo_pin).DifferentialLevels.Value(chVod) = Shmoo_value
'                Case "VICM":  TheHdw.Digital.Pins(shmoo_pin).DifferentialLevels.Value(chVicm) = Shmoo_value
'                Case Else:
'                    TheExec.ErrorLogMessage "Not supported Shmoo Parameter Name: " & Shmoo_Param_Name
'            End Select
'            Shmoo_Spec = shmoo_pin & "(" & Shmoo_Param_Name & ")"
'        Case "Global Spec":
'            If Port_name <> "" Then ' XI0_Port,XI0_Freq_VAR, XI0_PA
'                Shmoo_Spec = Left(Port_name, Len(Port_name) - 5) & "_Freq_VAR"
'                FRC_pin_name = Left(Port_name, Len(Port_name) - 5) & "_PA"
''                Call VaryFreq(Port_name, Shmoo_value, Shmoo_Spec)
'                If LCase(Port_name) Like "xi0*" Then
'                    If XI0_GP <> "" Then
'                        Call VaryFreq("XI0_Port", Shmoo_value, "XI0_Freq_VAR")
'                    ElseIf XI0_Diff_GP <> "" Then
'                        Call VaryFreq("XI0_Diff_Port", Shmoo_value, "XI0_Diff_Freq_VAR")
'                    End If
'                Else
'                    If XI0_GP <> "" Then
'                        Call VaryFreq("XI0_Port", TheExec.Specs.ac("XI0_Freq_VAR").ContextValue, "XI0_Freq_VAR")
'                    ElseIf XI0_Diff_GP <> "" Then
'                        Call VaryFreq("XI0_Diff_Port", TheExec.Specs.ac("XI0_Diff_Freq_VAR").ContextValue, "XI0_Diff_Freq_VAR")
'                    End If
'                End If
''                If LCase(Port_name) Like "rt*" Then
''                    If RTCLK_GP <> "" Then
''                        Call VaryFreq("RT_CLK32768_Port", Shmoo_value, "RT_CLK32768_Freq_VAR")
''                    ElseIf RTCLK_Diff_GP <> "" Then
''                        Call VaryFreq("RT_CLK32768_Diff_Port", Shmoo_value, "RT_CLK32768_Diff_Freq_VAR")
''                    End If
''                Else
''                    If RTCLK_GP <> "" Then
''                        Call VaryFreq("RT_CLK32768_Port", TheExec.Specs.ac("RT_CLK32768_Freq_VAR").ContextValue, "RT_CLK32768_Freq_VAR")
''                    ElseIf RTCLK_Diff_GP <> "" Then
''                        Call VaryFreq("RT_CLK32768_Diff_Port", TheExec.Specs.ac("RT_CLK32768_Diff_Freq_VAR").ContextValue, "RT_CLK32768_Diff_Freq_VAR")
''                    End If
''                End If
''                FreqMeasDebug FRC_pin_name, 0.5, 0.01, 0.1             'Debug to print out freq in datalog
''                FreqMeasDebug "RT_CLK32768_PA", 0.5, 0.01, 0.1
'            Else
'                TheExec.Overlays.ApplyUniformSpecToHW Shmoo_Param_Name, Shmoo_value
'                Shmoo_Spec = Shmoo_Param_Name
'            End If
'        Case Else:
'            TheExec.ErrorLogMessage "Not supported Shmoo Parameter Name: " & Shmoo_Param_Type
'    End Select
'End Function
'20190416 end
Public Function SetupDCVI_ForceV(pin_name As String, arg_str As String)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SetupDCVI_ForceV"

    'PAD_MTR_ANALOG_TEST_P:SetupFV:Vprog,Irange,CustomizeWaitTime
    Dim arg_ary() As String
    Dim Vprog As Double, Iprog As Double, IRange As Double, CustomizeWaitTime As String
    Dim PowerType() As String
    Dim Factor    As Long
    Dim NumTypes  As Long
    Dim WaitTime  As Double

    arg_ary = Split(arg_str, ",")
    Vprog = CDbl(Spec_Evaluate_DC(arg_ary(0)))
    IRange = CDbl(Spec_Evaluate_DC(arg_ary(1)))
    CustomizeWaitTime = arg_ary(2)

    Call TheExec.DataManager.GetChannelTypes(pin_name, NumTypes, PowerType())

    Select Case PowerType(0)
        Case "DCVI"
            Factor = 1
        Case "DCVIMerged"
            Factor = 2
        Case Else
    End Select

    If IRange > 2 * Factor Then
        IRange = 2 * Factor
        WaitTime = 1.6 * ms
    ElseIf IRange > 1 * Factor Then
        IRange = 2 * Factor
        WaitTime = 1.6 * ms
    ElseIf IRange > 0.2 * Factor Then
        IRange = 1 * Factor
        WaitTime = 1.6 * ms
    ElseIf IRange > 0.02 * Factor Then
        IRange = 0.2 * Factor
        WaitTime = 260 * us
    ElseIf IRange > 0.002 * Factor Then
        IRange = 0.02 * Factor
        WaitTime = 1.5 * ms
    ElseIf IRange > 0.0002 * Factor Then
        IRange = 0.002 * Factor
        WaitTime = 11 * ms
    ElseIf IRange > 0.00002 * Factor Then
        IRange = 0.0002 * Factor
        WaitTime = 1.4 * ms
    Else
        IRange = 0.00002 * Factor
        WaitTime = 6 * ms
    End If

    With TheHdw.DCVI.Pins(pin_name)
        .Gate = False
        .Mode = tlDCVIModeVoltage
        .Voltage = Vprog
        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
        ''20161018 - Swap current and current range sequence to avoid mode alarm
        ''            .Current = Irange
        ''            .CurrentRange.Value = Irange
        .SetCurrentAndRange IRange, IRange
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With

    With TheHdw.DCVI.Pins(pin_name)
        .Meter.Mode = tlDCVIMeterCurrent
        .Meter.CurrentRange.Value = IRange
    End With
    If glb_Disable_CurrRangeSetting_Print = False Then
        TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Meter I range setting, " & pin_name & " =" & TheHdw.DCVI.Pins(pin_name).Meter.CurrentRange.Value)
        TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Curr_meas Force Volt value, " & pin_name & " =" & Format(TheHdw.DCVI.Pins(pin_name).Voltage, "0.000"))
    End If
    If CustomizeWaitTime <> "" Then
        WaitTime = CDbl(CustomizeWaitTime)
    End If
    TheHdw.Wait (WaitTime)

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function SetupDCVI_ForceI(pin_name As String, arg_str As String)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "SetupDCVI_ForceI"

    'PAD_MTR_ANALOG_TEST_P:SetupFI:Vprog,Iprog
    Dim arg_ary() As String
    Dim Vprog As Double, Iprog As Double
    arg_ary = Split(arg_str, ",")
    Vprog = CDbl(Spec_Evaluate_DC(arg_ary(0)))
    Iprog = CDbl(Spec_Evaluate_DC(arg_ary(1)))

    With TheHdw.DCVI.Pins(pin_name)    '' High impedence mode
        If Iprog = 0 Then
            '' 20150612 - High impedence mode
            ' Only required if force was previously connected
            .Disconnect tlDCVIConnectDefault
            ' Program the DCVI mapped to MyPin to high impedance mode
            .Mode = tlDCVIModeHighImpedance
            ' Connect only the sense to use with high impedance mode
            .Connect tlDCVIConnectHighSense
            .Meter.Mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
            .Current = 0
        Else
            .Mode = tlDCVIModeCurrent
            .Connect tlDCVIConnectDefault
            .Voltage = Vprog
            .Meter.Mode = tlDCVIMeterVoltage  '''Change by Martin for TTR 20151230
            ''20170526-Add FI condition
            .CurrentRange.AutoRange = True
            .Current = Iprog
        End If
        .VoltageRange.AutoRange = True
        '        thehdw.Wait (5 * ms)
        .Gate = True
    End With

    If glb_Disable_CurrRangeSetting_Print = False Then
        TheExec.Datalog.WriteComment (TheExec.DataManager.InstanceName & " =====> Volt_meas Force Current value, " & pin_name & " =" & Format(TheHdw.DCVI.Pins(pin_name).Current, "0.000"))
    End If
    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MbistRetentionLevelWait_ForChar(mS_Time As Double, Retention_Voltage() As SiteDouble, Retention_Pins As PinList, RampStep As Double, Optional RampWaitTime As Double = 0.001)
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "MbistRetentionLevelWait_ForChar"


    ' thehdw.Digital.ApplyLevelsTiming True, True, True, tlPowered  'SEC DRAM

    'Dim Retention_Voltage As Double: Retention_Voltage = 0.5
    ''SWLINZA20171103, for ramp up/down retention voltage

    'Dim Retention_Pins As New PinList
    Dim Retention_Pins_Ary() As String
    Dim Retention_Pins_count As Long
    Dim RampDown_Time As Double: RampDown_Time = RampWaitTime    'RampDown_Time = 0
    Dim RampDown_Step As Double
    Dim Original_voltage() As New SiteDouble
    Dim DropVoltage() As New SiteDouble
    Dim DropVoltage_perStep() As New SiteDouble
    Dim Voltage_from_HW As String
    Dim ApplyPins As String
    Dim i, j      As Integer
    Dim Flag_ApplyPower() As New SiteBoolean
    Dim Site      As Variant
    If RampStep = 0 Then
        RampDown_Step = 20    ' default RampDown_Step = 20
    Else
        RampDown_Step = RampStep
    End If

    TheExec.DataManager.DecomposePinList Retention_Pins, Retention_Pins_Ary(), Retention_Pins_count
    ReDim Original_voltage(Retention_Pins_count - 1) As New SiteDouble
    ReDim DropVoltage(Retention_Pins_count - 1) As New SiteDouble
    ReDim DropVoltage_perStep(Retention_Pins_count - 1) As New SiteDouble
    ReDim Flag_ApplyPower(Retention_Pins_count - 1) As New SiteBoolean

    TheExec.Datalog.WriteComment "********************************************************"
    TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"

    For Each Site In TheExec.Sites
        For i = 0 To Retention_Pins_count - 1
            Flag_ApplyPower(i) = True
            Original_voltage(i) = Format(TheHdw.DCVS.Pins(Retention_Pins_Ary(i)).Voltage.Main, 3)
            DropVoltage(i) = Original_voltage(i) - Retention_Voltage(i)
            DropVoltage_perStep(i) = Format((DropVoltage(i) / RampDown_Step), 3)
            If Format(DropVoltage(i), 3) = 0 Then Flag_ApplyPower(i) = False
        Next i

        '--------- Ramp down for retention voltage ------'
        For i = 0 To RampDown_Step - 1
            For j = 0 To Retention_Pins_count - 1
                If Flag_ApplyPower(j) = True Then
                    If i = RampDown_Step - 1 Then
                        TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main = Retention_Voltage(j)
                    Else
                        TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main = Original_voltage(j) - DropVoltage_perStep(j) * i
                    End If
                End If
            Next j
            TheHdw.Wait RampDown_Time / RampDown_Step  'remove extra wait time for each step
        Next i

        Voltage_from_HW = ""
        ApplyPins = ""
        Dim Current_PinCount As Long
        '--------- Read back retention voltage from HW ------'
        For j = 0 To Retention_Pins_count - 1
            If Flag_ApplyPower(j) = True Then
                If Current_PinCount = 0 Then
                    Voltage_from_HW = CStr(Format(TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main, 3))    '& " V"
                    ApplyPins = Retention_Pins_Ary(j)
                Else
                    Voltage_from_HW = Voltage_from_HW & "," & CStr(Format(TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main, 3))    '& " V"
                    ApplyPins = ApplyPins & "," & Retention_Pins_Ary(j)
                End If
                Current_PinCount = Current_PinCount + 1
            End If
        Next j
        TheExec.Datalog.WriteComment "--------- Site:" & Site & "----------"
        TheExec.Datalog.WriteComment "*print: MbistRetention Pins " & ApplyPins
        TheExec.Datalog.WriteComment "*print: MbistRetention Volt " & Voltage_from_HW
    Next Site

    '----- Retention Wait time 100 ms ------
    TheHdw.Wait mS_Time * 0.001
    'TheExec.Datalog.WriteComment "*************************************************"
    'TheExec.Datalog.WriteComment "*print: MbistRetention wait " & mS_Time & " ms*"
    'TheExec.Datalog.WriteComment "*print: MbistRetention Pins " & ApplyPins
    'TheExec.Datalog.WriteComment "*print: MbistRetention Volt " & Voltage_from_HW
    '    TheExec.Datalog.WriteComment "*print: MbistRetention Voltage " & Retintion_voltage & " V*"
    TheExec.Datalog.WriteComment "********************************************************"
    DebugPrintFunc ""
    ''
    '--------- Ramp up for retention voltage ------'
    For Each Site In TheExec.Sites
        For i = 0 To RampDown_Step - 1
            For j = 0 To Retention_Pins_count - 1
                If Flag_ApplyPower(j) = True Then
                    If i = RampDown_Step - 1 Then
                        TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main = Original_voltage(j)
                    Else
                        TheHdw.DCVS.Pins(Retention_Pins_Ary(j)).Voltage.Main = Retention_Voltage(j) + DropVoltage_perStep(j) * i
                    End If
                End If
            Next j
            TheHdw.Wait RampDown_Time / RampDown_Step    'remove extra wait time for each step
        Next i
    Next Site

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Get_Tname_FromFlowSheet(Flow_Instance_Tname As String, HIO_PinName_Updated As Boolean) As String
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Get_Tname_FromFlowSheet"

    Dim Current_FlowSheet_Loc As String
    Dim Current_Instance_Tname As String

    HIO_PinName_Updated = False
    Current_FlowSheet_Loc = TheExec.Flow.Raw.SheetInRun + ":" + CStr(TheExec.Flow.Raw.GetCurrentLineNumber + 5)
    Current_Instance_Tname = Application.Worksheets(TheExec.Flow.Raw.SheetInRun).Range("I" & CStr(TheExec.Flow.Raw.GetCurrentLineNumber + 5)).Value
    '    Debug.Print Current_FlowSheet_Loc + ":" + Current_Instance_Tname
    If Current_Instance_Tname <> "" Then
        '        Get_Tname_FromFlowSheet = Current_Instance_Tname
        '        Flow_Instance_Tname = Current_Instance_Tname

        If UCase(Current_Instance_Tname) Like "*HAC*" Then Exit Function
        If UCase(Current_Instance_Tname) Like "*HIO*" Then
            Dim TNameSeg() As String
            ReDim TNameSeg(9) As String
            Dim SetupName As String
            Dim X_ApplyToPin As String
            Dim Y_ApplyToPin As String

            TNameSeg = Split(Current_Instance_Tname, "_")
            SetupName = TheExec.DevChar.Setups.ActiveSetupName
            With TheExec.DevChar.Setups(SetupName)
                If .Shmoo.Axes.Count > 1 Then
                    X_ApplyToPin = .Shmoo.Axes.Item(tlDevCharShmooAxis_X).ApplyTo.Pins
                    Y_ApplyToPin = .Shmoo.Axes.Item(tlDevCharShmooAxis_Y).ApplyTo.Pins
                    TNameSeg(5) = Replace(X_ApplyToPin & "&" & Y_ApplyToPin, ",", "&")
                    '                    TNameSeg(5) = Replace(TNameSeg(5), ",", "&")
                Else
                    X_ApplyToPin = .Shmoo.Axes.Item(tlDevCharShmooAxis_X).ApplyTo.Pins
                    TNameSeg(5) = Replace(X_ApplyToPin, ",", "&")
                End If
            End With
            Flow_Instance_Tname = Merge_TName(TNameSeg)
            HIO_PinName_Updated = True
        Else
            Flow_Instance_Tname = Current_Instance_Tname
            '            HIO_PinName_Updated = False
        End If
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function



'Public Function CZ_Shmoo_Info(ByRef SetupName As String, ByRef X_StepName As String, ByRef X_ApplyToPin As String, ByRef X_RangeFrom As Double, ByRef X_RangeTo As Double, ByRef X_StepSize As Double, _
'                              Optional ByRef Y_StepName As String, Optional ByRef Y_ApplyToPin As String, Optional ByRef Y_RangeFrom As Double, Optional ByRef Y_RangeTo As Double, Optional ByRef Y_StepSize As Double) As String
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "CZ_Shmoo_Info"
'
'    If TheExec.DevChar.Setups.IsRunning = True Then
'        SetupName = TheExec.DevChar.Setups.ActiveSetupName
'        With TheExec.DevChar.Setups(SetupName)
'            If .Shmoo.Axes.Count > 1 Then
'                X_StepName = .Shmoo.Axes(tlDevCharShmooAxis_X).StepName
'                Y_StepName = .Shmoo.Axes(tlDevCharShmooAxis_Y).StepName
'                X_ApplyToPin = .Shmoo.Axes.Item(tlDevCharShmooAxis_X).ApplyTo.Pins
'                Y_ApplyToPin = .Shmoo.Axes.Item(tlDevCharShmooAxis_Y).ApplyTo.Pins
'                X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.from
'                Y_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Range.from
'                X_RangeTo = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.To
'                Y_RangeTo = .Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Range.To
'                X_StepSize = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.StepSize
'                Y_StepSize = .Shmoo.Axes(tlDevCharShmooAxis_Y).Parameter.Range.StepSize
'            Else
'                X_StepName = .Shmoo.Axes(tlDevCharShmooAxis_X).StepName
'                X_ApplyToPin = .Shmoo.Axes.Item(tlDevCharShmooAxis_X).ApplyTo.Pins
'                X_RangeFrom = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.from
'                X_RangeTo = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.To
'                X_StepSize = .Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.Range.StepSize
'            End If
'        End With
'    End If
'
'    Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'Public Function CZ_TNum_Increment(Optional Flow_TestNumber As String) As String
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "CZ_TNum_Increment"
'
'    Dim Site      As Variant
'
'    If TheExec.DevChar.Setups.IsRunning = True Then
'        For Each Site In TheExec.Sites.Active: Next Site
'        TheExec.Flow.TestNumber = TheExec.Sites.Item(Site).TestNumber + 1
'        '        Debug.Print "TestNumber is: " & CStr(TheExec.sites.Item(site).TestNumber)
'    End If
'
'    Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function



'Public Function CZ_TNum_Decrement(Optional Flow_TestNumber As String) As String
'    On Error GoTo ErrHandler
'    Dim funcName  As String:: funcName = "CZ_TNum_Decrement"
'
'    Dim Site      As Variant
'
'    If TheExec.DevChar.Setups.IsRunning = True Then
'        For Each Site In TheExec.Sites.Active: Next Site
'        TheExec.Flow.TestNumber = TheExec.Sites.Item(Site).TestNumber - 1
'        '        Debug.Print "TestNumber is: " & CStr(TheExec.sites.Item(site).TestNumber)
'    End If
'
'    Exit Function
'ErrHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function


Public Function Spec_Evaluate_DC(ByVal temp_pin_info As String) As String
    On Error GoTo ErrHandler
    Dim funcName  As String:: funcName = "Spec_Evaluate_DC"

    Dim temp_pin_name, calc_info As String
    Dim temp_pininfo_arr() As String
    Dim i         As Integer
    Dim fake_calc_info As String

    If InStr(temp_pin_info, "+") > 0 Or InStr(temp_pin_info, "-") > 0 Or InStr(temp_pin_info, "*") > 0 Or InStr(temp_pin_info, "/") > 0 Then
        calc_info = temp_pin_info

        temp_pin_info = Replace(temp_pin_info, "(", "")
        temp_pin_info = Replace(temp_pin_info, ")", "")
        temp_pin_info = Replace(temp_pin_info, "+", "~")
        temp_pin_info = Replace(temp_pin_info, "-", "~")
        temp_pin_info = Replace(temp_pin_info, "*", "~")
        temp_pin_info = Replace(temp_pin_info, "/", "~")

        fake_calc_info = Replace(calc_info, "(", " ( ")
        fake_calc_info = Replace(fake_calc_info, ")", " ) ")
        fake_calc_info = Replace(fake_calc_info, "+", " + ")
        fake_calc_info = Replace(fake_calc_info, "-", " - ")
        fake_calc_info = Replace(fake_calc_info, "*", " * ")
        fake_calc_info = Replace(fake_calc_info, "/", " / ")
        fake_calc_info = " " & fake_calc_info & " "

        temp_pininfo_arr = Split(temp_pin_info, "~")

        For i = 0 To UBound(temp_pininfo_arr)
            If InStr(temp_pininfo_arr(i), "_") <> 0 Then
                temp_pin_name = temp_pininfo_arr(i)
                temp_pininfo_arr(i) = CStr(TheExec.Specs.DC.Item(Mid(temp_pininfo_arr(i), 2)).ContextValue)
                calc_info = Replace(fake_calc_info, " " & temp_pin_name & " ", temp_pininfo_arr(i), , 1)
            End If
        Next i
        Spec_Evaluate_DC = CStr(Evaluate(calc_info))
    Else
        If (InStr(temp_pin_info, "_") = 1) Then
            If TheExec.Specs.DC.Contains(Mid(temp_pin_info, 2)) Then
                Spec_Evaluate_DC = CStr(TheExec.Specs.DC.Item(Mid(temp_pin_info, 2)).ContextValue)
            Else
                Spec_Evaluate_DC = CStr(temp_pin_info)
            End If
        Else
            Spec_Evaluate_DC = CStr(temp_pin_info)
        End If
    End If

    Exit Function
ErrHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ": please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
