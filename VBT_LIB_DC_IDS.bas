Attribute VB_Name = "VBT_LIB_DC_IDS"
Option Explicit
'Revision History:
'V0.0 initial bring up

'IDS meas pinlistdata
Public IDS_meas As New PinListData

'Delta IDS for iEDA
Public gS_delta_IDS_pcpu As New SiteVariant
Public gS_delta_IDS_ecpu As New SiteVariant
Public gS_delta_IDS_gpu As New SiteVariant
Public gS_delta_IDS_dcs_ddr As New SiteVariant
Public gS_delta_IDS_cpu_sram As New SiteVariant
Public gS_delta_IDS_ave As New SiteVariant

'0826_SPI_IDS measure
Private HexVS_data_SPI_IDS As New PinListData
Private UVS_Hi_data_SPI_IDS As New PinListData
Private UVS_Lo_data_SPI_IDS As New PinListData

'==================================20180928  global variable for eFuse
Public All_Power_data_IDS_GB As New PinListData
''Start, IDS INFO for Efuse and BinCut - Carter, 20190829
Public ids_info_ary() As IDS_INFO
Public gl_IDS_INFO_Dic As New Scripting.Dictionary
Type IDS_INFO
    Pin As String
    Pat As String
    LoLimit As String
    HiLimit As String
    MeasureValue As New SiteDouble
End Type

Type AutoRange_Info
    PinName As String
    SlotType As String
    MergeType As String
    Init_step As Long
    Range_List() As Double
    Accuracy_List() As Double
    WaitTime_List() As Double
    CurrentRange_Val As Double
End Type

''End, IDS INFO for Efuse and BinCut - Carter, 20190829

' This module should be used for VBT Tests.  All functions in this module
' will be available to be used from the Test Instance sheet.
' Additional modules may be added as needed (all starting with "VBT_").
'
' The required signature for a VBT Test is:
'
' Public Function FuncName(<arglist>) As Long
'   where <arglist> is any list of arguments supported by VBT Tests.
'
' See online help for supported argument types in VBT Tests.
'
'
' It is highly suggested to use error handlers in VBT Tests.  A sample
' VBT Test with a suggeseted error handler is shown below:
'
' Function FuncName() As Long
'     On Error GoTo errHandler
'
'     Exit Function
' errHandler:
'     If AbortTest Then Exit Function Else Resume Next
' End Function

Public Function DCVS_IDD_dynamic(Power_pin_HexVs As PinList, Power_pin_HDVS As PinList, infi_pat As Pattern, Current_range As Double, _
                             Optional Validating_ As Boolean) As Long 'Carter, 20190315
Dim MeasureI_hexvs As New PinListData
Dim MeasureI_hdvs As New PinListData
On Error GoTo errHandler

If Validating_ Then 'Carter, 20190315
        Call PrLoadPattern(infi_pat.Value)
        Exit Function    ' Exit after validation
End If
    
Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlUnpowered)  'SEC DRAM
'DisconnectVDDCA 'SEC DRAM
With TheHdw.DCVS.Pins(Power_pin_HexVs)
    .Meter.mode = tlDCVSMeterCurrent
    .Meter.CurrentRange.Value = Current_range
    .Gate = True
End With

With TheHdw.DCVS.Pins(Power_pin_HDVS)
    .Meter.mode = tlDCVSMeterCurrent
    .Meter.CurrentRange.Value = Current_range
    .Gate = True
End With

TheHdw.Wait 0.003
TheHdw.Patterns(infi_pat).Load
TheHdw.Patterns(infi_pat).start

'''Strobe the meter on the VCC pin and store it in an pinlistdata variable defined''''
MeasureI_hexvs = TheHdw.DCVS.Pins(Power_pin_HexVs).Meter.Read(tlStrobe, 1, 0.001)
MeasureI_hdvs = TheHdw.DCVS.Pins(Power_pin_HDVS).Meter.Read(tlStrobe, 1, 0.001)


TheHdw.Digital.Patgen.Halt

''''Setup OFFLINE Simulation by stuffing the pinlistdata variable with simulation data'''''''

    If TheExec.TesterMode = testModeOffline Then
        MeasureI_hexvs.Value = 0.06 + Rnd * 0.005
        MeasureI_hdvs.Value = 0.06 - Rnd * 0.005
    End If
''''''''''DATALOG RESULTS''''''''''''''''''''''''''''''''''

TheExec.Flow.TestLimit resultVal:=MeasureI_hexvs, Unit:=unitAmp, ForceResults:=tlForceFlow
TheExec.Flow.TestLimit resultVal:=MeasureI_hdvs, Unit:=unitAmp, ForceResults:=tlForceFlow


DebugPrintFunc infi_pat.Value

Exit Function
errHandler:
        TheExec.AddOutput "Error in the VBT Icc"
        If AbortTest Then Exit Function Else Resume Next
End Function







Public Function DCVS_IDS_main_auto_range_and_measure(CorePower_Pin As String, _
                                                     OtherPower_Pin As String, _
                                                     Power_data As PinListData, _
                                                     repeat_count As Long, _
                                                     FlowLimitForInitIRange As Boolean, _
                                            Optional Search_Step As String, _
                                            Optional OtherPowerAutoRange As Boolean, _
                                            Optional HexInitIRange1A_Pins As PinList, _
                                            Optional HexInitIRange100mA_Pins As PinList, _
                                            Optional UVSInitIRange800mA_Pins As PinList, _
                                            Optional UVSInitIRange200mA_Pins As PinList, _
                                            Optional UVSInitIRange20mA_Pins As PinList, _
                                            Optional UVSInitIRange2mA_Pins As PinList, _
                                            Optional UVSInitIRange200uA_Pins As PinList, Optional debug_print_pins As String)

    Dim i As Long, j As Long, All_Power_Pin As String
    Dim site As Variant, Pin As Variant, Val As Double
    Dim k As Long
    Dim Tname As String
    Dim Vmain As Double
    Dim p As Variant
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim CorePowerPin_Ary() As String, CorePowerPin_Cnt As Long
    Dim OtherPowerPin_Ary() As String, OtherPowerPin_Cnt As Long
    
    Dim ChannelType As Long
    Dim Channels() As String, NumberChannels As Long
    Dim NumberSites As Long, Error As String
    
    On Error GoTo errHandler
    
    ' Get the limits info
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    
    ' if no Use-Limits on this test, FlowLimitsInfo is nothing
    If FlowLimitsInfo Is Nothing Then
        TheExec.AddOutput "Could not get the limits info", vbRed, True
        Exit Function
    End If

    Dim Val_Hi() As String
    Dim Val_Lo() As String
    FlowLimitsInfo.GetHighLimits Val_Hi
    FlowLimitsInfo.GetLowLimits Val_Lo
    
    '20161121 create pin dictionary are selected init i range
    Dim Dict_Hex1AIRangePins As Scripting.Dictionary
    Set Dict_Hex1AIRangePins = New Scripting.Dictionary
    Dim Dict_Hex100mAIRangePins As Scripting.Dictionary
    Set Dict_Hex100mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVS800mAIRangePins As Scripting.Dictionary
    Set Dict_UVS800mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVS200mAIRangePins As Scripting.Dictionary
    Set Dict_UVS200mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVS20mAIRangePins As Scripting.Dictionary
    Set Dict_UVS20mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVS2mAIRangePins As Scripting.Dictionary
    Set Dict_UVS2mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVS200uAIRangePins As Scripting.Dictionary
    Set Dict_UVS200uAIRangePins = New Scripting.Dictionary
    
    If (FlowLimitForInitIRange = False) Then
        If HexInitIRange1A_Pins <> "" Then
            TheExec.DataManager.DecomposePinList HexInitIRange1A_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_Hex1AIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If HexInitIRange100mA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList HexInitIRange100mA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_Hex100mAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If UVSInitIRange800mA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVSInitIRange800mA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVS800mAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If UVSInitIRange200mA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVSInitIRange200mA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVS200mAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If UVSInitIRange20mA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVSInitIRange20mA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVS20mAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If UVSInitIRange2mA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVSInitIRange2mA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVS2mAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If UVSInitIRange200uA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVSInitIRange200uA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVS200uAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
    End If

    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim Merge_Type, Slot_Type As String
    Dim Split_Ary() As String
    Dim SattleTime As Double
    Dim WaitTime As Double
    Dim p_hexvs As String
    Dim p_uvs As String
    Dim A_HexVS() As String
    Dim A_UVS() As String
    Dim HexVS_Power_data As New PinListData
    Dim UVS_Power_data As New PinListData
    Dim IDS_ini_Current_range() As Double
    
    Dim SlotType As Scripting.Dictionary
    Set SlotType = New Scripting.Dictionary
    Dim InitStep As Scripting.Dictionary
    Set InitStep = New Scripting.Dictionary
    Dim PinVal As New PinData
    Dim DropRngSite As New SiteBoolean
    Dim AutoRangePin_Ary() As String
    
    Dim range_ary() As AutoRange_Info
    
    '20160113: debug with bruce to fix NC issue -- start
    TheExec.DataManager.DecomposePinList CorePower_Pin, CorePowerPin_Ary, CorePowerPin_Cnt
    TheExec.DataManager.DecomposePinList OtherPower_Pin, OtherPowerPin_Ary, OtherPowerPin_Cnt
    
    For i = 0 To CorePowerPin_Cnt - 1
        If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) <> "N/C" Then All_Power_Pin = All_Power_Pin & "," & CorePowerPin_Ary(i)
    Next i
    
    For i = 0 To OtherPowerPin_Cnt - 1
        If TheExec.DataManager.ChannelType(OtherPowerPin_Ary(i)) <> "N/C" Then All_Power_Pin = All_Power_Pin & "," & OtherPowerPin_Ary(i)
    Next i
    
    If All_Power_Pin <> "" Then All_Power_Pin = Right(All_Power_Pin, Len(All_Power_Pin) - 1)
    '20160113: debug with bruce to fix NC issue -- end
    
    Pin_Ary = Split(All_Power_Pin, ",")
    
'    All_Power_Pin = CorePower_Pin & "," & OtherPower_Pin
'    TheExec.DataManager.DecomposePinList All_Power_Pin, Pin_Ary, Pin_Cnt
    ReDim IDS_ini_Current_range(UBound(Pin_Ary)) As Double
    
    WaitTime = 100 * us
    
    ' Set init IRange
    For i = 0 To UBound(Pin_Ary)
        Merge_Type = TheExec.DataManager.ChannelType(Pin_Ary(i))
        SlotType.Add Pin_Ary(i), GetInstrument(Pin_Ary(i), 0)
        IDS_ini_Current_range(i) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.CurrentRange.Value
'        Debug.Print Pin_Ary(i) & ":" & SlotType(Pin_Ary(i))

        If LCase(SlotType(Pin_Ary(i))) = "hexvs" Then
            p_hexvs = p_hexvs & "," & Pin_Ary(i)
            If (FlowLimitForInitIRange = True) Then
                If Val_Hi(i) = "" Then Val = IDS_ini_Current_range(i) Else Val = Abs(Val_Hi(i)) 'if pins without limit use current range
                
                If Val < 0.01 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.01, 0.01  ' HexVS
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.05
                    SattleTime = 100 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 3
                    
                ElseIf Val < 0.1 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.1, 0.1    ' HexVS
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.1
                    SattleTime = 10 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 2
                    
                ElseIf Val < 1 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 1, 1    ' HexVS
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 1
                    SattleTime = 1 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 1
                    
                ElseIf Val < 15 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 15, 1   ' HexVS
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 15
                    SattleTime = 100 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                    
                Else
                    Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                    If Val > TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max Then Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges Val, Val
                    SattleTime = 100 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                    
                End If
            Else
                If Dict_Hex1AIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 1 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 1
                ElseIf Dict_Hex100mAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 10 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 2
                Else
                    Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                    If Val > TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max Then Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges Val, Val
                    SattleTime = 100 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                End If
            End If
            
        ElseIf LCase(SlotType(Pin_Ary(i))) = "vhdvs" Then
            p_uvs = p_uvs & "," & Pin_Ary(i)
            If (FlowLimitForInitIRange = True) Then
                If Val_Hi(i) = "" Then Val = IDS_ini_Current_range(i) Else Val = Abs(Val_Hi(i)) 'if pins without limit use current range
    
                If Val < 0.000004 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.000004, 0.000004
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.000004
                    SattleTime = 18 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 6
                    
                ElseIf Val < 0.00002 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.00002, 0.00002
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.00002
                    SattleTime = 4 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 5
                    
                ElseIf Val < 0.0002 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.0002, 0.0002
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.0002
                    SattleTime = 4 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 4
                    
                ElseIf Val < 0.002 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.002, 0.002
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.002
                    SattleTime = 3.5 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 3
                    
                ElseIf Val < 0.02 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.02, 0.02
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.02
                    SattleTime = 540 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 2
                    
                ElseIf Val < 0.04 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.04, 0.04
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.04
                    SattleTime = 260 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 1
                    
                ElseIf Val < 0.2 Then
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges 0.2, 0.2
                    TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.Value = 0.2
                    SattleTime = 210 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                
                Else
                    Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                    If Val > TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max Then Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges Val, Val
                    SattleTime = 210 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                End If
            Else
                If Dict_UVS800mAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 210 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                ElseIf Dict_UVS200mAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 210 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                ElseIf Dict_UVS20mAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 540 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 2
                ElseIf Dict_UVS2mAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 3.5 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 3
                ElseIf Dict_UVS200uAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 4 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 4
                Else
                    Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value
                    If Val > TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max Then Val = TheHdw.DCVS.Pins(Pin_Ary(i)).CurrentRange.max
                    TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges Val, Val
                    SattleTime = 210 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
    '                step_ary(i) = 0
                    InitStep.Add Pin_Ary(i), 0
                End If
            End If
        End If
'        Debug.Print i & ":" & CorePower_Pin_Ary(i)
    Next i
    
    If HexInitIRange1A_Pins <> "" Then
        TheHdw.DCVS.Pins(HexInitIRange1A_Pins).SetCurrentRanges 1, 1    ' HexVS
        TheHdw.DCVS.Pins(HexInitIRange1A_Pins).CurrentLimit.Source.FoldLimit.Level.Value = 1
    End If
    If HexInitIRange100mA_Pins <> "" Then
        TheHdw.DCVS.Pins(HexInitIRange100mA_Pins).SetCurrentRanges 0.1, 0.1    ' HexVS
        TheHdw.DCVS.Pins(HexInitIRange100mA_Pins).CurrentLimit.Source.FoldLimit.Level.Value = 0.1
    End If
    If UVSInitIRange800mA_Pins <> "" Then
        TheHdw.DCVS.Pins(UVSInitIRange800mA_Pins).SetCurrentRanges 0.8, 0.8
        TheHdw.DCVS.Pins(UVSInitIRange800mA_Pins).CurrentRange.Value = 0.8
    End If
    If UVSInitIRange200mA_Pins <> "" Then
        TheHdw.DCVS.Pins(UVSInitIRange200mA_Pins).SetCurrentRanges 0.2, 0.2
        TheHdw.DCVS.Pins(UVSInitIRange200mA_Pins).CurrentRange.Value = 0.2
    End If
    If UVSInitIRange20mA_Pins <> "" Then
        TheHdw.DCVS.Pins(UVSInitIRange20mA_Pins).SetCurrentRanges 0.02, 0.02
        TheHdw.DCVS.Pins(UVSInitIRange20mA_Pins).CurrentRange.Value = 0.02
    End If
    If UVSInitIRange2mA_Pins <> "" Then
        TheHdw.DCVS.Pins(UVSInitIRange2mA_Pins).SetCurrentRanges 0.002, 0.002
        TheHdw.DCVS.Pins(UVSInitIRange2mA_Pins).CurrentRange.Value = 0.002
    End If
    If UVSInitIRange200uA_Pins <> "" Then
        TheHdw.DCVS.Pins(UVSInitIRange200uA_Pins).SetCurrentRanges 0.0002, 0.0002
        TheHdw.DCVS.Pins(UVSInitIRange200uA_Pins).CurrentRange.Value = 0.0002
    End If
    
    If p_hexvs <> "" Then p_hexvs = Right(p_hexvs, Len(p_hexvs) - 1)
    If p_uvs <> "" Then p_uvs = Right(p_uvs, Len(p_uvs) - 1)
    A_HexVS = Split(p_hexvs, ",")
    A_UVS = Split(p_uvs, ",")
    
    TheHdw.Wait 0.01 'add 10ms.
    TheHdw.Wait WaitTime
    
    If p_hexvs <> "" Then HexVS_Power_data = TheHdw.DCVS.Pins(p_hexvs).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
    If p_uvs <> "" Then UVS_Power_data = TheHdw.DCVS.Pins(p_uvs).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
    
    If OtherPowerAutoRange = True Then
        ReDim AutoRangePin_Ary(UBound(Pin_Ary))
        AutoRangePin_Ary = Pin_Ary
    Else
        ReDim AutoRangePin(UBound(CorePowerPin_Ary))
        AutoRangePin_Ary = CorePowerPin_Ary
    End If
    
    Dim Stop_Step, StepNo As Integer
    If Search_Step = "" Then
        Stop_Step = 5
    ElseIf (CLng(Search_Step) >= 6) Then
        Stop_Step = 6
    Else
        Stop_Step = CLng(Search_Step)
    End If
    
    For j = 1 To Stop_Step
        WaitTime = 260 * us
        For i = 0 To UBound(AutoRangePin_Ary)
            If LCase(SlotType(AutoRangePin_Ary(i))) = "hexvs" Then
                
                StepNo = j + InitStep(AutoRangePin_Ary(i))
                PinVal = HexVS_Power_data.Pins(AutoRangePin_Ary(i)).Abs
                
                Select Case StepNo
                    Case 1: '15A => 1A
                            DropRngSite = PinVal.compare(LessThan, 1 - (0.05 + 0.15)) 'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 1, 1 ': thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 1
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 1 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 2: '1A => 100mA
                            DropRngSite = PinVal.compare(LessThan, 0.1 - (0.01 + 0.005)) 'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.1, 0.1 ': thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 1  'prevent ifold accuarcy issue due to false alarm
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 10 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 3: '100mA => 10mA
                            DropRngSite = PinVal.compare(LessThan, 0.01 - 0.0005) 'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.01, 0.01 ': thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.05
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 200 * ms ' prevent alarm issue +100ms settle time.
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                End Select
                
            ElseIf LCase(SlotType(AutoRangePin_Ary(i))) = "vhdvs" Then
            
                StepNo = j + InitStep(AutoRangePin_Ary(i))
                PinVal = UVS_Power_data.Pins(AutoRangePin_Ary(i)).Abs
                Dim maxCurrentRange As Double
                Select Case StepNo
                    Case 1: '' max current range => 700 mA
                            maxCurrentRange = TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.max
                            If maxCurrentRange > 0.7 Then
                                Dim AccuracyVal As Double
                                If maxCurrentRange > 5.3 Then ''for 5.6 A current range
                                    AccuracyVal = maxCurrentRange * 0.007 + 0.04
                                ElseIf maxCurrentRange > 2.5 Then ''for 2.8 A current range
                                    AccuracyVal = maxCurrentRange * 0.007 + 0.02
                                ElseIf maxCurrentRange > 1.1 Then ''for 1.4 A current range
                                    AccuracyVal = maxCurrentRange * 0.007 + 0.01
                                ElseIf maxCurrentRange > 0.75 Then ''for 0.8 A current range
                                    AccuracyVal = maxCurrentRange * 0.007 + 0.006
                                End If
                                DropRngSite = PinVal.compare(LessThan, 0.7 - AccuracyVal) 'Next IRange - Accuracy
                                If DropRngSite.Any(True) Then
                                    TheExec.sites.Selected = DropRngSite
                                    TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.7, 0.7: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.7 ': thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.7
                                    TheExec.sites.Selected = True
                                End If
                            End If
                            SattleTime = 260 * us
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 2: '700mA => 200mA (else 400mA => 200mA)
                            maxCurrentRange = TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.max
                            If maxCurrentRange > 0.7 Then
                            DropRngSite = PinVal.compare(LessThan, 0.2 - 0.02) 'Next IRange - Accuracy
                                If DropRngSite.Any(True) Then
                                    TheExec.sites.Selected = DropRngSite
                                    TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.2, 0.2: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.2 ': thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
                                    TheExec.sites.Selected = True
                                End If
                            Else
                                If maxCurrentRange > 0.2 Then ''400mA => 200mA
                                    DropRngSite = PinVal.compare(LessThan, 0.2 - (maxCurrentRange * 0.007 + 0.0024)) 'Next IRange - Accuracy
                                    If DropRngSite.Any(True) Then
                                        TheExec.sites.Selected = DropRngSite
                                        TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.2, 0.2: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.2 ': thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.2
                                        TheExec.sites.Selected = True
                                    End If
                                End If
                            End If
                            
'                            DropRngSite = PinVal.compare(LessThan, 0.7 - (0.02)).LogicalAnd(PinVal.compare(GreaterThan, 0.2 - (0.005)))
'                            If DropRngSite.Any(True) Then
'
'                                If TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.max > 0.7 Then
'                                    TheExec.Sites.Selected = DropRngSite
'                                    TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.7, 0.7: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.7: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.7
'                                    TheExec.Sites.Selected = True
'                                End If
'                            End If
                            SattleTime = 260 * us
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 3: '200mA => 20mA
                            DropRngSite = PinVal.compare(LessThan, 0.02 - (0.0012 + 0.001)) 'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.02, 0.02: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.02 ': thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 260 * us
                            If SattleTime > WaitTime Then WaitTime = SattleTime
'                    Case 3: '40mA =>20mA
'                            DropRngSite = PinVal.compare(LessThan, 0.02 - (0.00024 + 0.0003)) 'Next IRange - Accuracy
'                            If DropRngSite.Any(True) Then
'                                TheExec.Sites.Selected = DropRngSite
'                                TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.02, 0.02: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.02:: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.02
'                                TheExec.Sites.Selected = True
'                            End If
'                            SattleTime = 540 * us
'                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 4: '20mA =>2mA
                            DropRngSite = PinVal.compare(LessThan, 0.002 - (0.00012 + 0.0001)) 'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.002, 0.002: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.002 ':: thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.002
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 3.5 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 5: '2mA =>200uA
                            DropRngSite = PinVal.compare(LessThan, 0.0002 - (0.000012 + 0.00001)) 'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.0002, 0.0002: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.0002 ':: thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.0002
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 4 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 6: '200uA =>20uA
                            DropRngSite = PinVal.compare(LessThan, 0.00002 - (0.0000012 + 0.000001)) 'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.00002, 0.00002: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.00002 ':: thehdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.00002
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 4 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
'                    Case 6: '20uA =>4uA
'                            DropRngSite = PinVal.compare(LessThan, 0.000004 - (0.00000012 + 0.0000001)) 'Next IRange - Accuracy
'                            If DropRngSite.Any(True) Then
'                                TheExec.Sites.Selected = DropRngSite
'                                TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).SetCurrentRanges 0.000004, 0.000004: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value = 0.000004: TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentLimit.Source.FoldLimit.Level.Value = 0.000004
'                                TheExec.Sites.Selected = True
'                            End If
'                            SattleTime = 18 * ms
'                            If SattleTime > WaitTime Then WaitTime = SattleTime
                   
                End Select
                
            End If
            'Print current range to Eng monitor
'            For Each Site In TheExec.Sites
'                TheExec.Datalog.WriteComment "Site(" & Site & "), " & AutoRangePin_Ary(i) & ", Step " & j & ", Irange: " & TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value
'            Next Site
        Next i
        
        If StepNo = 6 Then j = Stop_Step
        
        TheHdw.Wait WaitTime
        
        ''Add measurement points to prevent error 20171011 (M9)
        If p_hexvs <> "" Then
            If TheHdw.Alarms.Check Then
                HexVS_Power_data = TheHdw.DCVS.Pins(p_hexvs).Meter.Read(tlStrobe, 2000, , tlDCVSMeterReadingFormatAverage)
            Else
                HexVS_Power_data = TheHdw.DCVS.Pins(p_hexvs).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
            End If
        End If
        If p_uvs <> "" Then UVS_Power_data = TheHdw.DCVS.Pins(p_uvs).Meter.Read(tlStrobe, 1, , tlDCVSMeterReadingFormatAverage)
        '-------------------------------------------debug print
        If debug_print_pins <> "" Then
            For i = 0 To UBound(AutoRangePin_Ary)
                If InStr(LCase(debug_print_pins), LCase(AutoRangePin_Ary(i))) > 0 Then
                    For Each site In TheExec.sites
                        If LCase(SlotType(AutoRangePin_Ary(i))) = "hexvs" Then
                            TheExec.Datalog.WriteComment "Site(" & site & "), " & AutoRangePin_Ary(i) & ", Step " & j & ", Irange: " & TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value & ", Current: " & HexVS_Power_data.Pins(AutoRangePin_Ary(i)).Value(site)
                        ElseIf LCase(SlotType(AutoRangePin_Ary(i))) = "vhdvs" Then
                            TheExec.Datalog.WriteComment "Site(" & site & "), " & AutoRangePin_Ary(i) & ", Step " & j & ", Irange: " & TheHdw.DCVS.Pins(AutoRangePin_Ary(i)).CurrentRange.Value & ", Current: " & UVS_Power_data.Pins(AutoRangePin_Ary(i)).Value(site)
                        End If
                    Next site
                End If
            Next i
        End If
        '-------------------------------------------
    Next j
    
    For i = 0 To UBound(A_HexVS)
        Power_data.AddPin (A_HexVS(i))
        Power_data.Pins(A_HexVS(i)) = HexVS_Power_data.Pins(A_HexVS(i))
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                Power_data.Pins(A_HexVS(i)).Value(site) = 0.01 + Rnd() * 0.0001
            Next site
        End If
    Next i
    
    For i = 0 To UBound(A_UVS)
        Power_data.AddPin (A_UVS(i))
        Power_data.Pins(A_UVS(i)) = UVS_Power_data.Pins(A_UVS(i))
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                Power_data.Pins(A_UVS(i)).Value(site) = 0.0005 + Rnd() * 0.0001
            Next site
        End If
    Next i
  

    For i = 0 To CorePowerPin_Cnt - 1: For j = 0 To repeat_count - 1
        If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) <> "N/C" Then

            Tname = TheExec.DataManager.instanceName & "_" & CorePowerPin_Ary(i) & "_" & j    'add pin name Aruba 2017/12/28
            Vmain = Format(TheHdw.DCVS.Pins(Power_data.Pins(CorePowerPin_Ary(i))).Voltage.Main.Value, "0.00")
'
            If TPModeAsCharz_GLB Then
                TheExec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", ForceVal:=Vmain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            Else
                TheExec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=Vmain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                TheExec.Datalog.WriteComment "Current I Range: " & CorePowerPin_Ary(i) & "--->" & TheHdw.DCVS.Pins(CorePowerPin_Ary(i)).Meter.CurrentRange.Value
            End If
        End If
    Next j: Next i

    For k = 0 To OtherPowerPin_Cnt - 1
        If TheExec.DataManager.ChannelType(OtherPowerPin_Ary(k)) <> "N/C" Then

            Tname = TheExec.DataManager.instanceName & "_" & OtherPowerPin_Ary(k)        'add pin name Aruba 2017/12/28

            Vmain = Format(TheHdw.DCVS.Pins(Power_data.Pins(OtherPowerPin_Ary(k))).Voltage.Main.Value, "0.00")

            If TPModeAsCharz_GLB Then
                TheExec.Flow.TestLimit resultVal:=Power_data.Pins(OtherPowerPin_Ary(k)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", ForceVal:=Vmain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            Else
                TheExec.Flow.TestLimit resultVal:=Power_data.Pins(OtherPowerPin_Ary(k)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:=Tname, ForceVal:=Vmain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                TheExec.Datalog.WriteComment "Current I Range: " & OtherPowerPin_Ary(k) & "--->" & TheHdw.DCVS.Pins(OtherPowerPin_Ary(k)).Meter.CurrentRange.Value
            End If
        End If
    Next k

    'recover range setup
    
    'Pin_Cnt = CorePowerPin_Cnt + OtherPowerPin_Cnt
    
    For i = 0 To UBound(Pin_Ary)
        TheHdw.DCVS.Pins(Pin_Ary(i)).SetCurrentRanges IDS_ini_Current_range(i), IDS_ini_Current_range(i)
    Next i
    
    TheHdw.Wait 0.003
'    TheHdw.Digital.ApplyLevelsTiming False, True, False, tlPowered    'SEC DRAM

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment ("In DCVS_auto_range: " & err.Description)
    If AbortTest Then Exit Function Else Resume Next
    Resume Next
End Function

Public Function DCVI_IDS_main_auto_range_and_measure(CorePower_Pin As String, _
                                                         Power_data As PinListData, _
                                                         repeat_count As Long, _
                                                         FlowLimitForInitIRange As Boolean, _
                                                         Optional Search_Step As String, _
                                                         Optional OtherPowerAutoRange As Boolean, _
                                                         Optional UVI80InitIRange1A_Pins As PinList, _
                                                         Optional UVI80InitIRange200mA_Pins As PinList, _
                                                         Optional UVI80InitIRange20mA_Pins As PinList, _
                                                         Optional UVI80InitIRange2mA_Pins As PinList, _
                                                         Optional UVI80InitIRange200uA_Pins As PinList, Optional debug_print_pins As String)

    Dim i As Long, j As Long, All_Power_Pin As String
    Dim site As Variant, Pin As Variant, Val As Double
    Dim k As Long
    Dim Tname As String
    Dim Vmain As Double
    Dim p As Variant
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim CorePowerPin_Ary() As String, CorePowerPin_Cnt As Long
    Dim OtherPowerPin_Ary() As String, OtherPowerPin_Cnt As Long
    
    Dim ChannelType As Long
    Dim Channels() As String, NumberChannels As Long
    Dim NumberSites As Long, Error As String
    
    Dim Powerpin_log As String ''20180315 Abel added
                                                                                                                                                                                                                                           
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "IDS_main_auto_range_and_measure"
    
    'Get the limits info
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    
    'if no Use-Limits on this test, FlowLimitsInfo is nothing
    If FlowLimitsInfo Is Nothing Then
        TheExec.AddOutput "Could not get the limits info", vbRed, True
        Exit Function
    End If

    Dim Val_Hi() As String
    Dim Val_Lo() As String
    FlowLimitsInfo.GetHighLimits Val_Hi
    FlowLimitsInfo.GetLowLimits Val_Lo

    '20161121 create pin dictionary are selected init i range
    Dim Dict_Hex1AIRangePins As Scripting.Dictionary
    Set Dict_Hex1AIRangePins = New Scripting.Dictionary
    Dim Dict_Hex100mAIRangePins As Scripting.Dictionary
    Set Dict_Hex100mAIRangePins = New Scripting.Dictionary

    Dim Dict_UVS800mAIRangePins As Scripting.Dictionary
    Set Dict_UVS800mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVS200mAIRangePins As Scripting.Dictionary
    Set Dict_UVS200mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVS20mAIRangePins As Scripting.Dictionary
    Set Dict_UVS20mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVS2mAIRangePins As Scripting.Dictionary
    Set Dict_UVS2mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVS200uAIRangePins As Scripting.Dictionary
    Set Dict_UVS200uAIRangePins = New Scripting.Dictionary
    
    ''For UVI80 Irange:1A/200mA/20mA/2mA/200uA (2017/7/31)
    Dim Dict_UVI80_1AIRangePins As Scripting.Dictionary
    Set Dict_UVI80_1AIRangePins = New Scripting.Dictionary
    Dim Dict_UVI80_200mAIRangePins As Scripting.Dictionary
    Set Dict_UVI80_200mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVI80_20mAIRangePins As Scripting.Dictionary
    Set Dict_UVI80_20mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVI80_2mAIRangePins As Scripting.Dictionary
    Set Dict_UVI80_2mAIRangePins = New Scripting.Dictionary
    Dim Dict_UVI80_200uAIRangePins As Scripting.Dictionary
    Set Dict_UVI80_200uAIRangePins = New Scripting.Dictionary
    
    If (FlowLimitForInitIRange = False) Then
        ''For UVI80 Irange:1A/200mA/20mA/2mA/200uA (2017/7/31)
        If UVI80InitIRange1A_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVI80InitIRange1A_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVI80_1AIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If UVI80InitIRange200mA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVI80InitIRange200mA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVI80_200mAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If UVI80InitIRange20mA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVI80InitIRange20mA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVI80_20mAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If UVI80InitIRange2mA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVI80InitIRange2mA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVI80_2mAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
        If UVI80InitIRange200uA_Pins <> "" Then
            TheExec.DataManager.DecomposePinList UVI80InitIRange200uA_Pins, Pin_Ary, Pin_Cnt
            For i = 0 To Pin_Cnt - 1
                Dict_UVI80_200uAIRangePins.Add Pin_Ary(i), i
            Next i
            Erase Pin_Ary
        End If
    End If

    Dim typesCount As Long
    Dim numericTypes() As Long
    Dim stringTypes() As String
    Dim Merge_Type, Slot_Type As String
    Dim Split_Ary() As String
    Dim SattleTime As Double
    Dim WaitTime As Double
    Dim p_hexvs As String
    Dim p_uvs As String
    Dim A_HexVS() As String
    Dim A_UVS() As String
    Dim HexVS_Power_data As New PinListData
    Dim UVS_Power_data As New PinListData

    ''For UVI80 Irange:1A/200mA/20mA/2mA/200uA (2017/7/31)
    Dim p_uvi80 As String
    Dim A_UVI80() As String
    Dim UVI80_Power_data As New PinListData

    Dim IDS_ini_Current_range() As Double

    Dim SlotType As Scripting.Dictionary
    Set SlotType = New Scripting.Dictionary
    Dim InitStep As Scripting.Dictionary
    Set InitStep = New Scripting.Dictionary
    Dim PinVal As New PinData
    Dim DropRngSite As New SiteBoolean
    Dim AutoRangePin_Ary() As String

    Dim range_ary() As AutoRange_Info
    
    TheExec.DataManager.DecomposePinList CorePower_Pin, CorePowerPin_Ary, CorePowerPin_Cnt

    For i = 0 To CorePowerPin_Cnt - 1
        If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) <> "N/C" Then All_Power_Pin = All_Power_Pin & "," & CorePowerPin_Ary(i)
    Next i

    If All_Power_Pin <> "" Then All_Power_Pin = Right(All_Power_Pin, Len(All_Power_Pin) - 1)

    Pin_Ary = Split(All_Power_Pin, ",")
    ReDim IDS_ini_Current_range(UBound(Pin_Ary)) As Double
    WaitTime = 100 * us

    ' Set init IRange
    For i = 0 To UBound(Pin_Ary)
        Merge_Type = TheExec.DataManager.ChannelType(Pin_Ary(i))
        SlotType.Add Pin_Ary(i), GetInstrument(Pin_Ary(i), 0)

        ''For UVI80 Irange:1A/200mA/20mA/2mA/200uA (2017/7/31)
        If TheExec.DataManager.ChannelType(Pin_Ary(i)) Like "*DCVI*" Then
           IDS_ini_Current_range(i) = TheHdw.DCVI.Pins(Pin_Ary(i)).current
        End If
'        Debug.Print Pin_Ary(i) & ":" & SlotType(Pin_Ary(i))
        If LCase(SlotType(Pin_Ary(i))) = "dc-07" Then
            p_uvi80 = p_uvi80 & "," & Pin_Ary(i)
            TheHdw.DCVI.Pins(Pin_Ary(i)).Meter.mode = tlDCVIMeterCurrent '20180115 Rick

            If (FlowLimitForInitIRange = True) Then
                If Val_Hi(i) = "" Then
                    Val = 0.002
                Else
                    Val = Abs(Val_Hi(i))
                End If

                If Val < 0.0002 Then
                    TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange 0.0002, 0.0002
                    TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.Value = 0.0002
                    SattleTime = 1.4 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 4

                ElseIf Val < 0.002 Then
                    TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange 0.002, 0.002
                    TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.Value = 0.002
                    SattleTime = 11 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 3

                ElseIf Val < 0.02 Then
                    TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange 0.02, 0.02
                    TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.Value = 0.02
                    SattleTime = 1.5 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 2

                ElseIf Val < 0.2 Then
                    TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange 0.2, 0.2
                    TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.Value = 0.2
                    SattleTime = 260 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 1

                ElseIf Val < 1 Then
                    TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange 1, 1
                    TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.Value = 1
                    SattleTime = 1.6 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0

                Else
                    Val = TheHdw.DCVI.Pins(Pin_Ary(i)).current
                    If Val > TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.max Then Val = TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.max
                    TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange Val, Val
                    SattleTime = 1.6 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                End If
            Else
                If Dict_UVI80_1AIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 1.6 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                ElseIf Dict_UVI80_200mAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 260 * us
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 1
                ElseIf Dict_UVI80_20mAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 1.5 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 2
                ElseIf Dict_UVI80_2mAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 11 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 3
                ElseIf Dict_UVI80_200uAIRangePins.Exists(Pin_Ary(i)) Then
                    SattleTime = 1.4 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 4
                Else
                    Val = TheHdw.DCVI.Pins(Pin_Ary(i)).current
                    If Val > TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.max Then Val = TheHdw.DCVI.Pins(Pin_Ary(i)).CurrentRange.max
                    TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange Val, Val
                    SattleTime = 1.6 * ms
                    If SattleTime > WaitTime Then WaitTime = SattleTime
                    InitStep.Add Pin_Ary(i), 0
                End If
            End If
        End If
'        Debug.Print i & ":" & CorePower_Pin_Ary(i)
    Next i

    If UVI80InitIRange1A_Pins <> "" Then
        TheHdw.DCVI.Pins(UVI80InitIRange1A_Pins).SetCurrentAndRange 1, 1
        TheHdw.DCVI.Pins(UVI80InitIRange1A_Pins).CurrentRange.Value = 1
    End If
    If UVI80InitIRange200mA_Pins <> "" Then
        TheHdw.DCVI.Pins(UVI80InitIRange200mA_Pins).SetCurrentAndRange 0.2, 0.2
        TheHdw.DCVI.Pins(UVI80InitIRange200mA_Pins).CurrentRange.Value = 0.2
    End If
    If UVI80InitIRange20mA_Pins <> "" Then
        TheHdw.DCVI.Pins(UVI80InitIRange20mA_Pins).SetCurrentAndRange 0.02, 0.02
        TheHdw.DCVI.Pins(UVI80InitIRange20mA_Pins).CurrentRange.Value = 0.02
    End If
    If UVI80InitIRange2mA_Pins <> "" Then
        TheHdw.DCVI.Pins(UVI80InitIRange2mA_Pins).SetCurrentAndRange 0.002, 0.002
        TheHdw.DCVI.Pins(UVI80InitIRange2mA_Pins).CurrentRange.Value = 0.002
    End If
    If UVI80InitIRange200uA_Pins <> "" Then
        TheHdw.DCVI.Pins(UVI80InitIRange200uA_Pins).SetCurrentAndRange 0.0002, 0.0002
        TheHdw.DCVI.Pins(UVI80InitIRange200uA_Pins).CurrentRange.Value = 0.0002
    End If

    If p_uvi80 <> "" Then p_uvi80 = Right(p_uvi80, Len(p_uvi80) - 1)
    A_UVI80 = Split(p_uvi80, ",")

    TheHdw.Wait 0.035 'add 10ms.
    TheHdw.Wait WaitTime

    ReDim AutoRangePin(UBound(CorePowerPin_Ary))
    AutoRangePin_Ary = CorePowerPin_Ary

    If p_uvi80 <> "" Then UVI80_Power_data = TheHdw.DCVI.Pins(p_uvi80).Meter.Read(tlStrobe, 1, , tlDCVIMeterReadingFormatAverage)
    
    Dim Stop_Step, StepNo As Integer
    If Search_Step = "" Then
        Stop_Step = 5
    ElseIf (CLng(Search_Step) >= 6) Then
        Stop_Step = 6
    Else
        Stop_Step = CLng(Search_Step)
    End If
    '========================================================================================auto range search
    For j = 1 To Stop_Step
        WaitTime = 260 * us
        For i = 0 To UBound(AutoRangePin_Ary)
            If LCase(SlotType(AutoRangePin_Ary(i))) = "dc-07" Then
                            StepNo = j + InitStep(AutoRangePin_Ary(i))
                            PinVal = UVI80_Power_data.Pins(AutoRangePin_Ary(i)).Abs

                Select Case StepNo
                    Case 1: '2A-->1A          '2A: ¡Ó(0.20% + 8mA + 300£gA/V)
                            DropRngSite = PinVal.compare(LessThan, 1 - ((0.004 + 0.008) * 2)) 'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVI.Pins(AutoRangePin_Ary(i)).SetCurrentAndRange 1, 1
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 1.6 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 2: '1A-->200mA      '1A:¡Ó(0.20% + 8mA + 300£gA/V)
                            DropRngSite = PinVal.compare(LessThan, 0.2 - ((0.002 + 0.008) * 2)) 'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVI.Pins(AutoRangePin_Ary(i)).SetCurrentAndRange 0.2, 0.2
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 1.6 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 3: '200mA => 20mA      '200mA:¡Ó(0.20% + 400£gA + 30£gA/V)
                            DropRngSite = PinVal.compare(LessThan, 0.02 - ((0.0004 + 0.0004) * 2))  'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVI.Pins(AutoRangePin_Ary(i)).SetCurrentAndRange 0.02, 0.02
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 260 * us
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 4: '20mA  => 2mA       '20mA:¡Ó(0.20% + 40£gA + 3£gA/V)
                            DropRngSite = PinVal.compare(LessThan, 0.002 - ((0.00004 + 0.00004) * 2))  'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVI.Pins(AutoRangePin_Ary(i)).SetCurrentAndRange 0.002, 0.002
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 1.5 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 5: '2mA=> 200uA    '2mA:¡Ó(0.20% + 4£gA + 300nA/V),
                            DropRngSite = PinVal.compare(LessThan, 0.0002 - ((0.000004 + 0.000004) * 2))    'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVI.Pins(AutoRangePin_Ary(i)).SetCurrentAndRange 0.0002, 0.0002
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 11 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                    Case 6: '200uA=> 20uA   '200uA:¡Ó(0.20% + 400nA + 30nA/V),
                            DropRngSite = PinVal.compare(LessThan, 0.00002 - ((0.0000004 + 0.0000004) * 2))  'Next IRange - Accuracy
                            If DropRngSite.Any(True) Then
                                TheExec.sites.Selected = DropRngSite
                                TheHdw.DCVI.Pins(AutoRangePin_Ary(i)).SetCurrentAndRange 0.00002, 0.00002
                                TheExec.sites.Selected = True
                            End If
                            SattleTime = 1.4 * ms
                            If SattleTime > WaitTime Then WaitTime = SattleTime
                End Select
            End If
        Next i
        '-------------------------------------------debug print
        If debug_print_pins <> "" Then
            For i = 0 To UVI80_Power_data.Pins.Count - 1
                If InStr(LCase(debug_print_pins), LCase(UVI80_Power_data.Pins(i).Name)) > 0 Then
                    For Each site In TheExec.sites
                        TheExec.Datalog.WriteComment "Site(" & site & "), " & UVI80_Power_data.Pins(i).Name & ", Step " & j & ", Irange: " & TheHdw.DCVI.Pins(UVI80_Power_data.Pins(i).Name).CurrentRange.Value & ", Current: " & UVI80_Power_data.Pins(i).Value(site)
                    Next site
                End If
            Next i
        End If
        '-------------------------------------------
        If StepNo = 6 Then j = Stop_Step
        TheHdw.Wait WaitTime
        If p_uvi80 <> "" Then UVI80_Power_data = TheHdw.DCVI.Pins(p_uvi80).Meter.Read(tlStrobe, 1, , tlDCVIMeterReadingFormatAverage)
    Next j
    '========================================================================================
    For i = 0 To UBound(A_UVI80)
        Power_data.AddPin (A_UVI80(i))
        Power_data.Pins(A_UVI80(i)) = UVI80_Power_data.Pins(A_UVI80(i))
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                Power_data.Pins(A_UVI80(i)).Value(site) = 0.0005 + Rnd() * 0.0001
            Next site
        End If
    Next i

    For i = 0 To CorePowerPin_Cnt - 1: For j = 0 To repeat_count - 1
        If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) <> "N/C" Then

            ''20180315 Abel change naming'            Tname = TheExec.DataManager.InstanceName & "_" & j
            Tname = TheExec.DataManager.instanceName 'No need _0
            If TheExec.DataManager.ChannelType(CorePowerPin_Ary(i)) Like "*DCVI*" Then
                Vmain = Format(TheHdw.DCVI.Pins(Power_data.Pins(CorePowerPin_Ary(i))).Voltage, "0.00")
            End If

            If TPModeAsCharz_GLB = True Then 'wc 180319 for charz
                TheExec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", ForceVal:=Vmain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
            Else
                  Powerpin_log = Replace(UCase(CorePowerPin_Ary(i)), "_", "")
                  TheExec.Flow.TestLimit resultVal:=Power_data.Pins(CorePowerPin_Ary(i)), scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", ForceVal:=Vmain, ForceUnit:=unitVolt, ForceResults:=tlForceFlow   ''20180315 Abel add power pin name
            End If
        End If
    Next j: Next i

    For i = 0 To UBound(Pin_Ary)
        If TheExec.DataManager.ChannelType(Pin_Ary(i)) Like "*DCVI*" Then
            TheHdw.DCVI.Pins(Pin_Ary(i)).SetCurrentAndRange IDS_ini_Current_range(i), IDS_ini_Current_range(i)
        End If
    Next i

    TheHdw.Wait 0.003

    Exit Function
    
errorEmpty:
   TheExec.Datalog.WriteComment ("Please check your Hilimit(FlowLimitsInfo.GetHighLimits):: Val_Hi(0) = """)
    Exit Function
    
    
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function IDS_main_current(patt As Pattern, _
                                      DCVS_Power_Pin As PinList, DCVI_Power_Pin As PinList, _
                                      DCVS_OtherPower_Pin As PinList, _
                                      repeat_count As Long, _
                                      FlowLimitForInitIRange As Boolean, _
                             Optional Search_Step As String, _
                             Optional DisableClock As Boolean = False, _
                             Optional FlagWait As Boolean = False, _
                             Optional OtherPowerAutoRange As Boolean = False, _
                             Optional CharInputString As String, _
                             Optional HexInitIRange1A_Pins As PinList, _
                             Optional HexInitIRange100mA_Pins As PinList, _
                             Optional UVSInitIRange800mA_Pins As PinList, _
                             Optional UVSInitIRange200mA_Pins As PinList, _
                             Optional UVSInitIRange20mA_Pins As PinList, _
                             Optional UVSInitIRange2mA_Pins As PinList, _
                             Optional UVSInitIRange200uA_Pins As PinList, _
                             Optional UVI80InitIRange1A_Pins As PinList, _
                             Optional UVI80InitIRange200mA_Pins As PinList, _
                             Optional UVI80InitIRange20mA_Pins As PinList, _
                             Optional UVI80InitIRange2mA_Pins As PinList, _
                             Optional UVI80InitIRange200uA_Pins As PinList, _
                             Optional DisableClockPortName As String, _
                             Optional RTOS_Setup As Boolean = False, Optional DisconnectClock As Boolean = True, Optional debug_print_pins As String, Optional NotPrintOutLimit As Boolean = False, Optional Interpose_Meas_before As String, Optional Interpose_Meas_after As String, _
                             Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", Optional CUS_Str_DigSrcData As String = "", Optional DictName As String, Optional Fuse_StoreName As String, Optional Validating_ As Boolean) 'Carter, 20190315


    Dim p As Variant
    Dim p_ary() As String
    Dim PinCnt As Long
    Dim MeasCurr_HexVS As New PinListData
    Dim MeasCurr As New PinListData
    Dim MeasCurr_copy As New PinListData
    Dim Power_pin As String
    Dim TestNum() As Long, Cnt1 As Long
    Dim i As Long, j As Long
    Dim repeat_judge As Long

    Dim All_Power_data As New PinListData
    Dim site As Variant

    Dim AllSitePass As Boolean
    Dim BurstResult As New SiteLong
    Dim CLK_Pins As String

    Dim rtnPatNames() As String
    Dim PatCnt As Long
    Dim InDSPwave As New DSPWave

    On Error GoTo errHandler
    
    If Validating_ Then 'Carter, 20190315
        Call PrLoadPattern(patt.Value)
        Exit Function    ' Exit after validation
    End If
    
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    'Add SPI init cindition. Merge NAND IDS and SPI IDS in this module. 20160903 ylliuj
'    If RTOS_Setup = True Then
'        SPI_Initial_Conds_Fun
'    End If

'    If UCase(TheExec.CurrentJob) Like "*CHAR*" Then
'        If CharInputString <> "" Then
'            Call SetForceCondition(CharInputString)
'        End If
'    End If

    'TheHdw.Digital.Patgen.TimeoutEnable = False
    TheHdw.Digital.Patgen.TimeOut = 10
    Call TheHdw.Patterns(patt).Load
    '-------------------------------------------DSSC Soruce
    If DigSrc_pin <> "" Then
       rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(patt, PatCnt)
        Call GeneralDigSrcSetting(CStr(rtnPatNames(0)), DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, _
                                DigSrc_Assignment, DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPwave)
    End If
    '-------------------------------------------

    If FlagWait = True Then
        Call TheHdw.Patterns(patt).start
        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)  'Meas during CPUA loop
    Else
        Call TheHdw.Patterns(patt).Test(pfAlways, 0, tlResultModeDomain)
        TheHdw.Digital.Patgen.HaltWait
    End If
    '-------------------------------------------AP & RF
    If DisableClock = True Then
        Call Disable_FRC(DisableClockPortName, DisconnectClock)
        Wait 0.005
    End If
    '-------------------------------------------pre_measure_store
    If Interpose_Meas_before <> "" Then
        Call SetForceCondition(Interpose_Meas_before & ";STOREPREMEAS")
    End If
    '-------------------------------------------

    Set All_Power_data_IDS_GB = Nothing
    If DCVS_Power_Pin <> "" Then
        DCVS_IDS_main_auto_range_and_measure CStr(DCVS_Power_Pin), CStr(DCVS_OtherPower_Pin), All_Power_data, repeat_count, FlowLimitForInitIRange, Search_Step, OtherPowerAutoRange, HexInitIRange1A_Pins, HexInitIRange100mA_Pins, UVSInitIRange800mA_Pins, UVSInitIRange200mA_Pins, UVSInitIRange20mA_Pins, UVSInitIRange2mA_Pins, UVSInitIRange200uA_Pins, debug_print_pins
    End If
    If DCVI_Power_Pin <> "" Then
        DCVI_IDS_main_auto_range_and_measure CStr(DCVI_Power_Pin), All_Power_data, repeat_count, FlowLimitForInitIRange, Search_Step, OtherPowerAutoRange, UVI80InitIRange1A_Pins, UVI80InitIRange200mA_Pins, UVI80InitIRange20mA_Pins, UVI80InitIRange2mA_Pins, UVI80InitIRange200uA_Pins, debug_print_pins
    End If

    All_Power_data_IDS_GB = All_Power_data.Copy
    Wait 0.005

'''    Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
'''    '============================== For efuse =================================
'''    'collect measure values
'''    If eFusePower_Pin <> "" Then
'''        TheExec.DataManager.DecomposePinList eFusePower_Pin, p_ary, PinCnt
'''        For Each Site In TheExec.sites
'''            For i = 0 To PinCnt - 1
'''                'For eFuse category naming rule was fixed
'''                Call auto_eFuse_IDS_SetWriteDecimal("CFG", "ids_" + LCase(p_ary(i)), All_Power_data.Pins(p_ary(i)).value(Site))
'''            Next i
'''        Next Site
'''    End If
'''    '===========================================================================
    '-------------------------------------------AP & RF
    If DisableClock = True Then
        Call Enable_FRC(DisableClockPortName, DisableClock)
        'If DebugFlag = True Then TheExec.Datalog.WriteComment "print: nWire connect, pin " & DisableClockPortName
    End If
    '-------------------------------------------
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
    TheHdw.Digital.Patgen.HaltWait
    
    If FlagWait = True Then
        Call HardIP_WriteFuncResult
    End If

    'TheHdw.Digital.Patgen.HaltWait
    DebugPrintFunc patt.Value
    '-------------------------------------------restore pre_measure
    If Interpose_Meas_after <> "" Then
        Call SetForceCondition(Interpose_Meas_after)
    ElseIf Interpose_Meas_before <> "" Then
        Call SetForceCondition("RESTOREPREMEAS")
    Else
    
    End If
    '-------------------------------------------
    'TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD    'clean all cpu flag
    If DictName <> "" Then Call AddStoredMeasurement(DictName, All_Power_data)
    If Fuse_StoreName <> "" Then Call IDS_Store2Dic(Fuse_StoreName, CStr(DCVS_Power_Pin), All_Power_data, patt)

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in IDS_Main_current"
    If AbortTest Then Exit Function Else Resume Next

End Function

